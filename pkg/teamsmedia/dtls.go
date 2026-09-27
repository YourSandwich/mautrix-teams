// mautrix-teams - A Matrix-Microsoft Teams puppeting bridge.
// Copyright (C) 2026 Sandwich
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

package teamsmedia

import (
	"context"
	"crypto/sha256"
	"crypto/tls"
	"crypto/x509"
	"errors"
	"fmt"
	"io"
	"net"
	"strings"
	"sync"
	"time"

	"github.com/pion/dtls/v3"
	"github.com/pion/dtls/v3/pkg/crypto/selfsign"
	dtlsnet "github.com/pion/dtls/v3/pkg/net"
	"github.com/pion/srtp/v3"
)

// A call directly between two Teams endpoints, such as a colleague's
// one-to-one call ringing the user, is keyed with DTLS as WebRTC does, where
// calls through Teams's media servers carry SDES keys in the SDP.

// dtlsCertificate is the bridge's DTLS certificate and its SDP fingerprint.
type dtlsCertificate struct {
	cert        tls.Certificate
	fingerprint string
}

func newDTLSCertificate() (*dtlsCertificate, error) {
	cert, err := selfsign.GenerateSelfSigned()
	if err != nil {
		return nil, err
	}
	return &dtlsCertificate{cert: cert, fingerprint: fingerprintOf(cert.Certificate[0])}, nil
}

// fingerprintOf is an SDP fingerprint: the hash and its bytes in upper-case
// hex, colon separated.
func fingerprintOf(der []byte) string {
	sum := sha256.Sum256(der)
	hex := make([]string, len(sum))
	for i, b := range sum {
		hex[i] = fmt.Sprintf("%02X", b)
	}
	return "sha-256 " + strings.Join(hex, ":")
}

// handshakeDTLS runs the DTLS handshake in the given role, checking the peer's
// certificate against the fingerprint of its SDP, and keys SRTP from it.
// Reads go through a packetMux from then on, which hands the peer's DTLS
// records to the DTLS connection.
func (t *transport) handshakeDTLS(ctx context.Context, remoteFingerprint string, client bool) error {
	mux := newPacketMux(t.conn)
	shared := []dtls.Option{
		dtls.WithCertificates(t.dtlsCert.cert),
		dtls.WithSRTPProtectionProfiles(dtls.SRTP_AES128_CM_HMAC_SHA1_80),
		dtls.WithExtendedMasterSecret(dtls.RequireExtendedMasterSecret),
		// The certificate is self-signed; the fingerprint vouches for it.
		dtls.WithInsecureSkipVerify(true),
		dtls.WithVerifyPeerCertificate(func(raw [][]byte, _ [][]*x509.Certificate) error {
			if len(raw) == 0 || !strings.EqualFold(fingerprintOf(raw[0]), remoteFingerprint) {
				return errors.New("teamsmedia: dtls certificate doesn't match the sdp's fingerprint")
			}
			return nil
		}),
	}
	pc, addr := dtlsnet.PacketConnFromConn(mux.dtls), mux.dtls.RemoteAddr()
	var conn *dtls.Conn
	var err error
	if client {
		opts := make([]dtls.ClientOption, len(shared))
		for i, o := range shared {
			opts[i] = o
		}
		conn, err = dtls.ClientWithOptions(pc, addr, opts...)
	} else {
		opts := []dtls.ServerOption{dtls.WithClientAuth(dtls.RequireAnyClientCert)}
		for _, o := range shared {
			opts = append(opts, o)
		}
		conn, err = dtls.ServerWithOptions(pc, addr, opts...)
	}
	if err != nil {
		return err
	}
	if err := conn.HandshakeContext(ctx); err != nil {
		return fmt.Errorf("dtls: %w", err)
	}
	state, ok := conn.ConnectionState()
	if !ok {
		return errors.New("teamsmedia: dtls handshake left no state")
	}
	config := srtp.Config{Profile: srtpProfile}
	if err := config.ExtractSessionKeysFromDTLS(&state, client); err != nil {
		return err
	}
	enc, err := srtp.CreateContext(config.Keys.LocalMasterKey, config.Keys.LocalMasterSalt, srtpProfile)
	if err != nil {
		return err
	}
	dec, err := srtp.CreateContext(config.Keys.RemoteMasterKey, config.Keys.RemoteMasterSalt, srtpProfile)
	if err != nil {
		return err
	}
	t.enc, t.conn = enc, mux.srtp
	t.decs.Store(&[]decryptor{{ctx: dec}})
	return nil
}

// packetMux splits a connection's packets between DTLS, whose first byte is
// 20 to 63, and SRTP and SRTCP, 128 to 191, as RFC 7983 does.
type packetMux struct {
	dtls, srtp *muxEndpoint
}

func newPacketMux(conn net.Conn) *packetMux {
	m := &packetMux{dtls: newMuxEndpoint(conn), srtp: newMuxEndpoint(conn)}
	go func() {
		defer m.dtls.closeOnce()
		defer m.srtp.closeOnce()
		buf := make([]byte, 1500)
		for {
			n, err := conn.Read(buf)
			if err != nil {
				return
			}
			switch first := buf[0]; {
			case n == 0:
			case first >= 20 && first <= 63:
				m.dtls.deliver(buf[:n])
			case first >= 128 && first <= 191:
				m.srtp.deliver(buf[:n])
			}
		}
	}()
	return m
}

// muxEndpoint is one side of a packetMux: it reads what the mux hands it and
// writes to the connection directly. Closing the connection is left to the
// transport.
type muxEndpoint struct {
	net.Conn
	packets chan []byte
	closed  chan struct{}
	once    sync.Once
}

func newMuxEndpoint(conn net.Conn) *muxEndpoint {
	return &muxEndpoint{Conn: conn, packets: make(chan []byte, 64), closed: make(chan struct{})}
}

// deliver drops the packet when the reader falls behind, as a full socket
// buffer would.
func (e *muxEndpoint) deliver(packet []byte) {
	select {
	case e.packets <- append([]byte(nil), packet...):
	default:
	}
}

func (e *muxEndpoint) closeOnce() {
	e.once.Do(func() { close(e.closed) })
}

func (e *muxEndpoint) Read(b []byte) (int, error) {
	select {
	case packet := <-e.packets:
		return copy(b, packet), nil
	case <-e.closed:
		return 0, io.EOF
	}
}

func (e *muxEndpoint) Close() error {
	e.closeOnce()
	return nil
}

// The DTLS connection sets deadlines while it retransmits; reads here end only
// with the connection.
func (e *muxEndpoint) SetDeadline(time.Time) error     { return nil }
func (e *muxEndpoint) SetReadDeadline(time.Time) error { return nil }
