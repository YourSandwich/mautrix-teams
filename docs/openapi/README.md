# OpenAPI descriptions

Two OpenAPI documents describe the HTTP surfaces a Teams client talks to.

| File | What it covers | Source |
|---|---|---|
| `graph-teams-v1.0.yaml` | The Teams-related subset of Microsoft Graph v1.0: chats, teams and channels, teamwork, the Teams app catalog, online meetings, presence, communications, and the drive and shares endpoints used for files | Generated from Microsoft's published Graph description |
| `teams-webclient.yaml` | A reference of the undocumented Teams web-client APIs for work and personal accounts, with example data: sign-in and tokens, authz and host discovery, chat service, AMS media, middle tier (profiles, pictures, calendar, Meet Now), CSA, people search, person card, SharePoint and OneDrive upload and download, meeting recordings and transcripts, Trouter push and the registrar, presence reads and writes, and calls and meetings (call control, roster, media renegotiation, video). `x-bridge-usage` tells which operations the bridge sends | Written by hand from source code and HAR captures of the Teams web client |

## The web-client API is not a contract

Microsoft does not document the APIs in `teams-webclient.yaml`. They are what the Teams web client
happens to call today, reverse-engineered from this repository's `pkg/msteams`, from other
open-source clients and from HAR captures of the web client. Any path, header or field can change or
disappear without notice, and shapes differ between tenants, regions and client versions. Treat the
file as a map of what the bridge does and what others have seen, not as a guarantee of server
behaviour.

## Evidence conventions in `teams-webclient.yaml`

- Every operation carries `x-bridge-usage`: `used` when `pkg/msteams` sends the request,
  `documented-only` when only another client's source or a capture shows it. Parameters, schema
  properties, servers and websocket frames carry the same marker where they differ from their
  operation.
- Every operation and schema carries `x-evidence`, a list of where each claim comes from:
  - `pkg/msteams/messages.go:FetchAttachment` (or `pkg/teamsmedia/...`, `pkg/connector/...`)
    points at this repository by function, type or constant name, so the reference survives edits.
  - `purple-teams@f7eb519 teams_messages.c:2978-3031` points at another project's file and lines
    at a pinned commit. Sources used: EionRobb/purple-teams `f7eb519`, eisbaw/ost `0892144`,
    IanTerzo/Squads `f977552`, fossteams/teams-api `60cd5c1` (including its JSON fixtures).
  - `captured: enterprise/calls` (and the other labels in the `info` "Captures" table) points at
    a HAR capture of the Teams web client, `enterprise/*` from `teams.cloud.microsoft` and
    `consumer/*` from `teams.live.com`, made on 2026-09-26. The exports strip `Authorization` and
    `Cookie`, so the bearer audience of a captured request is unknown unless noted. The captures
    are not committed; only shapes and synthetic stand-ins for their values appear in the file.
  - `observed: ...` is something the bridge or a manual request saw on the wire that no capture or
    source file holds, for example one authz `regionGtms` response recorded from a real tenant on
    2026-09-22 (only host names from it appear), the device-code client matrix, and call-control
    error codes the bridge logged.
- A field holding JSON inside a JSON string is `type: string`, with `x-embedded-json-schema`
  pointing at the decoded shape.
- OpenAPI 3.0 cannot describe websockets. The Trouter socket's frames live in the
  `x-websocket-frames` extension of the `trouterWebSocket` operation, and the bodies call control
  posts to Trouter are schemas referenced from there.
- URLs a server hands out at runtime (the calling links) are modelled as a server whose whole URL
  is one variable.
- Every example value is synthetic: ids follow `00000000-0000-0000-0000-00000000000n`
  (`8:orgid:...`, `8:live:.cid.0123456789abcdef`, `19:abc123@thread.v2`), people are
  `Alex Example` and `Sam Sample` at `contoso.example`, IP addresses come from `192.0.2.0/24` and
  `203.0.113.0/24`, and tokens, SRTP keys and passcodes are placeholders. The only real ids are
  Microsoft's own: first-party client ids, resource ids, the Microsoft-account tenant id and
  well-known bot MRIs. No real user, tenant, message or call data belongs in the file.

## Linting

```sh
redocly lint docs/openapi/teams-webclient.yaml \
  --skip-rule no-path-trailing-slash \
  --skip-rule operation-4xx-response \
  --skip-rule no-unused-components
python3 -c 'import yaml; from openapi_spec_validator import validate; validate(yaml.safe_load(open("docs/openapi/teams-webclient.yaml")))'
```

The skipped rules conflict with accuracy: several real paths end in `/` (`/socket.io/1/`,
`/v1/presence/getpresence/`, the presence `/v1/me/...` writes and others), most operations have no
4xx case a client treats specially (a `default` response describes the generic handling), and
schemas referenced only from `x-websocket-frames`, `x-embedded-json-schema` or a description (the
`Event/Call` XML) look unused to redocly. The second command needs PyYAML and
`openapi-spec-validator`.

## Regenerating `graph-teams-v1.0.yaml`

`tools/graph-teams-subset.py` cuts the Teams paths out of Microsoft's full Graph description
(`openapi/v1.0/openapi.yaml` in `microsoftgraph/msgraph-metadata`) and keeps every component they
reference. It needs PyYAML.

```sh
python3 docs/openapi/tools/graph-teams-subset.py            # latest master
python3 docs/openapi/tools/graph-teams-subset.py <ref>      # a branch, tag or commit
python3 docs/openapi/tools/graph-teams-subset.py <ref> <local openapi.yaml>
```

The ref is resolved to a commit SHA through the GitHub API and recorded in `info.x-source`, and the
output is written next to the `tools/` directory. The kept path prefixes are the `PREFIXES` list at
the top of the script; edit that list to widen or narrow the subset.
