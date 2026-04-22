# Test data and fixtures

## Synapse dev stack

`../compose.yaml` spins up Synapse + Postgres on `localhost:8008` so the bridge
can register against it. The bridge itself runs natively (built via
`./build.sh`) so iteration is fast.

First-run setup:

```sh
# 1. generate Synapse signing key and default homeserver.yaml
podman-compose --profile generate up synapse-generate

# 2. edit the generated config in the named volume if needed; make sure
#    'app_service_config_files' points at /data/appservices/*.yaml
podman volume inspect mautrix-teams-synapse   # see the path on your host

# 3. build the bridge and generate a registration
./build.sh -o mautrix-teams
./mautrix-teams -e -c config.yaml
$EDITOR config.yaml                # set server_name: matrix.test, address: http://127.0.0.1:8008
./mautrix-teams -g -c config.yaml -r registration.yaml

# 4. boot Synapse and the bridge
podman-compose up -d synapse
./mautrix-teams -c config.yaml -r registration.yaml
```

## JSON fixtures

Fixture files here are loaded by `httptest` servers in `pkg/msteams/*_test.go`
and give the unit tests realistic response shapes without reaching out to a
real Teams tenant.
