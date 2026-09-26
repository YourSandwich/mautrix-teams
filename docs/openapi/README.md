# OpenAPI descriptions

Two OpenAPI documents describe the HTTP surfaces the bridge talks to.

| File | What it covers | Source |
|---|---|---|
| `graph-teams-v1.0.yaml` | The Teams-related subset of Microsoft Graph v1.0: chats, teams and channels, teamwork, the Teams app catalog, online meetings, presence, communications, and the drive and shares endpoints used for files | Generated from Microsoft's published Graph description |
| `teams-webclient.yaml` | The undocumented Teams web-client APIs the bridge uses: authz, chat service, AMS media, middle tier, CSA, people search, person card, SharePoint download, Trouter and the registrar, presence subscriptions and outgoing call signalling, plus a few endpoints only other clients use (presence polling, ost's GET variant of the Trouter bootstrap) | Written by hand from source code |

## The web-client API is not a contract

Microsoft does not document the APIs in `teams-webclient.yaml`. They are what the Teams web client
happens to call today, reverse-engineered from this repository's `pkg/msteams` and from other
open-source clients. Any path, header or field can change or disappear without notice, and shapes
differ between tenants, regions and client versions. Treat the file as a map of what the bridge
does and what others have seen, not as a guarantee of server behaviour.

## Evidence conventions in `teams-webclient.yaml`

- Every operation carries `x-bridge-usage`: `used` when `pkg/msteams` sends the request,
  `documented-only` when only another client's source shows it. Parameters, schema properties,
  servers and websocket frames carry the same marker where they differ from their operation.
- Every operation and schema carries `x-evidence`, a list of where each claim comes from:
  - `pkg/msteams/messages.go:FetchAttachment` (or `pkg/teamsmedia/...`, `pkg/connector/...`)
    points at this repository by function, type or constant name, so the reference survives edits.
  - `purple-teams@f7eb519 teams_messages.c:2978-3031` points at another project's file and lines
    at a pinned commit. Sources used: EionRobb/purple-teams `f7eb519`, eisbaw/ost `0892144`,
    IanTerzo/Squads `f977552`, fossteams/teams-api `60cd5c1` (including its JSON fixtures).
  - `observed: ...` refers to one authz `regionGtms` response recorded from a real tenant on
    2026-09-22. It is not committed; only host names from it appear in the file.
- A field holding JSON inside a JSON string is `type: string`, with `x-embedded-json-schema`
  pointing at the decoded shape.
- OpenAPI 3.0 cannot describe websockets. The Trouter socket's frames live in the
  `x-websocket-frames` extension of the `trouterWebSocket` operation.
- URLs a server hands out at runtime (the calling links) are modelled as a server whose whole URL
  is one variable.
- Examples use placeholder ids such as `8:orgid:00000000-0000-0000-0000-000000000001` and
  `19:abc123@thread.v2`. No real user, tenant or message data belongs in the file.

## Linting

```sh
redocly lint docs/openapi/teams-webclient.yaml \
  --skip-rule no-path-trailing-slash \
  --skip-rule operation-4xx-response \
  --skip-rule no-unused-components
```

The skipped rules conflict with accuracy: two real paths end in `/` (`/socket.io/1/`,
`/v1/presence/getpresence/`), most operations have no 4xx case the code treats specially (a
`default` response describes the generic handling), and schemas referenced only from
`x-websocket-frames` or `x-embedded-json-schema` look unused to redocly.

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
