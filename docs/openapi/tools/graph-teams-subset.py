#!/usr/bin/env python3
"""Cut the Microsoft Teams surface out of Microsoft Graph's v1.0 OpenAPI description.

usage: graph-teams-subset.py [REF [SOURCE]]

REF is a microsoftgraph/msgraph-metadata git ref (default: master). It is resolved to a
commit SHA, which is recorded in info.x-source. SOURCE is a local copy of
openapi/v1.0/openapi.yaml at that commit; without it the file is downloaded.
Writes ../graph-teams-v1.0.yaml next to this tools/ directory.
"""
import datetime
import json
import sys
import urllib.request
from pathlib import Path

import yaml

# An entry keeps its path and every path below it; a trailing "$" keeps the exact path only.
PREFIXES = [
    "/chats",
    "/teams",
    "/teamwork",
    "/appCatalogs/teamsApps",
    "/me/chats",
    "/me/joinedTeams",
    "/me/onlineMeetings",
    "/me/presence",
    "/users/{user-id}/chats",
    "/users/{user-id}/joinedTeams",
    "/users/{user-id}/onlineMeetings",
    "/users/{user-id}/presence",
    "/users/{user-id}/teamwork",
    "/communications",
    "/shares/{sharedDriveItem-id}/driveItem",
    "/drives/{drive-id}/items/{driveItem-id}$",
    "/drives/{drive-id}/items/{driveItem-id}/content$",
    "/drives/{drive-id}/items/{driveItem-id}/createLink$",
    "/drives/{drive-id}/items/{driveItem-id}/invite$",
    "/drives/{drive-id}/items/{driveItem-id}/createUploadSession$",
    "/drives/{drive-id}/items/{driveItem-id}/children$",
]

# Every entity type derives from this schema, so its discriminator mapping lists the whole API.
# All other mappings are followed, keeping polymorphic subtypes such as eventMessageDetail's.
ROOT_ENTITY = "microsoft.graph.entity"

REPO = "microsoftgraph/msgraph-metadata"
OUT = Path(__file__).resolve().parents[1] / "graph-teams-v1.0.yaml"


def wanted(path):
    return any(
        (path == p[:-1]) if p.endswith("$") else (path == p or path.startswith(p + "/"))
        for p in PREFIXES
    )


def refs(node, mappings=True):
    """Yield the $ref targets in node, plus discriminator mapping targets if mappings is set."""
    if isinstance(node, dict):
        if "$ref" in node:
            yield node["$ref"]
        if mappings and "discriminator" in node:
            yield from node["discriminator"].get("mapping", {}).values()
        for value in node.values():
            yield from refs(value, mappings)
    elif isinstance(node, list):
        for value in node:
            yield from refs(value, mappings)


def closure(components, roots):
    seen, todo = set(), list(refs(roots))
    while todo:
        ref = todo.pop()
        if ref not in seen:
            seen.add(ref)
            _, _, kind, name = ref.split("/")
            todo.extend(refs(components[kind][name], mappings=name != ROOT_ENTITY))
    return seen


def main():
    ref = sys.argv[1] if len(sys.argv) > 1 else "master"
    sha = json.load(urllib.request.urlopen(f"https://api.github.com/repos/{REPO}/commits/{ref}"))["sha"]
    url = f"https://raw.githubusercontent.com/{REPO}/{sha}/openapi/v1.0/openapi.yaml"
    text = Path(sys.argv[2]).read_bytes() if len(sys.argv) > 2 else urllib.request.urlopen(url).read()
    src = yaml.load(text, Loader=getattr(yaml, "CSafeLoader", yaml.SafeLoader))

    paths = {p: item for p, item in src["paths"].items() if wanted(p)}
    kept = closure(src["components"], paths)
    components = {}
    for kind, entries in src["components"].items():
        chosen = {name: c for name, c in entries.items() if f"#/components/{kind}/{name}" in kept}
        if chosen:
            components[kind] = chosen
    disc = components["schemas"][ROOT_ENTITY]["discriminator"]
    disc["mapping"] = {k: v for k, v in disc["mapping"].items() if v in kept}
    ops = [op for item in paths.values() for op in item.values() if isinstance(op, dict)]
    used_tags = {tag for op in ops for tag in op.get("tags", [])}
    today = datetime.datetime.now(datetime.timezone.utc).date().isoformat()

    out = {
        **src,
        "info": {
            **src["info"],
            "title": "Microsoft Graph v1.0 - Teams subset",
            "x-source": {"url": url, "commit": sha, "generated": today},
        },
        "paths": paths,
        "components": components,
        "tags": [t for t in src["tags"] if t["name"] in used_tags],
    }
    dumper = getattr(yaml, "CSafeDumper", yaml.SafeDumper)
    with OUT.open("w", encoding="utf-8") as f:
        yaml.dump(out, f, Dumper=dumper, sort_keys=False, allow_unicode=True)
    print(f"{OUT}: {len(paths)} paths, {len(ops)} operations, {len(components['schemas'])} schemas, commit {sha}")


if __name__ == "__main__":
    main()
