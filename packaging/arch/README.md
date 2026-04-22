# Arch Linux packaging

Two PKGBUILDs live here:

- `PKGBUILD` - for the eventual AUR release. Pulls the source tarball from a
  `v$pkgver` GitHub tag; only useful once releases are cut.
- `PKGBUILD-git` - builds straight from the current git worktree. This is the
  one to use for testing changes locally or on a server.

## Why they live in a subdirectory

`makepkg` creates `src/` and `pkg/` next to the PKGBUILD as its build + staging
directories. The Go source tree has its own `pkg/` (standard Go layout), so
running `makepkg` from the repository root would overwrite the Go source.
Keeping packaging under `packaging/arch/` avoids the collision: makepkg's
`pkg/` ends up at `packaging/arch/pkg/`, which is `.gitignore`d.

## Building the `-git` variant

```sh
cd packaging/arch
makepkg -si             # build and install
makepkg -f              # rebuild (e.g. after a source change)
```

`pkgver()` automatically derives the version from `git rev-list` + `rev-parse`.

Note: `makepkg` rewrites the `pkgver=` line in-place after running `pkgver()`.
The committed value is always `0.1.0.r0.unknown`; after a build the working
tree will show a dirty PKGBUILD-git. Reset it with
`git checkout -- PKGBUILD-git` before committing other changes.

## Building the release variant

Do not use this before a `v0.1.0` tag exists on GitHub - the source fetch
will 404. Once a tag is cut:

```sh
cd packaging/arch
makepkg -si
```

## Clean up

```sh
cd packaging/arch
rm -rf src pkg *.pkg.tar.zst
```
