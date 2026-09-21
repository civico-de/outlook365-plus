#!/usr/bin/env bash
#
# Clone the foreign extension this add-in requires at runtime. It lands
# OUTSIDE the repository and is mounted into the dev stack by
# docker-compose.yml.
set -euo pipefail

# docker-compose.yml mounts this exact expression; keep the two in step.
DEST="${OUTLOOK365_EXT_DIR:-$HOME/.cache/outlook365-ext}"
mkdir -p "$DEST"

# The image refuses a mounted extension whose version differs from its
# civikitchen.yaml extension_sources pin, so the clone checks out that commit.
checkout() { # checkout <dir> <url> <commit>
  if [ ! -d "$DEST/$1/.git" ]; then
    echo "  + $DEST/$1"
    git init -q "$DEST/$1"
    git -C "$DEST/$1" remote add origin "$2"
  fi
  if [ "$(git -C "$DEST/$1" rev-parse -q --verify HEAD || true)" = "$3" ]; then
    echo "  = $DEST/$1 @ $3"
  else
    echo "  > $DEST/$1 @ $3"
    git -C "$DEST/$1" fetch -q --depth 1 origin "$3"
    git -C "$DEST/$1" checkout -q --detach FETCH_HEAD
  fi
}

pin="$(sed -n 's#.*dataprocessor/-/archive/\([0-9a-f]\{40\}\)/.*#\1#p' "$(dirname "$0")/../civikitchen.yaml")"
[ -n "$pin" ] || { echo "no dataprocessor pin in civikitchen.yaml" >&2; exit 1; }
checkout dataprocessor https://lab.civicrm.org/extensions/dataprocessor.git "$pin"

echo
echo "Foreign sources in $DEST"
