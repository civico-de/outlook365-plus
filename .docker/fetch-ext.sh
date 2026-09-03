#!/usr/bin/env bash
#
# Clone the foreign extension this add-in requires at runtime. It lands
# OUTSIDE the repository and is mounted into the dev stack by
# docker-compose.yml.
set -euo pipefail

DEST="${OUTLOOK365_EXT_DIR:-${XDG_CACHE_HOME:-$HOME/.cache}/outlook365-ext}"
mkdir -p "$DEST"

clone() { # clone <dir> <url>
  if [ -d "$DEST/$1/.git" ]; then
    echo "  = $DEST/$1 (present)"
  else
    echo "  + $DEST/$1"
    git clone --depth 1 "$2" "$DEST/$1"
  fi
}

clone dataprocessor https://lab.civicrm.org/extensions/dataprocessor.git

echo
echo "Foreign sources in $DEST"
