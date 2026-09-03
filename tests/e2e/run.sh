#!/usr/bin/env bash
#
# Runs every e2e suite against the .docker dev stack and leaves the raw
# evidence in tests/e2e/out/.
#
# Usage:  bash tests/e2e/run.sh
set -uo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
rc=0
for s in "$here"/[0-9][0-9]-*.sh; do
  echo "──────────────────────────────────────────────────────────────"
  bash "$s" || rc=1
  echo
done
exit "$rc"
