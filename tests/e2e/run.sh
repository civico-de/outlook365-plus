#!/usr/bin/env bash
#
# Runs every e2e suite against the .docker dev stack and leaves the raw
# evidence in tests/e2e/out/, cleared first so nothing is read from an
# earlier run.
#
# Usage:  bash tests/e2e/run.sh
set -uo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
. "$here/lib.sh"

# A suite run against a half-provisioned stack produces failures that look
# like defects. Assert the precondition instead.
if [ "$(curl -s -o /dev/null -w '%{http_code}' "$BASE_URL/civicrm")" = "000" ]; then
  echo "The stack at $BASE_URL is not answering. Start it with:" >&2
  echo "  cd .docker && docker compose up -d" >&2
  exit 1
fi
for ext in outlook365 dataprocessor keyring authx; do
  if ! cvx api4 Extension.get \
    "{\"where\":[[\"key\",\"=\",\"$ext\"],[\"status\",\"=\",\"installed\"]],\"checkPermissions\":false}" \
    | grep -q '"key"'; then
    echo "Extension $ext is not installed — the stack is still provisioning." >&2
    exit 1
  fi
done

rm -rf "${OUT_DIR:?}"/*
rc=0
for s in "$here"/[0-9][0-9]-*.sh; do
  echo "──────────────────────────────────────────────────────────────"
  bash "$s" || rc=1
  echo
done
exit "$rc"
