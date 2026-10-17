# shellcheck shell=bash
# Shared helpers for the add-in's e2e suites. Source right after
# `set -uo pipefail`:
#
#   . "$(dirname "$0")/lib.sh"
#
# Provides the env knobs, PASS/FAIL bookkeeping (check/finish), the cv and
# curl plumbing, and token minting. Tokens never reach a command line: curl
# reads its whole configuration, headers included, from stdin.

COMPOSE_DIR="${COMPOSE_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.docker" && pwd)}"
BASE_URL="${BASE_URL:-http://localhost:8200}"
OUT_DIR="${OUT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/out}"
mkdir -p "$OUT_DIR"

PASS=0
FAIL=0

check() { # check <name> <condition-result(0/1)> [detail]
  if [ "$2" = "0" ]; then
    echo "  ✔ $1"; PASS=$((PASS + 1))
  else
    echo "  ✘ $1  — ${3:-}"; FAIL=$((FAIL + 1))
  fi
}

note() { echo "  · $*"; }

finish() {
  echo
  echo "Result: $PASS passed, $FAIL failed"
  [ "$FAIL" = 0 ]
}

cvx() { # cvx <cv args...>
  (cd "$COMPOSE_DIR" && docker compose exec -T app cv "$@" 2>/dev/null)
}

appsh() { # appsh <shell code> — run in the app container
  (cd "$COMPOSE_DIR" && docker compose exec -T app bash -c "$1")
}

sql() { # sql <query> — PII-free queries only: ids, counts, metadata
  (cd "$COMPOSE_DIR" && docker compose exec -T db mariadb -N -B -ucivicrm -pcivicrm civicrm -e "$1" 2>/dev/null)
}

site_key() {
  cvx ev 'echo CIVICRM_SITE_KEY;'
}

# ── requests ─────────────────────────────────────────────────────────────
# hit <label> <url> [extra curl config lines...]
#
# The token is passed through $TOKEN in the environment and written into
# curl's stdin config, so it appears in no process list.
hit() {
  local label="$1" url="$2"; shift 2
  rm -f "$OUT_DIR/$label.body"
  {
    printf 'url = "%s"\n' "$url"
    printf 'output = "%s"\n' "$OUT_DIR/$label.body"
    printf 'write-out = "%%{http_code}"\n'
    printf 'silent\ngloboff\n'
    printf '%s\n' "$@"
  } | curl -K -
}

# body_lacks <label> <extended-regex> — 0 when the body exists, is non-empty
# and does not match. A missing body fails instead of passing silently.
body_lacks() {
  local f="$OUT_DIR/$1.body"
  [ -s "$f" ] || return 1
  ! grep -qE "$2" "$f"
}

auth_header() { # auth_header [header-name] — a curl config line carrying the token
  printf 'header = "%s: Bearer %s"' "${1:-X-Civi-Auth}" "$TOKEN"
}

# cfg_value <string> — escape for a double-quoted curl config value. JSON
# payloads carry quotes, which would otherwise close the value early.
cfg_value() {
  printf '%s' "$1" | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g'
}

# ── the shapes the add-in sends ──────────────────────────────────────────
# assets/civicrm.js posts entity/action/json as a form body to
# civicrm/ajax/rest, and params to civicrm/ajax/api4/<Entity>/<action>,
# with the credential in a header and nothing else.
addin_api3() { # addin_api3 <label> <entity> <action> <json> [extra curl lines...]
  local label="$1" entity="$2" action="$3" json="$4"; shift 4
  hit "$label" "$BASE_URL/civicrm/ajax/rest" \
    "$(printf 'data-urlencode = "entity=%s"' "$entity")" \
    "$(printf 'data-urlencode = "action=%s"' "$action")" \
    "$(printf 'data-urlencode = "json=%s"' "$(cfg_value "$json")")" \
    "$@"
}

addin_api4() { # addin_api4 <label> <entity> <action> <params> [extra curl lines...]
  local label="$1" entity="$2" action="$3" params="$4"; shift 4
  hit "$label" "$BASE_URL/civicrm/ajax/api4/$entity/$action" \
    "$(printf 'data-urlencode = "params=%s"' "$(cfg_value "$params")")" \
    "$@"
}

# ── tokens ───────────────────────────────────────────────────────────────
# mint_token <contactId> <permission>...
# Prints "<id> <plaintext>".
mint_token() {
  local contact="$1"; shift
  local scope
  scope="$(printf '%s\n' "$@" | python3 -c 'import json,sys; print(json.dumps([l.rstrip("\n") for l in sys.stdin if l.strip()]))')"
  cvx api4 KeyringToken.create \
    "{\"contactId\":$contact,\"label\":\"outlook365-e2e\",\"scope\":$scope,\"mode\":\"enforce\",\"checkPermissions\":false}" \
    | python3 -c 'import json,sys; row = json.load(sys.stdin)[0]; print(row["id"], row["token"])'
}

revoke_token() { # revoke_token <id>
  cvx api4 KeyringToken.revoke "{\"where\":[[\"id\",\"=\",$1]],\"checkPermissions\":false}" >/dev/null
}

# The subject the add-in authenticates as: a NON-superuser contact with a
# Standalone user and a role holding the permissions the add-in needs.
# Standalone's admin holds 'all CiviCRM permissions and ACLs', which
# short-circuits CRM_Core_Permission::check before a scope is ever consulted.
seed_subject() {
  appsh 'cat > /tmp/outlook365-seed.php && cv scr /tmp/outlook365-seed.php' < "$(dirname "${BASH_SOURCE[0]}")/seed.php" \
    | sed -n 's/^contactId=//p' | tail -1
}
