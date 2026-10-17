# shellcheck shell=bash
# The add-in's own e2e helpers, loaded by the managed lib.sh. Tokens never
# reach a command line: hit() hands curl its whole configuration on stdin.

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
  cv api4 KeyringToken.create \
    "{\"contactId\":$contact,\"label\":\"outlook365-e2e\",\"scope\":$scope,\"mode\":\"enforce\",\"checkPermissions\":false}" \
    | python3 -c 'import json,sys; row = json.load(sys.stdin)[0]; print(row["id"], row["token"])'
}

revoke_token() { # revoke_token <id>
  cv api4 KeyringToken.revoke "{\"where\":[[\"id\",\"=\",$1]],\"checkPermissions\":false}" >/dev/null
}

# The subject the add-in authenticates as: a NON-superuser contact with a
# Standalone user and a role holding the permissions the add-in needs.
# Standalone's admin holds 'all CiviCRM permissions and ACLs', which
# short-circuits CRM_Core_Permission::check before a scope is ever consulted.
seed_subject() {
  cv_scr "$CK_E2E_DIR/seed.php" | sed -n 's/^contactId=//p' | tail -1
}
