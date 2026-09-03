#!/usr/bin/env bash
#
# What a deployment has to satisfy for the header flow, now that the add-in
# sends no site key: the authx perm guard, and same-origin requests.
#
# Usage: bash tests/e2e/30-deployment.sh
set -uo pipefail
. "$(dirname "$0")/lib.sh"

CONTACT_ID="${CONTACT_ID:-$(seed_subject)}"
KEY="$(site_key)"
V3_CONTACTS='{"sequential":1,"options":{"limit":1}}'

echo "outlook365 e2e — deployment requirements"
echo

guards="$(cvx ev 'echo json_encode(Civi::settings()->get("authx_guards"));')"
check "authx_guards contains perm" \
  "$(printf '%s' "$guards" | grep -q '"perm"' && echo 0 || echo 1)" "got $guards"
for flow in xheader header; do
  cred="$(cvx ev "echo json_encode(Civi::settings()->get('authx_${flow}_cred'));")"
  check "authx_${flow}_cred accepts api_key" \
    "$(printf '%s' "$cred" | grep -q '"api_key"' && echo 0 || echo 1)" "got $cred"
done

read -r TOKEN_ID TOKEN < <(mint_token "$CONTACT_ID" 'access CiviCRM' 'access AJAX API' 'view all contacts')
export TOKEN

# The guards are OR'd, so a token whose contact holds 'authenticate with api
# key' needs no site key. Take the permission away and the same request must
# fail — that is why the fork can drop the site key but the deployment cannot
# drop the permission.
role_perms() { # role_perms <json-array>
  cvx ev "\Civi\Api4\Role::update(FALSE)->addWhere('name','=','outlook365_e2e')
    ->addValue('permissions', json_decode('$1', TRUE))->execute();" >/dev/null
}
without_auth="$(cvx ev "
  \$r = \Civi\Api4\Role::get(FALSE)->addWhere('name','=','outlook365_e2e')->execute()->single();
  echo json_encode(array_values(array_diff(\$r['permissions'], ['authenticate with api key'])));")"
full="$(cvx ev "
  \$r = \Civi\Api4\Role::get(FALSE)->addWhere('name','=','outlook365_e2e')->execute()->single();
  echo json_encode(\$r['permissions']);")"

role_perms "$without_auth"
code="$(addin_api3 no_perm_guard Contact get "$V3_CONTACTS" "$(auth_header)")"
check "without 'authenticate with api key' the header flow is refused" \
  "$(grep -qi 'Login not permitted\|guard' "$OUT_DIR/no_perm_guard.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/no_perm_guard.body")"
code="$(addin_api3 no_perm_guard_site_key Contact get "$V3_CONTACTS" "$(auth_header)" \
  "$(printf 'header = "X-Civi-Key: %s"' "$KEY")")"
check "the site key alone satisfies the guard instead" \
  "$([ "$code" = "200" ] && grep -q '"is_error":0' "$OUT_DIR/no_perm_guard_site_key.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/no_perm_guard_site_key.body")"
role_perms "$full"
code="$(addin_api3 perm_guard_restored Contact get "$V3_CONTACTS" "$(auth_header)")"
check "restoring the permission restores the header flow" \
  "$([ "$code" = "200" ] && grep -q '"is_error":0' "$OUT_DIR/perm_guard_restored.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/perm_guard_restored.body")"

# X-Civi-Auth is not a CORS-safelisted header, so a cross-origin call would be
# preflighted. Measure whether CiviCRM answers the preflight at all.
preflight="$(curl -s -o "$OUT_DIR/preflight.head" -D - -X OPTIONS \
  -H 'Origin: https://addin.example.org' \
  -H 'Access-Control-Request-Method: POST' \
  -H 'Access-Control-Request-Headers: x-civi-auth' \
  -w '%{http_code}' "$BASE_URL/civicrm/ajax/rest")"
check "the CORS preflight carries no Access-Control-Allow-Origin" \
  "$(printf '%s' "$preflight" | grep -qi 'access-control-allow-origin' && echo 1 || echo 0)" \
  "http status $(printf '%s' "$preflight" | tail -c 3)"
note "the task pane must therefore be configured with the origin that serves it"

revoke_token "$TOKEN_ID"
finish
