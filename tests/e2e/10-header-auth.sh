#!/usr/bin/env bash
#
# The add-in's own request shapes against keyring: the token is accepted in
# the X-Civi-Auth header and refused anywhere in the URL. No site key is sent.
#
# Usage: bash tests/e2e/10-header-auth.sh
set -uo pipefail
. "$(dirname "$0")/lib.sh"

CONTACT_ID="${CONTACT_ID:-$(seed_subject)}"
KEY="$(site_key)"
V3_CONTACTS='{"sequential":1,"options":{"limit":1}}'
V4_EMAIL='{"select":["contact_id"],"limit":1}'

echo "outlook365 e2e — header authentication"
echo
note "subject contact id $CONTACT_ID"

read -r TOKEN_ID TOKEN < <(mint_token "$CONTACT_ID" \
  'access CiviCRM' 'access AJAX API' 'view all contacts' 'edit all contacts' \
  'add contacts' 'view all activities' 'edit all activities')
export TOKEN

# ── what assets/civicrm.js sends ─────────────────────────────────────────
code="$(addin_api3 v3_header Contact get "$V3_CONTACTS" "$(auth_header)")"
check "APIv3 (civicrm/ajax/rest) authenticates on X-Civi-Auth alone" \
  "$([ "$code" = "200" ] && grep -q '"is_error":0' "$OUT_DIR/v3_header.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_header.body")"

# The two entities the task pane actually lists, served by the extension's
# own data processors.
for entity in Outlook365Contact Outlook365Group; do
  code="$(addin_api3 "v3_$entity" "$entity" get '{"sequential":1,"options":{"limit":1}}' "$(auth_header)")"
  check "APIv3 $entity.get answers on the header flow" \
    "$([ "$code" = "200" ] && grep -q '"is_error":0' "$OUT_DIR/v3_$entity.body" && echo 0 || echo 1)" \
    "http=$code $(head -c 200 "$OUT_DIR/v3_$entity.body")"
done

code="$(addin_api4 v4_header Email get "$V4_EMAIL" "$(auth_header)")"
check "APIv4 (civicrm/ajax/api4) authenticates on X-Civi-Auth alone" \
  "$([ "$code" = "200" ] && grep -q '"count"' "$OUT_DIR/v4_header.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v4_header.body")"

# A write, POST, attributed to the subject — the add-in creates contacts and
# logs e-mails as activities.
code="$(addin_api4 v4_write Activity create \
  "{\"values\":{\"source_contact_id\":$CONTACT_ID,\"activity_type_id:name\":\"Email\",\"subject\":\"outlook365-e2e\"}}" \
  "$(auth_header)")"
check "an APIv4 write on the header flow succeeds" \
  "$([ "$code" = "200" ] && grep -q '"count":1' "$OUT_DIR/v4_write.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v4_write.body")"
activity_source="$(sql "SELECT ac.contact_id FROM civicrm_activity_contact ac
  JOIN civicrm_activity a ON a.id = ac.activity_id
  WHERE a.subject = 'outlook365-e2e' AND ac.record_type_id = 2 ORDER BY a.id DESC LIMIT 1")"
check "the write is attributed to the token's contact" \
  "$([ "$activity_source" = "$CONTACT_ID" ] && echo 0 || echo 1)" \
  "source contact $activity_source, expected $CONTACT_ID"

# ── the shapes the fork removed ──────────────────────────────────────────
# Upstream put the token in api_key= on every APIv3 call.
code="$(hit v3_url_token "$BASE_URL/civicrm/ajax/rest?entity=Contact&action=get&api_key=$TOKEN&key=$KEY&json=%7B%22sequential%22%3A1%7D")"
check "the same token in api_key= is refused" \
  "$(grep -qi 'must be sent as a request header' "$OUT_DIR/v3_url_token.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_url_token.body")"
check "the refusal names the header to use" \
  "$(grep -q 'X-Civi-Auth' "$OUT_DIR/v3_url_token.body" && echo 0 || echo 1)" \
  "$(head -c 200 "$OUT_DIR/v3_url_token.body")"

# Upstream's api4() helper put the token in the query string of the APIv4 URL.
code="$(hit v4_url_token "$BASE_URL/civicrm/ajax/api4/Email/get?api_key=$TOKEN&key=$KEY&params=%7B%22select%22%3A%5B%22contact_id%22%5D%7D")"
check "the same token in the APIv4 query string is refused" \
  "$(grep -qi 'must be sent as a request header' "$OUT_DIR/v4_url_token.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v4_url_token.body")"

# A correct header does not excuse a token that also rode in the URL: by then
# it is in the access log and the browser history.
code="$(hit v3_header_and_url "$BASE_URL/civicrm/ajax/rest?entity=Contact&action=get&api_key=$TOKEN&json=%7B%22sequential%22%3A1%7D" "$(auth_header)")"
check "a token in the header AND in the URL is refused" \
  "$(grep -qi 'must be sent as a request header' "$OUT_DIR/v3_header_and_url.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_header_and_url.body")"

# The token must not travel in the POST body either.
code="$(addin_api3 v3_body_token Contact get "$V3_CONTACTS" \
  "$(printf 'data-urlencode = "api_key=%s"' "$TOKEN")")"
check "a token in the POST body is refused" \
  "$(grep -qi 'must be sent as a request header' "$OUT_DIR/v3_body_token.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_body_token.body")"

# ── the site key the fork stopped sending ────────────────────────────────
check "no request in this suite carried the site key" \
  "$(grep -rqs "$KEY" "$OUT_DIR"/v3_header.body "$OUT_DIR"/v4_header.body && echo 1 || echo 0)"
code="$(addin_api3 v3_with_site_key Contact get "$V3_CONTACTS" "$(auth_header)" \
  "$(printf 'header = "X-Civi-Key: %s"' "$KEY")")"
check "sending the site key as well changes nothing" \
  "$([ "$code" = "200" ] && grep -q '"is_error":0' "$OUT_DIR/v3_with_site_key.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_with_site_key.body")"

# ── revocation reaches the add-in on its next call ───────────────────────
revoke_token "$TOKEN_ID"
code="$(addin_api3 v3_revoked Contact get "$V3_CONTACTS" "$(auth_header)")"
check "a revoked token is refused" \
  "$(grep -q 'Unknown keyring token' "$OUT_DIR/v3_revoked.body" && echo 0 || echo 1)" \
  "http=$code $(head -c 200 "$OUT_DIR/v3_revoked.body")"

note "token id used: $TOKEN_ID (revoked)"
finish
