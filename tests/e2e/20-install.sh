#!/usr/bin/env bash
#
# The extension installs and its pages render on the stack's CiviCRM, with
# no dependency on a dead API-key extension.
#
# Usage: bash tests/e2e/20-install.sh
set -uo pipefail
. "$(dirname "$0")/lib.sh"

echo "outlook365 e2e — installability"
echo

note "$(cvx ev 'echo "CiviCRM " . CRM_Utils_System::version() . ", PHP " . PHP_VERSION . ", Smarty " . \Smarty\Smarty::SMARTY_VERSION;')"

cvx api4 Extension.get '{"where":[["key","IN",["outlook365","dataprocessor","keyring","authx"]]],"select":["key","status"],"checkPermissions":false}' \
  > "$OUT_DIR/extensions.json"
for ext in outlook365 dataprocessor keyring authx; do
  check "$ext is installed" \
    "$(python3 -c "
import json,sys
rows = json.load(open('$OUT_DIR/extensions.json'))
sys.exit(0 if any(r['key'] == '$ext' and r['status'] == 'installed' for r in rows) else 1)
" && echo 0 || echo 1)"
done

check "info.xml no longer requires com.cividesk.apikey" \
  "$(grep -q 'com.cividesk.apikey' "$(dirname "$0")/../../info.xml" && echo 1 || echo 0)"

# The permission the anonymous task pane needs; seed.php grants it.
seed_subject >/dev/null

for page in taskpane.html readtaskpane.html; do
  code="$(hit "page_$page" "$BASE_URL/civicrm/outlook365/$page")"
  check "civicrm/outlook365/$page renders" \
    "$([ "$code" = "200" ] && grep -q 'assets/civicrm.js' "$OUT_DIR/page_$page.body" && echo 0 || echo 1)" \
    "http=$code $(head -c 200 "$OUT_DIR/page_$page.body")"
  check "civicrm/outlook365/$page carries no unrendered Smarty" \
    "$(grep -qE '\{(ts|literal|crmScope|\$)' "$OUT_DIR/page_$page.body" && echo 1 || echo 0)" \
    "$(grep -oE '\{(ts|literal|crmScope|\$)[^}]*\}' "$OUT_DIR/page_$page.body" | head -3)"
done

# The data processors the add-in ships back its Outlook365* APIv3 entities.
for dp in outlook_365_contacts outlook_365_groups; do
  check "data processor $dp is installed and active" \
    "$([ "$(sql "SELECT COUNT(*) FROM civicrm_data_processor WHERE name = '$dp' AND is_active = 1")" = "1" ] && echo 0 || echo 1)"
done

finish
