#!/usr/bin/env bash
# Dev-stack provisioning: report the authx policy the add-in depends on.
set -euo pipefail

cv ev 'echo "outlook365: authx guards = " . json_encode(Civi::settings()->get("authx_guards")) . "\n";'
cv ev 'echo "outlook365: xheader cred = " . json_encode(Civi::settings()->get("authx_xheader_cred")) . "\n";'
cv ev 'echo "outlook365: site key = " . CIVICRM_SITE_KEY . "\n";'
