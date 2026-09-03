#!/usr/bin/env bash
#
# The localized UI strings describe the fields the settings dialog actually
# has: a site base URL and a keyring token, no site key and no REST endpoint.
# Static — needs node on the host, not the stack.
#
# Usage: bash tests/e2e/40-uistrings.sh
set -uo pipefail
. "$(dirname "$0")/lib.sh"

echo "outlook365 e2e — localized UI strings"
echo

if ! command -v node >/dev/null 2>&1; then
  echo "  ✘ node is required for this suite and was not found on PATH" >&2
  exit 1
fi

results="$OUT_DIR/uistrings.txt"
node - "$(dirname "$0")/../../assets/UIStrings.js" > "$results" <<'NODE'
const fs = require('fs');
const vm = require('vm');
const src = fs.readFileSync(process.argv[2], 'utf8');

// The data is plain objects of strings, which is all of $.extend(true, …)
// that getLocaleStrings exercises.
function extend(deep, target) {
  for (const source of Array.prototype.slice.call(arguments, 2)) {
    for (const key of Object.keys(source)) {
      const value = source[key];
      if (deep && value && typeof value === 'object') {
        if (!target[key] || typeof target[key] !== 'object') target[key] = {};
        extend(deep, target[key], value);
      } else {
        target[key] = value;
      }
    }
  }
  return target;
}

// getLocaleStrings merges into UIStrings.default in place, so every locale
// gets its own evaluation of the file.
function load() {
  const ctx = { $: { extend } };
  vm.createContext(ctx);
  vm.runInContext(src, ctx);
  return ctx.UIStrings;
}

const english = load().default.SettingsScreen;
const locales = ['nl-NL', 'es-ES', 'ca-ES', 'fr-CA'];
const stale = [
  [/\bREST\b/i, 'REST endpoint'],
  [/site[ -]?key|clave de sitio|clau de lloc/i, 'site key'],
  [/api[ -]?key/i, 'API key'],
];
const out = [];
const report = (ok, name, detail) =>
  out.push([ok ? 'PASS' : 'FAIL', name, detail || ''].join('|'));

for (const locale of ['en-US'].concat(locales)) {
  const screen = load().getLocaleStrings(locale).SettingsScreen;
  const hits = [];
  for (const [key, value] of Object.entries(screen)) {
    for (const [pattern, label] of stale) {
      if (pattern.test(value)) hits.push(`${key} names the ${label}: ${value}`);
    }
  }
  report(hits.length === 0, `${locale} settings strings name no removed field`, hits.join('; '));
}

for (const locale of locales) {
  const screen = load().getLocaleStrings(locale).SettingsScreen;
  const missing = ['NotConfigured', 'URL', 'Token', 'TokenKeep']
    .filter((key) => screen[key] !== english[key]);
  report(missing.length === 0, `${locale} falls back to the English field labels`, missing.join(', '));
}
console.log(out.join('\n'));
NODE
node_rc=$?

if [ "$node_rc" != "0" ]; then
  echo "  ✘ resolving assets/UIStrings.js failed" >&2
  cat "$results" >&2
  exit 1
fi

while IFS='|' read -r status name detail; do
  [ -n "$status" ] || continue
  check "$name" "$([ "$status" = "PASS" ] && echo 0 || echo 1)" "$detail"
done < "$results"

finish
