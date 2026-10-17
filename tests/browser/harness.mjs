// Stack access through tests/e2e/lib.sh, and a page wired to the office stub
// with every request logged and nothing leaving for an external host.

import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";
import { expect } from "@playwright/test";
import { officeStub } from "./office-stub.mjs";

const LIB = path.join(import.meta.dirname, "..", "e2e", "lib.sh");
const FABRIC = path.dirname(
  createRequire(import.meta.url).resolve("office-ui-fabric-js/package.json"),
);
export const BASE_URL = new URL(process.env.CIVICRM_BASE_URL || "http://localhost:8200").origin;

// civikitchen's playwright job names its stack's compose file in CK_COMPOSE_FILE.
const ENV = process.env.CK_COMPOSE_FILE
  ? { ...process.env, COMPOSE_FILE: path.resolve(process.env.CK_COMPOSE_FILE) }
  : process.env;

// Arguments reach lib.sh as positional parameters, never as shell text.
function lib(code, ...args) {
  return execFileSync("bash", ["-c", `set -euo pipefail; . "$0"; ${code}`, LIB, ...args], {
    env: ENV,
    encoding: "utf8",
    stdio: ["ignore", "pipe", "inherit"],
  }).trim();
}

export const sql = (query) => lib('sql "$1"', query);
export const cv = (entityAction, params) =>
  JSON.parse(
    lib('cv api4 "$1" "$2"', entityAction, JSON.stringify({ ...params, checkPermissions: false })),
  );

export function seedSubject() {
  const id = lib("seed_subject");
  if (!/^\d+$/.test(id)) {
    throw new Error(`seed.php returned no contact id: ${id}`);
  }
  return Number(id);
}

export function mintToken(contactId, permissions) {
  const [id, token] = lib('mint_token "$@"', String(contactId), ...permissions).split(" ");
  if (!/^\d+$/.test(id) || !token?.startsWith("kr_")) {
    throw new Error("mint_token returned no token");
  }
  return { id: Number(id), token };
}

export const revokeToken = (id) => lib('revoke_token "$1"', String(id));
export const siteKey = () => lib("site_key");

/**
 * Opens a task pane page with office.js replaced by the stub. Returns the
 * request log; assertNetwork() checks it against the credential rules.
 */
export async function openPane(page, route, fixture) {
  const requests = [];
  const external = [];
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  page.on("request", (request) => requests.push(request));

  await page.addInitScript(officeStub, {
    language: "en-US",
    roaming: {},
    categories: [],
    masterCategories: [],
    requirementSets: { DialogApi: "1.2" },
    ...fixture,
  });
  await page.context().route("**/*", (r) => {
    const url = new URL(r.request().url());
    if (url.origin === BASE_URL) {
      return r.fallback();
    }
    if (url.pathname.endsWith("/hosted/office.js")) {
      return r.fulfill({ contentType: "text/javascript", body: "" });
    }
    const fabric = url.pathname.match(/\/office-ui-fabric-js\/1\.4\.0\/((?:css|js)\/[\w.]+)$/);
    if (fabric) {
      return r.fulfill({ path: path.join(FABRIC, "dist", fabric[1]) });
    }
    // Fabric's CSS pulls its web fonts from a CDN; the layout does not depend on them.
    if (!url.href.startsWith("https://spoprod-a.akamaihd.net/files/fabric/assets/")) {
      external.push(url.href);
    }
    return r.abort();
  });

  await page.goto(`${BASE_URL}/civicrm/outlook365/${route}`);
  return { requests, external, errors };
}

/**
 * Every CiviCRM API call carries the token in X-Civi-Auth; no request carries
 * it anywhere else, and none sends the site key.
 */
export function assertNetwork(log, token, key) {
  const api = log.requests.filter((r) => r.url().startsWith(`${BASE_URL}/civicrm/ajax/`));
  expect(api.length, "the pane made CiviCRM API calls").toBeGreaterThan(0);
  for (const r of api) {
    expect(r.headers()["x-civi-auth"], `${r.method()} ${r.url()}`).toBe(`Bearer ${token}`);
  }
  for (const r of log.requests) {
    const where = `${r.method()} ${r.url()}`;
    const body = r.postData() || "";
    expect(r.url().includes(token), `token in the URL of ${where}`).toBe(false);
    expect(body.includes(token), `token in the body of ${where}`).toBe(false);
    expect(/[?&](api_key|key)=/.test(r.url()), `credential parameter in ${where}`).toBe(false);
    expect(r.url().includes(key) || body.includes(key), `site key in ${where}`).toBe(false);
    expect(r.headers()["x-civi-key"], `X-Civi-Key on ${where}`).toBeUndefined();
  }
  expect(log.external, "requests to external hosts").toEqual([]);
  expect(log.errors, "uncaught page errors").toEqual([]);
}

// The serialisable part of the stub's state: what the add-in stored or sent to Outlook.
export const officeState = (page) =>
  page.evaluate(() => {
    const { roaming, saved, dialogUrls, childMessages, recipients, categories, masterCategories } =
      window.__officeState;
    return { roaming, saved, dialogUrls, childMessages, recipients, categories, masterCategories };
  });

/**
 * Presses Tab until the locator holds the focus, and fails when it never does.
 * Focus reached this way is keyboard focus, so :focus-visible applies.
 */
export async function tabTo(page, locator, max = 40) {
  for (let i = 0; i < max; i += 1) {
    await page.keyboard.press("Tab");
    if (await locator.evaluate((el) => el === document.activeElement)) {
      return i + 1;
    }
  }
  throw new Error("not reachable with Tab");
}

// Whether the focused element draws a focus indicator.
export const focusRing = (page) =>
  page.evaluate(() => {
    const el = document.activeElement;
    return el.matches(":focus-visible") && getComputedStyle(el).outlineStyle !== "none";
  });
