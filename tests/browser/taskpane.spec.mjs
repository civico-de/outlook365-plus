// The compose task pane: contact and group lists, recipients, settings.
import { test, expect } from "@playwright/test";
import * as h from "./harness.mjs";

// The scope tests/e2e/10-header-auth.sh mints.
const SCOPE = [
  "access CiviCRM",
  "access AJAX API",
  "view all contacts",
  "edit all contacts",
  "add contacts",
  "view all activities",
  "edit all activities",
];
const run = `pw${Date.now().toString(36)}`;
const person = { name: `Alex Example ${run}`, email: `alex.${run}@example.org` };
const groupTitle = `Example group ${run}`;
const compose = { mode: "compose" };
const MARKUP = '<img src=x onerror="window.__xss=1">';
const marked = { name: `${MARKUP} Marked ${run}`, email: `marked.${run}@example.org` };

let subject;
let key;
let minted;
const configured = () => ({ civicrm_url: h.BASE_URL, civicrm_token: minted.token });

test.beforeAll(() => {
  subject = h.seedSubject();
  key = h.siteKey();
  const [contact] = h.cv("Contact.create", {
    values: { contact_type: "Individual", first_name: "Alex", last_name: `Example ${run}` },
  });
  h.cv("Email.create", {
    values: { contact_id: contact.id, email: person.email, is_primary: true },
  });
  const [group] = h.cv("Group.create", { values: { title: groupTitle } });
  h.cv("GroupContact.create", { values: { group_id: group.id, contact_id: contact.id } });
  const [markedContact] = h.cv("Contact.create", {
    values: { contact_type: "Individual", first_name: MARKUP, last_name: `Marked ${run}` },
  });
  h.cv("Email.create", {
    values: { contact_id: markedContact.id, email: marked.email, is_primary: true },
  });
  const [markedGroup] = h.cv("Group.create", { values: { title: marked.name } });
  h.cv("GroupContact.create", {
    values: { group_id: markedGroup.id, contact_id: markedContact.id },
  });
  minted = h.mintToken(subject, SCOPE);
});

test.afterAll(() => h.revokeToken(minted.id));

test("lists contacts, finds one by name and adds it as a recipient", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();

  await page.locator("#searchField").fill(`Example ${run}`);
  await page.locator("#searchField").press("Enter");
  const found = page.locator("#contacts .ms-Persona");
  await expect(found).toHaveCount(1);
  await expect(found.locator(".ms-Persona-primaryText")).toHaveText(person.name);

  await found.getByRole("button", { name: "To", exact: true }).click();
  await expect
    .poll(async () => (await h.officeState(page)).recipients.to)
    .toEqual([{ displayName: person.name, emailAddress: person.email }]);
  h.assertNetwork(log, minted.token, key);
});

test("expands a group and adds its selected members as Cc", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.locator(".ms-CommandButton--pivot .groups").click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");

  const group = page.locator("#groups .ms-Persona");
  await expect(group).toHaveCount(1);
  await group.locator(".group-name").click();
  const members = group.locator(".ms-ListItem");
  await expect(members).toHaveCount(1);
  await group.getByRole("button", { name: "Select All", exact: true }).click();
  await group.getByRole("button", { name: "Cc", exact: true }).click();
  await expect
    .poll(async () => (await h.officeState(page)).recipients.cc)
    .toEqual([{ displayName: person.name, emailAddress: person.email }]);
  h.assertNetwork(log, minted.token, key);
});

test("an unconfigured pane opens settings and stores URL and token", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose });
  const dialog = page.frameLocator("#office-dialog");
  await expect(dialog.locator("#civicrm-url")).toHaveValue(h.BASE_URL);
  await expect(dialog.locator("#settings-done")).toBeDisabled();
  await dialog.locator("#civicrm-token").fill(minted.token);
  await dialog.locator("#civicrm-token").dispatchEvent("change");
  await dialog.locator("#settings-done").click();

  await expect(page.locator("#office-dialog")).toHaveCount(0);
  const state = await h.officeState(page);
  expect(state.saved).toEqual({ civicrm_url: h.BASE_URL, civicrm_token: minted.token });
  expect(state.dialogUrls.join(" ")).not.toContain(minted.token);
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  h.assertNetwork(log, minted.token, key);
});

test("a revoked token shows an error instead of an empty list", async ({ page }) => {
  const revoked = h.mintToken(subject, SCOPE);
  h.revokeToken(revoked.id);
  const log = await h.openPane(page, "taskpane.html", {
    item: compose,
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: revoked.token },
  });
  await expect(page.locator("#civicrm-error")).toBeVisible();
  await expect(page.locator("#contacts .ms-Persona")).toHaveCount(0);
  h.assertNetwork(log, revoked.token, key);
});

test("contact names and group titles from CiviCRM render as text", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(`Marked ${run}`);
  await page.locator("#searchField").press("Enter");
  const found = page.locator("#contacts .ms-Persona");
  await expect(found).toHaveCount(1);
  await expect(found.locator(".ms-Persona-primaryText")).toHaveText(marked.name);
  await found.getByRole("button", { name: "To", exact: true }).click();
  await expect
    .poll(async () => (await h.officeState(page)).recipients.to)
    .toEqual([{ displayName: marked.name, emailAddress: marked.email }]);

  await page.locator(".ms-CommandButton--pivot .groups").click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(`Marked ${run}`);
  await page.locator("#searchField").press("Enter");
  const group = page.locator("#groups .ms-Persona");
  await expect(group).toHaveCount(1);
  await expect(group.locator(".group-name")).toHaveText(marked.name);
  await group.locator(".group-name").click();
  const members = group.locator(".ms-ListItem");
  await expect(members).toHaveCount(1);
  await expect(members).toHaveText(marked.name);
  await group.getByRole("button", { name: "Select All", exact: true }).click();
  await group.getByRole("button", { name: "Cc", exact: true }).click();
  await expect
    .poll(async () => (await h.officeState(page)).recipients.cc)
    .toEqual([{ displayName: marked.name, emailAddress: marked.email }]);
  await expect(page.locator("#contacts img, #groups img")).toHaveCount(0);
  expect(await page.evaluate(() => window.__xss)).toBeUndefined();
  h.assertNetwork(log, minted.token, key);
});
