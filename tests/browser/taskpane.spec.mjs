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

test("the tabs expose their labels and keep the selected state in sync", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  const contacts = page.getByRole("tab", { name: "Contacts", exact: true });
  const groups = page.getByRole("tab", { name: "Groups", exact: true });
  await expect(page.getByRole("tablist")).toHaveCount(1);
  await expect(contacts).toHaveAttribute("aria-controls", "contacts");
  await expect(groups).toHaveAttribute("aria-controls", "groups");
  await expect(contacts).toHaveAttribute("aria-selected", "true");
  await expect(groups).toHaveAttribute("aria-selected", "false");
  await expect(page.getByRole("button", { name: "Settings", exact: true })).toBeVisible();
  // The label is the tab's whole accessible name; the glyph stays out of it.
  await expect(page.locator(".pivot-tab i[aria-hidden='true']")).toHaveCount(3);

  await groups.click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await expect(groups).toHaveAttribute("aria-selected", "true");
  await expect(contacts).toHaveAttribute("aria-selected", "false");
  await expect(groups.locator("xpath=..")).toHaveClass(/is-active/);
  await expect(page.locator(".ms-CommandButton--pivot.is-active")).toHaveCount(1);
  await expect(page.locator("#groups")).toHaveAttribute("role", "tabpanel");
  await expect(page.locator("#civicrm-error")).toHaveAttribute("role", "alert");
  h.assertNetwork(log, minted.token, key);
});

test("a group is expanded, selected and addressed from the keyboard alone", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  await page.getByRole("tab", { name: "Groups", exact: true }).press("Enter");
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");

  const toggle = page.getByRole("button", { name: groupTitle, exact: true });
  await expect(toggle).toHaveAttribute("aria-expanded", "false");
  await h.tabTo(page, toggle);
  expect(await h.focusRing(page)).toBe(true);
  await page.keyboard.press("Enter");
  await expect(toggle).toHaveAttribute("aria-expanded", "true");

  const member = page.getByRole("checkbox", { name: person.name, exact: true });
  await expect(member).not.toBeChecked();
  await member.press(" ");
  await expect(member).toBeChecked();
  await expect(member.locator("xpath=..")).toHaveClass(/is-selected/);
  await page.getByRole("button", { name: "Unselect All", exact: true }).press("Enter");
  await expect(member).not.toBeChecked();
  await page.getByRole("button", { name: "Select All", exact: true }).press("Enter");
  await expect(member).toBeChecked();

  await page.getByRole("button", { name: "Cc", exact: true }).press("Enter");
  await expect
    .poll(async () => (await h.officeState(page)).recipients.cc)
    .toEqual([{ displayName: person.name, emailAddress: person.email }]);
  await toggle.press("Enter");
  await expect(toggle).toHaveAttribute("aria-expanded", "false");
  await expect(page.locator("#groups .ms-ListItem")).toHaveCount(0);
  h.assertNetwork(log, minted.token, key);
});
