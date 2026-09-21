// axe-core over both task panes. The add-in renders in one theme only —
// no stylesheet carries a prefers-color-scheme rule — so light mode is all of it.
import AxeBuilder from "@axe-core/playwright";
import { test, expect } from "@playwright/test";
import * as h from "./harness.mjs";

const SCOPE = [
  "access CiviCRM",
  "access AJAX API",
  "view all contacts",
  "edit all contacts",
  "add contacts",
];
const run = `ax${Date.now().toString(36)}`;
const person = { displayName: `Ada Axe ${run}`, emailAddress: `ada.${run}@example.org` };
const groupTitle = `Example axe group ${run}`;

let minted;
let groupId;

const scan = async (page) => new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
// The rule ids and the elements they fired on, so a failure names the defect.
const violations = (results) =>
  results.violations.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(" ")).join(", ")}`);

test.beforeAll(() => {
  const subject = h.seedSubject();
  const [contact] = h.cv("Contact.create", {
    values: { contact_type: "Individual", first_name: "Ada", last_name: `Axe ${run}` },
  });
  h.cv("Email.create", {
    values: { contact_id: contact.id, email: person.emailAddress, is_primary: true },
  });
  [{ id: groupId }] = h.cv("Group.create", { values: { title: groupTitle } });
  h.cv("GroupContact.create", { values: { group_id: groupId, contact_id: contact.id } });
  minted = h.mintToken(subject, SCOPE);
});

test.afterAll(() => {
  h.cv("Group.delete", { where: [["id", "=", groupId]] });
  h.revokeToken(minted.id);
});

const roaming = () => ({ civicrm_url: h.BASE_URL, civicrm_token: minted.token });

test("the compose pane has no accessibility violations", async ({ page }) => {
  await h.openPane(page, "taskpane.html", { item: { mode: "compose" }, roaming: roaming() });
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  expect(violations(await scan(page))).toEqual([]);

  await page.getByRole("tab", { name: "Groups" }).click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");
  await expect(page.locator("#groups .ms-Persona")).toHaveCount(1);
  // A search box in use shows its clear button, which Fabric ships unlabelled.
  await expect(page.locator(".ms-SearchBox-clear .ms-CommandButton-button").first()).toBeVisible();
  expect(violations(await scan(page))).toEqual([]);

  await page.getByRole("button", { name: groupTitle }).click();
  await expect(page.locator("#groups .ms-ListItem")).toHaveCount(1);
  expect(violations(await scan(page))).toEqual([]);
});

test("the read pane has no accessibility violations", async ({ page }) => {
  await h.openPane(page, "readtaskpane.html", {
    item: {
      mode: "read",
      from: person,
      to: [{ displayName: `Nemo ${run}`, emailAddress: `nemo.${run}@example.org` }],
      cc: [],
      bcc: [],
      subject: `Example axe subject ${run}`,
      body: "<p>Example body</p>",
      dateTimeCreated: "2026-01-02T03:04:05Z",
    },
    roaming: roaming(),
  });
  await expect(page.locator("#contacts .ms-ListItem")).toHaveCount(2);
  expect(violations(await scan(page))).toEqual([]);
});
