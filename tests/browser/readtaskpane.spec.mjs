// The read task pane: the mail's contacts, saving contacts and the mail.
import { test, expect } from "@playwright/test";
import * as h from "./harness.mjs";

const SCOPE = [
  "access CiviCRM",
  "access AJAX API",
  "view all contacts",
  "edit all contacts",
  "add contacts",
  "view all activities",
];
const run = `pw${Date.now().toString(36)}`;
const known = { displayName: `Kim Known ${run}`, emailAddress: `kim.${run}@example.org` };
const newcomer = (n) => ({
  displayName: `Nia New${n} ${run}`,
  emailAddress: `nia${n}.${run}@example.org`,
});
const SAVED = "Saved in CiviCRM";
const MARKUP = '<img src=x onerror="window.__xss=1">';
// A scope that can read and edit contacts but not create them.
const NO_ADD = SCOPE.filter((p) => p !== "add contacts");

let subject;
let key;
let minted;
let knownId;

const message = (n, extra = {}) => ({
  mode: "read",
  from: known,
  to: [newcomer(n)],
  cc: [],
  bcc: [],
  subject: `Example subject ${run} ${n}`,
  body: "<p>Example body</p>",
  dateTimeCreated: "2026-01-02T03:04:05Z",
  ...extra,
});
const open = (page, item) =>
  h.openPane(page, "readtaskpane.html", {
    item,
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: minted.token },
  });
const rows = (page) => page.locator("#contacts .ms-List > .ms-ListItem");
const contactsWithEmail = (email) =>
  h
    .sql(`SELECT contact_id FROM civicrm_email WHERE email = '${email}' ORDER BY contact_id`)
    .split("\n")
    .filter(Boolean)
    .map(Number);
// Source (record type 2) and targets (3) of the newest activity with this subject.
function activity(subjectLine) {
  const id = h.sql(
    `SELECT id FROM civicrm_activity WHERE subject = '${subjectLine}' ORDER BY id DESC LIMIT 1`,
  );
  const of = (type) =>
    h
      .sql(
        `SELECT contact_id FROM civicrm_activity_contact WHERE activity_id = ${Number(id)} AND record_type_id = ${type} ORDER BY contact_id`,
      )
      .split("\n")
      .filter(Boolean)
      .map(Number);
  return { id: Number(id), source: of(2), targets: of(3) };
}

// The pane's state once it is idle. jQuery runs deferred callbacks from zero-delay
// timers, so the state is read after timer turns until one turn changes nothing.
const paneState = (page) =>
  page
    .evaluate(async () => {
      const $ = window.jQuery;
      const read = () =>
        JSON.stringify({
          errorDialog: $(".ms-Dialog").is(":visible"),
          savedNotice: $("#saved-contact-help-text").is(":visible"),
          inFlight: $.active,
          sameDocument: window.__paneMarker === true,
        });
      let previous;
      let current = read();
      do {
        previous = current;
        await new Promise((resolve) => {
          setTimeout(resolve, 0);
        });
        current = read();
      } while (current !== previous);
      return JSON.parse(current);
    })
    // A reload destroys the context; the next poll reads the new document.
    .catch(() => null);
// A reload drops the marker.
const markDocument = (page) =>
  page.evaluate(() => {
    window.__paneMarker = true;
  });
const FAILED = { errorDialog: true, savedNotice: false, inFlight: 0, sameDocument: true };

test.beforeAll(() => {
  subject = h.seedSubject();
  key = h.siteKey();
  [{ id: knownId }] = h.cv("Contact.create", {
    values: { contact_type: "Individual", first_name: "Kim", last_name: `Known ${run}` },
  });
  h.cv("Email.create", {
    values: { contact_id: knownId, email: known.emailAddress, is_primary: true },
  });
  minted = h.mintToken(subject, SCOPE);
});

test.afterAll(() => h.revokeToken(minted.id));

test("shows which of the mail's addresses CiviCRM knows", async ({ page }) => {
  const log = await open(page, message(1));
  await expect(rows(page)).toHaveCount(2);
  await expect(rows(page).nth(0)).toHaveClass(/not_to_save/);
  await expect(rows(page).nth(0).locator(".ms-ListItem-secondaryText")).toHaveText(
    known.displayName,
  );
  await expect(rows(page).nth(0).locator("a")).toHaveAttribute(
    "href",
    new RegExp(`cid=${knownId}$`),
  );
  await expect(rows(page).nth(1)).toHaveClass(/to_save/);
  await expect(rows(page).nth(1).locator(".save-contact")).toBeVisible();
  h.assertNetwork(log, minted.token, key);
});

test("saves an unknown address as a contact", async ({ page }) => {
  const log = await open(page, message(2));
  await rows(page).nth(1).locator(".save-contact").click();
  await expect(page.locator("#saved-contact-help-text")).toBeVisible();
  await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);
  expect(contactsWithEmail(newcomer(2).emailAddress)).toHaveLength(1);
  h.assertNetwork(log, minted.token, key);
});

test("saves the mail as an activity of the token's contact and tags it", async ({ page }) => {
  const item = message(3);
  const log = await open(page, item);
  await expect(rows(page)).toHaveCount(2);
  await page.locator(".save-email").click();
  await expect.poll(async () => (await h.officeState(page)).categories).toEqual([SAVED]);

  const [created] = contactsWithEmail(newcomer(3).emailAddress);
  const saved = activity(item.subject);
  expect(saved.source).toEqual([subject]);
  expect(saved.targets).toEqual([knownId, created].sort((a, b) => a - b));
  expect((await h.officeState(page)).masterCategories).toEqual([SAVED]);
  h.assertNetwork(log, minted.token, key);
});

// Picks "new group" in the dialog; resolves once the pane has handled its answer.
async function addToNewGroup(page, groupTitle) {
  await expect(rows(page)).toHaveCount(2);
  await page.locator(".save-contact-all-group").click();
  const dialog = page.frameLocator("#office-dialog");
  await expect(
    dialog.locator(".group-select-dropdown option", { hasText: "Administrators" }),
  ).toBeAttached();
  await dialog.locator("#civicrm-group-field").fill(groupTitle);
  await dialog.locator("#group-select-done").click();
  await expect(page.locator("#office-dialog")).toHaveCount(0);
}
const groupMembers = (groupTitle) =>
  h
    .sql(`SELECT gc.contact_id FROM civicrm_group_contact gc JOIN civicrm_group g ON g.id = gc.group_id
    WHERE g.title = '${groupTitle}' AND gc.status = 'Added' ORDER BY gc.contact_id`)
    .split("\n")
    .filter(Boolean)
    .map(Number);

test("adds the selected contacts to a new group", async ({ page }) => {
  const withGroups = h.mintToken(subject, [...SCOPE, "edit groups"]);
  const groupTitle = `Example group ${run}`;
  const log = await h.openPane(page, "readtaskpane.html", {
    item: message(4),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: withGroups.token },
  });
  await addToNewGroup(page, groupTitle);
  // The pane reloads once everything is saved; the reloaded list knows the newcomer.
  await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);
  await expect(page.locator(".ms-Dialog")).toBeHidden();

  const [created] = contactsWithEmail(newcomer(4).emailAddress);
  expect(groupMembers(groupTitle)).toEqual([knownId, created].sort((a, b) => a - b));
  h.assertNetwork(log, withGroups.token, key);
  h.revokeToken(withGroups.id);
});

test("a scope without edit groups shows an error instead of a silent success", async ({ page }) => {
  const groupTitle = `Example refused group ${run}`;
  const log = await open(page, message(7));
  await addToNewGroup(page, groupTitle);
  await expect(page.locator(".ms-Dialog")).toBeVisible();
  expect(h.sql(`SELECT COUNT(*) FROM civicrm_group WHERE title = '${groupTitle}'`)).toBe("0");
  h.assertNetwork(log, minted.token, key);
});

test.describe("a site with hundreds of groups", () => {
  const COUNT = 400;
  const prefix = `Example bulk group ${run}`;
  const title = (i) =>
    `${prefix} ${String(i).padStart(3, "0")} with a deliberately long title that fills the list well past one line`;
  let groupIds;

  test.beforeAll(() => {
    const records = Array.from({ length: COUNT }, (_, i) => ({ title: title(i) }));
    groupIds = h.cv("Group.save", { records }).map((g) => g.id);
  });
  test.afterAll(() => h.cv("Group.delete", { where: [["title", "LIKE", `${prefix} %`]] }));

  test("the dialog lists every group and adds the contacts to one", async ({ page }) => {
    const withGroups = h.mintToken(subject, [...SCOPE, "edit groups"]);
    const log = await h.openPane(page, "readtaskpane.html", {
      item: message(8),
      roaming: { civicrm_url: h.BASE_URL, civicrm_token: withGroups.token },
    });
    await expect(rows(page)).toHaveCount(2);
    await page.locator(".save-contact-all-group").click();
    const dialog = page.frameLocator("#office-dialog");
    const seeded = dialog.locator(".ms-Dropdown-item", { hasText: prefix });
    await expect(seeded).toHaveCount(COUNT);
    await expect(dialog.locator(".group-select-dropdown option", { hasText: prefix })).toHaveCount(
      COUNT,
    );

    const target = COUNT - 1;
    await dialog.locator(".ms-Dropdown > .ms-Dropdown-title:not(.ms-Dropdown-truncator)").click();
    await dialog.locator(".ms-Dropdown-item", { hasText: title(target) }).click();
    await expect(dialog.locator(".group-select-dropdown")).toHaveValue(String(groupIds[target]));
    // Read before Done: the pane reloads once everything is saved.
    const state = await h.officeState(page);
    expect(state.dialogUrls).toHaveLength(1);
    expect(new URL(state.dialogUrls[0]).search).toBe("");
    // One messageChild, larger than the request line a web server accepts.
    expect(state.childMessages).toHaveLength(1);
    expect(state.childMessages[0]).toBeGreaterThan(8192);
    await dialog.locator("#group-select-done").click();
    await expect(page.locator("#office-dialog")).toHaveCount(0);
    await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);

    const [created] = contactsWithEmail(newcomer(8).emailAddress);
    const members = h
      .sql(`SELECT contact_id FROM civicrm_group_contact
        WHERE group_id = ${Number(groupIds[target])} AND status = 'Added' ORDER BY contact_id`)
      .split("\n")
      .map(Number);
    expect(members).toEqual([knownId, created].sort((a, b) => a - b));
    // No group was created on the way.
    expect(h.sql(`SELECT COUNT(*) FROM civicrm_group WHERE title LIKE '${prefix} %'`)).toBe(
      String(COUNT),
    );
    h.assertNetwork(log, withGroups.token, key);
    h.revokeToken(withGroups.id);
  });
});

test("a client without DialogApi 1.2 shows an error instead of an empty group list", async ({
  page,
}) => {
  const log = await h.openPane(page, "readtaskpane.html", {
    item: message(9),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: minted.token },
    requirementSets: { DialogApi: "1.1" },
  });
  await expect(rows(page)).toHaveCount(2);
  await page.locator(".save-contact-all-group").click();
  await expect(page.locator(".ms-Dialog")).toBeVisible();
  expect((await h.officeState(page)).dialogUrls).toEqual([]);
  h.assertNetwork(log, minted.token, key);
});

test("a revoked token shows the error dialog", async ({ page }) => {
  const revoked = h.mintToken(subject, SCOPE);
  h.revokeToken(revoked.id);
  const log = await h.openPane(page, "readtaskpane.html", {
    item: message(6),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: revoked.token },
  });
  await expect(page.locator(".ms-Dialog")).toBeVisible();
  await expect(page.locator("#civicrm-notice-text")).toContainText(
    "An error occurred communicating with CiviCRM",
  );
  h.assertNetwork(log, revoked.token, key);
});

test("a failed bulk save shows the error, no saved notice and no reload", async ({ page }) => {
  const noAdd = h.mintToken(subject, NO_ADD);
  const log = await h.openPane(page, "readtaskpane.html", {
    item: message(10),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: noAdd.token },
  });
  await expect(rows(page)).toHaveCount(2);
  await markDocument(page);
  await page.locator(".save-contact-all").click();
  await expect.poll(() => paneState(page)).toEqual(FAILED);
  expect(contactsWithEmail(newcomer(10).emailAddress)).toEqual([]);
  h.assertNetwork(log, noAdd.token, key);
  h.revokeToken(noAdd.id);
});

test("a failed contact save stops saving the mail: no activity, no tag", async ({ page }) => {
  const noAdd = h.mintToken(subject, NO_ADD);
  const item = message(11);
  const log = await h.openPane(page, "readtaskpane.html", {
    item,
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: noAdd.token },
  });
  await expect(rows(page)).toHaveCount(2);
  await markDocument(page);
  await page.locator(".save-email").click();
  await expect.poll(() => paneState(page)).toEqual(FAILED);
  const activityCalls = log.requests.filter((r) =>
    r.url().includes("/civicrm/ajax/api4/Activity/"),
  );
  expect(activityCalls.map((r) => r.url())).toEqual([]);
  expect(h.sql(`SELECT COUNT(*) FROM civicrm_activity WHERE subject = '${item.subject}'`)).toBe(
    "0",
  );
  expect((await h.officeState(page)).categories).toEqual([]);
  h.assertNetwork(log, noAdd.token, key);
  h.revokeToken(noAdd.id);
});

test("names from the mail, from Outlook and from CiviCRM render as text", async ({ page }) => {
  const sender = {
    displayName: `${MARKUP} Sender ${run}`,
    emailAddress: `sender.${run}@example.org`,
  };
  const marked = { displayName: "Marked", emailAddress: `marked.${run}@example.org` };
  const [{ id: markedId }] = h.cv("Contact.create", {
    values: { contact_type: "Individual", first_name: MARKUP, last_name: `Marked ${run}` },
  });
  h.cv("Email.create", { values: { contact_id: markedId, email: marked.emailAddress } });
  const log = await open(page, message(12, { from: sender, to: [marked] }));
  await expect(rows(page)).toHaveCount(2);
  await expect(rows(page).nth(0).locator(".ms-ListItem-secondaryText")).toHaveText(
    sender.displayName,
  );
  await expect(rows(page).nth(0).locator(".ms-ListItem-action")).toHaveAttribute(
    "data-civicrm-name",
    sender.displayName,
  );
  await expect(rows(page).nth(1).locator(".ms-ListItem-secondaryText")).toHaveText(
    `${MARKUP} Marked ${run}`,
  );
  await expect(page.locator("#contacts img")).toHaveCount(0);
  expect(await page.evaluate(() => window.__xss)).toBeUndefined();
  h.assertNetwork(log, minted.token, key);
});

test("the save-to-group dialog shows group titles as text", async ({ page }) => {
  const title = `${MARKUP} Group ${run}`;
  const [{ id: groupId }] = h.cv("Group.create", { values: { title } });
  const log = await open(page, message(13));
  await expect(rows(page)).toHaveCount(2);
  await page.locator(".save-contact-all-group").click();
  const dialog = page.frameLocator("#office-dialog");
  const item = dialog.locator(".ms-Dropdown-item", { hasText: title });
  await expect(item).toHaveCount(1);
  await expect(item).toHaveText(title);
  // Picking an item must add no element: markup parsed even briefly would be.
  const frame = page.frame({ url: /saveContactInGroup/ });
  await frame.evaluate(() => {
    window.__added = [];
    new MutationObserver((records) =>
      records.forEach((r) =>
        r.addedNodes.forEach((n) => n.nodeType === 1 && window.__added.push(n.tagName)),
      ),
    ).observe(document.body, { childList: true, subtree: true });
  });
  await dialog.locator(".ms-Dropdown > .ms-Dropdown-title:not(.ms-Dropdown-truncator)").click();
  await item.click();
  await expect(
    dialog.locator(".ms-Dropdown > .ms-Dropdown-title:not(.ms-Dropdown-truncator)"),
  ).toHaveText(title);
  await expect(dialog.locator(".group-select-dropdown")).toHaveValue(String(groupId));
  await expect(dialog.locator("img")).toHaveCount(0);
  expect(await frame.evaluate(() => window.__added)).toEqual([]);
  expect(await frame.evaluate(() => window.__xss)).toBeUndefined();
  h.assertNetwork(log, minted.token, key);
  h.cv("Group.delete", { where: [["id", "=", groupId]] });
});

test("the command bar exposes its label and a settings button", async ({ page }) => {
  const log = await open(page, message(14));
  await expect(rows(page)).toHaveCount(2);
  await expect(page.getByRole("heading", { name: "Contacts", exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Settings", exact: true })).toBeVisible();
  await expect(page.locator(".pivot-tab i[aria-hidden='true']")).toHaveCount(2);
  h.assertNetwork(log, minted.token, key);
});

test("the rows, the save buttons and settings answer the keyboard alone", async ({ page }) => {
  const log = await open(page, message(15));
  await expect(rows(page)).toHaveCount(2);
  const sender = page.getByRole("checkbox", { name: known.displayName, exact: true });
  await h.tabTo(page, sender);
  expect(await h.focusRing(page)).toBe(true);
  await expect(sender).toBeChecked();
  await page.keyboard.press(" ");
  await expect(sender).not.toBeChecked();
  await expect(rows(page).nth(0)).not.toHaveClass(/is-selected/);

  await page.getByRole("button", { name: "Unselect All", exact: true }).press("Enter");
  await expect(page.locator("#contacts .ms-ListItem.is-selected")).toHaveCount(0);
  await page.getByRole("button", { name: "Select All", exact: true }).press("Enter");
  await expect(page.locator("#contacts .ms-ListItem.is-selected")).toHaveCount(2);

  const save = rows(page).nth(1).getByRole("button", { name: "Save Contact to CiviCRM" });
  await save.press("Enter");
  await expect(page.locator("#saved-contact-help-text")).toBeVisible();
  expect(contactsWithEmail(newcomer(15).emailAddress)).toHaveLength(1);

  await page.getByRole("button", { name: "Settings", exact: true }).press("Enter");
  await expect(page.frameLocator("#office-dialog").locator("#civicrm-url")).toHaveValue(h.BASE_URL);
  h.assertNetwork(log, minted.token, key);
});

// APIv3 calls all share one URL, so the entity and action come from the form body.
const api3Calls = (log, entity, action) =>
  log.requests.filter((r) => {
    const body = new URLSearchParams(r.postData() || "");
    return body.get("entity") === entity && body.get("action") === action;
  });

test("save all contacts creates every selected unknown address and reloads", async ({ page }) => {
  const item = message(16, { to: [newcomer("16a"), newcomer("16b")] });
  const log = await open(page, item);
  await expect(rows(page)).toHaveCount(3);
  await markDocument(page);
  await page.locator(".save-contact-all").click();
  // The reload is the pane's own confirmation that every save came back.
  await expect.poll(async () => (await paneState(page))?.sameDocument).toBe(false);
  await expect(rows(page)).toHaveCount(3);
  await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);
  await expect(rows(page).nth(2)).toHaveClass(/not_to_save/);
  expect(contactsWithEmail(newcomer("16a").emailAddress)).toHaveLength(1);
  expect(contactsWithEmail(newcomer("16b").emailAddress)).toHaveLength(1);
  expect(api3Calls(log, "Contact", "create")).toHaveLength(2);
  h.assertNetwork(log, minted.token, key);
});

test("save all contacts stops at the first refusal instead of saving on", async ({ page }) => {
  const noAdd = h.mintToken(subject, NO_ADD);
  const item = message(17, { to: [newcomer("17a"), newcomer("17b")] });
  const log = await h.openPane(page, "readtaskpane.html", {
    item,
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: noAdd.token },
  });
  await expect(rows(page)).toHaveCount(3);
  await markDocument(page);
  await page.locator(".save-contact-all").click();
  await expect.poll(() => paneState(page)).toEqual(FAILED);
  // Only the first address was tried; the second never left the pane.
  expect(api3Calls(log, "Contact", "create")).toHaveLength(1);
  expect(contactsWithEmail(newcomer("17a").emailAddress)).toEqual([]);
  expect(contactsWithEmail(newcomer("17b").emailAddress)).toEqual([]);
  h.assertNetwork(log, noAdd.token, key);
  h.revokeToken(noAdd.id);
});
