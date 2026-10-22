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

test("a search sent before the first list arrives shows only its own results", async ({ page }) => {
  let release;
  const held = new Promise((resolve) => {
    release = resolve;
  });
  await page.route(`${h.BASE_URL}/civicrm/ajax/rest`, async (r) => {
    const body = new URLSearchParams(r.request().postData() || "");
    const unfiltered =
      body.get("entity") === "Outlook365Contact" &&
      body.get("action") === "get" &&
      !JSON.parse(body.get("json") || "{}").display_name;
    if (unfiltered) await held;
    return r.fallback();
  });
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });

  await page.locator("#searchField").fill(`Example ${run}`);
  await page.locator("#searchField").press("Enter");
  const found = page.locator("#contacts .ms-Persona");
  await expect(found).toHaveCount(1);
  release();
  // jQuery 3 runs .then callbacks in a timer queued before the request leaves jQuery.active,
  // so one timer after that has seen the stale list's callback run.
  await page.waitForFunction(() => window.jQuery.active === 0);
  await page.evaluate(
    () =>
      new Promise((resolve) => {
        setTimeout(resolve, 0);
      }),
  );
  await expect(found).toHaveCount(1);
  await expect(found.locator(".ms-Persona-primaryText")).toHaveText(person.name);
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

// The API call a request carries: entity, action and the decoded APIv3 params.
const apiCall = (request) => {
  const body = new URLSearchParams(request.postData() || "");
  return {
    entity: body.get("entity"),
    action: body.get("action"),
    params: JSON.parse(body.get("json") || "{}"),
  };
};
// Holds the first API call that `match` accepts until the returned release() is called.
async function holdFirst(page, match) {
  let release;
  const held = new Promise((resolve) => {
    release = resolve;
  });
  let taken = false;
  await page.route(`${h.BASE_URL}/civicrm/ajax/rest`, async (r) => {
    if (!taken && match(apiCall(r.request()))) {
      taken = true;
      await held;
    }
    return r.fallback();
  });
  return release;
}
// jQuery 3 runs each .then step in its own timer after the request leaves jQuery.active;
// a group's members pass three of them before expandGroups resumes, so wait out more.
async function settled(page) {
  await page.waitForFunction(() => window.jQuery.active === 0);
  await page.evaluate(async () => {
    for (let i = 0; i < 10; i++) {
      await new Promise((resolve) => {
        setTimeout(resolve, 0);
      });
    }
  });
}
const groupRow = (page, title) => page.locator("#groups .ms-Persona").filter({ hasText: title });

test("contacts that arrive before getfields still show their secondary fields", async ({
  page,
}) => {
  const release = await holdFirst(
    page,
    (c) => c.entity === "Outlook365Contact" && c.action === "getfields",
  );
  const answered = page.waitForResponse((r) => {
    const c = apiCall(r.request());
    return c.entity === "Outlook365Contact" && c.action === "get";
  });
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await answered;
  release();
  const first = page.locator("#contacts .ms-Persona").first();
  await expect(first).toBeVisible();
  await expect(first.locator(".ms-Persona-secondaryText strong").first()).toHaveText(/:$/);
  h.assertNetwork(log, minted.token, key);
});

test("a getfields failure from before a search does not show an error over its results", async ({
  page,
}) => {
  let release;
  const held = new Promise((resolve) => {
    release = resolve;
  });
  let taken = false;
  await page.route(`${h.BASE_URL}/civicrm/ajax/rest`, async (r) => {
    const c = apiCall(r.request());
    if (!taken && c.entity === "Outlook365Contact" && c.action === "getfields") {
      taken = true;
      await held;
      return r.fulfill({ status: 500, body: "" });
    }
    return r.fallback();
  });
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.locator("#searchField").fill(`Example ${run}`);
  await page.locator("#searchField").press("Enter");
  await expect(page.locator("#contacts .ms-Persona")).toHaveCount(1);
  release();
  await settled(page);
  await expect(page.locator("#civicrm-error")).toBeHidden();
  await expect(page.locator("#contacts .ms-Persona")).toHaveCount(1);
  h.assertNetwork(log, minted.token, key);
});

test("a group expanded, collapsed and expanded again while loading shows its members once", async ({
  page,
}) => {
  const release = await holdFirst(
    page,
    (c) => c.entity === "Outlook365Group" && c.action === "get",
  );
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.getByRole("tab", { name: "Groups", exact: true }).click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");
  const group = groupRow(page, groupTitle);
  await expect(group).toHaveCount(1);

  const toggle = group.locator(".group-toggle");
  await toggle.click();
  await toggle.click();
  await toggle.click();
  await expect(group.locator(".ms-ListItem")).toHaveCount(1);
  release();
  await settled(page);
  await expect(group.locator(".ms-List")).toHaveCount(1);
  await expect(group.locator(".search-in-group")).toHaveCount(1);
  await expect(group.locator(".ms-ListItem")).toHaveCount(1);
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  h.assertNetwork(log, minted.token, key);
});

test("a group's members failing after the list was reset show no error", async ({ page }) => {
  let release;
  const held = new Promise((resolve) => {
    release = resolve;
  });
  let taken = false;
  await page.route(`${h.BASE_URL}/civicrm/ajax/rest`, async (r) => {
    if (!taken && apiCall(r.request()).entity === "Outlook365Group") {
      taken = true;
      await held;
      return r.fulfill({ status: 500, body: "" });
    }
    return r.fallback();
  });
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.getByRole("tab", { name: "Groups", exact: true }).click();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");
  const group = groupRow(page, groupTitle);
  await group.locator(".group-toggle").click();
  await page.getByRole("tab", { name: "Contacts", exact: true }).click();
  await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
  release();
  await settled(page);
  await expect(page.locator("#civicrm-error")).toBeHidden();
  h.assertNetwork(log, minted.token, key);
});

test("a cleared in-group search lists every member, whatever answers last", async ({ page }) => {
  const miss = `nomatch${run}`;
  const release = await holdFirst(
    page,
    (c) =>
      c.entity === "Outlook365Group" && JSON.stringify(c.params.display_name || "").includes(miss),
  );
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.getByRole("tab", { name: "Groups", exact: true }).click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(groupTitle);
  await page.locator("#searchField").press("Enter");
  const group = groupRow(page, groupTitle);
  await group.locator(".group-toggle").click();
  await expect(group.locator(".ms-ListItem")).toHaveCount(1);

  const field = group.locator(".search-in-group input");
  await field.fill(miss);
  await field.press("Enter");
  await field.fill("");
  await field.press("Enter");
  await expect(group.locator(".ms-ListItem")).toHaveCount(1);
  release();
  await settled(page);
  await expect(group.locator(".ms-ListItem")).toHaveText(person.name);
  const last = log.requests
    .map(apiCall)
    .filter((c) => c.entity === "Outlook365Group")
    .at(-1);
  expect(last.params.display_name).toBeUndefined();
  h.assertNetwork(log, minted.token, key);
});

test("Cc on one expanded group adds only that group's selection", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.getByRole("tab", { name: "Groups", exact: true }).click();
  await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
  await page.locator("#searchField").fill(run);
  await page.locator("#searchField").press("Enter");
  const plain = groupRow(page, groupTitle);
  const other = groupRow(page, marked.name);
  await expect(page.locator("#groups .ms-Persona")).toHaveCount(2);
  for (const group of [plain, other]) {
    await group.locator(".group-toggle").click();
    await expect(group.locator(".ms-ListItem")).toHaveCount(1);
    await group.getByRole("button", { name: "Select All", exact: true }).click();
  }
  await plain.getByRole("button", { name: "Cc", exact: true }).click();
  await expect
    .poll(async () => (await h.officeState(page)).recipients.cc)
    .toEqual([{ displayName: person.name, emailAddress: person.email }]);
  await settled(page);
  expect((await h.officeState(page)).recipients.cc).toHaveLength(1);
  h.assertNetwork(log, minted.token, key);
});

// Which pages a filtered list asked the server for, in order.
const pageOffsets = (log, entity) =>
  log.requests
    .map((r) => new URLSearchParams(r.postData() || ""))
    .filter((b) => b.get("entity") === entity && b.get("action") === "get")
    .map((b) => JSON.parse(b.get("json")))
    .filter((p) => p.display_name || p.title)
    .map((p) => p.options.offset);
const scrollToEnd = (page) =>
  page.evaluate(() => {
    window.scrollTo(0, document.documentElement.scrollHeight);
  });

test.describe("a list longer than one page", () => {
  // One more than the page size, so the second page ends the list.
  const PAGE = 25;
  const COUNT = PAGE + 1;
  const tag = `Scroll ${run}`;

  test.beforeAll(() => {
    // The data processor behind the contact list joins an email address, so each one needs one.
    const created = h.cv("Contact.save", {
      records: Array.from({ length: COUNT }, (_, i) => ({
        contact_type: "Individual",
        first_name: `C${String(i).padStart(2, "0")}`,
        last_name: tag,
      })),
    });
    h.cv("Email.save", {
      records: created.map((c, i) => ({
        contact_id: c.id,
        email: `scroll${String(i).padStart(2, "0")}.${run}@example.org`,
        is_primary: true,
      })),
    });
    h.cv("Group.save", {
      records: Array.from({ length: COUNT }, (_, i) => ({
        title: `${tag} ${String(i).padStart(2, "0")}`,
      })),
    });
  });
  test.afterAll(() => {
    h.cv("Contact.delete", { where: [["last_name", "=", tag]] });
    h.cv("Group.delete", { where: [["title", "LIKE", `${tag} %`]] });
  });

  test("scrolling to the end loads the next page of contacts exactly once", async ({ page }) => {
    const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
    await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
    await page.locator("#searchField").fill(tag);
    await page.locator("#searchField").press("Enter");
    await expect(page.locator("#contacts .ms-Persona")).toHaveCount(PAGE);
    expect(pageOffsets(log, "Outlook365Contact")).toEqual([0]);

    await scrollToEnd(page);
    await expect(page.locator("#contacts .ms-Persona")).toHaveCount(COUNT);
    expect(pageOffsets(log, "Outlook365Contact")).toEqual([0, PAGE]);

    // The list is exhausted: further scrolling asks for nothing more.
    await scrollToEnd(page);
    await scrollToEnd(page);
    await page.waitForLoadState("networkidle");
    expect(pageOffsets(log, "Outlook365Contact")).toEqual([0, PAGE]);
    await expect(page.locator("#contacts .ms-Persona")).toHaveCount(COUNT);
    h.assertNetwork(log, minted.token, key);
  });

  test("scrolling again while the next page loads asks for it once", async ({ page }) => {
    const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
    await expect(page.locator("#contacts .ms-Persona").first()).toBeVisible();
    await page.locator("#searchField").fill(tag);
    await page.locator("#searchField").press("Enter");
    await expect(page.locator("#contacts .ms-Persona")).toHaveCount(PAGE);

    const release = await holdFirst(
      page,
      (c) =>
        c.entity === "Outlook365Contact" && c.action === "get" && c.params.options.offset === PAGE,
    );
    await scrollToEnd(page);
    await expect.poll(() => pageOffsets(log, "Outlook365Contact")).toEqual([0, PAGE]);
    await page.evaluate(() => {
      window.dispatchEvent(new Event("scroll"));
      window.dispatchEvent(new Event("scroll"));
    });
    release();
    await settled(page);
    expect(pageOffsets(log, "Outlook365Contact")).toEqual([0, PAGE]);
    await expect(page.locator("#contacts .ms-Persona")).toHaveCount(COUNT);
    h.assertNetwork(log, minted.token, key);
  });

  test("scrolling to the end loads the next page of groups exactly once", async ({ page }) => {
    const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
    await page.getByRole("tab", { name: "Groups", exact: true }).click();
    await expect(page.locator("#groups .ms-Persona").first()).toBeVisible();
    await page.locator("#searchField").fill(tag);
    await page.locator("#searchField").press("Enter");
    await expect(page.locator("#groups .ms-Persona")).toHaveCount(PAGE);
    expect(pageOffsets(log, "Group")).toEqual([0]);

    await scrollToEnd(page);
    await expect(page.locator("#groups .ms-Persona")).toHaveCount(COUNT);
    expect(pageOffsets(log, "Group")).toEqual([0, PAGE]);

    await scrollToEnd(page);
    await scrollToEnd(page);
    await page.waitForLoadState("networkidle");
    expect(pageOffsets(log, "Group")).toEqual([0, PAGE]);
    await expect(page.locator("#groups .ms-Persona")).toHaveCount(COUNT);
    h.assertNetwork(log, minted.token, key);
  });
});
