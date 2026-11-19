// Contact ACLs reach the task pane: a token without "view all contacts" lists only the
// contacts its ACL lets it view, in the contact search and in a group's members. A group
// lists a "do not trade" member the contact search leaves out.
import { test, expect } from "@playwright/test";
import * as h from "./harness.mjs";

const SCOPE = ["access CiviCRM", "access AJAX API"];
const run = `acl${Date.now().toString(36)}`;
const permitted = { name: `Pat Allowed ${run}`, email: `pat.${run}@example.org` };
const outside = { name: `Olli Outside ${run}`, email: `olli.${run}@example.org` };
const noTrade = {
  name: `Dana Allowed ${run}`,
  email: `dana.${run}@example.org`,
  do_not_trade: true,
};
const mixedTitle = `Mixed group ${run}`;
const compose = { mode: "compose" };

let subject;
let key;
let minted;
let mixed;
const created = { groups: [], contacts: [], acls: [], roles: [], roleValue: null };
const configured = () => ({ civicrm_url: h.BASE_URL, civicrm_token: minted.token });

function contact(person) {
  const [first, ...last] = person.name.split(" ");
  const [row] = h.cv("Contact.create", {
    values: {
      contact_type: "Individual",
      first_name: first,
      last_name: last.join(" "),
      do_not_trade: !!person.do_not_trade,
    },
  });
  h.cv("Email.create", { values: { contact_id: row.id, email: person.email, is_primary: true } });
  created.contacts.push(row.id);
  return row.id;
}

function group(title, contactIds) {
  const [row] = h.cv("Group.create", { values: { title } });
  created.groups.push(row.id);
  for (const id of contactIds) {
    h.cv("GroupContact.create", { values: { group_id: row.id, contact_id: id } });
  }
  return row.id;
}

test.beforeAll(() => {
  subject = h.seedSubject();
  key = h.siteKey();
  const permittedId = contact(permitted);
  const outsideId = contact(outside);
  const noTradeId = contact(noTrade);
  const allowed = group(`Allowed ${run}`, [permittedId, noTradeId]);
  const subjects = group(`ACL subjects ${run}`, [subject]);
  mixed = group(mixedTitle, [permittedId, outsideId, noTradeId]);

  // Core's ACL: members of `subjects` may view the contacts of `allowed`, nothing else.
  const [role] = h.cv("OptionValue.create", {
    values: { "option_group_id:name": "acl_role", label: `outlook365 e2e ${run}`, name: run },
  });
  created.roleValue = role.value;
  created.roles.push(role.id);
  h.cv("ACLEntityRole.create", {
    values: {
      acl_role_id: role.value,
      entity_table: "civicrm_group",
      entity_id: subjects,
      is_active: true,
    },
  });
  const [acl] = h.cv("ACL.create", {
    values: {
      name: `outlook365 e2e ${run}`,
      deny: false,
      entity_table: "civicrm_acl_role",
      entity_id: role.value,
      operation: "View",
      object_table: "civicrm_group",
      object_id: allowed,
      is_active: true,
    },
  });
  created.acls.push(acl.id);
  minted = h.mintToken(subject, SCOPE);
});

test.afterAll(() => {
  h.revokeToken(minted.id);
  h.cv("ACL.delete", { where: [["id", "IN", created.acls]] });
  h.cv("ACLEntityRole.delete", { where: [["acl_role_id", "=", created.roleValue]] });
  h.cv("OptionValue.delete", { where: [["id", "IN", created.roles]] });
  h.cv("Group.delete", { where: [["id", "IN", created.groups]] });
  h.cv("Contact.delete", { where: [["id", "IN", created.contacts]], useTrash: false });
});

test("the contact search lists only contacts the token's ACL may view", async ({ page }) => {
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await expect(page.locator("#loadingContacts")).toBeHidden();
  await page.locator("#searchField").fill(run);
  await page.locator("#searchField").press("Enter");
  const found = page.locator("#contacts .ms-Persona .ms-Persona-primaryText");
  await expect(found).toHaveText([permitted.name]);
  await expect(page.locator("#civicrm-error")).toBeHidden();
  h.assertNetwork(log, minted.token, key);
});

test("a group's members leave out contacts the token's ACL may not view, not those marked do not trade", async ({
  page,
}) => {
  // The group id is the caller's to choose; the listing is replaced so the pane asks for this one.
  await page.route(`${h.BASE_URL}/civicrm/ajax/rest`, async (r) => {
    const body = new URLSearchParams(r.request().postData() || "");
    if (body.get("entity") !== "Group" || body.get("action") !== "get") {
      return r.fallback();
    }
    return r.fulfill({
      contentType: "application/json",
      body: JSON.stringify({
        is_error: 0,
        count: 1,
        values: [{ id: String(mixed), name: run, title: mixedTitle }],
      }),
    });
  });
  const log = await h.openPane(page, "taskpane.html", { item: compose, roaming: configured() });
  await page.getByRole("tab", { name: "Groups", exact: true }).click();
  const row = page.locator("#groups .ms-Persona").filter({ hasText: mixedTitle });
  await row.locator(".group-toggle").click();
  await expect(row.locator(".group-toggle")).toHaveAttribute("aria-expanded", "true");
  await expect(row.locator(".ms-ListItem")).toHaveText([noTrade.name, permitted.name]);
  await expect(page.locator("#civicrm-error")).toBeHidden();
  h.assertNetwork(log, minted.token, key);
});
