// The read task pane: the mail's contacts, saving contacts, the mail and folders.
import { test, expect } from '@playwright/test';
import * as h from './harness.mjs';

const SCOPE = ['access CiviCRM', 'access AJAX API', 'view all contacts', 'edit all contacts', 'add contacts', 'view all activities', 'edit all activities'];
const run = `pw${Date.now().toString(36)}`;
const known = { displayName: `Kim Known ${run}`, emailAddress: `kim.${run}@example.org` };
const newcomer = (n) => ({ displayName: `Nia New${n} ${run}`, emailAddress: `nia${n}.${run}@example.org` });
const SAVED = 'Saved in CiviCRM';

let subject;
let key;
let minted;
let knownId;

const message = (n, extra = {}) => ({
  mode: 'read',
  from: known,
  to: [newcomer(n)],
  cc: [],
  bcc: [],
  subject: `Example subject ${run} ${n}`,
  body: '<p>Example body</p>',
  dateTimeCreated: '2026-01-02T03:04:05Z',
  ...extra,
});
const open = (page, item, rest) =>
  h.openPane(page, 'readtaskpane.html', { item, roaming: { civicrm_url: h.BASE_URL, civicrm_token: minted.token } }, rest);
const rows = (page) => page.locator('#contacts .ms-List > .ms-ListItem');
const contactsWithEmail = (email) =>
  h.sql(`SELECT contact_id FROM civicrm_email WHERE email = '${email}' ORDER BY contact_id`).split('\n').filter(Boolean).map(Number);
// Source (record type 2) and targets (3) of the newest activity with this subject.
function activity(subjectLine) {
  const id = h.sql(`SELECT id FROM civicrm_activity WHERE subject = '${subjectLine}' ORDER BY id DESC LIMIT 1`);
  const of = (type) =>
    h.sql(`SELECT contact_id FROM civicrm_activity_contact WHERE activity_id = ${Number(id)} AND record_type_id = ${type} ORDER BY contact_id`)
      .split('\n').filter(Boolean).map(Number);
  return { id: Number(id), source: of(2), targets: of(3) };
}

test.beforeAll(() => {
  subject = h.seedSubject();
  key = h.siteKey();
  [{ id: knownId }] = h.cv('Contact.create', { values: { contact_type: 'Individual', first_name: 'Kim', last_name: `Known ${run}` } });
  h.cv('Email.create', { values: { contact_id: knownId, email: known.emailAddress, is_primary: true } });
  minted = h.mintToken(subject, SCOPE);
});

test.afterAll(() => h.revokeToken(minted.id));

test('shows which of the mail\'s addresses CiviCRM knows', async ({ page }) => {
  const log = await open(page, message(1));
  await expect(rows(page)).toHaveCount(2);
  await expect(rows(page).nth(0)).toHaveClass(/not_to_save/);
  await expect(rows(page).nth(0).locator('.ms-ListItem-secondaryText')).toHaveText(known.displayName);
  await expect(rows(page).nth(0).locator('a')).toHaveAttribute('href', new RegExp(`cid=${knownId}$`));
  await expect(rows(page).nth(1)).toHaveClass(/to_save/);
  await expect(rows(page).nth(1).locator('.save-contact')).toBeVisible();
  h.assertNetwork(log, minted.token, key);
});

test('saves an unknown address as a contact', async ({ page }) => {
  const log = await open(page, message(2));
  await rows(page).nth(1).locator('.save-contact').click();
  await expect(page.locator('#saved-contact-help-text')).toBeVisible();
  await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);
  expect(contactsWithEmail(newcomer(2).emailAddress)).toHaveLength(1);
  h.assertNetwork(log, minted.token, key);
});

test('saves the mail as an activity of the token\'s contact and tags it', async ({ page }) => {
  const item = message(3);
  const log = await open(page, item);
  await expect(rows(page)).toHaveCount(2);
  await page.locator('.save-email').click();
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
  await page.locator('.save-contact-all-group').click();
  const dialog = page.frameLocator('#office-dialog');
  await expect(dialog.locator('.group-select-dropdown option', { hasText: 'Administrators' })).toBeAttached();
  await dialog.locator('#civicrm-group-field').fill(groupTitle);
  await dialog.locator('#group-select-done').click();
  await expect(page.locator('#office-dialog')).toHaveCount(0);
}
const groupMembers = (groupTitle) =>
  h.sql(`SELECT gc.contact_id FROM civicrm_group_contact gc JOIN civicrm_group g ON g.id = gc.group_id
    WHERE g.title = '${groupTitle}' AND gc.status = 'Added' ORDER BY gc.contact_id`).split('\n').filter(Boolean).map(Number);

test('adds the selected contacts to a new group', async ({ page }) => {
  const withGroups = h.mintToken(subject, [...SCOPE, 'edit groups']);
  const groupTitle = `Example group ${run}`;
  const log = await h.openPane(page, 'readtaskpane.html', {
    item: message(4),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: withGroups.token },
  });
  await addToNewGroup(page, groupTitle);
  // The pane reloads once everything is saved; the reloaded list knows the newcomer.
  await expect(rows(page).nth(1)).toHaveClass(/not_to_save/);
  await expect(page.locator('.ms-Dialog')).toBeHidden();

  const [created] = contactsWithEmail(newcomer(4).emailAddress);
  expect(groupMembers(groupTitle)).toEqual([knownId, created].sort((a, b) => a - b));
  h.assertNetwork(log, withGroups.token, key);
  h.revokeToken(withGroups.id);
});

test('a scope without edit groups shows an error instead of a silent success', async ({ page }) => {
  const groupTitle = `Example refused group ${run}`;
  const log = await open(page, message(7));
  await addToNewGroup(page, groupTitle);
  await expect(page.locator('.ms-Dialog')).toBeVisible();
  expect(h.sql(`SELECT COUNT(*) FROM civicrm_group WHERE title = '${groupTitle}'`)).toBe('0');
  h.assertNetwork(log, minted.token, key);
});

test('saves a mail folder through the Outlook REST API', async ({ page }) => {
  const mail = { Id: 'm1', Subject: `Folder mail ${run}`, Body: { Content: '<p>Example</p>' }, Categories: [],
    From: { EmailAddress: { Address: known.emailAddress } }, ToRecipients: [] };
  const patched = [];
  const log = await open(page, message(5), (request) => {
    const url = request.url();
    if (request.method() === 'PATCH') {
      patched.push(url);
      return { json: {} };
    }
    if (url.endsWith('/v2.0/me/MailFolders')) {
      return { json: { value: [{ Id: 'f1', DisplayName: 'Example folder' }] } };
    }
    return { json: { value: url.includes('/MailFolders/f1/messages') ? [mail] : [] } };
  });
  await page.locator('#target .ms-ListItem', { hasText: 'Example folder' }).click();
  await page.locator('#send-submit').click();
  await expect(page.locator('#saved-email-notice')).toBeVisible();

  const saved = activity(mail.Subject);
  expect(saved.source).toEqual([subject]);
  expect(saved.targets).toEqual([knownId]);
  expect(patched).toEqual([`${h.REST_URL}/v2.0/me/messages/m1`]);
  h.assertNetwork(log, minted.token, key);
});

test('a revoked token shows the error dialog', async ({ page }) => {
  const revoked = h.mintToken(subject, SCOPE);
  h.revokeToken(revoked.id);
  const log = await h.openPane(page, 'readtaskpane.html', {
    item: message(6),
    roaming: { civicrm_url: h.BASE_URL, civicrm_token: revoked.token },
  });
  await expect(page.locator('.ms-Dialog')).toBeVisible();
  await expect(page.locator('#civicrm-notice-text')).toContainText('An error occurred communicating with CiviCRM');
  h.assertNetwork(log, revoked.token, key);
});
