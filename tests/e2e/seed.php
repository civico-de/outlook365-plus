<?php

declare(strict_types = 1);

// Seeds the e2e subject: a NON-superuser contact, a Standalone user and a role
// holding the permissions the add-in's calls need, plus the anonymous grant
// that makes the task pane reachable from Outlook.
// Prints "contactId=<id>". Fixtures use example.org only.

$roleName = 'outlook365_e2e';
$perms = [
  'access CiviCRM',
  'access AJAX API',
  'authenticate with api key',
  'view all contacts',
  'edit all contacts',
  'add contacts',
  'access all custom data',
  'view all activities',
  'edit all activities',
  'access outlook 365 pages',
];

$role = \Civi\Api4\Role::get(FALSE)->addWhere('name', '=', $roleName)->execute()->first();
$role ??= \Civi\Api4\Role::create(FALSE)
  ->addValue('name', $roleName)
  ->addValue('label', 'outlook365 e2e subject')
  ->addValue('is_active', TRUE)
  ->addValue('permissions', $perms)
  ->execute()->first();
\Civi\Api4\Role::update(FALSE)->addWhere('id', '=', $role['id'])
  ->addValue('permissions', $perms)->execute();

// Outlook loads the task pane without a CiviCRM session, so the pages are
// served to anonymous visitors. Upstream's permission says exactly that.
$everyone = \Civi\Api4\Role::get(FALSE)->addWhere('name', '=', 'everyone')->execute()->single();
$anonPerms = $everyone['permissions'] ?? [];
if (!in_array('access outlook 365 pages', $anonPerms, TRUE)) {
  $anonPerms[] = 'access outlook 365 pages';
  \Civi\Api4\Role::update(FALSE)->addWhere('id', '=', $everyone['id'])
    ->addValue('permissions', $anonPerms)->execute();
}

$contact = \Civi\Api4\Contact::get(FALSE)
  ->addWhere('external_identifier', '=', 'outlook365-e2e-subject')->execute()->first();
if ($contact === NULL) {
  $contact = \Civi\Api4\Contact::create(FALSE)
    ->addValue('contact_type', 'Individual')
    ->addValue('first_name', 'Outlook')
    ->addValue('last_name', 'E2e')
    ->addValue('external_identifier', 'outlook365-e2e-subject')
    ->execute()->first();
  \Civi\Api4\Email::create(FALSE)
    ->addValue('contact_id', $contact['id'])
    ->addValue('email', 'outlook365-e2e@example.org')
    ->addValue('is_primary', TRUE)->execute();
}

$user = \Civi\Api4\User::get(FALSE)->addWhere('username', '=', 'outlook365-e2e')->execute()->first();
$user ??= \Civi\Api4\User::create(FALSE)
  ->addValue('username', 'outlook365-e2e')
  ->addValue('contact_id', $contact['id'])
  ->addValue('uf_name', 'outlook365-e2e@example.org')
  ->execute()->first();
\Civi\Api4\User::update(FALSE)->addWhere('id', '=', $user['id'])
  ->addValue('roles', [$role['id']])->addValue('is_active', TRUE)->execute();

echo 'contactId=' . $contact['id'] . "\n";
