<?php

declare(strict_types = 1);

// Seeds the e2e subject: a NON-superuser contact, a Standalone user and a role
// holding the permissions the add-in's calls need, plus the anonymous grant
// that makes the task pane reachable from Outlook.
// Prints "contactId=<id>". Fixtures use example.org only.

/**
 * The id of an APIv4 row. A seeder that silently seeded nothing produces
 * failures further on that look like defects in the add-in.
 *
 * @param array<array-key, mixed>|null $row
 */
function outlook365_seed_id(?array $row, string $what): int {
  if ($row === NULL || !isset($row['id']) || !is_numeric($row['id'])) {
    throw new RuntimeException("outlook365 seed: no id for {$what}");
  }
  return (int) $row['id'];
}

/**
 * A role's permission list, which APIv4 types as mixed.
 *
 * @param array<array-key, mixed> $row
 * @return list<string>
 */
function outlook365_seed_permissions(array $row): array {
  $permissions = $row['permissions'] ?? [];
  if (!is_array($permissions)) {
    return [];
  }
  return array_values(array_filter($permissions, 'is_string'));
}

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
  'edit groups',
  'access outlook 365 pages',
];

$role = \Civi\Api4\Role::get(checkPermissions: FALSE)->addWhere('name', '=', $roleName)->execute()->first();
$role ??= \Civi\Api4\Role::create(checkPermissions: FALSE)
  ->addValue('name', $roleName)
  ->addValue('label', 'outlook365 e2e subject')
  ->addValue(fieldName: 'is_active', value: TRUE)
  ->addValue('permissions', $perms)
  ->execute()
  ->first();
$roleId = outlook365_seed_id($role, 'role ' . $roleName);
\Civi\Api4\Role::update(checkPermissions: FALSE)
  ->addWhere('id', '=', $roleId)
  ->addValue('permissions', $perms)
  ->execute();

// Outlook loads the task pane without a CiviCRM session, so the pages are
// served to anonymous visitors. Upstream's permission says exactly that.
$everyone = \Civi\Api4\Role::get(checkPermissions: FALSE)->addWhere('name', '=', 'everyone')->execute()->single();
$anonPerms = outlook365_seed_permissions($everyone);
if (!in_array('access outlook 365 pages', $anonPerms, TRUE)) {
  $anonPerms[] = 'access outlook 365 pages';
  \Civi\Api4\Role::update(checkPermissions: FALSE)
    ->addWhere('id', '=', outlook365_seed_id($everyone, 'role everyone'))
    ->addValue('permissions', $anonPerms)
    ->execute();
}

$contact = \Civi\Api4\Contact::get(checkPermissions: FALSE)
  ->addWhere('external_identifier', '=', 'outlook365-e2e-subject')
  ->execute()
  ->first();
if ($contact === NULL) {
  $contact = \Civi\Api4\Contact::create(checkPermissions: FALSE)
    ->addValue('contact_type', 'Individual')
    ->addValue('first_name', 'Outlook')
    ->addValue('last_name', 'E2e')
    ->addValue('external_identifier', 'outlook365-e2e-subject')
    ->execute()
    ->first();
  \Civi\Api4\Email::create(checkPermissions: FALSE)
    ->addValue('contact_id', outlook365_seed_id($contact, 'the subject contact'))
    ->addValue('email', 'outlook365-e2e@example.org')
    ->addValue(fieldName: 'is_primary', value: TRUE)
    ->execute();
}
$contactId = outlook365_seed_id($contact, 'the subject contact');

$user = \Civi\Api4\User::get(checkPermissions: FALSE)->addWhere('username', '=', 'outlook365-e2e')->execute()->first();
$user ??= \Civi\Api4\User::create(checkPermissions: FALSE)
  ->addValue('username', 'outlook365-e2e')
  ->addValue('contact_id', $contactId)
  ->addValue('uf_name', 'outlook365-e2e@example.org')
  ->execute()
  ->first();
\Civi\Api4\User::update(checkPermissions: FALSE)
  ->addWhere('id', '=', outlook365_seed_id($user, 'the subject user'))
  ->addValue('roles', [$roleId])
  ->addValue(fieldName: 'is_active', value: TRUE)
  ->execute();

echo 'contactId=' . $contactId . "\n";
