# Changelog

Versions up to 2.2.7 are upstream releases of the Outlook 365 extension; 3.0.0 is the first
release of this fork.

## [Unreleased]

### Fixed

- A mail is saved in CiviCRM once. "Save Email" created the Email activity twice per click after a
  contact or the settings had been saved, and again on every further click, creating the unknown
  contacts again each time. It is now disabled while the save runs and once the mail is saved or
  carries the category "Saved in CiviCRM", and the pane says so.
- "Save Email" in the read pane can only be used once the contact list has loaded. A click while
  the list was loading, after the lookup had failed or before the add-in was configured filed the
  mail without its contacts.
- Saving the settings clears the read pane's contact list until the lookup on the configured site
  has finished, so no contact of the previous site is saved with the new site's credentials.
- Contacts created by "Save Email" show as known in the list afterwards.
- The read pane works on Outlook clients without item categories (before Mailbox 1.8); there
  the mail gets no "Saved in CiviCRM" category. When Outlook cannot read the categories, the pane
  shows an error instead of offering an already filed mail for saving again.
- A contact saved while "Save Email" reads the mail no longer lets the mail be filed without its
  contacts; the click is dropped and the redrawn list offers saving again.
- A contact saved from a sender name that is an email address keeps first and last name empty and
  shows its email address as its name, and a single-word name becomes the first name, instead of
  the last name "Unknown".
- The extension's documentation link points to the handbook.

## [3.1.2] - 2026-10-22

### Fixed

- A task pane no longer shows the answer to an older request after a newer search or a reset has
  replaced it.

## [3.1.1] - 2026-10-22

### Changed

- Requires PHP 8.2 or later.
- The documented token scope no longer lists `edit all activities`, which CiviCRM does not define.

### Fixed

- The task panes are accessible: tab labels and icon-only controls are exposed to assistive
  technology, and the contact and group lists can be operated from the keyboard.
- The group list in the compose pane loads further pages on scroll instead of repeating its first
  page.
- A list no longer takes in responses that arrive after the pane has reset it.

## [3.1.0] - 2026-09-21

### Removed

- "Save Folder in CiviCRM" in the read pane. It relied on the Outlook REST API, which Microsoft has
  retired.

### Fixed

- A failed CiviCRM call shows an error in the task panes, including APIv3 answers that report an
  error.
- Saving all contacts stops at the first failure instead of reporting success.
- No mail activity is created when saving one of its contacts failed.
- The save-to-group dialog gets its group list from the pane instead of reading the settings.
- Group search in the compose pane matches the group title it displays.
- Names from mails and from CiviCRM are shown as text in the task panes, never as markup.

## [3.0.1] - 2026-09-15

### Changed

- Internal maintenance; no change for sites.

## [3.0.0] - 2026-09-13

### Added

- The add-in authenticates with a personal keyring token, sent only in the `X-Civi-Auth` header and
  never in a URL or request body.
- The task pane can be pinned in Outlook.

### Changed

- Requires the keyring extension and CiviCRM 6.17.
- The add-in settings hold the site URL and one API token; the API key and site key fields are
  gone. The settings dialog only learns whether a token is stored, and changing the site URL asks
  for a new token.
- The site URL is the site's base URL instead of its AJAX endpoint.
- Mails are saved as activities through APIv4.
- The manifest's AppDomain works with the new Outlook.

### Removed

- The dependency on `com.cividesk.apikey`.

### Fixed

- The contact search in the compose pane submitted no search term.

## [2.2.7] - 2025-04-02

### Changed

- Saving a contact explicitly creates an Individual and formats the email address.
- Updated the getting started documentation.

### Fixed

- Deleted contacts are excluded.

## [2.2.6] - 2024-08-15

### Fixed

- The compose pane identifies groups by their ID instead of their name, so opening and searching a
  group works whatever its name.

## [2.2.5] - 2024-08-12

### Added

- French (Canada) translation.

### Changed

- The permission is declared in the current CiviCRM format.

## [2.2.4] - 2023-10-12

### Added

- Translations of the task panes.

### Changed

- Updated the getting started documentation.

## [2.2.3] - 2021-11-12

### Changed

- Internal maintenance; no change for sites.

## [2.2.2] - 2021-11-11

### Fixed

- Group members are listed only when their status in the group is Added.

## [2.2.1] - 2021-11-11

### Added

- First release in this changelog, summarising everything up to 2.2.1: an Outlook 365 add-in,
  generated as a downloadable manifest, that searches CiviCRM contacts and groups while writing a
  mail and adds them as To, Cc or Bcc recipients.
- A read pane that shows which sender and recipients are CiviCRM contacts, creates missing
  contacts, adds them to a new or existing group, and saves the mail as an Email activity with its
  recipients as targets.
- Two data processors that back the add-in's contact and group lists.

[Unreleased]: https://github.com/civico-de/outlook365/compare/v3.1.2...HEAD
[3.1.2]: https://github.com/civico-de/outlook365/compare/v3.1.1...v3.1.2
[3.1.1]: https://github.com/civico-de/outlook365/compare/v3.1.0...v3.1.1
[3.1.0]: https://github.com/civico-de/outlook365/compare/v3.0.1...v3.1.0
[3.0.1]: https://github.com/civico-de/outlook365/compare/v3.0.0...v3.0.1
[3.0.0]: https://github.com/civico-de/outlook365/compare/v2.2.7...v3.0.0
[2.2.7]: https://github.com/civico-de/outlook365/compare/v2.2.6...v2.2.7
[2.2.6]: https://github.com/civico-de/outlook365/compare/v2.2.5...v2.2.6
[2.2.5]: https://github.com/civico-de/outlook365/compare/v2.2.4...v2.2.5
[2.2.4]: https://github.com/civico-de/outlook365/compare/v2.2.3...v2.2.4
[2.2.3]: https://github.com/civico-de/outlook365/compare/v2.2.2...v2.2.3
[2.2.2]: https://github.com/civico-de/outlook365/compare/v2.2.1...v2.2.2
[2.2.1]: https://github.com/civico-de/outlook365/releases/tag/v2.2.1
