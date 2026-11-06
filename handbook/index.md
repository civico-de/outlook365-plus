---
title: Outlook 365 Integration
---

# Outlook 365 Integration

This extension makes CiviCRM an address book inside Microsoft Outlook 365, on the web and in the desktop client. It generates an Outlook add-in that connects Outlook to your CiviCRM site. Staff who write their mail in Outlook can address CiviCRM contacts and groups, add the people in a mail to CiviCRM and file a mail as an activity, without leaving Outlook.

The add-in can be deployed for the whole organisation or by each user.

## What it does

While writing a mail:

- Search CiviCRM contacts and add them as To, Cc or Bcc recipients.
- Search CiviCRM groups, open a group, pick some or all of its members and add them as recipients, so one group can go in To and another in Bcc.

While reading a mail:

- See which of the sender and recipients are already CiviCRM contacts.
- Create contacts for the ones that are not, one at a time or all selected at once.
- Add the selected people to an existing group or to a new one.
- Save the mail as an Email activity with the selected people as recipients.

See [Using the add-in](usage.md) for the details.

## How it signs in

Each user enters the CiviCRM site URL and a personal API token in the add-in's settings. The token is issued by the [keyring](https://github.com/civico-de/keyring) extension and travels only in a request header, never in a URL or form field. Every action the add-in performs is attributed to that user's contact.

This is a maintained fork of the [upstream extension](https://lab.civicrm.org/extensions/outlook365). The fork's difference is this sign-in: upstream sends the API key and the site key in every URL.

## Requirements

- CiviCRM at the version listed in the extension's `info.xml` or newer.
- PHP 8.2 or newer.
- The extensions [Data Processor](https://civicrm.org/extensions/data-processor) 1.1.0 or newer and [keyring](https://github.com/civico-de/keyring), and the core extension `authx`.
- Microsoft Outlook 365.

## Install

Download the release archive from the [releases page](https://github.com/civico-de/outlook365-plus/releases), unpack it into your CiviCRM extensions directory and enable it after its requirements:

```bash
cv ext:enable authx dataprocessor keyring outlook365
```

Enabling installs the two data processors the add-in reads contacts and groups through. Continue with [Setup](setup.md).
