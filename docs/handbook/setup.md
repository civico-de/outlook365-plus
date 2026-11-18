---
title: Setup
order: 10
---

# Setup

## CiviCRM, once

1. Enable `authx`, `dataprocessor`, `keyring` and this extension.
2. Grant the permission **Access Outlook 365 pages** to anonymous visitors. Outlook loads the add-in's panes without a CiviCRM session.
    - Standalone: the role **Everyone, including anonymous users**.
    - WordPress: **Anonymous User**.
    - Drupal: both **Anonymous User** and **Authenticated User**.
3. Leave the authx settings at their stock values: `authx_guards` contains `perm`, and `authx_xheader_cred` contains `api_key`.

## A token for each user

Every person who uses the add-in needs their own keyring token, issued on their contact's **API tokens** tab. The token is shown once, when it is issued; a lost token is replaced, never recovered. Hand it over on a channel that does not keep it.

The token's scope needs the permissions listed under [Token scope](reference.md#token-scope). The contact also needs the permission **authenticate with api key** through their role, on top of the token's own scope, because the add-in sends no site key.

A token belongs to one person: everything the add-in does is recorded as done by that contact, and the token can be revoked on its own.

## The manifest

In CiviCRM, open **Administer → Download Outlook 365 Manifest.xml** and save the file `manifest.xml`. It installs the add-in in Outlook. The add-in carries the name of your organisation's own contact in CiviCRM, and the manifest points Outlook at your site's URL, so download it from the site the add-in should talk to.

## Deploy the add-in

There are two ways. Microsoft changes its admin screens often, so the labels below may differ slightly.

### For the whole organisation

1. Open the Microsoft 365 admin center.
2. Go to **Settings → Integrated apps** and open the add-ins page.
3. Choose **Deploy Add-in**, then **Upload custom apps**.
4. Upload `manifest.xml`, choose who gets the add-in and deploy it.

### For one user

1. In Outlook, open a mail or start a new one.
2. Open the **…** menu and choose **Get Add-ins**.
3. Under **My add-ins**, choose **Add a custom add-in → Add from file** and pick `manifest.xml`.
4. Confirm the installation.

## First use

Each user configures the add-in once:

1. In Outlook, open a mail or start a new one.
2. Open the **…** menu and choose the add-in, shown with the CiviCRM logo and your organisation's name.
3. In the pane, choose **Settings**.
4. Enter the **CiviCRM Site URL** and the **CiviCRM API token (keyring)**, then **Save**.

The site URL is the base address of the site, for example `https://crm.example.org`, not a page below it. It must be the same site that serves the add-in's panes: the token travels in a request header, and a call to another site would need a cross-origin preflight that CiviCRM does not answer.

To change the URL later, open **Settings** again. Leaving the token field empty keeps the stored token as long as the site URL stays the same.
