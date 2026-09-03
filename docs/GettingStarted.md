# Getting started

This page explains how to set up CiviCRM and Outlook 365 integration.

## CiviCRM Installation

First, download and enable the extension dependencies:

* Data Processor, https://civicrm.org/extensions/data-processor
* authx, shipped with CiviCRM core
* keyring, which issues the API tokens the Add-in authenticates with

Then enable the **Outlook 365 Integration** extension

After enabling the dependencies and this extension, install the two required Data Processors which are also shown here, https://lab.civicrm.org/extensions/outlook365/-/tree/master/data-processors?ref_type=heads

* Outlook 365 Contacts
* Outlook 365 Groups

## Obtain the CiviCRM Site URL and an API token

### CiviCRM Site URL

The base URL of the CiviCRM site, for example https://your.domain.name — not a
path below it. It must be the same origin that serves the Add-in's task pane,
because the token travels in a request header and a cross-origin call would
need a CORS preflight CiviCRM does not answer.

### CiviCRM API token

Each user of the Outlook integration needs one keyring token on their CiviCRM
Contact, issued on the contact's **API tokens** tab. The plaintext is shown
once, at issue; a lost token is replaced, never recovered.

The contact also needs the **authenticate with api key** permission, on top of
whatever the token's own scope allows. The Add-in sends no site key, so that
permission is what satisfies authx's guard.

## Provide the Site URL and the token to the end user

Give each user their own token over a channel that does not log it. A token is
a credential for that one person: every action the Add-in performs is
attributed to them, and the token can be revoked on its own.

## CiviCRM Permissions

Outlook loads the task pane with no CiviCRM session, so the **Access Outlook
365 pages** permission has to reach anonymous visitors.

For **Standalone**, add it to the **Everyone, including anonymous users** role.

For **WordPress**, enable it for **Anonymous User**.

For **Drupal**, enable it for both **Anonymous User** and **Authenticated
User** roles.

## Outlook 365 Manifest file (manifest.xml)

In the CiviCRM Menu under *Administer*. Select the **Download Outlook 365 Manifest file.xml** option and download the **manifest.xml** file. This file is used to install the Outlook integration as an **Add-In** into Outlook 365.

## Setting up Outlook 365

There are **two options** for deploying the Outlook 365 Add-Ins, either for all users in the Organisation or on a per user basis.

*Note: Screenshots are not shown for these steps as Microsoft regularly changes the Microsoft 365 admin center.*

### Deploying for the Organisation

1. Go to the [Microsoft 365 admin center](https://admin.microsoft.com/AdminPortal)
2. Click **Settings**
3. Click **Integrated apps**
4. On the Integrated apps page, click the [Add-ins link](https://admin.microsoft.com/Adminportal#/Settings/AddIns)
5. Click the **Deploy Add-in** button
6. Click on Next
7. On the **Deploy a new add-in** page, under the **Deploy a custom add-in** click the **Upload custom apps** button
8. Select **manifest.xml** file from your computer and click **Upload**
9. On the **Configure add-in** page, select options as appropriate for you
10. Click on **Deploy**

The Add-in will be deployed in your Organisation for Outlook 365 users. See the **First use** section below.

### Deploying per user

1. Login into Outlook 365
2. **Open an email** or click on **New message**
3. On the **email page** click on the **...** action button
4. Select the option, **Get Add-Ins**
5. On the **Add-Ins for Outlook** page, select **My add-ins** and scroll to the bottom of the page
6. Select **Add a custom add-in** and select the **Add From file...** option
8. Select **manifest.xml** file from your computer
9. Click **Install** when prompted

The Add-in will now be available in this users Outlook 365. See the *First use configuration* section below.

## First use configuration

The Add-in must be configured for *each user* with the following settings:
* CiviCRM Site URL
* CiviCRM API token

See the *Obtain the CiviCRM Site URL and an API token* section above for more details.

1. Login into Outlook 365
2. **Open an email** or click on **New message**
3. On the **email page** click on the **...** action button
4. Click on icon with the **CiviCRM logo**, your Organisation name should be shown next to this icon
5. The **Add-in pane** will open. Click on the Settings in the **Add-in pane** 
6. Enter the following details: CiviCRM Site URL, CiviCRM API token
6. Click **Save**

If the details have been entered correctly, then the Add-in should now be ready to use.
