---
title: Using the add-in
order: 20
---

# Using the add-in

Open the add-in from the **…** menu of a mail. It opens a different pane depending on whether you are writing or reading the mail. Until the add-in is configured, the pane asks you to open **Settings**, see [First use](setup.md#first-use).

When a call to CiviCRM fails, for example because the token's scope does not allow it, the pane shows an error and the action is not completed.

## Writing a mail

The pane has two tabs, **Contacts** and **Groups**. Both lists load 25 entries at a time and load more as you scroll.

### Contacts

Type a name into the search field and press Enter. The list shows matching contacts with their primary email address. **To**, **Cc** and **Bcc** add the contact to that field of the mail; an address that is already there is not added twice. The contact icon opens the contact in CiviCRM.

The list leaves out contacts that are deleted or deceased, that have "do not email" or "do not trade" set, and that the signed-in user may not view.

### Groups

Search groups by title the same way. Open a group to list its members with their primary email address; a search field inside the group filters the members by name. Select members one by one or with **Select All**, then add the selected ones with **To**, **Cc** or **Bcc**. Each group's buttons only take that group's selection, so you can put one group in To and another in Bcc.

The settings icon next to a group opens the group's settings in CiviCRM.

## Reading a mail

The pane lists the sender and every To, Cc and Bcc recipient, all selected at first.

- An address that belongs to a CiviCRM contact shows the contact's name and a link to the contact.
- An address that does not has a save button. It creates a contact of type Individual with that email address. The display name from the mail is split at the first space (or, without a space, at the first comma) into first and last name. A name with neither, such as a bare email address, leaves first and last name empty, and CiviCRM shows the email address as the contact's name.

Below the list:

- **Save Contacts** creates contacts for all selected addresses that are not in CiviCRM yet.
- **Save Contacts to Group** asks for a group: an existing active group, or the name of a new one. It creates the missing contacts and adds every selected person to the group.
- **Save Email** files the mail in CiviCRM as an activity of type Email. Selected people who are not in CiviCRM yet are created first.

The Email activity holds:

| Field | Value |
|---|---|
| Subject | The subject of the mail |
| Details | The HTML body of the mail |
| Date | When the mail was created |
| Added by | The contact the token belongs to |
| With contact | The selected people |

After the activity is saved, the mail gets the Outlook category **Saved in CiviCRM**. The add-in creates the category in your mailbox the first time it needs it.
