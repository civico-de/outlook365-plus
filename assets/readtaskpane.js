// ck-allow-api3 -- upstream's APIv3 call sites, routed through assets/civicrm.js; migrating them belongs upstream
(function () {
  'use strict';
  Office.initialize = function (reason) {

    var requiredAttendees = Office.context.mailbox.item;
    var config = false;
    var saveDialog = null
    // Save Email needs a contact list drawn for the current config and no save in flight.
    var listReady = false;
    var savingEmail = false;
    var emailSaved = false;
    var listDrawing = 0;

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);

    // The site URL identifies the site a token belongs to, so both sides of
    // that comparison have to be spelled the same way.
    function normalizeUrl(url) {
      return String(url || '').replace(/\/+$/, '');
    }

    jQuery(document).ready(async function () {
      // Set localized text for UI elements.
      await reset();

      if (config) {
        showContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);
      // The button sits outside the redrawn contact list, so it is bound once here.
      $('.save-email').on('click', saveEmail);

      $("body").delegate(".selectAll", "click", function (event) {
        setSelected($(".ms-ListItem-selectionTarget"), true)
      });

      $("body").delegate(".unselectAll", "click", function (event) {
        setSelected($(".ms-ListItem-selectionTarget"), false)
      });

      $("body").delegate(".ms-ListItem-selectionTarget", "change", function () {
        setSelected(this, $(this).prop('checked'));
      });

      $("body").delegate(".ms-ListItem", "click", function (event) {
        if ($(event.target).closest('a, .ms-ListItem-action, .ms-ListItem-selectionTarget').length) {
          return;
        }
        let box = $(this).children('.ms-ListItem-selectionTarget');
        setSelected(box, !box.prop('checked'));
      });

    });

    // The row's checkbox holds the selection; the class only mirrors it for Fabric's styling.
    function setSelected(box, selected) {
      $(box).prop('checked', selected).closest('.ms-ListItem').toggleClass('is-selected', selected);
    }

    function formatDateTimeForCivi(date) {
      let hours = String(date.getHours()).padStart(2, "0");
      let minutes = String(date.getMinutes()).padStart(2, "0");
      let seconds = String(date.getSeconds()).padStart(2, "0");
      return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + " " + hours + ":" + minutes + ":" + seconds;
    }

    async function pushEmailActivity(subject, body, date, targetContactIds) {
      return CiviCRM.api4(config, 'Activity', 'create', {
        values: {
          "source_contact_id": "user_contact_id",
          "activity_type_id:name": "Email",
          "subject": subject,
          "details": body,
          "activity_date_time": formatDateTimeForCivi(date),
          "target_contact_id": targetContactIds,
        },
      }).then(() => true, () => {
        openDialog(dialogComponent);
        return false;
      });
    }

    async function reset() {
      config = await getConfig();
      if (!config) {
        $('.not-configured-warning').show();
        $('#search-form').hide();
      } else {
        $('.not-configured-warning').hide();
      }
      $('#loadingContacts').hide();
    }

    async function openSettingsDialog() {
      // Only non-secret values reach the dialog URL. A token in a query string
      // lands in the web server's access log and the browser history.
      var previous = config;
      var url = settingsDialogUrl;
      var prefill = previous
        ? {url: previous.url, hasToken: !!previous.token}
        : (crmBaseUrl ? {url: crmBaseUrl, hasToken: false} : null);
      if (prefill) {
        url = url + '?config=' + encodeURIComponent(JSON.stringify(prefill));
      }

      var dialogOptions = {width: 30, height: 40, displayInIframe: true};

      Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
        var settingsDialog = result.value;
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function (message) {
          var settings = JSON.parse(message.message);
          var sameSite = previous && normalizeUrl(settings.url) === previous.url;
          config = {
            url: normalizeUrl(settings.url),
            token: settings.token || (sameSite ? previous.token : '')
          };
          clearContacts();
          await setConfig(config, async function (result) {
            settingsDialog.close();
            settingsDialog = null;
            await reset();
            showContacts()
          });
        });
      });
    }

    // Names come from the mail's sender and from CiviCRM, so they only ever go in as text.
    function getListItem(item, res) {
      let known = res['exist'];
      let name = known ? res['contact_name'] : item.displayName;
      let action = $('<div class="ms-ListItem-action">')
        .attr({'data-civicrm-name': name, 'data-civicrm-email': item.emailAddress});
      if (known) {
        action.attr('data-civicrm-id', String(res['contact_id'])).append(
          $('<a target="_blank">').attr({'href': res.contact_url, 'title': UIText.ContactScreen.TitleURL}).append(
            $('<i class="ms-Icon ms-Icon--Contact" aria-hidden="true">')
          )
        );
      } else {
        action.append(
          $('<button type="button" class="icon-button save-contact">')
            .attr('title', UIText.SaveContactInGroupScreen.SaveContact)
            .append($('<i class="ms-Icon ms-Icon--Save" aria-hidden="true">'))
        );
      }
      return $('<li class="ms-ListItem is-selectable is-selected">')
        .addClass(known ? 'not_to_save' : 'to_save')
        .append($('<input type="checkbox" class="ms-ListItem-selectionTarget" checked>').attr('aria-label', name))
        .append($('<span class="ms-ListItem-secondaryText">').text(name))
        .append($('<span class="ms-ListItem-tertiaryText">').text(item.emailAddress))
        .append($('<div class="ms-ListItem-actions">').append(action));
    }

    async function saveContactToCRM(contact) {
      let contactName = splitContactName(contact.name);
      var contact_info
      // The email goes in with the contact: CiviCRM refuses an Individual without a name
      // or an email, and falls back to the email as display name.
      await CiviCRM.api3(config, 'Contact', 'create', {
        "first_name": contactName.firstName,
        'last_name': contactName.lastName,
        "email": contact.email,
        "contact_type": 'Individual',
      }).done(function (result) {
        contact_info = result
      }).catch(response => {
        openDialog(dialogComponent);
      });
      return contact_info ? contact_info.id : null;
    }

    async function saveContact() {
      $(this).prop('disabled', true);
      showSavingContactInfo();
      let action = $(this).closest('.ms-ListItem-action');
      let name = action.attr('data-civicrm-name')
      let email = action.attr('data-civicrm-email');
      let contact = {name: name, email: email};
      let saved = await saveContactToCRM(contact);
      $(this).prop('disabled', false);
      if (!saved) {
        $('#saving-contact-help-text').toggle(false);
        return;
      }
      // Re-rendering the list recreates the notice hidden, so show it afterwards.
      await showContacts();
      showContactSavedInfo();
    }

    async function confirmSaveAllContact(event) {
      $(this).prop('disabled', true);
      showSavingContactInfo();
      let saved = await saveAllContact();
      $(this).prop('disabled', false);
      if (!saved) {
        $('#saving-contact-help-text').toggle(false);
        return;
      }
      showContactSavedInfo();
      location.reload();
    }

    function showSavingContactInfo() {
      $('#saving-contact-help-text').toggle(true);
      $('#saved-contact-help-text').toggle(false);
    }

    function showContactSavedInfo() {
      $('#saving-contact-help-text').toggle(false);
      $('#saved-contact-help-text').toggle(true);
    }

    /**
     * Saves the selected unknown addresses; stops at the first failure, which
     * has shown the error dialog, and returns false.
     */
    async function saveAllContact() {
      let toSave = []
      $('.ms-List').children().each(function (index) {
        if ($(this).hasClass('to_save') && $(this).hasClass('is-selected')) {
          let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
          toSave.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email')])
        }
      });

      for (const [name, email] of toSave) {
        if (!(await saveContactToCRM({"name": name, "email": email}))) {
          return false;
        }
      }
      return true;
    }

    async function confirmSaveAllContactInGroup(event) {
      let toSave = [];
      $('.ms-List').children().each(function (index) {
        let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
        // name,email,already saved
        if ($(this).hasClass('is-selected')) {
          if ($(this).hasClass('to_save')) {
            toSave.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email'), $(this).hasClass('to_save')])
          } else {
            toSave.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email'), $(this).hasClass('to_save'), contact.data('civicrm-id')])
          }
        }
      });
      if (toSave.length == 0) {
        console.log("None selected")
        return
      }
      // The dialog cannot read the stored token; it gets the groups (ids and
      // titles) through messageChild, which needs DialogApi 1.2.
      if (!Office.context.requirements.isSetSupported('DialogApi', '1.2')) {
        openDialog(dialogComponent);
        return;
      }
      var groups = await CiviCRM.api3(config, 'Group', 'get', {
        "sequential": 1,
        "return": ["id", "title"],
        "is_active": '1',
        "is_hidden": '0',
        "options": {"limit": 0},
      }).then(data => data.is_error ? null : data.values.map(g => [g.id, g.title]), () => null);
      if (!groups) {
        openDialog(dialogComponent);
        return;
      }
      var dialogOptions = {width: 30, height: 40, displayInIframe: true};
      Office.context.ui.displayDialogAsync(saveContactInGroupDialogUrl, dialogOptions, async function (result) {
        var confirmDialog = result.value;
        confirmDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function (message) {
          let groupInfo = JSON.parse(message.message);
          // The dialog reports ready once it listens for the group list.
          if (groupInfo['ready'] === true) {
            confirmDialog.messageChild(JSON.stringify(groups));
            return;
          }
          var groupId
          var groupResult
          if (groupInfo['exist'] === true) {
            groupId = groupInfo["selectedGroup"]
          } else {
            groupResult = await addGroupToCRM(groupInfo['groupName'])
            groupId = groupResult && groupResult["id"]
          }

          // The first failure has shown the error dialog; stop there.
          for (var key in toSave) {
            let contactId = toSave[key][2] === true
              ? await saveContactToCRM({"name": toSave[key][0], "email": toSave[key][1]})
              : toSave[key][3];
            if (!groupId || !contactId || !(await saveContactGroup(contactId, groupId))) {
              confirmDialog.close();
              return;
            }
          }
          confirmDialog.close();
          confirmDialog = null;
          location.reload();
        });

      });

    }

    async function saveContactGroup(contactId, groupId) {
      return CiviCRM.api3(config, 'GroupContact', 'create', {
        "group_id": groupId,
        "contact_id": contactId,
      }).then(() => true, () => {
        openDialog(dialogComponent);
        return false;
      });
    }

    async function addGroupToCRM(name) {
      let groupInfo
      await CiviCRM.api3(config, 'Group', 'create', {
        "title": name,
        "name": name,
      }).done(function (result) {
        groupInfo = result
      }).catch(response => {
        openDialog(dialogComponent);
      });

      return groupInfo
    }

    async function checkContact(val) {
      var exist = null
      var contact_url = CRMContactURL
      var contact_name = null
      var contact_id
      await CiviCRM.api3(config, 'Email', 'get', {"email": val.emailAddress}).done(function (result) {
        exist = false
        if (result["count"] > 0) {
          exist = true
          let keys = Object.keys(result["values"])
          contact_id = result["values"][keys[0]]["contact_id"]
          contact_url += "&cid=" + String(contact_id)
        }
      }).catch(response => {
        openDialog(dialogComponent);
      });
      if (exist === null) {
        return null;
      }
      if (exist) {
        await CiviCRM.api3(config, 'Contact', 'getsingle', {"id": contact_id}).done(function (result) {
          contact_name = result["display_name"]
        }).catch(response => {
          openDialog(dialogComponent);
        });
      }

      return {"exist": exist, "contact_url": contact_url, "contact_name": contact_name, "contact_id": contact_id}
    }

    // Contact IDs belong to one site, so a list is dropped as soon as the config changes.
    function clearContacts() {
      listDrawing++;
      listReady = false;
      $("#contacts").empty();
      updateSaveEmail();
    }

    function updateSaveEmail() {
      let usable = listReady && !savingEmail && !emailSaved;
      $('.save-email').prop('disabled', !usable).attr('aria-disabled', String(!usable));
      $('#saved-email-help-text').prop('hidden', !emailSaved);
    }

    // Item categories need Mailbox 1.8; an older client can neither read nor set the tag.
    function hasCategories() {
      return Office.context.requirements.isSetSupported('Mailbox', '1.8');
    }

    // Resolves null when Outlook cannot tell, so an unverified mail is not offered for saving.
    function hasSavedCategory() {
      if (!hasCategories()) {
        return Promise.resolve(false);
      }
      return new Promise(resolve => requiredAttendees.categories.getAsync(result => resolve(
        result.status === Office.AsyncResultStatus.Failed
          ? null
          : result.value.some(category => category.displayName === 'Saved in CiviCRM')
      )));
    }

    async function showContacts() {
      clearContacts();
      let drawing = listDrawing;
      $('#loadingContacts').show();

      let list = $('<ul class="ms-List">');

      // Sender, then To, Cc and Bcc. A lookup that failed has shown the error
      // dialog; listing its address as unknown would invite a duplicate.
      let addresses = [requiredAttendees.from].concat(requiredAttendees.to, requiredAttendees.cc, requiredAttendees.bcc);
      for (const val of addresses) {
        let res = await checkContact(val)
        if (drawing !== listDrawing) {
          return;
        }
        if (!res) {
          $('#loadingContacts').hide();
          return;
        }
        list.append(getListItem(val, res));
      }
      let filed = await hasSavedCategory();
      if (drawing !== listDrawing) {
        return;
      }
      if (filed === null) {
        $('#loadingContacts').hide();
        openDialog(dialogComponent);
        return;
      }

      let html = '<button class="ms-Button ms-Button--small save-contact-all">' +
        '<span class="ms-Button-label">' + UIText.SaveContactInGroupScreen.Save + '</span>' +
        '</button>'
      html += '<button class="ms-Button ms-Button--small save-contact-all-group">' +
        '<span class="ms-Button-label">' + UIText.SaveContactInGroupScreen.Title + '</span>' +
        '</button>'
      html += '<p id="saving-contact-help-text" style="display: none;">' + UIText.SaveContactInGroupScreen.SavingText + '</p>';
      html += '<p id="saved-contact-help-text" style="display: none;">' + UIText.SaveContactInGroupScreen.SavedText + '</p>';

      html += '<br><br>';


      $("#contacts").empty().append(list, html)
      $(".save-contact").on('click', saveContact)
      $(".save-contact-all-group").on('click', confirmSaveAllContactInGroup)
      $(".save-contact-all").on('click', confirmSaveAllContact)
      $('#loadingContacts').hide();
      emailSaved = emailSaved || filed;
      listReady = true;
      updateSaveEmail();
    }

    async function saveEmail() {
      if (savingEmail || emailSaved || !listReady) {
        return;
      }
      savingEmail = true;
      updateSaveEmail();
      let site = config;
      let drawn = listDrawing;
      let createdContacts = false;
      try {
        createdContacts = await fileEmail(site, drawn);
      } finally {
        savingEmail = false;
        updateSaveEmail();
      }
      // Contacts created on the way are known now; their rows must not offer to create them again.
      if (createdContacts && config === site) {
        await showContacts();
      }
    }

    // Saves the selected contacts and the mail; resolves whether it created contacts.
    async function fileEmail(site, drawn) {
      let result = await new Promise(resolve => requiredAttendees.body.getAsync("html", resolve));
      // A settings change or a contact save has redrawn the list the selection came from.
      if (config !== site || listDrawing !== drawn) {
        return false;
      }
      var emailBody = result['value']
      var targetContactIds = []
      var pendingContacts = []

      $('.ms-List').children().each(function (index) {
        let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
        if ($(this).hasClass('is-selected')) {
          if ($(this).hasClass('to_save')) {
            pendingContacts.push(
              saveContactToCRM({"name": contact.attr('data-civicrm-name'), "email": contact.attr('data-civicrm-email')})
            );
          } else {
            targetContactIds.push(contact.data('civicrm-id'));
          }
        }
      });

      // A failed save has shown the error dialog; the activity would lack that contact.
      let savedIds = await Promise.all(pendingContacts);
      let created = savedIds.some(id => id !== null);
      if (savedIds.includes(null) || config !== site) {
        return created;
      }
      targetContactIds = targetContactIds.concat(savedIds);

      if (await pushEmailActivity(requiredAttendees.subject, emailBody, new Date(requiredAttendees.dateTimeCreated), targetContactIds)) {
        emailSaved = true;
        addSavedCategoryToCurrentMessage();
      }
      return created;
    }

    function addSavedCategoryToCurrentMessage() {
      if (!hasCategories()) {
        return;
      }
      // skip if the message has sent to CiviCRM
      Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("Action failed with error: " + asyncResult.error.message);
        } else {
          var categories = asyncResult.value;
          var hasSaved = false;
          categories.forEach(function (item) {
            if (item.displayName === 'Saved in CiviCRM') {
              hasSaved = true;
            }
          });
          if (!hasSaved) {
            Office.context.mailbox.item.categories.addAsync(['Saved in CiviCRM'], result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully added categories");
              } else {
                console.log("categories.addAsync call failed with error: " + result.error.message);
                // create the category if it is not exist and try again
                if (result.error.code === 9044) {
                  var masterCategoriesToAdd = [
                    {
                      "displayName": "Saved in CiviCRM",
                      "color": Office.MailboxEnums.CategoryColor.Preset0
                    }
                  ];

                  Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                      console.log("Successfully added categories to master list");
                      addSavedCategoryToCurrentMessage();
                    } else {
                      console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
                    }
                  });
                }
              }
            });
          }
        }
      });

    }

    async function getConfig() {
      var config = {};

      config.url = normalizeUrl(Office.context.roamingSettings.get('civicrm_url'));
      config.token = Office.context.roamingSettings.get('civicrm_token');
      if (config.url && config.token) {
        return config;
      }
      return null;
    }

    /**
     * Save the configuration.
     *
     * @param config
     * @param callback
     */
    async function setConfig(config, callback) {
      await Office.context.roamingSettings.set('civicrm_url', config.url);
      await Office.context.roamingSettings.set('civicrm_token', config.token);
      Office.context.roamingSettings.saveAsync(callback);
    }

    function splitContactName(name) {
      let separatorIndex = name.indexOf(' ');
      if (separatorIndex === -1) {
        separatorIndex = name.indexOf(',');
      }
      // A single word is a first name, as in CiviCRM's own inbound mail; an email address is no name.
      if (separatorIndex === -1) {
        return {firstName: name.includes('@') ? '' : name, lastName: ''};
      }
      return {firstName: name.substring(0, separatorIndex), lastName: name.substring(separatorIndex + 1)};
    }
  }
})();
