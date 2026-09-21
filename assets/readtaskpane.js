// ck-allow-api3 -- upstream's APIv3 call sites, routed through assets/civicrm.js; migrating them belongs upstream
(function () {
  'use strict';
  Office.initialize = function (reason) {

    var requiredAttendees = Office.context.mailbox.item;
    var config = false;
    var saveDialog = null
    var accessToken = null;

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

      Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function (result) {
        accessToken = result.value;
        var getMessageUrl = Office.context.mailbox.restUrl + '/v2.0/me/MailFolders';

        $.ajax({
          url: getMessageUrl,
          dataType: 'json',
          headers: {'Authorization': 'Bearer ' + accessToken}
        }).done(function (data) {
          // Message is passed in `item`.
          $('#target').empty();
          console.log(data);
          for (const item of data.value) {
            let html = $(`<li class="ms-ListItem is-selectable" name="${item.Id}" id="${item.Id}" tabindex="0">` +
              `<span class="ms-ListItem-secondaryText">${item.DisplayName}</span>` +
              '<div class="ms-ListItem-selectionTarget"></div>' +
              '</li>');
            html.appendTo('#target');
          }
          var ListItemElements = document.querySelectorAll(".ms-ListItem");
          for (var i = 0; i < ListItemElements.length; i++) {
            new fabric['ListItem'](ListItemElements[i]);
          }
        }).fail(function (error) {
          // Handle error.
          console.log(error);
        });

        $("#send-submit").on('click', saveEmailsInFolder);
      });

      // fetch all folders

      if (config) {
        showContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);

      $("body").delegate(".selectAll", "click", function (event) {
        $(".ms-ListItem").addClass("is-selected")
      });

      $("body").delegate(".unselectAll", "click", function (event) {
        $(".ms-ListItem").removeClass("is-selected")
      });

      $("body").delegate(".ms-ListItem", "click", function (event) {
        if ($(event.target).closest('a, .ms-ListItem-action').length) {
          return;
        }
        if ($(this).hasClass('is-selected')) {
          $(this).removeClass('is-selected');
        } else {
          $(this).addClass('is-selected');
        }
      });

    });

    /**
     * An event handler for save emails in the selected folders
     * @param event
     */
    function saveEmailsInFolder(event) {
      $(this).prop('disabled', true);
      savingEmailInFolderInfo();
      $('#folder-form li.is-selected').each((index, element) => {
        let folderID = element.getAttribute('name');
        let getMessageUrl = Office.context.mailbox.restUrl +
          '/v2.0/me/MailFolders/' + folderID + '/messages?$expand=SingleValueExtendedProperties';
        $.ajax({
          url: getMessageUrl,
          dataType: 'json',
          headers: {'Authorization': 'Bearer ' + accessToken}
        }).done(async function (data) {
          console.log(data);
          // Message is passed in `item`.
          if (!data.value.length) return;
          for (const email of data.value) {
            if (email.Categories.includes('Saved in CiviCRM')) {
              continue;
            }
            $.ajax({
              url: Office.context.mailbox.restUrl + "/v2.0/me/messages/" + email.Id,
              dataType: 'json',
              contentType: 'application/json',
              method: 'PATCH',
              headers: {'Authorization': 'Bearer ' + accessToken},
              data: JSON.stringify({
                Categories: [
                  "Saved in CiviCRM"
                ]
              })
            }).done(result => {
              console.log(result);
            });
            let emailAddresses = [];
            ['From'].concat(['ToRecipients', 'CcRecipients', 'BccRecipients']).forEach(function (field) {
              let recipients = field === 'From' ? (email.From ? [email.From] : []) : (email[field] || []);
              recipients.forEach(function (recipient) {
                if (recipient.EmailAddress && recipient.EmailAddress.Address) {
                  emailAddresses.push(recipient.EmailAddress.Address);
                }
              });
            });
            let targetContactIds = await getContactIdsByEmail(emailAddresses);
            await pushEmailActivity(email.Subject, email.Body.Content, new Date(), targetContactIds);
          }
          emailInFolderSavedInfo();
        }).fail(function (error) {
          // Handle error.
          console.log(error);
        });
      });
    }

    function savingEmailInFolderInfo() {
      $('#saving-email-notice').show();
      $('#saved-email-notice').hide();
    }

    function emailInFolderSavedInfo() {
      $('#saving-email-notice').hide();
      $('#saved-email-notice').show();
    }

    function formatDateTimeForCivi(date) {
      let hours = String(date.getHours()).padStart(2, "0");
      let minutes = String(date.getMinutes()).padStart(2, "0");
      let seconds = String(date.getSeconds()).padStart(2, "0");
      return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + " " + hours + ":" + minutes + ":" + seconds;
    }

    /**
     * Resolve a list of email addresses to CiviCRM contact IDs in a single API4 call,
     * so the caller doesn't need to look up each contact separately.
     */
    async function getContactIdsByEmail(emailAddresses) {
      let addresses = Array.from(new Set((emailAddresses || []).filter(Boolean)));
      let contactIds = [];
      if (!addresses.length) {
        return contactIds;
      }
      await CiviCRM.api4(config, 'Email', 'get', {
        select: ['contact_id'],
        where: [['email', 'IN', addresses]],
      }).done(function (result) {
        (result.values || []).forEach(function (row) {
          if (contactIds.indexOf(row.contact_id) === -1) {
            contactIds.push(row.contact_id);
          }
        });
      }).fail(function (error) {
        console.log(error);
      });
      return contactIds;
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
          await setConfig(config, function (result) {
            settingsDialog.close();
            settingsDialog = null;
            reset();
            showContacts()
          });
        });
      });
    }

    async function getListItem(item, res) {
      let html = ''
      let name
      let class_to_add
      let contact_id
      if (res['exist']) {
        name = res['contact_name']
        class_to_add = "not_to_save"
        contact_id = res['contact_id']
      } else {
        name = item.displayName
        class_to_add = "to_save"
      }
      html += '<li class="ms-ListItem is-selectable is-selected ' + class_to_add + '">' +
        '<span class="ms-ListItem-secondaryText">' + name + '</span>' +
        '<span class="ms-ListItem-tertiaryText">' + item.emailAddress + '</span>' +
        '<div class="ms-ListItem-actions">' +
        '<div class="ms-ListItem-action" data-civicrm-name="' + name + '" data-civicrm-email="' + item.emailAddress + '"'

      if (res['exist']) {
        html += ' data-civicrm-id="' + String(contact_id) + '">'
      } else {
        html += '>';
      }

      if (res['exist']) {
        html += '<a href="' + res.contact_url + '" target="_blank"><i class="ms-Icon ms-Icon--Contact" title="' + UIText.ContactScreen.TitleURL + '"></i></a>'
      } else {
        html += '<i class="ms-Icon ms-Icon--Save save-contact" title="' + UIText.SaveContactInGroupScreen.SaveContact + '"></i>'
      }

      html += '</div>' +
        '</div>' +
        '</li>'
      return html
    }

    async function saveContactToCRM(contact) {
      let contactName = splitContactName(contact.name);
      var contact_info
      await CiviCRM.api3(config, 'Contact', 'create', {
        "first_name": contactName.firstName,
        'last_name': contactName.lastName,
        "contact_type": 'Individual',
      }).done(function (result) {
        contact_info = result
      }).catch(response => {
        openDialog(dialogComponent);
      });
      if (!contact_info) {
        return null;
      }
      var contact_id = contact_info.id

      await CiviCRM.api3(config, 'Email', 'create', {
        "contact_id": contact_id,
        "email": contact.email,
      }).done(function (result) {
        console.log(result)
      }).catch(response => {
        openDialog(dialogComponent);
        contact_id = null;
      });
      return contact_id
    }

    async function saveContact(event) {
      $(this).prop('disabled', true);
      showSavingContactInfo();
      let name = $(event.target).parent().data('civicrm-name')
      let email = $(event.target).parent().data('civicrm-email');
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
      await saveAllContact();
      $(this).prop('disabled', false);
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

    async function saveAllContact(event) {
      let toSave = []
      $('.ms-List').children().each(function (index) {
        if ($(this).hasClass('to_save') && $(this).hasClass('is-selected')) {
          let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
          toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email')])
        }
      });


      for (const [key, val] of Object.entries(toSave)) {
        await saveContactToCRM({"name": val[0], "email": val[1]})
      }
    }

    async function confirmSaveAllContactInGroup(event) {
      let toSave = [];
      $('.ms-List').children().each(function (index) {
        let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
        // name,email,already saved
        if ($(this).hasClass('is-selected')) {
          if ($(this).hasClass('to_save')) {
            toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email'), $(this).hasClass('to_save')])
          } else {
            toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email'), $(this).hasClass('to_save'), contact.data('civicrm-id')])
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

    async function showContacts() {

      let html = '<ul class="ms-List">'

      // Sender, then To, Cc and Bcc. A lookup that failed has shown the error
      // dialog; listing its address as unknown would invite a duplicate.
      let addresses = [requiredAttendees.from].concat(requiredAttendees.to, requiredAttendees.cc, requiredAttendees.bcc);
      for (const val of addresses) {
        let res = await checkContact(val)
        if (!res) {
          return;
        }
        html += await getListItem(val, res)
      }
      html += '</ul>'

      html += '<button class="ms-Button ms-Button--small save-contact-all">' +
        '<span class="ms-Button-label">' + UIText.SaveContactInGroupScreen.Save + '</span>' +
        '</button>'
      html += '<button class="ms-Button ms-Button--small save-contact-all-group">' +
        '<span class="ms-Button-label">' + UIText.SaveContactInGroupScreen.Title + '</span>' +
        '</button>'
      html += '<p id="saving-contact-help-text" style="display: none;">' + UIText.SaveContactInGroupScreen.SavingText + '</p>';
      html += '<p id="saved-contact-help-text" style="display: none;">' + UIText.SaveContactInGroupScreen.SavedText + '</p>';

      html += '<br><br>';


      $("#contacts").html(html)
      $(".save-contact").on('click', saveContact)
      $(".save-contact-all-group").on('click', confirmSaveAllContactInGroup)
      $(".save-contact-all").on('click', confirmSaveAllContact)
      $(".save-email").on('click', saveEmail)

    }

    async function saveEmail() {
      var emailBody = null
      await requiredAttendees.body.getAsync(
        "html",
        async function callback(result) {
          emailBody = result['value']

          var targetContactIds = []
          var pendingContacts = []

          $('.ms-List').children().each(function (index) {
             let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
             // name,email,already saved
             if ($(this).hasClass('is-selected')) {
               if ($(this).hasClass('to_save')) {
                 pendingContacts.push(
                   saveContactToCRM({"name": contact.data('civicrm-name'), "email": contact.data('civicrm-email')})
                     .then(contactId => targetContactIds.push(contactId))
                 );
               } else {
                 targetContactIds.push(contact.data('civicrm-id'));
               }
             }
          });

          await Promise.all(pendingContacts);

          if (await pushEmailActivity(requiredAttendees.subject, emailBody, new Date(requiredAttendees.dateTimeCreated), targetContactIds)) {
            addSavedCategoryToCurrentMessage();
          }
        });
    }

    function addSavedCategoryToCurrentMessage() {
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
      let contact = {};
      if (separatorIndex === -1) {
        contact.lastName = 'Unknown';
      } else {
        contact.lastName = name.substring(separatorIndex + 1);
      }
      contact.firstName = name.substring(0, separatorIndex);
      return contact;
    }
  }
})();
