// ck-allow-api3 -- upstream's APIv3 call sites, routed through assets/civicrm.js; migrating them belongs upstream
(function () {
  'use strict';
  Office.initialize = function (reason) {
    var item = Office.context.mailbox.item;
    var currentOffset = 0;
    var currentTab = "contacts"
    var moreAvailable = true;
    var search = false;
    var config = false;

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);

    var fields = [];

    // The site URL identifies the site a token belongs to, so both sides of
    // that comparison have to be spelled the same way.
    function normalizeUrl(url) {
      return String(url || '').replace(/\/+$/, '');
    }

    jQuery(document).ready(function () {
      // Set localized text for UI elements.
      $("#loadingContacts .ms-Spinner-label").text(UIText.Loading);
      $("#searchField .ms-SearchBox-text").text(UIText.Search);
      $('#settings-prompt').html(UIText.NotConfigured);
      $('#settings-icon .label').text(UIText.Settings);

      reset();
      if (config) {
        loadNextContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);
      $('.ms-CommandButton--pivot span').on('click', handleTabChange);
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

    $("body").delegate(".selectAll", "click", function (event) {
      let parentElement = $(event.target).closest(".allData").parent()
      parentElement.children("ul").children(".ms-ListItem").addClass("is-selected")
    });

    $("body").delegate(".unselectAll", "click", function (event) {
      let parentElement = $(event.target).closest(".allData").parent()
      parentElement.children("ul").children(".ms-ListItem").removeClass("is-selected")
    });

    $('#searchField').on("keypress", function (e) {
      if (e.keyCode == 13) {
        reset();
        search = $(this).val();
        if (currentTab === "contacts") {
          loadNextContacts();
        } else if (currentTab === "groups") {
          loadNextGroups();
        }
        return false; // prevent the button click from happening
      }
    });

    $(window).scroll(function () {
      if ($(window).scrollTop() == $(document).height() - $(window).height()) {
        // ajax call get data from server and append to the div
        if (currentTab === "contacts") {
          loadNextContacts();
        } else if (currentTab === "groups") {
          loadNextGroups();
        }
      }
    });

    /**
     * Handle Tab Change
     */
    function handleTabChange(event) {
      let classes = event.target.className.split(" ")
      let targetTab = classes[classes.length - 1]
      if (targetTab === "settings") {
        return
      }
      currentTab = targetTab
      let parentselector = $($(event.target).parent()).parent()
      $(".ms-CommandButton.ms-CommandButton--pivot").removeClass("is-active")
      parentselector.addClass("is-active")
      $(".dataclass").empty()
      reset()
      if (currentTab === "contacts") {
        $("#search-form").show();
        $('#search-form span.ms-SearchBox-text').text(UIText.ContactScreen.Placeholder);
        loadNextContacts()
      } else if (currentTab === "groups") {
        $("#search-form").show();
        $('#search-form span.ms-SearchBox-text').text(UIText.GroupScreen.Placeholder);
        loadNextGroups()
      }

    }

    function getSearchForm(id) {
      let html = '<div id="group_search_' + id + '">' +
        '<div class="ms-SearchBox ms-SearchBox--commandBar search-in-group">' +
        '<input class="ms-SearchBox-field" id="groupSearchField" type="text" value="">' +
        '<label class="ms-SearchBox-label">' +
        '<i class="ms-SearchBox-icon ms-Icon ms-Icon--Search"></i>' +
        '<span class="ms-SearchBox-text">' + UIText.ContactScreen.Placeholder + '</span>' +
        '</label>' +
        '<div class="ms-CommandButton ms-SearchBox-clear ms-CommandButton--noLabel">' +
        '<button class="ms-CommandButton-button">' +
        '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Clear"></i></span>' +
        '<span class="ms-CommandButton-label"></span>' +
        '</button>' +
        '</div>' +
        '<div class="ms-CommandButton ms-SearchBox-exit ms-CommandButton--noLabel">' +
        '<button class="ms-CommandButton-button">' +
        '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--ChromeBack"></i></span>' +
        '<span class="ms-CommandButton-label"></span>' +
        '</button>' +
        '</div>' +
        '<div class="ms-CommandButton ms-SearchBox-filter ms-CommandButton--noLabel">' +
        '<button class="ms-CommandButton-button">' +
        '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Filter"></i></span>' +
        '<span class="ms-CommandButton-label"></span>' +
        '</button>' +
        '</div>' +
        '</div>' +
        '</div>'

      return html
    }

    /**
     * Reset the contact list
     */
    function reset() {
      config = getConfig();
      if (!config) {
        $('.not-configured-warning').show();
        $('#search-form').hide();
      } else {
        $('.not-configured-warning').hide();
        if (currentTab !== "settings") {
          $('#search-form').show();
        }
      }
      $('#loadingContacts').hide();
      $('#civicrm-error').hide();
      $('#contacts').html('');
      $('#groups').html('');
      currentOffset = 0;
      moreAvailable = true;
      search = false;
      fields = [];
      if (currentTab === "contacts") {
        getContactFields();
      }
    }

    /**
     * Retrieve next batch of contacts.
     */
    function getContactFields() {
      if (!config) {
        return;
      }
      $('#loadingContacts').show();
      CiviCRM.api3(config, 'Outlook365Contact', 'getfields', {
        "sequential": 1,
        "api_action": "get"
      }).done(function (result) {
        fields = [];
        if (!result.is_error) {
          for (var i in result.values) {
            var field = result.values[i];
            if (field.name != 'display_name') {
              fields.push({
                "name": field.name,
                "title": field.title
              });
            }
          }
        }
      }).fail(showError);
    }

    /**
     * A failed call shows an error; an empty list would read as "no data".
     */
    function showError() {
      $('#civicrm-error').show();
    }

    /**
     * Retrieve next batch of Groups.
     */
    function loadNextGroups() {

      if (!moreAvailable || !config) {
        return;
      }
      var params = {
        "sequential": 1,
        "return": ["id", "name", "title"],
        "options": {
          "offset": currentOffset,
          "limit": 25,
        }
      };
      if (search) {
        params.title = {"LIKE": '%' + search + '%'};
      }
      CiviCRM.api3(config, 'Group', 'get', params).done(addGroups).fail(showError);
    }

    /**
     * Add the group to the list.
     *
     * @param data
     */
    function addGroups(data) {

      if (data.is_error == 0) {
        console.log(data);
        for (var i in data.values) {
          var group = data.values[i];
          var idname = 'group_' + group.id.toString();
          var name = group.title;
          var id = group.id;
          var groupURL = CRMGroupURL + '&id=' + id;

          var header = $('<div class="ms-Persona-primaryText CiviCRM-Group-Email">')
            .attr({'data-civicrm-id': id, 'data-civicrm-name': name})
            .append($('<i class="ms-Icon ms-Icon--ChevronRight" style="padding: 6px"></i>').attr('id', idname + '-expand-groups'))
            .append($('<span class="group-name">').text(name))
            .append($('<div class="ms-ListItem-actions">').append($('<div class="ms-ListItem-action">').append(
              $('<a target="_blank">').attr('href', groupURL).append(
                $('<i class="ms-Icon ms-Icon--Settings">').attr('title', UIText.GroupScreen.TitleURL)
              )
            )));
          $('#groups').append(
            $('<div class="ms-Persona">').append($('<div class="ms-Persona-details">').append(header))
          );
          $('#' + idname + '-expand-groups').on('click', expandGroups);
        }
        $('.group-name').on('click', event => {
          $(event.target).siblings('i').trigger('click');
        });

        if (data.count < 25) {
          moreAvailable = false;
        }

      }
    }

    async function expandGroups(event) {
      if ($(this).parent().hasClass("expanded")) {
        let id = $(this).parent('.CiviCRM-Group-Email').data('civicrm-id');
        $(this).parent().removeClass('expanded')
        $(this).parent().children(".allData").remove()
        $(this).parent().children(".CiviCRM-Group-Email").remove()
        $(this).parent().children("#group_search_" + id).remove()
        $(this).parent().children("ul").remove()
        $(this).attr('class', 'ms-Icon ms-Icon--ChevronRight');
        return;
      }
      $(this).parent().addClass('expanded')
      let name = $(this).parent('.CiviCRM-Group-Email').attr('data-civicrm-name').replace(" ", "-").toLowerCase();
      let id = $(this).parent('.CiviCRM-Group-Email').data('civicrm-id');
      // $(this).parent().append('<div id="' + name + '-expanded-groups"></div>');
      var params = {
        "sequential": 1,
        "options": {
          "limit": 0,
        },
        "group_id": id,
      };
      let members = await getGroupContacts(params)

      var buttons = '<button class="ms-Button ms-Button--small to"><span class="ms-Button-label">' + UIText.To + '</span></button>' +
        '<button class="ms-Button ms-Button--small cc"><span class="ms-Button-label">' + UIText.Cc + '</span></button>' +
        '<button class="ms-Button ms-Button--small bcc"><span class="ms-Button-label">' + UIText.Bcc + '</span></button>';
      $(event.target).parent().append(
        $('<div class="CiviCRM-Group-Email">').attr({'data-civicrm-id': id, 'data-civicrm-name': name}).append(buttons)
      );
      $(event.target).parent().append('<div class="allData"><button class="ms-Button ms-Button--small selectAll"><span class="ms-Button-label">' + UIText.GroupScreen.SelectAll + '</span></button>' +
        '<button class="ms-Button ms-Button--small"><span class="ms-Button-label unselectAll">' + UIText.GroupScreen.UnselectAll + '</span></button></div>')
      $(event.target).parent().append(getSearchForm(id))
      $("#group_search_" + id + " .ms-SearchBox-text").text(UIText.GroupScreen.SearchContact);

      $(event.target).parent().append($('<ul class="ms-List">').addClass(id + '-list-email').append(members));
      initListItems(members);
      var SearchBoxElements = document.querySelectorAll(".ms-SearchBox");
      for (var i = 0; i < SearchBoxElements.length; i++) {
        new fabric['SearchBox'](SearchBoxElements[i]);
      }


      // to get smoothness change class after query
      $(this).attr('class', 'ms-Icon ms-Icon--ChevronDown');

      $("#groups .ms-Button.to").click(function (event) {
        let toSend = []
        var recipients = item.to
        $('.' + id + '-list-email').children().each(function (index) {
          if ($(this).hasClass('is-selected')) {
            let contact = $(this)
            toSend.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email')])
          }

        });
        for (var iter in toSend) {
          addReiever(recipients, toSend[iter][1], toSend[iter][0]);
        }
      });

      $("#groups .ms-Button.cc").click(function (event) {
        let toSend = []
        var recipients = item.cc
        $('.' + id + '-list-email').children().each(function (index) {
          if ($(this).hasClass('is-selected')) {
            let contact = $(this)
            toSend.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email')])
          }

        });
        for (var iter in toSend) {
          addReiever(recipients, toSend[iter][1], toSend[iter][0]);
        }
      });

      $("#groups .ms-Button.bcc").click(function (event) {
        let toSend = []
        var recipients = item.bcc
        $('.' + id + '-list-email').children().each(function (index) {
          if ($(this).hasClass('is-selected')) {
            let contact = $(this)
            toSend.push([contact.attr('data-civicrm-name'), contact.attr('data-civicrm-email')])
          }

        });
        for (var iter in toSend) {
          addReiever(recipients, toSend[iter][1], toSend[iter][0]);
        }
      });

      $("#group_search_" + id + " #groupSearchField").on("keypress", async function (e) {
        if (e.keyCode == 13) {
          let groupSearchVal = $(this).val();
          console.log(groupSearchVal)
          if (groupSearchVal != '') {
            params.display_name = {"LIKE": '%' + groupSearchVal + '%'};
          }
          let found = await getGroupContacts(params)
          $("." + id + '-list-email').empty().append(found)
          initListItems(found);
          return false; // prevent the button click from happening
        }
      });
    }

    /**
     * The group's members as list items; names and addresses go in as text.
     */
    async function getGroupContacts(params) {
      var items = [];
      await CiviCRM.api3(config, 'Outlook365Group', 'get', params).done(function (data) {
        for (var i in data.values) {
          var contact = data.values[i];
          var contact_url = CRMContactURL;
          contact_url += "&cid=" + String(contact.contact_id);
          items.push($('<li class="ms-ListItem is-selectable">')
            .attr({'data-civicrm-name': contact.display_name, 'data-civicrm-email': contact.email})
            .append($('<span class="ms-ListItem-primaryText">').text(contact.display_name))
            .append('<div class="ms-ListItem-selectionTarget"></div>')
            .append($('<div class="ms-ListItem-actions">').append($('<div class="ms-ListItem-action">').append(
              $('<a target="_blank">').attr('href', contact_url).append(
                $('<i class="ms-Icon ms-Icon--Contact">').attr('title', UIText.ContactScreen.TitleURL)
              )
            ))));
        }
      }).catch(showError);
      return items;
    }

    function initListItems(items) {
      items.forEach(function (item) {
        new fabric['ListItem'](item[0]);
      });
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
      }).fail(response => {
        openDialog(dialogComponent);
      });
      if (exist) {
        await CiviCRM.api3(config, 'Contact', 'getsingle', {"id": contact_id}).done(function (result) {
          contact_name = result["display_name"]
        }).fail(response => {
          openDialog(dialogComponent);
        });
      }

      return {"exist": exist, "contact_url": contact_url, "contact_name": contact_name, "contact_id": contact_id}
    }

    /**
     * Retrieve next batch of contacts.
     */
    function loadNextContacts() {
      if (!moreAvailable || !config) {
        return;
      }
      $('#loadingContacts').show();
      var params = {
        "sequential": 1,
        "options": {
          "offset": currentOffset,
          "limit": 25,
        }
      };
      if (search) {
        params.display_name = {"LIKE": '%' + search + '%'};
      }
      CiviCRM.api3(config, 'Outlook365Contact', 'get', params).done(addContacts).fail(showError);
    }

    /**
     * Add the contacts to the list.
     *
     * @param data
     */
    function addContacts(data) {
      $('#loadingContacts').hide();
      if (data.is_error == 0) {
        for (var i in data.values) {
          currentOffset++;
          var contact = data.values[i];
          var details = $('<div class="ms-Persona-details">')
            .append($('<div class="ms-Persona-primaryText">').text(contact.display_name))
            .append($('<div class="ms-ListItem-actions search-contact-profile">').append($('<div class="ms-ListItem-action">').append(
              $('<a target="_blank">').attr('href', CRMContactURL + "&cid=" + String(contact.contact_id)).append(
                $('<i class="ms-Icon ms-Icon--Contact">').attr('title', UIText.ContactScreen.TitleURL)
              )
            )))
            .append('<div class="ms-Persona-secondaryText"></div>');
          for (var fieldI in fields) {
            var value = contact[fields[fieldI].name];
            if (value) {
              details.append($('<div class="ms-Persona-secondaryText">')
                .append($('<strong>').text(fields[fieldI].title + ':'), '&nbsp;', document.createTextNode(value)));
            }
          }
          if (contact.email) {
            details.append($('<div class="CiviCRM-Email">')
              .attr({'data-civicrm-name': contact.display_name, 'data-civicrm-email': contact.email})
              .append(
                '<button class="ms-Button ms-Button--small to"><span class="ms-Button-label">' + UIText.To + '</span></button>' +
                '<button class="ms-Button ms-Button--small cc"><span class="ms-Button-label">' + UIText.Cc + '</span></button>' +
                '<button class="ms-Button ms-Button--small bcc"><span class="ms-Button-label">' + UIText.Bcc + '</span></button>'
              ));
          }
          $('#contacts').append($('<div class="ms-Persona">').append(details));
        }

        $("#contacts .ms-Button.to").click(function () {
          var email = $(this).parent('.CiviCRM-Email').attr('data-civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').attr('data-civicrm-name');
          var recipients = item.to;
          addReiever(recipients, email, name);
        });
        $("#contacts .ms-Button.cc").click(function () {
          var email = $(this).parent('.CiviCRM-Email').attr('data-civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').attr('data-civicrm-name');
          var recipients = item.cc;
          addReiever(recipients, email, name);
        });
        $("#contacts .ms-Button.bcc").click(function () {
          var email = $(this).parent('.CiviCRM-Email').attr('data-civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').attr('data-civicrm-name');
          var recipients = item.bcc;
          addReiever(recipients, email, name);
        });
      }
      if (data.count < 25) {
        moreAvailable = false;
      }
    }

    function addReiever(recipients, email, name) {
      var data = [
        {
          "displayName": name,
          "emailAddress": email
        }
      ];
      // Use asynchronous method getAsync to get each type of recipients
      // of the composed item. Each time, this example passes an anonymous
      // callback function that doesn't take any parameters.
      recipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          write(asyncResult.error.message);
        } else {
          // Async call to get to-recipients of the item completed.
          // Display the email addresses of the to-recipients.
          var emailAlreadyExists = false;
          for (var i = 0; i < asyncResult.value.length; i++) {
            if (asyncResult.value[i].emailAddress == email) {
              emailAlreadyExists = true;
              break;
            }
          }
          if (!emailAlreadyExists) {
            recipients.addAsync(data, function (asyncResult) {
            });
          }
        }
      });
    }

    /**
     * Function to open the settings dialog.
     */
    function openSettingsDialog() {
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
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (message) {
          var settings = JSON.parse(message.message);
          var sameSite = previous && normalizeUrl(settings.url) === previous.url;
          config = {
            url: normalizeUrl(settings.url),
            token: settings.token || (sameSite ? previous.token : '')
          };
          setConfig(config, function (result) {
            settingsDialog.close();
            settingsDialog = null;
            reset();
            if (currentTab === "contacts") {
              loadNextContacts();
            } else if (currentTab === "groups") {
              loadNextGroups();
            }
          });
        });
      });
    }

    /**
     * Load the configuration.
     *
     * @returns {null}
     */
    function getConfig() {
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
    function setConfig(config, callback) {
      Office.context.roamingSettings.set('civicrm_url', config.url);
      Office.context.roamingSettings.set('civicrm_token', config.token);

      Office.context.roamingSettings.saveAsync(callback);
    }

  };

})();
