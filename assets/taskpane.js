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
      // Fabric's search box ships its clear button without a label of any kind.
      $('.ms-SearchBox-clear .ms-CommandButton-button').attr('aria-label', UIText.Clear);
      $('#searchField').attr('aria-label', UIText.Search);
      $('#settings-icon .label').text(UIText.Settings);

      reset();
      if (config) {
        loadNextContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);
      $('[role="tab"]').on('click', handleTabChange);
    });

    // The row's checkbox holds the selection; the class only mirrors it for Fabric's styling.
    function setSelected(box, selected) {
      $(box).prop('checked', selected).closest('.ms-ListItem').toggleClass('is-selected', selected);
    }

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

    $("body").delegate(".selectAll", "click", function (event) {
      let parentElement = $(event.target).closest(".allData").parent()
      setSelected(parentElement.children("ul").children(".ms-ListItem").children(".ms-ListItem-selectionTarget"), true)
    });

    $("body").delegate(".unselectAll", "click", function (event) {
      let parentElement = $(event.target).closest(".allData").parent()
      setSelected(parentElement.children("ul").children(".ms-ListItem").children(".ms-ListItem-selectionTarget"), false)
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
    function handleTabChange() {
      // The tab names its panel in aria-controls; that id is also the tab's own name.
      currentTab = $(this).attr('aria-controls')
      $('[role="tab"]').attr('aria-selected', 'false')
      $(this).attr('aria-selected', 'true')
      $(".ms-CommandButton.ms-CommandButton--pivot").removeClass("is-active")
      $(this).closest(".ms-CommandButton").addClass("is-active")
      $(".dataclass").empty()
      reset()
      if (currentTab === "contacts") {
        $("#search-form").show();
        $('#search-form span.ms-SearchBox-text').text(UIText.ContactScreen.Placeholder);
        $('#searchField').attr('aria-label', UIText.ContactScreen.Placeholder);
        loadNextContacts()
      } else if (currentTab === "groups") {
        $("#search-form").show();
        $('#search-form span.ms-SearchBox-text').text(UIText.GroupScreen.Placeholder);
        $('#searchField').attr('aria-label', UIText.GroupScreen.Placeholder);
        loadNextGroups()
      }

    }

    function getSearchForm(id) {
      let html = '<div id="group_search_' + id + '">' +
        '<div class="ms-SearchBox ms-SearchBox--commandBar search-in-group">' +
        '<input class="ms-SearchBox-field" id="groupSearchField_' + id + '" type="text" value="">' +
        '<label class="ms-SearchBox-label" for="groupSearchField_' + id + '">' +
        '<i class="ms-SearchBox-icon ms-Icon ms-Icon--Search" aria-hidden="true"></i>' +
        '<span class="ms-SearchBox-text">' + UIText.ContactScreen.Placeholder + '</span>' +
        '</label>' +
        '<div class="ms-CommandButton ms-SearchBox-clear ms-CommandButton--noLabel">' +
        '<button class="ms-CommandButton-button" aria-label="' + UIText.Clear + '">' +
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

          // Chevron and title are one control, so the keyboard reaches the group the mouse opens.
          var toggle = $('<button type="button" class="group-toggle" aria-expanded="false">')
            .attr('id', idname + '-expand-groups')
            .append('<i class="ms-Icon ms-Icon--ChevronRight" style="padding: 6px" aria-hidden="true"></i>')
            .append($('<span class="group-name">').text(name));
          var header = $('<div class="ms-Persona-primaryText CiviCRM-Group-Email">')
            .attr({'data-civicrm-id': id, 'data-civicrm-name': name})
            .append(toggle)
            .append($('<div class="ms-ListItem-actions">').append($('<div class="ms-ListItem-action">').append(
              $('<a target="_blank">').attr({'href': groupURL, 'title': UIText.GroupScreen.TitleURL}).append(
                $('<i class="ms-Icon ms-Icon--Settings" aria-hidden="true">')
              )
            )));
          $('#groups').append(
            $('<div class="ms-Persona">').append($('<div class="ms-Persona-details">').append(header))
          );
          // The offset the next page starts at; without it every scroll repeats the first page.
          currentOffset++;
          $('#' + idname + '-expand-groups').on('click', expandGroups);
        }

        if (data.count < 25) {
          moreAvailable = false;
        }

      }
    }

    async function expandGroups() {
      // The event target may be the toggle's icon or its label, so the header is read from the button.
      let header = $(this).parent('.CiviCRM-Group-Email');
      let chevron = $(this).children('i');
      if (header.hasClass("expanded")) {
        let id = header.data('civicrm-id');
        header.removeClass('expanded')
        header.children(".allData").remove()
        header.children(".CiviCRM-Group-Email").remove()
        header.children("#group_search_" + id).remove()
        header.children("ul").remove()
        chevron.attr('class', 'ms-Icon ms-Icon--ChevronRight');
        $(this).attr('aria-expanded', 'false');
        return;
      }
      header.addClass('expanded')
      let name = header.attr('data-civicrm-name').replace(" ", "-").toLowerCase();
      let id = header.data('civicrm-id');
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
      header.append(
        $('<div class="CiviCRM-Group-Email">').attr({'data-civicrm-id': id, 'data-civicrm-name': name}).append(buttons)
      );
      header.append('<div class="allData"><button class="ms-Button ms-Button--small selectAll"><span class="ms-Button-label">' + UIText.GroupScreen.SelectAll + '</span></button>' +
        '<button class="ms-Button ms-Button--small unselectAll"><span class="ms-Button-label">' + UIText.GroupScreen.UnselectAll + '</span></button></div>')
      header.append(getSearchForm(id))
      $("#group_search_" + id + " .ms-SearchBox-text").text(UIText.GroupScreen.SearchContact);
      $("#groupSearchField_" + id).attr('aria-label', UIText.GroupScreen.SearchContact);

      header.append($('<ul class="ms-List">').addClass(id + '-list-email').append(members));
      var SearchBoxElements = document.querySelectorAll(".ms-SearchBox");
      for (var i = 0; i < SearchBoxElements.length; i++) {
        new fabric['SearchBox'](SearchBoxElements[i]);
      }


      // to get smoothness change class after query
      chevron.attr('class', 'ms-Icon ms-Icon--ChevronDown');
      $(this).attr('aria-expanded', 'true');

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

      $("#groupSearchField_" + id).on("keypress", async function (e) {
        if (e.keyCode == 13) {
          let groupSearchVal = $(this).val();
          console.log(groupSearchVal)
          if (groupSearchVal != '') {
            params.display_name = {"LIKE": '%' + groupSearchVal + '%'};
          }
          let found = await getGroupContacts(params)
          $("." + id + '-list-email').empty().append(found)
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
            .append($('<input type="checkbox" class="ms-ListItem-selectionTarget">').attr('aria-label', contact.display_name))
            .append($('<span class="ms-ListItem-primaryText">').text(contact.display_name))
            .append($('<div class="ms-ListItem-actions">').append($('<div class="ms-ListItem-action">').append(
              $('<a target="_blank">').attr({'href': contact_url, 'title': UIText.ContactScreen.TitleURL}).append(
                $('<i class="ms-Icon ms-Icon--Contact" aria-hidden="true">')
              )
            ))));
        }
      }).catch(showError);
      return items;
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
