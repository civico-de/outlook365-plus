(function () {
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);
    jQuery(document).ready(async function () {
      $('h1.title').text(UIText.SaveContactInGroupScreen.Title);
      $('.not-configured-warning .ms-MessageBar-text').html(UIText.SettingsScreen.NotConfigured);
      $('#group-select-done .ms-Button-label').text(UIText.SaveContactInGroupScreen.Done);
      $('#group-select-done').on('click', done);
      $('.group-select-dropdown').on('change', function(){
        var value = $('.group-select-dropdown').val();
        if (value == '-1') {
          $('#newGroupName').removeClass('is-disabled');
        } else {
          $('#newGroupName').addClass('is-disabled');
        }
      });
      $('.group-select-dropdown').change();
      $('#group-select-done').prop('disabled', true);
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, addGroups, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          throw new Error('Cannot receive the group list: ' + result.error.message);
        }
        Office.context.ui.messageParent(JSON.stringify({"ready": true}));
      });
    });
  };

  // The task pane answers "ready" with [[id, title], ...].
  function addGroups(arg) {
    var groups = JSON.parse(arg.message);
    groups.forEach(function (group) {
      $('.group-select-dropdown').append($('<option>').val(group[0]));
    });
    // Fabric builds its items from the select once, writing option text as HTML,
    // so the titles go in as text after it has built them.
    new fabric['Dropdown'](document.querySelector('.ms-Dropdown'));
    var options = $('.group-select-dropdown option');
    var items = $('.ms-Dropdown-item');
    groups.forEach(function (group, i) {
      options.eq(i + 1).text(group[1]);
      items.eq(i + 1).text(group[1]);
    });
    // Fabric also copies a picked item's text into the title as HTML.
    Object.defineProperty(document.querySelector('.ms-Dropdown > .ms-Dropdown-title:not(.ms-Dropdown-truncator)'), 'innerHTML', {
      get: function () { return this.textContent; },
      set: function (value) { this.textContent = value; },
    });
    $('#group-select-done').prop('disabled', false);
  }

  function done() {
    let selectedGroup = $(".group-select-dropdown").val()
    let fieldName = $("#civicrm-group-field").val()
    let exist = true
    if (selectedGroup == '-1') {
      exist = false
    }
    Office.context.ui.messageParent(JSON.stringify({
      "exist": exist,
      "selectedGroup": selectedGroup,
      "groupName": fieldName
    }));
  }
})();
