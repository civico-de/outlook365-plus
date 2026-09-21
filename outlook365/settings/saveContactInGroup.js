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
    JSON.parse(arg.message).forEach(function (group) {
      $('.group-select-dropdown').append($('<option>').val(group[0]).text(group[1]));
    });
    // Fabric builds its option list from the select once, so only now.
    new fabric['Dropdown'](document.querySelector('.ms-Dropdown'));
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
