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
      addGroups(JSON.parse(getParameterByName('groups')) || []);


      var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
      for (var i = 0; i < DropdownHTMLElements.length; ++i) {
        var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
      }

      // Check if warning should be displayed.
      // var contact = JSON.parse(getParameterByName('config'));
      // console.log(contact)
      // if (contact) {
      //   $('#civicrm-name').val(contact.name);
      //   $('#civicrm-email').val(contact.email);
      // }
      // change();

      // // When the Done button is selected, send the
      // // values back to the caller as a serialized
      // // object.
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
    });
  };

  // [[id, title], ...] as the task pane passed it.
  function addGroups(groups) {
    groups.forEach(function (group) {
      $('.group-select-dropdown').append($('<option>').val(group[0]).text(group[1]));
    });
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

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
