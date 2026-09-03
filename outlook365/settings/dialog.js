(function () {
  'use strict';

  // The token itself never reaches this dialog, only whether one is stored,
  // so an empty field means "keep the stored token".
  var hasToken = false;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);
    var accessToken;

    jQuery(document).ready(function () {
      $('h1.title').text(UIText.SettingsScreen.Title);
      $('.not-configured-warning .ms-MessageBar-text').html(UIText.SettingsScreen.NotConfigured);
      $('.ms-Label.url').text(UIText.SettingsScreen.URL);
      $('.ms-Label.token').text(UIText.SettingsScreen.Token);
      $('#settings-done .ms-Button-label').text(UIText.SettingsScreen.Done);
      $('#civicrm-url').attr("placeholder", UIText.SettingsScreen.URL_Placeholder);


      $('#civicrm-url').change(change);
      $('#civicrm-token').change(change);

      // Check if warning should be displayed.
      var config = JSON.parse(getParameterByName('config'));
      if (!config) {
        $('.not-configured-warning').show();
      } else {
        $('#civicrm-url').val(config.url);
        hasToken = !!config.hasToken;
        if (hasToken) {
          $('#civicrm-token').attr('placeholder', UIText.SettingsScreen.TokenKeep);
        }
      }

      change();

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', done);
    });
  };

  function readSettings() {
    return {
      url: $('#civicrm-url').val(),
      token: $('#civicrm-token').val()
    };
  }

  function change() {
    var settings = readSettings();
    $('#settings-done').prop('disabled', !(settings.url && (settings.token || hasToken)));
  }

  function done() {
    Office.context.ui.messageParent(JSON.stringify(readSettings()));
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
