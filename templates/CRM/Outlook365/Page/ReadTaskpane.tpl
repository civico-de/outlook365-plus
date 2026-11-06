{crmScope extensionKey='outlook365'}
<!DOCTYPE html>
<html lang="{$config->lcMessages|default:'en_US'|replace:'_':'-'}">

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CiviCRM</title>
  <script type="text/javascript">
    var crmBaseUrl = '{$site_base_url|escape:'javascript'}';
    var confirmDialogUrl = '{$baseurl}outlook365/settings/confirm.html';
    var settingsDialogUrl = '{$baseurl}outlook365/settings/dialog.html';
    var saveContactDialogUrl = '{$baseurl}outlook365/settings/saveContact.html';
    var saveContactInGroupDialogUrl = '{$baseurl}outlook365/settings/saveContactInGroup.html';
  </script>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
  <script type="text/javascript" src="{$baseurl}assets/jquery-3.4.1.min.js"></script>

  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />
  <script type="text/javascript" src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

  <link rel="stylesheet" href="{$baseurl}assets/taskpane.css"/>
</head>

<body class="ms-font-m ms-Fabric ms-landing-page">
<main class="ms-landing-page__main">
  <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
    <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning" style="display: none;">
      <div class="ms-font-l" id="settings-prompt"></div>
    </div>

    <div class="ms-CommandBar-mainArea">
      <div class="ms-CommandButton ms-CommandButton--pivot is-active">
          <!-- The pane shows one list and switches nothing, so this heads the list rather than acting as a tab. -->
          <h1 class="pivot-tab contacts"><i class="ms-Icon ms-Icon--Contact" aria-hidden="true"></i><span class="ms-font-l">{ts}Contacts{/ts}</span></h1>
      </div>
      <div class="ms-CommandButton ms-CommandButton--pivot" id="settings-icon">
          <button type="button" class="pivot-tab settings" aria-haspopup="dialog">
            <i class="ms-Icon ms-Icon--Settings" aria-hidden="true"></i><span class="ms-font-l">{ts}Settings{/ts}</span>
          </button>
      </div>
      </div>
    <div class="dataclass">
      <br>
    </div>
    <div>
      <button class="ms-Button ms-Button--small selectAll">
        <span class="ms-Button-label">{ts}Select All{/ts}</span>
      </button>
      <button class="ms-Button ms-Button--small unselectAll">
        <span class="ms-Button-label">{ts}Unselect All{/ts}</span>
      </button>
    </div>

    <div class="dataclass" id="contacts">
    </div>
    <div class="ms-Spinner" id="loadingContacts">
      <div class="ms-Spinner-label">
        <br>
        {ts}Loading...{/ts}
      </div>
    </div>
    <hr>

    <div class="ms-CommandBar-mainArea">
      <h2>{ts}Save Email{/ts}</h2>
      <p>{ts}Click the following button to save this email in CiviCRM as an Activity{/ts}</p>
      <button class="ms-Button ms-Button--medium save-email" disabled aria-disabled="true">
        <span class="ms-Button-label">{ts}Save Email{/ts}</span>
      </button>
      <div role="status">
        <p id="saved-email-help-text" hidden>{ts}This email is saved in CiviCRM.{/ts}</p>
      </div>
      </div>

  </section>
</main>
<!-- <footer class="ms-landing-page__footer ms-bgColor-neutralLighter ms-bgColor-neutralLight--hover">
  <div id="settings-icon" class="ms-landing-page__footer--left ms-bgColor-neutralLight--hover ms-fontColor-neutralDark ms-fontColor-neutralDarker--hover" aria-label="Settings" tabindex=0>
    <i class="ms-Icon enlarge ms-Icon--Settings "></i><span class="label"></span>
  </div>
</footer> -->

<div class="civicrm-notice">
  <div class="ms-Dialog">
    <div class="ms-Dialog-title">CiviCRM</div>
    <div class="ms-Dialog-content" id="civicrm-notice-div">
      <p class="ms-Dialog-subText" id="civicrm-notice-text">{ts}An error occurred communicating with CiviCRM. This action could not be completed{/ts}</p>
    </div>
    <div class="ms-Dialog-actions">
      <button class="ms-Button ms-Dialog-action ms-Button--primary">
        <span class="ms-Button-label">OK</span>
      </button>
    </div>
  </div>
</div>
<script type="text/javascript">
  {literal}
  var noticeDialog = document.querySelector(".ms-Dialog");
  var dialogComponent = new fabric['Dialog'](noticeDialog);
  function openDialog(dialog) {
    // Open the dialog
    dialog.open();
  }
  {/literal}
</script>

<script>
  var CRMContactURL = "{$contactURL}";
  {literal}
  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
  var SearchBoxElements = document.querySelectorAll(".ms-SearchBox");
  for (var i = 0; i < SearchBoxElements.length; i++) {
    new fabric['SearchBox'](SearchBoxElements[i]);
  }
  {/literal}
</script>
<script type="text/javascript" src="{$baseurl}assets/UIStrings.js"></script>
<script type="text/javascript" src="{$baseurl}assets/civicrm.js"></script>
<script type="text/javascript" src="{$baseurl}assets/readtaskpane.js"></script>
</body>

</html>
{/crmScope}
