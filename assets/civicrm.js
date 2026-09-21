/* Transport for every CiviCRM call the add-in makes.
   The credential travels in a request header, never in a URL or a body. */

window.CiviCRM = (function () {
  "use strict";

  function request(config, url, data) {
    // jQuery is loaded by every page that loads this file.
    return window.jQuery.ajax({
      url: url,
      type: "POST",
      dataType: "json",
      headers: { "X-Civi-Auth": "Bearer " + config.token },
      data: data,
    });
  }

  /**
   * APIv3 through the legacy AJAX endpoint.
   *
   * @param config - the add-in configuration (url, token)
   * @param entity - APIv3 entity name
   * @param action - APIv3 action name
   * @param params - the APIv3 params object
   */
  function api3(config, entity, action, params) {
    // APIv3 reports errors as HTTP 200 with is_error; reject those like a failed request.
    return request(config, config.url + "/civicrm/ajax/rest", {
      entity: entity,
      action: action,
      json: JSON.stringify(params),
    }).then(function (result) {
      return result.is_error ? window.jQuery.Deferred().reject(result) : result;
    });
  }

  /**
   * APIv4 through the AJAX endpoint.
   *
   * @param config - the add-in configuration (url, token)
   * @param entity - APIv4 entity name
   * @param action - APIv4 action name
   * @param params - the APIv4 params object (select/where/values/etc.)
   */
  function api4(config, entity, action, params) {
    const url =
      config.url +
      "/civicrm/ajax/api4/" +
      encodeURIComponent(entity) +
      "/" +
      encodeURIComponent(action);
    return request(config, url, { params: JSON.stringify(params) });
  }

  return { api3: api3, api4: api4 };
})();
