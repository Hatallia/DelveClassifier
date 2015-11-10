(function () {
  'use strict';

  var officeAddin = angular.module('officeAddin');

  officeAddin.config(['$httpProvider', 'adalAuthenticationServiceProvider', 'appId', 'sharePointUrl', adalConfigurator]);

  function adalConfigurator($httpProvider, adalProvider, appId, sharePointUrl) {
    var adalConfig = {
      tenant: 'common',
      clientId: appId,
      extraQueryParameter: 'nux=1',
      endpoints: {}
      // cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost. 
    };
    adalConfig.endpoints[sharePointUrl + '/_api/'] = sharePointUrl;
    adalProvider.init(adalConfig, $httpProvider);
  }
})();