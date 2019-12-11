/* global Msal */
(function() {
  'use strict';

  var KAC = window.kintoneAzureConnect;

  var azureOauth = {

    userAgentApplication: null,

    getUserInfo: function() {
      return this.userAgentApplication.getAccount().userName;
    },

    init: function() {
      var self = this;
      self.userAgentApplication = new Msal.UserAgentApplication(KAC.config);
    },

    signIn: function() {
      var self = this;
      return self.userAgentApplication.loginPopup(KAC.graphApiScorp);
    },

    signOut: function() {
      var self = this;
      self.userAgentApplication.logout();
    },

    callGraphApi: function() {
      var self = this;
      return self.userAgentApplication.acquireTokenSilent(KAC.graphApiScorp);
    }
  };

  window.azureOauth = azureOauth || {};
}());
