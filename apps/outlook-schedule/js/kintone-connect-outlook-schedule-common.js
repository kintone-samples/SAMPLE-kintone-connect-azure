window.kintoneAzureConnect = {

  config: {
    auth: {
      clientId: 'b689274d-3ed5-429d-aa7a-23a2d446af0e',
      authority: 'https://login.microsoftonline.com/common'
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  },

  graphApiScorp: {
    scopes: ['calendars.readwrite'],
  },

  eventUrl: 'https://graph.microsoft.com/v1.0/me/events',

  kintone: {
    fieldCode: {

      // Field code of subject
      subject: 'To_Do',

      // Field code of body
      body: 'Details',

      // Field code of start
      startDate: 'From',

      // Field code of end
      endDate: 'To',

      // Field code of eventId
      eventId: 'EventId',
      
      // Field code of attachFile
      attachFile: 'Attachments'
    }
  }
};


