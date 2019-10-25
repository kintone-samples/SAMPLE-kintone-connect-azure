# Authentication with kintone from Azure AD

Use Microsoft's authentication library (msal.js) to enable access to Microsoft Cloud service resources from kintone.

## Description
With Microsoft's authentication library, you can obtain the access token from the Azure Active Directory,
From kintone you can access the data in the Microsoft cloud using the Microsoft Graph API.
<br>ex) Send and receive Outlook mail from kintone and data linkage between kintone and Outlook schedule.

## Usage
1. Create kintone application

2. Register the application to Azure AD from the following site.  
   https://portal.azure.com/

>* 1. Select "Azure Active Directory"
>* 2. Select "App registrations"
>* 3. Configure App information. Redirect URI is "Web" and "https://{subdomain}.cybozu.com/k/{app ID}"
>* 4. In the Application page, select "Authentication", configure Advanced settings. Logout URL is also "https://{subdomain}.cybozu.com/k/{app ID}" and Select "Access Tokens" and "ID Tokens" on Implicit grant 

3. Download kintone-ui-component and kintone-js-sdk from releases.  

 >* kintone-ui-component v0.4.2: https://github.com/kintone/kintone-ui-component/releases/tag/v0.4.2
 >* kintone-js-sdk v0.7.0: https://github.com/kintone/kintone-js-sdk/releases/tag/v0.7.0

4. Set common file according to kintone environment

***In case of cooperation with Outlook Mail***

```javascript
window.kintoneAzureConnect = {

  config: {
    auth: {
      clientId: '00074e52-413b-4ab9-9698-b268f4693e68',
      authority: 'https://login.microsoftonline.com/common'
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  },
  
  graphApiScorp: {
    scopes: ['mail.read', 'mail.send']
  },

  mail: {
    mailGetUrl: 'https://graph.microsoft.com/v1.0/me/messages?$top=100',
    mailSendUrl: 'https://graph.microsoft.com/v1.0/me/sendmail'
  },

  kintone: {
    fieldCode: {

      // Field code of subject
      subject: 'subject',

      // Field code of content
      content: 'contents',

      // Field code of from
      from: 'from',

      // Field code of to
      to: 'TO',

      // Field code of cc
      cc: 'CC',

      // Field code of bcc
      bcc: 'BCC',

      // Field code of messageId
      messageId: 'messageId',

      // Field code of mailAccount
      mailAccount: 'mailAccount',

      // Field code of attachFile
      attachFile: 'attachFile'
    }
  }
};
```

***In case of cooperation with Outlook Schedule***

```javascript
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
```


6. Upload JavaScript for PC

***In case of cooperation with Outlook Mail***
* [msal.js](https://secure.aadcdn.microsoftonline-p.com/lib/1.0.0/js/msal.js)
* kintone-ui-component.min.js v0.4.2
* kintone-js-sdk.min.js v0.7.0
* [jquery.min.js v3.4.1](https://js.cybozu.com/jquery/3.4.1/jquery.min.js)
* [sweetalert2.min.js v8.17.6](https://js.cybozu.com/sweetalert2/v8.17.6/sweetalert2.min.js)
* [common-js-functions.min.js](common/common-js-functions.min.js)
* [kintone-connect-outlook_mail_common.js](apps/outlook-mail/js/kintone-connect-outlook_mail_common.js)
* [oauth.js](common/outlook-auth/js/oauth.js)
* [kintone-connect-outlook_mail.js](apps/outlook-mail/js/kintone-connect-outlook_mail.js)

***In case of cooperation with Outlook Schedule***
* [msal.js](https://secure.aadcdn.microsoftonline-p.com/lib/1.0.0/js/msal.js)
* kintone-ui-component.min.js v0.4.2
* kintone-js-sdk.min.js v0.7.0
* [jquery.min.js v3.4.1](https://js.cybozu.com/jquery/3.4.1/jquery.min.js)
* [sweetalert2.min.js v8.17.6](https://js.cybozu.com/sweetalert2/v8.17.6/sweetalert2.min.js)
* [common-js-functions.min.js](common/common-js-functions.min.js)
* [kintone-connect-outlook-schedule-common.js](apps/outlook-schedule/js/kintone-connect-outlook-schedule-common.js)
* [oauth.js](common/outlook-auth/js/oauth.js)
* [kintone-connect-outlook-schedule.js](apps/outlook-schedule/js/kintone-connect-outlook-schedule.js)


7. Upload CSS File for PC

***In case of cooperation with Outlook Mail***
* [sweetalert2.min.css v8.17.6](https://js.cybozu.com/sweetalert2/v8.17.6/sweetalert2.min.css)
* kintone-ui-component.min.css v0.4.2

***In case of cooperation with Outlook Schedule***
* [sweetalert2.min.css v8.17.6](https://js.cybozu.com/sweetalert2/v8.17.6/sweetalert2.min.css)
* kintone-ui-component.min.css v0.4.2

## Authentication flow
![overview image](img/AuthenticationFlow.png?raw=true)

## License
MIT

## Copyright
Copyright(c) Cybozu, Inc.
