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