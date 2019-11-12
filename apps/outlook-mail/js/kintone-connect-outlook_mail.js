jQuery.noConflict();
(function($) {
  'use strict';

  // common/common-js-functions.min.js
  var KC = window.kintoneCustomize;

  // common/outlook-auth/js/oauth.js
  var AO = window.azureOauth;

  // apps/outlook-mail/js/kintone-connect-outlook_mail_common.js
  var KAC = window.kintoneAzureConnect;

  // kintone-js-sdk
  var kintoneSDKRecord = new kintoneJSSDK.Record();
  var kintoneFile = new kintoneJSSDK.File();

  // Get value from kintone-connect-outlook_mail_common.js
  var SUBJECT_FIELD_CODE = KAC.kintone.fieldCode.subject;
  var CONTENT_FIELD_CODE = KAC.kintone.fieldCode.content;
  var FROM_FIELD_CODE = KAC.kintone.fieldCode.from;
  var TO_FIELD_CODE = KAC.kintone.fieldCode.to;
  var CC_FIELD_CODE = KAC.kintone.fieldCode.cc;
  var BCC_FIELD_CODE = KAC.kintone.fieldCode.bcc;
  var MESSAGE_ID_FIELD_CODE = KAC.kintone.fieldCode.messageId;
  var MAIL_ACCOUNT_FIELD_CODE = KAC.kintone.fieldCode.mailAccount;
  var ATTACH_FILE_FIELD_CODE = KAC.kintone.fieldCode.attachFile;

  var MAIL_GET_URL = KAC.mail.mailGetUrl;
  var MAIL_SEND_URL = KAC.mail.mailSendUrl;

  var storage = window.localStorage;

  var kintoneMailService = {

    lang: {
      en: {
        button: {
          signIn: 'Sign in Outlook',
          signOut: 'Sign out of Outlook',
          getMail: 'Receive email',
          sendmail: 'Send email',
          sendExec: 'Send',
          cancelExec: 'Cancel'
        },
        message: {
          info: {
            confirmSend: 'Do you want to send this email?'
          },
          warning: {
            noMail: 'There are no emails in your inbox.'
          },
          success: {
            sendExec: 'Your email has been sent successfully.'
          },
          error: {
            sendFailure: 'Failed to send email.',
            signInFailure: 'Failed to sign in Outlook.',
            getAccessTokenFailure: 'Failed to get access token.',
            accessOutlookFailure: 'Failed to access to outlook.',
            getOutlookMailFailure: 'Failed to get outlook mails.',
            addKintoneRecordFailure: 'Failed to add kintone Record.'
          }
        }
      },
      ja: {
        button: {
          signIn: 'Outlookにログイン',
          signOut: 'Outlookからログアウト',
          getMail: 'メール受信',
          sendmail: 'メール送信',
          sendExec: '送信',
          cancelExec: 'キャンセル'
        },
        message: {
          info: {
            confirmSend: 'メールを送信しますか?'
          },
          warning: {
            noMail: '受信箱にメールがありません'
          },
          success: {
            sendExec: 'メールの送信に成功しました'
          },
          error: {
            sendFailure: 'メールの送信に失敗しました',
            signInFailure: 'サインインできませんでした',
            getAccessTokenFailure: 'アクセストークンが取得できませんでした',
            accessOutlookFailure: 'Outlookにアクセス出来ませんでした',
            getOutlookMailFailure: 'Outlookメールの取得に失敗しました',
            addKintoneRecordFailure: 'kintoneレコードの登録に失敗しました'
          }
        }
      }
    },

    setting: {
      lang: 'ja',
      i18n: {},
      ui: {
        buttons: {
          signInOutlook: {
            text: 'signIn',
            type: 'submit'
          },
          signOut: {
            text: 'signOut',
            type: 'normal'
          },
          getMail: {
            text: 'getMail',
            type: 'normal'
          },
          sendMail: {
            text: 'sendmail',
            type: 'normal'
          }
        }
      }
    },

    data: {
      ui: {},
      mail: {
        profile: {
          emailAddress: ''
        }
      },
      isLoginOutlook: false
    },

    init: function() {
      this.setting.lang = kintone.getLoginUser().language || 'ja';
      this.setting.i18n = this.setting.lang in this.lang ? this.lang[this.setting.lang] : this.lang.en;
    },

    uiCreateForIndex: function(kintoneHeaderSpace) {
      if (typeof kintoneHeaderSpace === 'undefined') {
        return;
      }

      if (this.data.ui.kintoneCustomizeOutlookHeaderSigned !== undefined || this.data.ui.kintoneCustomizeOutlookHeaderNotSigned !== undefined) {
        return;
      }

      this.data.ui.kintoneCustomizeOutlookHeaderSigned = document.createElement('div');
      this.data.ui.kintoneCustomizeOutlookHeaderNotSigned = document.createElement('div');

      this.data.ui.kintoneCustomizeOutlookUserInfo = new kintoneUIComponent.Label({text: ''});

      this.data.ui.btnSignIn = this.createButton(this.setting.ui.buttons.signInOutlook, this.setting.i18n.button);
      this.data.ui.btnSignOut = this.createButton(this.setting.ui.buttons.signOut, this.setting.i18n.button);
      this.data.ui.btnGetmail = this.createButton(this.setting.ui.buttons.getMail, this.setting.i18n.button);

      this.data.ui.kintoneCustomizeOutlookHeaderSigned.style.display = 'none';

      this.data.ui.kintoneCustomizeOutlookHeaderNotSigned.appendChild(this.data.ui.btnSignIn.render());
      this.data.ui.kintoneCustomizeOutlookHeaderSigned.appendChild(this.data.ui.kintoneCustomizeOutlookUserInfo.render());
      this.data.ui.kintoneCustomizeOutlookHeaderSigned.appendChild(this.data.ui.btnSignOut.render());
      this.data.ui.kintoneCustomizeOutlookHeaderSigned.appendChild(this.data.ui.btnGetmail.render());

      kintoneHeaderSpace.appendChild(this.data.ui.kintoneCustomizeOutlookHeaderNotSigned);
      kintoneHeaderSpace.appendChild(this.data.ui.kintoneCustomizeOutlookHeaderSigned);

      this.data.ui.btnSignIn.element.style.display = 'inline-block';
      this.data.ui.btnSignIn.element.style.margin = '0 15px 15px 15px';

      this.data.ui.btnSignOut.element.style.display = 'inline-block';
      this.data.ui.btnSignOut.element.style.margin = '0 15px 15px 15px';

      this.data.ui.btnGetmail.element.style.display = 'inline-block';
      this.data.ui.btnGetmail.element.style.margin = '0 15px 15px 15px';

      this.data.ui.kintoneCustomizeOutlookUserInfo.element.style.display = 'inline-block';
      this.data.ui.kintoneCustomizeOutlookUserInfo.element.style.margin = '0 15px 15px 15px';

    },

    uicreateForDetail: function() {
      var kintoneDetailHeaderSpace = kintone.app.record.getHeaderMenuSpaceElement();
      if (!this.isExpireAccessToken()) {
        return;
      }
      kintone.app.record.setFieldShown(MESSAGE_ID_FIELD_CODE, false);
      kintone.app.record.setFieldShown(MAIL_ACCOUNT_FIELD_CODE, false);
      this.data.ui.btnSendmail = this.createButton(this.setting.ui.buttons.sendMail, this.setting.i18n.button);
      this.data.ui.btnSendmail.element.style.display = 'inline-block';
      this.data.ui.btnSendmail.element.style.margin = '15px 0 0 15px';

      kintoneDetailHeaderSpace.appendChild(this.data.ui.btnSendmail.render());
    },

    createButton: function(setting, lang) {
      var text = lang ? lang[setting.text] || setting.text || '' : setting.text || '';
      var type = setting.type;
      var uiButton = new kintoneUIComponent.Button({text: text, type: type});
      return uiButton;
    },

    isExpireAccessToken: function() {
      if (storage.getItem('SESSION_KEY_TO_ACCESS_TOKEN')) {
        return true;
      }
      return false;
    },

    isSignUserDispInfo: function() {
      if (storage.getItem('SIGN_USER_MAILACCOUNT')) {
        return true;
      }
      return false;
    }
  };

  // outlook api用処理
  var outlookAPI = {

    // 初期処理
    init: function() {

      AO.init();
      if (!kintoneMailService.isExpireAccessToken() || !kintoneMailService.isSignUserDispInfo()) {
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderNotSigned.style.display = 'inline-block';
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderSigned.style.display = 'none';
      } else {
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderNotSigned.style.display = 'none';
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderSigned.style.display = 'inline-block';
        kintoneMailService.data.ui.kintoneCustomizeOutlookUserInfo.setText(storage.getItem('SIGN_USER_MAILACCOUNT'));
        kintoneMailService.data.mail.profile.emailAddress = storage.getItem('SIGN_USER_MAILACCOUNT');
        kintoneMailService.data.isLoginOutlook = true;
      }
    },

    signIn: function() {
      var self = this;
      KC.ui.loading.show();

      AO.signIn().then(function(id_token) {
        self.callGraphApi();
        KC.ui.loading.hide();
      }, function(error) {
        Swal.fire({
          title: 'Error!',
          type: 'error',
          text: kintoneMailService.setting.i18n.message.error.signInFailure,
          allowOutsideClick: false
        });
        KC.ui.loading.hide();
      });
    },

    signOut: function() {
      KC.ui.loading.show();
      storage.clear();
      AO.signOut();
      KC.ui.loading.hide();
    },

    // In order to call the Graph API, an access token needs to be acquired.
    callGraphApi: function() {
      var self = this;

      AO.callGraphApi().then(function(token) {

        var userInfo = AO.getUserInfo();

        // 「getMail」「signout」ボタンを表示
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderNotSigned.style.display = 'none';
        kintoneMailService.data.ui.kintoneCustomizeOutlookHeaderSigned.style.display = 'inline-block';
        kintoneMailService.data.ui.kintoneCustomizeOutlookUserInfo.setText(userInfo);
        kintoneMailService.data.mail.profile.emailAddress = userInfo;
        kintoneMailService.data.isLoginOutlook = true;

        // セッションに入れておく
        storage.setItem('SESSION_KEY_TO_ACCESS_TOKEN', token.accessToken);
        storage.setItem('SIGN_USER_MAILACCOUNT', userInfo);

        KC.ui.loading.hide();

      }, function(error) {
        if (error) {
          Swal.fire({
            title: 'Error!',
            type: 'error',
            text: kintoneMailService.setting.i18n.message.error.getAccessTokenFailure,
            allowOutsideClick: false
          });
          self.userAgentApplication = null;
          KC.ui.loading.hide();
        }
      }).catch(function() {
        KC.ui.loading.hide();
      });
    },

    getMessageIDIndexIsNotRetrived: function(index, data) {
      var self = this;

      // 取得済みメールかチェック
      return self.checkMessagesIsNotExistsOnkintone(data[index].id).then(function(resp) {
        if (resp === false) {
          if (data.length <= index + 1) {
            // Out of index
            return null;
          }
          return self.getMessageIDIndexIsNotRetrived(index + 1, data);
        }
        return index;
      });
    },

    // 取得済みメールかチェック
    checkMessagesIsNotExistsOnkintone: function(messageID) {
      var dataRequestkintoneApp = {
        app: kintone.app.getId(),
        fields: ['$id'],
        query: MESSAGE_ID_FIELD_CODE + ' like "' + messageID + '"',
        totalCount: true
      };
      return kintoneSDKRecord.getRecords(dataRequestkintoneApp).then(function(response) {

        if (!response.records || response.records.length === 0) {
          return true;
        }
        return false;
      });
    },

    // outlookメール取得
    getMail: function() {
      var self = this;
      var accessToken;
      var header;
      KC.ui.loading.show();

      if (kintoneMailService.isExpireAccessToken()) {
        accessToken = storage.getItem('SESSION_KEY_TO_ACCESS_TOKEN');
      } else {
        KC.ui.loading.hide();
        return;
      }

      header = {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json',
        'outlook.body-content-type': 'html'
      };

      // OutlookのINBOXからメール取得
      kintone.proxy(MAIL_GET_URL, 'GET', header, {}).then(function(res) {
        var data = JSON.parse(res[0]).value;
        if (data === undefined) {
          Swal.fire({
            title: 'ERROR!',
            type: 'error',
            text: kintoneMailService.setting.i18n.message.error.accessOutlookFailure,
            allowOutsideClick: false
          });
          KC.ui.loading.hide();
          return;
        }

        // 受信箱にメールが存在しない場合
        if (data.length === 0) {
          Swal.fire({
            title: 'WARN!',
            type: 'warning',
            text: kintoneMailService.setting.i18n.message.warning.noMail,
            allowOutsideClick: false
          });
          KC.ui.loading.hide();
          return;
        }

        // 取得したメールをkintoneへ登録
        self.putMailToKintoneApp(0, data, accessToken).catch(function(err) {
          Swal.fire({
            title: 'ERROR!',
            type: 'error',
            text: kintoneMailService.setting.i18n.message.error.addKintoneRecordFailure,
            allowOutsideClick: false
          });
          KC.ui.loading.hide();
        });
      }, function(err) {
        Swal.fire({
          title: 'ERROR!',
          type: 'error',
          text: kintoneMailService.setting.i18n.message.error.getOutlookMailFailure,
          allowOutsideClick: false
        });
        KC.ui.loading.hide();
      });
    },

    // 取得したメールをkintoneへ登録
    putMailToKintoneApp: function(index, data, accessToken) {
      var self = this;

      return this.getMessageIDIndexIsNotRetrived(index, data).then(function(indexMessage) {

        if (indexMessage === null) {
          return indexMessage;
        }
        // まだ未取得のメールを登録
        return self.addMailIntoKintone(data[indexMessage], accessToken).then(function(result) {
          // Process next mail
          if (indexMessage + 1 < data.length) {
            return self.putMailToKintoneApp(indexMessage + 1, data, accessToken);
          }
          return null;
        });
      }).then(function(resp) {
        window.location.reload();
        KC.ui.loading.hide();
      });
    },

    addMailIntoKintone: function(data, accessToken) {
      var kintoneRecord = {};
      var toRecipents;
      var toRecipentsStr;
      var ccRecipents;
      var ccRecipentsStr;
      var bccRecipents;
      var bccRecipentsStr;


      // Subject
      kintoneRecord[SUBJECT_FIELD_CODE] = {
        value: data.subject
      };

      // From
      kintoneRecord[FROM_FIELD_CODE] = {
        value: data.from.emailAddress.address
      };

      // To
      toRecipents = data.toRecipients;
      toRecipentsStr = toRecipents.map(function(el) {
        return el.emailAddress.address;
      }).toString();
      kintoneRecord[TO_FIELD_CODE] = {
        value: toRecipentsStr || ''
      };

      // Cc
      ccRecipents = data.ccRecipients;
      ccRecipentsStr = ccRecipents.map(function(el) {
        return el.emailAddress.address;
      }).toString();
      kintoneRecord[CC_FIELD_CODE] = {
        value: ccRecipentsStr || ''
      };

      // Bcc
      bccRecipents = data.bccRecipients;
      bccRecipentsStr = bccRecipents.map(function(el) {
        return el.emailAddress.address;
      }).toString();
      kintoneRecord[BCC_FIELD_CODE] = {
        value: bccRecipentsStr || ''
      };

      // Body
      kintoneRecord[CONTENT_FIELD_CODE] = {
        value: data.body.content ? outlookAPI.removeStyleTagOnString(data.body.content) : data.bodyPreview
      };

      // messageId
      kintoneRecord[MESSAGE_ID_FIELD_CODE] = {
        value: data.id
      };

      // mailAccount
      kintoneRecord[MAIL_ACCOUNT_FIELD_CODE] = {
        value: kintoneMailService.data.mail.profile.emailAddress
      };

      // attachFile
      if (data.hasAttachments) {
        return outlookAPI.getAttach(data.id, accessToken).then(function(promiseArrayFiles) {
          return kintone.Promise.all(promiseArrayFiles);
        }).then(function(arrayFiles) {
          kintoneRecord[ATTACH_FILE_FIELD_CODE] = {
            value: arrayFiles
          };
        }).then(function() {
          return outlookAPI.addKintone(kintoneRecord);
        });
      }
      return outlookAPI.addKintone(kintoneRecord);
    },

    // kintoneへ登録
    addKintone: function(postParam) {
      var param = {
        app: kintone.app.getId(),
        record: postParam
      };
      return kintoneSDKRecord.addRecord(param);
    },

    // 添付ファイル取得
    getAttach: function(messageId, accessToken) {
      var self = this;
      var url = 'https://graph.microsoft.com/v1.0/me/messages/' + messageId + '/attachments';
      var header = {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      };

      // メールに添付されているファイルを取得
      return kintone.proxy(url, 'GET', header, {}).then(function(res) {
        var data = JSON.parse(res[0]).value;

        // kintoneへファイルをアップロード
        return self.uploadFileToKintone(messageId, data, 0, [], accessToken);
      });
    },

    // kintoneへファイルをアップロード
    uploadFileToKintone: function(messageId, attachData, index, attachDataArr, accessToken) {
      var self = this;
      var attachId = attachData[index].id;
      var url = 'https://graph.microsoft.com/v1.0/me/messages/' + messageId + '/attachments/' + attachId;
      var header = {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      };

      // 添付ファイル情報を取得
      return kintone.proxy(url, 'GET', header, {}).then(function(attachRes) {
        var data = JSON.parse(attachRes[0]);
        var param = {
          fileName: data.name,
          fileBlob: outlookAPI.convertBase64AttachmentToBlob(data.contentBytes, data.contentType),
        };

        // kintoneへファイルをアップロード
        return kintoneFile.upload(param).then(function(resp) {
          var fileKey = {};
          fileKey.fileKey = resp.fileKey;
          attachDataArr.push(fileKey);

          if (index + 1 < attachData.length) {
            return self.uploadFileToKintone(messageId, attachData, index + 1, attachDataArr);
          }
          return attachDataArr;
        });
      });
    },

    getKintoneAttach: function(attaches) {
      // 添付ファイル分
      var outlookAttachements = [];
      var outlookAttachement;
      var i;
      for (i = 0; i < attaches.length; i++) {
        // kintoneからファイルをダウンロード
        outlookAttachement = this.kintoneFileToOutlookAttachment(attaches[i]).catch(function(error) {});
        outlookAttachements.push(outlookAttachement);
      }
      return kintone.Promise.all(outlookAttachements);
    },

    kintoneFileToOutlookAttachment: function(kintoneFileObj) {
      var self = this;
      var url = '/k/v1/file.json?fileKey=' + kintoneFileObj.fileKey;
      return new kintone.Promise(function(resolve, reject) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        xhr.responseType = 'arraybuffer';
        xhr.onload = function() {
          var outlookAttachment = {
            '@odata.type': '#microsoft.graph.fileAttachment',
            'name': kintoneFileObj.name,
            'contentBytes': outlookAPI.convertArrayBufferToBase64(this.response)
          };
          resolve(outlookAttachment);
        };
        xhr.onerror = function(e) {
          reject(e);
        };
        xhr.send();
      }).then({}, function(error) {
        self.showError(error);
        return false;
      });
    },

    decodeAttachment: function(stringEncoded) {
      return atob(stringEncoded.replace(/-/g, '+').replace(/_/g, '/'));
    },

    // base64データをBlobデータに変換
    convertBase64AttachmentToBlob: function(base64String, contentType) {
      var i;
      var blob;
      var binary = outlookAPI.decodeAttachment(base64String);
      var len = binary.length;
      var arrBuffer = new ArrayBuffer(len);
      var fileOutput = new Uint8Array(arrBuffer);
      for (i = 0; i < len; i++) {
        fileOutput[i] = binary.charCodeAt(i);
      }
      blob = new Blob([arrBuffer], {
        type: (contentType || 'octet/stream') + ';charset=utf-8;'
      });
      return blob;
    },

    convertArrayBufferToBase64: function(arraybuffer) {
      var i;
      var binary = '',
        bytes = new Uint8Array(arraybuffer),
        len = bytes.byteLength;
      for (i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      return window.btoa(binary);
    },

    removeStyleTagOnString: function(text) {
      return text.replace(/<style.*style>|<style.*[[\s\S\t]*?.*style>/mg, '');
    },

    sendMailInit: function(kinRec) {
      var self = this;
      // Confirm whether to execute
      Swal.fire({
        title: kintoneMailService.setting.i18n.message.info.confirmSend,
        type: 'warning',
        confirmButtonColor: '#DD6B55',
        confirmButtonText: kintoneMailService.setting.i18n.button.sendExec,
        cancelButtonText: kintoneMailService.setting.i18n.button.cancelExec,
        showCancelButton: 'true',
        allowOutsideClick: false
      }).then(function(isConfirm) {
        if (isConfirm.dismiss !== 'cancel') {
          self.sendMail(kinRec);
        } else {
          KC.ui.loading.hide();
        }
      });
    },

    // メール送信
    sendMail: function(kintoneData) {
      var accessToken;
      var sendParam = {};

      var toRecipentsArr = [];
      var ccRecipentsArr = [];
      var bccRecipentsArr = [];

      KC.ui.loading.show();
      if (kintoneMailService.isExpireAccessToken()) {
        accessToken = storage.getItem('SESSION_KEY_TO_ACCESS_TOKEN');
      } else {
        return;
      }

      // 件名
      sendParam.subject = kintoneData[SUBJECT_FIELD_CODE].value;
      // 本文
      sendParam.body = {
        'contentType': 'html',
        'content': kintoneData[CONTENT_FIELD_CODE].value
      };

      // To
      if (kintoneData[TO_FIELD_CODE].value) {
        kintoneData[TO_FIELD_CODE].value.replace(/\s/g, '').split(',').forEach(function(email) {
          if (!email) {
            return;
          }
          toRecipentsArr.push({
            'emailAddress': {
              'address': email
            }
          });
        });
        sendParam.toRecipients = toRecipentsArr;
      }

      // Cc
      if (kintoneData[CC_FIELD_CODE].value) {
        kintoneData[CC_FIELD_CODE].value.replace(/\s/g, '').split(',').forEach(function(email) {
          if (!email) {
            return;
          }
          ccRecipentsArr.push({
            'emailAddress': {
              'address': email
            }
          });
        });
        sendParam.ccRecipients = ccRecipentsArr;
      }

      // Bcc
      if (kintoneData[BCC_FIELD_CODE].value) {
        kintoneData[BCC_FIELD_CODE].value.replace(/\s/g, '').split(',').forEach(function(email) {
          if (!email) {
            return;
          }
          bccRecipentsArr.push({
            'emailAddress': {
              'address': email
            }
          });
        });
        sendParam.bccRecipients = bccRecipentsArr;
      }
      outlookAPI.getKintoneAttach(kintoneData[ATTACH_FILE_FIELD_CODE].value).then(function(files) {
        // ここはあってた
        sendParam.attachments = files;
        return outlookAPI.sendOutlook(sendParam, accessToken);
      });
    },

    // Outlookへ登録
    sendOutlook: function(sendParam, accessToken) {

      var header = {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      };

      var message = {};
      message.message = sendParam;
      message.saveToSentItems = true;
      kintone.proxy(MAIL_SEND_URL, 'POST', header, message).then(function(respdata) {
        var responseDataJson = window.JSON.parse(!respdata[0] ? '{}' : respdata[0]);
        if (typeof responseDataJson.error !== 'undefined') {
          Swal.fire({
            title: 'Error!',
            type: 'error',
            text: kintoneMailService.setting.i18n.message.error.sendFailure,
            allowOutsideClick: false
          });
        } else {
          Swal.fire({
            title: 'SUCCESS!',
            type: 'success',
            text: kintoneMailService.setting.i18n.message.success.sendExec,
            allowOutsideClick: false
          });
        }
        KC.ui.loading.hide();
      }).catch(function(error) {
        Swal.fire({
          title: 'Error!',
          type: 'error',
          text: kintoneMailService.setting.i18n.message.error.sendFailure,
          allowOutsideClick: false
        });
        KC.ui.loading.hide();
      });
    }
  };

  // レコード一覧画面の表示時
  kintone.events.on('app.record.index.show', function(event) {

    kintoneMailService.init();

    /* create kintone ui */
    kintoneMailService.uiCreateForIndex(kintone.app.getHeaderSpaceElement());

    // 初期処理
    outlookAPI.init();

    // Singinボタン押下時
    kintoneMailService.data.ui.btnSignIn.on('click', function() {
      outlookAPI.signIn();
    });

    // Singoutボタン押下時
    kintoneMailService.data.ui.btnSignOut.on('click', function() {
      outlookAPI.signOut();
    });

    // GET MAILボタン押下時
    kintoneMailService.data.ui.btnGetmail.on('click', function() {
      outlookAPI.getMail();
    });
  });

  // レコード詳細画面の表示時
  kintone.events.on('app.record.detail.show', function(event) {
    var record = event.record;

    kintoneMailService.init();

    /* create kintone ui */
    kintoneMailService.uicreateForDetail();

    // SEND MAILボタン押下時
    kintoneMailService.data.ui.btnSendmail.on('click', function() {
      outlookAPI.sendMailInit(record);
    });
  });

  // レコード作成画面の表示時
  kintone.events.on('app.record.create.show', function(event) {
    var record = event.record;
    kintone.app.record.setFieldShown(MESSAGE_ID_FIELD_CODE, false);
    kintone.app.record.setFieldShown(MAIL_ACCOUNT_FIELD_CODE, false);
    record[FROM_FIELD_CODE].disabled = true;
    record[FROM_FIELD_CODE].value = storage.getItem('SIGN_USER_MAILACCOUNT');
    return event;
  });

  // レコード編集画面の表示時
  kintone.events.on('app.record.edit.show', function(event) {
    var record = event.record;
    kintone.app.record.setFieldShown(MESSAGE_ID_FIELD_CODE, false);
    kintone.app.record.setFieldShown(MAIL_ACCOUNT_FIELD_CODE, false);
    record[FROM_FIELD_CODE].disabled = true;
    return event;
  });

})(jQuery);
