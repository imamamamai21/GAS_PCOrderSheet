/* ========== 納品した時にBOTを Workplaceで通知するためのスクリプト(資産管理用) ============ */

/**
 * 資産管理への納品通知
 * メニューから「納品を通知する」が押された時呼び出される
 */
function postDeliveryBotForArms() {
  deliveryBot(pcOrderSheet, SHEET_ID__ARMS, WorkplaceApi.postBotForSDArms);
}

/**
 * 資産管理への納品通知 (V1)
 * メニューから「納品を通知する」が押された時呼び出される
 */
function postDeliveryBotForArmsV1() {
  deliveryBot(orderSheetV1, SHEET_ID__V1, WorkplaceApi.postBotForSDArms);
}

/**
 * 技術購買部への納品通知
 */
function postDeliveryBotForMedia() {
  deliveryBot(mediaOderSheet, SHEET_ID__MEDIA, WorkplaceApi.postBotForMediaSD);
}

/**
 * 納品通知テスト
 */
function postDeliveryBotForTest() {
  deliveryBot(orderSheetV1, SHEET_ID__MEDIA, WorkplaceApi.postBotForTest);
}

/**
 * 納品のBOT通知
 * @param {Sheet()} sheet    : シートオブジェクト
 * @param {string}  gid      : シートのGid
 * @param {void}    postFumc : WorkplaceApiのpost関数
 */
function deliveryBot(sheet, gid, postFumc) {
  var index = sheet.getIndex();
  
  // 納品通知に更新がある&&空欄じゃない時、配列に対象の情報を保存する
  var botMessages = sheet.values.map(function(value, i) {
    if (value[index.delivery] != '' && value[index.sdName] != '' && !value[index.kitting] && value[index.deliveryBotDate] === '') {
      var employee = value[index.employeeName] ? '_' + (value[index.employeeName].replace('　', '') || '●●') + 'さん_' : '';
      var task = value[index.taskNo] ? (value[index.taskNo] || 'タスクNO') : '';
      var text = '\n## ' + value[index.model] + '_' + value[index.num] + '台' + employee + task + '\n' +
          '```\n見積No : ' + value[index.estimate] + '\n' +
            '発注者 : ' + value[index.person] + '\n' +
              '発注管理番号 : ' + value[index.odr] + '\n' +
              '納品先 : ' + value[index.deliveryArea] + (index.deliveryName ? '(' + value[index.deliveryName] + ')' : '') + '\n' +
                '販売店 : ' + value[index.shop] + '\n```\n' +
                  (value[index.url] ? '[▶台帳URL](' + value[index.url] + ')\n' : '') + '\n ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~';
      return { text: text, row: i + 1 };
    }
    else return null;
  }).filter(function(value) { return value != null });
  
  if (botMessages.length === 0) {
    Browser.msgBox('納品を通知する対象がありません', '以下を確認してください\\n・納品とSD担当者の欄に書き込んでいますか？\\n・資産管理用と技術購買部用で間違えていませんか？\\n・対象のKT依頼にチェックがついていませんか？\\n・すでに納品通知欄に日付が書かれていませんか？', Browser.Buttons.OK);
    return;
  }
  const TITLE = '# ■■■PCが納品されました■■■\n [PC発注確認シート](' + MY_SHEET_URL + MY_SHEET_ID + '/edit#gid=' + gid + ')\n';
  const FOOTER = '\n担当者の方はご対応お願いいたします。'
  var text = TITLE + botMessages.map(function(value) { return value.text }).join('') + FOOTER;
  
  // feed投稿
  postFumc(text);
  
  // 通知したことをシートに記入
  var timeStamp = Utilities.formatDate(new Date(), 'JST', 'MM/dd');
  botMessages.forEach(function (value) {
    sheet.sheet.getRange(sheet.getRowKey('deliveryBotDate') + value.row).setValue(timeStamp);
  });
  
  Browser.msgBox('BOTを通知しました', text, Browser.Buttons.OK);
}

/**
 * workplaceのメモからworkplaceのIDを取得
 なんかメンションうまくいかないのでコメントアウト
function getID(name) {
  var sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName(WARKPLACE_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 4; i < 26; i++) {
    if (name === data[i][1]) return data[i][2];
  }
  return '';
}*/
