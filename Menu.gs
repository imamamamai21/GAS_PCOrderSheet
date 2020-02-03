/**
 * メニューを設定する
 * トリガー登録しています。(池田)
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();         // Uiクラスを取得する
  var menu = ui.createMenu('▼スクリプト');  // Uiクラスからメニューを作成する
  // メニューにアイテムを追加する
  menu.addItem('マニュアルを開く', 'openManual');
  menu.addItem('納品通知(資産管理)', 'postDeliveryBotForArms');
  menu.addItem('納品通知(技術購買部)', 'postDeliveryBotForMedia');
  menu.addItem('レンタルPC情報を記載する', 'inputWithRental');
  menu.addItem('新しく台帳登録する', 'onClickCreateRecord');
  menu.addToUi(); // メニューをUiクラスに追加する
}

/**
 * マニュアルを表示する
 */
function openManual() {
  var html = HtmlService
    .createTemplateFromFile('manual')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('マニュアル');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * レンタルPCの情報を記載する
 */
function inputWithRental() {
  var targetRow = Browser.inputBox('レンタルPC情報をシートに追加します。', 'レンタルPC発注管理表より対象となる行数を入力してください(最新のODRのシートを参照します)', Browser.Buttons.OK_CANCEL);
  if (targetRow !== 'cancel') inputRentalPc(targetRow);
}

/**
 * 台帳登録するボタン押下時
 */
function onClickCreateRecord() {
  kintoneRecord.createRecord();
}