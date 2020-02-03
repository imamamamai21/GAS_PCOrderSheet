var PcOrderSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('資産管理T_発注リスト');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
PcOrderSheet.prototype = {
  getRowKey: function(target) {
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') this.showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  createIndex: function() {
    const DAY = '発注日';
    var filterData = this.values.filter(function(value) {
      return value.indexOf(DAY) > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      this.showTitleError();
      return;
    }
    
    this.index = {
      date        : filterData.indexOf(DAY),
      person      : filterData.indexOf('発注者'),
      num         : filterData.indexOf('台数'),
      deliveryDate: filterData.indexOf('到着予定日'),
      employeeNum : filterData.indexOf('管理者 社員番号'),
      employeeName: filterData.indexOf('管理者 氏名'),
      deliveryArea: filterData.indexOf('納品先'),
      deliveryName: filterData.indexOf('納品先名'),
      shop        : filterData.indexOf('販売店'),
      odr         : filterData.indexOf('発注管理番号'),
      estimate    : filterData.indexOf('見積No'),
      taskNo      : filterData.indexOf('タスクNo'),
      cpu         : filterData.indexOf('CPU'),
      cpuFrequency: filterData.indexOf('CPU動作周波数'),
      productId   : filterData.indexOf('機種識別ID'),
      maker       : filterData.indexOf('メーカー'),
      product     : filterData.indexOf('製品名'),
      display     : filterData.indexOf('ディスプレイ'),
      pcType      : filterData.indexOf('種別'),
      model       : filterData.indexOf('モデル'),
      modelInfo   : filterData.indexOf('詳細品番'),
      key         : filterData.indexOf('キー'),
      memory      : filterData.indexOf('メモリ'),
      ssd         : filterData.indexOf('ストレージ(SSD)'),
      hdd         : filterData.indexOf('ストレージ(HDD)'),
      rentalPcNo  : filterData.indexOf('レンタルPC契約番号'),
      fixedAsset  : filterData.indexOf('固定資産'),
      isShare     : filterData.indexOf('共有PC'),
      paidInAdv   : filterData.indexOf('初回費用負担'),
      ringy       : filterData.indexOf('稟議番号'),
      url         : filterData.indexOf('台帳URL'),
      delivery    : filterData.indexOf('納品'),
      sdName      : filterData.indexOf('SD担当者'),
      deliveryBotDate: filterData.indexOf('納品通知日'),
      kitting     : filterData.indexOf('KT依頼')
    };
    return this.index;
  }
};
var pcOrderSheet = new PcOrderSheet();

function testPS() {
  var value = pcOrderSheet.values[19];
  var index = pcOrderSheet.getIndex();
  var row = pcOrderSheet.getRowKey('deliveryDate')
  Logger.log(Number(''));
}
