var MediaPcOderSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('技術購買部_発注リスト');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
MediaPcOderSheet.prototype = {
  getRowKey: function(target) {
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
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
      shop        : filterData.indexOf('販売店'),
      num         : filterData.indexOf('台数'),
      deliveryDate: filterData.indexOf('到着予定日'),
      deliveryArea: filterData.indexOf('納品先'),
      odr         : filterData.indexOf('発注管理番号'),
      estimate    : filterData.indexOf('見積No'),
      maker       : filterData.indexOf('メーカー'),
      model       : filterData.indexOf('機種'),
      key         : filterData.indexOf('キー'),
      rentalPcNo  : filterData.indexOf('レンタルPC契約番号'),
      fixedAsset  : filterData.indexOf('固定資産'),
      delivery    : filterData.indexOf('納品'),
      sdName      : filterData.indexOf('SD担当者'),
      deliveryBotDate: filterData.indexOf('納品通知日'),
      kitting     : filterData.indexOf('KT依頼')
    };
    return this.index;
  }
};
var mediaOderSheet = new MediaPcOderSheet();