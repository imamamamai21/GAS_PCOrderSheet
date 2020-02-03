var MasterSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('マスタデータ');
  this.values = this.sheet.getDataRange().getValues();
  this.index = {};
  this.titleRow = 4;
  
  this.showTitleError = function(key) {
    Browser.msgBox(ERROR_TEXT_NONKEY_TITLE, ERROR_TEXT_NONKEY_MESSAGE + key, Browser.Buttons.OK);
  }
}
  
MasterSheet.prototype = {
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
    var me = this;
    var keys = this.values[me.titleRow - 1];
    this.values[0].forEach(function (value, i) { me.index[keys[i]] = i });
    return this.index;
  },
  /**
   * PC機種マスタ:データの更新
   */
  updateModelData: function() {
    var modelData = KintoneApi.modelApi.getAllData();
    // test用 var modelData = [{pc_class:{type:"DROP_DOWN", value:"\u8868\u793a"}, pc_appendix:{type:"MULTI_LINE_TEXT", value:""}, pc_category:{type:"DROP_DOWN", value:"N"}, pc_product:{type:"SINGLE_LINE_TEXT", value:"MacBook Pro"}}];
    if (modelData.length === 0) return;
    
    var modelRow = this.getRowKey('model_id');
    var modelLastRow = this.sheet.getRange(modelRow + ':' + modelRow).getValues().filter(String).length;
    this.sheet.getRange(this.titleRow + 1, this.getIndex().model_id + 1, modelLastRow, this.getIndex().pc_class + 1).clearContent();
    
    this.editData(modelData);
  },
  /**
   * CPUマスタ:データの更新
   */
  updateCpuData: function() {
    var cpuData = KintoneApi.cpuApi.getAllData();
    if (cpuData.length === 0) return;
    
    var cpuRow = this.getRowKey('cpu');
    var lastRow = this.sheet.getRange(cpuRow + ':' + cpuRow).getValues().filter(String).length;
    this.sheet.getRange(this.titleRow + 1, this.getIndex()['レコード番号'] + 1, lastRow, this.getIndex().cpu + 1).clearContent();
    
    this.editData(cpuData);
  },
  /**
   * 保管場所マスタ:データの更新
   */
  updatePlaceData: function() {
    var placeData = KintoneApi.placeApi.getAllData();
    if (placeData.length === 0) return;
    
    var locationRow = this.getRowKey('location_code');
    var lastRow = this.sheet.getRange(locationRow + ':' + locationRow).getValues().filter(String).length;
    this.sheet.getRange(this.titleRow + 1, this.getIndex().location_code + 1, lastRow, this.getIndex().location_name + 1).clearContent();
    
    this.editData(placeData);
  },
  /**
   * データを書き込む
   */
  editData: function(data) {
    var sortObj = {};
    var me = this;
    // 列ごとのobjectに変換する
    data.forEach(function(values, i) {
      Object.keys(me.getIndex()).forEach(function(key) {
        if (!data[i][key]) return;
        if (!sortObj[key]) sortObj[key] = [];
        sortObj[key][i] = [data[i][key].value];
      });
    });
    // タイトルごとに書き込む
    Object.keys(sortObj).forEach(function (key) {
      me.sheet.getRange(me.titleRow + 1, me.getIndex()[key] + 1, sortObj[key].length, 1).setValues(sortObj[key]);
    });
  }
};
var masterSheet = new MasterSheet();

/**
 * マスタデータ自動更新
 * トリガー登録：池田 毎日0時〜1時
 */
function updateMasterData() {
  masterSheet.updateModelData();
  //masterSheet.updateCpuData();
  //masterSheet.updatePlaceData();
}