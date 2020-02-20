////////////////////////////////////////////////////////////
//////////////////// 台帳登録するためのコード ////////////////////
////////////////////////////////////////////////////////////

var KintoneRecord = function() {
  this.api = KintoneApi.caApi.api;
  
  /**
   * 作成したレコードのURLを作る & シートに登録する
   * @param {number} num 行数
   */
  this.createRecordUrl = function(num) {
    var data = pcOrderSheet.values[num - 1];
    var url = this.createNewRecord(data);
    pcOrderSheet.sheet.getRange(pcOrderSheet.getRowKey('url') + num).setValue(url);
    return url;
  }
  
  /**
   * データの過不足があったときに出すアラート
   * @param {number} 行数
   */
  this.showAleart = function(num) {
    Browser.msgBox('作成できませんでした', '指定された行数(' + num + ')が有効ではありません。', Browser.Buttons.OK);
  }
  
  /**
   * 作成された台帳の案内を出します
   * @param {[string]} urls 台帳URL
   */
  this.showNewRecordInfo = function(urls) {
    const subTextFixed = '<p>※稟議番号が現時点で不明な場合、後で必ず入れてください。</p>';
    var urlText = '';
    
    urls.forEach(function(url, index) {
      var num = urls.length > 1 ? '【' + (index + 1)  + '】' : '';
      urlText += '<p><a href="' + url +  '" target="blank">' + num + 'NEWレコード▶</a></p>';
    });
    var htmlOutput = HtmlService
        .createHtmlOutput(urlText + '<p>新しく作られた台帳に間違いがないか確認してください。</p>' + subTextFixed)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(450);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '台帳が作成されました');
  }
  /**
   * 台帳に新しいレコードを作成
   * @param {object} 表のデータ1行分
   * @return {string} 台帳のURL
   */
  this.createNewRecord = function(data) {
    var newNo = KintoneApi.groupApi.getNewGroupId();
    var index = pcOrderSheet.getIndex();
    var ringy = data[index.ringy].toString();
    
    var response = this.api.postRecord({
      capc_id       : { value: newNo },
      pc_hostname   : { value: 'CAグループPC管理番号' },
      pc_id         : { value: newNo },
      pc_status     : { value: '納品待ち' },
      shared        : { value: data[index.isShare] === 'Yes' ? ['共有PC'] : [] },
      location      : { value: data[index.deliveryArea] },
      delivery_date : { value: '' },
      user_id_lookup: { value: data[index.employeeNum] }, // 管理者 社員番号
      pc_master     : { value: data[index.productId] },
      model_number  : { value: data[index.modelInfo] },
      keyboard      : { value: data[index.key] },
      cpu           : { value: data[index.cpu] },
      cpu_clock     : { value: data[index.cpuFrequency] },
      memory        : { value: data[index.memory] },
      ssd           : { value: data[index.ssd] },
      hdd           : { value: data[index.hdd] },
      appendix      : { value: '見積もりNo: ' + data[index.estimate] + '\n発注タスクNo: ' + data[index.taskNo] + (data[index.odr] === '' ? '' : '\n発注管理番号: ' + data[index.odr]) }, // 備考
      ringy         : { value: (ringy != '' && ringy.length < 10 ) ? ('0000000000' + ringy).slice(-10) : data[index.ringy] }, // 稟議番号は10桁に揃える
      paid_in_adv   : { value: data[index.paidInAdv] === 'あり' ? ['あり'] : [] }, // 初回費用負担
      rental_status : { value: data[index.rentalPcNo] ? '継続'　: null }, // レンタルステータス
      rental_contractid: { value: data[index.rentalPcNo] }
    });
    return this.api.getUri(response.id); // 新しい台帳のURL
  }
  /**
   * popupで打ち込んだ数字を配列にして返す
   * @param string 例： '3'や'1,2' '2-8'など
   * @return [number]
   */
  this.getNumber = function(numsText) {
    if (numsText.indexOf('-') === -1) return numsText.split(',');
    
    var array = [];
    var numsAry = numsText.split('-');
    if (numsAry.length != 2) return [];
    
    for (var i = Number(numsAry[0]); i <= Number(numsAry[1]); i++){
      array.push(i);
    }
    return array;
  }
}

KintoneRecord.prototype = {
  /**
   * レコードを作成する
   * ①データをチェックする　②kintoneに登録する ③台帳のURLを作る ④シートに台帳URLを書き込む
   */
  createRecord: function() {
    var popup = Browser.inputBox('台帳を作成します', 
      '作りたいPC情報が書かれた行数を半角数字で打ち込んでください。\n※複数行分作りたい場合はスペースなしの`,(半角カンマ)`で続ける、もしくは`-(半角ハイフン)`でご記入ください。\n※カンマの場合はカンマでつなげる分全てです(1,3なら1,3行目)\n※ハイフンでつなげる場合は連番になります(1-3なら1,2,3行目)\n※複数で指定できるのはハイフンかカンマかどちらかのみです。',
      Browser.Buttons.OK_CANCEL);
    if(popup === 'cancel') return;
    
    var urls = [];
    var me = this;
    
    this.getNumber(popup).forEach(function (num) {
      if (!Number(num)) {
        me.showAleart(num);
        return;
      }
      var value = pcOrderSheet.values[num - 1];
      Logger.log('value = ' + value)
      var index = pcOrderSheet.getIndex();
      function isNg(key) { return value[index[key]] === '' }
      
      if (isNg('person')) {
        me.showAleart(num);
        return;
      }
      if (value[index.num] != 1) {
        Browser.msgBox('複数台の場合、台帳を作成できません', '指定された行数' + num + 'は複数台あります。手動で作成していただくか、1列につき1台ずつ記載してからやり直しください。', Browser.Buttons.OK);
        return;
      }
      if (!isNg('url')) {
        Browser.msgBox('すでに台帳は作成されています', '指定した行数(' + num + ')が間違っていないかご確認ください。', Browser.Buttons.OK);
        return;
      }
      if (value[index.fixedAsset] === 'Yes' && isNg('ringy')) { // 固定資産なのに稟議番号ない場合
        var alert = Browser.msgBox('稟議番号がありませんがこのまま作りますか？', '固定資産のため、キャンセルして' + num + '行目に稟議番号を入力してからやり直してください。子会社など稟議番号がない場合はこのまま進めてください。', Browser.Buttons.OK_CANCEL);
        if(alert === 'cancel') return;
      }
      if (num === '' || isNg('deliveryArea') || isNg('estimate') || isNg('taskNo') || isNg('cpu') || isNg('productId') ||  isNg('key') || isNg('memory') || isNg('ssd')) {
        Browser.msgBox('情報が足りないため作成できません', '空欄で黄色になっている箇所は必須項目です。' + num + '行目の足りない箇所に記載してからやり直してください。', Browser.Buttons.OK);
        return;
      }
      // 台帳を作る
      urls.push(me.createRecordUrl(num));
    });
    if(urls.length > 0) this.showNewRecordInfo(urls);
  }
}
var kintoneRecord = new KintoneRecord();
