/**
 * レンタルPC発注管理表からの情報を自動入力
 * 参照データ▶https://docs.google.com/spreadsheets/d/1gDB1Ub50TKvOzjo0XCbNMGoH1VVP5izu2OgwjlEqD9g/edit#gid=1455470312
 */
function inputRentalPc(targetRow) {
  var rentalPcSheet = new RentalPCOderSheet.RentalPcSheet();
  var dateRow = pcOrderSheet.getRowKey('date');
  // 発注日の記載がない行に書き込む
  var rowNum = pcOrderSheet.sheet.getRange(dateRow + ':' + dateRow).getValues().filter(String).length + 2;
  var referenceData = rentalPcSheet.getValues()[targetRow - 1];
  var index = rentalPcSheet.getIndex();

  var popup = Browser.msgBox(rowNum + '行目に作られます', '管理者社員番号・機種番号などを忘れずに入れてください。', Browser.Buttons.OK_CANCEL);
  if (popup != 'ok') return;
  
  var sheet = pcOrderSheet.sheet;
  Logger.log(referenceData[index.deliveryPlanDate])
  
  sheet.getRange(pcOrderSheet.getRowKey('date') + rowNum).setValue(Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"));
  sheet.getRange(pcOrderSheet.getRowKey('person') + rowNum).setValue(referenceData[index.name]);
  sheet.getRange(pcOrderSheet.getRowKey('num') + rowNum).setValue('1');
  sheet.getRange(pcOrderSheet.getRowKey('deliveryDate') + rowNum).setValue(referenceData[index.deliveryPlanDate]);
  sheet.getRange(pcOrderSheet.getRowKey('employeeName') + rowNum).setValue(referenceData[index.requesterName]);
  sheet.getRange(pcOrderSheet.getRowKey('deliveryArea') + rowNum).setValue(referenceData[index.deliveryArea]);
  sheet.getRange(pcOrderSheet.getRowKey('shop') + rowNum).setValue('横河レンタリース');
  sheet.getRange(pcOrderSheet.getRowKey('odr') + rowNum).setValue(rentalPcSheet.getOdr());
  sheet.getRange(pcOrderSheet.getRowKey('estimate') + rowNum).setValue(referenceData[index.estimateNo]);
  sheet.getRange(pcOrderSheet.getRowKey('taskNo') + rowNum).setValue(referenceData[index.taskNo]);
  sheet.getRange(pcOrderSheet.getRowKey('key') + rowNum).setValue(referenceData[index.key]);
  sheet.getRange(pcOrderSheet.getRowKey('rentalPcNo') + rowNum).setValue(referenceData[index.agreementNo]);
  sheet.getRange(pcOrderSheet.getRowKey('fixedAsset') + rowNum).setValue('No');
  sheet.getRange(pcOrderSheet.getRowKey('isShare') + rowNum).setValue('No');
  sheet.getRange(pcOrderSheet.getRowKey('paidInAdv') + rowNum).setValue('なし');
  sheet.getRange(pcOrderSheet.getRowKey('ringy') + rowNum).setValue(referenceData[index.ringiNo]);
}

function testren() {
  inputRentalPc(16)
}