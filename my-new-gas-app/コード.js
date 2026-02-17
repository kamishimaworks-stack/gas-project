/**
 * Webアプリを表示する
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('器高式野帳')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * データの保存（表全体を受け取ってシートに保存）
 */
function saveAllData(dataList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('野帳データ');
  
  if (!sheet) {
    sheet = ss.insertSheet('野帳データ');
    // ヘッダー行の設定（設計値FH、差分Diff、モードなどを追加）
    sheet.appendRow(['測点', 'BS(後視)', 'IH(器械高)', 'FS(前視)', 'GH(地盤高)', 'FH(設計値)', 'Diff(差)', '備考']);
  }
  
  // 既存データをクリアして書き直す（簡易実装）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (dataList && dataList.length > 0) {
    sheet.getRange(2, 1, dataList.length, dataList[0].length).setValues(dataList);
  }
  
  return '保存しました';
}