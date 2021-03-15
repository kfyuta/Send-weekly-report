function sendWeeklyReport() {
  const token = ScriptApp.getOAuthToken();   
  const url = "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
  const file = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}}).getBlob().setName(FILENAME); 
  MailApp.sendEmail(TO, SUBJECT, MESSAGE, {attachments:file}); 
}

function createNewWeeklyReport() {
  // ワークシートのIDを指定しオブジェクトを取得
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 2番目のシート（最新版）を取得しコピーを作成。
  const prevSheet = ss.getSheets()[1]; 
  const newSheet = ss.insertSheet(1, {template: prevSheet});

  // コピー元シートの7日後の日付をコピーしたシート名に設定。
  const date = prevSheet.getName();
  const pd = new Date(date.substring(0, 4) + "-" + date.substring(4, 6) + "-" + date.substring(6, 8));
  const newDate = new Date(pd.getFullYear(), pd.getMonth(), pd.getDate() + 7);
  newSheet.setName(Utilities.formatDate(newDate, "Asia/Tokyo", "yyyyMMdd"));

  // コピーしたシートの日付を1週間後に変更
  const target = newSheet.getRange(5, 2, 1, 3);
  const targetValues = target.getValues();
  targetValues[0][0] = calcNthAfterDays(targetValues[0][0], 7);
  targetValues[0][2] = calcNthAfterDays(targetValues[0][2], 7);
  target.setValues(targetValues);
}
 
// 第一引数のn日後(前)を求める
function calcNthAfterDays(prevDate, n) {
  const date = new Date(prevDate);
  date.setDate(date.getDate() + n);
  return date;
}