function clearData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("データ記録");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}
