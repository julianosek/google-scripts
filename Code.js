function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
}
