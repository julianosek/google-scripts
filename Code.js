function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSheet(); //get the currently active sheet

  freezeSheet(sheet);
  colourSheet(sheet);
  trimSheet(sheet);
}

function freezeSheet(sheet) {
  sheet.setFrozenRows(2); // freeze first two rows
  sheet.setFrozenColumns(1); // freeze first column
}

function colourSheet(sheet) {
  sheet.getRange("1:1").setBackground("#0b5394"); // colour title row
  sheet.getRange("2:2").setBackground("#7babf8"); // colour criteria row
}

function trimSheet(sheet) {
  const lastRow = sheet.getLastRow(); // last row with content, so inner empty rows are preserved
  const lastCol = sheet.getLastColumn(); // last column with content, so inner empty columns are preserved
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
  if (maxCols > lastCol) {
    sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
  }
}
