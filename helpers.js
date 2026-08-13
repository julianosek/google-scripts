const HEADER_ROWS = 2; // number of header rows (title + criteria)
const TITLE_ROW = 1;
const CRITERIA_ROW = 2;

// freezes the header rows and first column
function freezeSheet(sheet) {
  sheet.setFrozenRows(HEADER_ROWS); // freeze header rows
  sheet.setFrozenColumns(1); // freeze first column
}

// deletes empty rows/columns outside the table, leaving empty cells inside it intact
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

// converts any cell holding a boolean value into a checkbox
function addCheckboxes(sheet, dataRange) {
  const values = dataRange.getValues();

  for (let r = 0; r < values.length; r++) {
    const hasBoolean = values[r].some((value) => typeof value === "boolean");
    if (!hasBoolean) continue;

    for (let c = 0; c < values[r].length; c++) {
      if (typeof values[r][c] === "boolean") {
        sheet.getRange(r + 1, c + 1).insertCheckboxes();
      }
    }
  }
}

// auto-resizes every row and column so content is never cut off or overly spaced
function resizeToFit(sheet, dataRange) {
  const lastRow = dataRange.getLastRow();
  const lastCol = dataRange.getLastColumn();

  sheet.autoResizeRows(1, lastRow);
  for (let c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }
}
