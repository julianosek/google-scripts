const HEADER_ROWS = 2; // number of header rows (title + criteria)
const TITLE_ROW = 1;
const CRITERIA_ROW = 2;

const TITLE_BG_COLOUR = "#0b5394";
const TITLE_FONT_COLOUR = "#ffffff";
const CRITERIA_BG_COLOUR = "#7babf8";
const CRITERIA_FONT_COLOUR = "#073763";

// runs the full formatting pipeline on the active sheet
function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSheet(); //get the currently active sheet

  freezeSheet(sheet);
  colourSheet(sheet);
  trimSheet(sheet);

  const dataRange = sheet.getDataRange(); // recompute after trimming so it reflects the final table bounds
  formatText(sheet, dataRange);
  addCheckboxes(sheet, dataRange);
  resizeToFit(sheet, dataRange);
}

// freezes the header rows and first column
function freezeSheet(sheet) {
  sheet.setFrozenRows(HEADER_ROWS); // freeze header rows
  sheet.setFrozenColumns(1); // freeze first column
}

// applies background and font colours to the title and criteria rows
function colourSheet(sheet) {
  sheet.getRange(TITLE_ROW + ":" + TITLE_ROW).setBackground(TITLE_BG_COLOUR).setFontColor(TITLE_FONT_COLOUR);
  sheet.getRange(CRITERIA_ROW + ":" + CRITERIA_ROW).setBackground(CRITERIA_BG_COLOUR).setFontColor(CRITERIA_FONT_COLOUR);
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

// wraps, centers, and bolds text across the table
function formatText(sheet, dataRange) {
  dataRange
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, HEADER_ROWS, sheet.getMaxColumns()).setFontWeight("bold"); // bold the header rows

  sheet
    .getRange(TITLE_ROW, 1, 1, dataRange.getLastColumn())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // let the title row overflow instead of wrapping
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
