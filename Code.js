const HEADER_ROWS = 2; // number of header rows (title + criteria)
const TITLE_ROW = 1;
const CRITERIA_ROW = 2;

// colour theme per research type: title/criteria row background + font colours
const THEMES = {
  PRODUCT: {
    titleBg: "#0b5394",
    titleFont: "#ffffff",
    criteriaBg: "#7babf8",
    criteriaFont: "#073763",
  },
  SERVICE: {
    titleBg: "#674ea7",
    titleFont: "#ffffff",
    criteriaBg: "#b4a7d6",
    criteriaFont: "#20124d",
  },
  TRAVEL: {
    titleBg: "#ffff00",
    titleFont: "#000000",
    criteriaBg: "#fffeb3",
    criteriaFont: "#000000",
  },
};

// adds the custom menu when the spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Sheet Tools")
    .addSubMenu(
      ui
        .createMenu("Apply Formatting")
        .addItem("Product Research", "formatProductResearch")
        .addItem("Service Research", "formatServiceResearch")
        .addItem("Travel Research", "formatTravelResearch")
    )
    .addToUi();
}

// menu entry points, one per research type/theme
function formatProductResearch() {
  formatSheet(THEMES.PRODUCT);
}

function formatServiceResearch() {
  formatSheet(THEMES.SERVICE);
}

function formatTravelResearch() {
  formatSheet(THEMES.TRAVEL);
}

// runs the full formatting pipeline on the active sheet using the given theme
function formatSheet(theme) {
  const sheet = SpreadsheetApp.getActiveSheet(); //get the currently active sheet

  freezeSheet(sheet);
  colourSheet(sheet, theme);
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

// applies the theme's background and font colours to the title and criteria rows
function colourSheet(sheet, theme) {
  sheet.getRange(TITLE_ROW + ":" + TITLE_ROW).setBackground(theme.titleBg).setFontColor(theme.titleFont);
  sheet.getRange(CRITERIA_ROW + ":" + CRITERIA_ROW).setBackground(theme.criteriaBg).setFontColor(theme.criteriaFont);
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
