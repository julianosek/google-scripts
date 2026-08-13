// colour set per research type: title/criteria row background + font colours
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

// applies the theme's background and font colours to the title and criteria rows
function colourSheet(sheet, theme) {
  sheet.getRange(TITLE_ROW + ":" + TITLE_ROW).setBackground(theme.titleBg).setFontColor(theme.titleFont);
  sheet.getRange(CRITERIA_ROW + ":" + CRITERIA_ROW).setBackground(theme.criteriaBg).setFontColor(theme.criteriaFont);
}

// wraps, centers, and bolds text across the table
function formatText(sheet, dataRange) {
  dataRange
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, HEADER_ROWS, sheet.getMaxColumns()).setFontWeight("bold"); // bold the header rows
}

// lets the title row spill into empty neighbouring cells instead of wrapping
function overflowTitleRow(sheet, dataRange) {
  sheet
    .getRange(TITLE_ROW, 1, 1, dataRange.getLastColumn())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
}