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