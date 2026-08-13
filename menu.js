// runs when a user installs the add-on; sets up the menu right away
function onInstall(e) {
  onOpen(e);
}

// adds the add-on menu when the spreadsheet opens
function onOpen(e) {
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