// Generate menu when spreadsheet is opened
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Reports')
      .addItem('EOM Report', 'main')
      .addToUi();
}