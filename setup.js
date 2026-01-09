/**
 * Creates a time-driven trigger for the addInterestRow function.
 * This function is designed to be run manually to set up the trigger.
 * It will delete all existing triggers to prevent duplicates.
 */
function setupTriggers() {
  // Delete all existing triggers for this project.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }

  // Create a new trigger for the addInterestRow function.
  ScriptApp.newTrigger('addInterestRow')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(3) // Runs between 3am and 4am
      .create();
}

/**
 * Initializes the "Ledger" and "Configuration" sheets with headers.
 * This function is designed to be run manually.
 */
function initSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Ledger Sheet
  var ledgerSheet = spreadsheet.getSheetByName("Ledger");
  if (!ledgerSheet) {
    var sheet1 = spreadsheet.getSheetByName("Sheet1");
    if (sheet1 && sheet1.getLastRow() === 0 && sheet1.getLastColumn() === 0) {
      ledgerSheet = sheet1.setName("Ledger");
    } else {
      ledgerSheet = spreadsheet.insertSheet("Ledger");
    }
  }
  ledgerSheet.getRange("A1:D1").setValues([["Date", "Type", "Amount", "Balance"]]);

  // Bootstrap with an initial deposit if the sheet is new or empty except for header
  if (ledgerSheet.getLastRow() < 2) {
    ledgerSheet.appendRow([new Date(), "Deposit", 0, 0]);
  }

  // Setup Configuration Sheet
  var configSheet = spreadsheet.getSheetByName("Configuration");
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet("Configuration");
  }
  configSheet.getRange("A1:B5").setValues([
    ["APY", "65%"],
    ["Compounding Frequency (days)", 7],
    ["Interest rate per period", 0.01246575342],
    ["Title", "Piggy Bank"],
    ["Max Balance", 60]
  ]);
}

/**
 * Runs all setup functions to initialize the application.
 */
function initApp() {
  initSheets();
  setupTriggers();
}
