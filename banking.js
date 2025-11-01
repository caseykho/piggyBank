/**
 * Finds the last row in the "Ledger" sheet and adds a new row
 * to calculate the latest interest payment.
 */
function addInterestRow() {
  // Get the active spreadsheet and then the specific sheet named "Ledger".
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");

  // Stop the script if the "Ledger" sheet doesn't exist.
  if (!ledgerSheet) {
    throw new Error('Error: Sheet named "Ledger" could not be found.');
  }

  // Get the number of the last row that contains data.
  const lastRow = ledgerSheet.getLastRow();
  // The new row will be the next one down.
  const newRow = lastRow + 1;

  // --- Prepare the data for the new row ---

  // 1. Get the current date for the 'Date' column.
  const currentDate = new Date();

  // 2. Define the 'Type' column.
  const type = "Interest";

  // 3. Create the formula for the 'Amount' column.
  // This references the balance (column D) in the previous row.
  const amountFormula = `=D${lastRow}*Configuration!$B$3`;

  // 4. Create the formula for the 'Balance' column.
  // This adds the new interest amount (column C of the new row)
  // to the previous balance (column D of the last row).
  const balanceFormula = `=D${lastRow}+C${newRow}`;

  // --- Write the data into the new row ---
  ledgerSheet.getRange(newRow, 1).setValue(currentDate); // Column A: Date
  ledgerSheet.getRange(newRow, 2).setValue(type);        // Column B: Type
  ledgerSheet.getRange(newRow, 3).setFormula(amountFormula); // Column C: Amount
  ledgerSheet.getRange(newRow, 4).setFormula(balanceFormula);  // Column D: Balance
}

/**
 * A private helper function to add a new transaction row to the ledger.
 * @param {string} type The type of transaction (e.g., "Deposit", "Withdrawal").
 * @param {number} amount The positive transaction amount.
 * @param {string} operator The mathematical operator for the balance formula ('+' or '-').
 * @return {string} The new balance as a string.
 */
function _addLedgerEntry(type, amount, operator) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");

  if (!ledgerSheet) {
    throw new Error('Sheet named "Ledger" could not be found.');
  }
  
  if (typeof amount !== 'number' || amount <= 0) {
      throw new Error('Invalid transaction amount provided.');
  }

  const lastRow = ledgerSheet.getLastRow();
  const newRow = lastRow + 1;

  const currentDate = new Date();
  const balanceFormula = `=D${lastRow}${operator}C${newRow}`;

  // --- Write the data into the new row ---
  ledgerSheet.getRange(newRow, 1).setValue(currentDate); // Column A: Date
  ledgerSheet.getRange(newRow, 2).setValue(type);        // Column B: Type
  ledgerSheet.getRange(newRow, 3).setValue(amount);      // Column C: Amount
  ledgerSheet.getRange(newRow, 4).setFormula(balanceFormula);  // Column D: Balance
  
  // Ensure all pending changes are applied so we can read the new value.
  SpreadsheetApp.flush();
  
  // Get the newly calculated balance and return it.
  const newBalance = ledgerSheet.getRange(newRow, 4).getDisplayValue();
  return newBalance;
}

/**
 * Appends a new row for a deposit. Calls the private helper function.
 * @param {number} amount The amount to deposit.
 * @return {string} The new balance as a string.
 */
function addDepositRow(amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configurationSheet = ss.getSheetByName("Configuration");
  if (!configurationSheet) {
    throw new Error('Sheet named "Configuration" could not be found.');
  }
  const maxBalanceRange = configurationSheet.getRange("B5");
  const maxBalance = maxBalanceRange.getValue();

  const ledgerSheet = ss.getSheetByName("Ledger");
  if (!ledgerSheet) {
    throw new Error('Sheet named "Ledger" could not be found.');
  }
  const lastRow = ledgerSheet.getLastRow();
  const currentBalanceRange = ledgerSheet.getRange(lastRow, 4);
  const currentBalance = currentBalanceRange.getValue();

  if (currentBalance + amount > maxBalance) {
    throw new Error('Deposit failed: This transaction would exceed the maximum balance of ' + maxBalance);
  }

  return _addLedgerEntry("Deposit", amount, "+");
}

/**
 * Appends a new row for a withdrawal. Calls the private helper function.
 * @param {number} amount The amount to withdraw.
 * @return {string} The new balance as a string.
 */
function addWithdrawalRow(amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");
  if (!ledgerSheet) {
    throw new Error('Sheet named "Ledger" could not be found.');
  }
  const lastRow = ledgerSheet.getLastRow();
  const currentBalanceRange = ledgerSheet.getRange(lastRow, 4);
  const currentBalance = currentBalanceRange.getValue();

  if (amount > currentBalance) {
    throw new Error('Withdrawal failed: Insufficient funds. You tried to withdraw ' + amount + ' but your balance is ' + currentBalance);
  }

  return _addLedgerEntry("Withdrawal", amount, "-");
}
