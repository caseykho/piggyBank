/**
 * Finds the last row in the "Ledger" sheet and adds a new row
 * to calculate the latest interest payment.
 */
function addInterestRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");
  const configSheet = ss.getSheetByName("Configuration");

  if (!ledgerSheet) {
    throw new Error('Error: Sheet named "Ledger" could not be found.');
  }
  if (!configSheet) {
    throw new Error('Error: Sheet named "Configuration" could not be found.');
  }

  const lastRow = ledgerSheet.getLastRow();
  const newRow = lastRow + 1;

  // --- Get values for calculation ---
  const interestRate = configSheet.getRange("B3").getValue();
  if (typeof interestRate !== 'number' || interestRate <= 0) {
    throw new Error('Invalid interest rate. Please check the Configuration sheet.');
  }

  const lastBalance = ledgerSheet.getRange(lastRow, 4).getValue();
  if (typeof lastBalance !== 'number') {
    throw new Error('Could not read the last balance from the Ledger sheet.');
  }

  // --- Perform calculations ---
  const interestAmount = parseFloat((lastBalance * interestRate).toFixed(2));
  const newBalance = parseFloat((lastBalance + interestAmount).toFixed(2));

  // --- Prepare the data for the new row ---
  const currentDate = new Date();
  const type = "Interest";

  // --- Write the data into the new row ---
  ledgerSheet.getRange(newRow, 1).setValue(currentDate);      // Column A: Date
  ledgerSheet.getRange(newRow, 2).setValue(type);             // Column B: Type
  ledgerSheet.getRange(newRow, 3).setValue(interestAmount);   // Column C: Amount
  ledgerSheet.getRange(newRow, 4).setValue(newBalance);       // Column D: Balance
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
  const newBalanceValue = ledgerSheet.getRange(newRow, 4).getValue();
  return newBalanceValue.toFixed(2);
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
