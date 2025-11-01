function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");

  if (!ledgerSheet) {
    return HtmlService.createHtmlOutput("<h1>Error: Ledger sheet not found.</h1>");
  }

  const lastRow = ledgerSheet.getLastRow();
  const currentBalance = ledgerSheet.getRange(lastRow, 4).getDisplayValue();

  const htmlOutput = `
    <style>
      body { font-family: Arial, sans-serif; text-align: center; padding-top: 50px; background-color: #f0f0f0; }
      h1 { color: #4a4a4a; }
      .balance { font-size: 48px; color: #2c8b2c; font-weight: bold; margin-top: 20px; margin-bottom: 30px;}
      .button {
        border: none; color: white; padding: 15px 32px; text-align: center;
        text-decoration: none; display: inline-block; font-size: 16px;
        margin: 4px 10px; cursor: pointer; border-radius: 8px; transition: background-color 0.3s;
      }
      .deposit { background-color: #4CAF50; /* Green */ }
      .deposit:hover { background-color: #45a049; }
      .withdraw { background-color: #f44336; /* Red */ }
      .withdraw:hover { background-color: #da190b; }
      .button:disabled { background-color: #cccccc; cursor: not-allowed; }
    </style>
    <h1>Your Piggy Bank's Current Balance</h1>
    <div class="balance">${currentBalance}</div>
    <button class="button deposit" id="depositBtn" onclick="showDepositDialog()">Deposit</button>
    <button class="button withdraw" id="withdrawBtn" onclick="showWithdrawDialog()">Withdraw</button>
    
    <script>
      function showDepositDialog() {
        const amount = promptForAmount("Please enter the amount to deposit:");
        if (amount) {
          setButtonsDisabled(true);
          google.script.run
            .withSuccessHandler(onTransactionSuccess)
            .withFailureHandler(onTransactionFailure)
            .addDepositRow(amount);
        }
      }
      
      function showWithdrawDialog() {
        const amount = promptForAmount("Please enter the amount to withdraw:");
        if (amount) {
          setButtonsDisabled(true);
          google.script.run
            .withSuccessHandler(onTransactionSuccess)
            .withFailureHandler(onTransactionFailure)
            .addWithdrawalRow(amount);
        }
      }

      function promptForAmount(message) {
        const amountString = prompt(message, "10.00");
        if (amountString === null || amountString.trim() === "") return null;
        
        const amount = parseFloat(amountString);
        if (isNaN(amount) || amount <= 0) {
          alert("Invalid amount. Please enter a positive number.");
          return null;
        }
        return amount;
      }
      
      function onTransactionSuccess(newBalance) {
        document.querySelector('.balance').textContent = newBalance;
        alert("Transaction successful!");
        setButtonsDisabled(false);
      }
      
      function onTransactionFailure(error) {
        alert("Transaction failed: " + error.message);
        setButtonsDisabled(false);
      }

      function setButtonsDisabled(disabled) {
          document.getElementById('depositBtn').disabled = disabled;
          document.getElementById('withdrawBtn').disabled = disabled;
      }
    </script>
  `;
  
  return HtmlService.createHtmlOutput(htmlOutput).setTitle('Piggy Bank Balance');
}
