function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");

  if (!ledgerSheet) {
    return HtmlService.createHtmlOutput("<h1>Error: Ledger sheet not found.</h1>");
  }

  const lastRow = ledgerSheet.getLastRow();
  const currentBalanceValue = ledgerSheet.getRange(lastRow, 4).getValue();
  const currentBalance = typeof currentBalanceValue === 'number' ? currentBalanceValue.toFixed(2) : currentBalanceValue;

  const configSheet = ss.getSheetByName("Configuration");
  let title = "Piggy Bank"; // Default title
  if (configSheet) {
    title = configSheet.getRange("B4").getValue() || title;
  }

  const htmlOutput = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');
        body { 
          font-family: 'Roboto', sans-serif; 
          display: flex;
          flex-direction: column;
          align-items: center; 
          justify-content: center;
          height: 100vh;
          margin: 0;
          background-color: #f7f8fa;
          color: #333;
        }
        .container {
          text-align: center;
          background: white;
          padding: 40px;
          border-radius: 12px;
          box-shadow: 0 8px 16px rgba(0,0,0,0.1);
          width: 90%;
          max-width: 400px;
        }
        h1 { 
          color: #2c3e50; 
          margin-bottom: 10px;
        }
        .balance { 
          font-size: 48px; 
          color: #27ae60; 
          font-weight: bold; 
          margin-top: 10px; 
          margin-bottom: 30px;
        }
        .actions {
          display: flex;
          justify-content: center;
          gap: 15px;
          margin-bottom: 30px;
        }
        .button {
          border: none; 
          color: white; 
          padding: 15px 0;
          text-align: center;
          text-decoration: none; 
          font-size: 16px;
          font-weight: bold;
          cursor: pointer; 
          border-radius: 8px; 
          transition: background-color 0.3s, box-shadow 0.3s;
          flex: 1;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .deposit { background-color: #2ecc71; }
        .deposit:hover { background-color: #27ae60; }
        .withdraw { background-color: #e74c3c; }
        .withdraw:hover { background-color: #c0392b; }
        .cancel { background-color: #95a5a6; }
        .cancel:hover { background-color: #7f8c8d; }
        .button:disabled { 
          background-color: #bdc3c7; 
          cursor: not-allowed; 
          box-shadow: none;
        }
        .transaction-form {
          display: none;
          margin-top: 20px;
        }
        input[type="number"] {
          width: calc(100% - 24px);
          padding: 12px;
          margin-bottom: 15px;
          border: 1px solid #ddd;
          border-radius: 8px;
          font-size: 16px;
        }
        .form-buttons {
          display: flex;
          gap: 10px;
        }
        .message {
          margin-top: 20px;
          font-size: 14px;
          min-height: 20px;
        }
        .success { color: #27ae60; }
        .error { color: #e74c3c; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>${title}</h1>
        <div class="balance">${currentBalance}</div>
        <div class="actions">
          <button class="button deposit" id="depositBtn" onclick="showTransactionForm('deposit')">Deposit</button>
          <button class="button withdraw" id="withdrawBtn" onclick="showTransactionForm('withdraw')">Withdraw</button>
        </div>
        <div class="transaction-form" id="transactionForm">
          <form id="amountForm">
            <input type="number" id="amountInput" placeholder="Enter amount" step="0.01" min="0.01">
            <div class="form-buttons">
              <button type="submit" class="button" id="submitBtn">Submit</button>
              <button type="button" class="button cancel" onclick="hideTransactionForm()">Cancel</button>
            </div>
          </form>
        </div>
        <div class="message" id="messageArea"></div>
      </div>
      
      <script>
        let currentTransactionType = '';

        function showTransactionForm(type) {
          currentTransactionType = type;
          document.getElementById('transactionForm').style.display = 'block';
          document.querySelector('.actions').style.display = 'none';
          document.getElementById('submitBtn').className = 'button deposit';
          document.getElementById('amountInput').focus();
        }

        function hideTransactionForm() {
          document.getElementById('transactionForm').style.display = 'none';
          document.querySelector('.actions').style.display = 'flex';
          document.getElementById('amountInput').value = '';
          clearMessage();
        }

        function submitTransaction() {
          const amountInput = document.getElementById('amountInput');
          const amount = parseFloat(amountInput.value);

          if (isNaN(amount) || amount <= 0) {
            showMessage("Please enter a positive number.", "error");
            return;
          }

          setButtonsDisabled(true);
          const action = currentTransactionType === 'deposit' ? 'addDepositRow' : 'addWithdrawalRow';
          
          google.script.run
            .withSuccessHandler(onTransactionSuccess)
            .withFailureHandler(onTransactionFailure)[action](amount);
        }
        
        function onTransactionSuccess(newBalance) {
          document.querySelector('.balance').textContent = newBalance;
          showMessage("Transaction successful!", "success");
          setButtonsDisabled(false);
          setTimeout(hideTransactionForm, 1500);
        }
        
        function onTransactionFailure(error) {
          showMessage("Error: " + error.message, "error");
          setButtonsDisabled(false);
        }

        function setButtonsDisabled(disabled) {
            document.getElementById('depositBtn').disabled = disabled;
            document.getElementById('withdrawBtn').disabled = disabled;
            document.getElementById('submitBtn').disabled = disabled;
        }

        function showMessage(msg, type) {
          const messageArea = document.getElementById('messageArea');
          messageArea.textContent = msg;
          messageArea.className = 'message ' + type;
        }

        function clearMessage() {
          const messageArea = document.getElementById('messageArea');
          messageArea.textContent = '';
          messageArea.className = 'message';
        }

        document.getElementById('amountForm').addEventListener('submit', function(event) {
          event.preventDefault();
          submitTransaction();
        });
      </script>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(htmlOutput)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

