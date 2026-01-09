function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Ledger");

  if (!ledgerSheet) {
    return HtmlService.createHtmlOutput("<h1>Error: Ledger sheet not found.</h1>");
  }

  const lastRow = ledgerSheet.getLastRow();
  const currentBalanceValue = lastRow < 2 ? 0 : ledgerSheet.getRange(lastRow, 4).getValue();
  const currentBalance = typeof currentBalanceValue === 'number' ? currentBalanceValue.toFixed(2) : "0.00";

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
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        * {
          box-sizing: border-box;
        }

        body {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          margin: 0;
          background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
          color: #1a1a2e;
        }

        .container {
          text-align: center;
          background: #ffffff;
          padding: 48px 40px;
          border-radius: 20px;
          box-shadow: 0 4px 24px rgba(0, 0, 0, 0.08);
          width: 90%;
          max-width: 380px;
        }

        h1 {
          color: #1a1a2e;
          font-size: 22px;
          font-weight: 600;
          margin: 0 0 8px 0;
          letter-spacing: -0.3px;
        }

        .balance-label {
          font-size: 13px;
          color: #6b7280;
          text-transform: uppercase;
          letter-spacing: 1px;
          font-weight: 500;
          margin-bottom: 4px;
        }

        .balance-wrapper {
          margin: 24px 0 32px 0;
        }

        .balance {
          font-size: 52px;
          color: #1a1a2e;
          font-weight: 700;
          letter-spacing: -2px;
          line-height: 1;
        }

        .balance::before {
          content: '$';
          font-size: 28px;
          font-weight: 500;
          vertical-align: top;
          margin-right: 2px;
          color: #6b7280;
          letter-spacing: 0;
        }

        .actions {
          display: flex;
          justify-content: center;
          gap: 12px;
        }

        .button {
          border: none;
          color: white;
          padding: 14px 0;
          text-align: center;
          text-decoration: none;
          font-size: 15px;
          font-weight: 600;
          cursor: pointer;
          border-radius: 12px;
          transition: all 0.2s ease;
          flex: 1;
          letter-spacing: -0.2px;
        }

        .button:active {
          transform: scale(0.98);
        }

        .deposit {
          background: #22c55e;
        }
        .deposit:hover {
          background: #16a34a;
        }

        .withdraw {
          background: #64748b;
        }
        .withdraw:hover {
          background: #475569;
        }

        .cancel {
          background: #e2e8f0;
          color: #64748b;
        }
        .cancel:hover {
          background: #cbd5e1;
        }

        .button:disabled {
          background: #e2e8f0;
          color: #94a3b8;
          cursor: not-allowed;
          transform: none;
        }

        .transaction-form {
          display: none;
          margin-top: 24px;
        }

        input[type="number"] {
          width: 100%;
          padding: 14px 16px;
          margin-bottom: 12px;
          border: 2px solid #e2e8f0;
          border-radius: 12px;
          font-size: 16px;
          font-family: inherit;
          transition: border-color 0.2s ease;
          outline: none;
        }

        input[type="number"]:focus {
          border-color: #22c55e;
        }

        input[type="number"]::placeholder {
          color: #94a3b8;
        }

        .form-buttons {
          display: flex;
          gap: 10px;
        }

        .message {
          margin-top: 20px;
          font-size: 14px;
          font-weight: 500;
          min-height: 20px;
        }

        .success { color: #16a34a; }
        .error { color: #dc2626; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>${title}</h1>
        <div class="balance-wrapper">
          <div class="balance-label">Current Balance</div>
          <div class="balance">${currentBalance}</div>
        </div>
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

