function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSV Download')
    .addItem('Download CSV', 'showTabSelector')
    .addToUi();
}

function showTabSelector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  
  // In the showTabSelector() function, modify the HTML part:

const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        select { margin: 10px 0; padding: 8px; width: 100%; }
        button { 
          margin: 20px 0 10px 0; 
          padding: 8px; 
          width: 100%;
          background-color: #4285f4; 
          color: white; 
          border: none; 
          cursor: pointer; 
        }
        button:hover { background-color: #357abd; }
        label { display: block; margin-top: 10px; }

        .loader {
          border: 4px solid #f3f3f3;
          border-radius: 50%;
          border-top: 4px solid #4285f4;
          width: 30px;
          height: 30px;
          animation: spin 1s linear infinite;
          margin: 10px auto;
          display: none;
        }

        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        .loading-text {
          text-align: center;
          color: #666;
          margin: 10px 0;
          display: none;
        }
      </style>
    </head>
    <body>
      <div>
        <label for="sheetSelect">Select Sheet:</label>
        <select id="sheetSelect" onchange="getPayDates()">
          ${sheetNames.map(name => '<option value="' + name + '">' + name + '</option>').join('')}
        </select>
      </div>
      <div>
        <label for="payDateSelect">Select Pay Date:</label>
        <select id="payDateSelect" disabled>
          <option value="">First select a sheet</option>
        </select>
      </div>
      <div>
        <label for="payTypeSelect">Select Pay Type:</label>
        <select id="payTypeSelect">
          <option value="year">year</option>
          <option value="hour">hour</option>
        </select>
      </div>
      <div>
        <label for="statusSelect">Select Payroll Status:</label>
        <select id="statusSelect">
          <option value="pending">Pending</option>
          <option value="initiated">Initiated</option>
          <option value="paid">Paid</option>
        </select>
      </div>
      <button onclick="downloadCSV()">Download CSV</button>
      <div id="loader" class="loader"></div>
      <div id="loadingText" class="loading-text">Generating CSV file...</div>

      
      <script>
        function getPayDates() {
          const selectedSheet = document.getElementById('sheetSelect').value;
          const payDateSelect = document.getElementById('payDateSelect');
          payDateSelect.disabled = true;
          payDateSelect.innerHTML = '<option value="">Loading dates...</option>';
          
          google.script.run
            .withSuccessHandler(function(dates) {
              payDateSelect.innerHTML = '';
              dates.forEach(function(date) {
                const option = document.createElement('option');
                option.value = date;
                option.text = date;
                payDateSelect.appendChild(option);
              });
              payDateSelect.disabled = false;
            })
            .withFailureHandler(function(error) {
              alert('Error: ' + error.message);
              payDateSelect.innerHTML = '<option value="">Error loading dates</option>';
            })
            .getUniqueDates(selectedSheet);
        }
        
        function downloadCSV() {
          const selectedSheet = document.getElementById('sheetSelect').value;
          const selectedDate = document.getElementById('payDateSelect').value;
          const selectedPayType = document.getElementById('payTypeSelect').value;
          const selectedStatus = document.getElementById('statusSelect').value;
          
          // Show loader
          document.getElementById('loader').style.display = 'block';
          document.getElementById('loadingText').style.display = 'block';
          
          google.script.run
            .withSuccessHandler(function(csvContent) {
              // Hide loader
              document.getElementById('loader').style.display = 'none';
              document.getElementById('loadingText').style.display = 'none';
              
              if (!csvContent) {
                alert('No data found for selected criteria');
                return;
              }
              const blob = new Blob([csvContent], {type: 'text/csv'});
              const url = window.URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = 'PayDate_' + selectedDate.replace(/[\/\s,]/g, '_') + '_' + selectedPayType + '.csv';
              document.body.appendChild(a);
              a.click();
              document.body.removeChild(a);
              window.URL.revokeObjectURL(url);
            })
            .withFailureHandler(function(error) {
              // Hide loader on error
              document.getElementById('loader').style.display = 'none';
              document.getElementById('loadingText').style.display = 'none';
              alert('Error downloading CSV: ' + error.message);
            })
            .generateCSV(selectedSheet, selectedDate, selectedStatus, selectedPayType);
        }
      </script>
    </body>
  </html>
`;
  
  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Select Sheet and Pay Date');
}

function getUniqueDates(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const payDates = sheet.getRange('J1:J' + lastRow).getValues()
    .flat()
    .filter(date => date instanceof Date && !isNaN(date) && date.toString() !== '')
    .map(date => Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy'));
  
  return [...new Set(payDates)].sort((a, b) => new Date(b) - new Date(a));
}

function generateCSV(sheetName, selectedDate, selectedStatus, selectedPayType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  const xMarkers = data[0];
  const validColumnIndices = xMarkers.map((value, index) => 
    value === 'X' ? index : -1).filter(index => index !== -1);
  
  const headers = validColumnIndices.map(index => {
    const headerCell = data[2][index];
    return headerCell ? headerCell.toString().trim() : '';
  });
  
  headers.unshift('Contract Ref Number', 'Status');
  
  const filteredData = data.slice(3).filter(row => {
    const payDate = row[9];
    const payType = row[30].toString().toLowerCase(); // Column AE (index 30)
    if (!(payDate instanceof Date)) return false;
    const formattedRowDate = Utilities.formatDate(payDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
    return formattedRowDate === selectedDate && payType.includes(selectedPayType.toLowerCase());
  });
  
  const filteredRows = filteredData.map(row => {
    const rowData = validColumnIndices.map((index, colIndex) => {
      const cell = row[index];
      const header = headers[colIndex + 2];
      
      const columnLetter = getColumnLetter(index + 1);
      if (columnLetter === 'AT' || columnLetter === 'AV') {
        if (typeof cell === 'number') {
          const decimalStr = cell.toString().split('.');
          if (decimalStr.length > 1) {
            return cell.toFixed(Math.min(8, decimalStr[1].length));
          }
          return cell.toString();
        }
      }
      
      if (typeof cell === 'number') {
        const decimalStr = cell.toString().split('.');
        if (decimalStr.length > 1) {
          return cell.toFixed(Math.min(2, decimalStr[1].length));
        }
        return cell.toString();
      }
      
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      }
      
      if (cell === null || cell === undefined || cell === '') {
        return '';
      }
      
      if (typeof cell === 'string' && cell.toLowerCase().includes('hourly')) {
        return cell;
      }
      
      return typeof cell === 'string' ? 
        '"' + cell.replace(/"/g, '""').trim() + '"' : 
        cell.toString().trim();
    });
    
    const originalRef = row[3].toString();
    const contractRef = originalRef.replace('PSM', 'CNT').replace(/-[^-]*$/, '');
    
    rowData.unshift('"' + contractRef + '"', '"' + selectedStatus + '"');
    return rowData;
  });
  
  if (filteredRows.length === 0) {
    throw new Error('No data found for selected date and pay type');
  }
  
  let csvContent = headers.map(header => '"' + header + '"').join(',') + '\n';
  csvContent += filteredRows.map(row => row.join(',')).join('\n');
  
  return csvContent;
}

function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
