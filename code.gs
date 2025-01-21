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
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          select, button { margin: 10px 0; padding: 8px; width: 100%; }
          button { background-color: #4285f4; color: white; border: none; cursor: pointer; }
          button:hover { background-color: #357abd; }
          .hidden { display: none; }
          label { display: block; margin-top: 10px; }
        </style>
      </head>
      <body>
        <div>
          <label for="sheetSelect">Select Sheet:</label>
          <select id="sheetSelect">
            ${sheetNames.map(name => '<option value="' + name + '">' + name + '</option>').join('')}
          </select>
          <button onclick="getPayDates()">Next</button>
        </div>
        <div id="payDatesDiv" class="hidden">
          <label for="payDateSelect">Select Pay Date:</label>
          <select id="payDateSelect"></select>
          <button onclick="showStatusSelect()">Next</button>
        </div>
        <div id="statusDiv" class="hidden">
          <label for="statusSelect">Select Payroll Status:</label>
          <select id="statusSelect">
            <option value="pending">Pending</option>
            <option value="initiated">Initiated</option>
            <option value="paid">Paid</option>
          </select>
          <button onclick="downloadCSV()">Download CSV</button>
        </div>
        
        <script>
          function getPayDates() {
            const selectedSheet = document.getElementById('sheetSelect').value;
            google.script.run
              .withSuccessHandler(function(dates) {
                const select = document.getElementById('payDateSelect');
                select.innerHTML = '';
                dates.forEach(function(date) {
                  const option = document.createElement('option');
                  option.value = date;
                  option.text = date;
                  select.appendChild(option);
                });
                document.getElementById('payDatesDiv').classList.remove('hidden');
              })
              .withFailureHandler(function(error) {
                alert('Error: ' + error.message);
              })
              .getUniqueDates(selectedSheet);
          }
          
          function showStatusSelect() {
            document.getElementById('statusDiv').classList.remove('hidden');
          }
          
          function downloadCSV() {
            const selectedSheet = document.getElementById('sheetSelect').value;
            const selectedDate = document.getElementById('payDateSelect').value;
            const selectedStatus = document.getElementById('statusSelect').value;
            google.script.run
              .withSuccessHandler(function(csvContent) {
                if (!csvContent) {
                  alert('No data found for selected date');
                  return;
                }
                const blob = new Blob([csvContent], {type: 'text/csv'});
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'PayDate_' + selectedDate.replace(/[\/\s,]/g, '_') + '.csv';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
              })
              .withFailureHandler(function(error) {
                alert('Error downloading CSV: ' + error.message);
              })
              .generateCSV(selectedSheet, selectedDate, selectedStatus);
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

function generateCSV(sheetName, selectedDate, selectedStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  const xMarkers = data[0];
  const validColumnIndices = xMarkers.map((value, index) => 
    value === 'X' ? index : -1).filter(index => index !== -1);
  
  const headers = validColumnIndices.map(index => {
    const headerCell = data[2][index];
    return headerCell ? headerCell.toString().trim() : '';
  });
  
  // Add Status and Contract Ref Number as the first two headers
  headers.unshift('Contract Ref Number', 'Status');
  
  const filteredData = data.slice(3).filter(row => {
    const payDate = row[9];
    if (!(payDate instanceof Date)) return false;
    const formattedRowDate = Utilities.formatDate(payDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
    return formattedRowDate === selectedDate;
  });
  
  const filteredRows = filteredData.map(row => {
    const rowData = validColumnIndices.map((index, colIndex) => {
      const cell = row[index];
      const header = headers[colIndex + 2]; // +2 because we added two columns at beginning
      
      // Handle columns AT and AV (up to 8 decimal places)
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
      
      // Handle FX Rate and other numerical values (up to 2 decimal places)
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
    
    // Get the value from Column D (index 3) and transform it
    const originalRef = row[3].toString();
    const contractRef = originalRef.replace('PSM', 'CNT').replace(/-[^-]*$/, '');
    
    // Add Contract Ref Number and Status as the first two columns
    rowData.unshift('"' + contractRef + '"', '"' + selectedStatus + '"');
    return rowData;
  });
  
  if (filteredRows.length === 0) {
    throw new Error('No data found for selected date');
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
