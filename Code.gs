function onOpen(e) {     //creates custom extention for drag and drop functionality
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Open Sidebar', 'openSidebar')
      .addToUi();
}

function openSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CSV Importer');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function importCSV(importData, columns, sheetName) {
  try {
    var spreadsheetId = '1OmlSptVRov_HRGr1yc_pwldMMudIalFpMFQBr-DBEoI'; // SheetId associated with the Apps Script code
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);

    // Check if the sheet exists, if not, create one
    if (!sheet) {
      spreadsheet.insertSheet(sheetName);
      sheet = spreadsheet.getSheetByName(sheetName);
    }

    // Parse importData to a 2D array
    var csvData = importData.map(function(row) {
      return row.map(function(cell) {
        return cell.trim(); // Trim whitespace from each cell
      });
    });

    // Filter columns based on user selection
    var filteredData = csvData.map(function(row) {
      var filteredRow = [];
      columns.forEach(function(column) {
        var columnIndex = csvData[0].indexOf(column.trim());
        if (columnIndex !== -1) {
          filteredRow.push(row[columnIndex]);
        } else {
          filteredRow.push('');
        }
      });
      return filteredRow;
    });

    // Append data to the sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

    return 'Success';
  } catch (error) {
    return 'Error: ' + error.message;
  }
}
