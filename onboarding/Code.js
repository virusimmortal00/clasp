function addRowBelowButton() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var rowIndex = activeCell.getRow();
  
  // Insert a new row below the active row
  sheet.insertRowAfter(rowIndex);

  // Get the new row index
  var newRowIndex = rowIndex + 1;

  // Get current date and editor
  var currentDate = new Date();
  var editorName = Session.getActiveUser().getEmail(); // Retrieves the email of the active user

  // Adjust these column indices based on where you want the date and editor
  var dateColumn = 8; // Column A (1-indexed)
  var editorColumn = 9; // Column B (1-indexed)

  // Set date and editor in the new row
  sheet.getRange(newRowIndex, dateColumn).setValue(currentDate);
  sheet.getRange(newRowIndex, editorColumn).setValue(editorName);
}

function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return; // No data to process

  // Define the range starting from row 2
  var numRows = lastRow - 1;
  var dataRange = sheet.getRange(2, 1, numRows, 6); // Columns A to F
  var data = dataRange.getValues();

  // Create a map to store the latest date for each account name
  var accountLatestDate = {};

  // First, find the latest date for each account name
  for (var i = 0; i < data.length; i++) {
    var accountName = data[i][0]; // Column A
    var dateValue = data[i][5];   // Column F

    // Check if account name and date are valid
    if (accountName && dateValue instanceof Date && !isNaN(dateValue)) {
      var existingDate = accountLatestDate[accountName];
      if (!existingDate || dateValue > existingDate) {
        accountLatestDate[accountName] = dateValue;
      }
    }
  }

  // Now, update Column C with the latest date for each account name
  var updates = [];
  for (var i = 0; i < data.length; i++) {
    var accountName = data[i][0]; // Column A
    data[i][2] = accountLatestDate[accountName] || ''; // Column C
    updates.push([data[i][2]]);
  }

  // Write the updated dates back to Column C
  var updateRange = sheet.getRange(2, 3, updates.length, 1); // Column C
  updateRange.setValues(updates);
}
