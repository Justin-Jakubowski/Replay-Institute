function myFunction() {
  function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  // Check if the edited column is the "Raw Data" column (adjust column index if necessary)
  if (col == 5) { // Assuming "Raw Data" is in column E (5th column)
    var data = e.value;
    var rangesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ranges");
    var rangeValues = rangesSheet.getRange("A2:C5").getValues(); // Adjust the range as needed

    var category = null;

    // Determine category based on ranges
    for (var i = 0; i < rangeValues.length; i++) {
      if (data >= rangeValues[i][1] && data <= rangeValues[i][2]) {
        category = rangeValues[i][0];
        break;
      }
    }

    // Reset checkboxes
    sheet.getRange(row, 1, 1, 4).uncheck(); // Assuming checkboxes are in columns A to D

    // Check the appropriate checkbox based on the category
    if (category == "Elite") {
      sheet.getRange(row, 1).check(); // Check "Elite" column
    } else if (category == "Good") {
      sheet.getRange(row, 2).check(); // Check "Good" column
    } else if (category == "Average") {
      sheet.getRange(row, 3).check(); // Check "Average" column
    } else if (category == "Bad") {
      sheet.getRange(row, 4).check(); // Check "Bad" column
    }
  }
}
}
