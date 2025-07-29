function EnterToDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database"); // your database sheet

  var row = findAvailableRow();
  Logger.log(row);

  promptAndSaveInfo(); //save name & level to database
  copyLowerBodyToDatabase(row);
  copyHipsToDatabase(row);
  copySpineToDatabase(row);
  copyUpperBodyToDatabase(row);
  copyQualityToDatabase(row);
  calculateMobilityTotalScore(row);
  copyStrengthToDatabase(row);
  copyStrengthScoreToDatabase(row);
  calculateStrengthTotalScore(row);
  copyPowerToDatabase(row);
  copyPowerScoreToDatabase(row);
  calculatePowerTotalScore(row);
  copyRotationalToDatabase(row);
  copyRotationalScoreToDatabase(row);
  calculateRotationalTotalScore(row);
  copyArmStrengthToDatabase(row);
  copyArmStrengthScoreToDatabase(row);
  calculateArmStrengthTotalScore(row);
}

function clearScreeningSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var namedRanges = [
    "Strength",
    "Power",
    "Rotational",
    "ArmStrength"
  ];

  namedRanges.forEach(function(name) {
    var range = ss.getRangeByName(name);
    if (range) {
      range.clearContent(); // or use .clear() to also remove formatting if needed
    } else {
      Logger.log("Named range not found: " + name);
    }
  });

  Logger.log("All specified ranges cleared.");

  setMobilityCheckBoxesToFalse();
  setPowerCheckBoxesToFalse();
  setStrengthCheckBoxesToFalse();
  setRotationalCheckBoxesToFalse();
  setArmStrengthCheckBoxesToFalse();
}


function setMobilityCheckBoxesToFalse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("MobilityCheckBoxes");
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Create a 2D array filled with false
  var falseValues = Array.from({ length: numRows }, () => Array(numCols).fill(false));

  range.setValues(falseValues);
}

function setStrengthCheckBoxesToFalse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("StrengthCheckBoxes");
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Create a 2D array filled with false
  var falseValues = Array.from({ length: numRows }, () => Array(numCols).fill(false));

  range.setValues(falseValues);
}

function setRotationalCheckBoxesToFalse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("RotationalCheckBoxes");
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Create a 2D array filled with false
  var falseValues = Array.from({ length: numRows }, () => Array(numCols).fill(false));

  range.setValues(falseValues);
}

function setPowerCheckBoxesToFalse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("PowerCheckBoxes");
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Create a 2D array filled with false
  var falseValues = Array.from({ length: numRows }, () => Array(numCols).fill(false));

  range.setValues(falseValues);
}

function setArmStrengthCheckBoxesToFalse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("ArmStrengthCheckBoxes");
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  // Create a 2D array filled with false
  var falseValues = Array.from({ length: numRows }, () => Array(numCols).fill(false));

  range.setValues(falseValues);
}

function promptAndSaveInfo() {
  var html = HtmlService.createHtmlOutputFromFile('ExperienceLevel')
    .setWidth(300)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter User Info');
}

function processForm(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");
  var nextRow = findAvailableRow();

  // Log for debug
  Logger.log("Saving to Row: " + nextRow);

  // Save the form data
  sheet.getRange(nextRow, getColumnIndex("First Name")).setValue(formData.firstName);
  sheet.getRange(nextRow, getColumnIndex("Last Name")).setValue(formData.lastName);
  sheet.getRange(nextRow, getColumnIndex("Education Level")).setValue(formData.educationLevel);

  return "Thank you, " + formData.firstName + "'s information has been saved to row " + nextRow;
  //SpreadsheetApp.getUi().alert("Thank you, " + formData.firstName + " has been added to the Database at row " + nextRow);
}

function findAvailableRow(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

  // Get the values in column B (adjust this if needed)
  var columnValues = sheet.getRange("B3:B").getValues(); // Fixed variable name

  // Loop through the column and find the first empty row
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] === "" || columnValues[i][0] === null) {
      return i + 3; // +1 because row index starts at 1
    }
  }

  // If no empty row found, return the next available row
  return columnValues.length + 1;
}


// Utility to find a column by header name
function getColumnIndex(headerName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName) + 1;
}


function copyLowerBodyToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

  var currentColumn = getColumnIndex("Ankle Dorsiflexion"); // Finds the column for "Ankle Dosiflexion"
  
  // Get values from the named range "LowerBody"
  var lowerBodyRange = ss.getRangeByName("LowerBody");
  var lowerBodyValues = lowerBodyRange.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = lowerBodyValues
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyPowerToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

  var currentColumn = getColumnIndex("Lateral Jump"); // Finds the column for "LastName"
  
  // Get values from the named range "Power"
  var range = ss.getRangeByName("Power");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyPowerScoreToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Lateral Jump Score"); // Finds the column for "LastName"
  
  // Get values from the named range "Power"
  var range = ss.getRangeByName("PowerScore");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyRotationalToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Medball Shotput (6oz)"); // Finds the column for "LastName"
  
  // Get values from the named range "Rotational"
  var range = ss.getRangeByName("Rotational");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyRotationalScoreToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Medball Shotput (6oz) Score"); // Finds the column for "LastName"
  
  // Get values from the named range "Rotational"
  var range = ss.getRangeByName("RotationalScore");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyArmStrengthToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("IR Strength"); // Finds the column for "LastName"
  
  // Get values from the named range "ArmStrength"
  var range = ss.getRangeByName("ArmStrength");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyArmStrengthScoreToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

 
  var currentColumn = getColumnIndex("IR Strength Score"); // Finds the column for "LastName"
  
  // Get values from the named range "ArmStrength"
  var range = ss.getRangeByName("ArmStrengthScore");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}


function copyHipsToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Cossack Squat"); // Finds the column for "LastName"
  
  // Get values from the named range "Hips"
  var range = ss.getRangeByName("Hips");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyStrengthToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Squat(5RM)"); // Finds the column for "LastName"
  
  // Get values from the named range "Strength"
  var range = ss.getRangeByName("Strength");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyStrengthScoreToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

 
  var currentColumn = getColumnIndex("Squat(5RM) Score"); // Finds the column for "LastName"
  
  // Get values from the named range "Strength"
  var range = ss.getRangeByName("StrengthScore");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copySpineToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

  var currentColumn = getColumnIndex("Low Back Screen"); // Finds the column for "LastName"
  
  // Get values from the named range "Spine"
  var range = ss.getRangeByName("Spine");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyUpperBodyToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");


  var currentColumn = getColumnIndex("Liftoffs"); // Finds the column for "LastName"
  
  // Get values from the named range "UpperBody"
  var range = ss.getRangeByName("UpperBody");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function copyQualityToDatabase(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

 
  var currentColumn = getColumnIndex("Pec"); // Finds the column for "LastName"
  
  // Get values from the named range "Quality"
  var range = ss.getRangeByName("Quality");
  var values = range.getValues(); // 2D array

  // Flatten and filter out blanks
  var valuesToPaste = values
    .map(row => row[0])        // flatten
    .filter(val => val !== "" && val !== null); // remove blanks

  // Paste filtered values across starting at the next column after LastName
  for (var i = 0; i < valuesToPaste.length; i++) {
    sheet.getRange(row, currentColumn + i).setValue(valuesToPaste[i]);
  }

  Logger.log("Pasted values at row " + row + " starting from column " + (currentColumn));
}

function getColumnIndex(header){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Database");

  // Get the first row (header row) values
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index for the header name
  var columnIndex = headers.indexOf(header) + 1;
  

  if (columnIndex === 0){
      Browser.msgBox("Column" + header + " not found");
    return -1;
  }
  return columnIndex;
}



function calculateMobilityTotalScore(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");

  const targetColumns = [
    "Ankle Dorsiflexion", "Ankle Eversion & Inversion", "Tibial Rotation", "Thomas Test",
    "Cossack Squat", "Seated Active Hip Rotation", "Single Leg Stability", "Sliders",
    "Overhead Squat", "Low Back Screen", "Pelvic Dissociation", "Locked T-Spine Rotation",
    "Cervical Screen", "Liftoffs", "Scapular Movement", "Back to Wall Shoulder Flexion",
    "Porearm Pronation/Supination", "Field Goal Test", "Total Arc", "Knee Supported ER",
    "Pec", "Lat", "Trap"
  ];

  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  let total = 0;

  targetColumns.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = dataSheet.getRange(row, colIndex + 1).getValue();
      if (!isNaN(value)) total += Number(value);
    } else {
      Logger.log("Column not found: " + header);
    }
  });

  // Write total score to the "Mobility Total Score" column in the same row
  const totalScoreColIndex = headers.indexOf("Mobility Total Score");
  if (totalScoreColIndex !== -1) {
    dataSheet.getRange(row, totalScoreColIndex + 1).setValue(total);
  } else {
    Logger.log("Mobility Total Score column not found.");
  }

  Logger.log("Mobility Total Score written to Database sheet at row " + row + ": " + total);
}


function calculateStrengthTotalScore(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");

  const targetColumns = [
    "Squat(5RM) Score",
    "Trap Bar Deadlift (5RM) Score",
    "Dumbell Bench Press (5RM) Score",
    "Dumbell Row (5RM) Score",
    "Pull-ups (BW) Score"
  ];

  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  let total = 0;

  targetColumns.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = dataSheet.getRange(row, colIndex + 1).getValue();
      if (!isNaN(value)) total += Number(value);
    } else {
      Logger.log("Column not found: " + header);
    }
  });

  // Write total score to the "Strength Total Score" column in the same row
  const totalScoreColIndex = headers.indexOf("Strength Total Score");
  if (totalScoreColIndex !== -1) {
    dataSheet.getRange(row, totalScoreColIndex + 1).setValue(total);
  } else {
    Logger.log("Strength Total Score column not found.");
  }

  Logger.log("Strength Total Score written to Database sheet at row " + row + ": " + total);
}

function calculatePowerTotalScore(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");

  const targetColumns = [
    "Vertical Jump Score",
    "Pause Jump Score",
    "Reactive Jump Score",
    "Medball Reverse Throw Score",
    "Medball Shotput (6oz)"
  ];

  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  let total = 0;

  targetColumns.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = dataSheet.getRange(row, colIndex + 1).getValue();
      if (!isNaN(value)) total += Number(value);
    } else {
      Logger.log("Column not found: " + header);
    }
  });

  // Write total score to the "Power Total Score" column in the same row
  const totalScoreColIndex = headers.indexOf("Power Total Score");
  if (totalScoreColIndex !== -1) {
    dataSheet.getRange(row, totalScoreColIndex + 1).setValue(total);
  } else {
    Logger.log("Power Total Score column not found.");
  }

  Logger.log("Power Total Score written to Database sheet at row " + row + ": " + total);
}

function calculateRotationalTotalScore(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");
  const profileSheet = ss.getSheetByName("Player Profile");

  const targetColumns = [
    "Medball Shotput (6oz) Score",
    "Wide Stance Landmine Pivot Press (5RM) Score",
    "Rotational Medball Slam (4oz) Score"
  ];

  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  let total = 0;

  targetColumns.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = dataSheet.getRange(row, colIndex + 1).getValue();
      if (!isNaN(value)) total += Number(value);
    } else {
      Logger.log("Column not found: " + header);
    }
  });

  // Write total score to the "Rotational Total Score" column in the same row
  const totalScoreColIndex = headers.indexOf("Rotational Total Score");
  if (totalScoreColIndex !== -1) {
    dataSheet.getRange(row, totalScoreColIndex + 1).setValue(total);
  } else {
    Logger.log("Rotational Total Score column not found.");
  }

  Logger.log("Rotational Total Score written to Database sheet at row " + row + ": " + total);
}

function calculateArmStrengthTotalScore(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");

  const targetColumns = [
    "IR Strength Score",
    "IR ROM Score",
    "ER Strength Score",
    "ER ROM Score",
    "Flexion Score",
    "Total Arc Score",
    "Scaption Strength Score",
    "Total Strength Score",
    "Grip Strength Score"
  ];

  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  let total = 0;

  targetColumns.forEach(header => {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      const value = dataSheet.getRange(row, colIndex + 1).getValue();
      if (!isNaN(value)) total += Number(value);
    } else {
      Logger.log("Column not found: " + header);
    }
  });

  // Write total score to the "Arm Strength Total Score" column in the same row
  const totalScoreColIndex = headers.indexOf("Arm Strength Total Score");
  if (totalScoreColIndex !== -1) {
    dataSheet.getRange(row, totalScoreColIndex + 1).setValue(total);
  } else {
    Logger.log("Arm Strength Total Score column not found.");
  }

  Logger.log("Arm Strength Total Score written to Database sheet at row " + row + ": " + total);
}

/*function populateStrengthReport(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");
  const profileSheet = ss.getSheetByName("Player Profile");

  // Get headers from row 2
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  // Find the index of "Squat(5RM)"
  const startHeader = "Squat(5RM)";
  const startIndex = headers.indexOf(startHeader);
  if (startIndex === -1) {
    Logger.log("Header not found: " + startHeader);
    return;
  }

  // Get the named range
  const strengthRange = ss.getRangeByName("StrengthReport");
  if (!strengthRange) {
    Logger.log("Named range 'StrengthReport' not found.");
    return;
  }

  const numRows = strengthRange.getHeight(); // number of rows in the vertical range

  // Get values from the row, starting from the start header
  const valuesToCopy = dataSheet.getRange(row, startIndex + 1, 1, numRows).getValues()[0];

  // Transform into a vertical 2D array (each value in its own row)
  const verticalValues = valuesToCopy.map(val => [val]);

  // Paste vertically into the StrengthReport range
  strengthRange.setValues(verticalValues);

  Logger.log("Strength report populated vertically with: " + valuesToCopy.join(", "));
}

function populatePowerReport(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");
  const profileSheet = ss.getSheetByName("Player Profile");

  // Get headers from row 2
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  // Find the index of "Lateral Jump"
  const startHeader = "Lateral Jump";
  const startIndex = headers.indexOf(startHeader);
  if (startIndex === -1) {
    Logger.log("Header not found: " + startHeader);
    return;
  }

  // Get the PowerReport named range
  const powerRange = ss.getRangeByName("PowerReport");
  if (!powerRange) {
    Logger.log("Named range 'PowerReport' not found.");
    return;
  }

  const numRows = powerRange.getHeight(); // vertical length of the range

  // Get the values from the row starting from the Power section
  const valuesToCopy = dataSheet.getRange(row, startIndex + 1, 1, numRows).getValues()[0];

  // Convert to vertical array
  const verticalValues = valuesToCopy.map(val => [val]);

  // Paste into the PowerReport range
  powerRange.setValues(verticalValues);

  Logger.log("Power report populated vertically with: " + valuesToCopy.join(", "));
}

function populateRotationalReport(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");
  const profileSheet = ss.getSheetByName("Player Profile");

  // Get headers from row 2
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  // Find the index of "Medball Shotput (6oz)"
  const startHeader = "Medball Shotput (6oz)";
  const startIndex = headers.indexOf(startHeader);
  if (startIndex === -1) {
    Logger.log("Header not found: " + startHeader);
    return;
  }

  // Get the RotationalReport named range
  const rotationalRange = ss.getRangeByName("RotationalReport");
  if (!rotationalRange) {
    Logger.log("Named range 'RotationalReport' not found.");
    return;
  }

  const numRows = rotationalRange.getHeight(); // vertical length of the range

  // Get the values from the row starting from the Rotational section
  const valuesToCopy = dataSheet.getRange(row, startIndex + 1, 1, numRows).getValues()[0];

  // Convert to vertical array
  const verticalValues = valuesToCopy.map(val => [val]);

  // Paste into the RotationalReport range
  rotationalRange.setValues(verticalValues);

  Logger.log("Rotational report populated vertically with: " + valuesToCopy.join(", "));
}

function populateArmStrengthReport(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Database");
  const profileSheet = ss.getSheetByName("Player Profile");

  // Get headers from row 2
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];

  // Find the index of "IR Strength"
  const startHeader = "IR Strength";
  const startIndex = headers.indexOf(startHeader);
  if (startIndex === -1) {
    Logger.log("Header not found: " + startHeader);
    return;
  }

  // Get the ArmStrengthReport named range
  const armStrengthRange = ss.getRangeByName("ArmStrengthReport");
  if (!armStrengthRange) {
    Logger.log("Named range 'ArmStrengthReport' not found.");
    return;
  }

  const numRows = armStrengthRange.getHeight(); // vertical length of the range

  // Get the values from the row starting from the Arm Strength section
  const valuesToCopy = dataSheet.getRange(row, startIndex + 1, 1, numRows).getValues()[0];

  // Convert to vertical array
  const verticalValues = valuesToCopy.map(val => [val]);

  // Paste into the ArmStrengthReport range
  armStrengthRange.setValues(verticalValues);

  Logger.log("Arm Strength report populated vertically with: " + valuesToCopy.join(", "));
}

function onCreatePlayerProfileClick() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Create Report", "Enter the full name (exactly as it appears):", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Action cancelled.");
    return;
  }

  const fullName = response.getResponseText().trim();
  if (!fullName) {
    ui.alert("No name entered. Please try again.");
    return;
  }

  const row = findRowByFullName(fullName);

  if (row !== -1) {
    ui.alert("Found " + fullName + " at row: " + row);
    // Do something with the row here, e.g., calculate score
    calculateMobilityTotalScore(row);
    calculateStrengthTotalScore(row);
    calculatePowerTotalScore(row);
    calculateRotationalTotalScore(row);
    calculateArmStrengthTotalScore(row);
    populateStrengthReport(row);
    populatePowerReport(row);
    populateRotationalReport(row);
    populateArmStrengthReport(row);
  } else {
    ui.alert("Name not found. Please make sure it matches exactly.");
    return;
  }
}*/

function findRowByFullName(fullName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Database");

  const names = dbSheet.getRange("A3:A" + dbSheet.getLastRow()).getValues(); // assumes headers are on row 2

  for (let i = 0; i < names.length; i++) {
    if (names[i][0].toString().trim().toLowerCase() === fullName.toLowerCase()) {
      return i + 3; // Adjust for offset since we started at row 3
    }
  }

  return -1; // Not found
}

function populateSlideForPlayer(fullName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Database");
  const row = findRowByFullName(fullName);

  if (row === -1) {
  SpreadsheetApp.getUi().alert(fullName + " is not found in the Database");
    return;
  }

  SpreadsheetApp.getUi().alert(fullName + " is found at row " + row + " of the Database!");

  const headers = dbSheet.getRange(2, 1, 1, dbSheet.getLastColumn()).getValues()[0];
  const playerData = dbSheet.getRange(row, 1, 1, dbSheet.getLastColumn()).getValues()[0];

  // Extract scores
  const mobility = getValue(headers, playerData, "Mobility Total Score");
  const strength = getValue(headers, playerData, "Strength Total Score");
  const power = getValue(headers, playerData, "Power Total Score");
  const rotational = getValue(headers, playerData, "Rotational Total Score");
  const arm = getValue(headers, playerData, "Arm Strength Total Score");
  const totalScore = mobility + strength + power + rotational + arm;
  const squat = getValue(headers, playerData, "Squat(5RM)");
  const deadlift = getValue(headers, playerData, "Trap Bar Deadlift (5RM)");
  const benchPress = getValue(headers, playerData, "Dumbell Bench Press (5RM)");
  const dumbellRow = getValue(headers, playerData, "Dumbell Row (5RM)");
  const pullups = getValue(headers, playerData, "Pull-ups (BW)");
  const lateralJump = getValue(headers, playerData, "Lateral Jump");
  const broadJump = getValue(headers, playerData, "Broad Jump");
  const verticalJump = getValue(headers, playerData, "Vertical Jump");
  const pauseJump = getValue(headers, playerData, "Pause Jump");
  const reactiveJump = getValue(headers, playerData, "Reactive Jump");
  const medballReverse = getValue(headers, playerData, "Medball Reverse Throw");
  const shotput = getValue(headers, playerData, "Medball Shotput (6oz)");

  // Placeholder replacement map
  const placeholderMap = {
    "{{Player}}": fullName,
    "{{Mobility Score}}": mobility,
    "{{Strength Score}}": strength,
    "{{Power Score}}": power,
    "{{Rotational Score}}": rotational,
    "{{Arm Strength and Health Score}}": arm,
    "{{Total Score}}": totalScore,
    "{{Squat}}": squat,
    "{{Trap Bar Deadlift}}": deadlift,
    "{{Dumbbell Bench Press}}": benchPress,
    "{{Dumbbell Row}}": dumbellRow,
    "{{Pull-Ups}}": pullups,
    "{{Lateral Jump}}": lateralJump,
    "{{Broad Jump}}": broadJump,
    "{{Vertical Jump}}": verticalJump,
    "{{Pause Jump}}": pauseJump,
    "{{Reactive Jump}}": reactiveJump,
    "{{Medball Reverse Throw}}": medballReverse, 
    "{{Medball Shotput}}": shotput

    // Add more as needed, like "{{Squat}}", etc.
  };

  // Open the presentation and duplicate the template slide
  const presentationId = '1P9rCxEuuz1m8TDInrWWFCTfN3a8MqcbI3lYokGETx84';
  const presentation = SlidesApp.openById(presentationId);
  const templateSlide = presentation.getSlides()[0]; // Slide 1 is the template

  const newSlide = templateSlide.duplicate();
  //presentation.appendSlide(newSlide); // Puts the copy at the end

  // Replace placeholders with values in the new slide
  for (let placeholder in placeholderMap) {
    newSlide.replaceAllText(placeholder, String(placeholderMap[placeholder] ?? ""));
  }

  Logger.log("Slide created for: " + fullName);
}

function promptAndRun() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Generate Slide", "Enter the player's full name:", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const fullName = response.getResponseText().trim();
    if (fullName) {
      populateSlideForPlayer(fullName);
    } else {
      ui.alert("No name entered. Please try again.");
    }
  } else {
    ui.alert("Slide generation cancelled.");
  }
}


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Screening Tools')
    .addItem('Generate Slide for Player', 'promptAndRun')
    .addItem('Add to Database', 'EnterToDatabase')
    .addItem('Clear Screening Sheet', 'clearScreeningSheet')
    .addToUi();
}

function getValue(headers, dataRow, headerName) {
  const index = headers.indexOf(headerName);
  return index !== -1 ? dataRow[index] : "";
}



