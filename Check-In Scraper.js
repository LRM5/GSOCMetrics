function pullCheckInDataFromSelectedURLs() {
  logExecution(); // Logs the execution time at the start
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const checkInDataSheet = ss.getSheetByName("Check-In Data");
  const tripSheet = ss.getSheetByName("Trips");

  // Clear previous contents and setup headers
  checkInDataSheet.clearContents();
  checkInDataSheet.appendRow(["Workbook URL", "Sheet Name", "Check In Type", "Check In Date / Time", "Responses"]);

  const tripNames = new Set(tripSheet.getDataRange().getValues().flat()); // Pre-fetch and cache trip names

  const settingsRange = settingsSheet.getRange("A2:C" + settingsSheet.getLastRow());
  const settingsValues = settingsRange.getValues();

  const newData = []; // For bulk append
  const newTripNames = []; // To update trips sheet
  
  settingsValues.forEach(([isSelected, , workbookUrl]) => {
    if (isSelected && workbookUrl) {
      processWorkbook(workbookUrl, newData, newTripNames, tripNames);
    }
  });

  // Bulk append new data to the check-in data sheet and trips sheet
  if (newData.length) checkInDataSheet.getRange(checkInDataSheet.getLastRow() + 1, 1, newData.length, 5).setValues(newData);
  if (newTripNames.length) tripSheet.getRange(tripSheet.getLastRow() + 1, 1, newTripNames.length, 1).setValues(newTripNames.map(name => [name]));
}

function processWorkbook(workbookUrl, newData, newTripNames, tripNames) {
  const datePattern = /\d{2}\/\d{2}\/\d{2} \d{2}:\d{2} (PDT|PST|Local)/; // Precompiled regex
  try {
    const workbook = SpreadsheetApp.openByUrl(workbookUrl);
    workbook.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!excludeSheet(sheetName)) {
        if (!tripNames.has(sheetName)) {
          newTripNames.push([sheetName]); // Add sheetName wrapped in an array for setValues
          tripNames.add(sheetName); // Update the in-memory set
        }
        const dataRange = sheet.getDataRange().getValues();
        dataRange.forEach((row) => {
          row.forEach((cell, cellIndex) => {
            if (cellIndex < row.length - 2 && isDateCell(String(row[cellIndex + 1]), datePattern)) {
              const checkInType = cell; // The current cell is assumed to be the check-in type
              const rawCheckInDateTime = row[cellIndex + 1];
              const dateTimePart = rawCheckInDateTime.split(" / ")[1]; // Focus on the second date-time part after the slash

              let formattedDate = "Invalid Date";
              if (dateTimePart) {
                const parts = dateTimePart.split(" ");
                const datePart = parts[0];
                const timePart = parts[1];
                const checkInDate = new Date(datePart + " " + timePart);

                // Validate and format the extracted date
                if (!isNaN(checkInDate.getTime())) {
                  formattedDate = Utilities.formatDate(checkInDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
                } else {
                  Logger.log("Invalid date encountered for check-in type: " + checkInType);
                }
              }

              const checkInComment = row[cellIndex + 2].trim();
              if (checkInComment !== "") {
                newData.push([workbookUrl, sheetName, checkInType, formattedDate, checkInComment]);
              }
            }
          });
        });
      }
    });
  } catch (e) {
    Logger.log(`Error accessing ${workbookUrl}: ${e}`);
  }
}

function isDateCell(cellValue, datePattern) {
  return datePattern.test(cellValue);
}

function logExecution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheetName = "Log";
  let logSheet = ss.getSheetByName(logSheetName);

  // If the "Log" sheet doesn't exist, create it
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    // Optional: Append header row if you are creating the log sheet for the first time
    logSheet.appendRow(["Timestamp", "Action"]); 
  }

  // Append the current timestamp and a descriptive action to the "Log" sheet
  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  const actionDescription = "Check In Scraper Script executed"; // Customize this as needed
  
  logSheet.appendRow([formattedTimestamp, actionDescription]);
}

function excludeSheet(sheetName) {
  // Define a list of sheet names to be excluded from processing
  const excludedSheets = [
    "Cover Page",
    "Template 1 ",
    "Template 2 ",
    "Template 3 ",
    "Template 4 ",
    "edit-tracker",
    "Log" // Including the log sheet as well, assuming you don't want to process it
  ];

  // Check if the provided sheetName exists in the list of excludedSheets
  return excludedSheets.includes(sheetName);
}
