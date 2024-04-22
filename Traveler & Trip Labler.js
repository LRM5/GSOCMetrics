function updatePersonnelTrips() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tripsSheet = ss.getSheetByName("Trips");
  var travelersSheet = ss.getSheetByName("DB:Travelers");
  var outputSheet = ss.getSheetByName("NC Personnel and Trips");

  // Fetching trip data
  var tripDataRange = tripsSheet.getRange("A2:A" + tripsSheet.getLastRow());
  var tripDataValues = tripDataRange.getValues().filter(row => row[0].trim() !== "");
  var travelerNamesRange = travelersSheet.getRange("A2:A" + travelersSheet.getLastRow());
  var travelerNamesValues = travelerNamesRange.getValues().flat().filter(name => name.trim() !== "");

  var outputData = [];

  // Processing each trip entry
  tripDataValues.forEach(row => {
    var tripName = row[0].trim();
    var nameIdentifiers = extractLastNameIdentifiers(tripName);
    var matchedNames = nameIdentifiers.map(identifier => findMatchingName(travelerNamesValues, identifier))
                                      .filter(name => name !== "Name not found");

    outputData.push([tripName].concat(matchedNames.length > 0 ? matchedNames : ["Name not found"]));
  });

  clearAndWriteData(outputSheet, outputData);
}

function extractLastNameIdentifiers(cellContent) {
  // Modified to handle names in any case
  var matches = cellContent.toUpperCase().match(/\b[A-Z]+\b/g) || [];
  return matches.filter(match => match.length > 1); // Filter out single-letter matches that are likely initials
}

function findMatchingName(travelers, identifier) {
  // Direct comparison for exact matches first
  let directMatch = travelers.find(traveler => traveler.toLowerCase() === identifier.toLowerCase());
  if (directMatch) return directMatch;

  // If direct match not found, attempt matching with initial or partial first name
  let identifierParts = identifier.split(/\s+/); // Splitting identifier into parts (initials or first names and last name)
  let matches = travelers.filter(traveler => {
    let travelerParts = traveler.split(/\s+/);
    let lastNameMatch = travelerParts.some(part => part.toLowerCase() === identifierParts[identifierParts.length - 1].toLowerCase());
    let firstNameInitialMatch = identifierParts.length > 1 && travelerParts[0][0].toLowerCase() === identifierParts[0][0].toLowerCase();
    return lastNameMatch && (identifierParts.length === 1 || firstNameInitialMatch);
  });

  return matches.length > 0 ? matches.join(", ") : "Name not found";
}

function clearAndWriteData(sheet, data) {
  // Clear existing content from the sheet to avoid leftover data.
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }

  // Determine the maximum number of columns needed by any row in the data.
  const maxColumns = data.reduce((max, row) => Math.max(max, row.length), 0);

  // Pad each row of the data to have the same number of columns.
  const paddedData = data.map(row => {
    while (row.length < maxColumns) {
      row.push(""); // Pad with empty strings to match the longest row.
    }
    return row;
  });

  // Write the padded data to the sheet.
  if (paddedData.length > 0) {
    sheet.getRange(2, 1, paddedData.length, maxColumns).setValues(paddedData);
  }
}
