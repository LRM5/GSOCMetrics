unction compileTravelData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('2023 Traveler/Trip List'); // Ensure this name matches your sheet
  let outputSheet = ss.getSheetByName('Summary'); // Attempt to get the Summary sheet

  // Check if the Summary sheet exists, if not, create it
  if (!outputSheet) {
    outputSheet = ss.insertSheet('Summary');
  }

  const dataRange = dataSheet.getRange('A2:H' + dataSheet.getLastRow()).getValues();
  const travelerCountries = {};

  dataRange.forEach(row => {
    const travelers = row.slice(1, 6); // Assuming travelers are from columns B to G
    const country = row[7]; // Assuming country is in column H

    travelers.forEach(traveler => {
      if (traveler) {
        if (!travelerCountries[traveler]) {
          travelerCountries[traveler] = new Set();
        }
        travelerCountries[traveler].add(country);
      }
    });
  });

  const output = [];
  for (const [traveler, countriesSet] of Object.entries(travelerCountries)) {
    const countries = Array.from(countriesSet).join(' | ');
    output.push([traveler, countries]);
  }

  if (output.length > 0) {
    outputSheet.getRange('A1').setValue('Traveler');
    outputSheet.getRange('B1').setValue('Countries Visited');
    outputSheet.getRange('A2:B' + (output.length + 1)).setValues(output);
  }
}
