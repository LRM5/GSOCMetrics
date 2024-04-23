function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Metrics Scripts')
    .addSubMenu(ui.createMenu('Actions')
      .addItem('Check-In Scraper', 'pullCheckInDataFromSelectedURLs')
      .addItem('Traveler Trip Labeler', 'updatePersonnelTrips'))
    .addToUi();
}
