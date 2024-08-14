function sendWeeklyEmail() {
  // Set up the spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Combined Responses');
  
  // Get the data from the relevant columns (B and V to AB are columns 2 and 22 to 28)
  var dataRange = sheet.getRange(2, 5, sheet.getLastRow() - 1, 27); // Starts from row 2 to skip headers
  var data = dataRange.getValues();
  
  // Initialize variables for checking the columns and counting unprocessed events
  var unprocessedCount = 0;
  
  // Loop through each row in the data
  data.forEach(function(row) {
    
    var columnE = row[0]; // Column E
    var approved = row[19]; // Column X - 'Approved'
    var calendarUpdated = row[22]; // Column Z - 'Google Calendar Updated'
    var mobilize = row[23]; // Column AA - 'Mobilize'
    // Check if Column E is not blank, then check other columns
    if (columnE !== '') {
      if (approved === 'No' || approved === '' || 
          calendarUpdated === 'No' || calendarUpdated === '' || 
          mobilize === 'No' || mobilize === '') {
        unprocessedCount++;
      }
    }
  });

  // If there are unprocessed events, send an email
  if (unprocessedCount > 0) {
    var subject = "Weekly Alert: " + unprocessedCount + " Unprocessed Event(s) Detected";
    var body = "Hello Team,\n\nThere " + (unprocessedCount === 1 ? "is" : "are") + " currently " + unprocessedCount + " unprocessed event" + (unprocessedCount === 1 ? "" : "s") + " detected in the 'Combined Responses' sheet. Please review and process them accordingly.\n\nYou can access the spreadsheet here: " + spreadsheet.getUrl() + "\n\nBest regards,\nUT Dems";
    var recipient = "recipient@example.com"; // Replace with your email address
    
    MailApp.sendEmail(
      recipient, 
      subject, 
      body,
      {
        noReply: true,
        name: "Utah Dems"
      });
  }
}

// Set up a trigger to run this function weekly
function createWeeklyTrigger() {
  // Delete existing triggers for this function to prevent duplicates
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendWeeklyEmail') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create a new weekly trigger
  ScriptApp.newTrigger('sendWeeklyEmail')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}
