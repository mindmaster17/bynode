var sheetName = 'Sheet1'; // Change this to your actual sheet name
var scriptProp = PropertiesService.getScriptProperties();

function intialSetup() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    var sheet = doc.getSheetByName(sheetName);

    // Fetch existing user IDs
    var existingIds = sheet.getRange('F2:F' + sheet.getLastRow()).getValues().flat().filter(String);

    // Generate a unique user ID
    var userId = generateUniqueUserId(existingIds);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;

    var newRow = headers.map(function(header) {
      if (header === 'timestamp') {
        return new Date();
      } else if (header === 'userId') {
        return userId; // Add the unique user ID
      } else {
        return e.parameter[header];
      }
    });

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow, 'userId': userId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Function to generate a unique user ID
function generateUniqueUserId(existingIds) {
  var userId;
  do {
    const randomNumber = Math.floor(10000000 + Math.random() * 90000000); // Generate a random 8-digit number
    userId = `TMS${randomNumber}`; // Return the user ID in the desired format
  } while (existingIds.includes(userId)); // Repeat if the ID already exists
  return userId; // Return a unique user ID
}

function doGet(e) {
  const sheet = SpreadsheetApp.openById(scriptProp.getProperty('key')).getActiveSheet();
  const ids = sheet.getRange('F2:F' + sheet.getLastRow()).getValues(); // Adjust range according to your sheet layout
  const existingIds = ids.flat().filter(String); // Flatten and filter empty values
  return ContentService.createTextOutput(JSON.stringify({ ids: existingIds })).setMimeType(ContentService.MimeType.JSON);
}
