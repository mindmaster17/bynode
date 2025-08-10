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
    const randomNumber = Math.floor(10000000 + Math.random() * 90000000);
    userId = `TMS${randomNumber}`;
  } while (existingIds.includes(userId));
  return userId;
}

function doGet(e) {
  const sheet = SpreadsheetApp.openById(scriptProp.getProperty('key')).getActiveSheet();
  const ids = sheet.getRange('F2:F' + sheet.getLastRow()).getValues();
  const existingIds = ids.flat().filter(String);
  return ContentService.createTextOutput(JSON.stringify({ ids: existingIds })).setMimeType(ContentService.MimeType.JSON);
}

// Trigger function to send email upon status change
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() !== sheetName || range.getColumn() !== /* column index for status */) return;

  const newValue = range.getValue();
  if (newValue === 'approved') {
    const row = range.getRow();
    const email = sheet.getRange(row, /* column index for email */).getValue(); // Replace with actual column index for email
    const entryCode = generateUniqueEntryCode(sheet); // Generate unique entry code

    // Send email to user
    MailApp.sendEmail({
      to: email,
      subject: 'TMS Entry Code',
      body: `Congratulations! Your registration has been approved.\n\nYour entry code is: ${entryCode}\n\nThank you!`
    });

    // Optionally, store the entry code in the sheet
    sheet.getRange(row, /* column index for entry code */).setValue(entryCode); // Replace with actual column index for entry code
  }
}

// Function to generate a unique alphanumeric entry code
function generateUniqueEntryCode(sheet) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let entryCode;
  const existingCodes = sheet.getRange('G2:G' + sheet.getLastRow()).getValues().flat().filter(String); // Replace 'G' with your actual entry code column

  do {
    entryCode = '';
    for (let i = 0; i < 7; i++) {
      entryCode += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (existingCodes.includes(entryCode)); // Repeat if the code already exists

  return entryCode; // Return a unique entry code
}
