var sheetName = 'Sheet1'; // Change this to your actual sheet name
var scriptProp = PropertiesService.getScriptProperties();

function initialSetup() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'error': 'Unable to acquire lock' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    var sheet = doc.getSheetByName(sheetName);

    // Fetch existing user IDs
    var existingIds = sheet.getRange('F2:F' + sheet.getLastRow()).getValues().flat().filter(String);

    // Generate a unique user ID
    var userId = generateUniqueUserId(existingIds); // Generate only once and use consistently

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;

    var newRow = headers.map(function(header) {
      if (header === 'timestamp') {
        return new Date();
      } else if (header === 'userId') {
        return userId; // Consistently using the generated userId
      } else {
        return e.parameter[header] || ''; // Ensure missing parameters are handled
      }
    });

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    // Return the same user ID in the response
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

  // Replace with actual column index for status and email
  const statusColumnIndex = 7;  // Example: assuming status is in column G (index 7)
  const emailColumnIndex = 2;   // Example: assuming email is in column B (index 2)
  const entryCodeColumnIndex = 8; // Example: assuming entry code is in column H (index 8)

  if (sheet.getName() !== sheetName || range.getColumn() !== statusColumnIndex) return;

  const newValue = range.getValue();
  if (newValue === 'approved') {
    const row = range.getRow();
    const email = sheet.getRange(row, emailColumnIndex).getValue();
    const entryCode = generateUniqueEntryCode(sheet); // Generate unique entry code

    // Send email to user
    MailApp.sendEmail({
      to: email,
      subject: 'TMS Entry Code',
      body: `Congratulations! Your registration has been approved.\n\nYour entry code is: ${entryCode}\n\nThank you!`
    });

    // Store the entry code in the sheet
    sheet.getRange(row, entryCodeColumnIndex).setValue(entryCode);
  }
}

// Function to generate a unique alphanumeric entry code
function generateUniqueEntryCode(sheet) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let entryCode;
  const existingCodes = sheet.getRange('H2:H' + sheet.getLastRow()).getValues().flat().filter(String); // Replace 'H' with your actual entry code column

  do {
    entryCode = '';
    for (let i = 0; i < 7; i++) {
      entryCode += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (existingCodes.includes(entryCode)); // Repeat if the code already exists

  return entryCode; // Return a unique entry code
}
