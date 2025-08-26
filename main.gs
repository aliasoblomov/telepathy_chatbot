var cc = DataStudioApp.createCommunityConnector();

// [START get_config] 
function getConfig(request) {
 var config = cc.getConfig();

 config
  .newInfo()
  .setId('instructions')
  .setText(
   'This connector reads data from the specified sheet in your Google Sheet and logs messages, timestamp, userid, receptions, and send status to another sheet within the same Google Sheet.'
  );

 config
  .newTextInput()
  .setId('spreadsheetUrl')
  .setName('Google Sheet URL')
  .setPlaceholder('Enter the Google Sheet URL here')
  .setHelpText(
   'Provide the full URL of your Google Sheet. The connector will extract the spreadsheet ID from this URL.'
  );

 config
  .newTextInput()
  .setId('loggingSheetName')
  .setName('Logging Sheet Name')
  .setPlaceholder('Enter the name of the sheet for logging data')
  .setHelpText('The name of the sheet where data will be logged.');

 config
  .newTextInput()
  .setId('readingSheetName')
  .setName('Reading Sheet Name')
  .setPlaceholder('Enter the name of the sheet to read data from')
  .setHelpText('The name of the sheet from which data will be read.');

 config
  .newTextInput()
  .setId('message')
  .setName('Message')
  .setPlaceholder('Enter your message here')
  .setHelpText('This is the message to be logged.')
  .setAllowOverride(true);

  config
  .newSelectSingle()
  .setId('userid')
  .setName('from Who?')
  .setHelpText('Select the sender.')
  .addOption(config.newOptionBuilder().setLabel('Ali').setValue('ali'))
  .addOption(config.newOptionBuilder().setLabel('Arash').setValue('arash'))
  .addOption(config.newOptionBuilder().setLabel('Reza').setValue('reza'))
  .addOption(config.newOptionBuilder().setLabel('Mohsen').setValue('mohsen'))
  .setAllowOverride(true);

 config
  .newSelectMultiple()
  .setId('receptions')
  .setName('to Who?')
  .setHelpText('Select the receptions. Multiple selections allowed.')
  .addOption(config.newOptionBuilder().setLabel('Ali').setValue('ali@siavak.com'))
  .addOption(config.newOptionBuilder().setLabel('Sia').setValue('aliizadi27@gmail.com'))
  .addOption(config.newOptionBuilder().setLabel('Arash').setValue('aliizadi727@gmail.com'))
  .setAllowOverride(true);

  

 config
  .newSelectSingle()
  .setId('sendStatus')
  .setName('Send Status')
  .setHelpText('Select whether to send the message. Only sent messages are logged.')
  .addOption(config.newOptionBuilder().setLabel('Unsent').setValue('UNSENT'))
  .addOption(config.newOptionBuilder().setLabel('Sent').setValue('SENT'))
  .setAllowOverride(true);

 config.setDateRangeRequired(false);
 return config.build();
}
// [END get_config]

// [START get_schema] 
function getFields(request) {
 var fields = cc.getFields();
 var types = cc.FieldType;

 // Fetch headers from the sheet dynamically
 var headers = getSheetHeaders(request);

 headers.forEach(function (header) {
  fields
   .newDimension()
   .setId(header.toLowerCase().replace(/[^a-z0-9]/g, '_')) // Use lowercase header for consistency and replace special characters
   .setName(header)
   .setType(types.TEXT);
 });

 return fields;
}

function getSchema(request) {
 return { schema: getFields(request).build() };
}
// [END get_schema]

// [START get_data] 
function getData(request) {
 request.configParams = validateConfig(request.configParams);

 // Log message, receptions, userid, and timestamp to the sheet if sendStatus is 'SENT'
 if (request.configParams.sendStatus === 'SENT') {
  try {
   var loggedData = logDataToSheet(request.configParams);
   checkAndSendEmails(loggedData); // Check receptions and send emails
  } catch (e) {
   cc.newUserError()
    .setDebugText('Error logging data to Sheet. Exception details: ' + e)
    .setText('The connector has encountered an error. Please try again later.')
    .throwException();
  }
 }

 var requestedFields = getFields(request).forIds(
  request.fields.map(function (field) {
   return field.name.toLowerCase().replace(/[^a-z0-9]/g, '_'); // Match field IDs with lowercase and replaced special characters
  })
 );

 try {
  var apiResponse = fetchDataFromSheet(request);
  var data = getFormattedData(apiResponse, requestedFields);
 } catch (e) {
  cc.newUserError()
   .setDebugText('Error fetching data from Sheet. Exception details: ' + e)
   .setText(
    'The connector has encountered an unrecoverable error. Please try again later, or file an issue if this error persists.'
   )
   .throwException();
 }

 return {
  schema: requestedFields.build(),
  rows: data,
 };
}

// Fetches data from the specified Google Sheet
function fetchDataFromSheet(request) {
 var spreadsheetId = extractSpreadsheetId(request.configParams.spreadsheetUrl);
 var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(
  request.configParams.readingSheetName
 );
 var values = sheet.getDataRange().getValues();
 var headers = values.shift(); // Assuming the first row contains headers

 var data = values.map(function (row) {
  var rowData = {};
  row.forEach(function (value, index) {
   rowData[
    headers[index].toLowerCase().replace(/[^a-z0-9]/g, '_')
   ] = value; // Using lowercase headers for consistency and replace special characters
  });
  return rowData;
 });

 return data;
}

// Formats the data from the Google Sheet into the required format
function getFormattedData(sheetData, requestedFields) {
 return sheetData.map(function (rowData) {
  var formattedRow = requestedFields.asArray().map(function (requestedField) {
   return rowData[requestedField.getId()]; // Use the requested field ID as the key
  });
  return { values: formattedRow };
 });
}

// Fetches headers from the sheet dynamically
function getSheetHeaders(request) {
 var spreadsheetId = extractSpreadsheetId(request.configParams.spreadsheetUrl);
 var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(
  request.configParams.readingSheetName
 );
 var headers = sheet
  .getRange(1, 1, 1, sheet.getLastColumn())
  .getValues()[0];
 return headers;
}

// Extracts the spreadsheet ID from the Google Sheet URL
function extractSpreadsheetId(url) {
 var matches = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
 return matches ? matches[1] : null;
}

// Logs data to the specified Google Sheet and returns the logged data
function logDataToSheet(configParams) {
 var SPREADSHEET_ID = extractSpreadsheetId(configParams.spreadsheetUrl); // Get Spreadsheet ID from URL
 var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(
  configParams.loggingSheetName
 );
 var lastRow = sheet.getLastRow();

 // Prepare data to log
 var timestamp = formatTimestamp(new Date());
 var dataToLog = [
  configParams.message || 'No message entered',
  configParams.userid || 'No Sender provided',
  configParams.receptions || 'No reception selected',
  timestamp,
 ];

 // Set headers if the sheet is empty
 if (lastRow === 0) {
  sheet.appendRow([
   'Message',
   'Sender',
   'Receptions',
   'Timestamp'
  ]);
  lastRow = 1; // Update lastRow after adding headers
 }

 // Append data to the next available row
 sheet
  .getRange(lastRow + 1, 1, 1, dataToLog.length)
  .setValues([dataToLog]);

 // Return the logged data for future use if needed
 return {
  'Message': dataToLog[0],
  'Sender': dataToLog[1],
  'Receptions': dataToLog[2],
  'Timestamp': dataToLog[3]
  
 };
}

function formatTimestamp(date) {
 // Format the date to 'YYYY-MM-DD HH:MM:SS'
 var year = date.getFullYear();
 var month = ('0' + (date.getMonth() + 1)).slice(-2); // Months are zero-indexed
 var day = ('0' + date.getDate()).slice(-2);
 var hours = ('0' + date.getHours()).slice(-2);
 var minutes = ('0' + date.getMinutes()).slice(-2);
 var seconds = ('0' + date.getSeconds()).slice(-2);
 return year + '-' + month + '-' + day + ' ' + hours + ':' + minutes + ':' + seconds;
}

// Validate email address format
function isValidEmail(email) {
 var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
 return re.test(String(email).toLowerCase());
}

function checkAndSendEmails(loggedData) {
  var receptions = loggedData['Receptions'].split(',').map(function(email) {
    return email.trim();
  });

  var validReceptions = receptions.filter(function(email) {
    return isValidEmail(email);
  });

  if (validReceptions.length > 0) {
    var message = loggedData['Message'] || "No message provided";
    var sender = loggedData['Sender'] || "Someone";

    var subject = `LookerChatBot - ${sender}'s message awaits you!⌛`;

    // Cute and friendly email text with an anchor for the dashboard link
    var emailBody = `
      <p>Hello there,</p>
      <p>You've got a  message from <strong>${sender}</strong>!</p>
      <p><strong>"${message}"</strong></p>
      <p>Don't keep it waiting—click here to check it out: <a href="https://lookerstudio.google.com/u/0/reporting/17ad7e2e-89ff-4c06-9017-d7ff411b4a33/page/MaoBE/edit">Dashboard Link</a></p>
      <p>&nbsp;</p>
      <p><em>Your friendly LookerChatBot ☁️<em></p>
    `;

    // Send the email using HTML format
    GmailApp.sendEmail(validReceptions.join(','), subject, '', {htmlBody: emailBody});
  }
}




function validateConfig(configParams) {
 configParams = configParams || {};
 configParams.message = configParams.message || 'No message entered'; // Default value for "Message"
 configParams.userid = configParams.userid || 'No Sender provided'; // Default value for "Sender"
 configParams.sendStatus = configParams.sendStatus || 'UNSENT'; // Default value for "Send Status"
 return configParams;
}

function isAdminUser() {
 return false;
}
