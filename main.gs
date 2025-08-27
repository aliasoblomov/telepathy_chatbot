Temp: https://docs.google.com/spreadsheets/d/1gSj44Hc8yG1jRqjQS7GxlH6azPe9HC7tOtJT23Da2Ws/edit?gid=0#gid=0

var cc = DataStudioApp.createCommunityConnector();

// [START get_config]
function getConfig(request) {
  var config = cc.getConfig();

  config
    .newInfo()
    .setId('instructions')
    .setText(
      'This connector reads data from a specified Google Sheet and logs comments and a timestamp to another sheet. When you select "SENT", the comment is also sent to your webhook URL.'
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
    .setId('comment')
    .setName('Comment')
    .setPlaceholder('Enter your comment here')
    .setHelpText('This is the comment to be logged and sent.')
    .setAllowOverride(true);

  config
    .newTextInput()
    .setId('webhookUrl')
    .setName('Webhook URL')
    .setPlaceholder('Enter the webhook URL')
    .setHelpText('The URL to send the comment and timestamp to.');

  config
    .newSelectSingle()
    .setId('sendStatus')
    .setName('Send Status')
    .setHelpText('Select "SENT" to log the comment and send it.')
    .addOption(config.newOptionBuilder().setLabel(' ').setValue(''))
    .addOption(config.newOptionBuilder().setLabel('SENT').setValue('SENT'))
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

  // Log comment and timestamp to the sheet and send to webhook if sendStatus is 'SENT'
  if (request.configParams.sendStatus === 'SENT') {
    try {
      var loggedData = logDataToSheet(request.configParams);
      sendWebhook(request.configParams, loggedData); // Send message to webhook
    } catch (e) {
      cc.newUserError()
        .setDebugText('Error logging data or sending webhook. Exception details: ' + e)
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
// [END get_data]

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
  if (!url) return null;
  var matches = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return matches ? matches[1] : null;
}

// Logs data to the specified Google Sheet and returns the logged data
function logDataToSheet(configParams) {
  var SPREADSHEET_ID = extractSpreadsheetId(configParams.spreadsheetUrl);
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(
    configParams.loggingSheetName
  );
  var lastRow = sheet.getLastRow();

  // Prepare data to log
  var timestamp = formatTimestamp(new Date());

  var dataToLog = [
    configParams.comment || 'No comment entered',
    timestamp,
  ];

  // Set headers if the sheet is empty
  if (lastRow === 0) {
    sheet.appendRow([
      'Comment',
      'Timestamp'
    ]);
    lastRow = 1; // Update lastRow after adding headers
  }

  // Append data to the next available row
  sheet
    .getRange(lastRow + 1, 1, 1, dataToLog.length)
    .setValues([dataToLog]);

  // Return the logged data for future use
  return {
    'Comment': dataToLog[0],
    'Timestamp': dataToLog[1]
  };
}

// Sends data to the webhook
function sendWebhook(configParams, loggedData) {
  var payload = {
    comment: loggedData.Comment,
    timestamp: loggedData.Timestamp
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(configParams.webhookUrl, options);
  } catch (e) {
    Logger.log('Error sending payload to webhook. Exception details: ' + e);
    // Optionally, re-throw as a user error if you want to notify the user in Data Studio
    cc.newUserError()
      .setDebugText('Error sending to webhook. Details: ' + e)
      .setText('Failed to send the comment to the webhook URL.')
      .throwException();
  }
}

// Helper function to format the timestamp
function formatTimestamp(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2); // Months are zero-indexed
  var day = ('0' + date.getDate()).slice(-2);
  var hours = ('0' + date.getHours()).slice(-2);
  var minutes = ('0' + date.getMinutes()).slice(-2);
  var seconds = ('0' + date.getSeconds()).slice(-2);
  return year + '-' + month + '-' + day + ' ' + hours + ':' + minutes + ':' + seconds;
}

// Validate configuration parameters
function validateConfig(configParams) {
  configParams = configParams || {};
  configParams.comment = configParams.comment || 'No comment entered'; // Default value for "Comment"
  configParams.sendStatus = configParams.sendStatus || ''; // Default value for "Send Status"
  return configParams;
}

function isAdminUser() {
  return false;
}
