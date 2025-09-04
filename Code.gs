// 1. PASTE YOUR SPREADSHEET ID HERE
const spreadsheetId = "14QGIPHo6kf7xfoiEZgZQkKfN4z1UCWNPb4sp0tiVAZY"; 

// 2. DO NOT EDIT THE CODE BELOW THIS LINE
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

/**
 * Handles HTTP GET requests.
 * Used to fetch all data from a specific sheet.
 * @param {Object} e - The event parameter containing request details.
 * e.g., ?sheet=Customers or ?sheet=CustomerGroups
 */
function doGet(e) {
  try {
    const sheetName = e.parameter.sheet;
    if (!sheetName) {
      // If no sheet parameter, return a user-friendly HTML page.
      const htmlOutput = `
        <html>
          <body style="font-family: sans-serif; text-align: center; padding-top: 50px;">
            <h1>API Endpoint</h1>
            <p>This is the backend API for the customer management application.</p>
            <p>It is working correctly, but it is not meant to be accessed directly.</p>
            <p>Please open the <code>indexV9.html</code> file to use the application.</p>
          </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(htmlOutput)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Sheet not found: " + sheetName })).setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Get and remove header row
    
    const result = data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "data": result })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles HTTP POST requests.
 * Used to add a new row of data to a specific sheet.
 * The request body should be a JSON object.
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const sheetName = requestData.sheet;
    const dataRows = requestData.data; // This is an array of objects

    if (!sheetName) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Sheet name is missing in POST data." })).setMimeType(ContentService.MimeType.JSON);
    }
    if (!dataRows || !Array.isArray(dataRows)) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Data is missing or not an array in POST data." })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Sheet not found: " + sheetName })).setMimeType(ContentService.MimeType.JSON);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Process each row object sent from the client
    dataRows.forEach(rowData => {
      // Add a server-side timestamp
      rowData['วันที่สร้าง'] = new Date().toLocaleString("th-TH", { timeZone: 'Asia/Bangkok' });

      const newRow = headers.map(header => {
        let value = rowData[header];

        // Add a single quote to specific fields to force text format in Google Sheets
        const textFormatHeaders = ['ID/TAX ID', 'เบอร์โทรศัพท์', 'เบอร์มือถือ', 'รหัส'];
        if (textFormatHeaders.includes(header) && value) {
          return "'" + value;
        }

        return value !== undefined ? value : ""; // Map data to header order, use empty string if undefined
      });

      sheet.appendRow(newRow);
    });

    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "Data added successfully." })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Log the error for debugging
    console.error("doPost Error: " + error.toString() + " Stack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Saves new customer data sent from the web form via google.script.run.
 * @param {object} payload An object containing the sheet name and the data object to save.
 * @returns {object} A response object indicating success or failure.
 */
function saveCustomerData(payload) {
  try {
    const sheetName = payload.sheet;
    const dataToSave = payload.data;

    if (!sheetName || !dataToSave) {
      throw new Error("The request is missing the required 'sheet' or 'data' parameters.");
    }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet named '${sheetName}' could not be found.`);
    }

    // Get the headers from the first row of the sheet.
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Create a new row array in the correct column order.
    const newRow = headers.map(header => dataToSave[header] || "");
    
    // Append the new row to the sheet.
    sheet.appendRow(newRow);

    return { status: 'success', message: 'Data saved successfully.' };

  } catch (error) {
    Logger.log('saveCustomerData Error: ' + error.stack);
    // Return an error object that the client-side script can understand.
    return { status: 'error', message: error.message };
  }
}

/**
 * Allows including other server-side files (like CSS or JS) into the main HTML file.
 * @param {string} filename The name of the file to include.
 * @returns {string} The content of the included file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
