// Main GET handler: Serve the HTML page or fetch data dynamically.
function doGet(e) {
  const action = e?.parameter?.action || e?.action;
  const page = e?.parameter?.page || e?.page;

  // Handle fetching data from the spreadsheet dynamically
  if (action === "getData") {
    try {
      const data = getSheetData();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({
        error: "Failed to process data request",
        details: error.toString(),
        stack: error.stack
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } else if (action === "getRawData") {
    try {
      const rawData = getRawSheetData();
      return ContentService.createTextOutput(JSON.stringify(rawData))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({
        error: "Failed to get raw data",
        details: error.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Return the appropriate HTML page based on the 'page' parameter
  if (page === "dashboard") {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Ministry Treasury Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === "debug") {
    return HtmlService.createHtmlOutputFromFile('debug')
      .setTitle('Spreadsheet Debug Tool')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('reimbursement')
      .setTitle('Ministry Expense Reimbursement Form')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// Consolidate spreadsheet constants at the top for consistency
const SPREADSHEET_ID = '137QVV0PwLvwzG_XUJZ3lAAs4yinl5pxxUM5VMu9xI1A';
const UPLOAD_FOLDER_ID = '1FKeBJtE5ul8vbygraVxk0z_Erdkb7PjC';
const OUTPUT_FOLDER_ID = '1FKeBJtE5ul8vbygraVxk0z_Erdkb7PjC';
const SHEET_NAME = 'Reimbursement Data';
const REQUEST_COUNTER_SHEET = 'Config';
const TEMPLATE_DOC_ID = '113rMs76OrlckUWOW1Yzypgl4lUkY7EPM';

// Add Ministry Email mappings with the correct email addresses
const MINISTRY_EMAILS = {
  "MoT": "technology.ministry@ashoka.edu.in",
  "MAA": "academicaffairs.ministry@ashoka.edu.in",
  "MCWB": "cwb.ministry@ashoka.edu.in",
  "Sports": "sports.ministry@Ashoka.edu.in",
  "Jazbaa": "culturalministry@ashoka.edu.in",
  "Tarang": "environmentministry@ashoka.edu.in",
  "CLM": "campus.life@ashoka.edu.in",
  "SG": "sg@ashoka.edu.in",
  "PRD": "prd@ashoka.edu.in", 
  "EC": "auec@ashoka.edu.in"
};

// Handle POST requests for various actions.
function doPost(e) {
  try {
    const action = e.parameter.action;
    
    if (!action) {
      // Handle form submission when no action parameter is present
      const formData = JSON.parse(e.postData.contents);
      return ContentService.createTextOutput(
        JSON.stringify(processFormWithFiles(formData))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const payload = JSON.parse(e.postData.contents);

    if (action === "updateStatus") {
        // Use the unified updateRequestStatus function
        const data = JSON.parse(e.postData.contents);
        const requestId = data.requestId;
        const newStatus = data.status;
        const reason = data.reason;
        return ContentService.createTextOutput(
          JSON.stringify(updateRequestStatus(requestId, newStatus, reason))
        ).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "sendEmailNotification") {
      sendEmailNotification(payload.requestId, payload.status, payload.reason);
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, message: "Email notification sent." })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ error: "Invalid action specified" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error in POST request: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Add this function to test spreadsheet access directly 
function testSpreadsheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Log all sheet names to verify we're accessing the right spreadsheet
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    Logger.log("Available sheets: " + sheetNames.join(", "));
    // Try to access the specific sheet
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log("ERROR: 'Reimbursement Data' sheet not found. Available sheets: " + sheetNames);
      return { error: "Sheet not found", availableSheets: sheetNames };
    }
    // Get headers to check column names
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Headers in spreadsheet: " + headers.join(", "));
    // Check row count
    const lastRow = sheet.getLastRow();
    Logger.log("Total rows in sheet: " + lastRow);
    // Get a sample row if available
    if (lastRow > 1) {
      const sampleRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log("Sample data row: " + sampleRow.join(", "));
    }
    
    return {
      success: true,
      sheetNames: sheetNames,
      headers: headers,
      rowCount: lastRow
    };
  } catch (error) {
    Logger.log("ERROR accessing spreadsheet: " + error.toString());
    Logger.log("Stack trace: " + error.stack);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
* Gets data from the spreadsheet for the dashboard.
* This function processes and groups data by Request ID for the dashboard.
* @return {Array} The processed rows from the spreadsheet
*/
function getSheetData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name '${SHEET_NAME}' not found.`);
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) {
      Logger.log('No data found in the spreadsheet.');
      return [];
    }

    const headers = data[0];
    const rows = data.slice(1);

    return rows.map(row => {
      const rowData = {};
      headers.forEach((header, index) => {
        rowData[header] = row[index];
      });
      return rowData;
    });
  } catch (error) {
    Logger.log('Error in getSheetData:', error);
    return { error: error.message };
  }
}

/**
 * Gets raw data from the spreadsheet without any grouping or processing.
 * @return {Object} The raw rows and headers from the spreadsheet
 */
function getRawSheetData() {
  try {
    Logger.log("getRawSheetData function called - starting execution");
    
    // Verify spreadsheet ID is valid
    if (!SPREADSHEET_ID) {
      Logger.log("ERROR: Spreadsheet ID is missing");
      return { 
        error: "Spreadsheet ID is missing", 
        details: "Please check your configuration"
      };
    }
    
    // Access the spreadsheet with error handling
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      Logger.log(`Successfully opened spreadsheet with ID: ${SPREADSHEET_ID}`);
    } catch (ssError) {
      Logger.log(`ERROR: Failed to open spreadsheet: ${ssError}`);
      return { 
        error: "Failed to access spreadsheet", 
        details: "Check permissions and spreadsheet ID",
        message: ssError.toString()
      };
    }

    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      const availableSheets = spreadsheet.getSheets().map(s => s.getName()).join(", ");
      Logger.log(`ERROR: Sheet '${SHEET_NAME}' not found. Available sheets: ${availableSheets}`);
      return { 
        error: "Sheet not found", 
        details: `'${SHEET_NAME}' sheet is missing. Available sheets: ${availableSheets}`
      };
    }
    
    // Get all data including headers
    const dataRange = sheet.getDataRange();
    if (!dataRange) {
      Logger.log("ERROR: Could not get data range from sheet");
      return { 
        error: "Could not get data range", 
        details: "Sheet might be empty or inaccessible"
      };
    }
    
    const data = dataRange.getValues();
    
    // Ensure we always have valid headers and rows
    const headers = data && data.length > 0 ? data[0] : [];
    const rows = data && data.length > 1 ? data.slice(1) : [];
    
    Logger.log(`Retrieved ${rows.length} rows of data (excluding header)`);
    
    // Log a sample of what we're returning for debugging
    Logger.log(`Headers: ${headers.join(", ")}`);
    if (rows.length > 0) {
      Logger.log(`First row data sample: ${rows[0].slice(0, 3).join(", ")}...`);
    }
    
    // Return a clean, explicitly formatted object
    const result = {
      success: true,
      headers: headers,
      rows: rows
    };
    
    // Log what we're returning to make debugging easier
    Logger.log(`Returning data with ${headers.length} headers and ${rows.length} rows`);
    
    return result;
  } catch (error) {
    Logger.log("ERROR in getRawSheetData: " + error);
    Logger.log("Stack trace: " + error.stack);
    return { 
      error: "Error retrieving spreadsheet data", 
      details: error.message,
      stack: error.stack
    };
  }
}

// Add a new function to get ALL data from spreadsheet
function getAllSpreadsheetData() {
  try {
    Logger.log("getAllSpreadsheetData called");
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log("ERROR: Sheet not found");
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheetNames = spreadsheet.getSheets().map(s => s.getName()).join(", ");
      return { 
        error: "Sheet not found", 
        details: `'${SHEET_NAME}' sheet is missing. Available sheets: ${sheetNames}`
      };
    }
    
    // Get all data including headers
    const dataRange = sheet.getDataRange();
    if (!dataRange) {
      Logger.log("ERROR: Could not get data range from sheet");
      return { 
        error: "Could not get data range", 
        details: "Sheet might be empty or inaccessible"
      };
    }
    
    const data = dataRange.getValues();
    Logger.log(`Retrieved ${data.length} rows (including header)`);
    
    // Return all raw data - make sure we return a clean object that can be serialized
    return JSON.parse(JSON.stringify(data));
  } catch (error) {
    Logger.log("ERROR in getAllSpreadsheetData: " + error);
    Logger.log("Stack trace: " + error.stack);
    return { 
      error: "Error retrieving spreadsheet data", 
      details: error.message,
      stack: error.stack
    };
  }
}

// Process raw data from the spreadsheet - updated to match exact column order
function processRawData(data) {
  console.log('Processing raw data', data);
  
  // Verify data structure
  if (!data || !data.headers || !data.rows || !Array.isArray(data.headers) || !Array.isArray(data.rows)) {
    console.error('Invalid data structure:', data);
    return [];
  }
  
  // Define the exact column names from the spreadsheet in correct order
  const COLUMN_NAMES = {
    REQUEST_ID: 'Request ID',
    NAME: 'Name',
    EMAIL: 'Email',
    MINISTRY: 'Ministry',
    DESIGNATION: 'Designation',
    REGISTRATION_NUMBER: 'Registration Number',
    UPI_ID: 'UPI ID',
    PAYMENT_CATEGORY: 'Payment Category',
    AMOUNT: 'Amount',
    JUSTIFICATION: 'Justification',
    PAYMENT_MODE: 'Payment Mode',
    PAYMENT_DATE: 'Payment Date',
    RECEIPT_LINK: 'Receipt Link',
    APPROVAL_STATUS: 'Approval Status',
    DATE_OF_SUBMISSION: 'Date of Submission',
    REASON: 'Reason'
  };
  
  // Create column mapping with improved case-insensitive matching
  const colMap = {};
  const caseInsensitiveColMap = {};
  
  // Map each header to its column index - both exact and case-insensitive
  data.headers.forEach((header, index) => {
    if (!header) return;
    
    const headerStr = header.toString().trim();
    colMap[headerStr] = index; // Exact match
    caseInsensitiveColMap[headerStr.toLowerCase()] = index; // Case-insensitive match
  });
  
  // Log column data for debugging
  console.log('Column mapping (exact case):', colMap);
  console.log('Raw headers:', data.headers);
  
  // Group data by Request ID
  const requestsMap = new Map();
  
  data.rows.forEach((row, rowIndex) => {
    try {
      // Skip empty rows
      if (!row || row.length === 0 || row.every(cell => !cell)) return;
      
      // Improved helper to safely get values from columns (case-insensitive if needed)
      const getValue = (columnName) => {
        // Try exact match first
        if (colMap[columnName] !== undefined) {
          return row[colMap[columnName]];
        }
        // Try case-insensitive match next
        if (caseInsensitiveColMap[columnName.toLowerCase()] !== undefined) {
          return row[caseInsensitiveColMap[columnName.toLowerCase()]];
        }
        // Nothing found
        return null;
      };

      // Get Request ID using improved getValue helper
      const requestId = getValue(COLUMN_NAMES.REQUEST_ID);
      
      if (!requestId) {
        console.log(`Row ${rowIndex + 1} has no valid Request ID, skipping`);
        return;
      }
      
      // Log a sample of the first request for debugging
      if (rowIndex === 0) {
        console.log(`Sample request ID: ${requestId}`);
      }
      
      // Extract expense details using improved getValue helper
      const expense = {
        Category: getValue(COLUMN_NAMES.PAYMENT_CATEGORY) || '',
        Amount: parseFloat(getValue(COLUMN_NAMES.AMOUNT) || 0) || 0,
        Justification: getValue(COLUMN_NAMES.JUSTIFICATION) || '',
        PaymentMode: getValue(COLUMN_NAMES.PAYMENT_MODE) || '',
        PaymentDate: getValue(COLUMN_NAMES.PAYMENT_DATE) || '',
        ReceiptLink: getValue(COLUMN_NAMES.RECEIPT_LINK) || ''
      };
      
      // Extract other fields using improved getValue helper
      const name = getValue(COLUMN_NAMES.NAME) || '';
      const email = getValue(COLUMN_NAMES.EMAIL) || '';
      const upi = getValue(COLUMN_NAMES.UPI_ID) || '';
      const ministry = getValue(COLUMN_NAMES.MINISTRY) || '';
      const designation = getValue(COLUMN_NAMES.DESIGNATION) || '';
      const regNumber = getValue(COLUMN_NAMES.REGISTRATION_NUMBER) || '';
      const reason = getValue(COLUMN_NAMES.REASON) || '';
      const status = getValue(COLUMN_NAMES.APPROVAL_STATUS) || 'Pending';
      const submissionDate = getValue(COLUMN_NAMES.DATE_OF_SUBMISSION) || '';
      
      // Group by request ID
      if (requestsMap.has(requestId)) {
        const existingRequest = requestsMap.get(requestId);
        existingRequest.Expenses.push(expense);
        existingRequest.TotalAmount += expense.Amount;
      } else {
        requestsMap.set(requestId, {
          RequestID: requestId,
          Name: name,
          Email: email,
          UPI: upi,
          Ministry: ministry,
          Designation: designation,
          RegistrationNumber: regNumber,
          ApprovalStatus: status,
          SubmissionDate: submissionDate,
          Reason: reason,
          Expenses: [expense],
          TotalAmount: expense.Amount
        });
      }
    } catch (error) {
      console.error(`Error processing row ${rowIndex + 1}:`, error);
    }
  });
  
  // Return the grouped data as an array
  const result = Array.from(requestsMap.values());
  console.log(`Processed ${result.length} unique requests`);
  
  // Log the first record for debugging
  if (result.length > 0) {
    console.log('First processed record:', result[0]);
  }
  
  return result;
}

/**
 * Ensure a folder exists for the given request ID and return the folder object.
 * If the folder doesn't exist, it will be created inside the UPLOAD_FOLDER_ID.
 * @param {string} requestId - The unique request ID.
 * @return {Folder} The folder object for the request ID.
 */
function getOrCreateRequestFolder(requestId) {
  const parentFolder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  const folders = parentFolder.getFoldersByName(requestId);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create a new folder for the request ID
  return parentFolder.createFolder(requestId);
}

/**
 * Save a receipt file with chronological numbering in the request folder
 * @param {string} requestId - The unique request ID
 * @param {object} fileInfo - Object containing file data
 * @param {number} index - Index number for sequential naming
 * @return {object} File information including URL
 */
function saveReceiptWithSequentialName(requestId, fileInfo, index) {
  try {
    // Get or create the request-specific folder
    const requestFolder = getOrCreateRequestFolder(requestId);
    
    // Extract the file extension from the original filename
    const originalName = fileInfo.filename;
    const fileExtension = originalName.substring(originalName.lastIndexOf('.')) || '';
    
    // Create the new sequential filename (e.g., "1.pdf", "2.jpg")
    const newFileName = `${index + 1}${fileExtension}`;
    
    // Create a blob from the base64 data
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileInfo.data), 
      fileInfo.mimeType, 
      newFileName
    );

    // Save the file to the request-specific folder
    const file = requestFolder.createFile(blob);

    // Set file sharing permissions
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      Logger.log("Could not set file sharing permissions: " + e.toString());
      // Continue even if sharing permissions couldn't be set
    }

    Logger.log(`File saved with sequential name: ${newFileName}, link: ${file.getUrl()}`);
    return {
      name: newFileName,
      url: file.getUrl(),
      id: file.getId()
    };
  } catch (error) {
    Logger.log(`Error saving receipt file: ${error.toString()}`);
    throw error;
  }
}

/**
 * Compile all receipts in the request folder into a single document with embedded images.
 * Ensures proper formatting with each receipt on its own page, centered and properly sized.
 * @param {string} requestId - The unique request ID.
 * @param {Array} receiptLinks - Optional array of receipt file URLs to compile.
 * @return {string} - URL of the compiled document.
 */
function compileReceiptsPDF(requestId, receiptLinks) {
  try {
    Logger.log(`Compiling receipts with embedded images for request ID: ${requestId}`);
    const requestFolder = getOrCreateRequestFolder(requestId);
    const MAX_RECEIPTS = 10;
    
    // Create a Google Doc to hold the receipts
    const doc = DocumentApp.create(`${requestId}_Receipts_Compilation`);
    const body = doc.getBody();
    
    // Set page margins for better layout
    body.setMarginBottom(36).setMarginLeft(36).setMarginRight(36).setMarginTop(36);
    
    // Calculate available content area width (in points)
    const pageWidth = 8.5 * 72; // 8.5 inches converted to points
    const contentWidth = pageWidth - 72; // Accounting for left and right margins
    
    // Add a title
    body.appendParagraph(`Receipt Compilation for Request: ${requestId}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Track whether we added any receipts
    let receiptsAdded = 0;
    const processedFiles = [];
    
    // Process receipt links if provided
    if (receiptLinks && receiptLinks.length > 0) {
      Logger.log(`Processing ${receiptLinks.length} receipt links`);
      
      // Limit number of receipts if needed
      const processLinks = receiptLinks.length > MAX_RECEIPTS ? 
        receiptLinks.slice(0, MAX_RECEIPTS) : receiptLinks;
      
      // Add each receipt to the document
      body.appendParagraph("Attached Receipts:")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      for (let i = 0; i < processLinks.length; i++) {
        const link = processLinks[i];
        try {
          // Extract the file ID from the link
          const fileIdMatch = link.match(/[-\w]{25,}/);
          if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            if (file) {
              const receiptName = file.getName();
              const mimeType = file.getMimeType();
              
              // Add receipt information with centered alignment
              const headerText = `Receipt ${i+1}: ${receiptName}`;
              const headerPara = body.appendParagraph(headerText)
                .setHeading(DocumentApp.ParagraphHeading.HEADING3)
                .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              
              // Insert the actual content based on file type
              if (mimeType.startsWith('image/')) {
                // For images, insert them directly with proper sizing
                const image = file.getBlob();
                
                // Add image with center alignment
                const imageElement = body.appendImage(image);
                
                // Resize image to fit within page width if needed
                const imageWidth = imageElement.getWidth();
                const imageHeight = imageElement.getHeight();
                
                if (imageWidth > contentWidth) {
                  // Calculate new dimensions that preserve aspect ratio
                  const ratio = contentWidth / imageWidth;
                  const newWidth = Math.floor(imageWidth * ratio);
                  const newHeight = Math.floor(imageHeight * ratio);
                  
                  // Resize the image
                  imageElement.setWidth(newWidth);
                  imageElement.setHeight(newHeight);
                }
                
                // Center the image
                const imageParagraph = imageElement.getParent();
                if (imageParagraph) {
                  imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                }
                
                receiptsAdded++;
                processedFiles.push(file);
              } else if (mimeType === MimeType.PDF) {
                // For PDFs, provide a clear link
                body.appendParagraph(`PDF Receipt: ${receiptName}`)
                  .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                
                // Add a button-like link to the PDF
                const linkPara = body.appendParagraph("View PDF")
                  .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
                  .setLinkUrl(file.getUrl());
                
                // Style the link as a button
                linkPara.setFontFamily("Arial")
                      .setFontSize(12)
                      .setBold(true);
                
                // Save the file for later processing
                processedFiles.push(file);
                receiptsAdded++;
              } else {
                // For other file types, add a link
                body.appendParagraph(`File type not supported for embedding. View at: ${file.getUrl()}`)
                  .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
                  .setLinkUrl(file.getUrl());
              }
              
              // Only add a page break if this isn't the last receipt
              if (i < processLinks.length - 1) {
                body.appendPageBreak();
              }
            }
          }
        } catch (error) {
          Logger.log(`Error processing receipt link: ${error.toString()}`);
          body.appendParagraph(`Error processing receipt: ${error.toString()}`)
            .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          // Continue with other receipts
        }
      }
    } else {
      // No specific links provided, get all files from folder
      const files = requestFolder.getFiles();
      body.appendParagraph("Receipts in Request Folder:")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      let fileCount = 0;
      while (files.hasNext() && fileCount < MAX_RECEIPTS) {
        const file = files.next();
        const mimeType = file.getMimeType();
        
        // Only include PDFs and images
        if (mimeType === MimeType.PDF || mimeType.startsWith('image/')) {
          const receiptName = file.getName();
          
          // Add receipt information
          body.appendParagraph(`Receipt ${fileCount+1}: ${receiptName}`)
            .setHeading(DocumentApp.ParagraphHeading.HEADING3)
            .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          
          // Insert based on file type
          if (mimeType.startsWith('image/')) {
            // For images, insert directly with proper sizing
            const image = file.getBlob();
            const imageElement = body.appendImage(image);
            
            // Resize image to fit within page width if needed
            const imageWidth = imageElement.getWidth();
            const imageHeight = imageElement.getHeight();
            
            if (imageWidth > contentWidth) {
              // Calculate new dimensions that preserve aspect ratio
              const ratio = contentWidth / imageWidth;
              const newWidth = Math.floor(imageWidth * ratio);
              const newHeight = Math.floor(imageHeight * ratio);
              
              // Resize the image
              imageElement.setWidth(newWidth);
              imageElement.setHeight(newHeight);
            }
            
            // Center the image
            const imageParagraph = imageElement.getParent();
            if (imageParagraph) {
              imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            }
          } else if (mimeType === MimeType.PDF) {
            // For PDFs, provide a link
            body.appendParagraph(`PDF Receipt: ${receiptName}`)
              .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            
            body.appendParagraph(`View PDF: ${file.getUrl()}`)
              .setLinkUrl(file.getUrl())
              .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          }
          
          // Save for later processing
          processedFiles.push(file);
          
          // Only add a page break if this isn't the last file
          if (fileCount < MAX_RECEIPTS - 1 && files.hasNext()) {
            body.appendPageBreak();
          }
          
          fileCount++;
          receiptsAdded++;
        }
      }
      Logger.log(`Found ${fileCount} files in the request folder`);
    }
    
    // Save and close the document
    doc.saveAndClose();
    
    if (receiptsAdded === 0) {
      Logger.log('No valid receipts found to compile.');
      // Clean up the empty document
      DriveApp.getFileById(doc.getId()).setTrashed(true);
      return null;
    }
    
    // Instead of converting to PDF, get the document file and move it to the request folder
    const docFile = DriveApp.getFileById(doc.getId());
    const compilationDoc = docFile.makeCopy(`${requestId}_Receipts_Compilation`, requestFolder);
    
    // Remove the original document from root
    docFile.setTrashed(true);
    
    // Set sharing permissions on the document
    compilationDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    Logger.log(`Compiled receipts document saved: ${compilationDoc.getUrl()}`);
    return compilationDoc.getUrl();
  } catch (error) {
    Logger.log(`Error compiling receipts into document: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return null;
  }
}

/**
 * Sends an enhanced email notification to the requestor with relevant attachments
 * Includes improved error handling for when documents aren't available
 * @param {string} requestId - The request ID
 * @param {string} status - The approval status
 * @param {string} reason - The reason for rejection (if applicable)
 * @param {string} requestorEmail - The email address of the requestor
 * @param {string} requestorName - The name of the requestor
 * @param {string} ministry - The ministry name
 * @param {string} documentLink - URL to the approval document (if available)
 * @param {string} pdfLink - URL to the compiled receipts PDF (if available)
 * @param {number} totalAmount - The total approved amount
 * @param {string} folderLink - URL to the request folder containing all files (if available)
 * @return {Object} Result object indicating success or failure
 */
function sendEnhancedEmailNotification(
  requestId, status, reason, requestorEmail, requestorName, ministry, documentLink, pdfLink, totalAmount, folderLink
) {
  try {
    Logger.log(`Sending enhanced email notification for request ${requestId} to ${requestorEmail}`);
    Logger.log(`Links available - Document: ${!!documentLink ? 'Yes' : 'No'}, Receipts: ${!!pdfLink ? 'Yes' : 'No'}, Folder: ${!!folderLink ? 'Yes' : 'No'}`);
    
    if (!requestorEmail) {
      return {
        success: false,
        message: "No recipient email address provided"
      };
    }
    
    // Format amount properly
    const formattedAmount = totalAmount ? `₹${parseFloat(totalAmount).toFixed(2)}` : "N/A";
    
    // Create appropriate subject based on status
    const isApproved = status.toLowerCase().includes('accept') || status.toLowerCase().includes('approve');
    const emailSubject = isApproved 
      ? `Expense Reimbursement Approved (${requestId})`
      : `Expense Reimbursement Request ${status} (${requestId})`;
    
    // Create email body with HTML formatting
    let htmlBody = `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px;">`;
    htmlBody += `<h2 style="color: #333; border-bottom: 1px solid #ddd; padding-bottom: 10px;">Reimbursement Request ${status}</h2>`;
    
    // Different content based on approved/rejected status
    if (isApproved) {
      htmlBody += `<p>Dear ${requestorName || "Requestor"},</p>`;
      htmlBody += `<p>Your reimbursement request (ID: <strong>${requestId}</strong>) has been <span style="color: green; font-weight: bold;">APPROVED</span>.</p>`;
      htmlBody += `<p>The approved amount of <strong>${formattedAmount}</strong> will be processed for payment.</p>`;
      
      // Add document links if available
      if (documentLink) {
        Logger.log(`Including approval document link in email: ${documentLink}`);
        htmlBody += `<p><strong>Approval Document:</strong> <a href="${documentLink}" style="background-color: #4CAF50; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px 0;">View Approval Document</a></p>`;
      } else {
        Logger.log(`Warning: No approval document link available for request ${requestId}`);
        htmlBody += `<p><em>Note: The approval document is being processed and will be available soon. You can contact the treasurer for more information.</em></p>`;
      }
      
      // Add receipt PDF link if available
      if (pdfLink) {
        Logger.log(`Including receipt compilation link in email: ${pdfLink}`);
        htmlBody += `<p><strong>Compiled Receipts:</strong> <a href="${pdfLink}" style="background-color: #2196F3; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px 0;">View Receipts</a></p>`;
      }
      
      // Add folder link if available
      if (folderLink) {
        htmlBody += `<p><strong>All Request Files:</strong> <a href="${folderLink}" style="background-color: #FF9800; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px 0;">View All Files</a></p>`;
      }
      
      htmlBody += `<p>If you have any questions about when you'll receive your payment, please contact the treasury team.</p>`;
    } else {
      htmlBody += `<p>Dear ${requestorName || "Requestor"},</p>`;
      htmlBody += `<p>Your reimbursement request (ID: <strong>${requestId}</strong>) has been <span style="color: red; font-weight: bold;">REJECTED</span>.</p>`;
      
      // Include reason if provided
      if (reason) {
        htmlBody += `<p><strong>Reason for rejection:</strong> ${reason}</p>`;
      }
      
      // Add folder link even in rejection for access to submitted files
      if (folderLink) {
        htmlBody += `<p><strong>Your Submitted Files:</strong> <a href="${folderLink}" style="background-color: #FF9800; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px 0;">View Folder</a></p>`;
      }
      
      htmlBody += `<p>If you believe this is an error or want to resubmit with corrections, please contact the treasury team.</p>`;
    }
    
    // Add request details
    htmlBody += `<div style="background-color: #f9f9f9; padding: 15px; margin: 20px 0; border-left: 4px solid #ddd;">`;
    htmlBody += `<h3 style="margin-top: 0;">Request Details</h3>`;
    htmlBody += `<p><strong>Request ID:</strong> ${requestId}</p>`;
    htmlBody += `<p><strong>Ministry:</strong> ${ministry || "N/A"}</p>`;
    htmlBody += `<p><strong>Total Amount:</strong> ${formattedAmount}</p>`;
    htmlBody += `<p><strong>Status:</strong> ${status}</p>`;
    htmlBody += `</div>`;
    
    // Contact information
    htmlBody += `<div style="margin-top: 20px;">`;
    htmlBody += `<p><strong>Treasury Contact:</strong> <a href="mailto:treasury@ashoka.edu.in">treasury@ashoka.edu.in</a></p>`;
    htmlBody += `</div>`;
    
    // Footer
    htmlBody += `<p style="border-top: 1px solid #ddd; margin-top: 20px; padding-top: 10px; font-size: 12px; color: #777;">`;
    htmlBody += `This is an automated message from the Ashoka Student Government Treasury system. Please do not reply to this email.`;
    htmlBody += `</p>`;
    htmlBody += `</div>`;
    
    // Create plain text version for email clients that don't support HTML
    let plainBody = `Reimbursement Request ${status}\n\n`;
    plainBody += `Dear ${requestorName || "Requestor"},\n\n`;
    plainBody += isApproved 
      ? `Your reimbursement request (ID: ${requestId}) has been APPROVED.\n\n`
      : `Your reimbursement request (ID: ${requestId}) has been REJECTED.\n\n`;
    if (!isApproved && reason) {
      plainBody += `Reason for rejection: ${reason}\n\n`;
    }
    if (isApproved) {
      plainBody += `The approved amount of ${formattedAmount} will be processed for payment.\n\n`;
      if (documentLink) {
        plainBody += `Approval Document: ${documentLink}\n\n`;
      } else {
        plainBody += `Note: The approval document is being processed and will be available soon.\n\n`;
      }
      if (pdfLink) {
        plainBody += `Compiled Receipts: ${pdfLink}\n\n`;
      }
      if (folderLink) {
        plainBody += `All Request Files: ${folderLink}\n\n`;
      }
    } else if (folderLink) {
      plainBody += `Your Submitted Files: ${folderLink}\n\n`;
    }
    plainBody += `Request Details:\n`;
    plainBody += `- Request ID: ${requestId}\n`;
    plainBody += `- Ministry: ${ministry || "N/A"}\n`;
    plainBody += `- Total Amount: ${formattedAmount}\n`;
    plainBody += `- Status: ${status}\n\n`;
    plainBody += `Treasury Contact: treasury@ashoka.edu.in\n\n`;
    plainBody += `This is an automated message from the Ashoka Student Government Treasury system. Please do not reply to this email.`;
    
    // Prepare email options
    const emailOptions = {
      to: requestorEmail,
      subject: emailSubject,
      htmlBody: htmlBody,
      body: plainBody
    };
    
    // Send the email to the requestor
    try {
      MailApp.sendEmail(emailOptions);
      Logger.log(`Email sent to requestor: ${requestorEmail}`);
    } catch (e) {
      Logger.log(`Error sending email to requestor: ${e.toString()}`);
      return {
        success: false,
        message: `Failed to send email: ${e.toString()}`
      };
    }
    
    // Send a separate email to the ministry
    const ministryEmail = MINISTRY_EMAILS[ministry];
    if (ministryEmail) {
      // Create a copy of the email options for the ministry
      const emailOptionsMinistry = { ...emailOptions };
      emailOptionsMinistry.to = ministryEmail;

      try {
        MailApp.sendEmail(emailOptionsMinistry);
        Logger.log(`Email sent to ministry: ${ministryEmail}`);
      } catch (e) {
        Logger.log(`Error sending email to ministry: ${e.toString()}`);
      }
    }
    
    return {
      success: true,
      message: `Email notification sent successfully to ${requestorEmail}`,
      attachmentsCount: 0,
      recipient: requestorEmail,
      status: status
    };
  } catch (error) {
    Logger.log(`ERROR in sendEnhancedEmailNotification: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return {
      success: false,
      message: `Failed to send email notification: ${error.toString()}`,
      error: error.toString()
    };
  }
}

/**
 * Ensure a folder exists for the given request ID and return the folder object.
 * If the folder doesn't exist, it will be created inside the UPLOAD_FOLDER_ID.
 * @param {string} requestId - The unique request ID.
 * @return {Folder} The folder object for the request ID.
 */
function getOrCreateRequestFolder(requestId) {
  const parentFolder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  const folders = parentFolder.getFoldersByName(requestId);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create a new folder for the request ID
  return parentFolder.createFolder(requestId);
}

// Enhanced function to update the status of a request - with async processing
function updateRequestStatus(requestData) {
  try {
    // Normalize parameters to object format for consistency
    let requestId, newStatus, reason, asyncProcessing;
    if (typeof requestData === 'object') {
      requestId = requestData.requestId;
      newStatus = requestData.status;
      reason = requestData.reason || '';
      asyncProcessing = requestData.asyncProcessing || false;
    } else {
      // Legacy format - convert to standardized format
      requestId = arguments[0];
      newStatus = arguments[1];
      reason = arguments[2] || '';
      asyncProcessing = false;
      // Log deprecation warning
      Logger.log("WARNING: Using deprecated parameter format for updateRequestStatus. Please use object format.");
    }

    const requestFolder = getOrCreateRequestFolder(requestId);
    Logger.log(`Updating request status: ID=${requestId}, Status=${newStatus}, Reason=${reason}, Async=${asyncProcessing}`);

    if (!requestId || !newStatus) {
      return { 
        success: false,
        message: "Request ID and new status are required" 
      };
    }

    // Get the spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log("ERROR: Sheet not found");
      return { 
        success: false,
        message: "Sheet not found. Check configuration." 
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find the column indices for the fields we need to update
    const requestIdColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('request') && 
      header.toString().toLowerCase().includes('id'));
    const statusColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('approval') && 
      header.toString().toLowerCase().includes('status'));
    const reasonColIndex = headers.findIndex(header => 
      header.toString().toLowerCase() === 'reason');

    if (requestIdColIndex === -1 || statusColIndex === -1) {
      Logger.log(`ERROR: Required columns not found. requestIdColIndex=${requestIdColIndex}, statusColIndex=${statusColIndex}`);
      return { 
        success: false,
        message: "Required columns not found in the spreadsheet. Check column names." 
      };
    }

    // Find the rows that match the request ID
    let rowsUpdated = 0;
    let requestorEmail = null;
    let requestorName = null;
    let ministry = null;
    let designation = null;
    let registrationNumber = null;
    let expenseDetails = [];
    let receiptLinks = [];
    let totalAmount = 0;

    // Get all relevant column indices
    const emailColIndex = headers.findIndex(header => header.toString().toLowerCase() === 'email');
    const nameColIndex = headers.findIndex(header => header.toString().toLowerCase() === 'name');
    const ministryColIndex = headers.findIndex(header => header.toString().toLowerCase() === 'ministry');
    const amountColIndex = headers.findIndex(header => header.toString().toLowerCase() === 'amount');
    const categoryColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('category') || 
      header.toString().toLowerCase().includes('expense category') || 
      header.toString().toLowerCase().includes('payment category'));
    const receiptColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('receipt'));
    const justificationColIndex = headers.findIndex(header => 
      header.toString().toLowerCase() === 'justification');
    const paymentDateColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('payment date') || 
      header.toString().toLowerCase().includes('expense date'));
    const designationColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('designation'));
    const regNumberColIndex = headers.findIndex(header => 
      header.toString().toLowerCase().includes('registration number') || 
      header.toString().toLowerCase().includes('reg number'));

    // Gather all data for this request - this is fast and needs to be done synchronously
    for (let i = 1; i < data.length; i++) {
      if (data[i][requestIdColIndex] === requestId) {
        // Update status column
        sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
        // Update reason column if it exists and a reason is provided
        if (reasonColIndex !== -1 && reason) {
          sheet.getRange(i + 1, reasonColIndex + 1).setValue(reason);
        }

        // Collect requestor information
        if (emailColIndex !== -1 && !requestorEmail) {
          requestorEmail = data[i][emailColIndex];
        }
        if (nameColIndex !== -1 && !requestorName) {
          requestorName = data[i][nameColIndex];
        }
        if (ministryColIndex !== -1 && !ministry) {
          ministry = data[i][ministryColIndex];
        }
        if (designationColIndex !== -1 && !designation) {
          designation = data[i][designationColIndex];
        }
        if (regNumberColIndex !== -1 && !registrationNumber) {
          registrationNumber = data[i][regNumberColIndex];
        }

        // Collect expense details and receipt links
        if (amountColIndex !== -1) {
          const amount = parseFloat(data[i][amountColIndex]) || 0;
          totalAmount += amount;
          const expenseDetail = {
            amount: amount,
            category: categoryColIndex !== -1 ? data[i][categoryColIndex] : 'Not specified',
            justification: justificationColIndex !== -1 ? data[i][justificationColIndex] : '',
            paymentDate: paymentDateColIndex !== -1 ? data[i][paymentDateColIndex] : '',
            billNumber: '', // Default empty, could extract from justification in some cases
            receiptLink: receiptColIndex !== -1 ? data[i][receiptColIndex] : ''
          };

          // Extract possible bill number from justification if present (format: "Bill #12345")
          if (expenseDetail.justification && expenseDetail.justification.includes('Bill #')) {
            const match = expenseDetail.justification.match(/Bill #([A-Za-z0-9-]+)/);
            if (match && match[1]) {
              expenseDetail.billNumber = match[1];
            }
          }

          expenseDetails.push(expenseDetail);
          if (expenseDetail.receiptLink) {
            receiptLinks.push(expenseDetail.receiptLink);
          }
        }
        rowsUpdated++;
      }
    }

    if (rowsUpdated === 0) {
      Logger.log(`ERROR: Request with ID ${requestId} not found in the spreadsheet`);
      return { 
        success: false,
        message: `Request with ID ${requestId} not found in the spreadsheet`
      };
    }

    Logger.log(`Successfully updated ${rowsUpdated} rows for request ${requestId}`);

    // For async processing, immediately return success and handle the rest in a separate process
    if (asyncProcessing) {
      // Create a task ID for tracking
      const taskId = `task_${requestId}_${new Date().getTime()}`;
      // Store task in Properties service or Cache service for status tracking
      const taskData = {
        requestId: requestId,
        status: newStatus,
        reason: reason,
        requestorEmail: requestorEmail,
        requestorName: requestorName,
        ministry: ministry,
        totalAmount: totalAmount,
        started: new Date().toISOString(),
        completed: false
      };
      // Store in Properties service (limited to 9KB per property)
      PropertiesService.getScriptProperties().setProperty(taskId, JSON.stringify(taskData));

      // Start the background processing using time-based trigger
      // This avoids UI timeouts since the docs & email generation can be slow
      const trigger = ScriptApp.newTrigger('processApprovalBackgroundTask')
        .timeBased()
        .after(1000) // Run 1 second later
        .create();

      // Store trigger ID with task data
      const updatedTask = JSON.parse(PropertiesService.getScriptProperties().getProperty(taskId));
      updatedTask.triggerId = trigger.getUniqueId();
      PropertiesService.getScriptProperties().setProperty(taskId, JSON.stringify(updatedTask));
      
      // Also store the task ID in a property named after the request for fast lookups
      PropertiesService.getScriptProperties().setProperty(`req_${requestId}`, taskId);

      // Return immediately with the task ID for polling
      return { 
        success: true,
        message: `Request status updated to ${newStatus}. Document generation and email notifications in progress.`,
        asyncTaskId: taskId,
        rowsUpdated: rowsUpdated
      };
    }

    // If not async, continue with synchronous processing as before
    // Additional processing for approvals - create document and compile receipts
    let documentLink = null;
    let pdfLink = null;
    let folderLink = null;
    let documentCreationResults = { success: false };
    let receiptCompilationResults = { success: false };

    // Get folder link regardless of status
    try {
      const requestFolder = getOrCreateRequestFolder(requestId);
      folderLink = requestFolder.getUrl();
      Logger.log(`Got folder link: ${folderLink}`);
    } catch (folderError) {
      Logger.log(`Warning: Failed to get folder link: ${folderError.toString()}`);
      // Continue even if getting folder link fails
    }

    if (newStatus.toLowerCase().includes('accept') || newStatus.toLowerCase().includes('approve')) {
      try {
        // Create approval document from template
        Logger.log(`Creating approval document for request ${requestId}`);
        const documentInfo = createApprovalDocument(requestId, {
          requestorName,
          ministry,
          amount: totalAmount,
          designation,
          registrationNumber,
          expenses: expenseDetails
        });
        documentLink = documentInfo.url;
        documentCreationResults = {
          success: true,
          url: documentLink,
          id: documentInfo.id
        };
        Logger.log(`Successfully created approval document: ${documentLink}`);

        // Compile receipts if available
        if (receiptLinks.length > 0) {
          Logger.log(`Compiling ${receiptLinks.length} receipts into PDF for request ${requestId}`);
          try {
            pdfLink = compileReceiptsPDF(requestId, receiptLinks);
            Logger.log(`Receipt compilation result: ${pdfLink ? "Success" : "Failed"}`);
          } catch (receiptError) {
            Logger.log(`Error in receipt compilation: ${receiptError.toString()}`);
            // Continue even if receipt compilation fails
          }

          if (pdfLink) {
            receiptCompilationResults = {
              success: true,
              url: pdfLink,
              count: receiptLinks.length
            };
            Logger.log(`Successfully compiled receipts into PDF: ${pdfLink}`);
          } else {
            receiptCompilationResults = {
              success: false,
              error: "PDF compilation failed",
              count: receiptLinks.length
            };
            Logger.log(`Failed to compile receipts into PDF`);
          }
        } else {
          Logger.log(`No receipts to compile for request ${requestId}`);
          receiptCompilationResults = {
            success: true,
            message: "No receipts to compile",
            count: 0
          };
        }
      } catch (docError) {
        Logger.log(`WARNING: Failed to create approval documents: ${docError.toString()}`);
        documentCreationResults = {
          success: false,
          error: docError.toString()
        };
        // Continue processing even if document creation fails
      }
    }

    // Send enhanced email notification to the requestor
    let emailResults = { success: false };
    if (requestorEmail) {
      try {
        Logger.log(`Sending email notification to ${requestorEmail}`);
        emailResults = sendEnhancedEmailNotification(
          requestId,
          newStatus,
          reason,
          requestorEmail,
          requestorName,
          ministry,
          documentLink,
          pdfLink,
          totalAmount,
          folderLink  // Add folder link to the email
        );
        if (emailResults.success) {
          Logger.log(`Enhanced email notification sent to ${requestorEmail} with ${emailResults.attachmentsCount || 0} attachments`);
        } else {
          Logger.log(`Failed to send email notification: ${emailResults.message}`);
        }
      } catch (emailError) {
        Logger.log(`WARNING: Failed to send enhanced email notification: ${emailError.toString()}`);
        emailResults = {
          success: false,
          error: emailError.toString()
        };
        // Continue processing even if email fails
      }
    } else {
      Logger.log(`No email address available for notification`);
      emailResults = {
        success: false,
        error: "No email address available"
      };
    }

    return {
      success: true,
      message: `Request status updated successfully. ${rowsUpdated} rows affected.`,
      documentLink: documentLink,
      receiptsPDF: pdfLink,
      folderLink: folderLink,  // Include folder link in the response
      emailNotification: emailResults,
      documentCreation: documentCreationResults,
      receiptCompilation: receiptCompilationResults
    };
  } catch (error) {
    Logger.log(`ERROR in updateRequestStatus: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return {
      success: false,
      message: `Error updating request status: ${error.toString()}`
    };
  }
}

// Function to process approval tasks in the background
function processApprovalBackgroundTask() {
  try {
    Logger.log("Starting background task processing");

    // Get all triggers and find the one that called this function
    const allTriggers = ScriptApp.getProjectTriggers();
    const currentTrigger = allTriggers.find(trigger => 
      trigger.getHandlerFunction() === 'processApprovalBackgroundTask');

    if (currentTrigger) {
      // Delete the trigger once we're running
      ScriptApp.deleteTrigger(currentTrigger);

      // Find which task belongs to this trigger
      const triggerId = currentTrigger.getUniqueId();
      Logger.log(`Processing task with trigger ID: ${triggerId}`);

      // Scan properties to find task with this trigger ID
      const scriptProperties = PropertiesService.getScriptProperties();
      const allProps = scriptProperties.getProperties();
      let taskId = null;
      let taskData = null;

      // Find the task with this trigger ID
      for (const key in allProps) {
        if (key.startsWith('task_')) {
          try {
            const data = JSON.parse(allProps[key]);
            if (data && data.triggerId === triggerId) {
              taskId = key;
              taskData = data;
              break;
            }
          } catch (e) {
            Logger.log(`Error parsing task data for ${key}: ${e}`);
          }
        }
      }

      if (!taskId || !taskData) {
        Logger.log("No task found for this trigger");
        return;
      }

      Logger.log(`Found task ${taskId} for request ${taskData.requestId}`);

      // Now process the task based on the status
      if (taskData.status.toLowerCase().includes('accept') || 
          taskData.status.toLowerCase().includes('approve')) {
        // Process approval
        processApprovalTask(taskId, taskData);
      } else {
        // Process rejection (just email notification)
        processRejectionTask(taskId, taskData);
      }
    } else {
      Logger.log("Trigger not found for this execution");
    }
  } catch (error) {
    Logger.log(`ERROR in processApprovalBackgroundTask: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

// Function to process approval tasks - FIXED to handle document creation failures better
function processApprovalTask(taskId, taskData) {
  try {
    Logger.log(`Processing approval task for request ${taskData.requestId}`);
    const scriptProperties = PropertiesService.getScriptProperties();
    let documentLink = null;
    let documentId = null;
    let receiptsPDF = null;
    let folderLink = null;

    // Get the request folder and its URL
    try {
      const requestFolder = getOrCreateRequestFolder(taskData.requestId);
      folderLink = requestFolder.getUrl();
      taskData.folderLink = folderLink;
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (folderError) {
      Logger.log(`ERROR getting folder link: ${folderError}`);
      taskData.folderError = folderError.toString();
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    }

    // Declare receiptLinks for use in this function
    let receiptLinks = [];

    // 1. Create approval document - FIXED to handle failures better
    try {
      Logger.log(`Creating approval document for request ${taskData.requestId}`);
      
      // Get additional data needed for document creation
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      const requestIdColIndex = headers.findIndex(header => 
        header.toString().toLowerCase().includes('request') && 
        header.toString().toLowerCase().includes('id'));
      const amountColIndex = headers.findIndex(header => header.toString().toLowerCase() === 'amount');
      const categoryColIndex = headers.findIndex(header => header.toString().toLowerCase().includes('category'));
      const justificationColIndex = headers.findIndex(header => header.toString().toLowerCase().includes('justification'));
      const paymentDateColIndex = headers.findIndex(header => 
        header.toString().toLowerCase().includes('payment date') || 
        header.toString().toLowerCase().includes('expense date'));
      const receiptColIndex = headers.findIndex(header => header.toString().toLowerCase().includes('receipt'));

      let expenseDetails = [];

      // Gather expense details and also collect receipt links from the spreadsheet
      for (let i = 1; i < data.length; i++) {
        if (data[i][requestIdColIndex] === taskData.requestId) {
          const amount = parseFloat(data[i][amountColIndex]) || 0;
          const expenseDetail = {
            Amount: amount,
            Category: categoryColIndex !== -1 ? data[i][categoryColIndex] : 'Not specified',
            Justification: justificationColIndex !== -1 ? data[i][justificationColIndex] : '',
            PaymentDate: paymentDateColIndex !== -1 ? data[i][paymentDateColIndex] : '',
            billNumber: '', // Default empty, could extract from justification in some cases
            receiptLink: receiptColIndex !== -1 ? data[i][receiptColIndex] : ''
          };
          expenseDetails.push(expenseDetail);
          if (expenseDetail.receiptLink) {
            receiptLinks.push(expenseDetail.receiptLink);
          }
        }
      }

      // Now try to create the approval document with the enhanced function
      const documentInfo = createApprovalDocument(taskData.requestId, {
        requestorName: taskData.requestorName,
        ministry: taskData.ministry,
        amount: taskData.totalAmount,
        designation: taskData.designation,
        registrationNumber: taskData.registrationNumber,
        expenses: expenseDetails
      });
      documentLink = documentInfo.url;
      documentId = documentInfo.id;
      taskData.documentLink = documentLink;
      taskData.documentCreated = true;
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (docError) {
      Logger.log(`ERROR creating document: ${docError}`);
      taskData.documentError = docError.toString();
      taskData.documentCreationAttempted = true;  // Mark that we tried
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
      // Even if document creation fails, continue with other steps
    }

    // Verify the link immediately
    try {
      const docFile = DriveApp.getFileById(documentId);
      if (!docFile) {
        Logger.log(`WARNING: Created document with ID ${documentId} couldn't be accessed`);
        taskData.documentError = "Document creation failed - no URL returned";
        scriptProperties.setProperty(taskId, JSON.stringify(taskData));
      }
    } catch (verifyError) {
      Logger.log(`ERROR verifying document link: ${verifyError.toString()}`);
      taskData.documentError = "Document creation failed - no URL returned";
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    }

    // 2. Process receipt compilation - With better error handling
    try {
      Logger.log(`Compiling receipts for request ${taskData.requestId}`);
      
      // If no receipt links were captured in the loop above, explicitly search the sheet
      if (!receiptLinks || receiptLinks.length === 0) {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_NAME);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        const requestIdColIndex = headers.findIndex(header => 
          header.toString().toLowerCase().includes('request') && 
          header.toString().toLowerCase().includes('id'));
        const receiptColIndex = headers.findIndex(header => 
          header.toString().toLowerCase().includes('receipt'));

        receiptLinks = [];
        for (let i = 1; i < data.length; i++) {
          if (data[i][requestIdColIndex] === taskData.requestId &&
              receiptColIndex !== -1 &&
              data[i][receiptColIndex]) {
            receiptLinks.push(data[i][receiptColIndex]);
          }
        }
      }

      if (receiptLinks && receiptLinks.length > 0) {
        // Try to compile the receipts
        receiptsPDF = compileReceiptsPDF(taskData.requestId, receiptLinks);

        if (receiptsPDF) {
          taskData.receiptsPDF = receiptsPDF;
          taskData.receiptsCompiled = true;
        } else {
          taskData.receiptsError = "Failed to compile receipts - no PDF was generated";
        }
      } else {
        taskData.receiptsInfo = "No receipt links found to compile";
      }
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (pdfError) {
      Logger.log(`ERROR compiling receipts: ${pdfError}`);
      taskData.receiptsError = pdfError.toString();
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
      // Continue even if receipt compilation fails
    }

    // 3. Send email notification - WITH attachments even if document creation failed
    try {
      Logger.log(`Sending approval notification for request ${taskData.requestId}`);
      Logger.log(`Links to include - Document: ${documentLink || 'None'}, Receipts: ${receiptsPDF || 'None'}`);
      
      const emailResult = sendEnhancedEmailNotification(
        taskData.requestId,
        taskData.status,
        taskData.reason,
        taskData.requestorEmail,
        taskData.requestorName,
        taskData.ministry,
        documentLink,
        receiptsPDF,
        taskData.totalAmount,
        folderLink  // Add folder link to the email
      );

      taskData.emailSent = emailResult.success;
      taskData.emailResult = emailResult;
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (emailError) {
      Logger.log(`ERROR sending email: ${emailError}`);
      taskData.emailError = emailError.toString();
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    }

    // Mark task as completed
    taskData.completed = true;
    taskData.completedTimestamp = new Date().toISOString();
    scriptProperties.setProperty(taskId, JSON.stringify(taskData));

    Logger.log(`Completed processing approval task for ${taskData.requestId}`);
    return true;
  } catch (error) {
    Logger.log(`ERROR in processApprovalTask: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return false;
  }
}

// Function to process rejection tasks - FIXED to actually send emails
function processRejectionTask(taskId, taskData) {
  try {
    Logger.log(`Processing rejection task for request ${taskData.requestId}`);
    const scriptProperties = PropertiesService.getScriptProperties();
    let folderLink = null;

    // Get the request folder and its URL
    try {
      const requestFolder = getOrCreateRequestFolder(taskData.requestId);
      folderLink = requestFolder.getUrl();
      taskData.folderLink = folderLink;
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (folderError) {
      Logger.log(`ERROR getting folder link: ${folderError}`);
      taskData.folderError = folderError.toString();
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    }

    // Send email notification for rejection
    try {
      Logger.log(`Sending rejection notification for request ${taskData.requestId}`);
      Logger.log(`Links to include - Document: None, Receipts: None, Folder: ${folderLink || 'None'}`);
      
      // Actually send the email notification
      const emailResult = sendEnhancedEmailNotification(
        taskData.requestId,
        taskData.status,
        taskData.reason,
        taskData.requestorEmail,
        taskData.requestorName,
        taskData.ministry,
        null, // No document link for rejections
        null, // No receipts PDF for rejections
        taskData.totalAmount,
        folderLink // Add folder link to the rejection email
      );

      // Update task data with progress
      taskData.emailSent = emailResult.success;
      taskData.emailResult = emailResult;
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    } catch (emailError) {
      Logger.log(`ERROR sending rejection email: ${emailError}`);
      taskData.emailError = emailError.toString();
      scriptProperties.setProperty(taskId, JSON.stringify(taskData));
    }

    // Mark task as completed
    taskData.completed = true;
    taskData.completedTimestamp = new Date().toISOString();
    scriptProperties.setProperty(taskId, JSON.stringify(taskData));

    Logger.log(`Completed processing rejection task for ${taskData.requestId}`);
    return true;
  } catch (error) {
    Logger.log(`ERROR in processRejectionTask: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return false;
  }
}

// Function to check the status of an async task
function checkAsyncTaskStatus(taskId) {
  try {
    Logger.log(`Checking status of task: ${taskId}`);
    if (!taskId || !taskId.startsWith('task_')) {
      return { 
        error: "Invalid task ID", 
        completed: false 
      };
    }

    // Get task data from Properties
    const taskDataString = PropertiesService.getScriptProperties().getProperty(taskId);
    if (!taskDataString) {
      return { 
        error: "Task not found", 
        completed: false 
      };
    }

    // Parse task data
    const taskData = JSON.parse(taskDataString);

    // Return task status
    return { 
      requestId: taskData.requestId,
      started: taskData.started,
      completed: taskData.completed,
      documentCreated: taskData.documentCreated || false,
      documentLink: taskData.documentLink || null,
      receiptsCompiled: taskData.receiptsCompiled || false,
      receiptsPDF: taskData.receiptsPDF || null,
      emailSent: taskData.emailSent || false,
      errors: {
        document: taskData.documentError || null,
        receipts: taskData.receiptsError || null,
        email: taskData.emailError || null
      }
    };
  } catch (error) {
    Logger.log(`ERROR in checkAsyncTaskStatus: ${error.toString()}`);
    return { 
      error: error.toString(), 
      completed: false 
    };
  }
}

/**
 * Process form submission with file uploads - updated to save files with sequential naming.
 */
function processFormWithFiles(formData) {
  try {
    Logger.log("Received form data: " + JSON.stringify(formData.metadata || formData));

    // Access the spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet not found. Check the sheet name.");
    }

    // Initialize headers if the sheet is empty - using exact column order
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Request ID", "Name", "Email", "Ministry", "Designation", "Registration Number", 
        "UPI ID", "Payment Category", "Amount", "Justification", "Payment Mode", 
        "Payment Date", "Receipt Link", "Approval Status", "Date of Submission", "Reason"
      ]);
    }

    // Generate a unique request ID
    const timestamp = new Date();
    const requestId = "REQ-" + timestamp.getTime();
    Logger.log("Generated Request ID: " + requestId);

    // Extract requestor details
    const name = formData.requestor ? formData.requestor.name : formData.name;
    const email = formData.requestor ? formData.requestor.email : formData.email;
    const ministry = formData.ministry;
    const pocDesignation = formData.requestor ? formData.requestor.pocDesignation : formData.pocDesignation;
    const pocReg = formData.requestor ? formData.requestor.pocReg : formData.pocReg;
    const pocUpi = formData.requestor ? formData.requestor.pocUpi : formData.pocUpi;

    // Process each expense entry
    const expenses = formData.expenses || [];
    const fileData = formData.files || [];
    let totalAmount = 0;

    if (expenses.length === 0) {
      Logger.log("No expenses found in the request");
      return { 
        success: false, 
        message: "No expense data found in your request" 
      };
    }

    Logger.log(`Processing ${expenses.length} expenses with ${fileData.length} file uploads`);

    // Associate files with expenses based on position if fileData exists
    expenses.forEach((expense, index) => {
      totalAmount += parseFloat(expense.amount) || 0;
      let receiptLink = "";

      // Handle file upload if available
      if (fileData && fileData[index]) {
        try {
          // Save file with sequential naming
          const fileInfo = saveReceiptWithSequentialName(requestId, fileData[index], index);
          receiptLink = fileInfo.url;
          Logger.log(`File uploaded as ${fileInfo.name}: ${receiptLink}`);
        } catch (fileError) {
          Logger.log(`Error uploading file for expense ${index + 1}: ${fileError.toString()}`);
          // Continue without receipt if file upload fails
        }
      }

      // Append the row with expense details - in exact column order
      sheet.appendRow([
        requestId,                         // Request ID
        name,                              // Name
        email,                             // Email
        ministry,                          // Ministry
        pocDesignation,                    // Designation
        pocReg,                            // Registration Number
        pocUpi,                            // UPI ID
        expense.category,                  // Payment Category
        expense.amount,                    // Amount
        expense.justification,             // Justification
        expense.paymentMode,               // Payment Mode
        expense.expenseDate,               // Payment Date
        receiptLink,                       // Receipt Link
        "Pending",                         // Approval Status
        timestamp,                         // Date of Submission
        ""                                 // Reason (empty initially)
      ]);

      Logger.log(`Added expense: ${expense.category}, Amount: ${expense.amount}`);
    });

    // Send confirmation email to the requester
    try {
      // Format the expense list for the email
      const expenseList = expenses.map((expense, index) => 
        `${index + 1}. ${expense.category}: ₹${expense.amount.toFixed(2)}`
      ).join("\n");

      const emailSubject = `Reimbursement Request Submitted (ID: ${requestId})`;
      let emailBody = `Dear ${name},\n\n`;
      emailBody += `Your reimbursement request (ID: ${requestId}) has been received and is pending review.\n\n`;
      emailBody += `Ministry: ${ministry}\n`;
      emailBody += `Total Amount: ₹${totalAmount.toFixed(2)}\n\n`;
      emailBody += `Expenses:\n${expenseList}\n\n`;
      emailBody += `You will be notified when your request is approved or declined.\n\n`;
      emailBody += `Thank you,\nAshoka Student Government Treasury`;

      // Try-catch for better error handling around email
      try {
        MailApp.sendEmail({
          to: email,
          subject: emailSubject,
          body: emailBody
        });
        Logger.log(`Confirmation email sent to ${email}`);
      } catch (specificEmailError) {
        Logger.log(`Failed to send confirmation email - specific error: ${specificEmailError.toString()}`);
        // Continue processing even if email fails
      }
    } catch (emailError) {
      Logger.log(`Warning: Failed to prepare or send confirmation email: ${emailError.toString()}`);
      Logger.log(`Stack trace: ${emailError.stack}`);
      // Continue processing even if email fails
    }

    // Notify the appropriate ministry if applicable
    if (ministry && MINISTRY_EMAILS[ministry]) {
      try {
        const ministryEmail = MINISTRY_EMAILS[ministry];
        const ministrySubject = `New Reimbursement Request ${requestId}`;
        let ministryBody = `A new reimbursement request has been submitted by ${name}.\n\n`;
        ministryBody += `Request ID: ${requestId}\n`;
        ministryBody += `Total Amount: ₹${totalAmount.toFixed(2)}\n\n`;
        ministryBody += `This is an automatic notification for your records. The request is pending approval by the treasurer.\n\n`;
        ministryBody += `Thank you,\nAshoka Student Government Treasury`;

        try {
          MailApp.sendEmail({
            to: ministryEmail,
            subject: ministrySubject,
            body: ministryBody
          });
          Logger.log(`Ministry notification sent to ${ministryEmail}`);
        } catch (specificMinistryEmailError) {
          Logger.log(`Failed to send ministry email - specific error: ${specificMinistryEmailError.toString()}`);
        }
      } catch (ministryEmailError) {
        Logger.log(`Warning: Failed to send ministry notification: ${ministryEmailError.toString()}`);
        // Continue processing even if email fails
      }
    }

    // Return success with the request ID
    return { 
      success: true, 
      requestId: requestId,
      message: "Reimbursement request submitted successfully" 
    };
  } catch (error) {
    Logger.log("Error in processFormWithFiles: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    return { 
      success: false, 
      message: error.message || "An error occurred during form submission" 
    };
  }
}

/**
 * Create an approval document and save it in the request-specific folder.
 * Generates the document from scratch instead of using a template.
 * @param {string} requestId - The unique request ID
 * @param {object} requestData - The request data including expenses
 * @return {object} Document info with ID and URL
 */
function createApprovalDocument(requestId, requestData) {
  try {
    Logger.log(`Creating approval document for request ID: ${requestId}`);

    // Get or create the folder for this request
    const requestFolder = getOrCreateRequestFolder(requestId);

    // Create a new Google Doc
    const doc = DocumentApp.create("Out of Pocket Expense Reimbursement Form");
    const body = doc.getBody();

    // Set page margins
    body.setMarginBottom(36).setMarginLeft(36).setMarginRight(36).setMarginTop(36);

    // Add header with underline
    const headerParagraph = body.appendParagraph("ASHOKA UNIVERSITY");
    headerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    headerParagraph.setFontSize(16);
    headerParagraph.setBold(true);
    headerParagraph.setUnderline(true);

    const subheaderParagraph = body.appendParagraph("Out of Pocket Expenses Reimbursement Form");
    subheaderParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subheaderParagraph.setFontSize(14);
    subheaderParagraph.setBold(true);
    subheaderParagraph.setUnderline(true);

    // Add space
    body.appendParagraph("\n");

    // Add today's date
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, "GMT+5:30", "dd/MM/yyyy");
    const dateParagraph = body.appendParagraph(`Date: ${formattedDate}`);
    dateParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Get nearest Thursday for submission date
    const nextThursday = getNearestThursday();
    const thursdayDate = Utilities.formatDate(nextThursday, "GMT+5:30", "dd/MM/yyyy");
    const submissionParagraph = body.appendParagraph(`Date of Submission: ${thursdayDate}`);
    submissionParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Add basic information
    const infoParagraph1 = body.appendParagraph(`Name of the Employee: ${requestData.requestorName || ""}`);
    infoParagraph1.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Format designation as "Ministry, Designation"
    const ministryName = requestData.ministry || "";
    const designationText = requestData.designation || "";
    const formattedDesignation = ministryName + (designationText ? `, ${designationText}` : "");
    const infoParagraph2 = body.appendParagraph(`Designation: ${formattedDesignation}`);
    infoParagraph2.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    const infoParagraph3 = body.appendParagraph(`Employee Code: ${requestData.registrationNumber || ""}`);
    infoParagraph3.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Add space
    body.appendParagraph("\n");

    // Create a table for expenses
    const table = body.appendTable();

    // Add header row
    const headerRow = table.appendTableRow();
    const headers = ["S.No.", "Date", "Bill Number", "Description", "Amount Rs.", "Receipt Attached?", "Remarks"];
    headers.forEach(header => {
      const cell = headerRow.appendTableCell(header);
      cell.getChild(0).asParagraph().setBold(true);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cell.setBackgroundColor("#D9D9D9");
    });

    // Process expenses
    const expenses = requestData.expenses || [];
    let totalAmount = 0;

    // Sort expenses by date if available
    if (expenses.length > 0 && expenses[0].paymentDate) {
      expenses.sort((a, b) => {
        try {
          const dateA = new Date(a.paymentDate);
          const dateB = new Date(b.paymentDate);
          return dateA - dateB;
        } catch (e) {
          return 0;
        }
      });
    }

    // Add expense rows
    expenses.forEach((expense, index) => {
      const row = table.appendTableRow();

      // S.No.
      row.appendTableCell((index + 1).toString())
         .getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      // Date
      let dateStr = "";
      if (expense.paymentDate) {
        try {
          const expDate = new Date(expense.paymentDate);
          dateStr = Utilities.formatDate(expDate, "GMT+5:30", "dd/MM/yyyy");
        } catch (e) {
          dateStr = expense.paymentDate;
        }
      }
      row.appendTableCell(dateStr);
      
      // Bill Number (expense number)
      row.appendTableCell(expense.billNumber || "");

      // Description (expense category)
      row.appendTableCell(expense.category || "");

      // Amount
      const amount = parseFloat(expense.amount) || 0;
      totalAmount += amount;
      row.appendTableCell(amount.toFixed(2))
         .getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

      // Receipt Attached?
      const hasReceipt = expense.receiptLink ? "Yes" : "No";
      row.appendTableCell(hasReceipt)
         .getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      // Remarks (justification)
      row.appendTableCell(expense.justification || "");
    });

    // If no expenses, add a blank row
    if (expenses.length === 0) {
      const row = table.appendTableRow();
      row.appendTableCell("").appendTableCell("").appendTableCell("")
         .appendTableCell("").appendTableCell("").appendTableCell("")
         .appendTableCell("");
    }

    // Add space
    body.appendParagraph("\n");

    // Add total
    const totalParagraph = body.appendParagraph(`Total Amount: Rs. ${totalAmount.toFixed(2)}`);
    totalParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    totalParagraph.setBold(true);

    // Add total in words
    const amountInWords = convertNumberToWords(totalAmount);
    const wordsLine = body.appendParagraph(`Total Amount in words: Rupees ${amountInWords}`);
    wordsLine.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Add space
    body.appendParagraph("\n\n");

    // Add signature area with the three required fields
    const sigParagraph = body.appendParagraph("Employee Signature: ________________________");
    sigParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      
    const accountsParagraph = body.appendParagraph("Checked by Accounts: ________________________");
    accountsParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    const approvalParagraph = body.appendParagraph("Approved by (Authorised Signatory): ________________________");
    approvalParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    // Save and close the document to finalize it
    doc.saveAndClose();

    // Get the file to move it to the request folder
    const file = DriveApp.getFileById(doc.getId());

    // Make a copy in the request folder with appropriate name
    const newFile = file.makeCopy(`Approval Document - ${requestId}`, requestFolder);

    // Delete the original from root
    file.setTrashed(true);

    // Set sharing permissions on the new document
    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      Logger.log(`Warning: Could not set document sharing permissions: ${sharingError}`);
      // Continue even if setting permissions fails
    }

    Logger.log(`Generated approval document from scratch: ${newFile.getUrl()}`);
    return { id: newFile.getId(), url: newFile.getUrl() };
  } catch (error) {
    Logger.log(`Error creating approval document for request ID ${requestId}: ${error}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

// Helper function to convert numbers to words for the amount
function convertNumberToWords(amount) {
  // Format the amount with 2 decimal places
  const formattedAmount = amount.toFixed(2);

  // Split the amount into rupees and paise
  const parts = formattedAmount.split('.');
  const rupees = parseInt(parts[0]);
  const paise = parseInt(parts[1]);

  // Define arrays for number words
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 
                'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  // Function to convert a number less than 1000 to words
  function convertLessThanOneThousand(num) {
    if (num === 0) {
      return '';
    }

    let result = '';

    if (num < 20) {
      result = ones[num];
    } else if (num < 100) {
      result = tens[Math.floor(num / 10)] + (num % 10 !== 0 ? ' ' + ones[num % 10] : '');
    } else {
      result = ones[Math.floor(num / 100)] + ' Hundred' + (num % 100 !== 0 ? ' and ' + convertLessThanOneThousand(num % 100) : '');
    }

    return result;
  }

  // Convert the rupees part
  let result = '';

  if (rupees === 0) {
    result = 'Zero';
  } else {
    // Handle crores
    const crore = Math.floor(rupees / 10000000);
    if (crore > 0) {
      result += convertLessThanOneThousand(crore) + ' Crore ';
      rupees %= 10000000;
    }

    // Handle lakhs
    const lakh = Math.floor(rupees / 100000);
    if (lakh > 0) {
      result += convertLessThanOneThousand(lakh) + ' Lakh ';
      rupees %= 100000;
    }

    // Handle thousands
    const thousand = Math.floor(rupees / 1000);
    if (thousand > 0) {
      result += convertLessThanOneThousand(thousand) + ' Thousand ';
      rupees %= 1000;
    }

    // Handle remaining part
    if (rupees > 0) {
      result += convertLessThanOneThousand(rupees);
    }
  }
  
  // Add the paise part if it exists
  if (paise > 0) {
    result += ' and ' + convertLessThanOneThousand(paise) + ' Paise';
  }
  
  return result + ' Only';
}

/**
 * Get nearest Thursday for the submission date
 * Returns the upcoming Thursday, or if today is Thursday, returns today
 * @return {Date} Date object for the nearest Thursday
 */
function getNearestThursday() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 is Sunday, 4 is Thursday
  const daysUntilThursday = dayOfWeek <= 4 ? 4 - dayOfWeek : 11 - dayOfWeek;  // Calculate days until Thursday (or 0 if today is Thursday)
  const thursday = new Date(today);  // Create a new date by adding the days until Thursday
  thursday.setDate(today.getDate() + daysUntilThursday);
  return thursday;
}
