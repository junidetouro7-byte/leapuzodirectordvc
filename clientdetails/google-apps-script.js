/**
 * Google Apps Script to handle Onboarding Form submissions.
 * 
 * Instructions:
 * 1. Open Google Sheets (create a new one or use your existing sheet).
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code and paste this script.
 * 4. Replace the FOLDER_ID placeholder below with the ID of the Google Drive folder where you want files to be saved.
 * 5. Click Save (disk icon).
 * 6. Click Deploy > New Deployment.
 * 7. Choose type: Web App.
 *    - Execute As: Me (your-email)
 *    - Who has access: Anyone
 * 8. Click Deploy, authorize permissions, and copy the generated "Web app URL".
 * 9. Paste the Web app URL in your onboarding form's Javascript script URL variable.
 */

// Replace with the Google Drive folder ID where uploaded logos should be saved.
const FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 1. Create or Find Headers
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() > 0 ? sheet.getLastColumn() : 1).getValues()[0];
    if (headers[0] === "") {
      // Setup default headers if sheet is empty
      headers = [
        "Timestamp",
        "Legal Name of the Company",
        "Business Sector",
        "Registered Business Address & Pincode",
        "Corporate Business Address & Pincode",
        "Constitution",
        "GST No:",
        "MSME (If applicable)",
        "Company Pan Number",
        "Official Contact Number",
        "Official Email Address",
        "Website URL(if available)",
        "Primary Contact Person Name",
        "Primary Contact Person Designation",
        "Primary Contact Person Number",
        "Primary Contact Person Email ID",
        "Secondary Contact Person Name",
        "Secondary Contact Person Designation",
        "Secondary Contact Person Number",
        "Secondary Contact Person Email ID",
        "Briefly describe your business.",
        "What products or services do you offer? ",
        "What are your primary business goals?",
        "Company Logo Link",
        "Brand Guidelines"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // 2a. Handle Logo Upload to Google Drive if present
    let logoUrl = "";
    if (data.logoBase64 && data.logoFileName) {
      let folder;
      try {
        if (FOLDER_ID && FOLDER_ID !== "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE" && FOLDER_ID !== "") {
          folder = DriveApp.getFolderById(FOLDER_ID);
        } else {
          folder = DriveApp.getRootFolder();
        }
      } catch (err) {
        folder = DriveApp.getRootFolder();
      }

      // Determine Content Type based on extension
      let contentType = "image/png"; // default
      const ext = data.logoFileName.toLowerCase().split('.').pop();
      if (ext === "pdf") {
        contentType = "application/pdf";
      } else if (ext === "jpg" || ext === "jpeg") {
        contentType = "image/jpeg";
      }

      const bytes = Utilities.base64Decode(data.logoBase64);
      const blob = Utilities.newBlob(bytes, contentType, data.logoFileName);
      const file = folder.createFile(blob);
      
      // Make the file publicly viewable so it can be opened from the sheet link
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      logoUrl = file.getUrl();
    }

    // 2b. Handle Brand Guidelines Upload to Google Drive if present
    let brandUrl = "";
    if (data.brandBase64 && data.brandFileName) {
      let folder;
      try {
        if (FOLDER_ID && FOLDER_ID !== "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE" && FOLDER_ID !== "") {
          folder = DriveApp.getFolderById(FOLDER_ID);
        } else {
          folder = DriveApp.getRootFolder();
        }
      } catch (err) {
        folder = DriveApp.getRootFolder();
      }

      // Determine Content Type based on extension
      let contentType = "application/octet-stream"; // default
      const ext = data.brandFileName.toLowerCase().split('.').pop();
      if (ext === "pdf") {
        contentType = "application/pdf";
      } else if (ext === "zip") {
        contentType = "application/zip";
      } else if (ext === "png") {
        contentType = "image/png";
      } else if (ext === "jpg" || ext === "jpeg") {
        contentType = "image/jpeg";
      } else if (ext === "txt") {
        contentType = "text/plain";
      }

      const bytes = Utilities.base64Decode(data.brandBase64);
      const blob = Utilities.newBlob(bytes, contentType, data.brandFileName);
      const file = folder.createFile(blob);
      
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      brandUrl = file.getUrl();
    }
    
    // 3. Map submission fields to headers
    const rowValues = [];
    const timestamp = new Date();
    
    headers.forEach(header => {
      if (header === "Timestamp") {
        rowValues.push(timestamp);
      } else if (header === "Company Logo Link" || header === "Upload your primary company logo. (High Quality)") {
        rowValues.push(logoUrl);
      } else if (header === "Brand Guidelines") {
        rowValues.push(brandUrl || data[header] || "");
      } else {
        // Match the data fields sent from form
        const val = data[header] || "";
        rowValues.push(val);
      }
    });
    
    // 4. Append to spreadsheet
    sheet.appendRow(rowValues);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      logoUrl: logoUrl
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
