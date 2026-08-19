function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    // 1. Open the active spreadsheet and sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 2. Define the headings in the exact order of the form fields
    var headers = [
      "Timestamp",
      "Employee ID",
      "Employee Name",
      "Department",
      "Designation",
      "Date of Joining",
      "Gender",
      "Date of Birth",
      "Blood Group",
      "Employment Type",
      "Work Location",
      "Primary Contact No",
      "Alternate Contact Number",
      "Email ID",
      "Permanent Address",
      "Current Address",
      "Emergency Contact Name",
      "Emergency Contact Phone Number",
      "Emergency Contact Relation",
      "Emergency Contact 2 Name",
      "Emergency Contact 2 Phone Number",
      "Emergency Contact 2 Relation",
      "Aadhaar Card Link",
      "PAN Card Link",
      "Driving License Link",
      "Bank Name",
      "Bank Branch",
      "Account Number",
      "IFSC Code",
      "UPI ID",
      "Cancelled Cheque Link"
    ];
    
    // 3. Check if header row exists, write headers if sheet is empty or headers are not present
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.appendRow(headers);
      // Format headers: Bold, background color, frozen row
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#971dcc"); // Matching theme purple
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }
    
    // 4. Create/get folder in Google Drive for uploads
    var folderName = "Employee Details Uploads";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    // Helper function to decode and save base64 files
    function saveFile(base64Data, fileName) {
      if (!base64Data || base64Data === "") return "";
      try {
        var splitData = base64Data.split(',');
        var contentType = splitData[0].match(/:(.*?);/)[1];
        var rawBase64 = splitData[1];
        var decoded = Utilities.base64Decode(rawBase64);
        var blob = Utilities.newBlob(decoded, contentType, fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      } catch (err) {
        return "Upload Error: " + err.toString();
      }
    }
    
    // 5. Upload files and get URLs
    var employeeName = data["Employee Name"] || "Employee";
    
    var aadhaarLink = saveFile(data.AadhaarFile, employeeName + "_Aadhaar_" + (data.AadhaarFileName || "file"));
    var panLink = saveFile(data.PANFile, employeeName + "_PAN_" + (data.PANFileName || "file"));
    var dlLink = saveFile(data.DrivingLicenseFile, employeeName + "_DrivingLicense_" + (data.DrivingLicenseFileName || "file"));
    var chequeLink = saveFile(data.CancelledChequeFile, employeeName + "_CancelledCheque_" + (data.CancelledChequeFileName || "file"));
    
    // 6. Map JSON keys to the headers order (accounting for exact names/spaces from index.html payload)
    var rowValues = [];
    
    // Timestamp
    rowValues.push(new Date());
    
    // Form fields mapped in exact order
    rowValues.push(data["Employee ID"] || "");
    rowValues.push(data["Employee Name"] || "");
    rowValues.push(data["Department"] || "");
    rowValues.push(data["Designation"] || "");
    rowValues.push(data["Date of Joining"] || "");
    rowValues.push(data["Gender"] || "");
    rowValues.push(data["Date of Birth"] || "");
    rowValues.push(data["Blood Group"] || "");
    rowValues.push(data["Employement Type"] || data["Employment Type"] || "");
    rowValues.push(data["Work Location"] || "");
    
    // Note: The form name attribute contains trailing spaces: "Primary Contact No  "
    rowValues.push(data["Primary Contact No  "] || data["Primary Contact No"] || "");
    rowValues.push(data["Alternate Contact Number"] || "");
    rowValues.push(data["Email ID"] || "");
    rowValues.push(data["Permanent Address"] || "");
    rowValues.push(data["Current Address"] || "");
    
    // Emergency Contact 1
    rowValues.push(data["Emergency contact Name"] || "");
    rowValues.push(data["Emergency contact phone number"] || "");
    rowValues.push(data["Emergency contact Relation"] || "");
    
    // Emergency Contact 2
    rowValues.push(data["Emergency Contact 2 name"] || "");
    rowValues.push(data["Emergency contact 2 phone number"] || "");
    rowValues.push(data["Emergency Contact 2 Relation"] || "");
    
    // Files links
    rowValues.push(aadhaarLink);
    rowValues.push(panLink);
    rowValues.push(dlLink);
    
    // Bank Details
    rowValues.push(data["Bank Name"] || "");
    rowValues.push(data["Bank Branch"] || "");
    rowValues.push(data["Account Number"] || "");
    rowValues.push(data["IFSC Code"] || "");
    rowValues.push(data["UPI ID"] || "");
    
    // Cancelled Cheque Link
    rowValues.push(chequeLink);
    
    // 7. Append row to spreadsheet
    sheet.appendRow(rowValues);
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
