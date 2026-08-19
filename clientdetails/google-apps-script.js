/**
 * Google Apps Script — Leapuzo Client Onboarding Form
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet (create a new one or use existing).
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code and paste this entire script.
 * 4. Replace YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE below with your Drive folder ID.
 *    (Folder ID is the long string in the Drive folder URL after /folders/)
 * 5. Click Save (disk icon).
 * 6. Click Deploy > New Deployment.
 * 7. Select type: Web App.
 *    - Execute As: Me (your email)
 *    - Who has access: Anyone
 * 8. Click Deploy, authorize permissions, copy the Web App URL.
 * 9. Paste that URL in clientdetails/index.html wherever scriptURL is defined.
 */

// ─── CONFIG ────────────────────────────────────────────────────────────────
const FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";
// ────────────────────────────────────────────────────────────────────────────

// All form headings in exact display order
const HEADERS = [
  "Timestamp",

  // ── Company Details ──────────────────────────────────────────────────────
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

  // ── Primary Contact ──────────────────────────────────────────────────────
  "Primary Contact Person Name",
  "Primary Contact Person Designation",
  "Primary Contact Person Number",
  "Primary Contact Person Email ID",

  // ── Secondary Contact ────────────────────────────────────────────────────
  "Secondary Contact Person Name",
  "Secondary Contact Person Designation",
  "Secondary Contact Person Number",
  "Secondary Contact Person Email ID",

  // ── Business Overview ────────────────────────────────────────────────────
  "Briefly describe your business.",
  "What products or services do you offer? ",
  "What are your primary business goals?",
  "Long-term goals (3-5 years)",
  "Who is your ideal customer?",
  "What influences their purchase?",
  "Why should they choose you?",
  "What are your USP's?",
  "Top 5 competitors",

  // ── Social Media Credentials ─────────────────────────────────────────────
  "Instagram Username / Email",
  "instagramPassword",
  "Facebook Username / Email",
  "facebookPassword",
  "YouTube Google Account Email",
  "youtubePassword",
  "LinkedIn Email",
  "X Username / Email",
  "xPassword",

  // ── Domain ───────────────────────────────────────────────────────────────
  "Domain Provider",
  "Domain Login Email / Username",
  "domainPassword",

  // ── Agency Expectations ──────────────────────────────────────────────────
  "What has frustrated you with previous agencies?",
  "What do you definitely NOT want?",
  "What type of design do you dislike?",
  "What expectations do you have from us?",

  // ── Uploaded Files ───────────────────────────────────────────────────────
  "Company Logo Link",
  "Brand Guidelines Link"
];

// ─── HELPER: Get or create Drive folder ─────────────────────────────────────
function getFolder() {
  try {
    if (FOLDER_ID && FOLDER_ID !== "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE") {
      return DriveApp.getFolderById(FOLDER_ID);
    }
  } catch (e) {}
  return DriveApp.getRootFolder();
}

// ─── HELPER: Upload base64 file to Drive, return public URL ─────────────────
function uploadFile(base64, fileName) {
  if (!base64 || !fileName) return "";
  const ext = fileName.toLowerCase().split(".").pop();
  const mimeMap = {
    pdf: "application/pdf",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    zip: "application/zip",
    txt: "text/plain"
  };
  const contentType = mimeMap[ext] || "application/octet-stream";
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = getFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ─── MAIN POST HANDLER ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header row if sheet is blank
    if (sheet.getLastColumn() === 0 || sheet.getRange(1, 1).getValue() === "") {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

      // Style the header row: bold, frozen, purple background, white text
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#971dcc");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, HEADERS.length, 200);
    }

    // Upload files and get Drive URLs
    const logoUrl  = uploadFile(data.logoBase64,  data.logoFileName);
    const brandUrl = uploadFile(data.brandBase64, data.brandFileName);

    // Map data to header order
    const row = HEADERS.map(header => {
      if (header === "Timestamp")          return new Date();
      if (header === "Company Logo Link")  return logoUrl;
      if (header === "Brand Guidelines Link") return brandUrl;

      // Handle HTML entity in field name (& was stored as &amp; in HTML)
      const decoded = header.replace(/&amp;/g, "&");
      return data[header] || data[decoded] || "";
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", logoUrl }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── TEST FUNCTION (run manually from Apps Script editor to verify) ──────────
function testSetup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Active spreadsheet: " + SpreadsheetApp.getActiveSpreadsheet().getName());
  Logger.log("Drive folder accessible: " + getFolder().getName());
  Logger.log("Headers count: " + HEADERS.length);
  Logger.log("Setup looks correct.");
}
