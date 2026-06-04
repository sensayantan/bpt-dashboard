function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('BPT Performance Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Data for Section 1
function getSection1Data() {
  try {
    const ss = SpreadsheetApp.openById("1OY94D17F43-q__X-wdISQxYozUX4YctMg4vTRUxh5GU");
    const sheet = ss.getSheetByName("Project P/L (FY25-Q4)"); 
    const values = sheet.getDataRange().getValues();
    return JSON.stringify(values);
  } catch (e) {
    return JSON.stringify([]); 
  }
}

/**
 * Fetches data from the "BPT - Progress Chaser" worksheet.
 * Sourced from row 2 as headers.
 */
function getProgressData() {
  try {
    const ss = SpreadsheetApp.openById("1RhWKg-gUfbYxX0m4YmuESsq3Klpj5X5ymHSejPCwieI");
    const sheet = ss.getSheetByName("BPT - Progress Chaser");
    // Start from row 2 per your requirement
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return JSON.stringify(values);
  } catch (e) {
    return JSON.stringify({error: e.toString()});
  }
}
