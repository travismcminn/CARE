/**
 * C.A.R.E. public fundraising status feed.
 * Deploy this as a Google Apps Script web app from the CARE Donor Tracker.
 *
 * Deployment: Execute as Me; Who has access: Anyone.
 * This returns ONLY the public donation total and timestamp.
 * Never return donor rows, names, email addresses, or payment IDs.
 */
function doGet() {
  const spreadsheet = SpreadsheetApp.openById('1Fy3SguQBWcTeuCFdvtbjey2V89slLmsee4vNXkQMLKg');
  const sheet = spreadsheet.getSheetByName('Dashboard');
  if (!sheet) throw new Error('Dashboard tab not found');
  const amount = sheet.getRange('B3').getValue(); // Net Donations Received
  const raised = typeof amount === 'number'
    ? amount
    : Number(String(amount).replace(/[$,]/g, ''));
  if (!Number.isFinite(raised) || raised < 0) throw new Error('Invalid fundraising total');

  const payload = {
    raised: Math.round(raised * 100) / 100,
    updated: new Date().toISOString()
  };

  return ContentService
    .createTextOutput('window.CARE_fundraisingUpdate(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
