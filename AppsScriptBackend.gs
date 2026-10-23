/**
 * Google Apps Script: read-only live feed for the English SPM Dashboard.
 * Bound to "English Performance Dashboard - SPM Cohorts".
 *
 * DEPLOYMENT (one-time):
 * Extensions > Apps Script > replace Code.gs with this file.
 * Deploy > New deployment > Web app.
 * Execute as: Me; Who has access: Anyone.
 * Copy the /exec URL into the "Connect Sheet" panel of the GitHub dashboard.
 *
 * IMPORTANT: The spreadsheet itself can remain Restricted. However,
 * this publicly deployed endpoint exposes the returned students' names,
 * assessment results and item scores to anyone who opens the feed URL.
 * Do not add write operations or credentials to this script.
 */
const DASHBOARD_SPREADSHEET_ID = '1HHL92HTiXhSvXAE-1g3mRIVUdyM16agTJw1lPf7pEf0';

function getDashboardTab_(spreadsheet, name) {
  const tab = spreadsheet.getSheetByName(name);
  if (!tab) throw new Error('Required tab is missing: ' + name);
  const lastRow = tab.getLastRow();
  const lastColumn = tab.getLastColumn();
  if (!lastRow || !lastColumn) return [];
  return tab.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
}

function toDashboardObjects_(matrix) {
  if (!matrix.length) return [];
  const headers = matrix[0].map(String);
  return matrix.slice(1).filter(row => row.some((v, index) =>
    headers[index] && String(v).trim() !== ''
  )).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = row[index] || '';
    });
    return obj;
  });
}

function buildDashboardResponse_() {
  const ss = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  const results = toDashboardObjects_(getDashboardTab_(ss, 'RESULTS_DATA'))
    .filter(row => row.Student_Name);
  const targets = toDashboardObjects_(getDashboardTab_(ss, 'TARGETS'));
  const interventions = toDashboardObjects_(getDashboardTab_(ss, 'INTERVENTION'));
  const items = toDashboardObjects_(getDashboardTab_(ss, 'ITEM_ANALYSIS_DATA'))
    .filter(row => row.Student_Name);
  const tracker = getDashboardTab_(ss, 'HEADCOUNT_TRACKER');
  const settings = getDashboardTab_(ss, 'SETTINGS');
  const passRow = settings.find(row => String(row[0]).trim() === 'Official Pass Mark (0–100)');
  const passMark = passRow ? Number(passRow[1]) : NaN;
  if (!Number.isFinite(passMark) || passMark < 0 || passMark > 100) {
    throw new Error('Pass mark not found or invalid in SETTINGS.');
  }
  return {
    ok: true,
    results: results,
    targets: targets,
    interventions: interventions,
    items: items,
    tracker: tracker,
    passMark: passMark,
    updated: new Date().toISOString()
  };
}

function doGet(e) {
  const prefix = String((e && e.parameter && e.parameter.prefix) || '');
  // Only a simple function identifier is accepted; never evaluate input.
  if (prefix && !/^[A-Za-z_$][A-Za-z0-9_$]{0,75}$/.test(prefix)) {
    return ContentService.createTextOutput('Invalid callback name')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  let payload;
  try {
    payload = buildDashboardResponse_();
  } catch (error) {
    payload = { ok: false, error: String(error && error.message || error) };
  }
  const result = JSON.stringify(payload);
  return ContentService.createTextOutput(prefix ? prefix + '(' + result + ');' : result)
    .setMimeType(prefix ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
