/**
 * Runs when the spreadsheet is opened to add our "Config" menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Config')
    .addItem('Set Directory Spreadsheet…', 'showDirectoryDialog')
    .addToUi();
}

/**
 * Prompts the user to paste a Sheets URL or ID,
 * extracts the ID, and saves it as a script property.
 */
function showDirectoryDialog() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Configure Directory Spreadsheet',
    'Paste the full Google Sheets URL or just the Spreadsheet ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  const id    = extractSpreadsheetId(input);
  
  if (!id) {
    ui.alert('❌ Invalid URL or ID. Please try again.');
    return;
  }
  
  PropertiesService
    .getScriptProperties()
    .setProperty('DIRECTORY_SPREADSHEET_ID', id);
  
  ui.alert('✅ DIRECTORY_SPREADSHEET_ID set to:\n' + id);
}

/**
 * Helpers: pulls an ID out of either
 *   • a /d/URL segment, or
 *   • a bare ID string
 */
function extractSpreadsheetId(input) {
  const urlMatch = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }
  // basic sanity check for a bare ID
  if (/^[a-zA-Z0-9-_]+$/.test(input)) {
    return input;
  }
  return null;
}

/**
 * Fetches raw data from the required Google Sheets.
 * Reads "Service Attendance", "Event Attendance", and "Attendance Stats"
 * from the active spreadsheet, and "Directory" from an external spreadsheet.
 * Uses getDataRange() to fetch all data with content.
 * Includes error handling and logging for debugging.
 *
 * @returns {object} An object containing the data arrays:
 *   { sData, eData, dData, statsData },
 * or undefined on failure.
 */
function getDataFromSheets() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  // --- get the Directory ID from script props (fail fast if missing) ---
  const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
  if (!directoryId) {
    throw new Error(
      "🚨 Missing Script Property: 'DIRECTORY_SPREADSHEET_ID'.\n" +
      "Use Config → Set Directory Spreadsheet… to configure it."
    );
  }
  
  // --- Get local sheets ---
  const serviceSheet = ss.getSheetByName("Service Attendance");
  const eventSheet   = ss.getSheetByName("Event Attendance");
  const statsSheet   = ss.getSheetByName("Attendance Stats");
  
  if (!serviceSheet || !eventSheet || !statsSheet) {
    if (!serviceSheet) Logger.log("❌ 'Service Attendance' not found.");
    if (!eventSheet)   Logger.log("❌ 'Event Attendance' not found.");
    if (!statsSheet)   Logger.log("❌ 'Attendance Stats' not found.");
    return;
  }
  
  // --- Load data from local sheets ---
  let sData, eData, statsData;
  try {
    sData     = serviceSheet.getDataRange().getValues();
    eData     = eventSheet.getDataRange().getValues();
    statsData = statsSheet.getDataRange().getValues();
    Logger.log("✅ Local sheets loaded.");
  } catch (err) {
    Logger.log("❌ Error reading local sheets: " + err.message);
    return;
  }
  
  // --- Load Directory from external spreadsheet by ID ---
  let directorySS;
  try {
    directorySS = SpreadsheetApp.openById(directoryId);
    Logger.log("✅ External spreadsheet opened via ID.");
  } catch (err) {
    Logger.log("❌ Could not open external spreadsheet ID=" + directoryId + " : " + err.message);
    return;
  }
  
  const directorySheet = directorySS.getSheetByName("Directory");
  if (!directorySheet) {
    Logger.log("❌ 'Directory' sheet not found in external spreadsheet.");
    return;
  }
  
  let dData;
  try {
    dData = directorySheet.getDataRange().getValues();
    Logger.log("✅ Directory sheet loaded.");
  } catch (err) {
    Logger.log("❌ Error reading Directory sheet: " + err.message);
    return;
  }
  
  Logger.log("✅ All required sheets loaded successfully.");
  return { sData, eData, dData, statsData };
}
