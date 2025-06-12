/**
 * Calculates activity levels based on the sum of values from Column E and Column K.
 * Writes the activity level ("Core", "Active", "Inactive", or blank) to Column L.
 * Non-numeric inputs are treated as 0 for the sum.
 * If both inputs are blank/non-numeric text, the output activity level is blank.
 */
function updateActivityLevels() {
  // Configuration:
  const sheetName = "attendance stats";         // Name of the sheet
  const sourceColumnLetterE = "E";              // Column for the first value
  const sourceColumnLetterK = "K";              // Column for the second value
  const targetColumnForActivity = "L";        // Column to write the activity level (12th column)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found.`);
    // SpreadsheetApp.getUi().alert(`Error: Sheet '${sheetName}' not found.`); // Commented out
    return;
  }

  const firstDataRow = 2; // Assuming data starts from row 2 (row 1 has headers)
  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) {
    Logger.log("No data to process in the sheet.");
    // SpreadsheetApp.getUi().alert("No data to process."); // Commented out
    return;
  }

  // Read data from source columns
  const valuesE = sheet.getRange(`${sourceColumnLetterE}${firstDataRow}:${sourceColumnLetterE}${lastRow}`).getValues();
  const valuesK = sheet.getRange(`${sourceColumnLetterK}${firstDataRow}:${sourceColumnLetterK}${lastRow}`).getValues();

  const activityLevelsOutput = []; // Array to hold the calculated activity levels

  // Iterate through each row
  for (let i = 0; i < valuesE.length; i++) {
    const rawValueE = valuesE[i][0];
    const rawValueK = valuesK[i][0];

    const numE = Number(rawValueE);
    const numK = Number(rawValueK);

    const valEForSum = isNaN(numE) ? 0 : numE;
    const valKForSum = isNaN(numK) ? 0 : numK;

    const combinedCount = valEForSum + valKForSum;
    let currentActivityLevel;

    const eIsEffectivelyBlankOrInvalid = (rawValueE === "" || (rawValueE.toString().trim() !== "" && isNaN(numE)));
    const kIsEffectivelyBlankOrInvalid = (rawValueK === "" || (rawValueK.toString().trim() !== "" && isNaN(numK)));

    if (eIsEffectivelyBlankOrInvalid && kIsEffectivelyBlankOrInvalid) {
      currentActivityLevel = "";
      Logger.log(`Row ${i + firstDataRow}: E='${rawValueE}', K='${rawValueK}' -> Both effectively blank/invalid. Activity: blank.`);
    } else if (combinedCount >= 12) {
      currentActivityLevel = "Core";
      Logger.log(`Row ${i + firstDataRow}: E='${rawValueE}', K='${rawValueK}' -> Sum=${combinedCount} -> Activity: Core`);
    } else if (combinedCount >= 3) {
      currentActivityLevel = "Active";
      Logger.log(`Row ${i + firstDataRow}: E='${rawValueE}', K='${rawValueK}' -> Sum=${combinedCount} -> Activity: Active`);
    } else {
      currentActivityLevel = "Inactive";
      Logger.log(`Row ${i + firstDataRow}: E='${rawValueE}', K='${rawValueK}' -> Sum=${combinedCount} -> Activity: Inactive`);
    }
    activityLevelsOutput.push([currentActivityLevel]);
  }

  if (activityLevelsOutput.length > 0) {
    const targetColumnIndex = sheet.getRange(`${targetColumnForActivity}1`).getColumn();
    sheet.getRange(firstDataRow, targetColumnIndex, activityLevelsOutput.length, 1).setValues(activityLevelsOutput);
    Logger.log(`Activity levels based on sum of E and K written to Column ${targetColumnForActivity}. Processed ${activityLevelsOutput.length} rows.`);
    // SpreadsheetApp.getUi().alert(`Activity levels have been updated in Column ${targetColumnForActivity}.`); // Commented out
  } else {
    Logger.log("No activity levels were calculated to write.");
  }
}
