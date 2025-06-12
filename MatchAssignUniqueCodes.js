/**
 * Extracts a purely numeric ID if the input is already a number or a string representing a non-negative integer.
 * It NO LONGER parses "BEL" prefixes. Any string containing non-numeric characters (like "BEL123")
 * will be treated as an invalid/missing ID.
 * @param {any} codeValue The value from the sheet that might be a numeric ID.
 * @returns {number|null} The extracted number, or null if not a valid purely numeric ID.
 */
function extractNumericBel(codeValue) {
  if (typeof codeValue === 'number' && Number.isInteger(codeValue) && codeValue >= 0) {
    return codeValue; // It's already a valid numeric ID
  }
  if (typeof codeValue === 'string') {
    const numStr = codeValue.trim();

    if (numStr === '') {
      return null; // Empty string is not a valid ID
    }

    // Ensure it's purely digits before parsing.
    // This will cause "BEL123", "BEL-123", etc., to be rejected and return null.
    if (!/^\d+$/.test(numStr)) {
      // You can comment out this log if it becomes too verbose due to many non-numeric entries
      Logger.log(`ℹ️ extractNumericBel: Input "${codeValue}" is not a purely numeric string. Will treat as missing/invalid ID.`);
      return null;
    }

    const num = parseInt(numStr, 10);
    // Ensure the parsed number is not NaN and is non-negative.
    if (!isNaN(num) && num >= 0) {
      return num;
    } else {
      Logger.log(`⚠️ extractNumericBel: Could not parse valid non-negative number from numeric string "${numStr}". Parsed: ${num}`);
      return null;
    }
  }
  // If codeValue is not a string or a number, or if it's a string that didn't meet criteria.
  return null;
}

/**
 * Fetches data using getDataFromSheets, matches/assigns purely NUMERIC IDs,
 * and formats attendance records into a standardized 11-column structure.
 *
 * This function first collects all existing PURELY NUMERIC IDs from the Directory,
 * Event Attendance (Column A), and Service Attendance (Column A) sheets using the modified extractNumericBel.
 * It builds a comprehensive name-to-NUMERIC_ID mapping (belMap). It then finds
 * the highest number among these numeric codes and generates new unique NUMERIC IDs
 * sequentially from the next number when a person doesn't have an existing valid numeric code.
 *
 * Assumes data object from getDataFromSheets contains:
 * - sData: Array of rows from Service Attendance
 * - eData: Array of rows from Event Attendance
 * - dData: Array of rows from Directory (assumed ID in row[1], Name in row[2])
 *
 * The function returns an array of arrays, where each inner array has
 * at least 11 elements, conforming to the "Event Attendance" column structure:
 * [0: Numeric ID (number), 1: Full Name, ..., 10: Timestamp]
 *
 * @returns {Array<Array<any>>} An array of formatted attendance records, or empty array if data loading fails.
 */
function matchOrAssignBelCodes() { // Consider renaming to matchOrAssignNumericIds for clarity
  // Get the data from the sheets using the helper function
  const data = getDataFromSheets();
  if (!data) {
    Logger.log("❌ No data loaded from sheets in matchOrAssignBelCodes. Exiting.");
    return [];
  }

  const { sData, eData, dData } = data;

  const belMap = new Map(); // Map to store normalized name -> NUMERIC ID mappings
  const allUsedCodes = new Set(); // Set to keep track of all existing NUMERIC IDs found
  const normalize = name => name?.toString().trim().toLowerCase();

  // --- Step 1: Populate belMap and allUsedCodes (with NUMBERS) from the Directory sheet ---
  if (dData && dData.length > 1) {
    dData.slice(1).forEach((row, index) => {
      if (row.length > 2) {
        const originalBel = row[1]; // Original ID from Column B
        const name = normalize(row[2]); // Name from Column C
        const numericBel = extractNumericBel(originalBel); // Uses the MODIFIED function

        if (name) { // Name is mandatory
          if (numericBel !== null) { // A valid PURELY NUMERIC ID was extracted
            if (!belMap.has(name)) {
              belMap.set(name, numericBel);
            }
            allUsedCodes.add(numericBel);
          } else if (originalBel && originalBel.toString().trim() !== '') {
            // Log if there was some value in the ID column that wasn't a valid number
            Logger.log(`ℹ️ Directory: Row ${index + 2}: Value "${originalBel}" in ID column is not a plain number and will be ignored. A new numeric ID may be generated for "${name}" if needed.`);
          }
        } else {
          Logger.log(`⚠️ Directory: Row ${index + 2}: Skipping row due to missing Name. Row data: ${JSON.stringify(row)}`);
        }
      } else {
        Logger.log(`⚠️ Directory: Row ${index + 2}: Skipping row due to insufficient columns (${row.length} found). Expected at least 3. Row data (partial): ${JSON.stringify(row.slice(0,3))}`);
      }
    });
    Logger.log(`✅ Populated numeric ID map and used codes from Directory (${dData.length > 1 ? dData.length - 1 : 0} data rows processed).`);
  } else {
    Logger.log("⚠️ matchOrAssignBelCodes: Directory data (dData) is empty or missing headers.");
  }

  // --- Step 2: Add existing NUMERIC IDs from Attendance sheets and update belMap ---
  const attendanceDataRaw = [];
  if (eData && eData.length > 1) attendanceDataRaw.push(...eData.slice(1));
  if (sData && sData.length > 1) attendanceDataRaw.push(...sData.slice(1));

  attendanceDataRaw.forEach((row, idx) => {
    if (row.length > 1) {
      const originalBelFromRow = row[0]; // Original value from Column A
      const name = normalize(row[1]); // Name from Column B
      const numericBelFromRow = extractNumericBel(originalBelFromRow); // Uses the MODIFIED function

      if (numericBelFromRow !== null) {
        allUsedCodes.add(numericBelFromRow);
        if (name && !belMap.has(name)) {
          belMap.set(name, numericBelFromRow);
        }
      } else if (originalBelFromRow && originalBelFromRow.toString().trim() !== '' && name) {
         Logger.log(`ℹ️ Attendance Sheets (Source Row approx. ${idx + 1}): Value "${originalBelFromRow}" for name "${name}" in ID column is not a plain number and will be ignored. A new ID may be generated.`);
      }
    }
  });
  Logger.log(`✅ Added existing numeric IDs from Attendance sheets to used codes set. Total unique used NUMERIC codes found: ${allUsedCodes.size}. Total NUMERIC ID mappings found: ${belMap.size}`);

  // --- Step 3: Initialize NUMERIC ID Code generator ---
  let highestFoundNum = 0;
  allUsedCodes.forEach(code => { // 'code' here is already a number
    if (typeof code === 'number' && code > highestFoundNum) {
      highestFoundNum = code;
    }
  });

  let codeCounter = highestFoundNum + 1;
  Logger.log(`✅ Initialized NUMERIC ID code counter to: ${codeCounter}. Highest valid number found among existing ID codes was ${highestFoundNum}.`);

  // --- Step 4: Define the function to generate the next unique NUMERIC ID code ---
  // Consider renaming to generateNumericId
  const generateBEL = () => {
    while (true) {
      const currentNumericCode = codeCounter; // This is the number to try
      if (!allUsedCodes.has(currentNumericCode)) {
        allUsedCodes.add(currentNumericCode); // Add the number to the set
        codeCounter++; // Increment the main counter for the next call
        return currentNumericCode; // Return the newly generated unique number
      }
      codeCounter++;
      if (codeCounter > highestFoundNum + 20000) { // Safety break
        Logger.log(`ERROR: ID code counter (${codeCounter}) significantly exceeds highest found number (${highestFoundNum}). Potential issue.`);
        throw new Error("ID code counter exceeded a safe threshold, potential infinite loop or too many users without pre-existing IDs.");
      }
    }
  };

  // --- Step 5: Process Attendance Data and Assign Final NUMERIC IDs ---
  const results = [];
  attendanceDataRaw.forEach(row => {
    if (row.length < 2) {
      Logger.log(`⚠️ Processing Attendance: Skipping row due to insufficient columns for Name. Row data: ${JSON.stringify(row)}`);
      return;
    }
    const name = normalize(row[1]);
    if (!name) {
      Logger.log(`⚠️ Processing Attendance: Skipping row due to missing Name in Column B. Row data: ${JSON.stringify(row)}`);
      return;
    }

    let numericBel; // This will hold the purely numeric ID
    if (belMap.has(name)) {
      numericBel = belMap.get(name); // This is a number from the map
    } else {
      numericBel = generateBEL(); // This is a new number from the generator
      belMap.set(name, numericBel); // Add the new mapping
      Logger.log(`✅ Generated new NUMERIC ID ${numericBel} for name "${row[1]}".`);
    }

    let formattedRow = Array(11).fill("");
    formattedRow[0] = numericBel; // Column A: Assigned PURELY NUMERIC ID

    // Map data based on source sheet structure
    if (row.length >= 11 && typeof row[10] !== 'undefined') { // eData like
      formattedRow[1] = row[1];  // Full Name
      formattedRow[2] = row[2];  // Event Name
      formattedRow[3] = row[3];  // Event ID
      formattedRow[4] = row[4];  // First Name
      formattedRow[5] = row[5];  // Last Name
      formattedRow[6] = row[6];  // Email
      formattedRow[7] = row[7];  // Phone
      formattedRow[8] = row[8];  // Form Sheet
      formattedRow[9] = row[9];  // Role
      formattedRow[10] = row[10]; // Timestamp
    } else if (row.length >= 8 && typeof row[4] !== 'undefined') { // sData like
      formattedRow[1] = row[1];         // Full Name
      formattedRow[2] = "Sunday Service"; // Event Name
      formattedRow[3] = "Service";        // Event ID
      formattedRow[4] = row[2];         // First Name (from sData C)
      formattedRow[5] = row[3];         // Last Name (from sData D)
      formattedRow[6] = row[6];         // Email (from sData G)
      // formattedRow[7] (Phone) and formattedRow[8] (Form Sheet) remain ""
      formattedRow[9] = "";             // Role
      formattedRow[10] = row[4];        // Timestamp (from sData E)
    } else {
      Logger.log(`⚠️ Processing Attendance: Skipping row for "${name}" (ID: ${numericBel}) with unrecognized structure. Row data: ${JSON.stringify(row)}`);
      return; // Skip if structure is unexpected
    }
    results.push(formattedRow);
  });

  Logger.log(`✅ Total attendance records matched and formatted with NUMERIC IDs: ${results.length}`);
  return results; // Each row[0] in this array will be a number
}

// NOTE: This function assumes getDataFromSheets() is defined elsewhere and returns
// { sData, eData, dData }.
// You would need to define getDataFromSheets(), for example, a MOCK version:
function getDataFromSheets() {
  // This is a MOCK function. You need to implement this to fetch your actual sheet data.
  Logger.log("ℹ️ getDataFromSheets: Using MOCK data. Implement actual data fetching.");
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Assuming this script runs in Google Sheets environment
  const dirSheet = ss.getSheetByName("Directory");
  const eventSheet = ss.getSheetByName("Event Attendance");
  const serviceSheet = ss.getSheetByName("Service Attendance");

  // Simulating data fetching:
  // Replace with actual getLastRow and getRange().getValues()
  // The provided code (matchOrAssignBelCodes) uses slice(1) on dData, eData, sData,
  // implying it expects these arrays to include a header row that slice(1) will skip.
  return {
    dData: dirSheet ? dirSheet.getDataRange().getValues() : [['Timestamp', 'BEL', 'Full Name']], // Header + example data structure
    eData: eventSheet ? eventSheet.getDataRange().getValues() : [['BEL', 'Full Name', 'Event', 'ID', 'First', 'Last', 'Email', 'Phone', 'Form', 'Role', 'Timestamp']], // Header + data
    sData: serviceSheet ? serviceSheet.getDataRange().getValues() : [['BEL', 'Full Name', 'First', 'Last', 'Timestamp', 'Status', 'Email', 'Notes']] // Header + data
  };
}

// Example of how you might call this and log the results (for testing):
/*
function testRun() {
  const formattedAttendance = matchOrAssignBelCodes();
  if (formattedAttendance.length > 0) {
    Logger.log("Sample of formatted data (first 5 rows):");
    formattedAttendance.slice(0, 5).forEach(row => Logger.log(JSON.stringify(row)));
    // Here you would typically write `formattedAttendance` to a new sheet or update an existing one.
    // For example:
    // const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formatted Output");
    // if (outputSheet) {
    //   outputSheet.clearContents(); // Clear old data
    //   outputSheet.getRange(1, 1, formattedAttendance.length, formattedAttendance[0].length).setValues(formattedAttendance);
    //   Logger.log(`✅ Data written to "Formatted Output" sheet.`);
    // } else {
    //   Logger.log(`⚠️ Sheet "Formatted Output" not found.`);
    // }
  } else {
    Logger.log("No data processed or returned.");
  }
}
*/