/**
 * Sunday Registration Checklist System
 * Creates a user-friendly checkbox interface for registration teams
 * to quickly check-in regular attendees, with automatic transfer to Sunday Service sheet
 */

const LOCAL_ID_SHEETS = ["Sunday Service", "Sunday Registration", "Event Attendance", "Service Attendance"];

function createSundayRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Sunday Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      'Sunday Registration sheet already exists. Do you want to recreate it?',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      ss.deleteSheet(regSheet);
    } else {
      return;
    }
  }
  regSheet = ss.insertSheet("Sunday Registration");
  setupRegistrationSheetLayout(regSheet);
  populateRegistrationList(regSheet);
  // addRegistrationMenu(); // Not strictly needed here as onOpen will handle it, but harmless
  Logger.log("✅ Sunday Registration sheet created successfully! Person IDs are populated via new logic.");
  SpreadsheetApp.getUi().alert(
    'Registration Sheet Created!',
    'Sunday Registration sheet has been created and populated with active members.\n\n' +
    'Person IDs in Column A are fetched/generated based on Directory, local sheets, or new.\n\n' +
    'The registration team can now:\n' +
    '1. Enter the service date in cell B2\n' +
    '2. Check the boxes for attendees\n' +
    '3. Click "Submit Attendance" from the "📋 Sunday Check-in" menu to transfer to Service Attendance sheet\n\n' +
    'Menus have been added/updated for easy access to functions.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupRegistrationSheetLayout(sheet) {
  sheet.clear();
  sheet.getRange("A1").setValue("🏛️ SUNDAY SERVICE REGISTRATION").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:E1").merge().setHorizontalAlignment("center");
  sheet.getRange("A2").setValue("📅 Service Date:");
  sheet.getRange("B2").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  sheet.getRange("A3").setValue("📝 Instructions: Check the box next to each person who is present today");
  sheet.getRange("A3:E3").merge();
  sheet.getRange("A4").setValue("🔄 Refresh List");
  sheet.getRange("B4").setValue("✅ Submit Attendance");
  sheet.getRange("C4").setValue("🧹 Clear All Checks");
  sheet.getRange("D4").setValue("Status: Ready");

  const headers = ["ID", "Full Name", "First Name", "Last Name", "✓ Present"];
  sheet.getRange("A5:E5").setValues([headers]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 80);
  sheet.hideColumns(1);

  sheet.getRange("A1:E4").setBackground("#f8f9fa");
  sheet.getRange("A2:B2").setBackground("#e3f2fd");
  sheet.getRange("A4:D4").setBackground("#fff3e0");
  sheet.setFrozenRows(5);
  Logger.log("✅ Registration sheet layout created.");
}

function findHighestIdInLocalSheets(sheetNamesArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let highestId = 0;
  sheetNamesArray.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 1) {
        let startDataRow = 1;
        if (sheetName === "Sunday Registration" && lastRow >= 6) {
          startDataRow = 6;
        } else if ((sheetName === "Event Attendance" || sheetName === "Sunday Service") && lastRow >= 2) {
          const headerValue = String(sheet.getRange(1, 1).getDisplayValue() || "").trim().toLowerCase();
          const headerValueColB = String(sheet.getRange(1, 2).getDisplayValue() || "").trim().toLowerCase();
          if ((headerValue === "person id" || headerValue === "id") &&
            (headerValueColB.includes("full name") || headerValueColB.includes("name"))) {
            startDataRow = 2;
          }
        }
        if (lastRow >= startDataRow) {
          const ids = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, 1).getValues();
          ids.forEach(row => {
            const id = parseInt(row[0]);
            if (!isNaN(id) && id > highestId) {
              highestId = id;
            }
          });
        }
      }
    } else {
      Logger.log(`Sheet "${sheetName}" not found for local ID generation base.`);
    }
  });
  Logger.log(`Highest current ID found across local sheets (${sheetNamesArray.join(', ')}): ${highestId}`);
  return highestId;
}

function findHighestIdInDirectory() {
  let highestId = 0;
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) {
      Logger.log('⚠️ DIRECTORY_SPREADSHEET_ID script property not set (for findHighestIdInDirectory).');
      return 0;
    }
    const directorySS = SpreadsheetApp.openById(directoryId);
    const directorySheet = directorySS.getSheetByName("Directory");
    if (directorySheet) {
      const lastRow = directorySheet.getLastRow();
      if (lastRow >= 2) {
        const ids = directorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
        ids.forEach(row => {
          const id = parseInt(row[0]);
          if (!isNaN(id) && id > highestId) {
            highestId = id;
          }
        });
      }
      Logger.log(`Highest ID found in external Directory: ${highestId}`);
    } else {
      Logger.log('⚠️ "Directory" sheet not found in the external spreadsheet (for findHighestIdInDirectory).');
    }
  } catch (error) {
    Logger.log(`❌ Error in findHighestIdInDirectory: ${error.message}`);
  }
  return highestId;
}

function getDirectoryDataMap() {
  const directoryDataMap = new Map();
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) {
      Logger.log('⚠️ DIRECTORY_SPREADSHEET_ID script property not set for getDirectoryDataMap. Please set it via the Config menu.');
      return directoryDataMap;
    }
    const directorySS = SpreadsheetApp.openById(directoryId);
    const directorySheet = directorySS.getSheetByName("Directory");
    if (directorySheet) {
      const directoryValues = directorySheet.getDataRange().getValues();
      if (directoryValues.length > 1) {
        const headers = directoryValues[0].map(h => String(h || "").trim().toLowerCase());
        const idColIndex = 0;
        const nameColIndex = 1;

        let emailColIndex = headers.indexOf("email");
        let firstNameColIndex = headers.indexOf("first name");
        if (firstNameColIndex === -1) firstNameColIndex = headers.indexOf("firstname");
        let lastNameColIndex = headers.indexOf("last name");
        if (lastNameColIndex === -1) lastNameColIndex = headers.indexOf("lastname");

        for (let i = 1; i < directoryValues.length; i++) {
          const row = directoryValues[i];
          const personId = String(row[idColIndex] || "").trim();
          const fullName = String(row[nameColIndex] || "").trim();
          if (personId && fullName) {
            const normalizedFullName = fullName.toUpperCase();
            directoryDataMap.set(normalizedFullName, {
              id: personId,
              email: emailColIndex !== -1 ? String(row[emailColIndex] || "").trim() : "",
              firstName: firstNameColIndex !== -1 ? String(row[firstNameColIndex] || "").trim() : "",
              lastName: lastNameColIndex !== -1 ? String(row[lastNameColIndex] || "").trim() : ""
            });
          }
        }
      }
      Logger.log(`Directory data map created with ${directoryDataMap.size} entries.`);
    } else {
      Logger.log('⚠️ "Directory" sheet not found in the external spreadsheet for getDirectoryDataMap.');
    }
  } catch (error) {
    Logger.log(`❌ Error in getDirectoryDataMap: ${error.message}. Ensure the ID is correct and you have access.`);
  }
  return directoryDataMap;
}

function getLocalSheetIdMap(sheetName, idColNum = 1, nameColNum = 2) {
  const localIdMap = new Map();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 1) {
      const data = sheet.getRange(1, 1, lastRow, Math.max(idColNum, nameColNum)).getValues();
      let dataStartRowIndex = 0;

      if (lastRow > 1) {
        const headerIdCell = String(data[0][idColNum - 1] || "").trim().toLowerCase();
        const headerNameCell = String(data[0][nameColNum - 1] || "").trim().toLowerCase();
        if ((headerIdCell === "id" || headerIdCell === "person id") &&
          (headerNameCell.includes("name"))) {
          dataStartRowIndex = 1;
        }
      }

      for (let i = dataStartRowIndex; i < data.length; i++) {
        const row = data[i];
        const personId = String(row[idColNum - 1] || "").trim();
        const fullName = String(row[nameColNum - 1] || "").trim();
        if (personId && fullName) {
          localIdMap.set(fullName.toUpperCase(), personId);
        }
      }
    }
    Logger.log(`Local ID map created for "${sheetName}" with ${localIdMap.size} entries.`);
  } else {
    Logger.log(`⚠️ Local sheet "${sheetName}" not found for ID lookup.`);
  }
  return localIdMap;
}

function populateRegistrationList(regSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!regSheet) {
    regSheet = ss.getSheetByName("Sunday Registration");
    if (!regSheet) { Logger.log("❌ Sunday Registration sheet not found for populateRegistrationList"); return; }
  }
  const statsSheet = ss.getSheetByName("Attendance Stats");
  if (!statsSheet) {
    Logger.log("❌ Attendance Stats sheet not found for populateRegistrationList. Cannot populate.");
    if (SpreadsheetApp.getUi()) { // Only alert if in UI context
      SpreadsheetApp.getUi().alert("Error", "Attendance Stats sheet not found. Cannot populate registration list.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return;
  }
  const statsData = statsSheet.getDataRange().getValues();
  if (statsData.length < 2) { Logger.log("❌ No data in Attendance Stats sheet"); return; }

  const directoryMap = getDirectoryDataMap();
  if (directoryMap.size === 0 && PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID')) {
    Logger.log("⚠️ Directory map is empty but DIRECTORY_SPREADSHEET_ID is set. Check Directory sheet content or ID validity.");
  }
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);
  const sundayServiceIdMap = getLocalSheetIdMap("Sunday Service", 1, 2);

  // --- MODIFIED LOGIC START ---
  // Start with the highest ID from the Directory. If directory ID is not available or 0,
  // then check local sheets.
  let nextGeneratedId = findHighestIdInDirectory();
  if (nextGeneratedId === 0) {
    nextGeneratedId = findHighestIdInLocalSheets(LOCAL_ID_SHEETS);
  }
  // --- MODIFIED LOGIC END ---

  Logger.log(`Initial base for nextGeneratedId (starting with Directory, then local): ${nextGeneratedId}`);

  const activeMembersData = [];
  const today = new Date();
  const threeMonthsAgo = new Date(today.getTime() - (90 * 24 * 60 * 60 * 1000));
  const processedNewPersonsInThisRun = new Map();

  for (let i = 1; i < statsData.length; i++) {
    const statsRow = statsData[i];
    const fullNameFromStats = String(statsRow[1] || "").trim();
    let firstNameFromStats = String(statsRow[2] || "").trim();
    let lastNameFromStats = String(statsRow[3] || "").trim();
    const month = statsRow[5];
    const lastDateStr = statsRow[7];
    const activityLevel = statsRow[10];

    if (!fullNameFromStats) continue;

    let includeMember = false;
    let lastDate = null;
    if (lastDateStr) {
      if (lastDateStr instanceof Date) {
        lastDate = lastDateStr;
      } else {
        try {
          lastDate = new Date(lastDateStr);
          if (isNaN(lastDate.getTime())) lastDate = null;
        } catch (e) { lastDate = null; }
      }
    }
    if (activityLevel === "Core" || activityLevel === "Active") { includeMember = true; }
    else if (activityLevel === "Inactive" && lastDate && lastDate > threeMonthsAgo) { includeMember = true; }
    else if (month > 0) { includeMember = true; }

    if (includeMember) {
      let personId;
      let finalFirstName = firstNameFromStats;
      let finalLastName = lastNameFromStats;
      const normalizedFullName = fullNameFromStats.toUpperCase();

      const directoryEntry = directoryMap.get(normalizedFullName);
      const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
      const serviceEntryId = sundayServiceIdMap.get(normalizedFullName);
      const alreadyProcessedNew = processedNewPersonsInThisRun.get(normalizedFullName);

      if (directoryEntry && directoryEntry.id) {
        personId = directoryEntry.id;
        finalFirstName = directoryEntry.firstName || finalFirstName;
        finalLastName = directoryEntry.lastName || finalLastName;
      } else if (eventEntryId) {
        personId = eventEntryId;
      } else if (serviceEntryId) {
        personId = serviceEntryId;
      } else if (alreadyProcessedNew) {
        personId = alreadyProcessedNew.id;
        finalFirstName = alreadyProcessedNew.firstName;
        finalLastName = alreadyProcessedNew.lastName;
      } else {
        nextGeneratedId++;
        personId = String(nextGeneratedId);
        processedNewPersonsInThisRun.set(normalizedFullName, { id: personId, firstName: finalFirstName, lastName: finalLastName });
        directoryMap.set(normalizedFullName, {
          id: personId,
          email: "",
          firstName: finalFirstName,
          lastName: finalLastName
        });
        Logger.log(`Generated new ID ${personId} for ${fullNameFromStats}.`);
      }
      activeMembersData.push([personId, fullNameFromStats, finalFirstName, finalLastName, false]);
    }
  }

  activeMembersData.sort((a, b) => (String(a[3]) || "").toLowerCase().localeCompare((String(b[3]) || "").toLowerCase()));

  const lastDataRowOnSheet = regSheet.getLastRow();
  if (lastDataRowOnSheet > 5) {
    regSheet.getRange(6, 1, lastDataRowOnSheet - 5, 5).clearContent().clearFormat();
  }
  if (activeMembersData.length > 0) {
    const startRow = 6;
    regSheet.getRange(startRow, 1, activeMembersData.length, 5).setValues(activeMembersData);
    const checkboxRange = regSheet.getRange(startRow, 5, activeMembersData.length, 1);
    checkboxRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    regSheet.getRange(startRow, 1, activeMembersData.length, 5).setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, startRow, activeMembersData.length); // Use updated refresh
  }
  regSheet.getRange("D4").setValue(`Status: ${activeMembersData.length} members loaded`);
  Logger.log(`✅ Registration list populated with ${activeMembersData.length} members. IDs fetched/generated.`);
}

function addPersonToRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();

  const nameResponse = ui.prompt('Add Person', 'Enter the full name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const fullNameEntered = String(nameResponse.getResponseText() || "").trim();
  if (!fullNameEntered) {
    ui.alert('Input Error', 'Please enter a valid name', ui.ButtonSet.OK);
    return;
  }

  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    const existingNames = regSheet.getRange(6, 2, lastDataRow - 5, 1).getValues();
    if (existingNames.some(row => row[0] && String(row[0]).trim().toLowerCase() === fullNameEntered.toLowerCase())) {
      ui.alert('Duplicate Entry', 'This person is already in the registration list.', ui.ButtonSet.OK);
      return;
    }
  }

  const directoryMap = getDirectoryDataMap();
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);
  const sundayServiceIdMap = getLocalSheetIdMap("Sunday Service", 1, 2);
  const normalizedFullName = fullNameEntered.toUpperCase();

  let personIdToAdd;
  let firstNameToAdd = "";
  let lastNameToAdd = "";

  const directoryEntry = directoryMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
  const serviceEntryId = sundayServiceIdMap.get(normalizedFullName);

  if (directoryEntry && directoryEntry.id) {
    personIdToAdd = directoryEntry.id;
    firstNameToAdd = directoryEntry.firstName;
    lastNameToAdd = directoryEntry.lastName;
  } else if (eventEntryId) {
    personIdToAdd = eventEntryId;
  } else if (serviceEntryId) {
    personIdToAdd = serviceEntryId;
  } else {
    // --- MODIFIED LOGIC START ---
    // Calculate the highest ID from directory first, then local sheets and current registration sheet
    let currentHighestOverallId = findHighestIdInDirectory();
    currentHighestOverallId = Math.max(currentHighestOverallId, findHighestIdInLocalSheets(LOCAL_ID_SHEETS));

    if (lastDataRow >= 6) {
      const currentSheetIds = regSheet.getRange(6, 1, lastDataRow - 5, 1).getValues();
      currentSheetIds.forEach(row => {
        const id = parseInt(row[0]);
        if (!isNaN(id) && id > currentHighestOverallId) {
          currentHighestOverallId = id;
        }
      });
    }
    // --- MODIFIED LOGIC END ---
    currentHighestOverallId++;
    personIdToAdd = String(currentHighestOverallId);
    Logger.log(`Generated new ID ${personIdToAdd} for manually added ${fullNameEntered}.`);
  }

  if (!firstNameToAdd && fullNameEntered) {
    const nameParts = fullNameEntered.split(/\s+/);
    firstNameToAdd = nameParts[0] || "";
    lastNameToAdd = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  const newRowData = [personIdToAdd, fullNameEntered, firstNameToAdd, lastNameToAdd, false];
  regSheet.getRange(nextSheetRow, 1, 1, 5).setValues([newRowData]);
  regSheet.getRange(nextSheetRow, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  const newRowRange = regSheet.getRange(nextSheetRow, 1, 1, 5);
  newRowRange.setBorder(true, true, true, true, true, true);
  refreshRowFormatting(regSheet); // Refresh all formatting

  ui.alert('Person Added!', `${fullNameEntered} has been added with ID ${personIdToAdd}.`, ui.ButtonSet.OK);
  Logger.log(`✅ Manually added ${fullNameEntered} (ID: ${personIdToAdd}) to registration list.`);
}

function addNewMemberToRegistration(memberData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) { Logger.log("Error: Sunday Registration sheet not found for addNewMember."); return; }

  const fullNameFromForm = String(memberData.fullName || "").trim();
  if (!fullNameFromForm) { Logger.log("Cannot add member: Full name missing from memberData."); return; }

  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    const existingNames = regSheet.getRange(6, 2, lastDataRow - 5, 1).getValues();
    if (existingNames.some(row => row[0] && String(row[0]).trim().toLowerCase() === fullNameFromForm.toLowerCase())) {
      Logger.log(`Member ${fullNameFromForm} already exists in registration list. Not adding from external data.`); return;
    }
  }

  const directoryMap = getDirectoryDataMap();
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);
  const sundayServiceIdMap = getLocalSheetIdMap("Sunday Service", 1, 2);
  const normalizedFullName = fullNameFromForm.toUpperCase();

  let personIdToAdd;
  let firstNameToAdd = String(memberData.firstName || "").trim();
  let lastNameToAdd = String(memberData.lastName || "").trim();

  const directoryEntry = directoryMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
  const serviceEntryId = sundayServiceIdMap.get(normalizedFullName);

  if (directoryEntry && directoryEntry.id) {
    personIdToAdd = directoryEntry.id;
    firstNameToAdd = directoryEntry.firstName || firstNameToAdd;
    lastNameToAdd = directoryEntry.lastName || lastNameToAdd;
  } else if (eventEntryId) {
    personIdToAdd = eventEntryId;
  } else if (serviceEntryId) {
    personIdToAdd = serviceEntryId;
  } else {
    // --- MODIFIED LOGIC START ---
    // Calculate the highest ID from directory first, then local sheets and current registration sheet
    let currentHighestOverallId = findHighestIdInDirectory();
    currentHighestOverallId = Math.max(currentHighestOverallId, findHighestIdInLocalSheets(LOCAL_ID_SHEETS));

    if (lastDataRow >= 6) {
      const currentSheetIds = regSheet.getRange(6, 1, lastDataRow - 5, 1).getValues();
      currentSheetIds.forEach(row => {
        const id = parseInt(row[0]);
        if (!isNaN(id) && id > currentHighestOverallId) {
          currentHighestOverallId = id;
        }
      });
    }
    // --- MODIFIED LOGIC END ---
    currentHighestOverallId++;
    personIdToAdd = String(currentHighestOverallId);
  }

  if (!firstNameToAdd && fullNameFromForm) {
    const nameParts = fullNameFromForm.split(/\s+/);
    firstNameToAdd = nameParts[0] || "";
    lastNameToAdd = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  const newRow = [personIdToAdd, fullNameFromForm, firstNameToAdd, lastNameToAdd, false];
  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  regSheet.getRange(nextSheetRow, 1, 1, 5).setValues([newRow]);
  regSheet.getRange(nextSheetRow, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  refreshRowFormatting(regSheet); // Refresh all formatting
  Logger.log(`✅ Added new member ${fullNameFromForm} (ID: ${personIdToAdd}) from external data.`);
}

function submitRegistrationAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const serviceDateValue = regSheet.getRange("B2").getValue();
  if (!serviceDateValue || !(serviceDateValue instanceof Date) || isNaN(serviceDateValue.getTime())) {
    SpreadsheetApp.getUi().alert("Input Error", "Please enter a valid service date in cell B2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // --- MODIFIED LINE START ---
  // Get the spreadsheet's timezone to format the date correctly
  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const formattedServiceDate = Utilities.formatDate(serviceDateValue, spreadsheetTimezone, "MM/dd/yyyy");
  // --- MODIFIED LINE END ---

  regSheet.getRange("D4").setValue("Status: Processing...");

  try {
    const serviceSheet = ss.getSheetByName("Service Attendance");
    if (!serviceSheet) {
      SpreadsheetApp.getUi().alert("Error", "'Service Attendance' sheet not found. Please create it.", SpreadsheetApp.getUi().ButtonSet.OK);
      regSheet.getRange("D4").setValue("Status: Error - Service Attendance sheet missing");
      throw new Error("'Service Attendance' sheet not found");
    }

    const directoryMap = getDirectoryDataMap();

    const lastRegDataRow = regSheet.getLastRow();
    if (lastRegDataRow < 6) {
      regSheet.getRange("D4").setValue("Status: No members to process");
      SpreadsheetApp.getUi().alert("No Members", "No members listed to process for attendance.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const regData = regSheet.getRange(6, 1, lastRegDataRow - 5, 5).getValues();
    const attendanceEntries = [];
    let checkedCount = 0;

    for (const row of regData) {
      const [personId, fullName, firstName, lastName, isChecked] = row;

      if (isChecked === true && fullName && String(fullName).trim() !== "") {
        let email = "";
        const normalizedFullName = String(fullName).trim().toUpperCase();
        const directoryEntry = directoryMap.get(normalizedFullName);
        if (directoryEntry && directoryEntry.email) {
          email = directoryEntry.email;
        } else {
          Logger.log(`Email not found in Directory for ${fullName} (ID: ${personId}). Will submit blank email.`);
        }

        const notes = "";
        attendanceEntries.push([
          personId,
          fullName,
          firstName || "",
          lastName || "",
          // --- MODIFIED LINE START ---
          formattedServiceDate, // Use the formatted date string here
          // --- MODIFIED LINE END ---
          "No",
          email,
          notes,
          new Date()
        ]);
        checkedCount++;
      }
    }

    if (attendanceEntries.length === 0) {
      regSheet.getRange("D4").setValue("Status: No members checked");
      SpreadsheetApp.getUi().alert("No Checks", "No members were checked for attendance.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const nextRowServiceSheet = findLastRowWithData(serviceSheet) + 1;
    serviceSheet.getRange(nextRowServiceSheet, 1, attendanceEntries.length, 9).setValues(attendanceEntries);
    // Ensure the date column in 'Service Attendance' is formatted as a date
    serviceSheet.getRange(nextRowServiceSheet, 5, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy");

    regSheet.getRange(6, 5, lastRegDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue(`Status: ${checkedCount} attendees submitted`);
    SpreadsheetApp.getUi().alert(
      'Attendance Submitted!',
      `Successfully submitted attendance for ${checkedCount} members to 'Service Attendance' sheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log(`✅ Successfully submitted ${checkedCount} attendance entries.`);
  } catch (error) {
    regSheet.getRange("D4").setValue("Status: Error occurred");
    Logger.log(`❌ Error submitting attendance: ${error.message}\n${error.stack || ""}`);
    SpreadsheetApp.getUi().alert("Error", `Error submitting attendance: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function clearAllChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    regSheet.getRange(6, 5, lastDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue("Status: All checks cleared");
    Logger.log("✅ All checkboxes cleared");
  } else {
    regSheet.getRange("D4").setValue("Status: No checks to clear");
    Logger.log("ℹ️ No data rows found to clear checks from.");
  }
}

function addRegistrationMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Sunday Check-in')
    .addItem('🔄 Refresh Member List', 'populateRegistrationList')
    .addItem('✅ Submit Attendance', 'submitRegistrationAttendance')
    .addItem('🧹 Clear All Checks', 'clearAllChecks')
    .addSeparator()
    .addItem('➕ Add Attendee (Quick Add)', 'addPersonToRegistration')
    .addItem('🔲 Add/Reformat Checkboxes', 'addCheckboxesToRegistration')
    .addItem('🗑️ Remove Attendee', 'removePersonFromRegistration')
    .addItem('Sort by Last Name', 'sortRegistrationByLastName')
    .addSeparator()
    .addItem('🆕 Create Empty Registration Sheet', 'createEmptyRegistrationSheet')
    .addToUi();
  Logger.log("✅ Sunday Check-in menu definition attempted by addRegistrationMenu.");
}

function addCheckboxesToRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastRow = regSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert("No Data", "No data found below row 5 to add checkboxes to. Please add member data starting row 6.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dataRangeForNames = regSheet.getRange(6, 2, lastRow - 5, 1);
  const nameValues = dataRangeForNames.getValues();
  let rowsWithActualNames = 0;
  for (let i = 0; i < nameValues.length; i++) {
    if (String(nameValues[i][0] || "").trim() !== "") {
      rowsWithActualNames = i + 1;
    }
  }
  if (rowsWithActualNames === 0 && lastRow >= 6) {
    SpreadsheetApp.getUi().alert("No Names Found", "No names found in Column B (Full Name) from row 6 downwards. Cannot add checkboxes meaningfully.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  } else if (rowsWithActualNames === 0) {
    SpreadsheetApp.getUi().alert("No Data", "No data rows found to add checkboxes to.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    const checkboxRange = regSheet.getRange(6, 5, rowsWithActualNames, 1);
    checkboxRange.clearContent().setValue(false).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

    const dataFormattingRange = regSheet.getRange(6, 1, rowsWithActualNames, 5);
    dataFormattingRange.setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, 6, rowsWithActualNames);

    regSheet.getRange("D4").setValue(`Status: ${rowsWithActualNames} members ready`);
    SpreadsheetApp.getUi().alert('Checkboxes Added/Reformatted!', `Successfully added/reformatted checkboxes for ${rowsWithActualNames} member rows.\n\nSheet is ready:\n1. Enter service date in B2\n2. Check attendance\n3. Click Submit Attendance`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`✅ Added/Reformatted checkboxes to ${rowsWithActualNames} rows`);
  } catch (error) {
    Logger.log(`❌ Error adding/reformatting checkboxes: ${error.message}`);
    SpreadsheetApp.getUi().alert("Error", `Error adding/reformatting checkboxes: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createEmptyRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Sunday Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Sheet Already Exists', 'Sunday Registration sheet already exists. Recreate it as empty?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) ss.deleteSheet(regSheet);
    else return;
  }
  regSheet = ss.insertSheet("Sunday Registration");
  setupRegistrationSheetLayout(regSheet);
  Logger.log("✅ Empty Sunday Registration sheet created.");
  SpreadsheetApp.getUi().alert(
    'Empty Registration Sheet Created!',
    'Sunday Registration sheet is ready for manual data entry (Columns A-E for data, starting row 6).\n\n' +
    'Person IDs (Col A) will be fetched from Directory or generated if you use Refresh/Add Attendee.\n\n' +
    'INSTRUCTIONS:\n' +
    '1. Paste directory data starting row 6 (Full Name in Col B, First in C, Last in D - ID will be handled by other functions)\n' +
    '2. Use "📋 Sunday Check-in" → "🔲 Add/Reformat Checkboxes" to set up column E.\n' +
    '3. Enter service date in B2 and start checking attendance!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removePersonFromRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow < 6) {
    ui.alert('No People', 'No people in the list to remove (list is empty below row 5).', ui.ButtonSet.OK);
    return;
  }

  const nameResponse = ui.prompt('Remove Person', 'Enter the FULL NAME of the person to remove (case-insensitive):', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const nameToRemove = String(nameResponse.getResponseText() || "").trim().toLowerCase();
  if (!nameToRemove) {
    ui.alert("No Name Entered", "No name entered to remove.", ui.ButtonSet.OK);
    return;
  }

  const allData = regSheet.getRange(6, 1, lastDataRow - 5, 5).getValues();
  let rowToDeleteInSheet = -1;
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][1] && String(allData[i][1]).trim().toLowerCase() === nameToRemove) {
      rowToDeleteInSheet = i + 6;
      break;
    }
  }

  if (rowToDeleteInSheet > 0) {
    regSheet.deleteRow(rowToDeleteInSheet);
    ui.alert('Person Removed!', `'${nameResponse.getResponseText().trim()}' has been removed.`, ui.ButtonSet.OK);
    Logger.log(`✅ Removed '${nameResponse.getResponseText().trim()}' from registration list, row ${rowToDeleteInSheet}`);
    refreshRowFormatting(regSheet);
  } else {
    ui.alert('Not Found', `Person '${nameResponse.getResponseText().trim()}' not found in the registration list.`, ui.ButtonSet.OK);
  }
}

function refreshRowFormatting(sheet, startDataRow = 6, numRowsInput = -1) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sunday Registration");
  if (!sheet) return;

  let numRowsToFormat = numRowsInput;
  if (numRowsToFormat === -1) {
    const lastSheetRowWithContent = findLastRowWithData(sheet); // Use robust last row detection
    if (lastSheetRowWithContent < startDataRow) {
      Logger.log("No data rows to format in refreshRowFormatting.");
      return;
    }
    numRowsToFormat = lastSheetRowWithContent - startDataRow + 1;
  }

  if (numRowsToFormat <= 0) {
    Logger.log("Calculated numRowsToFormat is <= 0 in refreshRowFormatting.");
    return;
  }

  for (let i = 0; i < numRowsToFormat; i++) {
    const currentRowInSheet = startDataRow + i;
    const rowRange = sheet.getRange(currentRowInSheet, 1, 1, 5);
    if (i % 2 === 1) {
      rowRange.setBackground("#f5f5f5");
    } else {
      rowRange.setBackground("white");
    }
  }
  Logger.log(`Row formatting refreshed for ${numRowsToFormat} rows starting at ${startDataRow}.`);
}


function findLastRowWithData(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  const maxCol = sheet.getMaxColumns(); // Use getMaxColumns to be safe
  if (maxCol === 0) return 0;

  for (let r = lastRow; r >= 1; r--) {
    const range = sheet.getRange(r, 1, 1, maxCol);
    if (!range.isBlank()) {
      return r;
    }
  }
  return 0;
}

function sortRegistrationByLastName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastDataRow = findLastRowWithData(regSheet); // Use robust last row
  if (lastDataRow < 6) {
    SpreadsheetApp.getUi().alert("No Data", "No data to sort (list is empty below row 5).", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const numDataRows = lastDataRow - 5;
  if (numDataRows <= 0) {
    SpreadsheetApp.getUi().alert("No Data", "No data rows to sort.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dataRange = regSheet.getRange(6, 1, numDataRows, 5);
  dataRange.sort({ column: 4, ascending: true });
  refreshRowFormatting(regSheet, 6, numDataRows);
  SpreadsheetApp.getUi().alert("List sorted by Last Name.");
  Logger.log("✅ Registration list sorted by last name.");
}

// ***** MASTER onOpen(e) FUNCTION FOR THE ENTIRE PROJECT *****
function onOpen(e) {
  Logger.log("Master onOpen triggered. AuthMode: " + (e ? e.authMode : 'N/A Event Object'));

  try {
    addRegistrationMenu();
    Logger.log("Call to addRegistrationMenu completed from onOpen.");
  } catch (error) {
    Logger.log("Error during addRegistrationMenu in onOpen: " + error.message + " Stack: " + error.stack);
  }

  try {
    if (typeof addTransferMenu === "function") {
      addTransferMenu();
      Logger.log("Call to addTransferMenu completed from onOpen.");
    } else {
      Logger.log("addTransferMenu function not found. Make sure it's defined in one of the project's .gs files.");
    }
  } catch (error) {
    Logger.log("Error during addTransferMenu in onOpen: " + error.message + " Stack: " + error.stack);
  }

  try {
    // Check if showDirectoryDialog function exists before trying to add the menu item
    if (typeof showDirectoryDialog === "function") {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('⚙️ Config')
        .addItem('Set Directory Spreadsheet URL…', 'showDirectoryDialog') // Changed menu text
        .addToUi();
      Logger.log("⚙️ Config menu added by onOpen.");
    } else {
      Logger.log("showDirectoryDialog function not found. Config menu item not added.");
    }
  } catch (error) {
    Logger.log("Error adding Config menu in onOpen: " + error.message + " Stack: " + error.stack);
  }
}

/**
 * Extracts the spreadsheet ID from a Google Sheet URL.
 * @param {string} url The full Google Sheet URL.
 * @returns {string} The extracted spreadsheet ID, or null if not found.
 */
function extractSpreadsheetIdFromUrl(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Allows the user to set the Directory Spreadsheet ID via a UI prompt,
 * accepting a full URL and extracting the ID.
 * This ID is stored in Script Properties.
 */
function showDirectoryDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentId = PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID');
  const promptMessage = currentId ?
    `Current Directory Spreadsheet ID: ${currentId}\n\nEnter the full URL of your main Directory file (or leave blank to keep current):` :
    'Please enter the full Google Sheet URL of your main Directory file:';

  const result = ui.prompt(
    'Set Directory Spreadsheet URL', // Changed prompt title
    promptMessage,
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    const urlOrId = result.getResponseText().trim();
    let finalId = '';

    if (urlOrId.includes('docs.google.com/spreadsheets/d/')) {
      // It's a URL, extract the ID
      finalId = extractSpreadsheetIdFromUrl(urlOrId);
      if (!finalId) {
        ui.alert('Invalid URL', 'Could not extract a valid Spreadsheet ID from the provided URL. Please ensure it\'s a valid Google Sheet URL.', ui.ButtonSet.OK);
        Logger.log(`❌ Invalid URL provided: ${urlOrId}`);
        return;
      }
    } else if (urlOrId) {
      // Assume it's a direct ID if not a URL, but log a warning if it doesn't look like a typical ID
      finalId = urlOrId;
      if (finalId.length < 30) { // Google Sheet IDs are typically long strings
        Logger.log(`⚠️ Short string provided, assuming it's a direct ID: ${finalId}`);
      }
    }

    if (finalId) {
      PropertiesService.getScriptProperties().setProperty('DIRECTORY_SPREADSHEET_ID', finalId);
      ui.alert('Success!', `Directory Spreadsheet ID has been set to: ${finalId}`, ui.ButtonSet.OK);
      Logger.log(`Directory Spreadsheet ID set to: ${finalId}`);
    } else if (urlOrId === "" && currentId) {
      // User cleared the field but there was a current ID. No change.
      ui.alert('No Change', 'Directory Spreadsheet ID was not changed from current value.', ui.ButtonSet.OK);
    } else if (urlOrId === "" && !currentId) {
      ui.alert('No ID Entered', 'Directory Spreadsheet ID was not set.', ui.ButtonSet.OK);
    }
  }
}