function updateAttendanceStatsSheet() {
  const finalData = calculateAttendanceStats();

  if (!finalData || finalData.length === 0) {
    Logger.log("❌ No final data to update the 'Attendance Stats' sheet.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet) {
    Logger.log("❌ 'Attendance Stats' sheet not found. Cannot update.");
    return;
  }

  // Format the data for writing to the sheet
  const output = finalData.map(row => {
    const [
      bel,
      fullName,
      ,
      ,
      quarter,
      month,
      volunteer,
      lastDate,
      lastEvent,
      total
    ] = row;

    const nameParts = fullName ? String(fullName).trim().split(/\s+/) : [];
    const firstName = nameParts.length > 0 ? nameParts[0] : "";
    const lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";

    let formattedDate = "";
    if (lastDate instanceof Date && !isNaN(lastDate.getTime())) {
      formattedDate = Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (lastDate) {
      formattedDate = String(lastDate);
      Logger.log(`⚠️ Invalid date for BEL ${bel}: ${lastDate}.`);
    }

    return [
      bel,
      fullName,
      firstName,
      lastName,
      quarter,
      month,
      volunteer,
      formattedDate,
      lastEvent,
      total
    ];
  });

  // Get how many rows and columns to write
  const numRows = output.length;
  const numCols = output[0].length;

  // Write the data starting at row 2, column 1 — only overwrite those rows
  sheet.getRange(2, 1, numRows, numCols).setValues(output);

  Logger.log(`✅ Wrote ${numRows} rows to 'Attendance Stats'. Existing rows below are untouched.`);
}
