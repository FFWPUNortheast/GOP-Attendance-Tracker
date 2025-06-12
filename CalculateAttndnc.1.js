/**
 * Calculates attendance statistics based on formatted raw data.
 * Groups entries by BEL code and summarizes attendance
 * for the current month, quarter, and year,
 * including volunteer instances and last attended date.
 *
 * Assumes rawData is an array of arrays where each inner array
 * represents a row conforming to the 11-column "Event Attendance" structure:
 * [0: BEL, 1: Full Name, 2: Event Name, 3: Event ID, 4: First Name, 5: Last Name, 6: Email, 7: Phone, 8: Form Sheet, 9: Role, 10: Timestamp (Date)]
 *
 * @returns {Array<Array<any>>} An array of arrays containing summarized attendance statistics per individual, or empty array if no data to process.
 */
function calculateAttendanceStats() {
  // Get the correctly formatted raw data from matchOrAssignBelCodes
  const rawData = matchOrAssignBelCodes();

  // Check if data was received from matchOrAssignBelCodes
  if (!rawData || rawData.length === 0) {
    Logger.log("❌ No data received from matchOrAssignBelCodes or data is empty after formatting.");
    return []; // Return empty array if no data
  }

  // Get current date information for filtering and calculations
  const now = new Date();
  const currentMonth = now.getMonth(); // 0 for January, 11 for December
  const currentQuarter = Math.floor(currentMonth / 3); // 0 for Q1, 3 for Q4
  const currentYear = now.getFullYear();

  // Map to group attendance entries by BEL code
  // Key: BEL code (string), Value: Array of attendance record objects
  const grouped = new Map();

  // Process each row of the formatted raw data received from matchOrAssignBelCodes
  rawData.forEach(row => {
    // --- Data Extraction with Correct Column Mapping (Expecting 11+ columns) ---
    // This check ensures the row has the minimum expected number of columns.
    // Based on the corrected matchOrAssignBelCodes, rows should have >= 11 columns.
    if (row.length < 11) {
      // This should ideally not happen if matchOrAssignBelCodes worked correctly and getDataFromSheets got enough columns.
      Logger.log(`⚠️ calculateAttendanceStats: Skipping row due to insufficient columns (${row.length} found). Expected at least 11. Row data (partial): ${JSON.stringify(row.slice(0, 11))}`);
      return; // Skip this row if it doesn't meet the minimum column requirement
    }

    // Extract values based on the 11-column "Event Attendance" structure (0-indexed)
    // These indices match the output format from the corrected matchOrAssignBelCodes
    const bel = row[0];         // Column A: ID Code (BEL)
    const name = row[1];        // Column B: Full Name
    const eventName = row[2];   // Column C: Event Name
    const eventId = row[3];     // Column D: Event ID
    const role = row[9];        // Column J: Role
    const dateStr = row[10];    // Column K: Timestamp (Date)
    // --- End Data Extraction ---

    // Parse the date string. Handles various date formats.
    // Apps Script getValues() might return Date objects directly for timestamps.
    let date;
    if (dateStr instanceof Date) {
        date = dateStr; // Use the Date object if already parsed
    } else if (typeof dateStr === 'number') {
         // Handle potential serial numbers for dates if getValues returns them (less common for timestamps)
         // This conversion might need adjustment based on how Apps Script handles serial numbers and timezones
         date = new Date((dateStr - (25567 + 2)) * 86400 * 1000); // Simple serial to Date conversion (might need refinement)
    }
     else {
        // Attempt to parse if it's a string or other type. String() handles null/undefined safely.
        date = new Date(String(dateStr));
    }


    // Validate the parsed date. isNaN(date.getTime()) is the standard way.
    if (isNaN(date.getTime())) {
      Logger.log(`⚠️ Skipping invalid date: "${dateStr}" found for BEL ${bel}. Full row data: ${JSON.stringify(row)}`);
      return; // Skip this row if the date is invalid
    }

    // Determine if the entry is for a Sunday Service or involves a Volunteer role
    // Use typeof checks for safety before string methods to avoid errors on null/undefined/non-strings
    const isSundayService = typeof eventName === 'string' && /sunday service/i.test(eventName);
    const isVolunteer = typeof role === 'string' && String(role).toLowerCase().includes("volunteer");

    // Create a unique key for each event instance for counting unique attendance
    // Use dateString for Sunday Service to treat each Sunday as a distinct event for stats
    // For other events, use a combination of name and ID
     // Ensure eventName and eventId are strings for key creation safety
    const eventNameKey = typeof eventName === 'string' ? eventName : 'UnknownEvent';
    const eventIdKey = typeof eventId === 'string' ? eventId : 'UnknownID';
    const eventKey = isSundayService ? `sunday service-${date.toDateString()}` : `${eventNameKey}-${eventIdKey}`;


    // Create a structured record object for this attendance entry
    const record = {
      name, // Uses the 'name' variable correctly assigned from row[1]
      date, // Uses the 'date' variable correctly parsed from row[10]
      eventKey,
      month: date.getMonth(),
      quarter: Math.floor(date.getMonth() / 3),
      year: date.getFullYear(),
      isVolunteer, // Uses the 'isVolunteer' flag based on 'role' from row[9]
      isSundayService, // Uses the 'isSundayService' flag based on 'eventName' from row[2]
    };

    // Group the record by BEL code
     // Ensure bel is treated as a string for map key consistency
    const belString = String(bel);
    if (!grouped.has(belString)) {
      grouped.set(belString, []); // Initialize an array for the BEL code if it doesn't exist
    }
    grouped.get(belString).push(record); // Add the current record to the array for this BEL code
  });

  // Array to store the final summary statistics
  const summary = [];

  // Process the grouped data to calculate summary statistics for each individual
  grouped.forEach((records, bel) => {
    // Initialize sets and counters for this individual's stats
    const uniqueEvents = new Set(); // To count total unique events attended across all time
    const monthEvents = new Set();  // To count unique events attended this current month/year
    const quarterEvents = new Set(); // To count unique events attended this current quarter/year
    let volunteerCount = 0;         // To count how many times the individual volunteered this current year

    // Iterate through the records for the current individual to populate sets and count volunteers
    records.forEach(r => {
      // Add event key to the set of all unique events attended by this person
      uniqueEvents.add(r.eventKey);

      // Check if the event is in the current year and month
      if (r.year === currentYear && r.month === currentMonth) {
        monthEvents.add(r.eventKey);
      }

      // Check if the event is in the current year and quarter
      if (r.year === currentYear && r.quarter === currentQuarter) {
        quarterEvents.add(r.eventKey);
      }

      // Count volunteer instances for the current year
      if (r.isVolunteer && r.year === currentYear) {
        volunteerCount++;
      }
    });

    // Sort records by date in descending order to find the latest attendance record
    // Use getTime() for reliable numeric comparison of Date objects
    records.sort((a, b) => b.date.getTime() - a.date.getTime());

    // Get information from the most recent attendance record
    // Ensure there's at least one record after filtering/skipping before accessing index 0
    const mostRecentRecord = records.length > 0 ? records[0] : null;

    let fullName = '';
    let lastDate = ''; // Will store the Date object
    let lastEventName = '';

    if (mostRecentRecord) {
        fullName = mostRecentRecord.name; // Full name from the latest record
        lastDate = mostRecentRecord.date; // Latest attendance date (as Date object)
        const lastEventKey = mostRecentRecord.eventKey; // Latest event key

        // Attempt to extract the event name from the last event key
        // Assumes the format is "EventName-EventId" or "sunday-DateString"
        const lastEventParts = lastEventKey.split('-');
        // Take the part before the first hyphen. If no hyphen, use the whole key.
        lastEventName = lastEventParts.length > 0 ? lastEventParts[0] : lastEventKey;
    }

    // Total unique events attended across all time is the size of the uniqueEvents set
    const totalUniqueEvents = uniqueEvents.size;

    // Add the calculated statistics for this individual to the summary array
    // The order here must match the columns written to in updateAttendanceStatsSheet (10 columns A-J)
    summary.push([
      bel,                  // Column A: BEL (from the grouped key)
      fullName,             // Column B: Full Name (from the most recent record)
      "",                   // Column C: First name placeholder (will be filled in update function)
      "",                   // Column D: Last name placeholder (will be filled in update function)
      quarterEvents.size,    // Column E: Count of unique events attended this quarter
      monthEvents.size,      // Column F: Count of unique events attended this month
      volunteerCount,        // Column G: Count of times volunteered this year
      lastDate,              // Column H: Last date attended (as Date object)
      lastEventName,         // Column I: Last event name (extracted from key)
      totalUniqueEvents      // Column J: Total count of unique events attended
    ]);
  });

  Logger.log("✅ Attendance stats calculated for: " + summary.length + " individuals.");

  return summary; // Return the array of summary statistics
}
