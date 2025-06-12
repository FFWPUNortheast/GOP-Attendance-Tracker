# GOP Attendance Tracker

This repository contains a collection of Google Apps Script modules for managing attendance tracking in Google Sheets.

The scripts automate tasks such as assigning unique IDs, consolidating form responses, computing statistics, and creating registration checklists.

## Overview

The project expects a spreadsheet with tabs such as **Sunday Service**, **Service Attendance**, **Event Attendance**, **Attendance Stats**, **Directory**, and **Sunday Registration**. Custom menus are added when the spreadsheet opens so you can transfer data or configure the external directory sheet.

Core modules include:

- **ActivityLevel.js** – Calculates member activity level by summing recent attendance counts.
- **CalculateAttndnc.1.js** – Summarizes attendance records for statistics by month, quarter, and year.
- **MatchAssignUniqueCodes.js** – Matches attendees to numeric IDs or generates new ones when needed.
- **NeedFollowUp.js** – Flags first–time attendees and people who may need follow‑up based on gaps in attendance.
- **PullSSAttendance.js** – Moves new check‑ins from the *Sunday Service* form sheet into *Service Attendance*.
- **SundayServiceChecklist.js** – Builds a *Sunday Registration* sheet with checkboxes for quick check‑ins.
- **Trigger.js** – Updates the stats sheet whenever changes occur.
- **LoadAndCleanData.js** – Loads raw data from multiple sheets and lets you set the external Directory spreadsheet ID.
- **UniqueCodeSndySrvc.js** – Processes form submissions on the Sunday Service sheet and assigns numeric IDs.
- **UpdateAttndce.js** – Writes computed statistics to the *Attendance Stats* tab.

See the individual files for detailed comments on how each function works.

## Setup

1. Open the script in the Apps Script editor and ensure each file is present.
2. When the spreadsheet opens, use the **Config** menu to set the ID of your external Directory spreadsheet.
3. Run `createSundayRegistrationSheet` to generate the registration checklist, or use the provided menu.
4. Optional triggers, such as `onAttendanceSheetsChange` or `onFormSubmitTransfer`, can be set up to automate updates.

## Usage Tips

- Many functions rely on specific column layouts. Review the expected columns in each sheet before running the scripts.
- Check the log output (`Logger.log`) if something doesn’t work as expected—most functions include detailed logging.
- Keep the Directory spreadsheet up to date so that ID assignment works smoothly.

This codebase can be extended to support additional events or custom reporting needs. Reading through the script comments and the earlier descriptions of how the workflow operates will help you adapt it to your organization.
