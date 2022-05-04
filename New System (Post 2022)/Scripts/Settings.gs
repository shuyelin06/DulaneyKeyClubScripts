/*
  This document will contain all of the constant variables used by the program
*/

// Get Value from the Connections Sheet
const ConnectionsSheetName = "Connections"; // Sheet in Officer Center with IDs (for connecting)
function getValue(cell) { return OfficerCenter.getSheetByName( ConnectionsSheetName ).getRange( cell ).getValue(); }

/* --- MIME Types --- */
const SpreadsheetMimeType = "application/vnd.google-apps.spreadsheet"; // MIME for Spreadsheets
const FormMimeType = "application/vnd.google-apps.form"; // MIME for forms

/* --- ID Cells --- */
const GeneralFolderCell = "C2";
const BoardFolderCell = "C3";
const EventsFolderCell = "C4";
const OfficerFolderCell = "C5";

const PointsSpreadsheetCell = "C7";
const InformationSpreadsheetCell = "C8";

/* --- IDs --- */
const DriveID = "";

const ExampleEventFolderID  = ""; // Event Folder Template
const EventReplacement = "[Sample Event]"; // String to be replaced with the event name

const ExampleServiceYearID = ""; // Service Year Folder Template
const ServiceYearReplacement = "[Service Year]"; // String to be replaced with the service year

const GeneralFolderID = getValue( GeneralFolderCell ); // General Folder ID
const BoardFolderID = getValue( BoardFolderCell ); // Board Folder ID
const EventsFolderID = getValue( EventsFolderCell ); // Events Folder ID
const OfficerFolderID = getValue( OfficerFolderCell ); // Officer Folder ID

const PointSpreadsheetID = getValue( PointsSpreadsheetCell ); // Point Spreadsheet for this year 
const MemberInformationID = getValue( InformationSpreadsheetCell ); // Member Info Spreadsheet for this year

/* --- Other Formatting Variables --- */
// Forms
const DefaultResponseSheetName = "Form Responses 1"; // Default sheet name for form responses
const NewResponseSheetName = "Responses"; // New sheet name for form responses

// Write-Ups
const WriteupSheetName = "Write-Up";
const WriteupTitleText = "Write-Up for ";

// Point Spreadsheet
const Conversions = {
  "Mar": "March",
  "Apr": "April",
  "May": "May",
  "Jun": "June",
  "Jul": "July",
  "Aug": "August",
  "Sep": "September",
  "Oct": "October",
  "Nov": "November",
  "Dec": "December",
  "Jan": "January",
  "Feb": "February"
}
const GeneralSheetName = "General";

const GeneralSheetLast = "A";
const GeneralSheetFirst = "B";
const GeneralSheetGrade = "C";
const GeneralSheetDues = "D";
const GeneralSheetPoints = "E";

const MonthSheetLookup = "A";
const MonthSheetLast = "B";
const MonthSheetFirst = "C";
const MonthSheetPoints = "D";

// Info Spreadsheet
const InformationSheetName = "Summary";

const InformationSheetFirst = "A";
const InformationSheetLast = "B";
const InformationSheetGrade = "C";
const InformationSheetDues = "E";

/* --- Text Styles --- */
const NameText = SpreadsheetApp.newTextStyle()
  .setBold(false)
  .setFontSize(10)
  .build();
const OtherText = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(11)
  .build();

const SectionText = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(14)
  .build();
