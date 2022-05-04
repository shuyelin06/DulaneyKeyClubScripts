/*
Creator: Shu-Ye Joshua Lin (2020-2021 DHS Key Club Membership Secretary)
Collaborators: Julie Lin (2020-2021 DHS Key Club Recording Secretary)

This script was created the summer between the 2019-2020 and 2020-2021 school years (during the COVID-19 pandemic). Please refer to the script documentation document for help.
Created in collaboration with the other Key Club officers and advisors of the 2020-2021 year. 
*/

// Write-up spreadsheet + its UI
const writeup = SpreadsheetApp.openById(writeupID);
const ui = SpreadsheetApp.getUi();

// Specific cells and information that goes into them to create a new write-up from scratch
var writeupFormat = {
  "C6": "Event Name",
  "C8": "Date of Event",
  "C10": "Time of Event",
  "C12": "Chairs",
  "C14": "Location",
  "C16": "Event Description",
  "F6": "Num. of Vol.",
  "F8": "Total Points",
  "F10": "Total Service Hours",
  "F12": "Event ID",
  "F14": "Event Entered",
  "F16": "Email"
}

var volFormat = {
  "C20": "Volunteer #",
  "D20": "Last",
  "E20": "First",
  "F20": "Points",
  "G20": "Hours/Donations"
}

// Custom text for the write-ups
const writeUpText = SpreadsheetApp.newTextStyle()
  .setBold(false)
  .setFontSize(12)
  .setFontFamily("Century Gothic")
  .build();

// Data validation configurations (makes sure all write-ups are consistent and allows for member dropdowns)
const firstVal = SpreadsheetApp.newDataValidation()
  .requireValueInRange(writeup.getSheetByName("Names(Data)").getRange("A2:A"))
  .setHelpText("If a warning appears, check that you have the correct name. If the volunteer has not registered to be a member, they will not show up.")  
  .build();
const lastVal = SpreadsheetApp.newDataValidation()
  .requireValueInRange(writeup.getSheetByName("Names(Data)").getRange("B2:B"))
  .setHelpText("If a warning appears, check that you have the correct name. If the volunteer has not registered to be a member, they will not show up.")
  .build();
const dateVal = SpreadsheetApp.newDataValidation()
  .requireDate()
  .setAllowInvalid(false)
  .setHelpText("Enter a valid date/time.")
  .build();
const pointVal = SpreadsheetApp.newDataValidation()
  .requireNumberBetween(0.5, 5)
  .setAllowInvalid(true)
  .setHelpText("While there will be exceptions, the official point cap is at 5 points max. Please contact the membership secretary for further information")
  .build();

const volunteerCheck = SpreadsheetApp.newDataValidation()
  .requireValueInRange(writeup.getSheetByName("Names(Data)").getRange("D2:D"))
  .setHelpText("Member not found. Please check to make sure you have the correct name.")
  .build();