/*
Creator: Shu-Ye Joshua Lin (2020-2021 DHS Key Club Membership Secretary)
Collaborators: Julie Lin (2020-2021 DHS Key Club Recording Secretary)

This script was created the summer between the 2019-2020 and 2020-2021 school years (during the COVID-19 pandemic). Please refer to the script documentation document for help.
Created in collaboration with the other Key Club officers and advisors of the 2020-2021 year. 

NOTE: All resources created by the '20-'21 Membership Secretary was created on this format. While you have freedom to do as you please,
I implore you to not change the format itself, as it may render some resources unuseable.
*/
// NOTE: The onOpen trigger may not run on adblock.

var data = {
  // Email of the current membership secretary
  "Email": "shuyelin06@gmail.com",
  "ClearKey": "ABCD",
  // Information for the point spreadsheet
  "Points": {
    "Sheets": {
      "Main": "General",
      "Month": {
        "LastName": "B",
        "FirstName": "C"
      }
    },
    // Columns for the main sheet
    "Columns": {
      "LastName": "A",
      "FirstName": "B",
      "Grade": "C",
      "Dues": "D",
      "TotalPoints": "E"
    }
  },
  // Information for the writeup spreadsheet
  "WriteUps": {
    "Cells": {
      "Name": "D6",
      "Date": "D8",
      "Chairs": "D12",
      "VolunteerNumber": "G6",
      "EventID": "G12",
      "Entered": "G14",
      "ChairEmail": "G16",
      // Volunteer cells should begin on the first volunteer row (not the header)
      "VolunteerLast": "D21",
      "VolunteerFirst": "E21",
      "VolunteerPoints": "F21"
    },
    "Columns": {
      "LastName": "D",
      "FirstName": "E"
    }
  },
  // Information for the member info spreadsheet
  "MemberInfo": {
    // Variable for how dues are shown on the information spreadsheet (if paid, they are marked with the word __)
    "DuesValue": "Yes",
    "Sheets": {
      "Main": "Member Information",
      "Form": "Form Data"
    },
    "Columns": {
      "Email": "A",
      "FirstName": "B",
      "LastName": "C",
      "Grade": "D",
      "Dues": "I"
    }
  }
}

// Defining the point, writeup and membership info spreadsheet, respectively
const pointSpreadsheet = SpreadsheetApp.openById(pointsID);
const writeupSpreadsheet = SpreadsheetApp.openById(writeupID);
const memberInfoSpreadsheet = SpreadsheetApp.openById(memberInfoID);

// UI of the point spreadsheet
var ui = SpreadsheetApp.getUi();

// Creates the menu to interact with
function onOpen(){
  ui.createMenu('Point Management')
    .addSubMenu(ui.createMenu("Members")
      .addItem('Enter a member', 'inputMemberInfo')
                .addSeparator()
      .addItem('Sync Member Information', 'syncInfo')
                .addSeparator()
      .addItem('Sort Data', 'sortData')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu("DANGER ZONE")
      .addItem('Clear Spreadsheet', 'clean')
    )  
    .addSeparator()
    .addItem('Enter a WriteUp', 'enterWriteup')  
    .addSeparator()
    .addItem('Format a Point Form', 'format')
    .addToUi();
}

// Data for formatting (used in the member functions and spreadsheet formatting)
var sheetFormat = {
  "March": [["#ff0000", "#ffffff", "#fddce8", "#3d85c6"],["#e06666", "#ffffff", "#f4cccc", "#3d85c6"], 'F'],
  "April": [["#ff5520", "#ffffff", "#ffc0ad", "#3d85c6"],["#ff7a51", "#ffffff", "#fff2cc", "#3d85c6"], 'G'],
  "May":[["#ff9900", "#ffffff", "#fce5cd", "#3d85c6"],["#f1c232", "#ffffff", "#fff2cc", "#3d85c6"],'H'],
  "June": [["#f1c232", "#ffffff", "#fff2cc", "#3d85c6"],["#ffd966", "#ffffff", "#fff2cc", "#3d85c6"], 'I'],
  "September": [["#8bc34a", "#ffffff", "#eef7e3", "#3d85c6"],["#93c47d", "#ffffff", "#d9ead3", "#3d85c6"], 'J'],
  "October": [["#76a5af", "#ffffff", "#d0e0e3", "#3d85c6"],["#a2c4c9", "#ffffff", "#d0e0e3", "#3d85c6"], 'K'],
  "November": [["#3c78d8", "#ffffff", "#c9daf8", "#3d85c6"],["#a4c2f4", "#ffffff", "#c9daf8", "#3d85c6"], 'L'],
  "December": [["#8989eb", "#ffffff", "#e8e7fc", "#3d85c6"],["#b4a7d6", "#ffffff", "#d9d2e9", "#3d85c6"], 'M'],
  "January": [["#b4a7d6", "#ffffff", "#e8e7fc", "#3d85c6"],["#d9d2e9", "#ffffff", "#eae6f3", "#3d85c6"], 'N'],
  "February": [["#c27ba0", "#ffffff", "#ead1dc", "#3d85c6"],["#d5a6bd", "#ffffff", "#f0e0e7", "#3d85c6"], 'O']
}

// Text Styles
const sectionText = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(14)
  .build();

const namesText = SpreadsheetApp.newTextStyle()
  .setBold(false)
  .setFontSize(10)
  .build();

const otherText = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontSize(11)
  .build();