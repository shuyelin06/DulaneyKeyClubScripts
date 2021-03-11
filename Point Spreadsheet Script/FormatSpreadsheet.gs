/*
This is the first function created for the script, created by Shu-Ye Lin
NOTE: THIS IS NOT INTENDED TO BE USED AT ALL. This function is only present in the event that the entire point spreadsheet is ruined. 

If attempting to clean the spreadsheet for reuse the following year, please refer to the export data function.
*/
function format(){
  var prompt = ui.prompt("Please enter the ID of the spreadsheet you would like to format");
  if(prompt.getSelectedButton() != ui.Button.OK) return;

  var spreadsheet = SpreadsheetApp.openById(prompt.getResponseText());

  createPointForm(spreadsheet);
}

// Creates a Point form
function createPointForm(spreadsheet){
  // Limits the running of this function to the membership secretary's email
  if(Session.getActiveUser().getEmail() != data["Email"]){
    ui.alert("You do not have the permissions to enter events!");
    return;
  }

  var defaultFormat = ["#3c78d8", "#ffffff", "#c9daf8", "#3d85c6"];
  
  var generalFunctions = {
    "A1": "Last Name",
    "B1": "First Name",
    "C1": "Grade",
    "D1": "Dues Paid",
    "E1": "Total Points"
  }

  var sheet = spreadsheet.insertSheet();
  
  sheet.clear();
  sheet.setName(data["Points"]["Sheets"]["Main"]);
  
  alternatingColors(defaultFormat[0], defaultFormat[1], defaultFormat[2], defaultFormat[3], 'A:D', spreadsheet);
  alternatingColors("#cc4125", "#ffffff", "#f4cccc", "#3d85c6", 'E:E', spreadsheet);
  
  // Setting header cells in the general sheet
  for(var key in generalFunctions){
    sheet.getRange(key).setValue(generalFunctions[key]).setTextStyle(sectionText).setHorizontalAlignment("Center");
  }

  // Resizing and formatting
  sheet.autoResizeColumns(1, sheet.getMaxColumns()-1);
  sheet.setFrozenColumns(4);
  sheet.setFrozenRows(1);
  sheet.deleteRows(8, sheet.getMaxRows()-8);
  sheet.deleteColumns(16, sheet.getMaxColumns()-15);

  // Set footer cells in the general sheet
  sheet.getRange("C8").setValue("Total Cumulative").setTextStyle(otherText).setHorizontalAlignment("Left");
  sheet.getRange("E8").setValue('=SUM(INDIRECT("E2:E"&row()-1))').setTextStyle(otherText).setHorizontalAlignment("Center");
  
  // Setting colors for each month point, their header and footer
  for(var key in sheetFormat){
    alternatingColors(sheetFormat[key][0][0], sheetFormat[key][0][1], sheetFormat[key][0][2], sheetFormat[key][0][3], sheetFormat[key][2] + ":" + sheetFormat[key][2], spreadsheet);
    sheet.getRange(sheetFormat[key][2] + "1").setValue(key + " Points").setTextStyle(sectionText).setHorizontalAlignment("Center");
    sheet.getRange(sheetFormat[key][2] + "8").setValue('=SUM(INDIRECT("'+sheetFormat[key][2]+'2:'+sheetFormat[key][2]+'"&row()-1))').setTextStyle(otherText).setHorizontalAlignment("Center");
  }

  sheet.autoResizeColumns(1, sheet.getMaxColumns());
  createSheets(sheetFormat, defaultFormat, spreadsheet);
}

// Creates all of the sheets for every month
function createSheets(format, defaultFormat, spreadsheet){
  for(var key in format){
    var s = spreadsheet.insertSheet(key);

    // Formats the month sheet
    alternatingColors(defaultFormat[0], defaultFormat[1], defaultFormat[2], defaultFormat[3], 'A:D', spreadsheet);
    alternatingColors(format[key][0][0], format[key][0][1], format[key][0][2], format[key][0][3], 'E:E', spreadsheet);
    alternatingColors(format[key][1][0], format[key][1][1], format[key][1][2], format[key][1][3], 'F:I', spreadsheet);

    s.setFrozenColumns(4);
    s.setFrozenRows(1);
    s.deleteColumns(7, s.getMaxColumns()-6);
    s.deleteRows(8, s.getMaxRows()-8);

    // Value adding
    addSheetFunctions(key, spreadsheet);
  }
}

// Add functions to each month sheet
function addSheetFunctions(month, spreadsheet){
  var sheetFunctions = {
    "A1": "VLOOKUP NAME",
    "B1": "Last Name",
    "C1": "First Name",
    "D1": "Grade"
  }

  let sheet = spreadsheet.getActiveSheet();

  // Value setting
  for (var key in sheetFunctions){
    sheet.getRange(key).setValue(sheetFunctions[key]).setTextStyle(sectionText).setHorizontalAlignment("Center");
  }
  sheet.getRange("E1").setValue(month + " Points").setTextStyle(sectionText).setHorizontalAlignment("Center");
  sheet.getRange("E8").setValue('=SUM(INDIRECT(ADDRESS(2, COLUMN(), 4) & ":" &ADDRESS(row()-1, COLUMN(), 4)))').setTextStyle(otherText).setHorizontalAlignment("Center");
  sheet.autoResizeColumns(1, sheet.getMaxColumns());
  sheet.getRange("B8").setValue("Total Cumulative").setTextStyle(otherText).setHorizontalAlignment("Left");
}

// Adds bandings to a given range in a sheet
function alternatingColors(header, firstRow, secondRow, footer, range, spreadsheet){
  spreadsheet.getRange(range).activate().applyRowBanding();

  var banding = spreadsheet.getActiveRange().getBandings()[0];

  banding.setRange(spreadsheet.getRange(range))
    .setHeaderRowColor(header)
    .setFirstRowColor(firstRow)
    .setSecondRowColor(secondRow)
    .setFooterRowColor(footer);
}