// Should be assigned to the New Member (Manual) button
function newMember() {
  var firstName = UserInterface.prompt("Enter the Member's First Name", UserInterface.ButtonSet.OK_CANCEL);
  if( firstName.getSelectedButton() != UserInterface.Button.OK ) {
    UserInterface.alert("No first name given"); return;
  } else firstName = firstName.getResponseText();

  var lastName = UserInterface.prompt("Enter the Member's Last Name", UserInterface.ButtonSet.OK_CANCEL);
  if( lastName.getSelectedButton() != UserInterface.Button.OK ) {
    UserInterface.alert("No last name given"); return;
  } else lastName = lastName.getResponseText();

  insertMember(firstName, lastName);
}

// Inserts a Member
function insertMember(firstName, lastName){
  var rowPositions = {}; // Output of the Function

  const PointSpreadsheet = SpreadsheetApp.openById( PointSpreadsheetID );
  var generalSheet = PointSpreadsheet.getSheetByName(GeneralSheetName);

  // Setting Values in General Sheet
  var row = openPosition(generalSheet, GeneralSheetFirst, GeneralSheetLast);
  rowPositions["General"] = row;

  generalSheet.getRange(GeneralSheetFirst + row)
    .setValue(firstName)
    .setTextStyle(NameText)
    .setHorizontalAlignment("left"); // First Name
  generalSheet.getRange(GeneralSheetLast + row)
    .setValue(lastName)
    .setTextStyle(NameText)
    .setHorizontalAlignment("left"); // Last Name
  generalSheet.getRange(GeneralSheetPoints + row)
    .setValue('=SUM($F' + row + ':$O' + row + ')')
    .setTextStyle(OtherText)
    .setHorizontalAlignment("center"); // Total Points

  var i = 1;
  for(var key in Conversions){ // VLOOKUP Functions
    var month = Conversions[key];

    generalSheet.getRange(GeneralSheetPoints + row).offset(0, i)
      .setValue('=VLOOKUP(TRIM($A$' + row + ')&", "&TRIM($B$' + row + '), ' + month + '!A:D, 4, FALSE)')
      .setTextStyle(OtherText)
      .setHorizontalAlignment("center");

    i++;
  }

  // Setting Values in Month Sheets
  for(var key in Conversions) {
    var month = Conversions[key];
    var monthSheet = PointSpreadsheet.getSheetByName( month );

    row = openPosition(monthSheet, MonthSheetFirst, MonthSheetLast);
    rowPositions[month] = row;

    monthSheet.getRange(MonthSheetLookup + row) // Lookup
      .setValue('=TRIM(B' + row + ')&", "&TRIM(C' + row + ')') 
      .setTextStyle(OtherText)
      .setHorizontalAlignment("left");
    monthSheet.getRange(MonthSheetLast + row) // Last Name
      .setValue(lastName)
      .setTextStyle(NameText)
      .setHorizontalAlignment("left"); 
    monthSheet.getRange(MonthSheetFirst + row) // First Name
      .setValue(firstName)
      .setTextStyle(NameText)
      .setHorizontalAlignment("left");
    monthSheet.getRange(MonthSheetPoints + row) // Point Sum
      .setValue('=SUM(E' + row + ':' + row + ')')
      .setTextStyle(OtherText)
      .setHorizontalAlignment("center");
  }

  // Return Row Positions
  return rowPositions;
}

// Sorts the entire range of members with their points into ascending order
// Should be assigned to the Sort Data (Manual) button
function sortData(){
  const PointSpreadsheet = SpreadsheetApp.openById( PointSpreadsheetID );

  var sheets = PointSpreadsheet.getSheets();
  for(var i = 0; i < sheets.length; i++){
    var sheet = sheets[i];
    sheet.getRange(2, 1, sheet.getMaxRows()-2, sheet.getMaxColumns()).sort({column: 1, ascending: true});
  }
}

// Find an open row, or create one
function openPosition(sheet, firstCol, lastCol){
  for( var i=1; i < sheet.getMaxRows(); i++ ){
    if(sheet.getRange(firstCol + i).getValue() == "" && sheet.getRange(lastCol + i).getValue() == "") return i;
  }

  var rowNum = sheet.getMaxRows();
  sheet.insertRowBefore(rowNum);
  return rowNum;
}