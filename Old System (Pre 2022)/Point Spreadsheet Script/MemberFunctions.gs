// Uses the UI to input one individual (using the writeups will do this automatically)
function inputMemberInfo(){
  let last = ui.prompt("Enter the Member's Last Name", ui.ButtonSet.OK_CANCEL);
  if(last.getSelectedButton() == ui.Button.CANCEL || last.getSelectedButton() == ui.Button.CLOSE) return;
  let first = ui.prompt("Enter the Member's First Name", ui.ButtonSet.OK_CANCEL);
  if(first.getSelectedButton() == ui.Button.CANCEL || first.getSelectedButton() == ui.Button.CLOSE) return;
  let grade = ui.prompt("Enter the Member's Grade Level", ui.ButtonSet.OK_CANCEL);
  if(grade.getSelectedButton() == ui.Button.CANCEL || grade.getSelectedButton() == ui.Button.CLOSE) return;
  let dues = ui.prompt("Has the member paid dues? (X if yes, blank if no)", ui.ButtonSet.OK_CANCEL);
  if(dues.getSelectedButton() == ui.Button.CANCEL || dues.getSelectedButton() == ui.Button.CLOSE) return;
  Member(last.getResponseText(), first.getResponseText(), grade.getResponseText(), dues.getResponseText(), openPosition());
  sortData();
}

// Inserts a member
function Member(Last, First, Grade, Dues, Row){
  let generalSheet = pointSpreadsheet.getSheetByName(data["Points"]["Sheets"]["Main"]);
  
  // Setting values for a member in the General Sheet
  generalSheet.getRange(data["Points"]["Columns"]["LastName"] + Row).setValue(Last).setTextStyle(namesText).setHorizontalAlignment("Left");
  generalSheet.getRange(data["Points"]["Columns"]["FirstName"] + Row).setValue(First).setTextStyle(namesText).setHorizontalAlignment("Left");
  generalSheet.getRange(data["Points"]["Columns"]["Grade"] + Row).setValue(Grade).setTextStyle(otherText).setHorizontalAlignment("Center");
  generalSheet.getRange(data["Points"]["Columns"]["Dues"] + Row).setValue(Dues).setTextStyle(otherText).setHorizontalAlignment("Center");
  generalSheet.getRange(data["Points"]["Columns"]["TotalPoints"] + Row).setValue('=SUM(F' + Row + ':O' + Row + ')').setTextStyle(otherText).setHorizontalAlignment("Center");
  
  // Create VLOOKUP functions for each month in the general sheet
  for(var key in sheetFormat){
    generalSheet.getRange(sheetFormat[key][2]+Row).setValue('=VLOOKUP(TRIM(A' + Row + ')&", "&TRIM(B' + Row + '), ' + key + '!A:E, 5, FALSE)').setTextStyle(otherText).setHorizontalAlignment("Center");
  }
  
  // Sets values for each month sheet
  for(var key in sheetFormat){
    let monthSheet = pointSpreadsheet.getSheetByName(key);
    monthSheet.getRange("A" + Row).setValue('=TRIM(B' + Row + ')&", "&TRIM(C' + Row + ')').setTextStyle(otherText).setHorizontalAlignment("Left");
    monthSheet.getRange("B" + Row).setValue(Last).setTextStyle(namesText).setHorizontalAlignment("Left");
    monthSheet.getRange("C" + Row).setValue(First).setTextStyle(namesText).setHorizontalAlignment("Left");
    monthSheet.getRange("D" + Row).setValue(Grade).setTextStyle(otherText).setHorizontalAlignment("Center");
    monthSheet.getRange("E" + Row).setValue('=SUM(F' + Row + ':' + Row + ')').setTextStyle(otherText).setHorizontalAlignment("Center");
  }

  sortData();
}

// In the General Sheet, searches for an empty row to insert a member into.
function openPosition(){
  let sheet = pointSpreadsheet.getSheetByName(data["Points"]["Sheets"]["Main"]).activate();
  for(var i=1; i<sheet.getMaxRows(); i++){
    if(sheet.getRange(data["Points"]["Columns"]["LastName"]+i).getValue() == "" || sheet.getRange(data["Points"]["Columns"]["LastName"]+i).getValue() == " ") return i;
  }
  let rowNum = sheet.getMaxRows();
  sheet.insertRowBefore(rowNum);
  for(var key in sheetFormat){
    let temp = pointSpreadsheet.getSheetByName(key);
    temp.insertRowBefore(rowNum)
  }
  return rowNum;
}

// Sorts the entire range of members with their points into ascending order
function sortData(){
  var sheets = pointSpreadsheet.getSheets();

  for(var i=0; i<sheets.length; i++){
    var sheet = sheets[i];

    sheet.getRange(2, 1, sheet.getMaxRows()-2, sheet.getMaxColumns()).sort({column: 1, ascending: true});
  }
}