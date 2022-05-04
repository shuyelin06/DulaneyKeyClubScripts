/*
  Log the Points from a Write-Up
*/
// Should be assigned to the Log Write-Up Button
function logWriteup() {
  // Obtaining the Write-Up Sheet ID
  var writeupID = UserInterface.prompt("Please Enter the Write-Up Sheet ID", UserInterface.ButtonSet.OK_CANCEL);
  if(writeupID.getSelectedButton() != UserInterface.Button.OK) return;
  else {
    writeupID = writeupID.getResponseText();
    console.log("Writeup ID: " + writeupID);
  }

  // Get the Write-Up Sheet
  var writeup = SpreadsheetApp.openById( writeupID ).getSheetByName( WriteupSheetName );
  var eventName = getWriteupName(writeup);
  console.log("Event Name: " + eventName);
  
  // Perform Basic Checks
  if( !canEnter(writeup) ) {
    console.log("Failed completion check");
    UserInterface.alert("Write-up failed completion check - cannot be entered");
    return;
  }

  // Get Appropriate Sheet in Point Spreadsheet
  var month = getEventMonth(writeup);
  var pointSheet = SpreadsheetApp.openById( PointSpreadsheetID ).getSheetByName( month );
  console.log("Event Month: " + month);

  // Create New Column for Event
  var col = pointSheet.getMaxColumns(); // Find Last Column
  pointSheet.insertColumnBefore(col); // Insert Column Before Last Column
  pointSheet.getRange(1, col) // Set Event Name
    .setValue(eventName)
    .setTextStyle(SectionText)
    .setHorizontalAlignment("center"); 
  pointSheet.autoResizeColumn(col);
  console.log("Event Column Created");

  // Log Points
  var data = writeup.getRange("C24:E").getValues();
  for(var i = 0; i < data.length; i++) {
    var firstName = data[i][0];
    var lastName = data[i][1];
    var points = parseInt(data[i][2]);

    // Basic Validation Check
    if( firstName === "" || lastName === ""  || isNaN(points)) continue;
    // Log Member
    else {
      var memberRow = memberSearch(firstName, lastName, MonthSheetFirst, MonthSheetLast, pointSheet);
      if(memberRow == -1) {
        memberRow = insertMember(firstName, lastName)[month];
      }

      pointSheet.getRange(memberRow, col).setValue(points);
    }
  }

  // Update Write-Up
  writeup.getRange("G11").setValue("Yes");

  // Sort Data
  sortData();
}

function getWriteupName(writeup) {
  const WriteupTitle = "E3";

  var title = writeup.getRange(WriteupTitle).getValue();
  return title.replaceAll( WriteupTitleText , "");
}

function canEnter(writeup) { // WIP
  const EventEntered = "G11";
  const Complete = "F17";
  const Volunteers = "G5";
  const EmptyCheck = [
    "D5", "D7", "D9", "D11", "D13", "D15", "G13"
  ]
  
  // Complete Check
  if(writeup.getRange(Complete).getValue() != "TRUE") {
    UserInterface.alert("Complete not checked");
    return false;
  }

  // Event Entered Check
  if(writeup.getRange(EventEntered).getValue() == "Yes") {
    UserInterface.alert("Event already entered");
    return false;
  }

  // Empty Check
  for(var i = 0; i < EmptyCheck.length; i++) {
    if( removeWhiteSpace(writeup.getRange(EmptyCheck[i]).getValue()) == "" ) {
      UserInterface.alert("There is an empty required cell");
      return false;
    }
  }

  // Volunteer Check 
  if(writeup.getRange(Volunteers).getValue() == 0) {
    UserInterface.alert("There are no volunteers");
    return false;
  }

  return true;
}

function getEventMonth(writeup) {
  const WriteupDate = "D5";

  var date = writeup.getRange( WriteupDate ).getValue().toString().split(" ")[1];
  for(var key in Conversions){
    if(date === key) return Conversions[key];
  }
  return null;
}