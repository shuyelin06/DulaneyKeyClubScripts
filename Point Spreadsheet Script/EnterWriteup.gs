/*
This function was created by Shu-Ye Lin near the beginning of the '20-'21 school year.
This function dramatically simplifies the process in entering write-ups to the point spreadsheet.

Intended to work with the ACCOMPANYING write-up spreadsheet (format developed by Julie Lin, our wonderful Recording Secretary). 
Some variables will be defined in the chance that something wrong occurs. 
*/
// Enters in the data from a given writeup sheet (using its ID)
function enterWriteup(){
  // Limits the running of this function to the membership secretary's email
  if(Session.getActiveUser().getEmail() != data["Email"]){
    ui.alert("You do not have the permissions to enter events!");
    return;
  }
  
  // Identifying the write-up sheet based on its ID
  let response = ui.prompt("Please enter the sheet ID", ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton() != ui.Button.OK) return;
  
  // Converting id extracted from user response of type String to type Int
  const id = parseInt(response.getResponseText(), 10);

  // Searching for writeup in the spreadsheet
  var sheet;
  for(var i in writeupSpreadsheet.getSheets()){
    if(writeupSpreadsheet.getSheets()[i].getRange(data["WriteUps"]["Cells"]["EventID"]).getValue() == id) sheet = writeupSpreadsheet.getSheets()[i];
  }

  // Failsafes
  if (sheet == null){
    ui.alert("Error: Write-Up could not be found");
    return;
  } 
  else if(sheet.getRange(data["WriteUps"]["Cells"]["Entered"]).getValue() == "Yes"){
    ui.alert("Error: Write-up has already been entered");
    return;
  } 
  else {
    // Checks if all fields are filled out and the write-up is locked/protected
    let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    
    if(finishCheck(sheet) || protection.length <= 0){
      ui.alert("Error: Write-Up has not yet been finished.");
      return;
    }
  }

  // Pulling the month and name from the write-up sheet
  const month = dateConversion(
    sheet.getRange(data["WriteUps"]["Cells"]["Date"])
    .getValue()
    .toString()
    .split(" ")[1]
  );
  const name = sheet.getRange(data["WriteUps"]["Cells"]["Name"]).getValue();
  
  // Create an event using the createEvent() function
  createEvent(month, name);
  
  // Checking if the member in the write-up exists. If not, adds a member.
  for(var i=0; i<sheet.getRange(data["WriteUps"]["Cells"]["VolunteerNumber"]).getValue(); i++){
    // Pulling the last name, first name and points from one volunteer
    const last = sheet.getRange(data["WriteUps"]["Cells"]["VolunteerLast"]).offset(i, 0).getValue();
    const first = sheet.getRange(data["WriteUps"]["Cells"]["VolunteerFirst"]).offset(i, 0).getValue();

    var index =  binarySearch(first, last, data["Points"]["Columns"]["FirstName"], data["Points"]["Columns"]["LastName"], pointSpreadsheet.getSheetByName(data["Points"]["Sheets"]["Main"]));

    if(index == -1){
      Member(last, first, "", "", openPosition());
    }
  }

  // Gives the given member points
  for(var i=0; i<sheet.getRange(data["WriteUps"]["Cells"]["VolunteerNumber"]).getValue(); i++){
    const last = sheet.getRange(data["WriteUps"]["Cells"]["VolunteerLast"]).offset(i, 0).getValue();
    const first = sheet.getRange(data["WriteUps"]["Cells"]["VolunteerFirst"]).offset(i, 0).getValue();
    const points = sheet.getRange(data["WriteUps"]["Cells"]["VolunteerPoints"]).offset(i, 0).getValue();

    if(last == "" || first == "") continue;

    let rowNumber = binarySearch(first,last,data["Points"]["Sheets"]["Month"]["FirstName"],data["Points"]["Sheets"]["Month"]["LastName"],pointSpreadsheet.getSheetByName(month));

    enterInPoints(rowNumber, points, month);
  }

  // Setting the entered value to yes
  sheet.getRange(data["WriteUps"]["Cells"]["Entered"]).setValue("Yes");
  
  //Contacting the event chair and (optional) members
  contactChair(sheet.getRange(data["WriteUps"]["Cells"]["ChairEmail"]).getValue(), sheet.getRange(data["WriteUps"]["Cells"]["Name"]).getValue());
  
  var contactCheck = ui.alert("Do you want to notify the members about this event being entered?", ui.ButtonSet.YES_NO);
  if(contactCheck == ui.Button.YES) contactMembers(sheet);
}

// Check if a given writeup sheet is complete or not
function finishCheck(sheet){
  for(var key in data["WriteUps"]["Cells"]){
    if(sheet.getRange(data["WriteUps"]["Cells"][key]).isBlank()) return true;
  }
  return false;
}

// Reads the month from the write-up date.
function dateConversion(monthNum){
  var conversion = {
    "Jan": "January",
    "Feb": "February",
    "Mar": "March",
    "Apr": "April",
    "May": "May",
    "Jun": "June",
    "Jul": "June",
    "Aug": "September",
    "Sep": "September",
    "Oct": "October",
    "Nov": "November",
    "Dec": "December"
  }
  for(var key in conversion){
    if(monthNum === key) return conversion[key];
  }
}

// Creates a new event in the point spreadsheet
function createEvent(Month, eventName){
  var sheet = pointSpreadsheet.getSheetByName(Month);
  var lastCol = sheet.getMaxColumns();
  sheet.insertColumnAfter(lastCol);
  
  // Converts the number to the corresponding letter (to be used in A1 notation)
  const colChar = String.fromCharCode(64+lastCol);
  
  // Value setting for the event
  sheet.getRange(colChar+"1").setValue(eventName).setTextStyle(otherText).setHorizontalAlignment("Center");
  sheet.getRange(colChar+sheet.getMaxRows()).setValue('=SUM(INDIRECT(ADDRESS(2, COLUMN(), 4) & ":" &ADDRESS(row()-1, COLUMN(), 4)))').setTextStyle(otherText).setHorizontalAlignment("Center");
  sheet.autoResizeColumn(lastCol);
}

// Enters the points for a given row in a given month. *NOTE*: Intended to enter in points for the second-to-last column. Please do not change the event placement
function enterInPoints(Row, Points, Month){
  var sheet = pointSpreadsheet.getSheetByName(Month);
  let colChar = String.fromCharCode(64+sheet.getLastColumn());

  console.log(colChar+Row);

  sheet.getRange(colChar+Row).setValue(Points).setTextStyle(otherText).setHorizontalAlignment("Center");
}