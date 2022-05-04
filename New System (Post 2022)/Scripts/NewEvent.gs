/*
  Creates a New Event Folder
*/
// Should be assigned to the Create New Event button
function newEvent() {
  // Constant Variable Creation
  const EventsFolder = DriveApp.getFolderById(EventsFolderID); // Folder of all the events
  const SampleEventFolder = DriveApp.getFolderById(ExampleEventFolderID); // Event Template Folder

  // Prompt User for Event Name
  var prompt = UserInterface.prompt("Please Enter the Name of the Event");
  var eventName = EventReplacement;
  if(prompt.getSelectedButton() != UserInterface.Button.OK ) {
    UserInterface.alert("No Event Name Provided");
    return;
  } else {
    eventName = prompt.getResponseText();
  }

  // Copy Files to EventFolder
  var newEventFolder = copyFiles(SampleEventFolder, EventsFolder);

  // Modify the Names
  modifyNames(newEventFolder, EventReplacement, eventName);

  // Format the Write-Up
  var writeupName = modifyName("[Sample Event] Write-Up", EventReplacement, eventName);
  formatWriteup( SpreadsheetApp.openById(findFile(newEventFolder, writeupName).getId()), eventName);

  // Success Message
  UserInterface.alert("Success!");
}

function findFile(folder, name) {
  var files = folder.getFiles();
  while( files.hasNext() ) {
    var file = files.next();

    if( file.getName() == name ) return file;
  }

  var folders = folder.getFolders();
  while( folders.hasNext() ) {
    var folder = folders.next();

    return findFile(folder, name);
  } 
  
}
function formatWriteup(writeupFile, eventName) {
  console.log("formatting");
  const SummarySheetName = "Write-Up";
  const MemberSheetName = "Members";

  var writeupSheet = writeupFile.getSheetByName( SummarySheetName );
  var memberSheet = writeupFile.getSheetByName( MemberSheetName );

  // Update Write-Up Title
  const WriteupTitle = "E3";
  writeupSheet.getRange( WriteupTitle).setValue("Write-Up for " + eventName);

  // Update MemberList
  const InformationRange = "A2:B";
  const InformationSheet = "Summary";

  const MemberFirstCol = "A"; // First Name Column, A
  const MemberLastCol  = "B"; // Last Name Column, B
  const FirstMemberRow = "3"; // First Member, Row 3

  var informationSheet = SpreadsheetApp.openById( MemberInformationID ).getSheetByName( InformationSheet );
  var information = informationSheet.getRange( InformationRange + informationSheet.getLastRow() ).getValues();

  for( var i = 0; i < information.length; i++ ) {
    for(var j = 0; j < information[0].length; j++ ) {
      memberSheet.getRange(MemberFirstCol + FirstMemberRow).offset(i, j).setValue( information[i][j] );
    }
  }
}