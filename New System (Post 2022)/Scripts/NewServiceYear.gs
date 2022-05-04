/*
  Copies all of the folders / files for a new service year
  Copying includes all files, and all forms along with their spreadsheet destinations

  EXCEPTIONS: Functions used in sheets that reference external sheets are not accounted for 
*/
// Should be assigned to the "New Service Year" Button
function newServiceYear() {
  // Constant Variable Creation
  const Drive = DriveApp.getFolderById(DriveID);
  const SampleServiceYear = DriveApp.getFolderById(ExampleServiceYearID);

  // Prompt User for Event Name
  var prompt = UserInterface.prompt("Please Enter the Service Year \n (ex. '22-'23)");
  var serviceYear = "Service Year";
  if(prompt.getSelectedButton() != UserInterface.Button.OK ) {
    UserInterface.alert("No Service Year Provided");
    return;
  } else {
    serviceYear = prompt.getResponseText();
  }

  // Copy files from the sample service year to the drive
  console.log(" --- Copying Folder --- ");
  var newServiceFolder = copyFiles(SampleServiceYear, Drive);
  console.log(" --- Finished Copying --- ");

  // Modify the names of the files
  console.log(" --- Modifying Names --- ");
  modifyNames(newServiceFolder, ServiceYearReplacement, serviceYear);
  console.log(" --- Finished Modifying --- ");

  // Link spreadsheets and forms
  console.log(" --- Linking --- ");
  linkFiles(newServiceFolder, serviceYear);
  console.log(" --- Finished Linking --- ");
}

// Link spreadsheets and forms together
function linkFiles(newServiceYear, serviceYear) {
  const SampleServiceYear = DriveApp.getFolderById(ExampleServiceYearID);

  // Find all forms in template folder
  console.log("Finding Forms");
  var forms = getForms( SampleServiceYear, [] );

  // Find all forms that have a response destination
  console.log("Filtering for Linked Forms");
  var links = getLinks(forms);
  
  // For every linked spreadsheet, copy it under a modified name
  console.log("Copying Linked Spreadsheets");
  const FormIndex = 0;
  const SpreadsheetIndex = 1;
  var newFiles = [];
  for(var i = 0; i < links.length; i++) {
    var file = links[i][SpreadsheetIndex];

    var newSpreadsheet = file.makeCopy( modifyName(file.getName(), ServiceYearReplacement, serviceYear) );
    var newForm = DriveApp.getFileById( FormApp.openByUrl( SpreadsheetApp.openById(newSpreadsheet.getId()).getFormUrl() ).getId() );
    newForm.setName( (newForm.getName().replaceAll("Copy of ", "")).replaceAll(ServiceYearReplacement, serviceYear) );

    newFiles.push( [ links[i][FormIndex].getParents().next().getName(), newForm ] );
    newFiles.push( [ links[i][SpreadsheetIndex].getParents().next().getName(), newSpreadsheet ] );
  }

  // Move Files
  console.log("Moving Forms and Spreadsheets");
  moveFiles(newServiceYear, newFiles);
}

function getForms(sourceFolder, forms) { // Find all forms in a given source folder and associated sub-folders
  var files = sourceFolder.getFilesByType( FormMimeType );
  while( files.hasNext() ) {
    var file = files.next();
    forms.push( file.getId() );
  }

  var folders = sourceFolder.getFolders();
  while( folders.hasNext() ) {
    getForms( folders.next(), forms );
  }

  return forms;
}
function getLinks(forms) { // Find all forms that are linked to a spreadsheet
  var links = [];
  
  for(var i = 0; i < forms.length; i++) {
    var formFile = DriveApp.getFileById( forms[i] );
    var form = FormApp.openById( forms[i] );

    try {
      var destinationFile = DriveApp.getFileById( form.getDestinationId() );
      links.push( [ formFile, destinationFile ] );
    } catch(exception) {
      console.log("Form does not have a destination");
    }
  }

  return links;
}
function moveFiles(sourceFolder, newFiles) { // Find all files associated with the link names previously found
  const FolderIndex = 0;
  const FileIndex = 1;
  for(var i = 0; i < newFiles.length; i++) {
    if(newFiles[i][FolderIndex] == sourceFolder.getName()) {
      console.log("File Successfully Moved");
      newFiles[i][FileIndex].moveTo(sourceFolder);
    }
  }

  var folders = sourceFolder.getFolders();
  while( folders.hasNext() ) {
    var folder = folders.next();
    moveFiles(folder, newFiles);
  }  
}