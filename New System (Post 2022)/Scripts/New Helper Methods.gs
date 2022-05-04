/*
  Helper Methods to create/copy new files and folders
*/

// Recursive Method to Copy copyFolder into destinationFolder
function copyFiles(copyFolder, destinationFolder) {
  // New Folder to Insert Files Into
  console.log("Creating New Folder: " + copyFolder.getName());
  var newFolder = destinationFolder.createFolder( copyFolder.getName() );

  // Copy Files in this Folder
  var files = copyFolder.getFiles();
  while( files.hasNext() ) {
    var file = files.next();

    // Does not copy spreadsheets with forms attached to them
    if( file.getMimeType() == SpreadsheetMimeType ) {
      if( SpreadsheetApp.openById( file.getId()).getFormUrl() != null ) continue;
    } 
    // Does not copy forms with spreadsheet destinations
    else if( file.getMimeType() == FormMimeType ) {
      try{
        FormApp.openById( file.getId() ).getDestinationId();
        continue;
      } catch(exception) {}
    }

    console.log("Copying " + file.getName());
    file.makeCopy( file.getName(), newFolder );
  }

  // Call Method on Folders in this Folder
  var folders = copyFolder.getFolders();
  while( folders.hasNext() ) {
    var folder = folders.next();
    copyFiles(folder, newFolder);
  }

  // Return NewFolder
  return newFolder;
}

// Replace the given targetString with a replacementString for all files in a given folder
function modifyNames(sourceFolder, targetString, replacementString) {
   // Modify Source Folder
  sourceFolder.setName( modifyName(sourceFolder.getName(), targetString, replacementString) );

  // Modify Files in Folder
  var files = sourceFolder.getFiles();
  while( files.hasNext() ) {
    var file = files.next();
    file.setName( modifyName(file.getName(), targetString, replacementString) );
  }

  // Modify Folders in Folder
  var folders = sourceFolder.getFolders();
  while( folders.hasNext() ) {
    var folder = folders.next();
    modifyNames(folder, targetString, replacementString);
  }
}
function modifyName(fileName, targetString, replacementString) { // Returns a Modified File Name
  var newName = fileName.replaceAll(targetString, replacementString);
  return newName;
}