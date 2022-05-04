/*
  Scripts created by Josh Lin

  The point of the scripts is to streamline officer duties by automating 
  many tasks that officers face. The scripts were created at the end of the '21-'22 Service Year.
*/
// Main Variables
const OfficerCenter = SpreadsheetApp.getActiveSpreadsheet();
const UserInterface = SpreadsheetApp.getUi();

// Ran Upon Document Open
function onOpen() {
  createUI();
}

// Create User Interface
function createUI() {
  UserInterface.createMenu("Scripts");
}
