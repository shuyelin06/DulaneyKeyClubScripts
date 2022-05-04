/** @OnlyCurrentDoc */

function NewVolunteer() {
  const FirstMemberRow = 24;

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  sheet.insertRowsBefore( FirstMemberRow, 1);
}