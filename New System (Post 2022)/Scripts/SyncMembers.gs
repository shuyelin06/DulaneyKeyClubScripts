// Sync information spreadsheet info with point sheet using a binary search algorithm
// Should be assigned to the Sync Member Information button
function syncInfo(){
  const PointSheet = SpreadsheetApp.openById( PointSpreadsheetID ).getSheetByName( GeneralSheetName );
  const InfoSheet = SpreadsheetApp.openById( MemberInformationID ).getSheetByName( InformationSheetName );
  
  // Counter to track how many members were synced
  var n = 0;
  
  for( var i = 2; i < PointSheet.getMaxRows(); i++ ){
    // Pulling the first and last name from the point spreadsheet
    const FirstName = PointSheet.getRange(GeneralSheetFirst + i).getValue();
    const LastName = PointSheet.getRange(GeneralSheetLast + i).getValue();
    
    // Search for member row in the info sheet
    var memberRow = memberSearch(FirstName, LastName, InformationSheetFirst, InformationSheetLast, InfoSheet);

    if(memberRow == -1) continue;

    // Pulling the grade and dues data from the member info spreadsheet
    const Grade = InfoSheet.getRange(InformationSheetGrade + memberRow).getValue().split(" ")[0];
    const Dues = InfoSheet.getRange(InformationSheetDues + memberRow).getValue();
    
    // Pulling the grade and dues from the point spreadsheet
    const FoundGrade = PointSheet.getRange(GeneralSheetGrade + i).getValue();
    const FoundDues = PointSheet.getRange(GeneralSheetDues + i).getValue();
    
    // Check if the grade and dues from the two spreadsheets match, if not change the values on the point spreadsheet
    if(Grade != FoundGrade || Dues != FoundDues){
      UserInterface.alert(FirstName + " " + LastName + "\n" + 
                "Displayed Dues: " + FoundDues + "    --    " + "Current Dues: " + Dues + "\n" + 
                "Displayed Grade: " + FoundGrade + "    --    " + "Current Grade: " + Grade);
      
      // Setting values in the point spreadsheet for grade and dues accordingly
      PointSheet.getRange(GeneralSheetGrade + i)
        .setValue(Grade)
        .setTextStyle(OtherText)
        .setHorizontalAlignment("center");
      PointSheet.getRange(GeneralSheetDues + i)
        .setValue(Dues)
        .setTextStyle(OtherText)
        .setHorizontalAlignment("center");
      
      // Incrementing n to track how many members were synced
      n++;
    }
    
  }
  
  // Informational panels to show what script did
  if(n == 0) UserInterface.alert("No member information changed.");
  else UserInterface.alert(n + " members have been synced.");
}