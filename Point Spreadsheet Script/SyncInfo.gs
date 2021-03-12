/*
This function was created and completed on November 30, 2020 by Shu-Ye Joshua Lin. Documentation as of now has not been added.

This function takes the point and member info spreadsheets and syncs the information, so that the point spreadsheet displays correct grades and dues.
Many variables are defined in the syncInfo() function, for easy access. In the event column numbers or data values change, you may change them accordingly.
*/

// Sync information spreadsheet info with point sheet using a binary search algorithm
function syncInfo(){
  // Limits the running of this function to the membership secretary's email
  if(Session.getActiveUser().getEmail() != memSec){
    ui.alert("You do not have the permissions to sync information!");
    return;
  }
  
  var pointSheet = pointSpreadsheet.getSheetByName(data["Points"]["Sheets"]["Main"]);
  var infoSheet = memberInfoSpreadsheet.getSheetByName(data["MemberInfo"]["Sheets"]["Main"]);
  
  // Counter to track how many members were synced
  let n = 0;
  
  for(var i=2; i<pointSheet.getMaxRows(); i++){
    // Pulling the first and last name from the point spreadsheet
    const firstName = pointSheet.getRange(data["Points"]["Columns"]["FirstName"] + i).getValue();
    const lastName = pointSheet.getRange(data["Points"]["Columns"]["LastName"] + i).getValue();
    
    // Binary Search will return what the member row number is, and if member not found -1
    let memberRow = binarySearch(firstName, lastName, data["MemberInfo"]["Columns"]["FirstName"], data["MemberInfo"]["Columns"]["LastName"], infoSheet);
    
    if(memberRow != -1){
      // Pulling the grade and dues data from the member info spreadsheet
      const grade = infoSheet.getRange(data["MemberInfo"]["Columns"]["Grade"] + memberRow).getValue().split(" ")[0];
      let dues = "";
      if(infoSheet.getRange(data["MemberInfo"]["Columns"]["Dues"] + memberRow).getValue() == data["MemberInfo"]["DuesValue"]) dues = "X";
      
      // Pulling the grade and dues from the point spreadsheet
      const foundDues = pointSheet.getRange(data["Points"]["Columns"]["Dues"] + i).getValue();
      const foundGrade = pointSheet.getRange(data["Points"]["Columns"]["Grade"] + i).getValue();
      
      // Check if the grade and dues from the two spreadsheets match, if not change the values on the point spreadsheet
      if(grade != foundGrade || dues != foundDues){
        ui.alert(firstName + " " + lastName + "\n" + 
                 "Displayed Dues: " + foundDues + "    --    " + "Current Dues: " + dues + "\n" + 
                 "Displayed Grade: " + foundGrade + "    --    " + "Current Grade: " + grade);
        
        // Setting values in the point spreadsheet for grade and dues accordingly
        pointSheet.getRange(data["Points"]["Columns"]["Grade"] + i).setValue(grade);
        pointSheet.getRange(data["Points"]["Columns"]["Dues"] + i).setValue(dues);
        
        // Incrementing n to track how many members were synced
        n++;
      }
    }
    
  }
  
  // Informational panels to show what script did
  if(n == 0) ui.alert("No member information changed.");
  else ui.alert(n + " members have been synced.");
}
