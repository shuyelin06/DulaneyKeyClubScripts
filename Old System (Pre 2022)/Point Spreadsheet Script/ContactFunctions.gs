function contactMembers(writeupSheet) {
  var infoSheet = memberInfoSpreadsheet.getSheetByName("Member Information");
  
  // Defining variables for use later
  const eventName = writeupSheet.getRange(data["WriteUps"]["Cells"]["Name"]).getValue();
  
  // Array of emails to email after
  let contactList = [];
  
  for(var i = 21; i<=writeupSheet.getLastRow(); i++){
    const lastName = writeupSheet.getRange(data["WriteUps"]["Columns"]["LastName"] + i).getValue();
    const firstName = writeupSheet.getRange(data["WriteUps"]["Columns"]["FirstName"] + i).getValue();
    
    // Use of binary search function to find the corresponding email for a given first and last name
    let memberRow = binarySearch(firstName, lastName, data["MemberInfo"]["Columns"]["FirstName"], data["MemberInfo"]["Columns"]["LastName"], infoSheet);
    if(memberRow == -1) continue;
    
    // Use of regex (.split method) to make sure in the event that there are multiple emails separated by a comma, that only the first one is taken.
    const email = infoSheet.getRange(data["MemberInfo"]["Columns"]["Email"] + memberRow).getValue().split(",")[0];
    
    contactList.push(email);
  }
  
  MailApp.sendEmail(contactList.join(","), "Event Entered",
                    "Event: " + eventName + ", has been entered into the points sheet!" + "\n" +
                    "\n" + "You are receiving this message because you were one of the volunteers for this event, and have received points because of it!"
                   );
  
  ui.alert("Members Notified");
}

// Contacts the chair that their write-up was entered
function contactChair(email, writeupName){
  MailApp.sendEmail(email, "Your Write-Up Has Been Entered!", 
                   "Your Dulaney High Key Club Write-Up for, " + writeupName + " has been entered by the Membership Secretary!" + "\n" + 
                    "Your write-up will soon be exported to the corresponding month spreadsheet, if you ever need to access it again." + "\n" +
                    "*This is part of an automated script to simply notify you when your write-up has been entered. If you have any concerns, please respond to this email*"
                   );
}