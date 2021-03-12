// Custom "Writeups" menu for board members to use.
function onOpen(){
  ui.createMenu('Writeups')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Membership Secretary')
      .addItem('Export Write-Up', 'exportWriteup'))  
    .addSeparator()  
    .addItem('Create New Write-Up', 'createWriteup')
    .addSeparator()
    .addItem('Lock Current Write-Up', 'finished')
    .addToUi();
}

// Creates and formats a writeup sheet
function createWriteup(){
  var numVol = ui.prompt("Please Enter the Number of Volunteers", ui.ButtonSet.OK_CANCEL);
  if(numVol.getSelectedButton() == ui.Button.CANCEL || numVol.getSelectedButton() == ui.Button.CLOSE) return;
  var sheet = writeup.insertSheet();
  let n = parseInt(numVol.getResponseText(), 10);
  let d = new Date();
  
  // Formatting
  // Formatting for the main frame
  sheet.getRange("B2:H17")
    .setBorder(true, true, true, true, false, false)
    .setBackground("#efefef")
    .setTextStyle(writeUpText)
    .setHorizontalAlignment("Center")
    .setVerticalAlignment("middle");
  
  // Alignment for the left hand information sections
  sheet.getRange("D6:D16")
    .setHorizontalAlignment("Left");
  
  // Formatting for volunteer frame 
  sheet.getRange("B19:H"+(n+21))
    .setBorder(true, true, true, true, false, false)
    .setTextStyle(writeUpText)
    .setBackground("#efefef")
    .setHorizontalAlignment("Center")
    .setVerticalAlignment("middle");
  
  // Formatting for the volunteer grid
  sheet.getRange("C20:G"+(n+20))
    .setBorder(true, true, true, true, true, true, "#d9d9d9", null);
  
  // Formatting for checking volunteer names
  sheet.getRange("J2:K"+(n+21))
    .setBorder(true, true, true, true, true, true, "d9d9d9", null)
    .setTextStyle(writeUpText)
    .setBackground("#efefef")
    .setHorizontalAlignment("Center")
    .setVerticalAlignment("middle");
  
  // Values and data validation for the write-up main frame
  sheet.getRange("E3").setValue("Event Write-Up");
  sheet.getRange("E4").setValue('Need help? Refer to the "Guide" sheet.');
  for(var key in writeupFormat){
    sheet.getRange(key).setValue(writeupFormat[key]);
    sheet.getRange(key).offset(0, 1).setBorder(false, false, true, false, false, false);
  }
  sheet.getRange("G6").setValue(n);
  sheet.getRange("G8").setValue('=SUM(F21:F)');
  sheet.getRange("G10").setValue('=SUM(G21:G)');
  sheet.getRange("G12").setValue(d.getTime());
  sheet.getRange("G14").setValue("No");
  sheet.getRange("D8").setDataValidation(dateVal);

  // Values and data validation for the volunteer frame
  sheet.getRange("G16").setValue(Session.getActiveUser().getEmail());
  for(var key in volFormat){
    sheet.getRange(key).setValue(volFormat[key]);
  }
  for(var i=0; i<numVol.getResponseText(); i++){
    sheet.getRange("C21").offset(i,0).setValue("#" + (i+1));
    sheet.getRange("D21").offset(i,0).setDataValidation(lastVal);
    sheet.getRange("E21").offset(i,0).setDataValidation(firstVal);
    sheet.getRange("F21").offset(i,0).setDataValidation(pointVal);
  }

  // Values and data validation for volunteer check frame
  sheet.getRange("K2").setValue("Member Name");
  for(var i=0; i<n; i++){
    var associatedRow = 21+i;
    sheet.getRange("J3").offset(i, 0).setValue("#"+(i+1));
    sheet.getRange("K3").offset(i, 0).setDataValidation(volunteerCheck);
    sheet.getRange("K3").offset(i, 0).setFormula("=CONCATENATE(E" + associatedRow + ",\" \" ," + "D" + associatedRow + ")");
  }

  // Column and Row Editing
  sheet.deleteColumns(13, sheet.getMaxColumns()-12);
  sheet.deleteRows(22+n, sheet.getMaxRows()-(22+n));

  sheet.setColumnWidth(11, 150);
  sheet.setColumnWidths(3, 5, 150);
  sheet.setRowHeights(2, n+20, 30);
  sheet.setHiddenGridlines(true);
  
  ui.alert("Your write-up has been created under the name: \n" + sheet.getName() + "\nNote: Don't worry about the name. Once you lock the write-up, the name will change automatically!");
} 

// Locks a writeup sheet from editing
function finished(){
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Check if the person locking is the creator of the write-up (rejects if not the same)
  if(sheet.getRange("G16").getValue() != Session.getActiveUser().getEmail() && memSec != Session.getActiveUser().getEmail()){
    ui.alert("You do not have the permissions to lock this write-up \nOnly the current membership secretary and the write-up creator can lock it!");
    return;
  }
  
  // Check if the user is sure that they want to lock the write-up
  var response = ui.alert("Are you sure you want to lock this write-up? \n*PLEASE NOTE THAT YOU WILL NO LONGER BE ABLE TO EDIT IT*", ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) return;
  
  // Add sheet protection and configure its settings so only the membership secretary + write-up creator can edit it
  try{
    var protection = sheet.protect().setDescription("Finished write-ups cannot be edited by any users!");
    protection.addEditor(memSec);
    protection.addEditor(sheet.getRange("G16").getValue());
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()){
      protection.setDomainEdit(false);
    }
  } catch (error){
    ui.alert("There was an error in protecting the sheet.");
  }
  
  // Changes the name of the sheet
  sheet.setName(sheet.getRange("D6").getValue());
  
  // Sends an email to the membership secretary about the completed write-up
  MailApp.sendEmail(memSec, "Write-Up Submitted", "Write-Up " + sheet.getName() + " by " + Session.getActiveUser() + " has been submitted!");
  
  ui.alert("Success! This write-up has been locked from editing!");
}