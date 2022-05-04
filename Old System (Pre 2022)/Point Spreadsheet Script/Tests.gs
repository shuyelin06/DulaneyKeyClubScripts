/*
  Test binary search with the general sheet. If the search gives a row different from that of the real row of a member, it will output the member and row number. 
*/
function binarySearchGeneral(){
  var generalSheet = pointSpreadsheet.getSheetByName("General");
  for(var i=0; i<generalSheet.getLastRow()-1; i++){
    const first = generalSheet.getRange("B"+1).offset(i, 0).getValue();
    const last = generalSheet.getRange("A"+1).offset(i, 0).getValue();

    var found = binarySearch(first, last, "B", "A", generalSheet);

    if((i+1)!=found){
      console.log(first + last);
      console.log(found);
    }
  }
}

/*
  Test binary search with the information sheet. If the search gives a row different from that of the real row of a member, it will output the member and row number for debugging.
*/
function testInfoSheet(){
  var infoSheet = memberInfoSpreadsheet.getSheetByName("Member Information");
  for(var i=0; i<infoSheet.getLastRow()-1; i++){
    const first = infoSheet.getRange("B"+1).offset(i, 0).getValue();
    const last = infoSheet.getRange("C"+1).offset(i, 0).getValue();

    var found = binarySearch(first, last, "B", "C", infoSheet);

    if((i+1)!=found){
      console.log(first + last)
      console.log(found);
    }
  }
  console.log("Done.");
}