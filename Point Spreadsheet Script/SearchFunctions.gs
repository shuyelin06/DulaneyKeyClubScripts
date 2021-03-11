/*
  This is my child. This file contains the script for a binary search for a last and first name of a member in a given sheet with given first and last name columns.
  CAN ONLY BE USED IN A SORTED SHEET. Make sure you sort the sheet (Descending A-Z) before calling binarysearch
*/
// Attempt at using binary search to do the syncinfo function faster
function binarySearch(firstName, lastName, firstCol, lastCol, sheet){
  // Initializing i and j, needed for the binary search
  let i = sheet.getLastRow()/2;
  let j = 0;
  
  // Initializing the return variable at -1, so that if member is not found the function returns -1
  let output = -1;
  
  // Binary search for the name
  let notFound = true;
  while(notFound){
    let foundLast = sheet.getRange(lastCol+Math.ceil(i)).getValue();
    
    // Failsafe break statement in the event that the member does not exist
    if(Math.abs(i-j)<0.5){
      notFound = false;
      continue;
    }
    
    // Check if last name matches, if not then increments i accordingly to continue searching
    if(lastName == foundLast){
      // searchAround() function searches above and below a radius around the row number found with a matching last name, to find the corresponding first name
      let n = searchAround(firstName, sheet, i, firstCol, lastCol);
      if(n != -1){
        output = n;
      }
      notFound = false;
    } 
    else if (lastName > foundLast){
      let a = i;
      i += Math.abs(i-j)/2;
      j = a;
    } 
    else i = (j+i)/2;
    
    
  }
  // Output will be the row number of the member, or will be -1 if member is not found.
  console.log("Search done");
  return output;
}

// Because there are multiple members with the same last name, with a given last name this function will search for the position of the first name we are looking for.
function searchAround(first, sheet, row, firstCol, lastCol){
  // The scope in which the searchAround function will search around.
  const radius = 6;
  
  let i = Math.ceil(row);
  const last = sheet.getRange(lastCol + i).getValue();
  let j = 0;
  
  // r is the return value for this function
  let r = -1;
  
  // Initializing the while loop
  let notFound = true;
  
  while(notFound){
    // Initializing the found last and first for the rows
    let foundFirst = sheet.getRange(firstCol + i).getValue();
    let foundLast = sheet.getRange(lastCol + i).getValue();
    
    if(j > radius * 2){
      notFound = false;
      continue;
    }
    
    if(first == foundFirst && last == foundLast){
        r = i;
        notFound = false;
    } 
    else {
      j++;
      // Modulus operator allows alternation so that we alternate between values above and below the intial center
      if(j % 2 == 0) i -= j;
      else i += j;
    }
      
  }
  
  // Return statement for r
  return r;
}