/* Created by Andy Holbert - updated on GitHub 9/21/2023. Unauthorized distribution prohibited. This will be used on Google Apps Script only.
REQUIREMENTS: This program requires this template listed https://docs.google.com/spreadsheets/d/1e3BVSQPxp2OfCD12tUriCA1NL0ii6gUk8mKLD8m-UtY/edit?usp=sharing. Following this template will allow for the program to be run by students.

PURPOSE: This program takes survey data from students to assign team roles based on preferences submitted by each team. This program takes data directly through Google Forms and automates the entire process. There is also an instruction guide for setting up the Google Form, but this contains private data on the instruction guide. Reach out to Andy Holbert for this information through GitHub.

*/

// BEGIN CODE


// Creates a Ui for running the program

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Run Setup')
      .addItem('Run Initial Setup', 'importSheet')
      .addToUi();
}

function importSheet() {
  
  // Helper
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var sourceurlprompt = ui.prompt("Please enter the URL of the responses spreadsheet. Please ensure this is set to 'anyone with the link can view.'");

  var sourceurl = sourceurlprompt.getResponseText();

  ss.getSheetByName('Entries_From_Standard').getRange('A1').setValue('=IMPORTRANGE("' + sourceurl + '","A:I")');
  ui.alert("Sheet successfully setup. Click OK to continue.");

  ui.alert("This sheet will now format all data to match with your survey. Press OK to continue.");

// This entire section numbers rows on a sheet based on their team to allow for updated formula calculation for data import. The maximum allowed team members for this program (again) is 5.

  var startvalue = sheet.getRange('F2').getValue();
  var startwithone = sheet.getRange('V2').setValue(1);
  var lastrow = sheet.getLastRow();
  var valuebase = startvalue;
  
  var cellcounter = 3;
  var valuetocheck = sheet.getRange(cellcounter,6).getValue();
  var cellnumber = 1;
  

  // If the team number of the next row is the same as the previous, a number +1 is assigned to the row. If this is not the case, the next row value is reset to 1. The "Assigned Role" cell has a condition indicating what each of these numbers mean for an if formula listed in columns W,X,Y,Z,AA. 
  while (cellcounter <= lastrow) {
    while (valuetocheck == valuebase) {
      var cellnumber = cellnumber + 1;
      var updatenextcell = sheet.getRange(cellcounter,22).setValue(cellnumber);
      var cellcounter = cellcounter + 1;
      var valuetocheck = sheet.getRange(cellcounter,6).getValue();
    };
    var valuebase = sheet.getRange(cellcounter,6).getValue();
    var cellnumber = 1;
    var updatenextcell = sheet.getRange(cellcounter,22).setValue(cellnumber);
    var cellcounter = cellcounter + 1
  };

// 

// Alert that the process is complete
ui.alert("Import is complete! If you have any questions or experience any errors, email Andy Holbert at holberat@miamioh.edu. Thank you!");

}
