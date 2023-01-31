/* 

Script for moving one row of Google Sheet to another
Currently is very specific to a particular spreadsheet, but can be eeasily modified for
any purpose. 

*/ 


function onEdit(event) {
    // assumes source data in sheet named Needed
    // target sheet of move to named Acquired
    // test column with yes/no is col 4 or D
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = event.source.getActiveSheet();
    var r = event.source.getActiveRange();
  
    if(s.getName() == "Mike Active POGOs" || s.getName() == "Nate Active POGOs" && r.getColumn() == 4 && r.getValue() == "8 - Closed") {
      var row = r.getRow();
      var numColumns = s.getLastColumn();
      var targetSheet = ss.getSheetByName("Closed POGOs");
      var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      s.getRange(row, 1, 1, numColumns).moveTo(target);
      s.deleteRow(row);
    }
  }