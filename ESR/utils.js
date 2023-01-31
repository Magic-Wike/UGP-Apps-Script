//some non essential utility functions i may come back to later 

// // takes columnID in A1 Notation and returns Range() with all non empty cells
// // this is kind of like hitting ctrl+shift+[dwn_Arrow]..must be easier way
// // '3' (row 3) is hard coded. the topmost data row in current formatting

// function getDataRange(column) {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var s = ss.getActiveSheet();
//   var col = s.getRange(column+'3:'+column);
//   var last_row = col.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow().toString();
//   var dRange = s.getRange(column+'3:'+column+last_row);
//   return dRange;
// }

// goes through getDataRange column and returns list of emails (or whatever really)
// function getCustomerEmails() {
//   let columnD = getDataRange("D").getDisplayValues()
//   let emails = []
//   for (const e in columnD) {
//     if (!(e in emails)) {
//       emails.push(e)
//     } 
//   return emails
//   }
// }

function testDate() {
    const d = new Date();
    const m = d.toLocaleString('default', { month : 'long' })
    Logger.log(d);
    Logger.log(m);
  }
  