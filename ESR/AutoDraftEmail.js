/* checks that all statuses for every order over threshold are not blank (customer has been contacted) for each location. and returns "Complete" or "Incomplete" for that location. This is then displayed on "Status by Location" sheet for RMs to easily check 
their teams. */

// **add overdue check? return "Overdue" if past 15th of the month. need to bone up on javascript datetime.

const staffedStoreCodes = ["BB","BL","CLT","DET","EL","EV","IC","IND","KZ","LN","LX","MA","MG","MKE","MN","NOR","PB","SU","TOL","UNC","WL"];
var lastRefresh = new Date();

function checkComplete(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName(sheetName);
  let test2 = [];
  // arbitrary starting point to find last row. this is the first cell of orderNum column, hard coded
  let first = s.getRange(3, 8);
  // ctrl+shift+down to find last actual row of data
  let lastRow = first.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  // uses above to isolate the range of data we are working with...customer name - status
  // numRows&numColumns are reduced to account for ignored rows/colums (r.g. ignoring column 1 (year) + blank rows with formulas and titles)
  let range = s.getRange(3, 2, lastRow-2, s.getDataRange().getLastColumn()-1);
  let values = range.getValues();
  // initiate array to store all rows where OVERTHRESHOLD == true via for loop
  // this will effectively store all of the Filter Views
  let overThresholdRows = [];
  for (let row = 0; row <= values.length-1; row++) {
    let threshold = values[row][22];
    // identifying admin last name so that we can ignore Seth's orders (which are under BB)
    let adminLast = values[row][8];
    if (threshold === true && adminLast != "Greene") {
      overThresholdRows.push(values[row]);
      test2.push(values[row][6]);
      }
    }
  // loops through all overThresholdRows...if any statuses are not updated, store code is added to "incomplete" array -- to be complete, all statuses must be updated
  let incomplete = [];
  let needFollowUp = [];
  for (let r = 0; r <= overThresholdRows.length-1; r++) {
    let storeCode = overThresholdRows[r][6].toString().trim();
    let status = overThresholdRows[r][23];
    // if (staffedStoreCodes.includes(storeCode) && status == "Unknown (Contacted)" && !(needFollowUp.includes(storeCode)) && [last updated time logic?])
    if (staffedStoreCodes.includes(storeCode) && status == "" && !(incomplete.includes(storeCode))) {
      incomplete.push(storeCode);
    }
  }
  incomplete.sort();
  let mgTest = incomplete.filter(x => x=="MG")
  Logger.log(mgTest.length)
  // final loop, pushes "complete/incomplete" to returnStatus array -- this is what will be returned and displayed on actual Status Sheet
  const returnStatus = [];
  const statusSheet = ss.getSheetByName("Status by Location");
  const checkerCodes = statusSheet.getRange("A2:A22").getDisplayValues();
  for (let row = 0; row <= checkerCodes.length-1; row++) {
    if (incomplete.includes(checkerCodes[row][0])) {
      returnStatus.push('Incomplete');
    } else {
      returnStatus.push('Complete');
    }
  }
  // returns Complete/Incomplete as array. Array's are displayed vertically in sheet, so should append nicely in the column
  return returnStatus;
  
  /* NOTE: this function feels very inefficient w/ all the separate loops..need to come back and refactor this at some point
  there must be an easier/better way that i am too much of a noob to think of currently...will revisit */
}



/* function for the refresh button on Status by Location page
   will go through each column and check if data present, if so, column needs refreshed
   flushes spreadsheet and refreshes the formula in each column w/ data present */

function refreshButton() {
  // logs time of execution for updating the lastRefresh variable -- used in onOpen(e) function to auto refresh if over 24 since last
  var currentTime = new Date();
  // grab main sheet to look at and edit, "Status by Location"
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Status by Location');
  // append all sheet names to array, will be used to fill out formula
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetNames = [];
  for (i = 0; i<sheets.length; i++) { sheetNames.push(sheets[i].getName()) }
  /*   
    *** hard coded range of cells w/ data *** 
  */
  var range = s.getRange("B1:H22");
  var data = range.getDisplayValues();
  var numCols = range.getNumColumns();
  var firstDataRow = data[1];
  var activeMonths = [];
  // loop through each column and check if first Data row (row 2) is empty, if not it will need refreshed
  // ** edited to only refresh the last 3 months of data for efficiency (numCols-3) **
  for (i=numCols-3;i<numCols;i++) {
    if (data[1][i] != "") { 
      activeMonths.push(data[0][i]);
      var cell = range.getCell(2, i+1);
      let formula = cell.getFormula();
      cell.setValue("Refreshing...");
      SpreadsheetApp.flush();
      cell.setValue(formula);
    }
  }
  // sets the value of 'Last Refreshed' cell in sheet (hard coded) for storing Date object to be checked against in onOpen function + shows user last refresh time
  var refreshDisplay = s.getRange("B24");
  refreshDisplay.setValue(currentTime);

};


// /* 
// assuming user will be clicking on email to create draft, get email data from row
// -returns dictionary with various order data */
function parseRowData(r=null) {
  s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // -optionally, fn will accept a row number as argument, if none looks for current selected cell
  if (r === null) {
    var c = s.getCurrentCell();
    var col = c.getColumn();
    var row = c.getRow();
  } else {
    const emailCol = 4
    var row = r;
    var c = s.getRange(r, emailCol);
    var col = emailCol
  }
  // all locations are relative to emailCol...which should be user selection 
    var custEmail = c.getDisplayValue();
    var custFirst = titleCase(s.getRange(row, col-2).getDisplayValue());
    var custLast = titleCase(s.getRange(row, col-1).getDisplayValue());
    var adminFirst = titleCase(s.getRange(row, col+5).getDisplayValue());
    var adminLast = titleCase(s.getRange(row, col+6).getDisplayValue());
    var jobName = s.getRange(row, col+2).getDisplayValue();
    var overThreshold = s.getRange(row, col+20).getValue();
    var status = s.getRange(row, col+21).getDisplayValue();
    var orderNum = s.getRange(row, col+1).getDisplayValue();
    var orderData = {
      'custEmail':custEmail,
      'custFirst':custFirst,
      'custLast':custLast,
      'adminFirst':adminFirst,
      'adminLast':adminLast,
      'jobName':jobName,
      'orderNum':orderNum,
      'overThreshold': overThreshold,
      'status': status
    }
    return orderData;
    // **note to self: call orderData items w/ dot.notation, not like dict[python]
}

// // make names (or any string) title case...so much work in js, not gonna bother with job names, dont plan to use them right now
function titleCase(str) {
  let cleanStr = str.trim();
  let newStr = cleanStr.slice(0, 1).toUpperCase() + cleanStr.slice(-(cleanStr.length-1)).toLowerCase();
  return newStr;
}

// row3162_data = parseRowData(3162);
// row134_data = parseRowData(134);
// Logger.log(Object.values(row134_data));
// Logger.log(row134_data.overThreshold === true);

function draftEmail(orderData) {
  var d = orderData;
  var signature = Gmail.Users.Settings.SendAs.list("me").sendAs.filter(function(account){if(account.isDefault){return true}})[0].signature;
  var pogoLink = "https://undergroundshirts.com/pages/pop-up-online-store";
  var emailBody = `<p>Hey ${d['custFirst']},<p>`;
  emailBody += `<p>Around this time last year, you placed an order for <i>**${d['jobName']}**</i>, so I thought I would reach out to see if you need any help placing a similar order this year.<p>`;
  emailBody += `<p>While we’ve moved past a majority of events being virtual, we do know there are still circumstances that may require virtual events, or simply people who would just rather order their swag online! Our <a href=${pogoLink}> Online Pop Up Shops</a> are a perfect tool to use! In short, we create an online order form, your customers pay for their orders, and we can ship them directly to each person!<p>`;
  emailBody += '<p>We are looking forward to working with you again! Let me know if you have any questions.<p>';
  emailBody += signature;
  var emailSubj = 'Hoping to help you again this year! Quick check in from Underground Printing.'
  var adminEmail = Session.getActiveUser().getEmail(); // gets email of current user 
  Logger.log(adminEmail);
  GmailApp.createDraft( // call Gmail to create the draft..'null' skips required 'body' argument, using htmlBody instead
    d['custEmail'],
    emailSubj,null,{htmlBody: emailBody});
}

function promptDraft() {
  const range = SpreadsheetApp.getCurrentCell();
  var row = range.getRow();
  Logger.log('Selection Row: '+row);
  var ui = SpreadsheetApp.getUi();
  if (range.getColumn() === 4) {
    var email = range.getDisplayValue();
    Logger.log('Email?: '+email);
    var trigger = ui.alert(`Would you like to draft an email to ${email}?`, ui.ButtonSet.YES_NO);
    if (trigger == ui.Button.NO) {
      Logger.log('Cancelled by User.');
      ui.alert('Email Draft Cancelled.', ui.ButtonSet.OK);
    } else if (trigger == ui.Button.YES) {
      rowData = parseRowData(row);
      Logger.log(Object.values(parseRowData));
      draftEmail(rowData);
      ui.alert('Email drafted succesfully. Check your Drafts folder!');
    }
  } else {
    ui.alert('No email selected. Please click on a customer email and try again.', ui.ButtonSet.OK);
  }
}

function onOpen(e) {
  // creates Script menu for executing Draft Email function and adds to UI
  var menu = SpreadsheetApp.getUi();
  menu.createMenu("Scripts")
  .addItem("Draft Email", 'promptDraft')
  .addToUi();
  // auto refresh check...will auto refresh "Staus" list if it's been over 24 hours since last refresh
  // because this is an onOpen trigger, should run this check every time a user with edit access opens the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("Status by Location");
  // resfreshDisplay = cell in sheet containing last updated Date object (hard coded)
  var refreshDisplay = s.getRange("B24");
  var lastRefresh = null;
  var now = new Date()
  // if there is no value in sheet, update lastRefresh to now and update the cell, else set lastRefresh to value in cell
  if (!refreshDisplay.getDisplayValue()) {
    lastRefresh = now.getTime()
    refreshDisplay.setValue(now)
  } else {
    lastRefresh=refreshDisplay.getValue().getTime()
  }
  // get time between lastRefresh and now in ms, converts to hours
  var msSinceLastRefresh = now - lastRefresh;
  Logger.log("ms since refresh: "+msSinceLastRefresh);
  var hoursSinceRefresh = msSinceLastRefresh / (60 * 60 * 1000);
  Logger.log("hourse since refresh:" +hoursSinceRefresh);
  // if lastRefresh has been over 24 hours, run refreshButton and update value in sheet
  if (hoursSinceRefresh >= 24) {
    refreshDisplay.setValue(now);
    refreshButton();
    Logger.log(`Refreshed at ${now}`)
  }
}
