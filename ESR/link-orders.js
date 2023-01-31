import 'google apps script';

/*
goes through column E (order numbers) and replaces contents w/ link to the order in UGP system
simply adds the rest of the Sheets formula to the order number already in the cell
*/
function link_orders() {
  let hyper = '=hyperlink("https://admin.undergroundshirts.com/c2/mongo_order/mongo_orders/view/';
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = s.getActiveSheet();
  var order_num_col = sheet.getRange("E3:E");
  // this is same as ctrl+dwn_Arrow...gets str of last row # w/ data
  var last_row = order_num_col.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow().toString();
  var order_nums_reduced = sheet.getRange("E3:E" + last_row);
  var order_nums_raw = order_nums_reduced.getValues();
  Logger.log(`Creating hyperlinks for ${order_nums_raw.length} cells`)
  // for every order num, turn into link
  for (i=0; i<order_nums_raw.length; i++) {
    var cell = sheet.getRange(i+3, 5)
    cell.setValue(hyper+order_nums_raw[i]+'", "'+order_nums_raw[i]+'")')
  }
  Logger.log("Order numbers linked succesfully!")
}


