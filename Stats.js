function statsData() {
var conn = Jdbc.getConnection(url, username, password);
var stmt = conn.createStatement();

var spreadsheet = SpreadsheetApp.getActive();
let dashboard = spreadsheet.getSheetByName('Dashboard');
var cell = dashboard.getRange('B5').getValues();
cell = (cell[0][0] === "" || cell[0][0] === null || cell[0][0] === "#N/A") ? 18794 : cell;

var results = stmt.executeQuery('SELECT `skills-belaying` AS `Belaying Skills`, COUNT(*) AS `How many this week` FROM wp_member_db db JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id WHERE `product_id` = ' + cell + ' AND  `cc_location`="' + cc_location + '"  AND `status` IN ("wc-processing", "wc-onhold", "wc-on-hold") GROUP BY `skills-belaying`');

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Stats');
  sheet.clearContents();
  sheet.clearFormats();

appendToSheet(sheet,results);
setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","0","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","2","#ffd898")
setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","5","#fad02c")
setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","10","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","20","#ffd898")
setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","30","#fad02c")



sheet.setColumnWidth(2, 150);
sheet.setColumnWidth(7, 300);
  setColoursFormat(sheet, "E1:E1000", "learner-lead-belayer", "#FFFF00")
  setColoursFormat(sheet, "E1:E1000", "lead-belayer", "#5CFF5C")
  setColoursFormat(sheet, "E1:E1000", "top", "#ADD8E6")
  setColoursFormat(sheet, "E1:E1000", "No-belaying", "#ffcccb")
setNumberFormat(sheet, "D3:D1000", "0");

results.close();
stmt.close();
//sheet.autoResizeColumns(1, numCols+1);

} 

