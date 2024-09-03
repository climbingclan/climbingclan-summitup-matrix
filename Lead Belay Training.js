
function leadBelayTrainingData() {
var conn = Jdbc.getConnection(url, username, password);
var stmt = conn.createStatement();

var spreadsheet = SpreadsheetApp.getActive();
let dashboard = spreadsheet.getSheetByName('Dashboard');
var cell = dashboard.getRange('B5').getValues();
cell = (cell[0][0] === "" || cell[0][0] === null || cell[0][0] === "#N/A") ? 18794 : cell;

 var results = stmt.executeQuery('select distinct `first_name` "Name",`nickname` "fbname",  CAST(`stats_volunteer_for_numerator_cached` AS UNSIGNED INTEGER) "Volunteering",`scores_volunteer_score_cached` "Receptiveness",`skills-belaying` "Belaying Skills",`climbing-indoors-toproping-grades` "Indoor TR",`admin-wednesday-requests-notes` as `Requests and notes`,  `climbing-indoors-leading-grades` "Indoor Lead",`skills-sport-climbing` "Sport Skills"  from wp_member_db db LEFT JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id JOIN wp_member_db_scores s on s.user_id = db.id where `product_id`=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold")  AND `skills-belaying` in ("Top-rope-belaying","learner-lead-belayer") AND (`skills-sport-climbing` IS NULL OR `skills-sport-climbing`<>"%Lead%") order by CAST(`stats_volunteer_for_numerator_cached` AS UNSIGNED INTEGER) DESC, `skills-belaying` DESC, `skills-trad-climbing` ASC,`climbing-happy-to-supervise` DESC');

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Lead Training');
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



function topRopeTrainingData() {
var conn = Jdbc.getConnection(url, username, password);
var stmt = conn.createStatement();

var spreadsheet = SpreadsheetApp.getActive();
let dashboard = spreadsheet.getSheetByName('Dashboard');
var cell = dashboard.getRange('B5').getValues();
cell = (cell[0][0] === "" || cell[0][0] === null || cell[0][0] === "#N/A") ? 18794 : cell;


 var results = stmt.executeQuery('select distinct `first_name` "Name",`nickname` "fbname",  CAST(`stats_attendance_attended_cached` AS UNSIGNED INTEGER) "Attended",`skills-belaying` "Belaying Skills",`admin-wednesday-requests-notes` as `Requests and notes`, CAST(`stats_volunteer_for_numerator_cached` AS UNSIGNED INTEGER) "Volunteered"  from wp_member_db db LEFT JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id JOIN wp_member_db_scores s on s.user_id = db.id where `product_id`=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold")  AND (`skills-belaying` IS NULL OR `skills-belaying` in ("No-belaying", "No-Belaying") ) order by CAST(`stats_attendance_attended_cached` AS UNSIGNED INTEGER) DESC, `skills-belaying` DESC, `skills-trad-climbing` ASC,`climbing-happy-to-supervise` DESC');

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('TR Belay Training');
  sheet.clearContents();
  sheet.clearFormats();

appendToSheet(sheet,results);

setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","0","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","2","#ffd898")
setColoursFormatLessThanOrEqualTo(sheet, "C3:C1000","10","#FFFFFF")

//setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","10","#ff75d8")
//setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","20","#ffd898")
//setColoursFormatLessThanOrEqualTo(sheet, "D3:D1000","30","#fad02c")
sheet.setColumnWidth(2, 150);
sheet.setColumnWidth(5, 300);
  setColoursFormat(sheet, "D1:D1000", "No-belaying", "#ffcccb")


results.close();
stmt.close();
//sheet.autoResizeColumns(1, numCols+1);

} 