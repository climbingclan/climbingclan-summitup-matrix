function volunteerData() {

 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();

// Volunteering Report
//
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Dashboard');
  var cell = sheet.getRange('B5').getValues();
cell = (cell[0][0] === "" || cell[0][0] === null || cell[0][0] === "#N/A") ? 18794 : cell;

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Volunteering');





// start of volunteering function
function volunteering(flip, title,cell)
{
//sheet.appendRow([title]);
//let row = sheet.getLastRow();
//console.log(row);
//sheet.getRange(row, 1, 2, 24).setFontWeight("bold");
console.log(cell)
if (flip==="volunteers") {
 var results = stmt.executeQuery('select `first_name` "First Name",`last_name` "Last Name",pd.cc_volunteer "Selected Roles",volunteer_indoor_preevent_facebook_promo "FB promo", volunteer_indoor_event_reporter "Rprtr", volunteer_indoor_check_in "Check-in", volunteer_indoor_welcome_liaison "WelcLn", volunteer_indoor_rounding_up "Rnd Up", volunteer_indoor_before_cake_tidying "PreCake", volunteer_indoor_after_cake_tidying "PostCake",volunteer_indoor_give_announcements "Anncments Vol", competency_indoor_announcements "Anncments Compet", competency_indoor_pairing "Pair", competency_indoor_trip_director "Event Dir", scores_volunteer_score_cached "Receptiveness",scores_volunteer_reliability_score_cached "V Reliability%",stats_attendance_indoor_wednesday_attended_cached "Attended", `admin-wednesday-requests-notes` as `Requests and notes`, pd.order_id "Order ID", pd.user_id "User ID",competency_indoor_floorwalker "Flwlkr",competency_indoor_skillsharer,competency_indoor_checkin from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id join wp_member_db_volunteering vl on pd.user_id = vl.id where product_id=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold")  AND stats_attendance_indoor_wednesday_attended_cached >= 2 AND cc_compliance_last_date_of_climbing BETWEEN DATE_SUB(NOW(), INTERVAL 2 MONTH) AND NOW() order by FIELD(pd.cc_volunteer,  "none", "Tuesday Facebook Promo","tuesdaypromo2", "Wednesday Facebook Promo","wednesdaypromo2", "check in", "welcome liaison", "pairing","rounding up", "Pre-cake Preparation","announcements",  "after cake tidy", "event director", "Event Reporter","Wednesday Evening Trainer", "floorwalker","Friday Facebook Promo","Summit Up Promo") asc, pd.cc_volunteer asc, CAST(db.scores_volunteer_score_cached AS UNSIGNED INTEGER) asc,CAST(db.stats_attendance_indoor_wednesday_attended_cached AS UNSIGNED INTEGER) desc,volunteer_indoor_event_reporter,volunteer_indoor_check_in')
}
else if (flip==="nonvolunteers"){
 var results = stmt.executeQuery('select `first_name` "First Name",`nickname` "Facebook Name",pd.cc_volunteer "Selected Roles",volunteer_indoor_preevent_facebook_promo "FB promo", volunteer_indoor_event_reporter "Rprtr", volunteer_indoor_check_in "Check-in", volunteer_indoor_welcome_liaison "WelcLn", volunteer_indoor_rounding_up "Rnd Up", volunteer_indoor_before_cake_tidying "PreCake", volunteer_indoor_after_cake_tidying "PostCake",volunteer_indoor_give_announcements "Anncments", volunteer_indoor_floorwalker "Flwlkr", volunteer_indoor_pairing "Pair", volunteer_indoor_event_director "Event Dir", scores_volunteer_score_cached "Receptiveness",stats_attendance_indoor_wednesday_attended_cached "Attended",`admin-wednesday-requests-notes` as `Requests and notes`, pd.order_id "Order ID", pd.user_id "User ID" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id join wp_member_db_volunteering vl on pd.user_id = vl.id where product_id=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold") AND  `admin-first-timer-question`="no" AND (`admin-can-you-help`="" AND pd.cc_volunteer="none") AND scores_volunteer_score_cached<>"" AND CAST(stats_attendance_indoor_wednesday_attended_cached AS UNSIGNED INTEGER)>="2" order by pd.cc_volunteer asc, CAST(db.scores_volunteer_score_cached AS UNSIGNED INTEGER) asc, CAST(db.stats_attendance_indoor_wednesday_attended_cached AS UNSIGNED INTEGER) desc ') 
}
 //sheet.clearContents();

appendToSheet(sheet,results);





} //end of volunteering

// full options
//help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening
  volunteering("volunteers", "Volunteers",cell);
  sheet.clearFormats();

  setNumberFormat(sheet, "O2:O1000", "0");

setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","50","#fad02c")
setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","80","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","90","#ffd898")

setColoursFormatLessThanOrEqualTo(sheet, "O2:O1000","10","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "O2:O1000","20","#ffd898")
setColoursFormatLessThanOrEqualTo(sheet, "O2:O1000","30","#fad02c")
setColoursFormat(sheet, "C2:C1000","none","#DAF7A6 ")
setColoursFormat(sheet, "C2:C1000","Selected","#FFFFFF")
setColoursFormat(sheet, "C2:C1000","","#e0ffff")
setTextFormat(sheet,"D2:N1000","No","#a9a9a9")

setColoursFormat(sheet, "C2:C1000","none","#DAF7A6 ")
setColoursFormat(sheet, "L2:N1000","Signed","#DAF7A6 ")

//volunteering("nonvolunteers", "People who haven't volunteered or been assigned");

sheet.setColumnWidth(1, 80);
sheet.setColumnWidth(2, 90);
sheet.setColumnWidth(3, 80);
sheet.setColumnWidth(4, 30);
sheet.setColumnWidth(6, 50);
sheet.setColumnWidth(7, 60);
sheet.setColumnWidth(8, 70);
sheet.setColumnWidth(9, 70);
sheet.setColumnWidth(10, 60);
sheet.setColumnWidth(11, 60);
sheet.setColumnWidth(12, 60);
sheet.setColumnWidth(13, 60);
sheet.setColumnWidth(14, 60);
sheet.setColumnWidth(15, 80);
sheet.setColumnWidth(16, 80);
sheet.setColumnWidth(17, 80);
sheet.setColumnWidth(18, 200);
sheet.setColumnWidth(19, 60);
sheet.setColumnWidth(20, 60);




setWrapped(sheet,"r2:r1000");


//var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
 

//setNumberFormat(sheet, "O3:O1000", "[<=75]'Try_Give_Break';[<=30]'Available';[<=20]'Try_Assign'");

//[<=75]"Available";[Color22][<=30]"Assign If Possible";0;[Magenta]_(@_)

return "done"
} 

