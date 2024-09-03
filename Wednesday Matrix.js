function runWednesdayMatrix() {
  
  runOutputTab();
  volunteerData();
  badgesData();
  leadBelayTrainingData();
  topRopeTrainingData();
  //runVolunteeringTab();
  runRolesTab();
  //runTestTab();
  statsData();
  onOpen();

  readEvents();

}




function runOutputTab() {
  console.log("output tab")

var conn = Jdbc.getConnection(url, username, password);
var stmt = conn.createStatement();

var spreadsheet = SpreadsheetApp.getActive();
let dashboard = spreadsheet.getSheetByName('Dashboard');
var cell = dashboard.getRange('B5').getValue() || "18794";
cell = (cell === "" || cell === null || cell === "#N/A") ? 18794 : cell;

  var results = stmt.executeQuery('select `admin-first-timer-indoor` AS New,"Arrived","BaseC","Paired", `first_name` as "Name",`last_name` as "Surname", `skills-belaying` as `Belaying Skills`,pd.cc_volunteer AS "Volunteer",`admin-wednesday-requests-notes` as `Requests and notes`,scores_attendance_reliability_score_cached "Reliability%",  `pd`.`order_id` as `Order ID`, `pd`.`user_id` as `Clan ID`   from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where `product_id`=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold") order by `first_name` ASC');
  //console.log(results);

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Output');
  //console.log(typeof(results))
  sheet.clearContents();
appendToSheet(sheet,results);
  sheet.clearFormats();
setWrapped(sheet,"i2:i1000");

  setColoursFormat(sheet, "A1:A1000", "Yes", "#ffcccb")
  setTextFormat(sheet, "A1:A1000", "No", "#a9a9a9")
  setColoursFormat(sheet, "B1:D1000", "True", "#90ee90")
  setColoursFormat(sheet, "G1:G1000", "learner-lead-belayer", "#FFFF00")
  setColoursFormat(sheet, "G1:G1000", "lead-belayer", "#5CFF5C")
  setColoursFormat(sheet, "G1:G1000", "top", "#ADD8E6")
  setColoursFormat(sheet, "G1:G1000", "No-belaying", "#ffcccb")
  sheet.setColumnWidth(9, 300);
  setTextFormat(sheet, "h2:h1000", "none", "#a9a9a9")
setColoursFormat(sheet, "h2:h1000","Check","#ADFF2F")
setColoursFormat(sheet, "h2:h1000","welcome","#ADFF2F")
setColoursFormat(sheet, "h2:h1000","pairing","#ADFF2F")
setColoursFormat(sheet, "h2:h1000","Event","#FF00FF")
  setTextFormat(sheet, "h2:h1000", "Tuesday", "#DEB887")
  setTextFormat(sheet, "h2:h1000", "Wednesday", "#DEB887")

setColoursFormat(sheet, "h2:h1000","Pre Cake","#F0E68C")
setColoursFormat(sheet, "h2:h1000","After Cake","#F0E68C")
setColoursFormat(sheet, "h2:h1000","Rounding","#F0E68C")
setColoursFormat(sheet, "h2:h1000","announce","#FF00FF")

setColoursFormat(sheet, "h2:h1000","floor","#FFEFD5")


setColoursFormat(sheet, "h2:h1000","Climbing Reporter","#ADFF2F")
setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","50","#fad02c")

 setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","80","#ff75d8")
setColoursFormatLessThanOrEqualTo(sheet, "P2:P1000","90","#ffd898")
    sheet.setColumnWidth(1, 31);
      sheet.setColumnWidth(5, 120);
      sheet.setColumnWidth(6, 150);

  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 300);

  var range = SpreadsheetApp.getActive().getRange("Output!B2:D150");
      sheet.appendRow(["","TRUE","TRUE","TRUE","Please ask latecomers and people not on the list to sign up right now at www.climbingclan.com and show you their confirmation page or confirmation email on their phone (or borrowed phone)"]);
      sheet.appendRow(["","TRUE","TRUE","TRUE","Once you've seen their confirmation page or email, you can add them below this line"]);
      sheet.appendRow(["","TRUE","TRUE","TRUE","Do this even with people who say they've already signed up"]);
      sheet.appendRow(["","TRUE","TRUE","TRUE","---------------------------------","----","-----","------"]);


  range.insertCheckboxes();





  //results.close();
  //stmt.close();

}


function runRolesTab() {
  console.log("Roles TAb")
var conn = Jdbc.getConnection(url, username, password);
var stmt = conn.createStatement();

var spreadsheet = SpreadsheetApp.getActive();
let dashboard = spreadsheet.getSheetByName('Dashboard');
var cell = dashboard.getRange('B5').getValues();
cell = (cell[0][0] === "" || cell[0][0] === null || cell[0][0] === "#N/A") ? 18794 : cell;
  var sheet = spreadsheet.getSheetByName('Roles');
  sheet.clearContents();
  sheet.clearFormats();
    sheet.appendRow(["Role","Ideal number of people doing it"]);
    var row = sheet.getLastRow();
    sheet.getRange(row, 1, 1, 5).setFontWeight("bold");

  // start of volunteering function
  function roles(querystring, title,assignmentnumber) {
    sheet.appendRow([title,assignmentnumber]);
    var row = sheet.getLastRow();
    sheet.getRange(row, 1, 1, 5).setFontWeight("bold");
    //var lastRow = sheet.getLastRow(row+1,1);

    var results = stmt.executeQuery('select `first_name` "First Name",`nickname` "Facebook Name", pd.order_id "Order ID"  from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=' + cell + ' AND `cc_location`="' + cc_location + '"  AND status in ("wc-processing", "wc-onhold", "wc-on-hold") AND pd.`cc_volunteer` LIKE "%' + querystring + '%" order by `first_name` ASC');
appendToSheet(sheet,results);




  } //end of volunteering

  // full options
  //help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening
  roles("director", "Event Director","1");
            sheet.appendRow([" ",""]);

      sheet.appendRow(["Promo Roles",""]);

  roles("tuesdaypromo", "Tuesday Promo","1");
  roles("wednesdaypromo", "Wednesday Promo","1");
              sheet.appendRow([" ",""]);

        sheet.appendRow(["Help around check-in time",""]);

  roles("check-in", "Check-in","2-3");
  roles("welcome", "Welcome Liaison","0-2");
  roles("pairing", "Pairing","2");
              sheet.appendRow([" ",""]);

          sheet.appendRow(["Help around announcements and cake",""]);

  roles("rounding", "Rounding Up","2");
  roles("precake", "Pre Announcements and Cake Admin","1");
  roles("announcements", "Do Announcements","1");
  roles("postcake", "Post Announcements and Cake Admin","1-2");
              sheet.appendRow([" ",""]);

          sheet.appendRow(["Help online afterwards",""]);

  roles("postpromo1", "Event Reporter","1");
  roles("postpromo2", "Friday Promo", "1");
              sheet.appendRow([" ",""]);

          sheet.appendRow(["Other",""]);

  roles("floorwalker", "Be floorwalker","0-Mark");

  //roles("happy to look after someone new to the clan", "Look after someone new to the Clan");
  //roles("happy to help introduce new climber", "Look after new climber");
  //roles("be event dir", "Be Event Director");
  setColoursFormat(sheet, "b1:b1000", "", "#e0ffff")







  //close SQL
  //results.close();
  //stmt.close();

  //end read data function
}
