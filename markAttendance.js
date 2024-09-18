
function markAttended(){
  const attendancetype = "Attended"
  const attendanceshow = "Attended"
  const metakey = "cc_attendance"
  const orderstatus= "completed"
  const metavalue = "attended"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markCancelled(){
  const attendancetype = "Cancel"
  const attendanceshow = "Cancelled"
  const metakey = "cc_attendance"
  const orderstatus= "completed"
  const metavalue = "cancelled"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markNoShow(){
  const attendancetype = "NoShow"
  const attendanceshow = "NoShow"
  const orderstatus= "completed"
  const metakey = "cc_attendance"
  const metavalue = "noshow"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markDuplicate(){
  const attendancetype = "Duplicate"
  const attendanceshow = "Duplicated"
  const orderstatus= "completed"
  const metakey = "cc_attendance"
  const metavalue = "duplicate"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markLateBail(){
  const attendancetype = "Late Bail"
  const attendanceshow = "Late Bail"
  const orderstatus= "completed"
  const metakey = "cc_attendance"
  const metavalue = "late-bail"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markNoRegisterShow(){
  const attendancetype = "No Register Show"
  const attendanceshow = "No Register Show"
  const orderstatus= "completed"
  const metakey = "cc_attendance"
  const metavalue = "noregistershow"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markVolunteerLateBail(){
  const attendancetype = "Late Bail"
  const attendanceshow = "Late Bail"
  const metakey = "cc_attendance"
  const metavalue = "late-bail"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markVolunteerCancelled(){
  const attendancetype = "Cancel"
  const attendanceshow = "Cancelled"
  const metakey = "cc_attendance"
  const orderstatus= "completed"
  const metavalue = "cancelled"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markVolunteerNoShow(){
  const attendancetype = "NoShow"
  const attendanceshow = "NoShow"
  const orderstatus= "completed"
  const metakey = "cc_attendance"
  const metavalue = "noshow"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markVolunteerAttended(){
  const attendancetype = "Attended"
  const orderstatus= "completed"
  const metavalue = "attended"

  markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue)
}

function markAttendance(attendancetype, attendanceshow, orderstatus, metakey, metavalue ) {

 let spreadsheet = SpreadsheetApp.getActive();
 let sheet = spreadsheet.getSheetByName('Output');
 console.log(sheet.getName())
  let active_range = sheet.getActiveRange();
  let currentRow = active_range.getRowIndex();
  //let currentRow = "14";
  console.log(currentRow);


  if(currentRow <=1){Browser.msgBox('Select an actual signup', Browser.Buttons.OK); return;}
    if(currentRow >=200){Browser.msgBox('Select an actual signup', Browser.Buttons.OK); return;}


  let order_id = sheet.getRange(currentRow, 11,1,1).getValue();  /// get submission ID 1 BV ( was 67)
  let first_name = sheet.getRange(currentRow, 5,1,1).getValue();  /// get submission ID 1 BV ( was 67)

  console.log(order_id);

// Input validation - prevent attendance being marked from any sheet other than "Output"
if(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() != 'Output'){Browser.msgBox('Attendance can only be marked from the "Output" sheet', Browser.Buttons.OK); return;}

if(order_id === ""){Browser.msgBox('No Order ID Found', Browser.Buttons.OK); return;}

  if (Browser.msgBox("Mark " + attendancetype + " on " +first_name + "'s place? \n Order " + order_id, Browser.Buttons.OK_CANCEL) == "ok") {



let cc_attendance_setter =  Session.getActiveUser().getEmail();

let data = {"meta_data": [
    {"key": metakey,
    "value": metavalue},
    {"key": "cc_attendance_set_by",
    "value": cc_attendance_setter }
  ],
   "status": orderstatus
};
console.log(orderstatus);

  let encodedAuthInformation = Utilities.base64Encode(apiusername+ ":" + apipassword);
  let headers = {"Authorization" : "Basic " + encodedAuthInformation};
  let options = {
  'method' : 'post',
  'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(data)
};
let url="https://www."+ apidomain + "/wp-json/wc/v3/orders/" + order_id

let response = UrlFetchApp.fetch(url, options);
  console.log(response);


//remove old data


let blankArray =[[attendanceshow,order_id,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,]];  /// set a blank variable to delete row (45 values)
sheet.getRange(currentRow, 6,1,45).setValues(blankArray);   // paste the blank variables into the cells to delete contents
}

}



