function markGivenBadge() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Badges & Bands');
  var active_range = sheet.getActiveRange();
  var currentRow = active_range.getRowIndex();
  console.log(currentRow);

  if (currentRow <= 2 || currentRow >= 100) {
    Browser.msgBox('Select an actual signup', Browser.Buttons.OK);
    return;
  }

  var user_id = sheet.getRange(currentRow, 5, 1, 1).getValue();
  var first_name = sheet.getRange(currentRow, 2, 1, 1).getValue();

  console.log(user_id);
  
  if (user_id === "" || user_id === "user_id") {
    Browser.msgBox('No User ID Found', Browser.Buttons.OK);
    return;
  } 

  let cc_attendance_setter = Session.getActiveUser().getEmail();

  let metakey = "milestones_3_marked";
  let metavalue = "given";
  let metakey2 = "milestones_3_badge";
  let datetime = Date.now();

  var data = {
    "meta_data": [
      {"key": metakey, "value": metavalue}, 
      {"key": metakey2, "value": metavalue}, 
      {"key": "milestones_3_badge_marked_given_at", "value": datetime}, 
      {"key": "milestones_3_badge_marked_given_by", "value": cc_attendance_setter}
    ]
  };

  let returndata = pokeToWooUserMeta(data, user_id);
  returndata = returndata.getContentText();
  returndata = JSON.parse(returndata);

  let search = returndata.meta_data.find(({key}) => key == "milestones_3_badge")?.value;

  Logger.log(search); 

  if (search === metavalue) {
    sheet.getRange(currentRow, 4, 1, 1).setValue("Given");
    sheet.getRange(currentRow, 1, 1, 1).setValue("TRUE");
    return metavalue;
  } else {
    Logger.log("ERROR" + search);
    return "ERROR" + search;
  }
}

function markGivenBand() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Badges & Bands');
  var active_range = sheet.getActiveRange();
  var currentRow = active_range.getRowIndex();
  console.log(currentRow);

  if (currentRow <= 2 || currentRow >= 100) {
    Browser.msgBox('Select an actual signup', Browser.Buttons.OK);
    return;
  }

  var user_id = sheet.getRange(currentRow, 5, 1, 1).getValue();
  var first_name = sheet.getRange(currentRow, 2, 1, 1).getValue();

  console.log(user_id);
  
  if (user_id === "" || user_id === "user_id") {
    Browser.msgBox('No User ID Found', Browser.Buttons.OK);
    return;
  } 

  let cc_attendance_setter = Session.getActiveUser().getEmail();

  let metakey = "milestones_5_marked";
  let metavalue = "given";
  let metakey2 = "milestones_5_band";
  let datetime = Date.now();

  var data = {
    "meta_data": [
      {"key": metakey, "value": metavalue}, 
      {"key": metakey2, "value": metavalue}, 
      {"key": "milestones_5_band_marked_given_at", "value": datetime}, 
      {"key": "milestones_5_band_marked_given_by", "value": cc_attendance_setter}
    ]
  };

  let returndata = pokeToWooUserMeta(data, user_id);
  returndata = returndata.getContentText();
  returndata = JSON.parse(returndata);

  let search = returndata.meta_data.find(({key}) => key == "milestones_5_band")?.value;

  Logger.log(search); 

  if (search === metavalue) {
    sheet.getRange(currentRow, 4, 1, 1).setValue("Given");
    sheet.getRange(currentRow, 1, 1, 1).setValue("TRUE");
    return metavalue;
  } else {
    Logger.log("ERROR" + search);
    return "ERROR" + search;
  }
}
