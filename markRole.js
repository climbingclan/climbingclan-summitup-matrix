function markVolunteerTuesdayPromo1(){
let cc_volunteer_old = "tuesdaypromo1"
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  //Logger.log(volunteer)
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerWednesdayPromo1(){
  let cc_volunteer_old = "wednesdaypromo1"
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerPostPromo1(){
  let cc_volunteer_old = "postpromo1"
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerCheckIn(){
  let cc_volunteer_old = "check-in";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerWelcomeLiaison(){
  let cc_volunteer_old = "welcome_liaison";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerPairing(){
  let cc_volunteer_old = "pairing";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerTripDirector(){
  let cc_volunteer_old = "trip_director";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerAnnouncements(){
  let cc_volunteer_old = "announcements";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerRoundingUp(){
  let cc_volunteer_old = "rounding_up";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerPreCakeAdmin(){
  let cc_volunteer_old = "precakeadmin";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerPostCakeAdmin(){
  let cc_volunteer_old = "postcakeadmin";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerFloorwalker(){
  let cc_volunteer_old = "floorwalker";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerEarlyPromo1(){
  let cc_volunteer_old = "earlypromo1";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerSummitUpPromo(){
  let cc_volunteer_old = "Summit Up Promo";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerPostPromo2(){
  let cc_volunteer_old = "postpromo2";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerTuesdayNightAdmin(){
  let cc_volunteer_old = "tuesdaynightadmin";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerWednesdayTrainer(){
  let cc_volunteer_old = "Wednesday Evening Trainer";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
}

function markVolunteerClear(){
    if (Browser.msgBox("This wont notify the person automatically that their roles has been cancelled - you will want to do that", Browser.Buttons.OK_CANCEL) == "ok") {
  let cc_volunteer_old = "none";
  let volunteer = searchEmailsSheet(cc_volunteer_old);
  assignRole(volunteer, cc_volunteer_old);
    }
}

/**
 * Assigns a volunteer role to a selected individual in a spreadsheet and updates the corresponding WordPress order.
 * This function retrieves the selected row in the 'Volunteering' sheet, checks for the validity of the row and order ID,
 * then prompts the user for confirmation before assigning the volunteer role. It updates the WordPress order with the
 * new role details and adds a note to the order. The function also handles errors during the WordPress data update.
 *
 * @param {Object} volunteer - An object containing volunteer details and role information.
 *   Expected properties:
 *   - cc_volunteer_old: legacy volunteer role name
 *   - cc_volunteer: New volunteer role
 *   - cc_volunteer_role_description_sentence: Description of the role
 *   - cc_volunteer_role_description_time: Time information for the role
 *   - cc_volunteer_reminder_time: Reminder time for the role
 *   - cc_volunteer_role_description_post_event_sentence: Post-event description of the role
 *   - cc_volunteer_role_instructions_url: URL for role instructions
 * @param {string} cc_volunteer_old - A string representing the previous volunteer role (deprecated, use volunteer.cc_volunteer_old instead).
 *
 * @throws {Error} Throws an error if the WordPress order update or note addition fails.
 */
function assignRole(volunteer, cc_volunteer_old ) {

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Volunteering');
  var active_range = sheet.getActiveRange();
  var currentRow = active_range.getRowIndex();
  //var currentRow = "5";
  console.log("Row " + currentRow);


  if(currentRow <=1){Browser.msgBox('Select an actual signup', Browser.Buttons.OK); return;}
    if(currentRow >=100){Browser.msgBox('Select an actual signup', Browser.Buttons.OK); return;}


  var order_id = Number(sheet.getRange(currentRow, 19,1,1).getValue());  /// get submission ID 1 BV ( was 67)
  var first_name = String(sheet.getRange(currentRow, 1,1,1).getValue());  /// get submission ID 1 BV ( was 67)

  console.log(order_id);
  
// Input validation - prevent roles being allocated from any sheet other than "Volunteering"
if(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() != 'Volunteering'){Browser.msgBox('Roles can only be assigned from the "Volunteering" sheet', Browser.Buttons.OK); return;}

if(order_id === "" || order_id ===  "order_id"){Browser.msgBox('No Order ID Found', Browser.Buttons.OK); return;} 
if (Browser.msgBox("Assign " + volunteer.cc_volunteer + " to " +first_name + "? \n Order " + order_id, Browser.Buttons.OK_CANCEL) == "ok") { 

let cc_role_assigner =  Session.getActiveUser().getEmail();
let noteToOrder = "#!Assigned Role by: " + cc_role_assigner
let currentTime = new Date().getTime();

let data = {"meta_data": [
    {"key": "cc_volunteer_old",
    "value": volunteer.cc_volunteer_old}, 
    {"key": "cc_volunteer",
    "value": volunteer.cc_volunteer}, 
    {"key": "cc_volunteer_role_description_sentence",
    "value": volunteer.cc_volunteer_role_description_sentence}, 
    {"key": "cc_volunteer_role_description_time",
    "value": volunteer.cc_volunteer_role_description_time}, 
    {"key": "cc_volunteer_reminder_time",
    "value": volunteer.cc_volunteer_reminder_time}, 
    {"key": "cc_volunteer_role_description_post_event_sentence",
    "value": volunteer.cc_volunteer_role_description_post_event_sentence}, 
    {"key": "cc_volunteer_role_instructions_url",
    "value": volunteer.cc_volunteer_role_instructions_url}, 
    {"key": "cc_volunteer_role_assigned_by",
    "value": cc_role_assigner },
    {"key": "cc_volunteer_role_assigned_at",
    "value": currentTime }
  ], 
//   "status": orderstatus
};


let orderDataOutcome = pokeToWordPressOrders(data, order_id)
let orderNoteOutcome = pokeNoteToOrder(order_id, noteToOrder)

if (orderDataOutcome === "Error" || orderNoteOutcome === "Error") {
  throw new Error("One or more variables contain an 'Error' string.")
  return
}
  
var blankArray =[[volunteer.cc_volunteer]];  
sheet.getRange(currentRow, 3,1,1).setValues(blankArray);   // paste the blank variables into the cells to delete contents
}

return;
}


