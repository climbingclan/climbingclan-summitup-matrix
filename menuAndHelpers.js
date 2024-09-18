function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mark an Attendance')
      .addItem('Mark Cancelled', 'markCancelled')
      .addItem('Mark Late Bail', 'markLateBail')
      .addItem('Mark No Show', 'markNoShow')
      .addSeparator()
      .addItem('Mark Duplicate', 'markDuplicate')
      .addSeparator()
      .addItem('Mark Attended', 'markAttended')
      .addSeparator()
     .addItem('Mark ALL Attended', 'markAttendedAndCloseEvent')
     // .addItem('NOT WORKING - Did not Register but showed', 'markNoRegisterShow')

      .addToUi();

   // create menu for dispatch  functions
  ui.createMenu('Assign a Role')
      .addItem('Assign Tuesday Promo', 'markVolunteerTuesdayPromo1')
      .addItem('Assign Wednesday Promo', 'markVolunteerWednesdayPromo1')
      .addSeparator()
      .addItem('Assign Check-in', 'markVolunteerCheckIn')
      .addItem('Assign Welcome Liaison', 'markVolunteerWelcomeLiaison')
      .addItem('Assign Pairing', 'markVolunteerPairing')
      .addItem('Assign Event Director', 'markVolunteerTripDirector')
      .addSeparator()

      .addItem('Assign Rounding Up', 'markVolunteerRoundingUp')
      .addItem('Assign Announcements', 'markVolunteerAnnouncements')

      .addItem('Assign Pre Cake Prep', 'markVolunteerPreCakeAdmin')
      .addItem('Assign After Cake Tidy', 'markVolunteerPostCakeAdmin')
      .addSeparator()
      .addItem('Assign FloorWalker', 'markVolunteerFloorwalker')
      .addItem('Assign Trainer on Pre-agreed Training', 'markVolunteerWednesdayTrainer')
      .addSeparator()
      .addItem('Assign Event Reporter (postpromo1)', 'markVolunteerPostPromo1')

      .addItem('Assign Friday Promo', 'markVolunteerPostPromo2')
      .addSeparator()
      //.addItem('Assign EarlyPromo1', 'markVolunteerEarlyPromo1')
      .addItem('Assign Summit Up Saturday Promo', 'markVolunteerSummitUpPromo')
      //.addItem('Assign Tuesday Night Admin', 'markVolunteerTuesdayNightAdmin')
      //.addItem('Assign TuesdayPromo2', 'markVolunteerTuesdayPromo2')
            .addSeparator()
      .addItem('Unassign Role', 'markVolunteerClear')

      .addToUi();

  ui.createMenu('Refresh Matrix')
      .addItem('Refresh All', 'readData')
      .addSeparator()
      .addItem('Refresh Output', 'refreshOutput')
      //.addItem('Refresh Stats', 'refreshStats')
      .addItem('Refresh Lead Training', 'refreshLeadBelayTraining')
      .addItem('Refresh TR Belay Training', 'refreshTopRopeTraining')
      .addItem('Refresh Event Listings', 'refreshEventListings')
      //.addItem('Refresh Roles', 'refreshRoles')
      //.addItem('Refresh Volunteer Intent', 'refreshVolunteerIntent')
      .addSeparator()
      .addItem('Refresh Volunteering', 'refreshVolunteering')
      //.addItem('Refresh Badges', 'refreshBadges')
      .addToUi();

        ui.createMenu('Badges & bands')
      .addItem('Mark Given Badge', 'markGivenBadge')
            .addItem('Mark Given Badge', 'markGivenBand')

      .addItem('Refresh Badges', 'badgesData')



      .addToUi();

        ui.createMenu('Give Volunteer Competencies')
      .addItem('Check-In - SignOff', 'giveCheckIn')
      .addItem('Announcements - SignOff', 'giveAnnouncements')
      .addItem('Floorwalker - SignOff', 'giveFloorwalker')
        .addSeparator()
      .addItem('Pairing - Eager To Learn', 'givePairingToLearn')
      .addItem('Pairing - In Training', 'givePairingTraining')
      .addItem('Pairing - SignOff', 'givePairing')
      .addSeparator()
      .addItem('Trip Director - Eager To Learn', 'giveTripDirectorToLearn')
      .addItem('Trip Director - In Training', 'giveTripDirectorTraining')
      .addItem('Trip Director - SignOff', 'giveTripDirector')
            .addItem('Trip Director - Remove', 'giveTripDirectorToRemove')

      //      .addSeparator()

      //.addItem('Set Learner Lead Belayer', 'setSkillsDownLearnerLeadBelaying')
      //.addItem('Set Non Belayer', 'setSkillsDownNoBelaying')


      .addToUi();

}

function refreshOutput() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readOutput(stmt);
  stmt.close();
  conn.close();
}

function refreshStats() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readStats(stmt);
  stmt.close();
  conn.close();
}

function refreshLeadBelayTraining() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readLeadBelayTraining(stmt);
  stmt.close();
  conn.close();
}

function refreshTopRopeTraining() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readTopRopeTraining(stmt);
  stmt.close();
  conn.close();
}

function refreshEventListing() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readEventListing(stmt);
  stmt.close();
  conn.close();
}

function refreshVolunteerIntent() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readVolunteerIntent(stmt);
  stmt.close();
  conn.close();
}

function refreshRoles() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readRoles(stmt);
  stmt.close();
  conn.close();
}

function refreshVolunteering() {
  readVolunteerData();
}

function refreshBadges() {
  readBadgesData();
}

function readVolunteerData() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  readVolunteers(stmt);

  stmt.close();
  conn.close();
}

function readBadgesData() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  readBadgesNeeded(stmt);
  readBandsNeeded(stmt);
  readBadgesGiven(stmt);

  stmt.close();
  conn.close();
}

function refreshEventListings() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  readEventListing(stmt);

  stmt.close();
  conn.close();
}

