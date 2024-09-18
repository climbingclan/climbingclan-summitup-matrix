function readData() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  const reports = [
    readOutput,
    readVolunteers,
    readLeadBelayTraining,
    readEventListing,
    readBadgesNeeded,
    readBandsNeeded,
    readTopRopeTraining,
    readStats,
    readRoles,
    readVolunteerIntent,
  ];

  reports.forEach(report => report(stmt));

  stmt.close();
  conn.close();
}

