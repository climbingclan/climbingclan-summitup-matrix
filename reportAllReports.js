function readData() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();

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

	for (const report of reports) {
		report(stmt);
	}

	stmt.close();
	conn.close();
}
