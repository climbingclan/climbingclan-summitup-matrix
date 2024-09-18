function readVolunteers(stmt) {
  makeReport(stmt, {
    sheetName: "Volunteering",
    title: "Volunteers",
    query: `
      SELECT 
        db.\`first_name\` AS "First Name",
        db.\`last_name\` AS "Last Name",
        pd.cc_volunteer AS "Selected Roles",
        volunteer_indoor_preevent_facebook_promo AS "FB promo",
        volunteer_indoor_event_reporter AS "Rprtr",
        volunteer_indoor_check_in AS "Check-in",
        volunteer_indoor_welcome_liaison AS "WelcLn",
        volunteer_indoor_rounding_up AS "Rnd Up",
        volunteer_indoor_before_cake_tidying AS "PreCake",
        volunteer_indoor_after_cake_tidying AS "PostCake",
        volunteer_indoor_give_announcements AS "Anncments Vol",
        competency_indoor_announcements AS "Anncments Compet",
        competency_indoor_pairing AS "Pair",
        competency_indoor_trip_director AS "Event Dir",
        scores_volunteer_score_cached AS "Receptiveness",
        scores_volunteer_reliability_score_cached AS "V Reliability%",
        stats_attendance_indoor_wednesday_attended_cached AS "Attended",
        \`admin-wednesday-requests-notes\` AS \`Requests and notes\`,
        pd.order_id AS "Order ID",
        pd.user_id AS "User ID",
        competency_indoor_floorwalker AS "Flwlkr",
        competency_indoor_skillsharer,
        competency_indoor_checkin
      FROM wp_member_db db
      JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id
      JOIN wp_member_db_volunteering vl ON pd.user_id = vl.id
      WHERE product_id = ${cell}
        AND \`cc_location\` = "${cc_location}"
        AND status IN ("wc-processing", "wc-onhold", "wc-on-hold")
        AND stats_attendance_indoor_wednesday_attended_cached >= 2
        AND cc_compliance_last_date_of_climbing BETWEEN DATE_SUB(NOW(), INTERVAL 2 MONTH) AND NOW()
      ORDER BY 
        FIELD(pd.cc_volunteer, "none", "Tuesday Facebook Promo", "tuesdaypromo2", "Wednesday Facebook Promo", "wednesdaypromo2", "check in", "welcome liaison", "pairing", "rounding up", "Pre-cake Preparation", "announcements", "after cake tidy", "event director", "Event Reporter", "Wednesday Evening Trainer", "floorwalker", "Friday Facebook Promo", "Summit Up Promo") ASC,
        pd.cc_volunteer ASC,
        CAST(db.scores_volunteer_score_cached AS UNSIGNED INTEGER) ASC,
        CAST(db.stats_attendance_indoor_wednesday_attended_cached AS UNSIGNED INTEGER) DESC,
        volunteer_indoor_event_reporter,
        volunteer_indoor_check_in
    `,
    formatting: [
      { type: 'numberFormat', column: "Receptiveness", format: "0" },
      { type: 'colorLessThanOrEqual', column: "V Reliability%", value: "50", color: colors.yellow },
      { type: 'colorLessThanOrEqual', column: "V Reliability%", value: "80", color: colors.pink },
      { type: 'colorLessThanOrEqual', column: "V Reliability%", value: "90", color: colors.lightYellow },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "10", color: colors.pink },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "20", color: colors.lightYellow },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "30", color: colors.yellow },
      { type: 'color', column: "Selected Roles", search: "none", color: colors.lightGreen },
      { type: 'color', column: "Selected Roles", search: "Selected", color: colors.white },
      { type: 'color', column: "Selected Roles", search: "", color: colors.lightBlue },
      { type: 'text', column: "FB promo", search: "No", color: colors.grey },
      { type: 'text', column: "Rprtr", search: "No", color: colors.grey },
      { type: 'text', column: "Check-in", search: "No", color: colors.grey },
      { type: 'text', column: "WelcLn", search: "No", color: colors.grey },
      { type: 'text', column: "Rnd Up", search: "No", color: colors.grey },
      { type: 'text', column: "PreCake", search: "No", color: colors.grey },
      { type: 'text', column: "PostCake", search: "No", color: colors.grey },
      { type: 'text', column: "Anncments Vol", search: "No", color: colors.grey },
      { type: 'text', column: "Anncments Compet", search: "No", color: colors.grey },
      { type: 'text', column: "Pair", search: "No", color: colors.grey },
      { type: 'text', column: "Event Dir", search: "No", color: colors.grey },
      { type: 'text', column: "Flwlkr", search: "No", color: colors.grey },
      { type: 'columnWidth', column: "Requests and notes", width: 200 },
      { type: 'wrap', column: "Requests and notes" }
    ]
  });
}


function volunteerData() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  readVolunteers(stmt);

  stmt.close();
  conn.close();

  return "done";
}

