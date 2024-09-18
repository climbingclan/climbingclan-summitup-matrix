function readOutput(stmt) {
  makeReport(stmt, {
    sheetName: "Output",
    query: `
      SELECT 
        db.\`admin-first-timer-indoor\` AS New, 
        "Arrived", 
        "BaseC", 
        "Paired", 
        db.\`first_name\` as "Name", 
        db.\`last_name\` as "Surname", 
        db.\`skills-belaying\` as \`Belaying Skills\`, 
        pd.cc_volunteer AS "Volunteer", 
        db.\`admin-wednesday-requests-notes\` as \`Requests and notes\`, 
        db.scores_attendance_reliability_score_cached "Reliability%", 
        \`pd\`.\`order_id\` as \`Order ID\`, 
        \`pd\`.\`user_id\` as \`Clan ID\` 
      FROM wp_member_db db 
      JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id 
      WHERE \`product_id\`="${cell}" 
        AND \`cc_location\`="${cc_location}" 
        AND status IN ("wc-processing", "wc-onhold", "wc-on-hold") 
      ORDER BY \`db\`.\`first_name\` ASC
    `,
    formatting: [
      { type: 'color', column: "New", search: "Yes", color: colors.lightRed },
      { type: 'text', column: "New", search: "No", color: colors.grey },
      { type: 'color', column: "Arrived", search: "True", color: colors.lightGreen },
      { type: 'color', column: "BaseC", search: "True", color: colors.lightGreen },
      { type: 'color', column: "Paired", search: "True", color: colors.lightGreen },
      { type: 'color', column: "Belaying Skills", search: "learner-lead-belayer", color: colors.brightYellow },
      { type: 'color', column: "Belaying Skills", search: "lead-belayer", color: colors.green },
      { type: 'color', column: "Belaying Skills", search: "top", color: colors.lightBlue },
      { type: 'color', column: "Belaying Skills", search: "No-belaying", color: colors.lightRed },
      { type: 'text', column: "Volunteer", search: "none", color: colors.grey },
      { type: 'color', column: "Volunteer", search: "Check", color: colors.lightGreen },
      { type: 'color', column: "Volunteer", search: "welcome", color: colors.lightGreen },
      { type: 'color', column: "Volunteer", search: "pairing", color: colors.lightGreen },
      { type: 'color', column: "Volunteer", search: "Event", color: colors.pink },
      { type: 'text', column: "Volunteer", search: "Tuesday", color: colors.lightYellow },
      { type: 'text', column: "Volunteer", search: "Wednesday", color: colors.lightYellow },
      { type: 'color', column: "Volunteer", search: "Pre Cake", color: colors.yellow },
      { type: 'color', column: "Volunteer", search: "After Cake", color: colors.yellow },
      { type: 'color', column: "Volunteer", search: "Rounding", color: colors.yellow },
      { type: 'color', column: "Volunteer", search: "announce", color: colors.pink },
      { type: 'color', column: "Volunteer", search: "floor", color: colors.lightYellow },
      { type: 'color', column: "Volunteer", search: "Climbing Reporter", color: colors.lightGreen },
      { type: 'colorLessThanOrEqual', column: "Reliability%", value: "50", color: colors.yellow },
      { type: 'colorLessThanOrEqual', column: "Reliability%", value: "80", color: colors.pink },
      { type: 'colorLessThanOrEqual', column: "Reliability%", value: "90", color: colors.lightYellow },
      { type: 'columnWidth', column: "New", width: 31 },
      { type: 'columnWidth', column: "Name", width: 120 },
      { type: 'columnWidth', column: "Surname", width: 150 },
      { type: 'columnWidth', column: "Volunteer", width: 90 },
      { type: 'columnWidth', column: "Requests and notes", width: 300 },
      { type: 'wrap', column: "Requests and notes" }
    ]
  });

  // Add additional rows
  let sheet = SpreadsheetApp.getActive().getSheetByName("Output");
  sheet.appendRow(["","TRUE","TRUE","TRUE","Please ask latecomers and people not on the list to sign up right now at www.climbingclan.com and show you their confirmation page or confirmation email on their phone (or borrowed phone)"]);
  sheet.appendRow(["","TRUE","TRUE","TRUE","Once you've seen their confirmation page or email, you can add them below this line"]);
  sheet.appendRow(["","TRUE","TRUE","TRUE","Do this even with people who say they've already signed up"]);
  sheet.appendRow(["","TRUE","TRUE","TRUE","---------------------------------","----","-----","------"]);

  // Insert checkboxes
  let range = sheet.getRange("B2:D150");
  range.insertCheckboxes();
}
