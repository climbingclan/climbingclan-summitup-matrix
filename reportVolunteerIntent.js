function readVolunteerIntent(stmt) {
  makeReport(stmt, {
    sheetName: "Intent",
    query: `
      SELECT DISTINCT 
        \`first_name\` AS "Name",
        \`nickname\` AS "fbname",
        admin_parthian_clan_join_admin_team AS "volunteer?",
        CAST(\`stats_volunteer_for_numerator_cached\` AS UNSIGNED INTEGER) AS "Volunteering",
        \`scores_volunteer_score_cached\` AS "Receptiveness",
        scores_volunteer_reliability_score_cached AS "V Reliability%",
        stats_attendance_indoor_wednesday_attended_cached AS "attended indoors",
        scores_attendance_reliability_score_cached AS "Attendance %",
        \`skills-belaying\` AS "Belaying Skills",
        \`climbing-indoors-skills-passing-on\` AS "Skills passing on",
        cc_compliance_last_date_of_climbing AS "last climbed"
      FROM wp_member_db db 
      WHERE admin_parthian_clan_join_admin_team IS NOT NULL 
        AND cc_compliance_last_date_of_climbing BETWEEN DATE_SUB(NOW(), INTERVAL 2 MONTH) AND NOW() 
      ORDER BY admin_parthian_clan_join_admin_team ASC, 
        CAST(\`stats_volunteer_for_numerator_cached\` AS UNSIGNED INTEGER) DESC, 
        CAST(\`stats_attendance_indoor_wednesday_attended_cached\` AS UNSIGNED INTEGER) DESC, 
        CAST(\`scores_volunteer_score_cached\` AS UNSIGNED INTEGER) DESC
    `,
    formatting: [
      { type: 'colorLessThanOrEqual', column: "Volunteering", value: "0", color: colors.pink },
      { type: 'colorLessThanOrEqual', column: "Volunteering", value: "2", color: colors.lightYellow },
      { type: 'colorLessThanOrEqual', column: "Volunteering", value: "5", color: colors.yellow },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "10", color: colors.pink },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "20", color: colors.lightYellow },
      { type: 'colorLessThanOrEqual', column: "Receptiveness", value: "30", color: colors.yellow },
      { type: 'color', column: "Belaying Skills", search: "learner-lead-belayer", color: colors.brightYellow },
      { type: 'color', column: "Belaying Skills", search: "lead-belayer", color: colors.green },
      { type: 'color', column: "Belaying Skills", search: "top", color: colors.lightBlue },
      { type: 'color', column: "Belaying Skills", search: "No-belaying", color: colors.lightRed },
      { type: 'numberFormat', column: "Volunteering", format: "0" },
      { type: 'columnWidth', column: "fbname", width: 150 },
      { type: 'columnWidth', column: "Skills passing on", width: 300 },
      { type: 'wrap', column: "Skills passing on" }
    ]
  });
}
