function readLeadBelayTraining(stmt) {
  makeReport(stmt, {
    sheetName: "Lead Training",
    query: `
      SELECT DISTINCT 
        db.\`first_name\` "Name", 
        db.\`nickname\` "fbname", 
        CAST(db.\`stats_volunteer_for_numerator_cached\` AS UNSIGNED INTEGER) "Volunteering", 
        db.\`scores_volunteer_score_cached\` "Receptiveness", 
        db.\`skills-belaying\` "Belaying Skills", 
        db.\`climbing-indoors-toproping-grades\` "Indoor TR", 
        db.\`admin-wednesday-requests-notes\` as \`Requests and notes\`, 
        db.\`climbing-indoors-leading-grades\` "Indoor Lead", 
        db.\`skills-sport-climbing\` "Sport Skills" 
      FROM wp_member_db db 
      LEFT JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id 
      JOIN wp_member_db_scores s ON s.user_id = db.id 
      WHERE \`product_id\`=${cell} 
        AND \`cc_location\`="${cc_location}" 
        AND status IN ("wc-processing", "wc-onhold", "wc-on-hold") 
        AND db.\`skills-belaying\` IN ("Top-rope-belaying","learner-lead-belayer") 
        AND (db.\`skills-sport-climbing\` IS NULL OR db.\`skills-sport-climbing\` NOT LIKE "%Lead%") 
      ORDER BY CAST(db.\`stats_volunteer_for_numerator_cached\` AS UNSIGNED INTEGER) DESC, 
        db.\`skills-belaying\` DESC, 
        db.\`skills-trad-climbing\` ASC, 
        db.\`climbing-happy-to-supervise\` DESC
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
      { type: 'numberFormat', column: "Receptiveness", format: "0" },
      { type: 'columnWidth', column: "fbname", width: 150 },
      { type: 'columnWidth', column: "Requests and notes", width: 300 },
      { type: 'wrap', column: "Requests and notes" }
    ]
  });
}
