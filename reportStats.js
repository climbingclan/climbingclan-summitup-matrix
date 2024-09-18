function readStats(stmt) {
  makeReport(stmt, {
    sheetName: "Stats",
    query: `
      SELECT 
        \`skills-belaying\` AS \`Belaying Skills\`, 
        COUNT(*) AS \`How many this week\` 
      FROM wp_member_db db 
      JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id 
      WHERE \`product_id\` = ${cell} 
        AND \`cc_location\`="${cc_location}" 
        AND \`status\` IN ("wc-processing", "wc-onhold", "wc-on-hold") 
      GROUP BY \`skills-belaying\`
    `,
    formatting: [
      { type: 'colorLessThanOrEqual', column: "How many this week", value: "5", color: colors.yellow },
      { type: 'colorLessThanOrEqual', column: "How many this week", value: "2", color: colors.lightYellow },
      { type: 'colorLessThanOrEqual', column: "How many this week", value: "0", color: colors.pink },
      { type: 'color', column: "Belaying Skills", search: "learner-lead-belayer", color: colors.brightYellow },
      { type: 'color', column: "Belaying Skills", search: "lead-belayer", color: colors.green },
      { type: 'color', column: "Belaying Skills", search: "top", color: colors.lightBlue },
      { type: 'color', column: "Belaying Skills", search: "No-belaying", color: colors.lightRed },
      { type: 'numberFormat', column: "How many this week", format: "0" },
      { type: 'columnWidth', column: "Belaying Skills", width: 150 }
    ]
  });
}
