function readRoles(stmt) {
  makeReport(stmt, {
    sheetName: "Roles",
    query: `
      SELECT 
        'Event Director' AS Role, 
        1 AS 'Ideal number of people doing it'
      UNION ALL
      SELECT 'Tuesday Promo', 1
      UNION ALL
      SELECT 'Wednesday Promo', 1
      UNION ALL
      SELECT 'Check-in', 3
      UNION ALL
      SELECT 'Welcome Liaison', 2
      UNION ALL
      SELECT 'Pairing', 2
      UNION ALL
      SELECT 'Rounding Up', 2
      UNION ALL
      SELECT 'Pre Announcements and Cake Admin', 1
      UNION ALL
      SELECT 'Do Announcements', 1
      UNION ALL
      SELECT 'Post Announcements and Cake Admin', 2
      UNION ALL
      SELECT 'Event Reporter', 1
      UNION ALL
      SELECT 'Friday Promo', 1
      UNION ALL
      SELECT 'Be floorwalker', 1
    `,
    formatting: [
      { type: 'color', column: "Ideal number of people doing it", search: "", color: colors.lightBlue }
    ]
  });
}
