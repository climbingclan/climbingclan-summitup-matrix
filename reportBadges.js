

function readBadgesNeeded(stmt) {
  makeReport(stmt, {
    sheetName: "Badges",
    query: `
      SELECT 
        "Given Badge" AS "Given Badge",
        db.\`first_name\` AS "First Name",
        db.\`nickname\` AS "Facebook Name",
        db.stats_volunteer_for_numerator_cached AS "Volunteered For",
        db.id AS "Clan ID"
      FROM wp_member_db db
      JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id
      WHERE product_id=${cell}
        AND \`cc_location\`="${cc_location}"
        AND status IN ("wc-processing", "wc-onhold", "wc-on-hold")
        AND ((db.\`stats_volunteer_for_numerator_cached\`>=3) OR (db.\`stats_volunteer_for_numerator_cached\`=2 AND pd.cc_volunteer<>"none"))
        AND ((db.milestones_3_badge IS NULL) OR (db.milestones_3_badge ="due"))
      ORDER BY db.\`first_name\`, CAST(db.stats_volunteer_for_numerator_cached AS UNSIGNED INTEGER) DESC
    `,
    formatting: [
      { type: 'color', column: "Given Badge", search: "", color: colors.lightGreen },
      { type: 'numberFormat', column: "Volunteered For", format: "0" },
      { type: 'columnWidth', column: "Facebook Name", width: 150 },
      { type: 'columnWidth', column: "Clan ID", width: 100 },
    ],
    title: "People who need badges"
  });

  // Add checkboxes to column A from row 2 downwards
  let sheet = SpreadsheetApp.getActive().getSheetByName("Badges");
  let lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    let range = sheet.getRange(2, 1, lastRow - 1, 1);
    range.insertCheckboxes();
  }
}


function readBadgesGiven(stmt) {
  makeReport(stmt, {
    sheetName: "Badges Given",
    query: `
      SELECT 
        "Given Badge" AS "Given Badge",
        db.\`first_name\` AS "First Name",
        db.\`nickname\` AS "Facebook Name",
        db.milestones_3_badge_marked_given_by AS "Given by",
        FROM_UNIXTIME((db.milestones_3_badge_marked_given_at)/1000, "%d %M %Y") AS "Given on"
      FROM wp_member_db db
      JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id
      WHERE product_id=${cell}
        AND \`cc_location\`="${cc_location}"
        AND status IN ("wc-processing", "wc-onhold", "wc-on-hold")
        AND db.milestones_3_badge="given"
      ORDER BY db.\`first_name\`, CAST(db.stats_volunteer_for_numerator_cached AS UNSIGNED INTEGER) DESC
    `,
    formatting: [
      { type: 'color', column: "Given Badge", search: "", color: colors.lightYellow },
      { type: 'columnWidth', column: "Facebook Name", width: 150 },
      { type: 'columnWidth', column: "Given by", width: 150 },
      { type: 'columnWidth', column: "Given on", width: 120 },
    ],
    title: "People who have been given badges"
  });
}
