
function readBandsNeeded(stmt) {
    makeReport(stmt, {
        sheetName: "Bands",
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
        AND db.\`skills-belaying\` = "lead-belayer"
        AND ((db.\`stats_volunteer_for_numerator_cached\`>=5) OR (db.\`stats_volunteer_for_numerator_cached\`=4 AND pd.cc_volunteer<>"none"))
        AND ((db.milestones_5_band IS NULL) OR (db.milestones_5_band ="due"))
      ORDER BY db.\`first_name\`, CAST(db.stats_volunteer_for_numerator_cached AS UNSIGNED INTEGER) DESC
    `,
        formatting: [
            { type: 'color', column: "Given Badge", search: "", color: colors.lightBlue },
            { type: 'numberFormat', column: "Volunteered For", format: "0" },
            { type: 'columnWidth', column: "Facebook Name", width: 150 },
            { type: 'columnWidth', column: "Clan ID", width: 100 },
        ],
        title: "People who need bands"
    });

    // Add checkboxes to column A from row 2 downwards
    let sheet = SpreadsheetApp.getActive().getSheetByName("Bands");
    let lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        let range = sheet.getRange(2, 1, lastRow - 1, 1);
        range.insertCheckboxes();
    }
}
