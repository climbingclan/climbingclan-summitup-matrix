/**
 * Reads event data from a database and updates a Google Sheet.
 * The function connects to a database, retrieves event data based on specific criteria,
 * and populates the 'Event List' sheet in the active Google Spreadsheet with the results.
 */
function readEvents() {
    // Establish database connection
    var conn = Jdbc.getConnection(url, username, password);
    var stmt = conn.createStatement();

    // Get the active spreadsheet and select the 'Dashboard' sheet
    var spreadsheet = SpreadsheetApp.getActive();
    var dashboardSheet = spreadsheet.getSheetByName('Dashboard');
    var cell = dashboardSheet.getRange('B5').getValues();

    // Query to select distinct events from the database
    var query = 'select distinct order_item_name, product_id from wp_order_product_customer_lookup ' +
                'where (memberbot_order_category LIKE "%indoor%" OR memberbot_order_category LIKE "%wednesday%" OR memberbot_order_category LIKE "%welcoming%") ' +
                'AND `cc_location`="' + cc_location + '" AND cc_attendance="pending" order by product_id desc';
    var results = stmt.executeQuery(query);

    // Select the 'Event List' sheet and clear its contents
    var eventListSheet = spreadsheet.getSheetByName('Event List');
    eventListSheet.clearContents();

    // Append the query results to the 'Event List' sheet
    appendToSheet(eventListSheet, results);

    // Close the SQL result set and statement
    results.close();
    stmt.close();
}
