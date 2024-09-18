function readEventListing(stmt) {
  const useNewConnection = !stmt;
  
  if (useNewConnection) {
    var conn = Jdbc.getConnection(url, username, password);
    stmt = conn.createStatement();
  }

  makeReport(stmt, {
    sheetName: "Event List",
    query: `
      SELECT DISTINCT 
        order_item_name AS "Trip Name", 
        product_id AS "ID"
      FROM wp_order_product_customer_lookup 
      WHERE (memberbot_order_category LIKE "%indoor%" 
        OR memberbot_order_category LIKE "%wednesday%" 
        OR memberbot_order_category LIKE "%welcoming%")
        AND cc_location="${cc_location}" 
        AND cc_attendance="pending" 
      ORDER BY product_id DESC
    `,
    formatting: []
  });

  if (useNewConnection) {
    stmt.close();
    conn.close();
  }
}

function eventListing() {
  readEventListing();
}
