//      .addItem('Finalise Assignments', 'test')
// MarkAttended and Close
// https://woocommerce.github.io/woocommerce-rest-api-docs/#batch-update-orders

// Get all IDs
// SELECT order_id  FROM wp_order_product_customer_lookup where product_id=$id AND status='wc-processing' AND cc_attendance='pending'
// Make on hold
function sendAllCragAssignments(tellme) {
    // Define the action based on the passed argument
    const action = tellme;
    
    // Establish a database connection
    const sconn = Jdbc.getConnection(url, username, password);
    const sstmt = sconn.createStatement();

    // Retrieve product_id from the 'Dashboard' sheet
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getSheetByName('Dashboard');
    const product_id = sheet.getRange('B5').getValues();

    // Execute a query to find orders with specific criteria
    const order_results = sstmt.executeQuery('SELECT distinct order_id from wp_order_product_customer_lookup where product_id="' + product_id + '"  AND `cc_location`="' + cc_location + '" AND status="wc-processing" AND cc_attendance="pending" LIMIT 99');
    
    // Get the email of the active user and current Unix time
    const active_user = Session.getActiveUser().getEmail();
    const currentUnixTime = Date.now();

    // Process each order result
    while (order_results.next()) {
        let scores_arr = [];
        for (let col = 0; col < 1; col++) {
            scores_arr.push(order_results.getString(col + 1));
        }
        console.log(scores_arr);

        const order_id = scores_arr[0];

        // If action is 'close', prepare and log the data
        if (action === "close") {
            const data = {
                "status": "completed",
                "meta_data": [
                    {
                        "key": "cc_attendance_set_by",
                        "value": active_user
                    },
                    {
                        "key": "cc_attendance_set_at",
                        "value": currentUnixTime
                    },
                    {
                        "key": "cc_attendance",
                        "value": "attended"
                    }
                ]
            };
            Logger.log(data);
            pokeToWordPressOrders(data, order_id);
        }
    }

    // If action is 'close', set product status to 'private'
    if (action === "close") {
      try{
        const data = {
            "status": "private",
            "meta_data": [
                {
                    "key": "cc_post_set_private_set_by",
                    "value": active_user
                },
                {
                    "key": "cc_post_set_private_set_at",
                    "value": currentUnixTime
                }
            ]
        };
        console.log("data");
        pokeToWordPressProducts(data, product_id);
      } catch (e){
        console.log(e)
      }
    } else {
        console.log("All done");
    }
}


//console.log(cell);


function markAttendedAndCloseEvent() {
    if (Browser.msgBox("This will mark all those who haven't been cancelled as ATTENDED ", Browser.Buttons.OK_CANCEL) == "ok") {
        if (Browser.msgBox("This should be done on Wednesday evening after climbing", Browser.Buttons.OK_CANCEL) == "ok") {
            if (Browser.msgBox("This cannot be undone", Browser.Buttons.OK_CANCEL) == "ok") {
                SpreadsheetApp.getActiveSpreadsheet().toast('Signup now being closed', 'Status', 5);

                sendAllCragAssignments("close");
                SpreadsheetApp.getActiveSpreadsheet().toast('Mark all attended completed successfully', 'Status', 5);
                SpreadsheetApp.getActiveSpreadsheet().toast('You can now close this window ', 'Status', 5);
                SpreadsheetApp.getActiveSpreadsheet().toast('Thanks for all you have done for the Clan this evening', 'Status', 150);

                try {
                    Browser.msgBox("Signup closed successfully", Browser.Buttons.OK);
                    runWednesdayMatrix();
                } catch (e) {
console.log(e)                }

                // Additional code here if needed
            }
        }
    }
}


