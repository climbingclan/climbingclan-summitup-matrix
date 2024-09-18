


/**
 * Sends a POST request to update a specific order in WordPress using the WooCommerce API.
 *
 * @param {Object} data - The data to be sent to the WordPress API, representing order updates.
 * @param {number|string} order_id - The ID of the order to be updated.
 * @returns {number|string} Returns the ID of the updated order if successful, otherwise returns 'Error'.
 */
function pokeToWordPressOrders(data, order_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiUrl = "https://www." + apidomain + "/wp-json/wc/v3/orders/" + order_id

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  //console.log(responseData)
  if (response.getResponseCode() == 200) {
    return responseData.id;
  } else {
    return 'Error';
  }
}

/**
 * Sends a POST request to update a specific product in WordPress using the WooCommerce API.
 *
 * @param {Object} data - The data to be sent to the WordPress API, representing product updates.
 * @param {number|string} product_id - The ID of the product to be updated.
 * @returns {number|string} Returns the ID of the updated product if successful, otherwise returns 'Error'.
 */
function pokeToWordPressProducts(data, product_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiUrl = "https://www." + apidomain + "/wp-json/wc/v3/products/" + product_id

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  if (response.getResponseCode() == 200) {
    return responseData.id;
  } else {
    return 'Error';
  }
}

/**
 * Sends a POST request to update user meta data in WordPress using the WooCommerce API.
 *
 * @param {Object} data - The data to be sent to the WordPress API, representing user meta updates.
 * @param {number|string} user_id - The ID of the user whose meta data is to be updated.
 * @returns {Object|string} Returns the response data object if successful, otherwise returns 'Error'.
 */
function pokeToWooUserMeta(data, user_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiUrl = "https://www." + apidomain + "/wp-json/wc/v3/customers/" + user_id

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  if (response.getResponseCode() == 200) {
    return responseData;
  } else {
    return 'Error';
  }
}

/**
 * Sends a POST request to add a note to a specific order in WordPress using the WooCommerce API.
 *
 * @param {number|string} orderNumber - The number of the order to which the note is to be added.
 * @param {string} noteText - The text of the note to be added to the order.
 * @returns {number|string} Returns the ID of the note if successful, otherwise returns 'Error'.
 */
function pokeNoteToOrder(orderNumber, noteText) {
    var apiUrl = 'https://www.' + apidomain + '/wp-json/wc/v3/orders/' + orderNumber + '/notes';
  var encodedAuthInformation = Utilities.base64Encode(apiusername + ':' + apipassword);
  var headers = {'Authorization': 'Basic ' + encodedAuthInformation, 'Content-Type': 'application/json'};
  var payload = {'note': noteText};
  var options = {'method': 'post', 'headers': headers, 'payload': JSON.stringify(payload)};
  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  //console.log(response.getResponseCode())
  if (response.getResponseCode() == 201) {
    //console.log(responseData)
    return responseData.id;
  } else {
    return 'Error';
  }
}
