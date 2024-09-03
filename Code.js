var server = '18.168.242.164';
var port = 3306;
var dbName = 'bitnami_wordpress';
var username = 'gsheets';
var password = 'eyai4yohF4uX8eeP7phoob';
var url = 'jdbc:mysql://'+server+':'+port+'/'+dbName;
var cc_location = "Summit Up Climbing Centre";


  function setColoursFormat(sheet,cellrange,search, colour) { 
  // Adds a conditional format rule to a sheet that causes all cells in range A1:B3 to turn red
  // if they contain a number between 1 and 10.

  let range = sheet.getRange(cellrange);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(search)
    .setBackground(colour)
   // .setTextStyle(0, 5, bold)
    .setRanges([range])
    .build()
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  }

//setColoursFormatLessThanOrEqualTo(sheet, "O3:O1000",">=","30","#e0ffff")
  function setColoursFormatLessThanOrEqualTo(sheet,cellrange, search, colour) { 
  // Adds a conditional format rule to a sheet that causes all cells in range A1:B3 to turn red
  // if they contain a number between 1 and 10.
search = Number(search);
  let range = sheet.getRange(cellrange);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(search)
    .setBackground(colour)
   // .setTextStyle(0, 5, bold)
    .setRanges([range])
    .build()
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  }


//setNumberFormat(sheet, "O3:O1000", "Rule")
  function setNumberFormat(sheet,cellrange, format) { 

  let range = sheet.getRange(cellrange);
  range.setNumberFormat(format);
  
  }



    function setTextFormat(sheet,cellrange,search, colour) { 
  // Adds a conditional format rule to a sheet that causes all cells in range A1:B3 to turn red
  // if they contain a number between 1 and 10.

  let range = sheet.getRange(cellrange);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(search)
  //  .setBackground(colour)
    .setFontColor(colour)
    .setRanges([range])
    .build()
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  }

    function setWrapped(sheet,cellrange) { 
  var cellrange = sheet.getRange(cellrange);
  cellrange.setWrap(true);
    }
    
function getIP() {
  var url = "http://api.ipify.org";
  var json = UrlFetchApp.fetch(url);
  Logger.log(json);
}

function appendToSheet(sheet, results) {
  let metaData = results.getMetaData();
  let numCols = metaData.getColumnCount();
  const rows = [];

  // First row with column labels
  const colLabels = [];
  for (let col = 0; col < numCols; col++) {
    colLabels.push(metaData.getColumnLabel(col + 1));
  }
  rows.push(colLabels);

  // Remaining rows with results
  while (results.next()) {
    const row = [];
    for (let col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    rows.push(row);
  }
  sheet.clearContents();
  // Find the last row containing a value
  const lastRow = sheet.getDataRange().getLastRow();

  // Set the values of the rows starting from the row below the last row containing a value
  // or the top row if it is empty
  let startRow = lastRow + 1;
  const topRowValues = sheet.getRange(1, 1, 1, numCols).getValues();
  if (topRowValues[0].every(value => value === "")) {
    startRow = 1;
  }

  sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
   // Set the font size of the rows with column labels to 18
  sheet.getRange(1, 1, 1, numCols).setFontSize(14);
  sheet.autoResizeColumns(1, numCols + 1);
}

function appendToSheetWithNoClear(sheet, results) {
  let metaData = results.getMetaData();
  let numCols = metaData.getColumnCount();
  const rows = [];

  // First row with column labels
  const colLabels = [];
  for (let col = 0; col < numCols; col++) {
    colLabels.push(metaData.getColumnLabel(col + 1));
  }
  rows.push(colLabels);

  // Remaining rows with results
  while (results.next()) {
    const row = [];
    for (let col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    rows.push(row);
  }
  //sheet.clearContents();
  // Find the last row containing a value
  const lastRow = sheet.getDataRange().getLastRow();

  // Set the values of the rows starting from the row below the last row containing a value
  // or the top row if it is empty
  let startRow = lastRow + 1;
  const topRowValues = sheet.getRange(1, 1, 1, numCols).getValues();
  if (topRowValues[0].every(value => value === "")) {
    startRow = 1;
  }

  sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
   // Set the font size of the rows with column labels to 18
  sheet.getRange(1, 1, 1, numCols).setFontSize(14);
  sheet.autoResizeColumns(1, numCols + 1);
}

function searchEmailsSheet(ccVolunteerOld) {
  var sheetName = 'Emails';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = dataRange.getNumRows();
  var numCols = dataRange.getNumColumns();
  var returnObj = {};
  for (var i = 1; i < numRows; i++) {
    if (values[i][0] == 'cc_volunteer_old') {
      var foundColumn = -1;
      for (var j = 1; j < numCols; j++) {
        if (values[i][j] == ccVolunteerOld) {
          foundColumn = j;
          break;
        }
      }
      if (foundColumn == -1) {
        throw new Error('Column not found');
      }
      for (var k = 1; k < numRows; k++) {
        returnObj[values[k][0]] = values[k][foundColumn];
      }
      break;
    }
  }
  return returnObj;
}
