const server = '18.168.242.164';
const port = 3306;
const dbName = 'bitnami_wordpress';
const username = 'gsheets';
const password = 'eyai4yohF4uX8eeP7phoob';
const url = 'jdbc:mysql://'+server+':'+port+'/'+dbName;
const cc_location = "Parthian Climbing Manchester";
const apidomain="climbingclan.com";
const apiusername="ck_3f8cd172e7aed36533d434e04e8c0b2affe19075";
const apipassword="cs_817f3cd22ae28bc33fa716a6fdfd707188c0409b";
const cell = setupCell("Dashboard", "B5");

const colors = {
  lightRed: "#ffcccb",
  lightGreen: "#90ee90",
  lightYellow: "#ffd898",
  lightBlue: "#ADD8E6",
  grey: "#a9a9a9",
  pink: "#ff75d8",
  yellow: "#fad02c",
  white: "#FFFFFF",
  green: "#5CFF5C",
  brightYellow: "#FFFF00"
};

function setupSheet(name) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(name);
  sheet.clearContents();
  sheet.clearFormats();
  return sheet;
}

function setupCell(name, range) {
  var spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(name);
  let cellValue = sheet.getRange(range).getValue();

  return cellValue;
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

  sheet.getRange(1, 1, rows.length, numCols).setValues(rows);

  // Set the font size of the rows with column labels to 14
  sheet.getRange(1, 1, 1, numCols).setFontSize(14);
  sheet.autoResizeColumns(1, numCols);
}

function getColumnRange(sheet, columnHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndex = headers.indexOf(columnHeader) + 1;
  if (columnIndex === 0) {
    throw new Error(`Column '${columnHeader}' not found in the sheet.`);
  }
  return sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1);
}

function setColoursFormat(sheet, columnHeader, search, colour) {
  let range = getColumnRange(sheet, columnHeader);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(search)
    .setBackground(colour)
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function setTextFormat(sheet, columnHeader, search, colour) {
  let range = getColumnRange(sheet, columnHeader);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(search)
    .setFontColor(colour)
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function setWrapped(sheet, columnHeader) {
  var range = getColumnRange(sheet, columnHeader);
  range.setWrap(true);
  sheet.setColumnWidth(range.getColumn(), 300); // Set column width to 300 pixels
}

function setColoursFormatLessThanOrEqualTo(sheet, columnHeader, search, colour) {
  search = Number(search);
  let range = getColumnRange(sheet, columnHeader);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(search)
    .setBackground(colour)
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function setNumberFormat(sheet, columnHeader, format) {
  let range = getColumnRange(sheet, columnHeader);
  range.setNumberFormat(format);
}

function makeReport(stmt, reportConfig) {
  let sheet = setupSheet(reportConfig.sheetName);
    if (isNaN(cell) || cell === "") {
      throw new Error("Invalid event selected - please try again");
    }
  var results = stmt.executeQuery(reportConfig.query.replace(/\${cell}/g, cell));

  appendToSheet(sheet, results);

  if (reportConfig.formatting) {
    reportConfig.formatting.forEach(format => {
      if (format.type === 'color') {
        setColoursFormat(sheet, format.column, format.search, format.color);
      } else if (format.type === 'text') {
        setTextFormat(sheet, format.column, format.search, format.color);
      } else if (format.type === 'wrap') {
        setWrapped(sheet, format.column);
      } else if (format.type === 'numberFormat') {
        setNumberFormat(sheet, format.column, format.format);
      } else if (format.type === 'colorLessThanOrEqual') {
        setColoursFormatLessThanOrEqualTo(sheet, format.column, format.value, format.color);
      } else if (format.type === 'columnWidth') {
        setColumnWidth(sheet, format.column, format.width);
      }
    });
  }

  results.close();
}

function setColumnWidth(sheet, columnHeader, width) {
  var range = getColumnRange(sheet, columnHeader);
  sheet.setColumnWidth(range.getColumn(), width);
}

function refreshOutput() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  readOutput(stmt);
  stmt.close();
  conn.close();
}

function getIP() {
  var url = "http://api.ipify.org";
  var json = UrlFetchApp.fetch(url);
  Logger.log(json);
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
  // Set the font size of the rows with column labels to 14
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
