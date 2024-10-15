const scriptProperties = PropertiesService.getScriptProperties();

const server = scriptProperties.getProperty("cred_server");
const port = Number.parseInt(scriptProperties.getProperty("cred_port"), 10);
const dbName = scriptProperties.getProperty("cred_dbName");
const username = scriptProperties.getProperty("cred_username");
const password = scriptProperties.getProperty("cred_password");
const url = `jdbc:mysql://${server}:${port}/${dbName}`;
const apidomain = scriptProperties.getProperty("cred_apidomain");
const apiusername = scriptProperties.getProperty("cred_apiusername");
const apipassword = scriptProperties.getProperty("cred_apipassword");

const cc_location = "Summit Up Climbing Centre";
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
	brightYellow: "#FFFF00",
};

function setupSheet(name) {
	const spreadsheet = SpreadsheetApp.getActive();
	const sheet = spreadsheet.getSheetByName(name);
	sheet.clearContents();
	sheet.clearFormats();
	return sheet;
}

function setupCell(name, range) {
	const spreadsheet = SpreadsheetApp.getActive();
	const sheet = spreadsheet.getSheetByName(name);
	const cellValue = sheet.getRange(range).getValue();

	return cellValue;
}

function appendToSheet(sheet, results) {
	const metaData = results.getMetaData();
	const numCols = metaData.getColumnCount();
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
	const lastRow = sheet.getLastRow();
	const numRows = Math.max(5, lastRow - 1);
	return sheet.getRange(2, columnIndex, numRows, 1);
}

function setColoursFormat(sheet, columnHeader, search, colour) {
	const range = getColumnRange(sheet, columnHeader);
	const rule = SpreadsheetApp.newConditionalFormatRule()
		.whenTextContains(search)
		.setBackground(colour)
		.setRanges([range])
		.build();
	const rules = sheet.getConditionalFormatRules();
	rules.push(rule);
	sheet.setConditionalFormatRules(rules);
}

function setTextFormat(sheet, columnHeader, search, colour) {
	const range = getColumnRange(sheet, columnHeader);
	const rule = SpreadsheetApp.newConditionalFormatRule()
		.whenTextContains(search)
		.setFontColor(colour)
		.setRanges([range])
		.build();
	const rules = sheet.getConditionalFormatRules();
	rules.push(rule);
	sheet.setConditionalFormatRules(rules);
}

function setWrapped(sheet, columnHeader) {
	const range = getColumnRange(sheet, columnHeader);
	range.setWrap(true);
	sheet.setColumnWidth(range.getColumn(), 300); // Set column width to 300 pixels
}

function setColoursFormatLessThanOrEqualTo(
	sheet,
	columnHeader,
	search,
	colour,
) {
	const numberSearch = Number(search);
	const range = getColumnRange(sheet, columnHeader);
	const rule = SpreadsheetApp.newConditionalFormatRule()
		.whenNumberLessThanOrEqualTo(numberSearch)
		.setBackground(colour)
		.setRanges([range])
		.build();
	const rules = sheet.getConditionalFormatRules();
	rules.push(rule);
	sheet.setConditionalFormatRules(rules);
}

function setNumberFormat(sheet, columnHeader, format) {
	const range = getColumnRange(sheet, columnHeader);
	range.setNumberFormat(format);
}

function makeReport(stmt, reportConfig) {
	const sheet = setupSheet(reportConfig.sheetName);

	let query = reportConfig.query;

	// Only replace ${cell} if it exists in the query
	if (query.includes("${cell}")) {
		if (Number.isNaN(Number(cell)) || cell === "") {
			throw new Error("Invalid event selected - please try again");
		}
		query = query.replace(/\${cell}/g, cell);
	}

	const results = stmt.executeQuery(query);

	appendToSheet(sheet, results);

	if (reportConfig.formatting) {
		for (const format of reportConfig.formatting) {
			if (format.type === "color") {
				setColoursFormat(sheet, format.column, format.search, format.color);
			} else if (format.type === "text") {
				setTextFormat(sheet, format.column, format.search, format.color);
			} else if (format.type === "wrap") {
				setWrapped(sheet, format.column);
			} else if (format.type === "numberFormat") {
				setNumberFormat(sheet, format.column, format.format);
			} else if (format.type === "colorLessThanOrEqual") {
				setColoursFormatLessThanOrEqualTo(
					sheet,
					format.column,
					format.value,
					format.color,
				);
			} else if (format.type === "columnWidth") {
				setColumnWidth(sheet, format.column, format.width);
			}
		}
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
	const metaData = results.getMetaData();
	const numCols = metaData.getColumnCount();
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
	if (topRowValues[0].every((value) => value === "")) {
		startRow = 1;
	}

	sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
	// Set the font size of the rows with column labels to 14
	sheet.getRange(1, 1, 1, numCols).setFontSize(14);
	sheet.autoResizeColumns(1, numCols + 1);
}

function searchEmailsSheet(ccVolunteerOld) {
	const sheetName = "Emails";
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
	const dataRange = sheet.getDataRange();
	const values = dataRange.getValues();
	const numRows = dataRange.getNumRows();
	const numCols = dataRange.getNumColumns();
	const returnObj = {};
	for (let i = 1; i < numRows; i++) {
		if (values[i][0] === "cc_volunteer_old") {
			let foundColumn = -1;
			for (let j = 1; j < numCols; j++) {
				if (values[i][j] === ccVolunteerOld) {
					foundColumn = j;
					break;
				}
			}
			if (foundColumn === -1) {
				throw new Error("Column not found");
			}
			for (let k = 1; k < numRows; k++) {
				returnObj[values[k][0]] = values[k][foundColumn];
			}
			break;
		}
	}
	return returnObj;
}
