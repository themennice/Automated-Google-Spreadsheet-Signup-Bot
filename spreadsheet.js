const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');

const creds = require('./client_secret.json');

async function accessSpreadsheet() {
	const doc = new GoogleSpreadsheet('1K-KIbHwlHGcnfgqZvyX7FZdNM85cVK5AATm7aPG9Bi4');
	await promisify(doc.useServiceAccountAuth)(creds);
	const info = await promisify(doc.getInfo)();
	const sheet = info.worksheets[0];
	console.log("Title: " + sheet.title +", Rows: " + sheet.rowCount);

	const anjanasheet = info.worksheets[2];
	console.log("Title: " + anjanasheet.title +", Rows: " + anjanasheet.rowCount);

	const cell_name = await promisify(anjanasheet.getCells)({
		'min-row' : 11,
		'max-row' : 11,
		'min-col' : 1,
		'max-col' : 3,
		'return-empty': true
		});

	cell_name[1].value = "Denys Dziubii";
	cell_name[1].save();
	cell_name[2].value = "ddziubii@sfu.ca";
	cell_name[2].save();
}

accessSpreadsheet();