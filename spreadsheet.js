const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');

const creds = require('./client_secret.json');

async function accessSpreadsheet() {

	// // actual sign up sheet
	// const doc = new GoogleSpreadsheet('1XLcU1dGu9d34gDDG9wXJ7EcskJfdhyioqzoJCSvWY2g');

	// practice sign up sheet
	const doc = new GoogleSpreadsheet('1K-KIbHwlHGcnfgqZvyX7FZdNM85cVK5AATm7aPG9Bi4');

	await promisify(doc.useServiceAccountAuth)(creds);
	const info = await promisify(doc.getInfo)();
	const sheet = info.worksheets[0];
	console.log("Title: " + sheet.title +", Rows: " + sheet.rowCount);

	const anjanasheet = info.worksheets[2];
	console.log("Title: " + anjanasheet.title +", Rows: " + anjanasheet.rowCount);

	for(var i = 0; i < 9; i++)
	{
		const cell_name = await promisify(anjanasheet.getCells)({
			'min-row' : 11 + i,
			'max-row' : 11 + i,
			'min-col' : 1,
			'max-col' : 3,
			'return-empty': true
			});
		if(cell_name[1].value == "Denys Dziubii")
			break;

		if(cell_name[1].value == false)
		{
			cell_name[1].value = "Denys Dziubii";
			cell_name[1].save();
			cell_name[2].value = "ddziubii@sfu.ca";
			cell_name[2].save();
			break;
		}
	}
}
// add functionality to find the longest available time after your slot in case your meeting runs long
// add html and CSS to create an interactive software
// send discord notifications or telegram messages what time you were booked for, messenger
accessSpreadsheet();