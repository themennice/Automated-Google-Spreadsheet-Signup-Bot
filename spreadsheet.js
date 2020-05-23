const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const nodemailer = require('nodemailer');

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

	const date = await promisify(anjanasheet.getCells)({
			'min-row' : 6,
			'max-row' : 6,
			'min-col' : 2,
			'max-col' : 2,
			'return-empty': true
			});
	console.log(date[0].value);

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
			email(cell_name[0], date[0]);
			break;
		}
	}
}

function email(time, date){
    var emailAddr = 'dziubiiden@gmail.com';

    var transporter = nodemailer.createTransport(
        {
            service: 'gmail',
            auth: { user: 'noreply.skypal@gmail.com', pass: 'SkyPal*0*' }
        });

    var subjects = 'BUS 360W session booked for ' + time.value;
    console.log(subjects);

    var mailOptions = {
        from: 'noreply.skypal@gmail.com',
        to: emailAddr,
        subject: subjects.toString(),
        text: 'Hello,\n\nYou have signed up for office hours for BUS 360W at ' + time.value + ' on ' + date.value + '.\n\nRegards, \n\nGoogle Spreadsheet Signup Bot',
        //html: 'Hello, <br><br> You have signed up for office hours for BUS 360W. <br><br>Regards, <br><br> Google Spreadsheet Signup Bot'
    };

    transporter.sendMail(mailOptions, function (error, info) {
        if (error)
            console.log(error);
        else
            console.log('Email sent: ' + info.response);
    });
}

// add functionality to find the longest available time after your slot in case your meeting runs long
// add html and CSS to create an interactive software
// send discord notifications or telegram messages what time you were booked for, messenger
accessSpreadsheet();