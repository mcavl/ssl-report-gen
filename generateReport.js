const util = require('util');
const exec = util.promisify(require('child_process').exec);
const fs = require('fs');
const readline = require('readline');
const PromisePool = require('es6-promise-pool');
const moment = require('moment');
const limit = moment().add(3, 'month');
const xl = require('excel4node');


// Excel workbook
const serverWorkBook = new xl.Workbook();
// Excel sheet
const serversWorkSheet = serverWorkBook.addWorksheet('Servers List');

// save servername to be processed
const lines = [];

// today
const now = moment(Date.now());

// styles to be applied to the cells
const validCertStyle = serverWorkBook.createStyle({
	font: {
		size: 12
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		bgColor: '#009414',
		fgColor: '#009414',
	}	
});
const invalidCert = serverWorkBook.createStyle({
	font: {
		size: 12,
		color: "white"
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		bgColor: '#a10800',
		fgColor: '#a10800',
	}	
});

const alertCert = serverWorkBook.createStyle({
	font: {
		size: 12,
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		bgColor: '#dbd800',
		fgColor: '#dbd800',
	}	
});

const strikethroughStyle = serverWorkBook.createStyle({
	font: {
		size: 12,
		effects: "striketrhough",
		color: "red"
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		bgColor: '#ababa7',
		fgColor: '#ababa7',
	}	
});

// write excel file
let writeExcel = function(server, date, cellNum) {
	let dateMessage = "Server could not be verified.";
	let style = alertCert;
	
	if (date !== null) {
		let validity = moment(new Date(date));
		dateMessage = validity.format('LLL');
		if (validity.isBefore(now)) {
			style = invalidCert;
		} else if (validity.isAfter(limit)) {
			style = validCertStyle;
		}
	} else {
		style = strikethroughStyle;
	}
	
	serversWorkSheet
		.cell((cellNum + 1), 1)
		.string(server)
		.style(style);
	
	serversWorkSheet
		.cell((cellNum + 1), 2)
		.string(dateMessage)
		.style(style);
}

async function getExpiringDate(index) {
	const servername = lines[index];
	const { stdout, stderr } = await exec(`./checkSSL.sh ${servername}`);
	const certExpiryDate = stderr ? null : stdout;
	writeExcel(servername, certExpiryDate, index);
}

let count = 0;

let promiseProducer = function () {
	if (count < lines.length) {
		return getExpiringDate(count++);
	} else {
		return null;
	}
}

// The number of promises to process simultaneously. 
let concurrency = 10;

// Create a pool. 
let pool = new PromisePool(promiseProducer, concurrency);

const lineReader = readline.createInterface({
	input: fs.createReadStream('DATA'),
	crlfDelay: Infinity
});

lineReader
	.on('line', async (line) => {
		lines.push(line);
	})
	.on('close', function(){
		// Start the pool. 
		var poolPromise = pool.start()

		// Wait for the pool to settle. 
		poolPromise.then(function () {
			console.log('Check SSL done');
			serverWorkBook.write('servers.xlsx'); // Writes the file ExcelFile.xlsx to the process.cwd();
			
		}, function (error) {
			console.log('Some promise rejected: ' + error.message)
		});
	});