var XLSX = require('xlsx'); // npm install xlsx
var fs = require('fs'); // npm install fs

var JSON_FILE_PATH = "jsonFile/excelOutput.json";
var EXCEL_FILE_PATH = "excelFile/multpileSheet.xlsx";
var CONFIG_FILE = "sheetConfig.json";
var workbook = XLSX.readFile(EXCEL_FILE_PATH);
var sheetNames = workbook.SheetNames;

// sheetConfig file writing
var config = '[{"sheetCount":'+sheetNames.length+'}]';
var configFile = fs.writeFile(CONFIG_FILE, config);

var file = fs.statSync(JSON_FILE_PATH);
if(file.size > 0) {
	fs.writeFile(JSON_FILE_PATH,''); // to delete the old data in json file
}
fs.appendFileSync(JSON_FILE_PATH, "{", 'utf8');

for(var i = 0; i<sheetNames.length; i++) {
		
	name = sheetNames[i]; // to get sheet name
	var sheet = workbook.Sheets[name]; // select the sheet to read
	
	fs.appendFileSync(JSON_FILE_PATH, '"'+name+'":' ,'utf8'); // to write sheet Name in file
	fs.appendFileSync(JSON_FILE_PATH,"[", 'utf8');
	
	sheet = XLSX.utils.sheet_to_json(sheet);
	
	var sum = 0;
	var sep = "";
	
	for (var cell in sheet) {
		
		data = sheet[cell];
	
		const content = JSON.stringify(data);
		
		fs.appendFileSync(JSON_FILE_PATH, sep+JSON.stringify(data), "utf8");
		
		if(!sep)
			sep = ",\n";
		
		sum += 1;
	
	}
	if(i == sheetNames.length-1)
		fs.appendFileSync(JSON_FILE_PATH, "]", 'utf8');
	else
		fs.appendFileSync(JSON_FILE_PATH, "],", 'utf8');
	console.log(sum+" Rows readed from sheet "+name);
	
}
fs.appendFileSync(JSON_FILE_PATH, "}", 'utf8');
	
//console.log("File Saved!!")	
