

var fs = require('fs');
var xlsx = require('xlsx');
var moment = require('moment');

let excel = xlsx.readFile("subscribers.xlsx");
let rowNum = Object.keys(excel.Sheets.Sheet1).length/9;
let dateFormat = "YYYY-MM-DD HH:mm:ss";
let sheet = [];
let header = [
    "licensePlateLetter",
    "licensePlateNumber",
    "licensePlateProvince",
    "title",
    "name",
    "phoneNumber",
    "start",
    "validThrough"]
sheet.push(header);
for(row = 1; row <= rowNum; row++) {
    let alphabet = "A";
    let rowData = [];
    for(column = 1; column <= 9; column++) {
        let preValue = excel.Sheets.Sheet1[alphabet + row];
        let value = preValue !== undefined ? preValue["w"] : undefined;
        value = trimWhiteSpace(value);

        if((column === 8 && value !== undefined) || column === 9) {
            value = moment(value,"YYYY-MM-DD").startOf("day").format(dateFormat);   
        } else if (column === 8 && value === undefined) {
            value = moment().startOf("day").format(dateFormat);
        } else if(column === 5) {
            value = value + " " + excel.Sheets.Sheet1[incrementAlphabet(alphabet) + row]["w"];
        } 

        if(column !== 6) {
            rowData.push(value);
        }
        alphabet = incrementAlphabet(alphabet);
    }
    sheet.push(rowData);
}

let csvSheet = xlsx.utils.aoa_to_sheet(sheet);
let csvBuild = xlsx.utils.sheet_to_csv(csvSheet);
//console.log(csvSheet);
let workBook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workBook,csvBuild, "subscribers csv")
console.log(workBook);

//xlsx.writeFile(workBook, 'subscribers.csv');
fs.writeFile(__dirname + "/test.csv", workBook.Sheets['subscribers csv'], function (err) {

});

function incrementAlphabet(c) {
    return String.fromCharCode(c.charCodeAt(0) + 1);
}

function trimWhiteSpace(str) {
    return str !== undefined && str !== null ? str.trim(): str;
}
console.log(Object.keys(excel.Sheets.Sheet1).length);




