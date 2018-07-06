

var fs = require('fs');
var xlsx = require('xlsx');
var moment = require('moment');

let excel = xlsx.readFile("subscribers.xlsx");
let rowNum = Object.keys(excel.Sheets.Sheet1).length/9;
let dateFormat = "YYYY-MM-DD HH:mm:ss";
let sheet = [];
let header = [
    "license_plate_letter",
    "license_plate_number",
    "license_plate_province",
    "title",
    "name",
    "phone_number",
    "start_date",
    "valid_through",
    "start_time",
    "end_time",
    "can_concurrent_parking",
    "id",
    "parking_lot_id",
    "status",
    "created_date",
    "updated_date"]
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
        } else if( column === 7) {
            value = trimWhiteSpace(value);
            value = removeWhiteSpace(value);
        } else if( column === 1) {
            value = trimWhiteSpace(value);
            value = removeWhiteSpace(value);
        }

        if(column !== 6) {
            rowData.push(value);
        }
        alphabet = incrementAlphabet(alphabet);
    }
    rowData.push(moment().startOf("day").format("HH:mm:ss"));
    rowData.push(moment().startOf("day").format("HH:mm:ss"));
    rowData.push("false");
    rowData.push(row);
    rowData.push(1);
    rowData.push("ACTIVE");
    rowData.push(moment().startOf("day").format(dateFormat));
    rowData.push(moment().startOf("day").format(dateFormat));
    sheet.push(rowData);
}

let csvSheet = xlsx.utils.aoa_to_sheet(sheet);
let csvBuild = xlsx.utils.sheet_to_csv(csvSheet);
//console.log(csvSheet);
let workBook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workBook,csvBuild, "subscribers csv")
console.log(workBook);

//xlsx.writeFile(workBook, 'subscribers.csv');
fs.writeFile(__dirname + "/subscribers.csv", workBook.Sheets['subscribers csv'], function (err) {
    console.log(err);
});

function incrementAlphabet(c) {
    return String.fromCharCode(c.charCodeAt(0) + 1);
}

function trimWhiteSpace(str) {
    return str !== undefined && str !== null ? str.trim(): str;
}

function removeWhiteSpace(str) {
    if(str.indexOf(' ') >=0)
        return str.replace(/\s/g, '');
    else
        return str;
}
console.log(Object.keys(excel.Sheets.Sheet1).length);




