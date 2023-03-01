const xlsx = require('xlsx');

const emails = [];
var workbook = xlsx.readFile(``);
var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'Emails';
var worksheet = workbook.Sheets[first_sheet_name];
const columnName = Object.keys(worksheet).find(
  (key) => worksheet[key].v === address_of_cell
);

for (let key in worksheet) {
  if (key.toString()[0] === columnName[0]) {
    emails.push(worksheet[key].v);
  }
}
console.log('Result list', emails);
