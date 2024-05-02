/*
1. Takes in command line arguement of the of the following:
- File path of the central table
- File path of the template file
- Po / PR
- The excel Row to copy data from

2. the logic 
- open the central table excel file
- open the template file depending on user input (PO / PR)
- read the a certain row of data in the central excel file according to user input
- Put that row of data in a object
    - Venue: HongKong ; Project: Beta; Cost: $100,000;
- fill the template PO/PR wiht the data
- Save as a new file
- close both files
*/

const secrets = require('./secrets.json');
const XLSX = require("xlsx");


/* Command line arguement */
const args = process.argv.slice(2);

/* Handle exceptions at Command line arguement */



/* Read the excel file */
let filename = args[0];
let workbook = XLSX.readFile(filename);

//This is the first worksheet of the file
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

/* Extract the data */
let row = args[1]
let columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];


const cellAddresses = columns.reduce((result,column) => {
    let cellAddress = column.concat(row);
    result[cellAddress] = "";
    return result;
},{});

for (let key in cellAddresses) {

    cellValue = worksheet[key]?.v;
    cellAddresses[key] = cellValue;

}

/*Update the key with the keys in secet*/
const combinedObj = {};

const obj1Keys = Object.keys(cellAddresses);
const obj2Keys = Object.keys(secrets);

for (let i = 0; i < obj1Keys.length; i++) {
  const key1 = obj1Keys[i];
  const key2 = obj2Keys[i];
  combinedObj[key2] = cellAddresses[key1];
}

console.log(combinedObj)

