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

console.log(args);

/* Handle exceptions at Command line arguement */



/* Read the excel file */
let filename = args[0];

let workbook = XLSX.readFile(filename);
//This is the first worksheet of the file
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

/* Extract the data */

const cellAddresses = ["A11", "A12", "A13"];
const cellObjects = cellAddresses.map(cellAddress => {
    cellValue = worksheet[cellAddress]?.v;
    return {cellAddress : cellValue};
});

console.log(cellObjects);
