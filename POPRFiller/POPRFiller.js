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
if (isNaN(args[1])) {
    throw new TypeError("The row you typed is not a number");
  }


/* Read the excel file */
let filename = args[0];
let template = args[2];
let workbook = XLSX.readFile(filename);

//This is the first worksheet of the file
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

/* Extract the data */
let row = args[1]
let columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
let indexRow = "7";

const extractedObj = {};

const secretKeys = Object.keys(secrets);

/*
Iterate both secretKeys and columns in a single loop
Update the addressValue with column per iteration,
Get the value using worksheet[CellAddress]
assign the new value to the extractedObj[secretKey] 
*/

for (let i = 0; i < columns.length; i++) {
  //const secretKey = secretKeys[i];
  const column = columns[i];

  const keyAddress = column.concat(indexRow); 
  const keyValue = worksheet[keyAddress]?.v || "";

  const cellAddress = column.concat(row);
  const cellValue = worksheet[cellAddress]?.v || "";

  extractedObj[keyValue] = cellValue;
}

console.log(extractedObj);

//pass in the tempalte to PO / PR function
args[3] == "po" ? handlePO(template) : handlePR(template);

function handlePO(template) {
    /* Open the template */
    let filename = template;
    let workbook = XLSX.readFile(filename);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    /* replace the value in the respective field in the tempalte */
    let PO = {
        
    }

    /* Save as a new file */

}

function handlePR(template) {


}
