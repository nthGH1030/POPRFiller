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
//const secretValue = Object.values(secrets);

/* Command line arguement */
const args = process.argv.slice(2);



/* Handle exceptions at Command line arguement */
if (isNaN(args[1])) {
    throw new TypeError("The row you typed is not a number");
  }


/* Read the excel file */
let filename = args[0];
let row = args[1];
let template = secrets.templatePO;
let workbook = XLSX.readFile(filename);

//This is the first worksheet of the file
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

/* Extract the data */

let columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
let indexRow = "7";

const extractedObj = {};


for (let i = 0; i < columns.length; i++) {
  const column = columns[i];

  const keyAddress = column.concat(indexRow); 
  const keyValue = worksheet[keyAddress]?.v || "";

  const cellAddress = column.concat(row);
  const cellValue = worksheet[cellAddress]?.v || "";

  extractedObj[keyValue] = cellValue;
}

console.log(extractedObj);

//Call PO or PR 
args[2] === 'po' ? handlePO(template, extractedObj, secrets)
  : args[2] === 'pr' ? handlePR(template)
  : (() => { throw new Error("You can only input pr or po"); })();

function handlePO(template, extractedObj, secrets) {
    /* Open the template */
    let filename = template;
    let workbook = XLSX.readFile(filename);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    /* replace the value in the respective field in the tempalte */
    let PO = {
        "PO Number": "F3",
        "Entity": "C9",
        "Description": "C14",
        "Type of expense": "C17",
        "Approved PO amount": "C19",
        "Vendor": "C37",
        "staff": "C44"
    }

    for (let [key, value] of Object.entries(PO)) {
        if (key in extractedObj) {
            // Get the corresponding cell address
            let cellAddress = value
            // Replace the cell value
            worksheet[cellAddress].v = extractedObj[key];
        }
    }

    /* Save as a new file */
    // Save the new workbook as a new file
    XLSX.writeFile(workbook, secrets.path);
}

function handlePR(template) {


}
