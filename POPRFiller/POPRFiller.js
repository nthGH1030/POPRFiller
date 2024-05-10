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
//const XLSX = require("xlsx");
const ExcelJS = require('exceljs');
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
let templatePO = secrets.templatePO;
let templatePR = secrets.templatePR;
let centralSheet = 'POPR summary';

async function readExcelFile(filename, sheetName) {
    try{
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filename);
        const worksheet = workbook.getWorksheet(sheetName);
        //console.log(worksheet)
        return worksheet
    } catch (error) {
        console.log('Error:', error);
    }
}

readExcelFile(filename, centralSheet)
    .then((worksheet) => {
    /* Extract the data */

    let columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
    let indexRow = "7";

    const extractedObj = {};

    for (let i = 0; i < columns.length; i++) {
    const column = columns[i];

    const keyAddress = column.concat(indexRow); 
    const keyValue = worksheet.getCell(keyAddress)?.value || "";

    const cellAddress = column.concat(row);
    const cellValue = worksheet.getCell(cellAddress)?.value || "";

    extractedObj[keyValue] = cellValue;
    }

    //console.log(extractedObj);

    //Call PO or PR
    args[2] === 'po' ? handlePO(templatePO, extractedObj, secrets, worksheet)
    : args[2] === 'pr' ? handlePR(templatePR, extractedObj, secrets)
    : (() => { throw new Error("You can only input pr or po"); })();
    })


.catch((error) => {
    console.log('Error:', error);
});

async function handlePO(templatePO, extractedObj, secrets, worksheet) {
    try {
      let POSheet = 'Purchase Requisition';
  
      /* Open the template */
      const templateWorksheet = await readExcelFile(templatePO, POSheet);
  
      /* Replace the value in the respective field in the template */
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
          let cellAddress = value;
          // Replace the cell value
          templateWorksheet.getCell(cellAddress).value = extractedObj[key];
        }
      }
  
      /* Save as a new file */
      const outputFilename = 'newfile.xlsx';
      await templateWorksheet.workbook.xlsx.writeFile(outputFilename);
      console.log('Workbook saved as a new file:', outputFilename);
    } catch (error) {
      console.log('Error:', error);
    }
  }

function handlePR(templatePR, extractedObj, secrets) {
    /* Open the template */
    let filename = templatePR;
    let workbook = XLSX.readFile(filename);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

       /* replace the value in the respective field in the tempalte */
    let PR = {        
        'Entity': 'C7',
        'PO Number': 'D13',
        'Vendor': 'C16',
        'Capex Nature': 'C36',
        'Purchase description / Payment Certification reason': 'C25',
        'Approved PO amount': 'D39',
        'Delivery date': 'C19',
        'Invoice number': 'D31'
    }

    for (let [key, value] of Object.entries(PR)) {
        if (key in extractedObj) {
            // Get the corresponding cell address
            let cellAddress = value
            // Replace the cell value
            worksheet[cellAddress].v = extractedObj[key];
        }
    }

    /* Save as a new file */
    // Save the new workbook as a new file
    XLSX.writeFile(workbook, secrets.PRpath);

}
