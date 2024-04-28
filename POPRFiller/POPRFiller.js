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

const secrets = requires('./secrets.json');
var XLSX = require("xlsx");