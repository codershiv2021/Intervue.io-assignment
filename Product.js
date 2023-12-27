//Make sure to install xlsx library

//import the required Node.js File System (fs) and xlsx libraries.

const fs = require('fs');
const XLSX = require('xlsx');


// Function we'll use to convert nested JSON to Excel Workbook
function convertJsonToExcel(jsonData) {

//workbook is created using XLSX.utils.book_new() to create a new Excel workbook.
  const workbook = XLSX.utils.book_new();


  // Iterate over each key in the JSON
  Object.keys(jsonData).forEach(sheetName => {
    const sheetData = jsonData[sheetName];

    // Create a worksheet for each key
    const worksheet = XLSX.utils.json_to_sheet(sheetData);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  // Save the workbook to a file
  XLSX.writeFile(workbook, 'output.xlsx');
  console.log('Excel file generated successfully.');
}

// Read the JSON file
const jsonFilePath = 'path/to/my/json-file.json'; // path to the JSON file
fs.readFile(jsonFilePath, 'utf8', (err, data) => {

//checks for errors and, if there are none, parses the JSON data and calls the convertJsonToExcel function

  if (err) {
    console.error('Error reading JSON file:', err);
    return;
  }

  try {
    const jsonData = JSON.parse(data);
    convertJsonToExcel(jsonData);
  } catch (parseError) {
    console.error('Error parsing JSON:', parseError);
  }
});


