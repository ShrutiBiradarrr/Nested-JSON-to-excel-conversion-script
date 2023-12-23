const fs = require('fs');
const XLSX = require('xlsx');

const json_data = {
 
   "name": "The Reading Nook",
  "location": "123 Book St, Bibliopolis",
  "isOpen": true,
  "numberOfSections": 2,
  "contact": null,
  "popularGenres": ["Fiction", "Mystery", "Sci-Fi", "Non-Fiction"],
  "test": {
    "test1": "Test 1",
    "test2": {
      "test3": "Test 3"
    }
  },
  "sections": [
    {
      "sectionName": "Section 1",
      "books": [
        {
          "title": "Journey to the Unknown",

          "author": "Alice Wonder",
          "price": 12.99,
          "isAvailable": true
        },
        {
          "title": "Mystery of the Ancient Map",
          "author": "Clive Cussler",
          "price": 15.5,
          "isAvailable": false
        }
      ]
    },
    {
      "sectionName": "Section 2",
      "books": [
        {
          "title": "The Reality of Myths",
          "author": "Helen Troy",
          "price": 18.25,
          "isAvailable": true
        }
      ]
    }
  ]
};

// Function to flat the structure of nested objects
const flat_structure = (obj, parentKey = '', defaultValue = '') => {
    let result = {};
    for (let key in obj) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        const nestedSheetName = `${parentKey}${key}`;
        result[parentKey + key] = nestedSheetName;
        const nested_data = flat_structure(obj[key], '', defaultValue);
        insert_nested_sheet(nested_data, nestedSheetName);
      } else {
        result[`${parentKey}${key}`] = obj[key] !== null ? obj[key] : defaultValue;
      }
    }
    return result;
  };

// Function to insert data into nested sheet
const insert_nested_sheet = (nested_data, SheetName) => {
    SheetName = String(SheetName); // Ensure SheetName is treated as a string
    let index = 1;
    let modifiedSheetName = SheetName;
    
    while (nestedWorkbook.SheetNames.indexOf(modifiedSheetName) >= 0) {
      modifiedSheetName = `${SheetName}_${index++}`;
    }
  
    const ws = XLSX.utils.json_to_sheet([nested_data]);
    const nestedSheet = XLSX.utils.aoa_to_sheet(XLSX.utils.sheet_to_json(ws, { header: 1 }));
    XLSX.utils.book_append_sheet(nestedWorkbook, nestedSheet, modifiedSheetName);
  };
  

// new workbook
const mainWorkbook = XLSX.utils.book_new();
const nestedWorkbook = XLSX.utils.book_new();

// Flatten the JSON data with a default value 
// to handle null values
const flattenedData = flat_structure(json_data, '', 'N/A');

// Create a new worksheet for the main sheet
const mainSheet = XLSX.utils.json_to_sheet([flattenedData]);

// Append the main sheet to the main workbook
XLSX.utils.book_append_sheet(mainWorkbook, mainSheet, 'MainSheet');

//nested sheet
const nested_output_path = 'nested_output.xlsx';
XLSX.writeFile(nestedWorkbook, nested_output_path, { bookType: 'xlsx', bookSST: false, type: 'file' });

//main sheet
const main_output_path = 'main_output.xlsx';
XLSX.writeFile(mainWorkbook, main_output_path, { bookType: 'xlsx', bookSST: false, type: 'file' });

console.log(`Main data : ${main_output_path}`);
console.log(`Nested data : ${nested_output_path}`);
