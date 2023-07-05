const XLSX = require("xlsx");
const fs = require("fs");

const filePath = "./Products - Copy.xlsx";

// Column (A = 0, B = 1, C = 2, ...)
const columnIndex = 1; // column B

try {
  // Check if the file exists
  if (!fs.existsSync(filePath)) {
    throw new Error("File not found");
  }
  const workbook = XLSX.readFile(filePath);

  // Using first sheet
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Get all cell addresses the column
  const columnAddresses = Object.keys(worksheet).filter((address) =>
    address.startsWith(`${XLSX.utils.encode_col(columnIndex)}`)
  );

  let duplicateValues = [];
  let uniqueValues = new Set();

  //function to check for dupliactes
  const checkDuplicates = () => {
    uniqueValues = new Set();
    duplicateValues = [];

    columnAddresses.forEach((address) => {
      const cell = worksheet[address];
      const value = cell ? cell.v : undefined;

      if (uniqueValues.has(value)) {
        // Duplicate value found
        duplicateValues.push(value);
      } else {
        // unique value
        uniqueValues.add(value);
      }
    });
  };

  console.log("start algo");
  // repeat algo until therese is only one duplicated value
  checkDuplicates();
  while (duplicateValues.length > 1) {
    var firstVal = duplicateValues[1];
    var counter = 1;

    columnAddresses.forEach((address) => {
      const cell = worksheet[address];
      const value = cell ? cell.v : undefined;
      if (firstVal === value) {
        const formattedCounter = counter.toString().padStart(3, "0");
        worksheet[address].v = worksheet[address].v + "-" + formattedCounter;
        counter++;
      }
    });

    XLSX.writeFile(workbook, filePath);
    console.log(firstVal + " done");
    checkDuplicates();
  }

  console.log("end algo");
  // printing
  if (duplicateValues.length > 0) {
    console.log("Duplicate values found:");
    console.log(duplicateValues);
  } else {
    console.log("No duplicate values found.");
  }
} catch (error) {
  console.error(`Error: ${error.message}`);
}
