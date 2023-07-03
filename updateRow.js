const XLSX = require("xlsx");

// Load the workbook
const workbook = XLSX.readFile("data.xlsx");

// Select the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Find the last row
const lastRow = XLSX.utils.decode_range(worksheet["!ref"]).e.r;
console.log("last row=>", lastRow);

// Define the data for the new row
const newRowIndex = lastRow + 1;
const newPhoneNumber = "1234567890";
const newFullName = "John Doe";
const newAum = 1000000;

// Set the values in the new row for phoneNumber, fullName, and aum
worksheet[`A${newRowIndex}`] = { t: "s", v: newPhoneNumber };
worksheet[`B${newRowIndex}`] = { t: "s", v: newFullName };
worksheet[`C${newRowIndex}`] = { t: "n", v: newAum };
console.log(worksheet[`D${lastRow}`]?.v);
let value = worksheet[`D${lastRow}`]?.v;
console.log("value",value)
const newValue = getNewAgentCode(value)
// worksheet["D29"]?.v
console.log("newValue",newValue)
worksheet[`D${newRowIndex}`] = { t: "s", v: newValue };
//  worksheet[cellCoordinates]?.v

// Copy the values from the last row for the remaining columns
for (let i = 4; i <= 27; i++) {
  const column = XLSX.utils.encode_col(i);
  const lastCellAddress = `${column}${lastRow}`;
  // console.log("lastCellAddress", lastCellAddress);
  const newCellAddress = `${column}${newRowIndex}`;
  // console.log("newCellAddress", newCellAddress);
  worksheet[newCellAddress] = { ...worksheet[lastCellAddress] };
}

// Update the range
worksheet["!ref"] = XLSX.utils.encode_range({
  s: { c: 0, r: 0 },
  e: { c: 27, r: newRowIndex },
});

// Save the workbook
XLSX.writeFile(workbook, "data.xlsx");

console.log("New row added successfully.");

function getNewAgentCode(value) {
  const teamMembersLength = 4;
  // const str = "AB-1";
  var numericValue = parseInt(value.match(/\d+/)[0]);
  console.log(numericValue);

  if (numericValue == teamMembersLength) {
    numericValue = 0;
  }

  const incrementedValue = numericValue + 1;
  console.log(incrementedValue);
  const result = value.replace(/\d+/, incrementedValue);
  console.log(result);
  return result;
  // console.log(numericValue); // Output: 1
}
