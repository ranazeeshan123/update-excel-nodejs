const XLSX = require("xlsx");

// Load the Excel file
const workbook = XLSX.readFile("data.xlsx");

// Select the worksheet
const worksheet = workbook.Sheets["Sheet1"]; // Replace 'Sheet1' with the actual sheet name

// Specify the CNIC and date to search for
const cnicToSearch = "06467648154"; // Replace with the desired CNIC value
var dateToSearch = "08 Jun 2023"; // Replace with the desired date value
dateToSearch = dateToSearch.toUpperCase();

let targetCellAddress = null;
let dateColumn = null;
var cellDateOld = "";
// Find the CNIC in the first column (column 'A')
for (const cellAddress in worksheet) {
  if (cellAddress[0] === "!") continue; // Skip non-cell addresses

  const cell = worksheet[cellAddress];
  const cellValue = cell.v;
  // Check if the cell value matches the desired CNIC
  if (cellValue === cnicToSearch && cellAddress.startsWith("A")) {
    console.log("cellValue in condition => ", cellValue);
    const rowNumber = parseInt(cellAddress.substring(1));
    console.log("rowNumber => ", rowNumber);

    // Search for the date in the first row
    for (const column in worksheet) {
      if (column[0] === "!") continue; // Skip non-cell addresses

      // Check if the cell address is in the first row (row 1)
      if (column.substring(1) === "1") {
        const cell = worksheet[column];
        console.log("=============");
        console.log("cell", cell);
        const cellValue = cell.v;
        // const cellDate = XLSX.SSF.format("yyyy-mm-dd", cellValue);
        // cellDateOld = cell.v;
        console.log("cellValue before ", cellValue);
        console.log("dateToSearch before ", dateToSearch);

        // Check if the cell value matches the desired date
        if (cellValue.toUpperCase() == dateToSearch) {
          const columnName = column.charAt(0);
          console.log("columnName", columnName);
          dateColumn = columnName;
          targetCellAddress = columnName + rowNumber;
          break;
        }
      }
    }

    break;
  }
}

if (targetCellAddress) {
  console.log("Matching cell address:", targetCellAddress);
  // Update the value of the matching cell
  //   const targetCellAddress = "B2";
  const number = targetCellAddress.match(/\d+/)[0];
  console.log(number);
  worksheet[targetCellAddress].v = 52000;
  //   const formattedDate = dateToSearch
  //     .toLocaleDateString("en-US", {
  //       year: "numeric",
  //       month: "2-digit",
  //       day: "2-digit",
  //     })
  //     .replace(/\//g, "-");
  //   worksheet[dateColumn + "1"].v = formattedDate;
  //   console.log(worksheet[dateColumn + "1"].v)
  console.log("dateColumn =>", dateColumn + "1");
  // Adjust column widths
  worksheet["!cols"] = [
    { width: 15 },
    { width: 20 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
  ];
  // Save the changes to the same Excel file
  XLSX.writeFile(workbook, "data.xlsx");
} else {
  console.log("No matching cell found.");
}
