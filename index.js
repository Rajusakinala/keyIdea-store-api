const express = require("express");
const xlsx = require("xlsx");
const path = require("path");
const cors = require("cors");
const app = express();
const PORT = 3000;
app.use(cors());
// Endpoint to get data from two sheets
app.get("/get-excel-data", (req, res) => {
  const filePath = path.join(__dirname, "Wedding Products.xlsx"); // Replace with your actual file path
  const workbook = xlsx.readFile(filePath);

  // Access the sheets by their names or indexes
  const firstSheetName = workbook.SheetNames[0]; // First sheet
  const secondSheetName = workbook.SheetNames[1]; // Second sheet

  const firstSheet = workbook.Sheets[firstSheetName];
  const secondSheet = workbook.Sheets[secondSheetName];

  // Convert sheet data to JSON format
  const firstSheetData = xlsx.utils.sheet_to_json(firstSheet);
  const secondSheetData = xlsx.utils.sheet_to_json(secondSheet);

  // Send the data as a response
  const totalPages = 8;
  const pageLimit = 12;

  res.send({
    pageNumber: req.query.pageNumber,
    // totalPages: Math.round(firstSheetData.length / pageLimit),
    limit: pageLimit,
    totalmatchedRecords: firstSheetData.filter((ele) => {
      return req.query.gender == ele.prodmeta_section;
    }).length,
    data: firstSheetData
      .filter((ele) => {
        return req.query.gender == ele.prodmeta_section;
      })
      .slice(
        (req.query.pageNumber - 1) * pageLimit,
        req.query.pageNumber * pageLimit
      ),
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
