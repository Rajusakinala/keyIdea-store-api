const express = require("express");
const xlsx = require("xlsx");
const path = require("path");
const cors = require("cors");
// var bodyParser = require("body-parser");
const app = express();
app.use(express.json());
const PORT = 4000;
app.use(cors());
// app.use(bodyParser.json());
// Endpoint to get data from two sheets
app.get("/", (req, res) => {
  res.send("home page");
});
app.get("/test", (req, res) => {
  res.send("Test");
});
app.post("/get-excel-data", (req, res) => {
  console.log("req.body", req.body);
  const filePath = path.join(__dirname, "Wedding Products.xlsx");
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

  const pageLimit = req.query?.limit ?? 12;

  let result = firstSheetData.filter((ele) => {
    return (
      req.query.gender == ele.prodmeta_section &&
      req.body.price[0] < ele.attr_14k_regular &&
      ele.attr_14k_regular < req.body.price[1]
    );
  });
  if (req.body.styles.length === 0) {
  } else {
    result = result.filter((ele) => {
      return req.body.styles.some((keyword) =>
        ele.prod_subcategory.includes(keyword)
      );
    });
    // console.log("styl", styleResultfilter);
  }

  res.send({
    pageNumber: req.query.pageNumber,
    limit: pageLimit,
    totalPages: Math.round(result.length / pageLimit),
    totalmatchedRecords: result.length,
    data: result.slice(
      (req.query.pageNumber - 1) * pageLimit,
      req.query.pageNumber * pageLimit
    ),
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
