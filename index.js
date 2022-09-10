// Requiring module
const reader = require("xlsx");
const excel = require("exceljs");
const express = require("express");
const bodyParser = require("body-parser");
const app = express();
const port = 3000;

// Reading our test file
// const file = reader.readFile("./Data/test.xlsx");

const ds1 = require("./Data/sheet1");
const ds2 = require("./Data/sheet2");

app.use(bodyParser.urlencoded({ extended: false }));

// parse application/json
app.use(bodyParser.json());

// const ws = reader.utils.json_to_sheet(ds1);
// const ws2 = reader.utils.json_to_sheet(ds2);

// reader.utils.book_append_sheet(file, ws, "Raw");
// reader.utils.book_append_sheet(file, ws2, "Sheet3");

// Writing to our file
// reader.writeFile(file, "./Data/test.xlsx");

// console.log(ds1);

const exportUser = async (req, res) => {
  // WRITE DOWNLOAD EXCEL LOGIC
};
// module.exports = exportUser;

app.post("/", async (req, res) => {
  let data = req.body;

  // let {header , data } = req.body

  //it can be passed in json as well
  let workSheeetColumnDets = [
    [
      { header: "Id", key: "id", width: 15 },
      { header: "Date", key: "Date", width: 25 },
    ],
    [
      { header: "Id", key: "id", width: 15 },
      { header: "Amount", key: "Amount", width: 25 },
    ],
  ];
  let workbook = new excel.Workbook();
  data.forEach(async (singleSheet, index) => {
    let worksheet = workbook.addWorksheet("sheet" + (index + 1));
    console.log(singleSheet);
    worksheet.columns = workSheeetColumnDets[index];
    await worksheet.addRows(singleSheet);
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=" + "tutorials.xlsx"
  );

  await workbook.xlsx.write(res);

  res.status(200).end();
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});

//-------------------------------------------------------Code for erxcel download dfrom json---------------------------
//   let worksheet1 = workbook.addWorksheet("ds1");
//   let worksheet2 = workbook.addWorksheet("ds2");
//   worksheet1.columns = [
//     { header: "Id", key: "id", width: 15 },
//     { header: "Date", key: "Date", width: 25 },
//   ];
//   worksheet2.columns = [
//     { header: "Id", key: "id", width: 15 },
//     { header: "Amount", key: "Amount", width: 25 },
//   ];
//   // Add Array Rows
//   await worksheet1.addRows(ds1);
//   await worksheet2.addRows(ds2);

//-------------------------------------------------------Code for erxcel download dfrom json---------------------------
