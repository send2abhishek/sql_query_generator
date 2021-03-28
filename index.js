const XLSX = require("xlsx");
const fs = require("fs");
const workbook = XLSX.readFile("test.xlsx");

// return array of sheets present in the excel file
const worksheet = workbook.SheetNames;
const firstSheet = worksheet[0];

// select the sheet which you want to read
const selectedSheet = workbook.Sheets[firstSheet];

let supplier = [];

Object.keys(selectedSheet).forEach((sheet) => {
  if (selectedSheet[sheet].v !== undefined) {
    supplier.push({
      supplier: selectedSheet[sheet].v,
    });
  }
});

var stream = fs.createWriteStream("query.txt", { flags: "a" });

supplier.forEach((element, index) => {
  let sql = `INSERT INTO Dim_Normalized_Supplier  VALUES (${index + 1},'${
    element.supplier
  }','Kumar, Abhishek','abhishek.kumar3@tyson.com','2021-03-26 19:02:58','2021-03-26 19:02:58');`;

  stream.write(sql + "\n");
});
stream.end();
