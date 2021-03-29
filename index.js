const XLSX = require("xlsx");
const fs = require("fs");
const workbook = XLSX.readFile("test.xlsx");

// return array of sheets present in the excel file
const worksheet = workbook.SheetNames;
const firstSheet = worksheet[0];

// select the sheet which you want to read
const selectedSheet = workbook.Sheets[firstSheet];

// console.log("selectedSheet", selectedSheet);
let supplier = [];

Object.keys(selectedSheet).forEach((sheet, index) => {
  if (
    selectedSheet[`A${index + 1}`] !== undefined &&
    selectedSheet[`B${index + 1}`] !== undefined &&
    selectedSheet[`C${index + 1}`] !== undefined
  ) {
    supplier.push({
      name: selectedSheet[`A${index + 1}`].v,
      category: selectedSheet[`B${index + 1}`].v,
      lag: selectedSheet[`C${index + 1}`].v,
    });
  }
});
console.log("supplier", supplier);
// var stream = fs.createWriteStream("query.txt", { flags: "a" });

// supplier.forEach((element, index) => {
//   let sql = `INSERT INTO Dim_Normalized_Supplier  VALUES (${index + 1},'${
//     element.supplier
//   }','Kumar, Abhishek','abhishek.kumar3@tyson.com','2021-03-26 19:02:58','2021-03-26 19:02:58');`;

//   stream.write(sql + "\n");
// });
// stream.end();
