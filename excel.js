const express = require("express");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const app = express();
const port = 3000;
const fs=require('fs')
const filePath = './excelData1.json';


app.use(bodyParser.json());

app.post("/api/jsonToExcel", (req, res) => {
  const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  const workSheet = xlsx.utils.json_to_sheet(jsonData);
  const workBook=xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workBook,workSheet,"sheetName")
  xlsx.write(workBook,{bookType:"xlsx",type:"buffer"})
  xlsx.write(workBook,{bookType:"xlsx",type:"binary"})
  xlsx.writeFile(workBook,"XLFILE.xlsx")
  res.send("json to excel convertion is done")
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
