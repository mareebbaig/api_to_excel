var xl = require("excel4node");
// import axios from "axios";
const axios = require("axios");

var wb = new xl.Workbook();

// const data = [
//   {
//     name: "Shadab Shaikh",
//     email: "shadab@gmail.com",
//     mobile: "1234567890",
//   },
// ];

async function Start(page) {
  const data = await axios.get(
    `https://api.empireflippers.com/api/v1/listings/list?limit=200&page=${page}`,
  );
  var FileData = [];
  for (let i = 0; i < data.data.data.listings.length; i++) {
    let niche = "";
    data.data.data.listings[i].niches.forEach((n) => {
      niche += n.niche + " ";
    });
    let obj = {
      Id: data.data.data.listings[i].id,
      Title: data.data.data.listings[i].public_title,
      netProfit: data.data.data.listings[i].average_monthly_net_profit,
      grossrevenue: data.data.data.listings[i].average_monthly_gross_revenue,
      reasonForSale: data.data.data.listings[i].reason_for_sale,
      niches: niche,
      countries: data.data.data.listings[0].country,
    };
    FileData.push(obj);
  }
  const ws = wb.addWorksheet(`Worksheet Name${page}`);

  const headingColumnNames = [
    "Id",
    "Title",
    "netProfit",
    "grossrevenue",
    "reasonForSale",
    "niches",
    "countries",
  ];

  let headingColumnIndex = 2;
  headingColumnNames.forEach((heading) => {
    ws.cell(1, headingColumnIndex++).string(heading);
  });

  let rowIndex = 2;

  FileData.forEach((record) => {
    let columnIndex = 2;
    Object.keys(record).forEach((columnName) => {
      if (typeof record[columnName] == "string") {
        ws.cell(rowIndex, columnIndex).string(record[columnName]);
      } else if (typeof record[columnName] == "number") {
        ws.cell(rowIndex, columnIndex).number(record[columnName]);
      }
      columnIndex++;
    });
    rowIndex++;
  });

  wb.write(`ExcelFile${page}.xlsx`, function (err, stats) {
    if (err) {
      console.error(err);
    }
  });
}

Start(12);
