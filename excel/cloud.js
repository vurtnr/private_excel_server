const Excel = require("exceljs");
const dayjs = require("dayjs");
const path = require('path')
const { includeColumn } = require('../util')


module.exports = {
  cloud: async () => {
    const writeWorkBook = new Excel.Workbook();
    await writeWorkBook.xlsx.readFile(path.join(__dirname, "../template/tracking.xlsx"));
    let worksheet = writeWorkBook.getWorksheet(1);
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, "../upload/order_detail.xlsx"));
    let detailSheet = workbook.getWorksheet(1);
    await workbook.xlsx.readFile(path.join(__dirname, "../template/contrast.xlsx"));
    let contrastSheet = workbook.getWorksheet(1);

    new Promise((resolve, reject) => {
      const map = {};
      contrastSheet.eachRow(function (row, rowNumber) {
        if (rowNumber === 1) {
          return;
        }
        const row_value = row.values;
        row_value.shift();
        map[row_value[2]] = row_value;
      });
      resolve(map);
    }).then(async (res) => {
      const header = worksheet.getRow(1).values;
      header.shift();
      const header_array = [];
      header.forEach((o, i) => {
        if (typeof o === "object") {
          header_array.push(o.richText[0].text);
        } else {
          header_array.push(o);
        }
      });
      const idx_array = [];
      header_array.forEach((o, i) => {
        includeColumn.includes(o) && idx_array.push(i);
      });
      const final_list = [];
      detailSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          return;
        }
        let arr = new Array(header.length).fill("");
        const values = row.values;
        values.shift();
        const contrastObj = res[values[2]];
        const client_model = contrastObj[1];
        const type = contrastObj[contrastObj.length - 1];
        const status = values[values.length - 1] === "Y" ? "Y" : "";
        const date = dayjs(values[0]).format("M/D");
        let value_arr = [
          status,
          type,
          values[1],
          client_model,
          values[2],
          values[4],
          parseInt(values[5]),
          date,
        ];
        idx_array.forEach((o, i) => {
          arr[o] = value_arr[i];
        });
        final_list.push(arr);
      });
      for (let l of final_list) {
        worksheet.addRow(l);
      }
      await writeWorkBook.xlsx.writeFile(
        path.join(__dirname, `../files/${global.currentFileName}.xlsx`)
      );
    });
  }
}