const Excel = require("exceljs");
const path = require('path')
const _ = require("lodash");
const { normalMergeCounts,normalTurnArrayToObject } = require('../util')


module.exports = {
  years_data: async () => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, "../upload/years_data.xlsx"));
    let worksheet1 = workbook.getWorksheet(1);
    let all_array = [];
    worksheet1.eachRow((row, rowNumber) => {
      let values = row.values;
      values.shift();
      if (rowNumber === 1) return;
      let year = values[0].substring(0, 4);
      let month = values[0].substring(5, 7);

      let arr = [
        values[2],
        parseInt(year),
        parseInt(month),
        parseInt(values[5]),
      ];
      all_array = [...all_array, arr];
    });
    let obj = normalTurnArrayToObject(all_array);
    let all_years_data = [];
    Object.keys(obj).map((key) => {
      let temp = normalTurnArrayToObject(obj[key]);
      Object.keys(temp).forEach((item) => {
        let temp_list = normalMergeCounts(temp[item]).sort(
          (a, b) => a[0] - b[0]
        );
        let newArray = new Array(12);
        for (let tem of temp_list) {
          newArray[tem[0] - 1] = tem[1];
        }
        let final_arr = [key, item, ...newArray];
        all_years_data.push(final_arr);
      });
    });
    let header = [
      "产品型号",
      "年份",
      "1月",
      "2月",
      "3月",
      "4月",
      "5月",
      "6月",
      "7月",
      "8月",
      "9月",
      "10月",
      "11月",
      "12月",
    ];
    header = header.map((i) => ({ name: i }));
    const writeWorkbook = new Excel.Workbook();
    const workSheetOne = writeWorkbook.addWorksheet("数据一览", {
      properties: {
        defaultColWidth: 15,
      },
    });
    workSheetOne.addTable({
      name: "MyTable",
      ref: "A1",
      headerRow: true,
      columns: header,
      rows: all_years_data,
    });
    await writeWorkbook.xlsx.writeFile(
      path.join(__dirname, `../files/${global.currentFileName}.xlsx`)
    );
  }
}