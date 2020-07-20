const Excel = require("exceljs");
const _ = require("lodash");
const path = require("path");
const {
  countMergeArray,
  countSignArray,
  createNewWorkbook,
  mergeCount,
} = require("../util");

module.exports = {
  merge: async () => {
    const workbook = new Excel.Workbook();

    await workbook.xlsx.readFile(path.join(__dirname, "../upload/client.xlsx"));
    let worksheet1 = workbook.getWorksheet(1);
    await workbook.xlsx.readFile(
      path.join(__dirname, "../template/contrast.xlsx")
    );
    let worksheet2 = workbook.getWorksheet(1);
    await workbook.xlsx.readFile(path.join(__dirname, "../upload/us.xlsx"));
    let worksheet3 = workbook.getWorksheet(1);

    new Promise((resolve, reject) => {
      let contrast_map_client = {};
      let contrast_map_us = {};
      worksheet2.eachRow(function (row, rowNumber) {
        if (rowNumber === 1) {
          return;
        }
        const row_value = row.values;
        row_value.shift();
        contrast_map_client[row_value[1]] = row_value;
        contrast_map_us[row_value[2]] = row_value;
      });
      global.maps = contrast_map_us;
      resolve({ contrast_map_client, contrast_map_us });
    })
      .then((res) => {
        const { contrast_map_client, contrast_map_us } = res;
        let client_header = [];
        let client_values_positive = []; //正数
        let client_values_negative = []; //负数
        worksheet1.eachRow(function (row, rowNumber) {
          let values = row.values;
          values.shift();
          if (rowNumber === 1) {
            client_header = values;
            client_header[0] = "亮迪型号";
            client_header[1] = "亮迪规格";
            client_header[2] = "亮迪单价";
            client_header.splice(0, 0, "巨数型号");
            client_header.splice(4, 0, "巨数单价");
            return;
          }
          const sd = contrast_map_client[values[0]][2];
          const price = contrast_map_client[values[0]][3];
          values.splice(0, 0, sd);
          values.splice(4, 0, price);
          // values.splice(0, 0, rowNumber);
          if (values[values.length - 1] > 0) {
            client_values_positive.push(values);
          } else {
            client_values_negative.push(values);
          }
        });
        return {
          client_header,
          client_values_positive,
          client_values_negative,
          contrast_map_us,
        };
      })
      .then((res) => {
        let {
          client_header,
          client_values_positive,
          client_values_negative,
          contrast_map_us,
        } = res;

        let us_header = [];
        let us_values_positive = [];
        let us_values_negative = [];

        worksheet3.eachRow(function (row, rowNumber) {
          let values = row.values;
          values.shift();
          if (rowNumber === 1) {
            us_header = values;
            return;
          }
          const price = contrast_map_us[values[0]]
            ? contrast_map_us[values[0]][3]
            : 0;
          values.splice(2, 0, price);
          // values.splice(0, 0, rowNumber);
          if (values[values.length - 1] > 0) {
            us_values_positive.push(values);
          } else {
            us_values_negative.push(values);
          }
        });
        let client_positive_tmp = _.cloneDeep(client_values_positive);
        let client_negative_tmp = _.cloneDeep(client_values_negative);
        let us_positive_tmp = _.cloneDeep(us_values_positive);
        let us_negative_tmp = _.cloneDeep(us_values_negative);
        let final_header = client_header.concat(us_header);
        final_header = final_header.map((item) => ({ name: item }));
        let original_single_array = countSignArray(
          client_values_positive,
          client_values_negative,
          us_values_positive,
          us_values_negative
        );
        let original_merge_array = countMergeArray(
          client_positive_tmp,
          client_negative_tmp,
          us_positive_tmp,
          us_negative_tmp
        );
        final_header.splice(6, 2);
        let third_table_array_merge = _.cloneDeep(original_merge_array);
        third_table_array_merge.map((item) => {
          item[0] = item[0].substring(0, item[0].length - 2);
        });
        let third_table = mergeCount(third_table_array_merge);
        createNewWorkbook(
          final_header,
          original_merge_array,
          original_single_array,
          third_table
        );
      });
  },
};
