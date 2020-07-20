const _ = require("lodash");
const Excel = require("exceljs");
const path = require("path");

const turnArrayToObject = (original) => {
  let tmp_original = _.cloneDeep(original);
  let final_obj = {};
  for (let item of tmp_original) {
    item[item.length - 1] = parseInt(item[item.length - 1]);
    if (!final_obj[item[0]]) {
      final_obj[item[0]] = [];
      final_obj[item[0]].push(item);
    } else {
      final_obj[item[0]].push(item);
    }
  }
  Object.keys(final_obj).forEach((key) => {
    return final_obj[key].sort(function (a, b) {
      return b[b.length - 1] - a[a.length - 1];
    });
  });
  return final_obj;
};
const mergeCount = (original, start, price_idx) => {
  let new_original = _.cloneDeep(original);
  let obj = {},
    array = [];
  for (let item of new_original) {
    let id =
      start || price_idx
        ? item[start].substring(0, item[start].length - 2)
        : item[0];
    let key = start || price_idx ? id + "-" + item[price_idx] : id;
    if (!obj[key]) {
      array.push(item);
      obj[key] = item;
    } else {
      for (let arr of array) {
        if (start || price_idx) {
          if (
            arr[start] === item[start] &&
            arr[price_idx] === item[price_idx]
          ) {
            arr[arr.length - 1] =
              parseInt(arr[arr.length - 1]) + parseInt(item[item.length - 1]);
          }
        } else {
          if (arr[0] === item[0]) {
            arr[arr.length - 1] =
              parseInt(arr[arr.length - 1]) + parseInt(item[item.length - 1]);
            arr[arr.length - 2] =
              parseInt(arr[arr.length - 2]) + parseInt(item[item.length - 2]);
          }
        }
      }
    }
  }
  return array;
};

const countMergeArray = (
  client_values_positive,
  client_values_negative,
  us_values_positive,
  us_values_negative
) => {
  // 合并客户正数数组
  let client_merge_positive = mergeCount(client_values_positive, 0, 3);
  // 合并客户负数数组
  let client_merge_negative = mergeCount(client_values_negative, 0, 3);
  // 合并公司正数数组
  let us_merge_positive = mergeCount(us_values_positive, 0, 2);
  // 合并公司负数数组
  let us_merge_negative = mergeCount(us_values_negative, 0, 2);
  let client_merge_positive_map = turnArrayToObject(client_merge_positive);
  let client_merge_negative_map = turnArrayToObject(client_merge_negative);
  let us_merge_positive_map = turnArrayToObject(us_merge_positive);
  let us_merge_negative_map = turnArrayToObject(us_merge_negative);

  // 合并公司与客户的正数数组
  let merge_positive_array = signCountAllValues(
    client_merge_positive_map,
    us_merge_positive_map
  );
  // 合并公司与客户的负数数组
  let merge_negative_array = signCountAllValues(
    client_merge_negative_map,
    us_merge_negative_map
  );

  let merge_final_array = merge_positive_array.concat(merge_negative_array);
  let merge_final_map = {};
  /**
   * 因为数组是乱序
   * 偷懒没用任何算法
   * 直接先转换成object对象
   * 再通过Object.values转化成数组
   */

  merge_final_map = turnArrayToObject(merge_final_array);
  Object.keys(merge_final_map).forEach((key, idx) => {
    let count = idx % 2 === 0 ? 0 : 1;
    merge_final_map[key].map((item) => {
      item[0] = item[0] + "_" + count;
    });
  });
  let original_merge_array = [];
  Object.values(merge_final_map).forEach((array) => {
    for (let arr of array) {
      arr.splice(6, 2);
      original_merge_array.push(arr);
    }
  });
  return original_merge_array;
};
const countSignArray = (
  client_values_positive,
  client_values_negative,
  us_values_positive,
  us_values_negative
) => {
  let all_client_positive = turnArrayToObject(client_values_positive);
  let all_client_negative = turnArrayToObject(client_values_negative);
  let all_us_positive = turnArrayToObject(us_values_positive);
  let all_us_negative = turnArrayToObject(us_values_negative);

  let single_positive_array = signCountAllValues(
    all_client_positive,
    all_us_positive
  );

  let single_negative_array = signCountAllValues(
    all_client_negative,
    all_us_negative
  );

  let single_final_array = single_positive_array.concat(single_negative_array);
  let single_final_map = {};
  single_final_map = turnArrayToObject(single_final_array);
  Object.keys(single_final_map).forEach((key, idx) => {
    let count = idx % 2 === 0 ? 0 : 1;
    single_final_map[key].map((item) => {
      item[0] = item[0] + "_" + count;
    });
  });
  let original_single_array = [];
  Object.values(single_final_map).forEach((array) => {
    for (let arr of array) {
      arr.splice(6, 2);
      original_single_array.push(arr);
    }
  });
  return original_single_array;
};

const signCountAllValues = (client_group_map, us_group_map) => {
  let client_group_map_tmp = _.cloneDeep(client_group_map); //深拷贝对象，不影响原有的数据
  let us_group_map_tmp = _.cloneDeep(us_group_map);
  let tableData = [];
  Object.keys(client_group_map_tmp).map((key, index) => {
    //根据对象key循环
    let client_array = client_group_map_tmp[key];
    // 当客户拥有对应型号数据而公司数据里没有时
    // 根据当前key值查出来的数据每条合并对应型号，对照表中的单价，数量为0
    let us_array = us_group_map_tmp[key];
    if (!us_array) {
      for (let i of client_array) {
        let array = i.concat([key, global.maps[key][3], 0]);
        tableData.push(array);
      }
    } else {
      // 两边数量相等
      if (client_array.length === us_array.length) {
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        for (let i in client_array_back) {
          let array = [];
          if (us_array_back[i]) {
            us_array_back[i].splice(2, 1);
            array = client_array_back[i].concat(us_array_back[i]);
          } else {
            array = client_array_back[i].concat([key, global.maps[key][3], 0]);
          }
          tableData.push(array);
        }
      } else if (client_array.length > us_array.length) {
        //客户的订单数量大于公司的订单数量
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        let array = [];
        for (let i = 0; i < us_array_back.length; i++) {
          us_array_back[i].splice(2, 1);
          array = client_array_back[i].concat(us_array_back[i]);
          tableData.push(array);
        }
        let cut_array = client_array_back.splice(us_array_back.length);
        for (let i of cut_array) {
          array = i.concat([key, global.maps[key][3], 0]);
          tableData.push(array);
        }
      } else if (client_array.length < us_array.length) {
        // 客户的订单数量小于公司的订单数量
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        for (let i = 0; i < client_array_back.length; i++) {
          us_array_back[i].splice(2, 1);
          array = client_array_back[i].concat(us_array_back[i]);
          tableData.push(array);
        }
        let cut_array = us_array_back.splice(client_array_back.length);
        for (let i of cut_array) {
          i.splice(2, 1);
          let client = [
            i[0],
            global.maps[i[0]][1],
            global.maps[i[0]][0],
            0,
            global.maps[i[0]][3],
            0,
            ...i,
          ];
          tableData.push(client);
        }
      }
      delete us_group_map_tmp[key];
    }
    delete client_group_map_tmp[key];
  });

  /**
   * 完成筛选后还遗留下来的数据进行另外的操作
   */
  const left_client_map = Object.keys(client_group_map_tmp);
  const left_us_map = Object.keys(us_group_map_tmp);
  if (left_client_map.length > 0) {
    left_client_map.forEach((key) => {
      left_client_map[key].forEach((item) => {
        item.splice(0, 1);
        let arr = item.concat([key, global.maps[key][3], 0]);
        tableData.push(arr);
      });
    });
  }
  if (left_us_map.length > 0) {
    left_us_map.forEach((key) => {
      us_group_map_tmp[key].forEach((item) => {
        let backup_item = _.cloneDeep(item);
        backup_item.splice(2, 1);
        if (!global.maps[backup_item[0]]) {
          return false;
        }

        let client_number = global.maps[backup_item[0]][1];
        let client_specifications = global.maps[backup_item[0]][0];
        let client_price = global.maps[backup_item[0]][3];
        let arr = [
          backup_item[0],
          client_number,
          client_specifications,
          0,
          client_price,
          0,
          ...backup_item,
        ];
        tableData.push(arr);
      });
    });
  }
  return tableData;
};

const createNewWorkbook = async (
  tableHeader,
  merge_data,
  single_data,
  third_table
) => {
  const workbook = new Excel.Workbook();
  const workSheetOne = workbook.addWorksheet("合并", {
    properties: {
      defaultColWidth: 25,
    },
  });
  const workSheetTwo = workbook.addWorksheet("单行", {
    properties: {
      defaultColWidth: 25,
    },
  });
  const workSheetThree = workbook.addWorksheet("对账", {
    properties: {
      defaultColWidth: 25,
    },
  });
  workSheetOne.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: tableHeader,
    rows: merge_data,
  });
  workSheetTwo.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: tableHeader,
    rows: single_data,
  });
  workSheetThree.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: tableHeader,
    rows: third_table,
  });

  const sheet_array = [workSheetOne, workSheetTwo, workSheetThree];
  sheet_array.forEach((sheet, idx) => {
    sheet.eachRow(function (row, rowNumber) {
      let values = row.values;
      values.shift();
      if (rowNumber === 1) return;

      let type =
        idx < sheet_array.length - 1
          ? parseInt(values[0].slice(-2).substring(1))
          : rowNumber % 2 === 0
          ? 0
          : 1;
      if (idx < sheet_array.length - 1) {
        values[0] = values[0].substring(0, values[0].length - 2);
        row.values = values;
      }
      const columns = [1, 2, 3];
      const color = type === 1 ? "e0861a" : "ffce7b";
      columns.forEach((i) => {
        row.getCell(i).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: "FF" + color,
          },
        };
        row.getCell(i).font = {
          color: {
            argb: "FF130c0e",
          },
        };
      });
      if (values[values.length - 2] !== values[values.length - 1]) {
        row.getCell(values.length).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: "FFed1941",
          },
        };
        row.getCell(values.length).font = {
          color: {
            argb: "FFFFFFFB",
          },
        };
        row.getCell(values.length - 1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: "FFed1941",
          },
        };
        row.getCell(values.length - 1).font = {
          color: {
            argb: "FFFFFFFB",
          },
        };
      }
      if (values[3] !== values[4]) {
        row.getCell(4).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: "FF426ab3",
          },
        };
        row.getCell(4).font = {
          color: {
            argb: "FFFFFFFB",
          },
        };
        row.getCell(5).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: "FF426ab3",
          },
        };
        row.getCell(5).font = {
          color: {
            argb: "FFFFFFFB",
          },
        };
      }
    });
  });
  await workbook.xlsx.writeFile(
    path.join(__dirname, `../files/${global.currentFileName}.xlsx`)
  );
};

const normalTurnArrayToObject = (original) => {
  let tmp_original = _.cloneDeep(original);
  let final_obj = {};
  for (let item of tmp_original) {
    let key = item[0];
    item.splice(0, 1);
    if (!final_obj[key]) {
      final_obj[key] = [];
      final_obj[key].push(item);
    } else {
      final_obj[key].push(item);
    }
  }
  return final_obj;
};

const normalMergeCounts = (array) => {
  let obj = {},
    list = [];
  for (let arr of array) {
    let key = arr[0];
    if (!obj[key]) {
      list.push(arr);
      obj[key] = arr;
    } else {
      for (let item of list) {
        if (item[0] === arr[0]) {
          item[1] = arr[1] + item[1];
        }
      }
    }
  }
  return list;
};

const includeColumn = [
  "状态",
  "类型",
  "销售订单号",
  "客户型号",
  "本厂型号",
  "规格",
  "订单数量",
  "销售下单",
];

module.exports = {
  countMergeArray,
  countSignArray,
  createNewWorkbook,
  mergeCount,
  normalTurnArrayToObject,
  normalMergeCounts,
  includeColumn,
};
