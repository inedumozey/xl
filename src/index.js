const excelToJson = require("./excelToJson");
const jsonToExcel = require("./jsonToExcel");

const xl = { jsonToExcel, excelToJson };

module.exports = { xl }