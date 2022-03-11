const sheetJs = require('./sheetjs');
const excelJs = require('./exceljs');
const populate = require('./populate');

const libs = [sheetJs, excelJs, populate];

const lib = libs[2];

module.exports = lib;
