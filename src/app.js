const sheetJs = require('./sheetjs');
const excelJs = require('./exceljs');

const lib = 0 ? 'sheetjs' : 'exceljs';

module.exports = lib === 'sheetjs' ? sheetJs : excelJs;
