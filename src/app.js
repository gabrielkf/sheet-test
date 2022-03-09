const XLSX = require('xlsx');
const { resolve } = require('path');

const filePath = resolve(__dirname, '..', 'assets', 'demo.xlsx');

function fromFile() {
  // const workbook = XLSX.readFile();
}

module.exports = { fromFile };
