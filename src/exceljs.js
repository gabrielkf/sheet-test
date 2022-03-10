const excel = require('exceljs');
const { resolve, join } = require('path');
const { createWriteStream } = require('fs');

const fileDir = resolve(__dirname, '..', 'assets');
const originalFile = join(fileDir, 'demo.xlsx');

async function editSheet() {
  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile(originalFile);

  const TAG = '{tag}';

  const sheet = workbook.worksheets[0];
  // const cell = sheet.findCell('A12');
  // console.log(cell);

  for (var i = 1; i <= sheet.actualRowCount; i++) {
    for (var j = 1; j <= sheet.actualColumnCount; j++) {
      const cell = sheet.getRow(i).getCell(j);

      if (cell.text === TAG) {
        cell.value = 'Editado';
      }
    }
  }

  // const newWorkbook = new excel.Workbook();
  // newWorkbook.addWorksheet(sheet);
  // newWorkbook.xlsx.writeFile('editedAlt.xlsx');
  workbook.xlsx.writeFile('editedAlt.xlsx');
}

module.exports = { editSheet };
