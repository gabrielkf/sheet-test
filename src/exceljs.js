const excel = require('exceljs');
const { resolve, join } = require('path');

const { CODES, TAGS } = require('./constants/replaceVariables');

const assets = resolve(__dirname, '..', 'assets');
const originalFile = join(assets, 'template.xlsx');
const logoFile = join(assets, 'bulbe.png');

const A4 = 9;

async function editSheet(replacements) {
  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile(originalFile);

  const sheet = workbook.worksheets[0];

  for (var i = 1; i <= sheet.rowCount; i++) {
    for (var j = 1; j <= sheet.columnCount; j++) {
      const cell = sheet.getRow(i).getCell(j);

      if (CODES.includes(cell.text)) {
        const key = TAGS.reduce((acc, t) => {
          acc = acc.replace(t, '');
          return acc;
        }, cell.text);

        cell.value = replacements[key];
      }
    }
  }

  addImage(workbook);

  workbook.xlsx.writeFile('editedAlt.xlsx');
}

async function addImage(workbook) {
  const sheet = workbook.worksheets[0];

  const logo = workbook.addImage({
    filename: logoFile,
    extension: 'png',
  });

  sheet.addImage(logo, 'B2: D4');

  workbook.xlsx.writeFile('withImage.xlsx');
}

module.exports = { editSheet };
