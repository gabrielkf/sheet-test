const Excel = require('exceljs');
const { createReadStream, createWriteStream } = require('fs');

const {
  CODES,
  TAGS,
  CELLS,
  ORIGINAL_FILE,
  LOGO,
} = require('./constants/replaceVariables');

const A4 = 9;

const writerOptions = {
  filename: '../editedStream.xlsx',
  useStyles: true,
  useSharedStrings: true,
};

// const writerOptions = {
//   filename: join(),
//   useStyles: true,
//   useSharedStrings: true,
// };

async function editSheet(replacements) {
  await workbook.xlsx.readFile(ORIGINAL_FILE, { useStyles: true });

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
  // Object.keys(CELLS).forEach(c => {
  //   const cell = sheet.findCell(c);
  //   cell.value = CELLS[c];
  // });

  addImage(workbook);

  workbook.xlsx.writeFile('editedAlt.xlsx');
}

async function addImage(workbook) {
  const sheet = workbook.worksheets[0];

  const logo = workbook.addImage({
    filename: LOGO,
    extension: 'png',
  });

  sheet.addImage(logo, 'B2: D4');

  workbook.xlsx.writeFile('withImage.xlsx');
}

module.exports = { editSheet };
