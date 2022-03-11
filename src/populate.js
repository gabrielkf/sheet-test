const Populate = require('xlsx-populate');

const {
  CODES,
  TAGS,
  CELLS,
  ORIGINAL_FILE,
  LOGO,
} = require('./constants/replaceVariables');

async function editSheet() {
  Populate.fromFileAsync(ORIGINAL_FILE).then(workbook => {
    replaceByCell(workbook);

    return workbook.toFileAsync('./out.xlsx');
  });
}

function replaceByCell(workbook) {
  Object.keys(CELLS).forEach(k => {
    workbook.sheet('Bulbe').cell(k).value(CELLS[k]);
  });
}

module.exports = { editSheet };
