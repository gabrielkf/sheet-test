const Populate = require('xlsx-populate');
const { convertToPdf } = require('./convert');

const {
  TAGS,
  TAG_MARKERS,
  CELLS,
  ORIGINAL_FILE,
} = require('./constants/replaceVariables');

async function editSheet(replacements) {
  Populate.fromFileAsync(ORIGINAL_FILE)
    .then(workbook => {
      replaceByTag(workbook, replacements);

      const [sheet] = workbook.sheets();
      const hCenter = sheet.printOptions('horizontalCentered');
      const printGrid = sheet.printGridLines();
      const margins = {
        top: sheet.pageMargins('top'),
        right: sheet.pageMargins('right'),
        bottom: sheet.pageMargins('bottom'),
        left: sheet.pageMargins('left'),
      };

      // console.log(hCenter);
      // console.log(printGrid);
      // console.log(margins);
      // console.log(sheet.pageBreaks());

      return workbook.outputAsync();
    })
    .then(buffer => {
      convertToPdf(buffer);
    });
}

function replaceByTag(workbook, replacements) {
  TAGS.forEach(tag =>
    workbook.find(TAG_MARKERS.join(tag), replacements[tag])
  );
}

function replaceByCell(workbook) {
  Object.keys(CELLS).forEach(k => {
    workbook.sheet('Bulbe').cell(k).value(CELLS[k]);
  });
}

module.exports = { editSheet };
