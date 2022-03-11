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
