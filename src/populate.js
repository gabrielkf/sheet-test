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
      //   .then(blob => {
      //   return blob.
      // });

      // return workbook.toFileAsync('./out.xlsx');
    })
    .then(data => {
      // console.log(data);
      convertToPdf(data);
    });

  // convertToPdf(data);
  // });

  // await convertToPdf(buffer);
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
