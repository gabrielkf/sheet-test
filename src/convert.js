const fs = require('fs');
const Libre = require('libreoffice-convert');

const { NEW_FILE } = require('./constants/replaceVariables');

Libre.convertAsync = require('util').promisify(Libre.convert);

async function convertToPdf(sheetBuffer) {
  // const sheetBuffer = fs.readFileSync(NEW_FILE);

  const pdfBuffer = await Libre.convertAsync(
    sheetBuffer,
    '.pdf',
    undefined
  );

  fs.writeFileSync('./out.pdf', pdfBuffer);
}

module.exports = { convertToPdf };
