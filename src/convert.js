const fs = require('fs');
const Libre = require('libreoffice-convert');

Libre.convertAsync = require('util').promisify(Libre.convert);

async function convertToPdf(sheetBuffer) {
  const pdfBuffer = await Libre.convertAsync(
    sheetBuffer,
    '.pdf',
    undefined
  );

  fs.writeFileSync('./out.pdf', pdfBuffer);
}

module.exports = { convertToPdf };
