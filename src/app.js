const xlsx = require('xlsx');
const { resolve, join } = require('path');

const fileDir = resolve(__dirname, '..', 'assets');
const filePath = join(fileDir, 'demo.xlsx');

function fromFile() {
  const workbook = xlsx.readFile(filePath);
  const plan1 = workbook.Sheets['Plan1'];
  const plan2 = workbook.Sheets['Plan2'];

  const data = xlsx.utils.sheet_to_json(plan1);

  const replaced = data.map(el => {
    if (el.modelo === '{tag}') {
      el.modelo = 'Replacement';
      return el;
    }

    return el;
  });

  const newBook = xlsx.utils.book_new();
  const newSheet = xlsx.utils.json_to_sheet(replaced);
  xlsx.utils.book_append_sheet(newBook, newSheet, 'NewData');

  xlsx.writeFile(newBook, 'edited.xlsx');
}

module.exports = { fromFile };
