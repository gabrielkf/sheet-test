const xlsx = require('sheetjs-style');
const { unlinkSync, existsSync, writeFileSync } = require('fs');
const { resolve, join } = require('path');

const fileDir = resolve(__dirname, '..', 'assets');
const filePath = join(fileDir, 'demo.xlsx');

async function editSheet() {
  const workbook = xlsx.readFile(filePath, {
    cellStyles: true,
    // sheetStubs: true,
  });

  const plan1 = workbook.Sheets['Plan1'];
  const plan2 = workbook.Sheets['Plan2'];

  // writeFileSync(join(fileDir, 'Plan1.json'), JSON.stringify(plan1));

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

  if (existsSync('edited.xlsx')) {
    unlinkSync('edited.xlsx');
  }

  xlsx.writeFile(newBook, 'edited.xlsx', {
    bookType: 'xlsx',
  });
}

module.exports = { editSheet };
