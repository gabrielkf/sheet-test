const { resolve, join } = require('path');

const ASSETS = resolve(__dirname, '..', '..', 'assets');
module.exports.ORIGINAL_FILE = join(ASSETS, 'template.xlsx');
module.exports.NEW_FILE = resolve(__dirname, '..', '..', 'out.xlsx');
module.exports.LOGO = join(ASSETS, 'bulbe.png');

module.exports.TAG_MARKERS = ['{', '}'];
module.exports.TAGS = [
  'representante',
  'representanteEndereco',
  'representanteCidade',
  'total',
  'vencimento',
  'locacao',
  'instalacao',
];

module.exports.CELLS = {
  B6: 'Julia Silva',
  B7: 'Av. dos Engenheiros, 1250/405',
  B8: 'Castelo - Belo Horizonte/MG',
  G4: 'R$ 5.4321,00',
  G7: '24/01/2022',
  D10: '12345',
  G10: '3005155000',
};
