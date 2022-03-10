const App = require('./src/app');

const replacement = {
  representante: 'Julia Silva',
  representanteEndereco: 'Av. dos Engenheiros, 1250/405',
  representanteCidade: 'Castelo - Belo Horizonte/MG',
  total: 'R$ 5.4321,00',
  vencimento: '24/01/2022',
  locacao: '12345',
  instalacao: '3005155000',
};

let run = (async function () {
  await App.editSheet(replacement);
})();
