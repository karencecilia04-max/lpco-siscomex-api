// Roda o RoboLPCO.gs na planilha simulada, do zero até a aba Novos cadastros preenchida.

import vm from 'node:vm';

import { carregarRobo } from './planilha_simulada.mjs';

const pasta = new URL('..', import.meta.url);

// Itens colados na aba Buscar: cada um cai num resultado diferente
export const ITENS_BUSCA = [
  ['295/80R22.5', 'MARCAUM', '', '', '', ''],
  ['175/70R13', 'MARCADOIS', 'URB 2', '', '', ''],
  ['LT225/75R16', 'MARCAUM', 'MD AT', '115/112S 10PR', '', ''],
  ['', '', '', '', 'PN-27580-TP500', ''],
  ['295/80R22.5', 'MARCAUM', 'MD500', '', '', '101'],
  ['295/80R22.5', 'MARCAUM', 'MD200', '152/148M 16PR', '', ''],
  ['315/80R22.5', 'MARCAUM', 'MD300', '156/150L 20PR', '', ''],
];

export function rodarRoteiro() {
  const { planilha, robo } = carregarRobo(new URL('RoboLPCO.gs', pasta), new URL('../exemplos/portal_ficticio.json', pasta));

  robo.prepararPlanilha();
  robo.gravarControle(planilha, 'CNPJ', '12345678000190');
  robo.atualizacaoDiaria();

  const buscar = planilha.getSheetByName('Buscar');
  buscar.getRange(9, 1, ITENS_BUSCA.length, ITENS_BUSCA[0].length).setValues(ITENS_BUSCA);
  robo.buscarLista();

  // A pessoa completa o INMETRO do primeiro novo cadastro e a descrição se atualiza
  const novos = planilha.getSheetByName('Novos cadastros');
  const N = vm.runInContext('N', robo);
  novos.getRange(9, N['Certificado INMETRO'] + 1, 1, 3).setValues([['OCP-0042/25', '001234/2025', '1A2B3C']]);
  robo.atualizarDescricoes();

  planilha.setActiveSheet(buscar);
  return { planilha, robo };
}
