// Testes da busca e dos novos cadastros do RoboLPCO.gs, fora do Google.
// As funções de busca são JavaScript puro: o arquivo é carregado num contexto do Node e
// recebe os produtos do catálogo fictício em exemplos/portal_ficticio.json.
// Rodar: node --test google_sheets/tests/busca.test.mjs

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const pasta = new URL('..', import.meta.url);
const gs = vm.createContext({ console });
vm.runInContext(readFileSync(new URL('RoboLPCO.gs', pasta), 'utf8'), gs);
const portal = JSON.parse(readFileSync(new URL('../exemplos/portal_ficticio.json', pasta), 'utf8'));

// Monta as linhas da aba Catálogo Produtos como o script faz depois de baixar a exportação
function linhasDoCatalogo() {
  vm.runInContext('var idxTeste = { raiz: "12345678", prodPorCodigo: new Map(), vincPorCodigo: new Map(), opPorChave: new Map() };', gs);
  const idx = gs.idxTeste;
  portal.produtos.forEach(p => idx.prodPorCodigo.set(Number(p.codigo), p));
  portal.vinculos.forEach(v => {
    const cod = Number(v.codigoProduto);
    if (!idx.vincPorCodigo.has(cod)) idx.vincPorCodigo.set(cod, []);
    idx.vincPorCodigo.get(cod).push(v);
  });
  portal.operadores.forEach(o => idx.opPorChave.set(`${o.codigoPais}|${o.codigo}`, o));

  const P = vm.runInContext('P', gs);
  return portal.produtos.map(p => {
    const linha = gs.linhaProduto(p.codigo, idx);
    const lpcos = Object.entries(portal.detalhes)
      .filter(([num, d]) => portal.consulta.find(c => c.numero === num).situacao === 'DEFERIDO' &&
        d.listaNcm.some(it => it.produto && it.produto.codigo === p.codigo))
      .map(([num]) => num);
    linha[P['Qtd LPCOs']] = lpcos.length;
    linha[P['LPCOs']] = lpcos.join(', ');
    return linha;
  });
}

const produtos = linhasDoCatalogo().map(l => gs.prepararProduto(l));
const P = vm.runInContext('P', gs);
const codigo = p => p.l[P['Código produto']];

function buscar(medida, marca = '', modelo = '', indice = '', interno = '', cod = '') {
  const filtro = gs.montarFiltro([medida, marca, modelo, indice, interno, cod]);
  return { filtro, achados: produtos.filter(p => gs.combina(p, filtro)).map(codigo) };
}

test('leitura do catálogo preenche medida, carga, lonas, marca, modelo e câmara', () => {
  const p = produtos.find(x => codigo(x) === 101).l;
  assert.equal(p[P['Medida']], '295/80R22.5');
  assert.equal(p[P['Índice de carga']], '154/149M');
  assert.equal(p[P['Lonas']], '18PR');
  assert.equal(p[P['Marca']], 'MARCAUM');
  assert.equal(p[P['Modelo']], 'MD200');
  assert.equal(p[P['Câmara']], 'TL');
});

test('busca exata pela medida e marca', () => {
  assert.deepEqual(buscar('295/80R22.5', 'MARCAUM').achados, [101]);
});

test('medida e índice podem vir juntos na mesma coluna', () => {
  assert.deepEqual(buscar('295/80 R 22,5 154/149M 18PR').achados, [101]);
});

test('busca pelo código interno', () => {
  assert.deepEqual(buscar('', '', '', '', 'PN-20555-EVECO').achados, [201]);
});

test('marca digitada por extenso bate com a marca do cadastro', () => {
  assert.deepEqual(buscar('205/55R16', 'MARCATRES PNEUS').achados, [201]);
});

test('marca diferente nunca é o mesmo pneu', () => {
  const { filtro, achados } = buscar('295/80R22.5', 'MARCADOIS', 'MD200');
  assert.deepEqual(achados, []);
  assert.ok(produtos.every(p => !gs.pontuar(p, filtro).possivel));
});

test('mesma medida, marca e modelo com outro índice aparece para análise', () => {
  const { filtro, achados } = buscar('295/80R22.5', 'MARCAUM', 'MD200', '152/148M 16PR');
  assert.deepEqual(achados, []);
  assert.deepEqual(gs.outroIndiceOuLonas(produtos, filtro).map(codigo), [101]);
});

test('sem resultado exato, o código do catálogo vira candidato com pontuação', () => {
  const { filtro, achados } = buscar('295/80R22.5', 'MARCAUM', 'OUTRO', '', '', '101');
  assert.deepEqual(achados, []);
  const r = gs.pontuar(produtos.find(p => codigo(p) === 101), filtro);
  assert.equal(r.possivel, true);
  assert.equal(r.pontos, 5 + 5 + 3);  // código, medida exata e marca
  assert.deepEqual([...r.faltam], ['modelo']);
});

test('medida de flotação não é cortada', () => {
  assert.deepEqual([...gs.lerMedidaCargaLonas('31X10.50R15LT 109S')], ['31X10.50R15LT', '109S', '']);
});

const N = vm.runInContext('N', gs);
function linhaNovo(campos) {
  const l = new Array(Object.keys(N).length).fill('');
  for (const [k, v] of Object.entries(campos)) l[N[k]] = v;
  return l;
}

test('descrição de novo cadastro no padrão do catálogo', () => {
  const r = gs.gerarDescricaoNovoProduto(linhaNovo({
    'NCM': '40112090', 'Medida': '295/80R22.5', 'Índice / Lonas': '154/149M 18PR', 'Marca': 'MARCAUM',
    'Modelo': 'MD300', 'Câmara (TL/TT)': 'TL', 'Certificado INMETRO': 'CERTIFICADO: OCP-0042/25',
    'Registro INMETRO': '001234/2025', 'Família INMETRO': '1A2B3C', 'Fabricante Código': 'OP0001',
  }));
  assert.equal(r.denominacao, 'PNEU DE CAMINHAO, MEDIDA: 295/80R22.5 154/149M 18PR, MARCA: MARCAUM');
  assert.match(r.descricao, /CONSTRUCAO RADIAL, MODELO: MD300, SEM CAMARA DE AR, CERTIFICADO: OCP-0042\/25,/);
  assert.deepEqual([...r.avisos], []);
});

test('pneu LT em NCM de caminhão usa o texto de light truck e avisa', () => {
  const r = gs.gerarDescricaoNovoProduto(linhaNovo({ 'NCM': '40112090', 'Medida': 'LT225/75R16', 'Marca': 'X' }));
  assert.match(r.denominacao, /^PNEU LIGHT TRUCK LT\/C/);
  assert.ok(r.avisos.some(a => /Exceção LT\/C/.test(a)));
  assert.ok(r.avisos.includes('Falta modelo'));
});

test('NCM estimada pelo tipo de medida', () => {
  assert.equal(gs.estimarNcm('205/55R16', '91V'), '40111000');
  assert.equal(gs.estimarNcm('295/80R22.5', '154/149M 18PR'), '40112090');
  assert.equal(gs.estimarNcm('LT225/75R16', '115/112S'), '40119090');
  assert.equal(gs.estimarNcm('18.4-30', '14PR'), '');
});

test('sugestões de NCM, câmara e fabricante vêm do catálogo', () => {
  const linhas = produtos.map(p => p.l);
  gs.fonteDeProdutos = () => null;
  gs.lerLinhas = () => linhas;
  const sug = new gs.Sugestoes({ getSheetByName: () => null });
  const novo = linhaNovo({ 'Medida': '295/80R22.5', 'Índice / Lonas': '152/148M', 'Marca': 'MARCAUM', 'Modelo': 'MD200' });
  const notas = sug.preencher(novo);
  assert.equal(novo[N['NCM']], '40112090');
  assert.equal(novo[N['Câmara (TL/TT)']], 'TL');
  assert.equal(novo[N['Fabricante Código']], 'OP0001');
  assert.ok(notas.some(n => /outro índice\/lonas: 101/.test(n)));
});
