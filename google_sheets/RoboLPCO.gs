// Robô LPCO: Portal Único Siscomex -> Google Sheets
//
// Mantém numa planilha do Google as LPCOs deferidas do CNPJ nas NCMs da aba Controle, os produtos
// do Catálogo de Produtos que aparecem nelas e o catálogo completo das NCMs escolhidas. Em cima
// desses dados, a aba Buscar responde se um item já tem cadastro e LPCO, e a aba Novos cadastros
// monta a denominação e a descrição dos itens que ainda precisam ser cadastrados.
//
// Instalar: numa planilha em branco, Extensões > Apps Script, colar este arquivo, salvar e
// recarregar a planilha. Menu Robô LPCO > Preparar planilha, preencher CNPJ e NCMs na aba
// Controle e depois Configuração > Configurar chave de acesso.
// A chave fica nas propriedades do usuário, nunca numa célula.

const ABA_PRODUTOS = 'Itens por LPCO';
const ABA_LPCOS = 'Banco de LPCO';
const ABA_CONTROLE = 'Controle';

const AMBIENTES = {
  producao: 'https://portalunico.siscomex.gov.br',
  validacao: 'https://val.portalunico.siscomex.gov.br',
};
const LIMITE_EXECUCAO_MS = 5 * 60 * 1000;  // Apps Script derruba em 6 min
const PAUSA_ENTRE_CHAMADAS_MS = 1000;
const DIAS_POR_JANELA = 30;

const CAB_PRODUTOS = [
  'Código produto', 'Versão na LPCO', 'Versão atual catálogo', 'Situação produto', 'NCM',
  'Descrição', 'Descrição detalhada', 'Código interno',
  'Cód. operador estrangeiro', 'Nome operador estrangeiro', 'País operador', 'TIN operador',
  'Qtd LPCOs', 'LPCOs', 'CNPJ raiz', 'Observação',
  'Medida', 'Índice de carga', 'Lonas', 'Marca', 'Modelo', 'Fim vigência LPCO mais recente', 'Câmara',
];
const CAB_LPCOS = [
  'Nº LPCO', 'Situação', 'Data registro', 'Início vigência', 'Fim vigência', 'Item', 'NCM',
  'Código produto', 'Versão produto', 'CNPJ raiz produto', 'DUIMP vinculada', 'Incluída em',
];
// Colunas (base 0) da aba Produtos
const P = {};
CAB_PRODUTOS.forEach((n, i) => { P[n] = i; });


// MENU
function onOpen() {
  // Buscar, limpar, trazer para Novos e atualizar descrições ficam nas caixas das próprias abas
  // (e o envio para Novos e a descrição são automáticos). O menu tem só o que não cabe nelas.
  const ui = SpreadsheetApp.getUi();
  const config = ui.createMenu('Configuração')
    .addItem('Configurar chave de acesso', 'configurarChave')
    .addItem('Recriar aba Buscar', 'criarAbaBuscar')
    .addItem('Recriar aba Novos cadastros', 'criarAbaNovos')
    .addSeparator()
    .addItem('Agendar atualização diária (7h)', 'agendarDiario')
    .addItem('Cancelar agendamento', 'cancelarAgendamento');
  ui.createMenu('Robô LPCO')
    .addItem('Preparar planilha', 'prepararPlanilha')
    .addItem('Atualizar agora (LPCOs e catálogo)', 'atualizacaoDiaria')
    .addItem('Baixar Excel de novos cadastros', 'baixarExcelNovos')
    .addSeparator()
    .addSubMenu(config)
    .addToUi();
}

// Menu: numa planilha em branco, cria a aba Controle e as abas de busca
function prepararPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(ABA_CONTROLE)) criarControle(ss);
  if (!ss.getSheetByName(ABAS_BUSCA.busca)) criarAbaBuscar();
  else if (!ss.getSheetByName(ABAS_BUSCA.novos)) criarAbaNovos();
  aplicarVisual();
  ss.setActiveSheet(ss.getSheetByName(ABA_CONTROLE));
  avisar('Planilha pronta. Preencha CNPJ e NCMs na aba Controle e configure a chave de acesso ' +
         '(Robô LPCO > Configuração).');
}

function configurarChave() {
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt('Chave de acesso do Portal Único', 'Client-Id (chave Pessoa Física, perfil IMPEXP):',
    ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK) return;
  const segredo = ui.prompt('Chave de acesso do Portal Único', 'Client-Secret:', ui.ButtonSet.OK_CANCEL);
  if (segredo.getSelectedButton() !== ui.Button.OK) return;
  const limpar = t => t.trim().replace(/^["']|["']$/g, '');
  const valores = { LPCO_CLIENT_ID: limpar(id.getResponseText()), LPCO_CLIENT_SECRET: limpar(segredo.getResponseText()) };
  if (valores.LPCO_CLIENT_ID === valores.LPCO_CLIENT_SECRET) {
    ui.alert('Client-Id e Client-Secret estão iguais. Confira e configure de novo.');
    return;
  }
  PropertiesService.getUserProperties().setProperties(valores);
  // Testa na hora, para não descobrir o erro só na atualização
  try {
    new PortalUnico(lerConfig(SpreadsheetApp.getActiveSpreadsheet())).autenticar(true);
    ui.alert('Chave salva só para o seu usuário e aceita pelo Portal Único. Agora use "Atualizar LPCOs deferidas".');
  } catch (e) {
    ui.alert(`Chave salva, mas o Portal Único recusou (Client-Id com ${valores.LPCO_CLIENT_ID.length} caracteres, ` +
             `Client-Secret com ${valores.LPCO_CLIENT_SECRET.length}).\n\n${e.message}`);
  }
}

function agendarDiario() {
  cancelarAgendamento();
  ScriptApp.newTrigger('atualizacaoDiaria').timeBased().everyDays(1).atHour(7).create();
  avisar('Atualização das LPCOs e do catálogo agendada para todo dia por volta das 7h.');
}

// Menu "Atualizar agora" e agendamento diário: LPCOs e catálogo, com um aviso só no fim
function atualizacaoDiaria() {
  avisosAcumulados = [];
  try {
    atualizarLpcos();
    atualizarCatalogo();
  } finally {
    const resumo = avisosAcumulados.join('\n\n');
    avisosAcumulados = null;
    avisar(resumo);
  }
}

function cancelarAgendamento() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['atualizarLpcos', 'atualizacaoDiaria'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

let avisosAcumulados = null;  // quando é uma lista, os avisos são juntados e mostrados de uma vez

function avisar(msg) {
  console.log(msg);
  if (avisosAcumulados) { avisosAcumulados.push(msg); return; }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* rodando por agendamento */ }
}


// ATUALIZAÇÃO
// Só uma atualização por vez na planilha (outra pessoa ou o agendamento podem estar rodando)
function comTrava(nome, acao) {
  const trava = LockService.getDocumentLock();
  if (!trava.tryLock(15000)) {
    avisar(`Outra atualização está em andamento nesta planilha. Tente "${nome}" de novo em alguns minutos.`);
    return;
  }
  try { acao(); } finally { trava.releaseLock(); }
}

function atualizarLpcos() { comTrava('Atualizar LPCOs deferidas', atualizarLpcosSemTrava); }
function atualizarCatalogo() { comTrava('Atualizar catálogo de produtos', atualizarCatalogoSemTrava); }

function atualizarLpcosSemTrava() {
  const inicioExec = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = lerConfig(ss);
  const hoje = inicioDoDia(new Date());

  // Busca a partir da última extração, voltando alguns dias de margem: uma LPCO
  // registrada antes pode ter sido deferida depois da última extração.
  let dataInicial = cfg.buscarDesde || hoje;
  if (cfg.ultimaExtracao) {
    dataInicial = new Date(inicioDoDia(cfg.ultimaExtracao).getTime() - cfg.diasMargem * 86400000);
  }
  const periodo = `${fmtData(dataInicial)} a ${fmtData(hoje)}`;

  const abaLpcos = obterAba(ss, ABA_LPCOS, CAB_LPCOS);
  const conhecidas = new Set(lerLinhas(abaLpcos).map(l => String(l[0])));
  const api = new PortalUnico(cfg);

  let resultado = 'OK', novasLinhas = [], novasLpcos = 0, completo = true;
  try {
    api.autenticar();
    const lista = listarDeferidas(api, cfg, dataInicial, hoje).filter(l => !conhecidas.has(l.numero));
    console.log(`${lista.length} LPCOs deferidas novas no período ${periodo}`);
    const agora = new Date();
    for (const l of lista) {
      if (Date.now() - inicioExec > LIMITE_EXECUCAO_MS) {
        completo = false;
        resultado = `Parcial: tempo esgotado (faltaram ${lista.length - novasLpcos} LPCOs). Rode de novo.`;
        break;
      }
      const det = api.get(`/talpco/api/ext/lpco/${l.numero}`);
      novasLinhas = novasLinhas.concat(linhasDaLpco(l, det, cfg.ncms, agora));
      novasLpcos++;
    }
  } catch (e) {
    completo = false;
    resultado = (e.limite ? 'Limite de acessos do portal: ' : 'Erro: ') + e.message;
    console.error(e.stack || e);
  }

  if (novasLinhas.length) {
    garantirLinhas(abaLpcos, abaLpcos.getLastRow() + novasLinhas.length);
    abaLpcos.getRange(abaLpcos.getLastRow() + 1, 1, novasLinhas.length, CAB_LPCOS.length).setValues(novasLinhas);
  }

  // Produtos novos ou atualizados: baixa o catálogo (3 chamadas) e refaz essas linhas
  const codigosNovos = [...new Set(novasLinhas.map(l => Number(l[7])).filter(Boolean))];
  if (codigosNovos.length) {
    try {
      const catalogos = atualizarProdutos(ss, api, cfg, codigosNovos, novasLinhas);
      // A exportação já foi baixada: aproveita para atualizar a aba Catálogo, se ela existir
      const daEmpresa = catalogos.find(c => c.raiz === cfg.cnpj.slice(0, 8));
      if (daEmpresa && ss.getSheetByName(ABA_CATALOGO)) gravarCatalogo(ss, cfg, daEmpresa);
    } catch (e) {
      completo = false;
      resultado += ` | Catálogo não atualizado: ${e.message}`;
      console.error(e.stack || e);
    }
  }
  recalcularLpcos(ss, ABA_PRODUTOS);
  recalcularLpcos(ss, ABA_CATALOGO);
  try { aplicarVisual(); } catch (e) { console.warn(e.message); }  // estende as faixas às linhas novas

  // Só avança a data da última extração quando a busca do período terminou inteira
  if (completo) gravarControle(ss, 'Última extração', new Date());
  gravarControle(ss, 'Período consultado na última execução', periodo);
  gravarControle(ss, 'LPCOs novas na última execução', novasLpcos);
  registrarHistorico(ss, [new Date(), periodo, novasLpcos, novasLinhas.length, codigosNovos.length, resultado]);
  try { mostrarDatas(ss); } catch (e) { console.warn(e.message); }

  avisar(`${resultado}\n\nPeríodo: ${periodo}\nLPCOs deferidas novas: ${novasLpcos}` +
         `\nItens incluídos: ${novasLinhas.length}\nProdutos atualizados: ${codigosNovos.length}`);
}

function listarDeferidas(api, cfg, inicio, fim) {
  const achadas = new Map();
  for (let ini = new Date(inicio); ini <= fim; ini = new Date(ini.getTime() + DIAS_POR_JANELA * 86400000)) {
    const fin = new Date(Math.min(ini.getTime() + (DIAS_POR_JANELA - 1) * 86400000, fim.getTime()));
    for (const ncm of cfg.ncms) {
      let offset = 0;
      while (true) {
        const dados = api.get('/talpco/api/ext/lpco/consulta', {
          'tipo-operacao': 'IMPORTACAO',
          'importador-exportador': cfg.cnpj,
          'situacao': 'DEFERIDO',
          'data-inicial-registro': fmtIso(ini),
          'data-final-registro': fmtIso(fin),
          'ncm-inicial': ncm,
          'ncm-final': ncm,
          'offset': offset,
        });
        if (!dados || !dados.length) break;
        const novos = dados.filter(d => !achadas.has(d.numero));
        novos.forEach(d => achadas.set(d.numero, d));
        if (!novos.length) break;  // evita laço se o portal ignorar o offset
        offset += dados.length;
      }
    }
  }
  return [...achadas.values()];
}

function linhasDaLpco(resumo, det, ncms, agora) {
  const itens = (det && det.listaNcm) || [];
  const linhas = [];
  for (const it of itens) {
    if (ncms.length && ncms.indexOf(String(it.ncm)) < 0) continue;
    const prod = it.produto || {};
    linhas.push([
      resumo.numero, resumo.situacao || 'Deferido', paraData(resumo.dataRegistro),
      paraData(resumo.dataInicioVigencia), paraData(resumo.dataFimVigencia),
      it.numeroItem, String(it.ncm || ''), prod.codigo || '', prod.versao || '', prod.cnpjRaiz || '',
      resumo.dueVinculada || '', agora,
    ]);
  }
  return linhas;
}


// ABAS PRODUTOS E CATÁLOGO (mesmas colunas, modelo Catalogo_Produtos_LPCOs)
// Produtos: só os produtos que aparecem em LPCOs deferidas.
// Catálogo: tudo que está cadastrado nas "NCMs do catálogo" da aba Controle.
const ABA_CATALOGO = 'Catálogo Produtos';

// Exportação completa do catálogo de um CNPJ raiz: 3 chamadas ao portal
function baixarCatalogo(api, raiz) {
  const prods = api.exportar(`/catp/api/ext/produto/exportar/${raiz}/true`);
  const vinculos = api.exportar(`/catp/api/ext/fabricante/exportar/${raiz}`);
  const operadores = api.exportar(`/catp/api/ext/operador-estrangeiro/exportar/${raiz}/true`);
  const idx = { raiz, prods, operadores, prodPorCodigo: new Map(), vincPorCodigo: new Map(), opPorChave: new Map() };
  prods.forEach(p => idx.prodPorCodigo.set(Number(p.codigo), p));
  vinculos.forEach(v => {
    const cod = Number(v.codigoProduto);
    if (!idx.vincPorCodigo.has(cod)) idx.vincPorCodigo.set(cod, []);
    idx.vincPorCodigo.get(cod).push(v);
  });
  operadores.forEach(o => idx.opPorChave.set(`${o.codigoPais}|${o.codigo}`, o));
  return idx;
}

function atualizarProdutos(ss, api, cfg, codigos, novasLinhas) {
  const raizes = [...new Set(novasLinhas.map(l => String(l[9] || cfg.cnpj.slice(0, 8)).padStart(8, '0')))];
  const aba = obterAba(ss, ABA_PRODUTOS, CAB_PRODUTOS);
  const linhas = lerLinhas(aba);
  const posicao = new Map(linhas.map((l, i) => [Number(l[P['Código produto']]), i]));
  const catalogos = [];
  for (const raiz of raizes) {
    const idx = baixarCatalogo(api, raiz);
    catalogos.push(idx);
    for (const cod of codigos) {
      const nova = linhaProduto(cod, idx);
      if (posicao.has(cod)) linhas[posicao.get(cod)] = nova;
      else { posicao.set(cod, linhas.length); linhas.push(nova); }
    }
  }
  linhas.sort((a, b) => Number(a[0]) - Number(b[0]));
  escreverLinhas(aba, linhas, CAB_PRODUTOS.length);
  return catalogos;
}

function linhaProduto(codigo, idx) {
  const prod = idx.prodPorCodigo.get(Number(codigo)) || null;
  const ops = [];
  for (const v of idx.vincPorCodigo.get(Number(codigo)) || []) {
    if (v.codigoOperadorEstrangeiro) {
      const o = idx.opPorChave.get(`${v.codigoPais}|${v.codigoOperadorEstrangeiro}`) || {};
      ops.push([v.codigoOperadorEstrangeiro, o.nome || '', v.codigoPais || '', o.tin || '']);
    } else if (v.cpfCnpjFabricante) {
      ops.push([v.cpfCnpjFabricante, 'fabricante nacional', 'BR', '']);
    } else {
      ops.push(['', 'fabricante desconhecido', v.codigoPais || '', '']);
    }
  }
  const p = prod || {};
  const lido = prod ? lerProdutoCatalogo(p.denominacao, p.descricao) : {};
  const linha = new Array(CAB_PRODUTOS.length).fill('');
  linha[P['Código produto']] = Number(codigo);
  linha[P['Versão atual catálogo']] = p.versao || '';
  linha[P['Situação produto']] = p.situacao || '';
  linha[P['NCM']] = p.ncm || '';
  linha[P['Descrição']] = p.denominacao || '';
  linha[P['Descrição detalhada']] = p.descricao || '';
  linha[P['Código interno']] = (p.codigosInterno || []).join('; ');
  linha[P['Cód. operador estrangeiro']] = ops.map(o => o[0]).join('\n');
  linha[P['Nome operador estrangeiro']] = ops.map(o => o[1]).join('\n');
  linha[P['País operador']] = ops.map(o => o[2]).join('\n');
  linha[P['TIN operador']] = ops.map(o => o[3]).join('\n');
  linha[P['CNPJ raiz']] = idx.raiz;
  linha[P['Medida']] = lido.medida || '';
  linha[P['Índice de carga']] = lido.carga || '';
  linha[P['Lonas']] = lido.lonas || '';
  linha[P['Marca']] = lido.marca || '';
  linha[P['Modelo']] = lido.modelo || '';
  linha[P['Câmara']] = prod ? lerCamara(`${p.denominacao} ${p.descricao}`) : '';
  return linha;
}

// Menu: baixa o catálogo inteiro e regrava a aba Catálogo
function atualizarCatalogoSemTrava() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = lerConfig(ss);
  const api = new PortalUnico(cfg);
  let resultado = 'OK', total = 0;
  try {
    api.autenticar();
    total = gravarCatalogo(ss, cfg, baixarCatalogo(api, cfg.cnpj.slice(0, 8)));
    try { aplicarVisual(); } catch (e) { console.warn(e.message); }
  } catch (e) {
    resultado = (e.limite ? 'Limite de acessos do portal: ' : 'Erro: ') + e.message;
    console.error(e.stack || e);
  }
  registrarHistorico(ss, [new Date(), 'Catálogo de produtos', '', '', total, resultado]);
  try { mostrarDatas(ss); } catch (e) { console.warn(e.message); }
  avisar(`${resultado}\n\nProdutos no catálogo (NCMs ${cfg.ncmsCatalogo.join(', ')}): ${total}`);
}

function gravarCatalogo(ss, cfg, idx) {
  const aba = obterAba(ss, ABA_CATALOGO, CAB_PRODUTOS);
  const linhas = idx.prods
    .filter(p => cfg.ncmsCatalogo.indexOf(String(p.ncm)) >= 0)
    .map(p => linhaProduto(p.codigo, idx))
    .sort((a, b) => Number(a[0]) - Number(b[0]));
  escreverLinhas(aba, linhas, CAB_PRODUTOS.length);
  gravarOperadores(ss, idx);
  recalcularLpcos(ss, ABA_CATALOGO);
  gravarControle(ss, 'Última atualização do catálogo', new Date());
  return linhas.length;
}

// Versões, LPCOs, quantidade, vigência e observação vêm da aba LPCOs (só deferidas)
function recalcularLpcos(ss, nomeAba) {
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return;
  const produtos = lerLinhas(aba);
  if (!produtos.length) return;
  const porProduto = new Map();
  for (const l of lerLinhas(obterAba(ss, ABA_LPCOS, CAB_LPCOS))) {
    const cod = Number(l[7]);
    if (!cod) continue;
    if (!porProduto.has(cod)) porProduto.set(cod, { versoes: new Set(), lpcos: new Set(), fim: null });
    const g = porProduto.get(cod);
    if (l[8] !== '') g.versoes.add(String(l[8]));
    g.lpcos.add(String(l[0]));
    const fim = paraData(l[4]);
    if (ehData(fim) && (!g.fim || fim > g.fim)) g.fim = fim;
  }
  for (const linha of produtos) {
    const g = porProduto.get(Number(linha[P['Código produto']])) ||
      { versoes: new Set(), lpcos: new Set(), fim: null };
    linha[P['Versão na LPCO']] = [...g.versoes].sort().join(', ');
    linha[P['Qtd LPCOs']] = g.lpcos.size;
    linha[P['LPCOs']] = [...g.lpcos].sort().join(', ');
    linha[P['Fim vigência LPCO mais recente']] = g.fim || '';
    const obs = [];
    const atual = String(linha[P['Versão atual catálogo']] || '');
    if (!atual && !linha[P['Descrição']]) obs.push('produto nao encontrado no catalogo');
    else if (atual && g.versoes.size && !g.versoes.has(atual)) obs.push('versao atual do catalogo diferente da usada na LPCO');
    if (atual && !linha[P['Cód. operador estrangeiro']] && !linha[P['Nome operador estrangeiro']]) {
      obs.push('sem operador estrangeiro vinculado');
    }
    linha[P['Observação']] = obs.join('; ');
  }
  escreverLinhas(aba, produtos, CAB_PRODUTOS.length);
}


// IDENTIDADE VISUAL
const COR = {
  tinta: '#14213D', tintaMedia: '#33415C', destaque: '#EB6834', destaqueClaro: '#FDEEE6',
  fundo: '#F3F5F8', borda: '#E1E5EB', cinza: '#5F6673', branco: '#FFFFFF',
  // cores só para status
  verde: '#2E7D5B', verdeClaro: '#E3F1EA', ambar: '#D98E04', ambarClaro: '#FBF0DA',
  vermelho: '#B8443A', vermelhoClaro: '#F6E0DD',
};
const FONTE_TITULO = 'Montserrat';

function aplicarVisual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estilo = [[ABA_CATALOGO, COR.tinta], [ABA_PRODUTOS, COR.tinta], [ABA_LPCOS, COR.tintaMedia]];
  for (const [nome, corAba] of estilo) {
    const aba = ss.getSheetByName(nome);
    if (!aba) continue;
    aba.setTabColor(corAba);
    estilizarTabela(aba, 1);
  }
  const ctrl = ss.getSheetByName(ABA_CONTROLE);
  if (ctrl) {
    ctrl.setTabColor(COR.cinza);
    ctrl.getRange('A1').setFontFamily(FONTE_TITULO).setFontSize(16).setFontWeight('bold').setFontColor(COR.tinta);
    const { linha } = linhaControle(ss, 'Histórico de execuções');
    if (linha) {
      ctrl.getRange(linha, 1).setFontFamily(FONTE_TITULO).setFontWeight('bold').setFontColor(COR.tinta);
      ctrl.getRange(linha + 1, 1, 1, 6).setBackground(COR.tinta).setFontColor(COR.branco).setFontWeight('bold');
      ctrl.getRange(3, 1, linha - 4, 1).setFontWeight('bold').setFontColor(COR.tinta);
      ctrl.getRange(3, 2, linha - 4, 1).setBackground(COR.destaqueClaro).setFontWeight('bold');
      ctrl.getRange(3, 3, linha - 4, 1).setFontColor(COR.cinza).setWrap(true);
    }
  }
}

function estilizarTabela(aba, linhaCab) {
  const ncol = Math.max(aba.getLastColumn(), 1);
  const nlin = Math.max(aba.getLastRow() - linhaCab + 1, 2);
  aba.getBandings().forEach(b => b.remove());
  aba.getRange(linhaCab, 1, nlin, ncol).applyRowBanding()
    .setHeaderRowColor(COR.tinta).setFirstRowColor(COR.branco).setSecondRowColor(COR.fundo);
  aba.getRange(linhaCab, 1, 1, ncol).setFontColor(COR.branco).setFontWeight('bold')
    .setFontFamily(FONTE_TITULO).setWrap(true).setVerticalAlignment('middle');
  aba.setRowHeight(linhaCab, 40);
  aba.setFrozenRows(linhaCab);
}


// ABA BUSCAR: lista de vários itens + caixa BUSCAR
// Busca exata primeiro (todos os campos informados iguais). Sem resultado exato, sugere "possíveis
// correspondências" só quando bate o código (interno ou do catálogo) ou quando medida, marca e modelo
// batem e a diferença está apenas em índice de carga, lonas ou detalhe da medida (LT, ZR, C).
// Marca ou modelo diferente nunca é o mesmo pneu. A pontuação (código 5, medida 5/4, marca 3,
// modelo 2, índice 2, lonas 1) só ordena os candidatos.

const ABAS_BUSCA = { busca: 'Buscar', resultado: 'Resultado da busca', novos: 'Novos cadastros' };
function ehAbaDeBusca(nome) { return nome === ABAS_BUSCA.busca; }
function ehAbaDeNovos(nome) { return nome === ABAS_BUSCA.novos; }

const BUSCA = {
  celBuscar: 'B6', celLimpar: 'D6',
  linhaCab: 8, primeira: 9, linhas: 1000,
  entradas: ['Medida', 'Marca', 'Modelo / Desenho', 'Índice / Lonas', 'Código interno', 'Código no catálogo'],
  saidas: ['Encontrados', 'Códigos no catálogo', 'LPCOs deferidas', 'Fim vigência', 'Situação', 'Decisão'],
};
const DECISOES = ['Equivalente - não cadastrar', 'Cadastrar novo item', 'Revisar'];
const CAB_RESULTADO = [
  'Item buscado', 'Linha', 'Código produto', 'Versão atual catálogo', 'Situação produto', 'NCM', 'Medida',
  'Índice de carga', 'Lonas', 'Câmara', 'Marca', 'Modelo', 'Código interno', 'Qtd LPCOs', 'LPCOs',
  'Fim vigência LPCO mais recente', 'Nome operador estrangeiro', 'Descrição', 'Descrição detalhada',
  'Registro INMETRO', 'Certificado INMETRO', 'Família INMETRO', 'Observação', 'Correspondência',
];
const LARGURAS_RESULTADO = [160, 50, 75, 70, 80, 80, 100, 85, 55, 60, 95, 170, 140, 55, 220, 95, 190,
  320, 420, 105, 170, 105, 200, 230];

// Valor de uma coluna do resultado; os dados INMETRO saem da descrição detalhada do cadastro
function valorResultado(coluna, p) {
  if (coluna in P) return p.l[P[coluna]];
  if (!p.inmetro) {
    p.inmetro = lerInmetro(p.l[P['Descrição detalhada']] || p.l[P['Descrição']]);
    if (!p.inmetro.some(Boolean)) p.inmetro = lerInmetro(p.l[P['Descrição']]);
  }
  const [registro, certificado, familia] = p.inmetro;
  return { 'Registro INMETRO': registro, 'Certificado INMETRO': certificado, 'Família INMETRO': familia }[coluna] || '';
}

// Cabeçalho das abas Buscar e Novos (título, datas e faixa colorida) nas linhas 1 a 4
function cabecalhoAba(aba, titulo, subtitulo, nCols) {
  aba.getRange(1, 1, 3, nCols).setBackground(COR.branco);
  [34, 30, 26].forEach((h, i) => aba.setRowHeight(i + 1, h));
  aba.getRange('A1').setValue(titulo)
    .setFontFamily(FONTE_TITULO).setFontSize(18).setFontWeight('bold').setFontColor(COR.tinta);
  aba.getRange('A2').setValue(subtitulo).setFontColor(COR.cinza).setFontSize(10);
  aba.getRange('C3').setValue('Última extração de LPCOs:').setFontColor(COR.cinza).setFontSize(10);
  aba.getRange('G3').setValue('Catálogo atualizado em:').setFontColor(COR.cinza).setFontSize(10)
    .setHorizontalAlignment('right');
  aba.getRange(4, 1, 1, nCols).setBackground(COR.destaque);
  aba.setRowHeight(4, 6);
  mostrarDatas(aba.getParent());
}

// Copia as datas da aba Controle para o topo das abas Buscar e Novos cadastros.
// Chamada ao criar as abas e ao fim de cada atualização (não depende de fórmula).
const CEL_DATAS = { 'E3': 'Última extração', 'H3': 'Última atualização do catálogo' };
function mostrarDatas(ss) {
  for (const aba of ss.getSheets()) {
    if (!ehAbaDeBusca(aba.getName()) && !ehAbaDeNovos(aba.getName())) continue;
    for (const [cel, rotulo] of Object.entries(CEL_DATAS)) {
      const valor = paraData(lerControle(ss, rotulo));
      aba.getRange(cel).setValue(ehData(valor) ? valor : 'sem registro').setNumberFormat('dd/mm/yyyy hh:mm')
        .setFontWeight('bold').setFontColor(COR.tinta).setHorizontalAlignment('left');
    }
  }
}

function botao(aba, rotulo, celRotulo, celCaixa, fundo) {
  aba.getRange(celRotulo).setValue(rotulo).setBackground(fundo).setFontColor(COR.tinta).setFontWeight('bold')
    .setFontFamily(FONTE_TITULO).setFontSize(11).setHorizontalAlignment('right').setVerticalAlignment('middle');
  aba.getRange(celCaixa).insertCheckboxes().setValue(false).setBackground(fundo)
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setFontSize(14);
}

function prepararAba(ss, nome, posicao) {
  let aba = ss.getSheetByName(nome);
  if (aba) { aba.clear(); aba.getImages().forEach(i => i.remove()); aba.getRange('A:Z').removeCheckboxes(); }
  else aba = ss.insertSheet(nome, posicao);
  aba.setHiddenGridlines(true);
  return aba;
}

function criarAbaBuscar() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = prepararAba(ss, g.busca, ss.getSheets().length);
  aba.setTabColor(COR.destaque);
  garantirLinhas(aba, BUSCA.primeira + BUSCA.linhas - 1);
  const nEnt = BUSCA.entradas.length, nTot = nEnt + BUSCA.saidas.length;
  cabecalhoAba(aba, 'Consulta de Produtos e LPCOs',
    'Catálogo de Produtos e LPCOs deferidas, Portal Único Siscomex', nTot);

  aba.getRange('A5').setValue('Digite ou cole um item por linha (pode colar as colunas Medida, Marca, Desenho e ' +
    'Índice/Lonas de outra planilha). Preencha só o que souber e marque BUSCAR. ' +
    `Itens com decisão "Cadastrar novo item" vão para a aba ${g.novos}.`)
    .setFontColor(COR.cinza).setFontStyle('italic');
  aba.setRowHeight(6, 36);
  botao(aba, 'BUSCAR', 'A6', BUSCA.celBuscar, COR.destaque);
  botao(aba, 'LIMPAR', 'C6', BUSCA.celLimpar, COR.fundo);

  aba.getRange(BUSCA.linhaCab, 1, 1, nTot).setValues([BUSCA.entradas.concat(BUSCA.saidas)]).setFontWeight('bold')
    .setFontFamily(FONTE_TITULO).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  aba.getRange(BUSCA.linhaCab, 1, 1, nEnt).setBackground(COR.tinta).setFontColor(COR.branco);
  aba.getRange(BUSCA.linhaCab, nEnt + 1, 1, BUSCA.saidas.length).setBackground(COR.destaque).setFontColor(COR.tinta);
  aba.setRowHeight(BUSCA.linhaCab, 38);
  aba.getRange(BUSCA.primeira, 1, BUSCA.linhas, nTot)
    .setBorder(true, true, true, true, true, true, COR.borda, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('top').setNumberFormat('@');
  aba.getRange(BUSCA.primeira, 1, BUSCA.linhas, nEnt).setBackground(COR.branco);
  aba.getRange(BUSCA.primeira, nEnt + 1, BUSCA.linhas, BUSCA.saidas.length).setBackground(COR.fundo)
    .setFontColor(COR.tinta).setWrap(true);
  aba.getRange(BUSCA.primeira, nEnt + 4, BUSCA.linhas, 1).setNumberFormat('dd/mm/yyyy');
  const colDecisao = aba.getRange(BUSCA.primeira, nTot, BUSCA.linhas, 1);
  colDecisao.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(DECISOES, true)
    .setAllowInvalid(false).build()).setBackground(COR.branco).setFontWeight('bold');
  aba.setFrozenRows(BUSCA.linhaCab);
  [130, 110, 150, 130, 130, 95, 85, 150, 230, 90, 230, 170].forEach((w, i) => aba.setColumnWidth(i + 1, w));

  obterAbaResultado(ss, g);
  if (!ss.getSheetByName(g.novos)) criarAbaNovos();
  aplicarVisual();
  ss.setActiveSheet(aba);
}

// Abas novas nascem com 1000 linhas: acrescenta as que faltarem antes de escrever
function garantirLinhas(aba, total) {
  const faltam = total - aba.getMaxRows();
  if (faltam > 0) aba.insertRowsAfter(aba.getMaxRows(), faltam);
}

function obterAbaResultado(ss, g) {
  let aba = ss.getSheetByName(g.resultado);
  if (!aba) {
    aba = ss.insertSheet(g.resultado, ss.getSheets().length);
    aba.setTabColor(COR.destaque);
  }
  // Abas criadas em versões anteriores ganham as colunas novas (descrição detalhada e INMETRO)
  const atual = aba.getRange(1, 1, 1, CAB_RESULTADO.length).getValues()[0];
  if (atual.some((v, i) => v !== CAB_RESULTADO[i])) {
    aba.getRange(1, 1, 1, CAB_RESULTADO.length).setValues([CAB_RESULTADO]);
    LARGURAS_RESULTADO.forEach((w, i) => aba.setColumnWidth(i + 1, w));
  }
  return aba;
}

// Caixas de seleção que funcionam como botões (gatilho simples: vale para qualquer editor)
function onEdit(e) {
  if (!e || !e.range) return;
  const nome = e.range.getSheet().getName();
  const linha = e.range.getRow(), nLinhas = e.range.getNumRows();
  const col = e.range.getColumn(), nCols = e.range.getNumColumns();
  // Decisão escolhida na busca: "Cadastrar novo item" vai direto para a aba Novos
  const colDecisao = BUSCA.entradas.length + BUSCA.saidas.length;
  if (ehAbaDeBusca(nome) && linha >= BUSCA.primeira && col <= colDecisao && col + nCols - 1 >= colDecisao) {
    sincronizarNovos(ABAS_BUSCA);
    return;
  }
  // Dado de cadastro digitado ou colado na aba Novos: refaz Denominação e Descrição dessas linhas
  if (ehAbaDeNovos(nome) && linha >= NOVOS.primeira && col <= N_ENTRADAS) {
    atualizarLinhasNovos(e.range.getSheet(), linha, nLinhas);
    return;
  }
  if (String(e.value).toUpperCase() !== 'TRUE') return;
  const cel = e.range.getA1Notation();
  let acao = null;
  if (ehAbaDeBusca(nome)) acao = { [BUSCA.celBuscar]: buscarLista, [BUSCA.celLimpar]: limparLista }[cel];
  else if (ehAbaDeNovos(nome)) {
    acao = { [NOVOS.celGerar]: gerarNovosDaBusca, [NOVOS.celAtualizar]: atualizarDescricoes,
             [NOVOS.celLimpar]: limparNovos }[cel];
  }
  if (!acao) return;
  e.range.setValue(false);
  acao();
}

function limparLista() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(g.busca);
  const nEnt = BUSCA.entradas.length, nSai = BUSCA.saidas.length;
  aba.getRange(BUSCA.primeira, 1, BUSCA.linhas, nEnt + nSai).clearContent();
  aba.getRange(BUSCA.primeira, nEnt + 1, BUSCA.linhas, nSai - 1).setBackground(COR.fundo);
  const res = obterAbaResultado(ss, g);
  if (res.getLastRow() > 1) res.getRange(2, 1, res.getLastRow() - 1, CAB_RESULTADO.length).clearContent();
}

function fonteDeProdutos(ss) {
  const abaCat = ss.getSheetByName(ABA_CATALOGO);
  return abaCat && abaCat.getLastRow() > 1 ? abaCat : obterAba(ss, ABA_PRODUTOS, CAB_PRODUTOS);
}

function buscarLista() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(g.busca);
  if (!aba) throw new Error(`Crie a aba ${g.busca} pelo menu Robô LPCO.`);
  const nEnt = BUSCA.entradas.length, nSai = BUSCA.saidas.length;
  const entradas = aba.getRange(BUSCA.primeira, 1, BUSCA.linhas, nEnt).getDisplayValues();
  const fonte = fonteDeProdutos(ss);
  const produtos = lerLinhas(fonte).map(prepararProduto);

  const resumo = [], fundos = [], detalhe = [];
  entradas.forEach((ent, i) => {
    if (ent.every(v => String(v).trim() === '')) {
      resumo.push(new Array(nSai).fill(''));
      fundos.push(new Array(nSai - 1).fill(COR.fundo).concat([COR.branco]));
      return;
    }
    const filtro = montarFiltro(ent);
    let achados = produtos.filter(p => combina(p, filtro)).map(p => ({ p, tipo: 'Exata' }));
    let situacao, cor, decisao, outroIndice = [];
    if (achados.length) {
      const comLpco = achados.filter(a => Number(a.p.l[P['Qtd LPCOs']]) > 0).length;
      if (achados.length === 1) {
        situacao = comLpco ? 'Encontrado, com LPCO' : 'Encontrado, sem LPCO deferida';
        cor = comLpco ? COR.verdeClaro : COR.ambarClaro;
      } else {
        situacao = `${achados.length} cadastros: ver aba "${g.resultado}"`;
        cor = COR.ambarClaro;
      }
      decisao = 'Equivalente - não cadastrar';
    } else {
      // Nada exato: os 3 candidatos com mais pontos
      achados = produtos.map(p => Object.assign({ p }, pontuar(p, filtro)))
        .filter(a => a.possivel)
        .sort((a, b) => b.pontos - a.pontos || b.fatores.length - a.fatores.length)
        .slice(0, 3)
        .map(a => ({ p: a.p, tipo: `Possível (${a.pontos} pts: ${a.fatores.join(', ')}; diferente: ${a.faltam.join(', ') || '-'})` }));
      situacao = achados.length ? `Possível correspondência: conferir "${g.resultado}"` : 'Não encontrado';
      cor = achados.length ? COR.ambarClaro : COR.vermelhoClaro;
      decisao = achados.length ? 'Revisar' : 'Cadastrar novo item';
      if (!achados.length) {
        outroIndice = outroIndiceOuLonas(produtos, filtro);
        if (outroIndice.length) {
          situacao = 'Não encontrado com este índice/lonas. Já cadastrado com outro: ' +
            outroIndice.slice(0, 5).map(p => `${p.l[P['Código produto']]} (${indiceLonas(p)})`).join(', ') +
            (outroIndice.length > 5 ? ` ... (+${outroIndice.length - 5})` : '');
        }
      }
    }
    const fins = achados.map(a => a.p.l[P['Fim vigência LPCO mais recente']]).filter(ehData);
    const fim = fins.length ? new Date(Math.max(...fins.map(d => d.getTime()))) : '';
    const lpcos = [...new Set(achados.flatMap(a => String(a.p.l[P['LPCOs']] || '').split(/,\s*/).filter(Boolean)))];
    resumo.push([
      achados.length,
      achados.slice(0, 15).map(a => a.p.l[P['Código produto']]).join(', ') + (achados.length > 15 ? ' ...' : ''),
      lpcos.slice(0, 6).join(', ') + (lpcos.length > 6 ? ` ... (+${lpcos.length - 6})` : ''),
      fim, situacao, decisao,
    ]);
    fundos.push(new Array(nSai - 1).fill(cor).concat([COR.branco]));
    const descricaoBusca = ent.map(v => String(v).trim()).filter(Boolean).join(' | ');
    for (const a of achados) {
      detalhe.push([descricaoBusca, BUSCA.primeira + i]
        .concat(CAB_RESULTADO.slice(2, -1).map(c => valorResultado(c, a.p))).concat([a.tipo]));
    }
    if (!achados.length) {
      detalhe.push([descricaoBusca, BUSCA.primeira + i, 'não encontrado']
        .concat(new Array(CAB_RESULTADO.length - 4).fill('')).concat(['Nenhuma']));
    }
    // Para análise: mesma medida, marca e modelo, mas outro índice de carga ou outras lonas
    const buscado = [ent[3], ent[0]].join(' ');
    for (const p of outroIndice) {
      detalhe.push([descricaoBusca, BUSCA.primeira + i]
        .concat(CAB_RESULTADO.slice(2, -1).map(c => valorResultado(c, p)))
        .concat([`Outro produto: mesma medida/marca/modelo, índice/lonas diferente ` +
                 `(cadastro ${indiceLonas(p)} x buscado ${indiceBuscado(filtro, buscado)})`]));
    }
  });

  aba.getRange(BUSCA.primeira, nEnt + 1, BUSCA.linhas, nSai).setValues(resumo).setBackgrounds(fundos);
  const res = obterAbaResultado(ss, g);
  if (res.getLastRow() > 1) res.getRange(2, 1, res.getLastRow() - 1, CAB_RESULTADO.length).clearContent();
  if (detalhe.length) {
    garantirLinhas(res, detalhe.length + 1);
    res.getRange(2, 1, detalhe.length, CAB_RESULTADO.length).setValues(detalhe);
    res.getRange(2, CAB_RESULTADO.indexOf('Fim vigência LPCO mais recente') + 1, detalhe.length, 1).setNumberFormat('dd/mm/yyyy');
    estilizarTabela(res, 1);
  }
  const buscados = resumo.filter(r => r[0] !== '').length;
  // A decisão gravada pela busca não dispara o onEdit: envia aqui os "Cadastrar novo item"
  const enviados = sincronizarNovos(g);
  ss.toast(`${buscados} item(ns) buscado(s) na aba ${fonte.getName()}. ` +
           `${enviados} item(ns) sem cadastro enviado(s) para a aba ${g.novos}.`, 'Busca concluída', 8);
}

// Textos normalizados de cada produto, calculados uma vez por busca
function prepararProduto(l) {
  return {
    l,
    medidaTexto: soAlfanum(l[P['Medida']]),
    medida: String(l[P['Medida']] || '').replace(/\D/g, ''),
    carga: soAlfanum(l[P['Índice de carga']]),
    lonas: soAlfanum(l[P['Lonas']]),
    camara: String(l[P['Câmara']] || ''),
    marcaCampo: soAlfanum(l[P['Marca']]),
    marca: soAlfanum(`${l[P['Marca']]} ${l[P['Descrição']]}`),
    modelo: soAlfanum(`${l[P['Modelo']]} ${l[P['Descrição detalhada']]}`),
    modeloCampo: soAlfanum(l[P['Modelo']]),
    interno: soAlfanum(l[P['Código interno']]),
    codigo: String(l[P['Código produto']]).trim(),
  };
}

function montarFiltro([medida, marca, modelo, indice, interno, codigo]) {
  const f = {};
  // Medida e índice podem vir juntos ("295/80R22.5 152/149M 18PR") ou em colunas separadas
  const texto = `${medida} ${indice}`.trim();
  if (texto) {
    const [lida, carga, lonas] = lerMedidaCargaLonas(texto);
    if (String(medida).trim()) {
      f.medida = (lida || String(medida)).replace(/\D/g, '');
      f.medidaTexto = soAlfanum(lida || medida);
      f.medidaParcial = !lida;
    }
    if (carga) f.carga = soAlfanum(carga);
    if (lonas) f.lonas = lonas;
    if (!carga && !lonas && String(indice).trim() && !lida) {
      const [, c, l] = lerMedidaCargaLonas(`000/00R00 ${indice}`);
      if (c) f.carga = soAlfanum(c);
      if (l) f.lonas = l;
    }
    f.camara = lerCamara(indice);
  }
  if (soAlfanum(marca)) f.marca = soAlfanum(marca);
  if (soAlfanum(modelo)) f.modelo = soAlfanum(modelo);
  if (soAlfanum(interno)) f.interno = soAlfanum(interno);
  if (String(codigo).trim()) f.codigo = String(codigo).replace(/\D/g, '');
  return f;
}

function combina(p, f) {
  if (f.medida && (f.medidaParcial ? p.medida.indexOf(f.medida) !== 0 : p.medida !== f.medida)) return false;
  if (f.carga && p.carga !== f.carga) return false;
  if (f.lonas && p.lonas !== f.lonas) return false;
  if (f.camara && p.camara && p.camara !== f.camara) return false;
  if (f.marca && !marcaBate(p, f.marca)) return false;
  if (f.modelo && !modeloBate(p, f.modelo)) return false;
  if (f.interno && p.interno.indexOf(f.interno) < 0) return false;
  if (f.codigo && p.codigo !== f.codigo) return false;
  return true;
}

// Cadastros com a mesma medida, marca e modelo da busca, mas com outro índice de carga ou lonas.
// Só quando a busca informou índice/lonas e marca ou modelo.
function outroIndiceOuLonas(produtos, f) {
  if (!(f.carga || f.lonas) || !(f.marca || f.modelo) || !f.medida) return [];
  return produtos.filter(p =>
    p.medida === f.medida && (!f.marca || marcaBate(p, f.marca)) && (!f.modelo || modeloBate(p, f.modelo)) &&
    ((f.carga && p.carga && p.carga !== f.carga) || (f.lonas && p.lonas && p.lonas !== f.lonas)));
}
function indiceLonas(p) {
  return [p.l[P['Índice de carga']], p.l[P['Lonas']]].filter(Boolean).join(' ') || 'sem índice';
}
function indiceBuscado(f, texto) {
  const [, carga, lonas] = lerMedidaCargaLonas(`000/00R00 ${texto}`);
  return [carga, lonas].filter(Boolean).join(' ') || [f.carga, f.lonas].filter(Boolean).join(' ');
}

// Pontuação dos candidatos (índice e lonas também compõem o pneu)
function pontuar(p, f) {
  let pontos = 0, base = false;
  const fatores = [], faltam = [];
  const conta = (ok, peso, nome, ehBase) => {
    if (ok) { pontos += peso; fatores.push(nome); if (ehBase) base = true; } else faltam.push(nome);
  };
  if (f.interno) conta(p.interno === f.interno, 5, 'código interno', true);
  if (f.codigo) conta(p.codigo === f.codigo, 5, 'código catálogo', true);
  if (f.medida) {
    if (p.medidaTexto === f.medidaTexto) conta(true, 5, 'medida exata', true);
    else if (p.medida === f.medida || (f.medidaParcial && p.medida.indexOf(f.medida) === 0)) conta(true, 4, 'medida base', true);
    else faltam.push('medida');
  }
  if (f.marca) conta(marcaBate(p, f.marca), 3, 'marca');
  if (f.modelo) conta(modeloBate(p, f.modelo), 2, 'modelo');
  if (f.carga) conta(p.carga === f.carga, 2, 'índice');
  if (f.lonas) conta(p.lonas === f.lonas, 1, 'lonas');
  const tem = nome => fatores.indexOf(nome) >= 0;
  const porCodigo = tem('código interno') || tem('código catálogo');
  // Índice de carga e lonas informados diferenciam o pneu: precisam bater também
  const porDescricao = (tem('medida exata') || tem('medida base')) && Boolean(f.marca || f.modelo) &&
    (!f.marca || tem('marca')) && (!f.modelo || tem('modelo')) &&
    (!f.carga || tem('índice')) && (!f.lonas || tem('lonas'));
  return { pontos, base, fatores, faltam, possivel: porCodigo || porDescricao };
}

// "MARCAUM" no cadastro bate com "MARCAUM TYRE" digitado, e "MD" digitado bate com "MD200"
function marcaBate(p, marca) {
  return p.marca.indexOf(marca) >= 0 || (p.marcaCampo.length >= 2 && marca.indexOf(p.marcaCampo) === 0);
}
function modeloBate(p, modelo) {
  return p.modelo.indexOf(modelo) >= 0 || (p.modeloCampo.length >= 3 && modelo.indexOf(p.modeloCampo) >= 0);
}

function lerCamara(texto) {
  const t = semAcento(texto);
  if (/SEM CAM[AE]RAS?|TUBELESS|\bTL\b/.test(t)) return 'TL';
  if (/COM CAM[AE]RAS?|TUBE ?TYPE|\bTT\b/.test(t)) return 'TT';
  return '';
}


// ABA NOVOS CADASTROS: denominação e descrição no padrão do Catálogo de Produtos
// e exportação em Excel para o cadastro em lote
const ABA_OPERADORES = 'Operadores';
const NOVOS = { celGerar: 'B6', celAtualizar: 'D6', celLimpar: 'F6', linhaCab: 8, primeira: 9, linhas: 1000 };
const CAB_NOVOS = [
  'Linha da busca', 'Medida', 'Índice / Lonas', 'Marca', 'Modelo', 'NCM', 'Câmara (TL/TT)',
  'Certificado INMETRO', 'Registro INMETRO', 'Família INMETRO',
  'Fabricante Conhecido', 'Fabricante PU', 'Fabricante Código', 'Fabricante Nome', 'Fabricante País',
  'Fabricante Logradouro', 'Fabricante Cidade', 'Denominação (gerada)', 'Descrição (gerada)', 'Observações',
  'Chave da busca',
];
const N = {};
CAB_NOVOS.forEach((n, i) => { N[n] = i; });
const N_ENTRADAS = N['Fabricante Cidade'] + 1;  // colunas A..Q são editáveis
// Colunas do arquivo de cadastro em lote
const CAB_EXPORTACAO = ['Código', 'Denominação', 'Descrição', 'NCM', 'Fabricante Conhecido', 'Fabricante PU',
  'Fabricante Código', 'Fabricante Nome', 'Fabricante País', 'Fabricante Logradouro', 'Fabricante Cidade'];
const CAB_OPERADORES = ['Código', 'País', 'Nome', 'Logradouro', 'Cidade', 'TIN', 'Situação'];

const MAPA_NCM = {
  '40111000': ['PNEU DE PASSAGEIRO', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS USADOS EM AUTOMOVEIS DE PASSAGEIROS'],
  '40112010': ['PNEU DE ONIBUS OU CAMINHAO', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS UTILIZADOS EM ONIBUS OU CAMINHOES'],
  '40112090': ['PNEU DE CAMINHAO', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS UTILIZADOS EM CAMINHOES E ONIBUS'],
  '40117090': ['PNEU AGRICOLA', 'EX01 - PARA MAQUINAS E TRATORES AGRICOLAS, SENDO: PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS USADOS EM VEICULOS E MAQUINAS AGRICOLAS'],
  '40118090': ['PNEU OTR', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS USADOS EM VEICULOS E MAQUINAS PARA CONSTRUCAO, MINERACAO OU MANUTENCAO INDUSTRIAL'],
  '40119090': ['PNEU LIGHT TRUCK LT/C', 'OUTROS PNEUMATICOS NOVOS, DE BORRACHA, DO TIPO UTILIZADO EM VEICULOS UTILITARIOS, VANS E LIGHT TRUCK'],
  '40114000': ['PNEU DE MOTOCICLETA', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS USADOS EM MOTOCICLETAS'],
  '40115000': ['PNEU DE BICICLETA', 'PNEUMATICOS NOVOS, DE BORRACHA, DOS TIPOS USADOS EM BICICLETAS'],
};
const PAISES = {
  CN: 'CHINA, REPUBLICA POPULAR', IN: 'INDIA', VN: 'VIETNA', HK: 'HONG KONG', KR: 'COREIA (DO SUL), REPUBLICA DA',
  US: 'ESTADOS UNIDOS', KH: 'CAMBOJA', ID: 'INDONESIA', AE: 'EMIRADOS ARABES UNIDOS', MY: 'MALASIA',
  TH: 'TAILANDIA', SG: 'CINGAPURA', HU: 'HUNGRIA', PK: 'PAQUISTAO', PA: 'PANAMA', IT: 'ITALIA', ES: 'ESPANHA',
  MX: 'MEXICO', TW: 'TAIWAN (FORMOSA)', JP: 'JAPAO', TR: 'TURQUIA', LK: 'SRI LANKA', RU: 'RUSSIA, FEDERACAO DA',
  BR: 'BRASIL', DE: 'ALEMANHA', FR: 'FRANCA', RS: 'SERVIA', LU: 'LUXEMBURGO',
};

function criarAbaNovos() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = prepararAba(ss, g.novos, ss.getSheets().length);
  aba.setTabColor(COR.tintaMedia);
  garantirLinhas(aba, NOVOS.primeira + NOVOS.linhas - 1);
  const nTot = CAB_NOVOS.length;
  cabecalhoAba(aba, 'Novos cadastros no Catálogo',
    'Denominação e descrição no padrão do Catálogo de Produtos', nTot);
  aba.getRange('A5').setValue(`1) Na aba ${g.busca}, marque "Cadastrar novo item" na Decisão.  ` +
    '2) Marque GERAR para trazer ' +
    'esses itens com NCM, câmara e fabricante sugeridos pelo catálogo.  3) Revise e complete (INMETRO, NCM, fabricante) ' +
    'e marque ATUALIZAR.  4) Menu Robô LPCO > Baixar Excel de novos cadastros.')
    .setFontColor(COR.cinza).setFontStyle('italic');
  aba.setRowHeight(6, 36);
  // Rótulos curtos: as colunas A e C são estreitas e o Sheets corta o texto que não cabe
  botao(aba, 'GERAR', 'A6', NOVOS.celGerar, COR.destaque);
  botao(aba, 'ATUALIZAR', 'C6', NOVOS.celAtualizar, COR.fundo);
  botao(aba, 'LIMPAR', 'E6', NOVOS.celLimpar, COR.fundo);

  aba.getRange(NOVOS.linhaCab, 1, 1, nTot).setValues([CAB_NOVOS]).setFontWeight('bold').setFontFamily(FONTE_TITULO)
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  aba.getRange(NOVOS.linhaCab, 1, 1, N_ENTRADAS).setBackground(COR.tinta).setFontColor(COR.branco);
  aba.getRange(NOVOS.linhaCab, N_ENTRADAS + 1, 1, N['Observações'] - N_ENTRADAS + 1).setBackground(COR.destaque).setFontColor(COR.tinta);
  aba.setRowHeight(NOVOS.linhaCab, 38);
  aba.getRange(NOVOS.primeira, 1, NOVOS.linhas, nTot)
    .setBorder(true, true, true, true, true, true, COR.borda, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('top').setNumberFormat('@');
  aba.getRange(NOVOS.primeira, N_ENTRADAS + 1, NOVOS.linhas, nTot - N_ENTRADAS).setBackground(COR.fundo).setWrap(true);
  aba.getRange(NOVOS.primeira, N['Câmara (TL/TT)'] + 1, NOVOS.linhas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['TL', 'TT'], true).setAllowInvalid(true).build());
  aba.setFrozenRows(NOVOS.linhaCab);
  aba.setFrozenColumns(2);
  [70, 120, 120, 100, 150, 85, 70, 130, 110, 100, 90, 80, 95, 200, 160, 200, 110, 330, 520, 260]
    .forEach((w, i) => aba.setColumnWidth(i + 1, w));
  aba.hideColumns(N['Chave da busca'] + 1);  // evita trazer o mesmo item da busca duas vezes
}

// Botão GERAR / menu: traz os itens com decisão "Cadastrar novo item" que ainda não estão na aba Novos
function gerarNovosDaBusca() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const n = sincronizarNovos(g);
  ss.setActiveSheet(ss.getSheetByName(g.novos));
  ss.toast(`${n} item(ns) novo(s) trazido(s) da busca. Preencha certificado, registro, família e o que faltar: ` +
           'a descrição se atualiza sozinha.', 'Novos cadastros', 8);
}

// Acrescenta na aba Novos os itens da busca com decisão "Cadastrar novo item".
// Os que já estão lá (mesma chave) ficam como estão, com tudo o que foi preenchido.
function sincronizarNovos(g) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const busca = ss.getSheetByName(g.busca);
  if (!busca) throw new Error(`Crie a aba ${g.busca} pelo menu Robô LPCO.`);
  if (!ss.getSheetByName(g.novos)) criarAbaNovos();
  const aba = ss.getSheetByName(g.novos);
  const nEnt = BUSCA.entradas.length, nTot = CAB_NOVOS.length;
  const linhasBusca = busca.getRange(BUSCA.primeira, 1, BUSCA.linhas, nEnt + BUSCA.saidas.length).getDisplayValues();
  const atuais = aba.getRange(NOVOS.primeira, 1, NOVOS.linhas, nTot).getDisplayValues();
  const chaves = new Set(atuais.map(l => l[N['Chave da busca']]).filter(Boolean));
  const proxima = atuais.reduce((ult, l, i) => (l.slice(0, N_ENTRADAS).some(v => String(v).trim()) ? i + 1 : ult), 0);
  let sug = null;

  const novas = [];
  linhasBusca.forEach((l, i) => {
    if (l[nEnt + BUSCA.saidas.length - 1] !== 'Cadastrar novo item') return;
    const [medida, marca, modelo, indice] = l;
    const chave = `${BUSCA.primeira + i}|${soAlfanum(medida + indice + marca + modelo)}`;
    if (chaves.has(chave)) return;
    chaves.add(chave);
    sug = sug || new Sugestoes(ss);
    const linha = new Array(nTot).fill('');
    const [lida] = lerMedidaCargaLonas(`${medida} ${indice}`);
    // Índice/lonas digitados junto da medida ("295/80R22,5 12-PR") vão para a coluna própria
    const [, fimMedida] = lerMedida(medida);
    const resto = fimMedida >= 0 ? semAcento(medida).slice(fimMedida).trim() : '';
    linha[N['Linha da busca']] = String(BUSCA.primeira + i);
    linha[N['Medida']] = String(lida || medida).trim().toUpperCase();
    linha[N['Índice / Lonas']] = [resto, String(indice).trim().toUpperCase()].filter(Boolean).join(' ')
      .replace(/(\d)\s?-\s?PR\b/g, '$1PR');
    linha[N['Marca']] = String(marca).trim().toUpperCase();
    linha[N['Modelo']] = String(modelo).trim().toUpperCase();
    linha[N['Chave da busca']] = chave;
    novas.push(completarDescricao(linha, sug.preencher(linha)));
  });
  if (novas.length) {
    aba.getRange(NOVOS.linhaCab, 1, 1, nTot).setValues([CAB_NOVOS]);
    aba.getRange(NOVOS.primeira + proxima, 1, novas.length, nTot).setValues(novas);
  }
  return novas.length;
}

// Refaz Denominação, Descrição e Observações das linhas editadas (chamada pelo onEdit)
function atualizarLinhasNovos(aba, linha, nLinhas) {
  const ini = Math.max(linha, NOVOS.primeira);
  const fim = Math.min(linha + nLinhas - 1, NOVOS.primeira + NOVOS.linhas - 1);
  if (fim < ini) return;
  const saida = aba.getRange(ini, 1, fim - ini + 1, CAB_NOVOS.length).getDisplayValues().map(regerarLinhaNovos);
  aba.getRange(ini, N_ENTRADAS + 1, saida.length, saida[0].length).setValues(saida);
}

function regerarLinhaNovos(l) {
  const largura = N['Observações'] - N_ENTRADAS + 1;
  if (l.slice(1, N_ENTRADAS).every(v => String(v).trim() === '')) return new Array(largura).fill('');
  // Mantém os avisos das sugestões (NCM, câmara, fabricante) para o analista conferir
  const sugestoes = String(l[N['Observações']] || '').split('\n').filter(t => /sugerid|estimad|^Atenção/i.test(t));
  return completarDescricao(l.slice(), sugestoes).slice(N_ENTRADAS, N['Observações'] + 1);
}

function limparNovos() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(g.novos);
  aba.getRange(NOVOS.primeira, 1, NOVOS.linhas, CAB_NOVOS.length).clearContent();
  ss.toast('Lista de novos cadastros apagada.', 'Novos cadastros', 4);
}

// Recalcula Denominação e Descrição a partir do que está nas colunas editáveis
function atualizarDescricoes() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(g.novos);
  const nTot = CAB_NOVOS.length;
  const linhas = aba.getRange(NOVOS.primeira, 1, NOVOS.linhas, nTot).getDisplayValues();
  const n = linhas.filter(l => l.slice(1, N_ENTRADAS).some(v => String(v).trim())).length;
  const saida = linhas.map(regerarLinhaNovos);
  aba.getRange(NOVOS.primeira, N_ENTRADAS + 1, NOVOS.linhas, saida[0].length).setValues(saida);
  ss.toast(`${n} descrição(ões) atualizada(s).`, 'Novos cadastros', 5);
}

function completarDescricao(linha, notas) {
  const g = gerarDescricaoNovoProduto(linha);
  linha[N['Denominação (gerada)']] = g.denominacao;
  linha[N['Descrição (gerada)']] = g.descricao;
  linha[N['Observações']] = notas.concat(g.avisos).join('\n');
  return linha;
}

// Valor do INMETRO sem o rótulo (se colaram "CERTIFICADO: ..."), em maiúsculas e sem pontuação no fim
function semRotulo(valor, rotulo) {
  return String(valor || '').replace(/\s+/g, ' ').trim().replace(rotulo, '').replace(/[.;,]+$/, '').trim().toUpperCase();
}

// Denominação e descrição no padrão usado no Catálogo de Produtos
function gerarDescricaoNovoProduto(l) {
  const avisos = [];
  const ncm = String(l[N['NCM']] || '').replace(/\D/g, '');
  const medida = String(l[N['Medida']] || '').toUpperCase().replace(/,/g, '.').replace(/\s+/g, '');
  const indice = String(l[N['Índice / Lonas']] || '').toUpperCase().replace(/\s+/g, ' ').trim();
  const medidaVisual = [medida, indice.replace(/\b(TL|TT)\b/g, '').replace(/\s+/g, ' ').trim()].filter(Boolean).join(' ');
  const marca = String(l[N['Marca']] || '').toUpperCase().trim();
  const modelo = String(l[N['Modelo']] || '').toUpperCase().trim();
  const camara = String(l[N['Câmara (TL/TT)']] || '').toUpperCase().trim() || lerCamara(indice);
  const construcao = /\dZ?R\d/.test(medida) || /^\d{1,2}\.\d{2}R/.test(medida) ? 'RADIAL' : 'DIAGONAL';
  const ehLtOuC = /LT/.test(medida) || /\dC$/.test(medida) || /\/C/.test(medida);

  let dp, dsp;
  if (ncm === '40112090' && ehLtOuC) {
    [dp, dsp] = MAPA_NCM['40119090'];
    avisos.push('Exceção LT/C sobre 40112090: texto de light truck');
  } else {
    [dp, dsp] = MAPA_NCM[ncm] || ['PNEU', 'PNEUMATICOS NOVOS, DE BORRACHA'];
  }
  const partes = [dsp, `SENDO: CONSTRUCAO ${construcao}`];
  if (modelo) partes.push(`MODELO: ${modelo}`);
  if (camara === 'TL') partes.push('SEM CAMARA DE AR');
  else if (camara === 'TT') partes.push('COM CAMARA DE AR');
  const cert = semRotulo(l[N['Certificado INMETRO']], /^CERTIFICADO(\s+DE\s+CONFORMIDADE)?(\s+INMETRO)?\s*[:\-]?\s*/i);
  const reg = semRotulo(l[N['Registro INMETRO']],
    /^(N[UÚ]MERO\s+(DO\s+)?)?REGISTRO(\s+(DE\s+OBJETO|INMETRO|DO\s+PRODUTO|NO\s+INMETRO))*\s*[:\-]?\s*/i);
  const fam = semRotulo(l[N['Família INMETRO']], /^FAM[IÍ]LIA(\s+INMETRO)?\s*[:\-]?\s*/i);
  if (cert) partes.push(`CERTIFICADO: ${cert}`);
  if (reg) partes.push(`REGISTRO DE OBJETO: ${reg}`);
  if (fam) partes.push(`FAMILIA: ${fam}`);

  const denominacao = `${dp}, MEDIDA: ${medidaVisual}, MARCA: ${marca}`;
  const descricao = partes.join(', ').trim() + '.';
  if (!ncm) avisos.push('Falta NCM');
  else if (!MAPA_NCM[ncm]) avisos.push(`NCM ${ncm} fora do mapa de descrições: texto genérico`);
  if (!medida) avisos.push('Falta medida');
  if (!marca) avisos.push('Falta marca');
  if (!modelo) avisos.push('Falta modelo');
  if (!camara) avisos.push('Câmara (TL/TT) não informada');
  if (!cert || !reg || !fam) avisos.push('INMETRO incompleto (certificado, registro e família)');
  if (!l[N['Fabricante Código']] && !l[N['Fabricante Nome']]) avisos.push('Falta fabricante');
  if (denominacao.length > 120) avisos.push(`Denominação com ${denominacao.length} caracteres (máximo 120)`);
  if (descricao.length > 3700) avisos.push(`Descrição com ${descricao.length} caracteres (máximo 3700)`);
  return { denominacao, descricao, avisos };
}

// Sugestões tiradas do catálogo: NCM e câmara pela medida, fabricante pela marca
function Sugestoes(ss) {
  this.porMedida = new Map();
  this.porMarca = new Map();
  for (const l of lerLinhas(fonteDeProdutos(ss))) {
    const chave = String(l[P['Medida']] || '').replace(/\D/g, '');
    if (chave) {
      if (!this.porMedida.has(chave)) this.porMedida.set(chave, []);
      this.porMedida.get(chave).push(l);
    }
    const marca = soAlfanum(l[P['Marca']]);
    if (marca) {
      if (!this.porMarca.has(marca)) this.porMarca.set(marca, []);
      this.porMarca.get(marca).push(l);
    }
  }
  this.operadores = new Map();
  const abaOp = ss.getSheetByName(ABA_OPERADORES);
  if (abaOp) lerLinhas(abaOp).forEach(o => this.operadores.set(`${o[1]}|${o[0]}`, o));
}

Sugestoes.prototype.preencher = function (linha) {
  const notas = [];
  const mais = (lista, f) => {
    const cont = new Map();
    lista.map(f).filter(Boolean).forEach(v => cont.set(v, (cont.get(v) || 0) + 1));
    return [...cont.entries()].sort((a, b) => b[1] - a[1])[0] || null;
  };
  const mesmaMedida = this.porMedida.get(String(linha[N['Medida']]).replace(/\D/g, '')) || [];
  const filtro = montarFiltro([linha[N['Medida']], linha[N['Marca']], linha[N['Modelo']], linha[N['Índice / Lonas']], '', '']);
  const parecidos = outroIndiceOuLonas(mesmaMedida.map(prepararProduto), filtro);
  if (parecidos.length) {
    notas.push('Atenção: já existe cadastro desta medida/marca/modelo com outro índice/lonas: ' +
      parecidos.slice(0, 5).map(p => `${p.l[P['Código produto']]} (${indiceLonas(p)})`).join(', '));
  }
  const ncm = mais(mesmaMedida, l => String(l[P['NCM']] || ''));
  if (ncm) {
    linha[N['NCM']] = ncm[0];
    notas.push(`NCM sugerida: ${ncm[1]} de ${mesmaMedida.length} cadastro(s) com esta medida`);
  } else {
    const estimada = estimarNcm(linha[N['Medida']], linha[N['Índice / Lonas']]);
    if (estimada) {
      linha[N['NCM']] = estimada;
      notas.push('NCM estimada pelo tipo de medida (sem cadastro igual no catálogo): CONFERIR');
    }
  }
  const camaraTexto = lerCamara(linha[N['Índice / Lonas']]);
  const marca = soAlfanum(linha[N['Marca']]);
  const mesmaMarcaMedida = mesmaMedida.filter(l => soAlfanum(l[P['Marca']]) === marca);
  const camara = camaraTexto ? [camaraTexto] : mais(mesmaMarcaMedida.length ? mesmaMarcaMedida : mesmaMedida, l => l[P['Câmara']]);
  if (camara) {
    linha[N['Câmara (TL/TT)']] = camara[0];
    if (!camaraTexto) notas.push(`Câmara sugerida pelo catálogo: ${camara[0]}`);
  } else if (/\dZ?R\d/.test(String(linha[N['Medida']]).replace(/\s/g, ''))) {
    linha[N['Câmara (TL/TT)']] = 'TL';
    notas.push('Câmara estimada TL (radial): CONFERIR');
  }
  const daMarca = this.porMarca.get(marca) || [];
  const op = mais(daMarca, l => {
    const cod = String(l[P['Cód. operador estrangeiro']] || '').split('\n')[0];
    const pais = String(l[P['País operador']] || '').split('\n')[0];
    return cod ? `${pais}|${cod}|${String(l[P['Nome operador estrangeiro']] || '').split('\n')[0]}` : '';
  });
  if (op) {
    const [pais, cod, nome] = op[0].split('|');
    const cad = this.operadores.get(`${pais}|${cod}`) || [];
    linha[N['Fabricante Conhecido']] = 'SIM';
    linha[N['Fabricante Código']] = cod;
    linha[N['Fabricante Nome']] = nome || cad[2] || '';
    linha[N['Fabricante País']] = PAISES[pais] || pais;
    linha[N['Fabricante Logradouro']] = cad[3] || '';
    linha[N['Fabricante Cidade']] = cad[4] || '';
    notas.push(`Fabricante sugerido: ${op[1]} de ${daMarca.length} cadastro(s) da marca`);
  }
  return notas;
};

// Só quando não há cadastro com a mesma medida: LT ou C é light truck; aro 17.5/19.5/22.5 ou
// carga dupla com 12PR ou mais é caminhão; os demais radiais de aro inteiro são passeio.
function estimarNcm(medida, indice) {
  const m = semAcento(medida).replace(/\s/g, '');
  const radial = m.match(/^(LT)?\d{3}\/\d{2}Z?R(\d{2}(?:\.5)?)(C|LT)?$/);
  if (!radial) return '';
  if (radial[1] || radial[3]) return '40119090';
  const lonas = Number((semAcento(indice).match(/(\d{1,2})\s?PR/) || [])[1] || 0);
  if (/\.5$/.test(radial[2]) || (/\d\/\d/.test(semAcento(indice)) && lonas >= 12)) return '40112090';
  return '40111000';
}

// Menu: gera o Excel no padrão do sistema de cadastro e mostra o link para baixar
function baixarExcelNovos() {
  const g = ABAS_BUSCA;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(g.novos);
  if (!aba) throw new Error(`A aba ${g.novos} ainda não existe.`);
  const linhas = aba.getRange(NOVOS.primeira, 1, NOVOS.linhas, CAB_NOVOS.length).getDisplayValues()
    .filter(l => String(l[N['Denominação (gerada)']]).trim());
  if (!linhas.length) { avisar(`Não há descrições geradas na aba ${g.novos}.`); return; }
  const dados = linhas.map(l => ['', l[N['Denominação (gerada)']], l[N['Descrição (gerada)']],
    String(l[N['NCM']]).replace(/\D/g, ''), l[N['Fabricante Conhecido']], l[N['Fabricante PU']],
    l[N['Fabricante Código']], l[N['Fabricante Nome']], String(l[N['Fabricante País']]).toUpperCase(),
    l[N['Fabricante Logradouro']], l[N['Fabricante Cidade']]]);

  const nome = `novos_produtos_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm')}`;
  const arquivo = SpreadsheetApp.create(nome);
  const planilha = arquivo.getSheets()[0].setName('novos_produtos');
  planilha.getRange(1, 1, dados.length + 1, CAB_EXPORTACAO.length).setNumberFormat('@');  // NCM sempre como texto
  planilha.getRange(1, 1, 1, CAB_EXPORTACAO.length).setValues([CAB_EXPORTACAO]).setFontWeight('bold');
  planilha.getRange(2, 1, dados.length, CAB_EXPORTACAO.length).setValues(dados);
  SpreadsheetApp.flush();

  const id = arquivo.getId();
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;color:${COR.tinta}">
       <p style="margin:0 0 10px">Arquivo <b>${nome}.xlsx</b> com <b>${dados.length}</b> produto(s), aba
       <b>novos_produtos</b>, no padrão do sistema de cadastro.</p>
       <p><a href="https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx" target="_blank"
          style="display:inline-block;background:${COR.destaque};color:${COR.tinta};font-weight:bold;
          padding:10px 18px;border-radius:6px;text-decoration:none">Baixar Excel</a>
       &nbsp;<a href="https://docs.google.com/spreadsheets/d/${id}/edit" target="_blank"
          style="color:${COR.tintaMedia}">abrir no Google Sheets</a></p>
       <p style="font-size:12px;color:${COR.cinza}">Uma cópia fica no seu Google Drive com o mesmo nome.</p>
     </div>`).setWidth(460).setHeight(190);
  SpreadsheetApp.getUi().showModalDialog(html, 'Excel de novos cadastros');
}

function gravarOperadores(ss, idx) {
  const aba = obterAba(ss, ABA_OPERADORES, CAB_OPERADORES);
  const linhas = idx.operadores.map(o => [o.codigo || '', o.codigoPais || '', o.nome || '', o.logradouro || '',
    o.nomeCidade || '', o.tin || '', o.situacao || '']);
  escreverLinhas(aba, linhas, CAB_OPERADORES.length);
  try { aba.hideSheet(); } catch (e) { /* aba ativa não pode ser ocultada */ }
}


// PORTAL ÚNICO
function PortalUnico(cfg) {
  this.base = AMBIENTES[cfg.ambiente] || AMBIENTES.producao;
  this.cfg = cfg;
  this.headers = null;
}

// O portal recusa um novo login feito menos de 60 s depois do anterior (PLAT-ER2033):
// o token fica guardado no cache do usuário e é reaproveitado até expirar.
const CHAVE_CACHE_TOKEN = 'LPCO_TOKEN';
const VALIDADE_CACHE_TOKEN_S = 50 * 60;

PortalUnico.prototype.autenticar = function (forcarNovo) {
  const cache = CacheService.getUserCache();
  const chaveCache = `${CHAVE_CACHE_TOKEN}_${this.base}`;
  if (!forcarNovo) {
    const guardado = cache.get(chaveCache);
    if (guardado) { this.headers = JSON.parse(guardado); return; }
  }
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('LPCO_CLIENT_ID');
  const segredo = props.getProperty('LPCO_CLIENT_SECRET');
  if (!id || !segredo) throw new Error('Chave de acesso não configurada (menu Robô LPCO > Configurar chave de acesso).');
  const pedir = () => UrlFetchApp.fetch(`${this.base}/portal/api/autenticar/chave-acesso`, {
    method: 'post', muteHttpExceptions: true,
    headers: { 'Client-Id': id, 'Client-Secret': segredo, 'Role-Type': 'IMPEXP' },
  });
  let r = pedir();
  if (r.getResponseCode() !== 200 && r.getContentText().indexOf('PLAT-ER2033') >= 0) {
    Utilities.sleep(61000);  // login recente sem token guardado: espera a janela de 60 s
    r = pedir();
  }
  if (r.getResponseCode() !== 200) {
    throw new Error(`Falha na autenticação (${r.getResponseCode()}): ${r.getContentText().slice(0, 300)}`);
  }
  const h = cabecalhosMinusculos(r.getAllHeaders());
  this.headers = { 'Authorization': h['set-token'], 'X-CSRF-Token': h['x-csrf-token'] };
  cache.put(chaveCache, JSON.stringify(this.headers), VALIDADE_CACHE_TOKEN_S);
};

PortalUnico.prototype.requisitar = function (caminho, params) {
  const qs = params ? '?' + Object.keys(params).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&') : '';
  const url = this.base + caminho + qs;
  const esperas = [5000, 15000, 30000];
  for (let tentativa = 0; ; tentativa++) {
    Utilities.sleep(PAUSA_ENTRE_CHAMADAS_MS);
    let r;
    try {
      r = UrlFetchApp.fetch(url, { method: 'get', headers: this.headers, muteHttpExceptions: true });
    } catch (e) {  // queda de conexão
      if (tentativa >= esperas.length) throw e;
      Utilities.sleep(esperas[tentativa]);
      continue;
    }
    const codigo = r.getResponseCode();
    if (codigo === 401 && tentativa === 0) { this.autenticar(true); continue; }  // token expirado
    if (codigo >= 500 && tentativa < esperas.length) { Utilities.sleep(esperas[tentativa]); continue; }
    const texto = codigo === 200 || codigo === 206 ? '' : r.getContentText();
    if (texto.indexOf('PUCX-ER1001') >= 0) {
      const erro = new Error(JSON.parse(texto).message || texto);
      erro.limite = true;
      throw erro;
    }
    if (codigo === 204) return null;
    if (codigo !== 200 && codigo !== 206) throw new Error(`GET ${caminho} -> ${codigo}: ${texto.slice(0, 300)}`);
    return r;
  }
};

PortalUnico.prototype.get = function (caminho, params) {
  const r = this.requisitar(caminho, params);
  return r ? JSON.parse(r.getContentText()) : null;
};

// As exportações do catálogo vêm num .zip com um JSON dentro
PortalUnico.prototype.exportar = function (caminho) {
  const r = this.requisitar(caminho);
  if (!r) return [];
  const blob = r.getBlob().setContentType('application/zip');
  let arquivos;
  try { arquivos = Utilities.unzip(blob); } catch (e) { arquivos = [blob]; }
  let registros = [];
  for (const a of arquivos) {
    const texto = a.getDataAsString('UTF-8').replace(/^﻿/, '');
    if (!/^\s*[[{]/.test(texto)) continue;
    const dados = JSON.parse(texto);
    registros = registros.concat(Array.isArray(dados) ? dados : [dados]);
  }
  return registros;
};

function cabecalhosMinusculos(h) {
  const saida = {};
  Object.keys(h).forEach(k => { saida[k.toLowerCase()] = Array.isArray(h[k]) ? h[k][0] : h[k]; });
  return saida;
}


// CONTROLE E PLANILHA
function lerConfig(ss) {
  const ncms = String(lerControle(ss, 'NCMs') || '').split(/[,;\s]+/).map(n => n.replace(/\D/g, '')).filter(Boolean);
  const cnpj = String(lerControle(ss, 'CNPJ') || '').replace(/\D/g, '').padStart(14, '0');
  if (cnpj.length !== 14 || /^0+$/.test(cnpj)) throw new Error('Preencha o CNPJ na aba Controle.');
  if (!ncms.length) throw new Error('Preencha as NCMs na aba Controle.');
  const ultima = lerControle(ss, 'Última extração');
  return {
    cnpj, ncms,
    ambiente: String(lerControle(ss, 'Ambiente') || 'producao').trim(),
    diasMargem: Number(lerControle(ss, 'Dias de margem')) || 0,
    ultimaExtracao: paraData(ultima),
    buscarDesde: paraData(lerControle(ss, 'Buscar desde')),
    ncmsCatalogo: String(garantirControle(ss, 'NCMs do catálogo', NCMS_CATALOGO_PADRAO,
      'NCMs da aba Catálogo, separadas por vírgula.')).split(/[,;\s]+/).map(n => n.replace(/\D/g, '')).filter(Boolean),
  };
}

function linhaControle(ss, rotulo) {
  const aba = ss.getSheetByName(ABA_CONTROLE);
  if (!aba) throw new Error('Aba Controle não encontrada.');
  const valores = aba.getRange(1, 1, Math.max(aba.getLastRow(), 1), 1).getValues();
  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === rotulo) return { aba, linha: i + 1 };
  }
  return { aba, linha: 0 };
}

const NCMS_CATALOGO_PADRAO = '40111000, 40112090, 40119090, 40117090, 40118090';

function criarControle(ss) {
  const aba = ss.insertSheet(ABA_CONTROLE, 0);
  const linhas = [
    ['Robô LPCO - Controle', '', ''],
    ['', '', ''],
    ['Última extração', '', 'Data e hora da última atualização completa. A próxima busca começa nesta data.'],
    ['Período consultado na última execução', '', ''],
    ['LPCOs novas na última execução', '', ''],
    ['CNPJ', '', 'CNPJ do importador (14 dígitos).'],
    ['NCMs', '40111000, 40112090', 'NCMs das LPCOs consultadas, separadas por vírgula.'],
    ['Buscar desde', `01/01/${new Date().getFullYear()}`, 'Início da primeira extração (dd/mm/aaaa). Depois vale a última extração.'],
    ['Dias de margem', 30, 'Volta estes dias antes da última extração, para pegar LPCOs registradas antes e deferidas depois.'],
    ['Ambiente', 'producao', 'producao ou validacao'],
    ['NCMs do catálogo', NCMS_CATALOGO_PADRAO, 'NCMs da aba Catálogo Produtos, separadas por vírgula.'],
    ['Última atualização do catálogo', '', 'Quando a aba Catálogo Produtos foi baixada do Portal Único.'],
    ['', '', ''],
    ['Histórico de execuções', '', ''],
  ];
  // CNPJ, NCMs e datas digitadas como texto, para o Sheets não converter
  ['B6:B8', 'B11'].forEach(a => aba.getRange(a).setNumberFormat('@'));
  aba.getRange(1, 1, linhas.length, 3).setValues(linhas);
  aba.getRange(linhas.length + 1, 1, 1, 6)
    .setValues([['Data', 'Período', 'LPCOs novas', 'Itens incluídos', 'Produtos atualizados', 'Resultado']]);
  [250, 180, 620, 100, 130, 280].forEach((w, i) => aba.setColumnWidth(i + 1, w));
}

// Planilhas criadas antes desta versão não têm a linha: insere antes do histórico
function garantirControle(ss, rotulo, valorPadrao, nota) {
  const { aba, linha } = linhaControle(ss, rotulo);
  if (linha) return aba.getRange(linha, 2).getValue();
  const hist = linhaControle(ss, 'Histórico de execuções').linha;
  const nova = hist > 1 ? hist - 1 : aba.getLastRow() + 1;
  if (hist > 1) aba.insertRowsBefore(nova, 1);
  aba.getRange(nova, 1, 1, 3).setValues([[rotulo, valorPadrao, nota]]);
  return valorPadrao;
}

function lerControle(ss, rotulo) {
  const { aba, linha } = linhaControle(ss, rotulo);
  return linha ? aba.getRange(linha, 2).getValue() : '';
}

function gravarControle(ss, rotulo, valor) {
  garantirControle(ss, rotulo, '', '');
  const { aba, linha } = linhaControle(ss, rotulo);
  aba.getRange(linha, 2).setValue(valor);
}

function registrarHistorico(ss, linha) {
  const { aba, linha: cab } = linhaControle(ss, 'Histórico de execuções');
  if (!cab) return;
  aba.insertRowsAfter(cab + 1, 1);  // mais recente no topo
  aba.getRange(cab + 2, 1, 1, linha.length).setValues([linha]);
}

function obterAba(ss, nome, cabecalho) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.setFrozenRows(1);
  }
  // Mantém o cabeçalho em dia (abas antigas ganham as colunas novas, como Câmara)
  const atual = aba.getRange(1, 1, 1, cabecalho.length).getValues()[0];
  if (atual.some((v, i) => v !== cabecalho[i])) {
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
  }
  return aba;
}

function lerLinhas(aba) {
  const n = aba.getLastRow() - 1;
  if (n < 1) return [];
  return aba.getRange(2, 1, n, aba.getLastColumn()).getValues().filter(l => l[0] !== '');
}

function escreverLinhas(aba, linhas, largura) {
  const n = aba.getLastRow() - 1;
  if (n > 0) aba.getRange(2, 1, n, largura).clearContent();
  if (linhas.length) {
    garantirLinhas(aba, linhas.length + 1);
    // Linhas lidas de abas antigas podem ter menos colunas: completa com vazio
    aba.getRange(2, 1, linhas.length, largura).setValues(
      linhas.map(l => l.slice(0, largura).concat(new Array(Math.max(0, largura - l.length)).fill(''))));
  }
}

function colunaLetra(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function inicioDoDia(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function fmtIso(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmtData(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }

function ehData(v) { return Boolean(v) && typeof v.getTime === 'function' && !isNaN(v.getTime()); }

// 'dd/MM/yyyy', 'yyyy-MM-dd' ou 'yyyy-MM-ddTHH:mm...' -> Date (ou '' se vazio)
function paraData(v) {
  if (ehData(v)) return v;
  const s = String(v || '').trim();
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return '';
}


// LEITURA DA DESCRIÇÃO DO PNEU (mesmas regras do leitura_pneus.py)
function semAcento(t) {
  return String(t == null ? '' : t).normalize('NFKD').replace(/[̀-ͯ]/g, '').toUpperCase();
}
function soAlfanum(t) { return semAcento(t).replace(/[^0-9A-Z]/g, ''); }

const RE_RADIAL = /(?<![0-9])(LT|P)?\s?(\d{3})\s?\/\s?(\d{2})\s?(Z?R)\s?(\d{2})(?:\s?[.,]\s?(5)|\s(5)(?=\s))?(C(?![A-Z])|LT)?/;
const RE_RADIAL_SEM_PERFIL = /(?<![0-9/.])(LT|P)?(\d{3}|\d{1,2}[.,]\d{2})\s?(Z?R)\s?(\d{2})(?:[.,](5))?(C(?![A-Z])|LT)?/;
const RE_FLOTACAO = /(?<![0-9])(\d{2})\s?X\s?(\d{1,2})[.,](\d{1,2})\s?R\s?(\d{2})(LT)?/;
const RE_DIAGONAL = /(?<![0-9/.])(\d{2}00|\d{1,3}(?:[.,]\d{1,2})?L?(?:\/\d{2,3})?)\s?-\s?(\d{1,2}(?:[.,]5)?)(?![0-9])/;
const RE_CARGA = /(?<![0-9A-Z/.,])(\d{2,3}(?:\/\d{2,3})?)\s?([A-HJ-NP-Z]\d?)\s?(XL)?(?=$|[^0-9A-Z]|TL|TT|RF)/;
const RE_LONAS = /(?<![0-9])(\d{1,2})\s?-?\s?PR(?![A-Z])/;
const RE_LONAS_HIFEN = /^-(\d{2})(?![0-9])/;
const FIM_TRECHO = /,|;|\bMARCA\b|\bMODELO\b|\bCERTIFICADO\b|\bFAMILIA\b|\bREGISTRO\b| - /;
const PARADA = String.raw`(?=\s*(?:,|;|\s-\s|\.\s|\.?$|\n|\bCERTIFICADO\b|\bFAMILIA\b|\bREGISTRO\b|\bNUMERO\b|\bMODELO\b|\bMARCA\b|\bSEM CAMARA|\bCOM CAMARA|\bREF\b|\bTIPO\b|\bCCLASSTRIB\b))`;
const RE_MARCA = new RegExp(String.raw`\bMARCA\s*:?\s*([\s\S]+?)` + PARADA);
const RE_MODELO = new RegExp(String.raw`\bMODELO\s*:?\s*([\s\S]+?)` + PARADA);
const RE_REGISTRO = /\bREGISTRO\b[^0-9]{0,45}?(\d{3,6}\s?\/\s?\d{3,4})/;
const RE_CERTIFICADO = /\bCERTIFICADO\b(?:\s+DE\s+CONFORMIDADE)?(?:\s+INMETRO)?\s*:?\s*([0-9A-Z][0-9A-Z .\-/]*?)(?=\s*(?:,|;|\s-\s|\.\s|\.?$|\bREGISTRO\b|\bFAMILIA\b|\bNUMERO\b))/;
const RE_FAMILIA = /\bFAMILIA(?:\s+INMETRO)?\s*:?\s*(?:\(F\d+\)\s*)?([0-9][0-9A-Z]{3,11})/;
const SIGLAS_TECNICAS = new Set(['TT', 'TL', 'HD', 'SB', 'DOT', 'EU', 'M+S', 'RF', 'VF', 'IF', 'CFO']);

function lerMedida(texto) {
  const t = semAcento(texto);
  let m = RE_RADIAL.exec(t);
  if (m) {
    const aro = m[5] + (m[6] || m[7] ? '.5' : '');
    return [`${m[1] || ''}${m[2]}/${m[3]}${m[4]}${aro}${m[8] || ''}`, m.index + m[0].length];
  }
  // Flotação antes do radial sem perfil: senão 31X10.50R15 vira só 10.50R15
  m = RE_FLOTACAO.exec(t);
  if (m) return [`${m[1]}X${m[2]}.${m[3]}R${m[4]}${m[5] || ''}`, m.index + m[0].length];
  m = RE_RADIAL_SEM_PERFIL.exec(t);
  if (m) return [`${m[1] || ''}${m[2]}${m[3]}${m[4]}${m[5] ? '.5' : ''}${m[6] || ''}`, m.index + m[0].length];
  m = RE_DIAGONAL.exec(t);
  if (m) {
    let largura = m[1].replace(',', '.');
    if (/^\d{2}00$/.test(largura)) largura = largura.slice(0, 2) + '.00';
    return [`${largura}-${m[2].replace(',', '.')}`, m.index + m[0].length];
  }
  return ['', -1];
}

function lerMedidaCargaLonas(...textos) {
  for (const texto of textos) {
    const [medida, fim] = lerMedida(texto);
    if (!medida) continue;
    let t = semAcento(texto).slice(fim);
    const corte = FIM_TRECHO.exec(t);
    const trecho = corte ? t.slice(0, corte.index) : t;
    const separado = trecho.replace(/(?<=\d)(?=[A-Z])|(?<=PR)(?=\d)/g, ' ');
    let carga = '';
    const m = RE_CARGA.exec(trecho) || RE_CARGA.exec(separado);
    if (m) carga = `${m[1]}${m[2]}${m[3] || /\bXL\b/.test(trecho) ? ' XL' : ''}`;
    const l = RE_LONAS.exec(trecho) || RE_LONAS_HIFEN.exec(trecho);
    return [medida, carga, l ? `${parseInt(l[1], 10)}PR` : ''];
  }
  return ['', '', ''];
}

function limpar(t) { return t.replace(/\s+/g, ' ').trim().replace(/^[ .:\-/]+|[ .:\-/]+$/g, ''); }

function primeiroCampo(re, textos) {
  for (const texto of textos) {
    const m = re.exec(semAcento(texto));
    if (m && limpar(m[1])) return limpar(m[1]);
  }
  return '';
}

function modeloNoTrecho(texto) {
  const [medida, fim] = lerMedida(texto);
  if (!medida) return '';
  let t = semAcento(texto).slice(fim);
  const corte = FIM_TRECHO.exec(t);
  t = (corte ? t.slice(0, corte.index) : t).replace(/\([^)]*\)\S*/g, ' ');
  return t.split(/[\s_]+/).filter(x => x && !SIGLAS_TECNICAS.has(x) &&
    !new RegExp('^' + RE_LONAS.source + '$').test(x) && !/^(-\d+|PR)$/.test(x) &&
    !new RegExp('^(?:' + RE_CARGA.source + ')$').test(x)).join(' ');
}

function lerInmetro(texto) {
  const t = semAcento(texto);
  const reg = RE_REGISTRO.exec(t), cert = RE_CERTIFICADO.exec(t), fam = RE_FAMILIA.exec(t);
  return [reg ? reg[1].replace(/\s/g, '') : '', cert ? limpar(cert[1]) : '', fam ? fam[1] : ''];
}

function lerProdutoCatalogo(denominacao, descricao) {
  const [medida, carga, lonas] = lerMedidaCargaLonas(denominacao, descricao);
  let [registro, certificado, familia] = lerInmetro(descricao || denominacao);
  if (!(registro || certificado || familia)) [registro, certificado, familia] = lerInmetro(denominacao);
  const modelo = primeiroCampo(RE_MODELO, [descricao, denominacao]) || modeloNoTrecho(denominacao);
  const r = { medida, carga, lonas, marca: primeiroCampo(RE_MARCA, [denominacao, descricao]), modelo,
              registro, certificado, familia };
  if (r.carga && r.carga.indexOf('XL') < 0 && /\bXL\b/.test(semAcento(r.modelo))) r.carga += ' XL';
  return r;
}


