// Planilha do Google simulada em memória, para rodar o RoboLPCO.gs fora do Google.
// Imita só o que o script usa: valores, cores, fontes, larguras, caixas de seleção e
// inserção de linhas. O portal é trocado pelas respostas de exemplos/portal_ficticio.json.

import { readFileSync } from 'node:fs';
import vm from 'node:vm';

function letraParaColuna(letras) {
  return [...letras].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function colunaParaLetra(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

// Datas criadas dentro do script vêm de outro contexto do Node: instanceof Date não serve
export function ehData(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

export function formatarData(d, formato) {
  const dois = n => String(n).padStart(2, '0');
  return formato.toLowerCase()
    .replace('dd', dois(d.getDate())).replace('mm', dois(d.getMonth() + 1)).replace('yyyy', d.getFullYear())
    .replace('hh', dois(d.getHours())).replace(/:mm$/, ':' + dois(d.getMinutes()));
}

class Intervalo {
  constructor(aba, linha, coluna, nLinhas, nColunas) {
    Object.assign(this, { aba, linha, coluna, nLinhas, nColunas });
    // Métodos de formatação que a prévia não precisa representar
    const proxy = new Proxy(this, {
      get: (alvo, nome) => (nome in alvo ? alvo[nome] : () => proxy),
    });
    return proxy;
  }

  celulas() {
    const lista = [];
    for (let i = 0; i < this.nLinhas; i++) {
      for (let j = 0; j < this.nColunas; j++) lista.push([i, j, this.aba.celula(this.linha + i, this.coluna + j)]);
    }
    return lista;
  }

  matriz(f) {
    return Array.from({ length: this.nLinhas }, (_, i) =>
      Array.from({ length: this.nColunas }, (_, j) => f(this.aba.celula(this.linha + i, this.coluna + j))));
  }

  aplicar(propriedade, valor) {
    this.celulas().forEach(([, , c]) => { c[propriedade] = valor; });
    return this;
  }

  getValues() { return this.matriz(c => (c.v === undefined ? '' : c.v)); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValues() {
    return this.matriz(c => {
      if (ehData(c.v)) {
        const comHora = c.v.getHours() || c.v.getMinutes();
        const formato = c.formato && c.formato !== '@' ? c.formato : (comHora ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy');
        return formatarData(c.v, formato);
      }
      return c.v === undefined || c.v === null ? '' : String(c.v);
    });
  }
  setValues(valores) {
    valores.forEach((linha, i) => linha.forEach((v, j) => { this.aba.celula(this.linha + i, this.coluna + j).v = v; }));
    return this;
  }
  setValue(v) { return this.aplicar('v', v); }
  clearContent() { return this.aplicar('v', ''); }
  setBackground(cor) { return this.aplicar('fundo', cor); }
  setBackgrounds(cores) {
    cores.forEach((linha, i) => linha.forEach((cor, j) => { this.aba.celula(this.linha + i, this.coluna + j).fundo = cor; }));
    return this;
  }
  setFontColor(cor) { return this.aplicar('cor', cor); }
  setFontWeight(peso) { return this.aplicar('negrito', peso === 'bold'); }
  setFontStyle(estilo) { return this.aplicar('italico', estilo === 'italic'); }
  setFontSize(tamanho) { return this.aplicar('tamanho', tamanho); }
  setFontFamily(fonte) { return this.aplicar('fonte', fonte); }
  setHorizontalAlignment(a) { return this.aplicar('alinhamento', a); }
  setWrap(w) { return this.aplicar('quebra', w); }
  setNumberFormat(f) { return this.aplicar('formato', f); }
  insertCheckboxes() { return this.aplicar('caixa', true); }
  removeCheckboxes() { return this.aplicar('caixa', false); }
  setDataValidation(regra) { return this.aplicar('lista', regra && regra.valores); }
  getA1Notation() { return `${colunaParaLetra(this.coluna)}${this.linha}`; }
  applyRowBanding() {
    const faixa = { linha: this.linha, nColunas: this.nColunas };
    this.aba.faixas.push(faixa);
    const encadear = {
      setHeaderRowColor: c => { faixa.cabecalho = c; return encadear; },
      setFirstRowColor: c => { faixa.primeira = c; return encadear; },
      setSecondRowColor: c => { faixa.segunda = c; return encadear; },
    };
    return encadear;
  }
}

class Aba {
  constructor(planilha, nome) {
    Object.assign(this, { planilha, nome, celulas: new Map(), maxLinhas: 1000, larguras: {}, ocultas: new Set(),
                          faixas: [], cor: null, oculta: false });
    // Congelar linhas, altura de linha, linhas de grade: sem efeito na prévia
    const proxy = new Proxy(this, {
      get: (alvo, nome) => (nome in alvo ? alvo[nome] : () => proxy),
    });
    return proxy;
  }

  celula(l, c) {
    const chave = `${l},${c}`;
    if (!this.celulas.has(chave)) this.celulas.set(chave, {});
    return this.celulas.get(chave);
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const [ini, fim] = a.split(':');
      const ler = ref => {
        const m = ref.match(/^([A-Z]+)(\d*)$/);
        return [m[2] ? Number(m[2]) : 1, letraParaColuna(m[1]), !m[2]];
      };
      const [l1, c1, soColuna] = ler(ini);
      const [l2, c2] = fim ? ler(fim) : [l1, c1];
      return new Intervalo(this, l1, c1, soColuna ? this.maxLinhas : l2 - l1 + 1, c2 - c1 + 1);
    }
    return new Intervalo(this, a, b, c || 1, d || 1);
  }

  ultimo(indice) {
    let maior = 0;
    for (const [chave, c] of this.celulas) {
      if (c.v !== undefined && c.v !== '' && c.v !== null) maior = Math.max(maior, Number(chave.split(',')[indice]));
    }
    return maior;
  }

  getLastRow() { return this.ultimo(0); }
  getLastColumn() { return this.ultimo(1); }
  getMaxRows() { return this.maxLinhas; }
  getName() { return this.nome; }
  setName(n) { this.nome = n; return this; }
  getParent() { return this.planilha; }
  setTabColor(cor) { this.cor = cor; return this; }
  setColumnWidth(col, largura) { this.larguras[col] = largura; return this; }
  hideColumns(col) { this.ocultas.add(col); return this; }
  hideSheet() { this.oculta = true; return this; }
  getBandings() { return this.faixas.map(f => ({ remove: () => this.faixas.splice(this.faixas.indexOf(f), 1) })); }
  getImages() { return []; }
  clear() { this.celulas.clear(); this.faixas = []; return this; }

  insertRowsAfter(depois, n) {
    const movidas = new Map();
    for (const [chave, c] of this.celulas) {
      const [l, col] = chave.split(',').map(Number);
      movidas.set(l > depois ? `${l + n},${col}` : chave, c);
    }
    this.celulas = movidas;
    this.maxLinhas += n;
    return this;
  }
  insertRowsBefore(antes, n) { return this.insertRowsAfter(antes - 1, n); }
}

class Planilha {
  constructor() { this.abas = []; this.ativa = null; this.avisos = []; }
  getSheetByName(nome) { return this.abas.find(a => a.nome === nome) || null; }
  getSheets() { return this.abas; }
  insertSheet(nome, posicao) {
    const aba = new Aba(this, nome);
    this.abas.splice(posicao === undefined ? this.abas.length : posicao, 0, aba);
    return aba;
  }
  deleteSheet(aba) { this.abas.splice(this.abas.indexOf(aba), 1); }
  setActiveSheet(aba) { this.ativa = aba; return aba; }
  getActiveSheet() { return this.ativa || this.abas[0]; }
  toast(msg) { this.avisos.push(msg); }
}

function validacao() {
  const regra = {};
  const construtor = {
    requireValueInList: v => { regra.valores = v; return construtor; },
    setAllowInvalid: () => construtor,
    build: () => regra,
  };
  return construtor;
}

// Portal ficticio: mesma interface do PortalUnico do script
function portalFicticio(dados) {
  function PortalFicticio() {}
  PortalFicticio.prototype.autenticar = function () {};
  PortalFicticio.prototype.get = function (caminho, params) {
    if (caminho.endsWith('/consulta')) {
      if (params.offset) return null;
      return dados.consulta.filter(l => l.situacao === params.situacao &&
        l.dataRegistro >= params['data-inicial-registro'] && l.dataRegistro <= params['data-final-registro'] &&
        dados.detalhes[l.numero].listaNcm.some(it => it.ncm === params['ncm-inicial']));
    }
    return dados.detalhes[caminho.split('/').pop()] || null;
  };
  PortalFicticio.prototype.exportar = function (caminho) {
    if (caminho.includes('/produto/')) return dados.produtos;
    if (caminho.includes('/fabricante/')) return dados.vinculos;
    return dados.operadores;
  };
  return PortalFicticio;
}

// Carrega o RoboLPCO.gs num contexto com a planilha simulada
export function carregarRobo(arquivoGs, arquivoPortal) {
  const planilha = new Planilha();
  const contexto = vm.createContext({
    console: { log() {}, warn() {}, error() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => planilha,
      getUi: () => ({ alert: msg => planilha.avisos.push(msg) }),
      newDataValidation: validacao,
      BorderStyle: { SOLID: 'solid' },
      flush() {},
    },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo' },
    Utilities: {
      formatDate: (d, _, f) => formatarData(d, f.replace(/MM/g, 'mm').replace('HHmm', 'hhmm')),
      sleep() {},
    },
  });
  vm.runInContext(readFileSync(arquivoGs, 'utf8'), contexto);
  contexto.PortalUnico = portalFicticio(JSON.parse(readFileSync(arquivoPortal, 'utf8')));
  return { planilha, robo: contexto };
}

export { colunaParaLetra };
