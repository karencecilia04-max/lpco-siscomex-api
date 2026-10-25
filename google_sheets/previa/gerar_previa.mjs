// Gera a prévia visual da planilha: roda o RoboLPCO.gs com dados fictícios e desenha cada aba
// num HTML com a aparência do Google Sheets. Também grava modelo.json, usado por
// gerar_xlsx.py para montar uma planilha que pode ser importada no Google Drive.
//
// Uso: node google_sheets/previa/gerar_previa.mjs

import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

import { colunaParaLetra, ehData } from './planilha_simulada.mjs';
import { rodarRoteiro } from './roteiro.mjs';

const SAIDA = new URL('saida/', import.meta.url);
const LARGURA_PADRAO = 100;
const MAX_LINHAS = 34;

const escapar = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Cor de fundo de uma célula: a dela, ou a da faixa alternada aplicada pelo script
function fundo(aba, linha, celula) {
  if (celula.fundo) return celula.fundo;
  const faixa = aba.faixas.find(f => linha >= f.linha);
  if (!faixa) return '';
  if (linha === faixa.linha) return faixa.cabecalho;
  return (linha - faixa.linha) % 2 ? faixa.primeira : faixa.segunda;
}

function corTexto(aba, linha, celula) {
  if (celula.cor) return celula.cor;
  const faixa = aba.faixas.find(f => linha === f.linha);
  return faixa ? '#FFFFFF' : '';
}

function desenharAba(aba, linhas, colunas, valores, rolagem) {
  const visiveis = colunas.filter(c => !aba.ocultas.has(c) && (!rolagem || c <= rolagem.fixas || c >= rolagem.desde));
  const largura = c => aba.larguras[c] || LARGURA_PADRAO;
  const total = 40 + visiveis.reduce((soma, c) => soma + largura(c), 0);
  const cab = visiveis.map(c => `<th style="width:${largura(c)}px">${colunaParaLetra(c)}</th>`);
  const corpo = linhas.map(l => {
    const tds = visiveis.map((c, i) => {
      const cel = aba.celulas.get(`${l},${c}`) || {};
      const texto = valores[l - 1][c - 1];
      const estilo = [
        fundo(aba, l, cel) && `background:${fundo(aba, l, cel)}`,
        corTexto(aba, l, cel) && `color:${corTexto(aba, l, cel)}`,
        cel.negrito && 'font-weight:700',
        cel.italico && 'font-style:italic',
        cel.tamanho && `font-size:${Math.round(cel.tamanho * 1.25)}px`,
        cel.fonte === 'Montserrat' && 'font-family:Montserrat,Arial,sans-serif',
        cel.alinhamento && `text-align:${cel.alinhamento}`,
        typeof cel.v === 'number' && !cel.alinhamento && 'text-align:right',
      ].filter(Boolean).join(';');
      // Texto sem quebra passa por cima das células vazias ao lado, como no Sheets:
      // para a esquerda quando alinhado à direita, para a direita nos outros casos
      const aDireita = cel.alinhamento === 'right';
      const lado = aDireita ? visiveis[i - 1] : visiveis[i + 1];
      const vizinha = lado ? valores[l - 1][lado - 1] || (aba.celulas.get(`${l},${lado}`) || {}).caixa : true;
      const transborda = aDireita ? 'vaza esquerda' : 'vaza';
      const classe = [cel.quebra ? 'quebra' : (vizinha ? 'corta' : transborda), cel.lista && 'lista']
        .filter(Boolean).join(' ');
      const conteudo = cel.caixa ? '<input type="checkbox" disabled>' : escapar(texto).replace(/\n/g, '<br>');
      return `<td class="${classe}" style="${estilo}"><div>${conteudo}</div></td>`;
    });
    return `<tr><th class="num">${l}</th>${tds.join('')}</tr>`;
  });
  return `<table style="width:${total}px"><thead><tr><th class="canto"></th>${cab.join('')}</tr></thead><tbody>${corpo.join('')}</tbody></table>`;
}

function pagina(planilha, ativa, rolagem) {
  const aba = planilha.getSheetByName(ativa);
  const ultimaLinha = Math.min(Math.max(aba.getLastRow() + 2, 12), MAX_LINHAS);
  const ultimaColuna = Math.max(aba.getLastColumn(), 6);
  const linhas = Array.from({ length: ultimaLinha }, (_, i) => i + 1);
  const colunas = Array.from({ length: ultimaColuna }, (_, i) => i + 1);
  const valores = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
  const abas = planilha.getSheets().filter(a => !a.oculta).map(a =>
    `<span class="aba${a.nome === ativa ? ' ativa' : ''}" style="border-bottom-color:${a.cor || 'transparent'}">${escapar(a.nome)}</span>`);

  return `<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>${escapar(ativa)} - Robô LPCO</title>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700&display=swap" rel="stylesheet">
<style>
  body { margin: 0; font-family: Arial, sans-serif; font-size: 13px; color: #202124; background: #fff; }
  .topo { display: flex; align-items: center; gap: 12px; padding: 10px 16px 0; }
  .icone { width: 26px; height: 34px; background: #188038; border-radius: 3px; }
  .titulo { font-size: 18px; }
  .menus { padding: 4px 54px 8px; color: #444; }
  .menus span { margin-right: 14px; }
  .menus .robo { font-weight: 700; color: #188038; }
  .barra { height: 34px; background: #edf2fa; margin: 0 8px; border-radius: 18px; }
  .grade { overflow: hidden; border-top: 1px solid #c7c7c7; margin-top: 8px; }
  table { border-collapse: collapse; table-layout: fixed; }
  th { background: #f8f9fa; color: #5f6368; font-weight: 400; font-size: 11px; border: 1px solid #c7c7c7; height: 20px; }
  th.canto, th.num { width: 40px; }
  td { border: 1px solid #e2e2e2; padding: 0; height: 21px; vertical-align: middle; position: relative; }
  td div { padding: 2px 4px; white-space: nowrap; }
  td.corta div { overflow: hidden; }
  td.vaza div { position: absolute; left: 0; top: 0; bottom: 0; display: flex; align-items: center; z-index: 1; }
  td.vaza.esquerda div { left: auto; right: 0; }
  td.quebra div { white-space: normal; }
  td.lista div::after { content: ""; position: absolute; right: 4px; top: 9px; border: 4px solid transparent; border-top-color: #5f6368; }
  input[type=checkbox] { margin: 0; width: 15px; height: 15px; }
  .abas { position: fixed; bottom: 0; left: 0; right: 0; background: #f8f9fa; border-top: 1px solid #c7c7c7; padding: 6px 12px; }
  .aba { display: inline-block; padding: 6px 14px; margin-right: 2px; border-bottom: 3px solid transparent; color: #444; }
  .aba.ativa { background: #e1e9f7; color: #0b57d0; font-weight: 700; border-radius: 4px 4px 0 0; }
</style></head><body>
<div class="topo"><div class="icone"></div><div class="titulo">Robô LPCO <span style="color:#5f6368;font-size:14px">(prévia com dados fictícios)</span></div></div>
<div class="menus"><span>Arquivo</span><span>Editar</span><span>Ver</span><span>Inserir</span><span>Formatar</span><span>Dados</span><span>Ferramentas</span><span>Extensões</span><span>Ajuda</span><span class="robo">Robô LPCO</span></div>
<div class="barra"></div>
<div class="grade">${desenharAba(aba, linhas, colunas, valores, rolagem)}</div>
<div class="abas">${abas.join('')}</div>
</body></html>`;
}

// Modelo simples de cada aba (valores exibidos, cores e larguras) para o gerar_xlsx.py
function modelo(planilha) {
  return planilha.getSheets().map(aba => {
    const nl = aba.getLastRow(), nc = aba.getLastColumn();
    const valores = nl && nc ? aba.getRange(1, 1, nl, nc).getValues() : [];
    const exibidos = nl && nc ? aba.getRange(1, 1, nl, nc).getDisplayValues() : [];
    return {
      nome: aba.nome, oculta: aba.oculta, cor: aba.cor, larguras: aba.larguras, ocultas: [...aba.ocultas],
      celulas: valores.flatMap((linha, i) => linha.map((v, j) => {
        const cel = aba.celulas.get(`${i + 1},${j + 1}`) || {};
        const texto = ehData(v) ? exibidos[i][j] : v;
        return {
          l: i + 1, c: j + 1, v: cel.caixa ? '' : texto, fundo: fundo(aba, i + 1, cel), cor: corTexto(aba, i + 1, cel),
          negrito: Boolean(cel.negrito), italico: Boolean(cel.italico), quebra: Boolean(cel.quebra), lista: cel.lista || null,
        };
      })).filter(c => c.v !== '' || c.fundo || c.lista),
    };
  });
}

const { planilha, robo } = rodarRoteiro();
mkdirSync(SAIDA, { recursive: true });
const ARQUIVOS = {
  'Buscar': 'buscar', 'Resultado da busca': 'resultado', 'Novos cadastros': 'novos_cadastros',
  'Catálogo Produtos': 'catalogo', 'Banco de LPCO': 'banco_lpco', 'Controle': 'controle',
};
for (const [nome, arquivo] of Object.entries(ARQUIVOS)) {
  writeFileSync(new URL(`${arquivo}.html`, SAIDA), pagina(planilha, nome));
}
// Novos cadastros rolada até a denominação e a descrição geradas (colunas A e B ficam congeladas)
const N = vm.runInContext('N', robo);
writeFileSync(new URL('novos_descricoes.html', SAIDA),
  pagina(planilha, 'Novos cadastros', { fixas: 2, desde: N['Denominação (gerada)'] + 1 }));
writeFileSync(new URL('modelo.json', SAIDA), JSON.stringify(modelo(planilha)));
console.log(`Prévia gravada em ${fileURLToPath(SAIDA)}`);
