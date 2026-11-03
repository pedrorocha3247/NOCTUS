/**
 * Controle Bancário — NOCTUS. Exportar a tabela resumo (Fluxo × Relatório de débito) para Excel (03/11/2026).
 *
 * Gera um .xlsx de verdade (OOXML escrito à mão + ZIP sem compressão), no navegador, sem biblioteca:
 * o SheetJS que o NOCTUS já carrega não grava estilos, e o padrão de planilha do Rocha exige fonte, cores,
 * sem linhas de grade, filtro e painel congelado.
 *
 * Padrão (Vinho & Champanhe): Abadi 11 (12 só no título), sem itálico, sem linhas de grade, cabeçalho vinho
 * com texto champanhe, faixa dourada no topo, linha fina bege entre os dados, cabeçalho congelado (só linha),
 * filtro, paisagem. Todo número derivado é FÓRMULA (diferença, situação, totais, cartões, conferência);
 * valores digitados só para o que veio dos relatórios. As fórmulas levam também o valor já calculado, para
 * visualizadores que não recalculam; o Excel recalcula ao abrir.
 *
 * Duas abas: "Resumo" (uma linha por empresa) e "Conferência" (checagens por fórmula + parágrafo pronto).
 * Aqui só há conta e montagem de arquivo (sem tela, sem rede), para testar em Node.
 */

const COR = {
  vinho: "4E1F2C", champanhe: "F1E6D2", dourado: "A8864F", verde: "5E7D5B", ocre: "B98A2A", tijolo: "B23A3A",
  cinza: "6E6266", dado: "262626", bege: "EAE2DA", fundoProblema: "F5E1E1", fundoAtencao: "F7EDD6", vazio: "A6A6A6",
};
const FMT_VALOR = "#,##0.00;[Red]-#,##0.00";
const FMT_REAIS = '"R$ "#,##0.00';
const FMT_INT = "0";

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
const xml = (t) => String(t ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");   // caracteres que o XML não aceita
const letra = (i) => String.fromCharCode(65 + i);                // colunas A..Z (aqui são no máximo 6)
const fixed2 = (v) => r2(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ------------------------------------------------------------------ estilos */
function novoEstilo() {
  const fontes = [], fundos = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const bordas = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], fmts = [], xfs = [], dxfs = [];
  const acha = (lista, item) => { let i = lista.indexOf(item); if (i < 0) { lista.push(item); i = lista.length - 1; } return i; };
  const fonte = (tam, neg, cor) => acha(fontes,
    `<font>${neg ? "<b/>" : ""}<sz val="${tam}"/><color rgb="FF${cor}"/><name val="Abadi"/><family val="2"/></font>`);
  fonte(11, false, COR.dado);   // fonte 0 = padrão do arquivo: Abadi 11 (células vazias também)
  const lado = (nome, d) => d ? `<${nome} style="${d.estilo}"><color rgb="FF${d.cor}"/></${nome}>` : `<${nome}/>`;
  const registro = new Map();
  /** Devolve o índice do estilo de célula. Mesmo pedido, mesmo índice. */
  function xf({ tam = 11, neg = false, cor = COR.dado, fundo = null, borda = null, fmt = null, h = null, recuo = 0, quebra = false, v = "center" } = {}) {
    const chave = JSON.stringify([tam, neg, cor, fundo, borda, fmt, h, recuo, quebra, v]);
    if (registro.has(chave)) return registro.get(chave);
    const fo = fonte(tam, neg, cor);
    const fi = fundo ? acha(fundos, `<fill><patternFill patternType="solid"><fgColor rgb="FF${fundo}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const bo = borda ? acha(bordas, `<border>${lado("left", borda.left)}${lado("right", borda.right)}${lado("top", borda.top)}${lado("bottom", borda.bottom)}<diagonal/></border>`) : 0;
    let nf = 0;
    if (fmt) { let i = fmts.indexOf(fmt); if (i < 0) { fmts.push(fmt); i = fmts.length - 1; } nf = 164 + i; }
    const al = `<alignment${h ? ` horizontal="${h}"` : ""} vertical="${v}"${recuo ? ` indent="${recuo}"` : ""}${quebra ? ' wrapText="1"' : ""}/>`;
    xfs.push(`<xf numFmtId="${nf}" fontId="${fo}" fillId="${fi}" borderId="${bo}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
    registro.set(chave, xfs.length - 1);
    return xfs.length - 1;
  }
  /** Formato condicional: só cor, negrito e fundo (o tamanho vem da célula). */
  function dxf({ cor = null, neg = false, fundo = null }) {
    return acha(dxfs, `<dxf>${cor || neg ? `<font>${neg ? "<b/>" : ""}${cor ? `<color rgb="FF${cor}"/>` : ""}</font>` : ""}` +
      `${fundo ? `<fill><patternFill><bgColor rgb="FF${fundo}"/></patternFill></fill>` : ""}</dxf>`);
  }
  function arquivo() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (fmts.length ? `<numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xml(f)}"/>`).join("")}</numFmts>` : "") +
      `<fonts count="${fontes.length}">${fontes.join("")}</fonts>` +
      `<fills count="${fundos.length}">${fundos.join("")}</fills>` +
      `<borders count="${bordas.length}">${bordas.join("")}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>` +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
  }
  return { xf, dxf, arquivo };
}

/* ------------------------------------------------------------------ planilha */
function novaAba() {
  const linhas = new Map();   // n → { ht, celulas: Map(col → xml) }
  const linha = (n) => { if (!linhas.has(n)) linhas.set(n, { ht: null, celulas: new Map() }); return linhas.get(n); };
  const ref = (c, n) => `${letra(c)}${n}`;
  return {
    altura(n, ht) { linha(n).ht = ht; },
    /** Texto fixo. */
    texto(c, n, t, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xml(t)}</t></is></c>`); },
    numero(c, n, v, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}"><v>${r2(v)}</v></c>`); },
    vazia(c, n, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}"/>`); },
    /** Fórmula com o valor já calculado (número ou texto). */
    formula(c, n, f, valor, s) {
      const txt = typeof valor === "string";
      linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}"${txt ? ' t="str"' : ""}><f>${xml(f)}</f><v>${txt ? xml(valor) : r2(valor)}</v></c>`);
    },
    dados() {
      return [...linhas.keys()].sort((a, b) => a - b).map((n) => {
        const l = linhas.get(n);
        return `<row r="${n}"${l.ht ? ` ht="${l.ht}" customHeight="1"` : ""}>${[...l.celulas.keys()].sort((a, b) => a - b).map((c) => l.celulas.get(c)).join("")}</row>`;
      }).join("");
    },
    ultimaLinha: () => Math.max(...linhas.keys()),
  };
}

const rodape = (fonte) => {
  const f = '&"Abadi,Regular"&11';
  const t = String(fonte || "").replace(/&/g, "&&");
  return `<headerFooter><oddFooter>${xml(`&L${f}${t}&C${f}Gerado em &D&R${f}Página &P de &N`)}</oddFooter></headerFooter>`;
};

/* ------------------------------------------------------------------ ZIP (sem compressão) */
const TABELA_CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = (b) => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = TABELA_CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };

export function zipar(arquivos, agora = new Date()) {
  const enc = new TextEncoder();
  const hora = (agora.getHours() << 11) | (agora.getMinutes() << 5) | (agora.getSeconds() >> 1);
  const data = ((agora.getFullYear() - 1980) << 9) | ((agora.getMonth() + 1) << 5) | agora.getDate();
  const partes = [], central = [];
  let pos = 0;
  for (const [nome, conteudo] of arquivos) {
    const nb = enc.encode(nome), cb = typeof conteudo === "string" ? enc.encode(conteudo) : conteudo, crc = crc32(cb);
    const loc = new DataView(new ArrayBuffer(30));
    loc.setUint32(0, 0x04034b50, true); loc.setUint16(4, 20, true); loc.setUint16(6, 0x0800, true); loc.setUint16(8, 0, true);
    loc.setUint16(10, hora, true); loc.setUint16(12, data, true); loc.setUint32(14, crc, true);
    loc.setUint32(18, cb.length, true); loc.setUint32(22, cb.length, true); loc.setUint16(26, nb.length, true); loc.setUint16(28, 0, true);
    partes.push(new Uint8Array(loc.buffer), nb, cb);
    const cen = new DataView(new ArrayBuffer(46));
    cen.setUint32(0, 0x02014b50, true); cen.setUint16(4, 20, true); cen.setUint16(6, 20, true); cen.setUint16(8, 0x0800, true); cen.setUint16(10, 0, true);
    cen.setUint16(12, hora, true); cen.setUint16(14, data, true); cen.setUint32(16, crc, true);
    cen.setUint32(20, cb.length, true); cen.setUint32(24, cb.length, true); cen.setUint16(28, nb.length, true);
    cen.setUint32(42, pos, true);
    central.push(new Uint8Array(cen.buffer), nb);
    pos += 30 + nb.length + cb.length;
  }
  const tamCentral = central.reduce((s, p) => s + p.length, 0);
  const fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arquivos.length, true); fim.setUint16(10, arquivos.length, true);
  fim.setUint32(12, tamCentral, true); fim.setUint32(16, pos, true);
  const todos = [...partes, ...central, new Uint8Array(fim.buffer)];
  const saida = new Uint8Array(todos.reduce((s, p) => s + p.length, 0));
  let o = 0; for (const p of todos) { saida.set(p, o); o += p.length; }
  return saida;
}

/* ------------------------------------------------------------------ arquivo */
export function nomeArquivoResumo(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
  return `Controle_Bancario_Tabela_Resumo_${m ? `${m[3]}${m[2]}${m[1]}` : "data"}.xlsx`;
}

/**
 * @param {object} p
 *   iso, rotulo ("03/11/2026"), cols (montarResumo().cols), total, divergencias, avisos (montarResumo().avisos),
 *   arquivoFluxo, arquivoDebito (nomes dos arquivos importados, para o rodapé), agora (Date, só para teste)
 * @returns {Uint8Array} conteúdo do .xlsx
 */
export function gerarXlsxResumo({ iso, rotulo, cols, total, avisos = [], arquivoFluxo = "", arquivoDebito = "", agora = new Date() }) {
  if (!Array.isArray(cols) || !cols.length) throw new Error("tabela resumo vazia");
  if (total?.relatorio == null) throw new Error("sem relatório de débito");
  const est = novoEstilo();

  /* ============ aba Resumo ============ */
  const A = novaAba();
  const NCOL = 6;
  const R_CARTAO = 6, R_CAB = 9, R_INI = 10, R_FIM = R_INI + cols.length - 1, R_TOT = R_FIM + 1;
  const sFaixa = est.xf({ fundo: COR.dourado });
  A.altura(1, 4); for (let c = 0; c < NCOL; c++) A.vazia(c, 1, sFaixa);
  A.altura(2, 24); A.texto(0, 2, "Tabela resumo Fluxo × Relatório de débito", est.xf({ tam: 12, neg: true, cor: COR.vinho }));
  A.texto(0, 3, `Controle Bancário · Fluxo em ${rotulo} · ${cols.length} ${cols.length === 1 ? "empresa" : "empresas"}`, est.xf({ cor: COR.cinza }));

  // totais por fórmula (valores já calculados para quem não recalcula)
  const somaFluxo = r2(cols.reduce((s, c) => s + c.fluxo, 0));
  const somaRel = r2(cols.reduce((s, c) => s + (c.relatorio == null ? 0 : c.relatorio), 0));
  const difLinha = (c) => (c.relatorio == null ? null : r2(c.fluxo - c.relatorio));
  const situacao = (d) => (d == null ? "Não comparado" : d === 0 ? "Confere" : d > 0 ? "Fluxo maior" : "Relatório maior");
  const somaDif = r2(cols.reduce((s, c) => s + (difLinha(c) == null ? 0 : difLinha(c)), 0));
  const nDiverg = cols.filter((c) => { const d = difLinha(c); return d != null && d !== 0; }).length;

  // cartões (linhas 6 e 7): rótulo em maiúsculas e valor; barra esquerda na cor do status
  const cartao = (c, rotulo2, formula, valor, fmt, corBarra, corTexto) => {
    const borda = { left: { estilo: "medium", cor: corBarra } };
    A.texto(c, R_CARTAO, rotulo2, est.xf({ neg: true, cor: COR.cinza, borda, recuo: 1 }));
    A.formula(c, R_CARTAO + 1, formula, valor, est.xf({ neg: true, cor: corTexto, borda, fmt, h: "left", recuo: 1 }));
  };
  A.altura(R_CARTAO, 20); A.altura(R_CARTAO + 1, 22);
  cartao(0, "EMPRESAS", `COUNTA(A${R_INI}:A${R_FIM})`, cols.length, FMT_INT, COR.dourado, COR.dourado);
  cartao(2, "FLUXO", `C${R_TOT}`, somaFluxo, FMT_REAIS, COR.dourado, COR.dourado);
  cartao(3, "RELATÓRIO DE DÉBITO", `D${R_TOT}`, somaRel, FMT_REAIS, COR.dourado, COR.dourado);
  cartao(4, "DIFERENÇA", `E${R_TOT}`, somaDif, FMT_REAIS, somaDif === 0 ? COR.verde : COR.ocre, somaDif === 0 ? COR.verde : COR.ocre);
  cartao(5, "DIVERGÊNCIAS", `COUNTIF(F${R_INI}:F${R_FIM},"Fluxo maior")+COUNTIF(F${R_INI}:F${R_FIM},"Relatório maior")`, nDiverg, FMT_INT,
    nDiverg === 0 ? COR.verde : COR.tijolo, nDiverg === 0 ? COR.verde : COR.tijolo);

  // cabeçalho da tabela
  A.altura(R_CAB, 32);
  const cab = (h, recuo = 0) => est.xf({ neg: true, cor: COR.champanhe, fundo: COR.vinho, h, recuo, quebra: true });
  A.texto(0, R_CAB, "Empresa", cab("left", 1)); A.texto(1, R_CAB, "Cód.", cab("center"));
  A.texto(2, R_CAB, "Fluxo (R$)", cab("right")); A.texto(3, R_CAB, "Relatório de débito (R$)", cab("right"));
  A.texto(4, R_CAB, "Diferença (R$)", cab("right")); A.texto(5, R_CAB, "Situação", cab("center"));

  // dados
  const linhaFina = { bottom: { estilo: "thin", cor: COR.bege } };
  const sNome = est.xf({ borda: linhaFina, h: "left", recuo: 1 }), sCod = est.xf({ borda: linhaFina, h: "center" });
  const sVal = est.xf({ borda: linhaFina, fmt: FMT_VALOR, h: "right" }), sSit = est.xf({ borda: linhaFina, h: "center" });
  const sVazioDireita = est.xf({ cor: COR.vazio, borda: linhaFina, h: "right" }), sVazioCentro = est.xf({ cor: COR.vazio, borda: linhaFina, h: "center" });
  cols.forEach((c, i) => {
    const n = R_INI + i, d = difLinha(c);
    A.altura(n, 20);
    A.texto(0, n, c.empresa, sNome);
    if (c.codigo != null && c.codigo !== "") { if (/^\d+$/.test(String(c.codigo))) A.numero(1, n, Number(c.codigo), est.xf({ borda: linhaFina, h: "center", fmt: FMT_INT })); else A.texto(1, n, String(c.codigo), sCod); }
    else A.texto(1, n, "—", sVazioCentro);
    A.numero(2, n, c.fluxo, sVal);
    if (c.relatorio == null) A.texto(3, n, "—", sVazioDireita); else A.numero(3, n, c.relatorio, sVal);
    A.formula(4, n, `IF(ISNUMBER(D${n}),ROUND(C${n}-D${n},2),"—")`, d == null ? "—" : d, d == null ? sVazioDireita : sVal);
    A.formula(5, n, `IF(NOT(ISNUMBER(E${n})),"Não comparado",IF(E${n}=0,"Confere",IF(E${n}>0,"Fluxo maior","Relatório maior")))`, situacao(d), sSit);
  });

  // linha de TOTAL (SUBTOTAL respeita o filtro)
  const bTot = { top: { estilo: "thin", cor: COR.vinho } };
  A.altura(R_TOT, 22);
  A.texto(0, R_TOT, "TOTAL", est.xf({ neg: true, cor: COR.vinho, borda: bTot, h: "left", recuo: 1 }));
  A.vazia(1, R_TOT, est.xf({ neg: true, cor: COR.vinho, borda: bTot }));
  const sTot = est.xf({ neg: true, cor: COR.vinho, borda: bTot, fmt: FMT_VALOR, h: "right" });
  A.formula(2, R_TOT, `SUBTOTAL(9,C${R_INI}:C${R_FIM})`, somaFluxo, sTot);
  A.formula(3, R_TOT, `SUBTOTAL(9,D${R_INI}:D${R_FIM})`, somaRel, sTot);
  A.formula(4, R_TOT, `SUBTOTAL(9,E${R_INI}:E${R_FIM})`, somaDif, sTot);
  A.vazia(5, R_TOT, est.xf({ neg: true, cor: COR.vinho, borda: bTot }));

  // formatação condicional: o status colore a linha e a palavra
  const dxAt = est.dxf({ fundo: COR.fundoAtencao }), dxPr = est.dxf({ fundo: COR.fundoProblema });
  const dxOk = est.dxf({ cor: COR.verde }), dxOc = est.dxf({ cor: COR.ocre, neg: true }), dxTj = est.dxf({ cor: COR.tijolo, neg: true });
  const cfResumo =
    `<conditionalFormatting sqref="A${R_INI}:F${R_FIM}">` +
      `<cfRule type="expression" dxfId="${dxAt}" priority="1"><formula>$F${R_INI}="Fluxo maior"</formula></cfRule>` +
      `<cfRule type="expression" dxfId="${dxPr}" priority="2"><formula>$F${R_INI}="Relatório maior"</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="F${R_INI}:F${R_FIM}">` +
      `<cfRule type="cellIs" dxfId="${dxOk}" priority="3" operator="equal"><formula>"Confere"</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxOc}" priority="4" operator="equal"><formula>"Fluxo maior"</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxTj}" priority="5" operator="equal"><formula>"Relatório maior"</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="E${R_CARTAO + 1}">` +
      `<cfRule type="cellIs" dxfId="${dxOk}" priority="6" operator="equal"><formula>0</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxOc}" priority="7" operator="notEqual"><formula>0</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="F${R_CARTAO + 1}">` +
      `<cfRule type="cellIs" dxfId="${dxOk}" priority="8" operator="equal"><formula>0</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxTj}" priority="9" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`;

  const fonteTxt = [arquivoFluxo, arquivoDebito].filter(Boolean).map((t) => String(t).slice(0, 60)).join(" e ");
  const planilha1 =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    `<dimension ref="A1:F${R_TOT}"/>` +
    `<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">` +
      `<pane ySplit="${R_CAB}" topLeftCell="A${R_INI}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${R_INI}" sqref="A${R_INI}"/></sheetView></sheetViews>` +
    '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>' +
    '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="9" customWidth="1"/><col min="3" max="3" width="21" customWidth="1"/>' +
      '<col min="4" max="4" width="25" customWidth="1"/><col min="5" max="5" width="21" customWidth="1"/><col min="6" max="6" width="21" customWidth="1"/></cols>' +
    `<sheetData>${A.dados()}</sheetData>` +
    `<autoFilter ref="A${R_CAB}:F${R_FIM}"/>` +
    cfResumo +
    '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
    rodape(`Fonte: ${fonteTxt || "relatórios do Controle Bancário"}`) + '</worksheet>';

  /* ============ aba Conferência ============ */
  const B = novaAba();
  const rng = (col) => `Resumo!${col}${R_INI}:${col}${R_FIM}`;
  const semConta = r2(avisos.filter((a) => a.tipo === "semConta").reduce((s, a) => s + Number(a.total || 0), 0));
  const nomes = cols.map((c) => c.empresa);
  const repetidas = nomes.filter((n) => nomes.filter((m) => m === n).length > 1).length;
  const naoComp = cols.filter((c) => c.relatorio == null).length;
  // [nome, esperado, formula do encontrado, valor encontrado, formato]
  const checagens = [
    ["Quantidade: empresas na tabela", cols.length, `COUNTA(${rng("A")})`, cols.length, FMT_INT],
    ["Soma: fluxo das empresas × total do fluxo", total.fluxo, `SUM(${rng("C")})`, somaFluxo, FMT_VALOR],
    ["Soma: relatório das empresas × total do relatório de débito", total.relatorio, `SUM(${rng("D")})`, somaRel, FMT_VALOR],
    ["Integridade: fluxo menos relatório × soma das diferenças", total.diferenca, `SUM(${rng("E")})`, somaDif, FMT_VALOR],
    ["Duplicidade: empresas repetidas", 0, `SUMPRODUCT(--(COUNTIF(${rng("A")},${rng("A")})>1))`, repetidas, FMT_INT],
    ["Cobertura: empresas sem código, não comparadas", 0, `COUNTIF(${rng("F")},"Não comparado")`, naoComp, FMT_INT],
    ["Relatório: valor debitado em conta não cadastrada", 0, null, semConta, FMT_VALOR],
  ];
  const C_CAB = 8, C_INI = 9, C_FIM = C_INI + checagens.length - 1;
  B.altura(1, 4); for (let c = 0; c < 5; c++) B.vazia(c, 1, sFaixa);
  B.altura(2, 24); B.texto(0, 2, "Conferência da tabela resumo", est.xf({ tam: 12, neg: true, cor: COR.vinho }));
  B.texto(0, 3, `Controle Bancário · Fluxo em ${rotulo}`, est.xf({ cor: COR.cinza }));

  const falhas = checagens.filter(([, esp, , enc]) => r2(enc - esp) !== 0).length;
  const statusGeral = falhas === 0 ? "CONFERIDO" : "COM PENDÊNCIAS";
  const bStatus = { left: { estilo: "medium", cor: falhas === 0 ? COR.verde : COR.tijolo } };
  B.altura(5, 20); B.altura(6, 22);
  B.texto(0, 5, "STATUS GERAL", est.xf({ neg: true, cor: COR.cinza, borda: bStatus, recuo: 1 }));
  B.formula(0, 6, `IF(COUNTIF(E${C_INI}:E${C_FIM},"FALHA")=0,"CONFERIDO","COM PENDÊNCIAS")`, statusGeral,
    est.xf({ neg: true, cor: falhas === 0 ? COR.verde : COR.tijolo, borda: bStatus, h: "left", recuo: 1 }));

  B.altura(C_CAB, 32);
  B.texto(0, C_CAB, "Checagem", cab("left", 1)); B.texto(1, C_CAB, "Esperado", cab("right")); B.texto(2, C_CAB, "Encontrado", cab("right"));
  B.texto(3, C_CAB, "Diferença", cab("right")); B.texto(4, C_CAB, "Resultado", cab("center"));
  checagens.forEach(([nome, esp, f, enc, fmt], i) => {
    const n = C_INI + i, dif = r2(enc - esp), ok = dif === 0;
    const sv = est.xf({ borda: linhaFina, fmt, h: "right" });
    B.altura(n, 20);
    B.texto(0, n, nome, sNome);
    B.numero(1, n, esp, sv);
    if (f) B.formula(2, n, f, enc, sv); else B.numero(2, n, enc, sv);
    B.formula(3, n, `ROUND(C${n}-B${n},2)`, dif, sv);
    B.formula(4, n, `IF(ROUND(C${n}-B${n},2)=0,"OK","FALHA")`, ok ? "OK" : "FALHA", sSit);
  });

  const dxFalha = est.dxf({ cor: COR.tijolo, neg: true, fundo: COR.fundoProblema });
  const R_RES = C_FIM + 2, R_PAR = R_RES + 1;
  B.altura(R_RES, 24); B.texto(0, R_RES, "Resumo", est.xf({ tam: 12, neg: true, cor: COR.vinho }));
  const aresta = (c) => `Resumo!${c}${R_TOT}`;
  const paragrafo =
    `Foram conferidas ${cols.length} empresas: fluxo de R$ ${fixed2(somaFluxo)}, relatório de débito de R$ ${fixed2(somaRel)} e diferença de R$ ${fixed2(somaDif)}. ` +
    `${nDiverg === 0 ? "Nenhuma empresa com divergência." : `${nDiverg} empresa(s) com divergência.`} ` +
    `${falhas === 0 ? "Todas as checagens da planilha fecharam." : `${falhas} checagem(ns) com falha.`}`;
  const fParagrafo =
    `"Foram conferidas "&Resumo!A${R_CARTAO + 1}&" empresas: fluxo de R$ "&FIXED(${aresta("C")},2)&", relatório de débito de R$ "&FIXED(${aresta("D")},2)` +
    `&" e diferença de R$ "&FIXED(${aresta("E")},2)&". "&IF(Resumo!F${R_CARTAO + 1}=0,"Nenhuma empresa com divergência.",Resumo!F${R_CARTAO + 1}&" empresa(s) com divergência.")` +
    `&" "&IF(COUNTIF(E${C_INI}:E${C_FIM},"FALHA")=0,"Todas as checagens da planilha fecharam.",COUNTIF(E${C_INI}:E${C_FIM},"FALHA")&" checagem(ns) com falha.")`;
  B.altura(R_PAR, 64);
  const sPar = est.xf({ quebra: true, h: "left", v: "top", recuo: 1 });
  B.formula(0, R_PAR, fParagrafo, paragrafo, sPar);
  for (let c = 1; c < 5; c++) B.vazia(c, R_PAR, sPar);

  const planilha2 =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    `<dimension ref="A1:E${R_PAR}"/>` +
    `<sheetViews><sheetView showGridLines="0" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">` +
      `<pane ySplit="${C_CAB}" topLeftCell="A${C_INI}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${C_INI}" sqref="A${C_INI}"/></sheetView></sheetViews>` +
    '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>' +
    '<cols><col min="1" max="1" width="62" customWidth="1"/><col min="2" max="4" width="20" customWidth="1"/><col min="5" max="5" width="16" customWidth="1"/></cols>' +
    `<sheetData>${B.dados()}</sheetData>` +
    '<sheetProtection sheet="1" objects="1" scenarios="1"/>' +
    `<mergeCells count="1"><mergeCell ref="A${R_PAR}:E${R_PAR}"/></mergeCells>` +
    `<conditionalFormatting sqref="E${C_INI}:E${C_FIM}"><cfRule type="cellIs" dxfId="${dxOk}" priority="1" operator="equal"><formula>"OK"</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxFalha}" priority="2" operator="equal"><formula>"FALHA"</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="A6"><cfRule type="cellIs" dxfId="${dxOk}" priority="3" operator="equal"><formula>"CONFERIDO"</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxTj}" priority="4" operator="equal"><formula>"COM PENDÊNCIAS"</formula></cfRule></conditionalFormatting>` +
    '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
    rodape(`Fonte: ${fonteTxt || "relatórios do Controle Bancário"}`) + '</worksheet>';

  /* ============ pacote ============ */
  const titulo = "Tabela resumo Fluxo × Relatório de débito";
  const iso8601 = agora.toISOString().replace(/\.\d{3}Z$/, "Z");
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const arquivos = [
    ["[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
    ["_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`],
    ["docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${xml(titulo)}</dc:title><dc:subject>${xml(`Fluxo em ${rotulo}`)}</dc:subject><dc:creator>NOCTUS</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${iso8601}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso8601}</dcterms:modified></cp:coreProperties>`],
    ["docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>NOCTUS</Application></Properties>'],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS} xmlns:r="${REL}">` +
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews>' +
      '<sheets><sheet name="Resumo" sheetId="1" r:id="rId1"/><sheet name="Conferência" sheetId="2" r:id="rId2"/></sheets>' +
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Resumo!$A$${R_CAB}:$F$${R_FIM}</definedName>` +
      `<definedName name="_xlnm.Print_Titles" localSheetId="0">Resumo!$${R_CAB}:$${R_CAB}</definedName>` +
      `<definedName name="_xlnm.Print_Titles" localSheetId="1">'Conferência'!$${C_CAB}:$${C_CAB}</definedName></definedNames>` +
      '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'],
    ["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>` +
      `<Relationship Id="rId3" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", est.arquivo()],
    ["xl/worksheets/sheet1.xml", planilha1],
    ["xl/worksheets/sheet2.xml", planilha2],
  ];
  return zipar(arquivos, agora);
}
