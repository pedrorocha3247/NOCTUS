/**
 * Controle Bancário — NOCTUS. Exportar a tabela resumo (Fluxo × Relatório de débito) para Excel.
 *
 * 03/11/2026 (refeito a pedido do Rocha): a planilha reproduz a tela do sistema, na mesma ordem e estrutura:
 * totais no topo (Fluxo em dd/mm/aaaa, Relatório de débito, Diferença e o selo de divergências), os avisos,
 * e a tabela com as EMPRESAS NAS COLUNAS e as linhas Fluxo, Relatório e Diferença (diferença laranja quando o
 * fluxo é maior, vermelha quando o débito é maior, hífen verde quando bate). Visual claro, com as cores do
 * próprio sistema (tema claro), em Calibri.
 *
 * Gera um .xlsx de verdade (OOXML escrito à mão + ZIP sem compressão), no navegador, sem biblioteca.
 * Fluxo e Relatório são os valores dos relatórios; Diferença, totais e selo são FÓRMULAS (com o valor já
 * calculado gravado, para visualizadores que não recalculam; o Excel recalcula ao abrir).
 * Aqui só há conta e montagem de arquivo (sem tela, sem rede), para testar em Node.
 */

const COR = {
  texto: "111827", sec: "4B5563", mudo: "6B7280", borda: "D8DDE5", cab: "F1F3F6", zebra: "F8F9FB", vazio: "9CA3AF",
  verde: "0C6B3F", laranja: "B54708", laranjaFundo: "FDEBDD", erro: "C62828", erroFundo: "FBEAEA", erroBorda: "E5534B",
  aviso: "9A6700", avisoFundo: "FBF3E0", avisoBorda: "D9A441",
};
const FONTE = "Calibri";
const FMT_CONT = '#,##0.00;-#,##0.00;"-"';              // igual à tela: zero vira "-"
const FMT_REAIS = '"R$ "#,##0.00;-"R$ "#,##0.00';

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
const xml = (t) => String(t ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");   // caracteres que o XML não aceita
const letra = (i) => { let s = "", n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const fixed2 = (v) => r2(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ------------------------------------------------------------------ estilos */
function novoEstilo() {
  const fontes = [], fundos = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const bordas = ["<border><left/><right/><top/><bottom/><diagonal/></border>"], fmts = [], xfs = [], dxfs = [];
  const acha = (lista, item) => { let i = lista.indexOf(item); if (i < 0) { lista.push(item); i = lista.length - 1; } return i; };
  const fonte = (tam, neg, cor) => acha(fontes,
    `<font>${neg ? "<b/>" : ""}<sz val="${tam}"/><color rgb="FF${cor}"/><name val="${FONTE}"/><family val="2"/></font>`);
  fonte(11, false, COR.texto);   // fonte 0 = padrão do arquivo
  const lado = (nome, d) => d ? `<${nome} style="${d.estilo}"><color rgb="FF${d.cor}"/></${nome}>` : `<${nome}/>`;
  const registro = new Map();
  function xf({ tam = 11, neg = false, cor = COR.texto, fundo = null, borda = null, fmt = null, h = null, recuo = 0, quebra = false, v = "center" } = {}) {
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
  /** Formato condicional: cor, negrito e fundo. */
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
    texto(c, n, t, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xml(t)}</t></is></c>`); },
    /** Texto com trechos de estilos diferentes: partes = [{ t, tam, neg, cor }]. */
    rico(c, n, partes, s) {
      const runs = partes.map((p) => `<r><rPr>${p.neg ? "<b/>" : ""}<sz val="${p.tam}"/><color rgb="FF${p.cor}"/><rFont val="${FONTE}"/><family val="2"/></rPr>` +
        `<t xml:space="preserve">${xml(p.t)}</t></r>`).join("");
      linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}" t="inlineStr"><is>${runs}</is></c>`);
    },
    numero(c, n, v, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}"><v>${r2(v)}</v></c>`); },
    vazia(c, n, s) { linha(n).celulas.set(c, `<c r="${ref(c, n)}" s="${s}"/>`); },
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
  };
}

const rodape = (fonte) => {
  const f = `&"${FONTE},Regular"&10`;
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

/** Textos dos avisos: os mesmos que a tela mostra acima da tabela. erro = vermelho, senão âmbar. */
function textosAvisos(avisos, rotulo) {
  const out = [];
  for (const a of avisos || []) {
    if (a.tipo === "semFluxo") out.push({ erro: false, t: `${a.empresa} tem débito de R$ ${fixed2(a.total)} no relatório, sem valor no fluxo de ${rotulo}.` });
    if (a.tipo === "semConta") out.push({ erro: true, t: `R$ ${fixed2(a.total)} debitados em conta bancária não cadastrada no NOCTUS: não entraram em nenhuma empresa. Confira o relatório.` });
    if (a.tipo === "naoConfere") out.push({ erro: true, t: `${a.empresa}: o total impresso no relatório (R$ ${fixed2(a.total)}) não bate com a soma das solicitações (R$ ${fixed2(a.soma)}). Confira o relatório.` });
    if (a.tipo === "semCodigo") out.push({ erro: false, t: `${a.empresa} está no fluxo, mas não tem código de empresa cadastrado: não dá para comparar com o relatório.` });
  }
  return out;
}

/**
 * @param {object} p
 *   iso, rotulo ("03/11/2026"), cols (montarResumo().cols), total, avisos (montarResumo().avisos),
 *   lanc (opcional, uma contagem por coluna: lançamentos da empresa na data, mostrada como "NOME · n" igual à tela),
 *   arquivoFluxo, arquivoDebito (nomes dos arquivos importados, para o rodapé), agora (Date, só para teste)
 * @returns {Uint8Array} conteúdo do .xlsx
 */
export function gerarXlsxResumo({ iso, rotulo, cols, total, avisos = [], lanc = [], arquivoFluxo = "", arquivoDebito = "", agora = new Date() }) {
  if (!Array.isArray(cols) || !cols.length) throw new Error("tabela resumo vazia");
  if (total?.relatorio == null) throw new Error("sem relatório de débito");
  const est = novoEstilo();
  const A = novaAba();
  const merges = [];
  const N = cols.length;                       // empresas
  const ULT = Math.max(N, 6);                  // última coluna usada (índice): no mínimo G, para caber o topo
  const col = (i) => letra(i);
  const textosAv = textosAvisos(avisos, rotulo);

  // linhas
  const R_LAB = 2, R_VAL = 3, R_ESP = 4, R_AV = 5;
  const R_CAB = R_AV + textosAv.length, R_FLX = R_CAB + 1, R_REL = R_CAB + 2, R_DIF = R_CAB + 3;
  const intervalo = (r) => `B${r}:${col(N)}${r}`;

  // valores calculados (cópia para quem não recalcula)
  const difDe = (c) => (c.relatorio == null ? null : r2(c.fluxo - c.relatorio));
  const somaFluxo = r2(cols.reduce((s, c) => s + c.fluxo, 0));
  const somaRel = r2(cols.reduce((s, c) => s + (c.relatorio == null ? 0 : c.relatorio), 0));
  const somaDif = r2(somaFluxo - somaRel);
  const nDiverg = cols.filter((c) => { const d = difDe(c); return d != null && d !== 0; }).length;
  const textoSelo = nDiverg === 0 ? "Sem divergência" : `${nDiverg} ${nDiverg === 1 ? "divergência" : "divergências"}`;

  /* ---- topo: totais e selo ---- */
  A.altura(1, 10);
  A.altura(R_LAB, 20); A.altura(R_VAL, 34); A.altura(R_ESP, 12);
  const sRot = est.xf({ tam: 10, neg: true, cor: COR.sec, h: "left", v: "bottom", recuo: 1 });
  const sKpi = (cor = COR.texto) => est.xf({ tam: 18, neg: true, cor, fmt: FMT_REAIS, h: "left", recuo: 1 });
  A.texto(0, R_LAB, `FLUXO EM ${rotulo}`, sRot);
  A.formula(0, R_VAL, `SUM(${intervalo(R_FLX)})`, somaFluxo, sKpi());
  A.texto(1, R_LAB, "RELATÓRIO DE DÉBITO", sRot); A.vazia(2, R_LAB, sRot); merges.push(`B${R_LAB}:C${R_LAB}`);
  A.formula(1, R_VAL, `SUM(${intervalo(R_REL)})`, somaRel, sKpi()); A.vazia(2, R_VAL, sKpi()); merges.push(`B${R_VAL}:C${R_VAL}`);
  A.texto(3, R_LAB, "DIFERENÇA", sRot); A.vazia(4, R_LAB, sRot); merges.push(`D${R_LAB}:E${R_LAB}`);
  A.formula(3, R_VAL, `ROUND(A${R_VAL}-B${R_VAL},2)`, somaDif, sKpi()); A.vazia(4, R_VAL, sKpi()); merges.push(`D${R_VAL}:E${R_VAL}`);
  const sSelo = est.xf({ neg: true, cor: COR.texto, fundo: COR.cab, h: "center" });
  const nF = `SUMPRODUCT(--ISNUMBER(${intervalo(R_DIF)}),--(${intervalo(R_DIF)}<>0))`;
  A.formula(ULT - 1, R_VAL, `IF(${nF}=0,"Sem divergência",${nF}&IF(${nF}=1," divergência"," divergências"))`, textoSelo, sSelo);
  A.vazia(ULT, R_VAL, sSelo); merges.push(`${col(ULT - 1)}${R_VAL}:${col(ULT)}${R_VAL}`);
  const sEsp = est.xf({ borda: { bottom: { estilo: "thin", cor: COR.borda } } });
  for (let c = 0; c <= ULT; c++) A.vazia(c, R_ESP, sEsp);

  /* ---- avisos (mesmo texto da tela) ---- */
  textosAv.forEach((a, i) => {
    const n = R_AV + i;
    const s = est.xf({ cor: a.erro ? COR.erro : COR.aviso, fundo: a.erro ? COR.erroFundo : COR.avisoFundo, h: "left", recuo: 1, quebra: true,
      borda: { left: { estilo: "medium", cor: a.erro ? COR.erroBorda : COR.avisoBorda } } });
    A.altura(n, a.t.length > 120 ? 36 : 24);
    A.texto(0, n, a.t, s);
    for (let c = 1; c <= ULT; c++) A.vazia(c, n, s);
    merges.push(`A${n}:${col(ULT)}${n}`);
  });

  /* ---- tabela: EMPRESA | empresas nas colunas ---- */
  const bFina = { bottom: { estilo: "thin", cor: COR.borda } };
  const temSemFluxo = cols.some((c) => c.semFluxo);
  A.altura(R_CAB, temSemFluxo ? 38 : 28);
  const sCabRot = est.xf({ tam: 10, neg: true, cor: COR.sec, fundo: COR.cab, h: "left", recuo: 1, quebra: true, borda: bFina });
  const sCab = est.xf({ tam: 10, neg: true, cor: COR.sec, fundo: COR.cab, h: "right", recuo: 1, quebra: true, borda: bFina });
  A.texto(0, R_CAB, "EMPRESA", sCabRot);
  cols.forEach((c, i) => {
    const nome = String(c.empresa).toUpperCase() + (lanc[i] > 0 ? ` · ${lanc[i]}` : "");
    if (c.semFluxo) A.rico(i + 1, R_CAB, [{ t: nome, tam: 10, neg: true, cor: COR.sec }, { t: "\nSEM FLUXO", tam: 8, neg: true, cor: COR.aviso }], sCab);
    else A.texto(i + 1, R_CAB, nome, sCab);
  });

  A.altura(R_FLX, 26); A.altura(R_REL, 26); A.altura(R_DIF, 30);
  const sRotulo = (fundo, borda) => est.xf({ neg: true, cor: COR.texto, fundo, h: "left", recuo: 1, borda });
  const bTopo = { top: { estilo: "thin", cor: COR.borda }, bottom: { estilo: "thin", cor: COR.borda } };
  A.texto(0, R_FLX, "Fluxo", sRotulo(null, bFina));
  A.texto(0, R_REL, "Relatório", sRotulo(COR.zebra, bFina));
  A.texto(0, R_DIF, "Diferença", sRotulo(null, bTopo));
  const sFlx = est.xf({ neg: true, fmt: FMT_CONT, h: "right", recuo: 1, borda: bFina });
  const sRel = est.xf({ neg: true, fmt: FMT_CONT, h: "right", recuo: 1, fundo: COR.zebra, borda: bFina });
  const sRelVazio = est.xf({ neg: true, cor: COR.vazio, h: "right", recuo: 1, fundo: COR.zebra, borda: bFina });
  const sDif = est.xf({ neg: true, fmt: FMT_CONT, h: "right", recuo: 1, borda: bTopo });
  const sDifVazio = est.xf({ neg: true, cor: COR.vazio, h: "right", recuo: 1, borda: bTopo });
  cols.forEach((c, i) => {
    const k = i + 1, L = col(k), d = difDe(c);
    A.numero(k, R_FLX, c.fluxo, sFlx);
    if (c.relatorio == null) A.texto(k, R_REL, "—", sRelVazio); else A.numero(k, R_REL, c.relatorio, sRel);
    A.formula(k, R_DIF, `IF(ISNUMBER(${L}${R_REL}),ROUND(${L}${R_FLX}-${L}${R_REL},2),"—")`, d == null ? "—" : d, d == null ? sDifVazio : sDif);
  });

  /* ---- cores da diferença (iguais às da tela): laranja = fluxo maior, vermelho = débito maior, verde = bate ---- */
  const dxLar = est.dxf({ cor: COR.laranja, neg: true, fundo: COR.laranjaFundo });
  const dxVer = est.dxf({ cor: COR.erro, neg: true, fundo: COR.erroFundo });
  const dxOk = est.dxf({ cor: COR.verde, neg: true });
  const dxLarT = est.dxf({ cor: COR.laranja, neg: true }), dxVerT = est.dxf({ cor: COR.erro, neg: true });
  const t = `B${R_DIF}`;
  const cf =
    `<conditionalFormatting sqref="${intervalo(R_DIF)}">` +
      `<cfRule type="expression" dxfId="${dxLar}" priority="1"><formula>AND(ISNUMBER(${t}),${t}&gt;0)</formula></cfRule>` +
      `<cfRule type="expression" dxfId="${dxVer}" priority="2"><formula>AND(ISNUMBER(${t}),${t}&lt;0)</formula></cfRule>` +
      `<cfRule type="expression" dxfId="${dxOk}" priority="3"><formula>AND(ISNUMBER(${t}),${t}=0)</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="D${R_VAL}">` +
      `<cfRule type="cellIs" dxfId="${dxLarT}" priority="4" operator="greaterThan"><formula>0</formula></cfRule>` +
      `<cfRule type="cellIs" dxfId="${dxVerT}" priority="5" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>`;

  /* ---- larguras: pelo maior nome ou valor de cada coluna ---- */
  const larg = [];
  larg[0] = Math.max(26, `FLUXO EM ${rotulo}`.length * 1.25 + 4, `R$ ${fixed2(somaFluxo)}`.length * 2 + 2);
  for (let c = 1; c <= ULT; c++) {
    const e = cols[c - 1];
    if (!e) { larg[c] = 17; continue; }
    const nome = String(e.empresa).toUpperCase() + (lanc[c - 1] > 0 ? ` · ${lanc[c - 1]}` : "");
    const vals = [e.fluxo, e.relatorio, difDe(e)].filter((v) => v != null).map((v) => fixed2(v).length);
    larg[c] = Math.max(17, Math.ceil(nome.length * 1.15 + 3), Math.max(...vals) * 1.2 + 4);
  }
  const colsXml = larg.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${Math.round(w * 10) / 10}" customWidth="1"/>`).join("");

  const fonteTxt = [arquivoFluxo, arquivoDebito].filter(Boolean).map((x) => String(x).slice(0, 60)).join(" e ");
  const planilha =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    `<dimension ref="A1:${col(ULT)}${R_DIF}"/>` +
    '<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0">' +
      '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>' +
    `<cols>${colsXml}</cols>` +
    `<sheetData>${A.dados()}</sheetData>` +
    `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` +
    cf +
    '<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
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
      '<sheets><sheet name="Tabela resumo" sheetId="1" r:id="rId1"/></sheets>' +
      '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'],
    ["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", est.arquivo()],
    ["xl/worksheets/sheet1.xml", planilha],
  ];
  return zipar(arquivos, agora);
}
