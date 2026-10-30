/**
 * Controle Bancário (NOCTUS) — leitura do relatório "Fluxo de Caixa - Pag. Futuros".
 *
 * Parte 1 (30/10/2026, pedido do Rocha): identificar a data do relatório e as
 * empresas com valor nessa data. Funções puras: recebem a planilha já convertida
 * em linhas (matriz) e devolvem o resultado; quem lê o .xlsx é a tela.
 *
 * Formato do relatório (arquivo real de 29/10/2026):
 *   linha 3  → nomes das empresas (D3:P3)
 *   linha 4  → unidade (MATRIZ, SLIM, PMONEY…)
 *   linha 5+ → coluna C com a data "dd-mm" e um valor por empresa, como texto "98985,70"
 *   depois   → "Sub Total =>", "Total =>", "Total Geral =>" (ignorados)
 * A leitura procura a data e o cabeçalho PELO CONTEÚDO, não por endereço fixo,
 * para não quebrar se o sistema de origem acrescentar uma linha ou coluna.
 *
 * Regras de negócio:
 *   - a data vem do próprio arquivo ("03-11") e recebe o ANO VIGENTE (03/11/2026);
 *   - entra na tabela só a empresa com valor diferente de zero naquela data;
 *   - um valor que não dá para ler NÃO é tratado como zero: o arquivo é recusado,
 *     para nunca sumir um pagamento da tabela em silêncio.
 */

/** Erro com texto para mostrar à pessoa (não é falha técnica). */
export function avisoLeitura(msg) {
  const e = new Error(msg);
  e.aviso = true;
  return e;
}

const limpo = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
const pad2 = (n) => String(n).padStart(2, "0");
const arred2 = (n) => Math.round(n * 100) / 100;

/** "98985,70", "98.985,70", "1234.56", 98985.7, "(5,00)", "-5" → número; vazio → 0; ilegível → NaN. */
export function lerValor(v) {
  if (typeof v === "number") return Number.isFinite(v) ? arred2(v) : NaN;
  let t = limpo(v).replace(/^R\$\s?/i, "").replace(/\s/g, "");
  if (!t || t === "-" || t === "—") return 0;
  const neg = /^\(.*\)$/.test(t) || /^-/.test(t) || /-$/.test(t);
  t = t.replace(/[()]/g, "").replace(/^-|-$/g, "");
  if (t.includes(",")) {
    if (!/^\d{1,3}(\.\d{3})*,\d+$|^\d+,\d+$/.test(t)) return NaN;
    t = t.replace(/\./g, "").replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(t)) {
    t = t.replace(/\./g, "");          // "1.234.567" sem centavos
  } else if (!/^\d+(\.\d+)?$/.test(t)) {
    return NaN;
  }
  const n = Number(t);
  return Number.isFinite(n) ? arred2(neg ? -n : n) : NaN;
}

/**
 * Data do relatório. "03-11" ou "03/11" → { dia: 3, mes: 11 } com o ano vigente;
 * aceita também data completa ("03/11/2026") ou célula de data do Excel.
 * Devolve null quando a célula não é uma data.
 */
export function lerData(v, anoVigente) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return montarData(v.getDate(), v.getMonth() + 1, v.getFullYear());
  }
  const t = limpo(v);
  let m = /^(\d{1,2})[-/.](\d{1,2})$/.exec(t);
  if (m) return montarData(Number(m[1]), Number(m[2]), anoVigente, t);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(t);
  if (m) return montarData(Number(m[1]), Number(m[2]), Number(m[3]), t);
  return null;
}
function montarData(dia, mes, ano, bruto) {
  const d = new Date(ano, mes - 1, dia);
  if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) {
    return { invalida: true, bruto: bruto ?? `${dia}-${mes}` };
  }
  return {
    iso: `${ano}-${pad2(mes)}-${pad2(dia)}`,
    rotulo: `${pad2(dia)}/${pad2(mes)}/${ano}`,
    bruto: bruto ?? `${pad2(dia)}-${pad2(mes)}`,
  };
}

const ehRotuloTotal = (t) => /total/i.test(t);

/**
 * Lê a matriz de linhas (linhas[i][j] = valor da célula) e devolve
 * { datas: [{ iso, rotulo, bruto, empresas: [{ empresa, unidade, saidas }], total }], colunas }.
 * `empresas` vem só com quem tem valor ≠ 0, na ordem das colunas do relatório.
 */
export function lerFluxoCaixa(linhas, { anoVigente = new Date().getFullYear() } = {}) {
  if (!Array.isArray(linhas) || !linhas.length) {
    throw avisoLeitura("A planilha está vazia. Confira se é o relatório de fluxo de caixa exportado em Excel.");
  }
  const largura = Math.max(...linhas.map((l) => (Array.isArray(l) ? l.length : 0)));

  // 1. linhas de data: uma célula "dd-mm" e, à direita dela, pelo menos um valor
  const linhasData = [];
  for (let i = 0; i < linhas.length; i++) {
    const l = linhas[i] || [];
    const rotuloLinha = l.map(limpo).find((t) => t && !/^[\d\s.,()\-/R$]+$/i.test(t)) || "";
    if (ehRotuloTotal(rotuloLinha)) continue;           // "Sub Total =>" etc.
    for (let j = 0; j < l.length; j++) {
      const d = lerData(l[j], anoVigente);
      if (!d) continue;
      linhasData.push({ i, j, d });
      break;
    }
  }
  if (!linhasData.length) {
    throw avisoLeitura("Não encontrei a data do relatório (ex.: 03-11). Confira se é o relatório de fluxo de caixa de pagamentos futuros.");
  }
  const invalida = linhasData.find((x) => x.d.invalida);
  if (invalida) throw avisoLeitura(`A data "${invalida.d.bruto}" do relatório não existe no calendário. Confira o arquivo.`);

  const colData = linhasData[0].j;
  const primeira = linhasData[0].i;

  // 2. cabeçalho: linhas logo acima da primeira data com 2+ textos à direita da coluna da data.
  //    A de cima é o nome da empresa; a de baixo (se houver) é a unidade.
  const cab = [];
  for (let i = primeira - 1; i >= 0; i--) {
    const l = linhas[i] || [];
    const textos = l.slice(colData + 1).filter((v) => limpo(v) !== "").length;
    if (textos < 2) break;
    cab.unshift(i);
  }
  if (!cab.length) {
    throw avisoLeitura("Não encontrei a linha com os nomes das empresas acima da data. Confira se é o relatório de fluxo de caixa.");
  }
  const linhaEmpresa = linhas[cab[0]] || [];
  const linhaUnidade = cab.length > 1 ? linhas[cab[1]] || [] : [];

  const colunas = [];
  for (let j = colData + 1; j < largura; j++) {
    const nome = limpo(linhaEmpresa[j]).toUpperCase();
    if (!nome) continue;
    colunas.push({ j, empresa: nome, unidade: limpo(linhaUnidade[j]).toUpperCase() });
  }
  if (!colunas.length) throw avisoLeitura("O relatório não tem nenhuma empresa no cabeçalho.");

  // mesmo nome em duas colunas (ex.: MOMENTUM MATRIZ e MOMENTUM FILIAL): mostra a unidade junto
  const vezes = {};
  colunas.forEach((c) => { vezes[c.empresa] = (vezes[c.empresa] || 0) + 1; });
  colunas.forEach((c) => {
    c.nome = vezes[c.empresa] > 1 && c.unidade ? `${c.empresa} (${c.unidade})` : c.empresa;
  });

  // 3. valores por data
  const vistas = new Set();
  const datas = linhasData.map(({ i, d }) => {
    if (vistas.has(d.iso)) throw avisoLeitura(`A data ${d.rotulo} aparece duas vezes no relatório. Confira o arquivo.`);
    vistas.add(d.iso);
    const l = linhas[i] || [];
    const empresas = [];
    const ilegiveis = [];
    for (const c of colunas) {
      const n = lerValor(l[c.j]);
      if (Number.isNaN(n)) { ilegiveis.push(`${c.nome} ("${limpo(l[c.j])}")`); continue; }
      if (n !== 0) empresas.push({ empresa: c.nome, unidade: c.unidade, saidas: n });
    }
    if (ilegiveis.length) {
      throw avisoLeitura(`Não consegui ler o valor de ${ilegiveis.join(", ")} na data ${d.rotulo}. Confira o arquivo.`);
    }
    const total = arred2(empresas.reduce((s, e) => s + e.saidas, 0));
    return { iso: d.iso, rotulo: d.rotulo, bruto: d.bruto, empresas, total };
  });

  return { datas, colunas: colunas.length };
}

/** Converte o .xlsx (ArrayBuffer) em matriz de linhas com a biblioteca SheetJS (window.XLSX). */
export function linhasDoExcel(XLSX, buffer) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  // a aba do fluxo de caixa; se o nome mudar, usa a primeira
  const nome = wb.SheetNames.find((n) => /fluxo|pag/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[nome];
  if (!ws) throw avisoLeitura("O arquivo não tem nenhuma planilha.");
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
}
