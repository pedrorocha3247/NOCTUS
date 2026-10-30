/**
 * Controle Bancário — NOCTUS. Parte 2 (30/10/2026): relatório de débito em conta
 * (SCK, Excel) e tabela resumo Fluxo × Relatório.
 *
 * O relatório de débito é o MESMO da Conferência de Pagamentos, então a leitura
 * reaproveita o parseRelatorioExcel de lá: ele separa as solicitações por empresa
 * pela conta bancária (contas.js) e devolve, em validacaoPorEmpresa, o total
 * IMPRESSO no relatório ("Total ==>") e a soma das solicitações lidas de cada empresa.
 * Aqui só interessa o total por empresa.
 *
 * A ligação com o fluxo de caixa é pelo CÓDIGO oficial da empresa (confirmado pelo
 * Rocha em 30/10/2026), nunca por semelhança de nome.
 */
import { parseRelatorioExcel } from "../conferencia/parser.js";
import { CONTAS_EMPRESA } from "../conferencia/contas.js";

/** Nome da empresa no fluxo de caixa → código oficial (o mesmo de contas.js e poderes.js). */
export const CODIGO_FLUXO = Object.freeze({
  "KASIL": "30", "IRM": "90", "PRAIA VERDE": "13", "SLIM": "16", "RVM": "40", "M5": "70",
  "POSTO SB": "91", "PICK MONEY": "95", "MOMENTUM": "15", "REALIZA": "28",
  "MMH HOLDING": "96", "M3 ASSETS": "26", "ABRASMA": "27",
});
const NOME_FLUXO_POR_CODIGO = Object.fromEntries(Object.entries(CODIGO_FLUXO).map(([n, c]) => [c, n]));
const CODIGO_POR_NOME_CURTO = new Map(CONTAS_EMPRESA.map((c) => [c.empresa, c.codigo]));

const r2 = (v) => Math.round(Number(v || 0) * 100) / 100;
function aviso(msg) { const e = new Error(msg); e.aviso = true; return e; }

/** É o relatório de débito em conta (e não o fluxo de caixa)? Olha o título do SCK nas primeiras linhas. */
export function ehRelatorioDebito(linhas) {
  for (const linha of (linhas || []).slice(0, 25)) {
    for (const v of linha || []) {
      if (typeof v === "string" && /SOLICITA[ÇC][ÕO]ES\s+DE\s+PAGAMENTO/i.test(v)) return true;
    }
  }
  return false;
}

/**
 * Lê o relatório de débito em conta e devolve o total de cada empresa.
 * { iso, rotulo, empresas: [{ codigo, empresa, total, qtd, soma, confere }] }
 * Conta bancária fora de contas.js vira um item com codigo null ("CONTA NÃO CADASTRADA"):
 * o valor nunca é somado em silêncio a uma empresa.
 */
export function lerDebito(XLSX, buffer) {
  const { meta, validacaoPorEmpresa } = parseRelatorioExcel(buffer, XLSX);
  const ini = meta?.dataInicio, fim = meta?.dataFim;
  if (!ini) throw aviso("Não encontrei a data do relatório de débito em conta. Confira se é o relatório de Solicitações de Pagamentos exportado em Excel.");
  if (fim && fim !== ini) throw aviso(`O relatório de débito cobre um período (${ini} até ${fim}). Importe o relatório de um dia só.`);
  const [d, m, a] = ini.split("/");
  const iso = `${a}-${m}-${d}`;

  const empresas = [];
  let semConta = null;
  for (const [nomeCurto, v] of Object.entries(validacaoPorEmpresa || {})) {
    // sem total impresso lido, fica a soma das solicitações, marcada como "não confere"
    const total = v.totaisLidos > 0 ? r2(v.valorRelatorio) : r2(v.valorExtraido);
    const item = { total, qtd: v.totaisLidos > 0 ? v.qtdRelatorio : v.qtdExtraida, soma: r2(v.valorExtraido),
                   confere: v.totaisLidos > 0 && v.confere === true };
    if (!nomeCurto) { semConta = item; continue; }
    const codigo = CODIGO_POR_NOME_CURTO.get(nomeCurto) || null;
    empresas.push({ codigo, empresa: NOME_FLUXO_POR_CODIGO[codigo] || nomeCurto.toUpperCase(), ...item });
  }
  if (semConta) empresas.push({ codigo: null, empresa: "CONTA NÃO CADASTRADA", ...semConta });
  if (!empresas.length) throw aviso("Não encontrei nenhuma solicitação no relatório de débito em conta. Confira o arquivo.");
  return { iso, rotulo: ini, empresas, total: r2(empresas.reduce((s, e) => s + e.total, 0)) };
}

/**
 * Monta a tabela resumo de uma data: colunas = empresas do fluxo (mesma ordem) +
 * empresas que só aparecem no débito (marcadas "sem fluxo"). Diferença = Fluxo − Relatório.
 * debito == null → ainda não importado: Relatório e Diferença ficam null.
 */
export function montarResumo(fluxoEmpresas, debito) {
  const cols = [];
  const porChave = new Map();
  for (const e of fluxoEmpresas || []) {
    const base = String(e.empresa || "").replace(/\s*\(.*\)\s*$/, "").trim();
    const codigo = CODIGO_FLUXO[base] || null;
    const chave = codigo || `nome:${e.empresa}`;
    let c = porChave.get(chave);
    if (!c) {
      c = { codigo, empresa: codigo ? base : String(e.empresa || ""), fluxo: 0, relatorio: null, semFluxo: false, semCodigo: !codigo };
      porChave.set(chave, c);
      cols.push(c);
    }
    c.fluxo = r2(c.fluxo + Number(e.saidas || 0));
  }

  const avisos = [];
  if (debito) {
    for (const c of cols) if (!c.semCodigo) c.relatorio = 0;
    for (const d of debito.empresas || []) {
      if (!d.codigo) { avisos.push({ tipo: "semConta", total: r2(d.total), qtd: d.qtd }); continue; }
      let c = porChave.get(d.codigo);
      if (!c) {
        if (Math.abs(Number(d.total || 0)) < 0.005) continue;
        c = { codigo: d.codigo, empresa: d.empresa, fluxo: 0, relatorio: 0, semFluxo: true, semCodigo: false };
        porChave.set(d.codigo, c);
        cols.push(c);
        avisos.push({ tipo: "semFluxo", empresa: d.empresa, total: r2(d.total) });
      }
      c.relatorio = r2(c.relatorio + Number(d.total || 0));
      if (d.confere === false) avisos.push({ tipo: "naoConfere", empresa: d.empresa, total: r2(d.total), soma: r2(d.soma) });
    }
    for (const c of cols) if (c.semCodigo) avisos.push({ tipo: "semCodigo", empresa: c.empresa });
  }
  for (const c of cols) c.diferenca = c.relatorio == null ? null : r2(c.fluxo - c.relatorio);

  const fluxo = r2(cols.reduce((s, c) => s + c.fluxo, 0));
  const relatorio = debito ? r2(cols.reduce((s, c) => s + (c.relatorio || 0), 0)) : null;
  return {
    cols, avisos,
    total: { fluxo, relatorio, diferenca: relatorio == null ? null : r2(fluxo - relatorio) },
    divergencias: cols.filter((c) => c.diferenca != null && c.diferenca !== 0).length,
  };
}
