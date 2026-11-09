/**
 * Controle Bancário — NOCTUS. Aba "Saldo do Dia" (09/11/2026, pedido do Rocha, modelo M4 escolhido).
 *
 * Mostra SEMPRE as mesmas empresas e bancos da planilha de referência (fluxo_bancario.xlsx, 26 colunas em 13
 * empresas), tenham ou não lançamento no Fluxo de Caixa daquele dia, com um campo manual de saldo atual ao lado
 * de cada banco. O saldo é guardado por USUÁRIO + DATA + LINHA (controle_bancario_saldo_noctus), nunca ligado ao
 * último fluxo importado: dia novo = campos vazios; voltar a uma data = valores recuperados.
 *
 * O id da linha (L01..L26) é a chave estável gravada no banco. NÃO renumerar: conta nova entra como L27 no fim.
 * A lista vem da planilha do Rocha, não do cadastro de contas da Conferência (contas.js), porque a planilha tem
 * linhas sem conta (M3, MMH, ABRASMA), contas incompletas (MEI 881 e 525) e bancos antigos (Banespa, Real).
 *
 * Só a conta e a leitura do valor ficam aqui (sem tela, sem rede), para testar em Node.
 * Tudo é somado em centavos inteiros: dinheiro nunca é comparado em ponto flutuante.
 */

/** Ordem e nomes da planilha de referência. Banco padronizado (a planilha escreve "itau", "BRASIL", "BB"). */
export const LINHAS_SALDO = Object.freeze([
  ["L01", "KASIL", "Itaú", "58.073-1"],
  ["L02", "IRM", "Itaú", "58.703-3"],
  ["L03", "IRM", "BB", "13341-8"],
  ["L04", "PRAIA VERDE", "Bradesco", "3890-3"],
  ["L05", "PRAIA VERDE", "Itaú", "84283-4"],
  ["L06", "SLIM", "Itaú", "41382-6"],
  ["L07", "RVM", "Itaú", "58697-7"],
  ["L08", "M5", "Itaú", "14982-8"],
  ["L09", "POSTO", "Itaú", "14227-8"],
  ["L10", "POSTO", "BB", "12174-6"],
  ["L11", "PICK MONEY", "Santander", "130135988"],
  ["L12", "PICK MONEY", "Itaú", "41982-5"],
  ["L13", "MEI", "Itaú", "62222-8"],
  ["L14", "MEI", "Bradesco", "37.629-9"],
  ["L15", "MEI", "Bradesco", "881"],
  ["L16", "MEI", "Bradesco", "525"],
  ["L17", "MEI", "Banespa", "13.000.050-9"],
  ["L18", "MEI", "Real", "13.000.488-8"],
  ["L19", "MEI", "Santander", "13.000008-3"],
  ["L20", "MEI", "BB", "9868-X"],
  ["L21", "MEI", "CEF", "003.160-6"],
  ["L22", "REALIZA", "Itaú", "46780-8"],
  ["L23", "M3", "Itaú", ""],
  ["L24", "MMH", "Itaú", ""],
  ["L25", "ABRASMA", "Itaú", ""],
  ["L26", "PRAIA VERDE", "Daycoval", "703.673-8"],
].map(([id, empresa, banco, conta]) => Object.freeze({ id, empresa, banco, conta })));

export const IDS_SALDO = Object.freeze(LINHAS_SALDO.map((l) => l.id));
export const EMPRESAS_SALDO = Object.freeze([...new Set(LINHAS_SALDO.map((l) => l.empresa))]);

const MAX_CENTAVOS = 100000000000000;   // 1 trilhão em centavos: o banco guarda numeric(14,2), até 999.999.999.999,99

/**
 * Lê o saldo digitado. Vazio é válido (campo sem saldo): { ok:true, valor:null }.
 * Aceita "1.234,56", "-150,5", "−150,5", "R$ 10", "2000.50" e "0". Zero é válido (conta zerada). Negativo é válido (saldo devedor).
 * "2.000" é dois mil (ponto de milhar, regra do pt-BR). Mais de 2 casas decimais é recusado, não arredondado.
 */
export function lerSaldo(texto) {
  let t = String(texto ?? "").replace(/R\$/gi, "").replace(/\s+/g, "").replace(/[−‒–]/g, "-");
  if (!t) return { ok: true, valor: null };
  let sinal = 1;
  if (t[0] === "-") { sinal = -1; t = t.slice(1); } else if (t[0] === "+") t = t.slice(1);
  let num;
  if (t.includes(",")) {
    if (!/^\d{1,3}(\.\d{3})*,\d{1,2}$|^\d+,\d{1,2}$/.test(t)) return { ok: false, motivo: "Valor inválido. Use o formato 1.234,56." };
    num = t.replace(/\./g, "").replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(t)) {
    num = t.replace(/\./g, "");
  } else if (/^\d+(\.\d{1,2})?$/.test(t)) {
    num = t;
  } else {
    return { ok: false, motivo: "Valor inválido. Use o formato 1.234,56." };
  }
  const c = Math.round(parseFloat(num) * 100);
  if (!Number.isFinite(c)) return { ok: false, motivo: "Valor inválido. Use o formato 1.234,56." };
  if (c >= MAX_CENTAVOS) return { ok: false, motivo: "O valor é grande demais. Confira os zeros." };
  return { ok: true, valor: (c * sinal) / 100 || 0 };   // "|| 0" evita "-0"
}

/** "1.234,56" (negativo: "-1.234,56"). Vazio (null/undefined) vira "". */
export function formatarSaldo(v) {
  if (v == null || v === "") return "";
  return Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** Soma em centavos de uma lista de valores (null/undefined ficam de fora). */
export function somarSaldos(valores) {
  let c = 0, preenchidos = 0;
  for (const v of valores || []) {
    if (v == null) continue;
    c += Math.round(Number(v) * 100); preenchidos++;
  }
  return { total: c / 100, preenchidos };
}

/** A linha casa com a busca? (empresa, banco ou conta, sem acento nem caixa) */
export function linhaCasaComBusca(linha, busca) {
  const norm = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const q = norm(busca).trim();
  if (!q) return true;
  return norm(`${linha.empresa} ${linha.banco} ${linha.conta}`).includes(q);
}
