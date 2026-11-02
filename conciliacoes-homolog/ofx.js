/**
 * Conciliações Bancárias — NOCTUS. Leitura do extrato bancário em .OFX.
 *
 * Parte 1 (02/11/2026). O OFX do Itaú vem em SGML (versão 102), não em XML:
 * as etiquetas de valor NÃO são fechadas (`<TRNAMT>12864.26` e quebra de linha),
 * então não dá para usar um leitor de XML. A leitura é por etiqueta, linha a linha.
 *
 * Três armadilhas do arquivo real de 30/10/2026 (Itaú, ag. 0262 c/c 84.283-4),
 * todas tratadas aqui:
 *
 *   1. SEM SINAL. Os 6 lançamentos vieram com TRNTYPE CREDIT e TRNAMT positivo,
 *      inclusive os que não são entrada. O tipo do arquivo, sozinho, não é
 *      confiável: a direção sai do SINAL do valor quando ele existe (TRNAMT
 *      negativo = saída) e, quando todos vêm positivos, do TRNTYPE. Se o arquivo
 *      inteiro vier positivo e sem nenhum DEBIT, a leitura AVISA que o extrato
 *      não traz saídas, em vez de afirmar que não houve nenhuma.
 *   2. LINHAS DE SALDO COMO LANÇAMENTO. "SALDO ANTERIOR" e "SALDO TOTAL
 *      DISPONÍVEL DIA" estão dentro de <BANKTRANLIST> como se fossem movimento.
 *      Somá-las dobraria o extrato. Elas são separadas e o "SALDO TOTAL
 *      DISPONÍVEL DIA" vira o saldo do dia (escolha do Rocha, 02/11/2026).
 *   3. FUSO COM RÓTULO ERRADO. As datas vêm "20261030100000[-03:EST]": o
 *      deslocamento -03 está certo, mas o rótulo EST é -05. Um leitor que
 *      confie no rótulo VIRA O DIA. Aqui a data é lida pelos 8 primeiros
 *      dígitos (AAAAMMDD), sem nenhuma conversão de fuso.
 *
 * O LEDGERBAL não é usado como saldo do dia: no arquivo real ele tem DTASOF de
 * 02/11, três dias depois do extrato. Ele é devolvido à parte, com a data dele.
 */

export function avisoLeitura(msg) {
  const e = new Error(msg);
  e.aviso = true;
  return e;
}

const limpo = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
const r2 = (n) => Math.round(n * 100) / 100;
const pad2 = (n) => String(n).padStart(2, "0");

/** Linhas de saldo que o conversor do banco imprime dentro da lista de lançamentos. */
const EH_LINHA_SALDO = /^SALDO\b/i;
const EH_SALDO_DO_DIA = /^SALDO\s+(TOTAL\s+)?DISPON[IÍ]VEL(\s+DIA)?/i;
const EH_SALDO_ANTERIOR = /^SALDO\s+ANTERIOR/i;

/**
 * Texto do arquivo a partir dos bytes. O extrato real do Itaú declara
 * "CHARSET:1252" no cabeçalho mas grava o conteúdo em UTF-8 — seguir o
 * cabeçalho quebra os acentos e faz "SALDO TOTAL DISPONÍVEL DIA" deixar de ser
 * reconhecido como linha de saldo. Por isso tenta UTF-8 primeiro, no modo que
 * recusa byte inválido, e só então cai para windows-1252.
 */
export function textoDoOfx(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    try { return new TextDecoder("windows-1252").decode(bytes); }
    catch { return new TextDecoder("latin1").decode(bytes); }
  }
}

/** O texto é um arquivo OFX? */
export function ehOfx(texto) {
  const t = String(texto || "").slice(0, 4000).toUpperCase();
  return t.includes("<OFX>") || t.includes("OFXHEADER");
}

/** Valor de uma etiqueta SGML não fechada: <TRNAMT>12864.26 → "12864.26". */
function etiqueta(bloco, nome) {
  const m = new RegExp(`<${nome}>([^\\r\\n<]*)`, "i").exec(bloco);
  return m ? m[1].trim() : null;
}

/**
 * Data do OFX. Lê só AAAAMMDD, ignorando hora e fuso de propósito (ver o
 * comentário do topo: o rótulo do fuso vem errado e viraria o dia).
 */
export function lerDataOfx(bruto) {
  const t = limpo(bruto);
  const m = /^(\d{4})(\d{2})(\d{2})/.exec(t);
  if (!m) return null;
  const ano = Number(m[1]), mes = Number(m[2]), dia = Number(m[3]);
  const d = new Date(ano, mes - 1, dia);
  if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) return null;
  return { iso: `${m[1]}-${m[2]}-${m[3]}`, rotulo: `${pad2(dia)}/${pad2(mes)}/${ano}` };
}

/** "0262842834" na conta do Itaú é agência 0262 + conta 84283-4 (sem o traço). */
export function separarContaItau(acctid, bankid) {
  const so = String(acctid || "").replace(/\D/g, "");
  if (String(bankid || "").replace(/\D/g, "").replace(/^0+/, "") !== "341" || so.length < 8) return null;
  const agencia = so.slice(0, 4);
  const resto = so.slice(4);
  return { agencia, numero: `${resto.slice(0, -1)}-${resto.slice(-1)}` };
}

const soDigitos = (v) => String(v ?? "").replace(/\D/g, "");
/** Compara dois números de conta ignorando pontos, traços e zeros à esquerda. */
export function mesmaConta(a, b) {
  const x = soDigitos(a).replace(/^0+/, ""), y = soDigitos(b).replace(/^0+/, "");
  return !!x && x === y;
}

/**
 * Lê o extrato e devolve:
 *   { conta:{bankid,acctid,agencia,numero,tipo}, periodo:{inicio,fim}, moeda,
 *     lancamentos:[{data,tipo,valor,memo,fitid,direcao}],
 *     saldoAnterior, saldoDoDia, saldoLivro:{valor,data},
 *     entradas, saidas, alertas:[] }
 *
 * `entradas`/`saidas` somam só os lançamentos de verdade (as linhas de saldo
 * ficam de fora) e, quando `dataAlvo` é informada, só os daquele dia.
 */
export function lerExtratoOfx(texto, { dataAlvo = null } = {}) {
  if (!ehOfx(texto)) {
    throw avisoLeitura("Este arquivo não é um extrato OFX. Baixe o extrato do banco no formato .OFX e tente de novo.");
  }
  const alertas = [];
  const conteudo = String(texto);

  const bankid = etiqueta(conteudo, "BANKID") || "";
  const acctid = etiqueta(conteudo, "ACCTID") || "";
  if (!acctid) throw avisoLeitura("O extrato não informa o número da conta (etiqueta ACCTID). Confira o arquivo.");
  const itau = separarContaItau(acctid, bankid);
  const conta = { bankid, acctid, tipo: etiqueta(conteudo, "ACCTTYPE") || "",
                  agencia: itau?.agencia || "", numero: itau?.numero || "" };

  const periodo = { inicio: lerDataOfx(etiqueta(conteudo, "DTSTART")), fim: lerDataOfx(etiqueta(conteudo, "DTEND")) };
  const moeda = etiqueta(conteudo, "CURDEF") || "BRL";

  /* ---------- lançamentos ---------- */
  const blocos = conteudo.match(/<STMTTRN>[\s\S]*?<\/STMTTRN>/gi) || [];
  if (!blocos.length) {
    throw avisoLeitura("O extrato não tem nenhum lançamento (STMTTRN). Confira se baixou o período certo.");
  }
  const todos = [];
  const ilegiveis = [];
  for (const b of blocos) {
    const bruto = etiqueta(b, "TRNAMT");
    const n = Number(String(bruto).replace(",", "."));
    if (!Number.isFinite(n)) { ilegiveis.push(limpo(bruto)); continue; }
    const data = lerDataOfx(etiqueta(b, "DTPOSTED"));
    todos.push({
      data,
      tipo: (etiqueta(b, "TRNTYPE") || "").toUpperCase(),
      valor: r2(n),
      memo: limpo(etiqueta(b, "MEMO") || etiqueta(b, "NAME") || ""),
      fitid: limpo(etiqueta(b, "FITID") || ""),
    });
  }
  if (ilegiveis.length) {
    throw avisoLeitura(`Não consegui ler o valor de ${ilegiveis.length} lançamento(s) do extrato ("${ilegiveis.slice(0, 3).join('", "')}"). Confira o arquivo.`);
  }
  const semData = todos.filter((t) => !t.data).length;
  if (semData) alertas.push(`${semData} lançamento(s) do extrato estão sem data e ficaram de fora da soma.`);

  /* ---------- linhas de saldo ---------- */
  const linhasSaldo = todos.filter((t) => EH_LINHA_SALDO.test(t.memo));
  const anterior = linhasSaldo.find((t) => EH_SALDO_ANTERIOR.test(t.memo));
  const doDia = linhasSaldo.find((t) => EH_SALDO_DO_DIA.test(t.memo));
  const saldoAnterior = anterior ? { valor: anterior.valor, data: anterior.data } : null;

  /* ---------- direção ---------- */
  // Com sinal no arquivo, o sinal manda. Sem nenhum negativo e sem nenhum DEBIT,
  // o extrato não distingue entrada de saída: usa o TRNTYPE e avisa.
  const movimentos = todos.filter((t) => !EH_LINHA_SALDO.test(t.memo));
  const temNegativo = movimentos.some((t) => t.valor < 0);
  const temDebit = movimentos.some((t) => t.tipo === "DEBIT");
  for (const m of movimentos) {
    m.direcao = m.valor < 0 || m.tipo === "DEBIT" ? "saida" : "entrada";
    m.valor = Math.abs(m.valor);
  }
  if (!temNegativo && !temDebit) {
    alertas.push("O extrato veio sem nenhum lançamento negativo nem de débito: o arquivo não distingue entrada de saída, então as saídas ficam sem conferência. Confira o extrato no banco antes de concluir a conciliação.");
  }

  /* ---------- totais do dia ---------- */
  const doDiaAlvo = (t) => !dataAlvo || (t.data && t.data.iso === dataAlvo);
  const usados = movimentos.filter((t) => t.data && doDiaAlvo(t));
  const foraDoPeriodo = movimentos.filter((t) => t.data && !doDiaAlvo(t));
  if (dataAlvo && foraDoPeriodo.length) {
    alertas.push(`${foraDoPeriodo.length} lançamento(s) do extrato são de outra data e não entraram na soma de ${rotulo(dataAlvo)}.`);
  }
  const entradas = r2(usados.filter((t) => t.direcao === "entrada").reduce((s, t) => s + t.valor, 0));
  const saidas = r2(usados.filter((t) => t.direcao === "saida").reduce((s, t) => s + t.valor, 0));

  /* ---------- saldo do dia ---------- */
  let saldoDoDia = null;
  if (doDia) {
    saldoDoDia = { valor: doDia.valor, origem: "extrato", data: doDia.data };
  } else {
    const base = saldoAnterior ? saldoAnterior.valor : 0;
    saldoDoDia = { valor: r2(base + entradas - saidas), origem: "calculado", data: dataAlvo ? lerDataOfx(dataAlvo.replace(/-/g, "")) : periodo.fim };
    alertas.push(saldoAnterior
      ? "O extrato não traz a linha de saldo do dia: o saldo foi calculado somando o saldo anterior com os lançamentos."
      : "O extrato não traz saldo anterior nem saldo do dia: o saldo foi calculado só com os lançamentos do período.");
  }

  const balamt = etiqueta(conteudo, "BALAMT");
  const saldoLivro = balamt != null && Number.isFinite(Number(balamt))
    ? { valor: r2(Number(balamt)), data: lerDataOfx(etiqueta(conteudo, "DTASOF")) }
    : null;
  if (saldoLivro?.data && dataAlvo && saldoLivro.data.iso !== dataAlvo) {
    alertas.push(`O saldo oficial do arquivo (LEDGERBAL) é de ${saldoLivro.data.rotulo}, e não de ${rotulo(dataAlvo)}; por isso ele não foi usado na comparação.`);
  }

  return { conta, periodo, moeda, lancamentos: movimentos, linhasSaldo,
           saldoAnterior, saldoDoDia, saldoLivro, entradas, saidas, alertas };
}

const rotulo = (iso) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`;
