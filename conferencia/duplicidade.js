/**
 * Duplicidade de pagamento dentro de uma mesma conferência (04/11/2026, pedido do Rocha).
 *
 * Por que existe: o relatório do SCK pode trazer DUAS solicitações idênticas (mesmo favorecido, mesmo valor,
 * mesma destinação) que só diferem no número da solicitação (S.N). Até 04/11 o sistema só apontava duplicidade
 * por NOTA FISCAL (parser.js, marcarRepetidos), então um pagamento sem NF, como "AJUDA DE CUSTO", passava batido.
 * Teste do Rocha: duas solicitações de R$ 500,00 para o mesmo favorecido, com a mesma destinação, sem nenhum apontamento.
 *
 * Duas regras, ambas dentro da MESMA empresa (pagamentos de empresas diferentes não são duplicidade um do outro):
 *
 *   1. FORTE  — mesmo favorecido + mesmo valor + mesma destinação:
 *      "Possível pagamento em duplicidade: mesmo favorecido, valor e destinação da(s) solicitação(ões) ..."
 *
 *   2. ATENÇÃO — mesmo favorecido + mesmo valor, destinação DIFERENTE:
 *      "Mesmo favorecido e valor da(s) solicitação(ões) ..., com destinação diferente: confira se não é pagamento em duplicidade"
 *      (uma duplicidade pode vir com a destinação escrita um pouco diferente; no relatório real de 04/11, 3 pares em
 *      120 solicitações caíram aqui, todos aparentemente legítimos: o conferente os remove no ✕).
 *
 * Como se identifica cada coisa (tudo objetivo, nada por semelhança):
 *   - favorecido: os dígitos do CPF/CNPJ quando há 11 ou mais; sem documento, o nome normalizado (sem acento, caixa
 *     e pontuação). Sem documento e sem nome, a solicitação não é comparada.
 *   - valor: em centavos inteiros (nunca comparar decimal em ponto flutuante). Valor zero ou ilegível não é comparado.
 *   - destinação: texto normalizado (sem acento, maiúsculas, espaços e pontuação colapsados).
 *   - a forma de pagamento (TED, PIX, boleto) NÃO entra: o mesmo valor ao mesmo favorecido com a mesma destinação é
 *     suspeito mesmo que uma tenha saído por TED e a outra por PIX.
 *
 * É sinal para o conferente olhar, não veredito: nunca recusa, nunca muda status. Função pura: não lê nem grava nada.
 */

/** Sem acento, maiúsculas, tudo que não é letra ou número vira um espaço só. */
export const normalizarTexto = (t) =>
  String(t ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();

const soDigitos = (t) => String(t ?? "").replace(/\D/g, "");

/** Chave do favorecido: documento (11+ dígitos) ou, sem ele, o nome. "" = não dá para comparar. */
export function chaveFavorecido(s) {
  const d = soDigitos(s.cpfCnpj);
  if (d.length >= 11) return "D" + d;
  const nome = normalizarTexto(s.favorecido);
  return nome ? "N" + nome : "";
}

/** Valor em centavos inteiros, ou null se ilegível/zero. */
export function centavos(valor) {
  const n = typeof valor === "number" ? valor : parseFloat(String(valor ?? "").replace(/\./g, "").replace(",", "."));
  if (!Number.isFinite(n) || n === 0) return null;
  return Math.round(n * 100);
}

const ordemSn = (a, b) => (String(a).length - String(b).length) || String(a).localeCompare(String(b));
const lista = (sns) => [...sns].sort(ordemSn).join(", ");

export const textoDuplicidadeForte = (sns) =>
  `Possível pagamento em duplicidade: mesmo favorecido, valor e destinação ${sns.length > 1 ? "das solicitações" : "da solicitação"} ${lista(sns)}`;
export const textoDuplicidadeAtencao = (sns) =>
  `Mesmo favorecido e valor ${sns.length > 1 ? "das solicitações" : "da solicitação"} ${lista(sns)}, com destinação diferente: confira se não é pagamento em duplicidade`;

/**
 * Devolve um Map: S.N -> lista de apontamentos de duplicidade (vazio = nenhum).
 * `solicitacoes` pode misturar empresas: o agrupamento já separa por empresa.
 */
export function detectarDuplicidades(solicitacoes) {
  const grupos = new Map();
  for (const s of solicitacoes || []) {
    const fav = chaveFavorecido(s), val = centavos(s.valor);
    if (!fav || val === null || s.sn == null || s.sn === "") continue;
    const k = `${s.empresaCodigo || s.empresa || ""}|${fav}|${val}`;
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(s);
  }
  const out = new Map();
  for (const membros of grupos.values()) {
    if (membros.length < 2) continue;
    for (const s of membros) {
      const dest = normalizarTexto(s.destinacao);
      const mesmos = [], diferentes = [];
      for (const o of membros) {
        if (o === s || String(o.sn) === String(s.sn)) continue;
        (normalizarTexto(o.destinacao) === dest ? mesmos : diferentes).push(o.sn);
      }
      const msgs = [];
      if (mesmos.length) msgs.push(textoDuplicidadeForte(mesmos));
      if (diferentes.length) msgs.push(textoDuplicidadeAtencao(diferentes));
      if (msgs.length) out.set(s.sn, msgs);
    }
  }
  return out;
}

/** Grava em cada solicitação o campo `alertasDuplicidade` (lista; [] quando não há). Devolve quantas ficaram marcadas. */
export function aplicarDuplicidades(solicitacoes) {
  const achadas = detectarDuplicidades(solicitacoes);
  let marcadas = 0;
  for (const s of solicitacoes || []) {
    s.alertasDuplicidade = achadas.get(s.sn) || [];
    if (s.alertasDuplicidade.length) marcadas++;
  }
  return marcadas;
}
