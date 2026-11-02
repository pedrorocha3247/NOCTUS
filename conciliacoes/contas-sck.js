/**
 * Conciliações Bancárias — NOCTUS. Quais contas estão habilitadas para conciliar.
 *
 * O cadastro de contas bancárias NÃO é refeito aqui: ele continua sendo o
 * contas.js da Conferência de Pagamentos (CONTAS_EMPRESA), que já tem as 33
 * contas do grupo com banco, agência, conta, CNPJ e o código oficial da empresa.
 * Este arquivo só acrescenta o que falta para a conciliação: o CÓDIGO DA CONTA
 * NO SCK, aquele que aparece na linha "Conta: 1001283 - BANCO ITAÚ - C/C 84.283-4"
 * do relatório.
 *
 * Para habilitar uma conta nova basta acrescentar UMA LINHA em CODIGO_SCK_POR_CONTA.
 * Nada mais no módulo precisa mudar: a empresa aparece sozinha na tela, com a
 * conta nova embaixo dela.
 *
 * Primeira conta habilitada (02/11/2026, pedido do Rocha): Praia Verde, conta
 * 1001283 — Itaú, agência 0262, conta 84.283-4.
 */
import { CONTAS_EMPRESA } from "../conferencia/contas.js";

/** "<agência>/<conta>" (só dígitos) → código da conta no SCK. */
export const CODIGO_SCK_POR_CONTA = Object.freeze({
  "0262/842834": "1001283",   // Praia Verde · Itaú 341
});

const soDigitos = (v) => String(v ?? "").replace(/\D/g, "");
const chaveConta = (agencia, conta) => `${soDigitos(agencia)}/${soDigitos(conta)}`;

/** Código do banco que aparece no nome ("Itaú - 341" → "341"); vazio quando não há. */
export function codigoBanco(banco) {
  const m = /(\d{3})\s*$/.exec(String(banco || "").trim());
  return m ? m[1] : "";
}

/**
 * Contas habilitadas, já com o código do SCK e um id estável para a tela.
 * O id é o próprio código do SCK, que é único por conta.
 */
export function contasHabilitadas() {
  return CONTAS_EMPRESA
    .map((c) => {
      const codigoSck = CODIGO_SCK_POR_CONTA[chaveConta(c.agencia, c.conta)];
      return codigoSck ? { ...c, codigoSck, id: codigoSck, banco3: codigoBanco(c.banco) } : null;
    })
    .filter(Boolean);
}

/** Empresas com ao menos uma conta habilitada, em ordem alfabética. */
export function empresasHabilitadas() {
  const porEmpresa = new Map();
  for (const c of contasHabilitadas()) {
    if (!porEmpresa.has(c.codigo)) {
      porEmpresa.set(c.codigo, { codigo: c.codigo, empresa: c.empresa, aba: c.aba, contas: [] });
    }
    porEmpresa.get(c.codigo).contas.push(c);
  }
  return [...porEmpresa.values()].sort((a, b) => a.empresa.localeCompare(b.empresa, "pt-BR"));
}

/** Conta pelo id (= código do SCK). */
export function contaPorId(id) {
  return contasHabilitadas().find((c) => c.id === String(id)) || null;
}
