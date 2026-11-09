/**
 * Controle Bancário — NOCTUS. Parte 3 (03/11/2026, modelo A escolhido pelo Rocha):
 * lançamentos que justificam a diferença Fluxo × Relatório de débito.
 *
 * Regras definidas pelo Rocha (30/10/2026):
 *   - o lançamento tem SINAL: positivo explica "fluxo maior" e negativo explica
 *     "débito maior", para que a soma bata com Fluxo − Relatório;
 *   - categorias: Outros Bancos, Desconsiderados, Numerários e Compromissos
 *     (sem "pendentes");
 *   - em Outros Bancos a empresa é sempre a do lançamento e o segundo campo pede
 *     só o banco, entre as contas dessa empresa (contas.js);
 *   - excluir = exclusão lógica (nada é apagado).
 *
 * Aqui fica só a conta e a validação (sem tela, sem rede), para testar em Node.
 * Tudo é somado em centavos inteiros: nunca se compara dinheiro com ponto flutuante.
 */
import { CONTAS_EMPRESA } from "../conferencia/contas.js";

export const CATEGORIAS = Object.freeze([
  { id: "outros_bancos", nome: "Outros Bancos" },
  { id: "desconsiderados", nome: "Desconsiderados" },
  { id: "numerarios", nome: "Numerários" },
  { id: "compromissos", nome: "Compromissos" },
]);
export const NOME_CATEGORIA = Object.freeze(Object.fromEntries(CATEGORIAS.map((c) => [c.id, c.nome])));

const centavos = (v) => Math.round(Number(v || 0) * 100);
const deCentavos = (c) => c / 100;

/**
 * Lê o valor digitado no formato brasileiro ("2.000,00", "-150,5", "+5000") e também
 * "2000.50". Devolve { ok:true, valor } ou { ok:false, motivo }. Zero não é válido.
 * Ambiguidade resolvida pela regra do pt-BR: "2.000" é dois mil (ponto de milhar).
 */
export function lerValor(texto) {
  let t = String(texto ?? "").replace(/R\$/gi, "").replace(/\s+/g, "").replace(/[−‒–]/g, "-");
  if (!t) return { ok: false, motivo: "Informe o valor." };
  let sinal = 1;
  if (t[0] === "-") { sinal = -1; t = t.slice(1); } else if (t[0] === "+") t = t.slice(1);
  let num;
  if (t.includes(",")) {
    if (!/^\d{1,3}(\.\d{3})*,\d{1,2}$|^\d+,\d{1,2}$/.test(t)) return { ok: false, motivo: "Valor inválido. Use o formato 2.000,00." };
    num = t.replace(/\./g, "").replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(t)) {
    num = t.replace(/\./g, "");
  } else if (/^\d+(\.\d{1,2})?$/.test(t)) {
    num = t;
  } else {
    return { ok: false, motivo: "Valor inválido. Use o formato 2.000,00." };
  }
  const c = Math.round(parseFloat(num) * 100) * sinal;
  if (!Number.isFinite(c)) return { ok: false, motivo: "Valor inválido. Use o formato 2.000,00." };
  if (c === 0) return { ok: false, motivo: "O valor não pode ser zero." };
  if (Math.abs(c) >= 100000000000) return { ok: false, motivo: "O valor é grande demais. Confira os zeros." };
  return { ok: true, valor: deCentavos(c) };
}

/**
 * Contas de OUTRA empresa que também aparecem em "Outros Bancos" (pedido do Rocha, 09/11/2026): a SLIM (16) pode
 * escolher a conta Banco do Brasil Ag. 2155 C/C 9868 -X e a Bradesco Ag. 3395 C/C 37629-9, que no cadastro são da
 * MOMENTUM (15). Ficam AQUI e não em contas.js de propósito: a Conferência usa o número da conta como chave única
 * para descobrir a empresa dona de cada solicitação, e a mesma conta em duas empresas deixaria isso ambíguo.
 * O texto é igual ao do cadastro da Momentum, para o rótulo ser o mesmo nas duas empresas.
 */
const CONTAS_EXTRAS_OUTROS_BANCOS = Object.freeze({
  "16": Object.freeze([
    { banco: "Banco do Brasil - 01", agencia: "2155", conta: "9868 -X" },
    { banco: "Bradesco - 237", agencia: "3395", conta: "37629-9" },
  ]),
});

/** Contas bancárias de uma empresa (pelo código oficial), no formato do cadastro da Conferência, mais as extras acima. */
export function contasDaEmpresa(codigo) {
  const proprias = CONTAS_EMPRESA.filter((c) => String(c.codigo) === String(codigo));
  const extras = CONTAS_EXTRAS_OUTROS_BANCOS[String(codigo)] || [];
  const vistos = new Set();
  return [...proprias, ...extras].map((c) => ({ rotulo: `${c.banco} · Ag. ${c.agencia} · C/C ${c.conta}` }))
    .filter((c) => !vistos.has(c.rotulo) && vistos.add(c.rotulo));   // sem repetir se um dia a conta entrar no cadastro da empresa
}

/**
 * Valida o formulário "Novo lançamento". `empresas` = códigos aceitos (empresas da tabela resumo da data).
 * Devolve { ok, erros:{campo:mensagem}, lancamento }.
 */
export function validarLancamento({ categoria, valorTexto, empresaCodigo, empresas, banco, descricao }) {
  const erros = {};
  if (!NOME_CATEGORIA[categoria]) erros.categoria = "Escolha a categoria.";
  const v = lerValor(valorTexto);
  if (!v.ok) erros.valor = v.motivo;
  if (!empresaCodigo || !(empresas || []).some((e) => String(e.codigo) === String(empresaCodigo))) erros.empresa = "Escolha a empresa.";
  let bancoFinal = null;
  if (categoria === "outros_bancos") {
    const contas = empresaCodigo ? contasDaEmpresa(empresaCodigo) : [];
    if (empresaCodigo && !contas.length) erros.banco = "Esta empresa não tem conta bancária cadastrada no NOCTUS.";
    else if (!contas.some((c) => c.rotulo === banco)) erros.banco = "Escolha o banco.";
    else bancoFinal = banco;
  }
  const d = String(descricao ?? "").replace(/\s+/g, " ").trim();
  if (d.length < 3) erros.descricao = "Descreva o lançamento (mínimo de 3 letras).";
  else if (d.length > 300) erros.descricao = "A descrição passa de 300 caracteres.";
  if (Object.keys(erros).length) return { ok: false, erros };
  const empresa = empresas.find((e) => String(e.codigo) === String(empresaCodigo));
  return {
    ok: true, erros: {},
    lancamento: { categoria, valor: v.valor, empresa_codigo: String(empresaCodigo), empresa: empresa.empresa,
                  banco: bancoFinal, descricao: d },
  };
}

/** Soma por categoria e total (só lançamentos ativos). */
export function somarPorCategoria(lancs) {
  const por = Object.fromEntries(CATEGORIAS.map((c) => [c.id, 0]));
  for (const l of lancs || []) if (l.categoria in por) por[l.categoria] += centavos(l.valor);
  const total = Object.values(por).reduce((s, c) => s + c, 0);
  return { porCategoria: Object.fromEntries(Object.entries(por).map(([k, c]) => [k, deCentavos(c)])), total: deCentavos(total) };
}

/**
 * Falta justificar = diferença da tabela resumo − total lançado. O sinal do resultado segue a convenção
 * das cores da tabela resumo: positivo = fluxo maior ainda sem explicação (laranja); negativo = débito
 * maior ainda sem explicação (vermelho).
 * situacao: "semRelatorio" (diferença desconhecida), "justificado" (zero),
 *           "falta" (fluxo ainda maior que o explicado) ou "excesso" (lançou além da diferença).
 */
export function faltaJustificar(diferenca, total) {
  if (diferenca == null) return { situacao: "semRelatorio", falta: null };
  const f = centavos(diferenca) - centavos(total);
  return { situacao: f === 0 ? "justificado" : f > 0 ? "falta" : "excesso", falta: deCentavos(f) };
}

/** Agrupa os lançamentos por código de empresa (ordem de criação). */
export function porEmpresa(lancs) {
  const m = new Map();
  for (const l of lancs || []) {
    const k = String(l.empresa_codigo);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(l);
  }
  return m;
}

/** Total lançado de uma empresa e quanto falta para a diferença dela. */
export function situacaoDaEmpresa(diferenca, lancsEmpresa) {
  const { total } = somarPorCategoria(lancsEmpresa);
  return { lancado: total, ...faltaJustificar(diferenca, total) };
}
