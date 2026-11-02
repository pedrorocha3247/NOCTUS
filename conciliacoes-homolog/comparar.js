/**
 * Conciliações Bancárias — NOCTUS. Conferência da conta e comparação SCK × extrato.
 *
 * Parte 1 (02/11/2026). Decisões do Rocha no mesmo dia:
 *   - saldo do banco = a linha "SALDO TOTAL DISPONÍVEL DIA" do extrato (não o
 *     LEDGERBAL, que no arquivo real é de outra data);
 *   - saldo do SCK = o "Saldo Contábil" do Resumo da Conciliação, e a
 *     "Diferença Geral da Conciliação" do próprio SCK aparece à parte, sem
 *     entrar nessa conta;
 *   - arquivo de outra conta é RECUSADO (não é só aviso): conciliar a conta
 *     errada produziria um resultado que parece válido e não é.
 *
 * Diferença = BANCO − SCK, em todas as linhas. Positiva quer dizer que o banco
 * tem mais do que o SCK registrou, ou seja, falta lançar no SCK.
 */

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
const zero = (n) => Math.abs(Number(n || 0)) < 0.005;
function aviso(msg) { const e = new Error(msg); e.aviso = true; return e; }

const soDigitos = (v) => String(v ?? "").replace(/\D/g, "");
const igualConta = (a, b) => {
  const x = soDigitos(a).replace(/^0+/, ""), y = soDigitos(b).replace(/^0+/, "");
  return !!x && x === y;
};

/**
 * A conta escolhida na tela, o relatório do SCK e o extrato falam da mesma conta?
 * `conta` é a linha de contas.js com o codigoSck acrescentado.
 * Lança aviso (arquivo recusado) quando não bate.
 */
export function conferirConta(conta, { sck = null, ofx = null } = {}) {
  if (!conta) throw aviso("Escolha a empresa e a conta antes de importar os arquivos.");

  if (sck?.conta) {
    if (conta.codigoSck && sck.conta.codigoSck && String(conta.codigoSck) !== String(sck.conta.codigoSck)) {
      throw aviso(`O relatório é da conta ${sck.conta.codigoSck} do SCK e a tela está na conta ${conta.codigoSck} (${rotuloConta(conta)}). Importe o relatório da conta certa ou troque a conta na tela.`);
    }
    if (sck.conta.numero && !igualConta(sck.conta.numero, conta.conta)) {
      throw aviso(`O relatório é da conta ${sck.conta.numero} e a tela está na conta ${conta.conta}. Confira o arquivo.`);
    }
  }
  if (sck?.empresa && conta.codigo && sck.empresa.codigo && String(sck.empresa.codigo) !== String(conta.codigo)) {
    throw aviso(`O relatório é da empresa ${sck.empresa.codigo} - ${sck.empresa.nome} e a tela está em ${conta.empresa}. Troque a empresa na tela ou importe o relatório certo.`);
  }
  if (ofx?.conta) {
    const bancoConta = soDigitos((conta.banco || "").match(/\d{3}/)?.[0] || "");
    const bancoOfx = soDigitos(ofx.conta.bankid).replace(/^0+/, "");
    if (bancoConta && bancoOfx && bancoConta.replace(/^0+/, "") !== bancoOfx) {
      throw aviso(`O extrato é do banco ${bancoOfx} e a conta da tela é do banco ${bancoConta} (${rotuloConta(conta)}). Importe o extrato da conta certa.`);
    }
    if (ofx.conta.numero && !igualConta(ofx.conta.numero, conta.conta)) {
      throw aviso(`O extrato é da conta ${ofx.conta.agencia ? ofx.conta.agencia + " / " : ""}${ofx.conta.numero} e a tela está na conta ${conta.agencia} / ${conta.conta}. Importe o extrato da conta certa.`);
    }
    if (ofx.conta.agencia && conta.agencia && soDigitos(ofx.conta.agencia).replace(/^0+/, "") !== soDigitos(conta.agencia).replace(/^0+/, "")) {
      throw aviso(`O extrato é da agência ${ofx.conta.agencia} e a conta da tela é da agência ${conta.agencia}. Importe o extrato da conta certa.`);
    }
  }
  return true;
}

export const rotuloConta = (c) => `${c.banco} · Ag. ${c.agencia} · C/C ${c.conta}`;

/**
 * Monta o quadro comparativo. Recebe o resultado de lerConciliacaoSck e de
 * lerExtratoOfx (qualquer um dos dois pode faltar: a tela mostra o que tem).
 *
 * Devolve { data, linhas:[{chave,rotulo,sck,banco,diferenca,situacao}],
 *           situacao, pendencias:[], alertas:[], sckSozinho:{...} }
 *
 * situacao de cada linha e do todo:
 *   "conferido"    → diferença zero
 *   "divergente"   → diferença diferente de zero
 *   "indefinido"   → falta um dos lados, ou o extrato não distingue entrada de saída
 */
export function montarConciliacao({ sck = null, ofx = null } = {}) {
  const alertas = [...(sck?.alertas || []), ...(ofx?.alertas || [])];
  const pendencias = [];
  if (!sck) pendencias.push("Importe o relatório de conciliação do SCK.");
  if (!ofx) pendencias.push("Importe o extrato bancário em .OFX.");

  // o extrato veio sem distinguir entrada de saída? então as saídas dele não valem como zero
  const extratoSemSinal = !!ofx?.alertas?.some((a) => /n[ãa]o distingue entrada de sa[íi]da/i.test(a));

  const linhas = [
    linha("saldo", "Saldo do dia",
      sck?.resumo?.contabil ?? null,
      ofx?.saldoDoDia?.valor ?? null,
      sck && sck.resumo?.contabil == null ? "O relatório não traz o Saldo Contábil." : null),
    linha("entradas", "Entradas", sck ? sck.entradas : null, ofx ? ofx.entradas : null, null),
    linha("saidas", "Saídas", sck ? sck.saidas : null, ofx ? ofx.saidas : null,
      extratoSemSinal ? "O extrato não informa saídas; o valor do banco não pode ser usado." : null),
  ];

  const situacao = linhas.some((l) => l.situacao === "divergente") ? "divergente"
    : linhas.some((l) => l.situacao === "indefinido") ? "indefinido"
    : "conferido";

  const sckSozinho = sck ? {
    diferencaGeral: sck.resumo?.diferencaGeral ?? null,
    saldoExtratoNoSck: sck.resumo?.saldoExtrato ?? null,
    naoMarcados: sck.resumo?.naoMarcados ?? null,
    status: sck.status || "",
    formulaConfere: sck.resumo?.formulaConfere ?? null,
  } : null;

  return {
    data: sck?.data || ofx?.periodo?.fim || null,
    linhas, situacao, pendencias, alertas, sckSozinho,
    quantidade: { sck: sck?.lancamentos?.length || 0, ofx: ofx?.lancamentos?.length || 0 },
  };
}

function linha(chave, rotulo, sck, banco, impedimento) {
  const faltaUm = sck == null || banco == null;
  const diferenca = faltaUm ? null : r2(banco - sck);
  const situacao = impedimento || faltaUm ? "indefinido" : zero(diferenca) ? "conferido" : "divergente";
  return { chave, rotulo, sck, banco, diferenca, situacao, impedimento: impedimento || null };
}
