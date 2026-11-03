/**
 * Conciliações Bancárias — NOCTUS. Conferência por tipo de lançamento e
 * composição da diferença de saldo (Parte 1, 02/11/2026).
 *
 * Regras do Rocha ("Regras de comparação e identificação das divergências
 * bancárias", 02/11/2026) e o ajuste do mesmo dia: o que é rendimento ou saldo
 * anterior SEMPRE se refere aos "Lançamentos Adicionados" do relatório do SCK.
 *
 *   1.1 COBR IBBA (extrato) × "RECEBIMENTO BAIXA BANCÁRIA" (SCK), pelo mesmo valor.
 *   1.2 PIX QRS (extrato) × SOMA de todos os "PAGAMENTO VIA PIX" do SCK.
 *   1.3 PIX RECEBIDO (extrato): não existe no SCK, é lançamento manual.
 *       Vai para "Lançar", um a um, com data, histórico, identificação e valor.
 *   1.4 Saldo: SALDO TOTAL DISPONÍVEL DIA (banco) × Saldo Contábil (SCK), e a
 *       diferença é DECOMPOSTA, não só apontada.
 *   1.5 Rendimentos: os Lançamentos Adicionados do SCK vão para "Validar" (já
 *       estão na conciliação; não sugerir lançamento de novo, para não duplicar).
 *       Rendimento NOVO no extrato, que ainda não está nos adicionados, vai para
 *       "Adicionar na conciliação".
 *       O SALDO ANTERIOR do extrato é conferido com a Diferença Geral do SCK
 *       (ver `conferenciaSaldoAnterior`); é validação, não parcela da soma.
 *
 * A conta da composição, em todas as linhas: Diferença = BANCO − SCK.
 *   + o que está no banco e não está no contábil (PIX recebido, rendimento);
 *   − o que está no contábil e não está no banco;
 * e o que sobra depois de explicar tudo aparece como "Diferença sem origem
 * identificada" — nunca é escondido nem distribuído.
 */

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
const zero = (n) => Math.abs(Number(n || 0)) < 0.005;
const igual = (a, b) => zero(r2(a) - r2(b));

/* ------------------------------------------------------------ classificação */
/** Tipo do lançamento do EXTRATO pelo histórico (MEMO). */
export function tipoExtrato(memo) {
  const t = String(memo || "").toUpperCase();
  if (/^COBR\b/.test(t)) return "cobranca";
  if (/^PIX\s+QRS\b/.test(t)) return "pixQrs";
  if (/^PIX\s+RECEBIDO\b/.test(t)) return "pixRecebido";
  if (/\bREND(IMENTO)?S?\b/.test(t)) return "rendimento";
  return "outro";
}
/** Tipo do lançamento do SCK pelo histórico. */
export function tipoSck(historico) {
  const t = String(historico || "").toUpperCase();
  if (/RECEBIMENTO\s+BAIXA\s+BANC/.test(t)) return "baixaBancaria";
  if (/PAGAMENTO\s+VIA\s+PIX/.test(t)) return "pagamentoPix";
  return "outro";
}

/**
 * Exemplo fictício (o repositório é público: nada de cliente real aqui):
 * "PIX RECEBIDO FULANO D30/10 FULANO DE TAL 000.000.000-00"
 *   → { nome:"FULANO DE TAL", documento:"000.000.000-00" }
 * Se o formato mudar, devolve o que conseguir (o histórico completo continua na linha).
 */
export function identificarPix(memo) {
  const t = String(memo || "").replace(/\s+/g, " ").trim();
  const doc = /(\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}|\d{11}|\d{14})\s*$/.exec(t);
  const documento = doc ? doc[1] : "";
  const semDoc = doc ? t.slice(0, doc.index).trim() : t;
  const m = /\d{2}\/\d{2}\s+(.+)$/.exec(semDoc);
  const nome = m ? m[1].trim() : semDoc.replace(/^PIX\s+RECEBIDO\s*/i, "").trim();
  return { nome, documento };
}

/**
 * Histórico do lançamento na planilha de importação (Parte 2, regra do Rocha, 02/11/2026):
 *   extrato:  "PIX RECEBIDO FULANO D30/10 FULANO DE TAL 000.000.000-00"
 *   planilha: "PIX RECEBIDO 30/10 - FULANO DE TAL 000.000.000-00"
 * Tira o nome abreviado colado na data e põe " - " antes do nome completo.
 * `padrao:false` quando o histórico não tem a forma esperada: aí devolve o
 * original sem mexer, e a tela pede conferência (nunca inventa um histórico).
 * `dataConfere:false` quando o dd/mm do histórico difere da data do lançamento.
 */
export function historicoImportacao(memo, data = null) {
  const original = String(memo || "").replace(/\s+/g, " ").trim();
  const m = /^PIX\s+RECEBIDO\s+.*?(\d{2}\/\d{2})\s+(.+)$/i.exec(original);
  if (!m) return { texto: original, padrao: false, dataConfere: null };
  const { nome, documento } = identificarPix(original);
  if (!nome) return { texto: original, padrao: false, dataConfere: null };
  const ddmm = m[1];
  const dataConfere = data?.rotulo ? data.rotulo.slice(0, 5) === ddmm : null;
  return { texto: `PIX RECEBIDO ${ddmm} - ${nome}${documento ? ` ${documento}` : ""}`, padrao: true, dataConfere };
}

/* ------------------------------------------------------------ regras */
/**
 * Recebe o resultado de lerConciliacaoSck e de lerExtratoOfx (com dataAlvo).
 * Devolve null se faltar um dos dois.
 *
 * {
 *   regras:      [{ regra, rotulo, banco, sck, diferenca, situacao, detalhe }],
 *   componentes: [{ grupo, tipo, descricao, valor, tratamento, data, historico, nome, documento, fitid, linhas }],
 *   grupos:      { lancar, adicionar, validar, divergencia } (somas),
 *   diferencaSaldo, explicado, semOrigem, fecha,
 *   saldoAnterior: { banco, sck, situacao } | null
 * }
 */
export function comporDiferenca({ sck = null, ofx = null } = {}) {
  if (!sck || !ofx) return null;
  const dataAlvo = sck.data?.iso || null;
  const movBanco = (ofx.lancamentos || []).filter((t) => t.data && (!dataAlvo || t.data.iso === dataAlvo));
  const movSck = sck.lancamentos || [];

  const banco = movBanco.map((t) => ({ ...t, categoria: tipoExtrato(t.memo), usado: false }));
  const contabil = movSck.map((l) => ({ ...l, categoria: tipoSck(l.historico), usado: false }));
  const sinalB = (t) => (t.direcao === "saida" ? -t.valor : t.valor);
  const sinalS = (l) => (l.direcao === "saida" ? -l.valor : l.valor);

  const regras = [];
  const componentes = [];

  /* 1.1 cobrança × baixa bancária, um a um pelo valor */
  const cobrancas = banco.filter((t) => t.categoria === "cobranca");
  const baixas = contabil.filter((l) => l.categoria === "baixaBancaria");
  for (const c of cobrancas) {
    const par = baixas.find((b) => !b.usado && b.direcao === c.direcao && igual(b.valor, c.valor));
    c.usado = true;
    if (par) {
      par.usado = true;
      regras.push({ regra: "1.1", rotulo: "Cobrança × Baixa bancária", banco: c.valor, sck: par.valor,
                    diferenca: 0, situacao: "conferido", detalhe: par.historico });
    } else {
      regras.push({ regra: "1.1", rotulo: "Cobrança × Baixa bancária", banco: c.valor, sck: null,
                    diferenca: r2(sinalB(c)), situacao: "divergente",
                    detalhe: "Não há RECEBIMENTO BAIXA BANCÁRIA com o mesmo valor no SCK." });
      componentes.push({ grupo: "divergencia", tipo: "cobrancaSemBaixa", descricao: `Cobrança sem baixa no SCK (${c.memo})`,
                         valor: r2(sinalB(c)), tratamento: "Conferir a baixa bancária no SCK",
                         data: c.data, historico: c.memo, fitid: c.fitid });
    }
  }
  for (const b of baixas.filter((x) => !x.usado)) {
    b.usado = true;
    regras.push({ regra: "1.1", rotulo: "Cobrança × Baixa bancária", banco: null, sck: b.valor,
                  diferenca: r2(-sinalS(b)), situacao: "divergente",
                  detalhe: "Baixa bancária no SCK sem cobrança de mesmo valor no extrato." });
    componentes.push({ grupo: "divergencia", tipo: "baixaSemCobranca", descricao: "Baixa bancária no SCK sem cobrança no extrato",
                       valor: r2(-sinalS(b)), tratamento: "Conferir a baixa no SCK", data: b.data, historico: b.historico });
  }

  /* 1.2 PIX QRS × soma dos PAGAMENTO VIA PIX */
  const qrs = banco.filter((t) => t.categoria === "pixQrs");
  const pagPix = contabil.filter((l) => l.categoria === "pagamentoPix");
  if (qrs.length || pagPix.length) {
    const somaB = r2(qrs.reduce((s, t) => s + sinalB(t), 0));
    const somaS = r2(pagPix.reduce((s, l) => s + sinalS(l), 0));
    qrs.forEach((t) => { t.usado = true; });
    pagPix.forEach((l) => { l.usado = true; });
    const dif = r2(somaB - somaS);
    regras.push({ regra: "1.2", rotulo: "PIX QRS × Pagamentos via PIX", banco: somaB, sck: somaS, diferenca: dif,
                  situacao: zero(dif) ? "conferido" : "divergente",
                  detalhe: `${qrs.length} PIX QRS no extrato · ${pagPix.length} pagamento(s) via PIX no SCK` });
    if (!zero(dif)) {
      componentes.push({ grupo: "divergencia", tipo: "pixQrs", descricao: "Diferença entre PIX QRS e pagamentos via PIX",
                         valor: dif, tratamento: "Conferir os pagamentos via PIX do SCK" });
    }
  }

  /* 1.3 PIX recebidos: lançamento manual, um a um */
  const recebidos = banco.filter((t) => t.categoria === "pixRecebido");
  for (const p of recebidos) {
    p.usado = true;
    const id = identificarPix(p.memo);
    componentes.push({ grupo: "lancar", tipo: "pixRecebido", descricao: "PIX recebido", valor: r2(sinalB(p)),
                       tratamento: "Lançamento manual no SCK", data: p.data, historico: p.memo,
                       nome: id.nome, documento: id.documento, fitid: p.fitid });
  }
  if (recebidos.length) {
    regras.push({ regra: "1.3", rotulo: "PIX recebidos", banco: r2(recebidos.reduce((s, t) => s + sinalB(t), 0)),
                  sck: null, diferenca: r2(recebidos.reduce((s, t) => s + sinalB(t), 0)), situacao: "pendente",
                  detalhe: `${recebidos.length} PIX a lançar manualmente no SCK` });
  }

  /* 1.5 rendimentos: novos no extrato × já adicionados no SCK */
  const adicionados = (sck.adicionados?.linhas || []).map((l) => ({ ...l, usado: false }));
  const rendExtrato = banco.filter((t) => t.categoria === "rendimento");
  for (const r of rendExtrato) {
    r.usado = true;
    const ja = adicionados.find((a) => !a.usado && igual(a.valor, r.valor) && (!a.data || !r.data || a.data.iso === r.data.iso));
    if (ja) { ja.usado = true; ja.noExtratoDoDia = true; continue; }
    componentes.push({ grupo: "adicionar", tipo: "rendimentoNovo", descricao: "Rendimento novo no extrato",
                       valor: r2(sinalB(r)), tratamento: "Adicionar na conciliação do SCK",
                       data: r.data, historico: r.memo, fitid: r.fitid });
  }
  if (adicionados.length) {
    componentes.push({ grupo: "validar", tipo: "adicionados", descricao: "Rendimentos já adicionados na conciliação",
                       valor: r2(sck.adicionados.total), tratamento: "Validar: já constam na conciliação, não lançar de novo",
                       linhas: adicionados.map(({ data, historico, valor, direcao }) => ({ data, historico, valor, direcao })) });
  }

  /* o que não se encaixou em nenhuma regra */
  for (const t of banco.filter((x) => !x.usado)) {
    componentes.push({ grupo: "divergencia", tipo: "extratoSemRegra", descricao: `No extrato, sem correspondente (${t.memo})`,
                       valor: r2(sinalB(t)), tratamento: "Identificar e lançar ou justificar",
                       data: t.data, historico: t.memo, fitid: t.fitid });
  }
  for (const l of contabil.filter((x) => !x.usado)) {
    componentes.push({ grupo: "divergencia", tipo: "sckSemRegra", descricao: "No SCK, sem correspondente no extrato",
                       valor: r2(-sinalS(l)), tratamento: "Conferir o lançamento no SCK",
                       data: l.data, historico: l.historico });
  }

  /* 1.4 composição do saldo */
  const saldoBanco = ofx.saldoDoDia?.valor ?? null;
  const saldoSck = sck.resumo?.contabil ?? null;
  const diferencaSaldo = saldoBanco == null || saldoSck == null ? null : r2(saldoBanco - saldoSck);
  const soma = (g) => r2(componentes.filter((c) => c.grupo === g).reduce((s, c) => s + c.valor, 0));
  const grupos = { lancar: soma("lancar"), adicionar: soma("adicionar"), validar: soma("validar"), divergencia: soma("divergencia") };
  const explicado = r2(grupos.lancar + grupos.adicionar + grupos.validar + grupos.divergencia);
  const semOrigem = diferencaSaldo == null ? null : r2(diferencaSaldo - explicado);

  return {
    regras, componentes, grupos, diferencaSaldo, explicado, semOrigem,
    fecha: semOrigem != null && zero(semOrigem),
    saldoAnterior: conferenciaSaldoAnterior({ sck, ofx }),
  };
}

/**
 * Saldo anterior do extrato × Diferença Geral da Conciliação do SCK.
 *
 * Fundamento (fórmula impressa no próprio relatório):
 *   Diferença Geral = Saldo Contábil + Adicionados − Não Marcados − Saldo Extrato/Conta.
 * Com o Saldo Extrato/Conta ainda zerado (relatório "em elaboração"), e os não
 * marcados sendo o movimento do dia já contabilizado, isso dá
 *   saldo contábil de abertura + adicionados = saldo do banco na abertura,
 * que é o SALDO ANTERIOR do extrato. No arquivo real: 43.788,22 + 27,00 −
 * 43.808,08 = 7,14 = SALDO ANTERIOR. É INFERÊNCIA a partir da fórmula, por isso
 * só vale com o Saldo Extrato/Conta zerado; com ele preenchido, não compara.
 */
export function conferenciaSaldoAnterior({ sck = null, ofx = null } = {}) {
  const banco = ofx?.saldoAnterior?.valor ?? null;
  const dg = sck?.resumo?.diferencaGeral ?? null;
  if (banco == null || dg == null) return null;
  if (sck.resumo.saldoExtrato != null && !zero(sck.resumo.saldoExtrato)) {
    return { banco, sck: dg, diferenca: null, situacao: "indefinido",
             motivo: "O Saldo Extrato/Conta do SCK já está preenchido; a comparação direta não se aplica." };
  }
  const diferenca = r2(banco - dg);
  return { banco, sck: dg, diferenca, situacao: zero(diferenca) ? "conferido" : "divergente", motivo: null };
}

/**
 * Versão para gravar no servidor (Parte 2 vai usar): valores, datas, FITID e o
 * tipo de cada componente. Nome e CPF/CNPJ dos pagadores NÃO vão — a Parte 2
 * relê o extrato na tela quando for lançar, e o FITID liga uma coisa à outra.
 */
export function composicaoParaGravar(c) {
  if (!c) return null;
  return {
    diferencaSaldo: c.diferencaSaldo, explicado: c.explicado, semOrigem: c.semOrigem, fecha: c.fecha,
    grupos: c.grupos,
    regras: c.regras.map(({ regra, rotulo, banco, sck, diferenca, situacao }) => ({ regra, rotulo, banco, sck, diferenca, situacao })),
    componentes: c.componentes.map((x) => ({
      grupo: x.grupo, tipo: x.tipo, valor: x.valor, data: x.data?.iso || null, fitid: x.fitid || null,
      // históricos do SCK de rendimento não têm dado pessoal; os demais ficam de fora
      ...(x.tipo === "adicionados" ? { linhas: x.linhas.map((l) => ({ data: l.data?.iso || null, historico: l.historico, valor: l.valor })) } : {}),
    })),
    saldoAnterior: c.saldoAnterior,
  };
}
