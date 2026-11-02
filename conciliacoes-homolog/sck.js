/**
 * Conciliações Bancárias — NOCTUS. Leitura do relatório de conciliação do SCK
 * ("RELRegistroConc", gerado pelo Crystal Reports e exportado em Excel .xls).
 *
 * Parte 1 (02/11/2026). Regras confirmadas pelo Rocha no mesmo dia:
 *
 *   - NO RELATÓRIO DE CONCILIAÇÃO a coluna "Débito" são os RECEBIMENTOS
 *     (entradas) e a coluna "Crédito" são as SAÍDAS. É o contrário do extrato
 *     bancário, onde crédito é entrada. A inversão é a visão contábil da conta
 *     do banco: o que entra no banco debita a conta.
 *   - Por isso um histórico "PAGAMENTO VIA PIX (ITAÚ) ... DO TITULAR <nome>" na
 *     coluna Débito é um CLIENTE PAGANDO a parcela dele, ou seja, dinheiro
 *     entrando na empresa — e não um pagamento feito pela empresa.
 *   - A direção vem SEMPRE da coluna, nunca do texto do histórico. O texto
 *     entra só no alerta quando a coluna discorda dele, para o Rocha conferir.
 *
 * Estrutura do arquivo real de 30/10/2026 (Praia Verde, conta 1001283):
 *   "13 - PRAIA VERDE EMPREENDIMENTOS E PARTICIPAÇÕES S.A"   ← empresa (código + nome)
 *   "Conta:  1001283 - BANCO ITAÚ - C/C 84.283-4"            ← conta do SCK
 *   "Data:  30/10/2026"      "Status:  EM ELABORAÇÃO"
 *   "Lançamentos Marcados" / "Lançamentos Não Marcados" / "Lançamentos Adicionados"
 *   cabeçalho: Data | RLC | Lin | Nº Doc. | Nº CNPJ | Contra Partida | Histórico | Débito | Crédito
 *   "Totais =>" com a soma impressa
 *   "Resumo da Conciliação": Saldo dos Lançamentos Marcados, Saldo Contábil,
 *   (+) Adicionados, (-) Não Marcados, (-) Saldo Extrato/Conta, (=) Diferença Geral.
 *
 * Tudo é procurado PELO CONTEÚDO (rótulo e posição do cabeçalho), nunca por
 * endereço fixo de célula, para o relatório não quebrar se o Crystal deslocar
 * uma linha ou coluna. Valor ilegível RECUSA o arquivo: nunca vira zero calado.
 */

export function avisoLeitura(msg) {
  const e = new Error(msg);
  e.aviso = true;
  return e;
}

const limpo = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
const pad2 = (n) => String(n).padStart(2, "0");
const r2 = (n) => Math.round(n * 100) / 100;

/** "12.864,26" → 12864.26; vazio → 0; ilegível → NaN. Mesma regra do fluxo de caixa. */
export function lerValor(v) {
  if (typeof v === "number") return Number.isFinite(v) ? r2(v) : NaN;
  let t = limpo(v).replace(/^R\$\s?/i, "").replace(/\s/g, "");
  if (!t || t === "-" || t === "—") return 0;
  const neg = /^\(.*\)$/.test(t) || /^-/.test(t) || /-$/.test(t);
  t = t.replace(/[()]/g, "").replace(/^-|-$/g, "");
  if (t.includes(",")) {
    if (!/^\d{1,3}(\.\d{3})*,\d+$|^\d+,\d+$/.test(t)) return NaN;
    t = t.replace(/\./g, "").replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(t)) {
    t = t.replace(/\./g, "");
  } else if (!/^\d+(\.\d+)?$/.test(t)) {
    return NaN;
  }
  const n = Number(t);
  return Number.isFinite(n) ? r2(neg ? -n : n) : NaN;
}

/** Data do Crystal: texto "30/10/2026", Date do Excel ou número de série. */
export function lerDataCelula(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return iso(v.getDate(), v.getMonth() + 1, v.getFullYear());
  }
  const t = limpo(v);
  const m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(t);
  if (m) return iso(Number(m[1]), Number(m[2]), Number(m[3]));
  // número de série do Excel (1899-12-30 como dia 0), que é como o .xls guarda a data
  if (/^\d+(\.\d+)?$/.test(t)) {
    const n = Math.floor(Number(t));
    if (n > 20000 && n < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return iso(d.getUTCDate(), d.getUTCMonth() + 1, d.getUTCFullYear());
    }
  }
  return null;
}
function iso(dia, mes, ano) {
  const d = new Date(ano, mes - 1, dia);
  if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) return null;
  return { iso: `${ano}-${pad2(mes)}-${pad2(dia)}`, rotulo: `${pad2(dia)}/${pad2(mes)}/${ano}` };
}

/** É o relatório de conciliação do SCK (e não outro relatório)? */
export function ehRelatorioConciliacao(linhas) {
  let temConta = false, temConciliacao = false;
  for (const linha of (linhas || []).slice(0, 30)) {
    for (const v of linha || []) {
      const t = limpo(v);
      if (/^Conta:\s/i.test(t)) temConta = true;
      if (/^Concilia[çc][ãa]o$/i.test(t) || /Resumo da Concilia[çc][ãa]o/i.test(t)) temConciliacao = true;
    }
  }
  return temConta && temConciliacao;
}

/** Procura, em todas as células, a primeira que casa com o padrão; devolve { i, j, texto }. */
function achar(linhas, padrao, ate = Infinity) {
  for (let i = 0; i < Math.min(linhas.length, ate); i++) {
    const l = linhas[i] || [];
    for (let j = 0; j < l.length; j++) {
      const t = limpo(l[j]);
      if (t && padrao.test(t)) return { i, j, texto: t };
    }
  }
  return null;
}

/** Valor do resumo: o rótulo está numa célula e o número à direita, na mesma linha. */
function valorAoLado(linhas, i, j) {
  const l = linhas[i] || [];
  for (let k = j + 1; k < l.length; k++) {
    if (limpo(l[k]) === "") continue;
    const n = lerValor(l[k]);
    return Number.isNaN(n) ? { ilegivel: limpo(l[k]) } : { valor: n };
  }
  return null;
}

/**
 * Lê o relatório e devolve:
 *   { empresa:{codigo,nome}, conta:{codigoSck,descricao,banco,agencia,numero},
 *     data:{iso,rotulo}, status, responsavel,
 *     lancamentos:[{ data, historico, valor, direcao:"entrada"|"saida", secao, contraPartida, rlc }],
 *     adicionados:{ total, linhas:[{data,historico,valor,direcao}] },
 *     entradas, saidas, totalImpresso, resumo:{...}, alertas:[] }
 * `lancamentos` traz só o movimento do dia (Marcados e Não Marcados); os
 * "Lançamentos Adicionados" vêm à parte, em `adicionados`.
 *
 * `entradas` soma a coluna Débito e `saidas` soma a coluna Crédito (regra do Rocha).
 */
export function lerConciliacaoSck(linhas) {
  if (!Array.isArray(linhas) || !linhas.length) {
    throw avisoLeitura("A planilha está vazia. Confira se é o relatório de conciliação exportado do SCK.");
  }
  if (!ehRelatorioConciliacao(linhas)) {
    throw avisoLeitura("Este arquivo não parece o relatório de conciliação do SCK: não encontrei a linha \"Conta:\" nem o \"Resumo da Conciliação\".");
  }
  const alertas = [];

  /* ---------- cabeçalho ---------- */
  const cabConta = achar(linhas, /^Conta:\s*\S/i, 40);
  if (!cabConta) throw avisoLeitura("Não encontrei a linha \"Conta:\" no relatório. Confira o arquivo.");
  const conta = lerLinhaConta(cabConta.texto);

  const cabData = achar(linhas, /^Data:\s*\d/i, 40);
  const data = cabData ? lerDataCelula(limpo(cabData.texto).replace(/^Data:\s*/i, "")) : null;
  if (!data) throw avisoLeitura("Não encontrei a data do relatório (linha \"Data:\"). Confira o arquivo.");

  const cabEmpresa = achar(linhas, /^(\d{1,4})\s*-\s*\S.*/, 6);
  const empresa = cabEmpresa
    ? { codigo: /^(\d{1,4})/.exec(cabEmpresa.texto)[1], nome: cabEmpresa.texto.replace(/^\d{1,4}\s*-\s*/, "").trim() }
    : null;

  const cabStatus = achar(linhas, /^Status:\s*\S/i, 40);
  const status = cabStatus ? cabStatus.texto.replace(/^Status:\s*/i, "").trim() : "";
  const cabResp = achar(linhas, /^Respons[áa]vel:\s*\S/i, 40);
  const responsavel = cabResp ? cabResp.texto.replace(/^Respons[áa]vel:\s*/i, "").trim() : "";

  /* ---------- colunas dos lançamentos ---------- */
  const cabLanc = acharCabecalho(linhas);
  if (!cabLanc) {
    throw avisoLeitura("Não encontrei o cabeçalho dos lançamentos (Data / Histórico / Débito / Crédito). Confira o arquivo.");
  }

  /* ---------- lançamentos ---------- */
  const { lancamentos, totalImpresso, ilegiveis } = lerLancamentos(linhas, cabLanc);
  if (ilegiveis.length) {
    throw avisoLeitura(`Não consegui ler o valor de ${ilegiveis.length} lançamento(s) do relatório (${ilegiveis.slice(0, 3).join("; ")}). Confira o arquivo.`);
  }

  /*
   * "Lançamentos Adicionados" (regra do Rocha, 02/11/2026, relatório das 11:24):
   * são valores que estão no BANCO e não estão no contábil, acrescentados à mão
   * na conciliação — no arquivo real, três "REND PAGO APLIC AUT APR" de dias
   * anteriores (1,43 + 25,53 + 0,04 = 27,00). NÃO são movimento do dia: ficam
   * fora de entradas/saídas e são devolvidos à parte, para a composição da
   * diferença tratá-los como "a validar".
   */
  const doDia = lancamentos.filter((l) => l.secao !== "adicionados");
  const adicionadosLinhas = lancamentos.filter((l) => l.secao === "adicionados");
  const entradas = r2(doDia.filter((l) => l.direcao === "entrada").reduce((s, l) => s + l.valor, 0));
  const saidas = r2(doDia.filter((l) => l.direcao === "saida").reduce((s, l) => s + l.valor, 0));
  // com sinal: Débito (entrada no banco) soma, Crédito (saída) subtrai
  const adicionadosTotal = r2(adicionadosLinhas.reduce((s, l) => s + (l.direcao === "entrada" ? l.valor : -l.valor), 0));
  const adicionadosBruto = r2(adicionadosLinhas.reduce((s, l) => s + l.valor, 0));

  // conferência contra o "Totais =>" impresso pelo próprio SCK (soma das seções)
  const somaLida = r2(entradas + saidas + adicionadosBruto);
  if (totalImpresso != null && Math.abs(somaLida - totalImpresso) >= 0.005) {
    alertas.push(`A soma dos lançamentos lidos (${moeda(somaLida)}) é diferente do total impresso no relatório (${moeda(totalImpresso)}).`);
  }
  // histórico que discorda da coluna: não muda o cálculo, só avisa
  const discordantes = doDia.filter((l) => l.direcao === "entrada" && /^\s*(PAGAMENTO|TARIFA|TED ENVIAD|DEB[IÍ]TO)/i.test(l.historico)
    && !/DO TITULAR|RECEB/i.test(l.historico));
  if (discordantes.length) {
    alertas.push(`${discordantes.length} lançamento(s) estão na coluna Débito (entrada) mas o histórico começa com "pagamento" ou "tarifa". Confira se a classificação do SCK está certa.`);
  }
  // lançamento com data diferente da data do relatório
  // (os adicionados são de outros dias por natureza e não entram neste aviso)
  const foraDaData = doDia.filter((l) => l.data && l.data.iso !== data.iso);
  if (foraDaData.length) {
    alertas.push(`${foraDaData.length} lançamento(s) têm data diferente de ${data.rotulo} e foram somados mesmo assim, como o SCK imprimiu.`);
  }

  /* ---------- resumo da conciliação ---------- */
  const resumo = lerResumo(linhas);
  // o "(+) Saldo dos Lançamentos Adicionados" do resumo tem de bater com as linhas lidas
  if (resumo.adicionados != null && Math.abs(resumo.adicionados - adicionadosTotal) >= 0.005) {
    alertas.push(`Os Lançamentos Adicionados lidos somam ${moeda(adicionadosTotal)}, mas o resumo do relatório informa ${moeda(resumo.adicionados)}. Confira o relatório.`);
  }

  const adicionados = {
    total: adicionadosTotal,
    linhas: adicionadosLinhas.map(({ data: d, historico, valor, direcao }) => ({ data: d, historico, valor, direcao })),
  };

  return { empresa, conta, data, status, responsavel, lancamentos: doDia, adicionados,
           entradas, saidas, totalImpresso, resumo, alertas };
}

const moeda = (v) => Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/** "Conta:  1001283 - BANCO ITAÚ - C/C 84.283-4" → código do SCK, banco e número. */
export function lerLinhaConta(texto) {
  const t = limpo(texto).replace(/^Conta:\s*/i, "");
  const codigo = (/^(\d{3,10})\b/.exec(t) || [])[1] || "";
  const resto = t.replace(/^\d{3,10}\s*-\s*/, "");
  const mNum = /C\/?C\s*([\d.\-]+)|CONTA\s*([\d.\-]+)/i.exec(resto);
  const numero = (mNum ? (mNum[1] || mNum[2]) : "").replace(/\./g, "");
  const banco = resto.replace(/\s*-\s*C\/?C.*$/i, "").replace(/^BANCO\s+/i, "").trim();
  return { codigoSck: codigo, descricao: t, banco, numero };
}

/** Acha a linha do cabeçalho e as colunas de Data, Histórico, Débito e Crédito. */
function acharCabecalho(linhas) {
  for (let i = 0; i < linhas.length; i++) {
    const l = linhas[i] || [];
    const col = {};
    for (let j = 0; j < l.length; j++) {
      const t = limpo(l[j]).toLowerCase();
      if (t === "data" && col.data == null) col.data = j;
      if (t === "histórico" || t === "historico") col.historico = j;
      if (t === "débito" || t === "debito") col.debito = j;
      if (t === "crédito" || t === "credito") col.credito = j;
      if (t === "rlc") col.rlc = j;
      if (/^contra\s*partida$/.test(t)) col.contraPartida = j;
    }
    if (col.debito != null && col.credito != null && col.historico != null) return { i, col };
  }
  return null;
}

/**
 * Varre as linhas abaixo do cabeçalho. Cada lançamento é uma linha com histórico
 * e valor em Débito ou em Crédito. Como o Crystal mescla células, o valor pode
 * começar na coluna do cabeçalho ou um pouco à direita: procura dentro da faixa
 * que vai da coluna do rótulo até a coluna do rótulo seguinte.
 */
function lerLancamentos(linhas, cab) {
  const { i: iCab, col } = cab;
  const fimDebito = col.credito > col.debito ? col.credito : col.debito + 3;
  const lancamentos = [];
  const ilegiveis = [];
  let totalImpresso = null;
  let secao = "";

  // a varredura começa no alto do arquivo porque os títulos de seção ("Lançamentos
  // Não Marcados") ficam ACIMA do cabeçalho das colunas; só os lançamentos abaixo
  // do cabeçalho são coletados.
  for (let i = 0; i < linhas.length; i++) {
    const l = linhas[i] || [];
    const textoLinha = l.map(limpo).filter(Boolean).join(" ");
    if (!textoLinha) continue;

    if (/Resumo da Concilia[çc][ãa]o/i.test(textoLinha)) break;
    if (/Lan[çc]amentos\s+(Marcados|N[ãa]o Marcados|Adicionados)/i.test(textoLinha)) {
      // "Não existem Lançamentos X" é declaração de ausência: não abre seção com conteúdo
      if (!/N[ãa]o existem/i.test(textoLinha)) {
        secao = /N[ãa]o Marcados/i.test(textoLinha) ? "nao-marcados"
              : /Adicionados/i.test(textoLinha) ? "adicionados" : "marcados";
      }
      continue;
    }
    if (i <= iCab) continue;
    // cada seção repete o cabeçalho das colunas (no relatório com "Lançamentos
    // Adicionados" ele aparece duas vezes): cabeçalho não é lançamento
    if (l.some((v) => /^d[ée]bito$/i.test(limpo(v))) && l.some((v) => /^cr[ée]dito$/i.test(limpo(v)))) continue;
    if (/^Totais?\s*=>/i.test(limpo(l.find((v) => /^Totais?\s*=>/i.test(limpo(v))) || ""))) {
      const v = faixa(l, col.debito, fimDebito);
      if (v && v.ilegivel == null) totalImpresso = r2((totalImpresso || 0) + v.valor);
      const vc = faixa(l, col.credito, col.credito + 3);
      if (vc && vc.ilegivel == null && vc.valor) totalImpresso = r2((totalImpresso || 0) + vc.valor);
      continue;
    }

    const historico = limpo(l[col.historico] || procurarTexto(l, col.historico));
    const vDeb = faixa(l, col.debito, fimDebito);
    const vCre = faixa(l, col.credito, col.credito + 3);
    if (!vDeb && !vCre) continue;
    if (vDeb?.ilegivel != null) { ilegiveis.push(`"${vDeb.ilegivel}"`); continue; }
    if (vCre?.ilegivel != null) { ilegiveis.push(`"${vCre.ilegivel}"`); continue; }
    const valor = (vDeb?.valor || 0) || (vCre?.valor || 0);
    if (!valor) continue;

    lancamentos.push({
      data: lerDataCelula(l[col.data]) || null,
      rlc: limpo(l[col.rlc] ?? ""),
      contraPartida: limpo(l[col.contraPartida] ?? ""),
      historico,
      valor: Math.abs(valor),
      direcao: vDeb?.valor ? "entrada" : "saida",   // Débito = recebimento (regra do Rocha)
      secao,
    });
  }
  return { lancamentos, totalImpresso, ilegiveis };
}

/** Primeiro valor não vazio entre as colunas [de, ate). */
function faixa(l, de, ate) {
  if (de == null) return null;
  for (let j = de; j < Math.min(ate, l.length); j++) {
    const t = limpo(l[j]);
    if (t === "") continue;
    const n = lerValor(t);
    if (Number.isNaN(n)) return { ilegivel: t };
    return n === 0 ? null : { valor: n };
  }
  return null;
}
/** O histórico às vezes começa uma coluna depois, por causa das mesclagens. */
function procurarTexto(l, de) {
  for (let j = de; j < Math.min(de + 3, l.length); j++) {
    const t = limpo(l[j]);
    if (t && !/^[\d\s.,()\-/R$]+$/.test(t)) return t;
  }
  return "";
}

const ROTULOS_RESUMO = [
  ["marcados", /Saldo dos Lan[çc]amentos Marcados/i],
  ["contabil", /Saldo Cont[áa]bil/i],
  ["adicionados", /Saldo dos Lan[çc]amentos Adicionados/i],
  ["naoMarcados", /Saldo dos Lan[çc]amentos N[ãa]o Marcados/i],
  ["saldoExtrato", /Saldo Extrato\s*\/?\s*Conta/i],
  ["diferencaGeral", /Diferen[çc]a Geral da Concilia[çc][ãa]o/i],
];

/** Lê o bloco "Resumo da Conciliação". Campo ausente fica null (não vira zero). */
export function lerResumo(linhas) {
  const inicio = achar(linhas, /Resumo da Concilia[çc][ãa]o/i);
  const de = inicio ? inicio.i : 0;
  const r = {};
  for (const [chave, padrao] of ROTULOS_RESUMO) {
    r[chave] = null;
    for (let i = de; i < linhas.length; i++) {
      const l = linhas[i] || [];
      for (let j = 0; j < l.length; j++) {
        if (!padrao.test(limpo(l[j]))) continue;
        const v = valorAoLado(linhas, i, j);
        if (v && v.ilegivel == null) r[chave] = v.valor;
        break;
      }
      if (r[chave] != null) break;
    }
  }
  // confere a fórmula impressa pelo próprio SCK: contábil + adicionados - não marcados - extrato = diferença
  if ([r.contabil, r.adicionados, r.naoMarcados, r.saldoExtrato, r.diferencaGeral].every((v) => v != null)) {
    const calc = r2(r.contabil + r.adicionados - r.naoMarcados - r.saldoExtrato);
    r.formulaConfere = Math.abs(calc - r.diferencaGeral) < 0.005;
    r.diferencaCalculada = calc;
  } else {
    r.formulaConfere = null;
  }
  return r;
}

/** Converte o .xls (ArrayBuffer) em matriz de linhas com o SheetJS (window.XLSX). */
export function linhasDoExcel(XLSX, buffer) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw avisoLeitura("O arquivo não tem nenhuma planilha.");
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
}
