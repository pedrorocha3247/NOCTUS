/**
 * Controle Bancário — leitura da planilha "Financeiro" e utilitários.
 *
 * Tudo aqui roda NO NAVEGADOR e não faz rede: a planilha tem agência, conta,
 * CNPJ e valores do dia, e o repositório do NOCTUS é público. Por isso o
 * módulo LÊ o .xlsx que a pessoa escolhe (SheetJS, o mesmo da Conferência) e
 * só guarda o resultado no localStorage do próprio navegador.
 *
 * O arquivo é de funções puras (recebem o XLSX e o workbook), para poder ser
 * testado fora do navegador. A leitura se apoia nos RÓTULOS da planilha (cabeçalhos
 * "BANCO / Ag. / C/C", títulos "CONTAS ...", dias numerados mesclados) e não em
 * endereços fixos, porque as abas das empresas não têm todas as mesmas colunas
 * (IRM e Praia Verde têm uma coluna DÍGITO a mais, por exemplo).
 */

/* ----------------------------------------------------------------- texto */
export const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim();
export const semAcento = (s) =>
  norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const soDigitos = (s) => String(s ?? "").replace(/\D/g, "");

/** Abas de empresa lidas (nome da aba na planilha → id curto usado no módulo). */
export const ABAS_EMPRESA = [
  ["MEI", "momentum"], ["Kasil", "kasil"], ["Praia Verde", "praia-verde"],
  ["Slim", "slim"], ["Posto", "posto"], ["MMH", "mmh"], ["Realiza", "realiza"],
  ["M5", "m5"], ["IRM", "irm"], ["Pick Money", "pick-money"], ["M3", "m3"],
  ["RVM", "rvm"], ["Abrasma", "abrasma"],
];

/** Nomes que a planilha usa para a mesma empresa → id. Chave sem acento, minúscula. */
const SINONIMOS = {
  momentum: "momentum", mei: "momentum", "momentum matriz": "momentum",
  kasil: "kasil", "praia verde": "praia-verde", pv: "praia-verde",
  slim: "slim", posto: "posto", "posto plaza": "posto", "posto sb": "posto",
  mmh: "mmh", "mmh holding": "mmh", realiza: "realiza", modo: "realiza",
  m5: "m5", irm: "irm", "pick money": "pick-money", pkm: "pick-money",
  m3: "m3", "m3 assets": "m3", rvm: "rvm", abrasma: "abrasma",
};
export function idEmpresa(nome) {
  const k = semAcento(nome);
  if (SINONIMOS[k]) return SINONIMOS[k];
  for (const [chave, id] of Object.entries(SINONIMOS)) {
    if (k === chave || k.startsWith(chave + " ")) return id;
  }
  return null;
}

/* ------------------------------------------------------------------ CNPJ */
export function validarCnpj(valor) {
  const d = soDigitos(valor);
  if (d.length !== 14 || /^(\d)\1+$/.test(d)) return false;
  const dv = (base) => {
    const pesos = base.length === 12
      ? [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] : [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
    const r = [...base].reduce((s, x, i) => s + Number(x) * pesos[i], 0) % 11;
    return r < 2 ? "0" : String(11 - r);
  };
  return d[12] === dv(d.slice(0, 12)) && d[13] === dv(d.slice(0, 13));
}
const REGEX_CNPJ = /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g;

/* ----------------------------------------------------- células (SheetJS) */
const ender = (XLSX, r, c) => XLSX.utils.encode_cell({ r, c });
function cel(XLSX, ws, r, c) { return ws[ender(XLSX, r, c)]; }
function txt(XLSX, ws, r, c) {
  const x = cel(XLSX, ws, r, c);
  if (!x || x.t === "e") return "";
  if (x.t === "s") return norm(x.v);
  if (x.t === "n") return String(x.v);
  if (x.t === "b") return x.v ? "true" : "false";
  return norm(x.w ?? x.v);
}
function numero(XLSX, ws, r, c) {
  const x = cel(XLSX, ws, r, c);
  return x && x.t === "n" && Number.isFinite(x.v) ? x.v : null;
}
function intervalo(XLSX, ws) {
  return ws && ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { r: 0, c: 0 }, e: { r: -1, c: -1 } };
}

/* ------------------------------------------------------- contas bancárias */
/** "Itaú - 341 (recebimentos)" → { banco:"Itaú", codBanco:"341", obs:"recebimentos" } */
export function lerBanco(bruto) {
  const s = norm(bruto);
  const m = /^(.*?)[\s-]*(\d{1,3})?\s*(?:\((.*)\))?$/.exec(s);
  let banco = norm(m?.[1] ?? s).replace(/[\s-]+$/, "");
  if (!banco) banco = s;
  return { banco, codBanco: m?.[2] || "", obs: norm(m?.[3] || "") };
}

/**
 * Lê uma tabela "BANCO | Ag. | [DÍGITO] | C/C | [DÍGITO] | CNPJ" cujo cabeçalho
 * está na linha `lh`, começando na coluna `cb`. Termina na primeira linha sem banco.
 */
function lerTabelaBancos(XLSX, ws, nomeAba, lh, cb, colMax) {
  const mapa = []; // { c, tipo }
  let vazios = 0;
  for (let c = cb; c <= colMax && vazios < 2; c++) {
    const t = semAcento(txt(XLSX, ws, lh, c)).replace(/\./g, "");
    if (!t) { vazios++; continue; }
    vazios = 0;
    const tipo = t === "banco" ? "banco" : t === "ag" ? "ag" : t === "c/c" ? "cc"
      : t === "digito" ? "dig" : t === "cnpj" ? "cnpj" : null;
    if (tipo) mapa.push({ c, tipo });
  }
  const col = (tipo) => mapa.find((m) => m.tipo === tipo)?.c;
  const cBanco = col("banco"), cAg = col("ag"), cCc = col("cc"), cCnpj = col("cnpj");
  const digs = mapa.filter((m) => m.tipo === "dig").map((m) => m.c);
  const cAgDig = digs.find((c) => c < cCc), cCcDig = digs.find((c) => c > cCc);
  const linhas = [];
  let cnpjAtual = "";
  const fim = intervalo(XLSX, ws).e.r;
  for (let r = lh + 1; r <= fim; r++) {
    const bancoBruto = txt(XLSX, ws, r, cBanco);
    if (!bancoBruto) break;
    const ag = txt(XLSX, ws, r, cAg), cc = txt(XLSX, ws, r, cCc);
    if (!ag && !cc) continue; // linha de título solta (ex.: nome da empresa)
    const cnpjCel = cCnpj != null ? txt(XLSX, ws, r, cCnpj) : "";
    if (cnpjCel) cnpjAtual = cnpjCel;
    const agDig = cAgDig != null ? txt(XLSX, ws, r, cAgDig) : "";
    const ccDig = cCcDig != null ? txt(XLSX, ws, r, cCcDig) : "";
    const b = lerBanco(bancoBruto);
    linhas.push({
      ...b, bancoBruto,
      agencia: agDig ? `${ag}-${agDig}` : ag,
      conta: ccDig ? `${cc}-${ccDig}` : cc,
      cnpj: cnpjCel || cnpjAtual,
      origem: `${nomeAba}!${ender(XLSX, r, cBanco)}`,
    });
  }
  return linhas;
}

/** Chave que identifica uma conta bancária independente de formatação. */
export function chaveConta(c) {
  const cod = c.codBanco || semAcento(c.banco);
  const sz = (s) => soDigitos(s).replace(/^0+/, ""); // "00042-1" e "42-1" são a mesma conta
  return `${cod}|${sz(c.agencia)}|${sz(c.conta)}`;
}

/** Bloco "CONTAS <EMPRESA>" da própria aba da empresa (colunas H em diante). */
function contasDaAba(XLSX, ws, nomeAba) {
  const ref = intervalo(XLSX, ws);
  for (let r = 0; r <= Math.min(ref.e.r, 12); r++) {
    for (let c = 6; c <= Math.min(ref.e.c, 12); c++) {
      if (semAcento(txt(XLSX, ws, r, c)) === "banco") {
        const titulo = r > 0 ? txt(XLSX, ws, r - 1, c) : "";
        return { titulo, tituloEnder: `${nomeAba}!${ender(XLSX, r - 1, c)}`,
          linhas: lerTabelaBancos(XLSX, ws, nomeAba, r, c, ref.e.c) };
      }
    }
  }
  return { titulo: "", tituloEnder: "", linhas: [] };
}

/** Blocos "CONTAS <EMPRESA>" da aba "Todas as Contas" (duas colunas de blocos). */
function blocosTodasAsContas(XLSX, ws, nomeAba) {
  const ref = intervalo(XLSX, ws);
  const blocos = [];
  for (const c0 of [8, 13]) {
    for (let r = 0; r <= ref.e.r; r++) {
      const t = txt(XLSX, ws, r, c0);
      if (!/^contas\s/i.test(t)) continue;
      let lh = -1;
      for (let k = r + 1; k <= Math.min(r + 2, ref.e.r); k++) {
        if (semAcento(txt(XLSX, ws, k, c0)) === "banco") { lh = k; break; }
      }
      if (lh < 0) continue;
      blocos.push({ titulo: t, id: idEmpresa(t.replace(/^contas\s+/i, "")),
        linhas: lerTabelaBancos(XLSX, ws, nomeAba, lh, c0, c0 + 4) });
    }
  }
  return blocos;
}

/* ------------------------------------------- códigos, contrapartes, históricos */
function listaDuasColunas(XLSX, ws, r0, rFim, cNome, cValor, nomeAba) {
  const out = [];
  for (let r = r0; r <= rFim; r++) {
    const nome = txt(XLSX, ws, r, cNome), cod = txt(XLSX, ws, r, cValor);
    if (!nome || !/^\d{6,8}$/.test(cod)) continue;
    out.push({ nome, codigo: cod, origem: `${nomeAba}!${ender(XLSX, r, cValor)}` });
  }
  return out;
}

function contrapartesDaAba(XLSX, ws, nomeAba) {
  const ref = intervalo(XLSX, ws);
  const res = { inter: [], bancos: [] };
  let sec = null;
  for (let r = 0; r <= ref.e.r; r++) {
    const a = txt(XLSX, ws, r, 2), b = txt(XLSX, ws, r, 3);
    if (/^contas inter empresas/i.test(a) || /^conta entre empresas/i.test(a)) { sec = "inter"; continue; }
    if (/^contas bancos/i.test(a)) { sec = "bancos"; continue; }
    if (!sec || !a || /^(empresa|banco)$/i.test(a)) continue;
    if (/^\d{6,8}$/.test(b)) {
      res[sec].push({ nome: a, codigo: b, origem: `${nomeAba}!${ender(XLSX, r, 3)}` });
    }
  }
  return res;
}

function historicosDaAba(XLSX, ws, nomeAba) {
  const ref = intervalo(XLSX, ws);
  const out = [];
  let grupo = "";
  for (let r = 2; r <= ref.e.r; r++) {
    const t = txt(XLSX, ws, r, 5);
    if (!t || t === "-" || /^hist[oó]rico$/i.test(t) || /^movimento de caixa/i.test(t)) continue;
    if (/^transfer[eê]ncia/i.test(t)) out.push({ grupo, texto: t, origem: `${nomeAba}!${ender(XLSX, r, 5)}` });
    else grupo = t;
  }
  return out;
}

/* ---------------------------------------------------- CNPJs e empreendimentos */
function lerCnpjs(XLSX, ws) {
  const out = [];
  if (!ws) return out;
  const fim = intervalo(XLSX, ws).e.r;
  for (let r = 3; r <= fim; r++) {
    const nome = txt(XLSX, ws, r, 2), cnpj = txt(XLSX, ws, r, 3);
    if (nome && /\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/.test(cnpj)) {
      out.push({ nome, cnpj, municipio: txt(XLSX, ws, r, 4), origem: `CNPJs!${ender(XLSX, r, 3)}` });
    }
  }
  return out;
}

function lerEmpreendimentos(XLSX, ws) {
  const res = { empreendimentos: [], vendas: [] };
  if (!ws) return res;
  const fim = intervalo(XLSX, ws).e.r;
  let modo = null;
  for (let r = 2; r <= fim; r++) {
    const c = semAcento(txt(XLSX, ws, r, 2));
    if (c === "empreendimento") { modo = "emp"; continue; }
    if (c === "numero") { modo = "vendas"; continue; }
    const nome = txt(XLSX, ws, r, 2);
    if (!nome) continue;
    if (modo === "emp") {
      res.empreendimentos.push({ nome, cnpj: txt(XLSX, ws, r, 3), gerente: txt(XLSX, ws, r, 4),
        municipio: txt(XLSX, ws, r, 5), origem: `Empreendimentos!${ender(XLSX, r, 3)}` });
    } else if (modo === "vendas") {
      res.vendas.push({ numero: nome, nome: txt(XLSX, ws, r, 3), cnpj: txt(XLSX, ws, r, 4),
        origem: `Empreendimentos!${ender(XLSX, r, 4)}` });
    }
  }
  return res;
}

/* ------------------------------------------------------------- calendário */
/** Quebra "A + B (x + y) + C" em itens, ignorando o "+" que está dentro de parênteses. */
export function dividirItens(texto) {
  const itens = [];
  let prof = 0, atual = "";
  for (const ch of String(texto)) {
    if (ch === "(") prof++;
    if (ch === ")") prof = Math.max(0, prof - 1);
    if (ch === "+" && prof === 0) { itens.push(atual); atual = ""; continue; }
    atual += ch;
  }
  itens.push(atual);
  return itens.map(norm).filter(Boolean);
}

const REGEX_DIA_UTIL = /^(\d+)\s*[°ºo]?\s*dia\s*[uú]til\s*-\s*(.+)$/i;

function lerCalendario(XLSX, ws) {
  const res = { dias: {}, regras: [], origem: {} };
  if (!ws) return res;
  const merges = ws["!merges"] || [];
  for (const m of merges) {
    if (m.e.r - m.s.r !== 1 || m.s.c !== m.e.c) continue; // célula do dia: 2 linhas × 1 coluna
    const n = numero(XLSX, ws, m.s.r, m.s.c);
    if (n == null || !Number.isInteger(n) || n < 1 || n > 31) continue;
    let achado = "", onde = "";
    for (let r = m.e.r + 1; r <= m.e.r + 3 && !achado; r++) {
      for (let c = m.s.c - 2; c <= m.s.c && !achado; c++) {
        const t = txt(XLSX, ws, r, c);
        if (t) { achado = t; onde = ender(XLSX, r, c); }
      }
    }
    if (!achado) continue;
    for (const item of dividirItens(achado)) {
      const regra = REGEX_DIA_UTIL.exec(item);
      if (regra) {
        const n5 = Number(regra[1]), tx = norm(regra[2]);
        if (!res.regras.some((x) => x.n === n5 && x.texto === tx)) {
          res.regras.push({ n: n5, texto: tx, origem: `Pagamentos Mensais!${onde}` });
        }
      } else {
        (res.dias[n] ||= []).push({ texto: item, origem: `Pagamentos Mensais!${onde}` });
      }
    }
  }
  return res;
}

const CATEGORIAS = [
  { id: "imposto", rotulo: "Impostos", teste: /\b(pis|cofins|confins|darf|iss|fgts|dctfweb|guias?|imposto)\b/i },
  { id: "folha", rotulo: "Folha e subsídio", teste: /\b(folha|subs[ií]dio)\b/i },
  { id: "repasse", rotulo: "Repasses", teste: /repasses?|^mei p\//i },
  { id: "comissao", rotulo: "Comissões", teste: /comiss/i },
  { id: "contrato", rotulo: "Contratos", teste: /contratos?|ceripa/i },
  { id: "energia", rotulo: "Energia", teste: /elektro|cpfl/i },
  { id: "cambio", rotulo: "Câmbio e investimento", teste: /c[aâ]mbio|investimento/i },
  { id: "ordinario", rotulo: "Ordinários", teste: /ordin[aá]rios?/i },
];
export const ROTULOS_CATEGORIA = Object.fromEntries(
  [...CATEGORIAS, { id: "outros", rotulo: "Outros" }].map((c) => [c.id, c.rotulo]));
export function categoriaDe(texto) {
  return (CATEGORIAS.find((c) => c.teste.test(texto)) || { id: "outros" }).id;
}

/* ----------------------------------------------------------- dias úteis */
const pad = (n) => String(n).padStart(2, "0");
export const isoData = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export function daIso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function pascoa(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31), dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}
const somarDias = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);

/**
 * Feriados em que os bancos não funcionam: nacionais fixos, Carnaval (segunda e
 * terça), Sexta-feira Santa e Corpus Christi. NÃO inclui feriado estadual/municipal.
 */
export function feriadosBancarios(ano) {
  const p = pascoa(ano), f = new Map();
  const fixo = [[1, 1, "Confraternização Universal"], [4, 21, "Tiradentes"], [5, 1, "Dia do Trabalho"],
    [9, 7, "Independência"], [10, 12, "N. Sra. Aparecida"], [11, 2, "Finados"],
    [11, 15, "Proclamação da República"], [11, 20, "Consciência Negra"], [12, 25, "Natal"]];
  for (const [m, d, n] of fixo) f.set(isoData(new Date(ano, m - 1, d)), n);
  f.set(isoData(somarDias(p, -48)), "Carnaval");
  f.set(isoData(somarDias(p, -47)), "Carnaval");
  f.set(isoData(somarDias(p, -2)), "Sexta-feira Santa");
  f.set(isoData(somarDias(p, 60)), "Corpus Christi");
  return f;
}
export function motivoNaoUtil(d) {
  const w = d.getDay();
  if (w === 0) return "Domingo";
  if (w === 6) return "Sábado";
  return feriadosBancarios(d.getFullYear()).get(isoData(d)) || "";
}
export const ehDiaUtil = (d) => !motivoNaoUtil(d);
/** n-ésimo dia útil do mês (1-based). */
export function enesimoDiaUtil(ano, mes, n) {
  let cont = 0;
  for (let dia = 1; dia <= 31; dia++) {
    const d = new Date(ano, mes, dia);
    if (d.getMonth() !== mes) break;
    if (ehDiaUtil(d) && ++cont === n) return d;
  }
  return null;
}

/**
 * Monta o mês pedido: { [dia]: [{texto, categoria, origem, nota?}] }.
 * Os dias 29, 30 e 31 de um mês mais curto caem no último dia do mês; a regra
 * "N° dia útil" cai no N-ésimo dia útil real (a planilha a repete em 3 dias
 * porque ela varia de 5 a 7 conforme fins de semana e feriados).
 */
export function montarMes(dados, ano, mes) {
  const ultimo = new Date(ano, mes + 1, 0).getDate();
  const out = {};
  const por = (dia, item) => { (out[dia] ||= []).push(item); };
  for (const [diaStr, lista] of Object.entries(dados.calendario?.dias || {})) {
    const dia = Number(diaStr), alvo = Math.min(dia, ultimo);
    for (const it of lista) {
      por(alvo, { texto: it.texto, categoria: categoriaDe(it.texto), origem: it.origem,
        nota: dia > ultimo ? `Cadastrado para o dia ${dia}; este mês tem ${ultimo} dias.` : "" });
    }
  }
  for (const r of dados.calendario?.regras || []) {
    const d = enesimoDiaUtil(ano, mes, r.n);
    if (!d) continue;
    por(d.getDate(), { texto: r.texto, categoria: categoriaDe(r.texto), origem: r.origem,
      nota: `${r.n}º dia útil do mês`, rotuloDia: `${r.n}º dia útil` });
  }
  return out;
}

/* ------------------------------------------------- Pagamentos do Dia (modelo) */
function lerPagamentosDoDia(XLSX, ws) {
  const res = { empresas: [], vizinhas: [], transferencias: [], conciliacao: [], checklist: [], origem: {} };
  if (!ws) return res;
  const fim = intervalo(XLSX, ws).e.r;
  const NOME = "Pagamentos do Dia";
  // 1) tabela principal (cabeçalho "Empresa | Saídas | Numerários | Compromissos | ...")
  let lh = -1;
  for (let r = 0; r <= fim; r++) {
    if (semAcento(txt(XLSX, ws, r, 2)) === "empresa" && semAcento(txt(XLSX, ws, r, 3)) === "saidas") { lh = r; break; }
  }
  if (lh >= 0) {
    for (let r = lh + 1; r <= fim; r++) {
      const nome = txt(XLSX, ws, r, 2);
      if (!nome) break;
      const saidas = numero(XLSX, ws, r, 3) ?? 0, num = numero(XLSX, ws, r, 4) ?? 0,
        comp = numero(XLSX, ws, r, 5) ?? 0;
      res.empresas.push({
        nome, id: idEmpresa(nome),
        total: Math.round((saidas + num + comp) * 100) / 100, numerarios: num, compromissos: comp,
        outrosBancos: txt(XLSX, ws, r, 6) === "true", desconsiderados: txt(XLSX, ws, r, 7) === "true",
      });
    }
  }
  // 2) "Comunicação entre Empresas"
  let rc = -1;
  for (let r = 0; r <= fim; r++) if (/^comunica[cç][aã]o entre empresas/i.test(txt(XLSX, ws, r, 2))) { rc = r; break; }
  if (rc >= 0) {
    for (let r = rc + 2; r <= fim; r++) {
      const nome = txt(XLSX, ws, r, 2);
      if (!nome) { if (r > rc + 14) break; continue; }
      const com = [];
      for (let c = 3; c <= 9; c++) { const t = txt(XLSX, ws, r, c); if (t) com.push(t); }
      if (com.length) res.vizinhas.push({ nome, com });
    }
  }
  // 3) transferências a fazer: linha de nomes com 2 linhas numéricas e uma de diferença
  for (let r = 0; r <= fim; r++) {
    const nomes = [];
    for (let c = 9; c <= 22; c++) {
      const t = txt(XLSX, ws, r, c);
      if (t && numero(XLSX, ws, r + 1, c) != null && numero(XLSX, ws, r + 2, c) != null) nomes.push({ c, nome: t });
    }
    if (nomes.length >= 3 && numero(XLSX, ws, r + 3, nomes[0].c) != null) {
      for (const n of nomes) {
        res.transferencias.push({ nome: n.nome, id: idEmpresa(n.nome),
          aTransferir: numero(XLSX, ws, r + 1, n.c), transferido: numero(XLSX, ws, r + 2, n.c) });
      }
      break;
    }
  }
  // 4) conciliação: "EMPRESA | KASIL | ..." seguida de "Fluxo (banco)", "Relatório", "Diferença"
  for (let r = 0; r <= fim; r++) {
    if (semAcento(txt(XLSX, ws, r, 2)) !== "empresa") continue;
    if (!/^fluxo/i.test(txt(XLSX, ws, r + 1, 2))) continue;
    for (let c = 3; c <= 12; c++) {
      const nome = txt(XLSX, ws, r, c);
      if (!nome) break;
      const justif = [];
      for (let k = r + 4; k <= r + 7; k++) {
        const v = numero(XLSX, ws, k, c);
        if (v) justif.push({ descricao: "", valor: v });
      }
      res.conciliacao.push({ nome, id: idEmpresa(nome),
        fluxo: numero(XLSX, ws, r + 1, c), relatorio: numero(XLSX, ws, r + 2, c), justificativas: justif });
    }
    break;
  }
  // 5) checklist de fechamento ("banco ok", "relatório ok", ...)
  for (let r = 0; r <= fim; r++) {
    const t = txt(XLSX, ws, r, 2);
    if (/\sok$/i.test(t)) { const s = t.replace(/\s+ok$/i, ""); res.checklist.push(s.charAt(0).toUpperCase() + s.slice(1)); }
  }
  return res;
}

/* ----------------------------------------------------------- leitura geral */
export function lerPlanilha(XLSX, wb) {
  const abas = wb.SheetNames || [];
  const aviso = [];
  const temAba = (n) => abas.includes(n);
  const ws = (n) => wb.Sheets[n];
  const faltando = ["Pagamentos Mensais", "Pagamentos do Dia", "Todas as Contas", "CNPJs", "Empreendimentos"]
    .filter((n) => !temAba(n));
  const algumaEmpresa = ABAS_EMPRESA.some(([n]) => temAba(n));
  if (!algumaEmpresa && faltando.length > 3) {
    throw new Error("Esta não parece ser a planilha Financeiro (não encontrei as abas de empresas nem as de pagamentos).");
  }
  if (faltando.length) aviso.push(`Abas não encontradas: ${faltando.join(", ")}.`);

  const dados = { versao: 1, empresas: [], contas: [], codigos: [], contrapartes: {}, historicos: {},
    blocosTodas: [], cnpjs: [], empreendimentos: [], vendas: [], calendario: { dias: {}, regras: [] },
    dia: null, formulasFixas: [], erros: [] };

  for (const [aba, id] of ABAS_EMPRESA) {
    if (!temAba(aba)) { aviso.push(`Aba "${aba}" não encontrada.`); continue; }
    const w = ws(aba);
    const nome = txt(XLSX, w, 0, 2) || aba;
    const bloco = contasDaAba(XLSX, w, aba);
    dados.empresas.push({ id, aba, nome, tituloBloco: bloco.titulo, tituloEnder: bloco.tituloEnder });
    for (const l of bloco.linhas) dados.contas.push({ empresa: id, ...l });
    dados.contrapartes[id] = contrapartesDaAba(XLSX, w, aba);
    dados.historicos[id] = historicosDaAba(XLSX, w, aba);
  }

  if (temAba("Todas as Contas")) {
    const w = ws("Todas as Contas"), fim = intervalo(XLSX, w).e.r;
    dados.codigos = [
      ...listaDuasColunas(XLSX, w, 4, fim, 2, 3, "Todas as Contas").map((x) => ({ ...x, grupo: "Contas banco" })),
      ...listaDuasColunas(XLSX, w, 4, 13, 5, 6, "Todas as Contas").map((x) => ({ ...x, grupo: "Contas banco novas" })),
    ];
    dados.blocosTodas = blocosTodasAsContas(XLSX, w, "Todas as Contas");
  }

  dados.cnpjs = lerCnpjs(XLSX, ws("CNPJs"));
  const emp = lerEmpreendimentos(XLSX, ws("Empreendimentos"));
  dados.empreendimentos = emp.empreendimentos; dados.vendas = emp.vendas;
  dados.calendario = lerCalendario(XLSX, ws("Pagamentos Mensais"));
  dados.dia = lerPagamentosDoDia(XLSX, ws("Pagamentos do Dia"));

  // fórmulas com valor digitado dentro e células com erro, nas abas que o módulo usa
  const usadas = ["Pagamentos Mensais", "Pagamentos do Dia", "Todas as Contas", "CNPJs", "Empreendimentos",
    ...ABAS_EMPRESA.map(([n]) => n)].filter(temAba);
  for (const aba of usadas) {
    const w = ws(aba);
    for (const k of Object.keys(w)) {
      if (k[0] === "!") continue;
      const x = w[k];
      if (x.t === "e") dados.erros.push({ onde: `${aba}!${k}`, valor: x.w || "#ERRO" });
      else if (typeof x.f === "string" && /(^|[^A-Za-z$\d.])\d{3,}(\.\d+)?/.test(x.f.replace(/\$?[A-Z]{1,3}\$?\d+/g, ""))) {
        dados.formulasFixas.push({ onde: `${aba}!${k}`, formula: `=${x.f}` });
      }
    }
  }
  return { dados, aviso };
}

/* ------------------------------------------------------------- auditoria */
const fmtMoeda = (n) => n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const nomeEmpresa = (dados, id) => dados.empresas.find((e) => e.id === id)?.nome || id;

/**
 * Cruza as informações da própria planilha e devolve os achados.
 * Cada achado: { id, sev: "alta"|"media"|"baixa", titulo, onde, detalhe, impacto, sugestao }.
 */
export function auditar(dados) {
  const f = [];
  const add = (a) => f.push({ id: `A${f.length + 1}`, ...a });

  // 1) CNPJ com dígito verificador inválido, em qualquer lugar
  const todosCnpj = new Map(); // cnpj → [origens]
  const reg = (cnpj, origem) => {
    for (const m of String(cnpj || "").match(REGEX_CNPJ) || []) {
      if (!todosCnpj.has(m)) todosCnpj.set(m, []);
      todosCnpj.get(m).push(origem);
    }
  };
  dados.cnpjs.forEach((x) => reg(x.cnpj, x.origem));
  dados.empreendimentos.forEach((x) => reg(x.cnpj, x.origem));
  dados.vendas.forEach((x) => reg(x.cnpj, x.origem));
  dados.contas.forEach((x) => reg(x.cnpj, x.origem));
  for (const [cnpj, onde] of todosCnpj) {
    if (!validarCnpj(cnpj)) {
      add({ sev: "alta", titulo: `CNPJ com dígito verificador inválido: ${cnpj}`, onde: onde.slice(0, 3).join(", "),
        detalhe: "Os dois dígitos finais não batem com o cálculo oficial do CNPJ.",
        impacto: "Pagamento, boleto ou cadastro com CNPJ errado pode ser recusado ou ir para a empresa errada.",
        sugestao: "Conferir o número no cartão CNPJ e corrigir em todas as abas onde aparece." });
    }
  }

  // 2) CNPJ de empreendimento/venda que não bate com o das contas da empresa
  const cnpjsDaEmpresa = (id) => new Set(dados.contas.filter((c) => c.empresa === id).map((c) => c.cnpj).filter(Boolean));
  for (const v of dados.vendas) {
    const id = idEmpresa(v.nome);
    if (!id || !v.cnpj) continue;
    const doCad = cnpjsDaEmpresa(id);
    if (doCad.size && !doCad.has(v.cnpj)) {
      add({ sev: "alta", titulo: `CNPJ de ${v.nome} diferente entre abas`,
        onde: `${v.origem} e aba ${dados.empresas.find((e) => e.id === id)?.aba}`,
        detalhe: `Em Empreendimentos consta ${v.cnpj}; nas contas da empresa consta ${[...doCad].join(" / ")}. Os dois são CNPJs válidos, então a divergência não é de digitação óbvia.`,
        impacto: "Relatórios, comprovantes e conferências por CNPJ podem apontar para a entidade errada.",
        sugestao: "Confirmar qual é o CNPJ correto da empresa e atualizar a aba que está errada." });
    }
  }

  // 3) códigos de conta repetidos entre nomes diferentes
  const porCodigo = new Map();
  for (const c of dados.codigos) (porCodigo.get(c.codigo) || porCodigo.set(c.codigo, []).get(c.codigo)).push(c);
  for (const [codigo, lista] of porCodigo) {
    if (lista.length < 2) continue;
    const nomes = lista.map((x) => x.nome);
    const base = new Set(nomes.map((n) => semAcento(n).replace(/\bnova?\b/g, "").trim()));
    const renomeacao = base.size === 1;
    add({ sev: renomeacao ? "baixa" : "media",
      titulo: `Código de conta ${codigo} usado por ${nomes.join(" e ")}`,
      onde: lista.map((x) => x.origem).join(", "),
      detalhe: renomeacao ? "O mesmo código aparece com dois nomes que parecem a mesma conta (renomeação)."
        : "O mesmo código está cadastrado para empresas diferentes.",
      impacto: renomeacao ? "Baixo: só confunde quem consulta a lista."
        : "Lançamento escolhendo pelo código pode cair na conta da outra empresa.",
      sugestao: renomeacao ? "Manter um só nome para o código." : "Confirmar qual é a conta correta de cada empresa e corrigir a que está repetida." });
  }

  // 4) título do bloco de contas diferente da aba
  for (const e of dados.empresas) {
    const t = e.tituloBloco && idEmpresa(e.tituloBloco.replace(/^contas\s+/i, ""));
    if (e.tituloBloco && t && t !== e.id) {
      add({ sev: "media", titulo: `Aba ${e.aba}: bloco intitulado "${e.tituloBloco}"`, onde: e.tituloEnder,
        detalhe: `As contas listadas nessa aba têm o CNPJ de ${nomeEmpresa(dados, e.id)}, mas o título aponta para outra empresa (provável cópia de outra aba).`,
        impacto: "Quem confia no título pode usar uma conta da empresa errada.",
        sugestao: "Corrigir o título do bloco." });
    }
  }

  // 5) mesma conta bancária em empresas diferentes
  const porChave = new Map();
  for (const c of dados.contas) (porChave.get(chaveConta(c)) || porChave.set(chaveConta(c), []).get(chaveConta(c))).push(c);
  for (const lista of porChave.values()) {
    const ids = [...new Set(lista.map((x) => x.empresa))];
    if (ids.length > 1) {
      add({ sev: "alta", titulo: `Mesma conta (${lista[0].banco} ag. ${lista[0].agencia} c/c ${lista[0].conta}) em mais de uma empresa`,
        onde: lista.map((x) => x.origem).join(", "),
        detalhe: `Aparece em: ${ids.map((i) => nomeEmpresa(dados, i)).join(", ")}.`,
        impacto: "Uma conta só pertence a um CNPJ. Pagamento ou conciliação podem ser atribuídos à empresa errada.",
        sugestao: "Manter a conta só na empresa titular." });
    }
  }

  // 6) contas da aba da empresa x aba "Todas as Contas"
  for (const bloco of dados.blocosTodas) {
    if (!bloco.id) continue;
    const daAba = new Set(dados.contas.filter((c) => c.empresa === bloco.id).map(chaveConta));
    const dosBlocos = new Set(bloco.linhas.map((l) => chaveConta(l)));
    for (const l of bloco.linhas) {
      if (!daAba.has(chaveConta(l))) {
        add({ sev: "media", titulo: `${l.banco} ${l.agencia} c/c ${l.conta} só consta em "Todas as Contas"`, onde: l.origem,
          detalhe: `Não encontrei essa conta na aba de ${nomeEmpresa(dados, bloco.id)}.`,
          impacto: "As duas listas deveriam ser iguais; uma delas está desatualizada.",
          sugestao: "Incluir na aba da empresa ou retirar de Todas as Contas, conforme o caso." });
      }
    }
    for (const c of dados.contas.filter((x) => x.empresa === bloco.id)) {
      if (!dosBlocos.has(chaveConta(c))) {
        add({ sev: "media", titulo: `${c.banco} ${c.agencia} c/c ${c.conta} só consta na aba ${dados.empresas.find((e) => e.id === bloco.id)?.aba}`,
          onde: c.origem, detalhe: `Não encontrei essa conta no bloco "${bloco.titulo}" de Todas as Contas.`,
          impacto: "As duas listas deveriam ser iguais; uma delas está desatualizada.",
          sugestao: "Incluir em Todas as Contas ou retirar da aba da empresa, conforme o caso." });
      }
    }
  }

  // 7) vínculos entre empresas só em um sentido
  const vinc = {};
  for (const [id, cp] of Object.entries(dados.contrapartes)) {
    vinc[id] = new Set(cp.inter.map((x) => idEmpresa(x.nome)).filter(Boolean));
  }
  for (const [a, set] of Object.entries(vinc)) {
    for (const b of set) {
      if (b === a || !vinc[b] || vinc[b].has(a)) continue;
      add({ sev: "baixa", titulo: `${nomeEmpresa(dados, a)} lista ${nomeEmpresa(dados, b)}, mas a aba de ${nomeEmpresa(dados, b)} não lista ${nomeEmpresa(dados, a)}`,
        onde: `abas ${dados.empresas.find((e) => e.id === a)?.aba} e ${dados.empresas.find((e) => e.id === b)?.aba}`,
        detalhe: "Em \"Contas inter empresas\" o vínculo existe em um lado só.",
        impacto: "Quem usa a aba que não lista não encontra o código nem o histórico para lançar a transferência.",
        sugestao: "Incluir o vínculo na outra aba ou confirmar que a transferência não existe." });
    }
  }

  // 8) número repetido em Vendas
  const nums = new Map();
  for (const v of dados.vendas) (nums.get(v.numero) || nums.set(v.numero, []).get(v.numero)).push(v);
  for (const [n, lista] of nums) {
    if (lista.length > 1) {
      add({ sev: "media", titulo: `Número ${n} repetido em Empreendimentos (${lista.map((x) => x.nome).join(" e ")})`,
        onde: lista.map((x) => x.origem).join(", "),
        detalhe: "A coluna NÚMERO deveria identificar uma única entidade.",
        impacto: "Se o número for usado como chave no sistema, uma das empresas fica sem identificação própria.",
        sugestao: "Conferir o número correto de cada uma." });
    }
  }

  // 9) possível erro de digitação no calendário
  const susp = [];
  for (const lista of Object.values(dados.calendario.dias)) lista.forEach((i) => { if (/\bconfins\b/i.test(i.texto)) susp.push(i.origem); });
  if (susp.length) {
    add({ sev: "baixa", titulo: 'Termo "Confins" no calendário (provável "Cofins")', onde: [...new Set(susp)].join(", "),
      detalhe: `Aparece em ${susp.length} lançamentos.`, impacto: "Só leitura, mas atrapalha buscas e dá má impressão em relatórios.",
      sugestao: 'Trocar por "Cofins" na planilha.' });
  }

  // 10) valores digitados dentro de fórmulas e células com erro
  if (dados.formulasFixas.length) {
    add({ sev: "media", titulo: `${dados.formulasFixas.length} fórmula(s) com valor digitado dentro`,
      onde: dados.formulasFixas.slice(0, 6).map((x) => `${x.onde} (${x.formula})`).join("; "),
      detalhe: "Ex.: o total do relatório do dia está dentro da fórmula, em vez de numa célula própria.",
      impacto: "Sem rastreabilidade: ninguém vê de onde veio o número nem o que foi ajustado à mão (um \"+2000\" dentro da fórmula, por exemplo).",
      sugestao: "Separar o valor numa célula, com rótulo, e deixar a fórmula só referenciando. No módulo, o Pagamentos do dia já guarda cada parte separada." });
  }
  if (dados.erros.length) {
    add({ sev: "baixa", titulo: `${dados.erros.length} célula(s) com erro (${[...new Set(dados.erros.map((e) => e.valor))].join(", ")})`,
      onde: dados.erros.slice(0, 8).map((e) => e.onde).join(", "), detalhe: "Fórmulas quebradas nas abas lidas.",
      impacto: "Podem esconder um dado que deveria aparecer ali.", sugestao: "Corrigir ou apagar a fórmula." });
  }

  // 11) conciliação do dia que veio com diferença
  for (const c of dados.dia?.conciliacao || []) {
    if (c.fluxo != null && c.relatorio != null) {
      const dif = Math.round((c.fluxo - c.relatorio) * 100) / 100;
      const just = Math.round(c.justificativas.reduce((s, j) => s + j.valor, 0) * 100) / 100;
      const res = Math.round((dif - just) * 100) / 100;
      if (res !== 0) {
        add({ sev: "alta", titulo: `Conciliação de ${c.nome} com diferença de ${fmtMoeda(res)}`,
          onde: "Pagamentos do Dia",
          detalhe: `Fluxo (banco) ${fmtMoeda(c.fluxo)} contra relatório ${fmtMoeda(c.relatorio)}${just ? `, com ${fmtMoeda(just)} já justificados` : ", sem nenhuma justificativa registrada"}.`,
          impacto: "Valor que saiu do banco e não está no relatório (ou o contrário) sem explicação.",
          sugestao: "Localizar a solicitação que explica a diferença e registrá-la em Justificativas." });
      }
    }
  }
  const ordem = { alta: 0, media: 1, baixa: 2 };
  return f.sort((a, b) => ordem[a.sev] - ordem[b.sev]).map((x, i) => ({ ...x, id: `A${i + 1}` }));
}
