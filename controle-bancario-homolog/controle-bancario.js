/**
 * Controle Bancário — tela do módulo (NOCTUS, homologação).
 *
 * Por que tudo aqui roda no navegador: a planilha "Financeiro" tem agência, conta,
 * CNPJ e valores do dia, e o repositório do NOCTUS é público. Então a planilha
 * NÃO vai para o repositório nem para o servidor — a pessoa escolhe o .xlsx,
 * o módulo lê (SheetJS) e guarda o resultado só no localStorage deste navegador,
 * separado por usuário do NOCTUS (mesmo padrão da Conferência).
 */
import {
  semAcento, validarCnpj, categoriaDe, ROTULOS_CATEGORIA, isoData, daIso, motivoNaoUtil,
  montarMes, lerPlanilha, auditar,
} from "./planilha.js";

/* ============================================================== utilidades */
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const $ = (sel, raiz = document) => raiz.querySelector(sel);
const $$ = (sel, raiz = document) => [...raiz.querySelectorAll(sel)];
const round2 = (n) => Math.round(n * 100) / 100;
const fmt = (n) => (n == null || Number.isNaN(n) ? "" : n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const fmtR = (n) => (n == null ? "—" : n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }));
const pad = (n) => String(n).padStart(2, "0");
const MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
const SEMANA = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
const dataLonga = (d) => `${SEMANA[d.getDay()]}, ${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`;
const dataCurta = (d) => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const hojeIso = () => isoData(new Date());
const somarDias = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);

/** "1.234,56", "1234.56", "R$ 1.234,56", "(100,00)", "-5" → número; vazio → null; ilegível → NaN. */
function parseMoeda(s) {
  let t = String(s ?? "").replace(/[R$\s]/g, "");
  if (!t) return null;
  const neg = /^\(.*\)$/.test(t) || t.startsWith("-");
  t = t.replace(/[()-]/g, "");
  if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
  else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, "");
  else if (/^\d{1,3}\.\d{3}$/.test(t)) t = t.replace(".", ""); // "1.234" lido como mil duzentos e trinta e quatro (padrão BR)
  if (!/^\d+(\.\d+)?$/.test(t)) return NaN;
  return round2(neg ? -Number(t) : Number(t));
}

/** Realça (sem acento / sem caixa) o trecho `q` dentro de `texto`, já escapado. */
function destacar(texto, q) {
  const original = String(texto ?? "");
  const nq = semAcento(q);
  if (!nq) return esc(original);
  let plano = "";
  const mapa = [];
  for (let i = 0; i < original.length; i++) {
    const c = original[i].normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
    for (const ch of c) { plano += ch; mapa.push(i); }
  }
  const pos = plano.indexOf(nq);
  if (pos < 0) return esc(original);
  const ini = mapa[pos], fim = mapa[pos + nq.length - 1] + 1;
  return esc(original.slice(0, ini)) + "<mark>" + esc(original.slice(ini, fim)) + "</mark>" + esc(original.slice(fim));
}

/* ================================================================ armazenamento */
const CHAVE_SESSAO = "sipep.session";
function escopoUsuario() {
  try {
    const s = JSON.parse(localStorage.getItem(CHAVE_SESSAO) || "null");
    const base = String(s?.user?.username || s?.user?.id || "sem-login").trim().toLowerCase();
    return base.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40) || "sem-login";
  } catch { return "sem-login"; }
}
const ESCOPO = escopoUsuario();
const PREFIXO = `noctus.controlebancario.${ESCOPO}.`;
let armazenamentoOk = true;
function ler(k) {
  try { const v = localStorage.getItem(PREFIXO + k); return v ? JSON.parse(v) : null; } catch { return null; }
}
function gravar(k, v) {
  try { localStorage.setItem(PREFIXO + k, JSON.stringify(v)); return true; }
  catch { armazenamentoOk = false; return false; }
}
function apagar(k) { try { localStorage.removeItem(PREFIXO + k); } catch { /* sem armazenamento */ } }
function diasSalvos() {
  const out = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i) || "";
      if (k.startsWith(PREFIXO + "dia.")) out.push(k.slice((PREFIXO + "dia.").length));
    }
  } catch { /* sem armazenamento */ }
  return out.filter((x) => daIso(x)).sort().reverse();
}
function apagarTudo() {
  try {
    const ks = [];
    for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i) || ""; if (k.startsWith(PREFIXO)) ks.push(k); }
    ks.forEach((k) => localStorage.removeItem(k));
  } catch { /* sem armazenamento */ }
}

/* ===================================================================== estado */
const S = {
  aba: "painel",
  dados: ler("dados"), meta: ler("meta"), achados: [],
  cal: { ano: new Date().getFullYear(), mes: new Date().getMonth(), sel: null },
  contas: { sub: "empresa", emp: null, q: "", mapaSel: null, cnpjQ: "" },
  dia: { iso: hojeIso(), estado: null, confirmaTrazer: false },
  qual: { filtro: "todas" },
  imp: { previa: null, lendo: false, erro: "", confirmaRemover: false, aplicarDia: false, dataDia: hojeIso() },
  msg: null,
};
const cacheMes = new Map();
function recalcular() {
  cacheMes.clear();
  S.achados = S.dados ? auditar(S.dados) : [];
  if (S.dados && !S.contas.emp) S.contas.emp = S.dados.empresas[0]?.id || null;
}
const itensDoMes = (ano, mes) => {
  const k = `${ano}-${mes}`;
  if (!cacheMes.has(k)) cacheMes.set(k, S.dados ? montarMes(S.dados, ano, mes) : {});
  return cacheMes.get(k);
};
const itensDaData = (d) => itensDoMes(d.getFullYear(), d.getMonth())[d.getDate()] || [];
const nomeEmp = (id) => S.dados?.empresas.find((e) => e.id === id)?.nome || id;
const abaEmp = (id) => S.dados?.empresas.find((e) => e.id === id)?.aba || id;

/* ============================================================ blocos reutilizáveis */
function copiarBtn(txt, rotulo = "Copiar") {
  return `<button type="button" class="copiar" data-acao="copiar" data-txt="${esc(txt)}" title="Copiar ${esc(txt)}">${rotulo}</button>`;
}
async function copiarTexto(txt) {
  try { await navigator.clipboard.writeText(txt); return true; } catch { /* tenta o plano B */ }
  try {
    const ta = document.createElement("textarea");
    ta.value = txt; ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta); ta.select();
    const ok = document.execCommand("copy"); ta.remove(); return ok;
  } catch { return false; }
}
const sevRotulo = { alta: "Alta", media: "Média", baixa: "Baixa" };
const sevBadge = (s) => `<span class="sev sev--${s}">${sevRotulo[s]}</span>`;
const regexCnpj = /\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}/;
function cnpjBadge(c) {
  if (!c || !regexCnpj.test(c)) return "";
  return validarCnpj(c) ? "" : ' <span class="sev sev--alta" title="Os dígitos finais não batem com o cálculo oficial do CNPJ">dígito inválido</span>';
}

function avisoPrivacidade() {
  return `<div class="alerta info privado"><div><strong>Os dados ficam só neste navegador.</strong>
    A planilha é lida aqui mesmo e guardada no armazenamento local deste navegador, no seu usuário do NOCTUS.
    Ela não é enviada ao servidor nem ao GitHub (o repositório do NOCTUS é público). Quem usar este mesmo navegador e o mesmo usuário verá os dados;
    em outro computador ou navegador é preciso importar de novo.</div></div>`;
}
function vazioImportar(titulo = "Nenhuma planilha importada ainda") {
  return `<div class="cartao vazio-grande"><h2>${esc(titulo)}</h2>
    <p class="sub" style="max-width:560px;margin:0 auto 1.2rem">Importe a planilha <b>Financeiro</b> (.xlsx) para ver o calendário de pagamentos,
    as contas, os CNPJs e fazer o controle dos pagamentos do dia.</p>
    <button class="botao primario" data-acao="ir" data-aba="importar">Importar planilha</button></div>${avisoPrivacidade()}`;
}

/* ====================================================================== PAINEL */
function vPainel() {
  if (!S.dados) return vazioImportar();
  const d = S.dados, hoje = new Date();
  const bancos = new Set(d.contas.map((c) => c.codBanco || semAcento(c.banco)));
  const altas = S.achados.filter((a) => a.sev === "alta").length;
  const doMes = itensDoMes(hoje.getFullYear(), hoje.getMonth());
  const nMes = Object.values(doMes).reduce((s, l) => s + l.length, 0);
  const nu = motivoNaoUtil(hoje);
  const itensHoje = itensDaData(hoje);

  // próximos vencimentos (até 10 dias à frente, só dias com itens)
  const prox = [];
  for (let i = 1; i <= 10; i++) {
    const dt = somarDias(hoje, i), l = itensDaData(dt);
    if (l.length) prox.push({ dt, l });
  }
  // andamento do dia de hoje
  const est = ler("dia." + hojeIso());
  const etapas = d.dia?.checklist?.length ? d.dia.checklist : [];
  const feitas = est ? etapas.filter((e) => est.checks?.[e]).length : 0;
  const diverg = est ? contarDivergencias(est) : 0;

  const kpis = `<div class="kpis">
    <div class="kpi"><div class="r">Empresas</div><div class="v">${d.empresas.length}</div><div class="d">abas lidas da planilha</div></div>
    <div class="kpi"><div class="r">Contas bancárias</div><div class="v">${d.contas.length}</div><div class="d">em ${bancos.size} banco(s) diferentes</div></div>
    <div class="kpi"><div class="r">Lançamentos em ${MESES[hoje.getMonth()]}</div><div class="v">${nMes}</div><div class="d">no calendário de pagamentos</div></div>
    <div class="kpi ${altas ? "alerta-kpi" : "ok-kpi"}"><div class="r">Qualidade dos dados</div><div class="v">${S.achados.length}</div>
      <div class="d">${altas ? `${altas} de severidade alta` : "nenhum achado de severidade alta"}</div></div>
  </div>`;

  const hojeHtml = `<div class="cartao"><div class="cartao-topo"><h2>Hoje · ${esc(dataLonga(hoje))}</h2>
      <button class="botao pequeno" data-acao="ir-dia" data-iso="${hojeIso()}">Abrir pagamentos do dia</button></div>
    ${nu ? `<div class="alerta info">Hoje não é dia útil bancário (${esc(nu)}). Não considera feriado estadual ou municipal.</div>` : ""}
    ${itensHoje.length ? itensHoje.map(itemLinha).join("") : `<p class="sub">Nenhum pagamento recorrente cadastrado para hoje.</p>`}
    ${est ? `<p class="sub" style="margin:.8rem 0 0">Controle de hoje: <b>${feitas} de ${etapas.length}</b> etapas concluídas${diverg ? `, <span class="neg"><b>${diverg}</b> conciliação(ões) com diferença</span>` : ""}.</p>`
      : `<p class="sub" style="margin:.8rem 0 0">O controle de hoje ainda não foi iniciado.</p>`}
  </div>`;

  const proxHtml = `<div class="cartao"><div class="cartao-topo"><h2>Próximos 10 dias</h2>
      <button class="linkbtn" data-acao="ir" data-aba="calendario">Ver calendário</button></div>
    ${prox.length ? `<ul class="lista-simples">${prox.map(({ dt, l }) =>
      `<li><span><b>${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}</b> <span class="sub">${esc(SEMANA[dt.getDay()])}${motivoNaoUtil(dt) && !/^(Sábado|Domingo)$/.test(motivoNaoUtil(dt)) ? " · " + esc(motivoNaoUtil(dt)) : ""}</span></span>
        <span style="text-align:right">${l.slice(0, 3).map((i) => `<span class="chip cat-${i.categoria}" title="${esc(i.texto)}"><span>${esc(i.texto)}</span></span>`).join(" ")}${l.length > 3 ? ` <span class="cal-mais">+${l.length - 3}</span>` : ""}</span></li>`).join("")}</ul>`
      : `<p class="sub">Nada cadastrado nos próximos 10 dias.</p>`}
  </div>`;

  const aviso = armazenamentoOk ? "" : `<div class="alerta erro">Este navegador não deixou gravar os dados. O módulo funciona nesta tela, mas ao fechar a página será preciso importar de novo.</div>`;
  const meta = S.meta ? `<p class="sub">Planilha em uso: <b>${esc(S.meta.arquivo)}</b>, importada em ${esc(new Date(S.meta.importadoEm).toLocaleString("pt-BR"))}.
    Para atualizar, importe a versão nova em <button class="linkbtn" data-acao="ir" data-aba="importar">Importar planilha</button>.</p>` : "";
  return `${aviso}${kpis}<div class="grade2">${hojeHtml}${proxHtml}</div>${meta}`;
}
function itemLinha(i) {
  return `<div class="item-linha cat-${i.categoria}"><div><strong>${esc(i.texto)}</strong>
    <span class="sub">${esc(ROTULOS_CATEGORIA[i.categoria] || "")}${i.nota ? " · " + esc(i.nota) : ""} · planilha: ${esc(i.origem)}</span></div></div>`;
}

/* ================================================================== CALENDÁRIO */
function vCalendario() {
  if (!S.dados) return vazioImportar();
  const { ano, mes } = S.cal;
  const itens = itensDoMes(ano, mes);
  const ultimo = new Date(ano, mes + 1, 0).getDate();
  const deslocamento = (new Date(ano, mes, 1).getDay() + 6) % 7; // semana começa na segunda
  const hoje = hojeIso();
  const usadas = new Set();
  let cel = "";
  for (let i = 0; i < deslocamento; i++) cel += `<div class="cal-dia fora" aria-hidden="true"></div>`;
  for (let n = 1; n <= ultimo; n++) {
    const dt = new Date(ano, mes, n), iso = isoData(dt), l = itens[n] || [], nu = motivoNaoUtil(dt);
    l.forEach((i) => usadas.add(i.categoria));
    const cls = ["cal-dia", nu ? "nao-util" : "", iso === hoje ? "hoje" : "", S.cal.sel === n ? "sel" : "", !l.length && iso !== hoje && S.cal.sel !== n ? "vazio" : ""].join(" ");
    cel += `<button type="button" class="${cls}" data-acao="cal-sel" data-dia="${n}" aria-label="${esc(dataLonga(dt))}${l.length ? ", " + l.length + " lançamento(s)" : ""}">
      <span class="cal-num"><span>${n}</span><small>${esc(nu || SEMANA[dt.getDay()].slice(0, 3))}</small></span>
      ${l.slice(0, 3).map((i) => `<span class="chip cat-${i.categoria}" title="${esc(i.texto)}"><span>${esc(i.texto)}</span></span>`).join("")}
      ${l.length > 3 ? `<span class="cal-mais">+${l.length - 3} mais</span>` : ""}</button>`;
  }
  const cab = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"].map((x) => `<div class="cal-cab">${x}</div>`).join("");
  const legenda = [...usadas].map((c) => `<span class="cat-${c}"><i></i>${esc(ROTULOS_CATEGORIA[c] || c)}</span>`).join("");

  let detalhe = "";
  if (S.cal.sel) {
    const dt = new Date(ano, mes, S.cal.sel), l = itens[S.cal.sel] || [], nu = motivoNaoUtil(dt);
    detalhe = `<div class="cartao detalhe-dia"><div class="cartao-topo"><h2>${esc(dataLonga(dt))}</h2>
      <button class="botao pequeno" data-acao="ir-dia" data-iso="${isoData(dt)}">Abrir pagamentos deste dia</button></div>
      ${nu ? `<div class="alerta info">${esc(nu)} — sem expediente bancário. Os vencimentos que caem aqui costumam ser pagos no dia útil anterior ou seguinte, conforme a regra de cada um (a planilha não informa qual).</div>` : ""}
      ${l.length ? l.map(itemLinha).join("") : `<p class="sub">Nenhum lançamento cadastrado para este dia.</p>`}</div>`;
  }
  const total = Object.values(itens).reduce((s, l) => s + l.length, 0);
  return `<div class="cartao">
    <div class="cal-nav">
      <button class="botao" data-acao="cal-mes" data-d="-1" aria-label="Mês anterior">‹</button>
      <h2>${MESES[mes]} de ${ano}</h2>
      <button class="botao" data-acao="cal-mes" data-d="1" aria-label="Próximo mês">›</button>
      <button class="botao pequeno" data-acao="cal-hoje">Hoje</button>
      <span class="sub">${total} lançamento(s)</span><span style="flex:1"></span>
      <button class="botao pequeno" data-acao="cal-copiar">Copiar agenda do mês</button>
    </div>
    <div class="cal-grade">${cab}${cel}</div>
    <div class="legenda">${legenda}<span class="mudo">Dias listrados: fim de semana ou feriado nacional / Carnaval / Corpus Christi.</span></div>
  </div>${detalhe}`;
}
function agendaDoMes() {
  const { ano, mes } = S.cal, itens = itensDoMes(ano, mes);
  const linhas = [`Pagamentos de ${MESES[mes]} de ${ano}`];
  for (const n of Object.keys(itens).map(Number).sort((a, b) => a - b)) {
    const dt = new Date(ano, mes, n);
    linhas.push(`${pad(n)}/${pad(mes + 1)} (${SEMANA[dt.getDay()]}${motivoNaoUtil(dt) ? ", " + motivoNaoUtil(dt) : ""}): ${itens[n].map((i) => i.texto).join("; ")}`);
  }
  return linhas.join("\n");
}

/* ================================================================ CONTAS E CNPJs */
const ALIAS = {
  momentum: ["mei", "momentum"], "praia-verde": ["pv", "praia verde"], "pick-money": ["pick money", "pkm"],
  realiza: ["realiza", "modo"], mmh: ["mmh"], m3: ["m3"], m5: ["m5"], rvm: ["rvm"], irm: ["irm"],
  kasil: ["kasil"], slim: ["slim"], posto: ["posto"], abrasma: ["abrasma"],
};
const citaEmpresa = (texto, id) => {
  const t = semAcento(texto);
  return (ALIAS[id] || [semAcento(id)]).some((a) => new RegExp(`(^|[^a-z0-9])${a}([^a-z0-9]|$)`).test(t));
};

function vContas() {
  if (!S.dados) return vazioImportar();
  const subs = [["empresa", "Por empresa"], ["mapa", "Mapa de transferências"], ["cnpj", "CNPJs e empreendimentos"]];
  return `<div class="cartao"><div class="barra">
      <div class="subabas" role="tablist" style="margin:0">${subs.map(([k, r]) =>
        `<button class="subaba" role="tab" data-acao="sub" data-sub="${k}" aria-selected="${S.contas.sub === k}">${r}</button>`).join("")}</div>
      <span class="espaco"></span>
      <div style="min-width:240px;flex:1;max-width:360px"><input type="search" id="busca-contas" placeholder="Buscar conta, código, CNPJ, histórico…" value="${esc(S.contas.q)}" autocomplete="off"></div>
    </div></div>
    <div id="contas-corpo">${corpoContas()}</div>`;
}
function corpoContas() {
  if (S.contas.q.trim().length >= 2) return resultadosBusca(S.contas.q.trim());
  if (S.contas.sub === "mapa") return vMapa();
  if (S.contas.sub === "cnpj") return vCnpjs();
  return vEmpresa();
}

function vEmpresa() {
  const d = S.dados, id = S.contas.emp, emp = d.empresas.find((e) => e.id === id) || d.empresas[0];
  if (!emp) return `<div class="cartao"><p class="sub">Nenhuma empresa lida.</p></div>`;
  const lista = d.empresas.map((e) => {
    const n = d.contas.filter((c) => c.empresa === e.id).length;
    return `<button class="emp-btn" data-acao="emp" data-id="${esc(e.id)}" aria-pressed="${e.id === emp.id}">${esc(e.nome)}<small>${n}</small></button>`;
  }).join("");
  const contas = d.contas.filter((c) => c.empresa === emp.id);
  const cp = d.contrapartes[emp.id] || { inter: [], bancos: [] };
  const hist = d.historicos[emp.id] || [];
  const grupos = [...new Set(hist.map((h) => h.grupo))];
  const tabelaCod = (titulo, l) => l.length ? `<h3>${titulo}</h3><div class="tabela-wrap"><table><thead><tr><th>Nome</th><th>Código</th><th></th></tr></thead><tbody>
      ${l.map((x) => `<tr><td>${esc(x.nome)}</td><td class="mono">${esc(x.codigo)}</td><td class="num">${copiarBtn(x.codigo)}</td></tr>`).join("")}</tbody></table></div>` : "";
  return `<div class="contas-layout"><div class="empresas-lista">${lista}</div><div>
    <div class="cartao"><div class="cartao-topo"><h2>${esc(emp.nome)}</h2><span class="sub">aba “${esc(emp.aba)}” · ${contas.length} conta(s)</span></div>
    ${contas.length ? `<div class="tabela-wrap"><table><thead><tr><th>Banco</th><th>Agência</th><th>Conta</th><th>CNPJ</th><th>Obs.</th></tr></thead><tbody>
      ${contas.map((c) => `<tr><td>${esc(c.banco)}${c.codBanco ? ` <span class="mudo">(${esc(c.codBanco)})</span>` : ""}</td>
        <td class="mono">${esc(c.agencia)} ${c.agencia ? copiarBtn(c.agencia) : ""}</td>
        <td class="mono">${esc(c.conta)} ${c.conta ? copiarBtn(c.conta) : ""}</td>
        <td class="mono">${esc(c.cnpj)}${cnpjBadge(c.cnpj)} ${c.cnpj ? copiarBtn(c.cnpj) : ""}</td>
        <td class="sub">${esc(c.obs)}</td></tr>`).join("")}</tbody></table></div>` : `<p class="sub">Nenhuma conta lida nesta aba.</p>`}
    </div>
    ${(cp.inter.length || cp.bancos.length) ? `<div class="cartao"><h2>Códigos de contrapartida</h2>
      <div class="grade2"><div>${tabelaCod("Inter empresas", cp.inter) || '<p class="sub">Sem vínculos entre empresas.</p>'}</div><div>${tabelaCod("Bancos", cp.bancos)}</div></div></div>` : ""}
    ${hist.length ? `<div class="cartao"><h2>Textos de histórico</h2>
      ${grupos.map((g) => `<div class="hist-grupo">${esc(g)}</div>${hist.filter((h) => h.grupo === g).map((h) =>
        `<div class="hist-item"><span>${esc(h.texto)}</span>${copiarBtn(h.texto)}</div>`).join("")}`).join("")}</div>` : ""}
  </div></div>`;
}

/** Mapa circular dos vínculos "Contas inter empresas". Tracejado âmbar = só um lado lista o outro. */
function vMapa() {
  const d = S.dados, emps = d.empresas, n = emps.length;
  const W = 720, H = 560, cx = W / 2, cy = H / 2, R = 215;
  const pos = {};
  emps.forEach((e, i) => { const a = (i / n) * Math.PI * 2 - Math.PI / 2; pos[e.id] = { x: cx + R * Math.cos(a), y: cy + R * Math.sin(a) }; });
  const dir = new Set(); // "a>b"
  for (const e of emps) for (const x of d.contrapartes[e.id]?.inter || []) {
    const b = emps.find((o) => citaEmpresa(x.nome, o.id))?.id;
    if (b && b !== e.id) dir.add(`${e.id}>${b}`);
  }
  const pares = new Map();
  for (const k of dir) { const [a, b] = k.split(">"); const chave = [a, b].sort().join("|"); pares.set(chave, (pares.get(chave) || 0) + 1); }
  const sel = S.contas.mapaSel;
  const arestas = [...pares.entries()].map(([chave, qtd]) => {
    const [a, b] = chave.split("|"), p = pos[a], q = pos[b];
    const ativa = sel && (a === sel || b === sel);
    return `<line class="aresta ${qtd === 1 ? "uma-mao" : ""} ${sel ? (ativa ? "forte" : "dim") : ""}" x1="${p.x.toFixed(1)}" y1="${p.y.toFixed(1)}" x2="${q.x.toFixed(1)}" y2="${q.y.toFixed(1)}"/>`;
  }).join("");
  const grau = (id) => [...dir].filter((k) => k.startsWith(id + ">")).length;
  const vizinho = (id) => sel && (id === sel || dir.has(`${sel}>${id}`) || dir.has(`${id}>${sel}`));
  const nos = emps.map((e) => {
    const p = pos[e.id];
    return `<g class="no ${sel === e.id ? "sel" : ""} ${sel && !vizinho(e.id) ? "dim" : ""}" data-acao="no" data-id="${esc(e.id)}" tabindex="0" role="button" aria-label="${esc(e.nome)}">
      <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="32"/><text x="${p.x.toFixed(1)}" y="${(p.y + 1).toFixed(1)}">${esc(e.aba.length > 9 ? e.aba.slice(0, 8) + "…" : e.aba)}</text>
      <text class="s" x="${p.x.toFixed(1)}" y="${(p.y + 14).toFixed(1)}">${grau(e.id)} vínculo(s)</text></g>`;
  }).join("");

  let painel = `<p class="sub" style="margin:.8rem 0 0">Clique numa empresa para ver os vínculos e os textos de histórico que citam as duas.</p>`;
  if (sel) {
    const a = sel, saidas = emps.filter((o) => dir.has(`${a}>${o.id}`)), entradas = emps.filter((o) => dir.has(`${o.id}>${a}`));
    const outros = [...new Set([...saidas, ...entradas].map((o) => o.id))];
    painel = `<h3>${esc(nomeEmp(a))}</h3>${outros.length ? `<div class="tabela-wrap"><table><thead><tr><th>Com</th><th>Vínculo</th><th>Código na aba de ${esc(abaEmp(a))}</th><th>Código na aba da outra</th></tr></thead><tbody>
      ${outros.map((b) => {
        const ab = dir.has(`${a}>${b}`), ba = dir.has(`${b}>${a}`);
        const cod = (de, para) => (S.dados.contrapartes[de]?.inter || []).find((x) => citaEmpresa(x.nome, para))?.codigo || "";
        const c1 = cod(a, b), c2 = cod(b, a);
        return `<tr><td>${esc(nomeEmp(b))}</td><td>${ab && ba ? '<span class="pilula verde">nos dois lados</span>' : `<span class="pilula" style="color:var(--color-warn)">só ${ab ? esc(abaEmp(a)) : esc(abaEmp(b))} lista</span>`}</td>
          <td class="mono">${esc(c1)} ${c1 ? copiarBtn(c1) : ""}</td><td class="mono">${esc(c2)} ${c2 ? copiarBtn(c2) : ""}</td></tr>`;
      }).join("")}</tbody></table></div>` : `<p class="sub">Esta empresa não tem vínculos listados.</p>`}
      ${outros.map((b) => {
        const hs = [...(S.dados.historicos[a] || []), ...(S.dados.historicos[b] || [])]
          .filter((h) => citaEmpresa(h.texto, a) && citaEmpresa(h.texto, b));
        const unicos = [...new Map(hs.map((h) => [h.texto, h])).values()];
        return unicos.length ? `<div class="hist-grupo">${esc(abaEmp(a))} × ${esc(abaEmp(b))}</div>${unicos.map((h) =>
          `<div class="hist-item"><span>${esc(h.texto)}</span>${copiarBtn(h.texto)}</div>`).join("")}` : "";
      }).join("")}`;
  }
  return `<div class="cartao"><div class="cartao-topo"><h2>Transferências entre empresas</h2>
    <span class="sub">linha contínua: os dois lados listam · tracejada: só um lado lista</span></div>
    <svg class="mapa" viewBox="0 0 ${W} ${H}" role="img" aria-label="Mapa de vínculos entre empresas">${arestas}${nos}</svg>${painel}</div>`;
}

function vCnpjs() {
  return `<div class="cartao"><div class="busca" style="max-width:420px"><input type="search" id="busca-cnpj" placeholder="Filtrar por nome, CNPJ, gerente, município…" value="${esc(S.contas.cnpjQ)}" autocomplete="off"></div>
    <div id="cnpj-corpo">${corpoCnpj()}</div></div>`;
}
function tabelasCnpj(cn, em, ve) {
  const linhaCnpj = (c) => `<span class="mono">${esc(c)}${cnpjBadge(c)}</span> ${c ? copiarBtn(c) : ""}`;
  return `<h3>CNPJs (${cn.length})</h3><div class="tabela-wrap"><table><thead><tr><th>Nome</th><th>CNPJ</th><th>Município</th></tr></thead><tbody>
      ${cn.map((x) => `<tr><td>${esc(x.nome)}</td><td>${linhaCnpj(x.cnpj)}</td><td class="sub">${esc(x.municipio)}</td></tr>`).join("") || `<tr><td colspan="3" class="sub">Nada encontrado.</td></tr>`}</tbody></table></div>
    <h3>Empreendimentos (${em.length})</h3><div class="tabela-wrap"><table><thead><tr><th>Nome</th><th>CNPJ</th><th>Gerente</th><th>Município</th></tr></thead><tbody>
      ${em.map((x) => `<tr><td>${esc(x.nome)}</td><td>${linhaCnpj(x.cnpj)}</td><td>${esc(x.gerente)}</td><td class="sub">${esc(x.municipio)}</td></tr>`).join("") || `<tr><td colspan="4" class="sub">Nada encontrado.</td></tr>`}</tbody></table></div>
    <h3>Vendas — empresas numeradas (${ve.length})</h3><div class="tabela-wrap"><table><thead><tr><th>Nº</th><th>Nome</th><th>CNPJ</th></tr></thead><tbody>
      ${ve.map((x) => `<tr><td class="mono">${esc(x.numero)}</td><td>${esc(x.nome)}</td><td>${linhaCnpj(x.cnpj)}</td></tr>`).join("") || `<tr><td colspan="3" class="sub">Nada encontrado.</td></tr>`}</tbody></table></div>`;
}

function resultadosBusca(q) {
  const d = S.dados, nq = semAcento(q), dq = q.replace(/\D/g, "");
  const bate = (...campos) => {
    const t = campos.join(" ");
    return semAcento(t).includes(nq) || (dq.length >= 3 && t.replace(/\D/g, "").includes(dq));
  };
  const blocos = [];
  const add = (titulo, lista, fn) => { if (lista.length) blocos.push(`<h3>${titulo} (${lista.length})</h3>${lista.slice(0, 15).map(fn).join("")}${lista.length > 15 ? `<p class="sub">…e mais ${lista.length - 15}. Refine a busca.</p>` : ""}`); };
  const linha = (rotulo, texto, copia, extra = "") => `<div class="resultado-busca"><div class="barra"><span class="pilula">${esc(rotulo)}</span><span style="flex:1">${destacar(texto, q)}${extra}</span>${copia ? copiarBtn(copia) : ""}</div></div>`;
  add("Contas bancárias", d.contas.filter((c) => bate(nomeEmp(c.empresa), c.banco, c.agencia, c.conta, c.cnpj, c.obs)),
    (c) => linha(nomeEmp(c.empresa), `${c.banco} · ag. ${c.agencia} · c/c ${c.conta} · ${c.cnpj}`, c.conta));
  add("Códigos de conta", d.codigos.filter((c) => bate(c.nome, c.codigo)), (c) => linha(c.grupo, `${c.nome} — ${c.codigo}`, c.codigo));
  const inter = [];
  for (const e of d.empresas) for (const x of d.contrapartes[e.id]?.inter || []) if (bate(e.nome, x.nome, x.codigo)) inter.push({ e, x });
  add("Contrapartida inter empresas", inter, ({ e, x }) => linha(e.nome, `${x.nome} — ${x.codigo}`, x.codigo));
  const hs = [];
  for (const e of d.empresas) for (const h of d.historicos[e.id] || []) if (bate(h.texto)) hs.push({ e, h });
  add("Textos de histórico", hs, ({ e, h }) => linha(e.nome, h.texto, h.texto));
  add("CNPJs", d.cnpjs.filter((x) => bate(x.nome, x.cnpj, x.municipio)), (x) => linha("CNPJ", `${x.nome} — ${x.cnpj}`, x.cnpj, cnpjBadge(x.cnpj)));
  add("Empreendimentos", d.empreendimentos.filter((x) => bate(x.nome, x.cnpj, x.gerente, x.municipio)),
    (x) => linha("Empreend.", `${x.nome} — ${x.cnpj} · ${x.gerente} · ${x.municipio}`, x.cnpj, cnpjBadge(x.cnpj)));
  add("Vendas", d.vendas.filter((x) => bate(x.numero, x.nome, x.cnpj)), (x) => linha("Venda " + x.numero, `${x.nome} — ${x.cnpj}`, x.cnpj));
  return `<div class="cartao"><h2>Resultados para “${esc(q)}”</h2>${blocos.join("") || '<p class="sub">Nada encontrado.</p>'}</div>`;
}

/* ============================================================ PAGAMENTOS DO DIA */
const diaVazio = () => ({ v: 1, empresas: {}, transf: {}, concil: {}, checks: {}, extras: [] });
const chaveEmp = (e) => e.id || semAcento(e.nome).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
const estadoDia = () => (S.dia.estado ||= ler("dia." + S.dia.iso) || diaVazio());
const soma = (l, f) => round2(l.reduce((s, x) => s + (f(x) || 0), 0));
const saidaDe = (r = {}) => (r.total == null && r.numerarios == null && r.compromissos == null
  ? null : round2((r.total || 0) - (r.numerarios || 0) - (r.compromissos || 0)));
const difConcil = (c = {}) => (c.fluxo != null && c.relatorio != null ? round2(c.fluxo - c.relatorio) : null);
const resConcil = (c = {}) => { const d = difConcil(c); return d == null ? null : round2(d - soma(c.just || [], (j) => j.v)); };
function contarDivergencias(est) {
  return Object.values(est.concil || {}).filter((c) => { const r = resConcil(c); return r != null && r !== 0; }).length;
}
const linhasEmpresas = () => (S.dados.dia?.empresas?.length ? S.dados.dia.empresas.map((e) => ({ k: chaveEmp(e), nome: e.nome }))
  : S.dados.empresas.map((e) => ({ k: e.id, nome: e.nome })));
const linhasTransf = () => (S.dados.dia?.transferencias || []).map((e) => ({ k: chaveEmp(e), nome: e.nome }));
function chavesConcil(est) {
  const base = (S.dados.dia?.conciliacao || []).map((e) => ({ k: chaveEmp(e), nome: e.nome }));
  for (const k of est.extras || []) if (!base.some((b) => b.k === k)) base.push({ k, nome: nomeEmp(k) });
  return base;
}
const etapasDoDia = () => (S.dados.dia?.checklist?.length ? S.dados.dia.checklist : ["Banco", "Relatório", "Extrato", "Suprimento", "Comprovantes"]);
const campoValor = (campo, k, f, v, extra = "") =>
  `<input type="text" class="valor" inputmode="decimal" autocomplete="off" data-campo="${campo}" data-k="${esc(k)}" data-f="${f}" ${extra} value="${esc(fmt(v))}" placeholder="0,00" aria-label="${esc(f)} ${esc(k)}">`;
const classeNum = (n) => (n == null ? "" : n < 0 ? "neg" : "");

function vDia() {
  if (!S.dados) return vazioImportar();
  const dt = daIso(S.dia.iso), est = estadoDia(), etapas = etapasDoDia(), nu = motivoNaoUtil(dt);
  const feitas = etapas.filter((e) => est.checks[e]).length;
  const emps = linhasEmpresas(), transf = linhasTransf();
  const regs = emps.map((e) => est.empresas[e.k] || {});
  const salvos = diasSalvos();
  const nav = `<div class="dia-nav">
    <button class="botao" data-acao="dia-ant" aria-label="Dia anterior">‹</button>
    <input type="date" id="dia-data" value="${esc(S.dia.iso)}" aria-label="Data">
    <button class="botao" data-acao="dia-prox" aria-label="Próximo dia">›</button>
    <button class="botao pequeno" data-acao="dia-hoje">Hoje</button>
    <h2>${esc(dataLonga(dt))}</h2><span class="espaco" style="flex:1"></span>
    <span class="salvo" id="salvo">${ler("dia." + S.dia.iso) ? "Registro salvo neste navegador" : "Ainda sem registro"}</span>
    <button class="botao pequeno" data-acao="dia-copiar">Copiar resumo do dia</button></div>
    ${salvos.length ? `<p class="sub" style="margin-top:-.4rem">Dias com registro: ${salvos.slice(0, 12).map((i) => `<button class="linkbtn" data-acao="ir-dia" data-iso="${i}">${i.slice(8)}/${i.slice(5, 7)}</button>`).join(" · ")}</p>` : ""}
    ${nu ? `<div class="alerta info">${esc(nu)}: sem expediente bancário. Use esta data só se houver movimento mesmo assim.</div>` : ""}`;

  const check = `<div class="cartao"><div class="cartao-topo"><h2>Etapas do fechamento</h2><span class="sub" id="check-txt">${feitas} de ${etapas.length}</span></div>
    <div class="checklist">${etapas.map((e) => `<label class="check"><input type="checkbox" data-campo="check" data-k="${esc(e)}" ${est.checks[e] ? "checked" : ""}>${esc(e)}</label>`).join("")}</div>
    <div class="progresso"><i id="check-barra" style="width:${etapas.length ? Math.round((feitas / etapas.length) * 100) : 0}%"></i></div></div>`;

  const tabela = `<div class="cartao"><div class="cartao-topo"><h2>Pagamentos por empresa</h2>
      <span class="sub">Saídas = Total do relatório − Numerários − Compromissos</span></div>
    <div class="tabela-wrap"><table><thead><tr><th>Empresa</th><th class="num">Total do relatório</th><th class="num">Numerários</th><th class="num">Compromissos</th><th class="num">Saídas</th><th>Outros bancos</th><th>Desconsiderados</th></tr></thead><tbody>
    ${emps.map((e, i) => { const r = regs[i], s = saidaDe(r); return `<tr><td><b>${esc(e.nome)}</b></td>
      <td>${campoValor("emp", e.k, "total", r.total)}</td><td>${campoValor("emp", e.k, "numerarios", r.numerarios)}</td><td>${campoValor("emp", e.k, "compromissos", r.compromissos)}</td>
      <td class="num calc ${classeNum(s)}" data-out="saida" data-k="${esc(e.k)}">${s == null ? "—" : fmt(s)}</td>
      <td><input type="checkbox" data-campo="emp-flag" data-k="${esc(e.k)}" data-f="outros" ${r.outros ? "checked" : ""} aria-label="Outros bancos ${esc(e.nome)}"></td>
      <td><input type="checkbox" data-campo="emp-flag" data-k="${esc(e.k)}" data-f="desc" ${r.desc ? "checked" : ""} aria-label="Desconsiderados ${esc(e.nome)}"></td></tr>`; }).join("")}
    </tbody><tfoot><tr><td>Total</td>
      <td class="num" data-out="tot-total">${fmt(soma(regs, (r) => r.total))}</td><td class="num" data-out="tot-numerarios">${fmt(soma(regs, (r) => r.numerarios))}</td>
      <td class="num" data-out="tot-compromissos">${fmt(soma(regs, (r) => r.compromissos))}</td>
      <td class="num" data-out="tot-saida">${fmt(soma(regs, (r) => saidaDe(r)))}</td><td></td><td></td></tr></tfoot></table></div></div>`;

  const bloco = transf.length ? `<div class="cartao"><div class="cartao-topo"><h2>Transferências a fazer</h2>
    <span class="sub">Falta = A transferir − Transferido</span></div><div class="tabela-wrap"><table><thead><tr><th>Empresa</th><th class="num">A transferir</th><th class="num">Transferido</th><th class="num">Falta</th></tr></thead><tbody>
    ${transf.map((e) => { const r = est.transf[e.k] || {}; const f = (r.aTransferir == null && r.transferido == null) ? null : round2((r.aTransferir || 0) - (r.transferido || 0));
      return `<tr><td><b>${esc(e.nome)}</b></td><td>${campoValor("transf", e.k, "aTransferir", r.aTransferir)}</td><td>${campoValor("transf", e.k, "transferido", r.transferido)}</td>
      <td class="num calc ${f ? "neg" : f === 0 ? "pos" : ""}" data-out="falta" data-k="${esc(e.k)}">${f == null ? "—" : fmt(f)}</td></tr>`; }).join("")}
    </tbody></table></div>
    <p class="sub" style="margin:.7rem 0 0">Na planilha original estas linhas não têm rótulo; assumi “A transferir”, “Transferido” e “Falta”. Confirme se é isso.</p></div>` : "";

  const livres = S.dados.empresas.filter((e) => !chavesConcil(est).some((c) => c.k === e.id));
  const concil = `<div class="cartao"><div class="cartao-topo"><h2>Conciliação banco × relatório</h2>
    <span class="sub">Diferença = Fluxo do banco − Relatório · deve fechar em zero com as justificativas</span></div>
    <div class="concil-grade" id="concil-grade">${htmlConciliacoes(est)}</div>
    ${livres.length ? `<div class="barra" style="margin-top:1rem"><select id="concil-add" style="max-width:240px" aria-label="Empresa">${livres.map((e) => `<option value="${esc(e.id)}">${esc(e.nome)}</option>`).join("")}</select>
      <button class="botao pequeno" data-acao="concil-add">Adicionar empresa à conciliação</button></div>` : ""}</div>`;

  const trazer = S.dados.dia?.empresas?.length ? `<div class="barra" style="margin-top:-.3rem;margin-bottom:1rem">
    ${S.dia.confirmaTrazer ? `<span class="sub">Já há dados neste dia; os valores da planilha vão substituí-los.</span>
      <button class="botao pequeno perigo" data-acao="dia-trazer-sim">Substituir</button><button class="botao pequeno" data-acao="dia-trazer-nao">Cancelar</button>`
      : `<button class="botao pequeno" data-acao="dia-trazer" title="Copia para este dia os valores da aba Pagamentos do Dia da planilha importada">Trazer valores da planilha</button>`}</div>` : "";
  return `${nav}${check}${trazer}${tabela}${bloco}${concil}`;
}
function htmlConciliacoes(est) {
  return chavesConcil(est).map(({ k, nome }) => {
    const c = est.concil[k] || {}, dif = difConcil(c), res = resConcil(c);
    const cls = res == null ? "" : res === 0 ? "confere" : "diverge";
    const extra = (est.extras || []).includes(k);
    return `<div class="concil ${cls}" data-card="${esc(k)}"><header><strong>${esc(nome)}</strong><span data-out="status" class="pilula ${res === 0 ? "verde" : res ? "verm" : ""}">${statusConcil(res)}</span></header>
      <div class="linha-form"><label>Fluxo do banco</label>${campoValor("concil", k, "fluxo", c.fluxo)}</div>
      <div class="linha-form"><label>Relatório</label>${campoValor("concil", k, "relatorio", c.relatorio)}</div>
      <div class="linha-total"><span>Diferença</span><span data-out="dif" class="${classeNum(dif)}">${dif == null ? "—" : fmt(dif)}</span></div>
      <h3 style="margin:.9rem 0 .4rem">Justificativas</h3>
      ${(c.just || []).map((j, i) => `<div class="just"><input type="text" data-campo="just-d" data-k="${esc(k)}" data-i="${i}" value="${esc(j.d)}" placeholder="Descrição" aria-label="Descrição">
        <input type="text" class="valor" inputmode="decimal" data-campo="just-v" data-k="${esc(k)}" data-i="${i}" value="${esc(fmt(j.v))}" placeholder="0,00" aria-label="Valor">
        <button class="x" data-acao="just-rem" data-k="${esc(k)}" data-i="${i}" aria-label="Remover justificativa">✕</button></div>`).join("")}
      <div class="barra"><button class="botao pequeno" data-acao="just-add" data-k="${esc(k)}">+ Justificativa</button>${extra ? `<button class="linkbtn" data-acao="concil-rem" data-k="${esc(k)}">Remover empresa</button>` : ""}</div>
      <div class="linha-total"><span>Sem explicação</span><span data-out="res" class="${res ? "neg" : res === 0 ? "pos" : ""}">${res == null ? "—" : fmt(res)}</span></div></div>`;
  }).join("") || `<p class="sub">Nenhuma empresa na conciliação.</p>`;
}
const statusConcil = (res) => (res == null ? "preencher" : res === 0 ? "confere" : "diverge");

function atualizarCalculos(campo, k) {
  const est = estadoDia();
  const set = (sel, txt, cls) => { const el = $(sel); if (!el) return; el.textContent = txt; if (cls != null) el.className = cls; };
  if (campo === "emp") {
    const r = est.empresas[k] || {}, s = saidaDe(r);
    set(`[data-out="saida"][data-k="${CSS.escape(k)}"]`, s == null ? "—" : fmt(s), `num calc ${classeNum(s)}`);
    const regs = linhasEmpresas().map((e) => est.empresas[e.k] || {});
    set('[data-out="tot-total"]', fmt(soma(regs, (x) => x.total)));
    set('[data-out="tot-numerarios"]', fmt(soma(regs, (x) => x.numerarios)));
    set('[data-out="tot-compromissos"]', fmt(soma(regs, (x) => x.compromissos)));
    set('[data-out="tot-saida"]', fmt(soma(regs, (x) => saidaDe(x))));
  } else if (campo === "transf") {
    const r = est.transf[k] || {};
    const f = (r.aTransferir == null && r.transferido == null) ? null : round2((r.aTransferir || 0) - (r.transferido || 0));
    set(`[data-out="falta"][data-k="${CSS.escape(k)}"]`, f == null ? "—" : fmt(f), `num calc ${f ? "neg" : f === 0 ? "pos" : ""}`);
  } else if (campo === "concil" || campo === "just") {
    const c = est.concil[k] || {}, dif = difConcil(c), res = resConcil(c), card = $(`[data-card="${CSS.escape(k)}"]`);
    if (!card) return;
    card.className = `concil ${res == null ? "" : res === 0 ? "confere" : "diverge"}`;
    set(`[data-card="${CSS.escape(k)}"] [data-out="dif"]`, dif == null ? "—" : fmt(dif), classeNum(dif));
    set(`[data-card="${CSS.escape(k)}"] [data-out="res"]`, res == null ? "—" : fmt(res), res ? "neg" : res === 0 ? "pos" : "");
    set(`[data-card="${CSS.escape(k)}"] [data-out="status"]`, statusConcil(res), `pilula ${res === 0 ? "verde" : res ? "verm" : ""}`);
  }
}

let timerSalvar = null;
function agendarSalvar() {
  clearTimeout(timerSalvar);
  timerSalvar = setTimeout(salvarDia, 350);
}
function salvarDia() {
  clearTimeout(timerSalvar); timerSalvar = null;
  if (!S.dia.estado) return;
  const ok = gravar("dia." + S.dia.iso, S.dia.estado);
  const el = $("#salvo");
  if (el) el.textContent = ok ? `Salvo às ${new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}` : "Não consegui salvar neste navegador";
}
function mudarDia(iso) {
  salvarDia();
  S.dia.iso = iso; S.dia.estado = null; S.dia.confirmaTrazer = false;
}
function aplicarValoresDaPlanilha(est, dados) {
  const nz = (n) => (n ? n : null);
  est.empresas = {}; est.transf = {}; est.concil = {};
  for (const e of dados.dia?.empresas || []) {
    est.empresas[chaveEmp(e)] = { total: nz(e.total), numerarios: nz(e.numerarios), compromissos: nz(e.compromissos), outros: !!e.outrosBancos, desc: !!e.desconsiderados };
  }
  for (const t of dados.dia?.transferencias || []) est.transf[chaveEmp(t)] = { aTransferir: nz(t.aTransferir), transferido: nz(t.transferido) };
  est.extras = [];
  for (const c of dados.dia?.conciliacao || []) {
    est.concil[chaveEmp(c)] = { fluxo: c.fluxo, relatorio: c.relatorio, just: (c.justificativas || []).map((j) => ({ d: j.descricao, v: j.valor })) };
  }
}
function resumoDia() {
  const dt = daIso(S.dia.iso), est = estadoDia(), etapas = etapasDoDia(), L = [];
  L.push(`Pagamentos do dia — ${dataLonga(dt)}`);
  L.push(`Etapas: ${etapas.map((e) => `${e} ${est.checks[e] ? "✔" : "pendente"}`).join(" · ")}`);
  L.push("", "Empresa | Total | Numerários | Compromissos | Saídas");
  for (const e of linhasEmpresas()) {
    const r = est.empresas[e.k]; if (!r || (r.total == null && r.numerarios == null && r.compromissos == null)) continue;
    L.push(`${e.nome} | ${fmt(r.total || 0)} | ${fmt(r.numerarios || 0)} | ${fmt(r.compromissos || 0)} | ${fmt(saidaDe(r))}`);
  }
  const falt = linhasTransf().filter((e) => { const r = est.transf[e.k]; return r && (r.aTransferir || r.transferido); });
  if (falt.length) { L.push("", "Transferências (a transferir | transferido | falta)"); for (const e of falt) { const r = est.transf[e.k]; L.push(`${e.nome} | ${fmt(r.aTransferir || 0)} | ${fmt(r.transferido || 0)} | ${fmt(round2((r.aTransferir || 0) - (r.transferido || 0)))}`); } }
  const cc = chavesConcil(est).filter((c) => difConcil(est.concil[c.k]) != null);
  if (cc.length) {
    L.push("", "Conciliação");
    for (const { k, nome } of cc) {
      const c = est.concil[k], res = resConcil(c);
      L.push(`${nome}: fluxo ${fmt(c.fluxo)} · relatório ${fmt(c.relatorio)} · diferença ${fmt(difConcil(c))}${(c.just || []).length ? " · justificativas: " + c.just.map((j) => `${j.d || "sem descrição"} ${fmt(j.v || 0)}`).join("; ") : ""} · ${res === 0 ? "CONFERE" : "SEM EXPLICAÇÃO " + fmt(res)}`);
    }
  }
  return L.join("\n");
}

/* ================================================================== QUALIDADE */
function vQualidade() {
  if (!S.dados) return vazioImportar();
  const cont = { alta: 0, media: 0, baixa: 0 }; S.achados.forEach((a) => cont[a.sev]++);
  const f = S.qual.filtro, lista = S.achados.filter((a) => f === "todas" || a.sev === f);
  const btn = (k, r, n) => `<button class="subaba" data-acao="qual-filtro" data-f="${k}" aria-selected="${f === k}">${r} (${n})</button>`;
  return `<div class="cartao"><div class="cartao-topo"><h2>Qualidade dos dados da planilha</h2>
    <button class="botao pequeno" data-acao="qual-copiar">Copiar relatório</button></div>
    <p class="sub" style="margin-top:-.4rem">O módulo cruza a planilha com ela mesma (CNPJ, contas repetidas, códigos, abas que deveriam coincidir, fórmulas com valor digitado).
    Ele aponta o que não fecha; <b>quem decide qual lado está certo é você</b> — não há fonte externa para confirmar.</p>
    <div class="filtros">${btn("todas", "Todos", S.achados.length)}${btn("alta", "Alta", cont.alta)}${btn("media", "Média", cont.media)}${btn("baixa", "Baixa", cont.baixa)}</div>
    ${lista.length ? lista.map((a) => `<div class="achado sev-${a.sev}">${sevBadge(a.sev)} <span class="mudo">${esc(a.id)}</span><h4>${esc(a.titulo)}</h4>
      <dl><dt>Onde</dt><dd>${esc(a.onde)}</dd><dt>O quê</dt><dd>${esc(a.detalhe)}</dd><dt>Impacto</dt><dd>${esc(a.impacto)}</dd><dt>Sugestão</dt><dd>${esc(a.sugestao)}</dd></dl></div>`).join("")
      : `<div class="alerta ok">Nenhum achado nesta categoria.</div>`}</div>`;
}
function relatorioQualidade(lista) {
  return ["Qualidade dos dados — planilha Financeiro", `Gerado em ${new Date().toLocaleString("pt-BR")}`, ""]
    .concat(lista.flatMap((a) => [`[${sevRotulo[a.sev]}] ${a.id} — ${a.titulo}`, `  Onde: ${a.onde}`, `  O quê: ${a.detalhe}`, `  Impacto: ${a.impacto}`, `  Sugestão: ${a.sugestao}`, ""])).join("\n");
}

/* ===================================================================== IMPORTAR */
function contagens(d) {
  return { empresas: d.empresas.length, contas: d.contas.length, codigos: d.codigos.length, cnpjs: d.cnpjs.length, empreendimentos: d.empreendimentos.length,
    vendas: d.vendas.length, diasCalendario: Object.keys(d.calendario.dias).length, regras: d.calendario.regras.length };
}
function resumoGrade(c) {
  const it = [["Empresas", c.empresas], ["Contas bancárias", c.contas], ["Códigos de conta", c.codigos], ["CNPJs", c.cnpjs],
    ["Empreendimentos", c.empreendimentos], ["Empresas numeradas", c.vendas], ["Dias no calendário", c.diasCalendario + c.regras]];
  return `<div class="resumo-leitura">${it.map(([r, n]) => `<div><b>${n}</b><span>${r}</span></div>`).join("")}</div>`;
}
function vImportar() {
  const atual = S.dados ? `<div class="cartao"><div class="cartao-topo"><h2>Dados atuais</h2></div>
    <p class="sub">Planilha: <b>${esc(S.meta?.arquivo || "—")}</b>${S.meta ? ` · importada em ${esc(new Date(S.meta.importadoEm).toLocaleString("pt-BR"))}` : ""}</p>
    ${resumoGrade(contagens(S.dados))}
    <div class="barra">${S.imp.confirmaRemover
      ? `<span class="sub">Apaga a planilha e todos os registros de pagamentos do dia deste navegador. Não dá para desfazer.</span>
         <button class="botao pequeno perigo" data-acao="imp-remover-sim">Apagar tudo</button><button class="botao pequeno" data-acao="imp-remover-nao">Cancelar</button>`
      : `<button class="botao pequeno" data-acao="imp-remover">Remover dados deste navegador</button>`}</div></div>` : "";
  const p = S.imp.previa;
  const previa = p ? (() => {
    const al = { alta: 0, media: 0, baixa: 0 }; p.achados.forEach((a) => al[a.sev]++);
    const diaExiste = !!ler("dia." + S.imp.dataDia);
    return `<div class="cartao"><div class="cartao-topo"><h2>Conferir antes de importar</h2><span class="sub">${esc(p.nome)} · ${Math.round(p.tamanho / 1024)} KB</span></div>
      ${resumoGrade(contagens(p.dados))}
      ${p.aviso.map((a) => `<div class="alerta">${esc(a)}</div>`).join("")}
      <p class="sub">Achados de qualidade nesta planilha: <b>${al.alta}</b> alta, <b>${al.media}</b> média, <b>${al.baixa}</b> baixa (detalhes na aba Qualidade dos dados depois de importar).</p>
      ${p.dados.dia?.empresas?.length ? `<label class="check" style="margin:.4rem 0 .8rem"><input type="checkbox" id="imp-aplicar" ${S.imp.aplicarDia ? "checked" : ""}>
        Trazer também os valores da aba “Pagamentos do Dia” para o dia
        <input type="date" id="imp-data" value="${esc(S.imp.dataDia)}"></label>
        ${S.imp.aplicarDia && diaExiste ? `<div class="alerta">Já existe registro em ${esc(S.imp.dataDia.split("-").reverse().join("/"))}; ele será substituído.</div>` : ""}
        <p class="sub" style="margin-top:0">A planilha não diz a que data aqueles valores se referem, por isso a escolha é sua. Deixe desmarcado se eles forem de um dia antigo.</p>` : ""}
      <div class="barra"><button class="botao primario" data-acao="imp-confirmar">Importar</button><button class="botao" data-acao="imp-cancelar">Cancelar</button></div></div>`;
  })() : "";
  return `${atual}<div class="cartao"><h2>${S.dados ? "Atualizar com uma nova versão da planilha" : "Importar a planilha Financeiro"}</h2>
    <div class="solta" id="solta" data-acao="imp-escolher" tabindex="0" role="button"><strong>${S.imp.lendo ? "Lendo a planilha…" : "Clique para escolher o arquivo"}</strong>
      <span>Excel (.xlsx) — ou arraste o arquivo para cá. A aba “Rascunho” é ignorada.</span>
      <input type="file" id="arquivo-planilha" accept=".xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden></div>
    ${S.imp.erro ? `<div class="alerta erro" style="margin-top:1rem">${esc(S.imp.erro)}</div>` : ""}</div>${previa}${avisoPrivacidade()}`;
}
let promessaXlsx = null;
function carregarXLSX() {
  if (window.XLSX) return Promise.resolve();
  promessaXlsx ||= (async () => {
    for (const src of ["../conferencia/vendor/xlsx.full.min.js", "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"]) {
      try {
        await new Promise((ok, falha) => { const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = () => falha(new Error(src)); document.head.appendChild(s); });
        if (window.XLSX) return;
      } catch { /* tenta a próxima origem */ }
    }
    promessaXlsx = null;
    throw new Error("Não consegui carregar a biblioteca que lê planilhas. Verifique a conexão e tente de novo.");
  })();
  return promessaXlsx;
}
async function lerArquivo(file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { S.imp.erro = "Escolha um arquivo do Excel (.xlsx)."; render(); return; }
  S.imp.lendo = true; S.imp.erro = ""; S.imp.previa = null; render();
  try {
    await carregarXLSX();
    const wb = window.XLSX.read(await file.arrayBuffer(), { type: "array" });
    const { dados, aviso } = lerPlanilha(window.XLSX, wb);
    S.imp.previa = { nome: file.name, tamanho: file.size, dados, aviso, achados: auditar(dados) };
  } catch (e) {
    S.imp.erro = e?.message || "Não consegui ler este arquivo.";
  } finally { S.imp.lendo = false; render(); }
}
function confirmarImportacao() {
  const p = S.imp.previa; if (!p) return;
  S.dados = p.dados;
  S.meta = { arquivo: p.nome, tamanho: p.tamanho, importadoEm: new Date().toISOString(), contagem: contagens(p.dados) };
  const ok = gravar("dados", S.dados) && gravar("meta", S.meta);
  let extra = "";
  if (S.imp.aplicarDia && p.dados.dia?.empresas?.length) {
    const est = diaVazio(); aplicarValoresDaPlanilha(est, p.dados);
    if (gravar("dia." + S.imp.dataDia, est)) { extra = ` Valores do dia trazidos para ${S.imp.dataDia.split("-").reverse().join("/")}.`; if (S.dia.iso === S.imp.dataDia) S.dia.estado = est; }
  }
  S.imp.previa = null; S.imp.aplicarDia = false; S.contas.emp = null;
  recalcular();
  S.msg = ok ? { tipo: "ok", txt: `Planilha importada: ${S.dados.empresas.length} empresas e ${S.dados.contas.length} contas.${extra}` }
    : { tipo: "erro", txt: "A planilha foi lida, mas este navegador não deixou gravar. Ela vale só até fechar a página." };
  ir("painel");
}

/* ======================================================================= ROTEAMENTO */
const VISOES = { painel: vPainel, calendario: vCalendario, contas: vContas, dia: vDia, qualidade: vQualidade, importar: vImportar };
function render() {
  $$("#abas .aba").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.aba === S.aba)));
  const n = $("#n-qualidade");
  if (S.achados.length) { n.textContent = S.achados.length; n.classList.remove("oculto"); } else n.classList.add("oculto");
  const msg = S.msg ? `<div class="alerta ${S.msg.tipo || "ok"}">${esc(S.msg.txt)}</div>` : ""; S.msg = null;
  $("#app").innerHTML = msg + (VISOES[S.aba] || vPainel)();
}
function ir(aba) {
  if (S.aba === "dia" && aba !== "dia") salvarDia();
  S.aba = VISOES[aba] ? aba : "painel";
  try { history.replaceState(null, "", "#" + S.aba); } catch { /* sem history */ }
  render(); window.scrollTo(0, 0);
}
async function feedbackCopia(el, txt, rotuloOk = "Copiado") {
  const ok = await copiarTexto(txt);
  const antes = el.textContent;
  el.textContent = ok ? rotuloOk : "Não copiou"; if (ok) el.classList.add("ok");
  setTimeout(() => { el.textContent = antes; el.classList.remove("ok"); }, 1400);
}

document.addEventListener("click", (ev) => {
  const abaBtn = ev.target.closest("[data-aba].aba");
  if (abaBtn) { ir(abaBtn.dataset.aba); return; }
  const el = ev.target.closest("[data-acao]"); if (!el) return;
  const a = el.dataset.acao;
  if (a === "copiar") { feedbackCopia(el, el.dataset.txt); return; }
  if (a === "ir") return ir(el.dataset.aba);
  if (a === "ir-dia") { mudarDia(el.dataset.iso); return ir("dia"); }
  // calendário
  if (a === "cal-sel") { S.cal.sel = Number(el.dataset.dia); render(); $(".detalhe-dia")?.scrollIntoView({ block: "nearest", behavior: "smooth" }); return; }
  if (a === "cal-mes") { let m = S.cal.mes + Number(el.dataset.d), y = S.cal.ano; if (m < 0) { m = 11; y--; } if (m > 11) { m = 0; y++; } Object.assign(S.cal, { mes: m, ano: y, sel: null }); return render(); }
  if (a === "cal-hoje") { const h = new Date(); Object.assign(S.cal, { mes: h.getMonth(), ano: h.getFullYear(), sel: h.getDate() }); return render(); }
  if (a === "cal-copiar") return feedbackCopia(el, agendaDoMes());
  // contas
  if (a === "sub") { S.contas.sub = el.dataset.sub; S.contas.q = ""; return render(); }
  if (a === "emp") { S.contas.emp = el.dataset.id; return render(); }
  if (a === "no") { S.contas.mapaSel = S.contas.mapaSel === el.dataset.id ? null : el.dataset.id; $("#contas-corpo").innerHTML = corpoContas(); return; }
  // pagamentos do dia
  if (a === "dia-ant" || a === "dia-prox") { mudarDia(isoData(somarDias(daIso(S.dia.iso), a === "dia-ant" ? -1 : 1))); return render(); }
  if (a === "dia-hoje") { mudarDia(hojeIso()); return render(); }
  if (a === "dia-copiar") return feedbackCopia(el, resumoDia());
  if (a === "dia-trazer") {
    if (ler("dia." + S.dia.iso)) { S.dia.confirmaTrazer = true; return render(); }
    aplicarValoresDaPlanilha(estadoDia(), S.dados); salvarDia(); return render();
  }
  if (a === "dia-trazer-sim") { S.dia.confirmaTrazer = false; aplicarValoresDaPlanilha(estadoDia(), S.dados); salvarDia(); return render(); }
  if (a === "dia-trazer-nao") { S.dia.confirmaTrazer = false; return render(); }
  if (a === "concil-add") { const k = $("#concil-add")?.value; if (k) { const est = estadoDia(); (est.extras ||= []).push(k); est.concil[k] ||= {}; salvarDia(); render(); } return; }
  if (a === "concil-rem") { const est = estadoDia(), k = el.dataset.k; est.extras = (est.extras || []).filter((x) => x !== k); delete est.concil[k]; salvarDia(); return render(); }
  if (a === "just-add") { const est = estadoDia(), k = el.dataset.k; (est.concil[k] ||= {}).just ||= []; est.concil[k].just.push({ d: "", v: null }); agendarSalvar(); $("#concil-grade").innerHTML = htmlConciliacoes(est); return; }
  if (a === "just-rem") { const est = estadoDia(), k = el.dataset.k; est.concil[k]?.just?.splice(Number(el.dataset.i), 1); agendarSalvar(); $("#concil-grade").innerHTML = htmlConciliacoes(est); return; }
  // qualidade
  if (a === "qual-filtro") { S.qual.filtro = el.dataset.f; return render(); }
  if (a === "qual-copiar") return feedbackCopia(el, relatorioQualidade(S.achados.filter((x) => S.qual.filtro === "todas" || x.sev === S.qual.filtro)));
  // importar
  if (a === "imp-escolher") { if (!S.imp.lendo && ev.target.id !== "arquivo-planilha") $("#arquivo-planilha")?.click(); return; }
  if (a === "imp-confirmar") return confirmarImportacao();
  if (a === "imp-cancelar") { S.imp.previa = null; S.imp.erro = ""; return render(); }
  if (a === "imp-remover") { S.imp.confirmaRemover = true; return render(); }
  if (a === "imp-remover-nao") { S.imp.confirmaRemover = false; return render(); }
  if (a === "imp-remover-sim") {
    apagarTudo(); S.dados = null; S.meta = null; S.achados = []; S.dia.estado = null; S.imp.confirmaRemover = false; cacheMes.clear();
    S.msg = { tipo: "ok", txt: "Dados removidos deste navegador." }; return render();
  }
});
document.addEventListener("keydown", (ev) => {
  if ((ev.key === "Enter" || ev.key === " ") && ev.target.matches?.('.no, .solta')) { ev.preventDefault(); ev.target.dispatchEvent(new MouseEvent("click", { bubbles: true })); }
});

function corpoCnpj() {
  const d = S.dados, q = S.contas.cnpjQ;
  const casa = (...c) => !q.trim() || semAcento(c.join(" ")).includes(semAcento(q)) || c.join(" ").replace(/\D/g, "").includes(q.replace(/\D/g, "") || "\u0000");
  return tabelasCnpj(d.cnpjs.filter((x) => casa(x.nome, x.cnpj, x.municipio)),
    d.empreendimentos.filter((x) => casa(x.nome, x.cnpj, x.gerente, x.municipio)), d.vendas.filter((x) => casa(x.numero, x.nome, x.cnpj)));
}
document.addEventListener("input", (ev) => {
  const t = ev.target;
  if (t.id === "busca-contas") { S.contas.q = t.value; $("#contas-corpo").innerHTML = corpoContas(); return; }
  if (t.id === "busca-cnpj") { S.contas.cnpjQ = t.value; $("#cnpj-corpo").innerHTML = corpoCnpj(); return; }
  const campo = t.dataset?.campo; if (!campo || !S.dados) return;
  const est = estadoDia(), k = t.dataset.k, f = t.dataset.f;
  if (["emp", "transf", "concil", "just-v"].includes(campo)) {
    const v = parseMoeda(t.value);
    t.classList.toggle("invalido", Number.isNaN(v)); t.style.borderColor = Number.isNaN(v) ? "var(--color-error)" : "";
    if (Number.isNaN(v)) return;
    if (campo === "emp") { (est.empresas[k] ||= {})[f] = v; atualizarCalculos("emp", k); }
    else if (campo === "transf") { (est.transf[k] ||= {})[f] = v; atualizarCalculos("transf", k); }
    else if (campo === "concil") { (est.concil[k] ||= {})[f] = v; atualizarCalculos("concil", k); }
    else { const j = ((est.concil[k] ||= {}).just ||= [])[Number(t.dataset.i)]; if (j) j.v = v; atualizarCalculos("just", k); }
    agendarSalvar();
  } else if (campo === "just-d") {
    const j = ((est.concil[k] ||= {}).just ||= [])[Number(t.dataset.i)]; if (j) j.d = t.value; agendarSalvar();
  }
});
document.addEventListener("focusout", (ev) => {
  const t = ev.target;
  if (!t.matches?.('input.valor[data-campo]')) return;
  const v = parseMoeda(t.value);
  if (v != null && !Number.isNaN(v)) t.value = fmt(v);
});
document.addEventListener("change", (ev) => {
  const t = ev.target;
  if (t.id === "arquivo-planilha") { lerArquivo(t.files?.[0]); t.value = ""; return; }
  if (t.id === "dia-data") { if (daIso(t.value)) { mudarDia(t.value); render(); } return; }
  if (t.id === "imp-aplicar") { S.imp.aplicarDia = t.checked; render(); return; }
  if (t.id === "imp-data") { if (daIso(t.value)) { S.imp.dataDia = t.value; render(); } return; }
  const campo = t.dataset?.campo; if (!campo || !S.dados) return;
  const est = estadoDia();
  if (campo === "check") {
    est.checks[t.dataset.k] = t.checked; agendarSalvar();
    const et = etapasDoDia(), n = et.filter((e) => est.checks[e]).length;
    const b = $("#check-barra"), x = $("#check-txt");
    if (b) b.style.width = (et.length ? Math.round((n / et.length) * 100) : 0) + "%";
    if (x) x.textContent = `${n} de ${et.length}`;
  } else if (campo === "emp-flag") { (est.empresas[t.dataset.k] ||= {})[t.dataset.f] = t.checked; agendarSalvar(); }
});
for (const tipo of ["dragover", "dragenter"]) document.addEventListener(tipo, (ev) => { const z = ev.target.closest?.(".solta"); if (z) { ev.preventDefault(); z.classList.add("ativa"); } });
document.addEventListener("dragleave", (ev) => ev.target.closest?.(".solta")?.classList.remove("ativa"));
document.addEventListener("drop", (ev) => {
  const z = ev.target.closest?.(".solta"); if (!z) return;
  ev.preventDefault(); z.classList.remove("ativa"); lerArquivo(ev.dataTransfer?.files?.[0]);
});
window.addEventListener("pagehide", salvarDia);
document.addEventListener("visibilitychange", () => { if (document.hidden) salvarDia(); });

/* =========================================================================== início */
recalcular();
{
  const h = location.hash.replace("#", "");
  if (VISOES[h]) S.aba = h;
  const hj = new Date(); S.cal.sel = null;
  if (S.aba === "calendario") S.cal.sel = hj.getDate();
}
render();
