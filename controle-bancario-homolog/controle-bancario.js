/**
 * Controle Bancário — tela do módulo (NOCTUS, homologação).
 *
 * A planilha "Financeiro" tem agência, conta, CNPJ e valores, e o repositório do
 * NOCTUS é público: por isso ela não vai ao repositório nem ao servidor. A pessoa
 * escolhe o .xlsx, o módulo lê aqui mesmo (SheetJS) e guarda o resultado só no
 * armazenamento deste navegador, separado por usuário do NOCTUS.
 *
 * Layout pensado para ser lido rápido: cinco abas, uma pergunta por aba
 * ("o que pago hoje?", "o que vence no mês?", "qual a conta?", "o dia fechou?",
 * "a planilha está certa?"). Detalhe só aparece quando a pessoa pede.
 */
import { semAcento, validarCnpj, isoData, daIso, motivoNaoUtil, montarMes, lerPlanilha, auditar } from "./planilha.js";

/* ============================================================== utilidades */
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const $ = (s) => document.querySelector(s);
const round2 = (n) => Math.round(n * 100) / 100;
const fmt = (n) => (n == null || Number.isNaN(n) ? "" : n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const pad = (n) => String(n).padStart(2, "0");
const MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
const SEMANA = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
const SEM_CURTA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const maiuscula = (s) => s.charAt(0).toUpperCase() + s.slice(1);
const dataLonga = (d) => maiuscula(`${SEMANA[d.getDay()]}, ${d.getDate()} de ${MESES[d.getMonth()]}`);
const hojeIso = () => isoData(new Date());
const somarDias = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);

/** "1.234,56", "1234.56", "R$ 1.234,56", "-5" → número; vazio → null; ilegível → NaN. */
function parseMoeda(s) {
  let t = String(s ?? "").replace(/[R$\s]/g, "");
  if (!t) return null;
  const neg = /^\(.*\)$/.test(t) || t.startsWith("-");
  t = t.replace(/[()-]/g, "");
  if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
  else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, "");
  else if (/^\d{1,3}\.\d{3}$/.test(t)) t = t.replace(".", "");
  if (!/^\d+(\.\d+)?$/.test(t)) return NaN;
  return round2(neg ? -Number(t) : Number(t));
}

/** As 9 categorias do leitor viram 5 cores — mais que isso o olho não separa. */
const COR = { imposto: "imposto", folha: "folha", repasse: "repasse", comissao: "contrato", contrato: "contrato",
  energia: "outros", cambio: "outros", ordinario: "outros", outros: "outros" };
const LEGENDA = [["imposto", "Impostos"], ["folha", "Folha e subsídio"], ["repasse", "Repasses"], ["contrato", "Contratos e comissões"], ["outros", "Outros"]];
const bolinha = (cat) => `<span class="bolinha c-${COR[cat] || "outros"}"></span>`;

/* ================================================================ armazenamento */
function escopoUsuario() {
  try {
    const s = JSON.parse(localStorage.getItem("sipep.session") || "null");
    const base = String(s?.user?.username || s?.user?.id || "sem-login").trim().toLowerCase();
    return base.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40) || "sem-login";
  } catch { return "sem-login"; }
}
const PREFIXO = `noctus.controlebancario.${escopoUsuario()}.`;
let armazenamentoOk = true;
const ler = (k) => { try { const v = localStorage.getItem(PREFIXO + k); return v ? JSON.parse(v) : null; } catch { return null; } };
const gravar = (k, v) => { try { localStorage.setItem(PREFIXO + k, JSON.stringify(v)); return true; } catch { armazenamentoOk = false; return false; } };
function apagarTudo() {
  try {
    const ks = [];
    for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i) || ""; if (k.startsWith(PREFIXO)) ks.push(k); }
    ks.forEach((k) => localStorage.removeItem(k));
  } catch { /* sem armazenamento */ }
}

/* ===================================================================== estado */
const agora = new Date();
const S = {
  aba: "hoje",
  dados: ler("dados"), meta: ler("meta"), achados: [],
  cal: { ano: agora.getFullYear(), mes: agora.getMonth(), sel: agora.getDate() },
  contas: { modo: "empresa", emp: null, q: "" },
  dia: { iso: hojeIso(), estado: null },
  imp: { previa: null, lendo: false, erro: "", confirmaRemover: false },
  msg: null,
};
const cacheMes = new Map();
function recalcular() {
  cacheMes.clear();
  S.achados = S.dados ? auditar(S.dados) : [];
  if (S.dados && !S.dados.empresas.some((e) => e.id === S.contas.emp)) S.contas.emp = S.dados.empresas[0]?.id || null;
}
const itensDoMes = (ano, mes) => {
  const k = `${ano}-${mes}`;
  if (!cacheMes.has(k)) cacheMes.set(k, S.dados ? montarMes(S.dados, ano, mes) : {});
  return cacheMes.get(k);
};
const itensDaData = (d) => itensDoMes(d.getFullYear(), d.getMonth())[d.getDate()] || [];
const nomeEmp = (id) => S.dados?.empresas.find((e) => e.id === id)?.nome || id;

/* ============================================================ peças comuns */
const listaEventos = (l) => `<ul class="ev-lista">${l.map((i) =>
  `<li>${bolinha(i.categoria)}<span>${esc(i.texto)}${i.nota ? `<small>${esc(i.nota)}</small>` : ""}</span></li>`).join("")}</ul>`;

function semPlanilha() {
  return `<div class="cartao" style="text-align:center;padding:3rem 1.5rem">
    <h2 style="font-size:1.2rem;margin-bottom:.5rem">Importe a planilha Financeiro para começar</h2>
    <p class="sub" style="max-width:460px;margin:0 auto 1.4rem">O calendário, as contas e o fechamento do dia são montados a partir dela.</p>
    <button class="botao primario" data-acao="ir" data-aba="planilha">Importar planilha</button></div>`;
}

/* ========================================================================= HOJE */
function vHoje() {
  if (!S.dados) return semPlanilha();
  const hoje = new Date(), itens = itensDaData(hoje), fechado = motivoNaoUtil(hoje);
  const est = ler("dia." + hojeIso()), etapas = etapasDoDia();
  const feitas = est ? etapas.filter((e) => est.checks?.[e]).length : 0;
  const diverg = est ? contarDivergencias(est) : 0;

  const prox = [];
  for (let i = 1; i <= 10; i++) { const dt = somarDias(hoje, i), l = itensDaData(dt); if (l.length) prox.push({ dt, l }); }

  const altas = S.achados.filter((a) => a.sev === "alta").length;
  const aviso = altas ? `<button class="aviso-planilha" data-acao="ir" data-aba="planilha">
      <span>A planilha tem ${altas} ${altas === 1 ? "ponto" : "pontos"} para corrigir</span><span>Ver →</span></button>` : "";

  return `<div class="colunas">
    <section class="cartao">
      <p class="rotulo">Hoje</p><h2 style="font-size:1.4rem;margin-bottom:1.1rem">${esc(dataLonga(hoje))}</h2>
      ${fechado ? `<p class="nota">${esc(fechado)} — bancos fechados.</p>` : ""}
      ${itens.length ? listaEventos(itens) : `<p class="vazio">Nenhum pagamento programado para hoje.</p>`}
      <div class="hoje-status">
        <span>Fechamento do dia: <b>${feitas} de ${etapas.length}</b> etapas${diverg ? ` · <span class="neg">${diverg} com diferença</span>` : ""}</span>
        <button class="botao" data-acao="ir-dia" data-iso="${hojeIso()}">Abrir fechamento</button>
      </div>
    </section>
    <section class="cartao">
      <h2 style="margin-bottom:.6rem">Próximos dias</h2>
      ${prox.length ? `<ul class="agenda">${prox.map(({ dt, l }) => `<li>
        <span class="data"><b>${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}</b><small>${SEM_CURTA[dt.getDay()]}</small></span>
        <span class="itens">${l.map((i) => `<span>${bolinha(i.categoria)}${esc(i.texto)}</span>`).join("")}</span></li>`).join("")}</ul>`
        : `<p class="vazio">Nada programado nos próximos 10 dias.</p>`}
    </section>
  </div>${aviso ? `<div style="margin-top:1.1rem">${aviso}</div>` : ""}`;
}

/* =================================================================== CALENDÁRIO */
function vCalendario() {
  if (!S.dados) return semPlanilha();
  const { ano, mes, sel } = S.cal, itens = itensDoMes(ano, mes);
  const ultimo = new Date(ano, mes + 1, 0).getDate();
  const desloc = (new Date(ano, mes, 1).getDay() + 6) % 7;
  const hoje = hojeIso();
  let cel = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"].map((x) => `<div class="cal-cab">${x}</div>`).join("");
  for (let i = 0; i < desloc; i++) cel += `<div class="dia fora"></div>`;
  for (let n = 1; n <= ultimo; n++) {
    const dt = new Date(ano, mes, n), l = itens[n] || [];
    const cls = ["dia", motivoNaoUtil(dt) ? "fechado" : "", isoData(dt) === hoje ? "hoje" : "", sel === n ? "sel" : "",
      !l.length && isoData(dt) !== hoje && sel !== n ? "vazio" : ""].join(" ");
    cel += `<button class="${cls}" data-acao="cal-sel" data-dia="${n}" aria-label="${esc(dataLonga(dt))}">
      <span class="num">${n}</span>
      ${l.slice(0, 2).map((i) => `<span class="ev">${bolinha(i.categoria)}<span>${esc(i.texto)}</span></span>`).join("")}
      ${l.length > 2 ? `<span class="mais">+${l.length - 2}</span>` : ""}</button>`;
  }
  let detalhe = `<p class="vazio">Clique num dia para ver os pagamentos.</p>`;
  if (sel) {
    const dt = new Date(ano, mes, sel), l = itens[sel] || [], fechado = motivoNaoUtil(dt);
    detalhe = `<p class="rotulo">${esc(SEMANA[dt.getDay()])}</p><h2>${sel} de ${MESES[mes]}</h2>
      ${fechado ? `<p class="nota">${esc(fechado)} — bancos fechados.</p>` : ""}
      ${l.length ? listaEventos(l) : `<p class="vazio">Nenhum pagamento programado.</p>`}`;
  }
  return `<div class="cal-layout">
    <section class="cartao">
      <div class="cal-nav">
        <button class="icone" data-acao="cal-mes" data-d="-1" aria-label="Mês anterior">‹</button>
        <h2>${maiuscula(MESES[mes])} de ${ano}</h2>
        <button class="icone" data-acao="cal-mes" data-d="1" aria-label="Próximo mês">›</button>
        <button class="botao" data-acao="cal-hoje">Hoje</button>
      </div>
      <div class="cal-grade">${cel}</div>
      <div class="legenda">${LEGENDA.map(([c, r]) => `<span>${bolinha(c)}${r}</span>`).join("")}</div>
    </section>
    <aside class="cartao detalhe">${detalhe}</aside>
  </div>`;
}

/* ======================================================================= CONTAS */
function vContas() {
  if (!S.dados) return semPlanilha();
  const m = S.contas.modo;
  return `<div class="alternar" role="group">
      <button data-acao="modo" data-modo="empresa" aria-pressed="${m === "empresa"}">Contas por empresa</button>
      <button data-acao="modo" data-modo="cnpj" aria-pressed="${m === "cnpj"}">CNPJs</button></div>
    ${m === "cnpj" ? vCnpjs() : vEmpresa()}`;
}
function vEmpresa() {
  const d = S.dados, emp = d.empresas.find((e) => e.id === S.contas.emp) || d.empresas[0];
  if (!emp) return `<p class="vazio">Nenhuma empresa na planilha.</p>`;
  const contas = d.contas.filter((c) => c.empresa === emp.id);
  const cnpjs = [...new Set(contas.map((c) => c.cnpj).filter(Boolean))];
  const invalido = (c) => (c && /\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}/.test(c) && !validarCnpj(c) ? `<span class="selo-erro">CNPJ inválido</span>` : "");
  const cp = d.contrapartes[emp.id] || { inter: [], bancos: [] };
  const hist = d.historicos[emp.id] || [];
  const tabela = (l) => `<div class="tabela"><table><tbody>${l.map((x) =>
    `<tr><td>${esc(x.nome)}</td><td class="num">${esc(x.codigo)}</td></tr>`).join("")}</tbody></table></div>`;
  const temCodigos = cp.inter.length || cp.bancos.length || hist.length;
  return `<div class="contas-layout">
    <nav class="empresas">${d.empresas.map((e) => `<button data-acao="emp" data-id="${esc(e.id)}" aria-pressed="${e.id === emp.id}">
      ${esc(e.nome)}<small>${d.contas.filter((c) => c.empresa === e.id).length}</small></button>`).join("")}</nav>
    <div>
      <section class="cartao">
        <div class="cab"><div><h2 style="font-size:1.25rem">${esc(emp.nome)}</h2>
          ${cnpjs.length === 1 ? `<p class="sub">CNPJ ${esc(cnpjs[0])}${invalido(cnpjs[0])}</p>` : ""}</div>
          <span class="sub">${contas.length} ${contas.length === 1 ? "conta" : "contas"}</span></div>
        ${contas.length ? `<div class="contas">${contas.map((c) => `<div class="conta">
          <div class="banco">${esc(c.banco)}${c.obs || cnpjs.length > 1 ? `<small>${esc([c.obs, cnpjs.length > 1 ? "CNPJ " + c.cnpj : ""].filter(Boolean).join(" · "))}${invalido(c.cnpj)}</small>` : ""}</div>
          <div class="nums"><span><small>Agência</small>${esc(c.agencia)}</span><span><small>Conta</small>${esc(c.conta)}</span></div>
        </div>`).join("")}</div>` : `<p class="vazio">Nenhuma conta nesta aba.</p>`}
      </section>
      ${temCodigos ? `<details class="cartao dobra"><summary>Códigos e históricos para lançamento</summary><div class="miolo">
        ${cp.inter.length ? `<h3>Transferência para outra empresa</h3>${tabela(cp.inter)}` : ""}
        ${cp.bancos.length ? `<h3>Contas de banco</h3>${tabela(cp.bancos)}` : ""}
        ${hist.length ? `<h3>Textos de histórico</h3><div class="tabela"><table><tbody>${hist.map((h) => `<tr><td>${esc(h.texto)}</td></tr>`).join("")}</tbody></table></div>` : ""}
      </div></details>` : ""}
    </div></div>`;
}
function vCnpjs() {
  return `<section class="cartao"><input type="search" id="busca-cnpj" placeholder="Buscar por nome, CNPJ, gerente ou município" value="${esc(S.contas.q)}" autocomplete="off">
    <div id="cnpj-corpo">${corpoCnpj()}</div></section>`;
}
function corpoCnpj() {
  const d = S.dados, q = S.contas.q.trim(), nq = semAcento(q), dq = q.replace(/\D/g, "");
  const casa = (...c) => !q || semAcento(c.join(" ")).includes(nq) || (dq.length >= 3 && c.join(" ").replace(/\D/g, "").includes(dq));
  const cnpj = (c) => `${esc(c)}${c && /\d{2}\.?\d{3}/.test(c) && !validarCnpj(c) ? `<span class="selo-erro">inválido</span>` : ""}`;
  const cn = d.cnpjs.filter((x) => casa(x.nome, x.cnpj, x.municipio));
  const em = d.empreendimentos.filter((x) => casa(x.nome, x.cnpj, x.gerente, x.municipio));
  if (!cn.length && !em.length) return `<p class="vazio">Nada encontrado.</p>`;
  return `${cn.length ? `<div class="tabela"><table><thead><tr><th>Empresa</th><th>CNPJ</th></tr></thead><tbody>
      ${cn.map((x) => `<tr><td>${esc(x.nome)}</td><td style="white-space:nowrap">${cnpj(x.cnpj)}</td></tr>`).join("")}</tbody></table></div>` : ""}
    ${em.length ? `<h3 style="font-size:.75rem;text-transform:uppercase;letter-spacing:.06em;color:var(--mudo);margin:1.6rem 0 .3rem">Empreendimentos</h3>
      <div class="tabela"><table><thead><tr><th>Nome</th><th>CNPJ</th><th>Gerente</th><th>Município</th></tr></thead><tbody>
      ${em.map((x) => `<tr><td>${esc(x.nome)}</td><td style="white-space:nowrap">${cnpj(x.cnpj)}</td><td>${esc(x.gerente)}</td><td>${esc(x.municipio)}</td></tr>`).join("")}</tbody></table></div>` : ""}`;
}

/* ============================================================ FECHAMENTO DO DIA */
const diaVazio = () => ({ v: 1, empresas: {}, transf: {}, concil: {}, checks: {}, extras: [] });
const chaveEmp = (e) => e.id || semAcento(e.nome).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
const estadoDia = () => (S.dia.estado ||= ler("dia." + S.dia.iso) || diaVazio());
const soma = (l, f) => round2(l.reduce((s, x) => s + (f(x) || 0), 0));
const saidaDe = (r = {}) => (r.total == null && r.numerarios == null && r.compromissos == null
  ? null : round2((r.total || 0) - (r.numerarios || 0) - (r.compromissos || 0)));
const difConcil = (c = {}) => (c.fluxo != null && c.relatorio != null ? round2(c.fluxo - c.relatorio) : null);
const resConcil = (c = {}) => { const d = difConcil(c); return d == null ? null : round2(d - soma(c.just || [], (j) => j.v)); };
const contarDivergencias = (est) => Object.values(est.concil || {}).filter((c) => { const r = resConcil(c); return r != null && r !== 0; }).length;
const etapasDoDia = () => (S.dados?.dia?.checklist?.length ? S.dados.dia.checklist : ["Banco", "Relatório", "Extrato", "Suprimento", "Comprovantes"]);
const linhasEmpresas = () => (S.dados.dia?.empresas?.length ? S.dados.dia.empresas.map((e) => ({ k: chaveEmp(e), nome: e.nome }))
  : S.dados.empresas.map((e) => ({ k: e.id, nome: e.nome })));
const linhasTransf = () => (S.dados.dia?.transferencias || []).map((e) => ({ k: chaveEmp(e), nome: e.nome }));
function chavesConcil(est) {
  const base = (S.dados.dia?.conciliacao || []).map((e) => ({ k: chaveEmp(e), nome: e.nome }));
  for (const k of est.extras || []) if (!base.some((b) => b.k === k)) base.push({ k, nome: nomeEmp(k), extra: true });
  return base;
}
/** Nome como aparece nas abas de empresa ("IRM", "Posto Plaza"), não o "POSTO SB" em caixa alta da aba do dia. */
const nomeCurto = (n, k) => S.dados?.empresas.find((e) => e.id === k)?.nome || maiuscula(String(n).toLowerCase());
const campo = (tipo, k, f, v, rotulo) =>
  `<input type="text" class="valor" inputmode="decimal" autocomplete="off" data-campo="${tipo}" data-k="${esc(k)}" data-f="${f}" value="${esc(fmt(v))}" aria-label="${esc(rotulo)}">`;
const txtValor = (n) => (n == null ? "—" : fmt(n));
const faltaDe = (r = {}) => (r.aTransferir == null && r.transferido == null ? null : round2((r.aTransferir || 0) - (r.transferido || 0)));

function vFechamento() {
  if (!S.dados) return semPlanilha();
  const dt = daIso(S.dia.iso), est = estadoDia(), etapas = etapasDoDia(), fechado = motivoNaoUtil(dt);
  const feitas = etapas.filter((e) => est.checks[e]).length;
  const emps = linhasEmpresas(), regs = emps.map((e) => est.empresas[e.k] || {}), transf = linhasTransf();
  const livres = S.dados.empresas.filter((e) => !chavesConcil(est).some((c) => c.k === e.id));

  return `<div class="dia-nav">
      <button class="icone" data-acao="dia-mover" data-d="-1" aria-label="Dia anterior">‹</button>
      <button class="icone" data-acao="dia-mover" data-d="1" aria-label="Próximo dia">›</button>
      <h2>${esc(dataLonga(dt))}</h2>
      ${S.dia.iso !== hojeIso() ? `<button class="botao" data-acao="dia-hoje">Ir para hoje</button>` : ""}
      <span class="espaco"></span>
      <input type="date" id="dia-data" value="${esc(S.dia.iso)}" aria-label="Escolher data">
      <span class="salvo" id="salvo">Salvo automaticamente</span>
    </div>
    ${fechado ? `<p class="nota">${esc(fechado)} — bancos fechados.</p>` : ""}

    <section class="cartao">
      <div class="cab" style="margin-bottom:.8rem"><h2>Etapas</h2><span class="sub" id="etapas-txt">${feitas} de ${etapas.length} concluídas</span></div>
      <div class="etapas">${etapas.map((e) => `<label class="etapa"><input type="checkbox" data-campo="check" data-k="${esc(e)}" ${est.checks[e] ? "checked" : ""}>${esc(e)}</label>`).join("")}</div>
    </section>

    <section class="cartao">
      <div class="cab"><h2>Valores do relatório</h2><span class="sub">Saídas = Total − Numerários − Compromissos</span></div>
      <div class="tabela"><table>
        <thead><tr><th>Empresa</th><th class="num">Total</th><th class="num">Numerários</th><th class="num">Compromissos</th><th class="num">Saídas</th></tr></thead>
        <tbody>${emps.map((e, i) => { const r = regs[i], s = saidaDe(r); return `<tr><td>${esc(nomeCurto(e.nome, e.k))}</td>
          <td>${campo("emp", e.k, "total", r.total, "Total " + e.nome)}</td>
          <td>${campo("emp", e.k, "numerarios", r.numerarios, "Numerários " + e.nome)}</td>
          <td>${campo("emp", e.k, "compromissos", r.compromissos, "Compromissos " + e.nome)}</td>
          <td class="calc ${s < 0 ? "neg" : ""}" data-out="saida" data-k="${esc(e.k)}">${txtValor(s)}</td></tr>`; }).join("")}</tbody>
        <tfoot><tr><td>Total</td><td class="num" data-out="t-total">${fmt(soma(regs, (r) => r.total))}</td>
          <td class="num" data-out="t-numerarios">${fmt(soma(regs, (r) => r.numerarios))}</td>
          <td class="num" data-out="t-compromissos">${fmt(soma(regs, (r) => r.compromissos))}</td>
          <td class="num" data-out="t-saida">${fmt(soma(regs, (r) => saidaDe(r)))}</td></tr></tfoot>
      </table></div>
    </section>

    <section class="cartao">
      <div class="cab"><h2>Conciliação banco × relatório</h2><span class="sub">A diferença precisa ser zero ou estar justificada</span></div>
      <div class="concil-grade" id="concil-grade">${htmlConciliacoes(est)}</div>
      ${livres.length ? `<div class="concil-add"><select id="concil-add" aria-label="Adicionar empresa à conciliação">
        <option value="">+ Adicionar empresa</option>${livres.map((e) => `<option value="${esc(e.id)}">${esc(e.nome)}</option>`).join("")}</select></div>` : ""}
    </section>

    ${transf.length ? `<details class="cartao dobra"><summary>Transferências entre empresas</summary><div class="miolo">
      <div class="tabela"><table><thead><tr><th>Empresa</th><th class="num">A transferir</th><th class="num">Transferido</th><th class="num">Falta</th></tr></thead><tbody>
      ${transf.map((e) => { const r = est.transf[e.k] || {}, f = faltaDe(r); return `<tr><td>${esc(nomeCurto(e.nome, e.k))}</td>
        <td>${campo("transf", e.k, "aTransferir", r.aTransferir, "A transferir " + e.nome)}</td>
        <td>${campo("transf", e.k, "transferido", r.transferido, "Transferido " + e.nome)}</td>
        <td class="calc ${f ? "neg" : f === 0 ? "pos" : ""}" data-out="falta" data-k="${esc(e.k)}">${txtValor(f)}</td></tr>`; }).join("")}
      </tbody></table></div></div></details>` : ""}`;
}
function situacao(res) {
  if (res == null) return { cls: "", txt: "a preencher" };
  if (res === 0) return { cls: "ok", txt: "confere" };
  return { cls: "falta", txt: `faltam ${fmt(Math.abs(res))}` };
}
function htmlConciliacoes(est) {
  return chavesConcil(est).map(({ k, nome, extra }) => {
    const c = est.concil[k] || {}, dif = difConcil(c), res = resConcil(c), s = situacao(res);
    const just = c.just || [];
    return `<div class="concil ${res == null ? "" : res === 0 ? "confere" : "diverge"}" data-card="${esc(k)}">
      <div class="concil-cab"><b>${esc(nomeCurto(nome, k))}</b><span class="situacao ${s.cls}" data-out="situacao">${s.txt}</span></div>
      <label class="campo">Banco ${campo("concil", k, "fluxo", c.fluxo, "Banco " + nome)}</label>
      <label class="campo">Relatório ${campo("concil", k, "relatorio", c.relatorio, "Relatório " + nome)}</label>
      <div class="dif"><span>Diferença</span><span data-out="dif">${txtValor(dif)}</span></div>
      ${just.map((j, i) => `<div class="just"><input type="text" data-campo="just-d" data-k="${esc(k)}" data-i="${i}" value="${esc(j.d)}" placeholder="Motivo" aria-label="Motivo">
        <input type="text" class="valor" inputmode="decimal" data-campo="just-v" data-k="${esc(k)}" data-i="${i}" value="${esc(fmt(j.v))}" placeholder="Valor" aria-label="Valor">
        <button class="x" data-acao="just-rem" data-k="${esc(k)}" data-i="${i}" aria-label="Remover">✕</button></div>`).join("")}
      <button class="link" data-acao="just-add" data-k="${esc(k)}">+ Justificar diferença</button>
      ${extra ? ` <button class="link discreto" data-acao="concil-rem" data-k="${esc(k)}">Remover</button>` : ""}
    </div>`;
  }).join("");
}
function atualizarCalculos(tipo, k) {
  const est = estadoDia(), K = CSS.escape(k);
  const set = (sel, txt, cls) => { const el = document.querySelector(sel); if (!el) return; el.textContent = txt; if (cls != null) el.className = cls; };
  if (tipo === "emp") {
    const s = saidaDe(est.empresas[k]);
    set(`[data-out="saida"][data-k="${K}"]`, txtValor(s), `calc ${s < 0 ? "neg" : ""}`);
    const regs = linhasEmpresas().map((e) => est.empresas[e.k] || {});
    for (const f of ["total", "numerarios", "compromissos"]) set(`[data-out="t-${f}"]`, fmt(soma(regs, (r) => r[f])));
    set('[data-out="t-saida"]', fmt(soma(regs, (r) => saidaDe(r))));
  } else if (tipo === "transf") {
    const f = faltaDe(est.transf[k]);
    set(`[data-out="falta"][data-k="${K}"]`, txtValor(f), `calc ${f ? "neg" : f === 0 ? "pos" : ""}`);
  } else {
    const c = est.concil[k] || {}, res = resConcil(c), s = situacao(res), card = document.querySelector(`[data-card="${K}"]`);
    if (!card) return;
    card.className = `concil ${res == null ? "" : res === 0 ? "confere" : "diverge"}`;
    set(`[data-card="${K}"] [data-out="dif"]`, txtValor(difConcil(c)));
    set(`[data-card="${K}"] [data-out="situacao"]`, s.txt, `situacao ${s.cls}`);
  }
}
let timerSalvar = null;
const agendarSalvar = () => { clearTimeout(timerSalvar); timerSalvar = setTimeout(salvarDia, 350); };
function salvarDia() {
  clearTimeout(timerSalvar); timerSalvar = null;
  if (!S.dia.estado) return;
  const ok = gravar("dia." + S.dia.iso, S.dia.estado), el = $("#salvo");
  if (el) el.textContent = ok ? `Salvo às ${new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}` : "Não foi possível salvar neste navegador";
}
function mudarDia(iso) { salvarDia(); S.dia.iso = iso; S.dia.estado = null; }

/* ===================================================================== PLANILHA */
const GRUPOS = [["alta", "Corrigir", "pode levar a pagamento ou lançamento errado"],
  ["media", "Verificar", "informação que não bate entre abas"], ["baixa", "Detalhes", "ajustes de organização"]];
function vPlanilha() {
  const p = S.imp.previa;
  let topo;
  if (p) {
    const c = contagens(p.dados);
    topo = `<section class="cartao"><p class="rotulo">Conferir antes de importar</p><h2>${esc(p.nome)}</h2>
      <div class="numeros"><div><b>${c.empresas}</b><span>empresas</span></div><div><b>${c.contas}</b><span>contas bancárias</span></div>
        <div><b>${c.vencimentos}</b><span>vencimentos no mês</span></div><div><b>${c.cnpjs}</b><span>CNPJs</span></div></div>
      ${p.aviso.map((a) => `<p class="nota">${esc(a)}</p>`).join("")}
      <div class="acoes"><button class="botao primario" data-acao="imp-confirmar">Importar</button><button class="botao" data-acao="imp-cancelar">Cancelar</button></div></section>`;
  } else if (S.dados) {
    const c = contagens(S.dados);
    topo = `<section class="cartao"><div class="cab" style="margin-bottom:0"><div><p class="rotulo">Planilha em uso</p>
        <h2>${esc(S.meta?.arquivo || "Financeiro")}</h2>
        <p class="sub">${S.meta ? `Importada em ${esc(new Date(S.meta.importadoEm).toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" }))} · ` : ""}${c.empresas} empresas · ${c.contas} contas</p></div>
      <div class="acoes"><button class="botao" data-acao="imp-escolher">${S.imp.lendo ? "Lendo…" : "Trocar planilha"}</button></div></div>
      ${S.imp.erro ? `<p class="nota neg" style="margin:1rem 0 0">${esc(S.imp.erro)}</p>` : ""}</section>`;
  } else {
    topo = `<section class="cartao"><div class="solta" data-acao="imp-escolher" tabindex="0" role="button">
      <strong>${S.imp.lendo ? "Lendo a planilha…" : "Escolha a planilha Financeiro (.xlsx)"}</strong><span>ou arraste o arquivo para cá</span></div>
      ${S.imp.erro ? `<p class="nota neg" style="margin:1rem 0 0">${esc(S.imp.erro)}</p>` : ""}</section>`;
  }

  let revisao = "";
  if (S.dados && !p) {
    revisao = `<section class="cartao"><h2>Pontos para revisar na planilha</h2>
      ${S.achados.length ? GRUPOS.map(([sev, titulo, desc]) => {
        const l = S.achados.filter((a) => a.sev === sev);
        return l.length ? `<h3 class="grupo ${sev}">${titulo} <span class="n">${l.length}</span><small>— ${desc}</small></h3>
          ${l.map((a) => `<details class="achado"><summary>${esc(a.titulo)}</summary><div class="corpo">
            <p>${esc(a.detalhe)}</p><p><b>O que fazer:</b> ${esc(a.sugestao)}</p><p>Onde: ${esc(a.onde)}</p></div></details>`).join("")}` : "";
      }).join("") : `<p class="vazio" style="margin-top:.8rem">Nenhum problema encontrado.</p>`}</section>`;
  }

  const remover = S.dados && !p ? `<p class="privado">${S.imp.confirmaRemover
    ? `Apagar a planilha e os fechamentos deste navegador? <button class="link neg" data-acao="imp-remover-sim">Apagar</button> · <button class="link discreto" data-acao="imp-remover-nao">Cancelar</button>`
    : `Os dados ficam só neste navegador — não vão para o servidor nem para o GitHub. <button class="link discreto" data-acao="imp-remover">Remover dados</button>`}</p>`
    : `<p class="privado">Os dados ficam só neste navegador — não vão para o servidor nem para o GitHub.</p>`;
  return topo + revisao + remover;
}
function contagens(d) {
  const h = new Date(), mes = montarMes(d, h.getFullYear(), h.getMonth());
  return { empresas: d.empresas.length, contas: d.contas.length, cnpjs: d.cnpjs.length,
    vencimentos: Object.values(mes).reduce((s, l) => s + l.length, 0) };
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
    throw new Error("Não consegui carregar o leitor de planilhas. Verifique a conexão e tente de novo.");
  })();
  return promessaXlsx;
}
async function lerArquivo(file) {
  if (!file) return;
  S.aba = "planilha";
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { S.imp.erro = "Escolha um arquivo do Excel (.xlsx)."; render(); return; }
  S.imp.lendo = true; S.imp.erro = ""; S.imp.previa = null; render();
  try {
    await carregarXLSX();
    const wb = window.XLSX.read(await file.arrayBuffer(), { type: "array" });
    const { dados, aviso } = lerPlanilha(window.XLSX, wb);
    S.imp.previa = { nome: file.name, tamanho: file.size, dados, aviso };
  } catch (e) {
    S.imp.erro = e?.message || "Não consegui ler este arquivo.";
  } finally { S.imp.lendo = false; render(); }
}
function confirmarImportacao() {
  const p = S.imp.previa; if (!p) return;
  S.dados = p.dados;
  S.meta = { arquivo: p.nome, tamanho: p.tamanho, importadoEm: new Date().toISOString() };
  const ok = gravar("dados", S.dados) && gravar("meta", S.meta);
  S.imp.previa = null;
  recalcular();
  S.msg = ok ? { txt: `Planilha importada: ${S.dados.empresas.length} empresas e ${S.dados.contas.length} contas.` }
    : { tipo: "erro", txt: "A planilha foi lida, mas este navegador não deixou gravar. Ela vale só até fechar a página." };
  ir("hoje");
}

/* ===================================================================== NAVEGAÇÃO */
const VISOES = { hoje: vHoje, calendario: vCalendario, contas: vContas, fechamento: vFechamento, planilha: vPlanilha };
function render() {
  document.querySelectorAll("#abas .aba").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.aba === S.aba)));
  const altas = S.achados.filter((a) => a.sev === "alta").length, n = $("#n-planilha");
  n.textContent = altas; n.classList.toggle("oculto", !altas);
  const msg = S.msg ? `<div class="msg ${S.msg.tipo || ""}">${esc(S.msg.txt)}</div>` : ""; S.msg = null;
  const aviso = armazenamentoOk ? "" : `<div class="msg erro">Este navegador não está deixando gravar. Ao fechar a página, será preciso importar de novo.</div>`;
  $("#app").innerHTML = aviso + msg + (VISOES[S.aba] || vHoje)();
}
function ir(aba) {
  if (S.aba === "fechamento" && aba !== "fechamento") salvarDia();
  S.aba = VISOES[aba] ? aba : "hoje";
  try { history.replaceState(null, "", "#" + S.aba); } catch { /* sem history */ }
  render(); window.scrollTo(0, 0);
}

document.addEventListener("click", (ev) => {
  const aba = ev.target.closest(".aba[data-aba]");
  if (aba) return ir(aba.dataset.aba);
  const el = ev.target.closest("[data-acao]"); if (!el) return;
  const a = el.dataset.acao, est = () => estadoDia();
  switch (a) {
    case "ir": return ir(el.dataset.aba);
    case "ir-dia": mudarDia(el.dataset.iso); return ir("fechamento");
    case "cal-sel": S.cal.sel = Number(el.dataset.dia); return render();
    case "cal-mes": {
      let m = S.cal.mes + Number(el.dataset.d), y = S.cal.ano;
      if (m < 0) { m = 11; y--; } else if (m > 11) { m = 0; y++; }
      Object.assign(S.cal, { mes: m, ano: y, sel: null }); return render();
    }
    case "cal-hoje": { const h = new Date(); Object.assign(S.cal, { mes: h.getMonth(), ano: h.getFullYear(), sel: h.getDate() }); return render(); }
    case "modo": S.contas.modo = el.dataset.modo; return render();
    case "emp": S.contas.emp = el.dataset.id; return render();
    case "dia-mover": mudarDia(isoData(somarDias(daIso(S.dia.iso), Number(el.dataset.d)))); return render();
    case "dia-hoje": mudarDia(hojeIso()); return render();
    case "just-add": {
      const c = (est().concil[el.dataset.k] ||= {}); (c.just ||= []).push({ d: "", v: null });
      agendarSalvar(); $("#concil-grade").innerHTML = htmlConciliacoes(est());
      const ins = document.querySelectorAll(`[data-card="${CSS.escape(el.dataset.k)}"] .just input[data-campo="just-d"]`);
      ins[ins.length - 1]?.focus(); return;
    }
    case "just-rem":
      est().concil[el.dataset.k]?.just?.splice(Number(el.dataset.i), 1); agendarSalvar();
      $("#concil-grade").innerHTML = htmlConciliacoes(est()); return;
    case "concil-rem": { const e = est(); e.extras = (e.extras || []).filter((x) => x !== el.dataset.k); delete e.concil[el.dataset.k]; salvarDia(); return render(); }
    case "imp-escolher": if (!S.imp.lendo) $("#arquivo-planilha").click(); return;
    case "imp-confirmar": return confirmarImportacao();
    case "imp-cancelar": S.imp.previa = null; S.imp.erro = ""; return render();
    case "imp-remover": S.imp.confirmaRemover = true; return render();
    case "imp-remover-nao": S.imp.confirmaRemover = false; return render();
    case "imp-remover-sim":
      apagarTudo(); Object.assign(S, { dados: null, meta: null, achados: [] }); S.dia.estado = null; S.imp.confirmaRemover = false; cacheMes.clear();
      S.msg = { txt: "Dados removidos deste navegador." }; return render();
  }
});
document.addEventListener("keydown", (ev) => {
  if ((ev.key === "Enter" || ev.key === " ") && ev.target.matches?.(".solta")) { ev.preventDefault(); $("#arquivo-planilha").click(); }
});
document.addEventListener("input", (ev) => {
  const t = ev.target;
  if (t.id === "busca-cnpj") { S.contas.q = t.value; $("#cnpj-corpo").innerHTML = corpoCnpj(); return; }
  const tipo = t.dataset?.campo; if (!tipo || !S.dados) return;
  const est = estadoDia(), k = t.dataset.k, f = t.dataset.f;
  if (tipo === "just-d") { const j = est.concil[k]?.just?.[Number(t.dataset.i)]; if (j) j.d = t.value; return agendarSalvar(); }
  if (!["emp", "transf", "concil", "just-v"].includes(tipo)) return;
  const v = parseMoeda(t.value);
  t.classList.toggle("invalido", Number.isNaN(v));
  if (Number.isNaN(v)) return;
  if (tipo === "emp") (est.empresas[k] ||= {})[f] = v;
  else if (tipo === "transf") (est.transf[k] ||= {})[f] = v;
  else if (tipo === "concil") (est.concil[k] ||= {})[f] = v;
  else { const j = est.concil[k]?.just?.[Number(t.dataset.i)]; if (j) j.v = v; }
  atualizarCalculos(tipo === "just-v" ? "concil" : tipo, k);
  agendarSalvar();
});
document.addEventListener("focusout", (ev) => {
  const t = ev.target;
  if (!t.matches?.("input.valor")) return;
  const v = parseMoeda(t.value);
  if (v != null && !Number.isNaN(v)) t.value = fmt(v);
});
document.addEventListener("change", (ev) => {
  const t = ev.target;
  if (t.id === "arquivo-planilha") { lerArquivo(t.files?.[0]); t.value = ""; return; }
  if (t.id === "dia-data") { if (daIso(t.value)) { mudarDia(t.value); render(); } return; }
  if (t.id === "concil-add") {
    if (!t.value) return;
    const est = estadoDia(); (est.extras ||= []).push(t.value); est.concil[t.value] ||= {}; salvarDia(); return render();
  }
  if (t.dataset?.campo === "check") {
    const est = estadoDia(), et = etapasDoDia();
    est.checks[t.dataset.k] = t.checked; agendarSalvar();
    const x = $("#etapas-txt"); if (x) x.textContent = `${et.filter((e) => est.checks[e]).length} de ${et.length} concluídas`;
  }
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
{ const h = location.hash.replace("#", ""); if (VISOES[h]) S.aba = h; else if (!S.dados) S.aba = "planilha"; }
render();
