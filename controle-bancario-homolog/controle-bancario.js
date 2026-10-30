/**
 * Controle Bancário — NOCTUS. Parte 1 (30/10/2026): importação do relatório de
 * fluxo de caixa e tabela das empresas com saídas na data do relatório.
 *
 * Acesso: enquanto o módulo está em construção, SÓ o administrador principal
 * (Rocha) vê o conteúdo; para os demais a página continua "Módulo em
 * desenvolvimento". A trava da tela é só aparência: quem garante é o banco
 * (RLS da tabela controle_bancario_noctus aceita apenas o usuário do Rocha).
 *
 * Dados: cada data importada vira uma linha no servidor (usuário + data);
 * reimportar a mesma data substitui a anterior. Aparece em qualquer computador.
 *
 * Parte 2 (30/10/2026): o mesmo campo aceita o relatório de débito em conta (SCK).
 * O arquivo é reconhecido pelo conteúdo; o total de cada empresa vai para
 * controle_bancario_debito_noctus e a tabela resumo compara Fluxo × Relatório
 * da mesma data (ver debito.js).
 */
import { lerFluxoCaixa, linhasDoExcel } from "./planilha.js";
import { ehRelatorioDebito, lerDebito, montarResumo } from "./debito.js";

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";
const TABELA = "controle_bancario_noctus";
const TABELA_DEBITO = "controle_bancario_debito_noctus";
const ADMIN_USUARIO = "pedro rocha";
const ADMIN_EMAIL = "pedro.rocha@momentum.com.br";
const TAMANHO_MAXIMO = 10 * 1024 * 1024;   // 10 MB: o relatório real tem ~10 KB

const MSG_FALHA_COMUNICACAO =
  "Falha de comunicação com o servidor. Tente novamente em instantes; se o problema continuar, contate o administrador.";
const MSG_FALHA_GERAL =
  "Ocorreu uma falha inesperada. Tente novamente; se o problema continuar, contate o administrador.";

const $ = (id) => document.getElementById(id);
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const moeda = (v) => Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
/** formato contábil da planilha do Rocha: zero vira "-" */
const contabil = (v) => (Math.abs(Number(v || 0)) < 0.005 ? "-" : moeda(v));
const rotuloIso = (iso) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`;
const plural = (n, um, varios) => `${n} ${n === 1 ? um : varios}`;
const listaE = (xs) => (xs.length <= 1 ? xs.join("") : `${xs.slice(0, -1).join(", ")} e ${xs[xs.length - 1]}`);

/* ------------------------------------------------------------ erros */
function registrarErro(e, oQueFalhou) {
  try { window.NoctusErros?.registrar(e, oQueFalhou); } catch { /* sem registro, segue */ }
}
/** Texto seguro para a tela: aviso legível passa; falha técnica vira mensagem padrão e vai ao registro. */
function msgErro(e, oQueFalhou, padrao = MSG_FALHA_COMUNICACAO) {
  if (e?.aviso) return e.message;
  if (e?.message === "sem-login") return `${oQueFalhou}: sua sessão expirou. Entre de novo no NOCTUS e repita a operação.`;
  registrarErro(e, oQueFalhou);
  return `${oQueFalhou}. ${padrao}`;
}
/* Toda mensagem leva o × para removê-la (regra do Rocha, 30/10/2026). O clique é tratado em ligarTela(). */
const ICONE_X = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" ' +
  'stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>';
const alerta = (tipo, html) => `<div class="alerta ${tipo}" role="${tipo === "erro" ? "alert" : "status"}">` +
  `<span class="alerta__texto">${html}</span>` +
  `<button type="button" class="alerta__x" aria-label="Remover mensagem" title="Remover mensagem">${ICONE_X}</button></div>`;

/* ------------------------------------------------------------ sessão (mesmo login do painel) */
function sessaoGravada() {
  try { return JSON.parse(localStorage.getItem(CHAVE_SESSAO) || "null"); } catch { return null; }
}
function sessaoAtual() {
  const s = sessaoGravada();
  return s && s.accessToken && s.expiresAt && s.expiresAt > Date.now() ? s : null;
}
async function renovarSessao() {
  const s = sessaoGravada();
  if (!s?.refreshToken) return null;
  try {
    const r = await fetch(`${SB_URL}/auth/v1/token?grant_type=refresh_token`, {
      method: "POST",
      headers: { apikey: SB_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ refresh_token: s.refreshToken }),
    });
    const d = await r.json().catch(() => ({}));
    if (!r.ok || !d.access_token) return null;
    const nova = { ...s, accessToken: d.access_token, refreshToken: d.refresh_token || s.refreshToken,
                   expiresAt: Date.now() + (d.expires_in || 3600) * 1000 };
    try { localStorage.setItem(CHAVE_SESSAO, JSON.stringify(nova)); } catch { /* segue com a sessão em memória */ }
    return nova;
  } catch { return null; }
}
async function tokenServidor() {
  const s = sessaoAtual() || (await renovarSessao());
  return s?.accessToken || null;
}
function idDoToken(token) {
  try {
    const b64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(atob(b64)).sub || null;
  } catch { return null; }
}

/** O papel guardado no navegador diz que é o administrador principal? (só decide o que tentar) */
function ehPrincipalNoNavegador() {
  const u = sessaoGravada()?.user;
  if (!u || !Array.isArray(u.roles) || !u.roles.includes("admin")) return false;
  return [u.username, u.displayName, u.id].some((v) => String(v || "").trim().toLowerCase() === ADMIN_USUARIO);
}
/** Confirma no servidor quem é o dono do token: "sim", "nao" ou "falha" (sem comunicação). */
async function principalConfirmado() {
  const token = await tokenServidor();
  if (!token) return "nao";
  try {
    const r = await fetch(`${SB_URL}/auth/v1/user`, { headers: { apikey: SB_KEY, Authorization: `Bearer ${token}` } });
    if (r.status === 401 || r.status === 403) return "nao";
    if (!r.ok) return "falha";
    const j = await r.json();
    return String(j?.email || "").trim().toLowerCase() === ADMIN_EMAIL ? "sim" : "nao";
  } catch { return "falha"; }
}

async function servidor(caminho, opcoes = {}, tentouRenovar = false) {
  const token = await tokenServidor();
  if (!token) throw new Error("sem-login");
  const r = await fetch(`${SB_URL}/rest/v1/${caminho}`, {
    ...opcoes,
    headers: { apikey: SB_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opcoes.headers || {}) },
  });
  if (r.status === 401 && !tentouRenovar) {
    if (await renovarSessao()) return servidor(caminho, opcoes, true);
    throw new Error("sem-login");
  }
  if (!r.ok) throw new Error(`HTTP ${r.status} em ${caminho.split("?")[0]}`);
  return r;
}

/* ------------------------------------------------------------ estado */
// pedido do Rocha (30/10): no máximo 5 dias por vez (3 no celular, para caber a data com o ano), setas para andar
const ehCelular = () => { try { return window.matchMedia("(max-width:560px)").matches; } catch { return false; } };
const visiveis = () => (ehCelular() ? 3 : 5);
const estado = {
  datas: new Map(),   // iso → { iso, empresas, arquivo, atualizado_em, naoSalvo }
  debitos: new Map(), // iso → { iso, empresas: [{ codigo, empresa, total, qtd, soma, confere }], arquivo, naoSalvo }
  atual: null,        // iso escolhida
  inicioDatas: 0,      // 1ª data visível na faixa de datas (no máximo visiveis() por vez)
  calMes: [2026, 0],   // [ano, mês 0-11] mostrado no calendário
};

async function carregarSalvos() {
  const r = await servidor(`${TABELA}?select=data_relatorio,empresas,arquivo,importado_em,atualizado_em` +
    `&order=data_relatorio.desc&limit=60`);
  const linhas = await r.json();
  estado.datas.clear();
  for (const l of linhas) {
    estado.datas.set(l.data_relatorio, { iso: l.data_relatorio, empresas: Array.isArray(l.empresas) ? l.empresas : [],
      arquivo: l.arquivo, atualizado_em: l.atualizado_em });
  }
  // abre na data importada por último
  const ultima = linhas.slice().sort((a, b) => String(b.atualizado_em).localeCompare(String(a.atualizado_em)))[0];
  estado.atual = ultima ? ultima.data_relatorio : null;
}

async function carregarDebitos() {
  const r = await servidor(`${TABELA_DEBITO}?select=data_relatorio,empresas,arquivo,atualizado_em` +
    `&order=data_relatorio.desc&limit=60`);
  estado.debitos.clear();
  for (const l of await r.json()) {
    estado.debitos.set(l.data_relatorio, { iso: l.data_relatorio, empresas: Array.isArray(l.empresas) ? l.empresas : [],
      arquivo: l.arquivo, atualizado_em: l.atualizado_em });
  }
}

async function salvarDebito(deb, arquivo) {
  const token = await tokenServidor();
  const uid = token && idDoToken(token);
  if (!uid) throw new Error("sem-login");
  const r = await servidor(`${TABELA_DEBITO}?on_conflict=auth_user_id,data_relatorio`, {
    method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify([{ auth_user_id: uid, data_relatorio: deb.iso, empresas: deb.empresas,
      arquivo: String(arquivo || "").slice(0, 200) }]),
  });
  return r.json();
}

async function salvar(datas, arquivo) {
  const token = await tokenServidor();
  const uid = token && idDoToken(token);
  if (!uid) throw new Error("sem-login");
  const corpo = datas.map((d) => ({ auth_user_id: uid, data_relatorio: d.iso, empresas: d.empresas,
    arquivo: String(arquivo || "").slice(0, 200) }));
  const r = await servidor(`${TABELA}?on_conflict=auth_user_id,data_relatorio`, {
    method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify(corpo),
  });
  return r.json();
}

/* ------------------------------------------------------------ leitura do Excel */
async function carregarSheetJS() {
  if (window.XLSX) return;
  const origens = [
    "../conferencia/vendor/xlsx.full.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
  ];
  for (const src of origens) {
    try {
      await new Promise((ok, falha) => {
        const s = document.createElement("script");
        s.src = src; s.onload = ok; s.onerror = () => falha(new Error(src));
        document.head.appendChild(s);
      });
      if (window.XLSX) return;
    } catch { /* tenta a próxima origem */ }
  }
  throw new Error(`biblioteca SheetJS não carregou (${origens.join(" | ")})`);
}

async function importar(arquivo) {
  const msg = $("upload-msg");
  const solta = $("solta");
  const nome = String(arquivo?.name || "");
  if (!/\.(xlsx|xls)$/i.test(nome)) {
    msg.innerHTML = alerta("erro", "Envie o relatório em Excel (.xlsx): fluxo de caixa ou débito em conta.");
    return;
  }
  if (arquivo.size > TAMANHO_MAXIMO) {
    msg.innerHTML = alerta("erro", "O arquivo passa de 10 MB, muito maior que os relatórios do Controle Bancário. Confira se escolheu o arquivo certo.");
    return;
  }

  solta.classList.add("lendo");
  msg.innerHTML = `<p class="sub">Lendo o relatório…</p>`;
  let resultado;
  try {
    await carregarSheetJS();
    const buffer = await arquivo.arrayBuffer();
    const linhas = linhasDoExcel(window.XLSX, buffer);
    if (ehRelatorioDebito(linhas)) {
      const deb = lerDebito(window.XLSX, buffer);
      return await importarDebito(deb, nome);
    }
    resultado = lerFluxoCaixa(linhas);
  } catch (e) {
    solta.classList.remove("lendo");
    msg.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui ler o relatório", MSG_FALHA_GERAL)));
    return;
  }

  const jaExistiam = resultado.datas.filter((d) => estado.datas.has(d.iso) && !estado.datas.get(d.iso).naoSalvo);
  let erroSalvar = null;
  try {
    const salvos = await salvar(resultado.datas, nome);
    for (const l of salvos) {
      estado.datas.set(l.data_relatorio, { iso: l.data_relatorio, empresas: l.empresas, arquivo: l.arquivo,
        atualizado_em: l.atualizado_em });
    }
  } catch (e) {
    erroSalvar = e;
    for (const d of resultado.datas) {
      estado.datas.set(d.iso, { iso: d.iso, empresas: d.empresas, arquivo: nome, atualizado_em: null, naoSalvo: true });
    }
  }
  solta.classList.remove("lendo");
  estado.atual = resultado.datas[0].iso;
  desenhar();

  if (erroSalvar) {
    msg.innerHTML = alerta("erro", esc(msgErro(erroSalvar, "O relatório foi lido, mas não consegui salvá-lo no servidor")));
    return;
  }
  const ds = resultado.datas;
  let texto;
  if (ds.length === 1) {
    const d = ds[0];
    const verbo = jaExistiam.length ? "atualizado" : "importado";
    texto = `Relatório de ${d.rotulo} ${verbo}: ${plural(d.empresas.length, "empresa com saída", "empresas com saídas")}.`;
  } else {
    texto = `Relatório importado: ${ds.length} datas (${listaE(ds.map((d) => d.rotulo))}).`;
    if (jaExistiam.length) texto += ` ${listaE(jaExistiam.map((d) => d.rotulo))} já ${jaExistiam.length === 1 ? "existia e foi atualizada" : "existiam e foram atualizadas"}.`;
  }
  msg.innerHTML = alerta("ok", esc(texto));
}

/** Grava o relatório de débito já lido e mostra a comparação da mesma data. */
async function importarDebito(deb, nome) {
  const msg = $("upload-msg");
  const jaExistia = estado.debitos.has(deb.iso) && !estado.debitos.get(deb.iso).naoSalvo;
  let erroSalvar = null;
  try {
    const salvos = await salvarDebito(deb, nome);
    const l = salvos[0];
    estado.debitos.set(l.data_relatorio, { iso: l.data_relatorio, empresas: l.empresas, arquivo: l.arquivo, atualizado_em: l.atualizado_em });
  } catch (e) {
    erroSalvar = e;
    estado.debitos.set(deb.iso, { iso: deb.iso, empresas: deb.empresas, arquivo: nome, naoSalvo: true });
  }
  $("solta").classList.remove("lendo");
  const temFluxo = estado.datas.has(deb.iso);
  if (temFluxo) estado.atual = deb.iso;
  desenhar();

  if (erroSalvar) {
    msg.innerHTML = alerta("erro", esc(msgErro(erroSalvar, "O relatório de débito foi lido, mas não consegui salvá-lo no servidor")));
    return;
  }
  const empresas = deb.empresas.filter((e) => e.codigo);
  let texto = `Relatório de débito de ${deb.rotulo} ${jaExistia ? "atualizado" : "importado"}: ` +
    `${plural(empresas.length, "empresa", "empresas")}, total R$ ${moeda(deb.total)}.`;
  if (!temFluxo) texto += ` Ainda não há fluxo de caixa de ${deb.rotulo}; a comparação aparece quando ele for importado.`;
  msg.innerHTML = alerta("ok", esc(texto));
}

/* ------------------------------------------------------------ tela */
/** Anda a faixa de datas sem trocar a data aberta. */
function botaoData(iso) {
  const longo = rotuloIso(iso);
  return `<button type="button" class="data-btn" data-iso="${esc(iso)}" aria-pressed="${iso === estado.atual}" ` +
    `aria-label="${esc(longo)}">${esc(longo)}</button>`;
}
function desenharDatasSemMudar() {
  const isos = [...estado.datas.keys()].sort();
  const maxInicio = Math.max(0, isos.length - visiveis());
  const ini = Math.min(Math.max(0, estado.inicioDatas), maxInicio);
  estado.inicioDatas = ini;
  $("datas").innerHTML = isos.slice(ini, ini + visiveis()).map((iso) =>
    botaoData(iso)
  ).join("");
  $("datas-ant").disabled = ini === 0;
  $("datas-prox").disabled = ini >= maxInicio;
}

function desenhar() {
  const cartao = $("cartao-tabela");
  if (!estado.atual || !estado.datas.has(estado.atual)) {
    cartao.classList.add("oculto");
    $("cartao-resumo")?.classList.add("oculto");
    return;
  }
  cartao.classList.remove("oculto");

  // escolha da data: só aparece quando há mais de uma
  const isos = [...estado.datas.keys()].sort();
  const caixaDatas = $("datas");
  const navDatas = $("datas-nav") || caixaDatas;   // defensivo: HTML antigo não tem a faixa com setas
  if (isos.length > 1) {
    navDatas.classList.remove("oculto");
    // janela de no máximo 5 datas; a data aberta sempre fica dentro da janela
    const maxInicio = Math.max(0, isos.length - visiveis());
    const posAtual = isos.indexOf(estado.atual);
    let ini = Math.min(Math.max(0, estado.inicioDatas), maxInicio);
    if (posAtual >= 0 && (posAtual < ini || posAtual >= ini + visiveis())) ini = Math.min(maxInicio, Math.max(0, posAtual - visiveis() + 1));
    estado.inicioDatas = ini;
    caixaDatas.innerHTML = isos.slice(ini, ini + visiveis()).map((iso) =>
      botaoData(iso)
    ).join("");
    const ant = $("datas-ant"), prox = $("datas-prox");
    if (ant) { ant.disabled = ini === 0; ant.classList.toggle("oculto", isos.length <= visiveis()); }
    if (prox) { prox.disabled = ini >= maxInicio; prox.classList.toggle("oculto", isos.length <= visiveis()); }
  } else {
    navDatas.classList.add("oculto");
    caixaDatas.innerHTML = "";
  }

  const reg = estado.datas.get(estado.atual);
  const empresas = reg.empresas || [];
  const total = Math.round(empresas.reduce((s, e) => s + Number(e.saidas || 0), 0) * 100) / 100;
  // defensivo: na publicação o JS vai ao ar antes do HTML que tem esses campos
  const por = (id, texto) => { const el = $(id); if (el) el.textContent = texto; };
  por("tabela-titulo", `Saídas em ${rotuloIso(reg.iso)}`);
  por("tabela-total", `R$ ${moeda(total)}`);
  por("tabela-qtd", empresas.length ? plural(empresas.length, "empresa", "empresas") : "Nenhuma empresa");
  $("tabela-aviso").innerHTML = reg.naoSalvo
    ? alerta("erro", "Esta data ainda não foi salva no servidor. Importe o relatório de novo quando a comunicação voltar.")
    : "";
  $("tabela-corpo").innerHTML = empresas.length
    ? empresas.map((e) => `<tr><td>${esc(e.empresa)}</td><td class="saidas">${esc(moeda(e.saidas))}</td></tr>`).join("")
    : `<tr><td colspan="2" class="vazio">Nenhuma empresa com saída em ${esc(rotuloIso(reg.iso))}.</td></tr>`;
  desenharResumo(reg);
}

/** Tabela resumo (modelo B): Fluxo × Relatório de débito da data aberta. */
function desenharResumo(reg) {
  const cartao = $("cartao-resumo");
  if (!cartao) return;   // defensivo: na publicação o JS vai ao ar antes do HTML
  cartao.classList.remove("oculto");
  const deb = estado.debitos.get(reg.iso) || null;
  const res = montarResumo(reg.empresas || [], deb);
  const sinal = (v) => (v == null || v === 0 ? "" : v > 0 ? "pos" : "neg");
  const dica = (v) => (v > 0 ? "Fluxo maior que o relatório de débito" : v < 0 ? "Relatório de débito maior que o fluxo" : "");

  $("resumo-titulo").textContent = `Fluxo em ${rotuloIso(reg.iso)}`;
  $("resumo-fluxo").textContent = `R$ ${moeda(res.total.fluxo)}`;
  $("resumo-relatorio").textContent = deb ? `R$ ${moeda(res.total.relatorio)}` : "—";
  const dif = $("resumo-diferenca");
  dif.textContent = deb ? `R$ ${moeda(res.total.diferenca)}` : "—";
  dif.className = `kpi-val ${sinal(res.total.diferenca)}`;
  $("resumo-situacao").textContent = !deb ? "Aguardando o relatório de débito"
    : res.divergencias ? plural(res.divergencias, "divergência", "divergências") : "Sem divergência";

  const avisos = [];
  if (!deb) avisos.push(alerta("", `Importe o relatório de débito em conta de ${esc(rotuloIso(reg.iso))} para preencher a linha Relatório.`));
  else if (deb.naoSalvo) avisos.push(alerta("erro", "O relatório de débito desta data ainda não foi salvo no servidor. Importe de novo quando a comunicação voltar."));
  for (const a of res.avisos) {
    if (a.tipo === "semFluxo") avisos.push(alerta("", `${esc(a.empresa)} tem débito de R$ ${esc(moeda(a.total))} no relatório, sem valor no fluxo de ${esc(rotuloIso(reg.iso))}.`));
    if (a.tipo === "semConta") avisos.push(alerta("erro", `R$ ${esc(moeda(a.total))} debitados em conta bancária não cadastrada no NOCTUS: não entraram em nenhuma empresa. Confira o relatório.`));
    if (a.tipo === "naoConfere") avisos.push(alerta("erro", `${esc(a.empresa)}: o total impresso no relatório (R$ ${esc(moeda(a.total))}) não bate com a soma das solicitações (R$ ${esc(moeda(a.soma))}). Confira o relatório.`));
    if (a.tipo === "semCodigo") avisos.push(alerta("", `${esc(a.empresa)} está no fluxo, mas não tem código de empresa cadastrado: não dá para comparar com o relatório.`));
  }
  $("resumo-aviso").innerHTML = avisos.join("");

  const cols = res.cols;
  const cab = cols.map((c) => `<th scope="col"${c.semFluxo ? ' class="sem-fluxo"' : ""}>${esc(c.empresa)}` +
    `${c.semFluxo ? '<span class="marca">sem fluxo</span>' : ""}</th>`).join("");
  const linha = (rotulo, classe, valor) =>
    `<tr class="${classe}"><th scope="row">${rotulo}</th>${cols.map(valor).join("")}</tr>`;
  $("resumo-tabela").innerHTML =
    `<thead><tr><th scope="col">Empresa</th>${cab}</tr></thead><tbody>` +
    linha("Fluxo", "fluxo", (c) => `<td>${esc(contabil(c.fluxo))}</td>`) +
    linha("Relatório", "relatorio", (c) => `<td>${c.relatorio == null ? "—" : esc(contabil(c.relatorio))}</td>`) +
    linha("Diferença", "diferenca", (c) => c.diferenca == null ? "<td>—</td>"
      : `<td><span class="dif ${sinal(c.diferenca) || "zero"}"${dica(c.diferenca) ? ` title="${dica(c.diferenca)}"` : ""}>${esc(contabil(c.diferenca))}</span></td>`) +
    "</tbody>";
  ajustarSetasEmpresas();
}

/** Setas nas pontas da linha das empresas: só aparecem quando a tabela não cabe; somem na ponta. */
function ajustarSetasEmpresas() {
  const caixa = $("resumo-caixa"), ant = $("emp-ant"), prox = $("emp-prox");
  if (!caixa || !ant || !prox) return;
  const sobra = caixa.scrollWidth - caixa.clientWidth;
  const rotulo = caixa.querySelector("thead th");
  ant.style.left = `${(rotulo ? rotulo.offsetWidth : 0) + 6}px`;   // logo depois da coluna fixa "EMPRESA"
  ant.classList.toggle("oculto", sobra <= 1 || caixa.scrollLeft <= 1);
  prox.classList.toggle("oculto", sobra <= 1 || caixa.scrollLeft >= sobra - 1);
}
function andarEmpresas(sentido) {
  const caixa = $("resumo-caixa");
  if (!caixa) return;
  const rotulo = caixa.querySelector("thead th");
  const passo = Math.max(120, caixa.clientWidth - (rotulo ? rotulo.offsetWidth : 0) - 40);
  caixa.scrollBy({ left: sentido * passo, behavior: "smooth" });
}

/* ------------------------------------------------------------ calendário (modelo A, 30/10) */
const MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
const isoDe = (a, m, d) => `${a}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
function hojeIso() { const h = new Date(); return isoDe(h.getFullYear(), h.getMonth(), h.getDate()); }
function desenharCalendario() {
  const caixa = $("cal-dias");
  if (!caixa) return;
  const [a, m] = estado.calMes;
  $("cal-mes").textContent = `${MESES[m][0].toUpperCase()}${MESES[m].slice(1)} de ${a}`;
  const primeiro = new Date(a, m, 1).getDay();       // 0 = domingo
  const hoje = hojeIso();
  let html = "";
  for (let i = 0; i < 42; i++) {
    const dt = new Date(a, m, 1 - primeiro + i);
    const iso = isoDe(dt.getFullYear(), dt.getMonth(), dt.getDate());
    const doMes = dt.getMonth() === m;
    const tem = estado.datas.has(iso);
    const cls = ["cal-dia", tem ? "tem" : "", doMes ? "" : "fora", iso === hoje ? "hoje" : ""].filter(Boolean).join(" ");
    const rot = `${rotuloIso(iso)}${tem ? ", com fluxo importado" : ", sem fluxo"}`;
    html += `<button type="button" class="${cls}" data-iso="${iso}" aria-label="${rot}"` +
      `${tem ? ` aria-pressed="${iso === estado.atual}"` : " disabled"}>${dt.getDate()}</button>`;
  }
  caixa.innerHTML = html;
}
function abrirCalendario(abrir) {
  const pop = $("cal-pop"), btn = $("btn-cal");
  if (!pop || !btn) return;
  const ja = !pop.classList.contains("oculto");
  const vai = abrir ?? !ja;
  if (vai) {
    const base = estado.atual || hojeIso();
    estado.calMes = [Number(base.slice(0, 4)), Number(base.slice(5, 7)) - 1];
    desenharCalendario();
  }
  pop.classList.toggle("oculto", !vai);
  btn.setAttribute("aria-expanded", String(vai));
  if (vai) (pop.querySelector('.cal-dia[aria-pressed="true"]') || pop.querySelector(".cal-dia.tem") || $("cal-prox"))?.focus();
}
function andarMes(sentido) {
  let [a, m] = estado.calMes; m += sentido;
  if (m < 0) { m = 11; a--; } if (m > 11) { m = 0; a++; }
  estado.calMes = [a, m]; desenharCalendario();
}

function ligarTela() {
  const solta = $("solta");
  const input = $("arquivo");
  solta.addEventListener("click", () => input.click());
  solta.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); input.click(); }
  });
  input.addEventListener("change", () => {
    const f = input.files?.[0];
    input.value = "";
    if (f) importar(f);
  });
  ["dragenter", "dragover"].forEach((t) => solta.addEventListener(t, (ev) => { ev.preventDefault(); solta.classList.add("ativa"); }));
  ["dragleave", "drop"].forEach((t) => solta.addEventListener(t, (ev) => { ev.preventDefault(); solta.classList.remove("ativa"); }));
  solta.addEventListener("drop", (ev) => {
    const f = ev.dataTransfer?.files?.[0];
    if (f) importar(f);
  });
  // × de qualquer mensagem do módulo: remove só aquela mensagem
  $("modulo").addEventListener("click", (ev) => {
    const x = ev.target.closest(".alerta__x");
    if (x) x.closest(".alerta")?.remove();
  });
  const ligar = (id, fn) => { const el = $(id); if (el) el.addEventListener("click", fn); };
  ligar("datas-ant", () => { estado.inicioDatas = Math.max(0, estado.inicioDatas - visiveis()); desenharDatasSemMudar(); });
  ligar("datas-prox", () => { estado.inicioDatas += visiveis(); desenharDatasSemMudar(); });
  ligar("emp-ant", () => andarEmpresas(-1));
  ligar("emp-prox", () => andarEmpresas(1));
  $("resumo-caixa")?.addEventListener("scroll", ajustarSetasEmpresas, { passive: true });
  window.addEventListener("resize", ajustarSetasEmpresas);
  ligar("btn-cal", () => abrirCalendario());
  ligar("cal-ant", () => andarMes(-1));
  ligar("cal-prox", () => andarMes(1));
  ligar("cal-hoje", () => { const h = hojeIso(); estado.calMes = [Number(h.slice(0, 4)), Number(h.slice(5, 7)) - 1]; desenharCalendario(); });
  ligar("cal-dias", (ev) => {
    const d = ev.target.closest(".cal-dia.tem");
    if (!d) return;
    estado.atual = d.dataset.iso; abrirCalendario(false); desenhar();
  });
  document.addEventListener("keydown", (ev) => {
    const pop = $("cal-pop");
    if (ev.key === "Escape" && pop && !pop.classList.contains("oculto")) { abrirCalendario(false); $("btn-cal")?.focus(); }
  });
  document.addEventListener("click", (ev) => {
    const pop = $("cal-pop");
    if (pop && !pop.classList.contains("oculto") && !ev.target.closest("#cal-pop, #btn-cal")) abrirCalendario(false);
  });
  $("datas").addEventListener("click", (ev) => {
    const b = ev.target.closest(".data-btn");
    if (!b) return;
    estado.atual = b.dataset.iso;
    desenhar();
  });
}

/* ------------------------------------------------------------ início */
async function iniciar() {
  // mesmo comportamento da Conferência: quem ainda não cadastrou o e-mail de recuperação vai cadastrar
  if (sessaoGravada()?.precisaEmail) { location.replace("../conta/#email"); return; }
  if (!ehPrincipalNoNavegador()) return;          // demais usuários: "Módulo em desenvolvimento"
  const confirmado = await principalConfirmado();
  if (confirmado === "nao") return;
  $("em-desenvolvimento").classList.add("oculto");
  $("modulo").classList.remove("oculto");
  ligarTela();
  if (confirmado === "falha") {
    $("upload-msg").innerHTML = alerta("erro", esc(msgErro(new Error("confirmação do usuário sem resposta"),
      "Não consegui carregar os relatórios já importados")));
    return;
  }
  try {
    await carregarSalvos();
  } catch (e) {
    $("upload-msg").innerHTML = alerta("erro", esc(msgErro(e, "Não consegui carregar os relatórios já importados")));
    return;
  }
  try {
    await carregarDebitos();
  } catch (e) {
    $("upload-msg").innerHTML = alerta("erro", esc(msgErro(e, "Não consegui carregar os relatórios de débito já importados")));
  }
  desenhar();
}

iniciar().catch((e) => registrarErro(e, "abertura do Controle Bancário"));
