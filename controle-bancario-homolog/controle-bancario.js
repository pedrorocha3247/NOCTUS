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
 */
import { lerFluxoCaixa, linhasDoExcel } from "./planilha.js";

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";
const TABELA = "controle_bancario_noctus";
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
const alerta = (tipo, html) => `<div class="alerta ${tipo}" role="${tipo === "erro" ? "alert" : "status"}">${html}</div>`;

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
const estado = {
  datas: new Map(),   // iso → { iso, empresas, arquivo, atualizado_em, naoSalvo }
  atual: null,        // iso escolhida
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
    msg.innerHTML = alerta("erro", "Envie o relatório de fluxo de caixa em Excel (.xlsx).");
    return;
  }
  if (arquivo.size > TAMANHO_MAXIMO) {
    msg.innerHTML = alerta("erro", "O arquivo passa de 10 MB, muito maior que o relatório de fluxo de caixa. Confira se escolheu o arquivo certo.");
    return;
  }

  solta.classList.add("lendo");
  msg.innerHTML = `<p class="sub">Lendo o relatório…</p>`;
  let resultado;
  try {
    await carregarSheetJS();
    resultado = lerFluxoCaixa(linhasDoExcel(window.XLSX, await arquivo.arrayBuffer()));
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

/* ------------------------------------------------------------ tela */
function desenhar() {
  const cartao = $("cartao-tabela");
  if (!estado.atual || !estado.datas.has(estado.atual)) { cartao.classList.add("oculto"); return; }
  cartao.classList.remove("oculto");

  // escolha da data: só aparece quando há mais de uma
  const isos = [...estado.datas.keys()].sort();
  const caixaDatas = $("datas");
  if (isos.length > 1) {
    caixaDatas.classList.remove("oculto");
    caixaDatas.innerHTML = isos.map((iso) =>
      `<button type="button" class="data-btn" data-iso="${esc(iso)}" aria-pressed="${iso === estado.atual}">${esc(rotuloIso(iso))}</button>`
    ).join("");
  } else {
    caixaDatas.classList.add("oculto");
    caixaDatas.innerHTML = "";
  }

  const reg = estado.datas.get(estado.atual);
  const empresas = reg.empresas || [];
  const total = Math.round(empresas.reduce((s, e) => s + Number(e.saidas || 0), 0) * 100) / 100;
  $("tabela-titulo").textContent = `Empresas com saídas em ${rotuloIso(reg.iso)}`;
  $("tabela-resumo").textContent = empresas.length
    ? `${plural(empresas.length, "empresa", "empresas")} · total R$ ${moeda(total)}`
    : "Nenhuma empresa com saída nesta data.";
  $("tabela-aviso").innerHTML = reg.naoSalvo
    ? alerta("erro", "Esta data ainda não foi salva no servidor. Importe o relatório de novo quando a comunicação voltar.")
    : "";
  $("tabela-corpo").innerHTML = empresas.length
    ? empresas.map((e) => `<tr><td>${esc(e.empresa)}</td><td class="saidas">${esc(moeda(e.saidas))}</td></tr>`).join("")
    : `<tr><td colspan="2" class="vazio">Nenhuma empresa com saída em ${esc(rotuloIso(reg.iso))}.</td></tr>`;
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
    desenhar();
  } catch (e) {
    $("upload-msg").innerHTML = alerta("erro", esc(msgErro(e, "Não consegui carregar os relatórios já importados")));
  }
}

iniciar().catch((e) => registrarErro(e, "abertura do Controle Bancário"));
