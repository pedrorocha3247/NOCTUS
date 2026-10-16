/**
 * NOCTUS: página de conta.
 *
 *   #email      Cadastro do e-mail de recuperação. Aparece UMA vez para quem
 *               ainda está com o e-mail provisório (contas antigas e contas
 *               recém-criadas). Quem já tem e-mail confirmado volta direto.
 *   #recuperar  "Esqueci a senha": e-mail cadastrado → código de 6 dígitos →
 *               senha nova.
 *
 * Tudo usa o Supabase Auth direto (os códigos são gerados e conferidos por
 * ele). O e-mail sai pelo SMTP configurado no Supabase (Gmail do NOCTUS).
 *
 * Por que código e não link: o antivírus do Outlook corporativo abre os links
 * dos e-mails para testar, e isso gasta um link de uso único antes da pessoa
 * clicar. Código digitado não sofre disso.
 */

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";

const CHAVE_SESSAO = "sipep.session"; // mesma chave do app principal
const ESPERA_REENVIO = 60; // segundos: o Supabase recusa reenvio antes disso
const EMAIL_OK = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
// E-mails provisórios gerados pelo sistema: não são de uma pessoa.
const EMAIL_PROVISORIO = /@noctus\.(app|invalid)$|^noctus\.sistema\+/i;

const params = new URLSearchParams(location.search);
const VOLTA = params.get("volta") === "homolog" ? "../homolog/" : "../";

const $ = (id) => document.getElementById(id);
const app = $("app");
$("ano").textContent = new Date().getFullYear();

/* ------------------------------------------------------------------ util */

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function campo({ id, label, type = "text", value = "", extra = "", classe = "" }) {
  return `<div class="ui-field ${classe}">
    <input id="${id}" type="${type}" class="ui-field__input" placeholder=" " value="${esc(value)}" ${extra}>
    <label for="${id}" class="ui-field__label">${label}</label>
  </div>`;
}

function alerta(tipo, texto) {
  const icone = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`;
  return `<div class="ui-alert ui-alert--${tipo}" role="alert">${icone}<span>${texto}</span></div>`;
}

function mostrarErro(texto) {
  const caixa = $("msg");
  if (caixa) caixa.innerHTML = texto ? alerta("error", texto) : "";
}

function carregando(botao, sim, rotulo) {
  if (!botao) return;
  botao.disabled = sim;
  botao.innerHTML = sim
    ? `<span class="ui-button__spinner" aria-hidden="true"></span><span class="ui-button__label--loading">${rotulo}</span>`
    : `<span>${rotulo}</span>`;
}

/** Contagem regressiva no link "Reenviar código". */
function travarReenvio(botao) {
  if (!botao) return;
  let falta = ESPERA_REENVIO;
  botao.disabled = true;
  const tique = () => {
    if (!document.body.contains(botao)) return clearInterval(id);
    if (falta <= 0) {
      clearInterval(id);
      botao.disabled = false;
      botao.textContent = "Reenviar código";
      return;
    }
    botao.textContent = `Reenviar código (${falta}s)`;
    falta--;
  };
  const id = setInterval(tique, 1000);
  tique();
}

/** Chamada ao Supabase Auth. Devolve {ok, status, dados}. */
async function auth(caminho, { method = "POST", body, token } = {}) {
  const headers = { apikey: SB_KEY, "Content-Type": "application/json" };
  if (token) headers.Authorization = "Bearer " + token;
  let r;
  try {
    r = await fetch(SB_URL + "/auth/v1/" + caminho, { method, headers, body: body ? JSON.stringify(body) : undefined });
  } catch {
    return { ok: false, status: 0, dados: {} };
  }
  const dados = await r.json().catch(() => ({}));
  return { ok: r.ok, status: r.status, dados };
}

/** Traduz os erros do Supabase Auth para mensagens que a pessoa entende. */
function traduzir(res, contexto) {
  const d = res.dados || {};
  const cod = d.error_code || d.code || "";
  const msg = String(d.msg || d.message || d.error_description || d.error || "");
  if (res.status === 0) return "Sem conexão com o servidor. Verifique a internet e tente de novo.";
  const seg = msg.match(/after (\d+) seconds?/i);
  if (seg) return `Por segurança, espere ${seg[1]} segundos antes de pedir outro código.`;
  if (res.status === 429 || /rate.?limit/i.test(cod + msg))
    return "Muitos pedidos em pouco tempo. Aguarde alguns minutos e tente de novo.";
  if (cod === "email_exists" || /already been registered|already registered/i.test(msg))
    return "Este e-mail já está cadastrado em outra conta. Use outro e-mail.";
  if (cod === "email_address_invalid" || /invalid.*email|email.*invalid/i.test(msg)) return "E-mail inválido. Confira o endereço.";
  if (cod === "otp_expired" || /expired|invalid.*token|token.*invalid/i.test(msg))
    return "Código incorreto ou vencido. Confira os 6 dígitos ou peça um código novo.";
  if (cod === "same_password" || /different from the old/i.test(msg)) return "A nova senha precisa ser diferente da atual.";
  if (cod === "weak_password" || /weak|at least/i.test(msg)) return "Senha fraca. Use pelo menos 6 caracteres.";
  if (res.status === 401 || cod === "bad_jwt" || cod === "session_not_found")
    return contexto === "sessao" ? "Sua sessão expirou. Entre de novo para continuar." : "Não autorizado. Tente de novo.";
  if (cod === "email_address_not_authorized")
    return "O envio de e-mails ainda não está liberado para este endereço. Avise o administrador.";
  return "Não foi possível concluir agora. Tente de novo em instantes.";
}

/* --------------------------------------------------------------- sessão */

function lerSessao() {
  try {
    return JSON.parse(localStorage.getItem(CHAVE_SESSAO) || "null");
  } catch {
    return null;
  }
}

function gravarSessao(s) {
  localStorage.setItem(CHAVE_SESSAO, JSON.stringify(s));
}

/** Token de acesso válido; renova pelo refresh token se estiver vencendo. */
async function tokenValido() {
  const s = lerSessao();
  if (!s) return null;
  if (s.accessToken && s.expiresAt - Date.now() > 30_000) return s.accessToken;
  if (!s.refreshToken) return null;
  const r = await auth("token?grant_type=refresh_token", { body: { refresh_token: s.refreshToken } });
  if (!r.ok || !r.dados.access_token) return null;
  s.accessToken = r.dados.access_token;
  s.refreshToken = r.dados.refresh_token || s.refreshToken;
  s.expiresAt = Date.now() + (r.dados.expires_in || 3600) * 1000;
  gravarSessao(s);
  return s.accessToken;
}

async function sair() {
  const s = lerSessao();
  if (s?.accessToken) await auth("logout", { token: s.accessToken });
  localStorage.removeItem(CHAVE_SESSAO);
  location.replace(VOLTA);
}

/* ------------------------------------------- #email: cadastro do e-mail */

function telaEmail() {
  const s = lerSessao();
  if (!s) return location.replace(VOLTA); // sem login: volta para entrar
  if (!s.precisaEmail) return location.replace(VOLTA); // já cadastrou: não pede de novo

  const nome = esc(s.user?.displayName || "");
  app.innerHTML = `
    <h1 class="conta__titulo">Cadastre seu e-mail</h1>
    <p class="conta__texto">${nome ? `Olá, <strong>${nome}</strong>. ` : ""}Antes de acessar o NOCTUS, informe seu e-mail.
      Ele será usado para recuperar sua senha se você esquecer. <strong>Você só precisa fazer isso uma vez.</strong></p>
    <form class="conta__form" id="form" novalidate>
      <div id="msg"></div>
      ${campo({ id: "email", label: "E-mail", type: "email", value: s.emailPendente || "", extra: 'autocomplete="email" autofocus' })}
      <button type="submit" class="ui-button ui-button--primary" id="enviar"><span>Enviar código</span></button>
    </form>
    <div class="conta__links conta__links--centro"><button type="button" class="linkbtn" id="sair">Sair</button></div>`;

  $("sair").onclick = sair;
  $("form").onsubmit = async (ev) => {
    ev.preventDefault();
    mostrarErro("");
    const email = $("email").value.trim().toLowerCase();
    if (!EMAIL_OK.test(email) || EMAIL_PROVISORIO.test(email)) return mostrarErro("Informe um e-mail válido.");
    carregando($("enviar"), true, "Enviando...");
    const erro = await pedirTrocaEmail(email);
    carregando($("enviar"), false, "Enviar código");
    if (erro) return mostrarErro(erro);
    telaEmailCodigo(email);
  };

  // Veio do cadastro com o e-mail já digitado: manda o código direto.
  if (s.emailPendente && !sessionStorage.getItem("noctus.conta.enviado")) {
    sessionStorage.setItem("noctus.conta.enviado", "1");
    $("form").requestSubmit();
  }
}

/** Pede ao Supabase a troca do e-mail; ele manda o código para o e-mail novo. */
async function pedirTrocaEmail(email) {
  const token = await tokenValido();
  if (!token) {
    localStorage.removeItem(CHAVE_SESSAO);
    setTimeout(() => location.replace(VOLTA), 2500);
    return "Sua sessão expirou. Entre de novo para continuar.";
  }
  const r = await auth("user", { method: "PUT", token, body: { email } });
  return r.ok ? null : traduzir(r, "sessao");
}

function telaEmailCodigo(email) {
  app.innerHTML = `
    <h1 class="conta__titulo">Confirme seu e-mail</h1>
    <p class="conta__texto">Enviamos um código de 6 dígitos para <strong>${esc(email)}</strong>.
      Ele vale por 10 minutos. Se não aparecer, confira também o lixo eletrônico.</p>
    <form class="conta__form" id="form" novalidate>
      <div id="msg"></div>
      ${campo({ id: "codigo", label: "Código", classe: "conta__codigo", extra: 'inputmode="numeric" autocomplete="one-time-code" maxlength="6" autofocus' })}
      <button type="submit" class="ui-button ui-button--primary" id="confirmar"><span>Confirmar</span></button>
    </form>
    <div class="conta__links">
      <button type="button" class="linkbtn" id="outro">Usar outro e-mail</button>
      <button type="button" class="linkbtn" id="reenviar">Reenviar código</button>
    </div>`;

  travarReenvio($("reenviar"));
  $("outro").onclick = () => {
    const s = lerSessao();
    if (s) {
      delete s.emailPendente;
      gravarSessao(s);
    }
    telaEmail();
  };
  $("reenviar").onclick = async () => {
    mostrarErro("");
    const erro = await pedirTrocaEmail(email);
    if (erro) return mostrarErro(erro);
    $("msg").innerHTML = alerta("info", "Código reenviado.");
    travarReenvio($("reenviar"));
  };
  $("codigo").oninput = (e) => (e.target.value = e.target.value.replace(/\D/g, "").slice(0, 6));
  $("form").onsubmit = async (ev) => {
    ev.preventDefault();
    mostrarErro("");
    const codigo = $("codigo").value.trim();
    if (!/^\d{6}$/.test(codigo)) return mostrarErro("Digite os 6 números do código.");
    carregando($("confirmar"), true, "Confirmando...");
    const r = await auth("verify", { body: { type: "email_change", email, token: codigo } });
    carregando($("confirmar"), false, "Confirmar");
    if (!r.ok) return mostrarErro(traduzir(r));
    // Se a "troca segura" estivesse ligada, o Supabase pediria também o código
    // do e-mail antigo (que é provisório e não recebe nada).
    if (!r.dados.access_token && /other email/i.test(r.dados.msg || "")) {
      return mostrarErro("O servidor pediu confirmação também no e-mail antigo. Avise o administrador.");
    }
    const s = lerSessao() || {};
    if (r.dados.access_token) {
      s.accessToken = r.dados.access_token;
      s.refreshToken = r.dados.refresh_token || s.refreshToken;
      s.expiresAt = Date.now() + (r.dados.expires_in || 3600) * 1000;
    }
    delete s.precisaEmail;
    delete s.emailPendente;
    gravarSessao(s);
    sessionStorage.removeItem("noctus.conta.enviado");
    app.innerHTML = `
      <h1 class="conta__titulo">E-mail confirmado</h1>
      ${alerta("sucesso", `Pronto. Se esquecer a senha, o código vai para <strong>${esc(email)}</strong>.`)}
      <div style="margin-top:1.1rem"><a class="ui-button ui-button--primary" href="${VOLTA}">Continuar para o NOCTUS</a></div>`;
    setTimeout(() => location.replace(VOLTA), 2500);
  };
}

/* ----------------------------------------- #recuperar: esqueci a senha */

function telaRecuperar(emailInicial = "") {
  app.innerHTML = `
    <h1 class="conta__titulo">Esqueci a senha</h1>
    <p class="conta__texto">Informe o e-mail cadastrado na sua conta. Vamos enviar um código para você criar uma senha nova.</p>
    <form class="conta__form" id="form" novalidate>
      <div id="msg"></div>
      ${campo({ id: "email", label: "E-mail cadastrado", type: "email", value: emailInicial, extra: 'autocomplete="email" autofocus' })}
      <button type="submit" class="ui-button ui-button--primary" id="enviar"><span>Enviar código</span></button>
    </form>
    <div class="conta__links conta__links--centro"><a class="linkbtn" href="${VOLTA}">Voltar para o login</a></div>`;

  $("form").onsubmit = async (ev) => {
    ev.preventDefault();
    mostrarErro("");
    const email = $("email").value.trim().toLowerCase();
    if (!EMAIL_OK.test(email)) return mostrarErro("Informe um e-mail válido.");
    carregando($("enviar"), true, "Enviando...");
    const erro = await pedirCodigoRecuperacao(email);
    carregando($("enviar"), false, "Enviar código");
    if (erro) return mostrarErro(erro);
    telaRecuperarCodigo(email);
  };
}

/**
 * Pede o código de recuperação. A resposta para a pessoa é sempre a mesma,
 * exista ou não o e-mail: assim a tela não serve para descobrir quem tem
 * conta. Só erro de limite/conexão é mostrado.
 */
async function pedirCodigoRecuperacao(email) {
  if (EMAIL_PROVISORIO.test(email)) return null; // não é de ninguém: finge que enviou
  const r = await auth("recover", { body: { email } });
  if (r.ok) return null;
  if (r.status === 0 || r.status === 429 || /after \d+ seconds?/i.test(r.dados?.msg || "")) return traduzir(r);
  return null;
}

function telaRecuperarCodigo(email) {
  app.innerHTML = `
    <h1 class="conta__titulo">Crie uma senha nova</h1>
    <p class="conta__texto">Se <strong>${esc(email)}</strong> estiver cadastrado, você vai receber um código de 6 dígitos em instantes.
      Ele vale por 10 minutos. Confira também o lixo eletrônico.</p>
    <form class="conta__form" id="form" novalidate>
      <div id="msg"></div>
      ${campo({ id: "codigo", label: "Código", classe: "conta__codigo", extra: 'inputmode="numeric" autocomplete="one-time-code" maxlength="6" autofocus' })}
      ${campo({ id: "senha", label: "Nova senha", type: "password", extra: 'autocomplete="new-password"' })}
      ${campo({ id: "senha2", label: "Confirme a nova senha", type: "password", extra: 'autocomplete="new-password"' })}
      <button type="submit" class="ui-button ui-button--primary" id="salvar"><span>Salvar nova senha</span></button>
    </form>
    <div class="conta__links">
      <button type="button" class="linkbtn" id="outro">Usar outro e-mail</button>
      <button type="button" class="linkbtn" id="reenviar">Reenviar código</button>
    </div>`;

  travarReenvio($("reenviar"));
  $("outro").onclick = () => telaRecuperar(email);
  $("reenviar").onclick = async () => {
    mostrarErro("");
    const erro = await pedirCodigoRecuperacao(email);
    if (erro) return mostrarErro(erro);
    $("msg").innerHTML = alerta("info", "Se o e-mail estiver cadastrado, um código novo foi enviado.");
    travarReenvio($("reenviar"));
  };
  $("codigo").oninput = (e) => (e.target.value = e.target.value.replace(/\D/g, "").slice(0, 6));
  $("form").onsubmit = async (ev) => {
    ev.preventDefault();
    mostrarErro("");
    const codigo = $("codigo").value.trim();
    const senha = $("senha").value;
    const senha2 = $("senha2").value;
    if (!/^\d{6}$/.test(codigo)) return mostrarErro("Digite os 6 números do código.");
    if (senha.length < 6) return mostrarErro("A senha deve ter ao menos 6 caracteres.");
    if (senha.length > 72) return mostrarErro("A senha pode ter no máximo 72 caracteres.");
    if (senha !== senha2) return mostrarErro("As senhas informadas não coincidem.");

    carregando($("salvar"), true, "Salvando...");
    // 1) O código vira uma sessão temporária de recuperação.
    const v = await auth("verify", { body: { type: "recovery", email, token: codigo } });
    if (!v.ok || !v.dados.access_token) {
      carregando($("salvar"), false, "Salvar nova senha");
      return mostrarErro(traduzir(v));
    }
    // 2) Com ela, grava a senha nova.
    const t = v.dados.access_token;
    const u = await auth("user", { method: "PUT", token: t, body: { password: senha } });
    if (!u.ok) {
      carregando($("salvar"), false, "Salvar nova senha");
      return mostrarErro(traduzir(u));
    }
    // 3) Encerra TODAS as sessões abertas dessa conta (outros computadores
    //    inclusive): se alguém estava usando a conta, perde o acesso agora.
    await auth("logout?scope=global", { token: t });
    localStorage.removeItem(CHAVE_SESSAO);

    app.innerHTML = `
      <h1 class="conta__titulo">Senha alterada</h1>
      ${alerta("sucesso", "Pronto. Entre com o seu <strong>nome de usuário</strong> e a senha nova.")}
      <div style="margin-top:1.1rem"><a class="ui-button ui-button--primary" href="${VOLTA}">Ir para o login</a></div>`;
  };
}

/* ---------------------------------------------------------------- rotas */

function rota() {
  const etapa = location.hash.replace(/^#/, "");
  if (etapa === "email") return telaEmail();
  return telaRecuperar();
}

window.addEventListener("hashchange", rota);
rota();
