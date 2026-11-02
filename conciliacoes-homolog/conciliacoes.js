/**
 * Conciliações Bancárias — NOCTUS. Parte 1 (02/11/2026): escolha da empresa e da
 * conta, importação do relatório de conciliação do SCK e do extrato .OFX, e
 * comparação de saldo, entradas e saídas.
 *
 * Acesso: enquanto o módulo está em construção, SÓ o administrador principal
 * (Rocha) vê o conteúdo; para os demais a página continua "Módulo em
 * desenvolvimento". A trava da tela é só aparência: quem garante é o banco
 * (RLS da tabela conciliacao_bancaria_noctus aceita apenas o usuário do Rocha).
 *
 * Dados: uma linha por conta + data; reimportar a mesma data substitui. É
 * gravado só o RESULTADO (totais, saldos e o resumo do SCK) — os lançamentos
 * individuais não vão para o servidor, porque os históricos trazem nome e
 * CPF/CNPJ de clientes e, para comparar totais, não são necessários.
 *
 * As contas que aparecem aqui vêm de contas.js (Conferência) mais o código do
 * SCK, em contas-sck.js: habilitar uma conta nova é acrescentar uma linha lá.
 */
import { lerConciliacaoSck, linhasDoExcel } from "./sck.js";
import { lerExtratoOfx, textoDoOfx } from "./ofx.js";
import { montarConciliacao, conferirConta, rotuloConta } from "./comparar.js";
import { empresasHabilitadas, contaPorId } from "./contas-sck.js";

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";
const TABELA = "conciliacao_bancaria_noctus";
const ADMIN_USUARIO = "pedro rocha";
const ADMIN_EMAIL = "pedro.rocha@momentum.com.br";
const TAMANHO_MAXIMO = 10 * 1024 * 1024;   // 10 MB: os arquivos reais têm ~20 KB

const MSG_FALHA_COMUNICACAO =
  "Não consegui falar com o servidor agora. Confira a internet e tente de novo em alguns instantes.";
const MSG_FALHA_GERAL =
  "Alguma coisa não saiu como esperado. Tente de novo; se continuar, avise o Rocha.";

const $ = (id) => document.getElementById(id);
const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const moeda = (v) => Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
/** Diferença com sinal à mostra: o + importa tanto quanto o −. */
const comSinal = (v) => (Math.abs(Number(v || 0)) < 0.005 ? "0,00" : `${v > 0 ? "+" : "−"}${moeda(Math.abs(v))}`);
const rotuloIso = (iso) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`;

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
/* Toda mensagem leva o × para removê-la (regra do Rocha, 30/10/2026). */
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
  conta: null,        // conta escolhida (de contas-sck.js)
  sck: null,          // resultado de lerConciliacaoSck
  ofx: null,          // resultado de lerExtratoOfx
  nomeSck: "",
  nomeOfx: "",
  textoOfx: "",      // guardado para reler o extrato quando a data do SCK chegar depois
  naoSalvo: false,
};

/* ------------------------------------------------------------ servidor */
/** Grava o resultado da conciliação. Reimportar a mesma conta e data substitui. */
async function salvar() {
  const { conta, sck, ofx } = estado;
  const dataIso = sck?.data?.iso || ofx?.periodo?.fim?.iso;
  if (!conta || !dataIso) return null;
  const quadro = montarConciliacao({ sck, ofx });
  const linha = {
    conta_sck: conta.codigoSck,
    empresa_codigo: conta.codigo,
    data_conciliacao: dataIso,
    // só o resultado: nada de lançamento com nome ou CPF de cliente
    sck: sck ? { entradas: sck.entradas, saidas: sck.saidas, totalImpresso: sck.totalImpresso,
                 resumo: sck.resumo, status: sck.status, responsavel: sck.responsavel,
                 lancamentos: sck.lancamentos.length } : {},
    extrato: ofx ? { entradas: ofx.entradas, saidas: ofx.saidas,
                     saldoDoDia: ofx.saldoDoDia, saldoAnterior: ofx.saldoAnterior, saldoLivro: ofx.saldoLivro,
                     lancamentos: ofx.lancamentos.length } : {},
    comparacao: { situacao: quadro.situacao, linhas: quadro.linhas.map(({ chave, sck: s, banco, diferenca, situacao }) =>
                    ({ chave, sck: s, banco, diferenca, situacao })) },
    arquivo_sck: estado.nomeSck || null,
    arquivo_ofx: estado.nomeOfx || null,
    atualizado_em: new Date().toISOString(),
  };
  const r = await servidor(`${TABELA}?on_conflict=auth_user_id,conta_sck,data_conciliacao`, {
    method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify([linha]),
  });
  return (await r.json())[0] || null;
}

/** Traz a conciliação já gravada desta conta e data, se houver. */
async function carregarSalva(conta, dataIso) {
  const r = await servidor(`${TABELA}?conta_sck=eq.${encodeURIComponent(conta.codigoSck)}` +
                           `&data_conciliacao=eq.${dataIso}&select=*`);
  return (await r.json())[0] || null;
}

/* ------------------------------------------------------------ empresa e conta */
function montarEscolha() {
  const empresas = empresasHabilitadas();
  const selE = $("sel-empresa"), selC = $("sel-conta");
  selE.innerHTML = empresas.map((e) => `<option value="${esc(e.codigo)}">${esc(e.empresa)}</option>`).join("");
  const trocarEmpresa = () => {
    const emp = empresas.find((e) => e.codigo === selE.value) || empresas[0];
    selC.innerHTML = (emp?.contas || []).map((c) =>
      `<option value="${esc(c.id)}">${esc(rotuloConta(c))}</option>`).join("");
    trocarConta();
  };
  const trocarConta = () => {
    estado.conta = contaPorId(selC.value);
    estado.sck = estado.ofx = null;
    estado.nomeSck = estado.nomeOfx = estado.textoOfx = "";
    $("resumo-sck").innerHTML = "";
    $("resumo-ofx").innerHTML = "";
    $("upload-msg").innerHTML = "";
    $("cartao-arquivos").classList.toggle("oculto", !estado.conta);
    $("cartao-quadro").classList.add("oculto");
    $("conta-detalhe").textContent = estado.conta
      ? `Conta ${estado.conta.codigoSck} no SCK · CNPJ ${estado.conta.cnpj}`
      : "";
  };
  selE.addEventListener("change", trocarEmpresa);
  selC.addEventListener("change", trocarConta);
  trocarEmpresa();

  if (!empresas.length) {
    $("escolha-sub").textContent = "Nenhuma conta está habilitada para conciliação ainda.";
  }
}

/* ------------------------------------------------------------ SheetJS */
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

/* ------------------------------------------------------------ importação */
const ICONE_OK = '<svg class="ok" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" ' +
  'stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 6 9 17l-5-5"/></svg>';

function mostrarResumoArquivo(qual, nome, detalhe) {
  $(`resumo-${qual}`).innerHTML = ICONE_OK +
    `<span class="corpo"><span class="nome">${esc(nome)}</span><span class="detalhe">${esc(detalhe)}</span></span>` +
    `<button type="button" class="trocar" data-trocar="${qual}">Trocar</button>`;
}

async function importarSck(arquivo) {
  const msg = $("upload-msg"), solta = $("solta-sck");
  const nome = String(arquivo?.name || "");
  if (!/\.(xlsx|xls)$/i.test(nome)) {
    msg.innerHTML = alerta("erro", "Envie o relatório de conciliação do SCK exportado em Excel (.xls ou .xlsx).");
    return;
  }
  if (arquivo.size > TAMANHO_MAXIMO) {
    msg.innerHTML = alerta("erro", "O arquivo passa de 10 MB, muito maior que o relatório de conciliação. Confira se escolheu o arquivo certo.");
    return;
  }
  solta.classList.add("lendo");
  msg.innerHTML = `<p class="sub">Lendo o relatório do SCK…</p>`;
  try {
    await carregarSheetJS();
    const linhas = linhasDoExcel(window.XLSX, await arquivo.arrayBuffer());
    const lido = lerConciliacaoSck(linhas);
    conferirConta(estado.conta, { sck: lido });       // arquivo de outra conta é recusado
    estado.sck = lido;
    estado.nomeSck = nome;
    realinharExtrato();                              // o extrato pode ter sido lido antes, sem saber a data
    mostrarResumoArquivo("sck", nome,
      `${lido.data.rotulo} · ${lido.lancamentos.length} lançamento(s) · ${lido.status || "sem status"}`);
    if (estado.ofx) {
      mostrarResumoArquivo("ofx", estado.nomeOfx,
        `${lido.data.rotulo} · ${estado.ofx.lancamentos.length} lançamento(s)`);
    }
    msg.innerHTML = "";
  } catch (e) {
    msg.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui ler o relatório de conciliação", MSG_FALHA_GERAL)));
  } finally {
    solta.classList.remove("lendo");
  }
  await depoisDeImportar();
}

async function importarOfx(arquivo) {
  const msg = $("upload-msg"), solta = $("solta-ofx");
  const nome = String(arquivo?.name || "");
  if (!/\.ofx$/i.test(nome)) {
    msg.innerHTML = alerta("erro", "Envie o extrato no formato .OFX, baixado no site do banco.");
    return;
  }
  if (arquivo.size > TAMANHO_MAXIMO) {
    msg.innerHTML = alerta("erro", "O arquivo passa de 10 MB, muito maior que um extrato diário. Confira se escolheu o arquivo certo.");
    return;
  }
  solta.classList.add("lendo");
  msg.innerHTML = `<p class="sub">Lendo o extrato…</p>`;
  try {
    const texto = textoDoOfx(await arquivo.arrayBuffer());
    estado.textoOfx = texto;
    // a data da comparação é a do relatório do SCK, quando ele já veio
    const dataAlvo = estado.sck?.data?.iso || null;
    const lido = lerExtratoOfx(texto, { dataAlvo });
    conferirConta(estado.conta, { ofx: lido });       // extrato de outra conta é recusado
    estado.ofx = lido;
    estado.nomeOfx = nome;
    const quando = lido.periodo?.fim?.rotulo || lido.saldoDoDia?.data?.rotulo || "sem período";
    mostrarResumoArquivo("ofx", nome, `${quando} · ${lido.lancamentos.length} lançamento(s)`);
    msg.innerHTML = "";
  } catch (e) {
    msg.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui ler o extrato", MSG_FALHA_GERAL)));
  } finally {
    solta.classList.remove("lendo");
  }
  await depoisDeImportar();
}

/**
 * Relê o extrato quando o relatório do SCK chega depois dele: só então se sabe
 * qual é a data da conciliação, e o extrato pode ter lançamento de outro dia.
 */
function realinharExtrato() {
  if (!estado.sck || !estado.ofx || !estado.textoOfx) return;
  try { estado.ofx = lerExtratoOfx(estado.textoOfx, { dataAlvo: estado.sck.data.iso }); }
  catch { /* mantém a leitura anterior */ }
}

async function depoisDeImportar() {
  desenharQuadro();
  if (!estado.sck || !estado.ofx) return;
  try {
    await salvar();
    estado.naoSalvo = false;
  } catch (e) {
    estado.naoSalvo = true;
    $("upload-msg").innerHTML = alerta("erro",
      esc(msgErro(e, "A conciliação foi feita, mas não consegui salvá-la no servidor")));
  }
  desenharQuadro();
}

/* ------------------------------------------------------------ quadro (modelo A) */
const TEXTO_SITUACAO = { conferido: "Confere", divergente: "Divergente", indefinido: "Não identificado" };

function desenharQuadro() {
  const cartao = $("cartao-quadro");
  if (!estado.sck && !estado.ofx) { cartao.classList.add("oculto"); return; }
  cartao.classList.remove("oculto");

  const q = montarConciliacao({ sck: estado.sck, ofx: estado.ofx });
  const data = q.data?.rotulo || (q.data?.iso ? rotuloIso(q.data.iso) : "");
  $("quadro-titulo").textContent = data ? `Conciliação de ${data}` : "Conciliação";
  $("quadro-conta").textContent = estado.conta ? `${estado.conta.empresa} · ${rotuloConta(estado.conta)}` : "";

  const pilula = $("quadro-situacao");
  pilula.textContent = q.situacao === "conferido" ? "Tudo confere"
    : q.situacao === "divergente" ? "Com divergência" : "Falta conferir";
  pilula.className = `pilula selo ${q.situacao === "conferido" ? "confere" : q.situacao}`;

  const notas = {
    saldo: "SCK: Saldo Contábil. Banco: saldo do dia no extrato.",
    entradas: "Créditos do relatório (coluna Débito) × entradas do extrato.",
    saidas: "Saídas do relatório (coluna Crédito) × débitos do extrato.",
  };
  $("quadro-cartoes").innerHTML = q.linhas.map((l) => {
    const classeDif = l.situacao === "indefinido" ? "indefinido"
      : l.situacao === "conferido" ? "confere" : (l.diferenca < 0 ? "negativa" : "divergente");
    // com impedimento o lado do banco não vale: mostrar 0,00 faria parecer que bate
    const valor = (v, vale = true) => (v == null || !vale ? "—" : moeda(v));
    const valeBanco = !l.impedimento;
    return `<section class="kpi-cartao ${l.situacao}">
      <div class="kpi-cab"><h3>${esc(l.rotulo)}</h3>
        <span class="selo ${l.situacao === "conferido" ? "confere" : l.situacao}">${TEXTO_SITUACAO[l.situacao]}</span></div>
      <dl class="kpi-corpo">
        <div class="kpi-linha"><dt>SCK</dt><dd>${valor(l.sck)}</dd></div>
        <div class="kpi-linha"><dt>Banco</dt><dd>${valor(l.banco, valeBanco)}</dd></div>
        <div class="kpi-risco"></div>
        <div class="kpi-linha dif ${classeDif}"><dt>Diferença</dt>
          <dd>${l.diferenca == null || l.impedimento ? "—" : comSinal(l.diferenca)}</dd></div>
        <p class="kpi-nota">${esc(l.impedimento || notas[l.chave] || "")}</p>
      </dl>
    </section>`;
  }).join("");

  /* a Diferença Geral do próprio SCK fica à parte, sem se misturar à comparação */
  const partes = [];
  if (q.sckSozinho?.diferencaGeral != null) {
    partes.push(`O SCK aponta <strong>Diferença Geral da Conciliação de ${esc(moeda(q.sckSozinho.diferencaGeral))}</strong>` +
      (q.sckSozinho.saldoExtratoNoSck === 0 ? ", com o Saldo Extrato/Conta ainda zerado" : "") +
      (q.sckSozinho.status ? ` (status <strong>${esc(q.sckSozinho.status)}</strong>)` : "") +
      ". Esse número é do próprio SCK e não entra na comparação acima.");
  }
  if (estado.naoSalvo) partes.push("Esta conciliação <strong>ainda não foi salva no servidor</strong>.");
  if (q.pendencias.length) partes.push(esc(q.pendencias.join(" ")));
  $("quadro-nota").innerHTML = partes.join("<br>");

  $("quadro-aviso").innerHTML = q.alertas.map((a) => alerta("atencao", esc(a))).join("");
}

/* ------------------------------------------------------------ tela */
function ligarCampo(qual, aoEscolher) {
  const solta = $(`solta-${qual}`), entrada = $(`arq-${qual}`);
  solta.addEventListener("click", () => entrada.click());
  solta.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); entrada.click(); }
  });
  ["dragenter", "dragover"].forEach((ev) => solta.addEventListener(ev, (e) => {
    e.preventDefault(); solta.classList.add("ativa");
  }));
  ["dragleave", "drop"].forEach((ev) => solta.addEventListener(ev, (e) => {
    e.preventDefault(); solta.classList.remove("ativa");
  }));
  solta.addEventListener("drop", (e) => {
    const f = e.dataTransfer?.files?.[0];
    if (f) aoEscolher(f);
  });
  entrada.addEventListener("change", () => {
    const f = entrada.files?.[0];
    entrada.value = "";
    if (f) aoEscolher(f);
  });
}

function ligarTela() {
  montarEscolha();
  ligarCampo("sck", importarSck);
  ligarCampo("ofx", importarOfx);

  document.addEventListener("click", (e) => {
    const x = e.target.closest(".alerta__x");
    if (x) { x.closest(".alerta")?.remove(); return; }
    const trocar = e.target.closest("[data-trocar]");
    if (trocar) $(`arq-${trocar.dataset.trocar}`).click();
  });
}

/* ------------------------------------------------------------ início */
(async function iniciar() {
  if (!ehPrincipalNoNavegador()) return;            // os demais continuam vendo "em desenvolvimento"
  const confirmado = await principalConfirmado();
  if (confirmado !== "sim") return;                 // "falha" também não abre: o servidor é quem decide
  $("em-desenvolvimento").classList.add("oculto");
  $("modulo").classList.remove("oculto");
  ligarTela();
})();
