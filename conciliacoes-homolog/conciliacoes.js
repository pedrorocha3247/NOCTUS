/**
 * Conciliações Bancárias — NOCTUS. Parte 1 (02/11/2026): escolha da empresa e da
 * conta, importação do relatório de conciliação do SCK e do extrato .OFX, e
 * comparação de saldo, entradas e saídas.
 *
 * Acesso (03/11/2026): enquanto o módulo está em construção, só quem tem perfil de
 * ADMINISTRADOR vê o conteúdo; o perfil "user" continua vendo "Módulo em
 * desenvolvimento". A trava da tela é só aparência: quem garante é o banco (RLS da
 * tabela conciliacao_bancaria_noctus: administrador, segundo noctus_eh_admin(), e só
 * as próprias linhas de cada um).
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
import { comporDiferenca, composicaoParaGravar } from "./composicao.js";
import { linhasImportacao, planilhaImportacao, nomeArquivoImportacao } from "./importacao.js";
import { empresasHabilitadas, contaPorId } from "./contas-sck.js";

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";
const TABELA = "conciliacao_bancaria_noctus";
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
/** O papel guardado no navegador é de administrador? (só decide o que tentar; quem decide é o servidor) */
function ehAdminNoNavegador() {
  const u = sessaoGravada()?.user;
  return !!u && Array.isArray(u.roles) && u.roles.includes("admin");
}
/** Pergunta ao servidor se o dono do token é administrador (mesma regra do resto do NOCTUS: noctus_eh_admin()). "sim", "nao" ou "falha" (sem comunicação). */
async function adminConfirmado() {
  const token = await tokenServidor();
  if (!token) return "nao";
  try {
    const r = await fetch(`${SB_URL}/rest/v1/rpc/noctus_eh_admin`, {
      method: "POST",
      headers: { apikey: SB_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: "{}",
    });
    if (r.status === 401 || r.status === 403) return "nao";
    if (!r.ok) return "falha";
    return (await r.json()) === true ? "sim" : "nao";
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
  composicaoAberta: false,   // recolhida por padrão (pedido do Rocha, 02/11/2026)
  importacaoAnterior: null,  // registro da última planilha gerada para esta conta + data (servidor)
  confirmarDeNovo: false,    // o aviso de "já foi gerada" está aberto, esperando "Gerar mesmo assim"
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
                 lancamentos: sck.lancamentos.length, adicionados: sck.adicionados?.total ?? 0 } : {},
    extrato: ofx ? { entradas: ofx.entradas, saidas: ofx.saidas,
                     saldoDoDia: ofx.saldoDoDia, saldoAnterior: ofx.saldoAnterior, saldoLivro: ofx.saldoLivro,
                     lancamentos: ofx.lancamentos.length } : {},
    comparacao: { situacao: quadro.situacao, linhas: quadro.linhas.map(({ chave, sck: s, banco, diferenca, situacao }) =>
                    ({ chave, sck: s, banco, diferenca, situacao })),
                  // estruturada para a Parte 2 (sem nome nem CPF/CNPJ: o FITID liga ao extrato)
                  composicao: composicaoParaGravar(comporDiferenca({ sck, ofx })),
                  // registro da planilha de importação gerada (evita importar duas vezes no SCK)
                  ...(estado.importacaoAnterior ? { importacao: estado.importacaoAnterior } : {}) },
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
/* Modelo B (03/11/2026): lista de empresas e, ao lado, as contas da empresa marcada. Os dois <select> escondidos
   continuam sendo quem guarda a escolha; a lista só desenha e repassa o clique para eles. */
const semAcento = (t) => String(t ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const iniciais = (nome) => {
  const p = String(nome ?? "").trim().split(/\s+/).filter(Boolean);
  return (p.length > 1 ? p[0][0] + p[1][0] : (p[0] || "").slice(0, 2)).toUpperCase();
};
/** "Itaú - 341" → {nome: "Itaú", cod: "341"}; "Daycoval" → {nome: "Daycoval", cod: ""}. */
const partesBanco = (banco) => {
  const m = /^(.*?)\s*-\s*(\d+)\s*$/.exec(String(banco ?? "").trim());
  return m && m[1] ? { nome: m[1], cod: m[2] } : { nome: String(banco ?? "").trim(), cod: "" };
};
const ICONE_BANCO = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 9.5 12 4l9 5.5M5.5 10.5v7M10 10.5v7M14 10.5v7M18.5 10.5v7M3 20h18"/></svg>';

function montarEscolha() {
  const empresas = empresasHabilitadas();
  const selE = $("sel-empresa"), selC = $("sel-conta");
  const listaE = $("lista-empresas"), listaC = $("lista-contas"), busca = $("busca-empresa");
  selE.innerHTML = empresas.map((e) => `<option value="${esc(e.codigo)}">${esc(e.empresa)}</option>`).join("");

  // a busca só aparece quando a lista já é longa
  busca.closest(".busca-empresa").hidden = empresas.length <= 5;

  const desenharEmpresas = () => {
    const q = semAcento(busca.value).trim();
    const visiveis = empresas.filter((e) => !q || semAcento(e.empresa).includes(q));
    listaE.innerHTML = visiveis.length ? visiveis.map((e) =>
      `<li><label class="opcao empresa-item"><input type="radio" name="opt-empresa" value="${esc(e.codigo)}"` +
      `${e.codigo === selE.value ? " checked" : ""}><span class="opcao__corpo">` +
      `<span class="avatar" aria-hidden="true">${esc(iniciais(e.empresa))}</span>` +
      `<span class="opcao__nome">${esc(e.empresa)}</span>` +
      `<span class="opcao__nota">${e.contas.length} ${e.contas.length === 1 ? "conta" : "contas"}</span></span></label></li>`
    ).join("") : `<li class="sem-resultado">Nenhuma empresa encontrada.</li>`;
  };
  const desenharContas = (emp) => {
    $("contas-titulo").textContent = emp ? `Contas de ${emp.empresa}` : "";
    listaC.innerHTML = (emp?.contas || []).map((c) => {
      const b = partesBanco(c.banco);
      return `<label class="opcao conta-item"><input type="radio" name="opt-conta" value="${esc(c.id)}"` +
        `${c.id === selC.value ? " checked" : ""}><span class="opcao__corpo">` +
        `<span class="radio-pino" aria-hidden="true"></span>` +
        `<span class="banco-tile" aria-hidden="true">${ICONE_BANCO}</span>` +
        `<span class="conta-texto"><strong>${esc(b.nome)}${b.cod ? ` <span>· ${esc(b.cod)}</span>` : ""}</strong>` +
        `<span>Ag. ${esc(c.agencia)} · C/C ${esc(c.conta)}</span></span>` +
        `<span class="pilula-sck">SCK ${esc(c.codigoSck)}</span></span></label>`;
    }).join("");
    $("contas-dica").textContent = (emp?.contas || []).length === 1
      ? "Quando a empresa tiver mais de uma conta habilitada, cada uma aparece aqui como uma opção." : "";
  };
  const trocarEmpresa = () => {
    const emp = empresas.find((e) => e.codigo === selE.value) || empresas[0];
    selC.innerHTML = (emp?.contas || []).map((c) =>
      `<option value="${esc(c.id)}">${esc(rotuloConta(c))}</option>`).join("");
    desenharEmpresas();
    desenharContas(emp);
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
    $("cartao-composicao").classList.add("oculto");
    estado.composicaoAberta = false;
    estado.importacaoAnterior = null;
    estado.confirmarDeNovo = false;
    $("cartao-importacao").classList.add("oculto");
    $("importacao-msg").innerHTML = "";
    $("conta-detalhe").textContent = estado.conta ? `CNPJ ${estado.conta.cnpj}` : "";
  };
  listaE.addEventListener("change", (e) => {
    if (e.target.name !== "opt-empresa") return;
    selE.value = e.target.value;
    trocarEmpresa();
    listaE.querySelector("input:checked")?.focus();     // o desenho refeito não pode tirar o foco do teclado
  });
  listaC.addEventListener("change", (e) => {
    if (e.target.name !== "opt-conta") return;
    selC.value = e.target.value;
    trocarConta();
  });
  busca.addEventListener("input", desenharEmpresas);
  selE.addEventListener("change", trocarEmpresa);
  selC.addEventListener("change", () => { desenharContas(empresas.find((e) => e.codigo === selE.value)); trocarConta(); });
  trocarEmpresa();

  if (!empresas.length) {
    $("escolha-sub").textContent = "Nenhuma conta está habilitada para conciliação ainda.";
    $("escolha-corpo").hidden = true;
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
const ICONE_OK = '<svg class="ok" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" ' +
  'stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 6 9 17l-5-5"/></svg>';

function mostrarResumoArquivo(qual, nome, detalhe) {
  $(`resumo-${qual}`).innerHTML = `<span class="ok-circulo">${ICONE_OK}</span>` +
    `<span class="corpo"><span class="nome">${esc(nome)}</span><span class="detalhe">${esc(detalhe)}</span></span>` +
    `<button type="button" class="trocar" data-trocar="${qual}">Trocar arquivo</button>`;
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
  // a gravação substitui a coluna "comparacao" inteira: traz antes o registro de
  // planilha já gerada para esta conta e data, para ele não se perder
  try {
    const salva = await carregarSalva(estado.conta, estado.sck.data.iso);
    estado.importacaoAnterior = salva?.comparacao?.importacao || null;
  } catch { /* sem o registro: segue; no pior caso o aviso de geração repetida não aparece */ }
  estado.confirmarDeNovo = false;
  $("importacao-msg").innerHTML = "";   // arquivo novo: a mensagem da geração anterior não vale mais
  desenharImportacao();
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
/*
 * Avisos dos leitores que NÃO aparecem na tela (pedido do Rocha, 02/11/2026 20:54),
 * porque se repetem todo dia com o extrato do Itaú e não pedem ação:
 *   - extrato sem sinal: o cartão Saídas já mostra "Não identificado" e explica;
 *   - LEDGERBAL de outra data: informação técnica; o saldo usado é o do dia.
 * Continuam nos leitores (o "sem sinal" é o que faz o cartão Saídas não dizer
 * "confere"), só não viram mensagem.
 */
const SO_INFORMATIVO = [/n[ãa]o distingue entrada de sa[íi]da/i, /LEDGERBAL/i];
const TEXTO_SITUACAO = { conferido: "Confere", divergente: "Divergente", indefinido: "Não identificado" };

function desenharQuadro() {
  const cartao = $("cartao-quadro");
  desenharComposicao();
  desenharImportacao();
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

  /* A faixa da Diferença Geral do SCK e a linha dos Lançamentos Adicionados saíram
     (pedido do Rocha, 02/11/2026 20:54): os dois números já aparecem, com explicação,
     na composição da diferença (saldo anterior × Diferença Geral; adicionados em "Validar"). */
  const partes = [];
  if (estado.naoSalvo) partes.push("Esta conciliação <strong>ainda não foi salva no servidor</strong>.");
  if (q.pendencias.length) partes.push(esc(q.pendencias.join(" ")));
  $("quadro-nota").innerHTML = partes.join("<br>");

  $("quadro-aviso").innerHTML = q.alertas.filter((a) => !SO_INFORMATIVO.some((r) => r.test(a)))
    .map((a) => alerta("atencao", esc(a))).join("");
}

/* ------------------------------------------------------------ composição (regras 1.1 a 1.5) */
const TEXTO_REGRA = { conferido: "Conciliado", divergente: "Divergente", pendente: "A lançar", indefinido: "Não identificado" };
const GRUPOS = [
  ["lancar", "Lançar no SCK", "acao-lancar"],
  ["adicionar", "Adicionar na conciliação", "acao-adicionar"],
  ["validar", "Validar (não lançar de novo)", "acao-validar"],
  ["divergencia", "Divergências", "acao-divergencia"],
];
const valorOuTraco = (v) => (v == null ? "—" : moeda(v));
const dataCurta = (d) => d?.rotulo || (d?.iso ? rotuloIso(d.iso) : "");

function desenharComposicao() {
  const cartao = $("cartao-composicao");
  const c = comporDiferenca({ sck: estado.sck, ofx: estado.ofx });
  if (!c) { cartao.classList.add("oculto"); return; }
  cartao.classList.remove("oculto");
  aplicarRecolhimento();

  $("composicao-sub").textContent = c.diferencaSaldo == null ? ""
    : `Diferença do saldo (Banco − SCK): ${comSinal(c.diferencaSaldo)}`;
  const pilula = $("composicao-situacao");
  pilula.textContent = c.semOrigem == null ? "Falta o saldo" : c.fecha ? "Composição fecha" : "Diferença sem origem";
  pilula.className = `pilula selo ${c.semOrigem == null ? "indefinido" : c.fecha ? "confere" : "divergente"}`;

  /* conferência por tipo */
  $("tabela-regras").innerHTML = `<thead><tr><th>Tipo</th><th class="num">Banco</th><th class="num">SCK</th>` +
    `<th class="num">Diferença</th><th>Situação</th></tr></thead><tbody>` +
    (c.regras.length ? c.regras.map((r) => `<tr>
      <td>${esc(r.rotulo)}<span class="detalhe">${esc(r.detalhe || "")}</span></td>
      <td class="num" data-rotulo="Banco">${valorOuTraco(r.banco)}</td>
      <td class="num" data-rotulo="SCK">${valorOuTraco(r.sck)}</td>
      <td class="num" data-rotulo="Diferença">${r.diferenca == null ? "—" : comSinal(r.diferenca)}</td>
      <td><span class="selo ${r.situacao === "conferido" ? "confere" : r.situacao}">${TEXTO_REGRA[r.situacao]}</span></td>
    </tr>`).join("") : `<tr><td colspan="5">Nenhum lançamento de cobrança, PIX QRS ou PIX recebido no dia.</td></tr>`) +
    `</tbody>`;

  /* de onde vem a diferença */
  let corpo = "";
  for (const [grupo, titulo, classe] of GRUPOS) {
    const itens = c.componentes.filter((x) => x.grupo === grupo);
    const temSaldoAnterior = grupo === "validar" && c.saldoAnterior;
    if (!itens.length && !temSaldoAnterior) continue;
    corpo += `<tr class="grupo"><td colspan="3"><span class="${classe}">${esc(titulo)}</span></td></tr>`;
    for (const x of itens) {
      const detalhe = x.tipo === "pixRecebido"
        ? [dataCurta(x.data), x.nome, x.documento].filter(Boolean).join(" · ") + (x.historico ? ` — ${x.historico}` : "")
        : x.linhas ? `${x.linhas.length} lançamento(s) na seção Lançamentos Adicionados do SCK`
        : [dataCurta(x.data), x.historico].filter(Boolean).join(" · ");
      corpo += `<tr><td>${esc(x.descricao)}<span class="detalhe">${esc(detalhe)}</span></td>` +
        `<td class="num">${comSinal(x.valor)}</td><td>${esc(x.tratamento)}</td></tr>`;
      for (const l of x.linhas || []) {
        corpo += `<tr class="sub"><td>${esc([dataCurta(l.data), l.historico].filter(Boolean).join(" · "))}</td>` +
          `<td class="num">${comSinal(l.direcao === "saida" ? -l.valor : l.valor)}</td><td></td></tr>`;
      }
    }
    if (temSaldoAnterior) {
      const s = c.saldoAnterior;
      const resultado = s.situacao === "conferido" ? `confere com a Diferença Geral do SCK (${moeda(s.sck)})`
        : s.situacao === "divergente" ? `não bate com a Diferença Geral do SCK (${moeda(s.sck)}): diferença de ${comSinal(s.diferenca)}`
        : s.motivo;
      corpo += `<tr><td>Saldo anterior do dia (extrato)<span class="detalhe">Conferência, não entra na soma</span></td>` +
        `<td class="num">${moeda(s.banco)}</td>` +
        `<td><span class="selo ${s.situacao === "conferido" ? "confere" : s.situacao}">${TEXTO_SITUACAO[s.situacao]}</span> ` +
        `${esc(resultado)}</td></tr>`;
    }
  }
  if (c.semOrigem != null && !c.fecha) {
    const dica = !estado.sck?.adicionados?.linhas?.length
      ? "Investigar. O relatório não tem Lançamentos Adicionados: confira se faltam rendimentos na conciliação do SCK."
      : "Investigar: nenhuma regra explica este valor.";
    corpo += `<tr class="grupo"><td colspan="3"><span class="acao-divergencia">Sem origem identificada</span></td></tr>` +
      `<tr><td>Diferença sem origem identificada</td><td class="num">${comSinal(c.semOrigem)}</td><td>${esc(dica)}</td></tr>`;
  }
  if (!corpo) corpo = `<tr><td colspan="3">Nada a lançar, adicionar ou validar.</td></tr>`;

  $("tabela-composicao").innerHTML =
    `<thead><tr><th>Descrição</th><th class="num">Valor</th><th>Tratamento</th></tr></thead>` +
    `<tbody>${corpo}</tbody><tfoot>` +
    `<tr><td>Total explicado</td><td class="num">${comSinal(c.explicado)}</td><td></td></tr>` +
    (c.diferencaSaldo == null ? "" :
      `<tr class="total-dif"><td>Diferença do saldo (Banco − SCK)</td><td class="num">${comSinal(c.diferencaSaldo)}</td>` +
      `<td>${c.fecha ? "A composição explica toda a diferença." : "A composição não explica toda a diferença."}</td></tr>`) +
    `</tfoot>`;
}

/** Abre ou fecha o corpo da composição; o topo (título, diferença e situação) fica sempre à vista. */
function aplicarRecolhimento() {
  const aberta = estado.composicaoAberta;
  $("composicao-corpo").hidden = !aberta;
  const botao = $("composicao-alternar");
  botao.setAttribute("aria-expanded", String(aberta));
  botao.textContent = aberta ? "Ocultar composição" : "Mostrar composição";
}

/* ------------------------------------------------------------ Parte 2: planilha de importação */
const quandoBr = (iso) => {
  try {
    return new Date(iso).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
  } catch { return ""; }
};

function desenharImportacao() {
  const cartao = $("cartao-importacao");
  if (!estado.sck || !estado.ofx) { cartao.classList.add("oculto"); return; }
  cartao.classList.remove("oculto");
  let r;
  try { r = linhasImportacao({ sck: estado.sck, ofx: estado.ofx }); }
  catch (e) { $("importacao-sub").textContent = ""; $("btn-importacao").disabled = true;
              $("importacao-msg").innerHTML = alerta("erro", esc(msgErro(e, "Não consegui preparar os lançamentos", MSG_FALHA_GERAL))); return; }
  const qtd = r.linhas.length;
  $("importacao-sub").textContent = qtd
    ? `${qtd} PIX recebido${qtd > 1 ? "s" : ""} · R$ ${moeda(r.total)} · conta ${r.linhas[0].conta} · ${estado.sck.data.rotulo}`
    : `Nenhum PIX recebido no extrato de ${estado.sck.data.rotulo}`;
  $("btn-importacao").disabled = !qtd;
}

/** Monta e baixa a planilha; com geração anterior registrada, pede confirmação antes. */
async function gerarImportacao(forcar = false) {
  const msg = $("importacao-msg");
  let r;
  try { r = linhasImportacao({ sck: estado.sck, ofx: estado.ofx }); }
  catch (e) { msg.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui gerar a planilha", MSG_FALHA_GERAL))); return; }

  const ant = estado.importacaoAnterior;
  if (ant && !forcar) {
    estado.confirmarDeNovo = true;
    msg.innerHTML = alerta("atencao",
      `A planilha desta conta e data <strong>já foi gerada em ${esc(quandoBr(ant.geradaEm))}</strong> ` +
      `(${esc(String(ant.quantidade))} PIX, R$ ${esc(moeda(ant.total))}). Se ela já foi importada no SCK, ` +
      `importar de novo <strong>duplica os lançamentos</strong>.`) +
      `<div class="acoes-confirmar"><button type="button" class="botao-principal" id="btn-importacao-confirmar">Gerar mesmo assim</button>` +
      `<button type="button" class="botao-comp" id="btn-importacao-cancelar">Cancelar</button></div>`;
    return;
  }
  estado.confirmarDeNovo = false;

  try {
    const bytes = planilhaImportacao(r.linhas);
    const nome = nomeArquivoImportacao(r.linhas[0].conta, estado.sck.data.iso);
    const url = URL.createObjectURL(new Blob([bytes],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);

    const avisos = r.alertas.map((t) => alerta("atencao", esc(t))).join("");
    msg.innerHTML = alerta("ok", `Planilha <strong>${esc(nome)}</strong> gerada com ${r.linhas.length} PIX ` +
      `(R$ ${esc(moeda(r.total))}). Importe-a no SCK.`) + avisos;

    // registra a geração (só quantidade, total, FITID e hora: nada de nome ou CPF/CNPJ)
    estado.importacaoAnterior = { geradaEm: new Date().toISOString(), quantidade: r.linhas.length, total: r.total,
                                  fitids: r.linhas.map((l) => l.fitid), arquivo: nome };
    try { await salvar(); estado.naoSalvo = false; }
    catch (e) {
      msg.innerHTML += alerta("erro", esc(msgErro(e,
        "A planilha foi gerada, mas não consegui registrar a geração no servidor")));
    }
  } catch (e) {
    msg.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui gerar a planilha", MSG_FALHA_GERAL)));
  }
}

/* ------------------------------------------------------------ tela */
function ligarCampo(qual, aoEscolher) {
  const solta = $(`solta-${qual}`), entrada = $(`arq-${qual}`);
  solta.addEventListener("click", () => entrada.click());
  solta.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); entrada.click(); }
  });
  const alvo = solta.closest(".campo-arquivo") || solta;
  ["dragenter", "dragover"].forEach((ev) => alvo.addEventListener(ev, (e) => {
    e.preventDefault(); alvo.classList.add("ativa");
  }));
  ["dragleave", "drop"].forEach((ev) => alvo.addEventListener(ev, (e) => {
    e.preventDefault();
    if (ev === "dragleave" && e.relatedTarget && alvo.contains(e.relatedTarget)) return;   // passou para um filho: segue ativa
    alvo.classList.remove("ativa");
  }));
  alvo.addEventListener("drop", (e) => {
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
  $("btn-importacao").addEventListener("click", () => gerarImportacao(false));
  $("importacao-msg").addEventListener("click", (e) => {
    if (e.target.closest("#btn-importacao-confirmar")) gerarImportacao(true);
    else if (e.target.closest("#btn-importacao-cancelar")) { estado.confirmarDeNovo = false; $("importacao-msg").innerHTML = ""; }
  });
  $("composicao-alternar").addEventListener("click", () => {
    estado.composicaoAberta = !estado.composicaoAberta;
    aplicarRecolhimento();
  });

  document.addEventListener("click", (e) => {
    const x = e.target.closest(".alerta__x");
    if (x) { x.closest(".alerta")?.remove(); return; }
    const trocar = e.target.closest("[data-trocar]");
    if (trocar) $(`arq-${trocar.dataset.trocar}`).click();
  });
}

/* ------------------------------------------------------------ início */
(async function iniciar() {
  if (!ehAdminNoNavegador()) return;                // perfil "user": continua vendo "em desenvolvimento"
  const confirmado = await adminConfirmado();
  if (confirmado !== "sim") { liberarAviso(); return; }   // "falha" também não abre: o servidor é quem decide
  $("em-desenvolvimento").classList.add("oculto");
  $("modulo").classList.remove("oculto");
  liberarAviso();
  ligarTela();
})().catch(() => liberarAviso());

/** Tira a marca "decidindo" do HTML: até aqui o quadro "em desenvolvimento" ficava escondido para o administrador (03/11/2026). */
function liberarAviso() { document.documentElement.classList.remove("decidindo"); }
