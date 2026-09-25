/**
 * Módulo de Conferência de Pagamentos — NOCTUS
 *
 * Roda inteiro no navegador: o PDF é lido localmente, o progresso fica no
 * localStorage e a planilha é gerada no cliente. Nenhum dado sai da máquina.
 */
import * as pdfjsLib from "./vendor/pdf.min.mjs";
import { parseRelatorio, parseRelatorioExcel } from "./parser.js";
import { verificarPoder } from "./poderes.js";
import { rotuloEmpresa } from "./contas.js";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  new URL("./vendor/pdf.worker.min.mjs", import.meta.url).href;

const ORDEM = ["TRANSFERÊNCIA", "TED", "DÉBITO EM CONTA", "PIX", "BOLETO"];
const STATUS = [
  { valor: "Aprovado",                   id: "aprovado"   },
  { valor: "Aguardando esclarecimentos", id: "aguardando" },
  { valor: "Recusado",                   id: "recusado"   },
];
/**
 * Nomes usados em versões anteriores, para não perder os pareceres já dados.
 * "Aprovado com ressalva" virou "Aprovado" — a ressalva em si segue escrita
 * no texto do parecer, que é onde ela sempre esteve.
 */
const STATUS_ANTIGOS = {
  Conforme: "Aprovado",
  Ressalva: "Aprovado",
  "Aprovado com ressalva": "Aprovado",
  Retido: "Aguardando esclarecimentos",
  "Em dúvida": "Aguardando esclarecimentos",
  Devolvido: "Recusado",
};
const idDoStatus = (v) => (STATUS.find((s) => s.valor === v) || {}).id || "";
const CHAVE = "noctus.conferencia";
const CHAVE_ANTIGA = "sipep.conferencia";
const CHAVE_OTN = "noctus.otn";
const OTN_PADRAO = 130.30;

/** Backend dos anexos (Supabase) — mesma sessão de login do NOCTUS (localStorage). */
const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";

const $ = (id) => document.getElementById(id);
const moeda = (v) =>
  (v ?? 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/** Valor do OTN usado pra converter o limite de poder (em OTN) para reais. */
const otnAtual = () => {
  const v = parseFloat(localStorage.getItem(CHAVE_OTN));
  return isFinite(v) && v > 0 ? v : OTN_PADRAO;
};

let estado = { dados: null, itens: [], pareceres: {}, removidas: [], i: 0, mesclagem: null };

/* ---------------------------------------------------------------- persistência */
/**
 * A chave é a data do relatório mais a empresa. O SCK vai acrescentando
 * solicitações ao longo do dia, então o relatório da tarde é o MESMO lote da
 * manhã, com mais linhas — e não uma conferência nova. Mas cada empresa tem o
 * seu relatório, então o mesmo dia pode ter vários lotes, um por empresa.
 *
 * O relatório em EXCEL foge dessa regra na origem: sai com todas as empresas
 * juntas num arquivo só (ver parseRelatorioExcel, em parser.js). Até
 * 22/09/2026 isso virava UMA conferência mista, com uma chave fixa "todas" —
 * ruim pra emitir parecer por empresa. Desde 23/09/2026 (dividirPorEmpresa/
 * processarMultiEmpresa, em conferencia.js) cada empresa do Excel vira a sua
 * PRÓPRIA conferência, com a MESMA chave por código que o PDF sempre usou —
 * por isso este `if` abaixo não deveria mais disparar em uso normal; fica só
 * como salvaguarda (por exemplo se algum dia alguém chamar chaveLote/
 * codEmpresa direto com o `dados` cru do parser, sem passar por
 * dividirPorEmpresa primeiro). Lotes ".etodas" salvos ANTES dessa mudança são
 * migrados automaticamente (ver o fim de migrarChaves).
 */
const codEmpresa = (meta) => {
  if (meta?.multiEmpresa) return "todas";
  const m = String(meta?.empresaCodigo || meta?.empresa || "").match(/^\s*(\d{1,4})/);
  if (m) return m[1];
  const nome = String(meta?.empresaNome || meta?.empresa || "").trim();
  return nome ? nome.replace(/[^A-Za-z0-9]+/g, "-").slice(0, 24) : "sem-empresa";
};
const nomeEmpresa = (meta) =>
  meta?.empresa || meta?.empresaNome || "Empresa não identificada";

/**
 * Texto do período do relatório: "01/09/2026" para um relatório de um dia só,
 * "01/09/2026 a 18/09/2026" quando o relatório cobre vários dias (dataFim
 * diferente de dataInicio). Antes disso só dataInicio era mostrado em toda
 * parte — para um relatório de um dia só isso já era o período inteiro, mas
 * num relatório de várias semanas escondia os outros dias do usuário.
 */
const textoPeriodo = (meta) =>
  meta?.dataFim && meta.dataFim !== meta.dataInicio
    ? `${meta.dataInicio || ""} a ${meta.dataFim}`
    : (meta?.dataInicio || "");

/**
 * Chave "base" da conferência (sem usuário): identifica a conferência pelo
 * documento em si (data + empresa). Usada pelos anexos, que pertencem à
 * solicitação, não a quem está conferindo — não pode depender de login.
 */
const chaveLoteBase = (d) =>
  `${CHAVE}.${(d.meta.dataInicio || "sem-data").replace(/\//g, "-")}.e${codEmpresa(d.meta)}`;

/** Chave da conferência COM escopo de usuário — usada só para o progresso
 *  (pareceres) salvo neste navegador, que é isolado por pessoa. */
const chaveLote = (d) =>
  `${CHAVE}.${escopoUsuario()}.${(d.meta.dataInicio || "sem-data").replace(/\//g, "-")}.e${codEmpresa(d.meta)}`;

/** Campos cuja alteração invalida um parecer já dado. */
const assinatura = (s) => [s.tipo, s.valor, s.favorecido, s.cpfCnpj, s.destinacao,
                           s.poder, s.solicitante, s.competente].join("|");

/**
 * Junta o relatório recém-aberto com a conferência já gravada para aquela data.
 * O relatório novo é a verdade PARA A LISTA ATIVA: ele manda em estado.itens e
 * na validação de totais, que só pode bater com o que está impresso hoje. Os
 * pareceres já dados são preservados, menos os de solicitações que mudaram —
 * esses voltam a pendente, porque conferir R$ 10.000 não vale como parecer
 * para R$ 15.000.
 *
 * Pedido do Rocha em 25/09/2026: uma solicitação que JÁ tinha parecer (ex.:
 * Recusado, com o motivo escrito) e que some de um relatório mais novo não
 * pode perder essa informação — ele precisa continuar sabendo o motivo mesmo
 * depois que ela sumiu do SCK (retirada, cancelada etc.). Até aqui ela
 * simplesmente saía de solicitacoes/pareceres e o motivo ficava só na memória
 * de quem conferiu. Agora, junto com `sumiram` (a lista de S.N., que já
 * existia só pra aviso/contagem), esta função também devolve `removidas`:
 * uma cópia de cada solicitação sumida — todos os campos originais (S.N.,
 * valor, favorecido, tipo etc.) MAIS o status e o texto do parecer que ela
 * tinha — guardada À PARTE de estado.itens/pareceres (fora da validação de
 * totais do relatório atual) e acumulada a cada novo relatório, até que a
 * solicitação volte a aparecer (ver `reapareceram`, que restaura o parecer
 * salvo) ou alguém a remova manualmente da lista de removidas.
 *
 * Uma solicitação que NUNCA teve parecer e simplesmente some não vira
 * "removida" — não há motivo nenhum a preservar ali, e listá-la seria ruído
 * (ela nem é uma decisão do conferente, só uma linha que nunca chegou a ser
 * olhada).
 */
function mesclar(anterior, novo) {
  const antes = new Map((anterior.solicitacoes || []).map((s) => [s.sn, s]));
  const pareceresAntigos = anterior.pareceres || {};
  // solicitações que já tinham sumido numa mesclagem anterior deste mesmo
  // lote (dia+empresa) e ficaram guardadas em "removidas" — indexadas por
  // S.N. pra checar se alguma delas voltou a constar no relatório de agora.
  const removidasAntigas = new Map((anterior.removidas || []).map((r) => [r.sn, r]));
  const pareceres = {};
  const novas = [], alteradas = [], reapareceram = [];

  for (const s of novo.solicitacoes) {
    const removidaAntes = removidasAntigas.get(s.sn);
    if (removidaAntes) {
      // reapareceu: estava sumida (já com parecer dado) e voltou a constar
      // no relatório — restaura o parecer que tinha, com a mesma regra de
      // "alterada" se os dados-chave mudaram desde que ela sumiu.
      removidasAntigas.delete(s.sn);
      reapareceram.push(s.sn);
      pareceres[s.sn] = (removidaAntes.status && assinatura(removidaAntes) !== assinatura(s))
        ? { status: "", parecer: removidaAntes.parecer }
        : { status: removidaAntes.status, parecer: removidaAntes.parecer };
      if (!pareceres[s.sn].status) alteradas.push(s.sn);
      continue;
    }
    const anteriorS = antes.get(s.sn);
    if (!anteriorS) { novas.push(s.sn); continue; }
    const p = pareceresAntigos[s.sn];
    if (!p) continue;
    if (p.status && assinatura(anteriorS) !== assinatura(s)) {
      // guarda o texto para ele reaproveitar, mas o status volta a pendente
      pareceres[s.sn] = { status: "", parecer: p.parecer };
      alteradas.push(s.sn);
    } else {
      pareceres[s.sn] = p;
    }
  }

  const agora = new Set(novo.solicitacoes.map((s) => s.sn));
  const sumidasAgora = (anterior.solicitacoes || [])
    .filter((s) => !agora.has(s.sn) && pareceresAntigos[s.sn]?.status);
  const removidasNovas = sumidasAgora.map((s) => ({
    ...s,
    status: pareceresAntigos[s.sn].status,
    parecer: pareceresAntigos[s.sn].parecer,
    removidoEm: novo.meta?.dataInicio || anterior.meta?.dataInicio || "",
  }));
  const removidas = [...removidasAntigas.values(), ...removidasNovas];
  const sumiram = sumidasAgora.map((s) => s.sn);

  return { pareceres, novas, alteradas, sumiram, removidas, reapareceram };
}

function salvar() {
  try {
    localStorage.setItem(chaveLote(estado.dados), JSON.stringify({
      meta: estado.dados.meta, validacao: estado.dados.validacao,
      solicitacoes: estado.dados.solicitacoes, pareceres: estado.pareceres,
      removidas: estado.removidas || [], i: estado.i,
    }));
  } catch (e) { /* modo privado, cota cheia: a conferência continua, só não persiste */ }
}

/**
 * As conferências guardadas antes de o sistema virar NOCTUS continuam valendo:
 * migra as chaves antigas na primeira abertura, sem perder nenhum parecer.
 */
function migrarChaves() {
  try {
    for (const antiga of Object.keys(localStorage)) {
      if (!antiga.startsWith(CHAVE_ANTIGA + ".")) continue;
      const nova = CHAVE + antiga.slice(CHAVE_ANTIGA.length);
      if (!localStorage.getItem(nova)) localStorage.setItem(nova, localStorage.getItem(antiga));
      localStorage.removeItem(antiga);
    }
    // chaves antigas traziam a quantidade no fim (…02-09-2026.94): agora a data basta
    for (const chave of Object.keys(localStorage)) {
      const m = chave.match(new RegExp(`^${CHAVE}\\.(\\d{2}-\\d{2}-\\d{4})\\.\\d+$`));
      if (!m) continue;
      const destino = `${CHAVE}.${m[1]}`;
      const atual = localStorage.getItem(destino);
      if (!atual) {
        localStorage.setItem(destino, localStorage.getItem(chave));
      } else {
        // duas gravações do mesmo dia: fica a de lista maior, com os pareceres somados
        const a = JSON.parse(atual), b = JSON.parse(localStorage.getItem(chave));
        const base = (b.solicitacoes || []).length > (a.solicitacoes || []).length ? b : a;
        base.pareceres = { ...(b.pareceres || {}), ...(a.pareceres || {}) };
        localStorage.setItem(destino, JSON.stringify(base));
      }
      localStorage.removeItem(chave);
    }
    // chaves só-data (…09-09-2026): agora cada empresa tem o seu lote
    for (const chave of Object.keys(localStorage)) {
      const m = chave.match(new RegExp(`^${CHAVE}\\.(\\d{2}-\\d{2}-\\d{4})$`));
      if (!m) continue;
      const v = JSON.parse(localStorage.getItem(chave));
      const destino = `${chave}.e${codEmpresa(v.meta)}`;
      if (!localStorage.getItem(destino)) localStorage.setItem(destino, localStorage.getItem(chave));
      localStorage.removeItem(chave);
    }
    for (const chave of Object.keys(localStorage)) {
      if (!chave.startsWith(CHAVE + ".")) continue;
      const v = JSON.parse(localStorage.getItem(chave));
      let mudou = false;
      for (const p of Object.values(v.pareceres || {})) {
        if (STATUS_ANTIGOS[p.status]) { p.status = STATUS_ANTIGOS[p.status]; mudou = true; }
      }
      if (mudou) localStorage.setItem(chave, JSON.stringify(v));
    }
    // conferências salvas antes do cadastro por usuário: passam a pertencer ao admin
    // (único usuário que existia antes) na primeira vez que ele abrir o módulo depois
    // desta atualização. Um usuário novo cadastrado nunca herda essas conferências.
    if (ehAdmin()) {
      for (const chave of Object.keys(localStorage)) {
        const m = chave.match(new RegExp(`^${CHAVE}\\.(\\d{2}-\\d{2}-\\d{4})\\.e(.+)$`));
        if (!m) continue;
        const destino = `${CHAVE}.${escopoUsuario()}.${m[1]}.e${m[2]}`;
        if (!localStorage.getItem(destino)) localStorage.setItem(destino, localStorage.getItem(chave));
        localStorage.removeItem(chave);
      }
    }
    // lotes ".etodas": Excel com várias empresas juntas numa ÚNICA
    // conferência, do jeito que o sistema salvava antes de 23/09/2026 —
    // separa em uma conferência por empresa (mesma regra de
    // dividirPorEmpresa/processarMultiEmpresa que passou a valer pra
    // upload novo dali pra frente), preservando o escopo do usuário dono do
    // lote (não necessariamente o usuário atual, se o navegador for
    // compartilhado) e os pareceres já dados, reagrupados por empresa.
    for (const chave of Object.keys(localStorage)) {
      const m = chave.match(new RegExp(`^${CHAVE}\\.(.+)\\.(\\d{2}-\\d{2}-\\d{4})\\.etodas$`));
      if (!m) continue;
      const [, escopo, data] = m;
      const v = JSON.parse(localStorage.getItem(chave));
      if (!v?.meta?.multiEmpresa || !Array.isArray(v.solicitacoes)) continue;
      const partes = dividirPorEmpresa({ meta: v.meta, solicitacoes: v.solicitacoes, validacao: v.validacao });
      for (const parte of partes) {
        calcularPoderes(parte.meta, parte.solicitacoes); // OTN pode ter mudado desde o save original
        const snsDaParte = new Set(parte.solicitacoes.map((s) => s.sn));
        const pareceresDaParte = {};
        for (const [sn, p] of Object.entries(v.pareceres || {}))
          if (snsDaParte.has(sn)) pareceresDaParte[sn] = p;
        const destino = `${CHAVE}.${escopo}.${data}.e${codEmpresa(parte.meta)}`;
        if (!localStorage.getItem(destino))
          localStorage.setItem(destino, JSON.stringify({
            meta: parte.meta, validacao: parte.validacao,
            solicitacoes: parte.solicitacoes, pareceres: pareceresDaParte, i: 0,
          }));
      }
      localStorage.removeItem(chave);
    }
  } catch (e) { /* modo privado ou entrada corrompida: segue sem migrar */ }
}

function lotesSalvos() {
  const out = [];
  const prefixo = `${CHAVE}.${escopoUsuario()}.`;
  for (let k = 0; k < localStorage.length; k++) {
    const chave = localStorage.key(k);
    if (!chave || !chave.startsWith(prefixo)) continue;
    try {
      const v = JSON.parse(localStorage.getItem(chave));
      out.push({ chave, meta: v.meta, data: v.meta?.dataInicio || "sem data",
                 empresa: nomeEmpresa(v.meta), total: v.solicitacoes.length,
                 feitos: Object.values(v.pareceres || {}).filter((p) => p.status).length });
    } catch (e) { /* entrada corrompida: ignora */ }
  }
  const ord = (d) => (d || "").split("/").reverse().join("-");
  return out.sort((a, b) => ord(b.data).localeCompare(ord(a.data)) ||
                            a.empresa.localeCompare(b.empresa));
}

function carregar(chave, destino) {
  const v = JSON.parse(localStorage.getItem(chave));
  estado.dados = { meta: v.meta, validacao: v.validacao, solicitacoes: v.solicitacoes };
  estado.pareceres = v.pareceres || {};
  estado.removidas = v.removidas || [];
  estado.itens = ordenar(v.solicitacoes, estado.pareceres);
  estado.i = Math.min(v.i || 0, estado.itens.length - 1);
  recalcularPoderes();   // o OTN pode ter mudado desde que este lote foi salvo
  irPara("revisao");
  render();
  // "Ver resumo" abre o relatório do dia direto; a tela de conferência fica
  // montada atrás, para o botão "Continuar conferindo" cair no lugar certo.
  if (destino === "resumo") mostrarResumo();
}

/* ------------------------------------------------------------------- utilidades */
/**
 * Ordena pela forma de pagamento e valor (ordem "natural" da conferência) e,
 * quando `pareceres` é informado, faz um segundo passe estável que empurra as
 * solicitações ainda sem parecer para depois das já conferidas — sem misturar
 * uma coisa com a outra. Sort é estável, então esse segundo passe preserva a
 * ordem natural dentro de cada grupo (conferidas / pendentes).
 */
const ordenar = (ss, pareceres) => {
  const base = [...ss].sort((a, b) => {
    const ta = ORDEM.indexOf(a.tipo), tb = ORDEM.indexOf(b.tipo);
    return (ta < 0 ? 99 : ta) - (tb < 0 ? 99 : tb) || (b.valor || 0) - (a.valor || 0);
  });
  if (!pareceres) return base;
  const pendente = (s) => (pareceres[s.sn]?.status ? 0 : 1);
  return base.sort((a, b) => pendente(a) - pendente(b));
};

const feitos = () => estado.itens.filter((s) => estado.pareceres[s.sn]?.status).length;

/**
 * Troca de tela e ajusta o "voltar" do cabeçalho.
 *
 * O voltar é de UM passo: do resumo ou da conferência volta-se para a tela do
 * relatório (onde estão as conferências salvas) e só de lá se sai para o
 * NOCTUS. Sair do módulo inteiro no meio de uma conferência era saída demais
 * para um clique só.
 */
function irPara(tela) {
  for (const t of ["upload", "revisao", "resumo", "config"])
    $("tela-" + t).classList.toggle("oculto", t !== tela);
  telaAtual = tela;
  const a = $("voltar-topo");
  if (a) {
    a.textContent = tela === "upload" ? "Voltar" : "Voltar ao relatório";
    a.href = tela === "upload" ? "../#/" : "#";
    a.title = tela === "upload" ? "" : "Voltar para a tela do relatório";
  }
  window.scrollTo(0, 0);
}

let telaAtual = "upload";

function ligarVoltar() {
  const a = $("voltar-topo");
  if (!a) return;
  a.onclick = (e) => {
    if (telaAtual === "upload") return;   // deixa o link levar ao NOCTUS
    e.preventDefault();
    if (telaAtual === "config") { fecharConfig(); return; }
    irPara("upload");
    renderRetomar();                      // os contadores mudaram desde que saiu daqui
  };
}

/* -------------------------------------------------------------------- configurações */
/**
 * Tela de configurações do módulo. Hoje só o OTN mora aqui, mas é o lugar de
 * qualquer parâmetro que o conferente precise ajustar sem mexer no código.
 */
let telaAntesConfig = "upload";
/** true quando as Configurações foram abertas direto da tela inicial do NOCTUS
 *  (botão Configurações no painel), não de dentro do módulo. */
let configAbertaViaHome = false;

function abrirConfig() {
  if (!ehAdmin()) return;
  if (telaAtual !== "config") telaAntesConfig = telaAtual;
  $("cfg-otn").value = otnAtual();
  $("cfg-msg").innerHTML = "";
  irPara("config");
  if (configAbertaViaHome) {
    const a = $("voltar-topo");
    if (a) { a.textContent = "Voltar"; a.title = "Voltar para o NOCTUS"; }
  }
}

/**
 * Volta pra tela de onde as configurações foram abertas, já redesenhada — ou,
 * se vieram direto da tela inicial do NOCTUS, volta pra lá em vez de cair
 * dentro do módulo.
 */
function fecharConfig() {
  if (configAbertaViaHome) { window.location.href = "../#/"; return; }
  const destino = telaAntesConfig;
  irPara(destino);
  if (destino === "upload") renderRetomar();
  else if (destino === "resumo") mostrarResumo();
  else if (estado.dados) render();
}

/**
 * Recalcula o apontamento de poder de todas as solicitações com o OTN vigente.
 *
 * O apontamento fica em `alertaPoder`, separado dos `alertas` que vêm do
 * parser: assim, quando o OTN muda, dá pra refazer só essa conta sem reler o
 * PDF e sem duplicar apontamento nenhum.
 */
const LIMITE_OTN_UNICO = 10;

/**
 * Calcula (e grava em cada solicitação) o apontamento de poder e o de 10 OTN
 * único, pro par (meta, solicitacoes) dado — função pura, sem depender de
 * `estado`, pra poder rodar também sobre um pedaço dividido de um Excel
 * multiempresa (processarMultiEmpresa/migrarChaves) ANTES dele existir em
 * `estado.dados`. `recalcularPoderes()` abaixo é só o atalho pro caso normal
 * (a conferência atualmente aberta em `estado`).
 */
function calcularPoderes(meta, solicitacoes) {
  const otn = otnAtual();
  for (const s of solicitacoes) {
    // PDF: um empresaCodigo só, do documento inteiro. Excel (pode trazer
    // várias empresas juntas — ver parseRelatorioExcel): cada solicitação
    // já carrega o SEU empresaCodigo, quando a conta bancária do bloco foi
    // reconhecida em contas.js E a empresa já tem código confirmado ali
    // (ainda não é o caso de todas — ver ATENÇÃO em contas.js); usa esse
    // quando existir, cai para o do documento quando não (PDF, ou Excel de
    // uma empresa só). Sem nenhum dos dois, verificarPoder devolve null
    // sozinho (empresa sem relação de poderes carregada) — não precisa de
    // tratamento especial aqui.
    const empresaCodigo = s.empresaCodigo || meta.empresaCodigo;
    s.alertaPoder = verificarPoder(empresaCodigo, s.competente,
                                   s.poder, s.valor, otn);
    s.alerta10otn = s.valor > otn * LIMITE_OTN_UNICO
      ? "Solicitação ultrapassa a 10 OTN" : null;
  }
}

function recalcularPoderes() {
  if (!estado.dados) return;
  calcularPoderes(estado.dados.meta, estado.dados.solicitacoes);
}

function salvarOtn(v) {
  if (!isFinite(v) || v <= 0) {
    $("cfg-msg").innerHTML =
      `<div class="alerta erro">Informe um valor de OTN maior que zero.</div>`;
    return;
  }
  localStorage.setItem(CHAVE_OTN, String(v));
  recalcularPoderes();
  if (estado.dados) salvar();
  $("cfg-otn").value = otnAtual();
  $("cfg-msg").innerHTML =
    `<div class="alerta ok">OTN atualizado para R$ ${moeda(otnAtual())}.
      Os apontamentos de poder da conferência aberta foram recalculados.</div>`;
}

function ligarConfig() {
  const abrir = $("btn-config");
  if (abrir) {
    if (!ehAdmin()) abrir.style.display = "none";
    else abrir.onclick = () => abrirConfig();
  }
  $("cfg-salvar").onclick = () =>
    salvarOtn(parseFloat(String($("cfg-otn").value).replace(",", ".")));
  $("cfg-padrao").onclick = () => {
    localStorage.removeItem(CHAVE_OTN);
    recalcularPoderes();
    if (estado.dados) salvar();
    $("cfg-otn").value = otnAtual();
    $("cfg-msg").innerHTML =
      `<div class="alerta ok">OTN voltou ao padrão de R$ ${moeda(OTN_PADRAO)}.</div>`;
  };
  $("cfg-concluir").onclick = () => fecharConfig();
}

/* ------------------------------------------------------------------------ upload */
function ligarUpload() {
  const solta = $("solta"), input = $("arquivo");
  solta.onclick = () => input.click();
  solta.ondragover = (e) => { e.preventDefault(); solta.classList.add("ativa"); };
  solta.ondragleave = () => solta.classList.remove("ativa");
  solta.ondrop = (e) => {
    e.preventDefault(); solta.classList.remove("ativa");
    if (e.dataTransfer.files[0]) processar(e.dataTransfer.files[0]);
  };
  input.onchange = () => input.files[0] && processar(input.files[0]);

  renderRetomar();
}

/** Quais dias estão expandidos na lista de conferências salvas (data -> aberto). */
const diasAbertos = {};

/**
 * Lista as conferências guardadas neste navegador.
 * `confirmando` é a chave do lote que está pedindo confirmação de remoção —
 * confirmação inline, e não confirm() nativo, que congela a página.
 */
function renderRetomar(confirmando) {
  const caixa = $("retomar");
  const salvos = lotesSalvos();
  renderRodapeGeral();
  if (!salvos.length) { caixa.classList.add("oculto"); caixa.innerHTML = ""; return; }
  caixa.classList.remove("oculto");

  // agrupa por dia, com as empresas daquele dia embaixo
  const dias = [];
  for (const l of salvos) {
    const ultimo = dias[dias.length - 1];
    if (ultimo && ultimo.data === l.data) ultimo.lotes.push(l);
    else dias.push({ data: l.data, lotes: [l] });
  }

  // o dia mais recente abre; os anteriores ficam recolhidos, senão a lista
  // cresce indefinidamente. A escolha do conferente vale enquanto a aba viver.
  dias.forEach((d, i) => {
    if (!(d.data in diasAbertos)) diasAbertos[d.data] = i === 0;
    // o dia que está pedindo confirmação de remoção não pode estar escondido
    if (confirmando && d.lotes.some((l) => l.chave === confirmando)) diasAbertos[d.data] = true;
  });

  const linha = (l) => l.chave === confirmando ? `
    <div class="lote lote--confirma">
      <span>Remover a conferência de <b>${l.empresa}</b> em ${l.data}?
        <span class="sub">${l.feitos ? `Os ${l.feitos} pareceres já dados serão perdidos.`
                                     : "Nenhum parecer foi dado nela."}</span></span>
      <span class="lote__acoes">
        <button class="botao fantasma" data-acao="cancelar">Cancelar</button>
        <button class="botao perigo" data-acao="remover" data-chave="${l.chave}">Remover</button>
      </span>
    </div>` : `
    <div class="lote${l.feitos === l.total ? " lote--completa" : ""}">
      <span class="lote__id">${l.empresa}
        <span class="sub">${l.feitos} de ${l.total} conferidas</span></span>
      <span class="lote__acoes">
        <button class="botao fantasma" data-acao="resumo" data-chave="${l.chave}"
                title="Abrir o relatório de pareceres deste dia">Ver resumo</button>
        <button class="botao fantasma" data-acao="retomar" data-chave="${l.chave}">Retomar</button>
        <button class="lote__x" data-acao="perguntar" data-chave="${l.chave}"
                title="Remover esta conferência" aria-label="Remover esta conferência">✕</button>
      </span>
    </div>`;

  caixa.innerHTML =
    `<p class="sub" style="margin-bottom:.6rem">Conferências em andamento neste navegador:</p>` +
    dias.map((d) => {
      const aberto = diasAbertos[d.data];
      const pend = d.lotes.filter((l) => l.feitos < l.total).length;
      return `
      <div class="dia">
        <div class="dia__cab">
          <button class="dia__toggle" data-acao="alternar-dia" data-dia="${d.data}"
                  aria-expanded="${aberto}">
            <span class="dia__seta">${aberto ? "▾" : "▸"}</span>${d.data}
            <span class="sub">${d.lotes.length === 1 ? "1 empresa"
                                                     : d.lotes.length + " empresas"}${
              aberto ? "" : pend ? ` · ${pend} em aberto` : " · tudo conferido"}</span>
          </button>
          <button class="dia__imprimir" data-acao="imprimir-dia" data-dia="${d.data}"
                  title="Gerar planilha com o resumo das empresas de ${d.data}">Imprimir</button>
        </div>
        ${aberto ? d.lotes.map(linha).join("") : ""}
      </div>`;
    }).join("");

  caixa.querySelectorAll("button").forEach((b) => {
    b.onclick = () => {
      const { acao, chave, dia } = b.dataset;
      if (acao === "alternar-dia") {
        diasAbertos[b.dataset.dia] = !diasAbertos[b.dataset.dia];
        renderRetomar(confirmando);
        return;
      }
      if (acao === "imprimir-dia") { imprimirDia(b, dia); return; }
      if (acao === "retomar") carregar(chave);
      else if (acao === "resumo") carregar(chave, "resumo");
      else if (acao === "perguntar") renderRetomar(chave);
      else if (acao === "cancelar") renderRetomar();
      else if (acao === "remover") {
        try { localStorage.removeItem(chave); } catch (e) { /* nada a fazer */ }
        renderRetomar();
      }
    };
  });
}

/**
 * Divide um relatório em Excel com várias empresas juntas (dados.meta.
 * multiEmpresa) numa lista de pedaços — um por empresa —, cada um já no
 * mesmo formato {meta, solicitacoes, validacao} de um relatório de UMA
 * empresa só (PDF, ou Excel que já veio de uma empresa única). Pedido do
 * Rocha em 23/09/2026: antes, um Excel com N empresas virava UMA conferência
 * só ("Múltiplas empresas (...)"), com os pareceres de todas misturados —
 * ruim pra emitir parecer por empresa e pra bater com o jeito que o PDF
 * sempre funcionou (uma conferência por empresa). A validação de totais de
 * cada pedaço (validacaoPorEmpresa, calculada em parseRelatorioExcel a
 * partir do total IMPRESSO das contas bancárias daquela empresa) é real,
 * não a validação do documento inteiro repetida pra todas — ver o
 * comentário de somarBlocosImpressos em parser.js.
 *
 * Solicitações com conta de origem não cadastrada em contas.js (empresa não
 * identificada automaticamente) viram seu próprio pedaço, "Empresa não
 * identificada" — não desaparecem nem ficam misturadas dentro de nenhuma
 * empresa de verdade.
 */
function dividirPorEmpresa(dados) {
  const { meta, solicitacoes, validacaoPorEmpresa } = dados;

  // nome curto -> código de empresa: o primeiro que aparecer com código,
  // mesma regra que parseRelatorioExcel já usa pra montar o rótulo combinado.
  const codigoDoNome = new Map();
  for (const s of solicitacoes) {
    if (s.empresa && s.empresaCodigo && !codigoDoNome.has(s.empresa))
      codigoDoNome.set(s.empresa, s.empresaCodigo);
  }
  const rotuloDoNome = (nomeCurto) => rotuloEmpresa(codigoDoNome.get(nomeCurto)) || nomeCurto;

  const grupos = new Map(); // chave "" = sem empresa identificada
  for (const s of solicitacoes) {
    const chave = s.empresa || "";
    if (!grupos.has(chave)) grupos.set(chave, []);
    grupos.get(chave).push(s);
  }
  // ordem estável: empresas na ordem em que apareceram no arquivo, "sem
  // empresa" por último quando existir.
  const ordem = [...meta.empresas, ...(grupos.has("") ? [""] : [])];

  return ordem.filter((chave) => grupos.has(chave)).map((chave) => {
    const sols = grupos.get(chave);
    const semEmpresa = chave === "";
    const rotuloDaChave = (c) => (c === "" ? "Empresa não identificada" : rotuloDoNome(c));
    return {
      meta: {
        ...meta,
        empresa: rotuloDaChave(chave),
        empresaCodigo: semEmpresa ? null : (codigoDoNome.get(chave) || null),
        empresaNome: semEmpresa ? null : chave,
        empresas: semEmpresa ? [] : [chave],
        multiEmpresa: false,
      },
      solicitacoes: sols,
      validacao: validacaoPorEmpresa?.[chave] || dados.validacao,
    };
  });
}

/**
 * Excel com várias empresas juntas: cada empresa vira (ou atualiza) a sua
 * PRÓPRIA conferência salva, mesclada com a que já estava em andamento pra
 * ela — mesma lógica de mesclar() de sempre, uma vez por empresa. Não abre
 * nenhuma automaticamente — não haveria uma escolha óbvia entre N empresas —
 * só salva todas e volta pra tela inicial, com a lista de "Conferências em
 * andamento" já mostrando cada uma separada.
 */
function processarMultiEmpresa(dados) {
  const partes = dividirPorEmpresa(dados);
  const resumos = [];
  // Pedido do Rocha em 24/09/2026: quando o relatório novo sai com menos
  // solicitações do que o lote salvo (uma foi cancelada/removida no SCK
  // entre um upload e outro), isso NUNCA aparecia pro conferente neste
  // caminho (Excel com várias empresas): só novas/alteradas eram citadas por
  // empresa, sumida nenhuma. No caminho de empresa única (PDF ou Excel de
  // uma empresa só) esse aviso já existe (ver "aviso-mesclagem" em render());
  // aqui replica a mesma contagem, por auditoria — um parecer já dado que
  // some sem aviso é exatamente o tipo de coisa que não pode passar batido.
  //
  // Pedido do Rocha em 25/09/2026: além de avisar, a solicitação sumida que
  // JÁ tinha parecer não pode mais perder o motivo — mesclar() agora devolve
  // `removidas` (ver comentário da função), que cada empresa grava junto do
  // lote e a tela de Resumo lista à parte, fora dos totais/validação do
  // relatório atual.
  let totalSumiram = 0;
  for (const parte of partes) {
    // mesmo cálculo que um upload de empresa única já fazia antes de salvar
    // (recalcularPoderes) — sem isso, o apontamento de poder só apareceria
    // na planilha impressa depois que alguém abrisse ("Retomar") esta
    // empresa pelo menos uma vez.
    calcularPoderes(parte.meta, parte.solicitacoes);
    const chave = chaveLote(parte);
    const salvoAnterior = localStorage.getItem(chave);
    let pareceres = {}, removidas = [], notaMerge = "";
    if (salvoAnterior) {
      const r = mesclar(JSON.parse(salvoAnterior), parte);
      pareceres = r.pareceres;
      removidas = r.removidas;
      totalSumiram += r.sumiram.length;
      const partes2 = [r.novas.length && `${r.novas.length} nova(s)`,
                        r.alteradas.length && `${r.alteradas.length} alterada(s)`,
                        r.sumiram.length && `${r.sumiram.length} já conferida(s) saiu/saíram do relatório (${r.sumiram.join(", ")})`,
                       ].filter(Boolean);
      if (partes2.length) notaMerge = ` — ${partes2.join(", ")}`;
    }
    try {
      localStorage.setItem(chave, JSON.stringify({
        meta: parte.meta, validacao: parte.validacao,
        solicitacoes: parte.solicitacoes, pareceres, removidas, i: 0,
      }));
    } catch (e) { /* modo privado, cota cheia: esta empresa não persiste */ }
    resumos.push(`${nomeEmpresa(parte.meta)} (${parte.solicitacoes.length}${notaMerge})`);
  }

  // Pedido do Rocha em 23/09/2026: a mensagem antiga listava o detalhe de
  // TODAS as empresas na cara, ficando enorme — vira ruído toda vez que se
  // importa um Excel. Agora o texto principal é curto (dispensável com o
  // ✕, como o aviso de mesclagem acima) e o detalhe por empresa — que ainda
  // é informação de auditoria útil (quantas solicitações entraram em cada
  // uma, e quantas foram mescladas) — fica atrás de "Ver detalhes",
  // recolhido por padrão em vez de sumir de vez. O aviso de sumidas (linha
  // abaixo) segue essa mesma regra: uma frase curta no corpo principal —
  // pra não passar despercebido — com a lista de S.N. específica de cada
  // empresa só dentro de "Ver detalhes", igual ao resto do detalhe.
  const avisoSumiram = totalSumiram
    ? ` ${totalSumiram} solicitação(ões) que você já tinha conferido saiu/saíram do relatório — motivo e
       status ficam guardados em "Removidas do relatório", no Resumo de cada empresa. Ver detalhes.`
    : "";
  $("upload-msg").innerHTML = `<div class="alerta ok alerta--removivel">
    <b>Relatório importado com sucesso.</b> ${partes.length}
    empresa${partes.length > 1 ? "s" : ""} atualizada${partes.length > 1 ? "s" : ""},
    cada uma na sua própria conferência. Escolha uma abaixo para continuar.${avisoSumiram}
    <details class="upload-msg__detalhe"><summary>Ver detalhes</summary>${resumos.join("; ")}</details>
    <button class="alerta__x" id="btn-fecha-upload-msg" title="Dispensar">✕</button>
  </div>`;
  const btnFechaUpload = $("btn-fecha-upload-msg");
  if (btnFechaUpload) btnFechaUpload.onclick = () => { $("upload-msg").innerHTML = ""; };
  irPara("upload");
  renderRetomar();
}

/**
 * Aceita o relatório em PDF ou em Excel (.xlsx/.xls) — mesmo relatório,
 * dois formatos de exportação do SCK (ver comentário no topo de parser.js
 * para o porquê de preferir o Excel quando ele estiver disponível). Os
 * dois caminhos convergem no mesmo `dados` ({meta, solicitacoes,
 * validacao}) e dali pra baixo o fluxo é idêntico — exceto quando o Excel
 * traz várias empresas juntas, que se separa em várias conferências
 * (processarMultiEmpresa) em vez de virar uma só.
 */
async function processar(arquivo) {
  const msg = $("upload-msg");
  const nome = arquivo.name.toLowerCase();
  const ehPdf = nome.endsWith(".pdf");
  const ehExcel = nome.endsWith(".xlsx") || nome.endsWith(".xls");
  if (!ehPdf && !ehExcel) {
    msg.innerHTML = `<div class="alerta erro">Envie o relatório em PDF ou em Excel (.xlsx).</div>`;
    return;
  }
  msg.innerHTML = `<div class="alerta">Lendo o relatório…</div>`;
  try {
    let dados;
    if (ehPdf) {
      dados = await parseRelatorio(new Uint8Array(await arquivo.arrayBuffer()), pdfjsLib);
    } else {
      await carregarSheetJS();
      dados = parseRelatorioExcel(await arquivo.arrayBuffer(), window.XLSX);
    }
    if (!dados.solicitacoes.length) {
      msg.innerHTML = `<div class="alerta erro">Nenhuma solicitação encontrada.
        O layout do relatório mudou?</div>`;
      return;
    }

    if (dados.meta.multiEmpresa) {
      processarMultiEmpresa(dados);
      return;
    }

    msg.innerHTML = "";

    const salvo = localStorage.getItem(chaveLote(dados));
    estado.dados = dados;
    estado.i = 0;
    recalcularPoderes();

    if (salvo) {
      const r = mesclar(JSON.parse(salvo), dados);
      estado.pareceres = r.pareceres;
      estado.removidas = r.removidas;
      estado.mesclagem = r;
    } else {
      estado.pareceres = {};
      estado.removidas = [];
      estado.mesclagem = null;
    }
    // com os pareceres já mesclados: os ainda pendentes ficam depois dos já
    // conferidos, sem se misturar entre eles.
    estado.itens = ordenar(dados.solicitacoes, estado.pareceres);
    salvar();
    irPara("revisao");
    render();
  } catch (e) {
    msg.innerHTML = `<div class="alerta erro">Não consegui ler o relatório: ${e.message}</div>`;
  }
}

/* -------------------------------------------------------------------- conferência */
function render() {
  const { validacao } = estado.dados;
  const m = estado.mesclagem;
  $("aviso-mesclagem").innerHTML =
    !m || (!m.novas.length && !m.alteradas.length && !m.sumiram.length && !m.reapareceram.length)
    ? ""
    : `<div class="alerta ok">
        <b>Relatório atualizado.</b> Seus pareceres foram mantidos.
        ${m.novas.length ? `${m.novas.length} solicitação(ões) nova(s).` : ""}
        ${m.alteradas.length ? `${m.alteradas.length} mudou/mudaram desde a última conferência
           e voltaram a pendente — o texto do parecer ficou guardado.` : ""}
        ${m.sumiram.length ? `${m.sumiram.length} que você já tinha conferido saiu/saíram do
           relatório: ${m.sumiram.join(", ")} — motivo e status ficam guardados em
           "Removidas do relatório", no Resumo.` : ""}
        ${m.reapareceram.length ? `${m.reapareceram.length} que tinha(m) sumido voltou/voltaram a
           aparecer: ${m.reapareceram.join(", ")} — o parecer anterior foi restaurado.` : ""}
        <button class="alerta__x" id="btn-fecha-mesclagem" title="Dispensar">✕</button>
      </div>`;
  if (m) {
    const x = $("btn-fecha-mesclagem");
    if (x) x.onclick = () => { estado.mesclagem = null; $("aviso-mesclagem").innerHTML = ""; };
  }

  const avisoTotais = validacao.confere ? "" : `
      <div class="alerta erro"><b>Atenção:</b> o que extraí não bateu com os totais impressos
        (${validacao.qtdExtraida} × ${validacao.qtdRelatorio} solicitações,
        R$ ${moeda(validacao.valorExtraido)} × R$ ${moeda(validacao.valorRelatorio)}).
        Confira o relatório antes de emitir o parecer.</div>`;

  // Só o Excel pode ter uma conta de origem que não bateu com contas.js (ver
  // parseRelatorioExcel, em parser.js). O aviso "veio de um Excel com N
  // empresas juntas" (contexto de origemMultiEmpresa, sem nenhuma ação
  // pendente) foi removido a pedido do Rocha em 25/09/2026 — mesmo com o ✕
  // pra dispensar (pedido dele em 23/09/2026), o "dispensado" só vivia na
  // memória da aba (estado.avisoOrigemDispensado), então voltava toda vez
  // que a conferência era reaberta (via "Retomar") ou a página recarregava —
  // virou incômodo em vez de contexto único. "Empresa não identificada"
  // continua de fora dessa remoção, de propósito: aponta solicitação(ões)
  // específica(s) sem conta cadastrada, isso é um risco de verdade (pode ser
  // poder mal conferido) e não deve sumir da tela sem o Rocha resolver.
  const meta = estado.dados.meta;
  let avisoExcel = "";
  if (meta?.formato === "excel") {
    const semEmpresa = estado.dados.solicitacoes.filter((x) => !x.empresa).length;
    if (semEmpresa)
      avisoExcel += `<div class="alerta erro">${semEmpresa} solicitação(ões) com conta de origem não cadastrada em contas.js
        — empresa não identificada automaticamente (veja o apontamento em cada uma).</div>`;
  }

  $("aviso-extracao").innerHTML = avisoTotais + avisoExcel;

  const s = estado.itens[estado.i];
  const total = estado.itens.length;
  $("rev-tipo").textContent = s.tipo;
  $("rev-contador").textContent = `solicitação ${estado.i + 1} de ${total} · ${feitos()} conferidas`;
  $("rev-barra").style.width = (feitos() / total * 100).toFixed(1) + "%";

  const listaApont = apontamentosLista(s);
  const ocultosApont = alertasOcultosDe(s.sn);
  const visiveisApont = listaApont.filter((a) => !ocultosApont.includes(a));
  $("rev-alertas").innerHTML =
    visiveisApont.map((a) => `
      <div class="alerta alerta--removivel">
        <span>${escAnexo(a)}</span>
        <button type="button" class="alerta__x" data-remover-apont
                title="Remover este apontamento — some da tela, do resumo, da planilha e do relatório impresso">✕</button>
      </div>`).join("") +
    (ocultosApont.length
      ? `<div class="sub apont-ocultos">
           <button type="button" class="linkbtn" id="btn-restaurar-apont">Restaurar Observação</button>
         </div>`
      : "");
  $("rev-alertas").querySelectorAll("[data-remover-apont]").forEach((btn, i) => {
    btn.onclick = () => { ocultarApontamento(s.sn, visiveisApont[i]); render(); };
  });
  const btnRestaurarApont = $("btn-restaurar-apont");
  if (btnRestaurarApont) btnRestaurarApont.onclick = () => { restaurarApontamentos(s.sn); render(); };

  $("c-sn").textContent = s.sn;
  $("c-valor").textContent = "R$ " + moeda(s.valor);
  $("c-poder").textContent = s.poder || "—";
  $("c-solicitante").textContent = s.solicitante || "—";
  $("c-competente").textContent = s.competente || "—";
  $("c-favorecido").textContent =
    [s.favorecido, s.cpfCnpj].filter(Boolean).join("  ·  ") || "—";
  $("c-destinacao").textContent = s.destinacao || "—";

  const p = estado.pareceres[s.sn] || {};
  for (const st of STATUS) $("st-" + st.id).checked = p.status === st.valor;
  $("rev-parecer").value = p.parecer || "";
  $("btn-anterior").disabled = estado.i === 0;
  $("btn-proxima").textContent =
    estado.i + 1 < total ? "Salvar e próxima" : "Salvar e finalizar";

  carregarAnexos();
}

function salvarAtual() {
  const s = estado.itens[estado.i];
  const st = document.querySelector('input[name="status"]:checked');
  const anterior = estado.pareceres[s.sn] || {};
  estado.pareceres[s.sn] = {
    status: st ? st.value : "",
    parecer: $("rev-parecer").value.trim(),
    ...(anterior.alertasOcultos?.length ? { alertasOcultos: anterior.alertasOcultos } : {}),
  };
  salvar();
}

function avancar(passo) {
  const novo = estado.i + passo;
  if (novo < 0 || novo >= estado.itens.length) return false;
  estado.i = novo; render(); return true;
}

function ligarRevisao() {
  $("btn-proxima").onclick = () => {
    salvarAtual();
    if (!avancar(1)) mostrarResumo();
  };
  $("btn-anterior").onclick = () => { salvarAtual(); avancar(-1); };
  $("btn-pular").onclick = () => { if (!avancar(1)) mostrarResumo(); };
  $("btn-resumo").onclick = () => { salvarAtual(); mostrarResumo(); };

  document.addEventListener("keydown", (e) => {
    if ($("tela-revisao").classList.contains("oculto")) return;
    const digitando = e.target.tagName === "TEXTAREA";
    if (+e.key >= 1 && +e.key <= STATUS.length && !digitando) {
      $("st-" + STATUS[+e.key - 1].id).checked = true; e.preventDefault();
    } else if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      $("btn-proxima").click(); e.preventDefault();
    } else if (e.key === "ArrowRight" && !digitando) {
      salvarAtual(); avancar(1); e.preventDefault();
    } else if (e.key === "ArrowLeft" && !digitando) {
      salvarAtual(); avancar(-1); e.preventDefault();
    }
  });
}

/* ------------------------------------------------------------------------- resumo */
const APONTADAS = "@apontadas";   // valor de filtro que não colide com nome de forma de pagamento
/** Apontamentos do parser + o de poder, que é recalculado quando o OTN muda. */
const apontamentosLista = (s) =>
  [...(s.alertas || []), ...(s.alertaPoder ? [s.alertaPoder] : []),
   ...(s.alerta10otn ? [s.alerta10otn] : [])];

/**
 * Apontamentos automáticos que o conferente removeu manualmente desta solicitação
 * (falso positivo, já esclarecido etc.). Ficam guardados junto do parecer e somem de
 * tudo — tela, resumo, planilha e relatório impresso —, não só da tela de revisão.
 */
const alertasOcultosDe = (sn) => estado.pareceres[sn]?.alertasOcultos || [];
const apontamentosVisiveis = (s) =>
  apontamentosLista(s).filter((a) => !alertasOcultosDe(s.sn).includes(a));
const temApontamento = (s) => apontamentosVisiveis(s).length > 0;
const apontamentosDe = (s) => apontamentosVisiveis(s).join(" · ");
/**
 * true se o sistema já gerou algum apontamento pra esta solicitação, mesmo
 * que o conferente tenha removido todos depois. Usado só no filtro "com
 * apontamento" do resumo, pra ainda ser possível localizar essas solicitações
 * — o texto do apontamento em si continua sumindo de tudo o mais quando
 * removido (tela, planilha, impresso).
 */
const temApontamentoOriginal = (s) => apontamentosLista(s).length > 0;

function ocultarApontamento(sn, texto) {
  const p = estado.pareceres[sn] || { status: "", parecer: "" };
  const atuais = p.alertasOcultos || [];
  if (atuais.includes(texto)) return;
  estado.pareceres[sn] = { ...p, alertasOcultos: [...atuais, texto] };
  salvar();
}

function restaurarApontamentos(sn) {
  const p = estado.pareceres[sn];
  if (!p || !(p.alertasOcultos || []).length) return;
  estado.pareceres[sn] = { ...p, alertasOcultos: [] };
  salvar();
}

const SEM_STATUS = "Sem conferir";
const statusDoItem = (s) => estado.pareceres[s.sn]?.status || SEM_STATUS;
/** A linha de filtro por status fica recolhida ou não, por escolha do conferente. */
const CHAVE_FILTRO_ST = "noctus.filtro-status";
const filtroStatusAberto = () => localStorage.getItem(CHAVE_FILTRO_ST) !== "oculto";
const CLASSE_STATUS = { Aprovado: "g", "Aguardando esclarecimentos": "b",
                        Recusado: "v", [SEM_STATUS]: "n" };

let filtro = null;         // forma de pagamento (ou apontadas)
let filtroStatus = null;   // status do parecer, combinado com o filtro acima
let busca = "";            // texto livre buscado na observação/destinação
let ultimoAberto = null;   // cartão de onde o conferidor foi aberto

function mostrarResumo() {
  irPara("resumo");
  $("res-data").textContent =
    "· " + textoPeriodo(estado.dados.meta) +
    (estado.dados.meta.empresa ? " · " + estado.dados.meta.empresa : "");

  const contagem = Object.fromEntries(STATUS.map((s) => [s.valor, 0]));
  contagem[SEM_STATUS] = 0;
  for (const s of estado.itens) contagem[statusDoItem(s)]++;
  const classe = CLASSE_STATUS;
  $("res-totais").innerHTML =
    `<div><span>Solicitações</span><b>${estado.itens.length}</b></div>
     <div><span>Valor total</span><b>R$ ${moeda(estado.dados.validacao.valorExtraido)}</b></div>`;
  $("res-status").innerHTML = Object.entries(contagem).filter(([, n]) => n)
    .map(([k, n]) => `<div class="status-card status-card--${classe[k]}">
        <b>${n}</b><span>${k}</span></div>`).join("");

  $("res-pendencia").innerHTML = contagem[SEM_STATUS]
    ? `<div class="alerta">${contagem[SEM_STATUS]} solicitação(ões) ainda sem status.</div>`
    : `<div class="alerta ok">Todas as solicitações conferidas.</div>`;

  const apontadas = estado.itens.filter(temApontamentoOriginal).length;
  $("res-filtros").innerHTML =
    [`<span class="chip ${filtro ? "" : "on"}" data-t="">Todas</span>`]
      .concat(ORDEM.filter((t) => estado.itens.some((s) => s.tipo === t))
        .map((t) => `<span class="chip ${filtro === t ? "on" : ""}" data-t="${t}">${t}</span>`))
      .concat(apontadas ? [`<span class="chip chip--apontada ${filtro === APONTADAS ? "on" : ""}"
        data-t="${APONTADAS}" title="Solicitações que o sistema marcou"
        >${apontadas} com apontamento</span>`] : [])
      .join("");
  $("res-filtros").querySelectorAll(".chip").forEach((c) => {
    c.onclick = () => { filtro = c.dataset.t || null; mostrarResumo(); };
  });

  // Filtro de status: vale DENTRO da aba escolhida acima, com as contagens da
  // própria aba. Sem aba nova — quem está olhando "TED" quer ver os recusados
  // de TED, não trocar de lista.
  const naAba = estado.itens.filter((s) =>
    filtro === APONTADAS ? temApontamentoOriginal(s) : (!filtro || s.tipo === filtro));
  const porStatus = {};
  for (const s of naAba) porStatus[statusDoItem(s)] = (porStatus[statusDoItem(s)] || 0) + 1;
  if (filtroStatus && !porStatus[filtroStatus]) filtroStatus = null;   // sumiu nesta aba
  $("res-status-filtros").innerHTML = !filtroStatusAberto()
    ? `<button type="button" class="linkbtn" data-abrir="1">filtrar por status</button>`
    : [`<span class="chip-st ${filtroStatus ? "" : "on"}" data-s="">Todos os status</span>`]
        .concat([...STATUS.map((s) => s.valor), SEM_STATUS]
          .filter((n) => porStatus[n])
          .map((n) => `<span class="chip-st chip-st--${classe[n]} ${filtroStatus === n ? "on" : ""}"
            data-s="${n}">${n} <b>${porStatus[n]}</b></span>`))
        .concat([`<button type="button" class="linkbtn" data-abrir="0">ocultar</button>`])
        .join("");
  $("res-status-filtros").querySelectorAll(".chip-st").forEach((c) => {
    c.onclick = () => { filtroStatus = c.dataset.s || null; mostrarResumo(); };
  });
  $("res-status-filtros").querySelectorAll("[data-abrir]").forEach((b) => {
    b.onclick = () => {
      const abrir = b.dataset.abrir === "1";
      localStorage.setItem(CHAVE_FILTRO_ST, abrir ? "aberto" : "oculto");
      if (!abrir) filtroStatus = null;   // filtro escondido e ativo esconderia linha da lista
      mostrarResumo();
    };
  });

  const esc = (t) => String(t ?? "").replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
  // Busca por texto (canto direito da barra de filtros — pedido do Rocha em
  // 24/09/2026: "supondo coloco 'Almoço' irá trazer todas as solicitações
  // que na observação contenha almoço"). O relatório não tem um campo único
  // e sempre presente chamado "observação": a coluna "OBSERVAÇÃO" só existe
  // de fato nas famílias DÉBITO EM CONTA/BOLETO (vira s.complemento — ver
  // montarComplementoExcel/COLUNAS em parser.js); em TED/TRANSFERÊNCIA/PIX
  // esse mesmo s.complemento guarda dado bancário ou chave PIX, não uma
  // observação. Por isso a busca cobre DOIS campos de texto livre de cada
  // solicitação — s.destinacao (a "Destinação", sempre presente e é o texto
  // mostrado em todo cartão da lista) e s.complemento (a "Observação" de
  // fato, quando a solicitação for débito/boleto) — maximizando o que o
  // termo pode encontrar sem tocar em campos que não são texto descritivo
  // (favorecido, CPF/CNPJ, status etc.). Ponto que pode precisar de
  // validação com o Rocha: se algum dia ele quiser a busca restrita só à
  // coluna "Observação" literal (não à Destinação), o campo de busca abaixo
  // deve trocar de `textoBusca` para olhar só `s.complemento`.
  const termoBusca = busca.trim().toLowerCase();
  const textoBusca = (s) => `${s.destinacao || ""} ${s.complemento || ""}`.toLowerCase();
  const visiveis = naAba
    .filter((s) => !filtroStatus || statusDoItem(s) === filtroStatus)
    .filter((s) => !termoBusca || textoBusca(s).includes(termoBusca));
  $("res-qtd").textContent = `· ${visiveis.length}` +
    (filtro === APONTADAS ? " com apontamento" : filtro ? " em " + filtro : "") +
    (filtroStatus ? ` · ${filtroStatus.toLowerCase()}` : "") +
    // .textContent (não innerHTML) já escapa sozinho — nada de esc() aqui,
    // senão um termo com "&" apareceria como "&amp;" literal na tela.
    (termoBusca ? ` · observação contém "${busca.trim()}"` : "");
  $("res-corpo").innerHTML = visiveis
    .map((s) => {
      const p = estado.pareceres[s.sn] || {};
      return `<div class="parecer${p.status ? "" : " parecer--pendente"}"
                   data-sn="${esc(s.sn)}" role="button" tabindex="0"
                   title="Abrir esta solicitação para conferir">
        <div class="parecer__topo">
          <span class="parecer__sn">${s.sn}</span>
          <span class="parecer__valor">R$ ${moeda(s.valor)}</span>
          ${p.status
            ? `<span class="marca marca--${idDoStatus(p.status)}">${p.status}</span>`
            : `<span class="marca marca--nenhum">Sem conferir</span>`}
          <span class="parecer__abrir">Abrir</span>
        </div>
        <div class="parecer__fav">${esc(s.favorecido)}</div>
        <div class="parecer__dest">${esc(s.destinacao)}</div>
        ${temApontamento(s)
          ? `<div class="parecer__apontamento">${esc(apontamentosDe(s))}</div>` : ""}
        ${p.parecer ? `<div class="parecer__texto">${esc(p.parecer)}</div>` : ""}
      </div>`;
    }).join("");

  if (ultimoAberto) {
    const alvo = $("res-corpo").querySelector(`[data-sn="${CSS.escape(ultimoAberto)}"]`);
    if (alvo && !$("res-lista").hidden) alvo.scrollIntoView({ block: "center" });
  }

  // Pedido do Rocha em 25/09/2026: solicitações que sumiram de um relatório
  // mais novo mas já tinham parecer (ver mesclar()/estado.removidas) ficam
  // listadas aqui, num cartão à parte — de propósito FORA de res-totais/
  // res-status/validacao acima, que continuam refletindo só o relatório
  // atual. Cartão inteiro escondido quando não há nenhuma (a maioria dos
  // dias), pra não virar ruído permanente na tela.
  const removidas = estado.removidas || [];
  $("card-removidas").hidden = !removidas.length;
  $("res-removidas-qtd").textContent = removidas.length ? `· ${removidas.length}` : "";
  $("res-removidas-corpo").innerHTML = [...removidas]
    .sort((a, b) => (b.valor || 0) - (a.valor || 0))
    .map((r) => `<div class="parecer parecer--removida">
        <div class="parecer__topo">
          <span class="parecer__sn">${esc(r.sn)}</span>
          <span class="parecer__valor">R$ ${moeda(r.valor)}</span>
          <span class="marca marca--${idDoStatus(r.status)}">${esc(r.status)}</span>
        </div>
        <div class="parecer__fav">${esc(r.favorecido)}</div>
        <div class="parecer__dest">${esc(r.destinacao)}</div>
        ${r.parecer ? `<div class="parecer__texto">${esc(r.parecer)}</div>` : ""}
        <div class="parecer__removida-nota">Não consta mais no relatório importado em
          ${esc(r.removidoEm || "data não registrada")}.</div>
      </div>`).join("");
}

/** Abre no conferidor a solicitação do cartão clicado. */
function abrirSolicitacao(sn) {
  const i = estado.itens.findIndex((s) => s.sn === sn);
  if (i < 0) return;
  ultimoAberto = sn;
  estado.i = i;
  irPara("revisao");
  render();
}

function ligarResumo() {
  // clique em qualquer ponto do cartão abre a solicitação, mas sem atrapalhar
  // quem estiver só selecionando texto para copiar
  $("res-corpo").onclick = (e) => {
    if (String(window.getSelection())) return;
    const cartao = e.target.closest(".parecer");
    if (cartao) abrirSolicitacao(cartao.dataset.sn);
  };
  $("res-corpo").onkeydown = (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    const cartao = e.target.closest(".parecer");
    if (cartao) { e.preventDefault(); abrirSolicitacao(cartao.dataset.sn); }
  };
  $("btn-voltar").onclick = () => { irPara("revisao"); render(); };
  $("btn-imprimir").onclick = () => window.print();
  $("btn-planilha").onclick = gerarPlanilha;
  // Elemento fixo em HTML (nunca recriado pelo innerHTML de mostrarResumo),
  // de propósito: se ele fosse gerado dentro de #res-filtros/innerHTML, cada
  // tecla digitada destruiria e recriaria o <input>, perdendo o foco e o
  // cursor no meio da digitação. Por ficar fora desses containers, o valor e
  // o foco sobrevivem normalmente a cada re-render disparado pelo oninput.
  $("res-busca").oninput = (e) => { busca = e.target.value; mostrarResumo(); };
  $("btn-lista").onclick = () => {
    const lista = $("res-lista");
    lista.hidden = !lista.hidden;
    $("btn-lista").textContent = lista.hidden ? "Mostrar" : "Ocultar";
    $("btn-lista").setAttribute("aria-expanded", String(!lista.hidden));
  };
  $("btn-removidas").onclick = () => {
    const lista = $("res-removidas-lista");
    lista.hidden = !lista.hidden;
    $("btn-removidas").textContent = lista.hidden ? "Mostrar" : "Ocultar";
    $("btn-removidas").setAttribute("aria-expanded", String(!lista.hidden));
  };
}

/* ---------------------------------------------------------------------- planilha */
const AZUL = "FF1F3864";        // cabeçalho
const AMBAR = "FF9C6500";       // apontamento automático
const CINZA_LINHA = "FFD9D9D9"; // divisórias
const COR_STATUS = {
  "Aprovado": "FF15803D",
  "Aguardando esclarecimentos": "FF1D4ED8",
  "Recusado": "FFB91C1C",
};

/**
 * A biblioteca de planilha (~950KB) só é carregada quando alguém clica em gerar.
 * Tenta a cópia local primeiro — se ela não estiver no servidor, cai no CDN.
 */
async function carregarExcelJS() {
  if (window.ExcelJS) return;
  const origens = [
    "./vendor/exceljs.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
  ];
  for (const src of origens) {
    try {
      await new Promise((ok, falha) => {
        const s = document.createElement("script");
        s.src = src; s.onload = ok; s.onerror = () => falha(new Error(src));
        document.head.appendChild(s);
      });
      if (window.ExcelJS) return;
    } catch (e) { /* tenta a próxima origem */ }
  }
  throw new Error("não consegui carregar a biblioteca de planilha");
}

/**
 * Biblioteca de LEITURA de planilha (SheetJS, ~880KB) — usada só quando o
 * relatório é enviado em .xlsx/.xls, mesmo padrão de carregarExcelJS()
 * acima (que é só de ESCRITA, pros botões "Gerar planilha"/"Imprimir").
 * Tenta a cópia local primeiro (já está em vendor/), cai no CDN se faltar.
 */
async function carregarSheetJS() {
  if (window.XLSX) return;
  const origens = [
    "./vendor/xlsx.full.min.js",
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
    } catch (e) { /* tenta a próxima origem */ }
  }
  throw new Error("não consegui carregar a biblioteca de leitura de planilha");
}

function baixar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = nome; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ------------------------------------------------------------------ anexos */
/**
 * Anexos (PDF, .msg etc.) de cada solicitação, guardados no Supabase.
 * Usa a MESMA sessão que o login do NOCTUS grava em localStorage — este
 * módulo não tem tela de login própria; se não houver sessão válida, só
 * mostra um aviso pedindo para entrar no site principal.
 */
let anexosEstado = { chave: null, lista: [] };

function sessaoAtual() {
  const bruto = localStorage.getItem(CHAVE_SESSAO);
  if (!bruto) return null;
  try {
    const s = JSON.parse(bruto);
    if (!s.accessToken || !s.expiresAt || s.expiresAt < Date.now()) return null;
    return s;
  } catch {
    return null;
  }
}

/**
 * Renova a sessão de login usando o refreshToken salvo (mesmo grant_type que
 * o painel principal usa). Só é chamada de perto do vencimento do token e com
 * o conferente ativo — ver ligarRenovacaoDeSessao().
 */
async function renovarSessao() {
  const bruto = localStorage.getItem(CHAVE_SESSAO);
  if (!bruto) return null;
  let s;
  try {
    s = JSON.parse(bruto);
  } catch {
    return null;
  }
  if (!s.refreshToken) return null;
  let resp;
  try {
    resp = await fetch(SB_URL + "/auth/v1/token?grant_type=refresh_token", {
      method: "POST",
      headers: { apikey: SB_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ refresh_token: s.refreshToken }),
    });
  } catch {
    return null;
  }
  const dados = await resp.json().catch(() => ({}));
  if (!resp.ok || !dados.access_token) return null;
  const nova = {
    user: s.user,
    accessToken: dados.access_token,
    refreshToken: dados.refresh_token || s.refreshToken,
    expiresAt: Date.now() + (dados.expires_in || 3600) * 1000,
  };
  localStorage.setItem(CHAVE_SESSAO, JSON.stringify(nova));
  return nova;
}

/**
 * Mantém a sessão viva enquanto o conferente está usando o módulo: renova
 * pouco antes do token vencer, mas só se houve atividade recente. Depois de
 * LIMITE_INATIVIDADE sem clique/tecla/scroll, para de renovar e deixa a
 * sessão expirar no prazo normal — não mantém a sessão viva pra sempre com a
 * aba esquecida aberta e sem uso.
 */
function ligarRenovacaoDeSessao() {
  let ultimaAtividade = Date.now();
  const marcar = () => { ultimaAtividade = Date.now(); };
  ["click", "keydown", "mousemove", "scroll", "touchstart"].forEach((ev) =>
    window.addEventListener(ev, marcar, { passive: true }));

  const LIMITE_INATIVIDADE = 60 * 60 * 1000; // 60min sem atividade: para de renovar
  const MARGEM_RENOVACAO = 5 * 60 * 1000;    // renova quando faltar menos de 5min
  const INTERVALO = 60 * 1000;               // confere a cada 1min

  setInterval(async () => {
    const s = sessaoAtual();
    if (!s) return;
    if (Date.now() - ultimaAtividade > LIMITE_INATIVIDADE) return;
    if (s.expiresAt - Date.now() > MARGEM_RENOVACAO) return;
    await renovarSessao();
  }, INTERVALO);
}

/** Usuário logado (mesma sessão do painel principal), ou null se ninguém entrou. */
function usuarioAtual() {
  return sessaoAtual()?.user || null;
}

/** true só para quem tem o papel "admin" (hoje: só o usuário fixo original). */
function ehAdmin() {
  const u = usuarioAtual();
  return !!u && Array.isArray(u.roles) && u.roles.includes("admin");
}

/**
 * Identificador do usuário atual pra separar os dados de cada um no localStorage
 * (cada conferente só vê e mexe nas próprias conferências em andamento). Sem
 * sessão, cai num escopo fixo — mesmo comportamento de antes do cadastro existir.
 */
function escopoUsuario() {
  const u = usuarioAtual();
  const base = String(u?.username || u?.id || "sem-login").trim().toLowerCase();
  return base.replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40) || "sem-login";
}

const escAnexo = (t) =>
  String(t ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function tamanhoLegivel(bytes) {
  if (!(bytes >= 0)) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

async function anexosFetch(caminho, opcoes = {}) {
  const sessao = sessaoAtual();
  if (!sessao) throw new Error("sem sessão");
  const headers = Object.assign(
    { apikey: SB_KEY, Authorization: "Bearer " + sessao.accessToken },
    opcoes.headers || {}
  );
  return fetch(SB_URL + caminho, Object.assign({}, opcoes, { headers }));
}

async function listarAnexos(lote, sn) {
  const q = `lote=eq.${encodeURIComponent(lote)}&sn=eq.${encodeURIComponent(sn)}`;
  const r = await anexosFetch(
    `/rest/v1/anexos_conferencia?${q}&select=id,nome,tipo,tamanho,caminho,criado_em&order=criado_em.asc`
  );
  if (!r.ok) throw new Error("não consegui listar os anexos");
  return r.json();
}

async function subirAnexo(lote, sn, arquivo) {
  const nomeSeguro = arquivo.name.replace(/[^A-Za-z0-9._-]+/g, "_");
  const caminho = `${lote}/${sn}/${Date.now()}-${nomeSeguro}`;
  const up = await anexosFetch(`/storage/v1/object/anexos-conferencia/${caminho}`, {
    method: "POST",
    headers: { "Content-Type": arquivo.type || "application/octet-stream" },
    body: arquivo,
  });
  if (!up.ok) throw new Error("não consegui enviar o arquivo");
  const ins = await anexosFetch(`/rest/v1/anexos_conferencia`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({
      lote, sn, nome: arquivo.name, tipo: arquivo.type, tamanho: arquivo.size, caminho,
    }),
  });
  if (!ins.ok) throw new Error("o arquivo foi enviado, mas não consegui salvar o registro dele");
  return ins.json();
}

async function removerAnexo(anexo) {
  await anexosFetch(`/storage/v1/object/anexos-conferencia/${anexo.caminho}`, { method: "DELETE" });
  const del = await anexosFetch(`/rest/v1/anexos_conferencia?id=eq.${anexo.id}`, { method: "DELETE" });
  if (!del.ok) throw new Error("não consegui remover o anexo");
}

async function baixarAnexo(anexo) {
  const r = await anexosFetch(`/storage/v1/object/anexos-conferencia/${anexo.caminho}`);
  if (!r.ok) throw new Error("não consegui baixar o anexo");
  baixar(await r.blob(), anexo.nome);
}

function anexosAviso(texto, erro) {
  const el = $("anexos-status");
  if (!el) return;
  el.textContent = texto || "";
  el.classList.toggle("erro-texto", !!erro);
}

function renderAnexosLista(lista) {
  const caixa = $("anexos-lista");
  caixa.innerHTML = lista.length
    ? lista.map((a) => `
        <div class="anexo-item" data-id="${a.id}">
          <span class="anexo-item__nome" title="${escAnexo(a.nome)}">${escAnexo(a.nome)}</span>
          <span class="anexo-item__tam sub">${tamanhoLegivel(a.tamanho)}</span>
          <button type="button" class="linkbtn anexo-item__baixar">Abrir</button>
          <button type="button" class="anexo-item__remover" title="Remover anexo">✕</button>
        </div>`).join("")
    : `<div class="sub">Nenhum arquivo anexado ainda.</div>`;

  caixa.querySelectorAll(".anexo-item__baixar").forEach((btn) => {
    btn.onclick = () => {
      const id = btn.closest(".anexo-item").dataset.id;
      const a = lista.find((x) => String(x.id) === id);
      if (!a) return;
      btn.disabled = true;
      baixarAnexo(a)
        .catch((e) => anexosAviso(e.message, true))
        .finally(() => { btn.disabled = false; });
    };
  });
  caixa.querySelectorAll(".anexo-item__remover").forEach((btn) => {
    btn.onclick = () => {
      const id = btn.closest(".anexo-item").dataset.id;
      const a = lista.find((x) => String(x.id) === id);
      if (!a) return;
      btn.disabled = true;
      removerAnexo(a)
        .then(() => carregarAnexos())
        .catch((e) => { anexosAviso(e.message, true); btn.disabled = false; });
    };
  });
}

/** Recarrega a lista de anexos da solicitação atualmente aberta na revisão. */
async function carregarAnexos() {
  const s = estado.itens[estado.i];
  if (!s) return;
  const sessao = sessaoAtual();
  const semLogin = $("anexos-sem-login"), area = $("anexos-area");
  if (semLogin) semLogin.classList.toggle("oculto", !!sessao);
  if (area) area.classList.toggle("oculto", !sessao);
  if (!sessao) return;

  const lote = chaveLoteBase(estado.dados);
  const minhaChave = `${lote}::${s.sn}`;
  anexosEstado.chave = minhaChave;
  $("anexos-lista").innerHTML = `<div class="sub">Carregando anexos…</div>`;
  anexosAviso("", false);
  try {
    const lista = await listarAnexos(lote, s.sn);
    if (anexosEstado.chave !== minhaChave) return; // já navegou para outra solicitação
    anexosEstado.lista = lista;
    renderAnexosLista(lista);
  } catch {
    if (anexosEstado.chave !== minhaChave) return;
    $("anexos-lista").innerHTML = `<div class="sub">Não consegui carregar os anexos agora.</div>`;
  }
}

function ligarAnexos() {
  $("anexos-btn").onclick = () => $("anexos-input").click();
  $("anexos-input").onchange = async () => {
    const arquivo = $("anexos-input").files[0];
    $("anexos-input").value = "";
    if (!arquivo) return;
    const s = estado.itens[estado.i];
    const lote = chaveLoteBase(estado.dados);
    const minhaChave = `${lote}::${s.sn}`;
    $("anexos-btn").disabled = true;
    anexosAviso(`Enviando "${arquivo.name}"…`, false);
    try {
      await subirAnexo(lote, s.sn, arquivo);
      if (anexosEstado.chave === minhaChave) { anexosAviso("", false); await carregarAnexos(); }
    } catch (e) {
      anexosAviso(`Não consegui anexar: ${e.message}`, true);
    } finally {
      $("anexos-btn").disabled = false;
    }
  };
}

/** Mostra um erro no lugar do bloco de pendência, sem alert() nativo. */
function alertaNoResumo(texto) {
  $("res-pendencia").innerHTML = `<div class="alerta erro">${texto}</div>`;
}

const statusDe = (sn) => estado.pareceres[sn]?.status || "";
const statusNaPlanilha = (sn) => statusDe(sn) || "Sem conferir";
const parecerDe = (sn) => estado.pareceres[sn]?.parecer || "";

/** Largura de coluna pelo maior conteúdo, com piso e teto. */
function largura(valores, minimo, maximo) {
  const maior = valores.reduce((m, v) => Math.max(m, String(v ?? "").length), 0);
  return Math.min(Math.max(minimo, maior + 2), maximo);
}

function estilizarCabecalho(linha) {
  linha.height = 24;
  linha.eachCell((c) => {
    c.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
    c.alignment = { vertical: "middle", horizontal: "left" };
    c.border = { bottom: { style: "thin", color: { argb: AZUL } } };
  });
}

/** Aba de solicitações. `comTipo` inclui a coluna Tipo (usada só na aba TODAS). */
function abaSolicitacoes(wb, nome, itens, comTipo) {
  const ws = wb.addWorksheet(nome, {
    views: [{ showGridLines: false, state: "frozen", ySplit: 1 }],
  });
  const colunas = [
    { header: "S.N", key: "sn" },
    ...(comTipo ? [{ header: "Tipo", key: "tipo" }] : []),
    { header: "Solicitante", key: "solicitante" },
    { header: "Competente", key: "competente" },
    { header: "Favorecido", key: "favorecido" },
    { header: "Destinação", key: "destinacao" },
    { header: "Valor (R$)", key: "valor" },
    { header: "Status", key: "status" },
    { header: "Parecer", key: "parecer" },
    { header: "Apontamentos", key: "apontamentos" },
  ];
  ws.columns = colunas;

  for (const s of itens) {
    ws.addRow({
      sn: s.sn, tipo: s.tipo, solicitante: s.solicitante, competente: s.competente,
      valor: s.valor, favorecido: s.favorecido, destinacao: s.destinacao,
      apontamentos: apontamentosDe(s),
      status: statusNaPlanilha(s.sn), parecer: parecerDe(s.sn),
    });
  }

  // largura pelo conteúdo; as colunas longas quebram linha e o Excel ajusta a altura
  const w = (k, min, max) => largura(itens.map((s) => ({
    sn: s.sn, tipo: s.tipo, solicitante: s.solicitante, competente: s.competente,
    valor: moeda(s.valor), favorecido: s.favorecido,
    destinacao: s.destinacao, apontamentos: apontamentosDe(s),
    status: statusNaPlanilha(s.sn), parecer: parecerDe(s.sn),
  }[k])).concat(colunas.find((c) => c.key === k).header), min, max);

  ws.getColumn("sn").width = w("sn", 10, 14);
  if (comTipo) ws.getColumn("tipo").width = w("tipo", 14, 22);
  ws.getColumn("solicitante").width = w("solicitante", 16, 28);
  ws.getColumn("competente").width = w("competente", 16, 28);
  ws.getColumn("favorecido").width = w("favorecido", 22, 42);
  ws.getColumn("destinacao").width = 58;
  ws.getColumn("valor").width = 14;
  ws.getColumn("apontamentos").width = 38;
  ws.getColumn("status").width = 27;
  ws.getColumn("parecer").width = 46;

  estilizarCabecalho(ws.getRow(1));

  ws.eachRow((linha, n) => {
    if (n === 1) return;
    linha.eachCell((c) => {
      c.font = { name: "Calibri", size: 11 };
      c.alignment = { vertical: "top", wrapText: false };
      c.border = { bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
    });
    linha.getCell("valor").numFmt = "#,##0.00";
    linha.getCell("valor").alignment = { vertical: "top", horizontal: "right" };
    for (const k of ["solicitante", "competente", "favorecido", "destinacao", "apontamentos", "parecer"]) {
      linha.getCell(k).alignment = { vertical: "top", wrapText: true };
    }
    // o apontamento é do sistema, não do conferente: fica marcado, não escondido
    const ap = linha.getCell("apontamentos");
    if (ap.value) ap.font = { name: "Calibri", size: 11, bold: true, color: { argb: AMBAR } };
    const st = linha.getCell("status");
    st.font = { name: "Calibri", size: 11, bold: !!COR_STATUS[st.value],
                color: { argb: COR_STATUS[st.value] || "FF808080" } };
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colunas.length } };
  return ws;
}

function abaResumo(wb) {
  const ws = wb.addWorksheet("RESUMO", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 30 }, { width: 10 }, { width: 18 }, { width: 12 },
                { width: 28 }, { width: 12 }, { width: 14 }];

  const titulo = (linha, texto, tamanho) => {
    const c = ws.getCell(`A${linha}`);
    c.value = texto;
    c.font = { name: "Calibri", size: tamanho, bold: true, color: { argb: AZUL } };
  };
  const cabecalho = (linha, textos) => {
    const l = ws.getRow(linha);
    textos.forEach((t, i) => (l.getCell(i + 1).value = t));
    l.height = 22;
    for (let i = 1; i <= textos.length; i++) {
      const c = l.getCell(i);
      c.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
      c.alignment = { vertical: "middle", horizontal: i === 1 ? "left" : "center", wrapText: true };
    }
  };

  titulo(1, "CONFERÊNCIA DE PAGAMENTOS", 16);
  ws.getCell("A2").value =
    `Relatório de ${textoPeriodo(estado.dados.meta) || "—"} · ${estado.itens.length} solicitações · ` +
    `R$ ${moeda(estado.dados.validacao.valorExtraido)}`;
  ws.getCell("A2").font = { name: "Calibri", size: 11, color: { argb: "FF595959" } };
  ws.getCell("A3").value = estado.dados.meta.empresa || "";
  ws.getCell("A3").font = { name: "Calibri", size: 10, color: { argb: "FF808080" } };

  // quantas solicitações o próprio sistema marcou — evidência de que a
  // verificação automática rodou, e quanto ela achou
  const comApontamento = estado.itens.filter(temApontamento).length;
  ws.getCell("A4").value = comApontamento
    ? `${comApontamento} solicitação(ões) com apontamento automático — ver coluna "Apontamentos"`
    : "Nenhum apontamento automático neste lote";
  ws.getCell("A4").font = comApontamento
    ? { name: "Calibri", size: 11, bold: true, color: { argb: AMBAR } }
    : { name: "Calibri", size: 10, color: { argb: "FF808080" } };

  // ---- por forma de pagamento
  titulo(5, "POR FORMA DE PAGAMENTO", 12);
  const nomes = STATUS.map((s) => s.valor);
  cabecalho(6, ["Forma de pagamento", "Qtde", "Valor (R$)", ...nomes, "Sem conferir"]);

  let linha = 7;
  const tipos = ORDEM.filter((t) => estado.itens.some((s) => s.tipo === t));
  for (const t of tipos) {
    const doTipo = estado.itens.filter((s) => s.tipo === t);
    const conta = (v) => doTipo.filter((s) => (statusDe(s.sn) || "Sem conferir") === v).length;
    const l = ws.getRow(linha);
    l.getCell(1).value = t;
    l.getCell(2).value = doTipo.length;
    l.getCell(3).value = doTipo.reduce((a, s) => a + (s.valor || 0), 0);
    nomes.forEach((n, i) => (l.getCell(4 + i).value = conta(n)));
    l.getCell(4 + nomes.length).value = conta("Sem conferir");
    linha++;
  }

  const total = ws.getRow(linha);
  total.getCell(1).value = "TOTAL";
  total.getCell(2).value = { formula: `SUM(B7:B${linha - 1})` };
  total.getCell(3).value = { formula: `SUM(C7:C${linha - 1})` };
  for (let i = 0; i <= nomes.length; i++) {
    const col = String.fromCharCode(68 + i); // D em diante
    total.getCell(4 + i).value = { formula: `SUM(${col}7:${col}${linha - 1})` };
  }
  const fimTipos = linha;

  // ---- por status
  const inicioStatus = linha + 3;
  titulo(inicioStatus - 1, "POR STATUS", 12);
  cabecalho(inicioStatus, ["Status", "Qtde", "Valor (R$)", "% do valor"]);
  linha = inicioStatus + 1;
  const valorTotal = estado.itens.reduce((a, s) => a + (s.valor || 0), 0) || 1;
  for (const nome of [...nomes, "Sem conferir"]) {
    const doStatus = estado.itens.filter((s) => (statusDe(s.sn) || "Sem conferir") === nome);
    if (!doStatus.length) continue;
    const soma = doStatus.reduce((a, s) => a + (s.valor || 0), 0);
    const l = ws.getRow(linha);
    l.getCell(1).value = nome;
    l.getCell(2).value = doStatus.length;
    l.getCell(3).value = soma;
    l.getCell(4).value = soma / valorTotal;
    if (COR_STATUS[nome]) {
      l.getCell(1).font = { name: "Calibri", size: 11, bold: true, color: { argb: COR_STATUS[nome] } };
    }
    linha++;
  }

  // formatação das duas tabelas
  for (let n = 7; n < linha; n++) {
    const l = ws.getRow(n);
    l.eachCell((c, i) => {
      if (!c.font) c.font = { name: "Calibri", size: 11 };
      c.border = { bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
      if (i >= 2) c.alignment = { horizontal: "center" };
      if (i === 3) { c.numFmt = "#,##0.00"; c.alignment = { horizontal: "right" }; }
    });
    if (n >= inicioStatus + 1) l.getCell(4).numFmt = "0.0%";
  }
  const lt = ws.getRow(fimTipos);
  lt.eachCell((c) => {
    c.font = { name: "Calibri", size: 11, bold: true };
    c.border = { top: { style: "thin", color: { argb: AZUL } } };
  });
  lt.getCell(3).numFmt = "#,##0.00";

  return ws;
}

async function gerarPlanilha() {
  const btn = $("btn-planilha");
  const rotulo = btn.textContent;
  btn.disabled = true; btn.textContent = "Gerando…";
  try {
    await carregarExcelJS();
    const wb = new ExcelJS.Workbook();
    wb.creator = "NOCTUS — Conferência de Pagamentos";
    wb.created = new Date();

    abaResumo(wb);
    abaSolicitacoes(wb, "TODAS", estado.itens, true);
    for (const tipo of ORDEM) {
      const linhas = estado.itens.filter((s) => s.tipo === tipo);
      if (linhas.length) abaSolicitacoes(wb, tipo.slice(0, 31), linhas, false);
    }

    const buffer = await wb.xlsx.writeBuffer();
    const ref = textoPeriodo(estado.dados.meta).replace(/\//g, "-").replace(/ /g, "_");
    baixar(new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }), `Conferencia_Pagamentos_${ref}.xlsx`);
  } catch (e) {
    alertaNoResumo(`Não consegui gerar a planilha: ${e.message}`);
  } finally {
    btn.disabled = false; btn.textContent = rotulo;
  }
}

/* ------------------------------------------------- planilha de todas as conferências */

/**
 * Lê do navegador todas as conferências guardadas, já achatadas em linhas.
 *
 * A planilha do dia responde "o que eu conferi neste relatório". Esta responde
 * "o que passou por mim no período" — é ela que permite cruzar fornecedor,
 * valor e nota entre dias diferentes, coisa que a conferência de um lote só,
 * por definição, não enxerga.
 */
function consolidarLotes() {
  const lotes = [];
  const prefixo = `${CHAVE}.${escopoUsuario()}.`;
  for (let k = 0; k < localStorage.length; k++) {
    const chave = localStorage.key(k);
    if (!chave || !chave.startsWith(prefixo)) continue;
    try {
      const v = JSON.parse(localStorage.getItem(chave));
      if (!Array.isArray(v.solicitacoes)) continue;
      lotes.push({
        data: v.meta?.dataInicio || "sem data",
        empresa: nomeEmpresa(v.meta),
        solicitacoes: v.solicitacoes,
        pareceres: v.pareceres || {},
      });
    } catch (e) { /* entrada corrompida: fica de fora */ }
  }
  const ord = (d) => (d || "").split("/").reverse().join("-");
  lotes.sort((a, b) => ord(a.data).localeCompare(ord(b.data)) ||
                       a.empresa.localeCompare(b.empresa));

  const linhas = [];
  for (const l of lotes) {
    for (const s of ordenar(l.solicitacoes)) {
      const p = l.pareceres[s.sn] || {};
      // Lote de PDF: uma empresa só, l.empresa já resolve. Lote de Excel
      // (pode ter várias empresas juntas — ver parseRelatorioExcel): cada
      // solicitação carrega a SUA PRÓPRIA empresa (s.empresa), identificada
      // pela conta bancária do bloco onde ela apareceu no relatório — usa
      // essa quando existir, cai para l.empresa (PDF, ou Excel de 1 empresa
      // só) quando não.
      const empresaLinha = s.empresa || l.empresa;
      // achado em 23/09/2026 (bug do "Imprimir" do dia, ver abaResumoGeral):
      // desde que o Excel multiempresa virou uma conferência por empresa
      // (dividirPorEmpresa, em cima), l.empresa é o rótulo "código - NOME
      // OFICIAL" (rotuloEmpresa) mas s.empresa aqui embaixo é o nome curto
      // interno ("Praia Verde", não "13 - PRAIA VERDE...") — os dois
      // deixaram de ser a mesma string. plano.porEmpresa (montado a partir
      // de linhas[].empresa, ou seja, do nome CURTO) precisa dessa mesma
      // chave curta pra achar a aba do lote, não do rótulo bonito — por
      // isso guarda a chave aqui, separada do rótulo que continua em
      // l.empresa pra exibição.
      if (l.chaveEmpresa === undefined) l.chaveEmpresa = empresaLinha;
      linhas.push({
        data: l.data, empresa: empresaLinha, sn: s.sn, tipo: s.tipo,
        solicitante: s.solicitante, competente: s.competente, valor: s.valor,
        favorecido: s.favorecido, destinacao: s.destinacao,
        apontamentos: apontamentosDe(s),
        status: p.status || "Sem conferir", parecer: p.parecer || "",
      });
    }
  }
  return { lotes, linhas };
}

/* ------------------------------------------------- abas por empresa ("Imprimir" do dia)
 * PARÂMETROS dos nomes das abas. A ORDEM desta lista é a ordem das abas na planilha.
 * `chave` = palavra inteira (sem acento, maiúscula) procurada no nome da empresa do
 * relatório. Empresa que não casar com nenhuma chave ganha a aba "EMP <código>".
 * Só o botão "Imprimir" do dia usa isto; o consolidado geral segue como era. */
const ABAS_EMPRESA = [
  { nome: "Momentum",    chave: "MOMENTUM",    aba: "MM" },
  { nome: "RVM",         chave: "RVM",         aba: "RVM" },
  { nome: "Pick Money",  chave: "PICK MONEY",  aba: "PKM" },
  { nome: "Slim",        chave: "SLIM",        aba: "SLIM" },
  { nome: "Posto",       chave: "POSTO",       aba: "Posto" },
  { nome: "Kasil",       chave: "KASIL",       aba: "Kasil" },
  { nome: "Instituto",   chave: "INSTITUTO",   aba: "IRM" },
  { nome: "Realiza",     chave: "REALIZA",     aba: "REALIZA" },
  { nome: "Abrasma",     chave: "ABRASMA",     aba: "ABRASMA" },
  { nome: "MMH",         chave: "MMH",         aba: "MMH" },
  { nome: "M5",          chave: "M5",          aba: "M5" },
  { nome: "M3",          chave: "M3",          aba: "M3" },
  { nome: "Praia Verde", chave: "PRAIA VERDE", aba: "PV" },
];

const semAcento = (t) => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
const letraColuna = (n) => { let s = ""; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
// Link interno = fórmula HYPERLINK (salto padrão do Excel). Aspas simples: "M3" parece célula.
// String(...) nos dois: acontece de menos vir undefined aqui (achado em
// 23/09/2026 — ver a nota em consolidarLotes) e sem isso o HYPERLINK
// quebra a planilha inteira em vez de só ficar com uma célula estranha.
const linkAba = (aba, rotulo = aba) => ({
  formula: `HYPERLINK("#'${String(aba).replace(/'/g, "''")}'!A1","${String(rotulo).replace(/"/g, '""')}")`,
  result: String(rotulo),
});

function nomeAbaEmpresa(empresa) {
  const txt = semAcento(empresa);
  const achou = ABAS_EMPRESA.find((e) => new RegExp(`(^|[^A-Z0-9])${e.chave}([^A-Z0-9]|$)`).test(txt));
  if (achou) return achou.aba;
  const cod = (String(empresa).match(/^\s*(\d+)/) || [])[1];
  return cod ? `EMP ${cod}` : String(empresa).replace(/[\\/?*[\]:]/g, " ").slice(0, 31);
}

/** Decide a aba de cada empresa e a ordem das abas (ordem de ABAS_EMPRESA; sem parâmetro, no fim). */
function planoAbasEmpresa(linhas) {
  const porEmpresa = new Map();                       // nome completo da empresa -> nome da aba
  for (const emp of new Set(linhas.map((l) => l.empresa))) porEmpresa.set(emp, nomeAbaEmpresa(emp));
  const pos = (a) => { const i = ABAS_EMPRESA.findIndex((e) => e.aba === a); return i < 0 ? 999 : i; };
  const abas = [...new Set(porEmpresa.values())].sort((a, b) => pos(a) - pos(b));
  return { porEmpresa, abas };
}

/** Linha 1 das abas de dados: "← RESUMO | TODAS | MM | PKM | …" (aba atual em destaque, sem link). */
function barraNavegacao(ws, plano, atual, colunas) {
  const itens = [["← RESUMO", "RESUMO GERAL"], ["TODAS", "TODAS"], ...plano.abas.map((a) => [a, a])];
  const linha = ws.getRow(1);
  linha.height = 20;
  for (let i = 1; i <= Math.max(colunas, itens.length); i++) {
    linha.getCell(i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
  }
  itens.forEach(([rotulo, alvo], i) => {
    const c = linha.getCell(i + 1);
    c.alignment = { vertical: "middle", horizontal: "left" };
    if (alvo === atual) {
      c.value = rotulo;
      c.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
    } else {
      c.value = linkAba(alvo, rotulo);
      c.font = { name: "Calibri", size: 10, bold: true, underline: true, color: { argb: "FF1D4ED8" } };
    }
  });
}

const COLUNAS_GERAL = [
  { header: "Data", key: "data", width: 12 },
  { header: "Empresa", key: "empresa", width: 34 },
  { header: "S.N", key: "sn", width: 12 },
  { header: "Forma de pagamento", key: "tipo", width: 20 },
  { header: "Solicitante", key: "solicitante", width: 22 },
  { header: "Competente", key: "competente", width: 22 },
  { header: "Favorecido", key: "favorecido", width: 38 },
  { header: "Destinação", key: "destinacao", width: 54 },
  { header: "Valor (R$)", key: "valor", width: 14 },
  { header: "Status", key: "status", width: 24 },
  { header: "Parecer", key: "parecer", width: 44 },
  { header: "Apontamentos", key: "apontamentos", width: 38 },
];

/** Aba de dados. Com `plano` (Imprimir do dia): barra de links na linha 1, cabeçalho na 2,
 *  filtro cobrindo todas as linhas e área de impressão sem a barra. */
function abaGeral(wb, nome, linhas, plano = null) {
  const topo = plano ? 2 : 1;
  const ws = wb.addWorksheet(nome, {
    views: [{ showGridLines: false, state: "frozen", ySplit: topo }],
  });
  if (plano) {
    ws.columns = COLUNAS_GERAL.map(({ key, width }) => ({ key, width }));
    barraNavegacao(ws, plano, nome, COLUNAS_GERAL.length);
    ws.addRow(Object.fromEntries(COLUNAS_GERAL.map(({ key, header }) => [key, header])));
  } else {
    ws.columns = COLUNAS_GERAL.map(({ header, key, width }) => ({ header, key, width }));
  }
  for (const l of linhas) ws.addRow(l);
  estilizarCabecalho(ws.getRow(topo));

  ws.eachRow((linha, n) => {
    if (n <= topo) return;
    linha.eachCell((c) => {
      c.font = { name: "Calibri", size: 11 };
      c.alignment = { vertical: "top", wrapText: false };
      c.border = { bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
    });
    linha.getCell("valor").numFmt = "#,##0.00";
    linha.getCell("valor").alignment = { vertical: "top", horizontal: "right" };
    for (const k of ["empresa", "solicitante", "competente", "favorecido", "destinacao", "apontamentos", "parecer"]) {
      linha.getCell(k).alignment = { vertical: "top", wrapText: true };
    }
    const ap = linha.getCell("apontamentos");
    if (ap.value) ap.font = { name: "Calibri", size: 11, bold: true, color: { argb: AMBAR } };
    const st = linha.getCell("status");
    st.font = { name: "Calibri", size: 11, bold: !!COR_STATUS[st.value],
                color: { argb: COR_STATUS[st.value] || "FF808080" } };
  });

  const ultima = plano ? topo + linhas.length : topo;
  ws.autoFilter = { from: { row: topo, column: 1 }, to: { row: ultima, column: COLUNAS_GERAL.length } };
  if (plano) ws.pageSetup.printArea = `A${topo}:${letraColuna(COLUNAS_GERAL.length)}${ultima}`;
  return ws;
}

/** Resumo. Com `plano` (Imprimir do dia): ganha a coluna "Aba" (link p/ a aba da empresa) à
 *  esquerda de Empresa, e uma linha oculta com o atalho da aba PARAMETROS. */
function abaResumoGeral(wb, lotes, linhas, plano = null) {
  const o = plano ? 1 : 0;                       // deslocamento das colunas p/ a direita
  const ws = wb.addWorksheet("RESUMO GERAL", { views: [{ showGridLines: false }] });
  ws.columns = (o ? [12, 10, 40, 10, 16, 12, 14, 16] : [12, 40, 10, 16, 12, 14, 16])
    .map((width) => ({ width }));

  const titulo = (l, texto, tamanho) => {
    const c = ws.getCell(`A${l}`);
    c.value = texto;
    c.font = { name: "Calibri", size: tamanho, bold: true, color: { argb: AZUL } };
  };
  const cabecalho = (l, textos) => {
    const linha = ws.getRow(l);
    textos.forEach((t, i) => (linha.getCell(i + 1).value = t));
    linha.height = 22;
    for (let i = 1; i <= textos.length; i++) {
      const c = linha.getCell(i);
      c.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
      c.alignment = { vertical: "middle", horizontal: i <= 2 + o ? "left" : "center", wrapText: true };
    }
  };
  const fonteLink = { name: "Calibri", size: 11, bold: true, underline: true, color: { argb: "FF1D4ED8" } };

  const datas = [...new Set(lotes.map((l) => l.data))];
  const valorTotal = linhas.reduce((a, l) => a + (l.valor || 0), 0);
  const apontadas = linhas.filter((l) => l.apontamentos).length;
  const semConferir = linhas.filter((l) => l.status === "Sem conferir").length;

  titulo(1, "CONFERÊNCIA DE PAGAMENTOS — CONSOLIDADO", 16);
  ws.getCell("A2").value =
    `${lotes.length} conferência(s) · ${datas.length} dia(s) · ${linhas.length} solicitações · ` +
    `R$ ${moeda(valorTotal)}`;
  ws.getCell("A2").font = { name: "Calibri", size: 11, color: { argb: "FF595959" } };
  const periodo = datas.length > 1 ? `${datas[0]} a ${datas[datas.length - 1]}` : datas[0];
  ws.getCell("A3").value = datas.length
    ? `Período: ${periodo} · gerado em ${new Date().toLocaleString("pt-BR")}`
    : "";
  ws.getCell("A3").font = { name: "Calibri", size: 10, color: { argb: "FF808080" } };
  const ondeApontamentos = plano ? "ver coluna Apontamentos" : "ver aba APONTAMENTOS";
  ws.getCell("A4").value = apontadas
    ? `${apontadas} solicitação(ões) com apontamento automático — ${ondeApontamentos}`
    : "Nenhum apontamento automático no período";
  ws.getCell("A4").font = apontadas
    ? { name: "Calibri", size: 11, bold: true, color: { argb: AMBAR } }
    : { name: "Calibri", size: 10, color: { argb: "FF808080" } };
  ws.getCell("A5").value = semConferir
    ? `Atenção: ${semConferir} solicitação(ões) ainda sem parecer neste consolidado`
    : "Todas as solicitações do período estão conferidas";
  ws.getCell("A5").font = semConferir
    ? { name: "Calibri", size: 11, bold: true, color: { argb: "FFB91C1C" } }
    : { name: "Calibri", size: 10, color: { argb: "FF15803D" } };

  titulo(7, "POR DIA E EMPRESA", 12);
  cabecalho(8, o
    ? ["Data", "Aba", "Empresa", "Qtde", "Valor (R$)", "Conferidas", "Sem conferir", "Apontamentos"]
    : ["Data", "Empresa", "Qtde", "Valor (R$)", "Conferidas", "Sem conferir", "Apontamentos"]);
  // Colunas C..A(mapeando COLUNAS_GERAL): C = S.N, I = Valor, J = Status,
  // L = Apontamentos — usadas nas fórmulas abaixo. 100000 é só um limite
  // generoso de linhas (nenhuma aba real chega perto disso), pra não
  // precisar saber o tamanho exato da aba na hora de montar a fórmula
  // (as abas de dados só são escritas DEPOIS desta, ver imprimirDia/
  // gerarPlanilhaGeral — mas isso não afeta a fórmula, o Excel só resolve
  // a referência quando o arquivo é aberto).
  const escaparFormula = (t) => String(t).replace(/"/g, '""');
  // Primeira linha de dado na aba "TODAS" — ela sempre existe (ver
  // gerarPlanilhaGeral/imprimirDia, que escrevem "TODAS" nos dois casos),
  // então é usada tanto no fallback de empresa sem aba própria abaixo
  // quanto no quadro "POR STATUS NO PERÍODO", que nunca tem aba por
  // empresa e por isso sempre busca em "TODAS".
  const deTodas = (plano ? 2 : 1) + 1;
  let l = 9;
  for (const lote of lotes) {
    const chave = lote.chaveEmpresa ?? lote.empresa;
    const linha = ws.getRow(l);
    linha.getCell(1).value = lote.data;
    let aba = null;
    if (o) {
      aba = plano.porEmpresa.get(chave);
      if (aba) {
        linha.getCell(2).value = linkAba(aba);
        linha.getCell(2).font = fonteLink;
      } else {
        // não deveria acontecer (toda linha de "linhas" tem uma aba em
        // plano.porEmpresa — ver planoAbasEmpresa), mas se acontecer é
        // melhor a planilha sair sem o link do que não sair nenhuma.
        linha.getCell(2).value = "—";
      }
    }
    linha.getCell(2 + o).value = lote.empresa;

    // Qtde/Valor/Conferidas/Sem conferir/Apontamentos por FÓRMULA, buscando
    // na aba de onde os dados realmente vêm — pedido do Rocha em 23/09/2026,
    // depois do "Imprimir" do dia sair com esses números todos zerados. A
    // causa foi lote.empresa não bater mais com o texto de linha.empresa
    // (ver a nota em consolidarLotes, lá em cima) — e como o valor era só um
    // número já calculado aqui, o zero ficava indistinguível de "essa
    // empresa realmente não teve solicitação nenhuma nesse dia", sem jeito
    // de notar o problema sem abrir a aba e contar na mão. Com fórmula, o
    // número sempre reflete o que está de fato na aba de origem — e se essa
    // mesma inconsistência acontecer de novo, dá pra clicar na célula,
    // ver a fórmula e a aba que ela aponta, em vez de confiar cegamente
    // num valor que o site já entregou pronto.
    if (aba) {
      // Imprimir do dia: cada empresa tem sua própria aba (ver abaGeral) —
      // a fórmula busca direto nela, sem precisar filtrar por empresa (a
      // aba inteira já é só daquela empresa).
      const R = `'${aba}'!`;
      linha.getCell(3 + o).value = { formula: `COUNTA(${R}C3:C100000)` };
      linha.getCell(4 + o).value = { formula: `SUM(${R}I3:I100000)` };
      linha.getCell(5 + o).value = { formula: `COUNTIFS(${R}C3:C100000,"<>",${R}J3:J100000,"<>Sem conferir")` };
      linha.getCell(6 + o).value = { formula: `COUNTIFS(${R}C3:C100000,"<>",${R}J3:J100000,"Sem conferir")` };
      linha.getCell(7 + o).value = { formula: `COUNTIFS(${R}C3:C100000,"<>",${R}L3:L100000,"<>")` };
    } else {
      // Planilha consolidada geral (gerarPlanilhaGeral): não tem uma aba por
      // empresa, só "TODAS" com todo mundo junto — busca ali, filtrando
      // pela empresa dessa linha da tabela.
      const crit = `"${escaparFormula(chave)}"`;
      linha.getCell(3 + o).value = { formula: `COUNTIFS(TODAS!$B$${deTodas}:$B$100000,${crit})` };
      linha.getCell(4 + o).value = { formula: `SUMIFS(TODAS!$I$${deTodas}:$I$100000,TODAS!$B$${deTodas}:$B$100000,${crit})` };
      linha.getCell(5 + o).value = { formula:
        `COUNTIFS(TODAS!$B$${deTodas}:$B$100000,${crit},TODAS!$J$${deTodas}:$J$100000,"<>Sem conferir")` };
      linha.getCell(6 + o).value = { formula:
        `COUNTIFS(TODAS!$B$${deTodas}:$B$100000,${crit},TODAS!$J$${deTodas}:$J$100000,"Sem conferir")` };
      linha.getCell(7 + o).value = { formula:
        `COUNTIFS(TODAS!$B$${deTodas}:$B$100000,${crit},TODAS!$L$${deTodas}:$L$100000,"<>")` };
    }
    l++;
  }
  const total = ws.getRow(l);
  total.getCell(1).value = "TOTAL";
  for (const n of [3, 4, 5, 6, 7]) {
    const col = letraColuna(n + o);
    total.getCell(col).value = { formula: `SUM(${col}9:${col}${l - 1})` };
  }
  const fimLotes = l;

  const inicioStatus = l + 3;
  titulo(inicioStatus - 1, "POR STATUS NO PERÍODO", 12);
  cabecalho(inicioStatus, o
    ? ["Status", "", "", "Qtde", "Valor (R$)", "% do valor"]
    : ["Status", "", "Qtde", "Valor (R$)", "% do valor"]);
  l = inicioStatus + 1;
  // Mesmo pedido do Rocha em 23/09/2026 aplicado aqui: Qtde/Valor por
  // fórmula em vez de número já calculado. Diferente do quadro "POR DIA E
  // EMPRESA", este quadro não é por empresa — é o total do período inteiro
  // por status — então não tem aba própria pra buscar: sempre lê a aba
  // "TODAS", que gerarPlanilhaGeral() e imprimirDia() sempre escrevem (com
  // ou sem abas por empresa). "% do valor" divide pelo próprio TOTAL de
  // Valor (R$) do quadro acima (célula com fórmula também, ver "total"
  // logo abaixo) em vez de reusar valorTotal calculado em JS — assim toda
  // a cadeia, do valor de cada empresa até o percentual aqui embaixo,
  // deriva de fórmula, e o Excel mesmo garante que os dois batem.
  const refTotalValor = `$${letraColuna(4 + o)}$${fimLotes}`;
  for (const nome of [...STATUS.map((x) => x.valor), "Sem conferir"]) {
    const doStatus = linhas.filter((x) => x.status === nome);
    if (!doStatus.length) continue;
    const linha = ws.getRow(l);
    linha.getCell(1).value = nome;
    const critStatus = `"${escaparFormula(nome)}"`;
    linha.getCell(3 + o).value = { formula: `COUNTIF(TODAS!$J$${deTodas}:$J$100000,${critStatus})` };
    linha.getCell(4 + o).value = { formula:
      `SUMIF(TODAS!$J$${deTodas}:$J$100000,${critStatus},TODAS!$I$${deTodas}:$I$100000)` };
    linha.getCell(5 + o).value = { formula: `IFERROR(${letraColuna(4 + o)}${l}/${refTotalValor},0)` };
    if (COR_STATUS[nome]) {
      linha.getCell(1).font = { name: "Calibri", size: 11, bold: true, color: { argb: COR_STATUS[nome] } };
    }
    l++;
  }

  const hair = { bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
  for (let n = 9; n < l; n++) {
    const linha = ws.getRow(n);
    linha.eachCell((c, i) => {
      if (!c.font) c.font = { name: "Calibri", size: 11 };
      c.border = hair;
      if (i >= 3 + o) c.alignment = { horizontal: "center" };
      if (i === 4 + o) { c.numFmt = "#,##0.00"; c.alignment = { horizontal: "right" }; }
    });
    if (n >= inicioStatus + 1) {
      linha.getCell(5 + o).numFmt = "0.0%";
      if (o) for (const i of [2, 3]) linha.getCell(i).border = hair;   // linha contínua sob "Status"
    }
  }
  // A linha TOTAL leva borda de cima azul (separando das empresas) — mas o
  // loop acima já tinha desenhado a borda de baixo cinza (fechando a
  // tabela). c.border = {...} SUBSTITUI o objeto inteiro, então sem
  // reaplicar o bottom aqui a borda de baixo do TOTAL desaparecia e a
  // tabela ficava "aberta" na última linha — achado pelo Rocha em
  // 23/09/2026, junto com o pedido de fórmula no quadro de status.
  const bordaTotal = { top: { style: "thin", color: { argb: AZUL } }, bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
  const lt = ws.getRow(fimLotes);
  lt.eachCell((c) => {
    c.font = { name: "Calibri", size: 11, bold: true };
    c.border = bordaTotal;
  });
  if (o) lt.getCell(2).border = bordaTotal;
  lt.getCell(4 + o).numFmt = "#,##0.00";

  if (plano) {
    ws.pageSetup.printArea = `A1:${letraColuna(7 + o)}${l - 1}`;   // o papel não muda
    const rp = ws.getRow(l + 1);                                   // atalho p/ PARAMETROS, oculto
    rp.getCell(1).value = linkAba("PARAMETROS");
    rp.getCell(1).font = fonteLink;
    rp.getCell(2).value = "Padrão de nomes das abas por empresa";
    rp.hidden = true;
  }
  return ws;
}

/** Aba de referência (oculta) com o mapeamento usado. Conta quantas solicitações casam com cada
 *  chave e mostra as que ficaram sem mapeamento (deve ser 0). */
function abaParametros(wb, plano, totalLinhas) {
  const ws = wb.addWorksheet("PARAMETROS", { state: "hidden", views: [{ showGridLines: false }] });
  ws.columns = [{ width: 22 }, { width: 34 }, { width: 16 }, { width: 26 }];
  barraNavegacao(ws, plano, "PARAMETROS", 4);
  ws.getCell("A3").value = "PARÂMETROS — NOMES DAS ABAS POR EMPRESA";
  ws.getCell("A3").font = { name: "Calibri", size: 12, bold: true, color: { argb: AZUL } };
  ws.getCell("A4").value = "Referência do padrão usado para nomear as abas. Empresa sem solicitações no dia não gera aba.";
  ws.getCell("A4").font = { name: "Calibri", size: 10, color: { argb: "FF808080" } };
  const cab = ws.getRow(6);
  ["Empresa", "Palavra-chave no nome da empresa", "Nome da aba", "Solicitações em TODAS"]
    .forEach((t, i) => (cab.getCell(i + 1).value = t));
  cab.height = 32;
  for (let i = 1; i <= 4; i++) {
    const c = cab.getCell(i);
    c.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } };
    c.alignment = { vertical: "middle", horizontal: i <= 2 ? "left" : "center", wrapText: true };
  }
  const hair = { bottom: { style: "hair", color: { argb: CINZA_LINHA } } };
  const fim = 2 + totalLinhas;                      // última linha de dados em TODAS
  ABAS_EMPRESA.forEach((e, i) => {
    const r = 7 + i;
    const linha = ws.getRow(r);
    linha.getCell(1).value = e.nome;
    linha.getCell(2).value = e.chave;
    linha.getCell(3).value = e.aba;
    linha.getCell(4).value = { formula: `COUNTIF(TODAS!$B$3:$B$${fim},"*"&B${r}&"*")` };
    for (let c = 1; c <= 4; c++) {
      const cel = linha.getCell(c);
      cel.font = { name: "Calibri", size: 11 };
      cel.border = hair;
      if (c >= 3) cel.alignment = { horizontal: "center" };
    }
  });
  const ult = 6 + ABAS_EMPRESA.length;
  const tm = ws.getRow(ult + 1), sm = ws.getRow(ult + 2);
  tm.getCell(1).value = "Total mapeado";
  tm.getCell(4).value = { formula: `SUM(D7:D${ult})` };
  sm.getCell(1).value = "Sem mapeamento (deve ser 0)";
  sm.getCell(4).value = { formula: `COUNTA(TODAS!$B$3:$B$${fim})-D${ult + 1}` };
  for (const linha of [tm, sm]) {
    linha.getCell(1).font = { name: "Calibri", size: 11, bold: true };
    linha.getCell(4).font = { name: "Calibri", size: 11, bold: true };
    linha.getCell(4).alignment = { horizontal: "center" };
  }
  ws.addConditionalFormatting({
    ref: `D${ult + 2}`,
    rules: [{ type: "cellIs", operator: "notEqual", formulae: [0],
              style: { font: { bold: true, color: { argb: "FFB91C1C" } } } }],
  });
  return ws;
}

/** Atualiza a linha do rodapé da tela do relatório. */
function renderRodapeGeral() {
  const caixa = $("rodape-upload");
  if (!caixa) return;
  const { lotes } = consolidarLotes();
  caixa.classList.toggle("oculto", !lotes.length);
}

async function gerarPlanilhaGeral() {
  const btn = $("btn-planilha-geral");
  const rotulo = btn.textContent;
  const aviso = (t, erro) => {
    $("upload-msg").innerHTML = `<div class="alerta${erro ? " erro" : ""}">${t}
      <button class="alerta__x" id="btn-fecha-planilha-geral" title="Dispensar">✕</button></div>`;
    $("btn-fecha-planilha-geral").onclick = () => { $("upload-msg").innerHTML = ""; };
  };
  btn.disabled = true; btn.textContent = "Gerando…";
  try {
    const { lotes, linhas } = consolidarLotes();
    if (!linhas.length) throw new Error("não há conferência guardada neste navegador");
    await carregarExcelJS();
    const wb = new ExcelJS.Workbook();
    wb.creator = "NOCTUS — Conferência de Pagamentos";
    wb.created = new Date();

    abaResumoGeral(wb, lotes, linhas);
    abaGeral(wb, "TODAS", linhas);
    const apontadas = linhas.filter((l) => l.apontamentos);
    if (apontadas.length) abaGeral(wb, "APONTAMENTOS", apontadas);

    const buffer = await wb.xlsx.writeBuffer();
    const datas = [...new Set(lotes.map((l) => l.data))].map((d) => d.replace(/\//g, "-"));
    const ref = datas.length > 1 ? `${datas[0]}_a_${datas[datas.length - 1]}` : datas[0];
    baixar(new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }), `Conferencias_Consolidado_${ref}.xlsx`);
    aviso("Planilha consolidada gerada.", false);
  } catch (e) {
    aviso(`Não consegui gerar a planilha consolidada: ${e.message}`, true);
  } finally {
    btn.disabled = false; btn.textContent = rotulo;
  }
}

/** Gera a planilha Excel de um único dia — botão "Imprimir" na linha do dia
 *  da lista de conferências guardadas. Mesmo formato do consolidado geral
 *  (abaResumoGeral + abaGeral), só que filtrado para as empresas daquele dia e com
 *  uma aba por empresa (ver ABAS_EMPRESA). O consolidado geral não muda. */
async function imprimirDia(btn, data) {
  const rotulo = btn.textContent;
  const aviso = (t, erro) => {
    $("upload-msg").innerHTML = `<div class="alerta${erro ? " erro" : ""}">${t}
      <button class="alerta__x" id="btn-fecha-planilha-geral" title="Dispensar">✕</button></div>`;
    $("btn-fecha-planilha-geral").onclick = () => { $("upload-msg").innerHTML = ""; };
  };
  btn.disabled = true; btn.textContent = "Gerando…";
  try {
    const { lotes, linhas } = consolidarLotes();
    const lotesDia = lotes.filter((l) => l.data === data);
    const linhasDia = linhas.filter((l) => l.data === data);
    if (!linhasDia.length) throw new Error(`não há conferência guardada para ${data}`);
    await carregarExcelJS();
    const wb = new ExcelJS.Workbook();
    wb.creator = "NOCTUS — Conferência de Pagamentos";
    wb.created = new Date();

    // Formato do "Imprimir" do dia: uma aba por empresa (nomes em ABAS_EMPRESA), sem a aba
    // APONTAMENTOS (a coluna Apontamentos segue nas abas) e com links entre as abas.
    const plano = planoAbasEmpresa(linhasDia);
    abaResumoGeral(wb, lotesDia, linhasDia, plano);
    abaGeral(wb, "TODAS", linhasDia, plano);
    for (const aba of plano.abas) {
      abaGeral(wb, aba, linhasDia.filter((l) => plano.porEmpresa.get(l.empresa) === aba), plano);
    }
    abaParametros(wb, plano, linhasDia.length).state = "hidden";

    const buffer = await wb.xlsx.writeBuffer();
    baixar(new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }), `Conferencias_${data.replace(/\//g, "-")}.xlsx`);
    aviso(`Planilha de ${data} gerada.`, false);
  } catch (e) {
    aviso(`Não consegui gerar a planilha de ${data}: ${e.message}`, true);
  } finally {
    btn.disabled = false; btn.textContent = rotulo;
  }
}

/* ------------------------------------------------------------------------- início */
migrarChaves();
ligarVoltar();
ligarConfig();
ligarUpload();
$("btn-planilha-geral").onclick = gerarPlanilhaGeral;
ligarRevisao();
ligarAnexos();
ligarResumo();
ligarRenovacaoDeSessao();
if (new URLSearchParams(location.search).get("config") === "1") {
  configAbertaViaHome = true;
  abrirConfig();
}
