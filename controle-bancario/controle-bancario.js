/**
 * Controle Bancário — NOCTUS. Parte 1 (30/10/2026): importação do relatório de
 * fluxo de caixa e tabela das empresas com saídas na data do relatório.
 *
 * Acesso (03/11/2026): enquanto o módulo está em construção, só quem tem perfil de
 * ADMINISTRADOR vê o conteúdo; o perfil "user" continua vendo "Módulo em
 * desenvolvimento". A trava da tela é só aparência: quem garante é o banco (RLS das
 * tabelas controle_bancario_*: administrador, segundo noctus_eh_admin(), e só as
 * próprias linhas de cada um).
 *
 * Dados: cada data importada vira uma linha no servidor (usuário + data);
 * reimportar a mesma data substitui a anterior. Aparece em qualquer computador.
 *
 * Parte 2 (30/10/2026): o mesmo campo aceita o relatório de débito em conta (SCK).
 * O arquivo é reconhecido pelo conteúdo; o total de cada empresa vai para
 * controle_bancario_debito_noctus e a tabela resumo compara Fluxo × Relatório
 * da mesma data (ver debito.js).
 *
 * Exportar (03/11/2026): botão acima do selo de divergências gera o Excel da tabela resumo
 * (exportar.js: a tabela como está na tela, empresas nas colunas, com fórmulas); só com o relatório de débito importado.
 *
 * Parte 3 (03/11/2026, modelo A): lançamentos que justificam a diferença Fluxo × Relatório
 * (Outros Bancos, Desconsiderados, Numerários e Compromissos), com sinal, painel de soma,
 * "Falta justificar" e detalhe por empresa. Exclusão é lógica (ver lancamentos.js e o banco).
 */
import { lerFluxoCaixa, linhasDoExcel } from "./planilha.js";
import { ehRelatorioDebito, lerDebito, montarResumo } from "./debito.js";
import { gerarXlsxResumo, nomeArquivoResumo } from "./exportar.js";
import { CATEGORIAS, NOME_CATEGORIA, contasDaEmpresa, validarLancamento, somarPorCategoria, faltaJustificar,
         porEmpresa, situacaoDaEmpresa } from "./lancamentos.js";

const SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
const CHAVE_SESSAO = "sipep.session";
const TABELA = "controle_bancario_noctus";
const TABELA_DEBITO = "controle_bancario_debito_noctus";
const TABELA_LANC = "controle_bancario_lancamentos_noctus";
const TABELA_VALID = "controle_bancario_validacao_noctus";   // 03/11/2026: checklist opcional de validação por data e empresa
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
/* Editar lançamento (04/11/2026, pedido do Rocha): botão com ícone de lápis ao lado de Excluir. */
const ICONE_LAPIS = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" ' +
  'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20l1-4L16.5 4.5a2.1 2.1 0 013 3L8 19l-4 1z"/><path d="M14 7l3 3"/></svg>';
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
// pedido do Rocha (30/10): no máximo 5 dias por vez (3 no celular, para caber a data com o ano), setas para andar
const ehCelular = () => { try { return window.matchMedia("(max-width:560px)").matches; } catch { return false; } };
const visiveis = () => (ehCelular() ? 3 : 5);
const estado = {
  datas: new Map(),   // iso → { iso, empresas, arquivo, atualizado_em, naoSalvo }
  debitos: new Map(), // iso → { iso, empresas: [{ codigo, empresa, total, qtd, soma, confere }], arquivo, naoSalvo }
  atual: null,        // iso escolhida
  inicioDatas: 0,      // 1ª data visível na faixa de datas (no máximo visiveis() por vez)
  calMes: [2026, 0],   // [ano, mês 0-11] mostrado no calendário
  lancamentos: [],     // Parte 3: lançamentos ativos de todas as datas { id, data_relatorio, empresa_codigo, empresa, categoria, valor, banco, descricao }
  lancFalha: false,    // true quando os lançamentos não puderam ser lidos: o painel não mostra soma que poderia estar errada
  empresasResumo: [],  // empresas da tabela resumo da data aberta (com código), para o formulário de lançamento
  detalhe: null,       // código da empresa cujo detalhe de lançamentos está aberto
  detalheIso: null,    // data a que esse detalhe pertence
  confirmaExcluir: null, // id do lançamento aguardando "Sim, excluir"
  dialogoIso: null,    // data a que o formulário aberto pertence
  salvandoLanc: false,
  editandoId: null,    // id do lançamento em edição (o formulário está no modo "Editar lançamento")
  validacoes: new Map(), // Validação: "iso|chave" → { extrato, comprovantes, suprimento, baixa_compromisso }
  validFalha: false,     // true quando as validações não puderam ser lidas: nada fica verde e as caixas ficam travadas
  validAberta: false,    // a tabela de validação começa oculta (funcionalidade opcional)
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

/** Aviso legível (vai direto para a tela, sem ser tratado como falha técnica). */
function erroAviso(texto) { const e = new Error(texto); e.aviso = true; return e; }
const CAMPOS_LANC = "id,data_relatorio,empresa_codigo,empresa,categoria,valor,banco,descricao,criado_em";

async function carregarLancamentos() {
  const LIMITE = 5000;
  const r = await servidor(`${TABELA_LANC}?select=${CAMPOS_LANC}&excluido=eq.false&order=id.asc&limit=${LIMITE}`);
  const linhas = await r.json();
  // se bateu no limite, há lançamentos que não vieram: melhor recusar do que mostrar soma incompleta
  if (!Array.isArray(linhas) || linhas.length >= LIMITE) throw new Error("lançamentos: leitura incompleta");
  estado.lancamentos = linhas.map((l) => ({ ...l, valor: Number(l.valor), empresa_codigo: String(l.empresa_codigo) }));
}

async function salvarLancamento(lanc, iso) {
  const r = await servidor(`${TABELA_LANC}?select=${CAMPOS_LANC}`, {
    method: "POST",
    headers: { Prefer: "return=representation" },
    body: JSON.stringify([{ data_relatorio: iso, ...lanc }]),
  });
  const linhas = await r.json();
  if (!Array.isArray(linhas) || linhas.length !== 1) throw new Error("lançamento: servidor não devolveu a linha gravada");
  return { ...linhas[0], valor: Number(linhas[0].valor), empresa_codigo: String(linhas[0].empresa_codigo) };
}

/** Exclusão LÓGICA: marca excluido=true. Nada é apagado do banco. */
async function excluirLancamento(id) {
  const r = await servidor(`${TABELA_LANC}?id=eq.${Number(id)}&excluido=eq.false`, {
    method: "PATCH",
    headers: { Prefer: "return=representation" },
    body: JSON.stringify({ excluido: true }),
  });
  const linhas = await r.json();
  if (!Array.isArray(linhas) || linhas.length !== 1) throw new Error("lançamento: exclusão não confirmada pelo servidor");
}

/* ------------------------------------------------------------ Validação (checklist opcional, modelo 1, 03/11/2026) */
const ITENS_VALID = [["extrato", "Extrato"], ["comprovantes", "Comprovantes"], ["suprimento", "Suprimento"], ["baixa_compromisso", "Baixa Compromisso"]];
const VALID_VAZIA = { extrato: false, comprovantes: false, suprimento: false, baixa_compromisso: false };
/** Chave estável da empresa: o código oficial; sem código, o nome. */
const chaveEmpresa = (c) => (c.codigo != null && c.codigo !== "" ? String(c.codigo) : `nome:${c.empresa}`);
const validacaoDe = (iso, chave) => estado.validacoes.get(`${iso}|${chave}`) || null;
const itensMarcados = (v) => (v ? ITENS_VALID.filter(([k]) => v[k] === true).length : 0);
/** Concluída = os 4 itens marcados (e as validações lidas sem falha). */
const validada = (iso, chave) => !estado.validFalha && itensMarcados(validacaoDe(iso, chave)) === ITENS_VALID.length;

/** Lê todas as validações do usuário em páginas (nunca corta em silêncio). */
async function carregarValidacoes() {
  const PAGINA = 1000, MAX_PAGINAS = 50;
  const lidas = [];
  for (let p = 0; p < MAX_PAGINAS; p++) {
    const r = await servidor(`${TABELA_VALID}?select=data_relatorio,empresa_chave,extrato,comprovantes,suprimento,baixa_compromisso` +
      `&order=id.asc&limit=${PAGINA}&offset=${p * PAGINA}`);
    const linhas = await r.json();
    if (!Array.isArray(linhas)) throw new Error("validação: resposta inesperada do servidor");
    lidas.push(...linhas);
    if (linhas.length < PAGINA) {
      estado.validacoes.clear();
      for (const l of lidas) {
        estado.validacoes.set(`${l.data_relatorio}|${l.empresa_chave}`, {
          extrato: l.extrato === true, comprovantes: l.comprovantes === true,
          suprimento: l.suprimento === true, baixa_compromisso: l.baixa_compromisso === true });
      }
      estado.validFalha = false;
      return;
    }
  }
  throw new Error("validação: leitura incompleta");
}

async function gravarValidacao(iso, chave, empresa) {
  const v = validacaoDe(iso, chave);
  if (!v) return;
  const token = await tokenServidor();
  const uid = token && idDoToken(token);
  if (!uid) throw new Error("sem-login");
  const r = await servidor(`${TABELA_VALID}?on_conflict=auth_user_id,data_relatorio,empresa_chave`, {
    method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify([{ auth_user_id: uid, data_relatorio: iso, empresa_chave: chave, empresa: String(empresa).slice(0, 80),
      extrato: v.extrato, comprovantes: v.comprovantes, suprimento: v.suprimento, baixa_compromisso: v.baixa_compromisso }]),
  });
  const linhas = await r.json();
  if (!Array.isArray(linhas) || linhas.length !== 1) throw new Error("validação: servidor não confirmou a gravação");
}

const filaValid = new Map();   // "iso|chave" → promessa: as gravações da mesma empresa saem uma depois da outra
/** Marca/desmarca um item: a tela muda na hora; a gravação vai ao servidor em seguida, em ordem, sempre com o estado mais recente. */
function marcarValidacao(iso, chave, empresa, item, marcado) {
  const k = `${iso}|${chave}`;
  estado.validacoes.set(k, { ...(estado.validacoes.get(k) || VALID_VAZIA), [item]: marcado });
  atualizarValidacao(iso);
  const p = (filaValid.get(k) || Promise.resolve()).then(() => gravarValidacao(iso, chave, empresa)).catch(async (e) => {
    const msg = $("validacao-msg");
    if (msg) msg.insertAdjacentHTML("afterbegin", alerta("erro", esc(msgErro(e, `Não consegui gravar a validação de ${empresa}`))));
    // volta ao que o servidor realmente tem, para a tela não mostrar uma validação que não foi salva
    try { await carregarValidacoes(); } catch { estado.validFalha = true; }
    atualizarValidacao(iso);
  });
  filaValid.set(k, p);
}

/** Pinta de verde (só o texto) o nome, o Fluxo e o Relatório das empresas com os 4 itens marcados. Não mexe na estrutura da tabela. */
function pintarValidadas(iso) {
  for (const el of document.querySelectorAll("#resumo-caixa [data-chave]")) {
    el.classList.toggle("validada", validada(iso, el.dataset.chave));
  }
}

/** Tabela de validação (modelo 1): uma linha por empresa da tabela resumo, os 4 itens em colunas. */
function desenharValidacao(reg, res) {
  const bloco = $("validacao");
  if (!bloco) return;   // defensivo: na publicação o JS vai ao ar antes do HTML
  bloco.classList.remove("oculto");
  const aberta = estado.validAberta;
  const btn = $("btn-validacao");
  btn.setAttribute("aria-expanded", String(aberta));
  $("validacao-cartao").classList.toggle("oculto", !aberta);
  const cols = res.cols;
  const trava = estado.validFalha ? " disabled" : "";
  $("validacao-corpo").innerHTML = cols.map((c) => {
    const chave = chaveEmpresa(c);
    const v = validacaoDe(reg.iso, chave) || VALID_VAZIA;
    const n = itensMarcados(v);
    const caixas = ITENS_VALID.map(([k, rotulo]) =>
      `<td class="valid-cel"><label class="valid-marca"><input type="checkbox" class="valid-cb" data-item="${k}" data-chave="${esc(chave)}" ` +
      `data-empresa="${esc(c.empresa)}" aria-label="${esc(rotulo)} de ${esc(c.empresa)}"${v[k] ? " checked" : ""}${trava}></label></td>`).join("");
    const situacao = estado.validFalha ? "—" : n === ITENS_VALID.length ? "Validado" : `${n} de ${ITENS_VALID.length}`;
    return `<tr data-chave="${esc(chave)}"><th scope="row">${esc(c.empresa)}</th>${caixas}` +
      `<td class="valid-sit${n === ITENS_VALID.length && !estado.validFalha ? " ok" : ""}">${situacao}</td></tr>`;
  }).join("");
  pintarValidadas(reg.iso);
}

/** Atualiza situação, caixas e verde sem redesenhar a tabela resumo inteira (que mede colunas e quebra em faixas). */
function atualizarValidacao(iso) {
  const reg = estado.atual && estado.datas.get(estado.atual);
  if (!reg || reg.iso !== iso) return;
  const deb = estado.debitos.get(reg.iso) || null;
  desenharValidacao(reg, montarResumo(reg.empresas || [], deb));
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
    $("validacao")?.classList.add("oculto");
    $("painel-lanc")?.classList.add("oculto");
    return;
  }
  cartao.classList.remove("oculto");
  if (estado.detalhe != null && estado.detalheIso !== estado.atual) { estado.detalhe = null; estado.confirmaExcluir = null; }

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

/** Exportar para Excel: gera o .xlsx da tabela resumo da data aberta (exportar.js) e baixa no navegador. */
function exportarResumo() {
  const reg = estado.atual && estado.datas.get(estado.atual);
  const deb = reg && estado.debitos.get(reg.iso);
  if (!reg || !deb) return;
  try {
    const res = montarResumo(reg.empresas || [], deb);
    // "NOME · n" no cabeçalho, igual à tela: quantidade de lançamentos de cada empresa na data
    const lancPor = porEmpresa(lancamentosDaData(reg.iso));
    const lanc = res.cols.map((c) => (c.codigo != null ? (lancPor.get(String(c.codigo)) || []).length : 0));
    const bytes = gerarXlsxResumo({ iso: reg.iso, rotulo: rotuloIso(reg.iso), cols: res.cols, total: res.total, avisos: res.avisos, lanc,
      arquivoFluxo: reg.arquivo, arquivoDebito: deb.arquivo });
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = nomeArquivoResumo(reg.iso);
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  } catch (e) {
    $("resumo-aviso").insertAdjacentHTML("afterbegin", alerta("erro", esc(msgErro(e, "Não consegui gerar o Excel da tabela resumo"))));
  }
}

/** Tabela resumo (modelo B): Fluxo × Relatório de débito da data aberta. */
function desenharResumo(reg, colMinima = 0) {
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
  // 03/11/2026: exportar para Excel só com o relatório de débito importado (sem ele a coluna Relatório ficaria vazia)
  const btnExp = $("btn-exportar");   // defensivo: na publicação o JS pode ir ao ar antes do HTML
  if (btnExp) {
    btnExp.disabled = !deb;
    btnExp.title = deb ? "Gerar um Excel com a tabela resumo desta data" : "Importe o relatório de débito para exportar";
  }

  // Parte 3: lançamentos da data, painel e empresas disponíveis no formulário
  const lancsData = lancamentosDaData(reg.iso);
  const lancPorEmpresa = porEmpresa(lancsData);
  estado.empresasResumo = res.cols.filter((c) => c.codigo != null && c.codigo !== "").map((c) => ({ codigo: String(c.codigo), empresa: c.empresa, diferenca: c.diferenca }));
  desenharPainel(reg, res, deb, lancsData);
  desenharValidacao(reg, res);
  const nLanc = (c) => (c.codigo != null ? (lancPorEmpresa.get(String(c.codigo)) || []).length : 0);

  // 30/10: todas as empresas à vista, sem rolagem: o quadro cresce até a largura da tela e,
  // se ainda não couber, as empresas quebram em faixas (cada faixa com Fluxo/Relatório/Diferença).
  const cols = res.cols;
  const celular = ehCelular();
  // largura da coluna pelo maior conteúdo (nome ou valor), medida com a fonte da tabela: nada fica espremido nem cortado
  const medir = (() => { const cv = document.createElement("canvas").getContext("2d");
    const fonte = getComputedStyle(document.body).fontFamily;
    return (texto, peso, px) => { cv.font = `${peso} ${px}px ${fonte}`; return cv.measureText(texto).width; }; })();
  const base = parseFloat(getComputedStyle(document.documentElement).fontSize) || 16;
  const maior = cols.reduce((m, c) => Math.max(m,
    medir((String(c.empresa) + (nLanc(c) ? ` · ${nLanc(c)}` : "")).toUpperCase(), 700, base * 0.72) * 1.06,     // cabeçalho em maiúsculas com espaçamento (e " · n" dos lançamentos)
    ...[c.fluxo, c.relatorio].filter((v) => v != null).map((v) => medir(contabil(v), 700, base * (celular ? 0.86 : 0.88))),
    // a diferença vai dentro do selo colorido: soma a folga interna do selo (.45rem de cada lado)
    ...[c.diferenca].filter((v) => v != null).map((v) => medir(contabil(v), 700, base * (celular ? 0.86 : 0.88)) + base * 0.9)), 0);
  // + folga da célula (.75rem de cada lado) + margem de segurança para a fonte
  const ROTULO = celular ? 96 : 120, COL = Math.max(celular ? 96 : 104, Math.ceil(maior + base * 1.5 + 12), colMinima), SOBRA = 24;
  const pai = cartao.parentElement.clientWidth - (parseFloat(getComputedStyle(cartao.parentElement).paddingLeft) || 0)
    - (parseFloat(getComputedStyle(cartao.parentElement).paddingRight) || 0);
  const tela = document.documentElement.clientWidth - 32;
  const precisa = ROTULO + cols.length * COL + SOBRA;
  const larg = Math.max(pai, Math.min(precisa, tela));
  cartao.style.width = larg > pai ? `${larg}px` : "";
  cartao.style.maxWidth = larg > pai ? "none" : "";
  cartao.style.marginLeft = larg > pai ? `${Math.round((pai - larg) / 2)}px` : "";
  cartao.style.setProperty("--resumo-rotulo", `${ROTULO}px`);
  const porLinha = Math.max(1, Math.floor((larg - ROTULO - SOBRA) / COL));
  const faixas = Math.max(1, Math.ceil(cols.length / porLinha));
  const porFaixa = Math.ceil(cols.length / faixas);   // faixas equilibradas
  const tabela = (grupo, i) => {
    const vazias = "<td></td>".repeat(porFaixa - grupo.length);
    const nome = (c) => nLanc(c)
      ? `<button type="button" class="cab-lanc" data-cod="${esc(c.codigo)}" aria-haspopup="dialog" ` +
        `aria-label="${esc(c.empresa)}: ${plural(nLanc(c), "lançamento", "lançamentos")}. Ver detalhes">${esc(c.empresa)} · ${nLanc(c)}</button>`
      : esc(c.empresa);
    // Validação (03/11/2026): data-chave + classe "validada" só mudam a COR do texto da empresa concluída; a estrutura não muda
    const atr = (c, extra = "") => ` data-chave="${esc(chaveEmpresa(c))}"` +
      ` class="${[extra, validada(reg.iso, chaveEmpresa(c)) ? "validada" : ""].filter(Boolean).join(" ")}"`;
    const cab = grupo.map((c) => `<th scope="col"${atr(c, c.semFluxo ? "sem-fluxo" : "")} title="${esc(c.empresa)}">${nome(c)}` +
      `${c.semFluxo ? '<span class="marca">sem fluxo</span>' : ""}</th>`).join("") + "<th></th>".repeat(porFaixa - grupo.length);
    const linha = (rotulo, classe, valor) =>
      `<tr class="${classe}"><th scope="row">${rotulo}</th>${grupo.map(valor).join("")}${vazias}</tr>`;
    return `<table class="tabela-resumo"${i === 0 ? ' id="resumo-tabela"' : ""} aria-label="Fluxo × Relatório de débito por empresa` +
      `${faixas > 1 ? ` (parte ${i + 1} de ${faixas})` : ""}"><colgroup><col class="rotulo"></colgroup>` +
      `<thead><tr><th scope="col">Empresa</th>${cab}</tr></thead><tbody>` +
      linha("Fluxo", "fluxo", (c) => `<td${atr(c)}>${esc(contabil(c.fluxo))}</td>`) +
      linha("Relatório", "relatorio", (c) => `<td${atr(c)}>${c.relatorio == null ? "—" : esc(contabil(c.relatorio))}</td>`) +
      linha("Diferença", "diferenca", (c) => c.diferenca == null ? "<td>—</td>"
        : `<td><span class="dif ${sinal(c.diferenca) || "zero"}"${dica(c.diferenca) ? ` title="${dica(c.diferenca)}"` : ""}>${esc(contabil(c.diferenca))}</span></td>`) +
      "</tbody></table>";
  };
  const grupos = [];
  for (let i = 0; i < cols.length; i += porFaixa) grupos.push(cols.slice(i, i + porFaixa));
  const caixa = $("resumo-caixa") || $("resumo-tabela").parentElement;
  caixa.innerHTML = grupos.map(tabela).join("");

  // conferência com a fonte real já desenhada: se alguma célula ficou apertada (fonte do computador mais larga
  // que a medida, como no Windows do Rocha em 30/10), refaz uma vez com a coluna na largura que faltou
  if (!colMinima) {
    const FOLGA = 8;
    let falta = 0;
    for (const cel of caixa.querySelectorAll("thead th[scope=col], tbody td")) {
      const cs = getComputedStyle(cel);
      const util = cel.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
      const dif = cel.querySelector(".dif");
      const conteudo = dif ? dif.getBoundingClientRect().width + FOLGA
        : cel.scrollWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
      if (cel.textContent && conteudo > util + 0.5) falta = Math.max(falta, conteudo - util);
    }
    if (falta > 0) { desenharResumo(reg, Math.ceil(COL + falta + 2)); return; }
  }
  desenharDetalhe();
}

/* ------------------------------------------------------------ Parte 3: lançamentos (modelo A, 03/11/2026) */
const lancamentosDaData = (iso) => estado.lancamentos.filter((l) => l.data_relatorio === iso);
/** Valor com sinal explícito: "+5.000,00" explica fluxo maior; "−100,00" explica débito maior. */
const moedaSinal = (v) => (v > 0 ? "+" : v < 0 ? "−" : "") + moeda(Math.abs(v));

/** Selo "Falta justificar": laranja = fluxo maior sem explicação; vermelho = passou/débito maior; verde = justificado. */
function seloFalta(f) {
  if (f.situacao === "semRelatorio") return { classe: "", texto: "Aguardando o relatório de débito", dica: "" };
  if (f.situacao === "justificado") return { classe: "ok", texto: "Justificado", dica: "Os lançamentos explicam toda a diferença." };
  if (f.situacao === "falta") return { classe: "pos", texto: `R$ ${moeda(f.falta)}`, dica: "Ainda falta explicar fluxo maior que o relatório de débito." };
  return { classe: "neg", texto: `R$ −${moeda(Math.abs(f.falta))}`, dica: "Os lançamentos passam da diferença (ou há débito maior ainda sem explicação)." };
}

function desenharPainel(reg, res, deb, lancsData) {
  const aside = $("painel-lanc");
  if (!aside) return;   // defensivo: na publicação o JS vai ao ar antes do HTML
  aside.classList.remove("oculto");
  const cartao = $("lanc-cartao"), btn = $("btn-lanc");
  if (estado.lancFalha) {   // sem leitura confiável não se mostra soma nem se aceita novo lançamento
    cartao.classList.add("oculto");
    btn.disabled = true;
    return;
  }
  cartao.classList.remove("oculto");
  btn.disabled = false;
  const { porCategoria, total } = somarPorCategoria(lancsData);
  $("lanc-categorias").innerHTML = CATEGORIAS.map((c) =>
    `<div class="lanc-linha"><span>${esc(c.nome)}</span><span class="lanc-val">${esc(contabil(porCategoria[c.id]))}</span></div>`).join("");
  $("lanc-total").textContent = contabil(total);
  const dif = deb ? res.total.diferenca : null;
  const celDif = $("lanc-dif");
  celDif.textContent = dif == null ? "—" : contabil(dif);
  celDif.className = `lanc-val ${dif == null || dif === 0 ? "" : dif > 0 ? "pos" : "neg"}`;
  const selo = seloFalta(faltaJustificar(dif, total));
  const pill = $("lanc-falta");
  pill.textContent = selo.texto;
  pill.className = `lanc-pill ${selo.classe}`;
  pill.title = selo.dica;
}

/** Detalhe dos lançamentos de uma empresa (aparece sob o cabeçalho "NOME · n" da tabela resumo). */
function desenharDetalhe() {
  const el = $("lanc-detalhe");
  if (!el) return;
  const reg = estado.atual && estado.datas.get(estado.atual);
  const cod = estado.detalhe;
  const lancs = reg && cod != null ? (porEmpresa(lancamentosDaData(reg.iso)).get(String(cod)) || []) : [];
  const col = reg ? estado.empresasResumo.find((c) => c.codigo === String(cod)) : null;
  const btn = el.parentElement?.querySelector(`.cab-lanc[data-cod="${CSS.escape(String(cod))}"]`);
  if (!reg || cod == null || !lancs.length || !col || !btn) {
    estado.detalhe = null; estado.confirmaExcluir = null;
    el.classList.add("oculto"); el.innerHTML = "";
    return;
  }
  const sit = situacaoDaEmpresa(col.diferenca, lancs);
  const selo = seloFalta({ situacao: sit.situacao, falta: sit.falta });
  const itens = lancs.map((l) => {
    const confirma = estado.confirmaExcluir === l.id;
    return `<li class="lanc-item" data-id="${l.id}">` +
      `<div class="lanc-item__topo"><strong>${esc(NOME_CATEGORIA[l.categoria] || l.categoria)}</strong>` +
      `<span class="lanc-val ${l.valor > 0 ? "pos" : "neg"}">${esc(moedaSinal(l.valor))}</span></div>` +
      `<div class="lanc-item__desc">${esc(l.descricao)}</div>` +
      (l.banco ? `<div class="lanc-item__banco">${esc(l.banco)}</div>` : "") +
      (confirma
        ? `<div class="lanc-item__confirma"><span>Excluir este lançamento?</span>` +
          `<button type="button" class="lanc-mini lanc-mini--perigo" data-acao="excluir-sim" data-id="${l.id}">Sim, excluir</button>` +
          `<button type="button" class="lanc-mini" data-acao="excluir-nao">Cancelar</button></div>`
        : `<div class="lanc-item__acoes">` +
          `<button type="button" class="lanc-mini lanc-mini--discreto lanc-mini--icone" data-acao="editar" data-id="${l.id}" ` +
          `aria-label="Editar lançamento" title="Editar lançamento">${ICONE_LAPIS}</button>` +
          `<button type="button" class="lanc-mini lanc-mini--discreto" data-acao="excluir" data-id="${l.id}">Excluir</button></div>`) +
      `</li>`;
  }).join("");
  const antes = el.querySelector(".lanc-detalhe__msg")?.innerHTML || "";
  el.innerHTML =
    `<div class="lanc-detalhe__topo"><h3 id="lanc-detalhe-titulo">${esc(col.empresa)} · ${plural(lancs.length, "lançamento", "lançamentos")}</h3>` +
    `<button type="button" class="lanc-fechar" data-acao="fechar" aria-label="Fechar" title="Fechar">${ICONE_X}</button></div>` +
    `<div class="lanc-detalhe__msg" aria-live="polite">${antes}</div>` +
    `<ul class="lanc-itens">${itens}</ul>` +
    `<div class="lanc-detalhe__rodape"><div><span>Lançado </span><strong>${esc(moeda(sit.lancado))}</strong>` +
    `<span> de </span><strong>${col.diferenca == null ? "—" : esc(moeda(col.diferenca))}</strong></div>` +
    `<div class="lanc-detalhe__falta"><span>${sit.situacao === "justificado" ? "Situação" : "Falta"}</span>` +
    `<span class="lanc-pill ${selo.classe}" title="${esc(selo.dica)}">${esc(sit.situacao === "falta" ? moeda(sit.falta) : selo.texto)}</span></div></div>`;
  el.classList.remove("oculto");
  // posição: logo abaixo do cabeçalho clicado, sem passar das bordas do cartão
  const cartao = el.parentElement;
  const cr = cartao.getBoundingClientRect(), br = btn.getBoundingClientRect();
  const larg = el.offsetWidth;
  const esquerda = Math.max(8, Math.min(br.left - cr.left, cr.width - larg - 8));
  el.style.left = `${Math.round(esquerda)}px`;
  el.style.top = `${Math.round(br.bottom - cr.top + 6)}px`;
}

function alternarDetalhe(cod) {
  const mesmo = estado.detalhe === String(cod);
  estado.detalhe = mesmo ? null : String(cod);
  estado.detalheIso = estado.atual;
  estado.confirmaExcluir = null;
  const el = $("lanc-detalhe");
  if (el) el.querySelector(".lanc-detalhe__msg")?.replaceChildren();
  desenharDetalhe();
  if (estado.detalhe) {
    const el = $("lanc-detalhe");
    try { el?.scrollIntoView({ block: "nearest" }); } catch { /* sem rolagem automática, segue */ }
    el?.querySelector("[data-acao=fechar]")?.focus({ preventScroll: true });
  }
}

function fecharDetalhe(devolverFoco = false) {
  const cod = estado.detalhe;
  estado.detalhe = null; estado.confirmaExcluir = null;
  desenharDetalhe();
  if (devolverFoco && cod != null) document.querySelector(`.cab-lanc[data-cod="${CSS.escape(String(cod))}"]`)?.focus();
}

async function aoClicarDetalhe(ev) {
  const b = ev.target.closest("[data-acao]");
  if (!b) return;
  const acao = b.dataset.acao;
  if (acao === "fechar") { fecharDetalhe(true); return; }
  if (acao === "editar") {
    const l = estado.lancamentos.find((x) => x.id === Number(b.dataset.id));
    if (l) abrirDialogo(true, l);
    return;
  }
  if (acao === "excluir") { estado.confirmaExcluir = Number(b.dataset.id); desenharDetalhe(); return; }
  if (acao === "excluir-nao") { estado.confirmaExcluir = null; desenharDetalhe(); return; }
  if (acao === "excluir-sim") {
    const id = Number(b.dataset.id);
    b.disabled = true;
    const msg = $("lanc-detalhe").querySelector(".lanc-detalhe__msg");
    try {
      await excluirLancamento(id);
    } catch (e) {
      estado.confirmaExcluir = null;
      desenharDetalhe();
      const m = $("lanc-detalhe")?.querySelector(".lanc-detalhe__msg");
      if (m) m.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui excluir o lançamento")));
      return;
    }
    estado.lancamentos = estado.lancamentos.filter((l) => l.id !== id);
    estado.confirmaExcluir = null;
    const reg = estado.atual && estado.datas.get(estado.atual);
    if (reg) desenharResumo(reg);
    $("lanc-msg").innerHTML = alerta("ok", "Lançamento excluído. O registro continua guardado no banco como excluído.");
    if (msg) msg.replaceChildren();
  }
}

/* formulário "Novo lançamento" */
const CAMPOS_ERRO = ["categoria", "valor", "empresa", "banco", "descricao"];
function limparErrosLanc() {
  for (const c of CAMPOS_ERRO) { const e = $(`erro-${c}`); if (e) e.textContent = ""; }
  $("lanc-form-msg").replaceChildren();
}
function atualizarBancos() {
  const caixa = $("lanc-banco-caixa");
  const ehOutros = $("lanc-categoria").value === "outros_bancos";
  caixa.classList.toggle("oculto", !ehOutros);
  if (!ehOutros) return;
  const cod = $("lanc-empresa").value;
  const contas = cod ? contasDaEmpresa(cod) : [];
  const nome = estado.empresasResumo.find((c) => c.codigo === cod)?.empresa;
  $("lanc-banco-rotulo").textContent = nome ? `Banco (${nome})` : "Banco";
  const sel = $("lanc-banco");
  sel.innerHTML = (contas.length === 1 ? "" : `<option value="">${contas.length ? "Selecione o banco" : "Sem conta cadastrada"}</option>`) +
    contas.map((c) => `<option value="${esc(c.rotulo)}">${esc(c.rotulo)}</option>`).join("");
  sel.disabled = !contas.length;
  $("erro-banco").textContent = cod && !contas.length ? "Esta empresa não tem conta bancária cadastrada no NOCTUS." : "";
}
function abrirDialogo(abrir, editar = null) {
  const dlg = $("lanc-dialogo"), fundo = $("lanc-fundo"), btn = $("btn-lanc");
  if (!dlg) return;
  if (abrir) {
    const reg = estado.atual && estado.datas.get(estado.atual);
    if (!reg || estado.lancFalha) return;
    fecharDetalhe();
    estado.dialogoIso = editar ? editar.data_relatorio : reg.iso;
    estado.editandoId = editar ? editar.id : null;
    $("lanc-dialogo-titulo").textContent = editar ? "Editar lançamento" : "Novo lançamento";
    $("lanc-salvar").textContent = editar ? "Salvar alterações" : "Salvar lançamento";
    $("lanc-categoria").innerHTML = CATEGORIAS.map((c) => `<option value="${c.id}">${esc(c.nome)}</option>`).join("");
    $("lanc-empresa").innerHTML = estado.empresasResumo.map((c) => `<option value="${esc(c.codigo)}">${esc(c.empresa)}</option>`).join("");
    // empresa sugerida: a primeira que ainda tem diferença a justificar
    const lancPor = porEmpresa(lancamentosDaData(reg.iso));
    const pendente = estado.empresasResumo.find((c) => c.diferenca != null &&
      situacaoDaEmpresa(c.diferenca, lancPor.get(c.codigo) || []).falta !== 0);
    $("lanc-empresa").value = (pendente || estado.empresasResumo[0])?.codigo ?? "";
    $("lanc-valor").value = ""; $("lanc-descricao").value = "";
    if (editar) {   // formulário preenchido com o lançamento escolhido
      $("lanc-categoria").value = editar.categoria;
      $("lanc-empresa").value = editar.empresa_codigo;
      $("lanc-valor").value = (editar.valor < 0 ? "-" : "") + moeda(Math.abs(editar.valor));
      $("lanc-descricao").value = editar.descricao;
    }
    limparErrosLanc();
    atualizarBancos();
    if (editar && editar.banco) $("lanc-banco").value = editar.banco;
    if (!estado.empresasResumo.length) {
      $("lanc-form-msg").innerHTML = alerta("erro", "Não há empresa com código cadastrado nesta data para receber lançamento.");
    }
  }
  if (!abrir) {   // ao fechar, o formulário volta ao modo "Novo lançamento"
    estado.editandoId = null;
    $("lanc-dialogo-titulo").textContent = "Novo lançamento";
    $("lanc-salvar").textContent = "Salvar lançamento";
  }
  dlg.classList.toggle("oculto", !abrir);
  fundo.classList.toggle("oculto", !abrir);
  btn.setAttribute("aria-expanded", String(abrir));
  if (abrir) {
    try { dlg.scrollIntoView({ block: "nearest" }); } catch { /* sem rolagem automática, segue */ }
    $("lanc-valor").focus({ preventScroll: true });
  }
}

async function enviarLancamento(ev) {
  ev.preventDefault();
  if (estado.salvandoLanc) return;
  limparErrosLanc();
  const iso = estado.dialogoIso;
  const editando = estado.editandoId != null ? estado.lancamentos.find((l) => l.id === estado.editandoId) || null : null;
  const r = validarLancamento({
    categoria: $("lanc-categoria").value, valorTexto: $("lanc-valor").value, empresaCodigo: $("lanc-empresa").value,
    empresas: estado.empresasResumo, banco: $("lanc-banco").value, descricao: $("lanc-descricao").value,
  });
  if (!r.ok) {
    for (const [campo, texto] of Object.entries(r.erros)) { const e = $(`erro-${campo}`); if (e) e.textContent = texto; }
    const primeiro = CAMPOS_ERRO.find((c) => r.erros[c]);
    ({ categoria: "lanc-categoria", valor: "lanc-valor", empresa: "lanc-empresa", banco: "lanc-banco", descricao: "lanc-descricao" }[primeiro]
      && $({ categoria: "lanc-categoria", valor: "lanc-valor", empresa: "lanc-empresa", banco: "lanc-banco", descricao: "lanc-descricao" }[primeiro])?.focus());
    return;
  }
  estado.salvandoLanc = true;
  const salvar = $("lanc-salvar");
  salvar.disabled = true; salvar.textContent = "Salvando…";
  try {
    const gravado = await salvarLancamento(r.lancamento, iso);
    estado.lancamentos.push(gravado);
    // Editar = grava a versão nova e só depois exclui (logicamente) a anterior: o histórico fica no banco e nenhuma
    // regra do banco que proíbe alterar valor foi afrouxada. Se a exclusão falhar, a tela avisa (a anterior continua contando).
    let falhaExcluir = null;
    if (editando) {
      try {
        await excluirLancamento(editando.id);
        estado.lancamentos = estado.lancamentos.filter((l) => l.id !== editando.id);
      } catch (e2) { falhaExcluir = e2; }
    }
    abrirDialogo(false);
    $("btn-lanc").focus();
    const reg = estado.atual && estado.datas.get(estado.atual);
    if (reg) desenharResumo(reg);
    const resumoLanc = `${esc(NOME_CATEGORIA[gravado.categoria])}, ${esc(moedaSinal(gravado.valor))} (${esc(gravado.empresa)})`;
    $("lanc-msg").innerHTML = falhaExcluir
      ? alerta("erro", esc(msgErro(falhaExcluir, "A alteração foi salva, mas não consegui excluir a versão anterior, que continua somando. Exclua-a manualmente")))
      : alerta("ok", editando ? `Lançamento alterado: ${resumoLanc}. A versão anterior continua guardada no banco como excluída.` : `Lançamento salvo: ${resumoLanc}.`);
  } catch (e) {
    $("lanc-form-msg").innerHTML = alerta("erro", esc(msgErro(e, "Não consegui salvar o lançamento")));
  } finally {
    estado.salvandoLanc = false;
    salvar.disabled = false; salvar.textContent = estado.editandoId != null ? "Salvar alterações" : "Salvar lançamento";
  }
}

/** Ao mudar o tamanho da janela, refaz a tabela resumo (largura do quadro e faixas). */
let refazerResumo = null;
function aoRedimensionar() {
  clearTimeout(refazerResumo);
  refazerResumo = setTimeout(() => { const reg = estado.atual && estado.datas.get(estado.atual); if (reg) desenharResumo(reg); }, 120);
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
  ligar("btn-exportar", exportarResumo);
  // Validação (opcional): o botão mostra/oculta a tabela; as caixas gravam no servidor (defensivo: HTML antigo não tem o bloco)
  if ($("btn-validacao")) {
    ligar("btn-validacao", () => {
      estado.validAberta = !estado.validAberta;
      const reg = estado.atual && estado.datas.get(estado.atual);
      if (reg) atualizarValidacao(reg.iso);
    });
    $("validacao-corpo").addEventListener("change", (ev) => {
      const cb = ev.target.closest(".valid-cb");
      const reg = estado.atual && estado.datas.get(estado.atual);
      if (!cb || !reg || estado.validFalha) return;
      marcarValidacao(reg.iso, cb.dataset.chave, cb.dataset.empresa, cb.dataset.item, cb.checked);
    });
  }
  // Parte 3: lançamentos (defensivo: HTML antigo não tem o painel)
  if ($("btn-lanc")) {
    ligar("btn-lanc", () => abrirDialogo($("lanc-dialogo").classList.contains("oculto")));
    ligar("lanc-fechar", () => { abrirDialogo(false); $("btn-lanc").focus(); });
    ligar("lanc-cancelar", () => { abrirDialogo(false); $("btn-lanc").focus(); });
    ligar("lanc-fundo", () => abrirDialogo(false));
    $("lanc-form").addEventListener("submit", enviarLancamento);
    $("lanc-categoria").addEventListener("change", atualizarBancos);
    $("lanc-empresa").addEventListener("change", atualizarBancos);
    $("lanc-detalhe").addEventListener("click", aoClicarDetalhe);
    document.addEventListener("click", (ev) => {
      const b = ev.target.closest(".cab-lanc");
      if (b) { alternarDetalhe(b.dataset.cod); return; }
      if (estado.detalhe == null) return;
      // composedPath: o clique pode ter redesenhado (e soltado) o elemento clicado dentro do detalhe
      if (ev.composedPath().some((n) => n.id === "lanc-detalhe")) return;
      fecharDetalhe();
    });
    document.addEventListener("keydown", (ev) => {
      if (ev.key !== "Escape") return;
      if (!$("lanc-dialogo").classList.contains("oculto")) { abrirDialogo(false); $("btn-lanc").focus(); }
      else if (estado.detalhe != null) fecharDetalhe(true);
    });
  }
  window.addEventListener("resize", aoRedimensionar);
  // a largura das colunas é medida com a fonte Inter: quando ela termina de carregar, mede de novo
  try {
    document.fonts?.ready?.then(aoRedimensionar);
    document.fonts?.addEventListener?.("loadingdone", aoRedimensionar);   // fonte que chega depois da primeira medida
  } catch { /* sem a API de fontes, segue com a medida atual */ }
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
  if (!ehAdminNoNavegador()) return;              // perfil "user": continua vendo "Módulo em desenvolvimento"
  const confirmado = await adminConfirmado();
  if (confirmado === "nao") { liberarAviso(); return; }
  $("em-desenvolvimento").classList.add("oculto");
  $("modulo").classList.remove("oculto");
  liberarAviso();
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
  // Parte 3: se os lançamentos não carregarem, o painel não mostra soma (poderia enganar) e o aviso explica
  try {
    await carregarLancamentos();
  } catch (e) {
    estado.lancFalha = true;
    const m = $("lanc-msg");
    if (m) m.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui carregar os lançamentos")));
  }
  // Validação (opcional): se não carregar, nada fica verde, as caixas ficam travadas e o aviso explica (sem texto técnico)
  try {
    await carregarValidacoes();
  } catch (e) {
    estado.validFalha = true;
    const m = $("validacao-msg");
    if (m) m.innerHTML = alerta("erro", esc(msgErro(e, "Não consegui carregar as validações já feitas")));
  }
  desenhar();
}

/** Tira a marca "decidindo" do HTML: até aqui o quadro "em desenvolvimento" ficava escondido para o administrador (03/11/2026). */
function liberarAviso() { document.documentElement.classList.remove("decidindo"); }

iniciar().catch((e) => { liberarAviso(); registrarErro(e, "abertura do Controle Bancário"); });
