/*
 * NOCTUS — registro dos erros técnicos (pedido do Rocha, 21/10/2026).
 *
 * As telas mostram só "Falha de comunicação… contate o administrador"; o
 * detalhe técnico vem para cá e fica guardado no servidor, na tabela
 * erros_noctus. Nenhuma tela do sistema lê essa tabela (pedido do Rocha,
 * 22/10/2026): ela só é vista pelo painel do Supabase, como um log —
 * Table Editor > erros_noctus, ou SQL Editor.
 *
 * Como funciona:
 *  - registrar(erro, "o que a pessoa estava fazendo") põe o erro numa fila no
 *    navegador (localStorage) — assim nada se perde quando o erro É a falta de
 *    conexão, ou quando ainda não há login;
 *  - a fila é enviada ao servidor assim que houver sessão válida e internet
 *    (alguns segundos depois do erro, ao voltar a conexão e a cada minuto);
 *  - erros não tratados da página (window "error"/"unhandledrejection") também
 *    entram, menos os que não dizem nada ("Script error.", extensões do navegador).
 *  - Limites para não virar enxurrada: o mesmo erro repetido em 10 minutos só
 *    soma em "ocorrencias"; até 30 registros por carregamento de página; fila
 *    de no máximo 50 itens; textos cortados. O servidor ainda limita por usuário.
 *  - Não guarda dado de conferência (valores, CPF/CNPJ, favorecido): só a
 *    mensagem e a pilha do erro, a tela, o navegador e o usuário logado.
 */
(function () {
  "use strict";
  if (window.NoctusErros) return;

  var SB_URL = "https://wyyqjlwzphlhnplmirlg.supabase.co";
  var SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind5eXFqbHd6cGhsaG5wbG1pcmxnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyMTU0MDcsImV4cCI6MjA5OTc5MTQwN30.8TEY_BZAm7P9vwuU0Eh5jzw6rgzcPOpTXk_gBU5FQJc";
  var FILA = "noctus.erros.fila";
  var MAX_FILA = 50, MAX_POR_PAGINA = 30, JANELA_REPETIDO = 10 * 60 * 1000;
  var naPagina = 0, enviando = false, timer = null;

  function ler() {
    try { var f = JSON.parse(localStorage.getItem(FILA) || "[]"); return Array.isArray(f) ? f : []; }
    catch (e) { return []; }
  }
  function gravar(f) {
    try { localStorage.setItem(FILA, JSON.stringify(f.slice(-MAX_FILA))); } catch (e) { /* navegador cheio */ }
  }
  function sessao() {
    try {
      var b = localStorage.getItem("sipep.session") || sessionStorage.getItem("sipep.session");
      var s = b && JSON.parse(b);
      if (s && s.accessToken && s.expiresAt > Date.now() + 30000) return s;
    } catch (e) { /* sem sessão */ }
    return null;
  }
  function corta(t, n) {
    t = String(t == null ? "" : t);
    return t.length > n ? t.slice(0, n - 1) + "…" : t;
  }
  function modulo() {
    var p = location.pathname;
    if (/conferencia-homolog/.test(p)) return "conferencia-homolog";
    if (/conferencia/.test(p)) return "conferencia";
    if (/\/conta\//.test(p)) return "conta";
    if (/\/homolog\//.test(p)) return "painel-homolog";
    return "painel";
  }
  function texto(e) {
    if (e instanceof Error) return (e.name && e.name !== "Error" ? e.name + ": " : "") + (e.message || "");
    if (typeof e === "string") return e;
    if (e && typeof e.message === "string") return e.message;
    try { return JSON.stringify(e); } catch (x) { return String(e); }
  }

  function registrar(erro, operacao) {
    try {
      if (naPagina >= MAX_POR_PAGINA) return;
      var item = {
        ocorrido_em: new Date().toISOString(),
        modulo: modulo(),
        tela: corta(location.pathname + (location.hash || "").split("?")[0], 200),
        operacao: corta(operacao || "", 200),
        mensagem: corta((texto(erro) || "(sem mensagem)") + (erro && erro.detalheTecnico ? " — " + erro.detalheTecnico : ""), 1000),
        detalhe: corta((erro && erro.stack) || "", 4000),
        navegador: corta(navigator.userAgent, 300),
        usuario: null,
        ocorrencias: 1,
      };
      var s = sessao();
      if (s && s.user) item.usuario = corta(s.user.username || s.user.displayName || s.user.id || "", 80);
      var chave = item.modulo + "|" + item.operacao + "|" + item.mensagem, f = ler();
      for (var i = f.length - 1; i >= 0; i--) {
        if (f[i]._k === chave && Date.now() - Date.parse(f[i].ocorrido_em) < JANELA_REPETIDO) {
          f[i].ocorrencias = (f[i].ocorrencias || 1) + 1;
          gravar(f); agendar(); return;
        }
      }
      item._k = chave;
      f.push(item); naPagina++;
      gravar(f); agendar();
    } catch (x) { /* o registro de erro nunca pode derrubar a tela */ }
  }

  function agendar(ms) {
    clearTimeout(timer);
    timer = setTimeout(enviar, ms || 3000);
  }

  function enviar() {
    if (enviando || navigator.onLine === false) return;
    var s = sessao(), f = ler();
    if (!s || !f.length) return;
    enviando = true;
    var lote = f.slice(0, 20);
    var marca = function (it) { return it._k + "@" + it.ocorrido_em; };
    var corpo = lote.map(function (it) {
      return { ocorrido_em: it.ocorrido_em, modulo: it.modulo, tela: it.tela, operacao: it.operacao,
        mensagem: it.mensagem, detalhe: it.detalhe, navegador: it.navegador, usuario: it.usuario,
        ocorrencias: Math.min(it.ocorrencias || 1, 100000) };
    });
    fetch(SB_URL + "/rest/v1/erros_noctus", {
      method: "POST",
      headers: { apikey: SB_KEY, Authorization: "Bearer " + s.accessToken,
        "Content-Type": "application/json", Prefer: "return=minimal" },
      body: JSON.stringify(corpo),
      keepalive: true,
    }).then(function (r) {
      // enviado, ou recusado por conteúdo (400/422): sai da fila para não travá-la
      if (r.ok || r.status === 400 || r.status === 422) {
        var enviados = lote.map(marca), resto = ler().filter(function (it) { return enviados.indexOf(marca(it)) < 0; });
        gravar(resto);
        if (r.ok && resto.length) agendar(1000);
      }
    }).catch(function () { /* sem conexão: tenta de novo depois */ })
      .then(function () { enviando = false; });
  }

  window.addEventListener("error", function (ev) {
    var m = ev && ev.message ? String(ev.message) : "";
    if (!m || /^Script error\.?$/i.test(m) || /ResizeObserver loop/i.test(m)) return;
    if (ev.filename && !/^https?:\/\/(pedrorocha3247\.github\.io|localhost|127\.0\.0\.1)/.test(ev.filename)) return;
    registrar(ev.error || m, "erro não tratado na página");
  });
  window.addEventListener("unhandledrejection", function (ev) {
    registrar(ev ? ev.reason : "promessa rejeitada", "operação sem tratamento de erro");
  });
  window.addEventListener("online", function () { agendar(1500); });
  setTimeout(enviar, 4000);
  setInterval(function () { if (ler().length) enviar(); }, 60000);

  window.NoctusErros = { registrar: registrar, enviar: enviar };
})();
