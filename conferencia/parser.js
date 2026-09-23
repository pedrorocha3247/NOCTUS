/**
 * Extração do relatório "Solicitações de Pagamentos por Débito em Conta Corrente" (SCK).
 *
 * Roda no navegador, em cima do pdf.js. Nenhuma dependência de servidor:
 * o PDF nunca sai da máquina de quem está conferindo.
 *
 *   const dados = await parseRelatorio(arrayBuffer, pdfjsLib);
 *
 * ACEITA DOIS FORMATOS DE RELATÓRIO. Em 22/09/2026 o gerador mudou o layout
 * sem aviso: colunas deslocadas de forma NÃO uniforme (o vão entre PODER e
 * CPF/CNPJ cresceu, o vão entre BANCO e o fim da linha encolheu — não dá pra
 * corrigir com um deslocamento fixo), o texto de cada total passou a vir
 * rotulado ("Total por forma de pagamento ==>" / "Total por conta ==>" em
 * vez de só "Total ==>"), o número passou a usar separador de milhar "," e
 * decimal "." (era o contrário), "Ag." virou "AG." e o tipo "TRANSFERÊNCIA
 * BANCÁRIA CONTA CORRENTE" perdeu o sufixo. A pior parte: a heurística antiga
 * de família por posição do "S.N" (x >= 31 → família A) INVERTEU entre os
 * dois formatos — não dá pra distinguir um relatório do outro só pela
 * posição de uma palavra.
 *
 * Por isso a escolha de estratégia é por DOCUMENTO inteiro, decidida antes
 * de processar qualquer cabeçalho, por um sinal textual exclusivo do
 * formato novo ("Total por " — ver formatoNovo() abaixo):
 *
 *  - formato ANTIGO: mantém a tabela de colunas FIXA por família (COLUNAS),
 *    exatamente como já estava em produção e validada em vários relatórios
 *    reais — zero risco de regressão, porque é o mesmo caminho de código.
 *  - formato NOVO: lê a posição de cada rótulo no cabeçalho da PRÓPRIA
 *    seção (detectarColunas) em vez de uma tabela fixa em pixel, porque o
 *    deslocamento entre os dois formatos não é uniforme.
 *
 * A tentativa de usar detectarColunas também no formato antigo (só com
 * proporções ajustadas) foi feita e descartada: bateu 100% num relatório de
 * um dia mas divergiu em ~12% das solicitações (344 de 2907) de um relatório
 * de vários dias, em campos de texto livre (favorecido/destinação/
 * solicitante/competente) — a tabela fixa, tunada à mão contra muitos casos
 * reais ao longo do tempo, é mais confiável que um corte derivado só da
 * geometria do cabeçalho. Por isso o caminho antigo fica intocado.
 *
 * TAMBÉM ACEITA O RELATÓRIO EM EXCEL (.xlsx), pela função
 * parseRelatorioExcel() no fim deste arquivo — ver o comentário dela para
 * o porquê de o Excel ser preferível ao PDF quando os dois estão
 * disponíveis (valores numéricos nativos, colunas por nome em vez de
 * posição em pixel) e como a EMPRESA de cada solicitação é identificada
 * (o Excel sai com todas as empresas juntas, ao contrário do PDF).
 */

import { empresaPorConta, rotuloEmpresa } from "./contas.js";
import { verificarGrupoEconomico } from "./grupoEconomico.js";

// Duas famílias de layout convivem no relatório ANTIGO.
// A: TRANSFERÊNCIA / TED / PIX      (cabeçalho "S.N" em x >= 31)
// B: DÉBITO EM CONTA / BOLETO       (cabeçalho "S.N" em x <  31)
const COLUNAS = {
  A: [["sn",0,52],["filial",45,88],["valor",88,133],["solicitante",133,186],
      ["competente",186,240],["poder",240,270],["cpfCnpj",270,336],
      ["favorecido",336,458],["destinacao",458,615],["complemento",615,760],
      ["poderDispendio",760,9999]],
  // O corte entre FILIAL e VALOR na família B ficava em x=85, colado no início
  // real do valor (x≈84,2) quando o valor tem 7 dígitos inteiros (>= R$
  // 1.000.000,00) — a coluna é alinhada à direita, então um valor mais largo
  // "nasce" mais à esquerda. Nesse caso o texto do valor caía inteiro na
  // coluna FILIAL, o campo VALOR ficava vazio (null) e a solicitação era
  // contada na quantidade mas somava zero — foi o que fez o relatório de
  // 18/09/2026 (BOLETO de R$ 1.016.008,01) não bater no valor total mesmo com
  // a quantidade de solicitações correta. "MATRIZ" termina em x≈77,2, então
  // baixar o corte para 80 dá margem para valores de até 8 dígitos inteiros
  // sem invadir a coluna FILIAL.
  B: [["sn",0,50],["filial",45,80],["valor",80,123],["solicitante",123,176],
      ["competente",176,235],["poder",235,262],["cpfCnpj",262,324],
      ["favorecido",324,430],["destinacao",430,620],["complemento",620,770],
      ["poderDispendio",770,9999]],
};
const TIPOS = ["TRANSFERÊNCIA BANCÁRIA CONTA CORRENTE","TRANSFERÊNCIA BANCÁRIA","TRANSFERÊNCIA BANCÁRIA - CONTA CORRENTE","TED","DÉBITO EM CONTA","PIX","BOLETO"];
const TIPO_CURTO = {
  "TRANSFERÊNCIA BANCÁRIA CONTA CORRENTE": "TRANSFERÊNCIA",
  "TRANSFERÊNCIA BANCÁRIA": "TRANSFERÊNCIA", // formato novo (22/09/2026 em diante)
  "TRANSFERÊNCIA BANCÁRIA - CONTA CORRENTE": "TRANSFERÊNCIA", // formato Excel
};
const RUIDO = ["Total ==>","Total por","Página","Filial:","SOLICITAÇÕES DE","DDP -"];

// No formato antigo a linha é só "Total ==> N Solicitação(ões) R$ V"; a
// partir de 22/09/2026 o gerador passou a rotular cada total ("Total por
// forma de pagamento ==>" para o subtotal por tipo, "Total por conta ==>"
// para o total da conta) — o texto entre "Total" e "==>" varia, mas o
// padrão qtd/valor depois de "==>" é o mesmo, então basta deixar esse meio
// livre em vez de casar literalmente. A lógica que já existia de descartar
// a linha de total redundante (quando ela bate com a soma das anteriores do
// mesmo bloco) cobre tanto "Total ==>" solto quanto "Total por conta"
// somando os "Total por forma de pagamento" da mesma conta — não precisou
// mudar.
const RE_TOTAL = /Total(?:\s+por\s+[^=]+?)?\s*==>\s+([\d.]+)\s+Solicitação\(ões\)\s+R\$\s+([\d.,]+)/;

// Cabeçalho de seção por banco: "<NOME DO BANCO OU FAVORECIDO> Ag. <cod> - C/C <cod>".
// Nem todo cabeçalho começa com a palavra "BANCO" — em relatórios de mais de um
// dia aparecem seções cujo "banco" é na verdade um favorecido recorrente, como
// "MERCADOPAGO.COM REPRESENTACOES LTDA Ag. 0001 - C/C 67953790764". Exigir
// "BANCO " no início fazia essas seções nunca fecharem o bloco anterior, e os
// totais de duas seções distintas eram somados como se fossem uma só — foi o
// que impediu a validação de bater no relatório de 01 a 18/09/2026 (a seção do
// Santander seguida da do MercadoPago, nas páginas finais, virava um bloco só).
// "Ag." virou "AG." no formato novo — comparar sem diferenciar maiúsculas.
const ehCabecalhoBanco = (t) => /ag\./i.test(t) && /c\/c/i.test(t);

// Sinal exclusivo do formato novo (22/09/2026 em diante): cada linha de
// total ganhou um rótulo ("Total por forma de pagamento ==>", "Total por
// conta ==>"). O formato antigo nunca imprime "Total por" em lugar nenhum
// do relatório, então basta procurar essa substring uma vez no texto
// inteiro do documento antes de processar qualquer cabeçalho — decide qual
// dos dois caminhos de extração usar (tabela fixa vs. detecção dinâmica).
const formatoNovo = (textoCompleto) => /Total\s+por\s+/.test(textoCompleto);

// Campos de solicitação, na ordem em que aparecem no relatório (as duas
// famílias de layout diferem só no campo final: BANCO/AGÊNCIA/CONTA
// BANCÁRIA em TED/PIX/TRANSFERÊNCIA, OBSERVAÇÃO em DÉBITO EM CONTA/BOLETO)
// — usado só pelo caminho do formato NOVO (detectarColunas).
const ORDEM_CAMPOS = ["sn","filial","valor","solicitante","competente","poder","cpfCnpj","favorecido","destinacao","complemento","poderDispendio"];
const OBRIGATORIOS = ["sn","filial","valor","solicitante","competente","poder","cpfCnpj","favorecido","destinacao","poderDispendio"];

/**
 * FORMATO NOVO SÓ: lê as colunas de uma seção a partir do PRÓPRIO cabeçalho
 * impresso ("S.N FILIAL VALOR SOLICITANTE ... PODER DO DISPÊNDIO"), porque
 * o deslocamento de colunas do formato novo em relação ao antigo não é
 * uniforme (ver comentário no topo do arquivo) — uma tabela fixa nova
 * quebraria de novo no próximo ajuste de margem/fonte do gerador.
 *
 * @param {{texto:string,x:number,y:number}[]} palavrasHeader palavras da
 *   linha de cabeçalho (± tolerância de y — ver chamada), em qualquer ordem.
 * @param {{texto:string,x:number,y:number}[]} [palavrasPagina] todas as
 *   palavras da página (para achar "DISPÊNDIO" quando ele quebra de linha
 *   para fora da janela de `palavrasHeader` — ver comentário abaixo).
 * @param {number} [yCentro] y do cabeçalho (S.N), usado só para desempatar
 *   quando a mesma família de coluna se repete mais de uma vez na página.
 * @returns {{colunas: [string,number,number][], usedYs: Set<number>}}
 */
function detectarColunas(palavrasHeader, palavrasPagina, yCentro) {
  const marcos = [];
  const usedYs = new Set();      // y (arredondado) de toda palavra que compôs o cabeçalho —
                                  // usado pelo chamador para marcar exatamente essas linhas
                                  // como "ignorar", sem depender de uma janela fixa de y que
                                  // tanto pode faltar uma sublinha quebrada quanto engolir a
                                  // primeira linha de dado real logo abaixo do cabeçalho.
  const add = (campo, x, y) => { if (!marcos.some((m) => m.campo === campo)) marcos.push({ campo, x }); if (y != null) usedYs.add(Math.round(y)); };
  const p = [...palavrasHeader].sort((a, b) => a.x - b.x);

  // "PODER DO DISPÊNDIO" é um rótulo de 3 palavras que, quando a coluna é
  // estreita (família com BANCO/AGÊNCIA/CONTA BANCÁRIA à direita — pouco
  // espaço sobra), o gerador quebra em duas sublinhas: "PODER DO" numa
  // linha, "DISPÊNDIO" na linha de baixo, centralizadas verticalmente.
  // Dois problemas com isso:
  //  1) quando a quebra cabe dentro da janela de y já usada para juntar o
  //     cabeçalho, DISPÊNDIO pode nascer com x ligeiramente MENOR que o de
  //     PODER, invertendo a ordem no sort por x — por isso a busca de "DO"
  //     e "DISPÊNDIO" é por PROXIMIDADE (x/y perto de PODER), não por
  //     adjacência de índice no array ordenado;
  //  2) às vezes a quebra é maior (~8pt) que a própria janela de y usada
  //     para reunir o cabeçalho, e DISPÊNDIO fica de fora de
  //     `palavrasHeader` inteiramente — por isso, se não achar por perto
  //     dentro do cabeçalho, procura na página toda, mas travado por x
  //     (a palavra de baixo nasce quase exatamente sob "PODER", tolerância
  //     apertada) para não pegar por engano um valor de dado de outra linha
  //     nem a mesma família repetida em outra seção da página.
  const achar = (lista, origem, texto, maxDx, maxDy) =>
    lista.find((q) => q.texto === texto && Math.abs(q.x - origem.x) <= maxDx && Math.abs((q.y ?? 0) - (origem.y ?? 0)) <= maxDy);

  let familia = null; // "A" (BANCO/AGÊNCIA/CONTA BANCÁRIA) ou "B" (OBSERVAÇÃO)

  for (let i = 0; i < p.length; i++) {
    const t = p[i].texto, x = p[i].x, y = p[i].y;
    if (t === "S.N") add("sn", x, y);
    else if (t === "FILIAL") add("filial", x, y);
    else if (t === "VALOR") add("valor", x, y);
    else if (t === "SOLICITANTE") add("solicitante", x, y);
    else if (t === "COMPETENTE") add("competente", x, y);
    else if (t === "PODER" && achar(p, p[i], "DO", 40, 10)) {
      const doPerto = achar(p, p[i], "DO", 40, 10);
      const dispPerto = achar(p, p[i], "DISPÊNDIO", 40, 10) || achar(palavrasPagina || [], { x, y: yCentro }, "DISPÊNDIO", 8, 20);
      if (dispPerto) {
        add("poderDispendio", x, y);
        usedYs.add(Math.round(doPerto.y));
        usedYs.add(Math.round(dispPerto.y));
      } else add("poder", x, y);
    }
    else if (t === "PODER") add("poder", x, y);
    else if (t === "CPF") add("cpfCnpj", x, y);
    else if (t === "FAVORECIDO") add("favorecido", x, y);
    else if (t === "DESTINAÇÃO") add("destinacao", x, y);
    else if (t === "OBSERVAÇÃO") { add("complemento", x, y); familia = "B"; }  // família "DÉBITO/BOLETO"
    else if (t === "BANCO") { add("complemento", x, y); familia = "A"; }      // família "TED/PIX/TRANSFERÊNCIA"
  }

  const faltando = OBRIGATORIOS.filter((c) => !marcos.some((m) => m.campo === c));
  if (faltando.length)
    throw new Error(`Cabeçalho do relatório não reconhecido — colunas não encontradas: ${faltando.join(", ")}. O layout do relatório mudou?`);

  marcos.sort((a, b) => ORDEM_CAMPOS.indexOf(a.campo) - ORDEM_CAMPOS.indexOf(b.campo));

  // O rótulo de cada coluna no cabeçalho marca só o INÍCIO do texto do
  // rótulo — não o início real dos dados. Um corte no meio do vão entre
  // dois rótulos (50/50) funciona bem na maioria das transições, mas falha
  // nas duas abaixo, onde um lado é texto curto e previsível e o outro pode
  // ser bem mais largo que a distância entre os dois rótulos:
  //  - competente -> poder: PODER é só um código de 2 dígitos, então
  //    quase não precisa de folga à esquerda do próprio rótulo — dar a
  //    volta toda para COMPETENTE (nome da pessoa, pode ser longo).
  //  - destinacao -> complemento e complemento -> poderDispendio: aqui a
  //    proporção certa depende da FAMÍLIA da seção. Em "A" o complemento é
  //    BANCO/AGÊNCIA/CONTA (curto); em "B" é OBSERVAÇÃO (texto livre, pode
  //    ser longo) — por isso "B" recebe uma folga bem maior antes de
  //    PODER DO DISPÊNDIO do que "A".
  // As demais transições ficam no corte 50/50 padrão — forçar uma razão
  // diferente nelas piorou o relatório novo sem necessidade (o vão entre
  // rótulos não muda de forma uniforme entre seções/famílias).
  const RATIOS = {
    poder: 0.93, // competente -> poder
  };
  const RATIO_COMPLEMENTO = { A: 0.92, B: 0.71 };
  const RATIO_PODER_DISPENDIO = { A: 0.94, B: 1.0 };

  let inicio = 0;
  const colunas = [];
  for (let i = 0; i < marcos.length; i++) {
    const atual = marcos[i], proximo = marcos[i + 1];
    let fim = proximo ? (atual.x + proximo.x) / 2 : 9999;
    if (proximo) {
      const vao = proximo.x - atual.x;
      if (proximo.campo === "valor") fim = proximo.x - vao * 0.35;
      else if (proximo.campo === "complemento") fim = atual.x + vao * (RATIO_COMPLEMENTO[familia] ?? 0.8);
      else if (proximo.campo === "poderDispendio") fim = atual.x + vao * (RATIO_PODER_DISPENDIO[familia] ?? 0.95);
      else if (RATIOS[proximo.campo] != null) fim = atual.x + vao * RATIOS[proximo.campo];
    }
    colunas.push([atual.campo, inicio, fim]);
    inicio = fim;
  }
  colunas[colunas.length - 1][2] = 9999;
  return { colunas, usedYs };
}

const coluna = (colunas, x) => (colunas.find(([, i, f]) => x >= i && x < f) || ["poderDispendio"])[0];

// Formato antigo: separador de milhar "." e decimal "," (ex.: "2.798.927,37").
// Formato novo (22/09/2026 em diante): separador de milhar "," e decimal "."
// (ex.: "246,269.32"), inclusive em valores sem milhar ("808.55") — mudou
// junto com o layout de colunas. Em vez de fixar um dos dois estilos, decide
// pelo ÚLTIMO separador que aparece na string: esse é o decimal (só ele
// importa, os dígitos depois dele são as casas decimais); qualquer
// separador antes dele é milhar e é só removido. Com um único separador (ou
// nenhum), decide pelo que sobrou: só vírgula → decimal (antigo); só ponto
// ou nenhum → mantém como está (já é o formato que parseFloat entende).
function num(t) {
  let s = String(t).trim();
  if (!s) return null;
  const ultimaVirgula = s.lastIndexOf(",");
  const ultimoPonto = s.lastIndexOf(".");
  if (ultimaVirgula !== -1 && ultimoPonto !== -1) {
    s = ultimaVirgula > ultimoPonto
      ? s.replace(/\./g, "").replace(",", ".")   // vírgula decimal (antigo)
      : s.replace(/,/g, "");                     // ponto decimal (novo)
  } else if (ultimaVirgula !== -1) {
    s = s.replace(",", ".");
  }
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}

/** Palavras da página, com x e y no mesmo referencial do relatório (y cresce para baixo). */
async function palavrasDaPagina(pagina) {
  const conteudo = await pagina.getTextContent();
  const altura = pagina.view[3];
  const saida = [];
  for (const item of conteudo.items) {
    const texto = item.str;
    if (!texto || !texto.trim()) continue;
    const x = item.transform[4];
    const y = altura - item.transform[5];
    // um item pode trazer várias palavras; distribui pela largura para não
    // jogar tudo na coluna da primeira palavra
    const pedacos = texto.trim().split(/\s+/);
    const larguraChar = item.width / Math.max(texto.length, 1);
    let deslocamento = 0;
    for (const p of pedacos) {
      const inicio = texto.indexOf(p, deslocamento);
      saida.push({ texto: p, x: x + (inicio < 0 ? 0 : inicio) * larguraChar, y });
      deslocamento = (inicio < 0 ? deslocamento : inicio) + p.length;
    }
  }
  return saida;
}

/**
 * Agrupa palavras em linhas tolerando pequenas variações de y.
 *
 * O gerador do relatório desenha a linha de "Total ==>" — e, no formato
 * novo, também a linha de cabeçalho de coluna — em pedaços com alturas
 * levemente diferentes (1 a 4 pontos de distância). Agrupando por y exato, a
 * linha chega quebrada e o total (ou o cabeçalho) se perde; com tolerância,
 * ela volta inteira.
 */
function agruparTolerante(palavras, tolerancia) {
  const ordenadas = [...palavras].sort((a, b) => a.y - b.y || a.x - b.x);
  const linhas = [];
  let atual = null;
  for (const p of ordenadas) {
    if (!atual || p.y - atual.y > tolerancia) {
      atual = { y: p.y, palavras: [p] };
      linhas.push(atual);
    } else {
      atual.palavras.push(p);
    }
  }
  return linhas.map((l) => l.palavras.sort((a, b) => a.x - b.x).map((p) => p.texto).join(" "));
}

function agruparLinhas(palavras) {
  const mapa = new Map();
  for (const p of palavras) {
    const y = Math.round(p.y);
    if (!mapa.has(y)) mapa.set(y, []);
    mapa.get(y).push(p);
  }
  for (const lista of mapa.values()) lista.sort((a, b) => a.x - b.x);
  return mapa;
}

export async function parseRelatorio(arrayBuffer, pdfjsLib) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  // Lê todas as páginas uma vez só (usado tanto para decidir o formato
  // quanto para o parse em si — evita buscar cada página duas vezes).
  const palavrasPorPagina = [];
  let novo = false;
  for (let n = 1; n <= pdf.numPages; n++) {
    const pagina = await pdf.getPage(n);
    const palavras = await palavrasDaPagina(pagina);
    palavrasPorPagina.push(palavras);
    if (!novo && formatoNovo(palavras.map((p) => p.texto).join(" "))) novo = true;
  }

  const meta = { empresa: null, empresaCodigo: null, empresaNome: null, dataInicio: null, dataFim: null, emitidoEm: null, paginas: pdf.numPages };
  const solicitacoes = [];
  const blocos = [];
  let blocoAtual = [];
  let banco = null, tipo = null, familia = "A", colunasAtuais = null;

  for (let n = 1; n <= pdf.numPages; n++) {
    const palavras = palavrasPorPagina[n - 1];

    // passe tolerante: fecha blocos por banco e lê os totais impressos
    for (const L of agruparTolerante(palavras, 3)) {
      const texto = L.trim();
      if (ehCabecalhoBanco(texto)) {
        if (blocoAtual.length) blocos.push(blocoAtual);
        blocoAtual = [];
        continue;
      }
      const mt = texto.match(RE_TOTAL);
      if (mt) blocoAtual.push([parseInt(mt[1].replace(/\./g, ""), 10), num(mt[2])]);
    }

    const linhas = agruparLinhas(palavras);
    const ys = [...linhas.keys()].sort((a, b) => a - b);
    const texto = new Map(ys.map((y) => [y, linhas.get(y).map((p) => p.texto).join(" ")]));

    const ignorar = new Set();
    const paradas = [];   // linhas de total/cabeçalho: nenhum registro atravessa
    const ancoras = [];
    for (const y of ys) {
      const L = (texto.get(y) || "").trim();

      // caixa da empresa: "15 - MOMENTUM EMPREENDIMENTOS IMOBILIARIOS LTDA.",
      // sempre antes da faixa de datas. O código tem no máximo 4 dígitos, o que
      // separa essa linha das solicitações (S.N tem 7).
      if (!meta.empresa && !meta.dataInicio) {
        const me = L.match(/^(\d{1,4})\s*-\s*([A-Za-zÀ-ÿ][^\n]*)$/);
        if (me) { meta.empresa = L; meta.empresaCodigo = me[1]; meta.empresaNome = me[2].trim(); }
      }
      if (!meta.dataInicio) {
        // formato antigo imprime "...CORRENTE - 01/09/2026 a 18/09/2026" (a
        // minúsculo); o formato novo (ver comentário no topo do arquivo)
        // também mudou ISSO sem aviso — "...CORRENTE - 22/09/2026 A
        // 22/09/2026", A maiúsculo — achado em 23/09/2026 depois que um
        // relatório do formato novo saiu com "sem data" na lista de
        // conferências (a extração inteira funcionava, só a DATA que não
        // batia com o regex). [Aa] cobre os dois sem exigir /i no resto.
        const m = L.match(/CORRENTE - (\d{2}\/\d{2}\/\d{4}) [Aa] (\d{2}\/\d{2}\/\d{4})/);
        if (m) { meta.dataInicio = m[1]; meta.dataFim = m[2]; }
      }
      if (!meta.emitidoEm && /^\d{2}:\d{2}:\d{2}$/.test(L)) meta.emitidoEm = L;

      if (ehCabecalhoBanco(L)) {
        banco = L; ignorar.add(y);
        continue;
      }
      if (TIPOS.includes(L)) { tipo = TIPO_CURTO[L] || L; ignorar.add(y); continue; }
      // o valor numérico do total é lido no passe tolerante, acima
      // fragmentos de linha de total que caem em outro bucket de y
      if (/Total ==>|Total por|Solicitação\(ões\)/.test(L) || /^R\$\s*[\d.,]+$/.test(L)) { ignorar.add(y); paradas.push(y); continue; }
      if (RUIDO.some((p) => L.startsWith(p))) { ignorar.add(y); continue; }
      // a caixa da empresa se repete no alto de cada página
      if (meta.empresa && L === meta.empresa) { ignorar.add(y); continue; }
      if (/^S\.N\b/.test(L)) {
        ignorar.add(y);
        if (novo) {
          // formato novo: lê as colunas do próprio cabeçalho da seção — o
          // cabeçalho pode vir fragmentado em vários y (mesmo problema da
          // linha de total), junta tudo num raio de 6pt ao redor desta linha.
          const proximas = palavras.filter((pp) => Math.abs(pp.y - y) <= 6);
          const resultado = detectarColunas(proximas, palavras, y);
          colunasAtuais = resultado.colunas;
          // Ignora só as linhas que de fato compuseram o cabeçalho (o
          // próprio y mais qualquer sublinha quebrada, como o "DISPÊNDIO"
          // que às vezes desce para a linha de baixo — ver detectarColunas).
          // Uma janela fixa (ex.: y+14, usada no formato antigo) já chegou a
          // engolir por engano a primeira linha de dado real quando ela
          // nasce só ~12pt abaixo do cabeçalho, no formato novo —
          // derrubando a solicitação inteira em silêncio.
          for (const yy of resultado.usedYs) ignorar.add(yy);
        } else {
          // formato antigo: tabela fixa por família, decidida pela posição
          // do "S.N" — inalterado desde a versão já validada em produção.
          const sn = linhas.get(y).find((p) => p.texto === "S.N");
          familia = sn && sn.x >= 31 ? "A" : "B";
          colunasAtuais = COLUNAS[familia];
          for (const yy of ys) if (yy >= y - 25 && yy <= y + 14) ignorar.add(yy);
        }
        continue;
      }
      // Âncora da solicitação: o S.N, na primeira coluna. Só considera se já
      // vimos um cabeçalho nesta seção (sempre vem antes, no relatório real).
      if (!colunasAtuais) continue;
      const fimSN = colunasAtuais.find(([nome]) => nome === "sn")[2];
      if (linhas.get(y).some((p) => /^\d{6,8}$/.test(p.texto) && p.x < fimSN))
        ancoras.push({ y, banco, tipo, colunas: colunasAtuais });
    }

    ancoras.forEach((a, i) => {
      const ini = a.y - 6;
      let fim = i + 1 < ancoras.length ? ancoras[i + 1].y - 6 : Infinity;
      const parada = paradas.filter((p) => p > a.y).sort((u, v) => u - v)[0];
      if (parada !== undefined) fim = Math.min(fim, parada - 2);
      const buf = {};
      for (const y of ys) {
        if (y < ini || y >= fim || ignorar.has(y)) continue;
        for (const p of linhas.get(y)) {
          const c = coluna(a.colunas, p.x);
          (buf[c] = buf[c] || []).push({ y, x: p.x, t: p.texto });
        }
      }
      const reg = { bancoDebitado: a.banco, tipo: a.tipo };
      for (const [nome] of a.colunas) {
        const partes = (buf[nome] || []).sort((u, v) => u.y - v.y || u.x - v.x);
        reg[nome] = partes.map((p) => p.t).join(" ").trim();
      }
      if (/^\d{6,8}$/.test(reg.sn)) { reg.valor = num(reg.valor); solicitacoes.push(reg); }
    });
  }
  if (blocoAtual.length) blocos.push(blocoAtual);

  for (const s of solicitacoes) s.alertas = alertas(s);
  marcarRepetidos(solicitacoes);

  // confere o extraído contra os totais impressos no rodapé de cada seção
  let qtdRel = 0, valorRel = 0;
  const totaisLidos = blocos.reduce((a, b) => a + b.length, 0);
  for (let b of blocos) {
    if (b.length > 1) {
      const q = b.slice(0, -1).reduce((a, x) => a + x[0], 0);
      const v = b.slice(0, -1).reduce((a, x) => a + x[1], 0);
      if (b[b.length - 1][0] === q && Math.abs(b[b.length - 1][1] - v) < 0.01) b = b.slice(0, -1);
    }
    qtdRel += b.reduce((a, x) => a + x[0], 0);
    valorRel += b.reduce((a, x) => a + x[1], 0);
  }
  const valorExtraido = Math.round(solicitacoes.reduce((a, s) => a + (s.valor || 0), 0) * 100) / 100;
  valorRel = Math.round(valorRel * 100) / 100;

  return {
    meta, solicitacoes,
    validacao: {
      qtdExtraida: solicitacoes.length, valorExtraido,
      qtdRelatorio: qtdRel, valorRelatorio: valorRel,
      totaisLidos, blocos: blocos.length,
      confere: solicitacoes.length === qtdRel && Math.abs(valorExtraido - valorRel) < 0.01,
    },
  };
}

const docCadastrado = (t) => (t.match(/Notas Fiscais:\s*(\d+)/) || [])[1] || null;
const nfsCitadas = (t) => [...new Set([...t.matchAll(/\bNF\.?\s*n?[ºo]?\s*(\d+)/g)].map((m) => m[1]))];

/** Só o que é objetivo e verificável. O julgamento é do conferente. */
function alertas(s) {
  const out = [];
  if (s.solicitante.trim() && s.solicitante.trim() === s.competente.trim())
    out.push("Solicitante e Competente são a mesma pessoa");
  const doc = docCadastrado(s.destinacao), nfs = nfsCitadas(s.destinacao);
  if (doc && nfs.length && nfs.every((v) => parseInt(v, 10) !== parseInt(doc, 10)))
    out.push(`Documento cadastrado (${doc}) diverge da NF citada (${nfs.join("/")})`);
  if (!s.poder.trim()) out.push("Poder não informado");
  // Pedido do Rocha em 23/09/2026: alertar sempre que o CNPJ do favorecido
  // for de uma empresa do próprio grupo (ver grupoEconomico.js) — roda pra
  // toda solicitação, PDF ou Excel, porque as duas passam por alertas(s).
  const grupo = verificarGrupoEconomico(s.cpfCnpj);
  if (grupo) out.push(grupo);
  return out;
}

const numeroNF = (v) => { const n = parseInt(v, 10); return isNaN(n) ? null : String(n); };

/**
 * Cruza as notas fiscais do lote, por fornecedor.
 *
 * Duas situações distintas, que não podem virar o mesmo alerta:
 *
 *  1. O MESMO documento cadastrado em duas solicitações do mesmo CNPJ.
 *     Duplicidade de pagamento até prova em contrário.
 *
 *  2. A NF citada na destinação de uma solicitação é a nota de OUTRA
 *     solicitação do mesmo fornecedor — foi assim que a 1452750 do
 *     relatório da SLIM (doc 59053) apareceu citando a NF 115157, que é
 *     o documento da 1452749. Aqui não há conclusão: pode ser erro de
 *     digitação na destinação, referência legítima ou pagamento em
 *     duplicidade. É sinal para o conferente olhar, não veredito.
 *
 * O cruzamento é sempre dentro do mesmo CPF/CNPJ: numeração de nota fiscal
 * é por emitente, então o mesmo número em fornecedores diferentes é
 * coincidência, não indício.
 */
function marcarRepetidos(solicitacoes) {
  const comoDoc = new Map();     // "cnpj|nf" -> Set de S.N que têm essa NF cadastrada
  const emQualquer = new Map();  // "cnpj|nf" -> Set de S.N que citam ou cadastram essa NF
  const envolvidas = new Map();  // S.N -> Set de NFs da solicitação
  const junta = (mapa, k, sn) => { if (!mapa.has(k)) mapa.set(k, new Set()); mapa.get(k).add(sn); };

  for (const s of solicitacoes) {
    const doc = numeroNF(docCadastrado(s.destinacao) || "");
    const nfs = new Set(nfsCitadas(s.destinacao).map(numeroNF).filter(Boolean));
    if (doc) { nfs.add(doc); junta(comoDoc, `${s.cpfCnpj}|${doc}`, s.sn); }
    envolvidas.set(s.sn, nfs);
    for (const nf of nfs) junta(emQualquer, `${s.cpfCnpj}|${nf}`, s.sn);
  }

  for (const s of solicitacoes) {
    const doc = numeroNF(docCadastrado(s.destinacao) || "");
    for (const nf of envolvidas.get(s.sn) || []) {
      const k = `${s.cpfCnpj}|${nf}`;
      const mesmoDoc = [...(comoDoc.get(k) || [])].filter((sn) => sn !== s.sn);
      if (doc === nf && mesmoDoc.length) {
        s.alertas.push(`Documento ${nf} já cadastrado na(s) solicitação(ões) ${mesmoDoc.join(", ")}`);
        continue;
      }
      const outras = [...(emQualquer.get(k) || [])].filter((sn) => sn !== s.sn);
      if (outras.length)
        s.alertas.push(`NF ${nf} citada aqui também aparece na(s) solicitação(ões) ${outras.join(", ")}`);
    }
  }
}

// ---------------------------------------------------------------------
// FORMATO EXCEL (.xlsx) — alternativa ao PDF para o mesmo relatório.
//
// Por que preferir o Excel ao PDF quando os dois estão disponíveis:
//  - VALOR sai em célula numérica nativa — elimina de vez a ambiguidade
//    de separador de milhar/decimal que já causou dois bugs diferentes no
//    caminho do PDF (ver comentário no topo do arquivo).
//  - as colunas são lidas pelo TEXTO do cabeçalho de cada seção, não por
//    posição em pixel — imune à classe inteira de bug de "coluna vazando
//    pra vizinha" ou "família invertida" que consumiu a maior parte do
//    conserto do PDF.
//  - não há fragmentação de linha por quebra de página — cada solicitação
//    é uma linha de planilha só, ponto.
//
// A DIFERENÇA que importa: o PDF sai sempre filtrado por UMA empresa (a
// caixa "NN - NOME DA EMPRESA" no cabeçalho vira meta.empresaCodigo). O
// Excel sai com "Filial: 0 - Todos" — TODAS as empresas do grupo no mesmo
// arquivo, confirmado com o Rocha. A separação entre uma empresa e outra,
// dentro do arquivo, é a MESMA linha de cabeçalho de banco que o PDF já usa
// pra separar blocos (ehCabecalhoBanco, acima): cada bloco de solicitações
// começa com uma linha como
//
//   BANCO ITAU S/A Ag. 0262 - C/C 62.222-8
//
// e o número da conta ali identifica a empresa via empresaPorConta()
// (contas.js, tabela extraída da planilha de contas que o Rocha mandou).
// Por isso cada solicitação carrega sua PRÓPRIA empresa/empresaCodigo, em
// vez de um único meta.empresaCodigo por documento como no PDF — ver
// campos `empresa`/`empresaAba`/`empresaCodigo` no registro devolvido.
//
// Conferido por amostragem em 22/09/2026 (relatório real de um dia, 129
// solicitações, 11 blocos de conta, 7 empresas): as 3 variações de
// cabeçalho de coluna do relatório têm as colunas S.N até Destinação
// sempre nas mesmas posições relativas — só a ÚLTIMA coluna antes de
// "Poder do Dispêndio" muda de nome (e por vezes de posição):
//   família A   (TED / TRANSFERÊNCIA): Banco, Agência, Conta Bancária
//   família B   (DÉBITO EM CONTA / BOLETO): Observação
//   família PIX (PIX): Chave PIX
// Por isso as colunas são lidas pelo NOME do rótulo em cada seção — mesmo
// raciocínio do formato novo do PDF (detectarColunas), só que mais simples
// aqui porque é célula de planilha, não geometria de texto em pixel.
// ---------------------------------------------------------------------

const ROTULOS_COLUNA_EXCEL = {
  "S.N": "sn",
  "FILIAL PAGAMENTO": "filial",
  "VALOR": "valor",
  "SOLICITANTE": "solicitante",
  "COMPETENTE": "competente",
  "PODER": "poder",
  "CPF / CNPJ": "cpfCnpj",
  "CPF/CNPJ": "cpfCnpj",
  "FAVORECIDO / REAL FAVORECIDO": "favorecido",
  "FAVORECIDO": "favorecido",
  "DESTINAÇÃO": "destinacao",
  "BANCO": "banco",
  "AGÊNCIA": "agencia",
  "CONTA BANCÁRIA": "contaBancaria",
  "OBSERVAÇÃO": "observacao",
  "CHAVE PIX": "chavePix",
  "PODER DO DISPÊNDIO": "poderDispendio",
};

// Só maiúsculas + espaço normalizado — SEM tirar acento: os rótulos do
// relatório são um conjunto fixo e conhecido (visto por amostragem), e as
// chaves de ROTULOS_COLUNA_EXCEL já estão escritas com o acento certo. Tirar
// acento aqui exigiria tirar dos dois lados (senão "DESTINAÇÃO" normalizado
// vira "DESTINACAO" e não bate com a chave acentuada do dicionário) — mais
// simples manter os dois lados iguais.
const normalizarRotuloExcel = (t) => String(t || "").toUpperCase().replace(/\s+/g, " ").trim();

// "BANCO ITAU S/A Ag. 0262 - C/C 62.222-8" -> "62.222-8" (só o número da
// conta importa pra empresaPorConta — ver contas.js: a normalização de lá
// já tira ponto/espaço, então "62.222-8" e "62222 - 8" caem na mesma chave).
const RE_CABECALHO_CONTA_EXCEL = /Ag\.\s*[0-9.\-\sxX]+?\s*-\s*C\/C\s*([0-9.\-\sxX]+)\s*$/i;

/** Valor bruto de uma célula (null se vazia/inexistente). */
function celulaExcel(XLSX, ws, r, c) {
  if (c === undefined) return null;
  const cel = ws[XLSX.utils.encode_cell({ r, c })];
  return cel ? cel.v : null;
}

const textoCelulaExcel = (XLSX, ws, r, c) => {
  const v = celulaExcel(XLSX, ws, r, c);
  return v == null ? "" : String(v).trim();
};

/**
 * Lê o cabeçalho de colunas de UMA seção (a linha onde a célula "S.N"
 * aparece), pelo TEXTO de cada rótulo — não por índice fixo, porque a
 * posição das colunas muda entre as 3 famílias (ver comentário acima).
 * Devolve {colunas: {campo: índice da coluna}, familia: "A"|"B"|"PIX"}.
 */
function detectarColunasExcel(XLSX, ws, linha, colMax) {
  const colunas = {};
  let familia = null;
  for (let c = 0; c <= colMax; c++) {
    const v = celulaExcel(XLSX, ws, linha, c);
    if (typeof v !== "string") continue;
    const campo = ROTULOS_COLUNA_EXCEL[normalizarRotuloExcel(v)];
    if (!campo) continue;
    if (colunas[campo] === undefined) colunas[campo] = c;
    if (campo === "banco") familia = "A";
    else if (campo === "observacao") familia = "B";
    else if (campo === "chavePix") familia = "PIX";
  }
  const obrigatorios = ["sn","filial","valor","solicitante","competente","poder","cpfCnpj","favorecido","destinacao","poderDispendio"];
  const faltando = obrigatorios.filter((campo) => colunas[campo] === undefined);
  if (faltando.length)
    throw new Error(`Cabeçalho do relatório Excel não reconhecido — colunas não encontradas: ${faltando.join(", ")}. O layout da planilha mudou?`);
  return { colunas, familia };
}

/** Campo `complemento`: mesmo papel que no PDF (texto livre, não usado em mais nada do site — só guardado). */
function montarComplementoExcel(XLSX, ws, r, colunas, familia) {
  if (familia === "A") {
    const partes = [
      textoCelulaExcel(XLSX, ws, r, colunas.banco),
      textoCelulaExcel(XLSX, ws, r, colunas.agencia),
      textoCelulaExcel(XLSX, ws, r, colunas.contaBancaria),
    ];
    return partes.filter(Boolean).join(" - ");
  }
  if (familia === "B") return textoCelulaExcel(XLSX, ws, r, colunas.observacao);
  if (familia === "PIX") return textoCelulaExcel(XLSX, ws, r, colunas.chavePix);
  return "";
}

/**
 * Extração do relatório "Solicitações de Pagamentos por Débito em Conta
 * Corrente" (SCK) a partir do EXPORT EM EXCEL (.xlsx) do mesmo relatório —
 * ver o comentário logo acima para o porquê de preferir este caminho ao
 * PDF e para a diferença que importa (empresa por SOLICITAÇÃO, não por
 * documento inteiro).
 *
 *   const dados = parseRelatorioExcel(arrayBuffer, XLSX);
 *
 * XLSX é o SheetJS (vendor/xlsx.full.min.js), carregado pelo chamador —
 * mesmo padrão de pdfjsLib em parseRelatorio(). Síncrono (SheetJS não
 * precisa de await para ler um ArrayBuffer já em memória).
 */
export function parseRelatorioExcel(arrayBuffer, XLSX) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws["!ref"]) throw new Error("Planilha vazia ou em formato não reconhecido.");
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const colMax = range.e.c;

  const meta = {
    formato: "excel",
    empresa: null, empresaCodigo: null, empresaNome: null,
    empresas: [], multiEmpresa: false,
    filialEscopo: null, dataInicio: null, dataFim: null, emitidoEm: null,
  };
  const solicitacoes = [];
  const blocos = [];
  let blocoAtual = [];
  let contaAtual = null;   // texto cru do cabeçalho de banco ("BANCO X Ag. Y - C/C Z") — vira bancoDebitado
  let empresaAtual = null; // resultado de empresaPorConta() pro bloco atual (ou null se conta não cadastrada)
  let tipoAtual = null;
  let colunasAtuais = null;
  let familiaAtual = null;
  const empresasVistas = new Map(); // nome curto -> true, só pra listar em meta.empresas

  for (let r = range.s.r; r <= range.e.r; r++) {
    // Coluna B: onde o gerador sempre imprime as linhas "estruturais"
    // (título, "Filial:", cabeçalho de banco, tipo, "S.N" e "Total ==>") —
    // confirmado em todas as seções do relatório de amostra. As colunas de
    // DADO de cada solicitação, por outro lado, são lidas dinamicamente
    // pelo cabeçalho (detectarColunasExcel), não fixadas em B.
    const b = textoCelulaExcel(XLSX, ws, r, 1);
    if (!b) continue; // linha em branco (separador visual)

    if (ehCabecalhoBanco(b)) {
      // O bloco que está sendo fechado agora pertence ao empresaAtual de
      // ANTES desta linha (a conta do cabeçalho anterior) — por isso marca
      // com empresaAtual antes de reatribuí-lo duas linhas abaixo. É o que
      // permite separar o total impresso de cada empresa na hora de validar
      // (ver validacaoPorEmpresa, mais abaixo) — cada bloco é sempre de UMA
      // conta bancária só, e cada conta é de uma empresa só.
      if (blocoAtual.length)
        blocos.push({ empresa: empresaAtual ? empresaAtual.empresa : "", totais: blocoAtual });
      blocoAtual = [];
      contaAtual = b;
      const m = RE_CABECALHO_CONTA_EXCEL.exec(b);
      empresaAtual = m ? empresaPorConta(m[1]) : null;
      if (empresaAtual) empresasVistas.set(empresaAtual.empresa, true);
      colunasAtuais = null;
      continue;
    }
    if (TIPOS.includes(b)) { tipoAtual = TIPO_CURTO[b] || b; colunasAtuais = null; continue; }
    if (b === "S.N") {
      const { colunas, familia } = detectarColunasExcel(XLSX, ws, r, colMax);
      colunasAtuais = colunas;
      familiaAtual = familia;
      continue;
    }
    const mt = b.match(RE_TOTAL);
    if (mt) { blocoAtual.push([parseInt(mt[1].replace(/\./g, ""), 10), num(mt[2])]); continue; }
    if (b.startsWith("Filial:")) { if (!meta.filialEscopo) meta.filialEscopo = b.replace(/^Filial:\s*/, ""); continue; }
    if (b.startsWith("SOLICITAÇÕES DE")) {
      const md = b.match(/(\d{2}\/\d{2}\/\d{4})\s+at[ée]\s+(\d{2}\/\d{2}\/\d{4})/i);
      if (md) { meta.dataInicio = md[1]; meta.dataFim = md[2]; }
      continue;
    }
    if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(b)) { meta.emitidoEm = b; continue; }

    // Não é nenhuma linha estrutural conhecida — só pode ser dado de
    // solicitação, e só faz sentido se já vimos um cabeçalho de coluna
    // nesta seção (sempre vem antes, no relatório real).
    if (!colunasAtuais) continue;
    const snBruto = celulaExcel(XLSX, ws, r, colunasAtuais.sn);
    if (typeof snBruto !== "number" || !Number.isInteger(snBruto) || snBruto <= 0) continue;

    const valorBruto = celulaExcel(XLSX, ws, r, colunasAtuais.valor);
    const poderBruto = celulaExcel(XLSX, ws, r, colunasAtuais.poder);
    solicitacoes.push({
      sn: String(snBruto),
      filial: textoCelulaExcel(XLSX, ws, r, colunasAtuais.filial),
      valor: typeof valorBruto === "number" ? valorBruto : num(valorBruto),
      solicitante: textoCelulaExcel(XLSX, ws, r, colunasAtuais.solicitante),
      competente: textoCelulaExcel(XLSX, ws, r, colunasAtuais.competente),
      poder: typeof poderBruto === "number" ? String(Math.trunc(poderBruto)) : textoCelulaExcel(XLSX, ws, r, colunasAtuais.poder),
      cpfCnpj: textoCelulaExcel(XLSX, ws, r, colunasAtuais.cpfCnpj),
      favorecido: textoCelulaExcel(XLSX, ws, r, colunasAtuais.favorecido),
      destinacao: textoCelulaExcel(XLSX, ws, r, colunasAtuais.destinacao),
      complemento: montarComplementoExcel(XLSX, ws, r, colunasAtuais, familiaAtual),
      poderDispendio: textoCelulaExcel(XLSX, ws, r, colunasAtuais.poderDispendio),
      bancoDebitado: contaAtual,
      tipo: tipoAtual,
      empresa: empresaAtual ? empresaAtual.empresa : null,
      empresaAba: empresaAtual ? empresaAtual.aba : null,
      empresaCodigo: empresaAtual ? (empresaAtual.codigo || null) : null,
    });
  }
  if (blocoAtual.length)
    blocos.push({ empresa: empresaAtual ? empresaAtual.empresa : "", totais: blocoAtual });

  for (const s of solicitacoes) {
    s.alertas = alertas(s);
    if (s.bancoDebitado && !s.empresa)
      s.alertas.push(`Conta de origem não cadastrada em contas.js ("${s.bancoDebitado}") — empresa não identificada automaticamente, confira manualmente`);
  }
  marcarRepetidos(solicitacoes);

  meta.empresas = [...empresasVistas.keys()];
  meta.multiEmpresa = meta.empresas.length > 1;
  if (meta.empresas.length === 1) {
    meta.empresaNome = meta.empresas[0];
    const umaComCodigo = solicitacoes.find((s) => s.empresa === meta.empresaNome && s.empresaCodigo);
    meta.empresaCodigo = umaComCodigo ? umaComCodigo.empresaCodigo : null;
  }
  // meta.empresa é o rótulo de EXIBIÇÃO (conferencia.js: nomeEmpresa/
  // codEmpresa) — sempre preenchido, mesmo com várias empresas juntas, pra
  // nunca aparecer "Empresa não identificada" quando na verdade dá, sim,
  // pra dizer quais são. Mesmo formato "CÓDIGO - NOME" que o PDF usa (ver
  // meta.empresa em parseRelatorio, acima) — mas ali o texto vem PRONTO do
  // próprio SCK (impresso na caixa do cabeçalho); aqui ele é MONTADO a
  // partir do código + da relação oficial em contas.js (rotuloEmpresa()),
  // porque o Excel não traz esse texto em lugar nenhum do arquivo. Cai pro
  // nome curto de contas.js quando o código não tem nome oficial cadastrado.
  const rotuloDe = (nomeCurto) => {
    const cod = solicitacoes.find((s) => s.empresa === nomeCurto && s.empresaCodigo)?.empresaCodigo;
    return rotuloEmpresa(cod) || nomeCurto;
  };
  meta.empresa = meta.empresas.length === 0
    ? "Empresa não identificada"
    : meta.empresas.length === 1
    ? rotuloDe(meta.empresas[0])
    : `Múltiplas empresas (${meta.empresas.map(rotuloDe).join(", ")})`;

  // Confere o extraído contra os totais impressos no rodapé de cada bloco —
  // MESMA lógica do PDF (parseRelatorio, acima): cada bloco imprime um
  // total por subseção de tipo E um total agregado do bloco inteiro; se o
  // último total do bloco bater com a soma dos anteriores, é o agregado —
  // descarta antes de somar, senão os totais entram em dobro na conferência.
  //
  // somarBlocosImpressos() faz essa soma pra QUALQUER subconjunto de blocos
  // — usada abaixo tanto pro total do documento inteiro (validacao, como
  // sempre foi) quanto pro total de CADA empresa sozinha (validacaoPorEmpresa,
  // novo em 23/09/2026): como cada bloco já sai marcado com a empresa da
  // conta bancária dele (empresaPorConta, acima), dá pra somar só os blocos
  // de uma empresa e comparar com as solicitações extraídas SÓ dela — uma
  // conferência de verdade por empresa, não a mesma validação global
  // repetida pra todas quando o Excel divide numa conferência por empresa
  // (ver dividirPorEmpresa em conferencia.js).
  function somarBlocosImpressos(subset) {
    let qtd = 0, valor = 0, totaisLidos = 0;
    for (let bloco of subset) {
      let b = bloco.totais;
      totaisLidos += b.length;
      if (b.length > 1) {
        const q = b.slice(0, -1).reduce((a, x) => a + x[0], 0);
        const v = b.slice(0, -1).reduce((a, x) => a + x[1], 0);
        if (b[b.length - 1][0] === q && Math.abs(b[b.length - 1][1] - v) < 0.01) b = b.slice(0, -1);
      }
      qtd += b.reduce((a, x) => a + x[0], 0);
      valor += b.reduce((a, x) => a + x[1], 0);
    }
    return { qtd, valor: Math.round(valor * 100) / 100, totaisLidos };
  }
  function validar(solicitacoesSubset, blocosSubset) {
    const { qtd: qtdRelatorio, valor: valorRelatorio, totaisLidos } = somarBlocosImpressos(blocosSubset);
    const valorExtraido = Math.round(solicitacoesSubset.reduce((a, s) => a + (s.valor || 0), 0) * 100) / 100;
    return {
      qtdExtraida: solicitacoesSubset.length, valorExtraido,
      qtdRelatorio, valorRelatorio,
      totaisLidos, blocos: blocosSubset.length,
      confere: solicitacoesSubset.length === qtdRelatorio && Math.abs(valorExtraido - valorRelatorio) < 0.01,
    };
  }

  const validacao = validar(solicitacoes, blocos);

  // nome curto ("" = conta não cadastrada em contas.js) -> validação feita
  // só com as solicitações e os blocos (contas bancárias) daquela empresa.
  const validacaoPorEmpresa = {};
  for (const chave of new Set(blocos.map((b) => b.empresa))) {
    validacaoPorEmpresa[chave] = validar(
      solicitacoes.filter((s) => (s.empresa || "") === chave),
      blocos.filter((b) => b.empresa === chave),
    );
  }

  return { meta, solicitacoes, validacao, validacaoPorEmpresa };
}
