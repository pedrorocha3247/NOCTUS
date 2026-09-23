/**
 * CNPJs de empresas do MESMO GRUPO econômico (Momentum/Grupo Kasil) — usado
 * pra alertar quando o FAVORECIDO de uma solicitação de pagamento é, ele
 * mesmo, uma empresa do grupo. Pedido do Rocha em 23/09/2026: "se constar
 * qualquer um desse cnpj, deve alertar" — não é necessariamente irregular
 * (pode ser um repasse legítimo entre empresas do grupo), mas é bem
 * diferente de pagar um fornecedor externo, e merece atenção redobrada na
 * conferência — daí o apontamento automático, igual aos outros que este
 * site já gera (ver alertas(), em parser.js).
 *
 * Duas listas, as duas mandadas pelo Rocha em 23/09/2026:
 *
 *  1) Empreendimentos RSC/NV/SBRR — filiais do CNPJ raiz da Momentum
 *     (47.686.555), uma por empreendimento imobiliário, cada uma com seu
 *     gerente e município responsável.
 *  2) As mesmas empresas do grupo que este site já confere pelo código (13
 *     a 96 — ver NOME_OFICIAL_POR_CODIGO em contas.js), aqui com o CNPJ do
 *     lado de VENDAS.
 *
 * ATENÇÃO — duas divergências encontradas ao cruzar esta lista com
 * contas.js, registradas aqui pra não se perderem, mas NÃO corrigidas
 * sozinho: quem decide qual está certo é o Rocha.
 *   - código 27 (Abrasma): aqui o CNPJ é "03.945.625/0001-14", mas em
 *     CONTAS_EMPRESA (contas.js) está cadastrado "03.954.217/0001-29" —
 *     são dígitos diferentes, não é só formatação. Pode ser CNPJ de
 *     filial/matriz diferentes, ou um dos dois ter erro de digitação.
 *   - código 28: aqui o nome do empreendimento é "MODO", mas o nome oficial
 *     em contas.js é "Realiza — Soluções Imobiliárias e Financeiras Ltda."
 *     — mesmo CNPJ (67.648.402/0001-78) nos dois, então "MODO" é
 *     provavelmente nome fantasia/comercial, não uma empresa diferente.
 *
 * O código "91" aparece duas vezes de propósito: duas pessoas jurídicas
 * distintas (CNPJs diferentes — Posto Santa Bárbara e Rubens Meneghetti)
 * foram informadas pelo Rocha sob o mesmo código.
 */
export const CNPJS_GRUPO = [
  // --- Empreendimentos (filiais da Momentum) — RSC / NV / SBRR ---
  { rotulo: "RSC I", cnpj: "47.686.555/0009-50", gerente: "Fernando Prates", municipio: "Arandu" },
  { rotulo: "RSC II", cnpj: "47.686.555/0010-93", gerente: "Marcelo Augusto Florencio", municipio: "Itaí" },
  { rotulo: "RSC III", cnpj: "47.686.555/0011-74", gerente: "Marcelo Augusto Florencio", municipio: "Itaí" },
  { rotulo: "RSC IV", cnpj: "47.686.555/0012-55", gerente: "Fernando Prates", municipio: "Arandu" },
  { rotulo: "RSC V", cnpj: "47.686.555/0013-36", gerente: "Jefferson de Moura", municipio: "Paranapanema" },
  { rotulo: "RSC XIII", cnpj: "47.686.555/0003-64", gerente: "Vitor Tomazoli", municipio: "Paranapanema" },
  { rotulo: "NV I", cnpj: "47.686.555/0014-17", gerente: "Roberval Junior", municipio: "Porangaba" },
  { rotulo: "NV II", cnpj: "47.686.555/0015-06", gerente: "Jobson Alves", municipio: "Pardinho" },
  { rotulo: "SBRR", cnpj: "47.686.555/0008-79", gerente: "Eric Rodolfo Valentim", municipio: "Águas de Sta. Bárbara" },
  // --- Empresas do grupo (mesmos códigos de NOME_OFICIAL_POR_CODIGO em contas.js) — CNPJ do lado de vendas ---
  { rotulo: "13 - PRAIA VERDE", cnpj: "68.199.298/0001-44" },
  { rotulo: "15 - MOMENTUM MATRIZ", cnpj: "47.686.555/0001-00" },
  { rotulo: "16 - SLIM MATRIZ", cnpj: "54.363.213/0001-07" },
  { rotulo: "26 - M3", cnpj: "46.460.117/0001-59" },
  { rotulo: "27 - ABRASMA", cnpj: "03.945.625/0001-14" },
  { rotulo: "28 - MODO", cnpj: "67.648.402/0001-78" },
  { rotulo: "30 - KASIL", cnpj: "67.550.996/0001-80" },
  { rotulo: "40 - RVM", cnpj: "67.648.733/0001-08" },
  { rotulo: "70 - M5", cnpj: "25.247.689/0001-84" },
  { rotulo: "90 - IRM", cnpj: "05.161.107/0001-35" },
  { rotulo: "91 - POSTO SB", cnpj: "27.510.105/0001-47" },
  { rotulo: "91 - RUBENS MENEGHETTI", cnpj: "08.085.265/0001-41" },
  { rotulo: "95 - PICK MONEY", cnpj: "39.831.735/0001-00" },
  { rotulo: "96 - MMH HOLDING", cnpj: "48.958.204/0001-66" },
];

/** Tira tudo que não é dígito — o CNPJ do relatório pode vir formatado (12.345.678/0001-99) ou não. */
const normalizarCnpj = (s) => String(s || "").replace(/\D/g, "");

const CNPJS_GRUPO_POR_DIGITOS = new Map(CNPJS_GRUPO.map((e) => [normalizarCnpj(e.cnpj), e]));

/**
 * Verifica se o CNPJ do favorecido é de uma empresa do próprio grupo —
 * devolve o texto do apontamento, ou null quando não há nada a apontar.
 * Comparação por dígitos (normalizarCnpj), pra não depender do CNPJ do
 * relatório vir formatado exatamente igual ao cadastrado aqui.
 */
export function verificarGrupoEconomico(cpfCnpj) {
  const alvo = normalizarCnpj(cpfCnpj);
  if (!alvo) return null;
  const achado = CNPJS_GRUPO_POR_DIGITOS.get(alvo);
  if (!achado) return null;
  return `Favorecido é empresa do grupo (${achado.rotulo}) — confira se é repasse intercompany`;
}
