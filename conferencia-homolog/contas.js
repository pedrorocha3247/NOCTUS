/**
 * Tabela conta bancária → empresa, extraída da planilha "Constas de banco
 * empresas.xlsx" que o Rocha enviou em 22/09/2026 (cada bloco de colunas da
 * planilha é o "CONTAS <EMPRESA>" de uma das 13 empresas do grupo).
 *
 * Serve para identificar a empresa dona de cada bloco de solicitações no
 * relatório em EXCEL de Solicitações de Pagamentos: ao contrário do PDF
 * (que já sai filtrado por uma empresa, com a caixa "NN - NOME DA EMPRESA"
 * no cabeçalho), o Excel sai com "Filial: 0 - Todos" — TODAS as empresas
 * juntas num arquivo só. O que separa um bloco de solicitações do outro,
 * dentro do Excel, é a mesma linha de cabeçalho de banco que o parser de
 * PDF já reconhece (ehCabecalhoBanco, em parser.js): algo como
 *
 *   BANCO ITAU S/A Ag. 0262 - C/C 62.222-8
 *
 * empresaPorConta() casa o número da conta dessa linha com esta tabela.
 *
 * O número da conta SOZINHO já é suficiente como chave: conferido em
 * 22/09/2026, nenhuma das 33 contas cadastradas aqui é compartilhada por
 * duas empresas — não há colisão. Banco e agência ficam guardados só como
 * documentação / conferência manual, não entram na busca.
 *
 * "empresa" é o nome curto e "aba" é o código de aba usados em
 * ABAS_EMPRESA (conferencia.js) — mantidos iguais de propósito, para as
 * duas tabelas continuarem falando a mesma língua.
 *
 * "codigo" é o mesmo código de empresa que o parser de PDF lê da caixa
 * "NN - NOME DA EMPRESA" no cabeçalho (meta.empresaCodigo) e que
 * poderes.js usa como chave de PODERES — é o que permite conferir o limite
 * de poder de uma solicitação vinda do Excel. Confirmado pelo Rocha em
 * 22/09/2026, direto da relação oficial de empresas:
 *
 *   13 Praia Verde Empreendimentos e Participações S.A.
 *   15 Momentum Empreendimentos Imobiliários Ltda.
 *   16 Slim
 *   26 M3 Assets Holding S.A.
 *   27 Associação Bras. Def. do Desenv. Sust. do M. Ambiente (Abrasma)
 *   28 Realiza - Soluções Imobiliárias e Financeiras Ltda.
 *   30 Kasil
 *   40 RVM
 *   70 M5
 *   90 IRM (Instituto)
 *   91 Posto
 *   95 Pick Money
 *   96 MMH
 *
 * Se um código errado fosse usado aqui, o site conferiria o poder da
 * solicitação contra o limite de OUTRA empresa, silenciosamente — por isso
 * nenhum "codigo" foi cadastrado sem confirmação direta do Rocha (a versão
 * anterior desta tabela tinha só a Momentum confirmada). Uma conta nova,
 * de uma empresa ainda não listada aqui, some do meio-termo: empresaPorConta()
 * continua devolvendo null pra ela até entrar na tabela — nunca "codigo"
 * adivinhado por semelhança de nome.
 *
 * ATENÇÃO — é dado cadastrado à mão a partir de uma planilha que o Rocha
 * mandou manualmente, não uma integração com o banco: não há como o site
 * perceber sozinho se uma empresa abrir conta nova, encerrar uma conta
 * existente ou mudar de agência. Se uma conta do relatório não aparecer
 * aqui, empresaPorConta() devolve null (silêncio, não erro) — o chamador
 * decide o que fazer (ver parseRelatorioExcel, em parser.js). Revise esta
 * tabela sempre que o Rocha mandar uma versão nova da planilha de contas.
 */
export const CONTAS_EMPRESA = [
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "J17 - 451", agencia: "0001", conta: "1123-4", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Banco do Brasil - 01", agencia: "2155", conta: "9868 -X", cnpj: "47.686.555/0003-64" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Bradesco - 237", agencia: "3395", conta: "37629-9", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Bradesco - 399", agencia: "134", conta: "356881-4", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Bradesco", agencia: "2496", conta: "6525-0", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Santander - 33", agencia: "0721", conta: "13000050-9", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Santander - 356", agencia: "3566", conta: "13000488-8", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Santander - 353", agencia: "2146", conta: "13000008-3", cnpj: "47.686.555/0001-00" },
  { empresa: "Momentum", aba: "MM", codigo: "15", banco: "Itaú - 341", agencia: "0262", conta: "62222-8", cnpj: "47.686.555/0001-00" },
  { empresa: "Instituto", aba: "IRM", codigo: "90", banco: "Itaú - 341", agencia: "0262", conta: "58703-3", cnpj: "05.161.107/0001-35" },
  { empresa: "M5", aba: "M5", codigo: "70", banco: "Itaú - 341", agencia: "7307", conta: "14982-8", cnpj: "25.247.689/0001-84" },
  { empresa: "Realiza", aba: "REALIZA", codigo: "28", banco: "Itaú - 341", agencia: "7307", conta: "46780-8", cnpj: "67.648.402/0001-78" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "J17 - 451", agencia: "0001", conta: "00042-1", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "Daycoval", agencia: "0001-9", conta: "703673-8", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "UBS - 15", agencia: "0001", conta: "458693-0", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "XP - 102", agencia: "0001", conta: "458693-0", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "Cred Suisses (UBS)", agencia: "0001", conta: "17651-5", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "Bradesco - 237", agencia: "3395", conta: "3890-3", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "Itaú - 341", agencia: "0262", conta: "84283-4", cnpj: "68.199.298/0001-44" },
  { empresa: "Praia Verde", aba: "PV", codigo: "13", banco: "Itaú - 341", agencia: "3001", conta: "16974-1", cnpj: "68.199.298/0001-44" },
  { empresa: "Slim", aba: "SLIM", codigo: "16", banco: "Itaú - 341", agencia: "0262", conta: "41382-6", cnpj: "54.363.213/0001-07" },
  { empresa: "RVM", aba: "RVM", codigo: "40", banco: "Itaú - 341", agencia: "0262", conta: "58697-7", cnpj: "67.648.733/0001-08" },
  { empresa: "RVM", aba: "RVM", codigo: "40", banco: "Caixa - 104", agencia: "4206", conta: "003 577580154-3", cnpj: "67.648.733/0001-08" },
  { empresa: "MMH", aba: "MMH", codigo: "96", banco: "Itaú - 341", agencia: "7307", conta: "51397-3", cnpj: "48.958.204/0001-66" },
  { empresa: "Pick Money", aba: "PKM", codigo: "95", banco: "Itaú - 341", agencia: "7307", conta: "41982-5", cnpj: "39.831.735/0001-00" },
  { empresa: "Pick Money", aba: "PKM", codigo: "95", banco: "Itaú - 341", agencia: "3001", conta: "17985-6", cnpj: "39.831.735/0001-00" },
  { empresa: "Pick Money", aba: "PKM", codigo: "95", banco: "Santander - 33", agencia: "0004", conta: "13013598-8", cnpj: "39.831.735/0001-00" },
  { empresa: "M3", aba: "M3", codigo: "26", banco: "Itaú - 341", agencia: "7307", conta: "51185-2", cnpj: "46.460.117/0001-59" },
  { empresa: "Kasil", aba: "Kasil", codigo: "30", banco: "Itaú - 341", agencia: "0262", conta: "58073-1", cnpj: "67.550.996/0001-80" },
  { empresa: "Kasil", aba: "Kasil", codigo: "30", banco: "Bradesco - 237", agencia: "3395", conta: "09840-0", cnpj: "67.550.996/0001-80" },
  { empresa: "Abrasma", aba: "ABRASMA", codigo: "27", banco: "Itaú - 341", agencia: "7307", conta: "44683-6", cnpj: "03.954.217/0001-29" },
  { empresa: "Posto", aba: "Posto", codigo: "91", banco: "Itaú - 341", agencia: "7307", conta: "14227-8", cnpj: "27.510.105/0001-47" },
  { empresa: "Posto", aba: "Posto", codigo: "91", banco: "Banco do Brasil - 01", agencia: "6752", conta: "12174-6", cnpj: "27.510.105/0001-47" },
];

/** Tira espaço (inclusive não separável), ponto de milhar e caixa — "62.222-8" e "62222 - 8" viram "62222-8". */
const normalizarConta = (s) => String(s || "").replace(/ /g, " ").replace(/[\s.]/g, "").toUpperCase();

/**
 * Acha a empresa dona de uma conta, pelo número da conta (agência é
 * ignorada na busca — ver o porquê no comentário da tabela acima).
 * Devolve {empresa, aba, codigo, banco, agencia, conta, cnpj} ou null
 * quando a conta não está cadastrada. As 13 empresas cadastradas aqui têm
 * "codigo" confirmado (ver lista no comentário da tabela); uma empresa
 * nova, ainda não cadastrada, simplesmente não aparece — devolve null.
 */
export function empresaPorConta(conta) {
  const alvo = normalizarConta(conta);
  if (!alvo) return null;
  return CONTAS_EMPRESA.find((c) => normalizarConta(c.conta) === alvo) || null;
}
