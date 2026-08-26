import { atraso } from '@/services/comuns/baseMock'

/**
 * Desenvolvimento — Alteração de Produtos (SB1010), Aprovação Comercial de
 * SPD (SKFT1 + WorkFlowSapComercial), Lote Piloto V2 e Solicitação de
 * Projeto ao Desenvolvimento (SPD) V2.
 */

export interface ResultadoDesenvolvimento {
  sucesso: boolean
  mensagem: string
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const CLIENTES_DES = [
  { cod: '001234', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.' },
  { cod: '002567', loja: '01', nome: 'BEBIDAS DO VALE LTDA' },
  { cod: '003891', loja: '01', nome: 'FRIGORIFICO SERRA AZUL LTDA' },
  { cod: '004125', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.' }
]

export function listarClientesDesenvolvimento() {
  return CLIENTES_DES
}

/* ------------------------------------------------------------------ *
 * Alteração de Produtos (SB1010.B1_MSBLQL)
 * ------------------------------------------------------------------ */

/** Status do produto no Protheus */
export const STATUS_PRODUTO = [
  { codigo: 'A', descricao: 'Aprovado' },
  { codigo: 'C', descricao: 'Cancelado' },
  { codigo: 'P', descricao: 'Projeto' },
  { codigo: 'V', descricao: 'Verificado' },
  { codigo: 'W', descricao: 'Validado' },
  { codigo: 'E', descricao: 'Em Espera' }
]

export function descricaoStatusProduto(codigo: string) {
  return STATUS_PRODUTO.find((s) => s.codigo === codigo)?.descricao ?? ''
}

export interface Produto {
  codigo: string
  descricao: string
  cliente: string
  status: string
}

const produtos: Produto[] = Array.from({ length: 40 }, (_, indice) => {
  const cliente = CLIENTES_DES[indice % CLIENTES_DES.length]
  const estilos = ['CAIXA MAÇÃ', 'CAIXA BEBIDA', 'CHAPA ONDULADA', 'BANDEJA HORTIFRUTI', 'CAIXA FRIGORIFICADA']

  return {
    codigo: `${String(400_000 + indice * 7)}`,
    descricao: `${estilos[indice % estilos.length]} ${1000 + indice * 13}X${800 + indice * 5}`,
    cliente: cliente.nome,
    status: STATUS_PRODUTO[indice % STATUS_PRODUTO.length].codigo
  }
})

export async function listarProdutos(termo: string): Promise<Produto[]> {
  const busca = termo.trim().toLowerCase()

  const resultado = produtos
    .filter((p) => !busca || p.codigo.includes(busca) || p.descricao.toLowerCase().includes(busca))
    .map((p) => ({ ...p }))

  return atraso(resultado, 400)
}

/** alterarStatus — o legado altera em lote e dispara e-mail para SK_MAILFTU */
export async function alterarStatusProdutos(
  codigos: string[],
  novoStatus: string
): Promise<ResultadoDesenvolvimento> {
  if (codigos.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos um produto.' }, 200)
  }
  if (!novoStatus) {
    return atraso({ sucesso: false, mensagem: 'Selecione o novo status.' }, 200)
  }

  for (const codigo of codigos) {
    const produto = produtos.find((p) => p.codigo === codigo)
    if (produto) produto.status = novoStatus
  }

  return atraso(
    {
      sucesso: true,
      mensagem: `${codigos.length} produto(s) alterado(s) para ${descricaoStatusProduto(novoStatus)}.`
    },
    600
  )
}

/* ------------------------------------------------------------------ *
 * Aprovação Comercial de SPD
 * ------------------------------------------------------------------ */

/** Motivos de reprovação do index_comercial.asp */
export const MOTIVOS_REPROVACAO_COMERCIAL = [
  { codigo: 'D', descricao: 'Emissão em Duplicidade' },
  { codigo: 'P', descricao: 'Preço não Aprovado' },
  { codigo: 'E', descricao: 'Descontinuidade Projeto' },
  { codigo: 'R', descricao: 'Sem Retorno do Comercial' },
  { codigo: 'N', descricao: 'Sem Retorno Cliente' },
  { codigo: 'O', descricao: 'Proposta Descartada' },
  { codigo: 'F', descricao: 'Cadastro Cliente Finalizado' },
  { codigo: 'I', descricao: 'Processo Interno SK Teste' },
  { codigo: 'T', descricao: 'Teste' }
]

/** Motivos de aprovação — classificam a oportunidade comercial */
export const MOTIVOS_APROVACAO_COMERCIAL = [
  { codigo: '1', descricao: 'Cliente Novo - Plano' },
  { codigo: '2', descricao: 'Cliente Novo - Explorar' },
  { codigo: '3', descricao: 'Cliente Novo - Qualificar' },
  { codigo: '5', descricao: 'Cliente Novo - Desenvolver' },
  { codigo: '6', descricao: 'Cliente Novo - Propor' },
  { codigo: '7', descricao: 'Cliente Novo - Fechar' },
  { codigo: '8', descricao: 'Cliente Atual - Melhoria' },
  { codigo: '9', descricao: 'Cliente Atual - Novo SKU' }
]

export function descricaoMotivoComercial(codigo: string) {
  return (
    [...MOTIVOS_APROVACAO_COMERCIAL, ...MOTIVOS_REPROVACAO_COMERCIAL].find((m) => m.codigo === codigo)
      ?.descricao ?? ''
  )
}

export interface PreCadastroComercial {
  id: number
  sap: string
  ftOrigem: string
  planta: string
  cliente: string
  vendedor: string
  orcamento: string
  estrutura: string
  precoKg: number
  dataEmissao: string
  /** '' = pendente, 'C' = aprovado, 'R' = reprovado */
  statusComercial: string
  motivo: string
  dataStatus: string
  usuarioStatus: string
}

const preCadastros: PreCadastroComercial[] = Array.from({ length: 22 }, (_, indice) => {
  const cliente = CLIENTES_DES[indice % CLIENTES_DES.length]
  const situacao = ['', '', 'C', 'R'][indice % 4]

  return {
    id: 5200 + indice,
    sap: String(770_000 + indice * 11),
    ftOrigem: `FT${String(30_000 + indice * 3)}`,
    planta: ['00', '04', '05', '06', '07'][indice % 5],
    cliente: cliente.nome,
    vendedor: ['cpereira', 'rnunes'][indice % 2],
    orcamento: String(12_000 + indice * 37),
    estrutura: ['BC', 'B', 'C', 'BCA'][indice % 4],
    precoKg: Number((3.1 + ((indice * 17) % 90) / 100).toFixed(4)),
    dataEmissao: iso(indice + 2),
    statusComercial: situacao,
    motivo: situacao === 'C' ? '8' : situacao === 'R' ? 'P' : '',
    dataStatus: situacao ? iso(indice) : '',
    usuarioStatus: situacao ? 'cpereira' : ''
  }
})

export interface FiltrosComercial {
  sap: string
  ftOrigem: string
  cliente: string
  status: string
}

export const FILTROS_COMERCIAL_VAZIOS: FiltrosComercial = {
  sap: '',
  ftOrigem: '',
  cliente: '',
  status: ''
}

export async function listarPreCadastros(filtros: FiltrosComercial): Promise<PreCadastroComercial[]> {
  const cliente = filtros.cliente.trim().toLowerCase()

  const resultado = preCadastros
    .filter((p) => !filtros.sap.trim() || p.sap.includes(filtros.sap.trim()))
    .filter(
      (p) =>
        !filtros.ftOrigem.trim() ||
        p.ftOrigem.toLowerCase().includes(filtros.ftOrigem.trim().toLowerCase())
    )
    .filter((p) => !cliente || p.cliente.toLowerCase().includes(cliente))
    .filter((p) => (filtros.status === 'T' ? true : p.statusComercial === filtros.status))
    .map((p) => ({ ...p }))
    .sort((a, b) => b.dataEmissao.localeCompare(a.dataEmissao))

  return atraso(resultado, 400)
}

export async function decidirComercial(
  id: number,
  aprovar: boolean,
  motivo: string,
  usuario: string
): Promise<ResultadoDesenvolvimento> {
  if (!motivo) {
    return atraso(
      { sucesso: false, mensagem: aprovar ? 'Selecione o motivo da aprovação.' : 'Selecione o motivo da reprovação.' },
      200
    )
  }

  const registro = preCadastros.find((p) => p.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Pré-cadastro não encontrado.' }, 200)

  registro.statusComercial = aprovar ? 'C' : 'R'
  registro.motivo = motivo
  registro.dataStatus = new Date().toISOString().slice(0, 10)
  registro.usuarioStatus = usuario

  return atraso(
    { sucesso: true, mensagem: aprovar ? 'Pré-cadastro aprovado.' : 'Pré-cadastro reprovado.' },
    550
  )
}

/* ------------------------------------------------------------------ *
 * Lote Piloto V2
 * ------------------------------------------------------------------ */

export const STATUS_LOTE_PILOTO = [
  { codigo: 'AGP', descricao: 'Aguardando Pedido LP' },
  { codigo: 'AGPR', descricao: 'Aguardando Produção' },
  { codigo: 'ETC', descricao: 'Em Teste Cliente' },
  { codigo: 'CV', descricao: 'Convertido em Venda' },
  { codigo: 'ASP', descricao: 'Aprovado sem Pedido de Conversão' },
  { codigo: 'RPC', descricao: 'Reprovado pelo Cliente' },
  { codigo: 'RIR', descricao: 'Reprovado Internamente/Retrabalho' },
  { codigo: 'DESC', descricao: 'Descontinuado' },
  { codigo: 'SRV', descricao: 'Sem Retorno Vendas/Cliente' },
  { codigo: 'VMN', descricao: 'Validação Máquina Nova/Homologação' },
  { codigo: 'EXC', descricao: 'Excluído' }
]

export const MOTIVOS_LOTE_PILOTO = [
  { codigo: 'INOV', descricao: 'Inovação' },
  { codigo: 'MPT', descricao: 'Melhoria - Paletização e Transporte' },
  { codigo: 'NSA', descricao: 'Novo SKU - SW - Acompanhamento detalhado' },
  { codigo: 'NSL', descricao: 'Novo SKU - Layout de Impressão' },
  { codigo: 'NSD', descricao: 'Novo SKU - Dimensional' },
  { codigo: 'NSC', descricao: 'Novo SKU - Padrão de Cor / Cromia / Reticula' },
  { codigo: 'NSP', descricao: 'Novo SKU - Paletização e Transporte' }
]

export const TIPOS_LOTE_PILOTO = [
  { codigo: 'LP', descricao: 'Lote Piloto' },
  { codigo: 'LPI', descricao: 'Lote Piloto Industrial' },
  { codigo: 'AM', descricao: 'Amostra' }
]

export const TIPOS_ITEM = [
  { codigo: 'PA', descricao: 'Produto Acabado' },
  { codigo: 'CH', descricao: 'Chapa' },
  { codigo: 'PP', descricao: 'Produto em Processo' }
]

export function descricaoStatusLp(codigo: string) {
  return STATUS_LOTE_PILOTO.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

export function descricaoMotivoLp(codigo: string) {
  return MOTIVOS_LOTE_PILOTO.find((m) => m.codigo === codigo)?.descricao ?? codigo
}

export interface HistoricoStatusLp {
  statusAnterior: string
  statusPosterior: string
  alteradoPor: string
  dataInclusao: string
}

export interface JustificativaLp {
  tipo: string
  descricao: string
  quemIncluiu: string
  dataInclusao: string
}

export interface LotePiloto {
  id: number
  codigoPlanta: string
  dataInclusao: string
  dataStatus: string
  status: string
  codigoCliente: string
  codigoLoja: string
  nomeCliente: string
  tipoItem: string
  codigoProduto: string
  referencia: string
  quantidadePaletes: number
  quantidadePedidos: number
  motivo: string
  precisaLiberacao: boolean
  tipoLp: string
  pedidoKiwi: string
  lotePilotoAprovado: boolean
  usuario: string
  historicoStatus: HistoricoStatusLp[]
  justificativas: JustificativaLp[]
}

interface LotePilotoInterno extends LotePiloto {
  excluido: boolean
}

const lotesPiloto: LotePilotoInterno[] = Array.from({ length: 18 }, (_, indice) => {
  const cliente = CLIENTES_DES[indice % CLIENTES_DES.length]
  const status = STATUS_LOTE_PILOTO[indice % STATUS_LOTE_PILOTO.length]

  return {
    id: 4400 + indice,
    codigoPlanta: ['00', '04', '05', '06', '07'][indice % 5],
    dataInclusao: iso(indice + 3),
    dataStatus: iso(indice),
    status: status.codigo,
    codigoCliente: cliente.cod,
    codigoLoja: cliente.loja,
    nomeCliente: cliente.nome,
    tipoItem: TIPOS_ITEM[indice % TIPOS_ITEM.length].codigo,
    codigoProduto: String(400_000 + indice * 7),
    referencia: `REF-${String(9000 + indice)}`,
    quantidadePaletes: 1 + (indice % 9),
    quantidadePedidos: 1 + (indice % 4),
    motivo: MOTIVOS_LOTE_PILOTO[indice % MOTIVOS_LOTE_PILOTO.length].codigo,
    precisaLiberacao: indice % 3 === 0,
    tipoLp: TIPOS_LOTE_PILOTO[indice % TIPOS_LOTE_PILOTO.length].codigo,
    pedidoKiwi: String(66_000 + indice * 5),
    lotePilotoAprovado: indice % 4 === 0,
    usuario: ['jsilva', 'malves', 'bcarvalho'][indice % 3],
    historicoStatus:
      indice % 2 === 0
        ? [
            {
              statusAnterior: 'AGP',
              statusPosterior: status.codigo,
              alteradoPor: 'malves',
              dataInclusao: iso(indice)
            }
          ]
        : [],
    justificativas:
      indice % 5 === 0
        ? [
            {
              tipo: 'Reprovação',
              descricao: 'Cliente relatou variação dimensional acima da tolerância.',
              quemIncluiu: 'bcarvalho',
              dataInclusao: iso(indice)
            }
          ]
        : [],
    excluido: false
  }
})

export interface FiltrosLotePiloto {
  planta: string
  status: string
  cliente: string
}

export const FILTROS_LP_VAZIOS: FiltrosLotePiloto = { planta: '', status: '', cliente: '' }

export async function listarLotesPiloto(filtros: FiltrosLotePiloto): Promise<LotePiloto[]> {
  const cliente = filtros.cliente.trim().toLowerCase()

  const resultado = lotesPiloto
    .filter((l) => !l.excluido)
    .filter((l) => !filtros.planta || l.codigoPlanta === filtros.planta)
    .filter((l) => !filtros.status || l.status === filtros.status)
    .filter((l) => !cliente || l.nomeCliente.toLowerCase().includes(cliente))
    .map(({ excluido, ...resto }) => resto)
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 400)
}

export interface DadosLotePiloto {
  id?: number
  codigoPlanta: string
  codigoCliente: string
  codigoLoja: string
  nomeCliente: string
  tipoItem: string
  codigoProduto: string
  referencia: string
  quantidadePaletes: string
  quantidadePedidos: string
  motivo: string
  precisaLiberacao: string
  tipoLp: string
  pedidoKiwi: string
  lotePilotoAprovado: string
  status: string
  justificativa: string
}

export const DADOS_LP_VAZIOS: DadosLotePiloto = {
  codigoPlanta: '',
  codigoCliente: '',
  codigoLoja: '',
  nomeCliente: '',
  tipoItem: '',
  codigoProduto: '',
  referencia: '',
  quantidadePaletes: '',
  quantidadePedidos: '',
  motivo: '',
  precisaLiberacao: '',
  tipoLp: '',
  pedidoKiwi: '',
  lotePilotoAprovado: '',
  status: '',
  justificativa: ''
}

export async function salvarLotePiloto(
  dados: DadosLotePiloto,
  usuario: string
): Promise<ResultadoDesenvolvimento> {
  const obrigatorios: [keyof DadosLotePiloto, string][] = [
    ['codigoPlanta', 'Planta'],
    ['codigoCliente', 'Cód. Cliente'],
    ['nomeCliente', 'Nome do Cliente'],
    ['tipoItem', 'Tipo do Item'],
    ['codigoProduto', 'Produto'],
    ['quantidadePaletes', 'Qtde. de Palete'],
    ['quantidadePedidos', 'Qtde. de Pedidos LP'],
    ['motivo', 'Motivo'],
    ['precisaLiberacao', 'Precisa de Liberação'],
    ['tipoLp', 'Tipo de LP'],
    ['pedidoKiwi', 'Pedido KIWI'],
    ['status', 'Status']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }

  const hoje = new Date().toISOString().slice(0, 10)

  const campos = {
    codigoPlanta: dados.codigoPlanta,
    codigoCliente: dados.codigoCliente.trim(),
    codigoLoja: dados.codigoLoja.trim(),
    nomeCliente: dados.nomeCliente.trim(),
    tipoItem: dados.tipoItem,
    codigoProduto: dados.codigoProduto.trim(),
    referencia: dados.referencia.trim(),
    quantidadePaletes: Number(dados.quantidadePaletes),
    quantidadePedidos: Number(dados.quantidadePedidos),
    motivo: dados.motivo,
    precisaLiberacao: dados.precisaLiberacao === 'S',
    tipoLp: dados.tipoLp,
    pedidoKiwi: dados.pedidoKiwi.trim(),
    lotePilotoAprovado: dados.lotePilotoAprovado === 'S'
  }

  if (dados.id) {
    const lote = lotesPiloto.find((l) => l.id === dados.id)
    if (!lote) return atraso({ sucesso: false, mensagem: 'Lote piloto não encontrado.' }, 200)

    // Trocar o status gera registro no histórico, como no legado
    if (lote.status !== dados.status) {
      lote.historicoStatus.push({
        statusAnterior: lote.status,
        statusPosterior: dados.status,
        alteradoPor: usuario,
        dataInclusao: hoje
      })
      lote.dataStatus = hoje
    }

    if (dados.justificativa.trim()) {
      lote.justificativas.push({
        tipo: 'Alteração',
        descricao: dados.justificativa.trim(),
        quemIncluiu: usuario,
        dataInclusao: hoje
      })
    }

    Object.assign(lote, campos, { status: dados.status })
    return atraso({ sucesso: true, mensagem: 'Lote piloto atualizado com sucesso!' }, 500)
  }

  lotesPiloto.push({
    id: Math.max(0, ...lotesPiloto.map((l) => l.id)) + 1,
    dataInclusao: hoje,
    dataStatus: hoje,
    status: dados.status,
    usuario,
    historicoStatus: [],
    justificativas: dados.justificativa.trim()
      ? [
          {
            tipo: 'Inclusão',
            descricao: dados.justificativa.trim(),
            quemIncluiu: usuario,
            dataInclusao: hoje
          }
        ]
      : [],
    excluido: false,
    ...campos
  })

  return atraso({ sucesso: true, mensagem: 'Lote piloto criado com sucesso!' }, 550)
}

export async function excluirLotePiloto(id: number): Promise<ResultadoDesenvolvimento> {
  const lote = lotesPiloto.find((l) => l.id === id)
  if (!lote) return atraso({ sucesso: false, mensagem: 'Lote piloto não encontrado.' }, 200)

  lote.excluido = true
  lote.status = 'EXC'
  return atraso({ sucesso: true, mensagem: 'Lote piloto excluído com sucesso!' }, 400)
}

/* ------------------------------------------------------------------ *
 * Solicitação de Projeto ao Desenvolvimento (SPD) V2
 * ------------------------------------------------------------------ */

/** Etapas de andamento do projeto */
export const STATUS_ANDAMENTO_SPD = [
  'Análise do Projeto',
  'Desenvolvimento de Layout',
  'Aprovação de Layout',
  'Corte de Amostras',
  'Envio de Especificação/Amostra',
  'Amostra Enviada',
  'Amostra Aprovada',
  'Amostra Reprovada',
  'Falta de Chapas',
  'Orçamento',
  'Liberação Análise de Custo',
  'Aprovação Custos Gerência',
  'Liberação Comercial - Gerência',
  'Liberação PPCP/Produção'
]

/** Tipo de solicitação — classifica o processo de aprovação */
export const PROCESSOS_SPD = [
  'Inovação',
  'Melhoria Continua - SK',
  'Melhoria Continua - Cliente',
  'Alteração Layout',
  'Alteração Dimensional',
  'Alteração Estrutura',
  'Nova Referência conforme especificado'
]

export const MOTIVOS_REJEICAO_SPD = [
  'Emissão em Duplicidade',
  'Preço',
  'Descontinuidade do Projeto',
  'Sem Retorno Comercial',
  'Sem Retorno Cliente',
  'Proposta Descartada'
]

export const ESTRUTURAS_PAPELAO = ['B', 'C', 'E', 'BC', 'BCA', 'EB']

export const ESTILOS_CAIXA = ['FEFCO 0201', 'FEFCO 0203', 'FEFCO 0300', 'FEFCO 0409', 'FEFCO 0713']

/** Etapa de aprovação por área */
export type AreaAprovacao = 'COMERCIAL' | 'DESENVOLVIMENTO' | 'PCP'

export interface AprovacaoSpd {
  area: AreaAprovacao
  aprovado: boolean
  usuario: string
  data: string
  motivo: string
}

export interface AndamentoSpd {
  processo: string
  inicio: string
  fim: string
  quem: string
  complemento: string
}

export interface SolicitacaoSpd {
  id: number
  codPlanta: string
  codCliente: string
  codLoja: string
  nomeCliente: string
  vendedor: string
  tipoSolicitacao: string
  numeroFtCopia: string
  ft: string
  desenvolvedor: string
  referencia: string
  necessitaAcompanhamento: boolean
  nivelLotePiloto: string
  toleranciaMaxima: string
  toleranciaMinima: string
  estilo: string
  observacoes: string
  /* Dimensional */
  comprimento: string
  largura: string
  altura: string
  numeroMatriz: string
  clienteUsaTara: boolean
  toleranciaPeso: string
  /* Estrutura */
  estruturaPapelao: string
  gramatura: string
  larguraChapa: string
  comprimentoChapa: string
  areaBlank: string
  areaLiquida: string
  peso: string
  inkJet: boolean
  /* Testes físicos */
  gramaturaReferencial: string
  esmagamento: string
  ectColuna: string
  mullenReferencial: string
  espessura: string
  bctCompressao: string
  cobbInterno: string
  cobbExterno: string
  /* Fechamento e paletização */
  tipoFechamento: string
  ehAmarrado: boolean
  quantidadeFitilhos: string
  tipoFitilho: string
  pecasPorUnidade: string
  ehPaletizado: boolean
  palete: string
  alturaMaximaPalete: string
  temFacesOpostas: boolean
  temCantinho: boolean
  temGradeMadeira: boolean
  temFilmeStrech: boolean
  clienteUsaRack: boolean
  overhang: boolean
  quantidadeEspelhos: string
  /* Aplicação */
  conteudoEmbalado: string
  consumoMensal: string
  tipoLinhaProducaoCliente: string
  tipoProdutoPrimario: string
  pesoPrimario: string
  quantidadePrimarios: string
  suportacaoPrimario: string
  tempoEstocagem: string
  umidadeRelativa: string
  quantidadeCores: string
  /* Comercial */
  precoKg: string
  orcamento: string
  margemContribuicao: string
  resultado: string
  tipoSavings: string
  valorSavingSk: string
  valorSavingCliente: string
  valorEstimadoCliche: string
  valorEstimadoMatriz: string
  informacoesAdicionais: string
  /* Amostras */
  quantidadeAmostras: string
  amostraCep: string
  amostraResponsavel: string
  /* Fluxo */
  dataEmissao: string
  status: 'ABERTA' | 'EM_ANDAMENTO' | 'APROVADA' | 'REJEITADA' | 'CANCELADA'
  statusAndamento: string
  aprovacoes: AprovacaoSpd[]
  andamentos: AndamentoSpd[]
}

export const SPD_VAZIA: Omit<SolicitacaoSpd, 'id' | 'dataEmissao' | 'status' | 'statusAndamento' | 'aprovacoes' | 'andamentos' | 'ft'> = {
  codPlanta: '',
  codCliente: '',
  codLoja: '',
  nomeCliente: '',
  vendedor: '',
  tipoSolicitacao: '',
  numeroFtCopia: '',
  desenvolvedor: '',
  referencia: '',
  necessitaAcompanhamento: false,
  nivelLotePiloto: '',
  toleranciaMaxima: '',
  toleranciaMinima: '',
  estilo: '',
  observacoes: '',
  comprimento: '',
  largura: '',
  altura: '',
  numeroMatriz: '',
  clienteUsaTara: false,
  toleranciaPeso: '',
  estruturaPapelao: '',
  gramatura: '',
  larguraChapa: '',
  comprimentoChapa: '',
  areaBlank: '',
  areaLiquida: '',
  peso: '',
  inkJet: false,
  gramaturaReferencial: '',
  esmagamento: '',
  ectColuna: '',
  mullenReferencial: '',
  espessura: '',
  bctCompressao: '',
  cobbInterno: '',
  cobbExterno: '',
  tipoFechamento: '',
  ehAmarrado: false,
  quantidadeFitilhos: '',
  tipoFitilho: '',
  pecasPorUnidade: '',
  ehPaletizado: false,
  palete: '',
  alturaMaximaPalete: '',
  temFacesOpostas: false,
  temCantinho: false,
  temGradeMadeira: false,
  temFilmeStrech: false,
  clienteUsaRack: false,
  overhang: false,
  quantidadeEspelhos: '',
  conteudoEmbalado: '',
  consumoMensal: '',
  tipoLinhaProducaoCliente: '',
  tipoProdutoPrimario: '',
  pesoPrimario: '',
  quantidadePrimarios: '',
  suportacaoPrimario: '',
  tempoEstocagem: '',
  umidadeRelativa: '',
  quantidadeCores: '',
  precoKg: '',
  orcamento: '',
  margemContribuicao: '',
  resultado: '',
  tipoSavings: '',
  valorSavingSk: '',
  valorSavingCliente: '',
  valorEstimadoCliche: '',
  valorEstimadoMatriz: '',
  informacoesAdicionais: '',
  quantidadeAmostras: '',
  amostraCep: '',
  amostraResponsavel: ''
}

const spds: SolicitacaoSpd[] = Array.from({ length: 16 }, (_, indice) => {
  const cliente = CLIENTES_DES[indice % CLIENTES_DES.length]
  const status = (['ABERTA', 'EM_ANDAMENTO', 'APROVADA', 'REJEITADA'] as const)[indice % 4]

  return {
    ...SPD_VAZIA,
    id: 3300 + indice,
    codPlanta: ['00', '04', '05', '06', '07'][indice % 5],
    codCliente: cliente.cod,
    codLoja: cliente.loja,
    nomeCliente: cliente.nome,
    vendedor: ['cpereira', 'rnunes'][indice % 2],
    tipoSolicitacao: PROCESSOS_SPD[indice % PROCESSOS_SPD.length],
    desenvolvedor: ['bcarvalho', 'plima'][indice % 2],
    referencia: `REF-SPD-${String(7000 + indice)}`,
    ft: `FT${String(31_000 + indice * 4)}`,
    estilo: ESTILOS_CAIXA[indice % ESTILOS_CAIXA.length],
    comprimento: String(300 + indice * 12),
    largura: String(200 + indice * 8),
    altura: String(150 + indice * 5),
    estruturaPapelao: ESTRUTURAS_PAPELAO[indice % ESTRUTURAS_PAPELAO.length],
    gramatura: String(150 + indice * 7),
    precoKg: (3.2 + (indice % 12) / 10).toFixed(2),
    consumoMensal: String(10_000 + indice * 1_300),
    quantidadeCores: String(1 + (indice % 5)),
    dataEmissao: iso(indice + 5),
    status,
    statusAndamento: STATUS_ANDAMENTO_SPD[indice % STATUS_ANDAMENTO_SPD.length],
    aprovacoes:
      status === 'APROVADA'
        ? [
            { area: 'COMERCIAL', aprovado: true, usuario: 'cpereira', data: iso(indice + 3), motivo: 'Cliente Atual - Novo SKU' },
            { area: 'DESENVOLVIMENTO', aprovado: true, usuario: 'bcarvalho', data: iso(indice + 2), motivo: '' },
            { area: 'PCP', aprovado: true, usuario: 'jsilva', data: iso(indice + 1), motivo: '' }
          ]
        : status === 'REJEITADA'
          ? [{ area: 'COMERCIAL', aprovado: false, usuario: 'cpereira', data: iso(indice + 3), motivo: 'Preço' }]
          : [],
    andamentos:
      status === 'EM_ANDAMENTO'
        ? [
            {
              processo: 'Análise do Projeto',
              inicio: iso(indice + 4),
              fim: iso(indice + 2),
              quem: 'bcarvalho',
              complemento: 'Projeto viável com a estrutura proposta.'
            }
          ]
        : []
  }
})

export interface FiltrosSpd {
  planta: string
  cliente: string
  referencia: string
  status: string
}

export const FILTROS_SPD_VAZIOS: FiltrosSpd = { planta: '', cliente: '', referencia: '', status: '' }

export async function listarSpds(filtros: FiltrosSpd): Promise<SolicitacaoSpd[]> {
  const cliente = filtros.cliente.trim().toLowerCase()
  const referencia = filtros.referencia.trim().toLowerCase()

  const resultado = spds
    .filter((s) => !filtros.planta || s.codPlanta === filtros.planta)
    .filter((s) => !cliente || s.nomeCliente.toLowerCase().includes(cliente))
    .filter((s) => !referencia || s.referencia.toLowerCase().includes(referencia))
    .filter((s) => !filtros.status || s.status === filtros.status)
    .map((s) => ({ ...s }))
    .sort((a, b) => b.dataEmissao.localeCompare(a.dataEmissao))

  return atraso(resultado, 450)
}

export async function salvarSpd(
  dados: Partial<SolicitacaoSpd> & { id?: number },
  usuario: string
): Promise<ResultadoDesenvolvimento> {
  const obrigatorios: [keyof SolicitacaoSpd, string][] = [
    ['codPlanta', 'Planta'],
    ['nomeCliente', 'Nome do Cliente'],
    ['codCliente', 'Cód. Cliente'],
    ['tipoSolicitacao', 'Tipo de Solicitação'],
    ['desenvolvedor', 'Técnico/Desenvolvedor'],
    ['referencia', 'Referência'],
    ['comprimento', 'Comprimento'],
    ['largura', 'Largura'],
    ['altura', 'Altura'],
    ['estruturaPapelao', 'Estrutura de Papelão']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }

  if (dados.id) {
    const spd = spds.find((s) => s.id === dados.id)
    if (!spd) return atraso({ sucesso: false, mensagem: 'SPD não encontrada.' }, 200)
    if (spd.status === 'CANCELADA') {
      return atraso({ sucesso: false, mensagem: 'SPD cancelada não pode ser alterada.' }, 200)
    }

    Object.assign(spd, dados)
    return atraso({ sucesso: true, mensagem: 'SPD atualizada com sucesso!' }, 600)
  }

  const novoId = Math.max(0, ...spds.map((s) => s.id)) + 1
  spds.push({
    ...SPD_VAZIA,
    ...dados,
    id: novoId,
    ft: `FT${String(31_000 + novoId)}`,
    dataEmissao: new Date().toISOString().slice(0, 10),
    status: 'ABERTA',
    statusAndamento: STATUS_ANDAMENTO_SPD[0],
    aprovacoes: [],
    andamentos: [
      {
        processo: STATUS_ANDAMENTO_SPD[0],
        inicio: new Date().toISOString().slice(0, 10),
        fim: '',
        quem: usuario,
        complemento: 'Solicitação registrada.'
      }
    ]
  } as SolicitacaoSpd)

  return atraso({ sucesso: true, mensagem: 'SPD registrada com sucesso!' }, 650)
}

/** alterarStatusAndamento — fecha a etapa atual e abre a próxima */
export async function alterarAndamentoSpd(
  id: number,
  novoAndamento: string,
  complemento: string,
  usuario: string
): Promise<ResultadoDesenvolvimento> {
  const spd = spds.find((s) => s.id === id)
  if (!spd) return atraso({ sucesso: false, mensagem: 'SPD não encontrada.' }, 200)
  if (!novoAndamento) {
    return atraso({ sucesso: false, mensagem: 'Selecione o novo andamento.' }, 200)
  }

  const hoje = new Date().toISOString().slice(0, 10)
  const aberto = spd.andamentos.find((a) => !a.fim)
  if (aberto) aberto.fim = hoje

  spd.andamentos.push({
    processo: novoAndamento,
    inicio: hoje,
    fim: '',
    quem: usuario,
    complemento: complemento.trim()
  })
  spd.statusAndamento = novoAndamento
  if (spd.status === 'ABERTA') spd.status = 'EM_ANDAMENTO'

  return atraso({ sucesso: true, mensagem: `Andamento alterado para "${novoAndamento}".` }, 550)
}

/** Aprovação/rejeição por área — Comercial, Desenvolvimento e PCP */
export async function decidirSpd(
  id: number,
  area: AreaAprovacao,
  aprovar: boolean,
  motivo: string,
  usuario: string
): Promise<ResultadoDesenvolvimento> {
  const spd = spds.find((s) => s.id === id)
  if (!spd) return atraso({ sucesso: false, mensagem: 'SPD não encontrada.' }, 200)
  if (!aprovar && !motivo) {
    return atraso({ sucesso: false, mensagem: 'Selecione o motivo da rejeição.' }, 200)
  }
  if (spd.aprovacoes.some((a) => a.area === area)) {
    return atraso({ sucesso: false, mensagem: `A área ${area} já se manifestou nesta SPD.` }, 200)
  }

  spd.aprovacoes.push({
    area,
    aprovado: aprovar,
    usuario,
    data: new Date().toISOString().slice(0, 10),
    motivo
  })

  if (!aprovar) {
    spd.status = 'REJEITADA'
    return atraso({ sucesso: true, mensagem: `SPD rejeitada pela área ${area}.` }, 600)
  }

  // Aprovada só quando as três áreas se manifestam positivamente
  const areas: AreaAprovacao[] = ['COMERCIAL', 'DESENVOLVIMENTO', 'PCP']
  const todasAprovaram = areas.every((item) =>
    spd.aprovacoes.some((a) => a.area === item && a.aprovado)
  )

  spd.status = todasAprovaram ? 'APROVADA' : 'EM_ANDAMENTO'
  return atraso(
    {
      sucesso: true,
      mensagem: todasAprovaram
        ? 'SPD aprovada por todas as áreas!'
        : `Aprovação da área ${area} registrada.`
    },
    600
  )
}

export async function cancelarSpd(id: number, reverter = false): Promise<ResultadoDesenvolvimento> {
  const spd = spds.find((s) => s.id === id)
  if (!spd) return atraso({ sucesso: false, mensagem: 'SPD não encontrada.' }, 200)

  if (reverter) {
    if (spd.status !== 'CANCELADA') {
      return atraso({ sucesso: false, mensagem: 'Esta SPD não está cancelada.' }, 200)
    }
    spd.status = spd.andamentos.length > 1 ? 'EM_ANDAMENTO' : 'ABERTA'
    return atraso({ sucesso: true, mensagem: 'Cancelamento revertido.' }, 500)
  }

  if (spd.status === 'APROVADA') {
    return atraso({ sucesso: false, mensagem: 'SPD aprovada não pode ser cancelada.' }, 200)
  }

  spd.status = 'CANCELADA'
  return atraso({ sucesso: true, mensagem: 'SPD cancelada com sucesso!' }, 500)
}

/** editarTestesFisicos — tela dedicada do legado, só os 8 campos de ensaio */
export async function salvarTestesFisicos(
  id: number,
  testes: Pick<
    SolicitacaoSpd,
    | 'gramaturaReferencial'
    | 'esmagamento'
    | 'ectColuna'
    | 'mullenReferencial'
    | 'espessura'
    | 'bctCompressao'
    | 'cobbInterno'
    | 'cobbExterno'
  >
): Promise<ResultadoDesenvolvimento> {
  const spd = spds.find((s) => s.id === id)
  if (!spd) return atraso({ sucesso: false, mensagem: 'SPD não encontrada.' }, 200)

  Object.assign(spd, testes)
  return atraso({ sucesso: true, mensagem: 'Testes físicos atualizados com sucesso!' }, 500)
}
