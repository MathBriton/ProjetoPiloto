import { atraso } from '@/services/comuns/baseMock'

/**
 * Validacoes — validação de pedidos (SP_LIBERAPEDIDOSZF / SP_LIBERAPEDIDO),
 * validação de pedidos de papel (SP_LIBERAPEDIDOSZF_PAPEL) e a fila de
 * aprovação das solicitações de crédito.
 */

/* ------------------------------------------------------------------ *
 * Validação de pedidos
 * ------------------------------------------------------------------ */

/**
 * Códigos de bloqueio do GravaStatus.asp. O código e a descrição são os
 * mesmos gravados na SP_LIBERAPEDIDO.
 */
export const MOTIVOS_BLOQUEIO = {
  LI: { codigo: '', descricao: '', rotulo: 'Liberado' },
  PR: { codigo: '17', descricao: 'Preco menor que do Pedido anterior', rotulo: 'Bloqueado por preço' },
  CP: { codigo: '30', descricao: 'Condição de Pagamento', rotulo: 'Bloqueado por cond. pagto.' },
  DE: { codigo: '35', descricao: 'Data de Entrega', rotulo: 'Bloqueado por data de entrega' },
  LC: { codigo: '40', descricao: 'Bloqueado por Limite de Crédito', rotulo: 'Bloqueado por lim. crédito' }
} as const

/** Status que o usuario pode escolher na validacao */
export type StatusSelecionavel = keyof typeof MOTIVOS_BLOQUEIO

/** Inclui o estado inicial, antes de qualquer validacao */
export type StatusValidacao = StatusSelecionavel | 'PENDENTE'

/** Papel só oferece preço, condição de pagamento e liberado */
export const STATUS_PAPEL: StatusSelecionavel[] = ['PR', 'CP', 'LI']

/** Pedido normal oferece as cinco opções do ValidaPedido.asp */
export const STATUS_PEDIDO: StatusSelecionavel[] = ['PR', 'LC', 'DE', 'CP', 'LI']

export function rotuloStatus(status: StatusValidacao) {
  return status === 'PENDENTE' ? 'Pendente' : MOTIVOS_BLOQUEIO[status].rotulo
}

export interface PedidoValidacao {
  id: number
  pedido: string
  cliente: string
  loja: string
  nome: string
  ft: string
  produto: string
  descricaoProduto: string
  tipoProduto: string
  tipoOnda: string
  fabrica: string
  quantidade: number
  preco: number
  precoKg: number
  precoKgNet: number
  m2: number
  emissao: string
  entrega: string
  observacao: string
  status: StatusValidacao
  bloqueio: string
  quemEmitiu: string
  /** Preço do último pedido do cliente — base da regra de preço menor */
  ultimoPreco: number
  pesoTotal: number
  planta: string
  condicaoPagamento: string
  localEntrega: string
  /** TES de pedido livre de débito: só gerente comercial valida */
  tesLivreDebito: boolean
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const CLIENTES_PEDIDO = [
  { cliente: '001234', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', municipio: 'JUIZ DE FORA - MG' },
  { cliente: '002567', loja: '01', nome: 'BEBIDAS DO VALE LTDA', municipio: 'CATAGUASES - MG' },
  { cliente: '003891', loja: '01', nome: 'FRIGORIFICO SERRA AZUL LTDA', municipio: 'BENTO GONCALVES - RS' },
  { cliente: '004125', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.', municipio: 'MARANGUAPE - CE' },
  { cliente: '006432', loja: '01', nome: 'DISTRIBUIDORA LITORAL SUL LTDA', municipio: 'SAQUAREMA - RJ' }
]

/** Tipos de produto usados pelas regras de peso do GravaStatus.asp */
const TIPOS_PRODUTO = [
  { codigo: '901', descricao: 'CAIXA' },
  { codigo: '902', descricao: 'MANTA' },
  { codigo: '903', descricao: 'CHAPA' }
]

function gerarPedidos(qtd: number, papel: boolean, plantas: string[]): PedidoValidacao[] {
  return Array.from({ length: qtd }, (_, indice) => {
    const dono = CLIENTES_PEDIDO[indice % CLIENTES_PEDIDO.length]
    const tipo = TIPOS_PRODUTO[indice % TIPOS_PRODUTO.length]
    const preco = Number((2.4 + ((indice * 37) % 180) / 100).toFixed(4))
    // 1 em cada 3 vem com preço abaixo do anterior, para exercitar a regra
    const ultimoPreco = indice % 3 === 0 ? Number((preco + 0.35).toFixed(4)) : Number((preco - 0.12).toFixed(4))
    const quantidade = 1_000 + ((indice * 313) % 24_000)

    return {
      id: (papel ? 20_000 : 10_000) + indice,
      pedido: String((papel ? 700_000 : 500_000) + indice * 13),
      cliente: dono.cliente,
      loja: dono.loja,
      nome: dono.nome,
      ft: String(40_000 + indice * 7),
      produto: papel ? `PA${String(1000 + indice).slice(-4)}` : `${tipo.codigo}${String(100 + indice).slice(-3)}`,
      descricaoProduto: papel ? 'BOBINA KRAFTLINER 175G' : `${tipo.descricao} ONDULADA 1200X800`,
      tipoProduto: papel ? 'PAPEL' : tipo.codigo,
      tipoOnda: papel ? '' : ['SIMPLES', 'DUPLA'][indice % 2],
      fabrica: ['F1', 'F2'][indice % 2],
      quantidade,
      preco,
      precoKg: Number((preco / 1.8).toFixed(4)),
      precoKgNet: Number((preco / 2.05).toFixed(4)),
      m2: Number((quantidade * 0.96).toFixed(2)),
      emissao: iso(indice + 3),
      entrega: iso(-(5 + (indice % 20))),
      observacao: indice % 5 === 0 ? 'Cliente solicitou entrega antecipada.' : '',
      status: 'PENDENTE' as StatusValidacao,
      bloqueio: '',
      quemEmitiu: ['jsilva', 'malves', 'cpereira'][indice % 3],
      ultimoPreco,
      // Peso baixo em alguns registros aciona a regra de peso mínimo
      pesoTotal: indice % 4 === 0 ? 180 : 400 + ((indice * 53) % 900),
      planta: plantas[indice % plantas.length],
      condicaoPagamento: ['28 DDL', '21 DDL', '30 DDL'][indice % 3],
      localEntrega: dono.municipio,
      tesLivreDebito: indice % 9 === 0
    }
  })
}

const pedidos = gerarPedidos(28, false, ['00', '04'])
const pedidosPapel = gerarPedidos(14, true, ['00'])
const pedidosBrasil = gerarPedidos(22, false, ['05', '06', '07'])

/** Parâmetros de peso mínimo por tipo de produto (regras do GravaStatus.asp) */
export const PESO_MINIMO = { chapaSimples: 250, chapaDupla: 300, caixa: 280, manta: 200 }

export interface AlertaValidacao {
  tipo: 'preco' | 'peso' | 'tes'
  mensagem: string
}

/**
 * Regras que o legado dispara ao tentar liberar um pedido. Retorna os avisos
 * que exigem confirmação explícita antes de gravar.
 */
export function alertasDaLiberacao(pedido: PedidoValidacao): AlertaValidacao[] {
  const alertas: AlertaValidacao[] = []

  if (pedido.ultimoPreco > 0 && pedido.preco < pedido.ultimoPreco) {
    alertas.push({
      tipo: 'preco',
      mensagem: `Preço (${pedido.preco}) menor que o do pedido anterior (${pedido.ultimoPreco}).`
    })
  }

  const limitePeso =
    pedido.tipoProduto === '903'
      ? pedido.tipoOnda === 'DUPLA'
        ? PESO_MINIMO.chapaDupla
        : PESO_MINIMO.chapaSimples
      : pedido.tipoProduto === '901'
        ? PESO_MINIMO.caixa
        : pedido.tipoProduto === '902'
          ? PESO_MINIMO.manta
          : 0

  if (limitePeso > 0 && pedido.pesoTotal < limitePeso) {
    alertas.push({
      tipo: 'peso',
      mensagem: `Peso total (${pedido.pesoTotal} kg) abaixo do mínimo de ${limitePeso} kg para este produto.`
    })
  }

  if (pedido.tesLivreDebito) {
    alertas.push({
      tipo: 'tes',
      mensagem:
        'Pedido com TES destinada a PEDIDOS LIVRES DE DÉBITO. Somente o Gerente Comercial pode validar.'
    })
  }

  return alertas
}

export type OrigemPedidos = 'pedido' | 'papel' | 'brasil'

function base(origem: OrigemPedidos) {
  if (origem === 'papel') return pedidosPapel
  if (origem === 'brasil') return pedidosBrasil
  return pedidos
}

export async function listarPedidosValidacao(
  origem: OrigemPedidos,
  planta = ''
): Promise<PedidoValidacao[]> {
  const resultado = base(origem)
    .filter((p) => !planta || p.planta === planta)
    .map((p) => ({ ...p }))
    .sort((a, b) => a.entrega.localeCompare(b.entrega))

  return atraso(resultado, 450)
}

export async function obterPedido(origem: OrigemPedidos, id: number): Promise<PedidoValidacao | null> {
  const encontrado = base(origem).find((p) => p.id === id)
  return atraso(encontrado ? { ...encontrado } : null, 300)
}

/** Últimos pedidos do cliente — o quadro comparativo do ValidaPedido.asp */
export async function ultimosPedidosDoCliente(
  origem: OrigemPedidos,
  cliente: string,
  idAtual: number
): Promise<PedidoValidacao[]> {
  const resultado = base(origem)
    .filter((p) => p.cliente === cliente && p.id !== idAtual)
    .slice(0, 5)
    .map((p) => ({ ...p }))

  return atraso(resultado, 300)
}

export interface ResultadoValidacao {
  sucesso: boolean
  mensagem: string
}

export async function gravarStatusPedido(
  origem: OrigemPedidos,
  id: number,
  status: StatusValidacao,
  usuario: string
): Promise<ResultadoValidacao> {
  if (status === 'PENDENTE') {
    return atraso({ sucesso: false, mensagem: 'Selecione um status para o pedido.' }, 200)
  }

  const pedido = base(origem).find((p) => p.id === id)
  if (!pedido) return atraso({ sucesso: false, mensagem: 'Pedido não encontrado.' }, 200)

  const motivo = MOTIVOS_BLOQUEIO[status]
  pedido.status = status
  pedido.bloqueio = motivo.descricao
  pedido.quemEmitiu = pedido.quemEmitiu || usuario

  const mensagem =
    status === 'LI'
      ? `PEDIDO ${pedido.pedido} LIBERADO !!`
      : status === 'LC'
        ? 'CLIENTE BLOQUEADO POR LIMITE DE CRÉDITO !!'
        : `PEDIDO ${pedido.pedido} BLOQUEADO POR ${motivo.rotulo.replace('Bloqueado por ', '').toUpperCase()} !!`

  return atraso({ sucesso: true, mensagem }, 450)
}

/** "Confirma o envio de todos os pedidos liberados?" do pa.asp */
export async function enviarLiberados(origem: OrigemPedidos): Promise<ResultadoValidacao> {
  const liberados = base(origem).filter((p) => p.status === 'LI')
  if (liberados.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Nenhum pedido liberado para envio.' }, 200)
  }

  // O envio tira os pedidos da fila de validação
  for (const pedido of liberados) {
    const indice = base(origem).indexOf(pedido)
    if (indice >= 0) base(origem).splice(indice, 1)
  }

  return atraso(
    { sucesso: true, mensagem: `${liberados.length} pedido(s) enviado(s) com sucesso.` },
    700
  )
}

/* ------------------------------------------------------------------ *
 * Solicitações de Crédito — fila de aprovação
 * ------------------------------------------------------------------ */

/**
 * Cadeia de aprovação. Cada nível de validador enxerga um status:
 * 10 Analista (A/D), 20 Coordenador (C), 30 Gerente Tesouraria (E),
 * 40 Controller (F), 50 Gerente de Planta (G), 60 CFO (H).
 */
export const ETAPAS_CREDITO = [
  { status: 'A', nivel: '10', rotulo: 'Análise do Analista' },
  { status: 'C', nivel: '20', rotulo: 'Coordenador de Crédito' },
  { status: 'E', nivel: '30', rotulo: 'Gerente de Tesouraria' },
  { status: 'F', nivel: '40', rotulo: 'Divisional Controller' },
  { status: 'G', nivel: '50', rotulo: 'Gerente de Planta' },
  { status: 'H', nivel: '60', rotulo: 'CFO' }
] as const

export type StatusCredito = 'A' | 'C' | 'D' | 'E' | 'F' | 'G' | 'H' | 'APROVADO' | 'REJEITADO'

export function rotuloStatusCredito(status: StatusCredito) {
  if (status === 'APROVADO') return 'Aprovado'
  if (status === 'REJEITADO') return 'Rejeitado'
  if (status === 'D') return 'Aguardando aprovação do Analista'
  return ETAPAS_CREDITO.find((e) => e.status === status)?.rotulo ?? status
}

/** Status que cada nível de validador tem permissão de tratar */
export function statusDoNivel(nivel: string, aba: string) {
  if (nivel === '10') {
    if (aba === 'aprovacoes') return 'D'
    if (aba === 'transferencias') return 'TRANSFERENCIA'
    return 'A'
  }
  return ETAPAS_CREDITO.find((e) => e.nivel === nivel)?.status ?? ''
}

export interface AprovacaoEtapa {
  etapa: string
  validador: string
  data: string
  observacao: string
}

export interface SolicitacaoCredito {
  id: number
  tipo: 'AUMENTO' | 'TRANSFERENCIA'
  solicitante: string
  cliente: string
  cnpj: string
  plantaOrigem: string
  plantaDestino: string
  limiteAnteriorOrigem: number
  limiteAnteriorDestino: number
  valorSolicitado: number
  status: StatusCredito
  dataSolicitacao: string
  observacaoSolicitante: string
  /** Campos preenchidos pelo analista durante a análise */
  scoreSerasa: string
  classeRiscoSerasa: string
  limiteSugeridoSerasa: string
  scoreSw: string
  classeRiscoSw: string
  limiteSugeridoSw: string
  valorSeguradora: string
  faturamentoEstimado: string
  balancoAno: string
  aprovacoes: AprovacaoEtapa[]
}

const solicitacoes: SolicitacaoCredito[] = [
  {
    id: 8001,
    tipo: 'AUMENTO',
    solicitante: 'jsilva',
    cliente: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.',
    cnpj: '11222333000144',
    plantaOrigem: '00',
    plantaDestino: '',
    limiteAnteriorOrigem: 250_000,
    limiteAnteriorDestino: 0,
    valorSolicitado: 400_000,
    status: 'A',
    dataSolicitacao: iso(3),
    observacaoSolicitante: 'Cliente ampliou linha de produção e dobrou o volume mensal.',
    scoreSerasa: '',
    classeRiscoSerasa: '',
    limiteSugeridoSerasa: '',
    scoreSw: '',
    classeRiscoSw: '',
    limiteSugeridoSw: '',
    valorSeguradora: '',
    faturamentoEstimado: '',
    balancoAno: '',
    aprovacoes: []
  },
  {
    id: 8002,
    tipo: 'AUMENTO',
    solicitante: 'malves',
    cliente: 'FRIGORIFICO SERRA AZUL LTDA',
    cnpj: '33444555000166',
    plantaOrigem: '05',
    plantaDestino: '',
    limiteAnteriorOrigem: 400_000,
    limiteAnteriorDestino: 0,
    valorSolicitado: 620_000,
    status: 'C',
    dataSolicitacao: iso(9),
    observacaoSolicitante: 'Contrato anual firmado com rede de supermercados.',
    scoreSerasa: '742',
    classeRiscoSerasa: 'B',
    limiteSugeridoSerasa: '580000',
    scoreSw: '710',
    classeRiscoSw: 'B',
    limiteSugeridoSw: '600000',
    valorSeguradora: '450000',
    faturamentoEstimado: '1200000',
    balancoAno: '2025',
    aprovacoes: [
      { etapa: 'Análise do Analista', validador: 'plima', data: iso(7), observacao: 'Documentação completa. Score adequado.' }
    ]
  },
  {
    id: 8003,
    tipo: 'TRANSFERENCIA',
    solicitante: 'cpereira',
    cliente: 'DISTRIBUIDORA LITORAL SUL LTDA',
    cnpj: '66777888000199',
    plantaOrigem: '07',
    plantaDestino: '00',
    limiteAnteriorOrigem: 60_000,
    limiteAnteriorDestino: 250_000,
    valorSolicitado: 25_000,
    status: 'A',
    dataSolicitacao: iso(2),
    observacaoSolicitante: 'Cliente concentrou compras na planta de Pirapetinga.',
    scoreSerasa: '',
    classeRiscoSerasa: '',
    limiteSugeridoSerasa: '',
    scoreSw: '',
    classeRiscoSw: '',
    limiteSugeridoSw: '',
    valorSeguradora: '',
    faturamentoEstimado: '',
    balancoAno: '',
    aprovacoes: []
  },
  {
    id: 8004,
    tipo: 'AUMENTO',
    solicitante: 'bcarvalho',
    cliente: 'HORTIFRUTI NORDESTE COMERCIO S.A.',
    cnpj: '44555666000177',
    plantaOrigem: '06',
    plantaDestino: '',
    limiteAnteriorOrigem: 95_000,
    limiteAnteriorDestino: 0,
    valorSolicitado: 180_000,
    status: 'E',
    dataSolicitacao: iso(18),
    observacaoSolicitante: 'Expansão para novos estados do Nordeste.',
    scoreSerasa: '588',
    classeRiscoSerasa: 'C',
    limiteSugeridoSerasa: '120000',
    scoreSw: '560',
    classeRiscoSw: 'C',
    limiteSugeridoSw: '140000',
    valorSeguradora: '100000',
    faturamentoEstimado: '480000',
    balancoAno: '2025',
    aprovacoes: [
      { etapa: 'Análise do Analista', validador: 'plima', data: iso(16), observacao: 'Score limítrofe, recomendo limite reduzido.' },
      { etapa: 'Coordenador de Crédito', validador: 'malves', data: iso(14), observacao: 'De acordo com a análise.' }
    ]
  },
  {
    id: 8005,
    tipo: 'AUMENTO',
    solicitante: 'jsilva',
    cliente: 'BEBIDAS DO VALE LTDA',
    cnpj: '22333444000155',
    plantaOrigem: '00',
    plantaDestino: '',
    limiteAnteriorOrigem: 120_000,
    limiteAnteriorDestino: 0,
    valorSolicitado: 200_000,
    status: 'APROVADO',
    dataSolicitacao: iso(40),
    observacaoSolicitante: 'Sazonalidade de verão.',
    scoreSerasa: '820',
    classeRiscoSerasa: 'A',
    limiteSugeridoSerasa: '250000',
    scoreSw: '805',
    classeRiscoSw: 'A',
    limiteSugeridoSw: '220000',
    valorSeguradora: '200000',
    faturamentoEstimado: '900000',
    balancoAno: '2024',
    aprovacoes: [
      { etapa: 'Análise do Analista', validador: 'plima', data: iso(38), observacao: 'Cliente com histórico impecável.' },
      { etapa: 'Coordenador de Crédito', validador: 'malves', data: iso(36), observacao: 'Aprovado.' },
      { etapa: 'Gerente de Tesouraria', validador: 'bcarvalho', data: iso(34), observacao: 'Aprovado.' }
    ]
  }
]

export interface FiltrosCredito {
  planta: string
  aba: string
  termo: string
}

export async function listarSolicitacoes(
  filtros: FiltrosCredito,
  nivel: string
): Promise<SolicitacaoCredito[]> {
  const alvo = statusDoNivel(nivel, filtros.aba)
  const termo = filtros.termo.trim().toLowerCase()

  const resultado = solicitacoes
    .filter((s) => {
      if (alvo === 'TRANSFERENCIA') return s.tipo === 'TRANSFERENCIA' && s.status !== 'APROVADO'
      return s.status === alvo
    })
    .filter((s) => filtros.planta === 'TODOS' || s.plantaOrigem === filtros.planta)
    .filter((s) => !termo || s.cliente.toLowerCase().includes(termo) || s.cnpj.includes(termo))
    .map((s) => ({ ...s }))
    .sort((a, b) => a.dataSolicitacao.localeCompare(b.dataSolicitacao))

  return atraso(resultado, 450)
}

/** Todas as solicitações, para o histórico e o relatório */
export async function listarTodasSolicitacoes(filtros: FiltrosCredito): Promise<SolicitacaoCredito[]> {
  const termo = filtros.termo.trim().toLowerCase()

  const resultado = solicitacoes
    .filter((s) => filtros.planta === 'TODOS' || s.plantaOrigem === filtros.planta)
    .filter((s) => !termo || s.cliente.toLowerCase().includes(termo) || s.cnpj.includes(termo))
    .map((s) => ({ ...s }))
    .sort((a, b) => b.dataSolicitacao.localeCompare(a.dataSolicitacao))

  return atraso(resultado, 400)
}

export interface DadosValidacaoCredito {
  idSolicitacao: number
  rejeitar: boolean
  limiteValidado: string
  observacao: string
  scoreSerasa: string
  classeRiscoSerasa: string
  limiteSugeridoSerasa: string
  scoreSw: string
  classeRiscoSw: string
  limiteSugeridoSw: string
  valorSeguradora: string
  faturamentoEstimado: string
  balancoAno: string
}

/**
 * Avança (ou encerra) o fluxo de aprovação. Aprovar move para a próxima
 * etapa da cadeia; a última etapa conclui como APROVADO.
 */
export async function validarSolicitacao(
  dados: DadosValidacaoCredito,
  nivel: string,
  usuario: string
): Promise<ResultadoValidacao> {
  const solicitacao = solicitacoes.find((s) => s.id === dados.idSolicitacao)
  if (!solicitacao) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)

  const etapaAtual = ETAPAS_CREDITO.find((e) => e.status === solicitacao.status)
  if (!etapaAtual || etapaAtual.nivel !== nivel) {
    return atraso(
      { sucesso: false, mensagem: 'Seu nível não pode validar esta etapa da solicitação.' },
      200
    )
  }

  if (!dados.observacao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a observação da validação.' }, 200)
  }

  solicitacao.aprovacoes.push({
    etapa: etapaAtual.rotulo,
    validador: usuario,
    data: new Date().toISOString().slice(0, 10),
    observacao: dados.observacao.trim()
  })

  if (dados.rejeitar) {
    solicitacao.status = 'REJEITADO'
    return atraso({ sucesso: true, mensagem: 'Solicitação rejeitada. E-mail enviado ao solicitante.' }, 600)
  }

  // O analista grava os dados de score/limite ao concluir a análise
  if (nivel === '10') {
    Object.assign(solicitacao, {
      scoreSerasa: dados.scoreSerasa,
      classeRiscoSerasa: dados.classeRiscoSerasa,
      limiteSugeridoSerasa: dados.limiteSugeridoSerasa,
      scoreSw: dados.scoreSw,
      classeRiscoSw: dados.classeRiscoSw,
      limiteSugeridoSw: dados.limiteSugeridoSw,
      valorSeguradora: dados.valorSeguradora,
      faturamentoEstimado: dados.faturamentoEstimado,
      balancoAno: dados.balancoAno
    })
  }

  if (dados.limiteValidado.trim()) {
    solicitacao.valorSolicitado = Number(dados.limiteValidado.replace(/\./g, '').replace(',', '.'))
  }

  const indiceAtual = ETAPAS_CREDITO.findIndex((e) => e.status === solicitacao.status)
  const proxima = ETAPAS_CREDITO[indiceAtual + 1]

  if (proxima) {
    solicitacao.status = proxima.status as StatusCredito
    return atraso(
      { sucesso: true, mensagem: `Validado. Encaminhado para ${proxima.rotulo}.` },
      600
    )
  }

  solicitacao.status = 'APROVADO'
  return atraso(
    { sucesso: true, mensagem: 'Solicitação aprovada. Limite do cliente atualizado.' },
    700
  )
}
