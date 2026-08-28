import { atraso } from '@/services/comuns/baseMock'

/**
 * Cliente — Cadastro Unificado V2 (UnificadoCliente), Alterar Status do
 * Cliente, Alteração de Dados e Transferência/Extensão.
 *
 * As três telas de solicitação são fluxos de aprovação por departamento. O
 * departamento do usuário sai das mesmas colunas da CRM_DTI usadas em
 * Permissões Especiais: NivelFiscal, NivelControladoria, NivelE2OPEN (CDAG)
 * e NivelControlesInternos.
 */

export interface ResultadoCliente {
  sucesso: boolean
  mensagem: string
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

/** Departamentos que participam dos fluxos */
export const DEPARTAMENTOS = [
  { codigo: 'SOLICITANTE', descricao: 'Solicitante' },
  { codigo: 'FISCAL', descricao: 'Fiscal' },
  { codigo: 'CONTROLADORIA', descricao: 'Controladoria' },
  { codigo: 'CDAG', descricao: 'E2OPEN / CDAG' },
  { codigo: 'CONTROLES_INTERNOS', descricao: 'Controles Internos' },
  { codigo: 'DIRETOR_COMERCIAL', descricao: 'Diretor Comercial' },
  { codigo: 'FINANCEIRO', descricao: 'Financeiro' }
] as const

export type Departamento = (typeof DEPARTAMENTOS)[number]['codigo']

/* ------------------------------------------------------------------ *
 * Cadastro Unificado V2
 * ------------------------------------------------------------------ */

/** StatusCli do UnificadoCliente */
export const STATUS_UNIFICADO = [
  { codigo: 'P', descricao: 'Pendente', departamento: 'SOLICITANTE' },
  { codigo: 'S', descricao: 'Fiscal', departamento: 'FISCAL' },
  { codigo: 'C', descricao: 'Controladoria', departamento: 'CONTROLADORIA' },
  { codigo: 'E', descricao: 'E2OPEN / CDAG', departamento: 'CDAG' },
  { codigo: 'I', descricao: 'Controles Internos', departamento: 'CONTROLES_INTERNOS' },
  { codigo: 'F', descricao: 'Financeiro / Finalizado', departamento: 'FINANCEIRO' },
  { codigo: 'R', descricao: 'Devolvido', departamento: 'SOLICITANTE' },
  { codigo: 'X', descricao: 'Cancelado', departamento: '' },
  { codigo: 'Z', descricao: 'Cliente Prospect', departamento: 'SOLICITANTE' }
] as const

/** Ordem do fluxo — cada etapa encaminha para a seguinte */
export const FLUXO_UNIFICADO = ['P', 'S', 'C', 'E', 'I', 'F']

export function descricaoStatusUnificado(codigo: string) {
  return STATUS_UNIFICADO.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

/** Tipo de solicitação — 5 = Prospect, que dispensa vários campos */
export const TIPOS_SOLICITACAO_CLIENTE = [
  { codigo: '1', descricao: 'Cadastro de Clientes' },
  { codigo: '5', descricao: 'Cadastro de Cliente Prospect' }
]

export const SEGMENTOS = ['Alimentos', 'Bebidas', 'Hortifruti', 'Frigorífico', 'Químico', 'Outros']
export const CONTAS_CONCILIACAO = ['12100', '12110', '12120']
export const TIPOS_INDUSTRIA = ['Alimentícia', 'Bebidas', 'Farmacêutica', 'Química', 'Outros']
export const CONDICOES_PAGAMENTO = ['14 DDL', '21 DDL', '28 DDL', '30 DDL', '45 DDL']
export const NATUREZAS = ['Revenda', 'Consumo', 'Industrialização']
export const TIPOS_FRETE = ['CIF', 'FOB']
export const IDIOMAS = ['Português', 'Inglês', 'Espanhol']

export interface VistoUnificado {
  data: string
  usuario: string
  acao: string
  observacao: string
  leadtimeDias: number
}

export interface CadastroUnificado {
  codigo: number
  planta: string
  tipoSolicitacao: string
  dataInclusao: string
  status: string
  emitente: string
  /* Identificação */
  eshop: string
  grupoContas: string
  empresa: string
  razaoSocial: string
  nomeFantasia: string
  cnpj: string
  cpf: string
  inscricaoEstadual: string
  inscricaoMunicipal: string
  /* Endereço */
  rua: string
  numero: string
  bairro: string
  cep: string
  estado: string
  cidade: string
  pais: string
  kmDestino: string
  idioma: string
  telefone: string
  emailOrdemPagamento: string
  codigoHfm: string
  /* Fiscal */
  declarantePisCofins: boolean
  contribuinteIcms: boolean
  icmsDiferido: boolean
  suspensaoIpi: boolean
  tipoIndustria: string
  codigoSuframa: string
  /* Financeiro */
  contaConciliacao: string
  chaveOrdenacao: string
  condicaoPagamento: string
  metodoPagamento: string
  codigoJuros: string
  correspondenciaEmpresa: string
  limiteCredito: string
  volumeProjetadoM2: string
  /* Comercial */
  segmento: string
  codigoSap: string
  natureza: string
  tipo: string
  tipoFrete: string
  tes: string
  codigoVendedor: string
  transitTime: string
  toleranciaEntregaMaior: string
  toleranciaEntregaMenor: string
  /* Endereços auxiliares */
  entregaRua: string
  entregaNumero: string
  entregaBairro: string
  entregaCep: string
  entregaEstado: string
  entregaCidade: string
  entregaKm: string
  /* Anexos e trilha */
  anexos: string[]
  vistos: VistoUnificado[]
}

export const UNIFICADO_VAZIO: Omit<
  CadastroUnificado,
  'codigo' | 'dataInclusao' | 'status' | 'emitente' | 'vistos'
> = {
  planta: '',
  tipoSolicitacao: '',
  eshop: '',
  grupoContas: '',
  empresa: '',
  razaoSocial: '',
  nomeFantasia: '',
  cnpj: '',
  cpf: '',
  inscricaoEstadual: '',
  inscricaoMunicipal: '',
  rua: '',
  numero: '',
  bairro: '',
  cep: '',
  estado: '',
  cidade: '',
  pais: 'Brasil',
  kmDestino: '',
  idioma: 'Português',
  telefone: '',
  emailOrdemPagamento: '',
  codigoHfm: '',
  declarantePisCofins: false,
  contribuinteIcms: false,
  icmsDiferido: false,
  suspensaoIpi: false,
  tipoIndustria: '',
  codigoSuframa: '',
  contaConciliacao: '',
  chaveOrdenacao: '',
  condicaoPagamento: '',
  metodoPagamento: '',
  codigoJuros: '',
  correspondenciaEmpresa: '',
  limiteCredito: '',
  volumeProjetadoM2: '',
  segmento: '',
  codigoSap: '',
  natureza: '',
  tipo: '',
  tipoFrete: '',
  tes: '',
  codigoVendedor: '',
  transitTime: '',
  toleranciaEntregaMaior: '',
  toleranciaEntregaMenor: '',
  entregaRua: '',
  entregaNumero: '',
  entregaBairro: '',
  entregaCep: '',
  entregaEstado: '',
  entregaCidade: '',
  entregaKm: '',
  anexos: []
}

const RAZOES = [
  'INDUSTRIA ALIMENTICIA BOA VISTA S.A.',
  'BEBIDAS DO VALE LTDA',
  'FRIGORIFICO SERRA AZUL LTDA',
  'HORTIFRUTI NORDESTE COMERCIO S.A.',
  'DISTRIBUIDORA LITORAL SUL LTDA'
]

const unificados: CadastroUnificado[] = Array.from({ length: 20 }, (_, indice) => {
  const status = STATUS_UNIFICADO[indice % 6].codigo

  return {
    ...UNIFICADO_VAZIO,
    codigo: 9100 + indice,
    planta: ['00', '04', '05', '06', '07'][indice % 5],
    tipoSolicitacao: indice % 5 === 0 ? '5' : '1',
    dataInclusao: iso(indice + 2),
    status,
    emitente: ['jsilva', 'malves', 'cpereira'][indice % 3],
    razaoSocial: `${RAZOES[indice % RAZOES.length]} ${indice}`,
    nomeFantasia: RAZOES[indice % RAZOES.length].split(' ')[0],
    cnpj: String(11_222_333_000_100 + indice * 17),
    empresa: ['00', '04', '05'][indice % 3],
    cidade: ['JUIZ DE FORA', 'CATAGUASES', 'BENTO GONCALVES', 'MARANGUAPE'][indice % 4],
    estado: ['MG', 'MG', 'RS', 'CE'][indice % 4],
    segmento: SEGMENTOS[indice % SEGMENTOS.length],
    condicaoPagamento: CONDICOES_PAGAMENTO[indice % CONDICOES_PAGAMENTO.length],
    limiteCredito: String(50_000 + indice * 7_500),
    anexos: indice % 3 === 0 ? ['contrato-social.pdf', 'cartao-cnpj.pdf'] : [],
    vistos:
      indice % 2 === 0
        ? [
            {
              data: iso(indice + 1),
              usuario: 'plima',
              acao: 'Enviado para Fiscal',
              observacao: 'Documentação conferida.',
              leadtimeDias: 1
            }
          ]
        : []
  }
})

export interface FiltrosUnificado {
  cliente: string
  codigo: string
  status: string
  dataInicio: string
  dataFim: string
  planta: string
}

export const FILTROS_UNIFICADO_VAZIOS: FiltrosUnificado = {
  cliente: '',
  codigo: '',
  status: '',
  dataInicio: '',
  dataFim: '',
  planta: ''
}

export async function listarUnificados(filtros: FiltrosUnificado): Promise<CadastroUnificado[]> {
  const cliente = filtros.cliente.trim().toLowerCase()

  const resultado = unificados
    .filter((u) => !cliente || u.razaoSocial.toLowerCase().includes(cliente))
    .filter((u) => !filtros.codigo.trim() || String(u.codigo).includes(filtros.codigo.trim()))
    .filter((u) => !filtros.status || u.status === filtros.status)
    .filter((u) => !filtros.planta || u.planta === filtros.planta)
    .filter((u) => !filtros.dataInicio || u.dataInclusao >= filtros.dataInicio)
    .filter((u) => !filtros.dataFim || u.dataInclusao <= filtros.dataFim)
    .map((u) => ({ ...u }))
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 450)
}

export async function salvarUnificado(
  dados: Partial<CadastroUnificado> & { codigo?: number },
  usuario: string
): Promise<ResultadoCliente> {
  const prospect = dados.tipoSolicitacao === '5'

  // Prospect dispensa vários campos; o cadastro completo exige todos
  const obrigatorios: [keyof CadastroUnificado, string][] = prospect
    ? [
        ['planta', 'Planta'],
        ['tipoSolicitacao', 'Tipo de Solicitação'],
        ['empresa', 'Empresa'],
        ['razaoSocial', 'Razão Social'],
        ['rua', 'Rua'],
        ['numero', 'Número'],
        ['bairro', 'Bairro'],
        ['cep', 'CEP'],
        ['estado', 'Estado']
      ]
    : [
        ['planta', 'Planta'],
        ['tipoSolicitacao', 'Tipo de Solicitação'],
        ['eshop', 'Eshop'],
        ['grupoContas', 'Grupo de Contas'],
        ['empresa', 'Empresa'],
        ['razaoSocial', 'Razão Social'],
        ['nomeFantasia', 'Nome Fantasia'],
        ['rua', 'Rua'],
        ['numero', 'Número'],
        ['bairro', 'Bairro'],
        ['cep', 'CEP'],
        ['estado', 'Estado'],
        ['cidade', 'Cidade'],
        ['pais', 'País'],
        ['kmDestino', 'KM até o destino'],
        ['idioma', 'Idioma']
      ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }

  if (!String(dados.cnpj ?? '').trim() && !String(dados.cpf ?? '').trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o CNPJ ou o CPF.' }, 200)
  }

  if (dados.codigo) {
    const registro = unificados.find((u) => u.codigo === dados.codigo)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)
    if (registro.status === 'X') {
      return atraso({ sucesso: false, mensagem: 'Cadastro cancelado não pode ser alterado.' }, 200)
    }

    Object.assign(registro, dados)
    return atraso({ sucesso: true, mensagem: 'Cadastro atualizado com sucesso!' }, 600)
  }

  unificados.push({
    ...UNIFICADO_VAZIO,
    ...dados,
    codigo: Math.max(0, ...unificados.map((u) => u.codigo)) + 1,
    dataInclusao: new Date().toISOString().slice(0, 10),
    status: prospect ? 'Z' : 'P',
    emitente: usuario,
    vistos: []
  } as CadastroUnificado)

  return atraso({ sucesso: true, mensagem: 'Cadastro registrado com sucesso!' }, 650)
}

/** Registra o visto e move o cadastro para a próxima etapa do fluxo */
export async function avancarUnificado(
  codigo: number,
  observacao: string,
  usuario: string
): Promise<ResultadoCliente> {
  const registro = unificados.find((u) => u.codigo === codigo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)

  const indiceAtual = FLUXO_UNIFICADO.indexOf(registro.status)
  if (indiceAtual < 0) {
    return atraso({ sucesso: false, mensagem: 'Cadastro não está em um status do fluxo.' }, 200)
  }

  const proximo = FLUXO_UNIFICADO[indiceAtual + 1]
  if (!proximo) {
    return atraso({ sucesso: false, mensagem: 'Cadastro já está na etapa final.' }, 200)
  }

  registro.vistos.push({
    data: new Date().toISOString().slice(0, 10),
    usuario,
    acao: `Enviado para ${descricaoStatusUnificado(proximo)}`,
    observacao: observacao.trim(),
    leadtimeDias: 0
  })
  registro.status = proximo

  return atraso(
    { sucesso: true, mensagem: `Cadastro encaminhado para ${descricaoStatusUnificado(proximo)}.` },
    600
  )
}

/** Devolve ao solicitante (status R), exigindo justificativa */
export async function devolverUnificado(
  codigo: number,
  observacao: string,
  usuario: string
): Promise<ResultadoCliente> {
  if (!observacao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o motivo da devolução.' }, 200)
  }

  const registro = unificados.find((u) => u.codigo === codigo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)

  registro.vistos.push({
    data: new Date().toISOString().slice(0, 10),
    usuario,
    acao: 'Devolvido ao solicitante',
    observacao: observacao.trim(),
    leadtimeDias: 0
  })
  registro.status = 'R'

  return atraso({ sucesso: true, mensagem: 'Cadastro devolvido ao solicitante.' }, 550)
}

export async function cancelarUnificado(codigo: number, usuario: string): Promise<ResultadoCliente> {
  const registro = unificados.find((u) => u.codigo === codigo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)
  if (registro.status === 'F') {
    return atraso({ sucesso: false, mensagem: 'Cadastro finalizado não pode ser cancelado.' }, 200)
  }

  registro.vistos.push({
    data: new Date().toISOString().slice(0, 10),
    usuario,
    acao: 'Cancelado',
    observacao: '',
    leadtimeDias: 0
  })
  registro.status = 'X'

  return atraso({ sucesso: true, mensagem: 'Cadastro cancelado.' }, 500)
}

/** Troca direta de status — o UPDATE StatusCli do index.asp */
export async function alterarStatusUnificado(
  codigo: number,
  novoStatus: string,
  usuario: string
): Promise<ResultadoCliente> {
  if (!novoStatus) {
    return atraso({ sucesso: false, mensagem: 'Selecione o novo status.' }, 200)
  }

  const registro = unificados.find((u) => u.codigo === codigo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)

  const anterior = registro.status
  registro.status = novoStatus
  registro.vistos.push({
    data: new Date().toISOString().slice(0, 10),
    usuario,
    acao: `Status alterado de ${descricaoStatusUnificado(anterior)} para ${descricaoStatusUnificado(novoStatus)}`,
    observacao: '',
    leadtimeDias: 0
  })

  return atraso({ sucesso: true, mensagem: 'Status alterado com sucesso!' }, 500)
}

export async function anexarDocumento(codigo: number, arquivo: string): Promise<ResultadoCliente> {
  if (!arquivo.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o nome do documento.' }, 200)
  }

  const registro = unificados.find((u) => u.codigo === codigo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Cadastro não encontrado.' }, 200)

  registro.anexos.push(arquivo.trim())
  return atraso({ sucesso: true, mensagem: 'Documento anexado com sucesso!' }, 450)
}

/* ------------------------------------------------------------------ *
 * Solicitações (Alteração de Dados e Transferência/Extensão)
 * ------------------------------------------------------------------ */

/** Fluxo da Alteração de Dados do Cliente */
export const STATUS_ALTERACAO = [
  { codigo: 'EM_DIGITACAO', descricao: 'Em digitação', departamento: 'SOLICITANTE' },
  { codigo: 'PENDENTE_DIRETOR_COMERCIAL', descricao: 'Pendente Diretor Comercial', departamento: 'DIRETOR_COMERCIAL' },
  { codigo: 'PENDENTE_FINANCEIRO', descricao: 'Pendente Financeiro', departamento: 'FINANCEIRO' },
  { codigo: 'PENDENTE_CONTROLADORIA', descricao: 'Pendente Controladoria', departamento: 'CONTROLADORIA' },
  { codigo: 'PENDENTE_FINALIZACAO', descricao: 'Pendente Finalização', departamento: 'CONTROLADORIA' },
  { codigo: 'FINALIZADO', descricao: 'Finalizado', departamento: '' },
  { codigo: 'CANCELADO', descricao: 'Cancelado', departamento: '' }
] as const

/** Fluxo da Transferência/Extensão de Cliente */
export const STATUS_EXTENSAO = [
  { codigo: 'EM_DIGITACAO', descricao: 'Em digitação', departamento: 'SOLICITANTE' },
  { codigo: 'PENDENTE_FISCAL', descricao: 'Pendente Fiscal', departamento: 'FISCAL' },
  { codigo: 'PENDENTE_DIRETOR_COMERCIAL', descricao: 'Pendente Diretor Comercial', departamento: 'DIRETOR_COMERCIAL' },
  { codigo: 'PENDENTE_CDAG', descricao: 'Pendente CDAG', departamento: 'CDAG' },
  { codigo: 'PENDENTE_CONTROLADORIA', descricao: 'Pendente Controladoria', departamento: 'CONTROLADORIA' },
  { codigo: 'FINALIZADO', descricao: 'Finalizado', departamento: '' },
  { codigo: 'CANCELADO', descricao: 'Cancelado', departamento: '' }
] as const

/** Campos que a solicitação de alteração permite mudar */
export const CAMPOS_ALTERACAO = [
  'Razão Social',
  'Nome Fantasia',
  'Endereço',
  'Condição de Pagamento',
  'Limite de Crédito',
  'Tipo de Frete',
  'Código do Vendedor',
  'Transit Time',
  'Inscrição Estadual',
  'Segmento'
]

export type TipoSolicitacaoCliente = 'ALTERACAO' | 'EXTENSAO'

export interface AvaliacaoSolicitacao {
  etapa: string
  usuario: string
  data: string
  aprovado: boolean
  observacao: string
}

export interface SolicitacaoCliente {
  id: number
  tipo: TipoSolicitacaoCliente
  dataInclusao: string
  planta: string
  /** Transferência/extensão informa a planta de destino */
  plantaDestino: string
  codCliente: string
  codLoja: string
  razaoSocial: string
  cnpj: string
  emitente: string
  status: string
  justificativa: string
  /** Alteração: quais campos mudam e para quê */
  camposAlterados: { campo: string; valorAtual: string; valorNovo: string }[]
  /** Extensão: documentação anexada */
  documentos: string[]
  tipoContribuinte: string
  avaliacoes: AvaliacaoSolicitacao[]
}

const solicitacoes: SolicitacaoCliente[] = [
  {
    id: 8600,
    tipo: 'ALTERACAO',
    dataInclusao: iso(4),
    planta: '00',
    plantaDestino: '',
    codCliente: '001234',
    codLoja: '01',
    razaoSocial: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.',
    cnpj: '11222333000144',
    emitente: 'jsilva',
    status: 'PENDENTE_DIRETOR_COMERCIAL',
    justificativa: 'Cliente alterou o endereço da matriz e a condição de pagamento negociada.',
    camposAlterados: [
      { campo: 'Endereço', valorAtual: 'RUA A, 100', valorNovo: 'AV. BRASIL, 2500' },
      { campo: 'Condição de Pagamento', valorAtual: '28 DDL', valorNovo: '30 DDL' }
    ],
    documentos: [],
    tipoContribuinte: '',
    avaliacoes: []
  },
  {
    id: 8601,
    tipo: 'ALTERACAO',
    dataInclusao: iso(12),
    planta: '05',
    plantaDestino: '',
    codCliente: '003891',
    codLoja: '01',
    razaoSocial: 'FRIGORIFICO SERRA AZUL LTDA',
    cnpj: '33444555000166',
    emitente: 'malves',
    status: 'PENDENTE_CONTROLADORIA',
    justificativa: 'Aumento de limite acordado com a diretoria.',
    camposAlterados: [{ campo: 'Limite de Crédito', valorAtual: '400000', valorNovo: '600000' }],
    documentos: [],
    tipoContribuinte: '',
    avaliacoes: [
      {
        etapa: 'Diretor Comercial',
        usuario: 'cpereira',
        data: iso(10),
        aprovado: true,
        observacao: 'De acordo com o plano comercial.'
      },
      {
        etapa: 'Financeiro',
        usuario: 'plima',
        data: iso(8),
        aprovado: true,
        observacao: 'Score do cliente comporta o novo limite.'
      }
    ]
  },
  {
    id: 8700,
    tipo: 'EXTENSAO',
    dataInclusao: iso(6),
    planta: '00',
    plantaDestino: '05',
    codCliente: '002567',
    codLoja: '01',
    razaoSocial: 'BEBIDAS DO VALE LTDA',
    cnpj: '22333444000155',
    emitente: 'cpereira',
    status: 'PENDENTE_FISCAL',
    justificativa: 'Cliente passará a ser atendido também pela planta de Bento.',
    camposAlterados: [],
    documentos: ['cartao-cnpj.pdf', 'inscricao-estadual.pdf'],
    tipoContribuinte: 'Contribuinte ICMS',
    avaliacoes: []
  },
  {
    id: 8701,
    tipo: 'EXTENSAO',
    dataInclusao: iso(20),
    planta: '04',
    plantaDestino: '07',
    codCliente: '004125',
    codLoja: '01',
    razaoSocial: 'HORTIFRUTI NORDESTE COMERCIO S.A.',
    cnpj: '44555666000177',
    emitente: 'rnunes',
    status: 'PENDENTE_CDAG',
    justificativa: 'Abertura de centro de distribuição em Saquarema.',
    camposAlterados: [],
    documentos: ['contrato-social.pdf'],
    tipoContribuinte: 'Isento',
    avaliacoes: [
      {
        etapa: 'Fiscal',
        usuario: 'plima',
        data: iso(18),
        aprovado: true,
        observacao: 'Situação fiscal regular.'
      },
      {
        etapa: 'Diretor Comercial',
        usuario: 'cpereira',
        data: iso(15),
        aprovado: true,
        observacao: 'Aprovado.'
      }
    ]
  }
]

export const TIPOS_CONTRIBUINTE = ['Contribuinte ICMS', 'Não Contribuinte', 'Isento']

function fluxoDe(tipo: TipoSolicitacaoCliente) {
  return tipo === 'ALTERACAO' ? STATUS_ALTERACAO : STATUS_EXTENSAO
}

export function descricaoStatusSolicitacao(tipo: TipoSolicitacaoCliente, codigo: string) {
  return fluxoDe(tipo).find((s) => s.codigo === codigo)?.descricao ?? codigo
}

/** Departamento responsável pela etapa atual */
export function departamentoDaEtapa(tipo: TipoSolicitacaoCliente, codigo: string) {
  return fluxoDe(tipo).find((s) => s.codigo === codigo)?.departamento ?? ''
}

export interface FiltrosSolicitacaoCliente {
  planta: string
  status: string
  cliente: string
}

export const FILTROS_SOLICITACAO_VAZIOS: FiltrosSolicitacaoCliente = {
  planta: '',
  status: '',
  cliente: ''
}

export async function listarSolicitacoesCliente(
  tipo: TipoSolicitacaoCliente,
  filtros: FiltrosSolicitacaoCliente
): Promise<SolicitacaoCliente[]> {
  const cliente = filtros.cliente.trim().toLowerCase()

  const resultado = solicitacoes
    .filter((s) => s.tipo === tipo)
    .filter((s) => !filtros.planta || s.planta === filtros.planta)
    .filter((s) => !filtros.status || s.status === filtros.status)
    .filter((s) => !cliente || s.razaoSocial.toLowerCase().includes(cliente) || s.cnpj.includes(cliente))
    .map((s) => ({ ...s }))
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 400)
}

export interface DadosSolicitacaoCliente {
  id?: number
  tipo: TipoSolicitacaoCliente
  planta: string
  plantaDestino: string
  codCliente: string
  codLoja: string
  razaoSocial: string
  cnpj: string
  justificativa: string
  camposAlterados: { campo: string; valorAtual: string; valorNovo: string }[]
  documentos: string[]
  tipoContribuinte: string
}

export function solicitacaoVazia(tipo: TipoSolicitacaoCliente): DadosSolicitacaoCliente {
  return {
    tipo,
    planta: '',
    plantaDestino: '',
    codCliente: '',
    codLoja: '',
    razaoSocial: '',
    cnpj: '',
    justificativa: '',
    camposAlterados: [],
    documentos: [],
    tipoContribuinte: ''
  }
}

export async function salvarSolicitacaoCliente(
  dados: DadosSolicitacaoCliente,
  usuario: string
): Promise<ResultadoCliente> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.codCliente.trim()) {
    return atraso({ sucesso: false, mensagem: 'Cód. Cliente é obrigatório.' }, 200)
  }
  if (!dados.justificativa.trim()) {
    return atraso({ sucesso: false, mensagem: 'Justificativa é obrigatória.' }, 200)
  }

  if (dados.tipo === 'ALTERACAO' && dados.camposAlterados.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Informe ao menos um campo a alterar.' }, 200)
  }
  if (dados.tipo === 'EXTENSAO') {
    if (!dados.plantaDestino) {
      return atraso({ sucesso: false, mensagem: 'Planta de destino é obrigatória.' }, 200)
    }
    if (dados.plantaDestino === dados.planta) {
      return atraso(
        { sucesso: false, mensagem: 'A planta de destino deve ser diferente da origem.' },
        200
      )
    }
  }

  if (dados.id) {
    const registro = solicitacoes.find((s) => s.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)
    if (registro.status !== 'EM_DIGITACAO') {
      return atraso(
        { sucesso: false, mensagem: 'Só é possível editar solicitações em digitação.' },
        200
      )
    }

    Object.assign(registro, dados)
    return atraso({ sucesso: true, mensagem: 'Solicitação atualizada com sucesso!' }, 550)
  }

  solicitacoes.push({
    ...dados,
    id: Math.max(0, ...solicitacoes.map((s) => s.id)) + 1,
    dataInclusao: new Date().toISOString().slice(0, 10),
    emitente: usuario,
    status: 'EM_DIGITACAO',
    avaliacoes: []
  })

  return atraso({ sucesso: true, mensagem: 'Solicitação registrada com sucesso!' }, 600)
}

/** Avaliação da etapa: aprovar avança no fluxo, rejeitar cancela */
export async function avaliarSolicitacaoCliente(
  id: number,
  aprovar: boolean,
  observacao: string,
  usuario: string
): Promise<ResultadoCliente> {
  const registro = solicitacoes.find((s) => s.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)
  if (!observacao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a observação da avaliação.' }, 200)
  }

  const fluxo = fluxoDe(registro.tipo)
  const indiceAtual = fluxo.findIndex((s) => s.codigo === registro.status)
  const etapa = fluxo[indiceAtual]

  if (!etapa || !etapa.departamento) {
    return atraso({ sucesso: false, mensagem: 'Esta solicitação não está em etapa de avaliação.' }, 200)
  }

  registro.avaliacoes.push({
    etapa: descricaoStatusSolicitacao(registro.tipo, registro.status),
    usuario,
    data: new Date().toISOString().slice(0, 10),
    aprovado: aprovar,
    observacao: observacao.trim()
  })

  if (!aprovar) {
    registro.status = 'CANCELADO'
    return atraso({ sucesso: true, mensagem: 'Solicitação reprovada e cancelada.' }, 550)
  }

  const proxima = fluxo[indiceAtual + 1]
  registro.status = proxima?.codigo ?? 'FINALIZADO'

  return atraso(
    {
      sucesso: true,
      mensagem:
        registro.status === 'FINALIZADO'
          ? 'Solicitação finalizada!'
          : `Encaminhada para ${descricaoStatusSolicitacao(registro.tipo, registro.status)}.`
    },
    600
  )
}

/** Envia a solicitação em digitação para a primeira etapa de avaliação */
export async function enviarSolicitacaoCliente(id: number): Promise<ResultadoCliente> {
  const registro = solicitacoes.find((s) => s.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)
  if (registro.status !== 'EM_DIGITACAO') {
    return atraso({ sucesso: false, mensagem: 'Esta solicitação já foi enviada.' }, 200)
  }
  if (registro.tipo === 'EXTENSAO' && registro.documentos.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Anexe a documentação antes de enviar.' }, 200)
  }

  const fluxo = fluxoDe(registro.tipo)
  registro.status = fluxo[1].codigo

  return atraso(
    { sucesso: true, mensagem: `Enviada para ${descricaoStatusSolicitacao(registro.tipo, registro.status)}.` },
    550
  )
}

export async function cancelarSolicitacaoCliente(id: number): Promise<ResultadoCliente> {
  const registro = solicitacoes.find((s) => s.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)
  if (registro.status === 'FINALIZADO') {
    return atraso({ sucesso: false, mensagem: 'Solicitação finalizada não pode ser cancelada.' }, 200)
  }

  registro.status = 'CANCELADO'
  return atraso({ sucesso: true, mensagem: 'Solicitação cancelada.' }, 500)
}
