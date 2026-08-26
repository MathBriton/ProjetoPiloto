import { atraso, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * Financeiro — Alterar Risco (SA1010), Cadastro de Canhoto (CANHOTOS),
 * Validadores de Crédito (ValidadoresCredito) e as duas solicitações de
 * crédito.
 *
 * Cada planta consulta um DADOSADV próprio no legado; aqui a planta é só um
 * campo dos registros em memória.
 */

/** Plantas com DADOSADV dedicado (Alterar Risco, Canhoto, Transferência) */
export const PLANTAS_PROTHEUS = [
  { codigo: '00', nome: 'Pirapetinga' },
  { codigo: '04', nome: 'Uberaba' },
  { codigo: '05', nome: 'Bento Gonçalves' },
  { codigo: '06', nome: 'Maranguape' },
  { codigo: '07', nome: 'Saquarema' }
]

/** Lista ampliada usada na Solicitação de Aumento de Crédito */
export const PLANTAS_CREDITO = [
  ...PLANTAS_PROTHEUS,
  { codigo: '08', nome: 'Araçatuba' },
  { codigo: '09', nome: 'Blumenau' },
  { codigo: '10', nome: 'Pacajus' },
  { codigo: '11', nome: 'Porto Feliz' },
  { codigo: '12', nome: 'Três Barras' },
  { codigo: '13', nome: 'Florestal' }
]

export function nomePlanta(codigo: string) {
  return PLANTAS_CREDITO.find((p) => p.codigo === codigo)?.nome ?? codigo
}

/* ------------------------------------------------------------------ *
 * Clientes (SA1010)
 * ------------------------------------------------------------------ */

export interface ClienteFinanceiro {
  codigo: string
  loja: string
  nome: string
  cnpj: string
  municipio: string
  uf: string
  planta: string
  /** A1_RISCO */
  risco: string
  limiteAtual: number
  limiteAberto: number
  codigoPrincipal: string
  nomePrincipal: string
  condicaoPagamento: string
}

const clientes: ClienteFinanceiro[] = [
  { codigo: '001234', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', cnpj: '11222333000144', municipio: 'JUIZ DE FORA', uf: 'MG', planta: '00', risco: 'A', limiteAtual: 250_000, limiteAberto: 87_500, codigoPrincipal: '001234', nomePrincipal: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', condicaoPagamento: '28 DDL' },
  { codigo: '001234', loja: '02', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A. - FILIAL', cnpj: '11222333000225', municipio: 'CATAGUASES', uf: 'MG', planta: '00', risco: 'B', limiteAtual: 80_000, limiteAberto: 12_300, codigoPrincipal: '001234', nomePrincipal: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', condicaoPagamento: '28 DDL' },
  { codigo: '002567', loja: '01', nome: 'BEBIDAS DO VALE LTDA', cnpj: '22333444000155', municipio: 'CATAGUASES', uf: 'MG', planta: '00', risco: 'C', limiteAtual: 120_000, limiteAberto: 118_900, codigoPrincipal: '002567', nomePrincipal: 'BEBIDAS DO VALE LTDA', condicaoPagamento: '21 DDL' },
  { codigo: '003891', loja: '01', nome: 'FRIGORIFICO SERRA AZUL LTDA', cnpj: '33444555000166', municipio: 'BENTO GONCALVES', uf: 'RS', planta: '05', risco: 'A', limiteAtual: 400_000, limiteAberto: 145_000, codigoPrincipal: '003891', nomePrincipal: 'FRIGORIFICO SERRA AZUL LTDA', condicaoPagamento: '30 DDL' },
  { codigo: '004125', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.', cnpj: '44555666000177', municipio: 'MARANGUAPE', uf: 'CE', planta: '06', risco: 'P', limiteAtual: 95_000, limiteAberto: 95_000, codigoPrincipal: '004125', nomePrincipal: 'HORTIFRUTI NORDESTE COMERCIO S.A.', condicaoPagamento: '14 DDL' },
  { codigo: '005678', loja: '01', nome: 'LATICINIOS CAMPO BELO LTDA', cnpj: '55666777000188', municipio: 'UBERABA', uf: 'MG', planta: '04', risco: '', limiteAtual: 0, limiteAberto: 0, codigoPrincipal: '005678', nomePrincipal: 'LATICINIOS CAMPO BELO LTDA', condicaoPagamento: '28 DDL' },
  { codigo: '006432', loja: '01', nome: 'DISTRIBUIDORA LITORAL SUL LTDA', cnpj: '66777888000199', municipio: 'SAQUAREMA', uf: 'RJ', planta: '07', risco: 'D', limiteAtual: 60_000, limiteAberto: 21_400, codigoPrincipal: '006432', nomePrincipal: 'DISTRIBUIDORA LITORAL SUL LTDA', condicaoPagamento: '21 DDL' },
  { codigo: '007219', loja: '01', nome: 'EMBALAGENS RAPIDAS COMERCIO LTDA', cnpj: '77888999000110', municipio: 'UBERABA', uf: 'MG', planta: '04', risco: 'E', limiteAtual: 35_000, limiteAberto: 8_200, codigoPrincipal: '007219', nomePrincipal: 'EMBALAGENS RAPIDAS COMERCIO LTDA', condicaoPagamento: '14 DDL' }
]

export function saldoDisponivel(cliente: ClienteFinanceiro) {
  return cliente.limiteAtual - cliente.limiteAberto
}

export async function buscarClientesFinanceiro(
  planta: string,
  termo: string
): Promise<ClienteFinanceiro[]> {
  const busca = termo.trim().toLowerCase()

  const resultado = clientes
    .filter((c) => !planta || c.planta === planta)
    .filter((c) => !busca || c.nome.toLowerCase().includes(busca) || c.codigo.includes(busca))

  return atraso(resultado.map((c) => ({ ...c })), 400)
}

export async function carregarCliente(
  planta: string,
  codigo: string,
  loja: string
): Promise<ClienteFinanceiro | null> {
  const encontrado = clientes.find(
    (c) =>
      c.planta === planta &&
      c.codigo === codigo.trim() &&
      (loja.trim() === '' || c.loja === loja.trim())
  )
  return atraso(encontrado ? { ...encontrado } : null, 350)
}

/* ------------------------------------------------------------------ *
 * Alterar Risco
 * ------------------------------------------------------------------ */

export interface OpcaoRisco {
  valor: string
  rotulo: string
}

/**
 * Opções de risco do `_riscos1()` do legado — dependem do risco atual e de a
 * loja ter sido informada:
 * - sem risco atual → só A e E
 * - com risco e sem loja → A a F e P
 * - com risco e com loja → J e A
 */
export function opcoesRisco(riscoAtual: string, loja: string): OpcaoRisco[] {
  if (!riscoAtual.trim()) {
    return [
      { valor: 'A', rotulo: 'RISCO A' },
      { valor: 'E', rotulo: 'RISCO E' }
    ]
  }

  if (!loja.trim()) {
    return [
      { valor: 'A', rotulo: 'RISCO A' },
      { valor: 'B', rotulo: 'RISCO B' },
      { valor: 'C', rotulo: 'RISCO C' },
      { valor: 'D', rotulo: 'RISCO D' },
      { valor: 'E', rotulo: 'RISCO E' },
      { valor: 'F', rotulo: 'RISCO F - Bloqueia só Faturamento' },
      { valor: 'P', rotulo: 'RISCO P - Bloqueia Novos Pedidos' }
    ]
  }

  return [
    { valor: 'J', rotulo: 'RISCO J - Bloqueia Novos Pedidos desta loja' },
    { valor: 'A', rotulo: 'RISCO A' }
  ]
}

export interface ResultadoFinanceiro {
  sucesso: boolean
  mensagem: string
}

export async function gravarRisco(
  planta: string,
  codigo: string,
  loja: string,
  risco: string
): Promise<ResultadoFinanceiro> {
  if (!planta || !codigo || !risco) {
    return atraso({ sucesso: false, mensagem: 'Parâmetros insuficientes.' }, 200)
  }

  // O UPDATE do legado só filtra por loja quando ela é informada
  const alvos = clientes.filter(
    (c) => c.planta === planta && c.codigo === codigo && (loja === '' || c.loja === loja)
  )
  if (alvos.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Cliente não encontrado.' }, 200)
  }

  for (const cliente of alvos) cliente.risco = risco
  return atraso(
    { sucesso: true, mensagem: `Risco alterado para ${risco} em ${alvos.length} registro(s).` },
    500
  )
}

/* ------------------------------------------------------------------ *
 * Cadastro de Canhoto
 * ------------------------------------------------------------------ */

/** Escala de satisfação do canhoto (0 = não respondido) */
export const ESCALA_SATISFACAO = [
  { valor: 0, rotulo: 'Não respondido' },
  { valor: 5, rotulo: 'Muito Satisfeito' },
  { valor: 4, rotulo: 'Satisfeito' },
  { valor: 3, rotulo: 'Indiferente' },
  { valor: 2, rotulo: 'Insatisfeito' },
  { valor: 1, rotulo: 'Muito insatisfeito' }
]

export function rotuloSatisfacao(valor: number) {
  return ESCALA_SATISFACAO.find((e) => e.valor === valor)?.rotulo ?? '—'
}

export interface Canhoto {
  id: number
  planta: string
  notaFiscal: string
  serie: string
  data: string
  placa: string
  /** Avaliação do produto (carga) */
  produto: number
  /** Avaliação do serviço (veículo) */
  servico: number
  observacao: string
  incluidoPor: string
  alteradoPor: string
}

export interface DadosCanhoto {
  id?: number
  planta: string
  notaFiscal: string
  serie: string
  data: string
  placa: string
  produto: number | ''
  servico: number | ''
  observacao: string
}

interface CanhotoInterno extends Canhoto {
  excluido: boolean
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const canhotos: CanhotoInterno[] = Array.from({ length: 32 }, (_, indice) => ({
  id: 4000 + indice,
  planta: PLANTAS_PROTHEUS[indice % PLANTAS_PROTHEUS.length].codigo,
  notaFiscal: String(120_000 + indice * 11),
  serie: indice % 3 === 0 ? '1' : '2',
  data: iso(indice),
  placa: `MNO${String(3000 + indice).slice(-4)}`,
  produto: [5, 4, 3, 4, 5, 2, 0][indice % 7],
  servico: [4, 5, 3, 5, 4, 1, 0][indice % 7],
  observacao: indice % 6 === 0 ? 'Entrega com atraso de 2 horas.' : '',
  incluidoPor: ['admin', 'jsilva', 'malves'][indice % 3],
  alteradoPor: indice % 4 === 0 ? 'admin' : '',
  excluido: false
}))

export type CampoPesquisaCanhoto = '' | 'PLACA' | 'DATA' | 'NF'

export interface FiltrosCanhoto {
  planta: string
  campo: CampoPesquisaCanhoto
  valor: string
}

export async function listarCanhotos(filtros: FiltrosCanhoto): Promise<Canhoto[]> {
  const valor = filtros.valor.trim().toLowerCase()

  const resultado = canhotos
    .filter((c) => !c.excluido)
    .filter((c) => !filtros.planta || c.planta === filtros.planta)
    .filter((c) => {
      if (!valor || !filtros.campo) return true
      if (filtros.campo === 'PLACA') return c.placa.toLowerCase().includes(valor)
      if (filtros.campo === 'NF') return c.notaFiscal.includes(valor)
      return c.data === filtros.valor
    })
    .sort((a, b) => b.data.localeCompare(a.data))
    .map(({ excluido, ...resto }) => resto)

  return atraso(resultado, 400)
}

export async function salvarCanhoto(
  dados: DadosCanhoto,
  usuario: string
): Promise<ResultadoFinanceiro> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.notaFiscal.trim()) {
    return atraso({ sucesso: false, mensagem: 'Nota Fiscal é obrigatória.' }, 200)
  }
  if (!dados.data) return atraso({ sucesso: false, mensagem: 'Data é obrigatória.' }, 200)
  if (dados.produto === '' || dados.servico === '') {
    return atraso({ sucesso: false, mensagem: 'Informe as avaliações de produto e serviço.' }, 200)
  }

  const campos = {
    planta: dados.planta,
    notaFiscal: dados.notaFiscal.trim(),
    serie: dados.serie.trim(),
    data: dados.data,
    placa: dados.placa.trim().toUpperCase(),
    produto: Number(dados.produto),
    servico: Number(dados.servico),
    observacao: dados.observacao.trim()
  }

  if (dados.id) {
    const canhoto = canhotos.find((c) => c.id === dados.id)
    if (!canhoto) return atraso({ sucesso: false, mensagem: 'Canhoto não encontrado.' }, 200)

    Object.assign(canhoto, campos, { alteradoPor: usuario })
    return atraso({ sucesso: true, mensagem: 'Cadastro atualizado com sucesso!' }, 450)
  }

  const duplicado = canhotos.some(
    (c) => !c.excluido && c.planta === campos.planta && c.notaFiscal === campos.notaFiscal && c.serie === campos.serie
  )
  if (duplicado) {
    return atraso({ sucesso: false, mensagem: 'Já existe canhoto para esta nota e série.' }, 200)
  }

  canhotos.push({
    id: Math.max(0, ...canhotos.map((c) => c.id)) + 1,
    ...campos,
    incluidoPor: usuario,
    alteradoPor: '',
    excluido: false
  })

  return atraso({ sucesso: true, mensagem: 'Cadastro realizado com sucesso!' }, 450)
}

export async function excluirCanhoto(id: number): Promise<ResultadoFinanceiro> {
  const canhoto = canhotos.find((c) => c.id === id)
  if (!canhoto) return atraso({ sucesso: false, mensagem: 'Canhoto não encontrado.' }, 200)

  canhoto.excluido = true
  return atraso({ sucesso: true, mensagem: 'Cadastro apagado com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Validadores de Crédito
 * ------------------------------------------------------------------ */

/** ValidadoresCredito.NIVEL */
export const NIVEIS_VALIDADOR = [
  { valor: '10', rotulo: 'Analista de Crédito' },
  { valor: '20', rotulo: 'Coordenador de Crédito' },
  { valor: '30', rotulo: 'Gerente de Tesouraria' },
  { valor: '40', rotulo: 'Divisional Controller' },
  { valor: '50', rotulo: 'Gerente de Planta' },
  { valor: '60', rotulo: 'CFO' }
]

export function rotuloNivel(valor: string) {
  return NIVEIS_VALIDADOR.find((n) => n.valor === valor)?.rotulo ?? valor
}

export interface ValidadorCredito {
  /** VLD_COD — é o CRM_R_E_C_N_O do usuário */
  codigo: number
  nome: string
  plantas: string[]
  nivel: string
  valorMinimo: number
  valorMaximo: number
}

export interface DadosValidador {
  codigo?: number
  idUsuario: number | ''
  plantas: string[]
  nivel: string
  valorMinimo: string
  valorMaximo: string
}

interface ValidadorInterno extends ValidadorCredito {
  excluido: boolean
}

const validadores: ValidadorInterno[] = [
  { codigo: 1007, nome: 'PATRÍCIA LIMA', plantas: ['00', '04'], nivel: '10', valorMinimo: 0, valorMaximo: 50_000, excluido: false },
  { codigo: 1003, nome: 'MARIA ALVES', plantas: ['05'], nivel: '20', valorMinimo: 50_000, valorMaximo: 150_000, excluido: false },
  { codigo: 1008, nome: 'BRUNO CARVALHO', plantas: ['00', '05', '12'], nivel: '30', valorMinimo: 150_000, valorMaximo: 400_000, excluido: false },
  { codigo: 1001, nome: 'ADMINISTRADOR (FICTÍCIO)', plantas: ['00', '04', '05', '06', '07'], nivel: '60', valorMinimo: 400_000, valorMaximo: 9_999_999, excluido: false }
]

export async function listarValidadores(termo = ''): Promise<ValidadorCredito[]> {
  const busca = termo.trim().toLowerCase()

  const resultado = validadores
    .filter((v) => !v.excluido)
    .filter((v) => !busca || v.nome.toLowerCase().includes(busca) || rotuloNivel(v.nivel).toLowerCase().includes(busca))
    .sort((a, b) => Number(a.nivel) - Number(b.nivel) || a.nome.localeCompare(b.nome, 'pt-BR'))
    .map(({ excluido, ...resto }) => resto)

  return atraso(resultado, 350)
}

export async function salvarValidador(dados: DadosValidador): Promise<ResultadoFinanceiro> {
  if (!dados.idUsuario) return atraso({ sucesso: false, mensagem: 'Validador é obrigatório.' }, 200)
  if (!dados.nivel) return atraso({ sucesso: false, mensagem: 'Nível é obrigatório.' }, 200)
  if (dados.plantas.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos uma planta.' }, 200)
  }

  const minimo = Number(dados.valorMinimo)
  const maximo = Number(dados.valorMaximo)
  if (Number.isNaN(minimo) || Number.isNaN(maximo)) {
    return atraso({ sucesso: false, mensagem: 'Informe valores numéricos válidos.' }, 200)
  }
  if (minimo > maximo) {
    return atraso({ sucesso: false, mensagem: 'O valor mínimo não pode ser maior que o máximo.' }, 200)
  }

  const usuario = usuariosSwnet.find((u) => u.id === Number(dados.idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 200)

  const existente = validadores.find((v) => v.codigo === usuario.id && !v.excluido)
  if (existente && (!dados.codigo || dados.codigo === existente.codigo)) {
    Object.assign(existente, {
      plantas: [...dados.plantas],
      nivel: dados.nivel,
      valorMinimo: minimo,
      valorMaximo: maximo
    })
    return atraso({ sucesso: true, mensagem: 'Cadastro atualizado com sucesso!' }, 450)
  }

  validadores.push({
    codigo: usuario.id,
    nome: usuario.nome,
    plantas: [...dados.plantas],
    nivel: dados.nivel,
    valorMinimo: minimo,
    valorMaximo: maximo,
    excluido: false
  })

  return atraso({ sucesso: true, mensagem: 'Cadastro realizado com sucesso!' }, 450)
}

export async function excluirValidador(codigo: number): Promise<ResultadoFinanceiro> {
  const validador = validadores.find((v) => v.codigo === codigo)
  if (!validador) return atraso({ sucesso: false, mensagem: 'Validador não encontrado.' }, 200)

  validador.excluido = true
  return atraso({ sucesso: true, mensagem: 'Cadastro apagado com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Solicitação de Aumento de Crédito
 * ------------------------------------------------------------------ */

export type TipoMercado = '' | 'NACIONAL' | 'EXPORTACAO'

export interface DadosSolicitacaoCredito {
  planta: string
  codCliente: string
  codLoja: string
  outrasVendas: boolean
  precoM2: string
  precoKg: string
  volumeM2: string
  volumeUnidades: string
  volumeToneladas: string
  tipoMercado: TipoMercado
  limiteSolicitado: string
  observacao: string
  contratoSocial: string
  balancoDre: string
  documentosComplementares: string
}

export const SOLICITACAO_CREDITO_VAZIA: DadosSolicitacaoCredito = {
  planta: '',
  codCliente: '',
  codLoja: '',
  outrasVendas: false,
  precoM2: '',
  precoKg: '',
  volumeM2: '',
  volumeUnidades: '',
  volumeToneladas: '',
  tipoMercado: '',
  limiteSolicitado: '',
  observacao: '',
  contratoSocial: '',
  balancoDre: '',
  documentosComplementares: ''
}

export async function solicitarAumentoCredito(
  dados: DadosSolicitacaoCredito
): Promise<ResultadoFinanceiro & { protocolo?: string }> {
  if (!dados.planta || !dados.codCliente.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a planta e o código do cliente.' }, 200)
  }
  if (!dados.tipoMercado) {
    return atraso({ sucesso: false, mensagem: 'Selecione o tipo de mercado.' }, 200)
  }

  const limite = Number(dados.limiteSolicitado.replace(/\./g, '').replace(',', '.'))
  if (!limite || Number.isNaN(limite) || limite <= 0) {
    return atraso({ sucesso: false, mensagem: 'Informe um limite desejado válido.' }, 200)
  }
  if (!dados.volumeM2.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o potencial de volume MSM mensal.' }, 200)
  }
  // Documentos são obrigatórios quando não é venda de resíduos/aparas
  if (!dados.outrasVendas && (!dados.contratoSocial || !dados.balancoDre)) {
    return atraso(
      { sucesso: false, mensagem: 'Anexe o Contrato Social e o Balanço/DRE do exercício.' },
      200
    )
  }

  const protocolo = `SAC-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`
  return atraso(
    { sucesso: true, mensagem: `Solicitação registrada com o protocolo ${protocolo}.`, protocolo },
    800
  )
}

/* ------------------------------------------------------------------ *
 * Transferência de Crédito
 * ------------------------------------------------------------------ */

export interface DadosTransferenciaCredito {
  plantaOrigem: string
  codClienteOrigem: string
  codLojaOrigem: string
  plantaDestino: string
  codClienteDestino: string
  codLojaDestino: string
  valorTransferencia: string
}

export const TRANSFERENCIA_VAZIA: DadosTransferenciaCredito = {
  plantaOrigem: '',
  codClienteOrigem: '',
  codLojaOrigem: '',
  plantaDestino: '',
  codClienteDestino: '',
  codLojaDestino: '',
  valorTransferencia: ''
}

/**
 * TransferenciaCreditoValidar.asp — o saldo disponível na origem tem de
 * cobrir o valor transferido.
 */
export async function validarTransferencia(
  dados: DadosTransferenciaCredito
): Promise<ResultadoFinanceiro> {
  if (!dados.plantaOrigem || !dados.codClienteOrigem.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a planta e o cliente de origem.' }, 200)
  }
  if (!dados.plantaDestino || !dados.codClienteDestino.trim()) {
    return atraso({ sucesso: false, mensagem: 'Selecione o cliente de destino.' }, 200)
  }

  const mesmoCliente =
    dados.plantaOrigem === dados.plantaDestino &&
    dados.codClienteOrigem === dados.codClienteDestino &&
    dados.codLojaOrigem === dados.codLojaDestino
  if (mesmoCliente) {
    return atraso({ sucesso: false, mensagem: 'Origem e destino não podem ser o mesmo cliente.' }, 200)
  }

  const valor = Number(dados.valorTransferencia.replace(/\./g, '').replace(',', '.'))
  if (!valor || Number.isNaN(valor) || valor <= 0) {
    return atraso({ sucesso: false, mensagem: 'Informe um valor de transferência válido.' }, 200)
  }

  const origem = clientes.find(
    (c) =>
      c.planta === dados.plantaOrigem &&
      c.codigo === dados.codClienteOrigem &&
      c.loja === dados.codLojaOrigem
  )
  if (!origem) return atraso({ sucesso: false, mensagem: 'Cliente de origem não encontrado.' }, 200)

  if (valor > saldoDisponivel(origem)) {
    return atraso(
      { sucesso: false, mensagem: 'Valor maior que o saldo disponível no cliente de origem.' },
      200
    )
  }

  return atraso({ sucesso: true, mensagem: 'Transferência válida.' }, 400)
}

export async function transferirCredito(
  dados: DadosTransferenciaCredito
): Promise<ResultadoFinanceiro> {
  const validacao = await validarTransferencia(dados)
  if (!validacao.sucesso) return validacao

  const valor = Number(dados.valorTransferencia.replace(/\./g, '').replace(',', '.'))

  const origem = clientes.find(
    (c) =>
      c.planta === dados.plantaOrigem &&
      c.codigo === dados.codClienteOrigem &&
      c.loja === dados.codLojaOrigem
  )
  const destino = clientes.find(
    (c) =>
      c.planta === dados.plantaDestino &&
      c.codigo === dados.codClienteDestino &&
      c.loja === dados.codLojaDestino
  )

  if (!origem || !destino) {
    return atraso({ sucesso: false, mensagem: 'Cliente não encontrado.' }, 200)
  }

  origem.limiteAtual -= valor
  destino.limiteAtual += valor

  return atraso({ sucesso: true, mensagem: 'Transferência de crédito realizada com sucesso!' }, 700)
}
