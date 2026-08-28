import { atraso } from '@/services/comuns/baseMock'

/**
 * Pedidos — emissão, pedidos digitados, solicitação de alteração,
 * orçamento, parâmetros, entrega no fim de semana e observação em NF.
 *
 * As telas do legado vêm em pares Brasil / PiraUbe: mesma mecânica,
 * conjuntos de planta diferentes.
 */

export interface ResultadoPedido {
  sucesso: boolean
  mensagem: string
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

/** Regiões atendidas — cada uma com seu conjunto de plantas */
export type RegiaoPedido = 'BRASIL' | 'PIRAUBE'

export const PLANTAS_POR_REGIAO: Record<RegiaoPedido, { codigo: string; nome: string }[]> = {
  BRASIL: [
    { codigo: '05', nome: 'Bento Gonçalves' },
    { codigo: '06', nome: 'Maranguape' },
    { codigo: '07', nome: 'Saquarema' }
  ],
  PIRAUBE: [
    { codigo: '00', nome: 'Pirapetinga' },
    { codigo: '04', nome: 'Uberaba' }
  ]
}

export function plantasDaRegiao(regiao: RegiaoPedido) {
  return PLANTAS_POR_REGIAO[regiao]
}

/* ------------------------------------------------------------------ *
 * Tabelas de apoio
 * ------------------------------------------------------------------ */

export const TIPOS_ITEM_PEDIDO = [
  { codigo: 'CX', descricao: 'Caixa' },
  { codigo: 'CH', descricao: 'Chapa' },
  { codigo: 'PA', descricao: 'Papel / Bobina' },
  { codigo: 'CV', descricao: 'Corte e Vinco' },
  { codigo: 'ML', descricao: 'Maleta' }
]

export const CONDICOES_PAGAMENTO_PEDIDO = ['14 DDL', '21 DDL', '28 DDL', '30 DDL', '45 DDL']
export const TIPOS_FRETE_PEDIDO = ['CIF', 'FOB']
export const MOEDAS = ['BRL', 'USD', 'EUR']
export const MODAIS_FRETE = ['Rodoviário', 'Marítimo', 'Aéreo']
export const INCOTERMS = ['EXW', 'FOB', 'CIF', 'DAP', 'DDP']
export const CODIGOS_TES = ['501', '502', '511', '512', '601']
export const GRAMATURAS = ['110', '125', '150', '175', '200', '250']
export const TUBETES = ['76 mm', '100 mm', '150 mm']
export const ESTRUTURAS_PEDIDO = ['B', 'C', 'E', 'BC', 'BCA', 'EB']

export interface ClientePedido {
  codigo: string
  loja: string
  nome: string
  cnpj: string
  municipio: string
  uf: string
  /** Saldo de aparas disponível para abatimento */
  saldoAparas: number
  condicaoPagamento: string
}

const clientes: ClientePedido[] = [
  { codigo: '001234', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', cnpj: '11222333000144', municipio: 'JUIZ DE FORA', uf: 'MG', saldoAparas: 12_500, condicaoPagamento: '28 DDL' },
  { codigo: '002567', loja: '01', nome: 'BEBIDAS DO VALE LTDA', cnpj: '22333444000155', municipio: 'CATAGUASES', uf: 'MG', saldoAparas: 0, condicaoPagamento: '21 DDL' },
  { codigo: '003891', loja: '01', nome: 'FRIGORIFICO SERRA AZUL LTDA', cnpj: '33444555000166', municipio: 'BENTO GONCALVES', uf: 'RS', saldoAparas: 4_800, condicaoPagamento: '30 DDL' },
  { codigo: '004125', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.', cnpj: '44555666000177', municipio: 'MARANGUAPE', uf: 'CE', saldoAparas: 0, condicaoPagamento: '14 DDL' },
  { codigo: '006432', loja: '01', nome: 'DISTRIBUIDORA LITORAL SUL LTDA', cnpj: '66777888000199', municipio: 'SAQUAREMA', uf: 'RJ', saldoAparas: 2_100, condicaoPagamento: '21 DDL' }
]

export interface ProdutoPedido {
  codigo: string
  referencia: string
  tipoItem: string
  estrutura: string
  /** Preço de tabela; o orçamento pode sobrescrever */
  valorUnitario: number
}

const produtos: ProdutoPedido[] = [
  { codigo: '901101', referencia: 'CAIXA MAÇÃ 18KG', tipoItem: 'CX', estrutura: 'BC', valorUnitario: 3.45 },
  { codigo: '901102', referencia: 'CAIXA BEBIDA 12X1L', tipoItem: 'CX', estrutura: 'C', valorUnitario: 2.87 },
  { codigo: '903201', referencia: 'CHAPA ONDULADA 1200X800', tipoItem: 'CH', estrutura: 'B', valorUnitario: 1.92 },
  { codigo: '905301', referencia: 'BOBINA KRAFTLINER 175G', tipoItem: 'PA', estrutura: '', valorUnitario: 4.15 },
  { codigo: '907401', referencia: 'BANDEJA CORTE E VINCO', tipoItem: 'CV', estrutura: 'E', valorUnitario: 5.6 }
]

export async function buscarClientesPedido(termo: string): Promise<ClientePedido[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = busca
    ? clientes.filter((c) => c.nome.toLowerCase().includes(busca) || c.codigo.includes(busca))
    : clientes
  return atraso(resultado.map((c) => ({ ...c })), 350)
}

export async function buscarProdutosPedido(termo: string, tipoItem = ''): Promise<ProdutoPedido[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = produtos
    .filter((p) => !tipoItem || p.tipoItem === tipoItem)
    .filter((p) => !busca || p.referencia.toLowerCase().includes(busca) || p.codigo.includes(busca))
  return atraso(resultado.map((p) => ({ ...p })), 350)
}

/* ------------------------------------------------------------------ *
 * Emissão de pedido
 * ------------------------------------------------------------------ */

export interface ItemPedido {
  tipoItem: string
  codigoProduto: string
  referencia: string
  estrutura: string
  quantidade: string
  valorUnitario: string
  /** Dimensional — usado por caixa/chapa */
  largura: string
  comprimento: string
  vinco1: string
  vinco2: string
  vinco3: string
  vinco4: string
  vinco5: string
  /** Específicos de bobina */
  gramaturaBobina: string
  diametroBobina: string
  larguraBobina: string
  tubeteBobina: string
  dataEntrega: string
  dataPrevistaEntrega: string
  pedidoCompra: string
  itemPedido: string
  observacao: string
}

export const ITEM_PEDIDO_VAZIO: ItemPedido = {
  tipoItem: '',
  codigoProduto: '',
  referencia: '',
  estrutura: '',
  quantidade: '',
  valorUnitario: '',
  largura: '',
  comprimento: '',
  vinco1: '',
  vinco2: '',
  vinco3: '',
  vinco4: '',
  vinco5: '',
  gramaturaBobina: '',
  diametroBobina: '',
  larguraBobina: '',
  tubeteBobina: '',
  dataEntrega: '',
  dataPrevistaEntrega: '',
  pedidoCompra: '',
  itemPedido: '',
  observacao: ''
}

export interface DadosEmissaoPedido {
  planta: string
  codCliente: string
  codLoja: string
  nomeCliente: string
  tes: string
  codClienteRemessa: string
  codLojaRemessa: string
  nomeClienteRemessa: string
  saldoAparas: string
  condicaoPagamento: string
  moeda: string
  tipoFrete: string
  modalFrete: string
  incoterms: string
  lotePiloto: boolean
  itens: ItemPedido[]
}

export const EMISSAO_VAZIA: DadosEmissaoPedido = {
  planta: '',
  codCliente: '',
  codLoja: '',
  nomeCliente: '',
  tes: '',
  codClienteRemessa: '',
  codLojaRemessa: '',
  nomeClienteRemessa: '',
  saldoAparas: '',
  condicaoPagamento: '',
  moeda: 'BRL',
  tipoFrete: '',
  modalFrete: '',
  incoterms: '',
  lotePiloto: false,
  itens: []
}

export interface AvisoPedido {
  tipo: 'orcamento' | 'lotePiloto' | 'maquinaInativa' | 'pedidoDuplicado'
  mensagem: string
  bloqueia: boolean
}

/**
 * Reproduz as verificações que o EmissaoPedidoRepositorio.asp faz antes de
 * gravar: orçamento vigente, lote piloto, máquina inativa e pedido de
 * compra duplicado.
 */
export async function validarItensPedido(dados: DadosEmissaoPedido): Promise<AvisoPedido[]> {
  const avisos: AvisoPedido[] = []

  for (const item of dados.itens) {
    const produto = produtos.find((p) => p.codigo === item.codigoProduto)
    if (!produto) continue

    // verificarExisteOrcamento: sem orçamento vigente o preço é o de tabela
    const valorInformado = Number(item.valorUnitario || 0)
    if (valorInformado > 0 && Math.abs(valorInformado - produto.valorUnitario) > 0.001) {
      avisos.push({
        tipo: 'orcamento',
        mensagem: `${produto.referencia}: valor informado difere do orçamento vigente (${produto.valorUnitario}).`,
        bloqueia: false
      })
    }

    // verificarExisteLotePiloto
    if (dados.lotePiloto && produto.tipoItem === 'PA') {
      avisos.push({
        tipo: 'lotePiloto',
        mensagem: `${produto.referencia}: bobina não é elegível a lote piloto.`,
        bloqueia: true
      })
    }

    // verificarExisteMaquinaInativa
    if (produto.tipoItem === 'CV') {
      avisos.push({
        tipo: 'maquinaInativa',
        mensagem: `${produto.referencia}: a máquina de corte e vinco está inativa na planta selecionada.`,
        bloqueia: false
      })
    }
  }

  // verificarPedidoCompraDuplicado
  const compras = dados.itens.map((i) => i.pedidoCompra.trim()).filter(Boolean)
  const duplicados = compras.filter((valor, indice) => compras.indexOf(valor) !== indice)
  for (const duplicado of [...new Set(duplicados)]) {
    avisos.push({
      tipo: 'pedidoDuplicado',
      mensagem: `Pedido de compra ${duplicado} repetido em mais de um item.`,
      bloqueia: true
    })
  }

  return atraso(avisos, 500)
}

export interface PedidoDigitado {
  id: number
  regiao: RegiaoPedido
  planta: string
  pedido: string
  codCliente: string
  nomeCliente: string
  codigoProduto: string
  referencia: string
  quantidade: number
  preco: number
  dataEmissao: string
  dataEntrega: string
  emitidoPor: string
  excluido: boolean
}

const pedidosDigitados: PedidoDigitado[] = Array.from({ length: 30 }, (_, indice) => {
  const cliente = clientes[indice % clientes.length]
  const produto = produtos[indice % produtos.length]
  const regiao: RegiaoPedido = indice % 2 === 0 ? 'BRASIL' : 'PIRAUBE'
  const plantas = PLANTAS_POR_REGIAO[regiao]

  return {
    id: 2200 + indice,
    regiao,
    planta: plantas[indice % plantas.length].codigo,
    pedido: String(310_000 + indice * 7),
    codCliente: cliente.codigo,
    nomeCliente: cliente.nome,
    codigoProduto: produto.codigo,
    referencia: produto.referencia,
    quantidade: 500 + ((indice * 271) % 12_000),
    preco: produto.valorUnitario,
    dataEmissao: iso(indice + 1),
    dataEntrega: iso(-(3 + (indice % 25))),
    emitidoPor: ['jsilva', 'cpereira', 'rnunes'][indice % 3],
    excluido: false
  }
})

export async function emitirPedido(
  dados: DadosEmissaoPedido,
  regiao: RegiaoPedido,
  usuario: string
): Promise<ResultadoPedido & { pedido?: string }> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.codCliente) return atraso({ sucesso: false, mensagem: 'Cliente é obrigatório.' }, 200)
  if (!dados.tes) return atraso({ sucesso: false, mensagem: 'TES é obrigatória.' }, 200)
  if (!dados.condicaoPagamento) {
    return atraso({ sucesso: false, mensagem: 'Condição de pagamento é obrigatória.' }, 200)
  }
  if (dados.itens.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Inclua ao menos um item no pedido.' }, 200)
  }

  for (const [indice, item] of dados.itens.entries()) {
    if (!item.codigoProduto) {
      return atraso({ sucesso: false, mensagem: `Item ${indice + 1}: produto é obrigatório.` }, 200)
    }
    if (!Number(item.quantidade)) {
      return atraso({ sucesso: false, mensagem: `Item ${indice + 1}: quantidade inválida.` }, 200)
    }
    if (!item.dataEntrega) {
      return atraso({ sucesso: false, mensagem: `Item ${indice + 1}: data de entrega é obrigatória.` }, 200)
    }
  }

  const avisos = await validarItensPedido(dados)
  const bloqueio = avisos.find((a) => a.bloqueia)
  if (bloqueio) return atraso({ sucesso: false, mensagem: bloqueio.mensagem }, 200)

  const numero = String(Math.max(...pedidosDigitados.map((p) => Number(p.pedido))) + 7)
  const hoje = new Date().toISOString().slice(0, 10)

  for (const item of dados.itens) {
    pedidosDigitados.push({
      id: Math.max(0, ...pedidosDigitados.map((p) => p.id)) + 1,
      regiao,
      planta: dados.planta,
      pedido: numero,
      codCliente: dados.codCliente,
      nomeCliente: dados.nomeCliente,
      codigoProduto: item.codigoProduto,
      referencia: item.referencia,
      quantidade: Number(item.quantidade),
      preco: Number(item.valorUnitario || 0),
      dataEmissao: hoje,
      dataEntrega: item.dataEntrega,
      emitidoPor: usuario,
      excluido: false
    })
  }

  return atraso(
    { sucesso: true, mensagem: `Pedido ${numero} emitido com ${dados.itens.length} item(ns).`, pedido: numero },
    800
  )
}

export interface FiltrosDigitados {
  planta: string
  pedido: string
  cliente: string
  dataInicio: string
  dataFim: string
}

export const FILTROS_DIGITADOS_VAZIOS: FiltrosDigitados = {
  planta: '',
  pedido: '',
  cliente: '',
  dataInicio: '',
  dataFim: ''
}

export async function listarPedidosDigitados(
  regiao: RegiaoPedido,
  filtros: FiltrosDigitados
): Promise<PedidoDigitado[]> {
  const cliente = filtros.cliente.trim().toLowerCase()

  const resultado = pedidosDigitados
    .filter((p) => !p.excluido && p.regiao === regiao)
    .filter((p) => !filtros.planta || p.planta === filtros.planta)
    .filter((p) => !filtros.pedido.trim() || p.pedido.includes(filtros.pedido.trim()))
    .filter((p) => !cliente || p.nomeCliente.toLowerCase().includes(cliente))
    .filter((p) => !filtros.dataInicio || p.dataEmissao >= filtros.dataInicio)
    .filter((p) => !filtros.dataFim || p.dataEmissao <= filtros.dataFim)
    .map((p) => ({ ...p }))
    .sort((a, b) => b.dataEmissao.localeCompare(a.dataEmissao))

  return atraso(resultado, 400)
}

/** Exclusão de pedido digitado — o legado tem tela dedicada por planta */
export async function excluirPedidosDigitados(ids: number[]): Promise<ResultadoPedido> {
  if (ids.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos um pedido.' }, 200)
  }

  for (const id of ids) {
    const registro = pedidosDigitados.find((p) => p.id === id)
    if (registro) registro.excluido = true
  }

  return atraso({ sucesso: true, mensagem: `${ids.length} pedido(s) excluído(s).` }, 550)
}

/* ------------------------------------------------------------------ *
 * Solicitação de alteração de pedido
 * ------------------------------------------------------------------ */

/**
 * Tipos de alteração.
 *
 * ATENÇÃO: no legado a lista vem da procedure
 * SP_INPANET_HISTORICO_LISTA_OPERACOES_ALTERACOES_PEDIDO, que não está no
 * material. Os ids 2 (Referência) e 3 (Pedido do Cliente) estão
 * confirmados no HTML, e o 7 é o único que checa programação de produção.
 * Os demais rótulos vieram dos campos do formulário e precisam ser
 * conferidos contra a procedure.
 */
export const TIPOS_ALTERACAO_PEDIDO = [
  { id: '1', descricao: 'Observação' },
  { id: '2', descricao: 'Referência' },
  { id: '3', descricao: 'Pedido do Cliente' },
  { id: '4', descricao: 'Condição de Pagamento' },
  { id: '5', descricao: 'Quantidade' },
  { id: '6', descricao: 'Tipo de Frete' },
  { id: '7', descricao: 'Data de Entrega' }
]

/**
 * `ehObrigatorio` do repositório: define se a evidência é exigida.
 * Usuário nível 1 nunca precisa anexar; o tipo 7 só exige quando o pedido
 * já está produzido ou programado.
 */
export function evidenciaObrigatoria(
  tipoAlteracao: string,
  nivelUsuario: number,
  pedidoProgramado: boolean
) {
  if (nivelUsuario === 1) return false
  if (tipoAlteracao === '1' || tipoAlteracao === '4') return false
  if (tipoAlteracao === '7') return pedidoProgramado
  return true
}

export interface SolicitacaoAlteracaoPedido {
  id: number
  regiao: RegiaoPedido
  planta: string
  pedido: string
  codCliente: string
  nomeCliente: string
  codigoProduto: string
  referencia: string
  tipoAlteracao: string
  valorAtual: string
  valorNovo: string
  novaDataDe: string
  novaDataAte: string
  observacao: string
  evidencias: string[]
  status: string
  dataInclusao: string
  solicitante: string
}

export const STATUS_ALTERACAO_PEDIDO = [
  { codigo: 'PENDENTE', descricao: 'Pendente' },
  { codigo: 'APROVADA', descricao: 'Aprovada' },
  { codigo: 'REPROVADA', descricao: 'Reprovada' }
]

const solicitacoesAlteracao: SolicitacaoAlteracaoPedido[] = [
  {
    id: 9200,
    regiao: 'BRASIL',
    planta: '05',
    pedido: '310007',
    codCliente: '003891',
    nomeCliente: 'FRIGORIFICO SERRA AZUL LTDA',
    codigoProduto: '901101',
    referencia: 'CAIXA MAÇÃ 18KG',
    tipoAlteracao: '7',
    valorAtual: '',
    valorNovo: '',
    novaDataDe: iso(-10),
    novaDataAte: iso(-8),
    observacao: 'Cliente pediu antecipação por conta de campanha promocional.',
    evidencias: ['email-cliente.pdf'],
    status: 'PENDENTE',
    dataInclusao: iso(2),
    solicitante: 'cpereira'
  },
  {
    id: 9201,
    regiao: 'PIRAUBE',
    planta: '00',
    pedido: '310014',
    codCliente: '001234',
    nomeCliente: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.',
    codigoProduto: '903201',
    referencia: 'CHAPA ONDULADA 1200X800',
    tipoAlteracao: '5',
    valorAtual: '8000',
    valorNovo: '12000',
    novaDataDe: '',
    novaDataAte: '',
    observacao: 'Aumento de volume acordado com o comercial.',
    evidencias: [],
    status: 'APROVADA',
    dataInclusao: iso(9),
    solicitante: 'jsilva'
  }
]

export interface DadosAlteracaoPedido {
  regiao: RegiaoPedido
  planta: string
  pedido: string
  tipoAlteracao: string
  valorNovo: string
  novaDataDe: string
  novaDataAte: string
  observacao: string
  evidencias: string[]
}

export const ALTERACAO_PEDIDO_VAZIA: Omit<DadosAlteracaoPedido, 'regiao'> = {
  planta: '',
  pedido: '',
  tipoAlteracao: '',
  valorNovo: '',
  novaDataDe: '',
  novaDataAte: '',
  observacao: '',
  evidencias: []
}

/** obterPedido — traz os dados do pedido para o formulário de alteração */
export async function obterPedidoParaAlteracao(
  regiao: RegiaoPedido,
  planta: string,
  pedido: string
): Promise<(PedidoDigitado & { programado: boolean }) | null> {
  const encontrado = pedidosDigitados.find(
    (p) => !p.excluido && p.regiao === regiao && p.planta === planta && p.pedido === pedido.trim()
  )
  if (!encontrado) return atraso(null, 400)

  // fq_PedidoProduzidoOuProgramado — aqui derivado da proximidade da entrega
  const programado = encontrado.dataEntrega <= iso(-7)

  return atraso({ ...encontrado, programado }, 400)
}

export async function listarSolicitacoesAlteracao(
  regiao: RegiaoPedido,
  planta = ''
): Promise<SolicitacaoAlteracaoPedido[]> {
  const resultado = solicitacoesAlteracao
    .filter((s) => s.regiao === regiao)
    .filter((s) => !planta || s.planta === planta)
    .map((s) => ({ ...s }))
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 400)
}

export async function salvarSolicitacaoAlteracao(
  dados: DadosAlteracaoPedido,
  nivelUsuario: number,
  pedidoProgramado: boolean,
  usuario: string
): Promise<ResultadoPedido> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.pedido.trim()) return atraso({ sucesso: false, mensagem: 'Pedido é obrigatório.' }, 200)
  if (!dados.tipoAlteracao) {
    return atraso({ sucesso: false, mensagem: 'Tipo de alteração é obrigatório.' }, 200)
  }
  if (!dados.observacao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Observação é obrigatória.' }, 200)
  }

  if (
    evidenciaObrigatoria(dados.tipoAlteracao, nivelUsuario, pedidoProgramado) &&
    dados.evidencias.length === 0
  ) {
    return atraso(
      {
        sucesso: false,
        mensagem:
          dados.tipoAlteracao === '7'
            ? 'Pedido já programado: anexe a evidência da alteração.'
            : 'Este tipo de alteração exige evidência anexada.'
      },
      200
    )
  }

  const pedido = pedidosDigitados.find(
    (p) => p.regiao === dados.regiao && p.planta === dados.planta && p.pedido === dados.pedido.trim()
  )

  solicitacoesAlteracao.push({
    id: Math.max(0, ...solicitacoesAlteracao.map((s) => s.id)) + 1,
    regiao: dados.regiao,
    planta: dados.planta,
    pedido: dados.pedido.trim(),
    codCliente: pedido?.codCliente ?? '',
    nomeCliente: pedido?.nomeCliente ?? '',
    codigoProduto: pedido?.codigoProduto ?? '',
    referencia: pedido?.referencia ?? '',
    tipoAlteracao: dados.tipoAlteracao,
    valorAtual: pedido ? String(pedido.quantidade) : '',
    valorNovo: dados.valorNovo.trim(),
    novaDataDe: dados.novaDataDe,
    novaDataAte: dados.novaDataAte,
    observacao: dados.observacao.trim(),
    evidencias: [...dados.evidencias],
    status: 'PENDENTE',
    dataInclusao: new Date().toISOString().slice(0, 10),
    solicitante: usuario
  })

  return atraso({ sucesso: true, mensagem: 'Solicitação de alteração registrada!' }, 650)
}

export async function excluirSolicitacaoAlteracao(id: number): Promise<ResultadoPedido> {
  const indice = solicitacoesAlteracao.findIndex((s) => s.id === id)
  if (indice < 0) return atraso({ sucesso: false, mensagem: 'Solicitação não encontrada.' }, 200)
  if (solicitacoesAlteracao[indice].status !== 'PENDENTE') {
    return atraso({ sucesso: false, mensagem: 'Só é possível excluir solicitações pendentes.' }, 200)
  }

  solicitacoesAlteracao.splice(indice, 1)
  return atraso({ sucesso: true, mensagem: 'Solicitação excluída.' }, 450)
}

/* ------------------------------------------------------------------ *
 * Orçamento Brasil
 * ------------------------------------------------------------------ */

export const TIPOS_PRECO = ['Por milheiro', 'Por m²', 'Por kg', 'Unitário']
export const TIPOS_CAIXA = ['FEFCO 0201', 'FEFCO 0203', 'FEFCO 0300', 'FEFCO 0409']

export interface ItemOrcamento {
  tipoItem: string
  tipoCaixa: string
  codigoProduto: string
  estrutura: string
  referencia: string
  quantidade: string
  tipoPreco: string
  valorUnitario: string
  gramatura: string
  comprimentoInterno: string
  larguraInterna: string
  alturaInterna: string
  comprimentoChapa: string
  larguraChapa: string
  arranjo: string
}

export const ITEM_ORCAMENTO_VAZIO: ItemOrcamento = {
  tipoItem: '',
  tipoCaixa: '',
  codigoProduto: '',
  estrutura: '',
  referencia: '',
  quantidade: '',
  tipoPreco: '',
  valorUnitario: '',
  gramatura: '',
  comprimentoInterno: '',
  larguraInterna: '',
  alturaInterna: '',
  comprimentoChapa: '',
  larguraChapa: '',
  arranjo: ''
}

export interface Orcamento {
  id: number
  planta: string
  dataInclusao: string
  dataValidadeDe: string
  dataValidadeAte: string
  status: string
  clienteNovo: boolean
  codCliente: string
  codLoja: string
  nomeCliente: string
  cnpj: string
  codVendedor: string
  nomeVendedor: string
  telefone: string
  contato: string
  email: string
  cep: string
  endereco: string
  numero: string
  complemento: string
  bairro: string
  uf: string
  municipio: string
  tipoFrete: string
  tipoCaminhao: string
  condicoesPagamento: string
  icms: string
  aliquotaIcms: string
  pis: string
  cofins: string
  itens: ItemOrcamento[]
}

export const STATUS_ORCAMENTO = [
  { codigo: 'ATIVO', descricao: 'Ativo' },
  { codigo: 'DESATIVADO', descricao: 'Desativado' },
  { codigo: 'VENCIDO', descricao: 'Vencido' }
]

export const ORCAMENTO_VAZIO: Omit<Orcamento, 'id' | 'dataInclusao' | 'status'> = {
  planta: '',
  dataValidadeDe: '',
  dataValidadeAte: '',
  clienteNovo: false,
  codCliente: '',
  codLoja: '',
  nomeCliente: '',
  cnpj: '',
  codVendedor: '',
  nomeVendedor: '',
  telefone: '',
  contato: '',
  email: '',
  cep: '',
  endereco: '',
  numero: '',
  complemento: '',
  bairro: '',
  uf: '',
  municipio: '',
  tipoFrete: '',
  tipoCaminhao: '',
  condicoesPagamento: '',
  icms: '',
  aliquotaIcms: '',
  pis: '',
  cofins: '',
  itens: []
}

const orcamentos: Orcamento[] = Array.from({ length: 14 }, (_, indice) => {
  const cliente = clientes[indice % clientes.length]
  const vencido = indice % 5 === 0

  return {
    ...ORCAMENTO_VAZIO,
    id: 6600 + indice,
    planta: PLANTAS_POR_REGIAO.BRASIL[indice % 3].codigo,
    dataInclusao: iso(indice + 5),
    dataValidadeDe: iso(indice + 5),
    dataValidadeAte: vencido ? iso(indice) : iso(-(30 - (indice % 20))),
    status: vencido ? 'VENCIDO' : indice % 7 === 0 ? 'DESATIVADO' : 'ATIVO',
    codCliente: cliente.codigo,
    codLoja: cliente.loja,
    nomeCliente: cliente.nome,
    cnpj: cliente.cnpj,
    municipio: cliente.municipio,
    uf: cliente.uf,
    codVendedor: ['V001', 'V002'][indice % 2],
    nomeVendedor: ['CARLOS PEREIRA', 'RICARDO NUNES'][indice % 2],
    condicoesPagamento: cliente.condicaoPagamento,
    tipoFrete: TIPOS_FRETE_PEDIDO[indice % 2],
    itens: [
      {
        ...ITEM_ORCAMENTO_VAZIO,
        tipoItem: 'CX',
        codigoProduto: produtos[indice % produtos.length].codigo,
        referencia: produtos[indice % produtos.length].referencia,
        estrutura: 'BC',
        quantidade: String(5_000 + indice * 500),
        tipoPreco: 'Por milheiro',
        valorUnitario: String(produtos[indice % produtos.length].valorUnitario)
      }
    ]
  }
})

export interface FiltrosOrcamento {
  planta: string
  cliente: string
  status: string
}

export const FILTROS_ORCAMENTO_VAZIOS: FiltrosOrcamento = { planta: '', cliente: '', status: '' }

export async function listarOrcamentos(filtros: FiltrosOrcamento): Promise<Orcamento[]> {
  const cliente = filtros.cliente.trim().toLowerCase()
  const hoje = new Date().toISOString().slice(0, 10)

  const resultado = orcamentos
    .map((o) => ({
      ...o,
      // Validade vencida sobrepõe o status gravado
      status: o.status === 'ATIVO' && o.dataValidadeAte < hoje ? 'VENCIDO' : o.status
    }))
    .filter((o) => !filtros.planta || o.planta === filtros.planta)
    .filter((o) => !cliente || o.nomeCliente.toLowerCase().includes(cliente))
    .filter((o) => !filtros.status || o.status === filtros.status)
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 450)
}

export async function salvarOrcamento(
  dados: Partial<Orcamento> & { id?: number }
): Promise<ResultadoPedido> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.nomeCliente?.trim()) {
    return atraso({ sucesso: false, mensagem: 'Nome do cliente é obrigatório.' }, 200)
  }
  if (!dados.dataValidadeDe || !dados.dataValidadeAte) {
    return atraso({ sucesso: false, mensagem: 'Informe o período de validade.' }, 200)
  }
  if (dados.dataValidadeAte < dados.dataValidadeDe) {
    return atraso({ sucesso: false, mensagem: 'A validade final não pode ser anterior à inicial.' }, 200)
  }
  if (!dados.itens || dados.itens.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Inclua ao menos um item no orçamento.' }, 200)
  }

  if (dados.id) {
    const registro = orcamentos.find((o) => o.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Orçamento não encontrado.' }, 200)
    Object.assign(registro, dados)
    return atraso({ sucesso: true, mensagem: 'Orçamento atualizado com sucesso!' }, 600)
  }

  orcamentos.push({
    ...ORCAMENTO_VAZIO,
    ...dados,
    id: Math.max(0, ...orcamentos.map((o) => o.id)) + 1,
    dataInclusao: new Date().toISOString().slice(0, 10),
    status: 'ATIVO'
  } as Orcamento)

  return atraso({ sucesso: true, mensagem: 'Orçamento criado com sucesso!' }, 650)
}

export async function desativarOrcamento(id: number, motivo: string): Promise<ResultadoPedido> {
  if (!motivo.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o motivo da desativação.' }, 200)
  }

  const registro = orcamentos.find((o) => o.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Orçamento não encontrado.' }, 200)

  registro.status = 'DESATIVADO'
  return atraso({ sucesso: true, mensagem: 'Orçamento desativado.' }, 500)
}

export interface ResultadoCalculo {
  item: string
  custoMaterial: number
  custoConversao: number
  precoSugerido: number
  margem: number
}

/** OrcamentoCalculoTrimboxRepositorio — cálculo de custo vindo do TRIMBOX */
export async function calcularOrcamento(orcamento: Orcamento): Promise<ResultadoCalculo[]> {
  const resultado = orcamento.itens.map((item) => {
    const quantidade = Number(item.quantidade || 1)
    const gramatura = Number(item.gramatura || 150)
    const custoMaterial = Number(((quantidade * gramatura) / 1000).toFixed(2))
    const custoConversao = Number((custoMaterial * 0.35).toFixed(2))
    const valorUnitario = Number(item.valorUnitario || 0)
    const receita = valorUnitario * quantidade
    const custoTotal = custoMaterial + custoConversao

    return {
      item: item.referencia || item.codigoProduto,
      custoMaterial,
      custoConversao,
      precoSugerido: Number((custoTotal * 1.28).toFixed(4)),
      margem: receita > 0 ? Number((((receita - custoTotal) / receita) * 100).toFixed(2)) : 0
    }
  })

  return atraso(resultado, 800)
}

/* ------------------------------------------------------------------ *
 * Parâmetros de pedido
 * ------------------------------------------------------------------ */

/** Faixas mín./máx. por tipo de produto, usadas na crítica da emissão */
export interface ParametroPedido {
  planta: string
  tipoProduto: string
  minimo: string
  maximo: string
}

export const TIPOS_PRODUTO_PARAMETRO = ['Papel', 'Caixa', 'Chapa', 'Corte e vinco', 'Maleta']

const parametros: ParametroPedido[] = PLANTAS_POR_REGIAO.PIRAUBE.concat(
  PLANTAS_POR_REGIAO.BRASIL
).flatMap((planta) =>
  TIPOS_PRODUTO_PARAMETRO.map((tipo, indice) => ({
    planta: planta.codigo,
    tipoProduto: tipo,
    minimo: String(100 + indice * 50),
    maximo: String(20_000 + indice * 5_000)
  }))
)

export async function listarParametros(planta: string): Promise<ParametroPedido[]> {
  return atraso(
    parametros.filter((p) => p.planta === planta).map((p) => ({ ...p })),
    350
  )
}

export async function salvarParametros(
  planta: string,
  novos: ParametroPedido[]
): Promise<ResultadoPedido> {
  for (const item of novos) {
    const minimo = Number(item.minimo)
    const maximo = Number(item.maximo)

    if (Number.isNaN(minimo) || Number.isNaN(maximo)) {
      return atraso({ sucesso: false, mensagem: `${item.tipoProduto}: informe valores numéricos.` }, 200)
    }
    if (minimo > maximo) {
      return atraso(
        { sucesso: false, mensagem: `${item.tipoProduto}: o mínimo não pode ser maior que o máximo.` },
        200
      )
    }
  }

  for (const item of novos) {
    const registro = parametros.find(
      (p) => p.planta === planta && p.tipoProduto === item.tipoProduto
    )
    if (registro) {
      registro.minimo = item.minimo
      registro.maximo = item.maximo
    }
  }

  return atraso({ sucesso: true, mensagem: 'Parâmetros salvos com sucesso!' }, 600)
}

/* ------------------------------------------------------------------ *
 * Entrega no fim de semana
 * ------------------------------------------------------------------ */

export interface ClienteEntregaFds {
  codigo: string
  nome: string
  ativo: boolean
}

const entregaFds: ClienteEntregaFds[] = [
  { codigo: '001234', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', ativo: true },
  { codigo: '003891', nome: 'FRIGORIFICO SERRA AZUL LTDA', ativo: true },
  { codigo: '006432', nome: 'DISTRIBUIDORA LITORAL SUL LTDA', ativo: false }
]

export async function listarEntregaFds(): Promise<ClienteEntregaFds[]> {
  return atraso(entregaFds.map((c) => ({ ...c })), 300)
}

/** validarCliente + criar: o cliente precisa existir na base */
export async function salvarEntregaFds(codigo: string, ativo: boolean): Promise<ResultadoPedido> {
  const alvo = codigo.trim()
  if (!alvo) return atraso({ sucesso: false, mensagem: 'Informe o código do cliente.' }, 200)

  const cliente = clientes.find((c) => c.codigo === alvo)
  if (!cliente) return atraso({ sucesso: false, mensagem: 'Cliente não encontrado na base.' }, 200)

  const existente = entregaFds.find((c) => c.codigo === alvo)
  if (existente) {
    existente.ativo = ativo
    return atraso({ sucesso: true, mensagem: 'Cliente atualizado com sucesso!' }, 450)
  }

  entregaFds.push({ codigo: alvo, nome: cliente.nome, ativo })
  return atraso({ sucesso: true, mensagem: 'Cliente liberado para entrega no fim de semana.' }, 500)
}

export async function excluirEntregaFds(codigo: string): Promise<ResultadoPedido> {
  const indice = entregaFds.findIndex((c) => c.codigo === codigo)
  if (indice < 0) return atraso({ sucesso: false, mensagem: 'Cliente não encontrado.' }, 200)

  entregaFds.splice(indice, 1)
  return atraso({ sucesso: true, mensagem: 'Cliente removido da lista.' }, 400)
}

/* ------------------------------------------------------------------ *
 * Observação em NF
 * ------------------------------------------------------------------ */

const observacoesNf = new Map<string, string>([['310007', 'Entregar somente no turno da manhã.']])

export async function buscarObservacaoNf(
  pedido: string
): Promise<{ encontrado: boolean; nomeCliente: string; observacao: string }> {
  const alvo = pedido.trim()
  const registro = pedidosDigitados.find((p) => !p.excluido && p.pedido === alvo)

  if (!registro) return atraso({ encontrado: false, nomeCliente: '', observacao: '' }, 400)

  return atraso(
    {
      encontrado: true,
      nomeCliente: registro.nomeCliente,
      observacao: observacoesNf.get(alvo) ?? ''
    },
    400
  )
}

export async function salvarObservacaoNf(pedido: string, observacao: string): Promise<ResultadoPedido> {
  const alvo = pedido.trim()
  if (!alvo) return atraso({ sucesso: false, mensagem: 'Informe o número do pedido.' }, 200)

  const existe = pedidosDigitados.some((p) => !p.excluido && p.pedido === alvo)
  if (!existe) return atraso({ sucesso: false, mensagem: 'Pedido não encontrado.' }, 200)

  observacoesNf.set(alvo, observacao.trim())
  return atraso({ sucesso: true, mensagem: 'Observação gravada na nota fiscal.' }, 500)
}
