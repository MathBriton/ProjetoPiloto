import { atraso } from '@/services/comuns/baseMock'

/**
 * Expedicao — Checklist-Transporte (CHECKLIST_TRANSPORTES,
 * CHECKLIST_LISTAITENS, CHECKLIST_TIPOVEICULO) e Simulador-De-Cargas
 * (VW_QUANTIDADE_CAMINHAO, BI_CLIENTES).
 */

/* ------------------------------------------------------------------ *
 * Checklist de transporte
 * ------------------------------------------------------------------ */

/** CHECKLIST_TRANSPORTES.STATUS — fluxo do legado, nesta ordem */
export const STATUS_CHECKLIST = ['PATIO HJ', 'PORTARIA', 'EXPEDIÇÃO', 'FINALIZADO'] as const
export type StatusChecklist = (typeof STATUS_CHECKLIST)[number]

export interface TipoVeiculo {
  id: number
  descricao: string
}

export const TIPOS_VEICULO: TipoVeiculo[] = [
  { id: 1, descricao: 'Carreta Baú' },
  { id: 2, descricao: 'Carreta Sider' },
  { id: 3, descricao: 'Truck' },
  { id: 4, descricao: 'Toco' },
  { id: 5, descricao: 'Bitrem' },
  { id: 6, descricao: 'Rodotrem' },
  { id: 7, descricao: 'VUC' }
]

export interface ItemChecklist {
  id: number
  descricao: string
}

/** CHECKLIST_LISTAITENS — itens avaliados no veículo */
export const ITENS_CHECKLIST: ItemChecklist[] = [
  { id: 1, descricao: 'Assoalho em bom estado, sem furos ou frestas' },
  { id: 2, descricao: 'Lonas e cintas de amarração em boas condições' },
  { id: 3, descricao: 'Ausência de odores fortes no compartimento de carga' },
  { id: 4, descricao: 'Compartimento limpo e seco' },
  { id: 5, descricao: 'Ausência de pragas ou vestígios de infestação' },
  { id: 6, descricao: 'Teto e laterais sem pontos de infiltração' },
  { id: 7, descricao: 'Portas e travas funcionando corretamente' },
  { id: 8, descricao: 'Pneus em condições adequadas de uso' },
  { id: 9, descricao: 'Extintor de incêndio dentro da validade' },
  { id: 10, descricao: 'Motorista com EPI adequado' },
  { id: 11, descricao: 'Documentação do veículo em dia' },
  { id: 12, descricao: 'Ausência de carga anterior remanescente' }
]

export interface ChecklistTransporte {
  id: number
  nomeProprietario: string
  telefone: string
  nomeMotorista: string
  documento: string
  placaVeiculo: string
  dataPreenchimento: string
  idTipoVeiculo: number
  tipoVeiculo: string
  quantidadeEixo: number
  horarioInicio: string
  horarioTermino: string
  /** Conclusão geral: aprovado / reprovado / aprovado com ressalva */
  conclusao: 'APROVADO' | 'REPROVADO' | 'RESSALVA'
  destinoCarga: string
  status: StatusChecklist
  /** ids dos itens marcados como "Atende" */
  itensAtende: number[]
}

export interface DadosChecklist {
  id?: number
  nomeProprietario: string
  telefone: string
  nomeMotorista: string
  documento: string
  placaVeiculo: string
  dataPreenchimento: string
  idTipoVeiculo: number | ''
  quantidadeEixo: string
  horarioInicio: string
  horarioTermino: string
  conclusao: ChecklistTransporte['conclusao'] | ''
  destinoCarga: string
  itensAtende: number[]
}

export interface ResultadoExpedicao {
  sucesso: boolean
  mensagem: string
}

interface ChecklistInterno extends Omit<ChecklistTransporte, 'tipoVeiculo'> {
  ativo: boolean
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const PROPRIETARIOS = [
  { proprietario: 'TRANSPORTES SANTA CLARA LTDA', motorista: 'JOSE ANTONIO RIBEIRO', doc: '12345678901' },
  { proprietario: 'LOG SUL TRANSPORTES EIRELI', motorista: 'MARCOS VINICIUS SILVA', doc: '23456789012' },
  { proprietario: 'RODOVIARIO MINAS GERAIS S.A.', motorista: 'PAULO CESAR ALMEIDA', doc: '34567890123' },
  { proprietario: 'EXPRESSO NORDESTE LTDA', motorista: 'ANTONIO CARLOS LIMA', doc: '45678901234' },
  { proprietario: 'TRANSCARGA BRASIL LTDA', motorista: 'ROBERTO DIAS FERREIRA', doc: '56789012345' }
]

const DESTINOS = ['SÃO PAULO / SP', 'RIO DE JANEIRO / RJ', 'BELO HORIZONTE / MG', 'CURITIBA / PR', 'RECIFE / PE']

const checklists: ChecklistInterno[] = Array.from({ length: 24 }, (_, indice) => {
  const dono = PROPRIETARIOS[indice % PROPRIETARIOS.length]
  const conclusoes: ChecklistTransporte['conclusao'][] = ['APROVADO', 'APROVADO', 'RESSALVA', 'REPROVADO']

  return {
    id: 3000 + indice,
    nomeProprietario: dono.proprietario,
    telefone: `(32) 9${String(8000 + indice).slice(-4)}-${String(1000 + indice * 3).slice(-4)}`,
    nomeMotorista: dono.motorista,
    documento: dono.doc,
    placaVeiculo: `RST${String(2000 + indice).slice(-4)}`,
    dataPreenchimento: iso(indice),
    idTipoVeiculo: TIPOS_VEICULO[indice % TIPOS_VEICULO.length].id,
    quantidadeEixo: [2, 3, 4, 5, 6, 9][indice % 6],
    horarioInicio: `${String(6 + (indice % 10)).padStart(2, '0')}:00`,
    horarioTermino: `${String(8 + (indice % 10)).padStart(2, '0')}:30`,
    conclusao: conclusoes[indice % conclusoes.length],
    destinoCarga: DESTINOS[indice % DESTINOS.length],
    status: STATUS_CHECKLIST[indice % STATUS_CHECKLIST.length],
    itensAtende: ITENS_CHECKLIST.filter((_, i) => (indice + i) % 4 !== 0).map((item) => item.id),
    ativo: true
  }
})

function paraChecklist(registro: ChecklistInterno): ChecklistTransporte {
  const { ativo, ...resto } = registro
  return {
    ...resto,
    tipoVeiculo: TIPOS_VEICULO.find((t) => t.id === resto.idTipoVeiculo)?.descricao ?? ''
  }
}

export interface FiltrosChecklist {
  /** '' = todos; senão o status exato */
  status: string
  /** campo de busca: NP = proprietário, NM = motorista, PL = placa */
  campo: 'NP' | 'NM' | 'PL' | ''
  valor: string
}

export async function listarChecklists(filtros: FiltrosChecklist): Promise<ChecklistTransporte[]> {
  const valor = filtros.valor.trim().toLowerCase()

  const resultado = checklists
    .filter((c) => c.ativo)
    .filter((c) => !filtros.status || c.status === filtros.status)
    .filter((c) => {
      if (!valor || !filtros.campo) return true
      if (filtros.campo === 'NP') return c.nomeProprietario.toLowerCase().includes(valor)
      if (filtros.campo === 'NM') return c.nomeMotorista.toLowerCase().includes(valor)
      return c.placaVeiculo.toLowerCase().includes(valor)
    })
    .sort((a, b) => b.dataPreenchimento.localeCompare(a.dataPreenchimento) || b.id - a.id)
    .map(paraChecklist)

  return atraso(resultado, 400)
}

/** BuscarPlaca do legado: reaproveita os dados do último checklist da placa */
export async function buscarPorPlaca(placa: string): Promise<ChecklistTransporte | null> {
  const alvo = placa.trim().toUpperCase()
  const encontrado = checklists
    .filter((c) => c.ativo && c.placaVeiculo.toUpperCase() === alvo)
    .sort((a, b) => b.id - a.id)[0]

  return atraso(encontrado ? paraChecklist(encontrado) : null, 350)
}

export async function salvarChecklist(dados: DadosChecklist): Promise<ResultadoExpedicao> {
  const obrigatorios: [keyof DadosChecklist, string][] = [
    ['nomeProprietario', 'Nome do proprietário'],
    ['nomeMotorista', 'Nome do motorista'],
    ['documento', 'CPF/CNPJ'],
    ['placaVeiculo', 'Placa'],
    ['dataPreenchimento', 'Data de preenchimento'],
    ['destinoCarga', 'Destino da carga']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }
  if (!dados.idTipoVeiculo) {
    return atraso({ sucesso: false, mensagem: 'Tipo de veículo é obrigatório.' }, 200)
  }
  if (!dados.conclusao) {
    return atraso({ sucesso: false, mensagem: 'Informe a conclusão do checklist.' }, 200)
  }

  const campos = {
    nomeProprietario: dados.nomeProprietario.trim().toUpperCase(),
    telefone: dados.telefone.trim(),
    nomeMotorista: dados.nomeMotorista.trim().toUpperCase(),
    documento: dados.documento.replace(/\D/g, ''),
    placaVeiculo: dados.placaVeiculo.trim().toUpperCase(),
    dataPreenchimento: dados.dataPreenchimento,
    idTipoVeiculo: Number(dados.idTipoVeiculo),
    quantidadeEixo: Number(dados.quantidadeEixo || 0),
    horarioInicio: dados.horarioInicio,
    horarioTermino: dados.horarioTermino,
    conclusao: dados.conclusao,
    destinoCarga: dados.destinoCarga.trim().toUpperCase(),
    itensAtende: [...dados.itensAtende]
  }

  if (dados.id) {
    const registro = checklists.find((c) => c.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Checklist não encontrado.' }, 200)
    if (registro.status === 'FINALIZADO') {
      return atraso({ sucesso: false, mensagem: 'Checklist finalizado não pode ser alterado.' }, 200)
    }

    Object.assign(registro, campos)
    return atraso({ sucesso: true, mensagem: 'Checklist alterado com sucesso!' }, 500)
  }

  checklists.push({
    id: Math.max(0, ...checklists.map((c) => c.id)) + 1,
    status: 'PATIO HJ',
    ativo: true,
    ...campos
  })

  return atraso({ sucesso: true, mensagem: 'Checklist cadastrado com sucesso!' }, 500)
}

/** MudarStatus — o legado bloqueia mudanças em checklist já finalizado */
export async function mudarStatus(id: number, novoStatus: StatusChecklist): Promise<ResultadoExpedicao> {
  const registro = checklists.find((c) => c.id === id && c.ativo)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Checklist não encontrado.' }, 200)
  if (registro.status === 'FINALIZADO') {
    return atraso({ sucesso: false, mensagem: 'Checklist já finalizado.' }, 200)
  }

  registro.status = novoStatus
  return atraso({ sucesso: true, mensagem: `Status alterado para ${novoStatus}.` }, 350)
}

export async function excluirChecklist(id: number): Promise<ResultadoExpedicao> {
  const registro = checklists.find((c) => c.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Checklist não encontrado.' }, 200)

  registro.ativo = false
  return atraso({ sucesso: true, mensagem: 'Checklist excluído com sucesso!' }, 350)
}

/** CheckListTransporteIndicador.asp — contagem por status */
export async function indicadoresChecklist() {
  const ativos = checklists.filter((c) => c.ativo)
  return atraso(
    STATUS_CHECKLIST.map((status) => ({
      status,
      total: ativos.filter((c) => c.status === status).length
    })),
    250
  )
}

/* ------------------------------------------------------------------ *
 * Simulador de cargas
 * ------------------------------------------------------------------ */

/** Cada planta consulta um DADOSADV diferente no legado */
export const PLANTAS_SIMULADOR = [
  { codigo: 'PIRAPETINGA', nome: 'Pirapetinga' },
  { codigo: 'UBERABA', nome: 'Uberaba' },
  { codigo: 'BENTO', nome: 'Bento Gonçalves' },
  { codigo: 'MARANGUAPE', nome: 'Maranguape' },
  { codigo: 'SAQUAREMA', nome: 'Saquarema' }
]

export interface ClienteSimulador {
  codigo: string
  loja: string
  nome: string
  cnpj: string
  bairro: string
  municipio: string
  uf: string
}

export interface ProdutoSimulador {
  codigo: string
  referencia: string
}

/** VW_QUANTIDADE_CAMINHAO — capacidade por tipo de caminhão */
export interface LinhaSimulacao {
  tipo: string
  totalCaixas: number
  totalPalete: number
  m2Total: number
  totalCaixas2800: number
  totalPalete2800: number
  m2Total2800: number
}

const CLIENTES: ClienteSimulador[] = [
  { codigo: 'C00101', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', cnpj: '11222333000144', bairro: 'DISTRITO INDUSTRIAL', municipio: 'JUIZ DE FORA', uf: 'MG' },
  { codigo: 'C00102', loja: '01', nome: 'BEBIDAS DO VALE LTDA', cnpj: '22333444000155', bairro: 'CENTRO', municipio: 'CATAGUASES', uf: 'MG' },
  { codigo: 'C00103', loja: '02', nome: 'FRIGORIFICO SERRA AZUL LTDA', cnpj: '33444555000166', bairro: 'ZONA RURAL', municipio: 'BENTO GONCALVES', uf: 'RS' },
  { codigo: 'C00104', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.', cnpj: '44555666000177', bairro: 'PARQUE INDUSTRIAL', municipio: 'MARANGUAPE', uf: 'CE' },
  { codigo: 'C00105', loja: '01', nome: 'LATICINIOS CAMPO BELO LTDA', cnpj: '55666777000188', bairro: 'SAO JOSE', municipio: 'UBERABA', uf: 'MG' }
]

const PRODUTOS: ProdutoSimulador[] = [
  { codigo: 'P100234', referencia: 'CAIXA MAÇÃ 18KG' },
  { codigo: 'P100567', referencia: 'CAIXA BEBIDA 12X1L' },
  { codigo: 'P100891', referencia: 'CHAPA ONDULADA 1200X800' },
  { codigo: 'P101245', referencia: 'CAIXA FRIGORIFICADA 20KG' },
  { codigo: 'P101678', referencia: 'BANDEJA HORTIFRUTI 5KG' }
]

export async function buscarClientes(termo: string): Promise<ClienteSimulador[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = busca
    ? CLIENTES.filter(
        (c) => c.nome.toLowerCase().includes(busca) || c.codigo.toLowerCase().includes(busca)
      )
    : CLIENTES
  return atraso(resultado, 350)
}

export async function buscarProdutos(termo: string): Promise<ProdutoSimulador[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = busca
    ? PRODUTOS.filter(
        (p) => p.referencia.toLowerCase().includes(busca) || p.codigo.toLowerCase().includes(busca)
      )
    : PRODUTOS
  return atraso(resultado, 350)
}

/**
 * Consultar: devolve a capacidade por tipo de caminhão para o par
 * cliente + produto na planta escolhida.
 */
export async function simularCarga(
  planta: string,
  codigoCliente: string,
  codigoProduto: string
): Promise<LinhaSimulacao[]> {
  if (!planta || !codigoCliente || !codigoProduto) return atraso([], 200)

  // Semente determinística a partir do produto, para o resultado ser estável
  const semente = [...codigoProduto].reduce((soma, c) => soma + c.charCodeAt(0), 0)

  const tipos = ['TRUCK', 'CARRETA', 'BITREM', 'RODOTREM', 'VUC']
  const fatores = [0.55, 1, 1.6, 2.1, 0.3]

  const resultado = tipos.map((tipo, indice) => {
    const paletes = Math.round((20 + (semente % 9)) * fatores[indice])
    const caixasPorPalete = 40 + (semente % 25)
    const m2PorCaixa = 0.6 + ((semente % 7) / 10)

    return {
      tipo,
      totalPalete: paletes,
      totalCaixas: paletes * caixasPorPalete,
      m2Total: Number((paletes * caixasPorPalete * m2PorCaixa).toFixed(2)),
      // Colunas "2800" = altura de empilhamento alternativa
      totalPalete2800: Math.round(paletes * 1.15),
      totalCaixas2800: Math.round(paletes * 1.15) * caixasPorPalete,
      m2Total2800: Number((Math.round(paletes * 1.15) * caixasPorPalete * m2PorCaixa).toFixed(2))
    }
  })

  return atraso(resultado, 600)
}
