import type {
  DadosFornecedorAparas,
  DadosRecebimento,
  FiltrosRecebimento,
  FiltrosRelatorioRecebimento,
  FornecedorAparas,
  RecebimentoAparas
} from '@/pages/cavacoAparas/tipos'
import type { ResultadoCrud } from '@/pages/cadastrosFnc/tipos'
import { atraso, nomeEmpresa } from '@/services/comuns/baseMock'

/**
 * Cavaco-Aparas — RECEBIMENTO_FORNECEDORES e CONTROLE_RECEBIMENTO.
 *
 * Regras herdadas do legado mantidas aqui:
 * - fornecedor é excluído por soft delete (coluna EXCLUIDO);
 * - recebimento é excluído gravando DATA_EXCLUSAO;
 * - CNPJ não pode repetir no cadastro de fornecedor;
 * - a NF é verificada antes de gravar (VerificarNF), para não duplicar
 *   nota do mesmo fornecedor.
 */

interface FornecedorInterno extends FornecedorAparas {
  excluido: boolean
}

interface RecebimentoInterno extends Omit<RecebimentoAparas, 'nomePlanta'> {
  excluido: boolean
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const fornecedores: FornecedorInterno[] = [
  { codigo: '100234', cnpj: '12345678000190', nome: 'APARAS SÃO JOÃO LTDA', excluido: false },
  { codigo: '100235', cnpj: '98765432000110', nome: 'RECICLAGEM MINAS EIRELI', excluido: false },
  { codigo: '100236', cnpj: '45678912000155', nome: 'COMÉRCIO DE PAPÉIS BENTO LTDA', excluido: false },
  { codigo: '100237', cnpj: '32165498000177', nome: 'ECO APARAS DO SUL LTDA', excluido: false },
  { codigo: '100238', cnpj: '78912345000122', nome: 'CAVACO E MADEIRAS UBERABA LTDA', excluido: false },
  { codigo: '100239', cnpj: '65432198000133', nome: 'FORNECEDOR DESATIVADO LTDA', excluido: true }
]

const PRODUTOS = ['APARA BRANCA', 'APARA MISTA', 'APARA KRAFT', 'CAVACO PINUS', 'CAVACO EUCALIPTO']
const LOCAIS = ['PÁTIO 1', 'PÁTIO 2', 'GALPÃO A', 'GALPÃO B']

/** Gera uma massa determinística de recebimentos para as grades e o relatório */
function gerarRecebimentos(): RecebimentoInterno[] {
  const plantas = ['00', '04', '05', '12']
  const ativos = fornecedores.filter((f) => !f.excluido)

  return Array.from({ length: 48 }, (_, indice) => {
    const fornecedor = ativos[indice % ativos.length]
    const pesoNf = 12_000 + ((indice * 733) % 18_000)
    const temBalanca = indice % 7 !== 0

    return {
      id: 9000 + indice,
      planta: plantas[indice % plantas.length],
      notaFiscal: String(45_000 + indice * 3),
      serie: indice % 3 === 0 ? '1' : '2',
      dataNf: iso(indice + 2),
      dataTeste: iso(indice + 1),
      horaTeste: `${String(7 + (indice % 11)).padStart(2, '0')}:${String((indice * 13) % 60).padStart(2, '0')}`,
      codigoFornecedor: fornecedor.codigo,
      cnpj: fornecedor.cnpj,
      nomeFornecedor: fornecedor.nome,
      produto: PRODUTOS[indice % PRODUTOS.length],
      pesoNf,
      pesoBalanca: temBalanca ? pesoNf - ((indice * 17) % 400) : null,
      umidade: Number((8 + ((indice * 7) % 45) / 10).toFixed(1)),
      impureza: Number((((indice * 3) % 30) / 10).toFixed(1)),
      proibitivo: Number((((indice * 5) % 15) / 10).toFixed(1)),
      outros: Number((((indice * 11) % 20) / 10).toFixed(1)),
      local: LOCAIS[indice % LOCAIS.length],
      quantidade: 10 + (indice % 25),
      fardosAnalisados: 2 + (indice % 6),
      placa: `ABC${String(1000 + indice).slice(-4)}`,
      motorista: ['CARLOS SOUZA', 'MARCOS DIAS', 'PAULO REIS', 'ANTONIO LUZ'][indice % 4],
      observacao: indice % 5 === 0 ? 'Carga com material fora de especificação.' : '',
      incluidoPor: ['admin', 'jsilva', 'bcarvalho'][indice % 3],
      excluido: false
    }
  })
}

const recebimentos: RecebimentoInterno[] = gerarRecebimentos()

function proximoIdRecebimento() {
  return Math.max(0, ...recebimentos.map((r) => r.id)) + 1
}

/** Só dígitos, como o legado grava o CNPJ */
function limparCnpj(cnpj: string) {
  return cnpj.replace(/\D/g, '')
}

export function formatarCnpj(cnpj: string) {
  const digitos = limparCnpj(cnpj)
  if (digitos.length !== 14) return cnpj
  return digitos.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
}

/* ------------------------------------------------------------------ *
 * Fornecedores
 * ------------------------------------------------------------------ */

export async function listarFornecedores(termo = ''): Promise<FornecedorAparas[]> {
  const busca = termo.trim().toLowerCase()

  const ativos = fornecedores
    .filter((f) => !f.excluido)
    .filter(
      (f) =>
        !busca ||
        f.nome.toLowerCase().includes(busca) ||
        f.cnpj.includes(limparCnpj(busca)) ||
        f.codigo.includes(busca)
    )
    .map(({ excluido, ...resto }) => resto)

  return atraso(ativos, 300)
}

export async function salvarFornecedor(
  dados: DadosFornecedorAparas,
  editando: boolean
): Promise<ResultadoCrud> {
  const codigo = dados.codigo.trim()
  const cnpj = limparCnpj(dados.cnpj)
  const nome = dados.nome.trim().toUpperCase()

  if (!codigo) return atraso({ sucesso: false, mensagem: 'Código SAP é obrigatório.' }, 150)
  if (cnpj.length !== 14) {
    return atraso({ sucesso: false, mensagem: 'CNPJ deve ter 14 caracteres.' }, 150)
  }
  if (!nome) return atraso({ sucesso: false, mensagem: 'Razão social é obrigatória.' }, 150)

  const existentePorCodigo = fornecedores.find((f) => f.codigo === codigo && !f.excluido)

  if (editando) {
    if (!existentePorCodigo) {
      return atraso({ sucesso: false, mensagem: 'Fornecedor não encontrado.' }, 150)
    }
    const cnpjDuplicado = fornecedores.some(
      (f) => !f.excluido && f.cnpj === cnpj && f.codigo !== codigo
    )
    if (cnpjDuplicado) {
      return atraso({ sucesso: false, mensagem: 'Já existe outro fornecedor com este CNPJ.' }, 150)
    }

    existentePorCodigo.cnpj = cnpj
    existentePorCodigo.nome = nome
    return atraso({ sucesso: true, mensagem: 'Cadastro atualizado com sucesso!' }, 400)
  }

  if (existentePorCodigo) {
    return atraso({ sucesso: false, mensagem: 'Já existe um fornecedor com este código SAP.' }, 150)
  }
  if (fornecedores.some((f) => !f.excluido && f.cnpj === cnpj)) {
    return atraso({ sucesso: false, mensagem: 'Já existe um fornecedor com este CNPJ.' }, 150)
  }

  fornecedores.push({ codigo, cnpj, nome, excluido: false })
  return atraso({ sucesso: true, mensagem: 'Cadastro realizado com sucesso!' }, 400)
}

export async function excluirFornecedor(codigo: string): Promise<ResultadoCrud> {
  const fornecedor = fornecedores.find((f) => f.codigo === codigo)
  if (!fornecedor) return atraso({ sucesso: false, mensagem: 'Fornecedor não encontrado.' }, 150)

  fornecedor.excluido = true
  return atraso({ sucesso: true, mensagem: 'Registro apagado com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Controle de recebimento
 * ------------------------------------------------------------------ */

function paraLinha(registro: RecebimentoInterno): RecebimentoAparas {
  const { excluido, ...resto } = registro
  return { ...resto, nomePlanta: nomeEmpresa(resto.planta) }
}

export async function listarRecebimentos(filtros: FiltrosRecebimento): Promise<RecebimentoAparas[]> {
  const nf = filtros.notaFiscal.trim()
  const fornecedor = filtros.fornecedor.trim().toLowerCase()
  const produto = filtros.produto.trim().toLowerCase()

  const resultado = recebimentos
    .filter((r) => !r.excluido)
    .filter((r) => !nf || r.notaFiscal.includes(nf))
    .filter(
      (r) =>
        !fornecedor ||
        r.nomeFornecedor.toLowerCase().includes(fornecedor) ||
        r.cnpj.includes(limparCnpj(fornecedor))
    )
    .filter((r) => !produto || r.produto.toLowerCase().includes(produto))
    .filter((r) => !filtros.dataInicio || r.dataTeste >= filtros.dataInicio)
    .filter((r) => !filtros.dataFim || r.dataTeste <= filtros.dataFim)
    .sort((a, b) => b.dataTeste.localeCompare(a.dataTeste))
    .map(paraLinha)

  return atraso(resultado, 400)
}

/** VerificarNF do legado: a mesma NF não pode repetir para o mesmo fornecedor */
export function notaFiscalDuplicada(notaFiscal: string, codigoFornecedor: string, idAtual?: number) {
  return recebimentos.some(
    (r) =>
      !r.excluido &&
      r.notaFiscal === notaFiscal.trim() &&
      r.codigoFornecedor === codigoFornecedor &&
      r.id !== idAtual
  )
}

export async function salvarRecebimento(
  dados: DadosRecebimento,
  planta: string,
  usuario: string
): Promise<ResultadoCrud> {
  const obrigatorios: [keyof DadosRecebimento, string][] = [
    ['notaFiscal', 'Nota Fiscal'],
    ['serie', 'Série'],
    ['dataNf', 'Data da N.F'],
    ['dataTeste', 'Data do Teste'],
    ['horaTeste', 'Hora do Teste'],
    ['codigoFornecedor', 'Fornecedor'],
    ['produto', 'Produto'],
    ['pesoNf', 'Peso N.F'],
    ['umidade', '% Umidade'],
    ['motorista', 'Nome do Motorista'],
    ['placa', 'Placa']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 150)
    }
  }

  if (notaFiscalDuplicada(dados.notaFiscal, dados.codigoFornecedor, dados.id)) {
    return atraso(
      { sucesso: false, mensagem: 'Já existe um recebimento com esta N.F para o fornecedor informado.' },
      150
    )
  }

  const numero = (valor: string) => (valor.trim() === '' ? 0 : Number(valor))

  const campos = {
    notaFiscal: dados.notaFiscal.trim(),
    serie: dados.serie.trim(),
    dataNf: dados.dataNf,
    dataTeste: dados.dataTeste,
    horaTeste: dados.horaTeste,
    codigoFornecedor: dados.codigoFornecedor,
    cnpj: dados.cnpj,
    nomeFornecedor: dados.nomeFornecedor,
    produto: dados.produto.trim(),
    pesoNf: numero(dados.pesoNf),
    pesoBalanca: dados.pesoBalanca.trim() === '' ? null : numero(dados.pesoBalanca),
    umidade: numero(dados.umidade),
    impureza: numero(dados.impureza),
    proibitivo: numero(dados.proibitivo),
    outros: numero(dados.outros),
    local: dados.local.trim(),
    quantidade: numero(dados.quantidade),
    fardosAnalisados: numero(dados.fardosAnalisados),
    placa: dados.placa.trim().toUpperCase(),
    motorista: dados.motorista.trim().toUpperCase(),
    observacao: dados.observacao.trim()
  }

  if (dados.id) {
    const registro = recebimentos.find((r) => r.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Registro não encontrado.' }, 150)

    Object.assign(registro, campos)
    return atraso({ sucesso: true, mensagem: 'Registro alterado com sucesso!' }, 450)
  }

  recebimentos.push({
    id: proximoIdRecebimento(),
    planta,
    incluidoPor: usuario,
    excluido: false,
    ...campos
  })

  return atraso({ sucesso: true, mensagem: 'Registro incluído com sucesso!' }, 450)
}

/**
 * Atualização só do peso da balança — no legado é o que o usuário sem
 * NivelControladoria consegue fazer (ação "APB").
 */
export async function atualizarPesoBalanca(id: number, pesoBalanca: string): Promise<ResultadoCrud> {
  const registro = recebimentos.find((r) => r.id === id && !r.excluido)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Registro não encontrado.' }, 150)

  registro.pesoBalanca = pesoBalanca.trim() === '' ? null : Number(pesoBalanca)
  return atraso({ sucesso: true, mensagem: 'Peso da balança atualizado com sucesso!' }, 400)
}

export async function excluirRecebimento(id: number): Promise<ResultadoCrud> {
  const registro = recebimentos.find((r) => r.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Registro não encontrado.' }, 150)

  registro.excluido = true
  return atraso({ sucesso: true, mensagem: 'Registro excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Relatório
 * ------------------------------------------------------------------ */

export async function listarRelatorio(
  filtros: FiltrosRelatorioRecebimento
): Promise<RecebimentoAparas[]> {
  const fornecedor = filtros.fornecedor.trim()

  const resultado = recebimentos
    .filter((r) => !r.excluido)
    .filter((r) => filtros.planta === 'TODOS' || r.planta === filtros.planta)
    .filter((r) => !fornecedor || r.codigoFornecedor === fornecedor)
    .filter((r) => !filtros.dataInicio || r.dataTeste >= filtros.dataInicio)
    .filter((r) => !filtros.dataFim || r.dataTeste <= filtros.dataFim)
    .sort((a, b) => b.dataTeste.localeCompare(a.dataTeste))
    .map(paraLinha)

  return atraso(resultado, 500)
}
