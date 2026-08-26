/**
 * Tipos das telas de Cavaco-Aparas (tabelas RECEBIMENTO_FORNECEDORES e
 * CONTROLE_RECEBIMENTO, banco PIR-CORPORATIVO).
 */

export interface FornecedorAparas {
  /** RECEBIMENTO_FORNECEDORES.CODIGO — código SAP */
  codigo: string
  cnpj: string
  nome: string
}

export interface DadosFornecedorAparas {
  codigo: string
  cnpj: string
  nome: string
}

export interface RecebimentoAparas {
  id: number
  planta: string
  nomePlanta: string
  /** NF + Série */
  notaFiscal: string
  serie: string
  dataNf: string
  dataTeste: string
  horaTeste: string
  codigoFornecedor: string
  cnpj: string
  nomeFornecedor: string
  produto: string
  pesoNf: number
  pesoBalanca: number | null
  umidade: number
  impureza: number
  proibitivo: number
  outros: number
  local: string
  quantidade: number
  fardosAnalisados: number
  placa: string
  motorista: string
  observacao: string
  incluidoPor: string
}

export interface DadosRecebimento {
  id?: number
  notaFiscal: string
  serie: string
  dataNf: string
  dataTeste: string
  horaTeste: string
  codigoFornecedor: string
  cnpj: string
  nomeFornecedor: string
  produto: string
  pesoNf: string
  pesoBalanca: string
  umidade: string
  impureza: string
  proibitivo: string
  outros: string
  local: string
  quantidade: string
  fardosAnalisados: string
  placa: string
  motorista: string
  observacao: string
}

export interface FiltrosRecebimento {
  notaFiscal: string
  fornecedor: string
  produto: string
  dataInicio: string
  dataFim: string
}

export interface FiltrosRelatorioRecebimento {
  planta: string
  fornecedor: string
  dataInicio: string
  dataFim: string
}
