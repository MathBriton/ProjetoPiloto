/**
 * Tipos das telas de Cadastros-Orçamento (tabelas ANALISTAS_ORCAMENTO,
 * APROVADORES_ORCAMENTO e FAIXA_APROVADORES).
 */

export interface AnalistaOrcamento {
  /** CODIGO_USUARIO — é o CRM_R_E_C_N_O do usuário */
  id: number
  nome: string
  email: string
  /** Uma linha por planta no banco; agrupadas aqui como no STUFF/FOR XML do ASP */
  plantas: string[]
  dataInclusao: string
}

/** APROVADORES_ORCAMENTO.NIVEL */
export type NivelAprovador = '1' | '2' | '3'

export interface AprovadorOrcamento {
  /** ID_APROVADOR — é o CRM_R_E_C_N_O do usuário */
  id: number
  nome: string
  email: string
  nivel: NivelAprovador
  plantas: string[]
  dataInclusao: string
}

/** FAIXA_APROVADORES.TIPO_PRODUTO */
export type TipoProduto = 'CX' | 'CH' | 'PA' | 'CJ'

export interface FaixaAprovador {
  id: number
  planta: string
  nomePlanta: string
  tipoProduto: TipoProduto
  valorMinKg: number
  valorMaxKg: number
  valorMinM2: number
  valorMaxM2: number
  quantidadeAprovadores: number
}

export interface DadosAnalistaOrcamento {
  id?: number
  idUsuario: number | ''
  plantas: string[]
}

export interface DadosAprovadorOrcamento {
  id?: number
  idUsuario: number | ''
  nivel: NivelAprovador | ''
  plantas: string[]
}

export interface DadosFaixaAprovador {
  id?: number
  planta: string
  tipoProduto: TipoProduto | ''
  valorMinKg: string
  valorMaxKg: string
  valorMinM2: string
  valorMaxM2: string
  quantidadeAprovadores: string
}
