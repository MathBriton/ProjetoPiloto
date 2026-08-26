export interface UsuarioBusca {
  matricula: string
  filial: string
  nome: string
  login: string
  ativo: boolean
}

export interface ItemPermissao {
  codigo: string
  nome: string
  categoria: string
  habilitado: boolean
}

export interface PermissoesUsuario {
  servicos: ItemPermissao[]
  consultas: ItemPermissao[]
}

export type TipoPermissao = 'servico' | 'consulta'

export interface OpcaoSimples {
  id: number
  nome: string
}

export interface UsuarioCadastro {
  id: number
  login: string
  nomeCompleto: string
  email: string
  perfilId: number
  plantaIds: number[]
  ativo: boolean
  criadoEm: string
}

export type NovoUsuario = Omit<UsuarioCadastro, 'id' | 'criadoEm'>

/* ------------------------------------------------------------------ *
 * Usuário SWnet (espelha CRM_DTI do legado ASP)
 * ------------------------------------------------------------------ */

export type Sexo = 'M' | 'F'

export interface UsuarioSwnet {
  /** CRM_R_E_C_N_O */
  id: number
  /** CRM_COD */
  matricula: string
  /** CRM_NOME */
  nome: string
  /** CRM_LOGIN */
  login: string
  /** CRM_USERAD — no formato PREFIXO\login */
  userAd: string
  /** CRM_EMAILPRIMARIO */
  email: string
  /** RA_SEXO */
  sexo: Sexo
  /** RA_NASC (ISO yyyy-mm-dd) */
  nascimento: string
  /** CRM_RAMAL */
  ramal: string
  /** CRM_TIPO */
  tipo: string
  /** EMP_CODIGO */
  empresa: string
  /** CRM_FUNCAO */
  funcao: string
  /** CRM_VALNIVEL */
  representante: boolean
  /** Bloqueado */
  bloqueado: boolean
  /** CRM_D_E_L_E_T <> '*' */
  ativo: boolean
  /** CRM_DATCAD (ISO) */
  dataCadastro: string
  /** CRM_DATACES (ISO) — null quando nunca acessou */
  ultimoAcesso: string | null
  /** DataBloqueio (ISO) */
  dataBloqueio: string | null
}

export type PrefixoAd = 'GROUP' | 'WESTROCK'

export interface DadosCadastroUsuario {
  matricula: string
  nome: string
  sexo: Sexo
  nascimento: string
  email: string
  representante: boolean
  prefixoAd: PrefixoAd
  login: string
  senha: string
  tipo: string
  empresa: string
  ramal: string
  enviarEmail: boolean
}

export interface DadosAlteracaoUsuario extends Omit<DadosCadastroUsuario, 'enviarEmail'> {
  id: number
  bloqueado: boolean
}

/* ------------------------------------------------------------------ *
 * Permissões especiais (CRM_DTI — colunas de nível/flag)
 * ------------------------------------------------------------------ */

export type NivelAlteraPedido = '' | '0' | '14'
export type NivelSolicitacaoAlteracaoPedido = '' | 'V' | 'P'
export type NivelSpd = '' | 'C' | 'D'
export type CodigoPermissao = '' | 'C' | 'G' | 'L'

export interface PermissoesEspeciais {
  idUsuario: number
  nome: string
  nivelAlteraPedido: NivelAlteraPedido
  nivelSolicitacaoAlteracaoPedido: NivelSolicitacaoAlteracaoPedido
  nivelSpd: NivelSpd
  permissao: CodigoPermissao
  verTodasFncs: boolean
  alteraPo: boolean
  alteraPa: boolean
  ativaUsuario: boolean
  validaPreCadastroProduto: boolean
  eTools: boolean
  nivelLimiteCredito: boolean
  nivelControladoria: boolean
  nivelE2Open: boolean
  nivelControlesInternos: boolean
  nivelFinanceiro: boolean
  nivelDiretorComercial: boolean
  nivelFiscal: boolean
}

/* ------------------------------------------------------------------ *
 * Gerenciar menus (LINKS/SUBMENU e RELATORIOSCAPA/CAT_RELATORIOS)
 * ------------------------------------------------------------------ */

export type TipoMenu = 'servico' | 'consulta'

export interface PermissaoMenu {
  codigo: string
  nome: string
  url: string
  extensao: string
  /** Só existe em Consultas (REL_BANCO) */
  banco?: string
  ativo: boolean
}

export interface CategoriaMenu {
  codigo: string
  nome: string
  permissoes: PermissaoMenu[]
}

export interface MenusCompletos {
  servicos: CategoriaMenu[]
  consultas: CategoriaMenu[]
}

export interface DadosPermissaoMenu {
  tipo: TipoMenu
  codigoCategoria: string
  codigo?: string
  nome: string
  url: string
  banco: string
  extensao: string
}

/* ------------------------------------------------------------------ *
 * Relatório de permissões de usuário
 * ------------------------------------------------------------------ */

export interface AcessoConcedido {
  codSubcategoria: string
  nomeSubcategoria: string
  descricao: string
}

export interface LinhaRelatorioPermissao {
  id: number
  matricula: string
  codPlanta: string
  nomePlanta: string
  funcao: string
  login: string
  nome: string
  email: string
  dataCadastro: string
  servicos: AcessoConcedido[]
  relatorios: AcessoConcedido[]
}

export type CategoriaRelatorio = '' | '1' | '2'

export interface FiltrosRelatorioPermissao {
  filial: string
  categoria: CategoriaRelatorio
  subcategorias: string[]
  id: string
  matricula: string
  email: string
}

export interface Subcategoria {
  id: string
  idCategoria: '1' | '2'
  descricao: string
}

/* ------------------------------------------------------------------ *
 * Cancelamento (desligamento) de usuários — múltiplos sistemas
 * ------------------------------------------------------------------ */

export interface Sistema {
  /** chave usada no repositório legado (swnet, benTrimbox, ...) */
  id: string
  nome: string
}

export interface ContaSistema {
  /** id do registro no sistema de origem */
  id: number
  sistemaId: string
  sistema: string
  nome: string
  email: string
  dataCadastro: string | null
  dataBloqueio: string | null
  ativo: boolean
  /** true quando o e-mail não tem conta naquele sistema */
  semRegistro: boolean
}

export type AcaoCancelamento = 'habilitar' | 'desabilitar'

/* ------------------------------------------------------------------ *
 * Revisão de usuários (inatividade)
 * ------------------------------------------------------------------ */

export interface UsuarioInativo {
  id: number
  sistemaId: string
  sistema: string
  nome: string
  email: string
  ultimoAcesso: string | null
  diasSemAcessar: number
}
