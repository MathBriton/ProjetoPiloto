/** Tipos das telas de Cadastros-FNC (tabelas Setor, Defeito, Responsavel, AprovadorFNC). */

export interface SetorFnc {
  /** Setor.Id — o legado gera por MAX(Id)+1, a tabela não tem identity */
  id: number
  descricao: string
}

/** Defeito.Tipo — ADM = Administrativo, PRO = Industrial */
export type TipoDefeito = 'ADM' | 'PRO'

export interface DefeitoFnc {
  id: number
  tipo: TipoDefeito
  idSetor: number
  nomeSetor: string
  descricao: string
}

export interface AprovadorFnc {
  /** AprovadorFNC.Id — é o CRM_R_E_C_N_O do usuário escolhido */
  id: number
  nome: string
  email: string
  dataInclusao: string
}

/** Responsavel.Tipo — 1 = Analista, 2 = Responsável Setorial */
export type TipoResponsavel = '1' | '2'

export interface ResponsavelFnc {
  /** Responsavel.Id — é o CRM_R_E_C_N_O do usuário escolhido */
  id: number
  nome: string
  email: string
  tipo: TipoResponsavel
  idSetor: number
  nomeSetor: string
  /** Uma linha por planta no banco; aqui agrupadas como no STUFF/FOR XML do ASP */
  plantas: string[]
  dataInclusao: string
}

export interface DadosSetorFnc {
  id?: number
  descricao: string
}

export interface DadosDefeitoFnc {
  id?: number
  tipo: TipoDefeito | ''
  idSetor: number | ''
  descricao: string
}

export interface DadosResponsavelFnc {
  id?: number
  idUsuario: number | ''
  idSetor: number | ''
  tipo: TipoResponsavel | ''
  plantas: string[]
}

export interface ResultadoCrud {
  sucesso: boolean
  mensagem: string
}
