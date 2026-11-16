export interface Planta {
  id: number
  codigo: string
  nome: string
}

export interface UsuarioLogado {
  id: number
  login: string
  nomeCompleto: string
  perfil: string
  senhaProvisoria?: boolean
  permissoes: string[]
  plantas: Planta[]
  plantaAtualId: number
}

export interface LoginRetorno {
  sucesso: boolean
  mensagemErro?: string
  token?: string
  refreshToken?: string
  expiraEm?: string
  usuario?: UsuarioLogado
}
