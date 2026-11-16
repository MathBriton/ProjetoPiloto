import { createContext, useContext, useState, useCallback, useEffect, type ReactNode } from 'react'
import type { UsuarioLogado } from '@/types'
import { login as loginApi, logout as logoutApi } from '@/services/autenticacaoApi'
import { salvarTokens, limparTokens, EVENTO_SESSAO_EXPIRADA } from '@/services/apiClient'

interface AuthContextValor {
  usuario: UsuarioLogado | null
  autenticado: boolean
  carregando: boolean
  entrar: (login: string, senha: string) => Promise<{ sucesso: boolean; mensagemErro?: string }>
  sair: () => void
  temPermissao: (chave: string) => boolean
}

const AuthContext = createContext<AuthContextValor | undefined>(undefined)

const CHAVE_STORAGE = 'swnet:usuario'

function carregarUsuarioSalvo(): UsuarioLogado | null {
  const bruto = localStorage.getItem(CHAVE_STORAGE)
  return bruto ? (JSON.parse(bruto) as UsuarioLogado) : null
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [usuario, setUsuario] = useState<UsuarioLogado | null>(carregarUsuarioSalvo)
  const [carregando, setCarregando] = useState(false)

  const entrar = useCallback(async (loginDigitado: string, senha: string) => {
    setCarregando(true)
    try {
      const resultado = await loginApi(loginDigitado, senha)

      if (!resultado.sucesso || !resultado.usuario || !resultado.token) {
        return { sucesso: false, mensagemErro: resultado.mensagemErro ?? 'Falha ao autenticar.' }
      }

      salvarTokens(resultado.token, resultado.refreshToken)
      localStorage.setItem(CHAVE_STORAGE, JSON.stringify(resultado.usuario))
      setUsuario(resultado.usuario)

      return { sucesso: true }
    } finally {
      setCarregando(false)
    }
  }, [])

  const sair = useCallback(() => {
    void logoutApi()
    localStorage.removeItem(CHAVE_STORAGE)
    setUsuario(null)
  }, [])

  // Refresh falhou no apiClient (sessão revogada/expirada): volta para o login
  useEffect(() => {
    const aoExpirar = () => {
      limparTokens()
      localStorage.removeItem(CHAVE_STORAGE)
      setUsuario(null)
    }
    window.addEventListener(EVENTO_SESSAO_EXPIRADA, aoExpirar)
    return () => window.removeEventListener(EVENTO_SESSAO_EXPIRADA, aoExpirar)
  }, [])

  const temPermissao = useCallback(
    (chave: string) => usuario?.permissoes.includes(chave) ?? false,
    [usuario]
  )

  return (
    <AuthContext.Provider
      value={{ usuario, autenticado: !!usuario, carregando, entrar, sair, temPermissao }}
    >
      {children}
    </AuthContext.Provider>
  )
}

export function useAuth() {
  const contexto = useContext(AuthContext)
  if (!contexto) throw new Error('useAuth precisa estar dentro de <AuthProvider>')
  return contexto
}
