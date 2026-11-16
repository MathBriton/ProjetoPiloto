import type { LoginRetorno } from '@/types'

const BASE_URL = '/api'

export const CHAVE_TOKEN = 'swnet:token'
export const CHAVE_REFRESH_TOKEN = 'swnet:refreshToken'
export const EVENTO_SESSAO_EXPIRADA = 'swnet:sessao-expirada'

function obterToken(): string | null {
  return localStorage.getItem(CHAVE_TOKEN)
}

export function salvarTokens(token: string, refreshToken?: string) {
  localStorage.setItem(CHAVE_TOKEN, token)
  if (refreshToken) localStorage.setItem(CHAVE_REFRESH_TOKEN, refreshToken)
}

export function limparTokens() {
  localStorage.removeItem(CHAVE_TOKEN)
  localStorage.removeItem(CHAVE_REFRESH_TOKEN)
}

// Uma única renovação em andamento por vez: o backend rotaciona o refresh token
// e trata reuso como roubo (revoga a sessão), então requisições paralelas que
// recebem 401 precisam aguardar a mesma renovação em vez de disparar várias.
let renovacaoEmAndamento: Promise<boolean> | null = null

async function renovarToken(): Promise<boolean> {
  const refreshToken = localStorage.getItem(CHAVE_REFRESH_TOKEN)
  if (!refreshToken) return false

  renovacaoEmAndamento ??= (async () => {
    try {
      const resposta = await fetch(`${BASE_URL}/auth/refresh`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ refreshToken })
      })
      if (!resposta.ok) return false

      const corpo = (await resposta.json()) as LoginRetorno
      if (!corpo.sucesso || !corpo.token) return false

      salvarTokens(corpo.token, corpo.refreshToken)
      return true
    } catch {
      return false
    } finally {
      renovacaoEmAndamento = null
    }
  })()

  return renovacaoEmAndamento
}

async function requisitar<T>(caminho: string, opcoes: RequestInit = {}, jaRenovou = false): Promise<T> {
  const token = obterToken()

  const resposta = await fetch(`${BASE_URL}${caminho}`, {
    ...opcoes,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...opcoes.headers
    }
  })

  // Access token vencido/inválido: tenta renovar uma vez e repete a chamada
  if (resposta.status === 401 && token && !jaRenovou && !caminho.startsWith('/auth/')) {
    if (await renovarToken()) return requisitar<T>(caminho, opcoes, true)

    limparTokens()
    window.dispatchEvent(new Event(EVENTO_SESSAO_EXPIRADA))
  }

  if (!resposta.ok) {
    const corpo = await resposta.json().catch(() => null)
    throw new Error(corpo?.mensagemErro ?? `Erro ${resposta.status} ao chamar ${caminho}`)
  }

  if (resposta.status === 204) return undefined as T
  return resposta.json() as Promise<T>
}

export const apiClient = {
  get: <T>(caminho: string) => requisitar<T>(caminho),
  post: <T>(caminho: string, corpo?: unknown) =>
    requisitar<T>(caminho, { method: 'POST', body: corpo ? JSON.stringify(corpo) : undefined })
}
