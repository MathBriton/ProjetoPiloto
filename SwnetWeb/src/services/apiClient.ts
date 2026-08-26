const BASE_URL = '/api'

function obterToken(): string | null {
  return localStorage.getItem('swnet:token')
}

async function requisitar<T>(caminho: string, opcoes: RequestInit = {}): Promise<T> {
  const token = obterToken()

  const resposta = await fetch(`${BASE_URL}${caminho}`, {
    ...opcoes,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...opcoes.headers
    }
  })

  if (!resposta.ok) {
    const corpo = await resposta.json().catch(() => null)
    throw new Error(corpo?.mensagemErro ?? `Erro ${resposta.status} ao chamar ${caminho}`)
  }

  return resposta.json() as Promise<T>
}

export const apiClient = {
  get: <T>(caminho: string) => requisitar<T>(caminho),
  post: <T>(caminho: string, corpo?: unknown) =>
    requisitar<T>(caminho, { method: 'POST', body: corpo ? JSON.stringify(corpo) : undefined })
}
