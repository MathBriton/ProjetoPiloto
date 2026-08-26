import { apiClient } from './apiClient'
import type { LoginRetorno } from '@/types'

/**
 * Enquanto o backend real (SwNet.Api) não está de pé, o login fica mockado
 * aqui para permitir testar a navegação ponta a ponta no frontend.
 *
 * Usuário fictício: admin / Admin@123 (mesmo par usado no seed de dev do
 * backend em SwNet.Infraestrutura/Contexto/SwNetDbSeeder.cs).
 *
 * Para conectar na API real, troque USAR_MOCK para false.
 */
const USAR_MOCK = true

const USUARIO_FICTICIO = { login: 'admin', senha: 'Admin@123' }

const RETORNO_MOCK: LoginRetorno = {
  sucesso: true,
  token: 'token-ficticio-de-desenvolvimento',
  refreshToken: 'refresh-ficticio',
  expiraEm: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
  usuario: {
    id: 1,
    login: 'admin',
    nomeCompleto: 'Administrador (fictício)',
    perfil: 'Administrador',
    permissoes: [
      'acessos-usuarios.visualizar',
      'acessos-usuarios.editar',
      'usuarios.incluir',
      'usuarios.alterar',
      'usuarios.ativar-desativar',
      'usuarios.cancelar',
      'usuarios.criacao-ad',
      'usuarios.revisao',
      'menus.gerenciar',
      'permissoes-especiais.editar',
      'permissoes.relatorio',
      // Equivale ao NivelControladoria da CRM_DTI, que libera a edição completa
      // no Controle de Recebimento (Cavaco/Aparas)
      'recebimento.controladoria'
    ],
    plantaAtualId: 1,
    plantas: [
      { id: 1, codigo: 'JF01', nome: 'Juiz de Fora' },
      { id: 2, codigo: 'SP01', nome: 'São Paulo' }
    ]
  }
}

function atraso(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

export async function login(loginDigitado: string, senha: string): Promise<LoginRetorno> {
  if (USAR_MOCK) {
    await atraso(400)
    const credenciaisValidas =
      loginDigitado.trim().toLowerCase() === USUARIO_FICTICIO.login && senha === USUARIO_FICTICIO.senha

    return credenciaisValidas
      ? RETORNO_MOCK
      : { sucesso: false, mensagemErro: 'Usuário ou senha inválidos.' }
  }

  return apiClient.post<LoginRetorno>('/auth/login', { login: loginDigitado, senha })
}
