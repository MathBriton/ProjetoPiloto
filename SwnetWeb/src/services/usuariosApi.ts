import type { NovoUsuario, OpcaoSimples, UsuarioCadastro } from '@/pages/administrador/tipos'

/**
 * Mock em memória enquanto a SwNet.Api não expõe os endpoints de cadastro.
 * Os dados se perdem ao recarregar a página — é só para viabilizar o fluxo no frontend.
 */

const PERFIS: OpcaoSimples[] = [
  { id: 1, nome: 'Administrador' },
  { id: 2, nome: 'Supervisor' },
  { id: 3, nome: 'Operador' },
  { id: 4, nome: 'Consulta' }
]

const PLANTAS: OpcaoSimples[] = [
  { id: 1, nome: 'JF01 — Juiz de Fora' },
  { id: 2, nome: 'SP01 — São Paulo' }
]

let usuarios: UsuarioCadastro[] = [
  {
    id: 1,
    login: 'admin',
    nomeCompleto: 'Administrador (fictício)',
    email: 'admin@swnet.local',
    perfilId: 1,
    plantaIds: [1, 2],
    ativo: true,
    criadoEm: '2026-01-15T09:00:00.000Z'
  },
  {
    id: 2,
    login: 'jsilva',
    nomeCompleto: 'João Silva',
    email: 'joao.silva@swnet.local',
    perfilId: 3,
    plantaIds: [1],
    ativo: true,
    criadoEm: '2026-03-02T13:20:00.000Z'
  },
  {
    id: 3,
    login: 'malves',
    nomeCompleto: 'Maria Alves',
    email: 'maria.alves@swnet.local',
    perfilId: 2,
    plantaIds: [2],
    ativo: false,
    criadoEm: '2026-05-11T17:45:00.000Z'
  }
]

let proximoId = 4

function atraso<T>(valor: T, ms = 300): Promise<T> {
  return new Promise((resolve) => setTimeout(() => resolve(valor), ms))
}

export function listarPerfis() {
  return PERFIS
}

export function listarPlantas() {
  return PLANTAS
}

export async function listarUsuarios(termo = ''): Promise<UsuarioCadastro[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = busca
    ? usuarios.filter(
        (u) =>
          u.nomeCompleto.toLowerCase().includes(busca) ||
          u.login.toLowerCase().includes(busca) ||
          u.email.toLowerCase().includes(busca)
      )
    : usuarios
  return atraso([...resultado], 200)
}

export async function loginDisponivel(login: string, ignorarId?: number): Promise<boolean> {
  const alvo = login.trim().toLowerCase()
  return atraso(!usuarios.some((u) => u.login.toLowerCase() === alvo && u.id !== ignorarId), 150)
}

export async function criarUsuario(dados: NovoUsuario): Promise<UsuarioCadastro> {
  const novo: UsuarioCadastro = { ...dados, id: proximoId++, criadoEm: new Date().toISOString() }
  usuarios = [...usuarios, novo]
  return atraso(novo, 400)
}

export async function alternarAtivo(id: number): Promise<UsuarioCadastro | undefined> {
  usuarios = usuarios.map((u) => (u.id === id ? { ...u, ativo: !u.ativo } : u))
  return atraso(usuarios.find((u) => u.id === id), 250)
}
