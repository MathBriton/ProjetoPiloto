import type { ItemPermissao, PermissoesUsuario, TipoPermissao, UsuarioBusca } from '@/pages/administrador/tipos'

/**
 * Mock fiel ao AcessosUsuarios.asp legado, enquanto o backend (SwNet.Api)
 * ainda não expõe os endpoints reais de permissões. Sem lógica de backend
 * aqui — só dados em memória para viabilizar a navegação/interação no
 * frontend (busca de usuário, categorias de Serviços/Consultas, salvar,
 * importar permissões de outro usuário).
 */

const USUARIOS_MOCK: UsuarioBusca[] = [
  { matricula: '00123', filial: '01', nome: 'Administrador (fictício)', login: 'admin', ativo: true },
  { matricula: '00456', filial: '01', nome: 'João Silva', login: 'jsilva', ativo: true },
  { matricula: '00789', filial: '01', nome: 'Maria Alves', login: 'malves', ativo: true },
  { matricula: '00321', filial: '02', nome: 'Carlos Pereira', login: 'cpereira', ativo: true },
  { matricula: '00654', filial: '01', nome: 'Fernanda Souza', login: 'fsouza', ativo: false },
  { matricula: '00987', filial: '02', nome: 'Ricardo Nunes', login: 'rnunes', ativo: false }
]

const CATEGORIAS_SERVICOS = ['Cadastros', 'Administração', 'Financeiro', 'Manutenção']
const CATEGORIAS_CONSULTAS = ['Relatórios Gerenciais', 'Relatórios Operacionais', 'Auditoria']

const SERVICOS_POR_CATEGORIA: Record<string, string[]> = {
  Cadastros: ['Cadastro de Clientes', 'Cadastro de Fornecedores', 'Cadastro de Produtos'],
  Administração: ['Acessos Usuários', 'Perfis e Permissões', 'Log de Atividade'],
  Financeiro: ['Contas a Pagar', 'Contas a Receber', 'Fluxo de Caixa'],
  Manutenção: ['Ordens de Serviço', 'Plano de Manutenção']
}

const CONSULTAS_POR_CATEGORIA: Record<string, string[]> = {
  'Relatórios Gerenciais': ['Faturamento Mensal', 'Ranking de Clientes'],
  'Relatórios Operacionais': ['Relatório de Menu', 'Produção Diária', 'Apontamento de Horas'],
  Auditoria: ['Trilha de Alterações', 'Acessos por Usuário']
}

function construirItens(porCategoria: Record<string, string[]>, prefixo: string): ItemPermissao[] {
  let contador = 1
  const itens: ItemPermissao[] = []
  for (const categoria of Object.keys(porCategoria)) {
    for (const nome of porCategoria[categoria]) {
      itens.push({
        codigo: `${prefixo}${contador}`,
        nome,
        categoria,
        // Mock determinístico: ~40% dos itens já vêm habilitados
        habilitado: contador % 5 < 2
      })
      contador++
    }
  }
  return itens
}

function atraso<T>(valor: T, ms = 350): Promise<T> {
  return new Promise((resolve) => setTimeout(() => resolve(valor), ms))
}

export function listarCategorias() {
  return { servicos: CATEGORIAS_SERVICOS, consultas: CATEGORIAS_CONSULTAS }
}

export async function buscarUsuarios(termo: string, incluirInativos = false): Promise<UsuarioBusca[]> {
  const termoNormalizado = termo.trim().toLowerCase()
  const resultado = USUARIOS_MOCK.filter((usuario) => {
    if (!incluirInativos && !usuario.ativo) return false
    if (!termoNormalizado) return true
    return (
      usuario.nome.toLowerCase().includes(termoNormalizado) ||
      usuario.login.toLowerCase().includes(termoNormalizado)
    )
  })
  return atraso(resultado, 200)
}

export async function obterPermissoes(_matricula: string, _filial: string): Promise<PermissoesUsuario> {
  // TODO: quando a API .NET existir, trocar por GET /api/acessos-usuarios/{matricula}/{filial}
  return atraso({
    servicos: construirItens(SERVICOS_POR_CATEGORIA, 'SRV'),
    consultas: construirItens(CONSULTAS_POR_CATEGORIA, 'CON')
  })
}

export async function salvarPermissoes(
  _matricula: string,
  _filial: string,
  _servicos: ItemPermissao[],
  _consultas: ItemPermissao[]
): Promise<{ sucesso: boolean }> {
  // TODO: trocar por POST /api/acessos-usuarios/salvar quando o backend existir
  return atraso({ sucesso: true }, 500)
}

export async function importarPermissoes(
  _origem: UsuarioBusca,
  _destino: UsuarioBusca
): Promise<PermissoesUsuario> {
  // TODO: trocar por POST /api/acessos-usuarios/importar quando o backend existir
  // Mock: devolve as permissões do usuário de origem "aplicadas" ao destino
  return obterPermissoes(_origem.matricula, _origem.filial)
}

export type { ItemPermissao, PermissoesUsuario, TipoPermissao, UsuarioBusca }
