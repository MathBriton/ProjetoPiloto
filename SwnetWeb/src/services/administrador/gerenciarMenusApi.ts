import type {
  CategoriaMenu,
  DadosPermissaoMenu,
  MenusCompletos,
  PermissaoMenu,
  TipoMenu
} from '@/pages/administrador/tipos'
import { atraso } from '@/services/comuns/baseMock'

/**
 * CriarPermissao.asp (Gerenciar Menus).
 *
 * Serviços vêm de SUBMENU + LINKS e Consultas de CAT_RELATORIOS +
 * RELATORIOSCAPA. Excluir é soft delete (LIN_D_E_L_E_T / REL_D_E_L_E_T_E_D
 * = '*'), por isso aqui `ativo: false` em vez de remover do array.
 */

function perm(
  codigo: string,
  nome: string,
  url: string,
  banco?: string,
  ativo = true
): PermissaoMenu {
  return { codigo, nome, url, extensao: '.asp', banco, ativo }
}

const servicos: CategoriaMenu[] = [
  {
    codigo: 'S01',
    nome: 'Administrador',
    permissoes: [
      perm('101', 'Acessos Usuários', '../Usuario/AcessosUsuarios'),
      perm('102', 'Cadastro de Usuário', '../Usuario/CadUsuario'),
      perm('103', 'Alterar Usuário', '../Usuario/AltUsuario'),
      perm('104', 'Ativar/Desativar Usuário', '../Usuario/ExcUsuario'),
      perm('105', 'Gerenciar Menus', '../CriarPermissao/CriarPermissao'),
      perm('106', 'Permissões Especiais', '../Usuario/PermissaoUsuario'),
      perm('107', 'Menu Antigo de Acessos', '../Usuario/AcessosLegado', undefined, false)
    ]
  },
  {
    codigo: 'S02',
    nome: 'Cadastros FNC',
    permissoes: [
      perm('201', 'Cadastro de FNC', '../FNC/CadFNC'),
      perm('202', 'Aprovação de FNC', '../FNC/AprovFNC'),
      perm('203', 'Revisão de FNC', '../FNC/RevFNC')
    ]
  },
  {
    codigo: 'S03',
    nome: 'Compras',
    permissoes: [
      perm('301', 'Solicitação de Compra', '../Compras/SolCompra'),
      perm('302', 'Aprovação de Pedido', '../Compras/AprovPedido'),
      perm('303', 'Cotação de Fornecedores', '../Compras/Cotacao')
    ]
  },
  {
    codigo: 'S04',
    nome: 'Expedição',
    permissoes: [
      perm('401', 'Programação de Cargas', '../Expedicao/ProgCargas'),
      perm('402', 'Romaneio de Carga', '../Expedicao/Romaneio')
    ]
  },
  {
    codigo: 'S05',
    nome: 'Financeiro',
    permissoes: [
      perm('501', 'Contas a Pagar', '../Financeiro/ContasPagar'),
      perm('502', 'Contas a Receber', '../Financeiro/ContasReceber'),
      perm('503', 'Limite de Crédito', '../Financeiro/LimiteCredito')
    ]
  },
  {
    codigo: 'S06',
    nome: 'RH',
    permissoes: [
      perm('601', 'Ponto Eletrônico', '../RH/Ponto'),
      perm('602', 'Solicitação de Férias', '../RH/Ferias', undefined, false)
    ]
  }
]

const consultas: CategoriaMenu[] = [
  {
    codigo: 'C01',
    nome: 'Relatórios Gerenciais',
    permissoes: [
      perm('801', 'Faturamento Mensal', '../Rel/FaturamentoMensal', 'HELPDESK'),
      perm('802', 'Ranking de Clientes', '../Rel/RankingClientes', 'HELPDESK'),
      perm('803', 'Margem por Produto', '../Rel/MargemProduto', 'PROTHEUS')
    ]
  },
  {
    codigo: 'C02',
    nome: 'Relatórios Operacionais',
    permissoes: [
      perm('811', 'Produção Diária', '../Rel/ProducaoDiaria', 'TRIMBOX'),
      perm('812', 'Apontamento de Horas', '../Rel/ApontamentoHoras', 'TRIMBOX'),
      perm('813', 'Refugo por Máquina', '../Rel/RefugoMaquina', 'TRIMBOX'),
      perm('814', 'Consumo de Aparas', '../Rel/ConsumoAparas', 'TRIMPAPER', false)
    ]
  },
  {
    codigo: 'C03',
    nome: 'Auditoria',
    permissoes: [
      perm('821', 'Trilha de Alterações', '../Rel/TrilhaAlteracoes', 'CORPORATIVO'),
      perm('822', 'Acessos por Usuário', '../Rel/AcessosUsuario', 'CORPORATIVO'),
      perm('823', 'Relatório de Permissões', '../Rel/RelatorioPermissaoUsuarios', 'HELPDESK')
    ]
  },
  {
    codigo: 'C04',
    nome: 'Qualidade (SGQ)',
    permissoes: [
      perm('831', 'Não Conformidades', '../Rel/NaoConformidades', 'HELPDESK'),
      perm('832', 'Índice de Retrabalho', '../Rel/IndiceRetrabalho', 'HELPDESK')
    ]
  }
]

function base(tipo: TipoMenu) {
  return tipo === 'servico' ? servicos : consultas
}

function proximoCodigo(tipo: TipoMenu) {
  const codigos = base(tipo)
    .flatMap((c) => c.permissoes)
    .map((p) => Number(p.codigo))
    .filter((n) => !Number.isNaN(n))
  return String(Math.max(0, ...codigos) + 1)
}

function clonarCategorias(categorias: CategoriaMenu[]): CategoriaMenu[] {
  return categorias.map((c) => ({ ...c, permissoes: c.permissoes.map((p) => ({ ...p })) }))
}

export async function listarMenus(): Promise<MenusCompletos> {
  return atraso({ servicos: clonarCategorias(servicos), consultas: clonarCategorias(consultas) }, 400)
}

export async function criarPermissao(
  dados: DadosPermissaoMenu
): Promise<{ sucesso: boolean; mensagem: string }> {
  const erro = validar(dados)
  if (erro) return atraso({ sucesso: false, mensagem: erro }, 200)

  const categoria = base(dados.tipo).find((c) => c.codigo === dados.codigoCategoria)
  if (!categoria) return atraso({ sucesso: false, mensagem: 'Categoria não encontrada.' }, 200)

  categoria.permissoes.push({
    codigo: proximoCodigo(dados.tipo),
    nome: dados.nome.trim().slice(0, 50),
    url: dados.url.trim(),
    extensao: dados.extensao || '.asp',
    banco: dados.tipo === 'consulta' ? dados.banco.trim() : undefined,
    ativo: true
  })

  return atraso({ sucesso: true, mensagem: 'Permissão criada com sucesso.' }, 400)
}

export async function editarPermissao(
  dados: DadosPermissaoMenu
): Promise<{ sucesso: boolean; mensagem: string }> {
  const erro = validar(dados)
  if (erro) return atraso({ sucesso: false, mensagem: erro }, 200)
  if (!dados.codigo) {
    return atraso({ sucesso: false, mensagem: 'Código da permissão não informado.' }, 200)
  }

  const categorias = base(dados.tipo)
  const origem = categorias.find((c) => c.permissoes.some((p) => p.codigo === dados.codigo))
  const destino = categorias.find((c) => c.codigo === dados.codigoCategoria)
  if (!origem || !destino) {
    return atraso({ sucesso: false, mensagem: 'Categoria destino não encontrada.' }, 200)
  }

  const indice = origem.permissoes.findIndex((p) => p.codigo === dados.codigo)
  const atualizada: PermissaoMenu = {
    ...origem.permissoes[indice],
    nome: dados.nome.trim().slice(0, 50),
    url: dados.url.trim(),
    extensao: dados.extensao || '.asp',
    banco: dados.tipo === 'consulta' ? dados.banco.trim() : undefined
  }

  // Trocar de categoria = mover o registro (o legado só troca SUB_COD/CAT_COD)
  origem.permissoes.splice(indice, 1)
  destino.permissoes.push(atualizada)

  return atraso({ sucesso: true, mensagem: 'Permissão alterada com sucesso.' }, 400)
}

/** Soft delete, igual ao legado (marca como excluída em vez de apagar) */
export async function excluirPermissao(
  tipo: TipoMenu,
  codigo: string
): Promise<{ sucesso: boolean; mensagem: string }> {
  return definirAtivo(tipo, codigo, false, 'Permissão excluída com sucesso.')
}

export async function reativarPermissao(
  tipo: TipoMenu,
  codigo: string
): Promise<{ sucesso: boolean; mensagem: string }> {
  return definirAtivo(tipo, codigo, true, 'Permissão reativada com sucesso.')
}

async function definirAtivo(tipo: TipoMenu, codigo: string, ativo: boolean, mensagem: string) {
  const permissao = base(tipo)
    .flatMap((c) => c.permissoes)
    .find((p) => p.codigo === codigo)

  if (!permissao) return atraso({ sucesso: false, mensagem: 'Permissão não encontrada.' }, 200)

  permissao.ativo = ativo
  return atraso({ sucesso: true, mensagem }, 350)
}

function validar(dados: DadosPermissaoMenu): string | null {
  if (!dados.nome.trim()) return 'Nome da permissão não pode ser vazio.'
  if (!dados.codigoCategoria) return 'Categoria destino não informada.'
  if (!dados.url.trim()) {
    return dados.tipo === 'servico'
      ? 'Informe a URL (LIN_URL) para Serviços.'
      : 'Informe a URL (REL_URL) para Consultas.'
  }
  if (dados.tipo === 'consulta' && !dados.banco.trim()) {
    return 'Informe o Banco (REL_BANCO) para Consultas.'
  }
  return null
}
