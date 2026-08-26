import type {
  AcessoConcedido,
  CategoriaRelatorio,
  FiltrosRelatorioPermissao,
  LinhaRelatorioPermissao,
  Subcategoria
} from '@/pages/administrador/tipos'
import { atraso, nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'
import { listarMenus } from './gerenciarMenusApi'

/**
 * RelatorioPermissaoUsuariosRepositorio.asp — ações listarUsuarios,
 * listarSubcategorias e listarAcessos.
 *
 * As subcategorias saem das mesmas tabelas do Gerenciar Menus (SUBMENU para
 * serviços, CAT_RELATORIOS para relatórios), então aqui reaproveitamos aquele
 * mock em vez de duplicar a lista.
 */

/** Mapa determinístico usuário → códigos de permissão concedidos */
function codigosConcedidos(idUsuario: number, total: number) {
  const concedidos = new Set<number>()
  for (let i = 0; i < total; i++) {
    if ((idUsuario + i * 7) % 3 !== 0) concedidos.add(i)
  }
  return concedidos
}

export async function listarSubcategorias(categoria: CategoriaRelatorio): Promise<Subcategoria[]> {
  const menus = await listarMenus()
  const lista: Subcategoria[] = []

  if (categoria === '' || categoria === '1') {
    lista.push(
      ...menus.servicos.map((c) => ({ id: c.codigo, idCategoria: '1' as const, descricao: c.nome }))
    )
  }
  if (categoria === '' || categoria === '2') {
    lista.push(
      ...menus.consultas.map((c) => ({ id: c.codigo, idCategoria: '2' as const, descricao: c.nome }))
    )
  }

  return lista
}

export async function listarUsuariosComPermissoes(
  filtros: FiltrosRelatorioPermissao
): Promise<LinhaRelatorioPermissao[]> {
  const menus = await listarMenus()

  const servicosDisponiveis: AcessoConcedido[] = menus.servicos
    .filter((c) => filtros.categoria !== '1' || !filtros.subcategorias.length || filtros.subcategorias.includes(c.codigo))
    .flatMap((c) =>
      c.permissoes
        .filter((p) => p.ativo)
        .map((p) => ({ codSubcategoria: p.codigo, nomeSubcategoria: c.nome, descricao: p.nome }))
    )

  const relatoriosDisponiveis: AcessoConcedido[] = menus.consultas
    .filter((c) => filtros.categoria !== '2' || !filtros.subcategorias.length || filtros.subcategorias.includes(c.codigo))
    .flatMap((c) =>
      c.permissoes
        .filter((p) => p.ativo)
        .map((p) => ({ codSubcategoria: p.codigo, nomeSubcategoria: c.nome, descricao: p.nome }))
    )

  const linhas = usuariosSwnet
    // O legado só considera usuários ativos e não bloqueados (CRM_D_E_L_E_T = '' e DataBloqueio IS NULL)
    .filter((u) => u.ativo && !u.dataBloqueio)
    .filter((u) => !filtros.filial || u.empresa === filtros.filial)
    .filter((u) => !filtros.id || String(u.id).includes(filtros.id.trim()))
    .filter((u) => !filtros.matricula || u.matricula.includes(filtros.matricula.trim()))
    .filter((u) => !filtros.email || u.email.toLowerCase().includes(filtros.email.trim().toLowerCase()))
    .map<LinhaRelatorioPermissao>((u) => {
      const marcadosServico = codigosConcedidos(u.id, servicosDisponiveis.length)
      const marcadosRelatorio = codigosConcedidos(u.id + 5, relatoriosDisponiveis.length)

      return {
        id: u.id,
        matricula: u.matricula,
        codPlanta: u.empresa,
        nomePlanta: nomeEmpresa(u.empresa),
        funcao: u.funcao,
        login: u.login,
        nome: u.nome,
        email: u.email,
        dataCadastro: u.dataCadastro,
        servicos:
          filtros.categoria === '2' ? [] : servicosDisponiveis.filter((_, i) => marcadosServico.has(i)),
        relatorios:
          filtros.categoria === '1' ? [] : relatoriosDisponiveis.filter((_, i) => marcadosRelatorio.has(i))
      }
    })
    // O ASP descarta quem não tem nenhum acesso nas categorias filtradas
    .filter((linha) => linha.servicos.length > 0 || linha.relatorios.length > 0)

  return atraso(linhas, 500)
}
