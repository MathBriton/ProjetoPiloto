import { useEffect, useState } from 'react'
import {
  BarChart3,
  Eraser,
  FileDown,
  FileSpreadsheet,
  ScrollText,
  Search,
  SlidersHorizontal,
  Wrench
} from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, EstadoVazio, LinhaCarregando, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import {
  listarSubcategorias,
  listarUsuariosComPermissoes
} from '@/services/administrador/relatorioPermissoesApi'
import { EMPRESAS, codigoEmpresa, formatarData } from '@/services/comuns/baseMock'
import type {
  CategoriaRelatorio,
  FiltrosRelatorioPermissao,
  LinhaRelatorioPermissao,
  Subcategoria
} from './tipos'

/** Refatoração de Relatorio-De-Permissoes-De-Usuarios/RelatorioPermissaoUsuarios.asp. */

const FILTROS_VAZIOS: FiltrosRelatorioPermissao = {
  filial: '',
  categoria: '',
  subcategorias: [],
  id: '',
  matricula: '',
  email: ''
}

export function RelatorioPermissoes() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [filtros, setFiltros] = useState<FiltrosRelatorioPermissao>(FILTROS_VAZIOS)
  const [subcategorias, setSubcategorias] = useState<Subcategoria[]>([])
  const [linhas, setLinhas] = useState<LinhaRelatorioPermissao[]>([])
  const [carregando, setCarregando] = useState(false)
  const [pesquisaFeita, setPesquisaFeita] = useState(false)
  const [selecionados, setSelecionados] = useState<Set<number>>(new Set())
  const [detalhe, setDetalhe] = useState<LinhaRelatorioPermissao | null>(null)

  // As subcategorias dependem da categoria escolhida (mesma regra do ASP)
  useEffect(() => {
    listarSubcategorias(filtros.categoria).then(setSubcategorias)
  }, [filtros.categoria])

  function mudarFiltro<C extends keyof FiltrosRelatorioPermissao>(
    campo: C,
    valor: FiltrosRelatorioPermissao[C]
  ) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function mudarCategoria(categoria: CategoriaRelatorio) {
    setFiltros((atual) => ({ ...atual, categoria, subcategorias: [] }))
  }

  async function pesquisar() {
    setCarregando(true)
    setSelecionados(new Set())
    try {
      setLinhas(await listarUsuariosComPermissoes(filtros))
      setPesquisaFeita(true)
    } finally {
      setCarregando(false)
    }
  }

  function limpar() {
    setFiltros(FILTROS_VAZIOS)
    setLinhas([])
    setSelecionados(new Set())
    setPesquisaFeita(false)
  }

  function alternarLinha(id: number) {
    setSelecionados((atual) => {
      const novo = new Set(atual)
      if (novo.has(id)) novo.delete(id)
      else novo.add(id)
      return novo
    })
  }

  const todasMarcadas = linhas.length > 0 && linhas.every((l) => selecionados.has(l.id))

  function exportar(formato: 'excel' | 'pdf') {
    const quantidade = selecionados.size > 0 ? selecionados.size : linhas.length
    if (quantidade === 0) {
      mostrarToast('Não há registros para exportar.', 'warning')
      return
    }
    // TODO: o legado usa SheetJS/jsPDF; sem lib de exportação no frontend ainda
    mostrarToast(
      `Exportação em ${formato === 'excel' ? 'Excel' : 'PDF'} (${quantidade} registro(s)) será implementada em breve.`,
      'info'
    )
  }

  const subcategoriasDisponiveis = subcategorias.filter(
    (s) => !filtros.categoria || s.idCategoria === filtros.categoria
  )

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Relatório de Permissões de Usuário" />

      <Painel titulo="Filtros" icone={SlidersHorizontal}>
        <form
          className="grid gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            pesquisar()
          }}
        >
          <Selecao
            id="filtroFilial"
            rotulo="Filial"
            value={filtros.filial}
            onChange={(e) => mudarFiltro('filial', e.target.value)}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Todas as filiais</option>
            {EMPRESAS.map((empresa) => (
              <option key={empresa.id} value={codigoEmpresa(empresa.id)}>
                {empresa.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroCategoria"
            rotulo="Categoria"
            value={filtros.categoria}
            onChange={(e) => mudarCategoria(e.target.value as CategoriaRelatorio)}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Todas</option>
            <option value="1">Serviços</option>
            <option value="2">Relatórios</option>
          </Selecao>

          <Campo rotulo="Subcategorias" htmlFor="filtroSubcategoria" className="sm:col-span-4">
            <select
              id="filtroSubcategoria"
              multiple
              size={3}
              value={filtros.subcategorias}
              disabled={subcategoriasDisponiveis.length === 0}
              onChange={(e) =>
                mudarFiltro(
                  'subcategorias',
                  [...e.target.selectedOptions].map((o) => o.value)
                )
              }
              className={classesControle}
            >
              {subcategoriasDisponiveis.map((sub) => (
                <option key={`${sub.idCategoria}-${sub.id}`} value={sub.id}>
                  {sub.descricao}
                </option>
              ))}
            </select>
          </Campo>

          <Entrada
            id="filtroId"
            rotulo="Id"
            value={filtros.id}
            onChange={(e) => mudarFiltro('id', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroMatricula"
            rotulo="Matrícula"
            value={filtros.matricula}
            onChange={(e) => mudarFiltro('matricula', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroEmail"
            rotulo="E-mail"
            value={filtros.email}
            onChange={(e) => mudarFiltro('email', e.target.value)}
            larguraClasse="sm:col-span-4"
          />

          <div className="flex items-end gap-2 sm:col-span-2">
            <Botao variante="neutro" icone={Eraser} aoClicar={limpar} desabilitado={carregando}>
              Limpar
            </Botao>
            <Botao tipo="submit" icone={Search} desabilitado={carregando}>
              {carregando ? '...' : 'Pesquisar'}
            </Botao>
          </div>
        </form>
      </Painel>

      {pesquisaFeita && (
        <Painel
          titulo={`${linhas.length} registro(s)`}
          icone={ScrollText}
          acoes={
            <>
              <Botao variante="sucesso" icone={FileSpreadsheet} aoClicar={() => exportar('excel')}>
                Excel
              </Botao>
              <Botao variante="perigo" icone={FileDown} aoClicar={() => exportar('pdf')}>
                PDF
              </Botao>
            </>
          }
        >
          {linhas.length === 0 && !carregando ? (
            <EstadoVazio
              icone={ScrollText}
              titulo="Nenhum usuário com permissões nos filtros informados"
              descricao="Ajuste os filtros e pesquise novamente."
            />
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
                  <tr>
                    <th className="px-3 py-2">
                      <input
                        type="checkbox"
                        checked={todasMarcadas}
                        onChange={() =>
                          setSelecionados(todasMarcadas ? new Set() : new Set(linhas.map((l) => l.id)))
                        }
                        aria-label="Selecionar todos os registros"
                        className="h-4 w-4 accent-swnet-600"
                      />
                    </th>
                    <th className="px-3 py-2 font-semibold">#</th>
                    <th className="px-3 py-2 font-semibold">Matrícula</th>
                    <th className="px-3 py-2 font-semibold">Filial</th>
                    <th className="px-3 py-2 font-semibold">Função</th>
                    <th className="px-3 py-2 font-semibold">Login</th>
                    <th className="px-3 py-2 font-semibold">Nome</th>
                    <th className="px-3 py-2 font-semibold">E-mail</th>
                    <th className="px-3 py-2 text-center font-semibold">Serviços</th>
                    <th className="px-3 py-2 text-center font-semibold">Relatórios</th>
                    <th className="px-3 py-2 font-semibold">Cadastro</th>
                    <th className="px-3 py-2 font-semibold">Ações</th>
                  </tr>
                </thead>
                <tbody>
                  {carregando && <LinhaCarregando colunas={12} />}

                  {!carregando &&
                    linhas.map((linha) => (
                      <tr key={linha.id} className="border-b border-slate-100 last:border-0 dark:border-slate-800">
                        <td className="px-3 py-2">
                          <input
                            type="checkbox"
                            checked={selecionados.has(linha.id)}
                            onChange={() => alternarLinha(linha.id)}
                            aria-label={`Selecionar ${linha.nome}`}
                            className="h-4 w-4 accent-swnet-600"
                          />
                        </td>
                        <td className="px-3 py-2 text-slate-400">{linha.id}</td>
                        <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{linha.matricula}</td>
                        <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{linha.nomePlanta}</td>
                        <td className="px-3 py-2 text-xs text-slate-500">{linha.funcao || '—'}</td>
                        <td className="px-3 py-2 font-medium text-slate-700 dark:text-slate-200">{linha.login}</td>
                        <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{linha.nome}</td>
                        <td className="px-3 py-2 text-xs text-slate-500">{linha.email}</td>
                        <td className="px-3 py-2 text-center text-slate-600 dark:text-slate-300">
                          {linha.servicos.length}
                        </td>
                        <td className="px-3 py-2 text-center text-slate-600 dark:text-slate-300">
                          {linha.relatorios.length}
                        </td>
                        <td className="px-3 py-2 text-xs text-slate-500">{formatarData(linha.dataCadastro)}</td>
                        <td className="px-3 py-2">
                          <button
                            type="button"
                            onClick={() => setDetalhe(linha)}
                            className="rounded-lg px-2 py-1 text-xs font-medium text-swnet-600 hover:bg-swnet-50
                              dark:text-swnet-400 dark:hover:bg-slate-800"
                          >
                            Detalhes
                          </button>
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          )}
        </Painel>
      )}

      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Permissões de ${detalhe.nome}` : ''}
        icone={ScrollText}
        largura="lg"
        aoFechar={() => setDetalhe(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setDetalhe(null)}>
            Fechar
          </Botao>
        }
      >
        {detalhe && (
          <div className="grid gap-4 sm:grid-cols-2">
            <ListaAcessos titulo="Serviços" icone={Wrench} itens={detalhe.servicos} />
            <ListaAcessos titulo="Relatórios" icone={BarChart3} itens={detalhe.relatorios} />
          </div>
        )}
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

function ListaAcessos({
  titulo,
  icone: Icone,
  itens
}: {
  titulo: string
  icone: typeof Wrench
  itens: LinhaRelatorioPermissao['servicos']
}) {
  return (
    <div className="rounded-lg border border-slate-200 dark:border-slate-800">
      <header className="flex items-center gap-2 border-b border-slate-200 px-3 py-2 dark:border-slate-800">
        <Icone size={14} className="text-slate-400" />
        <h3 className="text-xs font-semibold uppercase tracking-wide text-slate-600 dark:text-slate-300">
          {titulo}
        </h3>
        <span className="ml-auto rounded-full bg-slate-100 px-2 py-0.5 text-xs text-slate-500 dark:bg-slate-800 dark:text-slate-400">
          {itens.length}
        </span>
      </header>

      {itens.length === 0 ? (
        <p className="px-3 py-6 text-center text-xs text-slate-400">Nenhum acesso concedido.</p>
      ) : (
        <ul className="max-h-72 overflow-y-auto">
          {itens.map((item) => (
            <li
              key={item.codSubcategoria}
              className="border-b border-slate-50 px-3 py-2 last:border-0 dark:border-slate-800/60"
            >
              <span className="block text-sm text-slate-700 dark:text-slate-200">{item.descricao}</span>
              <span className="block text-xs text-slate-400">{item.nomeSubcategoria}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
