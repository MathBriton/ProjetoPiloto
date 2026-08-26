import { useCallback, useEffect, useMemo, useState } from 'react'
import {
  BarChart3,
  ChevronRight,
  EyeOff,
  Pencil,
  PlusCircle,
  RotateCcw,
  Save,
  Trash2,
  Wrench
} from 'lucide-react'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  criarPermissao,
  editarPermissao,
  excluirPermissao,
  listarMenus,
  reativarPermissao
} from '@/services/administrador/gerenciarMenusApi'
import type {
  CategoriaMenu,
  DadosPermissaoMenu,
  MenusCompletos,
  PermissaoMenu,
  TipoMenu
} from './tipos'

/**
 * Refatoração de Gerenciar-Menus/CriarPermissao.asp.
 *
 * Dois painéis (Serviços e Consultas), cada um com categorias expansíveis e
 * suas permissões. Exclusão é soft delete: a permissão vira "inativa" e só
 * reaparece com o filtro de inativos ligado, onde pode ser reativada.
 */

type EscopoInativos = 'nenhum' | 'todos' | 'servicos' | 'consultas'

const VALORES_MODAL: DadosPermissaoMenu = {
  tipo: 'servico',
  codigoCategoria: '',
  nome: '',
  url: '',
  banco: '',
  extensao: '.asp'
}

export function GerenciarMenus() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [menus, setMenus] = useState<MenusCompletos>({ servicos: [], consultas: [] })
  const [carregando, setCarregando] = useState(true)
  const [filtro, setFiltro] = useState('')
  const [escopoInativos, setEscopoInativos] = useState<EscopoInativos>('nenhum')
  const [abertas, setAbertas] = useState<Set<string>>(new Set())

  const [modal, setModal] = useState<{ modo: 'criar' | 'editar'; dados: DadosPermissaoMenu } | null>(null)
  const [salvandoModal, setSalvandoModal] = useState(false)
  const [erroModal, setErroModal] = useState('')

  const [excluindo, setExcluindo] = useState<{ tipo: TipoMenu; permissao: PermissaoMenu } | null>(null)
  const [reativando, setReativando] = useState<{ tipo: TipoMenu; permissao: PermissaoMenu } | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setMenus(await listarMenus())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  function alternarCategoria(chave: string) {
    setAbertas((atual) => {
      const novo = new Set(atual)
      if (novo.has(chave)) novo.delete(chave)
      else novo.add(chave)
      return novo
    })
  }

  /** Aplica busca textual e o filtro de inativos, descartando categorias vazias */
  const filtrar = useCallback(
    (categorias: CategoriaMenu[], tipo: TipoMenu) => {
      const busca = filtro.trim().toLowerCase()
      const incluirInativos =
        escopoInativos === 'todos' ||
        (escopoInativos === 'servicos' && tipo === 'servico') ||
        (escopoInativos === 'consultas' && tipo === 'consulta')

      return categorias
        .map((categoria) => ({
          ...categoria,
          permissoes: categoria.permissoes.filter((p) => {
            if (!p.ativo && !incluirInativos) return false
            if (!busca) return true
            return p.nome.toLowerCase().includes(busca) || p.url.toLowerCase().includes(busca)
          })
        }))
        .filter((categoria) => categoria.permissoes.length > 0)
    },
    [filtro, escopoInativos]
  )

  const servicosVisiveis = useMemo(() => filtrar(menus.servicos, 'servico'), [menus.servicos, filtrar])
  const consultasVisiveis = useMemo(() => filtrar(menus.consultas, 'consulta'), [menus.consultas, filtrar])

  function abrirCriacao(tipo: TipoMenu, codigoCategoria: string) {
    setErroModal('')
    setModal({ modo: 'criar', dados: { ...VALORES_MODAL, tipo, codigoCategoria } })
  }

  function abrirEdicao(tipo: TipoMenu, categoria: CategoriaMenu, permissao: PermissaoMenu) {
    setErroModal('')
    setModal({
      modo: 'editar',
      dados: {
        tipo,
        codigoCategoria: categoria.codigo,
        codigo: permissao.codigo,
        nome: permissao.nome,
        url: permissao.url,
        banco: permissao.banco ?? '',
        extensao: permissao.extensao
      }
    })
  }

  async function salvarModal() {
    if (!modal) return

    setSalvandoModal(true)
    setErroModal('')
    try {
      const resultado =
        modal.modo === 'criar' ? await criarPermissao(modal.dados) : await editarPermissao(modal.dados)

      if (!resultado.sucesso) {
        setErroModal(resultado.mensagem)
        return
      }

      setModal(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvandoModal(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirPermissao(excluindo.tipo, excluindo.permissao.codigo)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  async function confirmarReativacao() {
    if (!reativando) return
    const resultado = await reativarPermissao(reativando.tipo, reativando.permissao.codigo)
    setReativando(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  const categoriasDoTipoNoModal =
    modal?.dados.tipo === 'consulta' ? menus.consultas : menus.servicos

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Gerenciar Menus"
        acoes={
          <>
            <div className="flex items-center gap-1 rounded-lg border border-slate-200 p-1 dark:border-slate-700">
              <span className="flex items-center gap-1 px-1.5 text-xs font-medium text-slate-500 dark:text-slate-400">
                <EyeOff size={13} /> Inativos
              </span>
              {(
                [
                  ['nenhum', 'Ocultar'],
                  ['todos', 'Todos'],
                  ['servicos', 'Serviços'],
                  ['consultas', 'Consultas']
                ] as [EscopoInativos, string][]
              ).map(([valor, rotulo]) => (
                <button
                  key={valor}
                  type="button"
                  onClick={() => setEscopoInativos(valor)}
                  className={`rounded-md px-2 py-1 text-xs font-medium transition-colors ${
                    escopoInativos === valor
                      ? 'bg-swnet-600 text-white'
                      : 'text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800'
                  }`}
                >
                  {rotulo}
                </button>
              ))}
            </div>

            <CampoPesquisa
              valor={filtro}
              aoMudar={setFiltro}
              placeholder="Pesquisar permissões..."
              larguraClasse="w-64"
            />
          </>
        }
      />

      {carregando ? (
        <p className="py-16 text-center text-sm text-slate-400">Carregando menus...</p>
      ) : (
        <div className="grid gap-4 lg:grid-cols-2">
          <PainelMenu
            titulo="Serviços"
            icone={Wrench}
            tipo="servico"
            categorias={servicosVisiveis}
            abertas={abertas}
            aoAlternarCategoria={alternarCategoria}
            aoCriar={abrirCriacao}
            aoEditar={abrirEdicao}
            aoExcluir={(permissao) => setExcluindo({ tipo: 'servico', permissao })}
            aoReativar={(permissao) => setReativando({ tipo: 'servico', permissao })}
          />
          <PainelMenu
            titulo="Consultas"
            icone={BarChart3}
            tipo="consulta"
            categorias={consultasVisiveis}
            abertas={abertas}
            aoAlternarCategoria={alternarCategoria}
            aoCriar={abrirCriacao}
            aoEditar={abrirEdicao}
            aoExcluir={(permissao) => setExcluindo({ tipo: 'consulta', permissao })}
            aoReativar={(permissao) => setReativando({ tipo: 'consulta', permissao })}
          />
        </div>
      )}

      <Modal
        aberto={modal !== null}
        titulo={modal?.modo === 'criar' ? 'Nova permissão' : 'Editar permissão'}
        icone={modal?.modo === 'criar' ? PlusCircle : Pencil}
        largura="lg"
        aoFechar={() => setModal(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setModal(null)} desabilitado={salvandoModal}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvarModal} desabilitado={salvandoModal}>
              {salvandoModal ? 'Salvando...' : 'Confirmar'}
            </Botao>
          </>
        }
      >
        {modal && (
          <div className="grid gap-4">
            {erroModal && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erroModal}
              </p>
            )}

            <div className="grid gap-4 sm:grid-cols-2">
              <Selecao
                id="tipoPermissao"
                rotulo="Tipo"
                value={modal.dados.tipo}
                disabled={modal.modo === 'editar'}
                onChange={(e) =>
                  setModal({
                    ...modal,
                    dados: { ...modal.dados, tipo: e.target.value as TipoMenu, codigoCategoria: '' }
                  })
                }
              >
                <option value="servico">Serviço</option>
                <option value="consulta">Consulta</option>
              </Selecao>

              <Selecao
                id="categoriaPermissao"
                rotulo="Categoria"
                required
                value={modal.dados.codigoCategoria}
                onChange={(e) =>
                  setModal({ ...modal, dados: { ...modal.dados, codigoCategoria: e.target.value } })
                }
              >
                <option value="">Selecione...</option>
                {categoriasDoTipoNoModal.map((categoria) => (
                  <option key={categoria.codigo} value={categoria.codigo}>
                    {categoria.nome}
                  </option>
                ))}
              </Selecao>
            </div>

            <Entrada
              id="nomePermissao"
              rotulo="Nome da permissão"
              required
              maxLength={50}
              placeholder="Ex: Consulta de Estoque"
              value={modal.dados.nome}
              onChange={(e) => setModal({ ...modal, dados: { ...modal.dados, nome: e.target.value } })}
            />

            <Entrada
              id="urlPermissao"
              rotulo={modal.dados.tipo === 'servico' ? 'URL (LIN_URL)' : 'URL (REL_URL)'}
              required
              placeholder="Ex: ../CriarPermissao/CriarPermissao"
              value={modal.dados.url}
              onChange={(e) => setModal({ ...modal, dados: { ...modal.dados, url: e.target.value } })}
            />

            <div className="grid gap-4 sm:grid-cols-2">
              {modal.dados.tipo === 'consulta' && (
                <Entrada
                  id="bancoPermissao"
                  rotulo="Banco (REL_BANCO)"
                  required
                  maxLength={50}
                  placeholder="Ex: HELPDESK"
                  value={modal.dados.banco}
                  onChange={(e) => setModal({ ...modal, dados: { ...modal.dados, banco: e.target.value } })}
                />
              )}
              <Entrada
                id="extensaoPermissao"
                rotulo="Extensão"
                maxLength={10}
                value={modal.dados.extensao}
                onChange={(e) => setModal({ ...modal, dados: { ...modal.dados, extensao: e.target.value } })}
              />
            </div>
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir permissão"
        mensagem={
          excluindo
            ? `A permissão "${excluindo.permissao.nome}" deixará de aparecer nos menus. Ela pode ser reativada depois pelo filtro de inativos.`
            : ''
        }
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ConfirmDialog
        aberto={reativando !== null}
        titulo="Reativar permissão"
        mensagem={reativando ? `A permissão "${reativando.permissao.nome}" voltará a aparecer nos menus.` : ''}
        rotuloConfirmar="Reativar"
        aoConfirmar={confirmarReativacao}
        aoCancelar={() => setReativando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

interface PainelMenuProps {
  titulo: string
  icone: typeof Wrench
  tipo: TipoMenu
  categorias: CategoriaMenu[]
  abertas: Set<string>
  aoAlternarCategoria: (chave: string) => void
  aoCriar: (tipo: TipoMenu, codigoCategoria: string) => void
  aoEditar: (tipo: TipoMenu, categoria: CategoriaMenu, permissao: PermissaoMenu) => void
  aoExcluir: (permissao: PermissaoMenu) => void
  aoReativar: (permissao: PermissaoMenu) => void
}

function PainelMenu({
  titulo,
  icone: Icone,
  tipo,
  categorias,
  abertas,
  aoAlternarCategoria,
  aoCriar,
  aoEditar,
  aoExcluir,
  aoReativar
}: PainelMenuProps) {
  const total = categorias.reduce((soma, c) => soma + c.permissoes.length, 0)

  return (
    <section className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <header className="flex items-center gap-2 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
        <Icone size={16} className="text-slate-400" />
        <h2 className="text-sm font-semibold text-slate-700 dark:text-slate-200">{titulo}</h2>
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-500 dark:bg-slate-800 dark:text-slate-400">
          {total}
        </span>
      </header>

      <div className="scrollbar-fina max-h-[32rem] overflow-y-auto p-2">
        {categorias.length === 0 && (
          <p className="px-2 py-8 text-center text-sm text-slate-400">Nenhuma permissão encontrada.</p>
        )}

        {categorias.map((categoria) => {
          const chave = `${tipo}:${categoria.codigo}`
          const aberta = abertas.has(chave)

          return (
            <div key={chave} className="mb-1 rounded-lg border border-slate-100 last:mb-0 dark:border-slate-800">
              <div className="flex items-center gap-2 px-2 py-2">
                <button
                  type="button"
                  onClick={() => aoAlternarCategoria(chave)}
                  aria-expanded={aberta}
                  className="flex min-w-0 flex-1 items-center gap-2 text-left"
                >
                  <ChevronRight
                    size={13}
                    className={`shrink-0 text-slate-400 transition-transform ${aberta ? 'rotate-90' : ''}`}
                  />
                  <span className="truncate text-sm font-medium text-slate-700 dark:text-slate-200">
                    {categoria.nome}
                  </span>
                  <span className="shrink-0 rounded-full bg-slate-100 px-1.5 py-0.5 text-[11px] text-slate-500 dark:bg-slate-800 dark:text-slate-400">
                    {categoria.permissoes.length}
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => aoCriar(tipo, categoria.codigo)}
                  title="Nova permissão"
                  aria-label={`Nova permissão em ${categoria.nome}`}
                  className="shrink-0 rounded-lg p-1.5 text-swnet-600 hover:bg-swnet-50 dark:text-swnet-400 dark:hover:bg-slate-800"
                >
                  <PlusCircle size={16} />
                </button>
              </div>

              {aberta && (
                <ul className="border-t border-slate-100 dark:border-slate-800">
                  {categoria.permissoes.map((permissao) => (
                    <li
                      key={permissao.codigo}
                      className="flex items-center gap-2 border-b border-slate-50 px-3 py-2 last:border-0 dark:border-slate-800/60"
                    >
                      <span className="min-w-0 flex-1">
                        <span className="flex items-center gap-2">
                          <span className="truncate text-sm text-slate-600 dark:text-slate-300">
                            {permissao.nome}
                          </span>
                          {!permissao.ativo && <Etiqueta texto="Inativa" tom="perigo" />}
                        </span>
                        <span className="block truncate text-xs text-slate-400">
                          {permissao.url}
                          {permissao.extensao}
                          {permissao.banco && ` • ${permissao.banco}`}
                        </span>
                      </span>

                      {permissao.ativo ? (
                        <>
                          <button
                            type="button"
                            onClick={() => aoEditar(tipo, categoria, permissao)}
                            title="Editar"
                            aria-label={`Editar ${permissao.nome}`}
                            className="shrink-0 rounded-lg p-1.5 text-slate-400 hover:bg-slate-100 hover:text-swnet-600 dark:hover:bg-slate-800"
                          >
                            <Pencil size={14} />
                          </button>
                          <button
                            type="button"
                            onClick={() => aoExcluir(permissao)}
                            title="Excluir"
                            aria-label={`Excluir ${permissao.nome}`}
                            className="shrink-0 rounded-lg p-1.5 text-slate-400 hover:bg-red-50 hover:text-red-600 dark:hover:bg-slate-800"
                          >
                            <Trash2 size={14} />
                          </button>
                        </>
                      ) : (
                        <button
                          type="button"
                          onClick={() => aoReativar(permissao)}
                          title="Reativar"
                          aria-label={`Reativar ${permissao.nome}`}
                          className="shrink-0 rounded-lg p-1.5 text-slate-400 hover:bg-emerald-50 hover:text-emerald-600 dark:hover:bg-slate-800"
                        >
                          <RotateCcw size={14} />
                        </button>
                      )}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          )
        })}
      </div>
    </section>
  )
}
