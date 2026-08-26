import { useCallback, useEffect, useState } from 'react'
import {
  ClipboardCheck,
  Download,
  Eraser,
  FileDown,
  FileSearch,
  FileSpreadsheet,
  RotateCcw,
  Save,
  Search
} from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  ANALISES,
  FILIAIS_NOTAS,
  classificarNota,
  importarFornecedores,
  importarNotas,
  listarNotas,
  nomeFilial,
  obterNota,
  reabrirNota,
  type FiltrosNotas,
  type NotaInsumo
} from '@/services/compras/notasInsumoApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Compras/Classificar-Notas.
 *
 * Junta numa tela só o que o legado espalha em cinco ASPs:
 * ListaNotasInsumo (grade + filtros), ClassificaNotasInsumo e
 * ReclassificaNotasInsumo (o modal de análises) e as duas rotinas de
 * importação.
 */

const FILTROS_VAZIOS: FiltrosNotas = {
  filial: '',
  razao: '',
  notaFiscal: '',
  serie: '',
  // O legado abre já filtrando as não classificadas
  classificadas: 'N'
}

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

interface Classificacao {
  nota: NotaInsumo
  marcadas: string[]
  documentacaoOk: boolean
  visualOk: boolean
}

export function ClassificarNotas() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [filtros, setFiltros] = useState<FiltrosNotas>(FILTROS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosNotas>(FILTROS_VAZIOS)
  const [notas, setNotas] = useState<NotaInsumo[]>([])
  const [carregando, setCarregando] = useState(true)
  const [importando, setImportando] = useState(false)

  const [classificacao, setClassificacao] = useState<Classificacao | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [reabrindo, setReabrindo] = useState<NotaInsumo | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setNotas(await listarNotas(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosNotas>(campo: C, valor: string) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function limpar() {
    setFiltros(FILTROS_VAZIOS)
    setFiltrosAplicados(FILTROS_VAZIOS)
  }

  async function importar(tipo: 'fornecedores' | 'notas') {
    setImportando(true)
    try {
      const resultado = tipo === 'fornecedores' ? await importarFornecedores() : await importarNotas()
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
      if (tipo === 'notas' && resultado.sucesso) await carregar()
    } finally {
      setImportando(false)
    }
  }

  async function abrirClassificacao(nota: NotaInsumo) {
    setErro('')
    const dados = await obterNota(nota.idNota)
    if (!dados) {
      mostrarToast('Nota não encontrada.', 'danger')
      return
    }

    setClassificacao({
      nota: dados.nota,
      marcadas: dados.marcadas,
      // Reclassificação já abre com as confirmações marcadas
      documentacaoOk: dados.nota.classificada,
      visualOk: dados.nota.classificada
    })
  }

  function alternarAnalise(codigo: string) {
    setClassificacao((atual) =>
      atual
        ? {
            ...atual,
            marcadas: atual.marcadas.includes(codigo)
              ? atual.marcadas.filter((item) => item !== codigo)
              : [...atual.marcadas, codigo]
          }
        : atual
    )
  }

  async function salvar() {
    if (!classificacao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await classificarNota(
        classificacao.nota.idNota,
        classificacao.marcadas,
        classificacao.documentacaoOk,
        classificacao.visualOk
      )

      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }

      setClassificacao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarReabertura() {
    if (!reabrindo) return

    const resultado = await reabrirNota(reabrindo.idNota)
    setReabrindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  function exportar(formato: 'csv' | 'pdf') {
    if (notas.length === 0) {
      mostrarToast('Não há notas para exportar.', 'warning')
      return
    }
    // TODO: o legado gera CSV/PDF no cliente; sem lib de exportação por ora
    mostrarToast(
      `Exportação em ${formato.toUpperCase()} (${notas.length} nota(s)) será implementada em breve.`,
      'info'
    )
  }

  const pontosMarcados = classificacao
    ? ANALISES.filter((a) => classificacao.marcadas.includes(a.codigo)).reduce((s, a) => s + a.peso, 0)
    : 0

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Classificar Notas de Insumo"
        acoes={
          <>
            <Botao
              variante="neutro"
              icone={Download}
              aoClicar={() => importar('fornecedores')}
              desabilitado={importando}
            >
              Importar Fornecedores
            </Botao>
            <Botao
              variante="neutro"
              icone={Download}
              aoClicar={() => importar('notas')}
              desabilitado={importando}
            >
              Importar Notas
            </Botao>
          </>
        }
      />

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Selecao
            id="filtroFilial"
            rotulo="Filial"
            value={filtros.filial}
            onChange={(e) => mudarFiltro('filial', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todas</option>
            {FILIAIS_NOTAS.map((filial) => (
              <option key={filial.codigo} value={filial.codigo}>
                {filial.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroRazao"
            rotulo="Razão Social"
            placeholder="Buscar por razão..."
            value={filtros.razao}
            onChange={(e) => mudarFiltro('razao', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroNotaFiscal"
            rotulo="Nota Fiscal"
            placeholder="Nº da nota"
            value={filtros.notaFiscal}
            onChange={(e) => mudarFiltro('notaFiscal', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroSerie"
            rotulo="Série"
            placeholder="Série"
            value={filtros.serie}
            onChange={(e) => mudarFiltro('serie', e.target.value)}
            larguraClasse="sm:col-span-1"
          />
          <Selecao
            id="filtroTipo"
            rotulo="Tipo"
            value={filtros.classificadas}
            onChange={(e) => mudarFiltro('classificadas', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todas</option>
            <option value="S">Classificadas</option>
            <option value="N">Não Classificadas</option>
          </Selecao>

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Filtrar">
              <span className="sr-only">Filtrar</span>
            </Botao>
            <Botao variante="neutro" icone={Eraser} aoClicar={limpar} titulo="Limpar filtros">
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel
        titulo={`${notas.length} nota(s)`}
        icone={FileSearch}
        acoes={
          <>
            <Botao variante="sucesso" icone={FileSpreadsheet} aoClicar={() => exportar('csv')}>
              CSV
            </Botao>
            <Botao variante="perigo" icone={FileDown} aoClicar={() => exportar('pdf')}>
              PDF
            </Botao>
          </>
        }
      >
        <Tabela
          itens={notas}
          chave={(nota) => nota.idNota}
          carregando={carregando}
          mensagemVazia="Nenhuma nota encontrada para os filtros informados"
          colunas={[
            { cabecalho: 'Planta', conteudo: (n) => nomeFilial(n.filial) },
            { cabecalho: 'Cód. Forn.', conteudo: (n) => n.codigo },
            { cabecalho: 'Razão Social', conteudo: (n) => n.razao },
            { cabecalho: 'Nota Fiscal', conteudo: (n) => n.notaFiscal },
            { cabecalho: 'Série', conteudo: (n) => n.serie },
            { cabecalho: 'Item', conteudo: (n) => n.item },
            { cabecalho: 'Data Lanç.', conteudo: (n) => formatarData(n.dataLancamento) },
            { cabecalho: 'Material', conteudo: (n) => n.material },
            { cabecalho: 'Descrição', conteudo: (n) => n.descricao },
            {
              cabecalho: 'Qtd.',
              conteudo: (n) => `${numero.format(n.quantidade)} ${n.unidade}`,
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Pontos',
              conteudo: (n) => (n.classificada ? `${n.pontos} / ${n.peso}` : '—'),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Situação',
              conteudo: (n) => (
                <Etiqueta
                  texto={n.classificada ? 'Classificada' : 'Pendente'}
                  tom={n.classificada ? 'sucesso' : 'alerta'}
                />
              )
            }
          ]}
          acoes={(nota) => (
            <>
              <BotaoAcao
                icone={ClipboardCheck}
                titulo={nota.classificada ? 'Reclassificar nota' : 'Classificar nota'}
                aoClicar={() => abrirClassificacao(nota)}
              />
              {nota.classificada && (
                <BotaoAcao
                  icone={RotateCcw}
                  titulo="Liberar para reclassificação"
                  tom="perigo"
                  aoClicar={() => setReabrindo(nota)}
                />
              )}
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={classificacao !== null}
        titulo={classificacao?.nota.classificada ? 'Reclassificar Nota' : 'Classificar Nota'}
        icone={ClipboardCheck}
        largura="xl"
        aoFechar={() => setClassificacao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setClassificacao(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {classificacao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <dl className="grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-5 dark:border-slate-700">
              {[
                ['Código', classificacao.nota.codigo],
                ['Razão Social', classificacao.nota.razao],
                ['Nota Fiscal', classificacao.nota.notaFiscal],
                ['Série', classificacao.nota.serie],
                ['Item', classificacao.nota.item]
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            <div className="overflow-hidden rounded-lg border border-slate-200 dark:border-slate-700">
              <table className="w-full text-left text-sm">
                <thead className="border-b border-slate-200 bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-400">
                  <tr>
                    <th className="px-3 py-2 font-semibold">Lançamento</th>
                    <th className="px-3 py-2 font-semibold">Material</th>
                    <th className="px-3 py-2 font-semibold">Descrição</th>
                    <th className="px-3 py-2 text-right font-semibold">Quantidade</th>
                    <th className="px-3 py-2 font-semibold">UM</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="px-3 py-2 text-slate-600 dark:text-slate-300">
                      {formatarData(classificacao.nota.dataLancamento)}
                    </td>
                    <td className="px-3 py-2 text-slate-600 dark:text-slate-300">
                      {classificacao.nota.material}
                    </td>
                    <td className="px-3 py-2 text-slate-600 dark:text-slate-300">
                      {classificacao.nota.descricao}
                    </td>
                    <td className="px-3 py-2 text-right text-slate-600 dark:text-slate-300">
                      {numero.format(classificacao.nota.quantidade)}
                    </td>
                    <td className="px-3 py-2 text-slate-600 dark:text-slate-300">
                      {classificacao.nota.unidade}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {([0, 1] as const).map((tipo) => (
              <Campo key={tipo} rotulo={tipo === 0 ? 'Documentação' : 'Inspeção Visual'}>
                <div className="grid gap-2">
                  {ANALISES.filter((analise) => analise.tipo === tipo).map((analise) => (
                    <label
                      key={analise.codigo}
                      className="flex cursor-pointer items-center gap-2 rounded-lg border border-slate-200 px-3 py-2
                        text-sm text-slate-600 hover:border-swnet-300 dark:border-slate-700 dark:text-slate-300"
                    >
                      <input
                        type="checkbox"
                        checked={classificacao.marcadas.includes(analise.codigo)}
                        disabled={salvando}
                        onChange={() => alternarAnalise(analise.codigo)}
                        className="h-4 w-4 shrink-0 accent-swnet-600"
                      />
                      <span className="flex-1">{analise.descricao}</span>
                      <span className="shrink-0 text-xs text-slate-400">{analise.peso} pts</span>
                    </label>
                  ))}
                </div>
              </Campo>
            ))}

            <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 p-3 dark:border-slate-700">
              <span className="text-sm font-medium text-slate-700 dark:text-slate-200">
                Total: {pontosMarcados} / {classificacao.nota.peso} pontos
              </span>

              <div className="flex flex-wrap gap-4">
                <label className="flex cursor-pointer items-center gap-2 text-sm text-emerald-700 dark:text-emerald-400">
                  <input
                    type="checkbox"
                    checked={classificacao.documentacaoOk}
                    disabled={salvando}
                    onChange={(e) =>
                      setClassificacao({ ...classificacao, documentacaoOk: e.target.checked })
                    }
                    className="h-4 w-4 accent-emerald-600"
                  />
                  Documentação conferida
                </label>

                <label className="flex cursor-pointer items-center gap-2 text-sm text-emerald-700 dark:text-emerald-400">
                  <input
                    type="checkbox"
                    checked={classificacao.visualOk}
                    disabled={salvando}
                    onChange={(e) => setClassificacao({ ...classificacao, visualOk: e.target.checked })}
                    className="h-4 w-4 accent-emerald-600"
                  />
                  Inspeção visual conferida
                </label>
              </div>
            </div>
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={reabrindo !== null}
        titulo="Liberar para reclassificação"
        mensagem={
          reabrindo
            ? `As análises da nota ${reabrindo.notaFiscal} serão apagadas e a nota voltará para pendente. Deseja continuar?`
            : ''
        }
        rotuloConfirmar="Liberar"
        aoConfirmar={confirmarReabertura}
        aoCancelar={() => setReabrindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
