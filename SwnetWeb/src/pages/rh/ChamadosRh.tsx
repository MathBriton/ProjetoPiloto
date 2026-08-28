import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  CheckCircle2,
  Eraser,
  Headset,
  MessageSquare,
  RotateCcw,
  Search,
  Send,
  Star,
  UserPlus
} from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  FILTROS_CHAMADO_VAZIOS,
  STATUS_CHAMADO,
  alterarCampoChamado,
  atribuirAtendente,
  avaliarChamado,
  descricaoNivel,
  descricaoStatusChamado,
  enviarMensagem,
  finalizarChamado,
  listarAtendentes,
  listarCategorias,
  listarChamados,
  listarDepartamentos,
  listarPrioridades,
  motivosPorNota,
  obterKpis,
  reabrirChamado,
  type AtendenteRh,
  type CategoriaChamado,
  type Chamado,
  type DepartamentoChamado,
  type EscopoChamados,
  type FiltrosChamado,
  type KpiChamados,
  type PrioridadeChamado
} from '@/services/rh/rhApi'
import { PLANTAS_CREDITO, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de RH/Chamados-RH e RH/Meus-Chamados-RH — no legado as duas
 * pastas têm os mesmos arquivos; o que muda é a grade usada
 * (ChamadoGradeAtendente vs ChamadoGradeSolicitante).
 */

const TOM_STATUS: Record<string, 'neutro' | 'info' | 'alerta' | 'sucesso' | 'perigo'> = {
  ABERTO: 'info',
  AGUARDANDO_SOLICITANTE: 'alerta',
  REABERTO: 'alerta',
  FINALIZADO: 'sucesso',
  CANCELADO: 'perigo'
}

interface ChamadosRhProps {
  escopo: EscopoChamados
  titulo: string
}

export function ChamadosRh({ escopo, titulo }: ChamadosRhProps) {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const ehFila = escopo === 'FILA'
  const login = usuario?.login ?? ''

  const [filtros, setFiltros] = useState<FiltrosChamado>(FILTROS_CHAMADO_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_CHAMADO_VAZIOS)
  const [chamados, setChamados] = useState<Chamado[]>([])
  const [kpis, setKpis] = useState<KpiChamados | null>(null)
  const [carregando, setCarregando] = useState(true)

  const [departamentos, setDepartamentos] = useState<DepartamentoChamado[]>([])
  const [categorias, setCategorias] = useState<CategoriaChamado[]>([])
  const [prioridades, setPrioridades] = useState<PrioridadeChamado[]>([])
  const [atendentes, setAtendentes] = useState<AtendenteRh[]>([])

  const [aberto, setAberto] = useState<Chamado | null>(null)
  const [novaMensagem, setNovaMensagem] = useState('')
  const [finalizando, setFinalizando] = useState<Chamado | null>(null)
  const [mensagemFinal, setMensagemFinal] = useState('')
  const [reabrindo, setReabrindo] = useState<Chamado | null>(null)
  const [motivoReabertura, setMotivoReabertura] = useState('')
  const [avaliando, setAvaliando] = useState<Chamado | null>(null)
  const [avaliacao, setAvaliacao] = useState({ nota: 0, motivos: [] as string[], comentario: '' })
  const [atribuindo, setAtribuindo] = useState<Chamado | null>(null)
  const [atendenteEscolhido, setAtendenteEscolhido] = useState('')
  const [processando, setProcessando] = useState(false)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, listaDep, listaCat, listaPri, listaAtend, indicadores] = await Promise.all([
        listarChamados(escopo, filtrosAplicados, login),
        listarDepartamentos(),
        listarCategorias(),
        listarPrioridades(),
        listarAtendentes(),
        ehFila ? obterKpis(filtrosAplicados.planta) : Promise.resolve(null)
      ])
      setChamados(lista)
      setDepartamentos(listaDep)
      setCategorias(listaCat)
      setPrioridades(listaPri)
      setAtendentes(listaAtend)
      setKpis(indicadores)
    } finally {
      setCarregando(false)
    }
  }, [escopo, filtrosAplicados, login, ehFila])

  useEffect(() => {
    carregar()
  }, [carregar])

  // Mantém o chamado aberto em sincronia após cada ação
  useEffect(() => {
    if (!aberto) return
    const atualizado = chamados.find((c) => c.id === aberto.id)
    if (atualizado) setAberto(atualizado)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chamados])

  function nomeDepartamento(id: number) {
    return departamentos.find((d) => d.id === id)?.nome ?? '—'
  }

  function nomeCategoria(id: number | null) {
    return id === null ? '—' : (categorias.find((c) => c.id === id)?.nome ?? '—')
  }

  function nivelPrioridade(id: number) {
    return prioridades.find((p) => p.id === id)?.nivel ?? ''
  }

  async function executar(acao: () => Promise<{ sucesso: boolean; mensagem: string }>) {
    setProcessando(true)
    try {
      const resultado = await acao()
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
      if (resultado.sucesso) await carregar()
      return resultado.sucesso
    } finally {
      setProcessando(false)
    }
  }

  async function responder() {
    if (!aberto) return
    const ok = await executar(() => enviarMensagem(aberto.id, novaMensagem, login, ehFila))
    if (ok) setNovaMensagem('')
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo={titulo} />

      {ehFila && kpis && (
        <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-5">
          {[
            { rotulo: 'Abertos', valor: kpis.abertos, tom: 'text-swnet-600 dark:text-swnet-400' },
            { rotulo: 'Aguardando solicitante', valor: kpis.aguardando, tom: 'text-amber-600 dark:text-amber-400' },
            { rotulo: 'Atrasados', valor: kpis.atrasados, tom: 'text-red-600 dark:text-red-400' },
            { rotulo: 'Finalizados', valor: kpis.finalizados, tom: 'text-emerald-600 dark:text-emerald-400' },
            { rotulo: 'Nota média', valor: kpis.notaMedia || '—', tom: 'text-slate-700 dark:text-slate-200' }
          ].map((kpi) => (
            <div
              key={kpi.rotulo}
              className="rounded-xl border border-slate-200 bg-white px-4 py-3 shadow-sm dark:border-slate-800 dark:bg-slate-900"
            >
              <span className="block text-xs font-medium uppercase tracking-wide text-slate-500 dark:text-slate-400">
                {kpi.rotulo}
              </span>
              <span className={`mt-1 block text-2xl font-bold ${kpi.tom}`}>{kpi.valor}</span>
            </div>
          ))}
        </div>
      )}

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Selecao
            id="filtroPlantaChamado"
            rotulo="Filial"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todas</option>
            {PLANTAS_CREDITO.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroStatusChamado"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todos</option>
            {STATUS_CHAMADO.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroDepChamado"
            rotulo="Departamento"
            value={filtros.idDepartamento}
            onChange={(e) => setFiltros({ ...filtros, idDepartamento: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todos</option>
            {departamentos.map((item) => (
              <option key={item.id} value={item.id}>
                {item.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroTermoChamado"
            rotulo="Assunto, nº ou solicitante"
            value={filtros.termo}
            onChange={(e) => setFiltros({ ...filtros, termo: e.target.value })}
            larguraClasse="sm:col-span-2"
          />

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Pesquisar">
              <span className="sr-only">Pesquisar</span>
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              titulo="Limpar filtros"
              aoClicar={() => {
                setFiltros(FILTROS_CHAMADO_VAZIOS)
                setFiltrosAplicados(FILTROS_CHAMADO_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${chamados.length} chamado(s)`} icone={Headset}>
        <Tabela
          itens={chamados}
          chave={(item) => item.id}
          carregando={carregando}
          mensagemVazia="Nenhum chamado encontrado"
          colunas={[
            { cabecalho: '#', conteudo: (c) => c.id, className: 'w-20' },
            { cabecalho: 'Filial', conteudo: (c) => nomePlanta(c.planta) },
            { cabecalho: 'Assunto', conteudo: (c) => c.assunto },
            ...(ehFila
              ? [{ cabecalho: 'Solicitante', conteudo: (c: Chamado) => c.nomeSolicitante }]
              : []),
            { cabecalho: 'Atendente', conteudo: (c) => c.atendente || 'não atribuído' },
            { cabecalho: 'Departamento', conteudo: (c) => nomeDepartamento(c.idDepartamento) },
            { cabecalho: 'Categoria', conteudo: (c) => nomeCategoria(c.idCategoria) },
            {
              cabecalho: 'Prioridade',
              conteudo: (c) => {
                const nivel = nivelPrioridade(c.idPrioridade)
                return (
                  <Etiqueta
                    texto={descricaoNivel(nivel)}
                    tom={nivel === 'URGENTE' ? 'perigo' : nivel === 'ALTA' ? 'alerta' : 'neutro'}
                  />
                )
              }
            },
            {
              cabecalho: 'Status',
              conteudo: (c) => (
                <Etiqueta texto={descricaoStatusChamado(c.status)} tom={TOM_STATUS[c.status] ?? 'neutro'} />
              )
            },
            { cabecalho: 'Data Criação', conteudo: (c) => formatarData(c.dataCriacao) },
            ...(ehFila
              ? [
                  {
                    cabecalho: 'Data Prevista',
                    conteudo: (c: Chamado) => {
                      const atrasado =
                        c.status !== 'FINALIZADO' &&
                        c.status !== 'CANCELADO' &&
                        c.dataPrevista < new Date().toISOString().slice(0, 10)
                      return (
                        <span className={atrasado ? 'font-semibold text-red-600 dark:text-red-400' : ''}>
                          {formatarData(c.dataPrevista)}
                        </span>
                      )
                    }
                  }
                ]
              : []),
            {
              cabecalho: 'Data Finalização',
              conteudo: (c) => (c.dataFinalizacao ? formatarData(c.dataFinalizacao) : '—')
            }
          ]}
          acoes={(item) => {
            const encerrado = item.status === 'FINALIZADO' || item.status === 'CANCELADO'

            return (
              <>
                <BotaoAcao
                  icone={MessageSquare}
                  titulo="Abrir conversa"
                  aoClicar={() => {
                    setNovaMensagem('')
                    setAberto(item)
                  }}
                />

                {ehFila && !encerrado && (
                  <>
                    <BotaoAcao
                      icone={UserPlus}
                      titulo="Atribuir atendente"
                      aoClicar={() => {
                        setAtendenteEscolhido(item.atendente)
                        setAtribuindo(item)
                      }}
                    />
                    <BotaoAcao
                      icone={CheckCircle2}
                      titulo="Finalizar chamado"
                      aoClicar={() => {
                        setMensagemFinal('')
                        setFinalizando(item)
                      }}
                    />
                  </>
                )}

                {!ehFila && item.status === 'FINALIZADO' && (
                  <>
                    {!item.avaliacao && (
                      <BotaoAcao
                        icone={Star}
                        titulo="Avaliar atendimento"
                        aoClicar={() => {
                          setAvaliacao({ nota: 0, motivos: [], comentario: '' })
                          setAvaliando(item)
                        }}
                      />
                    )}
                    <BotaoAcao
                      icone={RotateCcw}
                      titulo="Reabrir chamado"
                      aoClicar={() => {
                        setMotivoReabertura('')
                        setReabrindo(item)
                      }}
                    />
                  </>
                )}

                {ehFila && !encerrado && (
                  <BotaoAcao
                    icone={Ban}
                    titulo="Cancelar chamado"
                    tom="perigo"
                    aoClicar={() =>
                      executar(() => alterarCampoChamado(item.id, 'status', 'CANCELADO'))
                    }
                  />
                )}
              </>
            )
          }}
        />
      </Painel>

      {/* Conversa */}
      <Modal
        aberto={aberto !== null}
        titulo={aberto ? `Chamado ${aberto.id} — ${aberto.assunto}` : ''}
        icone={MessageSquare}
        largura="lg"
        aoFechar={() => setAberto(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setAberto(null)}>
            Fechar
          </Botao>
        }
      >
        {aberto && (
          <div className="grid gap-4">
            <dl className="grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-4 dark:border-slate-700">
              {[
                ['Solicitante', aberto.nomeSolicitante],
                ['Departamento', nomeDepartamento(aberto.idDepartamento)],
                ['Categoria', nomeCategoria(aberto.idCategoria)],
                ['Status', descricaoStatusChamado(aberto.status)],
                ['Tipo', aberto.tipo],
                ['Centro de Custo', aberto.centroCusto || '—'],
                ['Data Prevista', formatarData(aberto.dataPrevista)],
                ['Atendente', aberto.atendente || 'não atribuído']
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            <p className="rounded-lg bg-slate-50 px-3 py-2 text-sm text-slate-600 dark:bg-slate-800 dark:text-slate-300">
              {aberto.descricao}
            </p>

            {/* Ajustes rápidos do atendente */}
            {ehFila && aberto.status !== 'FINALIZADO' && aberto.status !== 'CANCELADO' && (
              <div className="grid gap-3 sm:grid-cols-3">
                <Selecao
                  id="mudarPrioridade"
                  rotulo="Prioridade"
                  value={String(aberto.idPrioridade)}
                  disabled={processando}
                  onChange={(e) => executar(() => alterarCampoChamado(aberto.id, 'idPrioridade', e.target.value))}
                >
                  {prioridades.map((item) => (
                    <option key={item.id} value={item.id}>
                      {descricaoNivel(item.nivel)} ({item.prazoHoras}h)
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="mudarDepartamento"
                  rotulo="Departamento"
                  value={String(aberto.idDepartamento)}
                  disabled={processando}
                  onChange={(e) =>
                    executar(() => alterarCampoChamado(aberto.id, 'idDepartamento', e.target.value))
                  }
                >
                  {departamentos.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="mudarCategoria"
                  rotulo="Categoria"
                  value={String(aberto.idCategoria ?? '')}
                  disabled={processando}
                  onChange={(e) => executar(() => alterarCampoChamado(aberto.id, 'idCategoria', e.target.value))}
                >
                  <option value="">Sem categoria</option>
                  {categorias.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>
              </div>
            )}

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Conversa
              </p>
              <div className="grid max-h-72 gap-2 overflow-y-auto">
                {aberto.mensagens.length === 0 && (
                  <p className="rounded-lg border border-dashed border-slate-200 px-3 py-6 text-center text-xs text-slate-400 dark:border-slate-700">
                    Nenhuma mensagem ainda.
                  </p>
                )}

                {aberto.mensagens.map((mensagem, indice) => (
                  <div
                    key={indice}
                    className={`max-w-[85%] rounded-lg px-3 py-2 text-sm ${
                      mensagem.ehAtendente
                        ? 'ml-auto bg-swnet-50 text-slate-700 dark:bg-swnet-950/40 dark:text-slate-200'
                        : 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-200'
                    }`}
                  >
                    <span className="mb-0.5 block text-xs font-semibold text-slate-500 dark:text-slate-400">
                      {mensagem.autor} &bull; {mensagem.dataHora}
                    </span>
                    {mensagem.texto}
                  </div>
                ))}
              </div>
            </div>

            {aberto.avaliacao && (
              <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-3 dark:border-emerald-900 dark:bg-emerald-500/10">
                <p className="text-sm font-medium text-emerald-800 dark:text-emerald-300">
                  Avaliação: {aberto.avaliacao.nota} / 5
                </p>
                {aberto.avaliacao.motivos.length > 0 && (
                  <p className="mt-1 text-xs text-emerald-700 dark:text-emerald-400">
                    {aberto.avaliacao.motivos.join(' • ')}
                  </p>
                )}
                {aberto.avaliacao.comentario && (
                  <p className="mt-1 text-xs text-emerald-700 dark:text-emerald-400">
                    &ldquo;{aberto.avaliacao.comentario}&rdquo;
                  </p>
                )}
              </div>
            )}

            {aberto.status !== 'FINALIZADO' && aberto.status !== 'CANCELADO' && (
              <div className="flex items-end gap-2">
                <Campo rotulo="Nova mensagem" htmlFor="novaMensagem" className="flex-1">
                  <textarea
                    id="novaMensagem"
                    rows={2}
                    value={novaMensagem}
                    disabled={processando}
                    onChange={(e) => setNovaMensagem(e.target.value)}
                    className={classesControle}
                  />
                </Campo>
                <Botao variante="sucesso" icone={Send} aoClicar={responder} desabilitado={processando}>
                  Enviar
                </Botao>
              </div>
            )}
          </div>
        )}
      </Modal>

      {/* Atribuir atendente */}
      <Modal
        aberto={atribuindo !== null}
        titulo="Atribuir atendente"
        icone={UserPlus}
        aoFechar={() => setAtribuindo(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAtribuindo(null)}>
              Cancelar
            </Botao>
            <Botao
              variante="sucesso"
              aoClicar={async () => {
                if (!atribuindo) return
                const ok = await executar(() => atribuirAtendente(atribuindo.id, atendenteEscolhido))
                if (ok) setAtribuindo(null)
              }}
              desabilitado={processando}
            >
              Atribuir
            </Botao>
          </>
        }
      >
        <Selecao
          id="atendenteEscolhido"
          rotulo="Atendente"
          required
          value={atendenteEscolhido}
          onChange={(e) => setAtendenteEscolhido(e.target.value)}
        >
          <option value="">Selecione...</option>
          {atendentes.map((item) => (
            <option key={item.id} value={item.usuario}>
              {item.nome} ({item.usuario})
            </option>
          ))}
        </Selecao>
      </Modal>

      {/* Finalizar */}
      <Modal
        aberto={finalizando !== null}
        titulo="Finalizar chamado"
        icone={CheckCircle2}
        aoFechar={() => setFinalizando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setFinalizando(null)}>
              Cancelar
            </Botao>
            <Botao
              variante="sucesso"
              aoClicar={async () => {
                if (!finalizando) return
                const ok = await executar(() => finalizarChamado(finalizando.id, mensagemFinal, login))
                if (ok) setFinalizando(null)
              }}
              desabilitado={processando}
            >
              Finalizar
            </Botao>
          </>
        }
      >
        <Campo rotulo="Mensagem de encerramento" htmlFor="mensagemFinal">
          <textarea
            id="mensagemFinal"
            rows={3}
            value={mensagemFinal}
            placeholder="Opcional — será enviada ao solicitante."
            onChange={(e) => setMensagemFinal(e.target.value)}
            className={classesControle}
          />
        </Campo>
      </Modal>

      {/* Reabrir */}
      <Modal
        aberto={reabrindo !== null}
        titulo="Reabrir chamado"
        icone={RotateCcw}
        aoFechar={() => setReabrindo(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setReabrindo(null)}>
              Cancelar
            </Botao>
            <Botao
              aoClicar={async () => {
                if (!reabrindo) return
                const ok = await executar(() => reabrirChamado(reabrindo.id, motivoReabertura, login))
                if (ok) setReabrindo(null)
              }}
              desabilitado={processando}
            >
              Reabrir
            </Botao>
          </>
        }
      >
        <div className="grid gap-3">
          <p className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800 dark:bg-amber-500/10 dark:text-amber-400">
            Reabrir remove a avaliação registrada.
          </p>
          <Campo rotulo="Motivo da reabertura" htmlFor="motivoReabertura" obrigatorio>
            <textarea
              id="motivoReabertura"
              rows={3}
              value={motivoReabertura}
              onChange={(e) => setMotivoReabertura(e.target.value)}
              className={classesControle}
            />
          </Campo>
        </div>
      </Modal>

      {/* Avaliação */}
      <Modal
        aberto={avaliando !== null}
        titulo="Avaliar atendimento"
        icone={Star}
        aoFechar={() => setAvaliando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAvaliando(null)}>
              Cancelar
            </Botao>
            <Botao
              variante="sucesso"
              aoClicar={async () => {
                if (!avaliando) return
                const ok = await executar(() =>
                  avaliarChamado(avaliando.id, avaliacao.nota, avaliacao.motivos, avaliacao.comentario)
                )
                if (ok) setAvaliando(null)
              }}
              desabilitado={processando || avaliacao.nota === 0}
            >
              Enviar avaliação
            </Botao>
          </>
        }
      >
        <div className="grid gap-4">
          <Campo rotulo="Sua nota" obrigatorio>
            <div className="flex gap-2">
              {[1, 2, 3, 4, 5].map((nota) => (
                <button
                  key={nota}
                  type="button"
                  onClick={() => setAvaliacao({ nota, motivos: [], comentario: avaliacao.comentario })}
                  aria-label={`Nota ${nota}`}
                  className={`flex h-11 w-11 items-center justify-center rounded-lg border text-sm font-semibold
                    transition-colors ${
                      avaliacao.nota >= nota
                        ? 'border-amber-400 bg-amber-50 text-amber-700 dark:border-amber-600 dark:bg-amber-500/10 dark:text-amber-400'
                        : 'border-slate-200 text-slate-400 hover:border-amber-300 dark:border-slate-700'
                    }`}
                >
                  <Star size={18} fill={avaliacao.nota >= nota ? 'currentColor' : 'none'} />
                </button>
              ))}
            </div>
          </Campo>

          {avaliacao.nota > 0 && (
            <>
              <p className="text-sm text-slate-600 dark:text-slate-300">
                {avaliacao.nota >= 4
                  ? 'Excelente! O que mais pesou?'
                  : 'Obrigado — o que podemos melhorar?'}
              </p>

              <div className="grid gap-2 sm:grid-cols-2">
                {motivosPorNota(avaliacao.nota).map((motivo) => {
                  const marcado = avaliacao.motivos.includes(motivo)
                  return (
                    <label
                      key={motivo}
                      className="flex cursor-pointer items-center gap-2 rounded-lg border border-slate-200 px-3 py-2
                        text-sm text-slate-600 dark:border-slate-700 dark:text-slate-300"
                    >
                      <input
                        type="checkbox"
                        checked={marcado}
                        onChange={() =>
                          setAvaliacao({
                            ...avaliacao,
                            motivos: marcado
                              ? avaliacao.motivos.filter((m) => m !== motivo)
                              : [...avaliacao.motivos, motivo]
                          })
                        }
                        className="h-4 w-4 accent-swnet-600"
                      />
                      {motivo}
                    </label>
                  )
                })}
              </div>

              <Campo rotulo="Comentário" htmlFor="comentarioAvaliacao">
                <textarea
                  id="comentarioAvaliacao"
                  rows={2}
                  value={avaliacao.comentario}
                  onChange={(e) => setAvaliacao({ ...avaliacao, comentario: e.target.value })}
                  className={classesControle}
                />
              </Campo>
            </>
          )}
        </div>
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

/** Ponto de entrada "Chamados RH" — a fila do atendente */
export function FilaChamadosRh() {
  return <ChamadosRh escopo="FILA" titulo="Chamados RH" />
}

/** Ponto de entrada "Meus Chamados RH" — a visão do solicitante */
export function MeusChamadosRh() {
  return <ChamadosRh escopo="MEUS" titulo="Meus Chamados RH" />
}
