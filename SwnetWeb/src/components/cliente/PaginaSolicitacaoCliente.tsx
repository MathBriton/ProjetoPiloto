import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  Contact,
  Eraser,
  FileSearch,
  Paperclip,
  Pencil,
  Plus,
  Save,
  Search,
  Send,
  ThumbsDown,
  ThumbsUp
} from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  CAMPOS_ALTERACAO,
  DEPARTAMENTOS,
  FILTROS_SOLICITACAO_VAZIOS,
  STATUS_ALTERACAO,
  STATUS_EXTENSAO,
  TIPOS_CONTRIBUINTE,
  avaliarSolicitacaoCliente,
  cancelarSolicitacaoCliente,
  departamentoDaEtapa,
  descricaoStatusSolicitacao,
  enviarSolicitacaoCliente,
  listarSolicitacoesCliente,
  salvarSolicitacaoCliente,
  solicitacaoVazia,
  type DadosSolicitacaoCliente,
  type Departamento,
  type FiltrosSolicitacaoCliente,
  type SolicitacaoCliente,
  type TipoSolicitacaoCliente
} from '@/services/cliente/clienteApi'
import { PLANTAS_PROTHEUS, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Tela compartilhada por Alterar-Dados-Do-Cliente e
 * Transferencia-Extensao-De-Cliente.
 *
 * No legado são duas pastas quase idênticas (Grade + FormularioGeral +
 * FormularioAvaliacao + Repositorio); o que muda é o fluxo de aprovação e
 * dois blocos de campo — a alteração lista os campos a mudar, a extensão
 * pede planta de destino e documentação.
 */

interface PaginaSolicitacaoClienteProps {
  tipo: TipoSolicitacaoCliente
  titulo: string
}

export function PaginaSolicitacaoCliente({ tipo, titulo }: PaginaSolicitacaoClienteProps) {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const fluxo = tipo === 'ALTERACAO' ? STATUS_ALTERACAO : STATUS_EXTENSAO

  /** Departamento do usuário — viria das colunas de nível da CRM_DTI */
  const [departamento, setDepartamento] = useState<Departamento>('SOLICITANTE')

  const [filtros, setFiltros] = useState<FiltrosSolicitacaoCliente>(FILTROS_SOLICITACAO_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_SOLICITACAO_VAZIOS)
  const [solicitacoes, setSolicitacoes] = useState<SolicitacaoCliente[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<DadosSolicitacaoCliente | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const [detalhe, setDetalhe] = useState<SolicitacaoCliente | null>(null)
  const [avaliando, setAvaliando] = useState<{ item: SolicitacaoCliente; aprovar: boolean } | null>(null)
  const [observacao, setObservacao] = useState('')
  const [cancelando, setCancelando] = useState<SolicitacaoCliente | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setSolicitacoes(await listarSolicitacoesCliente(tipo, filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [tipo, filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof DadosSolicitacaoCliente>(
    campo: C,
    valor: DadosSolicitacaoCliente[C]
  ) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function abrirEdicao(item: SolicitacaoCliente) {
    setErro('')
    setEdicao({
      id: item.id,
      tipo: item.tipo,
      planta: item.planta,
      plantaDestino: item.plantaDestino,
      codCliente: item.codCliente,
      codLoja: item.codLoja,
      razaoSocial: item.razaoSocial,
      cnpj: item.cnpj,
      justificativa: item.justificativa,
      camposAlterados: [...item.camposAlterados],
      documentos: [...item.documentos],
      tipoContribuinte: item.tipoContribuinte
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarSolicitacaoCliente(edicao, usuario?.login ?? '')
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function enviar(item: SolicitacaoCliente) {
    const resultado = await enviarSolicitacaoCliente(item.id)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'warning')
    if (resultado.sucesso) await carregar()
  }

  async function confirmarAvaliacao() {
    if (!avaliando) return

    const resultado = await avaliarSolicitacaoCliente(
      avaliando.item.id,
      avaliando.aprovar,
      observacao,
      usuario?.login ?? ''
    )
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setAvaliando(null)
      setObservacao('')
      await carregar()
    }
  }

  async function confirmarCancelamento() {
    if (!cancelando) return

    const resultado = await cancelarSolicitacaoCliente(cancelando.id)
    setCancelando(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  /* Campos alterados — só na solicitação de alteração */
  function adicionarCampo() {
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            camposAlterados: [...atual.camposAlterados, { campo: '', valorAtual: '', valorNovo: '' }]
          }
        : atual
    )
  }

  function mudarCampo(indice: number, chave: 'campo' | 'valorAtual' | 'valorNovo', valor: string) {
    setEdicao((atual) => {
      if (!atual) return atual
      const campos = [...atual.camposAlterados]
      campos[indice] = { ...campos[indice], [chave]: valor }
      return { ...atual, camposAlterados: campos }
    })
  }

  function removerCampo(indice: number) {
    setEdicao((atual) =>
      atual
        ? { ...atual, camposAlterados: atual.camposAlterados.filter((_, i) => i !== indice) }
        : atual
    )
  }

  function adicionarDocumento(nome: string) {
    if (!nome.trim()) return
    setEdicao((atual) =>
      atual ? { ...atual, documentos: [...atual.documentos, nome.trim()] } : atual
    )
  }

  const ehExtensao = tipo === 'EXTENSAO'

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo={titulo}
        acoes={
          <>
            <Selecao
              id="departamentoUsuario"
              rotulo=""
              value={departamento}
              onChange={(e) => setDepartamento(e.target.value as Departamento)}
              larguraClasse="w-56"
            >
              {DEPARTAMENTOS.map((item) => (
                <option key={item.codigo} value={item.codigo}>
                  Área: {item.descricao}
                </option>
              ))}
            </Selecao>

            {departamento === 'SOLICITANTE' && (
              <Botao
                variante="sucesso"
                icone={Plus}
                aoClicar={() => {
                  setErro('')
                  setEdicao(solicitacaoVazia(tipo))
                }}
              >
                Nova solicitação
              </Botao>
            )}
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
            id="filtroPlantaSol"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todas as plantas</option>
            {PLANTAS_PROTHEUS.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroStatusSol"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Todos</option>
            {fluxo.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroClienteSol"
            rotulo="Cliente ou CNPJ"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-3"
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
                setFiltros(FILTROS_SOLICITACAO_VAZIOS)
                setFiltrosAplicados(FILTROS_SOLICITACAO_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${solicitacoes.length} solicitação(ões)`} icone={Contact}>
        <Tabela
          itens={solicitacoes}
          chave={(item) => item.id}
          carregando={carregando}
          mensagemVazia="Nenhuma solicitação encontrada"
          colunas={[
            { cabecalho: '#', conteudo: (s) => s.id, className: 'w-20' },
            { cabecalho: 'Data Inclusão', conteudo: (s) => formatarData(s.dataInclusao) },
            { cabecalho: 'Planta', conteudo: (s) => nomePlanta(s.planta) },
            ...(ehExtensao
              ? [
                  {
                    cabecalho: 'Planta Destino',
                    conteudo: (s: SolicitacaoCliente) => nomePlanta(s.plantaDestino)
                  }
                ]
              : []),
            { cabecalho: 'Cliente', conteudo: (s) => s.razaoSocial },
            { cabecalho: 'CNPJ', conteudo: (s) => s.cnpj },
            {
              cabecalho: 'Status',
              conteudo: (s) => (
                <Etiqueta
                  texto={descricaoStatusSolicitacao(tipo, s.status)}
                  tom={
                    s.status === 'FINALIZADO'
                      ? 'sucesso'
                      : s.status === 'CANCELADO'
                        ? 'perigo'
                        : s.status === 'EM_DIGITACAO'
                          ? 'neutro'
                          : 'alerta'
                  }
                />
              )
            },
            { cabecalho: 'Emitente', conteudo: (s) => s.emitente }
          ]}
          acoes={(item) => {
            const emDigitacao = item.status === 'EM_DIGITACAO'
            const podeAvaliar = departamentoDaEtapa(tipo, item.status) === departamento && !emDigitacao
            const encerrada = item.status === 'FINALIZADO' || item.status === 'CANCELADO'

            return (
              <>
                <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(item)} />

                {emDigitacao && departamento === 'SOLICITANTE' && (
                  <>
                    <BotaoAcao icone={Pencil} titulo="Editar" aoClicar={() => abrirEdicao(item)} />
                    <BotaoAcao icone={Send} titulo="Enviar para avaliação" aoClicar={() => enviar(item)} />
                  </>
                )}

                {podeAvaliar && (
                  <>
                    <BotaoAcao
                      icone={ThumbsUp}
                      titulo="Aprovar etapa"
                      aoClicar={() => {
                        setObservacao('')
                        setAvaliando({ item, aprovar: true })
                      }}
                    />
                    <BotaoAcao
                      icone={ThumbsDown}
                      titulo="Reprovar"
                      tom="perigo"
                      aoClicar={() => {
                        setObservacao('')
                        setAvaliando({ item, aprovar: false })
                      }}
                    />
                  </>
                )}

                {!encerrada && (
                  <BotaoAcao
                    icone={Ban}
                    titulo="Cancelar solicitação"
                    tom="perigo"
                    aoClicar={() => setCancelando(item)}
                  />
                )}
              </>
            )
          }}
        />
      </Painel>

      {/* Formulário */}
      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar solicitação ${edicao.id}` : 'Nova solicitação'}
        icone={Contact}
        largura="xl"
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <>
            {erro && (
              <p className="mb-4 rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <Secao titulo="Cliente">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaSol"
                  rotulo={ehExtensao ? 'Planta de Origem' : 'Planta'}
                  required
                  value={edicao.planta}
                  disabled={salvando}
                  onChange={(e) => mudar('planta', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_PROTHEUS.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                {ehExtensao && (
                  <Selecao
                    id="plantaDestinoSol"
                    rotulo="Planta de Destino"
                    required
                    value={edicao.plantaDestino}
                    disabled={salvando}
                    onChange={(e) => mudar('plantaDestino', e.target.value)}
                    larguraClasse="sm:col-span-4"
                  >
                    <option value="">Selecione...</option>
                    {PLANTAS_PROTHEUS.map((planta) => (
                      <option key={planta.codigo} value={planta.codigo}>
                        {planta.nome}
                      </option>
                    ))}
                  </Selecao>
                )}

                <Entrada
                  id="codClienteSol"
                  rotulo="Cód. Cliente"
                  required
                  value={edicao.codCliente}
                  disabled={salvando}
                  onChange={(e) => mudar('codCliente', e.target.value)}
                  larguraClasse={ehExtensao ? 'sm:col-span-2' : 'sm:col-span-3'}
                />
                <Entrada
                  id="codLojaSol"
                  rotulo="Cód. Loja"
                  maxLength={2}
                  value={edicao.codLoja}
                  disabled={salvando}
                  onChange={(e) => mudar('codLoja', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />

                <Entrada
                  id="razaoSocialSol"
                  rotulo="Razão Social"
                  value={edicao.razaoSocial}
                  disabled={salvando}
                  onChange={(e) => mudar('razaoSocial', e.target.value)}
                  larguraClasse="sm:col-span-8"
                />
                <Entrada
                  id="cnpjSol"
                  rotulo="CNPJ/CPF"
                  value={edicao.cnpj}
                  disabled={salvando}
                  onChange={(e) => mudar('cnpj', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-4"
                />

                {ehExtensao && (
                  <Selecao
                    id="tipoContribuinte"
                    rotulo="Tipo de Contribuinte"
                    value={edicao.tipoContribuinte}
                    disabled={salvando}
                    onChange={(e) => mudar('tipoContribuinte', e.target.value)}
                    larguraClasse="sm:col-span-4"
                  >
                    <option value="">Selecione...</option>
                    {TIPOS_CONTRIBUINTE.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </Selecao>
                )}

                <Campo rotulo="Justificativa" htmlFor="justificativaSol" obrigatorio className="sm:col-span-12">
                  <textarea
                    id="justificativaSol"
                    rows={3}
                    value={edicao.justificativa}
                    disabled={salvando}
                    onChange={(e) => mudar('justificativa', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            {!ehExtensao && (
              <Secao titulo="Campos a Alterar">
                <div className="mb-3 flex justify-end">
                  <Botao variante="neutro" icone={Plus} aoClicar={adicionarCampo} desabilitado={salvando}>
                    Adicionar campo
                  </Botao>
                </div>

                {edicao.camposAlterados.length === 0 ? (
                  <p className="rounded-lg border border-dashed border-slate-200 px-3 py-6 text-center text-xs text-slate-400 dark:border-slate-700">
                    Nenhum campo informado.
                  </p>
                ) : (
                  <div className="grid gap-2">
                    {edicao.camposAlterados.map((item, indice) => (
                      <div key={indice} className="grid items-end gap-2 sm:grid-cols-12">
                        <Selecao
                          id={`campo-${indice}`}
                          rotulo="Campo"
                          value={item.campo}
                          disabled={salvando}
                          onChange={(e) => mudarCampo(indice, 'campo', e.target.value)}
                          larguraClasse="sm:col-span-4"
                        >
                          <option value="">Selecione...</option>
                          {CAMPOS_ALTERACAO.map((campo) => (
                            <option key={campo} value={campo}>
                              {campo}
                            </option>
                          ))}
                        </Selecao>
                        <Entrada
                          id={`atual-${indice}`}
                          rotulo="Valor Atual"
                          value={item.valorAtual}
                          disabled={salvando}
                          onChange={(e) => mudarCampo(indice, 'valorAtual', e.target.value)}
                          larguraClasse="sm:col-span-3"
                        />
                        <Entrada
                          id={`novo-${indice}`}
                          rotulo="Valor Novo"
                          value={item.valorNovo}
                          disabled={salvando}
                          onChange={(e) => mudarCampo(indice, 'valorNovo', e.target.value)}
                          larguraClasse="sm:col-span-4"
                        />
                        <div className="sm:col-span-1">
                          <Botao
                            variante="perigo"
                            aoClicar={() => removerCampo(indice)}
                            desabilitado={salvando}
                            titulo="Remover campo"
                          >
                            &times;
                          </Botao>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </Secao>
            )}

            {ehExtensao && (
              <Secao titulo="Documentação" icone={Paperclip}>
                <div className="mb-3 flex items-end gap-2">
                  <Entrada
                    id="novoDocumento"
                    rotulo="Nome do documento"
                    placeholder="ex.: cartao-cnpj.pdf"
                    disabled={salvando}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault()
                        adicionarDocumento((e.target as HTMLInputElement).value)
                        ;(e.target as HTMLInputElement).value = ''
                      }
                    }}
                    larguraClasse="flex-1"
                  />
                  <span className="pb-2 text-xs text-slate-400">Enter para adicionar</span>
                </div>

                {edicao.documentos.length === 0 ? (
                  <p className="rounded-lg border border-dashed border-slate-200 px-3 py-6 text-center text-xs text-slate-400 dark:border-slate-700">
                    Nenhum documento anexado. A extensão exige documentação antes do envio.
                  </p>
                ) : (
                  <ul className="grid gap-1.5">
                    {edicao.documentos.map((documento, indice) => (
                      <li
                        key={indice}
                        className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                          text-slate-600 dark:border-slate-700 dark:text-slate-300"
                      >
                        <Paperclip size={14} className="shrink-0 text-slate-400" />
                        <span className="flex-1 truncate">{documento}</span>
                        <button
                          type="button"
                          onClick={() =>
                            mudar('documentos', edicao.documentos.filter((_, i) => i !== indice))
                          }
                          disabled={salvando}
                          className="text-xs text-red-600 hover:underline dark:text-red-400"
                        >
                          remover
                        </button>
                      </li>
                    ))}
                  </ul>
                )}

                <p className="mt-2 text-xs text-slate-400">
                  O upload real depende do backend; por ora informe o nome do arquivo.
                </p>
              </Secao>
            )}
          </>
        )}
      </Modal>

      {/* Detalhes */}
      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Solicitação ${detalhe.id}` : ''}
        icone={FileSearch}
        largura="lg"
        aoFechar={() => setDetalhe(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setDetalhe(null)}>
            Fechar
          </Botao>
        }
      >
        {detalhe && (
          <div className="grid gap-4">
            <dl className="grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-3 dark:border-slate-700">
              {[
                ['Cliente', `${detalhe.codCliente}/${detalhe.codLoja} — ${detalhe.razaoSocial}`],
                ['CNPJ', detalhe.cnpj],
                ['Planta', nomePlanta(detalhe.planta)],
                ...(ehExtensao ? [['Planta Destino', nomePlanta(detalhe.plantaDestino)]] : []),
                ...(ehExtensao ? [['Tipo Contribuinte', detalhe.tipoContribuinte || '—']] : []),
                ['Emitente', detalhe.emitente],
                ['Status', descricaoStatusSolicitacao(tipo, detalhe.status)],
                ['Data', formatarData(detalhe.dataInclusao)]
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
              <strong>Justificativa:</strong> {detalhe.justificativa}
            </p>

            {detalhe.camposAlterados.length > 0 && (
              <div>
                <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                  Campos a alterar
                </p>
                <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                  <Tabela
                    itens={detalhe.camposAlterados}
                    chave={(c) => c.campo}
                    colunas={[
                      { cabecalho: 'Campo', conteudo: (c) => c.campo },
                      { cabecalho: 'Valor Atual', conteudo: (c) => c.valorAtual || '—' },
                      { cabecalho: 'Valor Novo', conteudo: (c) => c.valorNovo }
                    ]}
                  />
                </div>
              </div>
            )}

            {detalhe.documentos.length > 0 && (
              <div>
                <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                  Documentos
                </p>
                <ul className="grid gap-1.5">
                  {detalhe.documentos.map((documento) => (
                    <li
                      key={documento}
                      className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                        text-slate-600 dark:border-slate-700 dark:text-slate-300"
                    >
                      <Paperclip size={14} className="shrink-0 text-slate-400" />
                      {documento}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Trilha de avaliação
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.avaliacoes}
                  chave={(a) => a.etapa + a.data}
                  mensagemVazia="Nenhuma etapa avaliada ainda"
                  colunas={[
                    { cabecalho: 'Etapa', conteudo: (a) => a.etapa },
                    { cabecalho: 'Avaliador', conteudo: (a) => a.usuario },
                    { cabecalho: 'Data', conteudo: (a) => formatarData(a.data) },
                    {
                      cabecalho: 'Resultado',
                      conteudo: (a) => (
                        <Etiqueta
                          texto={a.aprovado ? 'Aprovado' : 'Reprovado'}
                          tom={a.aprovado ? 'sucesso' : 'perigo'}
                        />
                      )
                    },
                    { cabecalho: 'Observação', conteudo: (a) => a.observacao }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Avaliação */}
      <Modal
        aberto={avaliando !== null}
        titulo={avaliando?.aprovar ? 'Aprovar etapa' : 'Reprovar solicitação'}
        icone={avaliando?.aprovar ? ThumbsUp : ThumbsDown}
        aoFechar={() => setAvaliando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAvaliando(null)}>
              Cancelar
            </Botao>
            <Botao
              variante={avaliando?.aprovar ? 'sucesso' : 'perigo'}
              aoClicar={confirmarAvaliacao}
            >
              Confirmar
            </Botao>
          </>
        }
      >
        {avaliando && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              Solicitação {avaliando.item.id} — {avaliando.item.razaoSocial}
              <span className="block text-xs text-slate-400">
                Etapa: {descricaoStatusSolicitacao(tipo, avaliando.item.status)}
              </span>
            </p>

            {!avaliando.aprovar && (
              <p className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800 dark:bg-amber-500/10 dark:text-amber-400">
                Reprovar encerra a solicitação — ela será cancelada.
              </p>
            )}

            <Campo rotulo="Observação" htmlFor="observacaoAvaliacao" obrigatorio>
              <textarea
                id="observacaoAvaliacao"
                rows={3}
                value={observacao}
                onChange={(e) => setObservacao(e.target.value)}
                className={classesControle}
              />
            </Campo>
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={cancelando !== null}
        titulo="Cancelar solicitação"
        mensagem={cancelando ? `Deseja realmente cancelar a solicitação ${cancelando.id}?` : ''}
        rotuloConfirmar="Cancelar solicitação"
        aoConfirmar={confirmarCancelamento}
        aoCancelar={() => setCancelando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
