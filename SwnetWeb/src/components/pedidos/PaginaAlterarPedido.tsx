import { useCallback, useEffect, useState } from 'react'
import { FileSearch, Paperclip, Pencil, Plus, Save, Search, Trash2 } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  ALTERACAO_PEDIDO_VAZIA,
  STATUS_ALTERACAO_PEDIDO,
  TIPOS_ALTERACAO_PEDIDO,
  evidenciaObrigatoria,
  excluirSolicitacaoAlteracao,
  listarSolicitacoesAlteracao,
  obterPedidoParaAlteracao,
  plantasDaRegiao,
  salvarSolicitacaoAlteracao,
  type DadosAlteracaoPedido,
  type PedidoDigitado,
  type RegiaoPedido,
  type SolicitacaoAlteracaoPedido
} from '@/services/pedidos/pedidosApi'
import { nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Solicitação de alteração de pedido — compartilhada por
 * Alterar-Pedidos-Brasil e Alterar-Pedidos-PiraUbe.
 *
 * A regra central é a do `ehObrigatorio` do repositório: se a evidência
 * precisa ou não ser anexada, conforme o tipo de alteração, o nível do
 * usuário e — no tipo 7 — se o pedido já está produzido ou programado.
 */

interface PaginaAlterarPedidoProps {
  regiao: RegiaoPedido
  titulo: string
}

export function PaginaAlterarPedido({ regiao, titulo }: PaginaAlterarPedidoProps) {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const plantas = plantasDaRegiao(regiao)

  /** Nível do usuário — o legado lê da sessão; aqui fica ajustável */
  const [nivelUsuario, setNivelUsuario] = useState(2)

  const [filtroPlanta, setFiltroPlanta] = useState('')
  const [solicitacoes, setSolicitacoes] = useState<SolicitacaoAlteracaoPedido[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<Omit<DadosAlteracaoPedido, 'regiao'> | null>(null)
  const [pedidoCarregado, setPedidoCarregado] = useState<(PedidoDigitado & { programado: boolean }) | null>(
    null
  )
  const [buscandoPedido, setBuscandoPedido] = useState(false)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [novaEvidencia, setNovaEvidencia] = useState('')
  const [detalhe, setDetalhe] = useState<SolicitacaoAlteracaoPedido | null>(null)
  const [excluindo, setExcluindo] = useState<SolicitacaoAlteracaoPedido | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setSolicitacoes(await listarSolicitacoesAlteracao(regiao, filtroPlanta))
    } finally {
      setCarregando(false)
    }
  }, [regiao, filtroPlanta])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof Omit<DadosAlteracaoPedido, 'regiao'>>(
    campo: C,
    valor: Omit<DadosAlteracaoPedido, 'regiao'>[C]
  ) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  async function buscarPedido() {
    if (!edicao?.planta || !edicao.pedido.trim()) {
      mostrarToast('Informe a planta e o número do pedido.', 'warning')
      return
    }

    setBuscandoPedido(true)
    try {
      const encontrado = await obterPedidoParaAlteracao(regiao, edicao.planta, edicao.pedido)
      if (!encontrado) {
        setPedidoCarregado(null)
        mostrarToast('Pedido não encontrado nesta planta.', 'warning')
        return
      }
      setPedidoCarregado(encontrado)
    } finally {
      setBuscandoPedido(false)
    }
  }

  function adicionarEvidencia() {
    if (!novaEvidencia.trim() || !edicao) return
    mudar('evidencias', [...edicao.evidencias, novaEvidencia.trim()])
    setNovaEvidencia('')
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarSolicitacaoAlteracao(
        { ...edicao, regiao },
        nivelUsuario,
        pedidoCarregado?.programado ?? false,
        usuario?.login ?? ''
      )
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      setPedidoCarregado(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirSolicitacaoAlteracao(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  function descricaoTipo(id: string) {
    return TIPOS_ALTERACAO_PEDIDO.find((t) => t.id === id)?.descricao ?? id
  }

  const exigeEvidencia = edicao
    ? evidenciaObrigatoria(edicao.tipoAlteracao, nivelUsuario, pedidoCarregado?.programado ?? false)
    : false

  const ehAlteracaoData = edicao?.tipoAlteracao === '7'

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo={titulo}
        acoes={
          <>
            <Selecao
              id="nivelUsuarioAlt"
              rotulo=""
              value={String(nivelUsuario)}
              onChange={(e) => setNivelUsuario(Number(e.target.value))}
              larguraClasse="w-48"
            >
              <option value="1">Nível 1 (sem evidência)</option>
              <option value="2">Nível 2</option>
            </Selecao>
            <Botao
              variante="sucesso"
              icone={Plus}
              aoClicar={() => {
                setErro('')
                setPedidoCarregado(null)
                setEdicao({ ...ALTERACAO_PEDIDO_VAZIA })
              }}
            >
              Nova solicitação
            </Botao>
          </>
        }
      />

      <Painel titulo="Filtros">
        <Selecao
          id="filtroPlantaAlt"
          rotulo="Planta"
          value={filtroPlanta}
          onChange={(e) => setFiltroPlanta(e.target.value)}
          larguraClasse="max-w-xs"
        >
          <option value="">Todas as plantas</option>
          {plantas.map((planta) => (
            <option key={planta.codigo} value={planta.codigo}>
              {planta.nome}
            </option>
          ))}
        </Selecao>
      </Painel>

      <Painel titulo={`${solicitacoes.length} solicitação(ões)`} icone={Pencil}>
        <Tabela
          itens={solicitacoes}
          chave={(item) => item.id}
          carregando={carregando}
          mensagemVazia="Nenhuma solicitação de alteração"
          colunas={[
            { cabecalho: '#', conteudo: (s) => s.id, className: 'w-20' },
            { cabecalho: 'Data', conteudo: (s) => formatarData(s.dataInclusao) },
            { cabecalho: 'Planta', conteudo: (s) => nomePlanta(s.planta) },
            { cabecalho: 'Pedido', conteudo: (s) => s.pedido },
            { cabecalho: 'Cliente', conteudo: (s) => s.nomeCliente || '—' },
            { cabecalho: 'Referência', conteudo: (s) => s.referencia || '—' },
            { cabecalho: 'Tipo de Alteração', conteudo: (s) => descricaoTipo(s.tipoAlteracao) },
            {
              cabecalho: 'Evidência',
              conteudo: (s) =>
                s.evidencias.length > 0 ? (
                  <Etiqueta texto={`${s.evidencias.length} anexo(s)`} tom="info" />
                ) : (
                  '—'
                )
            },
            {
              cabecalho: 'Status',
              conteudo: (s) => (
                <Etiqueta
                  texto={STATUS_ALTERACAO_PEDIDO.find((x) => x.codigo === s.status)?.descricao ?? s.status}
                  tom={
                    s.status === 'APROVADA' ? 'sucesso' : s.status === 'REPROVADA' ? 'perigo' : 'alerta'
                  }
                />
              )
            },
            { cabecalho: 'Solicitante', conteudo: (s) => s.solicitante }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(item)} />
              {item.status === 'PENDENTE' && (
                <BotaoAcao
                  icone={Trash2}
                  titulo="Excluir solicitação"
                  tom="perigo"
                  aoClicar={() => setExcluindo(item)}
                />
              )}
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo="Nova solicitação de alteração"
        icone={Pencil}
        largura="lg"
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

            <Secao titulo="Pedido">
              <div className="grid items-end gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaAlt"
                  rotulo="Planta"
                  required
                  value={edicao.planta}
                  disabled={salvando}
                  onChange={(e) => {
                    mudar('planta', e.target.value)
                    setPedidoCarregado(null)
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {plantas.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="pedidoAlt"
                  rotulo="Pedido"
                  required
                  value={edicao.pedido}
                  disabled={salvando}
                  onChange={(e) => {
                    mudar('pedido', e.target.value)
                    setPedidoCarregado(null)
                  }}
                  larguraClasse="sm:col-span-4"
                />

                <div className="sm:col-span-4">
                  <Botao icone={Search} aoClicar={buscarPedido} desabilitado={buscandoPedido || salvando}>
                    {buscandoPedido ? 'Buscando...' : 'Buscar pedido'}
                  </Botao>
                </div>
              </div>

              {pedidoCarregado && (
                <dl className="mt-4 grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-4 dark:border-slate-700">
                  {[
                    ['Cliente', pedidoCarregado.nomeCliente],
                    ['Produto', pedidoCarregado.codigoProduto],
                    ['Referência', pedidoCarregado.referencia],
                    ['Quantidade', String(pedidoCarregado.quantidade)],
                    ['Entrega atual', formatarData(pedidoCarregado.dataEntrega)],
                    [
                      'Situação',
                      pedidoCarregado.programado ? 'Produzido / programado' : 'Ainda não programado'
                    ]
                  ].map(([rotulo, valor]) => (
                    <div key={rotulo}>
                      <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                      <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                        {valor}
                      </dd>
                    </div>
                  ))}
                </dl>
              )}
            </Secao>

            <Secao titulo="Alteração">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoAlteracao"
                  rotulo="Tipo de Alteração"
                  required
                  value={edicao.tipoAlteracao}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoAlteracao', e.target.value)}
                  larguraClasse="sm:col-span-6"
                >
                  <option value="">Selecione o tipo...</option>
                  {TIPOS_ALTERACAO_PEDIDO.map((tipo) => (
                    <option key={tipo.id} value={tipo.id}>
                      {tipo.descricao}
                    </option>
                  ))}
                </Selecao>

                {ehAlteracaoData ? (
                  <>
                    <Entrada
                      id="novaDataDe"
                      rotulo="Nova Data (De)"
                      type="date"
                      value={edicao.novaDataDe}
                      disabled={salvando}
                      onChange={(e) => mudar('novaDataDe', e.target.value)}
                      larguraClasse="sm:col-span-3"
                    />
                    <Entrada
                      id="novaDataAte"
                      rotulo="Nova Data (Até)"
                      type="date"
                      value={edicao.novaDataAte}
                      disabled={salvando}
                      onChange={(e) => mudar('novaDataAte', e.target.value)}
                      larguraClasse="sm:col-span-3"
                    />
                  </>
                ) : (
                  <Entrada
                    id="valorNovo"
                    rotulo="Para"
                    value={edicao.valorNovo}
                    disabled={salvando || !edicao.tipoAlteracao}
                    ajuda="Novo valor do campo"
                    onChange={(e) => mudar('valorNovo', e.target.value)}
                    larguraClasse="sm:col-span-6"
                  />
                )}

                <Campo rotulo="Observação" htmlFor="observacaoAlt" obrigatorio className="sm:col-span-12">
                  <textarea
                    id="observacaoAlt"
                    rows={4}
                    value={edicao.observacao}
                    disabled={salvando}
                    onChange={(e) => mudar('observacao', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Evidência" icone={Paperclip}>
              {exigeEvidencia && (
                <p className="mb-3 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800 dark:bg-amber-500/10 dark:text-amber-400">
                  {ehAlteracaoData
                    ? 'O pedido já está produzido ou programado: a evidência é obrigatória.'
                    : 'Este tipo de alteração exige evidência anexada.'}
                </p>
              )}

              <div className="mb-3 flex items-end gap-2">
                <Entrada
                  id="novaEvidencia"
                  rotulo="Nome do arquivo"
                  placeholder="ex.: email-cliente.pdf"
                  value={novaEvidencia}
                  disabled={salvando}
                  onChange={(e) => setNovaEvidencia(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault()
                      adicionarEvidencia()
                    }
                  }}
                  larguraClasse="flex-1"
                />
                <Botao variante="neutro" aoClicar={adicionarEvidencia} desabilitado={salvando}>
                  Anexar
                </Botao>
              </div>

              {edicao.evidencias.length === 0 ? (
                <p className="rounded-lg border border-dashed border-slate-200 px-3 py-4 text-center text-xs text-slate-400 dark:border-slate-700">
                  Nenhuma evidência anexada.
                </p>
              ) : (
                <ul className="grid gap-1.5">
                  {edicao.evidencias.map((evidencia, indice) => (
                    <li
                      key={indice}
                      className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                        text-slate-600 dark:border-slate-700 dark:text-slate-300"
                    >
                      <Paperclip size={14} className="shrink-0 text-slate-400" />
                      <span className="flex-1 truncate">{evidencia}</span>
                      <button
                        type="button"
                        onClick={() =>
                          mudar('evidencias', edicao.evidencias.filter((_, i) => i !== indice))
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
            </Secao>
          </>
        )}
      </Modal>

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
                ['Pedido', detalhe.pedido],
                ['Planta', nomePlanta(detalhe.planta)],
                ['Cliente', detalhe.nomeCliente || '—'],
                ['Produto', detalhe.codigoProduto || '—'],
                ['Referência', detalhe.referencia || '—'],
                ['Tipo', descricaoTipo(detalhe.tipoAlteracao)],
                ['Valor atual', detalhe.valorAtual || '—'],
                [
                  'Valor novo',
                  detalhe.tipoAlteracao === '7'
                    ? `${formatarData(detalhe.novaDataDe)} a ${formatarData(detalhe.novaDataAte)}`
                    : detalhe.valorNovo || '—'
                ],
                ['Solicitante', detalhe.solicitante]
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
              <strong>Observação:</strong> {detalhe.observacao}
            </p>

            {detalhe.evidencias.length > 0 && (
              <ul className="grid gap-1.5">
                {detalhe.evidencias.map((evidencia) => (
                  <li
                    key={evidencia}
                    className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                      text-slate-600 dark:border-slate-700 dark:text-slate-300"
                  >
                    <Paperclip size={14} className="shrink-0 text-slate-400" />
                    {evidencia}
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir solicitação"
        mensagem={excluindo ? `Deseja realmente excluir a solicitação ${excluindo.id}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
