import { useCallback, useEffect, useState } from 'react'
import { Check, Coins, FileSearch, History, Search, ThumbsDown, ThumbsUp } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  ETAPAS_CREDITO,
  listarSolicitacoes,
  listarTodasSolicitacoes,
  rotuloStatusCredito,
  validarSolicitacao,
  type DadosValidacaoCredito,
  type FiltrosCredito,
  type SolicitacaoCredito,
  type StatusCredito
} from '@/services/validacoes/validacoesApi'
import { NIVEIS_VALIDADOR, PLANTAS_CREDITO, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Validacoes/Solicitacoes-De-Credito.
 *
 * É a fila de aprovação das solicitações criadas em Financeiro. O que o
 * usuário enxerga depende do nível dele na cadeia (10 Analista → 60 CFO);
 * o nível 10 ainda tem três abas: análises, aprovações e transferências.
 */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

const ABAS = [
  { id: 'analises', rotulo: 'Análises Pendentes' },
  { id: 'aprovacoes', rotulo: 'Aprovações Pendentes' },
  { id: 'transferencias', rotulo: 'Transferências Pendentes' }
]

const VALIDACAO_VAZIA: Omit<DadosValidacaoCredito, 'idSolicitacao'> = {
  rejeitar: false,
  limiteValidado: '',
  observacao: '',
  scoreSerasa: '',
  classeRiscoSerasa: '',
  limiteSugeridoSerasa: '',
  scoreSw: '',
  classeRiscoSw: '',
  limiteSugeridoSw: '',
  valorSeguradora: '',
  faturamentoEstimado: '',
  balancoAno: ''
}

function tomStatusCredito(status: StatusCredito) {
  if (status === 'APROVADO') return 'sucesso' as const
  if (status === 'REJEITADO') return 'perigo' as const
  return 'alerta' as const
}

export function SolicitacoesCredito() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  /**
   * O nível do validador vem do cadastro de Validadores de Crédito. Sem
   * backend, deixamos o usuário escolher para conseguir percorrer a cadeia.
   */
  const [nivel, setNivel] = useState('10')
  const [aba, setAba] = useState('analises')
  const [filtros, setFiltros] = useState<FiltrosCredito>({ planta: 'TODOS', aba: 'analises', termo: '' })
  const [verTodas, setVerTodas] = useState(false)

  const [solicitacoes, setSolicitacoes] = useState<SolicitacaoCredito[]>([])
  const [carregando, setCarregando] = useState(true)

  const [detalhe, setDetalhe] = useState<SolicitacaoCredito | null>(null)
  const [validando, setValidando] = useState<SolicitacaoCredito | null>(null)
  const [dados, setDados] = useState(VALIDACAO_VAZIA)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const consulta = { ...filtros, aba }
      setSolicitacoes(
        verTodas ? await listarTodasSolicitacoes(consulta) : await listarSolicitacoes(consulta, nivel)
      )
    } finally {
      setCarregando(false)
    }
  }, [filtros, aba, nivel, verTodas])

  useEffect(() => {
    carregar()
  }, [carregar])

  function abrirValidacao(solicitacao: SolicitacaoCredito) {
    setErro('')
    setDados({
      ...VALIDACAO_VAZIA,
      limiteValidado: String(solicitacao.valorSolicitado),
      scoreSerasa: solicitacao.scoreSerasa,
      classeRiscoSerasa: solicitacao.classeRiscoSerasa,
      limiteSugeridoSerasa: solicitacao.limiteSugeridoSerasa,
      scoreSw: solicitacao.scoreSw,
      classeRiscoSw: solicitacao.classeRiscoSw,
      limiteSugeridoSw: solicitacao.limiteSugeridoSw,
      valorSeguradora: solicitacao.valorSeguradora,
      faturamentoEstimado: solicitacao.faturamentoEstimado,
      balancoAno: solicitacao.balancoAno
    })
    setValidando(solicitacao)
  }

  function mudar<C extends keyof typeof VALIDACAO_VAZIA>(campo: C, valor: (typeof VALIDACAO_VAZIA)[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  async function confirmar(rejeitar: boolean) {
    if (!validando) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await validarSolicitacao(
        { ...dados, rejeitar, idSolicitacao: validando.id },
        nivel,
        usuario?.login ?? ''
      )

      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }

      setValidando(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  /** O analista (nível 10) é quem preenche os dados de score e limite */
  const ehAnalista = nivel === '10'

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Solicitações de Crédito" />

      <Painel titulo="Filtros">
        <div className="grid items-end gap-4 sm:grid-cols-12">
          <Selecao
            id="nivelValidador"
            rotulo="Meu nível de validador"
            value={nivel}
            onChange={(e) => setNivel(e.target.value)}
            ajuda="Viria do cadastro de Validadores de Crédito"
            larguraClasse="sm:col-span-4"
          >
            {NIVEIS_VALIDADOR.map((item) => (
              <option key={item.valor} value={item.valor}>
                {item.valor} — {item.rotulo}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroPlanta"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="TODOS">Todas</option>
            {PLANTAS_CREDITO.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="termoCredito"
            rotulo="Cliente ou CNPJ"
            value={filtros.termo}
            onChange={(e) => setFiltros({ ...filtros, termo: e.target.value })}
            larguraClasse="sm:col-span-3"
          />

          <div className="sm:col-span-2">
            <Botao icone={Search} aoClicar={carregar} desabilitado={carregando}>
              {carregando ? '...' : 'Pesquisar'}
            </Botao>
          </div>
        </div>

        <label className="mt-3 flex cursor-pointer items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
          <input
            type="checkbox"
            checked={verTodas}
            onChange={(e) => setVerTodas(e.target.checked)}
            className="h-4 w-4 accent-swnet-600"
          />
          Ver todas as solicitações (histórico), ignorando a fila do meu nível
        </label>
      </Painel>

      {ehAnalista && !verTodas && (
        <div className="flex flex-wrap gap-1 rounded-xl border border-slate-200 bg-white p-1 dark:border-slate-800 dark:bg-slate-900">
          {ABAS.map((item) => (
            <button
              key={item.id}
              type="button"
              onClick={() => setAba(item.id)}
              className={`rounded-lg px-3 py-2 text-sm font-medium transition-colors ${
                aba === item.id
                  ? 'bg-swnet-600 text-white'
                  : 'text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800'
              }`}
            >
              {item.rotulo}
            </button>
          ))}
        </div>
      )}

      <Painel titulo={`${solicitacoes.length} solicitação(ões)`} icone={Coins}>
        <Tabela
          itens={solicitacoes}
          chave={(s) => s.id}
          carregando={carregando}
          mensagemVazia="Nenhuma solicitação nesta fila"
          colunas={[
            { cabecalho: '#', conteudo: (s) => s.id, className: 'w-20' },
            { cabecalho: 'Solicitante', conteudo: (s) => s.solicitante },
            { cabecalho: 'Cliente', conteudo: (s) => s.cliente },
            { cabecalho: 'CNPJ', conteudo: (s) => s.cnpj },
            { cabecalho: 'Planta Origem', conteudo: (s) => nomePlanta(s.plantaOrigem) },
            {
              cabecalho: 'Planta Destino',
              conteudo: (s) => (s.plantaDestino ? nomePlanta(s.plantaDestino) : '—')
            },
            {
              cabecalho: 'Limite Ant. Origem',
              conteudo: (s) => moeda.format(s.limiteAnteriorOrigem),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Limite Ant. Destino',
              conteudo: (s) => (s.plantaDestino ? moeda.format(s.limiteAnteriorDestino) : '—'),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Valor',
              conteudo: (s) => moeda.format(s.valorSolicitado),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Status',
              conteudo: (s) => (
                <Etiqueta texto={rotuloStatusCredito(s.status)} tom={tomStatusCredito(s.status)} />
              )
            },
            { cabecalho: 'Data', conteudo: (s) => formatarData(s.dataSolicitacao) }
          ]}
          acoes={(solicitacao) => {
            const etapa = ETAPAS_CREDITO.find((e) => e.status === solicitacao.status)
            const podeValidar = etapa?.nivel === nivel

            return (
              <>
                <BotaoAcao
                  icone={FileSearch}
                  titulo="Ver detalhes"
                  aoClicar={() => setDetalhe(solicitacao)}
                />
                {podeValidar && (
                  <BotaoAcao
                    icone={Check}
                    titulo="Validar solicitação"
                    aoClicar={() => abrirValidacao(solicitacao)}
                  />
                )}
              </>
            )
          }}
        />
      </Painel>

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
                ['Cliente', detalhe.cliente],
                ['CNPJ', detalhe.cnpj],
                ['Tipo', detalhe.tipo === 'AUMENTO' ? 'Aumento de limite' : 'Transferência'],
                ['Solicitante', detalhe.solicitante],
                ['Planta Origem', nomePlanta(detalhe.plantaOrigem)],
                ['Planta Destino', detalhe.plantaDestino ? nomePlanta(detalhe.plantaDestino) : '—'],
                ['Valor solicitado', moeda.format(detalhe.valorSolicitado)],
                ['Status', rotuloStatusCredito(detalhe.status)],
                ['Data', formatarData(detalhe.dataSolicitacao)]
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            {detalhe.observacaoSolicitante && (
              <p className="rounded-lg bg-slate-50 px-3 py-2 text-sm text-slate-600 dark:bg-slate-800 dark:text-slate-300">
                <strong>Justificativa:</strong> {detalhe.observacaoSolicitante}
              </p>
            )}

            {detalhe.scoreSerasa && (
              <Secao titulo="Análise de Crédito">
                <dl className="grid gap-3 sm:grid-cols-3">
                  {[
                    ['Score Serasa', detalhe.scoreSerasa],
                    ['Classe Risco Serasa', detalhe.classeRiscoSerasa],
                    ['Limite Sugerido Serasa', detalhe.limiteSugeridoSerasa],
                    ['Score SW', detalhe.scoreSw],
                    ['Classe Risco SW', detalhe.classeRiscoSw],
                    ['Limite Sugerido SW', detalhe.limiteSugeridoSw],
                    ['Valor Seguradora', detalhe.valorSeguradora],
                    ['Faturamento Estimado', detalhe.faturamentoEstimado],
                    ['Ano do Balanço', detalhe.balancoAno]
                  ].map(([rotulo, valor]) => (
                    <div key={rotulo}>
                      <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                      <dd className="mt-0.5 text-sm text-slate-700 dark:text-slate-200">{valor || '—'}</dd>
                    </div>
                  ))}
                </dl>
              </Secao>
            )}

            <div>
              <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                <History size={13} /> Trilha de aprovação
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.aprovacoes}
                  chave={(item) => item.etapa}
                  mensagemVazia="Nenhuma etapa validada ainda"
                  colunas={[
                    { cabecalho: 'Etapa', conteudo: (a) => a.etapa },
                    { cabecalho: 'Aprovador', conteudo: (a) => a.validador },
                    { cabecalho: 'Data', conteudo: (a) => formatarData(a.data) },
                    { cabecalho: 'Observação', conteudo: (a) => a.observacao }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      <Modal
        aberto={validando !== null}
        titulo={validando ? `Validar solicitação ${validando.id}` : ''}
        icone={Check}
        largura="lg"
        aoFechar={() => setValidando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setValidando(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao
              variante="perigo"
              icone={ThumbsDown}
              aoClicar={() => confirmar(true)}
              desabilitado={salvando}
            >
              Rejeitar
            </Botao>
            <Botao
              variante="sucesso"
              icone={ThumbsUp}
              aoClicar={() => confirmar(false)}
              desabilitado={salvando}
            >
              {salvando ? 'Gravando...' : 'Aprovar'}
            </Botao>
          </>
        }
      >
        {validando && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <p className="text-sm text-slate-600 dark:text-slate-300">
              <strong>{validando.cliente}</strong> — solicitou {moeda.format(validando.valorSolicitado)}
              {validando.plantaDestino
                ? ` em transferência de ${nomePlanta(validando.plantaOrigem)} para ${nomePlanta(validando.plantaDestino)}`
                : ` de limite em ${nomePlanta(validando.plantaOrigem)}`}
              .
            </p>

            {ehAnalista && (
              <Secao titulo="Dados da Análise">
                <div className="grid gap-4 sm:grid-cols-3">
                  <Entrada
                    id="scoreSerasa"
                    rotulo="Score Serasa"
                    value={dados.scoreSerasa}
                    disabled={salvando}
                    onChange={(e) => mudar('scoreSerasa', e.target.value)}
                  />
                  <Entrada
                    id="classeRiscoSerasa"
                    rotulo="Classe Risco Serasa"
                    maxLength={2}
                    value={dados.classeRiscoSerasa}
                    disabled={salvando}
                    onChange={(e) => mudar('classeRiscoSerasa', e.target.value.toUpperCase())}
                  />
                  <Entrada
                    id="limiteSugeridoSerasa"
                    rotulo="Limite Sugerido Serasa"
                    value={dados.limiteSugeridoSerasa}
                    disabled={salvando}
                    onChange={(e) => mudar('limiteSugeridoSerasa', e.target.value)}
                  />
                  <Entrada
                    id="scoreSw"
                    rotulo="Score SW"
                    value={dados.scoreSw}
                    disabled={salvando}
                    onChange={(e) => mudar('scoreSw', e.target.value)}
                  />
                  <Entrada
                    id="classeRiscoSw"
                    rotulo="Classe Risco SW"
                    maxLength={2}
                    value={dados.classeRiscoSw}
                    disabled={salvando}
                    onChange={(e) => mudar('classeRiscoSw', e.target.value.toUpperCase())}
                  />
                  <Entrada
                    id="limiteSugeridoSw"
                    rotulo="Limite Sugerido SW"
                    value={dados.limiteSugeridoSw}
                    disabled={salvando}
                    onChange={(e) => mudar('limiteSugeridoSw', e.target.value)}
                  />
                  <Entrada
                    id="valorSeguradora"
                    rotulo="Valor Seguradora"
                    value={dados.valorSeguradora}
                    disabled={salvando}
                    onChange={(e) => mudar('valorSeguradora', e.target.value)}
                  />
                  <Entrada
                    id="faturamentoEstimado"
                    rotulo="Faturamento Estimado"
                    value={dados.faturamentoEstimado}
                    disabled={salvando}
                    onChange={(e) => mudar('faturamentoEstimado', e.target.value)}
                  />
                  <Entrada
                    id="balancoAno"
                    rotulo="Ano do Balanço"
                    maxLength={4}
                    value={dados.balancoAno}
                    disabled={salvando}
                    onChange={(e) => mudar('balancoAno', e.target.value.replace(/\D/g, ''))}
                  />
                </div>
              </Secao>
            )}

            <Entrada
              id="limiteValidado"
              rotulo="Limite validado"
              value={dados.limiteValidado}
              disabled={salvando}
              ajuda="Deixe como está para manter o valor solicitado"
              onChange={(e) => mudar('limiteValidado', e.target.value)}
              larguraClasse="max-w-xs"
            />

            <Campo rotulo="Observação da validação" htmlFor="observacaoValidacao" obrigatorio>
              <textarea
                id="observacaoValidacao"
                rows={3}
                value={dados.observacao}
                disabled={salvando}
                placeholder="Parecer desta etapa..."
                onChange={(e) => mudar('observacao', e.target.value)}
                className={classesControle}
              />
            </Campo>
          </div>
        )}
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
