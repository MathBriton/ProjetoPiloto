import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  ClipboardList,
  Eraser,
  FileSearch,
  FlaskConical,
  History,
  Package,
  Pencil,
  Plus,
  RotateCcw,
  Save,
  Search,
  ThumbsDown,
  ThumbsUp
} from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Chave, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  ESTILOS_CAIXA,
  ESTRUTURAS_PAPELAO,
  FILTROS_SPD_VAZIOS,
  MOTIVOS_REJEICAO_SPD,
  PROCESSOS_SPD,
  SPD_VAZIA,
  STATUS_ANDAMENTO_SPD,
  alterarAndamentoSpd,
  cancelarSpd,
  decidirSpd,
  listarClientesDesenvolvimento,
  listarSpds,
  salvarSpd,
  salvarTestesFisicos,
  type AreaAprovacao,
  type FiltrosSpd,
  type SolicitacaoSpd as Spd
} from '@/services/desenvolvimento/desenvolvimentoApi'
import { PLANTAS_PROTHEUS, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Desenvolvimento/Solicitacao-De-SPD-V2.
 *
 * O formulário legado tem ~60 campos; aqui eles ficam agrupados nas mesmas
 * seções do original. O fluxo exige aprovação de três áreas (Comercial,
 * Desenvolvimento e PCP) — a rejeição de qualquer uma encerra a SPD.
 */

const CLIENTES = listarClientesDesenvolvimento()

const AREAS: AreaAprovacao[] = ['COMERCIAL', 'DESENVOLVIMENTO', 'PCP']

const TOM_STATUS: Record<Spd['status'], 'neutro' | 'info' | 'sucesso' | 'perigo' | 'alerta'> = {
  ABERTA: 'neutro',
  EM_ANDAMENTO: 'info',
  APROVADA: 'sucesso',
  REJEITADA: 'perigo',
  CANCELADA: 'alerta'
}

const ROTULO_STATUS: Record<Spd['status'], string> = {
  ABERTA: 'Aberta',
  EM_ANDAMENTO: 'Em andamento',
  APROVADA: 'Aprovada',
  REJEITADA: 'Rejeitada',
  CANCELADA: 'Cancelada'
}

type Rascunho = Partial<Spd> & { id?: number }

export function SolicitacaoSpd() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosSpd>(FILTROS_SPD_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosSpd>(FILTROS_SPD_VAZIOS)
  const [spds, setSpds] = useState<Spd[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<Rascunho | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const [detalhe, setDetalhe] = useState<Spd | null>(null)
  const [testes, setTestes] = useState<Spd | null>(null)
  const [andamento, setAndamento] = useState<Spd | null>(null)
  const [novoAndamento, setNovoAndamento] = useState({ processo: '', complemento: '' })
  const [decidindo, setDecidindo] = useState<{ spd: Spd; area: AreaAprovacao; aprovar: boolean } | null>(null)
  const [motivoDecisao, setMotivoDecisao] = useState('')
  const [cancelando, setCancelando] = useState<Spd | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setSpds(await listarSpds(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof Spd>(campo: C, valor: Spd[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function escolherCliente(codigo: string) {
    const cliente = CLIENTES.find((c) => c.cod === codigo)
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            codCliente: codigo,
            codLoja: cliente?.loja ?? '',
            nomeCliente: cliente?.nome ?? ''
          }
        : atual
    )
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarSpd(edicao, usuario?.login ?? '')
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

  async function salvarEnsaios() {
    if (!testes) return

    setSalvando(true)
    try {
      const resultado = await salvarTestesFisicos(testes.id, {
        gramaturaReferencial: testes.gramaturaReferencial,
        esmagamento: testes.esmagamento,
        ectColuna: testes.ectColuna,
        mullenReferencial: testes.mullenReferencial,
        espessura: testes.espessura,
        bctCompressao: testes.bctCompressao,
        cobbInterno: testes.cobbInterno,
        cobbExterno: testes.cobbExterno
      })
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
      if (resultado.sucesso) {
        setTestes(null)
        await carregar()
      }
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarAndamento() {
    if (!andamento) return

    const resultado = await alterarAndamentoSpd(
      andamento.id,
      novoAndamento.processo,
      novoAndamento.complemento,
      usuario?.login ?? ''
    )
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setAndamento(null)
      setNovoAndamento({ processo: '', complemento: '' })
      await carregar()
    }
  }

  async function confirmarDecisao() {
    if (!decidindo) return

    const resultado = await decidirSpd(
      decidindo.spd.id,
      decidindo.area,
      decidindo.aprovar,
      motivoDecisao,
      usuario?.login ?? ''
    )
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setDecidindo(null)
      setMotivoDecisao('')
      await carregar()
    }
  }

  async function confirmarCancelamento() {
    if (!cancelando) return

    const reverter = cancelando.status === 'CANCELADA'
    const resultado = await cancelarSpd(cancelando.id, reverter)
    setCancelando(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Solicitação de Projeto ao Desenvolvimento (SPD)"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ ...SPD_VAZIA })
            }}
          >
            Nova SPD
          </Botao>
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
            id="filtroPlantaSpd"
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

          <Entrada
            id="filtroClienteSpd"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroReferenciaSpd"
            rotulo="Referência"
            value={filtros.referencia}
            onChange={(e) => setFiltros({ ...filtros, referencia: e.target.value })}
            larguraClasse="sm:col-span-2"
          />
          <Selecao
            id="filtroStatusSpd"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todos</option>
            {Object.entries(ROTULO_STATUS).map(([codigo, rotulo]) => (
              <option key={codigo} value={codigo}>
                {rotulo}
              </option>
            ))}
          </Selecao>

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Pesquisar">
              <span className="sr-only">Pesquisar</span>
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              titulo="Limpar filtros"
              aoClicar={() => {
                setFiltros(FILTROS_SPD_VAZIOS)
                setFiltrosAplicados(FILTROS_SPD_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${spds.length} SPD(s)`} icone={ClipboardList}>
        <Tabela
          itens={spds}
          chave={(spd) => spd.id}
          carregando={carregando}
          mensagemVazia="Nenhuma SPD encontrada"
          colunas={[
            { cabecalho: '#', conteudo: (s) => s.id, className: 'w-20' },
            { cabecalho: 'Cliente', conteudo: (s) => s.nomeCliente },
            { cabecalho: 'Vendedor', conteudo: (s) => s.vendedor },
            { cabecalho: 'Referência', conteudo: (s) => s.referencia },
            { cabecalho: 'FT', conteudo: (s) => s.ft },
            {
              cabecalho: 'Status',
              conteudo: (s) => <Etiqueta texto={ROTULO_STATUS[s.status]} tom={TOM_STATUS[s.status]} />
            },
            { cabecalho: 'Processo', conteudo: (s) => s.tipoSolicitacao },
            { cabecalho: 'Andamento', conteudo: (s) => s.statusAndamento },
            { cabecalho: 'Data Emissão', conteudo: (s) => formatarData(s.dataEmissao) }
          ]}
          acoes={(spd) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(spd)} />

              {spd.status !== 'CANCELADA' && spd.status !== 'APROVADA' && (
                <>
                  <BotaoAcao
                    icone={Pencil}
                    titulo="Editar SPD"
                    aoClicar={() => {
                      setErro('')
                      setEdicao({ ...spd })
                    }}
                  />
                  <BotaoAcao
                    icone={FlaskConical}
                    titulo="Testes físicos"
                    aoClicar={() => setTestes({ ...spd })}
                  />
                  <BotaoAcao
                    icone={History}
                    titulo="Alterar andamento"
                    aoClicar={() => {
                      setNovoAndamento({ processo: '', complemento: '' })
                      setAndamento(spd)
                    }}
                  />
                </>
              )}

              <BotaoAcao
                icone={spd.status === 'CANCELADA' ? RotateCcw : Ban}
                titulo={spd.status === 'CANCELADA' ? 'Reverter cancelamento' : 'Cancelar SPD'}
                tom={spd.status === 'CANCELADA' ? 'neutro' : 'perigo'}
                aoClicar={() => setCancelando(spd)}
              />
            </>
          )}
        />
      </Painel>

      {/* Formulário completo */}
      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar SPD ${edicao.id}` : 'Nova SPD'}
        icone={ClipboardList}
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

            <Secao titulo="Identificação">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="codPlantaSpd"
                  rotulo="Planta"
                  required
                  value={edicao.codPlanta ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codPlanta', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_PROTHEUS.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="codClienteSpd"
                  rotulo="Cliente"
                  required
                  value={edicao.codCliente ?? ''}
                  disabled={salvando}
                  onChange={(e) => escolherCliente(e.target.value)}
                  larguraClasse="sm:col-span-5"
                >
                  <option value="">Selecione...</option>
                  {CLIENTES.map((cliente) => (
                    <option key={cliente.cod} value={cliente.cod}>
                      {cliente.cod} — {cliente.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="codLojaSpd"
                  rotulo="Loja"
                  maxLength={2}
                  value={edicao.codLoja ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codLoja', e.target.value)}
                  larguraClasse="sm:col-span-1"
                />
                <Entrada
                  id="vendedorSpd"
                  rotulo="Vendedor"
                  value={edicao.vendedor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('vendedor', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Selecao
                  id="tipoSolicitacao"
                  rotulo="Tipo de Solicitação"
                  required
                  value={edicao.tipoSolicitacao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoSolicitacao', e.target.value)}
                  larguraClasse="sm:col-span-5"
                >
                  <option value="">Selecione...</option>
                  {PROCESSOS_SPD.map((processo) => (
                    <option key={processo} value={processo}>
                      {processo}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="numeroFTCopia"
                  rotulo="Copiar de FT"
                  value={edicao.numeroFtCopia ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('numeroFtCopia', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="idDesenvolvedor"
                  rotulo="Técnico/Desenvolvedor"
                  required
                  value={edicao.desenvolvedor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('desenvolvedor', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="referenciaSpd"
                  rotulo="Referência"
                  required
                  value={edicao.referencia ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('referencia', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Selecao
                  id="estilo"
                  rotulo="Estilo"
                  value={edicao.estilo ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('estilo', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {ESTILOS_CAIXA.map((estilo) => (
                    <option key={estilo} value={estilo}>
                      {estilo}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="nivelLotePiloto"
                  rotulo="Nível de Lote Piloto"
                  value={edicao.nivelLotePiloto ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('nivelLotePiloto', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="toleranciaMaxima"
                  rotulo="Tolerância Máx. (%)"
                  value={edicao.toleranciaMaxima ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('toleranciaMaxima', e.target.value)}
                  larguraClasse="sm:col-span-1"
                />
                <Entrada
                  id="toleranciaMinima"
                  rotulo="Tolerância Mín. (%)"
                  value={edicao.toleranciaMinima ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('toleranciaMinima', e.target.value)}
                  larguraClasse="sm:col-span-1"
                />

                <div className="sm:col-span-12">
                  <Chave
                    rotulo="Necessita acompanhamento"
                    marcado={edicao.necessitaAcompanhamento ?? false}
                    desabilitado={salvando}
                    aoAlternar={(v) => mudar('necessitaAcompanhamento', v)}
                  />
                </div>

                <Campo rotulo="Observações" htmlFor="observacoesSpd" className="sm:col-span-12">
                  <textarea
                    id="observacoesSpd"
                    rows={2}
                    value={edicao.observacoes ?? ''}
                    disabled={salvando}
                    onChange={(e) => mudar('observacoes', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Dimensional" icone={Package}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="comprimento"
                  rotulo="Comprimento (mm)"
                  type="number"
                  required
                  value={edicao.comprimento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('comprimento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="largura"
                  rotulo="Largura (mm)"
                  type="number"
                  required
                  value={edicao.largura ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('largura', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="altura"
                  rotulo="Altura (mm)"
                  type="number"
                  required
                  value={edicao.altura ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('altura', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="numeroMatriz"
                  rotulo="Número da Matriz"
                  value={edicao.numeroMatriz ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('numeroMatriz', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="toleranciaPeso"
                  rotulo="Tolerância de Peso (± %)"
                  value={edicao.toleranciaPeso ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('toleranciaPeso', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <div className="sm:col-span-8">
                  <Chave
                    rotulo="Cliente usa tara"
                    marcado={edicao.clienteUsaTara ?? false}
                    desabilitado={salvando}
                    aoAlternar={(v) => mudar('clienteUsaTara', v)}
                  />
                </div>
              </div>
            </Secao>

            <Secao titulo="Estrutura de Papelão Ondulado">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="estruturaPapelao"
                  rotulo="Estrutura"
                  required
                  value={edicao.estruturaPapelao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('estruturaPapelao', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {ESTRUTURAS_PAPELAO.map((estrutura) => (
                    <option key={estrutura} value={estrutura}>
                      {estrutura}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="gramatura"
                  rotulo="Gramatura"
                  value={edicao.gramatura ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('gramatura', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="larguraChapa"
                  rotulo="Largura da Chapa (mm)"
                  value={edicao.larguraChapa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('larguraChapa', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="comprimentoChapa"
                  rotulo="Comprimento da Chapa (mm)"
                  value={edicao.comprimentoChapa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('comprimentoChapa', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="areaBlank"
                  rotulo="Área Blank (m²)"
                  value={edicao.areaBlank ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('areaBlank', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="areaLiquida"
                  rotulo="Área Líquida (m²)"
                  value={edicao.areaLiquida ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('areaLiquida', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="pesoSpd"
                  rotulo="Peso (g)"
                  value={edicao.peso ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('peso', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <div className="sm:col-span-3">
                  <Chave
                    rotulo="Ink Jet"
                    marcado={edicao.inkJet ?? false}
                    desabilitado={salvando}
                    aoAlternar={(v) => mudar('inkJet', v)}
                  />
                </div>
              </div>
            </Secao>

            <Secao titulo="Fechamento e Paletização">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="tipoFechamento"
                  rotulo="Fechamento"
                  value={edicao.tipoFechamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoFechamento', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="quantidadeFitilhos"
                  rotulo="Quantidade de Fitilhos"
                  type="number"
                  value={edicao.quantidadeFitilhos ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeFitilhos', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Selecao
                  id="tipoFitilho"
                  rotulo="Fitilhos na"
                  value={edicao.tipoFitilho ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoFitilho', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  <option value="LARGURA">Largura</option>
                  <option value="COMPRIMENTO">Comprimento</option>
                </Selecao>
                <Entrada
                  id="pecasPorUnidade"
                  rotulo="Pacotes de"
                  type="number"
                  value={edicao.pecasPorUnidade ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('pecasPorUnidade', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Entrada
                  id="palete"
                  rotulo="Palete"
                  value={edicao.palete ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('palete', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="alturaMaximaPalete"
                  rotulo="Altura Máx. do Palete (mm)"
                  type="number"
                  value={edicao.alturaMaximaPalete ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('alturaMaximaPalete', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="quantidadeEspelhos"
                  rotulo="Espelho"
                  type="number"
                  value={edicao.quantidadeEspelhos ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeEspelhos', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <div className="grid gap-2 sm:col-span-12 sm:grid-cols-4">
                  {(
                    [
                      ['ehAmarrado', 'Amarrado'],
                      ['ehPaletizado', 'Paletizado'],
                      ['temFacesOpostas', 'Faces Opostas'],
                      ['temCantinho', 'Cantinho'],
                      ['temGradeMadeira', 'Grade de Madeira'],
                      ['temFilmeStrech', 'Filme Strech'],
                      ['clienteUsaRack', 'Cliente usa Rack'],
                      ['overhang', 'Aceita overhang']
                    ] as [keyof Spd, string][]
                  ).map(([campo, rotulo]) => (
                    <Chave
                      key={campo}
                      rotulo={rotulo}
                      marcado={Boolean(edicao[campo])}
                      desabilitado={salvando}
                      aoAlternar={(v) => mudar(campo, v as never)}
                    />
                  ))}
                </div>
              </div>
            </Secao>

            <Secao titulo="Aplicação">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="conteudoEmbalado"
                  rotulo="Conteúdo a ser Embalado"
                  value={edicao.conteudoEmbalado ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('conteudoEmbalado', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="consumoMensal"
                  rotulo="Consumo Mensal (unidades)"
                  type="number"
                  value={edicao.consumoMensal ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('consumoMensal', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="tipoLinhaProducaoCliente"
                  rotulo="Linha de Produção do Cliente"
                  value={edicao.tipoLinhaProducaoCliente ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoLinhaProducaoCliente', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="tipoProdutoPrimario"
                  rotulo="Tipo de Produto Primário"
                  value={edicao.tipoProdutoPrimario ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoProdutoPrimario', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="pesoPrimario"
                  rotulo="Peso do Primário (kg)"
                  value={edicao.pesoPrimario ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('pesoPrimario', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="quantidadePrimarios"
                  rotulo="Quantidade de Primários"
                  type="number"
                  value={edicao.quantidadePrimarios ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadePrimarios', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="suportacaoPrimario"
                  rotulo="Suportação do Primário"
                  value={edicao.suportacaoPrimario ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('suportacaoPrimario', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="tempoEstocagem"
                  rotulo="Tempo de Estocagem (meses)"
                  type="number"
                  value={edicao.tempoEstocagem ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tempoEstocagem', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="umidadeRelativa"
                  rotulo="Umidade Relativa (%)"
                  value={edicao.umidadeRelativa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('umidadeRelativa', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="quantidadeCores"
                  rotulo="Quantidade de Cores"
                  type="number"
                  value={edicao.quantidadeCores ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeCores', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
              </div>
            </Secao>

            <Secao titulo="Comercial">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="precoKGSpd"
                  rotulo="Preço KG (R$)"
                  value={edicao.precoKg ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('precoKg', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="orcamentoSpd"
                  rotulo="Orçamento"
                  value={edicao.orcamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('orcamento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="margemContribuicao"
                  rotulo="Margem de Contribuição (%)"
                  value={edicao.margemContribuicao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('margemContribuicao', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="resultadoSpd"
                  rotulo="Resultado (%)"
                  value={edicao.resultado ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('resultado', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="tipoSavings"
                  rotulo="Savings"
                  value={edicao.tipoSavings ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoSavings', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorSavingSK"
                  rotulo="Valor Saving SK (R$)"
                  value={edicao.valorSavingSk ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('valorSavingSk', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorSavingCliente"
                  rotulo="Valor Saving Cliente (R$)"
                  value={edicao.valorSavingCliente ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('valorSavingCliente', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorEstimadoCliche"
                  rotulo="Valor Estimado Clichê (R$)"
                  value={edicao.valorEstimadoCliche ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('valorEstimadoCliche', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorEstimadoMatriz"
                  rotulo="Valor Estimado Matriz (R$)"
                  value={edicao.valorEstimadoMatriz ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('valorEstimadoMatriz', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Campo rotulo="Informações Adicionais" htmlFor="informacoesAdicionais" className="sm:col-span-12">
                  <textarea
                    id="informacoesAdicionais"
                    rows={2}
                    value={edicao.informacoesAdicionais ?? ''}
                    disabled={salvando}
                    onChange={(e) => mudar('informacoesAdicionais', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Amostras">
              <div className="grid gap-4 sm:grid-cols-3">
                <Entrada
                  id="quantidadeAmostras"
                  rotulo="Quantidade de Amostras"
                  type="number"
                  value={edicao.quantidadeAmostras ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeAmostras', e.target.value)}
                />
                <Entrada
                  id="amostraCep"
                  rotulo="CEP"
                  value={edicao.amostraCep ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('amostraCep', e.target.value)}
                />
                <Entrada
                  id="amostraResponsavel"
                  rotulo="Aos Cuidados"
                  value={edicao.amostraResponsavel ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('amostraResponsavel', e.target.value)}
                />
              </div>
            </Secao>
          </>
        )}
      </Modal>

      {/* Testes físicos */}
      <Modal
        aberto={testes !== null}
        titulo={testes ? `Testes Físicos — SPD ${testes.id}` : ''}
        icone={FlaskConical}
        largura="lg"
        aoFechar={() => setTestes(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setTestes(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvarEnsaios} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {testes && (
          <div className="grid gap-4 sm:grid-cols-2">
            {(
              [
                ['gramaturaReferencial', 'Gramatura Referencial (g/m²)'],
                ['esmagamento', 'Esmagamento (kgf/cm²)'],
                ['ectColuna', 'ECT/Coluna (kgf/cm)'],
                ['mullenReferencial', 'Mullen Referencial (kgf/cm²)'],
                ['espessura', 'Espessura (mm)'],
                ['bctCompressao', 'BCT/Compressão (kgf)'],
                ['cobbInterno', 'COBB Interno (g/m²H₂O)'],
                ['cobbExterno', 'COBB Externo (g/m²H₂O)']
              ] as [keyof Spd, string][]
            ).map(([campo, rotulo]) => (
              <Entrada
                key={campo}
                id={String(campo)}
                rotulo={rotulo}
                value={String(testes[campo] ?? '')}
                disabled={salvando}
                onChange={(e) => setTestes({ ...testes, [campo]: e.target.value })}
              />
            ))}
          </div>
        )}
      </Modal>

      {/* Alterar andamento */}
      <Modal
        aberto={andamento !== null}
        titulo={andamento ? `Andamento — SPD ${andamento.id}` : ''}
        icone={History}
        largura="lg"
        aoFechar={() => setAndamento(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAndamento(null)}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={confirmarAndamento}>
              Alterar andamento
            </Botao>
          </>
        }
      >
        {andamento && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              Andamento atual: <strong>{andamento.statusAndamento}</strong>
            </p>

            <Selecao
              id="novoAndamento"
              rotulo="Novo andamento"
              required
              value={novoAndamento.processo}
              onChange={(e) => setNovoAndamento({ ...novoAndamento, processo: e.target.value })}
            >
              <option value="">Selecione...</option>
              {STATUS_ANDAMENTO_SPD.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <Campo rotulo="Complemento" htmlFor="complementoAndamento">
              <textarea
                id="complementoAndamento"
                rows={2}
                value={novoAndamento.complemento}
                onChange={(e) => setNovoAndamento({ ...novoAndamento, complemento: e.target.value })}
                className={classesControle}
              />
            </Campo>
          </div>
        )}
      </Modal>

      {/* Detalhes com aprovações e andamentos */}
      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `SPD ${detalhe.id} — ${detalhe.referencia}` : ''}
        icone={FileSearch}
        largura="xl"
        aoFechar={() => setDetalhe(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setDetalhe(null)}>
            Fechar
          </Botao>
        }
      >
        {detalhe && (
          <div className="grid gap-4">
            <dl className="grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-4 dark:border-slate-700">
              {[
                ['Cliente', detalhe.nomeCliente],
                ['Planta', nomePlanta(detalhe.codPlanta)],
                ['Vendedor', detalhe.vendedor],
                ['Desenvolvedor', detalhe.desenvolvedor],
                ['Processo', detalhe.tipoSolicitacao],
                ['Estilo', detalhe.estilo || '—'],
                ['Dimensões (mm)', `${detalhe.comprimento} × ${detalhe.largura} × ${detalhe.altura}`],
                ['Estrutura', detalhe.estruturaPapelao],
                ['Preço Kg', detalhe.precoKg || '—'],
                ['Consumo Mensal', detalhe.consumoMensal || '—'],
                ['Status', ROTULO_STATUS[detalhe.status]],
                ['Andamento', detalhe.statusAndamento]
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Aprovações por área
              </p>
              <div className="grid gap-2 sm:grid-cols-3">
                {AREAS.map((area) => {
                  const registro = detalhe.aprovacoes.find((a) => a.area === area)
                  const podeDecidir =
                    !registro && detalhe.status !== 'CANCELADA' && detalhe.status !== 'REJEITADA'

                  return (
                    <div
                      key={area}
                      className="rounded-lg border border-slate-200 p-3 dark:border-slate-700"
                    >
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                        {area}
                      </p>

                      {registro ? (
                        <>
                          <Etiqueta
                            texto={registro.aprovado ? 'Aprovado' : 'Rejeitado'}
                            tom={registro.aprovado ? 'sucesso' : 'perigo'}
                          />
                          <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                            {registro.usuario} &bull; {formatarData(registro.data)}
                          </p>
                          {registro.motivo && (
                            <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                              {registro.motivo}
                            </p>
                          )}
                        </>
                      ) : podeDecidir ? (
                        <div className="mt-2 flex gap-1">
                          <Botao
                            variante="sucesso"
                            icone={ThumbsUp}
                            aoClicar={() => {
                              setMotivoDecisao('')
                              setDecidindo({ spd: detalhe, area, aprovar: true })
                            }}
                          >
                            <span className="sr-only">Aprovar</span>
                          </Botao>
                          <Botao
                            variante="perigo"
                            icone={ThumbsDown}
                            aoClicar={() => {
                              setMotivoDecisao('')
                              setDecidindo({ spd: detalhe, area, aprovar: false })
                            }}
                          >
                            <span className="sr-only">Rejeitar</span>
                          </Botao>
                        </div>
                      ) : (
                        <p className="mt-1 text-xs text-slate-400">Aguardando</p>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Andamento do processo
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.andamentos}
                  chave={(a) => a.processo + a.inicio}
                  mensagemVazia="Nenhum andamento registrado"
                  colunas={[
                    { cabecalho: 'Processo', conteudo: (a) => a.processo },
                    { cabecalho: 'Início', conteudo: (a) => formatarData(a.inicio) },
                    { cabecalho: 'Fim', conteudo: (a) => (a.fim ? formatarData(a.fim) : 'em aberto') },
                    {
                      cabecalho: 'Leadtime',
                      conteudo: (a) => {
                        if (!a.fim) return '—'
                        const dias = Math.round(
                          (new Date(a.fim).getTime() - new Date(a.inicio).getTime()) / 86_400_000
                        )
                        return `${dias} dia(s)`
                      }
                    },
                    { cabecalho: 'Quem', conteudo: (a) => a.quem },
                    { cabecalho: 'Complemento', conteudo: (a) => a.complemento || '—' }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Decisão de área */}
      <Modal
        aberto={decidindo !== null}
        titulo={
          decidindo
            ? `${decidindo.aprovar ? 'Aprovar' : 'Rejeitar'} — ${decidindo.area}`
            : ''
        }
        icone={decidindo?.aprovar ? ThumbsUp : ThumbsDown}
        aoFechar={() => setDecidindo(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setDecidindo(null)}>
              Cancelar
            </Botao>
            <Botao
              variante={decidindo?.aprovar ? 'sucesso' : 'perigo'}
              aoClicar={confirmarDecisao}
            >
              Confirmar
            </Botao>
          </>
        }
      >
        {decidindo && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              SPD {decidindo.spd.id} — {decidindo.spd.nomeCliente}
            </p>

            {decidindo.aprovar ? (
              <Entrada
                id="observacaoAprovacao"
                rotulo="Observação"
                value={motivoDecisao}
                onChange={(e) => setMotivoDecisao(e.target.value)}
                ajuda="Opcional"
              />
            ) : (
              <Selecao
                id="motivoRejeicao"
                rotulo="Motivo da rejeição"
                required
                value={motivoDecisao}
                onChange={(e) => setMotivoDecisao(e.target.value)}
              >
                <option value="">Selecione...</option>
                {MOTIVOS_REJEICAO_SPD.map((motivo) => (
                  <option key={motivo} value={motivo}>
                    {motivo}
                  </option>
                ))}
              </Selecao>
            )}
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={cancelando !== null}
        titulo={cancelando?.status === 'CANCELADA' ? 'Reverter cancelamento' : 'Cancelar SPD'}
        mensagem={
          cancelando
            ? cancelando.status === 'CANCELADA'
              ? `A SPD ${cancelando.id} voltará ao fluxo. Deseja continuar?`
              : `Deseja realmente cancelar a SPD ${cancelando.id}?`
            : ''
        }
        rotuloConfirmar={cancelando?.status === 'CANCELADA' ? 'Reverter' : 'Cancelar SPD'}
        aoConfirmar={confirmarCancelamento}
        aoCancelar={() => setCancelando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
