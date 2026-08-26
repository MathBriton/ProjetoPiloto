import { useCallback, useEffect, useState } from 'react'
import { Boxes, Eraser, FileSearch, History, Pencil, Plus, Save, Search, Trash2 } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  DADOS_LP_VAZIOS,
  FILTROS_LP_VAZIOS,
  MOTIVOS_LOTE_PILOTO,
  STATUS_LOTE_PILOTO,
  TIPOS_ITEM,
  TIPOS_LOTE_PILOTO,
  descricaoMotivoLp,
  descricaoStatusLp,
  excluirLotePiloto,
  listarClientesDesenvolvimento,
  listarLotesPiloto,
  salvarLotePiloto,
  type DadosLotePiloto,
  type FiltrosLotePiloto,
  type LotePiloto as Lote
} from '@/services/desenvolvimento/desenvolvimentoApi'
import { PLANTAS_PROTHEUS, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/** Refatoração de Desenvolvimento/Lote-Piloto-V2. */

const CLIENTES = listarClientesDesenvolvimento()

/** Status de desfecho positivo aparecem em verde; negativos, em vermelho */
const STATUS_POSITIVOS = ['CV', 'ASP']
const STATUS_NEGATIVOS = ['RPC', 'RIR', 'DESC', 'SRV', 'EXC']

function tomStatus(codigo: string) {
  if (STATUS_POSITIVOS.includes(codigo)) return 'sucesso' as const
  if (STATUS_NEGATIVOS.includes(codigo)) return 'perigo' as const
  return 'alerta' as const
}

export function LotePiloto() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosLotePiloto>(FILTROS_LP_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosLotePiloto>(FILTROS_LP_VAZIOS)
  const [lotes, setLotes] = useState<Lote[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<DadosLotePiloto | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [historico, setHistorico] = useState<Lote | null>(null)
  const [excluindo, setExcluindo] = useState<Lote | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setLotes(await listarLotesPiloto(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof DadosLotePiloto>(campo: C, valor: string) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  /** Escolher o cliente preenche código, loja e nome de uma vez */
  function escolherCliente(codigo: string) {
    const cliente = CLIENTES.find((c) => c.cod === codigo)
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            codigoCliente: codigo,
            codigoLoja: cliente?.loja ?? '',
            nomeCliente: cliente?.nome ?? ''
          }
        : atual
    )
  }

  function abrirEdicao(lote: Lote) {
    setErro('')
    setEdicao({
      id: lote.id,
      codigoPlanta: lote.codigoPlanta,
      codigoCliente: lote.codigoCliente,
      codigoLoja: lote.codigoLoja,
      nomeCliente: lote.nomeCliente,
      tipoItem: lote.tipoItem,
      codigoProduto: lote.codigoProduto,
      referencia: lote.referencia,
      quantidadePaletes: String(lote.quantidadePaletes),
      quantidadePedidos: String(lote.quantidadePedidos),
      motivo: lote.motivo,
      precisaLiberacao: lote.precisaLiberacao ? 'S' : 'N',
      tipoLp: lote.tipoLp,
      pedidoKiwi: lote.pedidoKiwi,
      lotePilotoAprovado: lote.lotePilotoAprovado ? 'S' : 'N',
      status: lote.status,
      justificativa: ''
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarLotePiloto(edicao, usuario?.login ?? '')
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

  async function confirmarExclusao() {
    if (!excluindo) return

    const resultado = await excluirLotePiloto(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Lote Piloto"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao(DADOS_LP_VAZIOS)
            }}
          >
            Novo lote piloto
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
            id="filtroPlantaLp"
            rotulo="Filial"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todas as filiais</option>
            {PLANTAS_PROTHEUS.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroStatusLp"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Todos</option>
            {STATUS_LOTE_PILOTO.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroClienteLp"
            rotulo="Cliente"
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
                setFiltros(FILTROS_LP_VAZIOS)
                setFiltrosAplicados(FILTROS_LP_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${lotes.length} lote(s) piloto`} icone={Boxes}>
        <Tabela
          itens={lotes}
          chave={(lote) => lote.id}
          carregando={carregando}
          mensagemVazia="Nenhum lote piloto encontrado"
          colunas={[
            { cabecalho: '#', conteudo: (l) => l.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (l) => nomePlanta(l.codigoPlanta) },
            { cabecalho: 'Data Inclusão', conteudo: (l) => formatarData(l.dataInclusao) },
            { cabecalho: 'Data Status', conteudo: (l) => formatarData(l.dataStatus) },
            {
              cabecalho: 'Status',
              conteudo: (l) => <Etiqueta texto={descricaoStatusLp(l.status)} tom={tomStatus(l.status)} />
            },
            { cabecalho: 'Cliente', conteudo: (l) => l.nomeCliente },
            { cabecalho: 'Produto', conteudo: (l) => l.codigoProduto },
            { cabecalho: 'Qtd. Paletes', conteudo: (l) => l.quantidadePaletes, alinhamento: 'direita' },
            { cabecalho: 'Qtd. Pedidos', conteudo: (l) => l.quantidadePedidos, alinhamento: 'direita' },
            { cabecalho: 'Motivo', conteudo: (l) => descricaoMotivoLp(l.motivo) },
            {
              cabecalho: 'Tipo LP',
              conteudo: (l) => TIPOS_LOTE_PILOTO.find((t) => t.codigo === l.tipoLp)?.descricao ?? l.tipoLp
            },
            { cabecalho: 'Usuário', conteudo: (l) => l.usuario }
          ]}
          acoes={(lote) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Histórico" aoClicar={() => setHistorico(lote)} />
              <BotaoAcao icone={Pencil} titulo="Editar lote piloto" aoClicar={() => abrirEdicao(lote)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir lote piloto"
                tom="perigo"
                aoClicar={() => setExcluindo(lote)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar Lote Piloto ${edicao.id}` : 'Novo Lote Piloto'}
        icone={Boxes}
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

            <Secao titulo="Cliente e Produto">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="codigoPlanta"
                  rotulo="Planta"
                  required
                  value={edicao.codigoPlanta}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoPlanta', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_PROTHEUS.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="codigoCliente"
                  rotulo="Cliente"
                  required
                  value={edicao.codigoCliente}
                  disabled={salvando}
                  onChange={(e) => escolherCliente(e.target.value)}
                  larguraClasse="sm:col-span-6"
                >
                  <option value="">Selecione...</option>
                  {CLIENTES.map((cliente) => (
                    <option key={cliente.cod} value={cliente.cod}>
                      {cliente.cod} — {cliente.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="codigoLoja"
                  rotulo="Loja"
                  maxLength={2}
                  value={edicao.codigoLoja}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoLoja', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />

                <Selecao
                  id="tipoItem"
                  rotulo="Tipo do Item"
                  required
                  value={edicao.tipoItem}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoItem', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_ITEM.map((tipo) => (
                    <option key={tipo.codigo} value={tipo.codigo}>
                      {tipo.descricao}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="codigoProduto"
                  rotulo="Produto"
                  required
                  value={edicao.codigoProduto}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoProduto', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="referencia"
                  rotulo="Referência"
                  value={edicao.referencia}
                  disabled={salvando}
                  onChange={(e) => mudar('referencia', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
              </div>
            </Secao>

            <Secao titulo="Lote">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="quantidadePaletes"
                  rotulo="Qtde. de Palete"
                  type="number"
                  min="1"
                  required
                  value={edicao.quantidadePaletes}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadePaletes', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="quantidadePedidos"
                  rotulo="Qtde. de Pedidos LP"
                  type="number"
                  min="1"
                  required
                  value={edicao.quantidadePedidos}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadePedidos', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="pedidoWRK"
                  rotulo="Pedido KIWI"
                  required
                  value={edicao.pedidoKiwi}
                  disabled={salvando}
                  onChange={(e) => mudar('pedidoKiwi', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Selecao
                  id="tipoLP"
                  rotulo="Tipo de LP"
                  required
                  value={edicao.tipoLp}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoLp', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_LOTE_PILOTO.map((tipo) => (
                    <option key={tipo.codigo} value={tipo.codigo}>
                      {tipo.descricao}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="motivo"
                  rotulo="Motivo"
                  required
                  value={edicao.motivo}
                  disabled={salvando}
                  onChange={(e) => mudar('motivo', e.target.value)}
                  larguraClasse="sm:col-span-6"
                >
                  <option value="">Selecione...</option>
                  {MOTIVOS_LOTE_PILOTO.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.descricao}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="precisaLiberacao"
                  rotulo="Precisa de Liberação"
                  required
                  value={edicao.precisaLiberacao}
                  disabled={salvando}
                  onChange={(e) => mudar('precisaLiberacao', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  <option value="S">Sim</option>
                  <option value="N">Não</option>
                </Selecao>

                <Selecao
                  id="lotePilotoAprovado"
                  rotulo="Lote Piloto Aprovado"
                  required
                  value={edicao.lotePilotoAprovado}
                  disabled={salvando}
                  onChange={(e) => mudar('lotePilotoAprovado', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  <option value="S">Sim</option>
                  <option value="N">Não</option>
                </Selecao>

                <Selecao
                  id="statusLp"
                  rotulo="Status"
                  required
                  value={edicao.status}
                  disabled={salvando}
                  onChange={(e) => mudar('status', e.target.value)}
                  larguraClasse="sm:col-span-6"
                >
                  <option value="">Selecione...</option>
                  {STATUS_LOTE_PILOTO.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.descricao}
                    </option>
                  ))}
                </Selecao>

                <Campo rotulo="Justificativa" htmlFor="justificativaLp" className="sm:col-span-12">
                  <textarea
                    id="justificativaLp"
                    rows={2}
                    value={edicao.justificativa}
                    disabled={salvando}
                    placeholder="Registrada no histórico de justificativas..."
                    onChange={(e) => mudar('justificativa', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>
          </>
        )}
      </Modal>

      <Modal
        aberto={historico !== null}
        titulo={historico ? `Histórico do Lote Piloto ${historico.id}` : ''}
        icone={History}
        largura="lg"
        aoFechar={() => setHistorico(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setHistorico(null)}>
            Fechar
          </Botao>
        }
      >
        {historico && (
          <div className="grid gap-4">
            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Histórico de status
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={historico.historicoStatus}
                  chave={(h) => h.dataInclusao + h.statusPosterior}
                  mensagemVazia="Nenhuma alteração de status registrada"
                  colunas={[
                    { cabecalho: 'Status Anterior', conteudo: (h) => descricaoStatusLp(h.statusAnterior) },
                    { cabecalho: 'Status Posterior', conteudo: (h) => descricaoStatusLp(h.statusPosterior) },
                    { cabecalho: 'Alterado Por', conteudo: (h) => h.alteradoPor },
                    { cabecalho: 'Data', conteudo: (h) => formatarData(h.dataInclusao) }
                  ]}
                />
              </div>
            </div>

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Justificativas
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={historico.justificativas}
                  chave={(j) => j.dataInclusao + j.tipo}
                  mensagemVazia="Nenhuma justificativa registrada"
                  colunas={[
                    { cabecalho: 'Tipo', conteudo: (j) => j.tipo },
                    { cabecalho: 'Justificativa', conteudo: (j) => j.descricao },
                    { cabecalho: 'Quem Incluiu', conteudo: (j) => j.quemIncluiu },
                    { cabecalho: 'Data', conteudo: (j) => formatarData(j.dataInclusao) }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir lote piloto"
        mensagem={excluindo ? `Deseja realmente excluir o lote piloto ${excluindo.id}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
