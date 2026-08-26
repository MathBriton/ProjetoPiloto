import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Eraser, Search, ThumbsDown, ThumbsUp } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  FILTROS_COMERCIAL_VAZIOS,
  MOTIVOS_APROVACAO_COMERCIAL,
  MOTIVOS_REPROVACAO_COMERCIAL,
  decidirComercial,
  descricaoMotivoComercial,
  listarPreCadastros,
  type FiltrosComercial,
  type PreCadastroComercial
} from '@/services/desenvolvimento/desenvolvimentoApi'
import { nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Desenvolvimento/Aprovacao-Comercial-SPD
 * (index_comercial.asp + EspelhoPreCadastro.asp).
 */

const decimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })

function rotuloStatus(status: string) {
  if (status === 'C') return 'Aprovado'
  if (status === 'R') return 'Reprovado'
  return 'Pendente'
}

export function AprovacaoComercialSpd() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosComercial>(FILTROS_COMERCIAL_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosComercial>(FILTROS_COMERCIAL_VAZIOS)
  const [registros, setRegistros] = useState<PreCadastroComercial[]>([])
  const [carregando, setCarregando] = useState(true)

  const [decidindo, setDecidindo] = useState<{ item: PreCadastroComercial; aprovar: boolean } | null>(null)
  const [motivo, setMotivo] = useState('')
  const [salvando, setSalvando] = useState(false)
  const [espelho, setEspelho] = useState<PreCadastroComercial | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setRegistros(await listarPreCadastros(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosComercial>(campo: C, valor: string) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function abrirDecisao(item: PreCadastroComercial, aprovar: boolean) {
    setMotivo('')
    setDecidindo({ item, aprovar })
  }

  async function confirmar() {
    if (!decidindo) return

    setSalvando(true)
    try {
      const resultado = await decidirComercial(
        decidindo.item.id,
        decidindo.aprovar,
        motivo,
        usuario?.login ?? ''
      )
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setDecidindo(null)
        await carregar()
      }
    } finally {
      setSalvando(false)
    }
  }

  const motivosDisponiveis = decidindo?.aprovar
    ? MOTIVOS_APROVACAO_COMERCIAL
    : MOTIVOS_REPROVACAO_COMERCIAL

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Aprovação Comercial — SPD" />

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Entrada
            id="filtroSap"
            rotulo="SAP"
            value={filtros.sap}
            onChange={(e) => mudarFiltro('sap', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroFtOrigem"
            rotulo="FT Origem"
            value={filtros.ftOrigem}
            onChange={(e) => mudarFiltro('ftOrigem', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroClienteComercial"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => mudarFiltro('cliente', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Selecao
            id="filtroStatusComercial"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => mudarFiltro('status', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="T">Status (Todos)</option>
            <option value="">Pendente</option>
            <option value="C">Aprovado</option>
            <option value="R">Reprovado</option>
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
                setFiltros(FILTROS_COMERCIAL_VAZIOS)
                setFiltrosAplicados(FILTROS_COMERCIAL_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${registros.length} pré-cadastro(s)`} icone={ClipboardCheck}>
        <Tabela
          itens={registros}
          chave={(item) => item.id}
          carregando={carregando}
          mensagemVazia="Nenhum pré-cadastro encontrado"
          colunas={[
            { cabecalho: 'SAP', conteudo: (p) => p.sap },
            {
              cabecalho: 'Status',
              conteudo: (p) => (
                <Etiqueta
                  texto={rotuloStatus(p.statusComercial)}
                  tom={
                    p.statusComercial === 'C' ? 'sucesso' : p.statusComercial === 'R' ? 'perigo' : 'alerta'
                  }
                />
              )
            },
            {
              cabecalho: 'Data / Motivo',
              conteudo: (p) =>
                p.statusComercial ? (
                  <span className="text-xs">
                    {formatarData(p.dataStatus)}
                    <span className="block text-slate-400">{descricaoMotivoComercial(p.motivo)}</span>
                  </span>
                ) : (
                  '—'
                )
            },
            { cabecalho: 'Emissão', conteudo: (p) => formatarData(p.dataEmissao) },
            { cabecalho: 'Cliente', conteudo: (p) => p.cliente },
            { cabecalho: 'Vendedor', conteudo: (p) => p.vendedor },
            { cabecalho: 'Orçamento', conteudo: (p) => p.orcamento },
            { cabecalho: 'Estrutura', conteudo: (p) => p.estrutura },
            {
              cabecalho: 'Preço Kg',
              conteudo: (p) => decimal.format(p.precoKg),
              alinhamento: 'direita'
            }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao
                icone={ClipboardCheck}
                titulo="Ver espelho do pré-cadastro"
                aoClicar={() => setEspelho(item)}
              />
              {!item.statusComercial && (
                <>
                  <BotaoAcao
                    icone={ThumbsUp}
                    titulo="Aprovar"
                    aoClicar={() => abrirDecisao(item, true)}
                  />
                  <BotaoAcao
                    icone={ThumbsDown}
                    titulo="Reprovar"
                    tom="perigo"
                    aoClicar={() => abrirDecisao(item, false)}
                  />
                </>
              )}
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={decidindo !== null}
        titulo={decidindo?.aprovar ? 'Aprovar pré-cadastro' : 'Reprovar pré-cadastro'}
        icone={decidindo?.aprovar ? ThumbsUp : ThumbsDown}
        aoFechar={() => setDecidindo(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setDecidindo(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao
              variante={decidindo?.aprovar ? 'sucesso' : 'perigo'}
              aoClicar={confirmar}
              desabilitado={salvando}
            >
              {salvando ? 'Gravando...' : decidindo?.aprovar ? 'Aprovar' : 'Reprovar'}
            </Botao>
          </>
        }
      >
        {decidindo && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              SAP <strong>{decidindo.item.sap}</strong> — {decidindo.item.cliente}
            </p>

            <Selecao
              id="motivoComercial"
              rotulo={decidindo.aprovar ? 'Motivo da aprovação' : 'Motivo da reprovação'}
              required
              value={motivo}
              disabled={salvando}
              onChange={(e) => setMotivo(e.target.value)}
            >
              <option value="">Selecione...</option>
              {motivosDisponiveis.map((item) => (
                <option key={item.codigo} value={item.codigo}>
                  {item.descricao}
                </option>
              ))}
            </Selecao>
          </div>
        )}
      </Modal>

      <Modal
        aberto={espelho !== null}
        titulo={espelho ? `Espelho do pré-cadastro ${espelho.sap}` : ''}
        icone={ClipboardCheck}
        largura="lg"
        aoFechar={() => setEspelho(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setEspelho(null)}>
            Fechar
          </Botao>
        }
      >
        {espelho && (
          <dl className="grid gap-3 sm:grid-cols-3">
            {[
              ['SAP', espelho.sap],
              ['FT Origem', espelho.ftOrigem],
              ['Planta', nomePlanta(espelho.planta)],
              ['Cliente', espelho.cliente],
              ['Vendedor', espelho.vendedor],
              ['Orçamento', espelho.orcamento],
              ['Estrutura', espelho.estrutura],
              ['Preço Kg', decimal.format(espelho.precoKg)],
              ['Emissão', formatarData(espelho.dataEmissao)],
              ['Status', rotuloStatus(espelho.statusComercial)],
              ['Motivo', descricaoMotivoComercial(espelho.motivo) || '—'],
              ['Decidido por', espelho.usuarioStatus || '—']
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
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
