import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Files, Send } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import { ModalValidacaoPedido } from '@/components/validacoes/ModalValidacaoPedido'
import { tomStatus } from './ValidarPedido'
import {
  STATUS_PAPEL,
  enviarLiberados,
  gravarStatusPedido,
  listarPedidosValidacao,
  rotuloStatus,
  type PedidoValidacao,
  type StatusValidacao
} from '@/services/validacoes/validacoesApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Validacoes/Papel (pa.asp + detalha_pa.asp +
 * gravastatus_pa.asp + envia_pa.asp).
 *
 * Mesma mecânica do Validar Pedido, mas com menos colunas e só três
 * status possíveis (preço, condição de pagamento e liberado).
 */

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })
const decimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })

export function ValidacaoPapel() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [pedidos, setPedidos] = useState<PedidoValidacao[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtro, setFiltro] = useState('')
  const [emValidacao, setEmValidacao] = useState<PedidoValidacao | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [confirmandoEnvio, setConfirmandoEnvio] = useState(false)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setPedidos(await listarPedidosValidacao('papel'))
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function confirmarValidacao(status: StatusValidacao) {
    if (!emValidacao) return

    setSalvando(true)
    try {
      const resultado = await gravarStatusPedido('papel', emValidacao.id, status, usuario?.login ?? '')
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setEmValidacao(null)
        await carregar()
      }
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarEnvio() {
    setConfirmandoEnvio(false)
    const resultado = await enviarLiberados('papel')
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'warning')
    if (resultado.sucesso) await carregar()
  }

  const busca = filtro.trim().toLowerCase()
  const visiveis = busca
    ? pedidos.filter(
        (p) => p.pedido.includes(busca) || p.nome.toLowerCase().includes(busca) || p.cliente.includes(busca)
      )
    : pedidos

  const liberados = pedidos.filter((p) => p.status === 'LI').length

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Validação de Pedidos de Papel"
        acoes={
          <>
            <CampoPesquisa
              valor={filtro}
              aoMudar={setFiltro}
              placeholder="Pedido ou cliente..."
              larguraClasse="w-56"
            />
            <Botao
              variante="sucesso"
              icone={Send}
              aoClicar={() => setConfirmandoEnvio(true)}
              desabilitado={liberados === 0}
            >
              Enviar liberados ({liberados})
            </Botao>
          </>
        }
      />

      <Painel titulo={`${visiveis.length} pedido(s) de papel`} icone={Files}>
        <Tabela
          itens={visiveis}
          chave={(pedido) => pedido.id}
          carregando={carregando}
          mensagemVazia="Nenhum pedido de papel aguardando validação"
          colunas={[
            { cabecalho: 'Pedido', conteudo: (p) => p.pedido },
            { cabecalho: 'Cliente', conteudo: (p) => p.cliente },
            { cabecalho: 'Loja', conteudo: (p) => p.loja },
            { cabecalho: 'Nome', conteudo: (p) => p.nome },
            { cabecalho: 'FT', conteudo: (p) => p.ft },
            { cabecalho: 'Produto', conteudo: (p) => p.produto },
            {
              cabecalho: 'Quant.',
              conteudo: (p) => numero.format(p.quantidade),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Preço',
              conteudo: (p) => (
                <span className={p.preco < p.ultimoPreco ? 'font-semibold text-amber-600 dark:text-amber-400' : ''}>
                  {decimal.format(p.preco)}
                </span>
              ),
              alinhamento: 'direita'
            },
            { cabecalho: 'Emissão', conteudo: (p) => formatarData(p.emissao) },
            { cabecalho: 'Entrega', conteudo: (p) => formatarData(p.entrega) },
            {
              cabecalho: 'Obs.',
              conteudo: (p) => (
                <span className="block max-w-40 truncate text-xs" title={p.observacao}>
                  {p.observacao || '—'}
                </span>
              )
            },
            {
              cabecalho: 'Status',
              conteudo: (p) => <Etiqueta texto={rotuloStatus(p.status)} tom={tomStatus(p.status)} />
            },
            {
              cabecalho: 'Bloqueio',
              conteudo: (p) => <span className="text-xs">{p.bloqueio || '—'}</span>
            }
          ]}
          acoes={(pedido) => (
            <BotaoAcao
              icone={ClipboardCheck}
              titulo="Validar pedido"
              aoClicar={() => setEmValidacao(pedido)}
            />
          )}
        />
      </Painel>

      <ModalValidacaoPedido
        pedido={emValidacao}
        origem="papel"
        opcoes={STATUS_PAPEL}
        salvando={salvando}
        aoFechar={() => setEmValidacao(null)}
        aoConfirmar={confirmarValidacao}
      />

      <ConfirmDialog
        aberto={confirmandoEnvio}
        titulo="Enviar pedidos liberados"
        mensagem={`Confirma o envio de todos os ${liberados} pedido(s) liberado(s)?`}
        rotuloConfirmar="Enviar"
        aoConfirmar={confirmarEnvio}
        aoCancelar={() => setConfirmandoEnvio(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
