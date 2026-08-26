import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Send } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import { ModalValidacaoPedido } from '@/components/validacoes/ModalValidacaoPedido'
import {
  STATUS_PEDIDO,
  enviarLiberados,
  gravarStatusPedido,
  listarPedidosValidacao,
  rotuloStatus,
  type PedidoValidacao,
  type StatusValidacao
} from '@/services/validacoes/validacoesApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Validacoes/Validar-Pedido
 * (ListaPedidosAValidar.asp + ValidaPedido.asp + GravaStatus.asp).
 */

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })
const decimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })

export function tomStatus(status: StatusValidacao) {
  if (status === 'LI') return 'sucesso' as const
  if (status === 'PENDENTE') return 'neutro' as const
  return 'perigo' as const
}

export function ValidarPedido() {
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
      setPedidos(await listarPedidosValidacao('pedido'))
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
      const resultado = await gravarStatusPedido('pedido', emValidacao.id, status, usuario?.login ?? '')
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
    const resultado = await enviarLiberados('pedido')
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'warning')
    if (resultado.sucesso) await carregar()
  }

  const busca = filtro.trim().toLowerCase()
  const visiveis = busca
    ? pedidos.filter(
        (p) =>
          p.pedido.includes(busca) ||
          p.nome.toLowerCase().includes(busca) ||
          p.cliente.includes(busca) ||
          p.produto.toLowerCase().includes(busca)
      )
    : pedidos

  const liberados = pedidos.filter((p) => p.status === 'LI').length

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Validar Pedido"
        acoes={
          <>
            <CampoPesquisa
              valor={filtro}
              aoMudar={setFiltro}
              placeholder="Pedido, cliente ou produto..."
              larguraClasse="w-64"
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

      <Painel titulo={`${visiveis.length} pedido(s) na fila`} icone={ClipboardCheck}>
        <Tabela
          itens={visiveis}
          chave={(pedido) => pedido.id}
          carregando={carregando}
          mensagemVazia="Nenhum pedido aguardando validação"
          colunas={[
            { cabecalho: 'Pedido', conteudo: (p) => p.pedido },
            { cabecalho: 'Cliente', conteudo: (p) => `${p.cliente}/${p.loja}` },
            { cabecalho: 'Nome', conteudo: (p) => p.nome },
            { cabecalho: 'FT', conteudo: (p) => p.ft },
            { cabecalho: 'Produto', conteudo: (p) => p.produto },
            { cabecalho: 'Tipo', conteudo: (p) => p.tipoProduto },
            { cabecalho: 'Fab.', conteudo: (p) => p.fabrica },
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
            {
              cabecalho: 'Preço/KG',
              conteudo: (p) => decimal.format(p.precoKg),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Preço/Kg Net',
              conteudo: (p) => decimal.format(p.precoKgNet),
              alinhamento: 'direita'
            },
            { cabecalho: 'M²', conteudo: (p) => numero.format(p.m2), alinhamento: 'direita' },
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
            },
            { cabecalho: 'Quem Emitiu', conteudo: (p) => p.quemEmitiu }
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
        origem="pedido"
        opcoes={STATUS_PEDIDO}
        salvando={salvando}
        aoFechar={() => setEmValidacao(null)}
        aoConfirmar={confirmarValidacao}
      />

      <ConfirmDialog
        aberto={confirmandoEnvio}
        titulo="Enviar pedidos liberados"
        mensagem={`Confirma o envio dos ${liberados} pedido(s) liberado(s)? Eles sairão da fila de validação.`}
        rotuloConfirmar="Enviar"
        aoConfirmar={confirmarEnvio}
        aoCancelar={() => setConfirmandoEnvio(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
