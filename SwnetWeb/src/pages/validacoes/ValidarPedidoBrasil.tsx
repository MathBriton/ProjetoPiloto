import { useCallback, useEffect, useState } from 'react'
import { ArrowLeft, Building2, ClipboardCheck, Send } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import { ModalValidacaoPedido } from '@/components/validacoes/ModalValidacaoPedido'
import { tomStatus } from './ValidarPedido'
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
 * Refatoração de Validacoes/Validar-Pedido-Brasil.
 *
 * O legado começa por SelecionaPlantaValidarPedido.asp (Bento, Maranguape
 * ou Saquarema) e só então abre a lista, porque cada planta tem sua própria
 * tela de detalhes. Aqui a seleção de planta é o primeiro passo da mesma tela.
 */

/** Cores por planta, como no radio group do legado */
const PLANTAS = [
  { codigo: '05', nome: 'Bento Gonçalves', classe: 'text-blue-700 dark:text-blue-400' },
  { codigo: '06', nome: 'Maranguape', classe: 'text-amber-700 dark:text-amber-400' },
  { codigo: '07', nome: 'Saquarema', classe: 'text-green-700 dark:text-green-400' }
]

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })
const decimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })

export function ValidarPedidoBrasil() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [planta, setPlanta] = useState('')
  const [pedidos, setPedidos] = useState<PedidoValidacao[]>([])
  const [carregando, setCarregando] = useState(false)
  const [filtro, setFiltro] = useState('')
  const [emValidacao, setEmValidacao] = useState<PedidoValidacao | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [confirmandoEnvio, setConfirmandoEnvio] = useState(false)

  const carregar = useCallback(async () => {
    if (!planta) return

    setCarregando(true)
    try {
      setPedidos(await listarPedidosValidacao('brasil', planta))
    } finally {
      setCarregando(false)
    }
  }, [planta])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function confirmarValidacao(status: StatusValidacao) {
    if (!emValidacao) return

    setSalvando(true)
    try {
      const resultado = await gravarStatusPedido('brasil', emValidacao.id, status, usuario?.login ?? '')
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
    const resultado = await enviarLiberados('brasil')
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
  const plantaAtual = PLANTAS.find((p) => p.codigo === planta)

  if (!planta) {
    return (
      <div className="space-y-4">
        <TituloPagina titulo="Validar Pedido Brasil" />

        <div className="mx-auto max-w-2xl">
          <Painel titulo="Selecione a planta" icone={Building2}>
            <div className="grid gap-3 sm:grid-cols-3">
              {PLANTAS.map((item) => (
                <button
                  key={item.codigo}
                  type="button"
                  onClick={() => setPlanta(item.codigo)}
                  className="flex flex-col items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 py-8
                    text-center shadow-sm transition-all hover:-translate-y-0.5 hover:border-swnet-300
                    hover:shadow-md dark:border-slate-800 dark:bg-slate-900"
                >
                  <Building2 size={30} strokeWidth={1.3} className={item.classe} />
                  <span className={`text-sm font-semibold ${item.classe}`}>{item.nome}</span>
                  <span className="text-xs text-slate-400">Planta {item.codigo}</span>
                </button>
              ))}
            </div>
          </Painel>
        </div>

        <ToastStack toasts={toasts} aoFechar={fecharToast} />
      </div>
    )
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo={`Validar Pedido Brasil — ${plantaAtual?.nome ?? planta}`}
        acoes={
          <>
            <Botao variante="neutro" icone={ArrowLeft} aoClicar={() => setPlanta('')}>
              Trocar planta
            </Botao>
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

      <Painel titulo={`${visiveis.length} pedido(s) na fila`} icone={ClipboardCheck}>
        <Tabela
          itens={visiveis}
          chave={(pedido) => pedido.id}
          carregando={carregando}
          mensagemVazia="Nenhum pedido aguardando validação nesta planta"
          colunas={[
            { cabecalho: '#', conteudo: (p) => p.id, className: 'w-20' },
            { cabecalho: 'Pedido', conteudo: (p) => p.pedido },
            { cabecalho: 'Cliente', conteudo: (p) => `${p.cliente}/${p.loja}` },
            { cabecalho: 'Nome', conteudo: (p) => p.nome },
            { cabecalho: 'FT', conteudo: (p) => p.ft },
            { cabecalho: 'Produto', conteudo: (p) => p.produto },
            { cabecalho: 'Tipo Produto', conteudo: (p) => p.tipoProduto },
            { cabecalho: 'Tipo Onda', conteudo: (p) => p.tipoOnda || '—' },
            {
              cabecalho: 'Quant.',
              conteudo: (p) => numero.format(p.quantidade),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Preço Net',
              conteudo: (p) => decimal.format(p.preco),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Preço KG Net',
              conteudo: (p) => decimal.format(p.precoKgNet),
              alinhamento: 'direita'
            },
            { cabecalho: 'M²', conteudo: (p) => numero.format(p.m2), alinhamento: 'direita' },
            { cabecalho: 'Emissão', conteudo: (p) => formatarData(p.emissao) },
            { cabecalho: 'Entrega', conteudo: (p) => formatarData(p.entrega) },
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
        origem="brasil"
        opcoes={STATUS_PEDIDO}
        salvando={salvando}
        aoFechar={() => setEmValidacao(null)}
        aoConfirmar={confirmarValidacao}
      />

      <ConfirmDialog
        aberto={confirmandoEnvio}
        titulo="Enviar pedidos liberados"
        mensagem={`Confirma o envio dos ${liberados} pedido(s) liberado(s)?`}
        rotuloConfirmar="Enviar"
        aoConfirmar={confirmarEnvio}
        aoCancelar={() => setConfirmandoEnvio(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
