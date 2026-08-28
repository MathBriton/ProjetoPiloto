import { useCallback, useEffect, useState } from 'react'
import { ClipboardList, Eraser, Search, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  FILTROS_DIGITADOS_VAZIOS,
  excluirPedidosDigitados,
  listarPedidosDigitados,
  plantasDaRegiao,
  type FiltrosDigitados,
  type PedidoDigitado,
  type RegiaoPedido
} from '@/services/pedidos/pedidosApi'
import { nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Pedidos digitados — compartilhada por Pedidos-Digitados-Brasil e
 * Pedidos-Digitados-PiraUbe. No legado a exclusão tem tela própria
 * (SelecionaPlantaExclusaoPedido.asp); aqui é a seleção em lote da grade.
 */

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })
const moeda = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 })

interface PaginaPedidosDigitadosProps {
  regiao: RegiaoPedido
  titulo: string
}

export function PaginaPedidosDigitados({ regiao, titulo }: PaginaPedidosDigitadosProps) {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const plantas = plantasDaRegiao(regiao)

  const [filtros, setFiltros] = useState<FiltrosDigitados>(FILTROS_DIGITADOS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_DIGITADOS_VAZIOS)
  const [pedidos, setPedidos] = useState<PedidoDigitado[]>([])
  const [carregando, setCarregando] = useState(true)
  const [selecionados, setSelecionados] = useState<Set<number>>(new Set())
  const [confirmando, setConfirmando] = useState(false)

  const carregar = useCallback(async () => {
    setCarregando(true)
    setSelecionados(new Set())
    try {
      setPedidos(await listarPedidosDigitados(regiao, filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [regiao, filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function alternar(id: number) {
    setSelecionados((atual) => {
      const novo = new Set(atual)
      if (novo.has(id)) novo.delete(id)
      else novo.add(id)
      return novo
    })
  }

  const todosMarcados = pedidos.length > 0 && pedidos.every((p) => selecionados.has(p.id))

  async function confirmarExclusao() {
    setConfirmando(false)
    const resultado = await excluirPedidosDigitados([...selecionados])
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo={titulo} />

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Selecao
            id="filtroPlantaDig"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todas</option>
            {plantas.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroPedidoDig"
            rotulo="Pedido"
            value={filtros.pedido}
            onChange={(e) => setFiltros({ ...filtros, pedido: e.target.value })}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroClienteDig"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroDataIniDig"
            rotulo="Emissão de"
            type="date"
            value={filtros.dataInicio}
            onChange={(e) => setFiltros({ ...filtros, dataInicio: e.target.value })}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroDataFimDig"
            rotulo="Até"
            type="date"
            value={filtros.dataFim}
            onChange={(e) => setFiltros({ ...filtros, dataFim: e.target.value })}
            larguraClasse="sm:col-span-2"
          />

          <div className="flex gap-2 sm:col-span-12 sm:justify-end">
            <Botao tipo="submit" icone={Search} desabilitado={carregando}>
              Pesquisar
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              aoClicar={() => {
                setFiltros(FILTROS_DIGITADOS_VAZIOS)
                setFiltrosAplicados(FILTROS_DIGITADOS_VAZIOS)
              }}
            >
              Limpar
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel
        titulo={`${pedidos.length} pedido(s) digitado(s)`}
        icone={ClipboardList}
        acoes={
          <Botao
            variante="perigo"
            icone={Trash2}
            aoClicar={() => setConfirmando(true)}
            desabilitado={selecionados.size === 0}
          >
            Excluir ({selecionados.size})
          </Botao>
        }
      >
        <Tabela
          itens={pedidos}
          chave={(item) => item.id}
          carregando={carregando}
          mensagemVazia="Nenhum pedido digitado encontrado"
          colunas={[
            {
              cabecalho: (
                <input
                  type="checkbox"
                  checked={todosMarcados}
                  onChange={() =>
                    setSelecionados(todosMarcados ? new Set() : new Set(pedidos.map((p) => p.id)))
                  }
                  aria-label="Selecionar todos"
                  className="h-4 w-4 accent-swnet-600"
                />
              ),
              conteudo: (p) => (
                <input
                  type="checkbox"
                  checked={selecionados.has(p.id)}
                  onChange={() => alternar(p.id)}
                  aria-label={`Selecionar pedido ${p.pedido}`}
                  className="h-4 w-4 accent-swnet-600"
                />
              ),
              className: 'w-12'
            },
            { cabecalho: 'Pedido', conteudo: (p) => p.pedido },
            { cabecalho: 'Planta', conteudo: (p) => nomePlanta(p.planta) },
            { cabecalho: 'Cliente', conteudo: (p) => p.codCliente },
            { cabecalho: 'Nome', conteudo: (p) => p.nomeCliente },
            { cabecalho: 'Produto', conteudo: (p) => p.codigoProduto },
            { cabecalho: 'Referência', conteudo: (p) => p.referencia },
            {
              cabecalho: 'Quant.',
              conteudo: (p) => numero.format(p.quantidade),
              alinhamento: 'direita'
            },
            { cabecalho: 'Preço', conteudo: (p) => moeda.format(p.preco), alinhamento: 'direita' },
            { cabecalho: 'Emissão', conteudo: (p) => formatarData(p.dataEmissao) },
            { cabecalho: 'Entrega', conteudo: (p) => formatarData(p.dataEntrega) },
            { cabecalho: 'Emitido por', conteudo: (p) => p.emitidoPor }
          ]}
        />
      </Painel>

      <ConfirmDialog
        aberto={confirmando}
        titulo="Excluir pedidos digitados"
        mensagem={`${selecionados.size} pedido(s) serão excluídos. Deseja continuar?`}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
