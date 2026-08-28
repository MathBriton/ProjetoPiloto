import { useState } from 'react'
import { AlertTriangle, ClipboardList, Package, Plus, Search, Send, Trash2, Truck } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Chave, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { Tabela } from '@/components/comuns/Tabela'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  CODIGOS_TES,
  CONDICOES_PAGAMENTO_PEDIDO,
  EMISSAO_VAZIA,
  ESTRUTURAS_PEDIDO,
  GRAMATURAS,
  INCOTERMS,
  ITEM_PEDIDO_VAZIO,
  MODAIS_FRETE,
  MOEDAS,
  TIPOS_FRETE_PEDIDO,
  TIPOS_ITEM_PEDIDO,
  TUBETES,
  buscarClientesPedido,
  buscarProdutosPedido,
  emitirPedido,
  plantasDaRegiao,
  validarItensPedido,
  type AvisoPedido,
  type ClientePedido,
  type DadosEmissaoPedido,
  type ItemPedido,
  type ProdutoPedido,
  type RegiaoPedido
} from '@/services/pedidos/pedidosApi'

/**
 * Emissão de pedido — compartilhada por Pedidos-Brasil-v2 e PedidosPiraUbe.
 *
 * Os campos do item mudam conforme o tipo: caixa e chapa pedem dimensional
 * e vincos; bobina (papel) troca isso pelos campos de gramatura, diâmetro,
 * largura e tubete.
 */

const moeda = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 })

interface PaginaEmissaoPedidoProps {
  regiao: RegiaoPedido
  titulo: string
}

export function PaginaEmissaoPedido({ regiao, titulo }: PaginaEmissaoPedidoProps) {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const plantas = plantasDaRegiao(regiao)

  const [dados, setDados] = useState<DadosEmissaoPedido>(EMISSAO_VAZIA)
  const [item, setItem] = useState<ItemPedido>(ITEM_PEDIDO_VAZIO)
  const [avisos, setAvisos] = useState<AvisoPedido[]>([])
  const [emitindo, setEmitindo] = useState(false)

  const [buscaAberta, setBuscaAberta] = useState<'cliente' | 'remessa' | 'produto' | null>(null)
  const [termo, setTermo] = useState('')
  const [resultadosCliente, setResultadosCliente] = useState<ClientePedido[]>([])
  const [resultadosProduto, setResultadosProduto] = useState<ProdutoPedido[]>([])
  const [buscando, setBuscando] = useState(false)

  function mudar<C extends keyof DadosEmissaoPedido>(campo: C, valor: DadosEmissaoPedido[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  function mudarItem<C extends keyof ItemPedido>(campo: C, valor: ItemPedido[C]) {
    setItem((atual) => ({ ...atual, [campo]: valor }))
  }

  async function abrirBusca(tipo: 'cliente' | 'remessa' | 'produto') {
    setBuscaAberta(tipo)
    setTermo('')
    setBuscando(true)
    try {
      if (tipo === 'produto') setResultadosProduto(await buscarProdutosPedido('', item.tipoItem))
      else setResultadosCliente(await buscarClientesPedido(''))
    } finally {
      setBuscando(false)
    }
  }

  async function pesquisar() {
    setBuscando(true)
    try {
      if (buscaAberta === 'produto') setResultadosProduto(await buscarProdutosPedido(termo, item.tipoItem))
      else setResultadosCliente(await buscarClientesPedido(termo))
    } finally {
      setBuscando(false)
    }
  }

  function selecionarCliente(cliente: ClientePedido) {
    if (buscaAberta === 'remessa') {
      setDados((atual) => ({
        ...atual,
        codClienteRemessa: cliente.codigo,
        codLojaRemessa: cliente.loja,
        nomeClienteRemessa: cliente.nome
      }))
    } else {
      setDados((atual) => ({
        ...atual,
        codCliente: cliente.codigo,
        codLoja: cliente.loja,
        nomeCliente: cliente.nome,
        saldoAparas: String(cliente.saldoAparas),
        // A condição de pagamento vem do cadastro do cliente
        condicaoPagamento: atual.condicaoPagamento || cliente.condicaoPagamento
      }))
    }
    setBuscaAberta(null)
  }

  function selecionarProduto(produto: ProdutoPedido) {
    setItem((atual) => ({
      ...atual,
      codigoProduto: produto.codigo,
      referencia: produto.referencia,
      estrutura: produto.estrutura,
      valorUnitario: atual.valorUnitario || String(produto.valorUnitario)
    }))
    setBuscaAberta(null)
  }

  function adicionarItem() {
    if (!item.codigoProduto) {
      mostrarToast('Selecione o produto antes de adicionar.', 'warning')
      return
    }
    if (!Number(item.quantidade)) {
      mostrarToast('Informe a quantidade.', 'warning')
      return
    }
    if (!item.dataEntrega) {
      mostrarToast('Informe a data de entrega.', 'warning')
      return
    }

    mudar('itens', [...dados.itens, item])
    setItem(ITEM_PEDIDO_VAZIO)
    setAvisos([])
  }

  async function validar() {
    if (dados.itens.length === 0) {
      mostrarToast('Inclua ao menos um item.', 'warning')
      return
    }

    const encontrados = await validarItensPedido(dados)
    setAvisos(encontrados)
    mostrarToast(
      encontrados.length === 0 ? 'Nenhuma inconsistência encontrada.' : `${encontrados.length} aviso(s).`,
      encontrados.some((a) => a.bloqueia) ? 'danger' : encontrados.length ? 'warning' : 'success'
    )
  }

  async function emitir() {
    setEmitindo(true)
    try {
      const resultado = await emitirPedido(dados, regiao, usuario?.login ?? '')
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setDados(EMISSAO_VAZIA)
        setItem(ITEM_PEDIDO_VAZIO)
        setAvisos([])
      }
    } finally {
      setEmitindo(false)
    }
  }

  const ehBobina = item.tipoItem === 'PA'
  const total = Number(item.quantidade || 0) * Number(item.valorUnitario || 0)
  const totalPedido = dados.itens.reduce(
    (soma, linha) => soma + Number(linha.quantidade || 0) * Number(linha.valorUnitario || 0),
    0
  )

  return (
    <div className="space-y-4">
      <TituloPagina titulo={titulo} />

      <Painel titulo="Cabeçalho do Pedido" icone={ClipboardList}>
        <Secao titulo="Cliente">
          <div className="grid items-end gap-4 sm:grid-cols-12">
            <Selecao
              id="plantaPedido"
              rotulo="Planta"
              required
              value={dados.planta}
              disabled={emitindo}
              onChange={(e) => mudar('planta', e.target.value)}
              larguraClasse="sm:col-span-3"
            >
              <option value="">Selecione...</option>
              {plantas.map((planta) => (
                <option key={planta.codigo} value={planta.codigo}>
                  {planta.nome}
                </option>
              ))}
            </Selecao>

            <Campo rotulo="Cód. Cliente" obrigatorio className="sm:col-span-3">
              <div className="flex">
                <input
                  type="text"
                  readOnly
                  value={dados.codCliente}
                  className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
                />
                <button
                  type="button"
                  onClick={() => abrirBusca('cliente')}
                  disabled={emitindo}
                  aria-label="Pesquisar cliente"
                  className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                    hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
                >
                  <Search size={15} />
                </button>
              </div>
            </Campo>

            <Entrada
              id="codLojaPedido"
              rotulo="Cód. Loja"
              value={dados.codLoja}
              disabled
              larguraClasse="sm:col-span-1"
            />
            <Entrada
              id="nomeClientePedido"
              rotulo="Nome Cliente"
              value={dados.nomeCliente}
              disabled
              larguraClasse="sm:col-span-5"
            />

            <Selecao
              id="tesPedido"
              rotulo="TES"
              required
              value={dados.tes}
              disabled={emitindo}
              onChange={(e) => mudar('tes', e.target.value)}
              larguraClasse="sm:col-span-2"
            >
              <option value="">Selecione...</option>
              {CODIGOS_TES.map((tes) => (
                <option key={tes} value={tes}>
                  {tes}
                </option>
              ))}
            </Selecao>

            <Campo rotulo="Cód. Cliente Remessa" className="sm:col-span-3">
              <div className="flex">
                <input
                  type="text"
                  readOnly
                  value={dados.codClienteRemessa}
                  className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
                />
                <button
                  type="button"
                  onClick={() => abrirBusca('remessa')}
                  disabled={emitindo}
                  aria-label="Pesquisar cliente de remessa"
                  className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                    hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
                >
                  <Search size={15} />
                </button>
              </div>
            </Campo>

            <Entrada
              id="lojaRemessa"
              rotulo="Loja Remessa"
              value={dados.codLojaRemessa}
              disabled
              larguraClasse="sm:col-span-1"
            />
            <Entrada
              id="nomeRemessa"
              rotulo="Nome Cliente Remessa"
              value={dados.nomeClienteRemessa}
              disabled
              larguraClasse="sm:col-span-4"
            />
            <Entrada
              id="saldoAparas"
              rotulo="Saldo Aparas"
              value={dados.saldoAparas}
              disabled
              larguraClasse="sm:col-span-2"
            />
          </div>
        </Secao>

        <Secao titulo="Condições Comerciais" icone={Truck}>
          <div className="grid gap-4 sm:grid-cols-12">
            <Selecao
              id="condPagtoPedido"
              rotulo="Condição de Pagamento"
              required
              value={dados.condicaoPagamento}
              disabled={emitindo}
              onChange={(e) => mudar('condicaoPagamento', e.target.value)}
              larguraClasse="sm:col-span-3"
            >
              <option value="">Selecione...</option>
              {CONDICOES_PAGAMENTO_PEDIDO.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <Selecao
              id="moedaPedido"
              rotulo="Moeda"
              value={dados.moeda}
              disabled={emitindo}
              onChange={(e) => mudar('moeda', e.target.value)}
              larguraClasse="sm:col-span-2"
            >
              {MOEDAS.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <Selecao
              id="tipoFretePedido"
              rotulo="Tipo de Frete"
              value={dados.tipoFrete}
              disabled={emitindo}
              onChange={(e) => mudar('tipoFrete', e.target.value)}
              larguraClasse="sm:col-span-2"
            >
              <option value="">Selecione...</option>
              {TIPOS_FRETE_PEDIDO.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <Selecao
              id="modalFrete"
              rotulo="Modal de frete"
              value={dados.modalFrete}
              disabled={emitindo}
              onChange={(e) => mudar('modalFrete', e.target.value)}
              larguraClasse="sm:col-span-2"
            >
              <option value="">Selecione...</option>
              {MODAIS_FRETE.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <Selecao
              id="incoterms"
              rotulo="Incoterms"
              value={dados.incoterms}
              disabled={emitindo}
              onChange={(e) => mudar('incoterms', e.target.value)}
              larguraClasse="sm:col-span-3"
            >
              <option value="">Selecione...</option>
              {INCOTERMS.map((item) => (
                <option key={item} value={item}>
                  {item}
                </option>
              ))}
            </Selecao>

            <div className="sm:col-span-12">
              <Chave
                rotulo="Pedido de Lote Piloto"
                marcado={dados.lotePiloto}
                desabilitado={emitindo}
                aoAlternar={(v) => mudar('lotePiloto', v)}
              />
            </div>
          </div>
        </Secao>
      </Painel>

      <Painel
        titulo="Adicionar Item"
        icone={Package}
        acoes={
          <Botao icone={Plus} aoClicar={adicionarItem} desabilitado={emitindo}>
            Adicionar ao pedido
          </Botao>
        }
      >
        <div className="grid gap-4 sm:grid-cols-12">
          <Selecao
            id="tipoItemPedido"
            rotulo="Tipo do Item"
            required
            value={item.tipoItem}
            disabled={emitindo}
            onChange={(e) => {
              mudarItem('tipoItem', e.target.value)
              mudarItem('codigoProduto', '')
              mudarItem('referencia', '')
            }}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Selecione...</option>
            {TIPOS_ITEM_PEDIDO.map((tipo) => (
              <option key={tipo.codigo} value={tipo.codigo}>
                {tipo.descricao}
              </option>
            ))}
          </Selecao>

          <Campo rotulo="Produto" obrigatorio className="sm:col-span-3">
            <div className="flex">
              <input
                type="text"
                readOnly
                value={item.codigoProduto}
                className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
              />
              <button
                type="button"
                onClick={() => abrirBusca('produto')}
                disabled={emitindo}
                aria-label="Pesquisar produto"
                className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                  hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
              >
                <Search size={15} />
              </button>
            </div>
          </Campo>

          <Entrada
            id="referenciaItem"
            rotulo="Referência"
            value={item.referencia}
            disabled
            larguraClasse="sm:col-span-4"
          />

          <Selecao
            id="estruturaItem"
            rotulo="Estrutura"
            value={item.estrutura}
            disabled={emitindo || ehBobina}
            onChange={(e) => mudarItem('estrutura', e.target.value)}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Selecione...</option>
            {ESTRUTURAS_PEDIDO.map((estrutura) => (
              <option key={estrutura} value={estrutura}>
                {estrutura}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="quantidadeItem"
            rotulo="Quantidade"
            type="number"
            min="1"
            required
            value={item.quantidade}
            disabled={emitindo}
            onChange={(e) => mudarItem('quantidade', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="valorUnitarioItem"
            rotulo="Valor Unitário"
            type="number"
            step="0.0001"
            value={item.valorUnitario}
            disabled={emitindo}
            onChange={(e) => mudarItem('valorUnitario', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="totalItem"
            rotulo="Total"
            value={total ? moeda.format(total) : ''}
            disabled
            larguraClasse="sm:col-span-2"
          />

          {ehBobina ? (
            <>
              <Selecao
                id="gramaturaBobina"
                rotulo="Gramatura Bobina"
                value={item.gramaturaBobina}
                disabled={emitindo}
                onChange={(e) => mudarItem('gramaturaBobina', e.target.value)}
                larguraClasse="sm:col-span-2"
              >
                <option value="">Selecione...</option>
                {GRAMATURAS.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </Selecao>
              <Entrada
                id="diametroBobina"
                rotulo="Diâmetro Bobina"
                value={item.diametroBobina}
                disabled={emitindo}
                onChange={(e) => mudarItem('diametroBobina', e.target.value)}
                larguraClasse="sm:col-span-2"
              />
              <Entrada
                id="larguraBobina"
                rotulo="Largura Bobina"
                value={item.larguraBobina}
                disabled={emitindo}
                onChange={(e) => mudarItem('larguraBobina', e.target.value)}
                larguraClasse="sm:col-span-2"
              />
              <Selecao
                id="tubeteBobina"
                rotulo="Tubete"
                value={item.tubeteBobina}
                disabled={emitindo}
                onChange={(e) => mudarItem('tubeteBobina', e.target.value)}
                larguraClasse="sm:col-span-2"
              >
                <option value="">Selecione...</option>
                {TUBETES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </Selecao>
            </>
          ) : (
            <>
              <Entrada
                id="larguraItem"
                rotulo="Largura"
                value={item.largura}
                disabled={emitindo}
                onChange={(e) => mudarItem('largura', e.target.value)}
                larguraClasse="sm:col-span-2"
              />
              <Entrada
                id="comprimentoItem"
                rotulo="Comprimento"
                value={item.comprimento}
                disabled={emitindo}
                onChange={(e) => mudarItem('comprimento', e.target.value)}
                larguraClasse="sm:col-span-2"
              />
              {([1, 2, 3, 4, 5] as const).map((numero) => (
                <Entrada
                  key={numero}
                  id={`vinco${numero}`}
                  rotulo={`Vinco ${numero}`}
                  value={item[`vinco${numero}` as keyof ItemPedido] as string}
                  disabled={emitindo}
                  onChange={(e) => mudarItem(`vinco${numero}` as keyof ItemPedido, e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
              ))}
            </>
          )}

          <Entrada
            id="dataEntregaItem"
            rotulo="Data de Entrega"
            type="date"
            required
            value={item.dataEntrega}
            disabled={emitindo}
            onChange={(e) => mudarItem('dataEntrega', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="dataPrevistaItem"
            rotulo="Data Prevista de Entrega"
            type="date"
            value={item.dataPrevistaEntrega}
            disabled={emitindo}
            onChange={(e) => mudarItem('dataPrevistaEntrega', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="pedidoCompraItem"
            rotulo="Pedido de Compra"
            value={item.pedidoCompra}
            disabled={emitindo}
            onChange={(e) => mudarItem('pedidoCompra', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="itemPedidoCompra"
            rotulo="Item Pedido"
            value={item.itemPedido}
            disabled={emitindo}
            onChange={(e) => mudarItem('itemPedido', e.target.value)}
            larguraClasse="sm:col-span-3"
          />

          <Entrada
            id="observacaoItem"
            rotulo="Observação"
            value={item.observacao}
            disabled={emitindo}
            onChange={(e) => mudarItem('observacao', e.target.value)}
            larguraClasse="sm:col-span-12"
          />
        </div>
      </Painel>

      <Painel
        titulo={`Itens do Pedido (${dados.itens.length})`}
        icone={Package}
        acoes={
          <>
            <span className="text-sm font-medium text-slate-600 dark:text-slate-300">
              Total: {moeda.format(totalPedido)}
            </span>
            <Botao variante="neutro" icone={AlertTriangle} aoClicar={validar} desabilitado={emitindo}>
              Validar
            </Botao>
            <Botao
              variante="sucesso"
              icone={Send}
              aoClicar={emitir}
              desabilitado={emitindo || dados.itens.length === 0}
            >
              {emitindo ? 'Emitindo...' : 'Emitir pedido'}
            </Botao>
          </>
        }
      >
        {avisos.length > 0 && (
          <div className="mb-4 grid gap-2">
            {avisos.map((aviso, indice) => (
              <p
                key={indice}
                className={`flex items-start gap-2 rounded-lg px-3 py-2 text-sm ${
                  aviso.bloqueia
                    ? 'bg-red-50 text-red-700 dark:bg-red-500/10 dark:text-red-400'
                    : 'bg-amber-50 text-amber-800 dark:bg-amber-500/10 dark:text-amber-400'
                }`}
              >
                <AlertTriangle size={15} className="mt-0.5 shrink-0" />
                <span className="flex-1">{aviso.mensagem}</span>
                <Etiqueta texto={aviso.bloqueia ? 'Bloqueia' : 'Aviso'} tom={aviso.bloqueia ? 'perigo' : 'alerta'} />
              </p>
            ))}
          </div>
        )}

        <Tabela
          itens={dados.itens}
          chave={(linha) => dados.itens.indexOf(linha)}
          mensagemVazia="Nenhum item adicionado ao pedido"
          colunas={[
            {
              cabecalho: 'Tipo',
              conteudo: (i) => TIPOS_ITEM_PEDIDO.find((t) => t.codigo === i.tipoItem)?.descricao ?? i.tipoItem
            },
            { cabecalho: 'Produto', conteudo: (i) => i.codigoProduto },
            { cabecalho: 'Referência', conteudo: (i) => i.referencia },
            { cabecalho: 'Estrutura', conteudo: (i) => i.estrutura || '—' },
            { cabecalho: 'Quantidade', conteudo: (i) => i.quantidade, alinhamento: 'direita' },
            {
              cabecalho: 'Valor Unit.',
              conteudo: (i) => moeda.format(Number(i.valorUnitario || 0)),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Total',
              conteudo: (i) => moeda.format(Number(i.quantidade || 0) * Number(i.valorUnitario || 0)),
              alinhamento: 'direita'
            },
            { cabecalho: 'Entrega', conteudo: (i) => i.dataEntrega },
            { cabecalho: 'Pedido Compra', conteudo: (i) => i.pedidoCompra || '—' }
          ]}
          acoes={(linha) => (
            <button
              type="button"
              onClick={() => mudar('itens', dados.itens.filter((outro) => outro !== linha))}
              disabled={emitindo}
              aria-label="Remover item"
              className="rounded-lg p-1.5 text-slate-400 hover:bg-red-50 hover:text-red-600 dark:hover:bg-slate-800"
            >
              <Trash2 size={14} />
            </button>
          )}
        />
      </Painel>

      <Modal
        aberto={buscaAberta !== null}
        titulo={buscaAberta === 'produto' ? 'Pesquisar Produto' : 'Pesquisar Cliente'}
        icone={Search}
        largura="lg"
        aoFechar={() => setBuscaAberta(null)}
      >
        <div className="grid gap-4">
          <div className="flex items-end gap-2">
            <Entrada
              id="termoBuscaPedido"
              rotulo={buscaAberta === 'produto' ? 'Código ou referência' : 'Código ou nome'}
              value={termo}
              disabled={buscando}
              onChange={(e) => setTermo(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  pesquisar()
                }
              }}
              larguraClasse="flex-1"
            />
            <Botao icone={Search} aoClicar={pesquisar} desabilitado={buscando}>
              Buscar
            </Botao>
          </div>

          {buscaAberta === 'produto' ? (
            <Tabela
              itens={resultadosProduto}
              chave={(p) => p.codigo}
              carregando={buscando}
              colunas={[
                { cabecalho: 'Código', conteudo: (p) => p.codigo },
                { cabecalho: 'Referência', conteudo: (p) => p.referencia },
                { cabecalho: 'Estrutura', conteudo: (p) => p.estrutura || '—' },
                {
                  cabecalho: 'Valor',
                  conteudo: (p) => moeda.format(p.valorUnitario),
                  alinhamento: 'direita'
                }
              ]}
              acoes={(p) => (
                <button
                  type="button"
                  onClick={() => selecionarProduto(p)}
                  className="rounded-lg px-2 py-1 text-xs font-medium text-swnet-600 hover:bg-swnet-50
                    dark:text-swnet-400 dark:hover:bg-slate-800"
                >
                  Selecionar
                </button>
              )}
              rotuloAcoes=""
            />
          ) : (
            <Tabela
              itens={resultadosCliente}
              chave={(c) => `${c.codigo}-${c.loja}`}
              carregando={buscando}
              colunas={[
                { cabecalho: 'Código', conteudo: (c) => c.codigo },
                { cabecalho: 'Loja', conteudo: (c) => c.loja },
                { cabecalho: 'Nome', conteudo: (c) => c.nome },
                { cabecalho: 'Município', conteudo: (c) => `${c.municipio}/${c.uf}` }
              ]}
              acoes={(c) => (
                <button
                  type="button"
                  onClick={() => selecionarCliente(c)}
                  className="rounded-lg px-2 py-1 text-xs font-medium text-swnet-600 hover:bg-swnet-50
                    dark:text-swnet-400 dark:hover:bg-slate-800"
                >
                  Selecionar
                </button>
              )}
              rotuloAcoes=""
            />
          )}
        </div>
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
