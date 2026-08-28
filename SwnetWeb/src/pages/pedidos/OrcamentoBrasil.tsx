import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  Calculator,
  Coins,
  Eraser,
  FileSearch,
  Package,
  Pencil,
  Plus,
  Save,
  Search,
  Trash2,
  Users
} from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Chave, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import {
  CONDICOES_PAGAMENTO_PEDIDO,
  ESTRUTURAS_PEDIDO,
  FILTROS_ORCAMENTO_VAZIOS,
  GRAMATURAS,
  ITEM_ORCAMENTO_VAZIO,
  ORCAMENTO_VAZIO,
  PLANTAS_POR_REGIAO,
  STATUS_ORCAMENTO,
  TIPOS_CAIXA,
  TIPOS_FRETE_PEDIDO,
  TIPOS_ITEM_PEDIDO,
  TIPOS_PRECO,
  buscarClientesPedido,
  calcularOrcamento,
  desativarOrcamento,
  listarOrcamentos,
  salvarOrcamento,
  type ClientePedido,
  type FiltrosOrcamento,
  type ItemOrcamento,
  type Orcamento,
  type ResultadoCalculo
} from '@/services/pedidos/pedidosApi'
import { nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Pedidos/Orçamento-Brasil.
 *
 * Inclui o cálculo de custo do OrcamentoCalculoTrimboxRepositorio, que no
 * legado consulta o TRIMBOX da planta para chegar ao preço sugerido.
 */

const moeda = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 })

const TOM_STATUS = { ATIVO: 'sucesso', DESATIVADO: 'neutro', VENCIDO: 'perigo' } as const

type Rascunho = Partial<Orcamento> & { id?: number }

export function OrcamentoBrasil() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [filtros, setFiltros] = useState<FiltrosOrcamento>(FILTROS_ORCAMENTO_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_ORCAMENTO_VAZIOS)
  const [orcamentos, setOrcamentos] = useState<Orcamento[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<Rascunho | null>(null)
  const [item, setItem] = useState<ItemOrcamento>(ITEM_ORCAMENTO_VAZIO)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const [buscaAberta, setBuscaAberta] = useState(false)
  const [termo, setTermo] = useState('')
  const [resultados, setResultados] = useState<ClientePedido[]>([])
  const [buscando, setBuscando] = useState(false)

  const [calculo, setCalculo] = useState<{ orcamento: Orcamento; linhas: ResultadoCalculo[] } | null>(null)
  const [calculando, setCalculando] = useState(false)
  const [desativando, setDesativando] = useState<Orcamento | null>(null)
  const [motivo, setMotivo] = useState('')
  const [detalhe, setDetalhe] = useState<Orcamento | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setOrcamentos(await listarOrcamentos(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof Orcamento>(campo: C, valor: Orcamento[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function mudarItem<C extends keyof ItemOrcamento>(campo: C, valor: ItemOrcamento[C]) {
    setItem((atual) => ({ ...atual, [campo]: valor }))
  }

  async function pesquisarCliente() {
    setBuscando(true)
    try {
      setResultados(await buscarClientesPedido(termo))
    } finally {
      setBuscando(false)
    }
  }

  function selecionarCliente(cliente: ClientePedido) {
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            codCliente: cliente.codigo,
            codLoja: cliente.loja,
            nomeCliente: cliente.nome,
            cnpj: cliente.cnpj,
            municipio: cliente.municipio,
            uf: cliente.uf,
            condicoesPagamento: atual.condicoesPagamento || cliente.condicaoPagamento
          }
        : atual
    )
    setBuscaAberta(false)
  }

  function adicionarItem() {
    if (!item.referencia.trim() && !item.codigoProduto.trim()) {
      mostrarToast('Informe o produto ou a referência do item.', 'warning')
      return
    }
    if (!Number(item.quantidade)) {
      mostrarToast('Informe a quantidade.', 'warning')
      return
    }

    mudar('itens', [...(edicao?.itens ?? []), item])
    setItem(ITEM_ORCAMENTO_VAZIO)
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarOrcamento(edicao)
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

  async function calcular(orcamento: Orcamento) {
    setCalculando(true)
    try {
      const linhas = await calcularOrcamento(orcamento)
      setCalculo({ orcamento, linhas })
    } finally {
      setCalculando(false)
    }
  }

  async function confirmarDesativacao() {
    if (!desativando) return

    const resultado = await desativarOrcamento(desativando.id, motivo)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setDesativando(null)
      setMotivo('')
      await carregar()
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Orçamento Brasil"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setItem(ITEM_ORCAMENTO_VAZIO)
              setEdicao({ ...ORCAMENTO_VAZIO })
            }}
          >
            Novo orçamento
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
            id="filtroPlantaOrc"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todas</option>
            {PLANTAS_POR_REGIAO.BRASIL.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroClienteOrc"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-4"
          />

          <Selecao
            id="filtroStatusOrc"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todos</option>
            {STATUS_ORCAMENTO.map((s) => (
              <option key={s.codigo} value={s.codigo}>
                {s.descricao}
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
              titulo="Limpar"
              aoClicar={() => {
                setFiltros(FILTROS_ORCAMENTO_VAZIOS)
                setFiltrosAplicados(FILTROS_ORCAMENTO_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${orcamentos.length} orçamento(s)`} icone={Coins}>
        <Tabela
          itens={orcamentos}
          chave={(o) => o.id}
          carregando={carregando}
          mensagemVazia="Nenhum orçamento encontrado"
          colunas={[
            { cabecalho: '#', conteudo: (o) => o.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (o) => nomePlanta(o.planta) },
            { cabecalho: 'Data Inclusão', conteudo: (o) => formatarData(o.dataInclusao) },
            { cabecalho: 'Validade De', conteudo: (o) => formatarData(o.dataValidadeDe) },
            { cabecalho: 'Validade Até', conteudo: (o) => formatarData(o.dataValidadeAte) },
            {
              cabecalho: 'Status',
              conteudo: (o) => (
                <Etiqueta
                  texto={STATUS_ORCAMENTO.find((s) => s.codigo === o.status)?.descricao ?? o.status}
                  tom={TOM_STATUS[o.status as keyof typeof TOM_STATUS] ?? 'neutro'}
                />
              )
            },
            { cabecalho: 'Qtd. Itens', conteudo: (o) => o.itens.length, alinhamento: 'centro' },
            { cabecalho: 'Cliente', conteudo: (o) => o.nomeCliente },
            { cabecalho: 'CNPJ', conteudo: (o) => o.cnpj },
            { cabecalho: 'Cidade', conteudo: (o) => `${o.municipio}/${o.uf}` },
            { cabecalho: 'Vendedor', conteudo: (o) => o.nomeVendedor }
          ]}
          acoes={(orcamento) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(orcamento)} />
              <BotaoAcao
                icone={Calculator}
                titulo="Calcular custo"
                aoClicar={() => calcular(orcamento)}
              />
              {orcamento.status === 'ATIVO' && (
                <>
                  <BotaoAcao
                    icone={Pencil}
                    titulo="Editar orçamento"
                    aoClicar={() => {
                      setErro('')
                      setItem(ITEM_ORCAMENTO_VAZIO)
                      setEdicao({ ...orcamento })
                    }}
                  />
                  <BotaoAcao
                    icone={Ban}
                    titulo="Desativar orçamento"
                    tom="perigo"
                    aoClicar={() => {
                      setMotivo('')
                      setDesativando(orcamento)
                    }}
                  />
                </>
              )}
            </>
          )}
        />
      </Painel>

      {/* Formulário */}
      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar orçamento ${edicao.id}` : 'Novo Orçamento'}
        icone={Coins}
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

            <Secao titulo="Cliente" icone={Users}>
              <div className="mb-3">
                <Chave
                  rotulo="Cliente Novo"
                  descricao="Sem cadastro no Protheus — preencha os dados manualmente."
                  marcado={edicao.clienteNovo ?? false}
                  desabilitado={salvando}
                  aoAlternar={(v) => mudar('clienteNovo', v)}
                />
              </div>

              <div className="grid items-end gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaOrc"
                  rotulo="Planta"
                  required
                  value={edicao.planta ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('planta', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_POR_REGIAO.BRASIL.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="codClienteOrc"
                  rotulo="Cód. Cliente"
                  value={edicao.codCliente ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('codCliente', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="codLojaOrc"
                  rotulo="Cód. Loja"
                  maxLength={2}
                  value={edicao.codLoja ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('codLoja', e.target.value)}
                  larguraClasse="sm:col-span-1"
                />

                <Entrada
                  id="nomeClienteOrc"
                  rotulo="Nome do Cliente"
                  required
                  value={edicao.nomeCliente ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('nomeCliente', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <div className="sm:col-span-2">
                  <Botao
                    variante="neutro"
                    icone={Search}
                    aoClicar={() => {
                      setTermo('')
                      setBuscaAberta(true)
                      pesquisarCliente()
                    }}
                    desabilitado={salvando}
                  >
                    Buscar
                  </Botao>
                </div>

                <Entrada
                  id="cnpjOrc"
                  rotulo="CNPJ"
                  value={edicao.cnpj ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('cnpj', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="codVendedorOrc"
                  rotulo="Cód. Vendedor"
                  value={edicao.codVendedor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codVendedor', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="nomeVendedorOrc"
                  rotulo="Nome do Vendedor"
                  value={edicao.nomeVendedor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeVendedor', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="contatoOrc"
                  rotulo="Contato"
                  value={edicao.contato ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('contato', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Entrada
                  id="telefoneOrc"
                  rotulo="Telefone"
                  value={edicao.telefone ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('telefone', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="emailOrc"
                  rotulo="E-mail"
                  type="email"
                  value={edicao.email ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('email', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="municipioOrc"
                  rotulo="Município"
                  value={edicao.municipio ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('municipio', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="ufOrc"
                  rotulo="UF"
                  maxLength={2}
                  value={edicao.uf ?? ''}
                  disabled={salvando || !edicao.clienteNovo}
                  onChange={(e) => mudar('uf', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-2"
                />
              </div>
            </Secao>

            <Secao titulo="Condições">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="validadeDe"
                  rotulo="Data Validade De"
                  type="date"
                  required
                  value={edicao.dataValidadeDe ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('dataValidadeDe', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="validadeAte"
                  rotulo="Data Validade Até"
                  type="date"
                  required
                  value={edicao.dataValidadeAte ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('dataValidadeAte', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Selecao
                  id="tipoFreteOrc"
                  rotulo="Tipo de Frete"
                  value={edicao.tipoFrete ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoFrete', e.target.value)}
                  larguraClasse="sm:col-span-2"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_FRETE_PEDIDO.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="tipoCaminhaoOrc"
                  rotulo="Tipo de Caminhão"
                  value={edicao.tipoCaminhao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoCaminhao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Selecao
                  id="condPagtoOrc"
                  rotulo="Condições de Pagamento"
                  value={edicao.condicoesPagamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('condicoesPagamento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {CONDICOES_PAGAMENTO_PEDIDO.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="icmsOrc"
                  rotulo="ICMS"
                  value={edicao.icms ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('icms', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="aliquotaIcmsOrc"
                  rotulo="Alíquota de ICMS"
                  value={edicao.aliquotaIcms ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('aliquotaIcms', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="pisOrc"
                  rotulo="PIS"
                  value={edicao.pis ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('pis', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="cofinsOrc"
                  rotulo="COFINS"
                  value={edicao.cofins ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('cofins', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
              </div>
            </Secao>

            <Secao titulo="Itens do Orçamento" icone={Package}>
              <div className="mb-3 grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoItemOrc"
                  rotulo="Tipo do item"
                  value={item.tipoItem}
                  disabled={salvando}
                  onChange={(e) => mudarItem('tipoItem', e.target.value)}
                  larguraClasse="sm:col-span-2"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_ITEM_PEDIDO.map((t) => (
                    <option key={t.codigo} value={t.codigo}>
                      {t.descricao}
                    </option>
                  ))}
                </Selecao>
                <Selecao
                  id="tipoCaixaOrc"
                  rotulo="Tipo de Caixa"
                  value={item.tipoCaixa}
                  disabled={salvando || item.tipoItem !== 'CX'}
                  onChange={(e) => mudarItem('tipoCaixa', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_CAIXA.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="codProdutoOrc"
                  rotulo="Cód. Produto"
                  value={item.codigoProduto}
                  disabled={salvando}
                  onChange={(e) => mudarItem('codigoProduto', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="referenciaOrc"
                  rotulo="Referência"
                  value={item.referencia}
                  disabled={salvando}
                  onChange={(e) => mudarItem('referencia', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />

                <Selecao
                  id="estruturaOrc"
                  rotulo="Estrutura"
                  value={item.estrutura}
                  disabled={salvando}
                  onChange={(e) => mudarItem('estrutura', e.target.value)}
                  larguraClasse="sm:col-span-2"
                >
                  <option value="">Selecione...</option>
                  {ESTRUTURAS_PEDIDO.map((e2) => (
                    <option key={e2} value={e2}>
                      {e2}
                    </option>
                  ))}
                </Selecao>
                <Selecao
                  id="gramaturaOrc"
                  rotulo="Gramatura"
                  value={item.gramatura}
                  disabled={salvando}
                  onChange={(e) => mudarItem('gramatura', e.target.value)}
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
                  id="quantidadeOrc"
                  rotulo="Quantidade"
                  type="number"
                  min="1"
                  value={item.quantidade}
                  disabled={salvando}
                  onChange={(e) => mudarItem('quantidade', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Selecao
                  id="tipoPrecoOrc"
                  rotulo="Tipo de Preço"
                  value={item.tipoPreco}
                  disabled={salvando}
                  onChange={(e) => mudarItem('tipoPreco', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_PRECO.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="valorUnitarioOrc"
                  rotulo="Valor Unitário"
                  type="number"
                  step="0.0001"
                  value={item.valorUnitario}
                  disabled={salvando}
                  onChange={(e) => mudarItem('valorUnitario', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Entrada
                  id="compIntOrc"
                  rotulo="Compr. Interno"
                  value={item.comprimentoInterno}
                  disabled={salvando}
                  onChange={(e) => mudarItem('comprimentoInterno', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="largIntOrc"
                  rotulo="Largura Interna"
                  value={item.larguraInterna}
                  disabled={salvando}
                  onChange={(e) => mudarItem('larguraInterna', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="altIntOrc"
                  rotulo="Altura Interna"
                  value={item.alturaInterna}
                  disabled={salvando}
                  onChange={(e) => mudarItem('alturaInterna', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="compChapaOrc"
                  rotulo="Compr. Chapa"
                  value={item.comprimentoChapa}
                  disabled={salvando}
                  onChange={(e) => mudarItem('comprimentoChapa', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="largChapaOrc"
                  rotulo="Largura Chapa"
                  value={item.larguraChapa}
                  disabled={salvando}
                  onChange={(e) => mudarItem('larguraChapa', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="arranjoOrc"
                  rotulo="Arranjo"
                  value={item.arranjo}
                  disabled={salvando}
                  onChange={(e) => mudarItem('arranjo', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />

                <div className="sm:col-span-12">
                  <Botao variante="neutro" icone={Plus} aoClicar={adicionarItem} desabilitado={salvando}>
                    Adicionar item
                  </Botao>
                </div>
              </div>

              <Tabela
                itens={edicao.itens ?? []}
                chave={(linha) => (edicao.itens ?? []).indexOf(linha)}
                mensagemVazia="Nenhum item no orçamento"
                colunas={[
                  {
                    cabecalho: 'Tipo',
                    conteudo: (i) =>
                      TIPOS_ITEM_PEDIDO.find((t) => t.codigo === i.tipoItem)?.descricao ?? i.tipoItem
                  },
                  { cabecalho: 'Produto', conteudo: (i) => i.codigoProduto || '—' },
                  { cabecalho: 'Referência', conteudo: (i) => i.referencia },
                  { cabecalho: 'Estrutura', conteudo: (i) => i.estrutura || '—' },
                  { cabecalho: 'Quantidade', conteudo: (i) => i.quantidade, alinhamento: 'direita' },
                  { cabecalho: 'Tipo Preço', conteudo: (i) => i.tipoPreco || '—' },
                  {
                    cabecalho: 'Valor Unit.',
                    conteudo: (i) => moeda.format(Number(i.valorUnitario || 0)),
                    alinhamento: 'direita'
                  }
                ]}
                acoes={(linha) => (
                  <button
                    type="button"
                    onClick={() =>
                      mudar('itens', (edicao.itens ?? []).filter((outro) => outro !== linha))
                    }
                    disabled={salvando}
                    aria-label="Remover item"
                    className="rounded-lg p-1.5 text-slate-400 hover:bg-red-50 hover:text-red-600 dark:hover:bg-slate-800"
                  >
                    <Trash2 size={14} />
                  </button>
                )}
              />
            </Secao>
          </>
        )}
      </Modal>

      {/* Busca de cliente */}
      <Modal
        aberto={buscaAberta}
        titulo="Pesquisar Cliente"
        icone={Search}
        largura="lg"
        aoFechar={() => setBuscaAberta(false)}
      >
        <div className="grid gap-4">
          <div className="flex items-end gap-2">
            <Entrada
              id="termoClienteOrc"
              rotulo="Código ou nome"
              value={termo}
              disabled={buscando}
              onChange={(e) => setTermo(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  pesquisarCliente()
                }
              }}
              larguraClasse="flex-1"
            />
            <Botao icone={Search} aoClicar={pesquisarCliente} desabilitado={buscando}>
              Buscar
            </Botao>
          </div>

          <Tabela
            itens={resultados}
            chave={(c) => `${c.codigo}-${c.loja}`}
            carregando={buscando}
            colunas={[
              { cabecalho: 'Código', conteudo: (c) => c.codigo },
              { cabecalho: 'Loja', conteudo: (c) => c.loja },
              { cabecalho: 'Nome', conteudo: (c) => c.nome },
              { cabecalho: 'CNPJ', conteudo: (c) => c.cnpj }
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
        </div>
      </Modal>

      {/* Cálculo de custo */}
      <Modal
        aberto={calculo !== null || calculando}
        titulo="Cálculo de custo (TRIMBOX)"
        icone={Calculator}
        largura="lg"
        aoFechar={() => setCalculo(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setCalculo(null)}>
            Fechar
          </Botao>
        }
      >
        {calculando ? (
          <p className="py-8 text-center text-sm text-slate-400">Consultando o TRIMBOX...</p>
        ) : (
          calculo && (
            <div className="grid gap-4">
              <p className="text-sm text-slate-600 dark:text-slate-300">
                Orçamento {calculo.orcamento.id} — {calculo.orcamento.nomeCliente}
              </p>

              <Tabela
                itens={calculo.linhas}
                chave={(l) => l.item}
                colunas={[
                  { cabecalho: 'Item', conteudo: (l) => l.item },
                  {
                    cabecalho: 'Custo Material',
                    conteudo: (l) => moeda.format(l.custoMaterial),
                    alinhamento: 'direita'
                  },
                  {
                    cabecalho: 'Custo Conversão',
                    conteudo: (l) => moeda.format(l.custoConversao),
                    alinhamento: 'direita'
                  },
                  {
                    cabecalho: 'Preço Sugerido',
                    conteudo: (l) => moeda.format(l.precoSugerido),
                    alinhamento: 'direita'
                  },
                  {
                    cabecalho: 'Margem',
                    conteudo: (l) => (
                      <Etiqueta
                        texto={`${l.margem}%`}
                        tom={l.margem >= 20 ? 'sucesso' : l.margem > 0 ? 'alerta' : 'perigo'}
                      />
                    ),
                    alinhamento: 'direita'
                  }
                ]}
              />
            </div>
          )
        )}
      </Modal>

      {/* Detalhes */}
      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Orçamento ${detalhe.id}` : ''}
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
                ['Cliente', detalhe.nomeCliente],
                ['CNPJ', detalhe.cnpj],
                ['Planta', nomePlanta(detalhe.planta)],
                ['Vendedor', detalhe.nomeVendedor || '—'],
                ['Validade', `${formatarData(detalhe.dataValidadeDe)} a ${formatarData(detalhe.dataValidadeAte)}`],
                ['Cond. Pagamento', detalhe.condicoesPagamento || '—']
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            <Tabela
              itens={detalhe.itens}
              chave={(linha) => detalhe.itens.indexOf(linha)}
              mensagemVazia="Nenhum item"
              colunas={[
                { cabecalho: 'Referência', conteudo: (i) => i.referencia },
                { cabecalho: 'Estrutura', conteudo: (i) => i.estrutura || '—' },
                { cabecalho: 'Quantidade', conteudo: (i) => i.quantidade, alinhamento: 'direita' },
                {
                  cabecalho: 'Valor Unit.',
                  conteudo: (i) => moeda.format(Number(i.valorUnitario || 0)),
                  alinhamento: 'direita'
                }
              ]}
            />
          </div>
        )}
      </Modal>

      {/* Desativar */}
      <Modal
        aberto={desativando !== null}
        titulo="Desativar orçamento"
        icone={Ban}
        aoFechar={() => setDesativando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setDesativando(null)}>
              Cancelar
            </Botao>
            <Botao variante="perigo" aoClicar={confirmarDesativacao}>
              Desativar
            </Botao>
          </>
        }
      >
        <Campo rotulo="Motivo da desativação" htmlFor="motivoDesativacao" obrigatorio>
          <textarea
            id="motivoDesativacao"
            rows={3}
            value={motivo}
            onChange={(e) => setMotivo(e.target.value)}
            className={classesControle}
          />
        </Campo>
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
