import { useState } from 'react'
import { Boxes, Eraser, FileDown, FileSpreadsheet, Search, Truck, X } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, EstadoVazio, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Tabela } from '@/components/comuns/Tabela'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import {
  PLANTAS_SIMULADOR,
  buscarClientes,
  buscarProdutos,
  simularCarga,
  type ClienteSimulador,
  type LinhaSimulacao,
  type ProdutoSimulador
} from '@/services/expedicao/expedicaoApi'

/**
 * Refatoração de Expedicao/Simulador-De-Cargas
 * (ConsultaTransporteDeCarga.asp + Repositorio).
 *
 * O usuário escolhe planta, cliente e produto — os dois últimos por modal de
 * busca — e a consulta devolve a capacidade por tipo de caminhão, com as
 * colunas normais e as de empilhamento 2800.
 */

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

export function SimuladorCargas() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [planta, setPlanta] = useState('')
  const [cliente, setCliente] = useState<ClienteSimulador | null>(null)
  const [produto, setProduto] = useState<ProdutoSimulador | null>(null)

  const [buscaAberta, setBuscaAberta] = useState<'cliente' | 'produto' | null>(null)
  const [termo, setTermo] = useState('')
  const [resultadosCliente, setResultadosCliente] = useState<ClienteSimulador[]>([])
  const [resultadosProduto, setResultadosProduto] = useState<ProdutoSimulador[]>([])
  const [buscando, setBuscando] = useState(false)

  const [linhas, setLinhas] = useState<LinhaSimulacao[]>([])
  const [consultando, setConsultando] = useState(false)
  const [consultaFeita, setConsultaFeita] = useState(false)

  async function abrirBusca(tipo: 'cliente' | 'produto') {
    setBuscaAberta(tipo)
    setTermo('')
    setBuscando(true)
    try {
      if (tipo === 'cliente') setResultadosCliente(await buscarClientes(''))
      else setResultadosProduto(await buscarProdutos(''))
    } finally {
      setBuscando(false)
    }
  }

  async function pesquisarNaBusca() {
    setBuscando(true)
    try {
      if (buscaAberta === 'cliente') setResultadosCliente(await buscarClientes(termo))
      else setResultadosProduto(await buscarProdutos(termo))
    } finally {
      setBuscando(false)
    }
  }

  async function consultar() {
    if (!planta || !cliente || !produto) {
      mostrarToast('Selecione a filial, o cliente e o produto.', 'warning')
      return
    }

    setConsultando(true)
    try {
      setLinhas(await simularCarga(planta, cliente.codigo, produto.codigo))
      setConsultaFeita(true)
    } finally {
      setConsultando(false)
    }
  }

  function limpar() {
    setPlanta('')
    setCliente(null)
    setProduto(null)
    setLinhas([])
    setConsultaFeita(false)
  }

  function exportar(formato: 'csv' | 'pdf') {
    if (linhas.length === 0) {
      mostrarToast('Não há resultado para exportar.', 'warning')
      return
    }
    // TODO: o legado gera CSV/PDF no cliente; sem lib de exportação por ora
    mostrarToast(`Exportação em ${formato.toUpperCase()} será implementada em breve.`, 'info')
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Simulador de Cargas" />

      <Painel titulo="Parâmetros da simulação" icone={Truck}>
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            consultar()
          }}
        >
          <Selecao
            id="cboPlanta"
            rotulo="Filial"
            required
            value={planta}
            disabled={consultando}
            onChange={(e) => setPlanta(e.target.value)}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Selecione...</option>
            {PLANTAS_SIMULADOR.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.nome}
              </option>
            ))}
          </Selecao>

          <Campo rotulo="Cliente" obrigatorio className="sm:col-span-4">
            <div className="flex">
              <input
                type="text"
                readOnly
                value={cliente ? `${cliente.codigo} — ${cliente.nome}` : ''}
                placeholder="Nenhum cliente selecionado"
                className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
              />
              {cliente && (
                <button
                  type="button"
                  onClick={() => setCliente(null)}
                  aria-label="Limpar cliente"
                  className="border border-l-0 border-slate-200 px-2 text-slate-400 hover:text-slate-600
                    dark:border-slate-700"
                >
                  <X size={15} />
                </button>
              )}
              <button
                type="button"
                onClick={() => abrirBusca('cliente')}
                disabled={consultando}
                aria-label="Buscar cliente"
                className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                  hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
              >
                <Search size={15} />
              </button>
            </div>
          </Campo>

          <Campo rotulo="Produto" obrigatorio className="sm:col-span-3">
            <div className="flex">
              <input
                type="text"
                readOnly
                value={produto ? `${produto.codigo} — ${produto.referencia}` : ''}
                placeholder="Nenhum produto selecionado"
                className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
              />
              {produto && (
                <button
                  type="button"
                  onClick={() => setProduto(null)}
                  aria-label="Limpar produto"
                  className="border border-l-0 border-slate-200 px-2 text-slate-400 hover:text-slate-600
                    dark:border-slate-700"
                >
                  <X size={15} />
                </button>
              )}
              <button
                type="button"
                onClick={() => abrirBusca('produto')}
                disabled={consultando}
                aria-label="Buscar produto"
                className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                  hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
              >
                <Search size={15} />
              </button>
            </div>
          </Campo>

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={consultando}>
              {consultando ? '...' : 'Consultar'}
            </Botao>
            <Botao variante="neutro" icone={Eraser} aoClicar={limpar} titulo="Limpar">
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      {consultaFeita && (
        <Painel
          titulo="Capacidade por tipo de caminhão"
          icone={Boxes}
          acoes={
            <>
              <Botao variante="sucesso" icone={FileSpreadsheet} aoClicar={() => exportar('csv')}>
                CSV
              </Botao>
              <Botao variante="perigo" icone={FileDown} aoClicar={() => exportar('pdf')}>
                PDF
              </Botao>
            </>
          }
        >
          {linhas.length === 0 && !consultando ? (
            <EstadoVazio
              icone={Boxes}
              titulo="Nenhuma capacidade encontrada"
              descricao="Não há dados de carga para essa combinação de filial, cliente e produto."
            />
          ) : (
            <Tabela
              itens={linhas}
              chave={(linha) => linha.tipo}
              carregando={consultando}
              colunas={[
                { cabecalho: 'Tipo', conteudo: (l) => l.tipo },
                {
                  cabecalho: 'Total Caixas',
                  conteudo: (l) => numero.format(l.totalCaixas),
                  alinhamento: 'direita'
                },
                {
                  cabecalho: 'Total Paletes',
                  conteudo: (l) => numero.format(l.totalPalete),
                  alinhamento: 'direita'
                },
                { cabecalho: 'M² Total', conteudo: (l) => numero.format(l.m2Total), alinhamento: 'direita' },
                {
                  cabecalho: 'Caixas (2800)',
                  conteudo: (l) => numero.format(l.totalCaixas2800),
                  alinhamento: 'direita'
                },
                {
                  cabecalho: 'Paletes (2800)',
                  conteudo: (l) => numero.format(l.totalPalete2800),
                  alinhamento: 'direita'
                },
                {
                  cabecalho: 'M² Total (2800)',
                  conteudo: (l) => numero.format(l.m2Total2800),
                  alinhamento: 'direita'
                }
              ]}
            />
          )}
        </Painel>
      )}

      <Modal
        aberto={buscaAberta !== null}
        titulo={buscaAberta === 'cliente' ? 'Busca de Cliente' : 'Busca de Produto'}
        icone={Search}
        largura="lg"
        aoFechar={() => setBuscaAberta(null)}
      >
        <div className="grid gap-4">
          <div className="flex items-end gap-2">
            <Entrada
              id="termoBusca"
              rotulo={buscaAberta === 'cliente' ? 'Nome ou código' : 'Descrição ou código'}
              value={termo}
              disabled={buscando}
              onChange={(e) => setTermo(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  pesquisarNaBusca()
                }
              }}
              larguraClasse="flex-1"
            />
            <Botao icone={Search} aoClicar={pesquisarNaBusca} desabilitado={buscando}>
              {buscando ? '...' : 'Buscar'}
            </Botao>
          </div>

          {buscaAberta === 'cliente' ? (
            <Tabela
              itens={resultadosCliente}
              chave={(item) => `${item.codigo}-${item.loja}`}
              carregando={buscando}
              colunas={[
                { cabecalho: 'Código', conteudo: (c) => c.codigo },
                { cabecalho: 'Loja', conteudo: (c) => c.loja },
                { cabecalho: 'Nome', conteudo: (c) => c.nome },
                { cabecalho: 'CNPJ', conteudo: (c) => c.cnpj },
                { cabecalho: 'Município', conteudo: (c) => `${c.municipio} / ${c.uf}` }
              ]}
              acoes={(item) => (
                <button
                  type="button"
                  onClick={() => {
                    setCliente(item)
                    setBuscaAberta(null)
                  }}
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
              itens={resultadosProduto}
              chave={(item) => item.codigo}
              carregando={buscando}
              colunas={[
                { cabecalho: 'Código', conteudo: (p) => p.codigo },
                { cabecalho: 'Descrição', conteudo: (p) => p.referencia }
              ]}
              acoes={(item) => (
                <button
                  type="button"
                  onClick={() => {
                    setProduto(item)
                    setBuscaAberta(null)
                  }}
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
