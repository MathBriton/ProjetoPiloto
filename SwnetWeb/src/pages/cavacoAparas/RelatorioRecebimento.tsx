import { useCallback, useEffect, useState } from 'react'
import { Eraser, FileDown, FileSpreadsheet, Search } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import { formatarCnpj, listarFornecedores, listarRelatorio } from '@/services/cavacoAparas/aparasApi'
import { EMPRESAS, codigoEmpresa, formatarData } from '@/services/comuns/baseMock'
import type { FiltrosRelatorioRecebimento, FornecedorAparas, RecebimentoAparas } from './tipos'

/** Refatoração de Cavaco-Aparas/Controle-De-Recebimento/RelatorioControleRecebimento.asp. */

const FILTROS_VAZIOS: FiltrosRelatorioRecebimento = {
  planta: 'TODOS',
  fornecedor: '',
  dataInicio: '',
  dataFim: ''
}

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

export function RelatorioRecebimento() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [filtros, setFiltros] = useState<FiltrosRelatorioRecebimento>(FILTROS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosRelatorioRecebimento>(FILTROS_VAZIOS)
  const [linhas, setLinhas] = useState<RecebimentoAparas[]>([])
  const [fornecedores, setFornecedores] = useState<FornecedorAparas[]>([])
  const [carregando, setCarregando] = useState(true)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [resultado, listaFornecedores] = await Promise.all([
        listarRelatorio(filtrosAplicados),
        listarFornecedores()
      ])
      setLinhas(resultado)
      setFornecedores(listaFornecedores)
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosRelatorioRecebimento>(campo: C, valor: string) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function limpar() {
    setFiltros(FILTROS_VAZIOS)
    setFiltrosAplicados(FILTROS_VAZIOS)
  }

  function exportar(formato: 'excel' | 'pdf') {
    if (linhas.length === 0) {
      mostrarToast('Não há registros para exportar.', 'warning')
      return
    }
    // TODO: o legado usa SheetJS/jsPDF; ainda não há lib de exportação no frontend
    mostrarToast(
      `Exportação em ${formato === 'excel' ? 'Excel' : 'PDF'} (${linhas.length} registro(s)) será implementada em breve.`,
      'info'
    )
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Relatório de Controle de Recebimento" />

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Selecao
            id="filtroPlanta"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => mudarFiltro('planta', e.target.value)}
            larguraClasse="sm:col-span-3"
          >
            <option value="TODOS">Todas as plantas</option>
            {EMPRESAS.map((empresa) => (
              <option key={empresa.id} value={codigoEmpresa(empresa.id)}>
                {empresa.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroFornecedor"
            rotulo="Fornecedor"
            value={filtros.fornecedor}
            onChange={(e) => mudarFiltro('fornecedor', e.target.value)}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Todos os fornecedores</option>
            {fornecedores.map((fornecedor) => (
              <option key={fornecedor.codigo} value={fornecedor.codigo}>
                {fornecedor.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroDataIni"
            rotulo="De"
            type="date"
            value={filtros.dataInicio}
            onChange={(e) => mudarFiltro('dataInicio', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroDataFim"
            rotulo="Até"
            type="date"
            value={filtros.dataFim}
            onChange={(e) => mudarFiltro('dataFim', e.target.value)}
            larguraClasse="sm:col-span-2"
          />

          <div className="flex gap-2 sm:col-span-1">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Filtrar">
              <span className="sr-only">Filtrar</span>
            </Botao>
            <Botao variante="neutro" icone={Eraser} aoClicar={limpar} titulo="Limpar filtros">
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel
        titulo={`${linhas.length} registro(s)`}
        icone={FileSpreadsheet}
        acoes={
          <>
            <Botao variante="sucesso" icone={FileSpreadsheet} aoClicar={() => exportar('excel')}>
              Excel
            </Botao>
            <Botao variante="perigo" icone={FileDown} aoClicar={() => exportar('pdf')}>
              PDF
            </Botao>
          </>
        }
      >
        <Tabela
          itens={linhas}
          chave={(linha) => linha.id}
          carregando={carregando}
          mensagemVazia="Nenhum recebimento encontrado para os filtros informados"
          colunas={[
            { cabecalho: 'Planta', conteudo: (r) => r.nomePlanta },
            { cabecalho: 'N.F', conteudo: (r) => r.notaFiscal },
            { cabecalho: 'Série', conteudo: (r) => r.serie },
            { cabecalho: 'Data N.F', conteudo: (r) => formatarData(r.dataNf) },
            { cabecalho: 'Data Teste', conteudo: (r) => formatarData(r.dataTeste) },
            { cabecalho: 'Hora', conteudo: (r) => r.horaTeste },
            { cabecalho: 'CNPJ', conteudo: (r) => formatarCnpj(r.cnpj) },
            { cabecalho: 'Fornecedor', conteudo: (r) => r.nomeFornecedor },
            { cabecalho: 'Produto', conteudo: (r) => r.produto },
            { cabecalho: 'Peso N.F', conteudo: (r) => numero.format(r.pesoNf), alinhamento: 'direita' },
            {
              cabecalho: 'Peso Bal.',
              conteudo: (r) => (r.pesoBalanca === null ? '—' : numero.format(r.pesoBalanca)),
              alinhamento: 'direita'
            },
            { cabecalho: '% Umid.', conteudo: (r) => numero.format(r.umidade), alinhamento: 'direita' },
            { cabecalho: '% Impurez.', conteudo: (r) => numero.format(r.impureza), alinhamento: 'direita' },
            { cabecalho: '% Proib.', conteudo: (r) => numero.format(r.proibitivo), alinhamento: 'direita' },
            { cabecalho: '% Outros', conteudo: (r) => numero.format(r.outros), alinhamento: 'direita' },
            { cabecalho: 'Local', conteudo: (r) => r.local },
            { cabecalho: 'Qtd.', conteudo: (r) => r.quantidade, alinhamento: 'direita' },
            { cabecalho: 'Fardos An.', conteudo: (r) => r.fardosAnalisados, alinhamento: 'direita' },
            { cabecalho: 'Placa', conteudo: (r) => r.placa },
            { cabecalho: 'Motorista', conteudo: (r) => r.motorista },
            {
              cabecalho: 'Obs',
              conteudo: (r) => (
                <span className="block max-w-48 truncate text-xs" title={r.observacao}>
                  {r.observacao || '—'}
                </span>
              )
            },
            { cabecalho: 'Incluído por', conteudo: (r) => r.incluidoPor }
          ]}
        />
      </Painel>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
