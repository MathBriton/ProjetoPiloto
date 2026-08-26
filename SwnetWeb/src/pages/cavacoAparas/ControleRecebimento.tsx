import { useCallback, useEffect, useState } from 'react'
import { ClipboardList, Eraser, Pencil, Plus, Save, Scale, Search, Trash2, Truck } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import { usePlanta } from '@/context/PlantaContext'
import {
  atualizarPesoBalanca,
  excluirRecebimento,
  formatarCnpj,
  listarFornecedores,
  listarRecebimentos,
  salvarRecebimento
} from '@/services/cavacoAparas/aparasApi'
import { formatarData } from '@/services/comuns/baseMock'
import type {
  DadosRecebimento,
  FiltrosRecebimento,
  FornecedorAparas,
  RecebimentoAparas
} from './tipos'

/**
 * Refatoração de Cavaco-Aparas/Controle-De-Recebimento
 * (CrudControleRecebimento.asp + Repositorio).
 *
 * O legado consulta `ObterNivel()` (coluna NivelControladoria da CRM_DTI)
 * para decidir o que o usuário pode fazer: com nível, edita o registro
 * inteiro; sem nível, só atualiza o peso da balança (ação "APB").
 */

const FILTROS_VAZIOS: FiltrosRecebimento = {
  notaFiscal: '',
  fornecedor: '',
  produto: '',
  dataInicio: '',
  dataFim: ''
}

const RECEBIMENTO_VAZIO: DadosRecebimento = {
  notaFiscal: '',
  serie: '',
  dataNf: '',
  dataTeste: '',
  horaTeste: '',
  codigoFornecedor: '',
  cnpj: '',
  nomeFornecedor: '',
  produto: '',
  pesoNf: '',
  pesoBalanca: '',
  umidade: '',
  impureza: '',
  proibitivo: '',
  outros: '',
  local: '',
  quantidade: '',
  fardosAnalisados: '',
  placa: '',
  motorista: '',
  observacao: ''
}

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

export function ControleRecebimento() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario, temPermissao } = useAuth()
  const { plantaAtual } = usePlanta()

  /** Equivalente ao ObterNivel() do repositório ASP */
  const podeEditarTudo = temPermissao('recebimento.controladoria')

  const [registros, setRegistros] = useState<RecebimentoAparas[]>([])
  const [fornecedores, setFornecedores] = useState<FornecedorAparas[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtros, setFiltros] = useState<FiltrosRecebimento>(FILTROS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosRecebimento>(FILTROS_VAZIOS)

  const [edicao, setEdicao] = useState<DadosRecebimento | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<RecebimentoAparas | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, listaFornecedores] = await Promise.all([
        listarRecebimentos(filtrosAplicados),
        listarFornecedores()
      ])
      setRegistros(lista)
      setFornecedores(listaFornecedores)
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosRecebimento>(campo: C, valor: string) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function limparFiltros() {
    setFiltros(FILTROS_VAZIOS)
    setFiltrosAplicados(FILTROS_VAZIOS)
  }

  function mudar<C extends keyof DadosRecebimento>(campo: C, valor: string) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  /** Selecionar o fornecedor preenche CNPJ e razão social, como no legado */
  function escolherFornecedor(codigo: string) {
    const fornecedor = fornecedores.find((f) => f.codigo === codigo)
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            codigoFornecedor: codigo,
            cnpj: fornecedor?.cnpj ?? '',
            nomeFornecedor: fornecedor?.nome ?? ''
          }
        : atual
    )
  }

  function abrirEdicao(registro: RecebimentoAparas) {
    setErro('')
    setEdicao({
      id: registro.id,
      notaFiscal: registro.notaFiscal,
      serie: registro.serie,
      dataNf: registro.dataNf,
      dataTeste: registro.dataTeste,
      horaTeste: registro.horaTeste,
      codigoFornecedor: registro.codigoFornecedor,
      cnpj: registro.cnpj,
      nomeFornecedor: registro.nomeFornecedor,
      produto: registro.produto,
      pesoNf: String(registro.pesoNf),
      pesoBalanca: registro.pesoBalanca === null ? '' : String(registro.pesoBalanca),
      umidade: String(registro.umidade),
      impureza: String(registro.impureza),
      proibitivo: String(registro.proibitivo),
      outros: String(registro.outros),
      local: registro.local,
      quantidade: String(registro.quantidade),
      fardosAnalisados: String(registro.fardosAnalisados),
      placa: registro.placa,
      motorista: registro.motorista,
      observacao: registro.observacao
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      // Sem nível de controladoria, o legado só grava o peso da balança
      const resultado =
        !podeEditarTudo && edicao.id
          ? await atualizarPesoBalanca(edicao.id, edicao.pesoBalanca)
          : await salvarRecebimento(edicao, plantaAtual?.codigo ?? '', usuario?.login ?? '')

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

    const resultado = await excluirRecebimento(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  const somenteBalanca = !podeEditarTudo && !!edicao?.id

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Controle de Recebimento" />

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Entrada
            id="filtroNF"
            rotulo="Nota Fiscal"
            value={filtros.notaFiscal}
            onChange={(e) => mudarFiltro('notaFiscal', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroFornecedor"
            rotulo="Fornecedor"
            value={filtros.fornecedor}
            onChange={(e) => mudarFiltro('fornecedor', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroProduto"
            rotulo="Produto"
            value={filtros.produto}
            onChange={(e) => mudarFiltro('produto', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroDataIni"
            rotulo="Teste de"
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
            <Botao variante="neutro" icone={Eraser} aoClicar={limparFiltros} titulo="Limpar filtros">
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel
        titulo={`Recebimentos (${registros.length})`}
        icone={ClipboardList}
        acoes={
          podeEditarTudo && (
            <Botao
              variante="sucesso"
              icone={Plus}
              aoClicar={() => {
                setErro('')
                setEdicao(RECEBIMENTO_VAZIO)
              }}
            >
              Novo
            </Botao>
          )
        }
      >
        {!podeEditarTudo && (
          <p className="mb-3 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800 dark:bg-amber-500/10 dark:text-amber-400">
            Seu usuário não tem nível de controladoria: é possível apenas atualizar o peso da balança.
          </p>
        )}

        <Tabela
          itens={registros}
          chave={(registro) => registro.id}
          carregando={carregando}
          colunas={[
            { cabecalho: 'N.F / Série', conteudo: (r) => `${r.notaFiscal} / ${r.serie}` },
            { cabecalho: 'Data N.F', conteudo: (r) => formatarData(r.dataNf) },
            {
              cabecalho: 'Data/Hora Teste',
              conteudo: (r) => `${formatarData(r.dataTeste)} ${r.horaTeste}`
            },
            { cabecalho: 'CNPJ', conteudo: (r) => formatarCnpj(r.cnpj) },
            { cabecalho: 'Razão Social', conteudo: (r) => r.nomeFornecedor },
            { cabecalho: 'Produto', conteudo: (r) => r.produto },
            { cabecalho: 'Peso N.F', conteudo: (r) => numero.format(r.pesoNf), alinhamento: 'direita' },
            {
              cabecalho: 'Peso Balança',
              conteudo: (r) =>
                r.pesoBalanca === null ? (
                  <span className="text-xs text-amber-600 dark:text-amber-400">pendente</span>
                ) : (
                  numero.format(r.pesoBalanca)
                ),
              alinhamento: 'direita'
            },
            { cabecalho: '% Umid.', conteudo: (r) => numero.format(r.umidade), alinhamento: 'direita' },
            { cabecalho: '% Impur.', conteudo: (r) => numero.format(r.impureza), alinhamento: 'direita' },
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
            }
          ]}
          acoes={(registro) => (
            <>
              <BotaoAcao
                icone={podeEditarTudo ? Pencil : Scale}
                titulo={podeEditarTudo ? 'Editar registro' : 'Atualizar peso da balança'}
                aoClicar={() => abrirEdicao(registro)}
              />
              {podeEditarTudo && (
                <BotaoAcao
                  icone={Trash2}
                  titulo="Excluir registro"
                  tom="perigo"
                  aoClicar={() => setExcluindo(registro)}
                />
              )}
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={
          somenteBalanca
            ? 'Atualizar Peso da Balança'
            : edicao?.id
              ? 'Editar Recebimento'
              : 'Novo Recebimento'
        }
        icone={somenteBalanca ? Scale : ClipboardList}
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

            <Secao titulo="Nota Fiscal" icone={ClipboardList}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="edtNF"
                  rotulo="Nota Fiscal"
                  required
                  value={edicao.notaFiscal}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('notaFiscal', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtSerie"
                  rotulo="Série"
                  required
                  maxLength={3}
                  value={edicao.serie}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('serie', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="edtDataNf"
                  rotulo="Data da N.F"
                  type="date"
                  required
                  value={edicao.dataNf}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('dataNf', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtDataTeste"
                  rotulo="Data do Teste"
                  type="date"
                  required
                  value={edicao.dataTeste}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('dataTeste', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="edtHoraTeste"
                  rotulo="Hora do Teste"
                  type="time"
                  required
                  value={edicao.horaTeste}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('horaTeste', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
              </div>
            </Secao>

            <Secao titulo="Fornecedor" icone={Truck}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="edtFornecedor"
                  rotulo="Fornecedor"
                  required
                  value={edicao.codigoFornecedor}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => escolherFornecedor(e.target.value)}
                  larguraClasse="sm:col-span-6"
                >
                  <option value="">Selecione um fornecedor...</option>
                  {fornecedores.map((fornecedor) => (
                    <option key={fornecedor.codigo} value={fornecedor.codigo}>
                      {fornecedor.codigo} — {fornecedor.nome}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="edtCnpj"
                  rotulo="CNPJ"
                  readOnly
                  value={edicao.cnpj ? formatarCnpj(edicao.cnpj) : ''}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtNomeForn"
                  rotulo="Razão Social"
                  readOnly
                  value={edicao.nomeFornecedor}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Carga e Análise" icone={Scale}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="edtProduto"
                  rotulo="Produto"
                  required
                  value={edicao.produto}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('produto', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="edtPesoNf"
                  rotulo="Peso N.F / Kg"
                  type="number"
                  step="0.01"
                  required
                  value={edicao.pesoNf}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('pesoNf', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="edtPesoBal"
                  rotulo="Peso Balança / Kg"
                  type="number"
                  step="0.01"
                  value={edicao.pesoBalanca}
                  disabled={salvando}
                  autoFocus={somenteBalanca}
                  onChange={(e) => mudar('pesoBalanca', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="edtUmidade"
                  rotulo="% Umidade"
                  type="number"
                  step="0.1"
                  required
                  value={edicao.umidade}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('umidade', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtImpureza"
                  rotulo="% Impureza"
                  type="number"
                  step="0.1"
                  value={edicao.impureza}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('impureza', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtProibitivo"
                  rotulo="% Proibitivo"
                  type="number"
                  step="0.1"
                  value={edicao.proibitivo}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('proibitivo', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="edtOutros"
                  rotulo="% Outros"
                  type="number"
                  step="0.1"
                  value={edicao.outros}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('outros', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Armazenagem e Transporte" icone={Truck}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="edtLocal"
                  rotulo="Local de Armazenagem"
                  value={edicao.local}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('local', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="edtQuantidade"
                  rotulo="Quantidade"
                  type="number"
                  value={edicao.quantidade}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('quantidade', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="edtFardos"
                  rotulo="Fardos Analisados"
                  type="number"
                  value={edicao.fardosAnalisados}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('fardosAnalisados', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="edtMotorista"
                  rotulo="Nome do Motorista"
                  required
                  value={edicao.motorista}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('motorista', e.target.value)}
                  larguraClasse="sm:col-span-8"
                />
                <Entrada
                  id="edtPlaca"
                  rotulo="Placa"
                  required
                  maxLength={8}
                  value={edicao.placa}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('placa', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="edtObs"
                  rotulo="Observação"
                  value={edicao.observacao}
                  disabled={salvando || somenteBalanca}
                  onChange={(e) => mudar('observacao', e.target.value)}
                  larguraClasse="sm:col-span-12"
                />
              </div>
            </Secao>
          </>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir recebimento"
        mensagem={
          excluindo
            ? `Deseja realmente excluir o recebimento da N.F ${excluindo.notaFiscal} (${excluindo.nomeFornecedor})?`
            : ''
        }
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
