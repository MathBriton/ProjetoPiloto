import { useCallback, useEffect, useState } from 'react'
import { Check, ChevronRight, ClipboardList, Eraser, ListChecks, Pencil, Plus, Save, Search, Trash2, X } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, OpcoesRadio, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  ITENS_CHECKLIST,
  STATUS_CHECKLIST,
  TIPOS_VEICULO,
  buscarPorPlaca,
  excluirChecklist,
  indicadoresChecklist,
  listarChecklists,
  mudarStatus,
  salvarChecklist,
  type ChecklistTransporte as Checklist,
  type DadosChecklist,
  type FiltrosChecklist,
  type StatusChecklist
} from '@/services/expedicao/expedicaoApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Expedicao/Checklist-Transporte.
 *
 * Reúne CheckListTransporte.asp (formulário), CheckListTransporteGrade.asp
 * (grade + mudança de status) e CheckListTransporteIndicador.asp (contadores).
 */

const FILTROS_VAZIOS: FiltrosChecklist = { status: '', campo: '', valor: '' }

const CHECKLIST_VAZIO: DadosChecklist = {
  nomeProprietario: '',
  telefone: '',
  nomeMotorista: '',
  documento: '',
  placaVeiculo: '',
  dataPreenchimento: new Date().toISOString().slice(0, 10),
  idTipoVeiculo: '',
  quantidadeEixo: '',
  horarioInicio: '',
  horarioTermino: '',
  conclusao: '',
  destinoCarga: '',
  itensAtende: []
}

const TOM_STATUS = {
  'PATIO HJ': 'neutro',
  PORTARIA: 'info',
  EXPEDIÇÃO: 'alerta',
  FINALIZADO: 'sucesso'
} as const

const TOM_CONCLUSAO = { APROVADO: 'sucesso', RESSALVA: 'alerta', REPROVADO: 'perigo' } as const

export function ChecklistTransporte() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [checklists, setChecklists] = useState<Checklist[]>([])
  const [indicadores, setIndicadores] = useState<{ status: string; total: number }[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtros, setFiltros] = useState<FiltrosChecklist>(FILTROS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosChecklist>(FILTROS_VAZIOS)

  const [edicao, setEdicao] = useState<DadosChecklist | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<Checklist | null>(null)
  const [detalhe, setDetalhe] = useState<Checklist | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, contadores] = await Promise.all([
        listarChecklists(filtrosAplicados),
        indicadoresChecklist()
      ])
      setChecklists(lista)
      setIndicadores(contadores)
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof DadosChecklist>(campo: C, valor: DadosChecklist[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function alternarItem(id: number) {
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            itensAtende: atual.itensAtende.includes(id)
              ? atual.itensAtende.filter((item) => item !== id)
              : [...atual.itensAtende, id]
          }
        : atual
    )
  }

  /** BuscarPlaca: repete os dados do último checklist daquela placa */
  async function preencherPelaPlaca() {
    if (!edicao?.placaVeiculo.trim()) return

    const anterior = await buscarPorPlaca(edicao.placaVeiculo)
    if (!anterior) {
      mostrarToast('Nenhum checklist anterior para esta placa.', 'info')
      return
    }

    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            nomeProprietario: anterior.nomeProprietario,
            telefone: anterior.telefone,
            nomeMotorista: anterior.nomeMotorista,
            documento: anterior.documento,
            idTipoVeiculo: anterior.idTipoVeiculo,
            quantidadeEixo: String(anterior.quantidadeEixo)
          }
        : atual
    )
    mostrarToast('Dados do último checklist desta placa carregados.', 'success')
  }

  function abrirEdicao(checklist: Checklist) {
    setErro('')
    setEdicao({
      id: checklist.id,
      nomeProprietario: checklist.nomeProprietario,
      telefone: checklist.telefone,
      nomeMotorista: checklist.nomeMotorista,
      documento: checklist.documento,
      placaVeiculo: checklist.placaVeiculo,
      dataPreenchimento: checklist.dataPreenchimento,
      idTipoVeiculo: checklist.idTipoVeiculo,
      quantidadeEixo: String(checklist.quantidadeEixo),
      horarioInicio: checklist.horarioInicio,
      horarioTermino: checklist.horarioTermino,
      conclusao: checklist.conclusao,
      destinoCarga: checklist.destinoCarga,
      itensAtende: [...checklist.itensAtende]
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarChecklist(edicao)
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

  async function avancarStatus(checklist: Checklist, novoStatus: StatusChecklist) {
    const resultado = await mudarStatus(checklist.id, novoStatus)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  async function confirmarExclusao() {
    if (!excluindo) return

    const resultado = await excluirChecklist(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Checklist de Transporte"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao(CHECKLIST_VAZIO)
            }}
          >
            Novo checklist
          </Botao>
        }
      />

      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
        {indicadores.map((indicador) => (
          <button
            key={indicador.status}
            type="button"
            onClick={() => {
              const proximo = { ...filtros, status: indicador.status }
              setFiltros(proximo)
              setFiltrosAplicados(proximo)
            }}
            className={`rounded-xl border bg-white px-4 py-3 text-left shadow-sm transition-colors
              dark:bg-slate-900 ${
                filtrosAplicados.status === indicador.status
                  ? 'border-swnet-400 dark:border-swnet-600'
                  : 'border-slate-200 hover:border-swnet-300 dark:border-slate-800'
              }`}
          >
            <span className="block text-xs font-medium uppercase tracking-wide text-slate-500 dark:text-slate-400">
              {indicador.status}
            </span>
            <span className="mt-1 block text-2xl font-bold text-slate-800 dark:text-white">
              {indicador.total}
            </span>
          </button>
        ))}
      </div>

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Selecao
            id="filtroStatus"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todos</option>
            {STATUS_CHECKLIST.map((status) => (
              <option key={status} value={status}>
                {status}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroCampo"
            rotulo="Pesquisar por"
            value={filtros.campo}
            onChange={(e) => setFiltros({ ...filtros, campo: e.target.value as FiltrosChecklist['campo'] })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Selecione...</option>
            <option value="NP">Nome do Proprietário</option>
            <option value="NM">Nome do Motorista</option>
            <option value="PL">Placa do Veículo</option>
          </Selecao>

          <Entrada
            id="filtroValor"
            rotulo="Valor"
            value={filtros.valor}
            disabled={!filtros.campo}
            onChange={(e) => setFiltros({ ...filtros, valor: e.target.value })}
            larguraClasse="sm:col-span-4"
          />

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Filtrar">
              <span className="sr-only">Filtrar</span>
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              titulo="Limpar filtros"
              aoClicar={() => {
                setFiltros(FILTROS_VAZIOS)
                setFiltrosAplicados(FILTROS_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${checklists.length} checklist(s)`} icone={ClipboardList}>
        <Tabela
          itens={checklists}
          chave={(checklist) => checklist.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (c) => c.id, className: 'w-20' },
            { cabecalho: 'Proprietário', conteudo: (c) => c.nomeProprietario },
            { cabecalho: 'Motorista', conteudo: (c) => c.nomeMotorista },
            { cabecalho: 'Placa', conteudo: (c) => c.placaVeiculo },
            { cabecalho: 'Destino', conteudo: (c) => c.destinoCarga },
            { cabecalho: 'Data', conteudo: (c) => formatarData(c.dataPreenchimento) },
            {
              cabecalho: 'Conclusão',
              conteudo: (c) => <Etiqueta texto={c.conclusao} tom={TOM_CONCLUSAO[c.conclusao]} />
            },
            {
              cabecalho: 'Status',
              conteudo: (c) => <Etiqueta texto={c.status} tom={TOM_STATUS[c.status]} />
            },
            {
              cabecalho: 'Itens',
              conteudo: (c) => (
                <span className="text-xs">
                  {c.itensAtende.length} / {ITENS_CHECKLIST.length}
                </span>
              ),
              alinhamento: 'centro'
            }
          ]}
          acoes={(checklist) => {
            const indiceAtual = STATUS_CHECKLIST.indexOf(checklist.status)
            const proximo = STATUS_CHECKLIST[indiceAtual + 1]

            return (
              <>
                <BotaoAcao
                  icone={ListChecks}
                  titulo="Ver itens avaliados"
                  aoClicar={() => setDetalhe(checklist)}
                />
                {proximo && (
                  <BotaoAcao
                    icone={ChevronRight}
                    titulo={`Avançar para ${proximo}`}
                    aoClicar={() => avancarStatus(checklist, proximo)}
                  />
                )}
                {checklist.status !== 'FINALIZADO' && (
                  <BotaoAcao icone={Pencil} titulo="Editar checklist" aoClicar={() => abrirEdicao(checklist)} />
                )}
                <BotaoAcao
                  icone={Trash2}
                  titulo="Excluir checklist"
                  tom="perigo"
                  aoClicar={() => setExcluindo(checklist)}
                />
              </>
            )
          }}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Checklist' : 'Novo Checklist de Transporte'}
        icone={ClipboardList}
        largura="xl"
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
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

            <Secao titulo="Transportador">
              <div className="grid gap-4 sm:grid-cols-12">
                <div className="sm:col-span-4">
                  <Entrada
                    id="txtPlaca"
                    rotulo="Placa do Veículo"
                    required
                    maxLength={8}
                    value={edicao.placaVeiculo}
                    disabled={salvando}
                    onChange={(e) => mudar('placaVeiculo', e.target.value.toUpperCase())}
                  />
                  <button
                    type="button"
                    onClick={preencherPelaPlaca}
                    disabled={salvando || !edicao.placaVeiculo.trim()}
                    className="mt-1 text-xs font-medium text-swnet-600 hover:underline disabled:opacity-50 dark:text-swnet-400"
                  >
                    Buscar dados do último checklist desta placa
                  </button>
                </div>

                <Entrada
                  id="txtNomeProprietario"
                  rotulo="Nome do Proprietário"
                  required
                  value={edicao.nomeProprietario}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeProprietario', e.target.value)}
                  larguraClasse="sm:col-span-8"
                />

                <Entrada
                  id="txtNomeMotorista"
                  rotulo="Nome do Motorista"
                  required
                  value={edicao.nomeMotorista}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeMotorista', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />
                <Entrada
                  id="txtCnpj"
                  rotulo="CPF / CNPJ"
                  required
                  maxLength={14}
                  value={edicao.documento}
                  disabled={salvando}
                  onChange={(e) => mudar('documento', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="txtTelefone"
                  rotulo="Telefone"
                  value={edicao.telefone}
                  disabled={salvando}
                  onChange={(e) => mudar('telefone', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Veículo e Carregamento">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="cboTipoVeiculo"
                  rotulo="Tipo de Veículo"
                  required
                  value={edicao.idTipoVeiculo}
                  disabled={salvando}
                  onChange={(e) => mudar('idTipoVeiculo', e.target.value ? Number(e.target.value) : '')}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_VEICULO.map((tipo) => (
                    <option key={tipo.id} value={tipo.id}>
                      {tipo.descricao}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="txtQuantidadeEixo"
                  rotulo="Qtd. de Eixos"
                  type="number"
                  min="1"
                  value={edicao.quantidadeEixo}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeEixo', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="txtDataPreenchimento"
                  rotulo="Data de Preenchimento"
                  type="date"
                  required
                  value={edicao.dataPreenchimento}
                  disabled={salvando}
                  onChange={(e) => mudar('dataPreenchimento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="txtHorarioInicio"
                  rotulo="Início Carregamento"
                  type="time"
                  value={edicao.horarioInicio}
                  disabled={salvando}
                  onChange={(e) => mudar('horarioInicio', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="txtHorarioTermino"
                  rotulo="Término Carregamento"
                  type="time"
                  value={edicao.horarioTermino}
                  disabled={salvando}
                  onChange={(e) => mudar('horarioTermino', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="txtDestinoCarga"
                  rotulo="Destino da Carga"
                  required
                  value={edicao.destinoCarga}
                  disabled={salvando}
                  onChange={(e) => mudar('destinoCarga', e.target.value)}
                  larguraClasse="sm:col-span-9"
                />
              </div>
            </Secao>

            <Secao titulo="Itens Avaliados" icone={ListChecks}>
              <div className="mb-3 flex items-center justify-between">
                <span className="text-xs text-slate-400">
                  {edicao.itensAtende.length} de {ITENS_CHECKLIST.length} atendem
                </span>
                <button
                  type="button"
                  disabled={salvando}
                  onClick={() =>
                    mudar(
                      'itensAtende',
                      edicao.itensAtende.length === ITENS_CHECKLIST.length
                        ? []
                        : ITENS_CHECKLIST.map((item) => item.id)
                    )
                  }
                  className="text-xs font-medium text-swnet-600 hover:underline disabled:opacity-50 dark:text-swnet-400"
                >
                  {edicao.itensAtende.length === ITENS_CHECKLIST.length
                    ? 'Desmarcar todos'
                    : 'Marcar todos como Atende'}
                </button>
              </div>

              <div className="overflow-hidden rounded-lg border border-slate-200 dark:border-slate-700">
                <table className="w-full text-left text-sm">
                  <thead className="border-b border-slate-200 bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-400">
                    <tr>
                      <th className="px-3 py-2 font-semibold">Item Avaliado</th>
                      <th className="w-24 px-3 py-2 text-center font-semibold">Atende</th>
                      <th className="w-28 px-3 py-2 text-center font-semibold">Não Atende</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ITENS_CHECKLIST.map((item) => {
                      const atende = edicao.itensAtende.includes(item.id)
                      return (
                        <tr
                          key={item.id}
                          className="border-b border-slate-100 last:border-0 dark:border-slate-800"
                        >
                          <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{item.descricao}</td>
                          <td className="px-3 py-2 text-center">
                            <input
                              type="radio"
                              name={`item-${item.id}`}
                              checked={atende}
                              disabled={salvando}
                              onChange={() => !atende && alternarItem(item.id)}
                              aria-label={`${item.descricao}: atende`}
                              className="h-4 w-4 accent-emerald-600"
                            />
                          </td>
                          <td className="px-3 py-2 text-center">
                            <input
                              type="radio"
                              name={`item-${item.id}`}
                              checked={!atende}
                              disabled={salvando}
                              onChange={() => atende && alternarItem(item.id)}
                              aria-label={`${item.descricao}: não atende`}
                              className="h-4 w-4 accent-red-600"
                            />
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>

              <div className="mt-4">
                <OpcoesRadio
                  rotulo="Conclusão"
                  nome="conclusao"
                  valor={edicao.conclusao}
                  desabilitado={salvando}
                  opcoes={[
                    { valor: 'APROVADO', rotulo: 'Aprovado' },
                    { valor: 'RESSALVA', rotulo: 'Aprovado com ressalva' },
                    { valor: 'REPROVADO', rotulo: 'Reprovado' }
                  ]}
                  aoMudar={(v) => mudar('conclusao', v as DadosChecklist['conclusao'])}
                />
              </div>
            </Secao>
          </>
        )}
      </Modal>

      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Itens avaliados — ${detalhe.placaVeiculo}` : ''}
        icone={ListChecks}
        largura="lg"
        aoFechar={() => setDetalhe(null)}
        rodape={
          <Botao variante="neutro" aoClicar={() => setDetalhe(null)}>
            Fechar
          </Botao>
        }
      >
        {detalhe && (
          <ul className="grid gap-1.5">
            {ITENS_CHECKLIST.map((item) => {
              const atende = detalhe.itensAtende.includes(item.id)
              return (
                <li
                  key={item.id}
                  className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                    dark:border-slate-700"
                >
                  {atende ? (
                    <Check size={15} className="shrink-0 text-emerald-600 dark:text-emerald-400" />
                  ) : (
                    <X size={15} className="shrink-0 text-red-600 dark:text-red-400" />
                  )}
                  <span className="flex-1 text-slate-600 dark:text-slate-300">{item.descricao}</span>
                  <Etiqueta texto={atende ? 'Atende' : 'Não atende'} tom={atende ? 'sucesso' : 'perigo'} />
                </li>
              )
            })}
          </ul>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir checklist"
        mensagem={
          excluindo
            ? `Deseja realmente excluir o checklist da placa ${excluindo.placaVeiculo} (${excluindo.nomeMotorista})?`
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
