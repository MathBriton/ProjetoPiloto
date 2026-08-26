import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  ClipboardCheck,
  ClipboardList,
  Eraser,
  FileSearch,
  Pencil,
  Plus,
  Save,
  Search,
  Send
} from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  DADOS_FNC_VAZIOS,
  DISPOSICOES,
  FILTROS_FNC_VAZIOS,
  SITUACOES_FNC,
  STATUS_FNC,
  TIPOS_DEFEITO,
  TIPOS_RECLAMACAO,
  avaliarFnc,
  calcularNqa,
  cancelarFnc,
  enviarParaResponsavel,
  listarDefeitosPorTipo,
  listarFncs,
  listarSetoresSgq,
  responderFnc,
  salvarFnc,
  type DadosFnc,
  type FiltrosFnc,
  type Fnc,
  type NotaVinculada
} from '@/services/sgq/sgqApi'
import { PLANTAS_CREDITO, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de SGQ/FNC-Brasil.
 *
 * O legado tem três formulários para a mesma FNC, conforme o papel do
 * usuário na tabela Responsavel (a mesma do Cadastros-FNC):
 * - Solicitante: registra a FNC e a envia ao responsável
 * - Responsável setorial: responde se procede, com justificativa
 * - Analista: avalia a eficácia e encerra
 */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })
const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

type Papel = 'SOLICITANTE' | 'RESPONSAVEL' | 'ANALISTA'

const SETORES = listarSetoresSgq()

function descricaoSituacao(codigo: string) {
  return SITUACOES_FNC.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

function descricaoStatus(codigo: string) {
  return STATUS_FNC.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

const TOM_SITUACAO: Record<string, 'neutro' | 'info' | 'sucesso' | 'perigo' | 'alerta'> = {
  '1': 'neutro',
  '2': 'info',
  '3': 'perigo',
  '4': 'sucesso',
  '5': 'alerta'
}

export function FncBrasil() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  /** Viria do cadastro de Responsável de FNC; sem backend, é escolhido aqui */
  const [papel, setPapel] = useState<Papel>('SOLICITANTE')

  const [filtros, setFiltros] = useState<FiltrosFnc>(FILTROS_FNC_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosFnc>(FILTROS_FNC_VAZIOS)
  const [fncs, setFncs] = useState<Fnc[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<DadosFnc | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const [detalhe, setDetalhe] = useState<Fnc | null>(null)
  const [encaminhando, setEncaminhando] = useState<Fnc | null>(null)
  const [responsavelDestino, setResponsavelDestino] = useState('')
  const [respondendo, setRespondendo] = useState<Fnc | null>(null)
  const [resposta, setResposta] = useState({ procede: true, descricao: '', justificativa: '' })
  const [avaliando, setAvaliando] = useState<Fnc | null>(null)
  const [avaliacao, setAvaliacao] = useState({ eficacia: '', aprovada: true })
  const [cancelando, setCancelando] = useState<Fnc | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setFncs(await listarFncs(filtrosAplicados, usuario?.login ?? ''))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados, usuario?.login])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosFnc>(campo: C, valor: FiltrosFnc[C]) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function mudar<C extends keyof DadosFnc>(campo: C, valor: DadosFnc[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function abrirEdicao(fnc: Fnc) {
    setErro('')
    setEdicao({
      id: fnc.id,
      codPlanta: fnc.codPlanta,
      tipoReclamacao: fnc.tipoReclamacao,
      idSetor: fnc.idSetor,
      responsavel: fnc.responsavel,
      codCliente: fnc.codCliente,
      codLoja: fnc.codLoja,
      nomeCliente: fnc.nomeCliente,
      cidadeEstado: fnc.cidadeEstado,
      vendedor: fnc.vendedor,
      dataOcorrencia: fnc.dataOcorrencia,
      tipoDefeito: fnc.tipoDefeito,
      defeitoId: fnc.defeitoId,
      descricaoDefeito: fnc.descricaoDefeito,
      disposicaoId: fnc.disposicaoId,
      descricaoDisposicao: fnc.descricaoDisposicao,
      quantidadeReclamada: String(fnc.quantidadeReclamada),
      valorReclamado: String(fnc.valorReclamado),
      contato: fnc.contato,
      email: fnc.email,
      telefone: fnc.telefone,
      notas: [...fnc.notas]
    })
  }

  function adicionarNota() {
    setEdicao((atual) =>
      atual
        ? {
            ...atual,
            notas: [...atual.notas, { numero: '', serie: '1', quantidade: 0, valor: 0 }]
          }
        : atual
    )
  }

  function mudarNota(indice: number, campo: keyof NotaVinculada, valor: string) {
    setEdicao((atual) => {
      if (!atual) return atual
      const notas = [...atual.notas]
      notas[indice] = {
        ...notas[indice],
        [campo]: campo === 'numero' || campo === 'serie' ? valor : Number(valor || 0)
      }
      return { ...atual, notas }
    })
  }

  function removerNota(indice: number) {
    setEdicao((atual) =>
      atual ? { ...atual, notas: atual.notas.filter((_, i) => i !== indice) } : atual
    )
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarFnc(edicao, usuario?.login ?? '')
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

  async function confirmarEncaminhamento() {
    if (!encaminhando) return

    const resultado = await enviarParaResponsavel(encaminhando.id, responsavelDestino)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setEncaminhando(null)
      setResponsavelDestino('')
      await carregar()
    }
  }

  async function confirmarResposta() {
    if (!respondendo) return

    const resultado = await responderFnc(
      respondendo.id,
      resposta.procede,
      resposta.descricao,
      resposta.justificativa,
      usuario?.login ?? ''
    )
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setRespondendo(null)
      setResposta({ procede: true, descricao: '', justificativa: '' })
      await carregar()
    }
  }

  async function confirmarAvaliacao() {
    if (!avaliando) return

    const resultado = await avaliarFnc(avaliando.id, avaliacao.eficacia, avaliacao.aprovada)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setAvaliando(null)
      setAvaliacao({ eficacia: '', aprovada: true })
      await carregar()
    }
  }

  async function confirmarCancelamento() {
    if (!cancelando) return

    const resultado = await cancelarFnc(cancelando.id)
    setCancelando(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  // NQA em tempo real enquanto o analista edita, como no formulário legado
  const totalNotas = edicao?.notas.reduce((soma, nota) => soma + nota.quantidade, 0) ?? 0
  const nqa = edicao
    ? calcularNqa(edicao.tipoDefeito, Number(edicao.quantidadeReclamada || 0), totalNotas)
    : null

  const defeitosDisponiveis = listarDefeitosPorTipo(edicao?.tipoDefeito ?? '')

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="FNC Brasil"
        acoes={
          <>
            <Selecao
              id="papelUsuario"
              rotulo=""
              value={papel}
              onChange={(e) => setPapel(e.target.value as Papel)}
              larguraClasse="w-56"
            >
              <option value="SOLICITANTE">Papel: Solicitante</option>
              <option value="RESPONSAVEL">Papel: Responsável Setorial</option>
              <option value="ANALISTA">Papel: Analista</option>
            </Selecao>
            {papel === 'SOLICITANTE' && (
              <Botao
                variante="sucesso"
                icone={Plus}
                aoClicar={() => {
                  setErro('')
                  setEdicao(DADOS_FNC_VAZIOS)
                }}
              >
                Nova FNC
              </Botao>
            )}
          </>
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
            id="filtroPlantaFnc"
            rotulo="Filial"
            value={filtros.planta}
            onChange={(e) => mudarFiltro('planta', e.target.value)}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Todas as filiais</option>
            {PLANTAS_CREDITO.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroSituacao"
            rotulo="Situação"
            value={filtros.situacao}
            onChange={(e) => mudarFiltro('situacao', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todos</option>
            {SITUACOES_FNC.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroStatusFnc"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => mudarFiltro('status', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todos</option>
            {STATUS_FNC.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="filtroAprovada"
            rotulo="Aprovada"
            value={filtros.aprovada}
            onChange={(e) => mudarFiltro('aprovada', e.target.value)}
            larguraClasse="sm:col-span-1"
          >
            <option value="">Todos</option>
            <option value="1">Sim</option>
            <option value="0">Não</option>
          </Selecao>

          <Selecao
            id="filtroTipoReclamacao"
            rotulo="Tipo"
            value={filtros.tipoReclamacao}
            onChange={(e) => mudarFiltro('tipoReclamacao', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todos</option>
            {TIPOS_RECLAMACAO.map((tipo) => (
              <option key={tipo} value={tipo}>
                {tipo}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroIdFNC"
            rotulo="Id FNC"
            value={filtros.idFnc}
            onChange={(e) => mudarFiltro('idFnc', e.target.value)}
            larguraClasse="sm:col-span-2"
          />

          <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-600 sm:col-span-6 dark:text-slate-300">
            <input
              type="checkbox"
              checked={filtros.apenasMinhasPendencias}
              onChange={(e) => mudarFiltro('apenasMinhasPendencias', e.target.checked)}
              className="h-4 w-4 accent-swnet-600"
            />
            Apenas minhas pendências
          </label>

          <div className="flex gap-2 sm:col-span-6 sm:justify-end">
            <Botao tipo="submit" icone={Search} desabilitado={carregando}>
              Pesquisar
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              aoClicar={() => {
                setFiltros(FILTROS_FNC_VAZIOS)
                setFiltrosAplicados(FILTROS_FNC_VAZIOS)
              }}
            >
              Limpar
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${fncs.length} FNC(s)`} icone={ClipboardList}>
        <Tabela
          itens={fncs}
          chave={(fnc) => fnc.id}
          carregando={carregando}
          mensagemVazia="Nenhuma FNC encontrada para os filtros informados"
          colunas={[
            { cabecalho: '#', conteudo: (f) => f.id, className: 'w-20' },
            { cabecalho: 'Data Inclusão', conteudo: (f) => formatarData(f.dataInclusao) },
            { cabecalho: 'Filial', conteudo: (f) => nomePlanta(f.codPlanta) },
            { cabecalho: 'Tipo de Reclamação', conteudo: (f) => f.tipoReclamacao },
            { cabecalho: 'Setor', conteudo: (f) => f.nomeSetor },
            { cabecalho: 'Responsável', conteudo: (f) => f.responsavel },
            { cabecalho: 'Cliente', conteudo: (f) => f.nomeCliente },
            { cabecalho: 'Loja', conteudo: (f) => f.codLoja },
            {
              cabecalho: 'Quantidade',
              conteudo: (f) => numero.format(f.quantidadeReclamada),
              alinhamento: 'direita'
            },
            { cabecalho: 'Disposição', conteudo: (f) => f.descricaoDisposicao },
            { cabecalho: 'Defeito', conteudo: (f) => f.descricaoDefeito },
            {
              cabecalho: 'Aprovada',
              conteudo: (f) => (
                <Etiqueta texto={f.aprovada ? 'Sim' : 'Não'} tom={f.aprovada ? 'sucesso' : 'neutro'} />
              )
            },
            {
              cabecalho: 'Coef. NQA',
              conteudo: (f) => {
                const total = f.notas.reduce((s, n) => s + n.quantidade, 0)
                const resultado = calcularNqa(f.tipoDefeito, f.quantidadeReclamada, total)
                if (resultado.classificacao === 'INDEFINIDO') return '—'
                return (
                  <Etiqueta
                    texto={`${resultado.percentual.toFixed(2)}% ${resultado.classificacao}`}
                    tom={resultado.classificacao === 'APROVADO' ? 'sucesso' : 'perigo'}
                  />
                )
              }
            },
            { cabecalho: 'Status', conteudo: (f) => descricaoStatus(f.status) },
            {
              cabecalho: 'Situação',
              conteudo: (f) => (
                <Etiqueta texto={descricaoSituacao(f.situacao)} tom={TOM_SITUACAO[f.situacao] ?? 'neutro'} />
              )
            }
          ]}
          acoes={(fnc) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(fnc)} />

              {papel === 'SOLICITANTE' && fnc.situacao === '1' && (
                <>
                  <BotaoAcao icone={Pencil} titulo="Editar FNC" aoClicar={() => abrirEdicao(fnc)} />
                  <BotaoAcao
                    icone={Send}
                    titulo="Enviar ao responsável"
                    aoClicar={() => {
                      setResponsavelDestino(fnc.responsavel)
                      setEncaminhando(fnc)
                    }}
                  />
                </>
              )}

              {papel === 'RESPONSAVEL' && fnc.situacao === '2' && (
                <BotaoAcao
                  icone={ClipboardCheck}
                  titulo="Responder FNC"
                  aoClicar={() => setRespondendo(fnc)}
                />
              )}

              {papel === 'ANALISTA' && fnc.status === '2' && (
                <BotaoAcao
                  icone={ClipboardCheck}
                  titulo="Avaliar eficácia"
                  aoClicar={() => setAvaliando(fnc)}
                />
              )}

              {papel !== 'RESPONSAVEL' && fnc.status !== '3' && (
                <BotaoAcao
                  icone={Ban}
                  titulo="Cancelar FNC"
                  tom="perigo"
                  aoClicar={() => setCancelando(fnc)}
                />
              )}
            </>
          )}
        />
      </Painel>

      {/* Formulário do solicitante / analista */}
      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar FNC ${edicao.id}` : 'Nova FNC'}
        icone={ClipboardList}
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

            <Secao titulo="Identificação">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="codPlanta"
                  rotulo="Planta"
                  required
                  value={edicao.codPlanta}
                  disabled={salvando}
                  onChange={(e) => mudar('codPlanta', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_CREDITO.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="tipoReclamacao"
                  rotulo="Tipo de Reclamação"
                  required
                  value={edicao.tipoReclamacao}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoReclamacao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_RECLAMACAO.map((tipo) => (
                    <option key={tipo} value={tipo}>
                      {tipo}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="idSetor"
                  rotulo="Setor"
                  value={edicao.idSetor}
                  disabled={salvando}
                  onChange={(e) => mudar('idSetor', e.target.value ? Number(e.target.value) : '')}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {SETORES.map((setor) => (
                    <option key={setor.id} value={setor.id}>
                      {setor.nome}
                    </option>
                  ))}
                </Selecao>
              </div>
            </Secao>

            <Secao titulo="Cliente">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="nomeCliente"
                  rotulo="Nome do Cliente"
                  required
                  value={edicao.nomeCliente}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeCliente', e.target.value)}
                  larguraClasse="sm:col-span-6"
                />
                <Entrada
                  id="cidadeEstado"
                  rotulo="Cidade/Estado"
                  value={edicao.cidadeEstado}
                  disabled={salvando}
                  onChange={(e) => mudar('cidadeEstado', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="codCliente"
                  rotulo="Cód. Cliente"
                  required
                  value={edicao.codCliente}
                  disabled={salvando}
                  onChange={(e) => mudar('codCliente', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="codLoja"
                  rotulo="Loja"
                  maxLength={2}
                  value={edicao.codLoja}
                  disabled={salvando}
                  onChange={(e) => mudar('codLoja', e.target.value)}
                  larguraClasse="sm:col-span-1"
                />

                <Entrada
                  id="vendedor"
                  rotulo="Vendedor"
                  required
                  value={edicao.vendedor}
                  disabled={salvando}
                  onChange={(e) => mudar('vendedor', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="dataOcorrencia"
                  rotulo="Data de Ocorrência"
                  type="date"
                  required
                  value={edicao.dataOcorrencia}
                  disabled={salvando}
                  onChange={(e) => mudar('dataOcorrencia', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="contato"
                  rotulo="Contato"
                  value={edicao.contato}
                  disabled={salvando}
                  onChange={(e) => mudar('contato', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="email"
                  rotulo="E-mail"
                  type="email"
                  value={edicao.email}
                  disabled={salvando}
                  onChange={(e) => mudar('email', e.target.value)}
                  larguraClasse="sm:col-span-6"
                />
                <Entrada
                  id="telefone"
                  rotulo="Telefone"
                  value={edicao.telefone}
                  disabled={salvando}
                  onChange={(e) => mudar('telefone', e.target.value)}
                  larguraClasse="sm:col-span-6"
                />
              </div>
            </Secao>

            <Secao titulo="Defeito e Disposição">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoDefeito"
                  rotulo="Tipo de Defeito"
                  required
                  value={edicao.tipoDefeito}
                  disabled={salvando}
                  onChange={(e) => {
                    mudar('tipoDefeito', e.target.value)
                    mudar('defeitoId', '')
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_DEFEITO.map((tipo) => (
                    <option key={tipo.codigo} value={tipo.codigo}>
                      {tipo.descricao}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="defeitoId"
                  rotulo="Classificação Defeito"
                  required
                  value={edicao.defeitoId}
                  disabled={salvando || !edicao.tipoDefeito}
                  onChange={(e) => mudar('defeitoId', e.target.value ? Number(e.target.value) : '')}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {defeitosDisponiveis.map((defeito) => (
                    <option key={defeito.id} value={defeito.id}>
                      {defeito.descricao}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="disposicaoId"
                  rotulo="Código Disposição"
                  required
                  value={edicao.disposicaoId}
                  disabled={salvando}
                  onChange={(e) => mudar('disposicaoId', e.target.value ? Number(e.target.value) : '')}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {DISPOSICOES.map((disposicao) => (
                    <option key={disposicao.id} value={disposicao.id}>
                      {disposicao.codigo} — {disposicao.descricao}
                    </option>
                  ))}
                </Selecao>

                <Campo rotulo="Descrição Defeito" htmlFor="descricaoDefeito" obrigatorio className="sm:col-span-12">
                  <textarea
                    id="descricaoDefeito"
                    rows={3}
                    value={edicao.descricaoDefeito}
                    disabled={salvando}
                    onChange={(e) => mudar('descricaoDefeito', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Notas Vinculadas">
              <div className="mb-3 flex items-center justify-between">
                <span className="text-xs text-slate-400">
                  Total das notas: {numero.format(totalNotas)}
                </span>
                <Botao variante="neutro" icone={Plus} aoClicar={adicionarNota} desabilitado={salvando}>
                  Vincular nota
                </Botao>
              </div>

              {edicao.notas.length === 0 ? (
                <p className="rounded-lg border border-dashed border-slate-200 px-3 py-6 text-center text-xs text-slate-400 dark:border-slate-700">
                  Nenhuma nota vinculada.
                </p>
              ) : (
                <div className="grid gap-2">
                  {edicao.notas.map((nota, indice) => (
                    <div key={indice} className="grid items-end gap-2 sm:grid-cols-12">
                      <Entrada
                        id={`nota-${indice}`}
                        rotulo="Nota Fiscal"
                        value={nota.numero}
                        disabled={salvando}
                        onChange={(e) => mudarNota(indice, 'numero', e.target.value)}
                        larguraClasse="sm:col-span-3"
                      />
                      <Entrada
                        id={`serie-${indice}`}
                        rotulo="Série"
                        maxLength={3}
                        value={nota.serie}
                        disabled={salvando}
                        onChange={(e) => mudarNota(indice, 'serie', e.target.value)}
                        larguraClasse="sm:col-span-2"
                      />
                      <Entrada
                        id={`qtd-${indice}`}
                        rotulo="Quantidade"
                        type="number"
                        min="0"
                        value={String(nota.quantidade)}
                        disabled={salvando}
                        onChange={(e) => mudarNota(indice, 'quantidade', e.target.value)}
                        larguraClasse="sm:col-span-3"
                      />
                      <Entrada
                        id={`valor-${indice}`}
                        rotulo="Valor"
                        type="number"
                        step="0.01"
                        value={String(nota.valor)}
                        disabled={salvando}
                        onChange={(e) => mudarNota(indice, 'valor', e.target.value)}
                        larguraClasse="sm:col-span-3"
                      />
                      <div className="sm:col-span-1">
                        <Botao
                          variante="perigo"
                          aoClicar={() => removerNota(indice)}
                          desabilitado={salvando}
                          titulo="Remover nota"
                        >
                          &times;
                        </Botao>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Secao>

            <Secao titulo="Reclamação">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="quantidadeReclamada"
                  rotulo="Quantidade Reclamada Total"
                  type="number"
                  min="0"
                  value={edicao.quantidadeReclamada}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeReclamada', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="valorReclamado"
                  rotulo="Valor Reclamado Total"
                  type="number"
                  step="0.01"
                  value={edicao.valorReclamado}
                  disabled={salvando}
                  onChange={(e) => mudar('valorReclamado', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                {nqa && nqa.classificacao !== 'INDEFINIDO' && (
                  <div className="sm:col-span-4">
                    <span className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300">
                      Coeficiente NQA
                    </span>
                    <div
                      className={`rounded-lg px-3 py-2 text-sm ${
                        nqa.classificacao === 'APROVADO'
                          ? 'bg-emerald-50 text-emerald-700 dark:bg-emerald-500/10 dark:text-emerald-400'
                          : 'bg-red-50 text-red-700 dark:bg-red-500/10 dark:text-red-400'
                      }`}
                    >
                      <strong>{nqa.percentual.toFixed(2)}%</strong> — {nqa.classificacao}
                      <span className="block text-xs opacity-80">{nqa.detalhe}</span>
                    </div>
                  </div>
                )}
              </div>
            </Secao>
          </>
        )}
      </Modal>

      {/* Detalhes */}
      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `FNC ${detalhe.id}` : ''}
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
                ['Cliente', `${detalhe.codCliente}/${detalhe.codLoja} — ${detalhe.nomeCliente}`],
                ['Filial', nomePlanta(detalhe.codPlanta)],
                ['Setor', detalhe.nomeSetor],
                ['Responsável', detalhe.responsavel],
                ['Vendedor', detalhe.vendedor],
                ['Ocorrência', formatarData(detalhe.dataOcorrencia)],
                ['Defeito', detalhe.descricaoDefeito],
                ['Disposição', detalhe.descricaoDisposicao],
                ['Qtd. Reclamada', numero.format(detalhe.quantidadeReclamada)],
                ['Valor Reclamado', moeda.format(detalhe.valorReclamado)],
                ['Situação', descricaoSituacao(detalhe.situacao)],
                ['Status', descricaoStatus(detalhe.status)]
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            {detalhe.eficacia && (
              <p className="rounded-lg bg-slate-50 px-3 py-2 text-sm text-slate-600 dark:bg-slate-800 dark:text-slate-300">
                <strong>Eficácia:</strong> {detalhe.eficacia}
              </p>
            )}

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Notas vinculadas
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.notas}
                  chave={(nota) => `${nota.numero}-${nota.serie}`}
                  mensagemVazia="Nenhuma nota vinculada"
                  colunas={[
                    { cabecalho: 'Nota', conteudo: (n) => n.numero },
                    { cabecalho: 'Série', conteudo: (n) => n.serie },
                    {
                      cabecalho: 'Quantidade',
                      conteudo: (n) => numero.format(n.quantidade),
                      alinhamento: 'direita'
                    },
                    { cabecalho: 'Valor', conteudo: (n) => moeda.format(n.valor), alinhamento: 'direita' }
                  ]}
                />
              </div>
            </div>

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Histórico de respostas
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.respostas}
                  chave={(r) => r.dataResposta + r.usuarioResposta}
                  mensagemVazia="Nenhuma resposta registrada"
                  colunas={[
                    { cabecalho: 'Data Envio', conteudo: (r) => formatarData(r.dataEnvio) },
                    { cabecalho: 'Data Resposta', conteudo: (r) => formatarData(r.dataResposta) },
                    { cabecalho: 'Usuário Envio', conteudo: (r) => r.usuarioEnvio },
                    { cabecalho: 'Usuário Resposta', conteudo: (r) => r.usuarioResposta },
                    {
                      cabecalho: 'Procede',
                      conteudo: (r) =>
                        r.procede === null ? (
                          '—'
                        ) : (
                          <Etiqueta
                            texto={r.procede ? 'Sim' : 'Não'}
                            tom={r.procede ? 'perigo' : 'sucesso'}
                          />
                        )
                    },
                    { cabecalho: 'Justificativa', conteudo: (r) => r.justificativa }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Encaminhar ao responsável */}
      <Modal
        aberto={encaminhando !== null}
        titulo="Enviar ao responsável"
        icone={Send}
        aoFechar={() => setEncaminhando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEncaminhando(null)}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Send} aoClicar={confirmarEncaminhamento}>
              Enviar
            </Botao>
          </>
        }
      >
        <Entrada
          id="responsavelDestino"
          rotulo="Responsável do setor"
          required
          value={responsavelDestino}
          ajuda="Login do responsável cadastrado em Cadastros FNC"
          onChange={(e) => setResponsavelDestino(e.target.value)}
        />
      </Modal>

      {/* Resposta do responsável setorial */}
      <Modal
        aberto={respondendo !== null}
        titulo={respondendo ? `Responder FNC ${respondendo.id}` : ''}
        icone={ClipboardCheck}
        largura="lg"
        aoFechar={() => setRespondendo(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setRespondendo(null)}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={confirmarResposta}>
              Registrar resposta
            </Botao>
          </>
        }
      >
        <div className="grid gap-4">
          <Campo rotulo="A reclamação procede?" obrigatorio>
            <div className="flex gap-4 pt-1">
              {[
                { valor: true, rotulo: 'Sim, procede' },
                { valor: false, rotulo: 'Não procede' }
              ].map((opcao) => (
                <label
                  key={String(opcao.valor)}
                  className="flex cursor-pointer items-center gap-2 text-sm text-slate-600 dark:text-slate-300"
                >
                  <input
                    type="radio"
                    name="procede"
                    checked={resposta.procede === opcao.valor}
                    onChange={() => setResposta({ ...resposta, procede: opcao.valor })}
                    className="h-4 w-4 accent-swnet-600"
                  />
                  {opcao.rotulo}
                </label>
              ))}
            </div>
          </Campo>

          <Campo rotulo="Descrição da análise" htmlFor="descricaoResposta">
            <textarea
              id="descricaoResposta"
              rows={2}
              value={resposta.descricao}
              onChange={(e) => setResposta({ ...resposta, descricao: e.target.value })}
              className={classesControle}
            />
          </Campo>

          <Campo rotulo="Justificativa" htmlFor="justificativaResposta" obrigatorio>
            <textarea
              id="justificativaResposta"
              rows={3}
              value={resposta.justificativa}
              placeholder="Causa raiz e ação tomada..."
              onChange={(e) => setResposta({ ...resposta, justificativa: e.target.value })}
              className={classesControle}
            />
          </Campo>
        </div>
      </Modal>

      {/* Avaliação do analista */}
      <Modal
        aberto={avaliando !== null}
        titulo={avaliando ? `Avaliar FNC ${avaliando.id}` : ''}
        icone={ClipboardCheck}
        largura="lg"
        aoFechar={() => setAvaliando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAvaliando(null)}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={confirmarAvaliacao}>
              Encerrar FNC
            </Botao>
          </>
        }
      >
        <div className="grid gap-4">
          <Campo rotulo="Eficácia / Observação" htmlFor="eficacia" obrigatorio>
            <textarea
              id="eficacia"
              rows={4}
              value={avaliacao.eficacia}
              placeholder="Avaliação da eficácia da ação tomada..."
              onChange={(e) => setAvaliacao({ ...avaliacao, eficacia: e.target.value })}
              className={classesControle}
            />
          </Campo>

          <label className="flex cursor-pointer items-center gap-2 text-sm text-slate-600 dark:text-slate-300">
            <input
              type="checkbox"
              checked={avaliacao.aprovada}
              onChange={(e) => setAvaliacao({ ...avaliacao, aprovada: e.target.checked })}
              className="h-4 w-4 accent-swnet-600"
            />
            Marcar FNC como aprovada
          </label>
        </div>
      </Modal>

      <ConfirmDialog
        aberto={cancelando !== null}
        titulo="Cancelar FNC"
        mensagem={cancelando ? `Deseja realmente cancelar a FNC ${cancelando.id}?` : ''}
        rotuloConfirmar="Cancelar FNC"
        aoConfirmar={confirmarCancelamento}
        aoCancelar={() => setCancelando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
