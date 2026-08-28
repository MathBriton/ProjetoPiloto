import { useCallback, useEffect, useState } from 'react'
import {
  Ban,
  Building2,
  Contact,
  Eraser,
  FileSearch,
  Landmark,
  Paperclip,
  Pencil,
  Plus,
  Save,
  Search,
  Send,
  Truck,
  Undo2
} from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Campo, Chave, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  CONDICOES_PAGAMENTO,
  CONTAS_CONCILIACAO,
  FILTROS_UNIFICADO_VAZIOS,
  IDIOMAS,
  NATUREZAS,
  SEGMENTOS,
  STATUS_UNIFICADO,
  TIPOS_FRETE,
  TIPOS_INDUSTRIA,
  TIPOS_SOLICITACAO_CLIENTE,
  anexarDocumento,
  avancarUnificado,
  cancelarUnificado,
  descricaoStatusUnificado,
  devolverUnificado,
  listarUnificados,
  salvarUnificado,
  UNIFICADO_VAZIO,
  type CadastroUnificado as Cadastro,
  type FiltrosUnificado
} from '@/services/cliente/clienteApi'
import { PLANTAS_PROTHEUS, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Cliente/Cadastro-Unificado-V2.
 *
 * O formulário legado muda de exigência conforme o tipo: o cadastro
 * completo (tipo 1) pede praticamente todos os campos, enquanto o Prospect
 * (tipo 5) dispensa a maior parte deles.
 */

const TOM_STATUS: Record<string, 'neutro' | 'info' | 'alerta' | 'sucesso' | 'perigo'> = {
  P: 'neutro',
  S: 'info',
  C: 'info',
  E: 'alerta',
  I: 'alerta',
  F: 'sucesso',
  R: 'perigo',
  X: 'perigo',
  Z: 'neutro'
}

type Rascunho = Partial<Cadastro> & { codigo?: number }

export function CadastroUnificado() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosUnificado>(FILTROS_UNIFICADO_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_UNIFICADO_VAZIOS)
  const [cadastros, setCadastros] = useState<Cadastro[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<Rascunho | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')

  const [detalhe, setDetalhe] = useState<Cadastro | null>(null)
  const [encaminhando, setEncaminhando] = useState<{ item: Cadastro; devolver: boolean } | null>(null)
  const [observacao, setObservacao] = useState('')
  const [anexando, setAnexando] = useState<Cadastro | null>(null)
  const [nomeAnexo, setNomeAnexo] = useState('')
  const [cancelando, setCancelando] = useState<Cadastro | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setCadastros(await listarUnificados(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof Cadastro>(campo: C, valor: Cadastro[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  const prospect = edicao?.tipoSolicitacao === '5'

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarUnificado(edicao, usuario?.login ?? '')
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

    const resultado = encaminhando.devolver
      ? await devolverUnificado(encaminhando.item.codigo, observacao, usuario?.login ?? '')
      : await avancarUnificado(encaminhando.item.codigo, observacao, usuario?.login ?? '')

    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setEncaminhando(null)
      setObservacao('')
      await carregar()
    }
  }

  async function confirmarAnexo() {
    if (!anexando) return

    const resultado = await anexarDocumento(anexando.codigo, nomeAnexo)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) {
      setAnexando(null)
      setNomeAnexo('')
      await carregar()
    }
  }

  async function confirmarCancelamento() {
    if (!cancelando) return

    const resultado = await cancelarUnificado(cancelando.codigo, usuario?.login ?? '')
    setCancelando(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Cadastro Unificado"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ ...UNIFICADO_VAZIO })
            }}
          >
            Novo cadastro
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
          <Entrada
            id="filtroClienteUni"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroCodigoUni"
            rotulo="Código"
            value={filtros.codigo}
            onChange={(e) => setFiltros({ ...filtros, codigo: e.target.value })}
            larguraClasse="sm:col-span-2"
          />
          <Selecao
            id="filtroStatusUni"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => setFiltros({ ...filtros, status: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            <option value="">TODOS</option>
            {STATUS_UNIFICADO.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao.toUpperCase()}
              </option>
            ))}
          </Selecao>
          <Selecao
            id="filtroPlantaUni"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todas</option>
            {PLANTAS_PROTHEUS.map((planta) => (
              <option key={planta.codigo} value={planta.codigo}>
                {planta.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroDataInicio"
            rotulo="Data Início"
            type="date"
            value={filtros.dataInicio}
            onChange={(e) => setFiltros({ ...filtros, dataInicio: e.target.value })}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroDataFim"
            rotulo="Data Fim"
            type="date"
            value={filtros.dataFim}
            onChange={(e) => setFiltros({ ...filtros, dataFim: e.target.value })}
            larguraClasse="sm:col-span-3"
          />

          <div className="flex gap-2 sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando} titulo="Pesquisar">
              <span className="sr-only">Pesquisar</span>
            </Botao>
            <Botao
              variante="neutro"
              icone={Eraser}
              titulo="Limpar filtros"
              aoClicar={() => {
                setFiltros(FILTROS_UNIFICADO_VAZIOS)
                setFiltrosAplicados(FILTROS_UNIFICADO_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${cadastros.length} cadastro(s)`} icone={Contact}>
        <Tabela
          itens={cadastros}
          chave={(item) => item.codigo}
          carregando={carregando}
          mensagemVazia="Nenhum cadastro encontrado"
          colunas={[
            { cabecalho: '#', conteudo: (c) => c.codigo, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (c) => nomePlanta(c.planta) },
            { cabecalho: 'Data Inclusão', conteudo: (c) => formatarData(c.dataInclusao) },
            {
              cabecalho: 'Status',
              conteudo: (c) => (
                <Etiqueta texto={descricaoStatusUnificado(c.status)} tom={TOM_STATUS[c.status] ?? 'neutro'} />
              )
            },
            { cabecalho: 'Cliente', conteudo: (c) => c.razaoSocial },
            { cabecalho: 'CPF/CNPJ', conteudo: (c) => c.cnpj || c.cpf || '—' },
            { cabecalho: 'Emitente', conteudo: (c) => c.emitente }
          ]}
          acoes={(item) => {
            const encerrado = item.status === 'X' || item.status === 'F'

            return (
              <>
                <BotaoAcao icone={FileSearch} titulo="Ver detalhes" aoClicar={() => setDetalhe(item)} />
                <BotaoAcao
                  icone={Paperclip}
                  titulo="Anexar documento"
                  aoClicar={() => {
                    setNomeAnexo('')
                    setAnexando(item)
                  }}
                />

                {!encerrado && (
                  <>
                    <BotaoAcao
                      icone={Pencil}
                      titulo="Editar cadastro"
                      aoClicar={() => {
                        setErro('')
                        setEdicao({ ...item })
                      }}
                    />
                    <BotaoAcao
                      icone={Send}
                      titulo="Encaminhar para próxima etapa"
                      aoClicar={() => {
                        setObservacao('')
                        setEncaminhando({ item, devolver: false })
                      }}
                    />
                    <BotaoAcao
                      icone={Undo2}
                      titulo="Devolver ao solicitante"
                      aoClicar={() => {
                        setObservacao('')
                        setEncaminhando({ item, devolver: true })
                      }}
                    />
                    <BotaoAcao
                      icone={Ban}
                      titulo="Cancelar cadastro"
                      tom="perigo"
                      aoClicar={() => setCancelando(item)}
                    />
                  </>
                )}
              </>
            )
          }}
        />
      </Painel>

      {/* Formulário */}
      <Modal
        aberto={edicao !== null}
        titulo={edicao?.codigo ? `Editar cadastro ${edicao.codigo}` : 'Novo Cadastro Unificado'}
        icone={Contact}
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

            {prospect && (
              <p className="mb-4 rounded-lg bg-swnet-50 px-3 py-2 text-xs text-swnet-700 dark:bg-swnet-950/40 dark:text-swnet-300">
                Cadastro Prospect: a maior parte dos campos fica opcional.
              </p>
            )}

            <Secao titulo="Tipo e Identificação">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaUni"
                  rotulo="Planta"
                  required
                  value={edicao.planta ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('planta', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_PROTHEUS.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="tipoSolicitacaoUni"
                  rotulo="Tipo de Cadastro"
                  required
                  value={edicao.tipoSolicitacao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoSolicitacao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_SOLICITACAO_CLIENTE.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.descricao}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="eshop"
                  rotulo="Eshop"
                  required={!prospect}
                  value={edicao.eshop ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('eshop', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="grupoContas"
                  rotulo="Grupo de Contas"
                  required={!prospect}
                  value={edicao.grupoContas ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('grupoContas', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Entrada
                  id="empresaUni"
                  rotulo="Empresa"
                  required
                  value={edicao.empresa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('empresa', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="razaoSocialUni"
                  rotulo="Razão Social"
                  required
                  value={edicao.razaoSocial ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('razaoSocial', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />
                <Entrada
                  id="nomeFantasia"
                  rotulo="Nome Fantasia"
                  required={!prospect}
                  value={edicao.nomeFantasia ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeFantasia', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="cnpjUni"
                  rotulo="CNPJ"
                  maxLength={14}
                  value={edicao.cnpj ?? ''}
                  disabled={salvando}
                  ajuda="Informe CNPJ ou CPF"
                  onChange={(e) => mudar('cnpj', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="cpfUni"
                  rotulo="CPF"
                  maxLength={11}
                  value={edicao.cpf ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('cpf', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="inscricaoEstadual"
                  rotulo="Inscrição Estadual"
                  value={edicao.inscricaoEstadual ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('inscricaoEstadual', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="inscricaoMunicipal"
                  rotulo="Inscrição Municipal"
                  value={edicao.inscricaoMunicipal ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('inscricaoMunicipal', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Endereço" icone={Building2}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="ruaUni"
                  rotulo="Rua"
                  required
                  value={edicao.rua ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('rua', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />
                <Entrada
                  id="numeroUni"
                  rotulo="Número"
                  required
                  value={edicao.numero ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('numero', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="bairroUni"
                  rotulo="Bairro"
                  required
                  value={edicao.bairro ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('bairro', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="cepUni"
                  rotulo="CEP"
                  required
                  maxLength={8}
                  value={edicao.cep ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('cep', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-2"
                />

                <Entrada
                  id="estadoUni"
                  rotulo="Estado"
                  required
                  maxLength={2}
                  value={edicao.estado ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('estado', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="cidadeUni"
                  rotulo="Cidade"
                  required={!prospect}
                  value={edicao.cidade ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('cidade', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="paisUni"
                  rotulo="País"
                  required={!prospect}
                  value={edicao.pais ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('pais', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="kmDestino"
                  rotulo="KM até o destino"
                  required={!prospect}
                  value={edicao.kmDestino ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('kmDestino', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Selecao
                  id="idiomaUni"
                  rotulo="Idioma"
                  required={!prospect}
                  value={edicao.idioma ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('idioma', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {IDIOMAS.map((idioma) => (
                    <option key={idioma} value={idioma}>
                      {idioma}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="telefoneUni"
                  rotulo="Telefone"
                  value={edicao.telefone ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('telefone', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="emailOrdemPagamento"
                  rotulo="E-mail para ordem de pagamento"
                  type="email"
                  value={edicao.emailOrdemPagamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('emailOrdemPagamento', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />
                <Entrada
                  id="codigoHfm"
                  rotulo="Código HFM (Intercompany)"
                  value={edicao.codigoHfm ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoHfm', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Fiscal" icone={Landmark}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoIndustria"
                  rotulo="Tipo de Indústria"
                  value={edicao.tipoIndustria ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoIndustria', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_INDUSTRIA.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="codigoSuframa"
                  rotulo="Código Suframa"
                  value={edicao.codigoSuframa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoSuframa', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <div className="grid gap-2 sm:col-span-12 sm:grid-cols-4">
                  {(
                    [
                      ['declarantePisCofins', 'Declarante de PIS/COFINS'],
                      ['contribuinteIcms', 'Contribuinte de ICMS'],
                      ['icmsDiferido', 'ICMS Diferido'],
                      ['suspensaoIpi', 'Suspensão/Isenção do IPI']
                    ] as [keyof Cadastro, string][]
                  ).map(([campo, rotulo]) => (
                    <Chave
                      key={campo}
                      rotulo={rotulo}
                      marcado={Boolean(edicao[campo])}
                      desabilitado={salvando}
                      aoAlternar={(v) => mudar(campo, v as never)}
                    />
                  ))}
                </div>
              </div>
            </Secao>

            <Secao titulo="Financeiro">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="contaConciliacao"
                  rotulo="Conta Conciliação"
                  value={edicao.contaConciliacao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('contaConciliacao', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {CONTAS_CONCILIACAO.map((conta) => (
                    <option key={conta} value={conta}>
                      {conta}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="chaveOrdenacao"
                  rotulo="Chave de Ordenação"
                  value={edicao.chaveOrdenacao ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('chaveOrdenacao', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Selecao
                  id="condicaoPagamentoUni"
                  rotulo="Condição de Pagamento"
                  value={edicao.condicaoPagamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('condicaoPagamento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {CONDICOES_PAGAMENTO.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="metodoPagamento"
                  rotulo="Método de Pagamento"
                  value={edicao.metodoPagamento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('metodoPagamento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />

                <Entrada
                  id="codigoJuros"
                  rotulo="Código de Juros"
                  value={edicao.codigoJuros ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoJuros', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="correspondenciaEmpresa"
                  rotulo="Correspondência (Empresa)"
                  value={edicao.correspondenciaEmpresa ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('correspondenciaEmpresa', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="limiteCreditoUni"
                  rotulo="Limite de Crédito"
                  value={edicao.limiteCredito ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('limiteCredito', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="volumeProjetadoM2"
                  rotulo="Volume projetado mensal (M²)"
                  value={edicao.volumeProjetadoM2 ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('volumeProjetadoM2', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Comercial e Logística" icone={Truck}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="segmentoUni"
                  rotulo="Segmento"
                  value={edicao.segmento ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('segmento', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {SEGMENTOS.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="codigoSap"
                  rotulo="Código SAP"
                  value={edicao.codigoSap ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoSap', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Selecao
                  id="naturezaUni"
                  rotulo="Natureza"
                  value={edicao.natureza ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('natureza', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {NATUREZAS.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </Selecao>
                <Entrada
                  id="tipoUni"
                  rotulo="Tipo"
                  value={edicao.tipo ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipo', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Selecao
                  id="tipoFreteUni"
                  rotulo="Tipo de Frete"
                  value={edicao.tipoFrete ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tipoFrete', e.target.value)}
                  larguraClasse="sm:col-span-2"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_FRETE.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="tesUni"
                  rotulo="TES"
                  value={edicao.tes ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('tes', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="codigoVendedorUni"
                  rotulo="Código do Vendedor"
                  value={edicao.codigoVendedor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoVendedor', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="transitTime"
                  rotulo="Transit Time (dias)"
                  value={edicao.transitTime ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('transitTime', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="toleranciaMaior"
                  rotulo="Tolerância entrega a maior"
                  value={edicao.toleranciaEntregaMaior ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('toleranciaEntregaMaior', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="toleranciaMenor"
                  rotulo="Tolerância entrega a menor"
                  value={edicao.toleranciaEntregaMenor ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('toleranciaEntregaMenor', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
              </div>
            </Secao>

            <Secao titulo="Endereço de Entrega">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="entregaRua"
                  rotulo="Rua"
                  value={edicao.entregaRua ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaRua', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="entregaNumero"
                  rotulo="Número"
                  value={edicao.entregaNumero ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaNumero', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="entregaBairro"
                  rotulo="Bairro"
                  value={edicao.entregaBairro ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaBairro', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="entregaCep"
                  rotulo="CEP"
                  maxLength={8}
                  value={edicao.entregaCep ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaCep', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="entregaEstado"
                  rotulo="Estado"
                  maxLength={2}
                  value={edicao.entregaEstado ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaEstado', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-2"
                />
                <Entrada
                  id="entregaCidade"
                  rotulo="Cidade"
                  value={edicao.entregaCidade ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaCidade', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="entregaKm"
                  rotulo="KM"
                  value={edicao.entregaKm ?? ''}
                  disabled={salvando}
                  onChange={(e) => mudar('entregaKm', e.target.value)}
                  larguraClasse="sm:col-span-2"
                />
              </div>
            </Secao>
          </>
        )}
      </Modal>

      {/* Detalhes */}
      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Cadastro ${detalhe.codigo}` : ''}
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
                ['Razão Social', detalhe.razaoSocial],
                ['Nome Fantasia', detalhe.nomeFantasia || '—'],
                ['CNPJ/CPF', detalhe.cnpj || detalhe.cpf || '—'],
                ['Planta', nomePlanta(detalhe.planta)],
                ['Cidade/UF', `${detalhe.cidade} / ${detalhe.estado}`],
                ['Segmento', detalhe.segmento || '—'],
                ['Cond. Pagamento', detalhe.condicaoPagamento || '—'],
                ['Limite de Crédito', detalhe.limiteCredito || '—'],
                ['Status', descricaoStatusUnificado(detalhe.status)]
              ].map(([rotulo, valor]) => (
                <div key={rotulo}>
                  <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                  <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                    {valor}
                  </dd>
                </div>
              ))}
            </dl>

            {detalhe.anexos.length > 0 && (
              <div>
                <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                  Anexos
                </p>
                <ul className="grid gap-1.5">
                  {detalhe.anexos.map((anexo) => (
                    <li
                      key={anexo}
                      className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                        text-slate-600 dark:border-slate-700 dark:text-slate-300"
                    >
                      <Paperclip size={14} className="shrink-0 text-slate-400" />
                      {anexo}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Histórico de vistos
              </p>
              <div className="rounded-lg border border-slate-200 dark:border-slate-700">
                <Tabela
                  itens={detalhe.vistos}
                  chave={(v) => v.data + v.acao}
                  mensagemVazia="Nenhum visto registrado"
                  colunas={[
                    { cabecalho: 'Data', conteudo: (v) => formatarData(v.data) },
                    { cabecalho: 'Usuário', conteudo: (v) => v.usuario },
                    { cabecalho: 'Ação', conteudo: (v) => v.acao },
                    { cabecalho: 'Observação', conteudo: (v) => v.observacao || '—' }
                  ]}
                />
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Encaminhar / devolver */}
      <Modal
        aberto={encaminhando !== null}
        titulo={encaminhando?.devolver ? 'Devolver ao solicitante' : 'Encaminhar para próxima etapa'}
        icone={encaminhando?.devolver ? Undo2 : Send}
        aoFechar={() => setEncaminhando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEncaminhando(null)}>
              Cancelar
            </Botao>
            <Botao
              variante={encaminhando?.devolver ? 'perigo' : 'sucesso'}
              aoClicar={confirmarEncaminhamento}
            >
              Confirmar
            </Botao>
          </>
        }
      >
        {encaminhando && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              Cadastro {encaminhando.item.codigo} — {encaminhando.item.razaoSocial}
              <span className="block text-xs text-slate-400">
                Status atual: {descricaoStatusUnificado(encaminhando.item.status)}
              </span>
            </p>

            <Campo
              rotulo="Observação"
              htmlFor="observacaoVisto"
              obrigatorio={encaminhando.devolver}
            >
              <textarea
                id="observacaoVisto"
                rows={3}
                value={observacao}
                placeholder={
                  encaminhando.devolver ? 'Motivo da devolução...' : 'Observação do visto (opcional)'
                }
                onChange={(e) => setObservacao(e.target.value)}
                className={classesControle}
              />
            </Campo>
          </div>
        )}
      </Modal>

      {/* Anexo */}
      <Modal
        aberto={anexando !== null}
        titulo="Anexar documento"
        icone={Paperclip}
        aoFechar={() => setAnexando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAnexando(null)}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={confirmarAnexo}>
              Anexar
            </Botao>
          </>
        }
      >
        <Entrada
          id="nomeAnexo"
          rotulo="Nome do documento"
          required
          placeholder="ex.: contrato-social.pdf"
          value={nomeAnexo}
          ajuda="O upload real depende do backend"
          onChange={(e) => setNomeAnexo(e.target.value)}
        />
      </Modal>

      <ConfirmDialog
        aberto={cancelando !== null}
        titulo="Cancelar cadastro"
        mensagem={cancelando ? `Deseja realmente cancelar o cadastro ${cancelando.codigo}?` : ''}
        rotuloConfirmar="Cancelar cadastro"
        aoConfirmar={confirmarCancelamento}
        aoCancelar={() => setCancelando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
