import { useCallback, useEffect, useState } from 'react'
import { Eraser, Files, Pencil, Plus, Printer, Save, Search, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  ESCALA_SATISFACAO,
  PLANTAS_PROTHEUS,
  excluirCanhoto,
  listarCanhotos,
  nomePlanta,
  rotuloSatisfacao,
  salvarCanhoto,
  type Canhoto,
  type CampoPesquisaCanhoto,
  type DadosCanhoto,
  type FiltrosCanhoto
} from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Financeiro/Cadastro-Canhoto
 * (grid_canhoto.asp + form_canhoto.asp + impr_canhoto.asp).
 */

const FILTROS_VAZIOS: FiltrosCanhoto = { planta: '00', campo: '', valor: '' }

const CANHOTO_VAZIO: DadosCanhoto = {
  planta: '00',
  notaFiscal: '',
  serie: '',
  data: new Date().toISOString().slice(0, 10),
  placa: '',
  produto: '',
  servico: '',
  observacao: ''
}

/** Nota 4-5 = satisfeito, 3 = neutro, 1-2 = insatisfeito, 0 = sem resposta */
function tomSatisfacao(valor: number) {
  if (valor === 0) return 'neutro' as const
  if (valor >= 4) return 'sucesso' as const
  if (valor === 3) return 'info' as const
  return 'perigo' as const
}

export function CadastroCanhoto() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosCanhoto>(FILTROS_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosCanhoto>(FILTROS_VAZIOS)
  const [canhotos, setCanhotos] = useState<Canhoto[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<DadosCanhoto | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<Canhoto | null>(null)
  const [impressao, setImpressao] = useState<Canhoto | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setCanhotos(await listarCanhotos(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar<C extends keyof DadosCanhoto>(campo: C, valor: DadosCanhoto[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarCanhoto(edicao, usuario?.login ?? '')
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

    const resultado = await excluirCanhoto(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Cadastro de Canhoto"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ ...CANHOTO_VAZIO, planta: filtrosAplicados.planta })
            }}
          >
            Novo canhoto
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
            id="cboPlanta"
            rotulo="Planta"
            value={filtros.planta}
            onChange={(e) => setFiltros({ ...filtros, planta: e.target.value })}
            larguraClasse="sm:col-span-3"
          >
            {PLANTAS_PROTHEUS.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="cboPesquisa"
            rotulo="Pesquisar por"
            value={filtros.campo}
            onChange={(e) =>
              setFiltros({ ...filtros, campo: e.target.value as CampoPesquisaCanhoto, valor: '' })
            }
            larguraClasse="sm:col-span-3"
          >
            <option value="">Selecione...</option>
            <option value="PLACA">Placa</option>
            <option value="DATA">Data</option>
            <option value="NF">Nota Fiscal</option>
          </Selecao>

          <Entrada
            id="valorPesquisa"
            rotulo="Valor"
            type={filtros.campo === 'DATA' ? 'date' : 'text'}
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

      <Painel titulo={`${canhotos.length} canhoto(s)`} icone={Files}>
        <Tabela
          itens={canhotos}
          chave={(canhoto) => canhoto.id}
          carregando={carregando}
          colunas={[
            { cabecalho: 'NF / Série', conteudo: (c) => `${c.notaFiscal} / ${c.serie}` },
            { cabecalho: 'Data', conteudo: (c) => formatarData(c.data) },
            { cabecalho: 'Placa', conteudo: (c) => c.placa },
            {
              cabecalho: 'Produto',
              conteudo: (c) => <Etiqueta texto={rotuloSatisfacao(c.produto)} tom={tomSatisfacao(c.produto)} />
            },
            {
              cabecalho: 'Serviço',
              conteudo: (c) => <Etiqueta texto={rotuloSatisfacao(c.servico)} tom={tomSatisfacao(c.servico)} />
            },
            { cabecalho: 'Incluído por', conteudo: (c) => c.incluidoPor },
            { cabecalho: 'Alterado por', conteudo: (c) => c.alteradoPor || '—' },
            {
              cabecalho: 'Observação',
              conteudo: (c) => (
                <span className="block max-w-48 truncate text-xs" title={c.observacao}>
                  {c.observacao || '—'}
                </span>
              )
            }
          ]}
          acoes={(canhoto) => (
            <>
              <BotaoAcao icone={Printer} titulo="Imprimir canhoto" aoClicar={() => setImpressao(canhoto)} />
              <BotaoAcao
                icone={Pencil}
                titulo="Editar canhoto"
                aoClicar={() => {
                  setErro('')
                  setEdicao({
                    id: canhoto.id,
                    planta: canhoto.planta,
                    notaFiscal: canhoto.notaFiscal,
                    serie: canhoto.serie,
                    data: canhoto.data,
                    placa: canhoto.placa,
                    produto: canhoto.produto,
                    servico: canhoto.servico,
                    observacao: canhoto.observacao
                  })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Apagar canhoto"
                tom="perigo"
                aoClicar={() => setExcluindo(canhoto)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Canhoto' : 'Novo Canhoto'}
        icone={Files}
        largura="lg"
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
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <div className="grid gap-4 sm:grid-cols-12">
              <Selecao
                id="plantaCanhoto"
                rotulo="Planta"
                required
                value={edicao.planta}
                disabled={salvando}
                onChange={(e) => mudar('planta', e.target.value)}
                larguraClasse="sm:col-span-4"
              >
                {PLANTAS_PROTHEUS.map((item) => (
                  <option key={item.codigo} value={item.codigo}>
                    {item.nome}
                  </option>
                ))}
              </Selecao>

              <Entrada
                id="edtNOTA"
                rotulo="Nota Fiscal"
                required
                value={edicao.notaFiscal}
                disabled={salvando}
                onChange={(e) => mudar('notaFiscal', e.target.value.replace(/\D/g, ''))}
                larguraClasse="sm:col-span-3"
              />
              <Entrada
                id="edtSERIE"
                rotulo="Série"
                maxLength={3}
                value={edicao.serie}
                disabled={salvando}
                onChange={(e) => mudar('serie', e.target.value)}
                larguraClasse="sm:col-span-2"
              />
              <Entrada
                id="dataCanhoto"
                rotulo="Data"
                type="date"
                required
                value={edicao.data}
                disabled={salvando}
                onChange={(e) => mudar('data', e.target.value)}
                larguraClasse="sm:col-span-3"
              />

              <Entrada
                id="placaCanhoto"
                rotulo="Placa"
                maxLength={8}
                value={edicao.placa}
                disabled={salvando}
                onChange={(e) => mudar('placa', e.target.value.toUpperCase())}
                larguraClasse="sm:col-span-4"
              />

              <Selecao
                id="edtPRODUTO"
                rotulo="Produto (Carga)"
                required
                value={edicao.produto}
                disabled={salvando}
                onChange={(e) => mudar('produto', e.target.value === '' ? '' : Number(e.target.value))}
                larguraClasse="sm:col-span-4"
              >
                <option value="">Selecione...</option>
                {ESCALA_SATISFACAO.map((item) => (
                  <option key={item.valor} value={item.valor}>
                    {item.rotulo}
                  </option>
                ))}
              </Selecao>

              <Selecao
                id="edtSERVICO"
                rotulo="Serviço (Veículo)"
                required
                value={edicao.servico}
                disabled={salvando}
                onChange={(e) => mudar('servico', e.target.value === '' ? '' : Number(e.target.value))}
                larguraClasse="sm:col-span-4"
              >
                <option value="">Selecione...</option>
                {ESCALA_SATISFACAO.map((item) => (
                  <option key={item.valor} value={item.valor}>
                    {item.rotulo}
                  </option>
                ))}
              </Selecao>

              <Entrada
                id="edtOBS"
                rotulo="Observação"
                value={edicao.observacao}
                disabled={salvando}
                onChange={(e) => mudar('observacao', e.target.value)}
                larguraClasse="sm:col-span-12"
              />
            </div>
          </div>
        )}
      </Modal>

      <Modal
        aberto={impressao !== null}
        titulo="Canhoto"
        icone={Printer}
        largura="md"
        aoFechar={() => setImpressao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setImpressao(null)}>
              Fechar
            </Botao>
            <Botao icone={Printer} aoClicar={() => window.print()}>
              Imprimir
            </Botao>
          </>
        }
      >
        {impressao && (
          <dl className="grid gap-3 sm:grid-cols-2">
            {[
              ['Planta', nomePlanta(impressao.planta)],
              ['Nota Fiscal / Série', `${impressao.notaFiscal} / ${impressao.serie}`],
              ['Data', formatarData(impressao.data)],
              ['Placa', impressao.placa],
              ['Produto (Carga)', rotuloSatisfacao(impressao.produto)],
              ['Serviço (Veículo)', rotuloSatisfacao(impressao.servico)],
              ['Incluído por', impressao.incluidoPor],
              ['Alterado por', impressao.alteradoPor || '—'],
              ['Observação', impressao.observacao || '—']
            ].map(([rotulo, valor]) => (
              <div key={rotulo}>
                <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                <dd className="mt-0.5 text-sm text-slate-700 dark:text-slate-200">{valor}</dd>
              </div>
            ))}
          </dl>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Apagar canhoto"
        mensagem={
          excluindo ? `Deseja realmente apagar o canhoto da NF ${excluindo.notaFiscal}?` : ''
        }
        rotuloConfirmar="Apagar"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
