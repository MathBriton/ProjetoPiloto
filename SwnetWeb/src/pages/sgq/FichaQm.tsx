import { useCallback, useEffect, useState } from 'react'
import { Eraser, Files, Pencil, Plus, Save, Search, Trash2 } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  DADOS_QM_VAZIOS,
  FILTROS_QM_VAZIOS,
  STATUS_QM,
  descricaoStatusQm,
  excluirFichaQm,
  listarFichasQm,
  salvarFichaQm,
  type DadosFichaQm,
  type FichaQm as Ficha,
  type FiltrosQm
} from '@/services/sgq/sgqApi'
import { PLANTAS_CREDITO, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/** Refatoração de SGQ/Ficha-De-QM (devoluções). */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })
const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })

/** Status finais (refaturado/refugado) contam como concluídos */
function tomStatus(codigo: string) {
  if (codigo.startsWith('AGUARDANDO')) return 'alerta' as const
  if (codigo.startsWith('SIMBOLICO')) return 'info' as const
  return 'sucesso' as const
}

export function FichaQm() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [filtros, setFiltros] = useState<FiltrosQm>(FILTROS_QM_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState<FiltrosQm>(FILTROS_QM_VAZIOS)
  const [fichas, setFichas] = useState<Ficha[]>([])
  const [carregando, setCarregando] = useState(true)

  const [edicao, setEdicao] = useState<DadosFichaQm | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<Ficha | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setFichas(await listarFichasQm(filtrosAplicados))
    } finally {
      setCarregando(false)
    }
  }, [filtrosAplicados])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudarFiltro<C extends keyof FiltrosQm>(campo: C, valor: string) {
    setFiltros((atual) => ({ ...atual, [campo]: valor }))
  }

  function mudar<C extends keyof DadosFichaQm>(campo: C, valor: string) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  function abrirEdicao(ficha: Ficha) {
    setErro('')
    setEdicao({
      id: ficha.id,
      codPlanta: ficha.codPlanta,
      idFnc: ficha.idFnc === null ? '' : String(ficha.idFnc),
      status: ficha.status,
      nomeCliente: ficha.nomeCliente,
      modoDefeito: ficha.modoDefeito,
      area: ficha.area,
      ordemProducao: ficha.ordemProducao,
      dataEntradaDevolucao: ficha.dataEntradaDevolucao,
      dataConclusao: ficha.dataConclusao,
      nfDevolucao: ficha.nfDevolucao,
      codigoVa: ficha.codigoVa,
      codigoQm: ficha.codigoQm,
      quantidadePaletes: String(ficha.quantidadePaletes),
      quantidadeRecebida: String(ficha.quantidadeRecebida),
      quantidadeRefugada: String(ficha.quantidadeRefugada),
      valor: String(ficha.valor),
      valorUnitarioMaterial: String(ficha.valorUnitarioMaterial)
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarFichaQm(edicao)
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

    const resultado = await excluirFichaQm(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Fichas de QM — Devoluções"
        acoes={
          <Botao
            variante="sucesso"
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao(DADOS_QM_VAZIOS)
            }}
          >
            Nova ficha
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
            id="filtroPlantaQm"
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

          <Entrada
            id="dataCriacaoDe"
            rotulo="Data Criação De"
            type="date"
            value={filtros.dataCriacaoDe}
            onChange={(e) => mudarFiltro('dataCriacaoDe', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="dataCriacaoAte"
            rotulo="Data Criação Até"
            type="date"
            value={filtros.dataCriacaoAte}
            onChange={(e) => mudarFiltro('dataCriacaoAte', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroClienteQm"
            rotulo="Cliente"
            value={filtros.nomeCliente}
            onChange={(e) => mudarFiltro('nomeCliente', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Selecao
            id="filtroStatusQm"
            rotulo="Status"
            value={filtros.status}
            onChange={(e) => mudarFiltro('status', e.target.value)}
            larguraClasse="sm:col-span-2"
          >
            <option value="">Todos os status</option>
            {STATUS_QM.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="filtroIdQm"
            rotulo="Id Ficha"
            value={filtros.id}
            onChange={(e) => mudarFiltro('id', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroIdFncQm"
            rotulo="Id FNC"
            value={filtros.idFnc}
            onChange={(e) => mudarFiltro('idFnc', e.target.value)}
            larguraClasse="sm:col-span-2"
          />
          <Entrada
            id="filtroOp"
            rotulo="Ordem de Produção"
            value={filtros.ordemProducao}
            onChange={(e) => mudarFiltro('ordemProducao', e.target.value)}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroNfDev"
            rotulo="NF Devolução"
            value={filtros.nfDevolucao}
            onChange={(e) => mudarFiltro('nfDevolucao', e.target.value)}
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
                setFiltros(FILTROS_QM_VAZIOS)
                setFiltrosAplicados(FILTROS_QM_VAZIOS)
              }}
            >
              <span className="sr-only">Limpar</span>
            </Botao>
          </div>
        </form>
      </Painel>

      <Painel titulo={`${fichas.length} ficha(s)`} icone={Files}>
        <Tabela
          itens={fichas}
          chave={(ficha) => ficha.id}
          carregando={carregando}
          mensagemVazia="Nenhuma ficha encontrada para os filtros informados"
          colunas={[
            { cabecalho: '#', conteudo: (f) => f.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (f) => nomePlanta(f.codPlanta) },
            { cabecalho: 'Cliente', conteudo: (f) => f.nomeCliente },
            { cabecalho: 'FNC', conteudo: (f) => f.idFnc ?? '—' },
            {
              cabecalho: 'Status',
              conteudo: (f) => <Etiqueta texto={descricaoStatusQm(f.status)} tom={tomStatus(f.status)} />
            },
            {
              cabecalho: 'Qtd. Paletes',
              conteudo: (f) => numero.format(f.quantidadePaletes),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Qtd. Recebida',
              conteudo: (f) => numero.format(f.quantidadeRecebida),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Qtd. Refugada',
              conteudo: (f) => numero.format(f.quantidadeRefugada),
              alinhamento: 'direita'
            },
            { cabecalho: 'Valor', conteudo: (f) => moeda.format(f.valor), alinhamento: 'direita' },
            {
              cabecalho: 'Valor Unit.',
              conteudo: (f) => moeda.format(f.valorUnitarioMaterial),
              alinhamento: 'direita'
            },
            { cabecalho: 'NF Devolução', conteudo: (f) => f.nfDevolucao },
            { cabecalho: 'Código VA', conteudo: (f) => f.codigoVa },
            { cabecalho: 'Código QM', conteudo: (f) => f.codigoQm },
            { cabecalho: 'Modo Defeito', conteudo: (f) => f.modoDefeito },
            { cabecalho: 'Área', conteudo: (f) => f.area },
            { cabecalho: 'Ordem Produção', conteudo: (f) => f.ordemProducao },
            { cabecalho: 'Entrada Devolução', conteudo: (f) => formatarData(f.dataEntradaDevolucao) },
            {
              cabecalho: 'Conclusão',
              conteudo: (f) => (f.dataConclusao ? formatarData(f.dataConclusao) : '—')
            },
            { cabecalho: 'Criação', conteudo: (f) => formatarData(f.dataCriacao) }
          ]}
          acoes={(ficha) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Editar ficha" aoClicar={() => abrirEdicao(ficha)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir ficha"
                tom="perigo"
                aoClicar={() => setExcluindo(ficha)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? `Editar Ficha ${edicao.id}` : 'Nova Ficha de QM'}
        icone={Files}
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
                  id="codPlantaQm"
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

                <Entrada
                  id="idFncQm"
                  rotulo="FNC"
                  value={edicao.idFnc}
                  disabled={salvando}
                  ajuda="Opcional"
                  onChange={(e) => mudar('idFnc', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-2"
                />

                <Selecao
                  id="statusQm"
                  rotulo="Status"
                  required
                  value={edicao.status}
                  disabled={salvando}
                  onChange={(e) => mudar('status', e.target.value)}
                  larguraClasse="sm:col-span-3"
                >
                  <option value="">Selecione...</option>
                  {STATUS_QM.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.descricao}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="nomeClienteQm"
                  rotulo="Cliente"
                  value={edicao.nomeCliente}
                  disabled={salvando}
                  onChange={(e) => mudar('nomeCliente', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
              </div>
            </Secao>

            <Secao titulo="Devolução">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="modoDefeito"
                  rotulo="Modo de Defeito"
                  value={edicao.modoDefeito}
                  disabled={salvando}
                  onChange={(e) => mudar('modoDefeito', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="area"
                  rotulo="Área"
                  value={edicao.area}
                  disabled={salvando}
                  onChange={(e) => mudar('area', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="ordemProducao"
                  rotulo="Ordem de Produção"
                  value={edicao.ordemProducao}
                  disabled={salvando}
                  onChange={(e) => mudar('ordemProducao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="dataEntradaDevolucao"
                  rotulo="Data Entrada Devolução"
                  type="date"
                  required
                  value={edicao.dataEntradaDevolucao}
                  disabled={salvando}
                  onChange={(e) => mudar('dataEntradaDevolucao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="dataConclusao"
                  rotulo="Data de Conclusão"
                  type="date"
                  required
                  value={edicao.dataConclusao}
                  disabled={salvando}
                  onChange={(e) => mudar('dataConclusao', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
                <Entrada
                  id="nfDevolucao"
                  rotulo="NF Devolução"
                  required
                  value={edicao.nfDevolucao}
                  disabled={salvando}
                  onChange={(e) => mudar('nfDevolucao', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-4"
                />

                <Entrada
                  id="codigoVA"
                  rotulo="Código VA"
                  required
                  value={edicao.codigoVa}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoVa', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-6"
                />
                <Entrada
                  id="codigoQM"
                  rotulo="Código QM"
                  required
                  value={edicao.codigoQm}
                  disabled={salvando}
                  onChange={(e) => mudar('codigoQm', e.target.value.toUpperCase())}
                  larguraClasse="sm:col-span-6"
                />
              </div>
            </Secao>

            <Secao titulo="Quantidades e Valores">
              <div className="grid gap-4 sm:grid-cols-12">
                <Entrada
                  id="quantidadePaletes"
                  rotulo="Quantidade Paletes"
                  type="number"
                  min="0"
                  required
                  value={edicao.quantidadePaletes}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadePaletes', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="quantidadeRecebida"
                  rotulo="Quantidade Recebida"
                  type="number"
                  min="0"
                  required
                  value={edicao.quantidadeRecebida}
                  disabled={salvando}
                  onChange={(e) => mudar('quantidadeRecebida', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="quantidadeRefugada"
                  rotulo="Quantidade Refugada"
                  type="number"
                  min="0"
                  required
                  value={edicao.quantidadeRefugada}
                  disabled={salvando}
                  ajuda="Não pode exceder a recebida"
                  onChange={(e) => mudar('quantidadeRefugada', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorUnitarioMaterial"
                  rotulo="Valor Unit. do Material"
                  type="number"
                  step="0.01"
                  value={edicao.valorUnitarioMaterial}
                  disabled={salvando}
                  onChange={(e) => mudar('valorUnitarioMaterial', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="valorQm"
                  rotulo="Valor"
                  type="number"
                  step="0.01"
                  required
                  value={edicao.valor}
                  disabled={salvando}
                  onChange={(e) => mudar('valor', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
              </div>
            </Secao>
          </>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir ficha"
        mensagem={excluindo ? `Deseja realmente excluir a ficha ${excluindo.id}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
