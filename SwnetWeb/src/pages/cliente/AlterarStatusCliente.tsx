import { useCallback, useEffect, useState } from 'react'
import { Eraser, FileSearch, Power, Search, Send } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  FILTROS_UNIFICADO_VAZIOS,
  STATUS_UNIFICADO,
  alterarStatusUnificado,
  descricaoStatusUnificado,
  listarUnificados,
  type CadastroUnificado,
  type FiltrosUnificado
} from '@/services/cliente/clienteApi'
import { PLANTAS_PROTHEUS, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Cliente/Alterar-Status-Do-Cliente.
 *
 * No legado a tela reaproveita o index.asp do Cadastro Unificado e faz o
 * `UPDATE UnificadoCliente SET StatusCli`. É a fila por departamento: cada
 * área filtra pelo seu status e move o cadastro adiante.
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

export function AlterarStatusCliente() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [filtros, setFiltros] = useState<FiltrosUnificado>(FILTROS_UNIFICADO_VAZIOS)
  const [filtrosAplicados, setFiltrosAplicados] = useState(FILTROS_UNIFICADO_VAZIOS)
  const [cadastros, setCadastros] = useState<CadastroUnificado[]>([])
  const [carregando, setCarregando] = useState(true)

  const [alterando, setAlterando] = useState<CadastroUnificado | null>(null)
  const [novoStatus, setNovoStatus] = useState('')
  const [salvando, setSalvando] = useState(false)
  const [detalhe, setDetalhe] = useState<CadastroUnificado | null>(null)

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

  async function confirmar() {
    if (!alterando) return

    setSalvando(true)
    try {
      const resultado = await alterarStatusUnificado(
        alterando.codigo,
        novoStatus,
        usuario?.login ?? ''
      )
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setAlterando(null)
        setNovoStatus('')
        await carregar()
      }
    } finally {
      setSalvando(false)
    }
  }

  /** Contadores por status, para a fila ficar visível de cara */
  const contadores = STATUS_UNIFICADO.map((status) => ({
    ...status,
    total: cadastros.filter((c) => c.status === status.codigo).length
  })).filter((item) => item.total > 0)

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Alterar Status do Cliente" />

      {contadores.length > 0 && (
        <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
          {contadores.map((item) => (
            <button
              key={item.codigo}
              type="button"
              onClick={() => {
                const proximo = { ...filtros, status: item.codigo }
                setFiltros(proximo)
                setFiltrosAplicados(proximo)
              }}
              className={`rounded-xl border bg-white px-3 py-2 text-left shadow-sm transition-colors
                dark:bg-slate-900 ${
                  filtrosAplicados.status === item.codigo
                    ? 'border-swnet-400 dark:border-swnet-600'
                    : 'border-slate-200 hover:border-swnet-300 dark:border-slate-800'
                }`}
            >
              <span className="block truncate text-xs font-medium text-slate-500 dark:text-slate-400">
                {item.descricao}
              </span>
              <span className="mt-0.5 block text-xl font-bold text-slate-800 dark:text-white">
                {item.total}
              </span>
            </button>
          ))}
        </div>
      )}

      <Painel titulo="Filtros">
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setFiltrosAplicados(filtros)
          }}
        >
          <Entrada
            id="filtroClienteStatus"
            rotulo="Cliente"
            value={filtros.cliente}
            onChange={(e) => setFiltros({ ...filtros, cliente: e.target.value })}
            larguraClasse="sm:col-span-3"
          />
          <Entrada
            id="filtroCodigoStatus"
            rotulo="Código"
            value={filtros.codigo}
            onChange={(e) => setFiltros({ ...filtros, codigo: e.target.value })}
            larguraClasse="sm:col-span-2"
          />
          <Selecao
            id="filtroStatusAlt"
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
            id="filtroPlantaStatus"
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

      <Painel titulo={`${cadastros.length} cadastro(s)`} icone={Power}>
        <Tabela
          itens={cadastros}
          chave={(item) => item.codigo}
          carregando={carregando}
          mensagemVazia="Nenhum cadastro encontrado"
          colunas={[
            { cabecalho: '#', conteudo: (c) => c.codigo, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (c) => nomePlanta(c.planta) },
            { cabecalho: 'Data Inclusão', conteudo: (c) => formatarData(c.dataInclusao) },
            { cabecalho: 'Cliente', conteudo: (c) => c.razaoSocial },
            { cabecalho: 'CPF/CNPJ', conteudo: (c) => c.cnpj || c.cpf || '—' },
            {
              cabecalho: 'Status atual',
              conteudo: (c) => (
                <Etiqueta texto={descricaoStatusUnificado(c.status)} tom={TOM_STATUS[c.status] ?? 'neutro'} />
              )
            },
            { cabecalho: 'Emitente', conteudo: (c) => c.emitente }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao icone={FileSearch} titulo="Ver histórico" aoClicar={() => setDetalhe(item)} />
              <BotaoAcao
                icone={Send}
                titulo="Alterar status"
                aoClicar={() => {
                  setNovoStatus('')
                  setAlterando(item)
                }}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={alterando !== null}
        titulo="Alterar status do cadastro"
        icone={Power}
        aoFechar={() => setAlterando(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setAlterando(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" aoClicar={confirmar} desabilitado={salvando || !novoStatus}>
              {salvando ? 'Gravando...' : 'Alterar'}
            </Botao>
          </>
        }
      >
        {alterando && (
          <div className="grid gap-4">
            <p className="text-sm text-slate-600 dark:text-slate-300">
              Cadastro {alterando.codigo} — {alterando.razaoSocial}
              <span className="block text-xs text-slate-400">
                Status atual: {descricaoStatusUnificado(alterando.status)}
              </span>
            </p>

            <Selecao
              id="novoStatusCliente"
              rotulo="Novo status"
              required
              value={novoStatus}
              disabled={salvando}
              onChange={(e) => setNovoStatus(e.target.value)}
            >
              <option value="">Selecione...</option>
              {STATUS_UNIFICADO.filter((item) => item.codigo !== alterando.status).map((item) => (
                <option key={item.codigo} value={item.codigo}>
                  {item.descricao}
                </option>
              ))}
            </Selecao>
          </div>
        )}
      </Modal>

      <Modal
        aberto={detalhe !== null}
        titulo={detalhe ? `Histórico do cadastro ${detalhe.codigo}` : ''}
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
          <div className="rounded-lg border border-slate-200 dark:border-slate-700">
            <Tabela
              itens={detalhe.vistos}
              chave={(v) => v.data + v.acao}
              mensagemVazia="Nenhuma movimentação registrada"
              colunas={[
                { cabecalho: 'Data', conteudo: (v) => formatarData(v.data) },
                { cabecalho: 'Usuário', conteudo: (v) => v.usuario },
                { cabecalho: 'Ação', conteudo: (v) => v.acao },
                { cabecalho: 'Observação', conteudo: (v) => v.observacao || '—' }
              ]}
            />
          </div>
        )}
      </Modal>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
