import { useState } from 'react'
import { ArrowUpDown, Ban, FileDown, History, Search } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, EstadoVazio, LinhaCarregando, TituloPagina } from '@/components/comuns/Comuns'
import { Selecao } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  desabilitarInativos,
  listarUsuariosInativos
} from '@/services/administrador/cancelamentoUsuariosApi'
import { SISTEMAS, formatarData } from '@/services/comuns/baseMock'
import type { UsuarioInativo } from './tipos'

/**
 * Refatoração de Revisao-De-Usuarios/RevisaoUsuarios.asp.
 *
 * Regras herdadas do legado:
 * - só é possível marcar contas com 60 dias ou mais sem acesso;
 * - a desabilitação em lote aceita um sistema por vez.
 */

const DIAS_MINIMOS_PARA_DESABILITAR = 60

const OPCOES_DIAS = [
  { valor: 0, rotulo: 'Sem filtro' },
  { valor: 60, rotulo: '60 dias ou mais' },
  { valor: 90, rotulo: '90 dias ou mais' },
  { valor: 180, rotulo: '180 dias ou mais' }
]

function chave(usuario: UsuarioInativo) {
  return `${usuario.sistemaId}:${usuario.id}`
}

export function RevisaoUsuarios() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [dias, setDias] = useState(0)
  const [sistemaId, setSistemaId] = useState('TODOS')
  const [usuarios, setUsuarios] = useState<UsuarioInativo[]>([])
  const [carregando, setCarregando] = useState(false)
  const [pesquisaFeita, setPesquisaFeita] = useState(false)
  const [ordemDecrescente, setOrdemDecrescente] = useState(true)
  const [selecionados, setSelecionados] = useState<Set<string>>(new Set())
  const [confirmando, setConfirmando] = useState(false)
  const [processando, setProcessando] = useState(false)

  async function pesquisar() {
    setCarregando(true)
    setSelecionados(new Set())
    try {
      setUsuarios(await listarUsuariosInativos(dias, sistemaId))
      setPesquisaFeita(true)
    } finally {
      setCarregando(false)
    }
  }

  /** Sistema já travado pela primeira seleção — o legado só grava um por vez */
  const sistemaSelecionado = [...selecionados][0]?.split(':')[0] ?? null

  function podeDesabilitar(usuario: UsuarioInativo) {
    return usuario.diasSemAcessar >= DIAS_MINIMOS_PARA_DESABILITAR
  }

  function alternar(usuario: UsuarioInativo) {
    const chaveUsuario = chave(usuario)

    if (!selecionados.has(chaveUsuario) && sistemaSelecionado && sistemaSelecionado !== usuario.sistemaId) {
      mostrarToast('Selecione usuários de apenas um sistema por vez para desabilitar.', 'warning')
      return
    }

    setSelecionados((atual) => {
      const novo = new Set(atual)
      if (novo.has(chaveUsuario)) novo.delete(chaveUsuario)
      else novo.add(chaveUsuario)
      return novo
    })
  }

  function alternarTodos() {
    if (selecionados.size > 0) {
      setSelecionados(new Set())
      return
    }

    // Sem seleção prévia, marca só o primeiro sistema elegível da lista
    const elegiveis = ordenados.filter(podeDesabilitar)
    const primeiroSistema = elegiveis[0]?.sistemaId
    if (!primeiroSistema) return

    setSelecionados(new Set(elegiveis.filter((u) => u.sistemaId === primeiroSistema).map(chave)))
  }

  async function executar() {
    setConfirmando(false)
    setProcessando(true)
    try {
      const alvos = usuarios
        .filter((u) => selecionados.has(chave(u)))
        .map((u) => ({ id: u.id, sistemaId: u.sistemaId }))

      const resultado = await desabilitarInativos(alvos)
      mostrarToast(`${resultado.alteradas.length} usuário(s) desabilitado(s) com sucesso!`, 'success')

      setUsuarios(await listarUsuariosInativos(dias, sistemaId))
      setSelecionados(new Set())
    } finally {
      setProcessando(false)
    }
  }

  function exportarPdf() {
    // TODO: o legado usa jsPDF; ainda não há lib de PDF no frontend
    mostrarToast('Exportação em PDF será implementada em breve.', 'info')
  }

  const ordenados = [...usuarios].sort((a, b) =>
    ordemDecrescente ? b.diasSemAcessar - a.diasSemAcessar : a.diasSemAcessar - b.diasSemAcessar
  )

  const elegiveis = ordenados.filter(podeDesabilitar).length

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Revisão de Usuários" />

      <Painel titulo="Filtros" icone={History}>
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            pesquisar()
          }}
        >
          <Selecao
            id="diasInativos"
            rotulo="Dias sem acessar"
            value={dias}
            disabled={carregando}
            onChange={(e) => setDias(Number(e.target.value))}
            larguraClasse="sm:col-span-4"
          >
            {OPCOES_DIAS.map((opcao) => (
              <option key={opcao.valor} value={opcao.valor}>
                {opcao.rotulo}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="sistemaFiltro"
            rotulo="Sistema"
            value={sistemaId}
            disabled={carregando}
            onChange={(e) => setSistemaId(e.target.value)}
            larguraClasse="sm:col-span-6"
          >
            <option value="TODOS">Todos os Sistemas</option>
            {SISTEMAS.map((sistema) => (
              <option key={sistema.id} value={sistema.id}>
                {sistema.nome}
              </option>
            ))}
          </Selecao>

          <div className="sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando}>
              {carregando ? 'Buscando...' : 'Pesquisar'}
            </Botao>
          </div>
        </form>
      </Painel>

      {pesquisaFeita && (
        <Painel
          titulo={`${usuarios.length} usuário(s) encontrado(s)`}
          acoes={
            <>
              {sistemaId === 'TODOS' && (
                <span className="hidden text-xs text-slate-400 lg:inline">
                  Selecione usuários de apenas um sistema por vez.
                </span>
              )}
              <Botao variante="neutro" icone={FileDown} aoClicar={exportarPdf}>
                PDF
              </Botao>
              <Botao
                variante="perigo"
                icone={Ban}
                aoClicar={() => setConfirmando(true)}
                desabilitado={selecionados.size === 0 || processando}
              >
                Desabilitar ({selecionados.size})
              </Botao>
            </>
          }
        >
          {usuarios.length === 0 ? (
            <EstadoVazio
              icone={History}
              titulo="Nenhum usuário encontrado!"
              descricao="Não foram encontrados usuários com os filtros selecionados."
            />
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
                  <tr>
                    <th className="px-3 py-2">
                      <input
                        type="checkbox"
                        checked={selecionados.size > 0}
                        onChange={alternarTodos}
                        disabled={elegiveis === 0}
                        aria-label="Selecionar usuários elegíveis"
                        className="h-4 w-4 accent-swnet-600"
                      />
                    </th>
                    <th className="px-3 py-2 font-semibold">#</th>
                    <th className="px-3 py-2 font-semibold">Nome</th>
                    <th className="px-3 py-2 font-semibold">E-mail</th>
                    <th className="px-3 py-2 font-semibold">Sistema</th>
                    <th className="px-3 py-2 font-semibold">Último acesso</th>
                    <th className="px-3 py-2 font-semibold">
                      <button
                        type="button"
                        onClick={() => setOrdemDecrescente((v) => !v)}
                        title="Clique para ordenar"
                        className="inline-flex items-center gap-1 uppercase hover:text-slate-700 dark:hover:text-slate-200"
                      >
                        Dias sem acessar <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th className="px-3 py-2 font-semibold">Alerta</th>
                  </tr>
                </thead>
                <tbody>
                  {carregando && <LinhaCarregando colunas={8} />}

                  {!carregando &&
                    ordenados.map((usuario) => {
                      const elegivel = podeDesabilitar(usuario)
                      return (
                        <tr
                          key={chave(usuario)}
                          className="border-b border-slate-100 last:border-0 dark:border-slate-800"
                        >
                          <td className="px-3 py-2">
                            <input
                              type="checkbox"
                              checked={selecionados.has(chave(usuario))}
                              disabled={!elegivel}
                              onChange={() => alternar(usuario)}
                              aria-label={`Selecionar ${usuario.nome}`}
                              className="h-4 w-4 accent-swnet-600 disabled:opacity-40"
                            />
                          </td>
                          <td className="px-3 py-2 text-slate-400">{usuario.id}</td>
                          <td className="px-3 py-2 text-slate-700 dark:text-slate-200">{usuario.nome}</td>
                          <td className="px-3 py-2 text-slate-500 dark:text-slate-400">{usuario.email}</td>
                          <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{usuario.sistema}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">
                            {formatarData(usuario.ultimoAcesso)}
                          </td>
                          <td className="px-3 py-2 font-medium text-slate-700 dark:text-slate-200">
                            {usuario.diasSemAcessar}
                          </td>
                          <td className="px-3 py-2">
                            <Etiqueta
                              texto={elegivel ? 'Revisar' : 'Em uso'}
                              tom={elegivel ? 'perigo' : 'sucesso'}
                            />
                          </td>
                        </tr>
                      )
                    })}
                </tbody>
              </table>
            </div>
          )}
        </Painel>
      )}

      <ConfirmDialog
        aberto={confirmando}
        titulo="Desabilitar usuários"
        mensagem={`${selecionados.size} usuário(s) perderão o acesso ao sistema selecionado. Deseja continuar?`}
        rotuloConfirmar="Desabilitar"
        aoConfirmar={executar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
