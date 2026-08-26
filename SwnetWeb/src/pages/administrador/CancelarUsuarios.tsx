import { useState } from 'react'
import { Ban, CheckCircle2, Search, UserSearch } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, EstadoVazio, LinhaCarregando, TituloPagina } from '@/components/comuns/Comuns'
import { Entrada } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  alterarContas,
  listarContasPorEmail,
  pesquisarEmails
} from '@/services/administrador/cancelamentoUsuariosApi'
import { formatarData } from '@/services/comuns/baseMock'
import type { AcaoCancelamento, ContaSistema } from './tipos'

/**
 * Refatoração de Cancelar-Usuarios/CancelamentoUsuario.asp (Desligamento de
 * Usuário). O fluxo do legado tem três etapas:
 *
 *   1. pesquisar por nome/e-mail em todos os bancos
 *   2. escolher quais e-mails únicos visualizar
 *   3. marcar contas no grid e habilitar/desabilitar em lote
 */

function chaveConta(conta: ContaSistema) {
  return `${conta.sistemaId}:${conta.id}`
}

export function CancelarUsuarios() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [nome, setNome] = useState('')
  const [email, setEmail] = useState('')
  const [pesquisando, setPesquisando] = useState(false)
  const [pesquisaFeita, setPesquisaFeita] = useState(false)

  const [emailsEncontrados, setEmailsEncontrados] = useState<string[]>([])
  const [emailsSelecionados, setEmailsSelecionados] = useState<Set<string>>(new Set())

  const [contas, setContas] = useState<ContaSistema[]>([])
  const [carregandoContas, setCarregandoContas] = useState(false)
  const [selecionadas, setSelecionadas] = useState<Set<string>>(new Set())

  const [confirmando, setConfirmando] = useState<AcaoCancelamento | null>(null)
  const [processando, setProcessando] = useState(false)

  async function pesquisar() {
    if (!nome.trim() && !email.trim()) {
      mostrarToast('Informe o nome ou o e-mail do usuário para pesquisar.', 'warning')
      return
    }

    setPesquisando(true)
    setContas([])
    setSelecionadas(new Set())
    try {
      const encontrados = await pesquisarEmails(nome, email)
      setEmailsEncontrados(encontrados)
      setEmailsSelecionados(new Set(encontrados))
      setPesquisaFeita(true)
    } finally {
      setPesquisando(false)
    }
  }

  function alternarEmail(alvo: string) {
    setEmailsSelecionados((atual) => {
      const novo = new Set(atual)
      if (novo.has(alvo)) novo.delete(alvo)
      else novo.add(alvo)
      return novo
    })
  }

  async function visualizarContas() {
    if (emailsSelecionados.size === 0) {
      mostrarToast('Selecione ao menos um e-mail.', 'warning')
      return
    }

    setCarregandoContas(true)
    setSelecionadas(new Set())
    try {
      setContas(await listarContasPorEmail([...emailsSelecionados]))
    } finally {
      setCarregandoContas(false)
    }
  }

  const contasSelecionaveis = contas.filter((c) => !c.semRegistro)
  const todasMarcadas =
    contasSelecionaveis.length > 0 && contasSelecionaveis.every((c) => selecionadas.has(chaveConta(c)))

  function alternarTodas() {
    setSelecionadas(todasMarcadas ? new Set() : new Set(contasSelecionaveis.map(chaveConta)))
  }

  function alternarConta(conta: ContaSistema) {
    setSelecionadas((atual) => {
      const novo = new Set(atual)
      const chave = chaveConta(conta)
      if (novo.has(chave)) novo.delete(chave)
      else novo.add(chave)
      return novo
    })
  }

  async function executar() {
    const acao = confirmando
    if (!acao) return

    setConfirmando(null)
    setProcessando(true)
    try {
      const alvos = contasSelecionaveis
        .filter((c) => selecionadas.has(chaveConta(c)))
        .map((c) => ({ id: c.id, sistemaId: c.sistemaId }))

      const resultado = await alterarContas(acao, alvos)

      if (resultado.alteradas.length === 0) {
        mostrarToast(
          'Os usuários selecionados já estão no estado desejado ou não foram localizados nos sistemas.',
          'warning'
        )
      } else {
        mostrarToast(
          `${resultado.alteradas.length} conta(s) ${acao === 'habilitar' ? 'habilitada(s)' : 'desabilitada(s)'} com sucesso.`,
          'success'
        )
      }

      // Recarrega o grid para refletir as datas de bloqueio
      setContas(await listarContasPorEmail([...emailsSelecionados]))
      setSelecionadas(new Set())
    } finally {
      setProcessando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Cancelar Usuários" />

      <Painel titulo="Desligamento de Usuário" icone={Ban}>
        <form
          className="grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            pesquisar()
          }}
        >
          <Entrada
            id="nomeUsuario"
            rotulo="Nome do usuário"
            placeholder="Nome do usuário..."
            value={nome}
            disabled={pesquisando}
            onChange={(e) => setNome(e.target.value)}
            larguraClasse="sm:col-span-5"
          />
          <Entrada
            id="emailUsuario"
            rotulo="E-mail do usuário"
            placeholder="E-mail do usuário..."
            value={email}
            disabled={pesquisando}
            onChange={(e) => setEmail(e.target.value)}
            larguraClasse="sm:col-span-5"
          />
          <div className="sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={pesquisando}>
              {pesquisando ? 'Buscando...' : 'Pesquisar'}
            </Botao>
          </div>
        </form>
      </Painel>

      {pesquisaFeita && emailsEncontrados.length === 0 && !pesquisando && (
        <Painel titulo="Resultado">
          <EstadoVazio
            icone={UserSearch}
            titulo="Nenhum usuário encontrado para os critérios informados"
            descricao="Tente outro nome ou parte do e-mail."
          />
        </Painel>
      )}

      {emailsEncontrados.length > 0 && (
        <Painel
          titulo={`${emailsEncontrados.length} usuário(s) único(s) encontrado(s)`}
          acoes={
            <Botao aoClicar={visualizarContas} desabilitado={carregandoContas}>
              {carregandoContas ? 'Carregando...' : 'Visualizar selecionados'}
            </Botao>
          }
        >
          <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
            {emailsEncontrados.map((item) => (
              <label
                key={item}
                className="flex cursor-pointer items-center gap-2 rounded-lg border border-slate-200 px-3 py-2
                  text-sm text-slate-600 hover:border-swnet-300 dark:border-slate-700 dark:text-slate-300"
              >
                <input
                  type="checkbox"
                  checked={emailsSelecionados.has(item)}
                  onChange={() => alternarEmail(item)}
                  className="h-4 w-4 accent-swnet-600"
                />
                <span className="truncate" title={item}>
                  {item}
                </span>
              </label>
            ))}
          </div>
        </Painel>
      )}

      {(contas.length > 0 || carregandoContas) && (
        <Painel
          titulo="Contas por sistema"
          acoes={
            <>
              <span className="text-xs text-slate-400">{selecionadas.size} selecionada(s)</span>
              <Botao
                variante="sucesso"
                icone={CheckCircle2}
                aoClicar={() => setConfirmando('habilitar')}
                desabilitado={selecionadas.size === 0 || processando}
              >
                Habilitar
              </Botao>
              <Botao
                variante="perigo"
                icone={Ban}
                aoClicar={() => setConfirmando('desabilitar')}
                desabilitado={selecionadas.size === 0 || processando}
              >
                Desabilitar
              </Botao>
            </>
          }
        >
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
                <tr>
                  <th className="px-3 py-2">
                    <input
                      type="checkbox"
                      checked={todasMarcadas}
                      onChange={alternarTodas}
                      aria-label="Selecionar todas as contas"
                      className="h-4 w-4 accent-swnet-600"
                    />
                  </th>
                  <th className="px-3 py-2 font-semibold">#</th>
                  <th className="px-3 py-2 font-semibold">Usuário</th>
                  <th className="px-3 py-2 font-semibold">E-mail</th>
                  <th className="px-3 py-2 font-semibold">Sistema</th>
                  <th className="px-3 py-2 font-semibold">Cadastro</th>
                  <th className="px-3 py-2 font-semibold">Bloqueio</th>
                  <th className="px-3 py-2 font-semibold">Status</th>
                </tr>
              </thead>
              <tbody>
                {carregandoContas && <LinhaCarregando colunas={8} mensagem="Consultando sistemas..." />}

                {!carregandoContas &&
                  contas.map((conta) => (
                    <tr
                      key={conta.semRegistro ? `vazio:${conta.sistemaId}` : chaveConta(conta)}
                      className="border-b border-slate-100 last:border-0 dark:border-slate-800"
                    >
                      <td className="px-3 py-2">
                        {!conta.semRegistro && (
                          <input
                            type="checkbox"
                            checked={selecionadas.has(chaveConta(conta))}
                            onChange={() => alternarConta(conta)}
                            aria-label={`Selecionar ${conta.nome} em ${conta.sistema}`}
                            className="h-4 w-4 accent-swnet-600"
                          />
                        )}
                      </td>
                      <td className="px-3 py-2 text-slate-400">{conta.semRegistro ? '—' : conta.id}</td>
                      <td className="px-3 py-2 text-slate-700 dark:text-slate-200">{conta.nome || '—'}</td>
                      <td className="px-3 py-2 text-slate-500 dark:text-slate-400">{conta.email || '—'}</td>
                      <td className="px-3 py-2 text-slate-600 dark:text-slate-300">{conta.sistema}</td>
                      <td className="px-3 py-2 text-xs text-slate-500">{formatarData(conta.dataCadastro)}</td>
                      <td className="px-3 py-2 text-xs text-slate-500">{formatarData(conta.dataBloqueio)}</td>
                      <td className="px-3 py-2">
                        {conta.semRegistro ? (
                          <Etiqueta texto="SEM REGISTRO" tom="neutro" />
                        ) : (
                          <Etiqueta
                            texto={conta.ativo ? 'ATIVO' : 'INATIVO'}
                            tom={conta.ativo ? 'sucesso' : 'perigo'}
                          />
                        )}
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </Painel>
      )}

      <ConfirmDialog
        aberto={confirmando !== null}
        titulo={confirmando === 'habilitar' ? 'Habilitar contas' : 'Desabilitar contas'}
        mensagem={
          confirmando === 'habilitar'
            ? `${selecionadas.size} conta(s) selecionada(s) voltarão a ter acesso aos respectivos sistemas. Deseja continuar?`
            : `${selecionadas.size} conta(s) selecionada(s) perderão o acesso aos respectivos sistemas. Deseja continuar?`
        }
        rotuloConfirmar={confirmando === 'habilitar' ? 'Habilitar' : 'Desabilitar'}
        aoConfirmar={executar}
        aoCancelar={() => setConfirmando(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
