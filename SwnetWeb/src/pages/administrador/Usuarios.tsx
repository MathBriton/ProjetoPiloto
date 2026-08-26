import { useEffect, useState } from 'react'
import { Search, UserPlus, X } from 'lucide-react'
import { alternarAtivo, criarUsuario, listarPerfis, listarPlantas, listarUsuarios } from '@/services/usuariosApi'
import { FormularioUsuario } from '@/components/usuarios/FormularioUsuario'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { useToast, ToastStack } from '@/components/Toast'
import type { NovoUsuario, UsuarioCadastro } from './tipos'

export function Usuarios() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const perfis = listarPerfis()
  const plantas = listarPlantas()

  const [usuarios, setUsuarios] = useState<UsuarioCadastro[]>([])
  const [termo, setTermo] = useState('')
  const [carregando, setCarregando] = useState(true)
  const [formularioAberto, setFormularioAberto] = useState(false)
  const [alvoStatus, setAlvoStatus] = useState<UsuarioCadastro | null>(null)

  useEffect(() => {
    const timeout = setTimeout(() => {
      setCarregando(true)
      listarUsuarios(termo)
        .then(setUsuarios)
        .finally(() => setCarregando(false))
    }, 250)
    return () => clearTimeout(timeout)
  }, [termo])

  function nomePerfil(id: number) {
    return perfis.find((p) => p.id === id)?.nome ?? '—'
  }

  function nomesPlantas(ids: number[]) {
    return ids.map((id) => plantas.find((p) => p.id === id)?.nome ?? id).join(', ') || '—'
  }

  async function salvarNovo(dados: NovoUsuario) {
    const criado = await criarUsuario(dados)
    setUsuarios((atual) => [...atual, criado])
    setFormularioAberto(false)
    mostrarToast(`Usuário ${criado.login} criado com sucesso.`, 'success')
  }

  async function confirmarStatus() {
    if (!alvoStatus) return
    const atualizado = await alternarAtivo(alvoStatus.id)
    if (atualizado) {
      setUsuarios((atual) => atual.map((u) => (u.id === atualizado.id ? atualizado : u)))
      mostrarToast(`Usuário ${atualizado.login} ${atualizado.ativo ? 'ativado' : 'inativado'}.`, 'success')
    }
    setAlvoStatus(null)
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3 px-1">
        <h1 className="text-xl font-bold text-slate-800 dark:text-white">Usuários</h1>

        <div className="flex items-center gap-2">
          <div className="relative">
            <Search size={14} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={termo}
              onChange={(e) => setTermo(e.target.value)}
              placeholder="Pesquisar usuário"
              className="w-56 rounded-lg border border-slate-200 bg-slate-50 py-1.5 pl-8 pr-7 text-xs
                focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                dark:border-slate-700 dark:bg-slate-800 dark:text-white"
            />
            {termo && (
              <button
                onClick={() => setTermo('')}
                aria-label="Limpar pesquisa"
                className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
              >
                <X size={13} />
              </button>
            )}
          </div>

          <button
            onClick={() => setFormularioAberto(true)}
            className="flex items-center gap-1.5 rounded-lg bg-swnet-600 px-3 py-2 text-sm font-medium text-white hover:bg-swnet-700"
          >
            <UserPlus size={16} /> Novo usuário
          </button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-900">
        <table className="w-full text-left text-sm">
          <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
            <tr>
              <th className="px-4 py-3 font-semibold">Login</th>
              <th className="px-4 py-3 font-semibold">Nome</th>
              <th className="px-4 py-3 font-semibold">E-mail</th>
              <th className="px-4 py-3 font-semibold">Perfil</th>
              <th className="px-4 py-3 font-semibold">Plantas</th>
              <th className="px-4 py-3 font-semibold">Status</th>
            </tr>
          </thead>
          <tbody>
            {carregando && (
              <tr>
                <td colSpan={6} className="px-4 py-10 text-center text-sm text-slate-400">
                  Carregando usuários...
                </td>
              </tr>
            )}

            {!carregando && usuarios.length === 0 && (
              <tr>
                <td colSpan={6} className="px-4 py-10 text-center text-sm text-slate-400">
                  Nenhum usuário encontrado.
                </td>
              </tr>
            )}

            {!carregando &&
              usuarios.map((usuario) => (
                <tr key={usuario.id} className="border-b border-slate-100 last:border-0 dark:border-slate-800">
                  <td className="px-4 py-3 font-medium text-slate-700 dark:text-slate-200">{usuario.login}</td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">{usuario.nomeCompleto}</td>
                  <td className="px-4 py-3 text-slate-500 dark:text-slate-400">{usuario.email}</td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">{nomePerfil(usuario.perfilId)}</td>
                  <td className="px-4 py-3 text-xs text-slate-500 dark:text-slate-400">{nomesPlantas(usuario.plantaIds)}</td>
                  <td className="px-4 py-3">
                    <button
                      onClick={() => setAlvoStatus(usuario)}
                      className={`rounded-full px-2.5 py-1 text-xs font-medium transition-colors ${
                        usuario.ativo
                          ? 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100 dark:bg-emerald-500/10 dark:text-emerald-400'
                          : 'bg-slate-100 text-slate-500 hover:bg-slate-200 dark:bg-slate-800 dark:text-slate-400'
                      }`}
                    >
                      {usuario.ativo ? 'Ativo' : 'Inativo'}
                    </button>
                  </td>
                </tr>
              ))}
          </tbody>
        </table>
      </div>

      <FormularioUsuario
        aberto={formularioAberto}
        aoFechar={() => setFormularioAberto(false)}
        aoSalvar={salvarNovo}
      />

      <ConfirmDialog
        aberto={alvoStatus !== null}
        titulo={alvoStatus?.ativo ? 'Inativar usuário' : 'Ativar usuário'}
        mensagem={
          alvoStatus
            ? `Deseja ${alvoStatus.ativo ? 'inativar' : 'ativar'} o usuário ${alvoStatus.nomeCompleto}?`
            : ''
        }
        aoConfirmar={confirmarStatus}
        aoCancelar={() => setAlvoStatus(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
