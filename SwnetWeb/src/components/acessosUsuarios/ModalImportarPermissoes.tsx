import { useEffect, useState } from 'react'
import { Copy, X, User } from 'lucide-react'
import { buscarUsuarios } from '@/services/acessosUsuariosApi'
import type { UsuarioBusca } from '@/pages/administrador/tipos'

interface ModalImportarPermissoesProps {
  aberto: boolean
  usuarioDestino: UsuarioBusca | null
  aoFechar: () => void
  aoConfirmar: (usuarioOrigem: UsuarioBusca) => void
}

export function ModalImportarPermissoes({ aberto, usuarioDestino, aoFechar, aoConfirmar }: ModalImportarPermissoesProps) {
  const [incluirInativos, setIncluirInativos] = useState(false)
  const [usuarios, setUsuarios] = useState<UsuarioBusca[]>([])
  const [origemSelecionada, setOrigemSelecionada] = useState<string>('')

  useEffect(() => {
    if (!aberto) return
    setIncluirInativos(false)
    setOrigemSelecionada('')
    buscarUsuarios('', false).then(setUsuarios)
  }, [aberto])

  useEffect(() => {
    if (!aberto) return
    buscarUsuarios('', incluirInativos).then(setUsuarios)
  }, [incluirInativos, aberto])

  if (!aberto) return null

  const opcoes = usuarios.filter(
    (u) => !usuarioDestino || `${u.matricula}-${u.filial}` !== `${usuarioDestino.matricula}-${usuarioDestino.filial}`
  )

  function confirmar() {
    const usuario = opcoes.find((u) => `${u.matricula}|${u.filial}` === origemSelecionada)
    if (!usuario) return
    aoConfirmar(usuario)
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 px-4" role="dialog" aria-modal="true">
      <div className="w-full max-w-lg rounded-xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <h2 className="flex items-center gap-2 text-sm font-semibold text-slate-800 dark:text-white">
            <Copy size={16} /> Importar Permissões
          </h2>
          <button onClick={aoFechar} aria-label="Fechar" className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200">
            <X size={17} />
          </button>
        </div>

        <div className="space-y-4 px-4 py-4">
          <div>
            <div className="mb-2 flex items-center justify-between">
              <label htmlFor="usuarioOrigem" className="flex items-center gap-1 text-sm font-medium text-slate-600 dark:text-slate-300">
                <User size={14} /> Usuário de origem
              </label>
              <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
                <input
                  type="checkbox"
                  checked={incluirInativos}
                  onChange={(e) => setIncluirInativos(e.target.checked)}
                  className="h-3.5 w-3.5 accent-swnet-600"
                />
                Incluir inativos
              </label>
            </div>

            <select
              id="usuarioOrigem"
              value={origemSelecionada}
              onChange={(e) => setOrigemSelecionada(e.target.value)}
              className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2 px-3 text-sm
                focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                dark:border-slate-700 dark:bg-slate-800 dark:text-white"
            >
              <option value="">Selecione um usuário...</option>
              {opcoes.map((u) => (
                <option key={`${u.matricula}-${u.filial}`} value={`${u.matricula}|${u.filial}`}>
                  {u.ativo ? '' : '⚠ '}
                  {u.nome} — {u.login}
                  {u.ativo ? '' : ' (inativo)'}
                </option>
              ))}
            </select>
          </div>

          <div className="rounded-lg bg-amber-50 px-3 py-3 text-center text-sm text-amber-700 dark:bg-amber-950/30 dark:text-amber-400">
            <strong>Atenção:</strong> todas as permissões atuais serão substituídas pelas do usuário selecionado.
          </div>
        </div>

        <div className="flex justify-end gap-2 border-t border-slate-200 px-4 py-3 dark:border-slate-800">
          <button
            onClick={aoFechar}
            className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-600
              hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            Cancelar
          </button>
          <button
            onClick={confirmar}
            disabled={!origemSelecionada}
            className="rounded-lg bg-swnet-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-swnet-700 disabled:cursor-not-allowed disabled:opacity-50"
          >
            Confirmar
          </button>
        </div>
      </div>
    </div>
  )
}
