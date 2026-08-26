import { AlertTriangle } from 'lucide-react'

interface ConfirmDialogProps {
  aberto: boolean
  titulo: string
  mensagem: string
  aoConfirmar: () => void
  aoCancelar: () => void
  rotuloConfirmar?: string
}

export function ConfirmDialog({
  aberto,
  titulo,
  mensagem,
  aoConfirmar,
  aoCancelar,
  rotuloConfirmar = 'Confirmar'
}: ConfirmDialogProps) {
  if (!aberto) return null

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 px-4" role="dialog" aria-modal="true">
      <div className="w-full max-w-sm rounded-xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <AlertTriangle size={17} className="text-amber-500" />
          <h2 className="text-sm font-semibold text-slate-800 dark:text-white">{titulo}</h2>
        </div>
        <div className="px-4 py-4">
          <p className="text-sm text-slate-600 dark:text-slate-300">{mensagem}</p>
        </div>
        <div className="flex justify-end gap-2 border-t border-slate-200 px-4 py-3 dark:border-slate-800">
          <button
            onClick={aoCancelar}
            className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-600
              hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            Cancelar
          </button>
          <button
            onClick={aoConfirmar}
            className="rounded-lg bg-swnet-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-swnet-700"
          >
            {rotuloConfirmar}
          </button>
        </div>
      </div>
    </div>
  )
}
