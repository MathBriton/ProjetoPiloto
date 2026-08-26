import { useCallback, useState } from 'react'
import { CheckCircle2, AlertTriangle, XCircle, Info, X } from 'lucide-react'

type TipoToast = 'success' | 'warning' | 'danger' | 'info'

interface Toast {
  id: number
  tipo: TipoToast
  mensagem: string
}

const ICONE_POR_TIPO: Record<TipoToast, typeof CheckCircle2> = {
  success: CheckCircle2,
  warning: AlertTriangle,
  danger: XCircle,
  info: Info
}

const COR_POR_TIPO: Record<TipoToast, string> = {
  success: 'text-emerald-600 dark:text-emerald-400',
  warning: 'text-amber-600 dark:text-amber-400',
  danger: 'text-red-600 dark:text-red-400',
  info: 'text-slate-500 dark:text-slate-400'
}

let proximoId = 1

export function useToast() {
  const [toasts, setToasts] = useState<Toast[]>([])

  const mostrarToast = useCallback((mensagem: string, tipo: TipoToast = 'info') => {
    const id = proximoId++
    setToasts((atual) => [...atual, { id, tipo, mensagem }])
    setTimeout(() => {
      setToasts((atual) => atual.filter((t) => t.id !== id))
    }, 4500)
  }, [])

  const fecharToast = useCallback((id: number) => {
    setToasts((atual) => atual.filter((t) => t.id !== id))
  }, [])

  return { toasts, mostrarToast, fecharToast }
}

export function ToastStack({ toasts, aoFechar }: { toasts: Toast[]; aoFechar: (id: number) => void }) {
  if (toasts.length === 0) return null

  return (
    <div className="fixed bottom-4 right-4 z-50 flex flex-col gap-2">
      {toasts.map((toast) => {
        const Icone = ICONE_POR_TIPO[toast.tipo]
        return (
          <div
            key={toast.id}
            role="status"
            className="flex w-80 items-start gap-2 rounded-lg border border-slate-200 bg-white p-3 shadow-lg
              dark:border-slate-800 dark:bg-slate-900"
          >
            <Icone size={18} className={`mt-0.5 shrink-0 ${COR_POR_TIPO[toast.tipo]}`} />
            <p className="flex-1 text-sm text-slate-700 dark:text-slate-200">{toast.mensagem}</p>
            <button
              onClick={() => aoFechar(toast.id)}
              aria-label="Fechar notificação"
              className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
            >
              <X size={15} />
            </button>
          </div>
        )
      })}
    </div>
  )
}
