import { useEffect, type ReactNode } from 'react'
import { X, type LucideIcon } from 'lucide-react'

/**
 * Modal genérico (os `.modal` do Bootstrap usados no Gerenciar Menus e nas
 * confirmações de ação). Para confirmações simples continue usando o
 * `ConfirmDialog` já existente.
 */

interface ModalProps {
  aberto: boolean
  titulo: string
  icone?: LucideIcon
  aoFechar: () => void
  rodape?: ReactNode
  largura?: 'sm' | 'md' | 'lg' | 'xl'
  children: ReactNode
}

const LARGURAS = {
  sm: 'max-w-sm',
  md: 'max-w-md',
  lg: 'max-w-2xl',
  xl: 'max-w-4xl'
}

export function Modal({ aberto, titulo, icone: Icone, aoFechar, rodape, largura = 'md', children }: ModalProps) {
  useEffect(() => {
    if (!aberto) return
    function aoTeclar(evento: KeyboardEvent) {
      if (evento.key === 'Escape') aoFechar()
    }
    document.addEventListener('keydown', aoTeclar)
    return () => document.removeEventListener('keydown', aoTeclar)
  }, [aberto, aoFechar])

  if (!aberto) return null

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-label={titulo}
    >
      <div
        className={`flex max-h-[90vh] w-full ${LARGURAS[largura]} flex-col overflow-hidden rounded-xl
          border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900`}
      >
        <header className="flex items-center gap-2 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          {Icone && <Icone size={16} className="text-swnet-600 dark:text-swnet-400" />}
          <h2 className="text-sm font-semibold text-slate-800 dark:text-white">{titulo}</h2>
          <button
            type="button"
            onClick={aoFechar}
            aria-label="Fechar"
            className="ml-auto rounded-lg p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600
              dark:hover:bg-slate-800 dark:hover:text-slate-200"
          >
            <X size={16} />
          </button>
        </header>

        <div className="flex-1 overflow-y-auto px-4 py-4">{children}</div>

        {rodape && (
          <footer className="flex justify-end gap-2 border-t border-slate-200 px-4 py-3 dark:border-slate-800">
            {rodape}
          </footer>
        )}
      </div>
    </div>
  )
}
