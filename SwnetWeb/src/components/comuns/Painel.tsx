import type { ReactNode } from 'react'
import type { LucideIcon } from 'lucide-react'

/**
 * Card com faixa de título — equivalente ao `.panel-header` / `.card-header`
 * repetido nas telas ASP do Administrador.
 */

interface PainelProps {
  titulo: string
  icone?: LucideIcon
  acoes?: ReactNode
  children: ReactNode
  className?: string
}

export function Painel({ titulo, icone: Icone, acoes, children, className = '' }: PainelProps) {
  return (
    <section
      className={`overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm
        dark:border-slate-800 dark:bg-slate-900 ${className}`}
    >
      <header
        className="flex items-center gap-2 border-b border-slate-200 px-4 py-3
          dark:border-slate-800"
      >
        {Icone && <Icone size={16} className="shrink-0 text-slate-400" />}
        <h2 className="text-xs font-semibold uppercase tracking-wide text-slate-600 dark:text-slate-300">
          {titulo}
        </h2>
        {acoes && <div className="ml-auto flex items-center gap-2">{acoes}</div>}
      </header>
      <div className="p-4">{children}</div>
    </section>
  )
}

/** Bloco interno de um painel (as sub-seções "Dados Pessoais", "Vínculo" etc.) */
export function Secao({ titulo, icone: Icone, children }: Omit<PainelProps, 'acoes' | 'className'>) {
  return (
    <fieldset className="mb-4 rounded-lg border border-slate-200 last:mb-0 dark:border-slate-800">
      <legend className="ml-3 flex items-center gap-1.5 px-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
        {Icone && <Icone size={13} />}
        {titulo}
      </legend>
      <div className="px-4 pb-4 pt-2">{children}</div>
    </fieldset>
  )
}
