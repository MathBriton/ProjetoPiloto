import { useEffect, useRef } from 'react'
import { ChevronDown } from 'lucide-react'
import type { ItemPermissao } from '@/pages/administrador/tipos'

interface CartaoCategoriaProps {
  categoria: string
  itens: ItemPermissao[]
  filtro: string
  aoAlternarItem: (codigo: string) => void
  aoAlternarCategoria: (categoria: string, marcar: boolean) => void
}

export function CartaoCategoria({ categoria, itens, filtro, aoAlternarItem, aoAlternarCategoria }: CartaoCategoriaProps) {
  const refCheckboxCategoria = useRef<HTMLInputElement>(null)

  const itensFiltrados = filtro
    ? itens.filter((item) => item.nome.toLowerCase().includes(filtro.toLowerCase()))
    : itens

  const total = itens.length
  const marcados = itens.filter((i) => i.habilitado).length
  const todosMarcados = marcados === total && total > 0
  const algunsMarcados = marcados > 0 && marcados < total

  useEffect(() => {
    if (refCheckboxCategoria.current) {
      refCheckboxCategoria.current.indeterminate = algunsMarcados
    }
  }, [algunsMarcados])

  if (filtro && itensFiltrados.length === 0) return null

  return (
    <article className="mb-2 overflow-hidden rounded-lg border border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
      <details open className="group">
        <summary className="flex cursor-pointer list-none items-center justify-between gap-2 border-b border-slate-200
          px-3 py-2 hover:bg-slate-50 dark:border-slate-800 dark:hover:bg-slate-800/60">
          <div className="flex items-center gap-2">
            <input
              ref={refCheckboxCategoria}
              type="checkbox"
              checked={todosMarcados}
              onChange={(e) => {
                e.stopPropagation()
                aoAlternarCategoria(categoria, e.target.checked)
              }}
              onClick={(e) => e.stopPropagation()}
              aria-label={`Selecionar categoria ${categoria}`}
              className="h-4 w-4 accent-swnet-600"
            />
            <span className="text-xs font-semibold text-slate-700 dark:text-slate-200">{categoria}</span>
            <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[0.65rem] font-medium text-slate-500 dark:bg-slate-800 dark:text-slate-400">
              {marcados}/{total}
            </span>
          </div>
          <ChevronDown size={14} className="shrink-0 text-slate-400 transition-transform group-open:rotate-180" />
        </summary>

        <div className="divide-y divide-slate-100 dark:divide-slate-800">
          {itensFiltrados.map((item) => (
            <label
              key={item.codigo}
              className={`flex cursor-pointer items-center gap-2 px-3 py-2 text-xs transition-colors hover:bg-swnet-50/60 dark:hover:bg-slate-800/40 ${
                item.habilitado ? 'bg-swnet-50/40 dark:bg-slate-800/30' : ''
              }`}
            >
              <input
                type="checkbox"
                checked={item.habilitado}
                onChange={() => aoAlternarItem(item.codigo)}
                className="h-4 w-4 accent-swnet-600"
              />
              <span className="text-slate-700 dark:text-slate-200">{item.nome}</span>
            </label>
          ))}
        </div>
      </details>
    </article>
  )
}
