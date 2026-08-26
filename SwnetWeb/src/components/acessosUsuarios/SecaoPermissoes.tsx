import type { LucideIcon } from 'lucide-react'
import type { ItemPermissao } from '@/pages/administrador/tipos'
import { CartaoCategoria } from './CartaoCategoria'

interface SecaoPermissoesProps {
  titulo: string
  icone: LucideIcon
  itens: ItemPermissao[]
  filtro: string
  aoAlternarItem: (codigo: string) => void
  aoAlternarCategoria: (categoria: string, marcar: boolean) => void
}

export function SecaoPermissoes({ titulo, icone: Icone, itens, filtro, aoAlternarItem, aoAlternarCategoria }: SecaoPermissoesProps) {
  const categorias = Array.from(new Set(itens.map((item) => item.categoria)))

  return (
    <section className="flex-1 rounded-xl border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-900">
      <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
        <Icone size={16} className="text-swnet-600 dark:text-swnet-400" />
        <h2 className="text-sm font-semibold text-slate-700 dark:text-slate-200">{titulo}</h2>
      </div>

      <div className="scrollbar-fina max-h-[60vh] overflow-y-auto p-3">
        {categorias.map((categoria) => (
          <CartaoCategoria
            key={categoria}
            categoria={categoria}
            itens={itens.filter((item) => item.categoria === categoria)}
            filtro={filtro}
            aoAlternarItem={aoAlternarItem}
            aoAlternarCategoria={aoAlternarCategoria}
          />
        ))}
      </div>
    </section>
  )
}
