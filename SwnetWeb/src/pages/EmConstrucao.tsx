import { useLocation } from 'react-router-dom'
import { Construction, FolderSearch } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { TituloPagina } from '@/components/comuns/Comuns'
import { ITENS_PLANOS } from '@/config/menu'

/**
 * Tela usada pelos menus que ainda não foram refatorados.
 *
 * O fonte ASP já está no material; a rota e o item de menu existem para a
 * navegação ficar completa, e apontamos a pasta de origem para facilitar
 * retomar de onde parou.
 */
export function EmConstrucao() {
  const location = useLocation()
  const item = ITENS_PLANOS.find((entrada) => entrada.caminho === location.pathname)

  return (
    <div className="space-y-4">
      <TituloPagina titulo={item?.rotulo ?? 'Tela pendente'} />

      <div className="mx-auto max-w-2xl">
        <Painel titulo="Refatoração pendente" icone={Construction}>
          <div className="flex flex-col items-center gap-4 px-4 py-8 text-center">
            <Construction size={44} strokeWidth={1.2} className="text-amber-500" />

            <div>
              <p className="text-sm font-medium text-slate-700 dark:text-slate-200">
                Esta tela ainda não foi refatorada.
              </p>
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                A rota e o item de menu já estão criados, e o fonte ASP está na pasta abaixo — só falta
                a implementação em React.
              </p>
            </div>

            {item?.origem && (
              <p
                className="flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2
                  text-xs text-slate-600 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-300"
              >
                <FolderSearch size={14} className="shrink-0 text-slate-400" />
                <code className="break-all">Projetos-A-Refatorar/Menu-Serviços-Sidebar/{item.origem}</code>
              </p>
            )}
          </div>
        </Painel>
      </div>
    </div>
  )
}
