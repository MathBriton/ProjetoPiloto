import { useEffect, useState } from 'react'
import { NavLink, useLocation } from 'react-router-dom'
import { ChevronsLeft, ChevronsRight, ChevronDown } from 'lucide-react'
import { useAuth } from '@/context/AuthContext'
import { ITENS_MENU, ehGrupo, type ItemMenu } from '@/config/menu'

interface SidebarProps {
  expandida: boolean
  aoAlternar: () => void
}

const classesItem = (ativo: boolean) =>
  `flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium transition-colors ${
    ativo
      ? 'bg-swnet-600 text-white shadow-sm'
      : 'text-slate-600 hover:bg-swnet-50 hover:text-swnet-700 dark:text-slate-300 dark:hover:bg-slate-800 dark:hover:text-swnet-300'
  }`

export function Sidebar({ expandida, aoAlternar }: SidebarProps) {
  const { temPermissao } = useAuth()
  const location = useLocation()
  const [gruposAbertos, setGruposAbertos] = useState<Set<string>>(new Set())

  // Abre automaticamente o grupo que contém a rota ativa
  useEffect(() => {
    for (const entrada of ITENS_MENU) {
      if (ehGrupo(entrada) && entrada.filhos.some((filho) => location.pathname.startsWith(filho.caminho))) {
        setGruposAbertos((atual) => new Set(atual).add(entrada.rotulo))
      }
    }
  }, [location.pathname])

  function alternarGrupo(rotulo: string) {
    if (!expandida) {
      aoAlternar()
    }
    setGruposAbertos((atual) => {
      const novo = new Set(atual)
      novo.has(rotulo) ? novo.delete(rotulo) : novo.add(rotulo)
      return novo
    })
  }

  function itemVisivel(item: ItemMenu) {
    return !item.permissaoNecessaria || temPermissao(item.permissaoNecessaria)
  }

  return (
    <aside
      className={`h-screen shrink-0 border-r border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-950
        flex flex-col transition-[width] duration-200 ease-in-out ${expandida ? 'w-60' : 'w-16'}`}
    >
      <div className="flex h-16 items-center justify-between border-b border-slate-200 px-3 dark:border-slate-800">
        {expandida && (
          <span className="text-lg font-bold tracking-tight text-swnet-700 dark:text-swnet-300">
            SWnet
          </span>
        )}
        <button
          type="button"
          onClick={aoAlternar}
          aria-label={expandida ? 'Retrair menu' : 'Expandir menu'}
          className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-swnet-50 hover:text-swnet-700
            dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-swnet-300"
        >
          {expandida ? <ChevronsLeft size={18} /> : <ChevronsRight size={18} />}
        </button>
      </div>

      <nav className="scrollbar-fina flex-1 space-y-1 overflow-y-auto p-2">
        {ITENS_MENU.map((entrada) => {
          if (!ehGrupo(entrada)) {
            if (!itemVisivel(entrada)) return null
            const Icone = entrada.icone
            return (
              <NavLink
                key={entrada.caminho}
                to={entrada.caminho}
                end={entrada.caminho === '/'}
                title={!expandida ? entrada.rotulo : undefined}
                className={({ isActive }) => classesItem(isActive)}
              >
                <Icone size={20} className="shrink-0" />
                {expandida && <span className="truncate">{entrada.rotulo}</span>}
              </NavLink>
            )
          }

          const filhosVisiveis = entrada.filhos.filter(itemVisivel)
          if (filhosVisiveis.length === 0) return null

          const GrupoIcone = entrada.icone
          const aberto = gruposAbertos.has(entrada.rotulo)
          const grupoAtivo = filhosVisiveis.some((filho) => location.pathname.startsWith(filho.caminho))

          return (
            <div key={entrada.rotulo}>
              <button
                type="button"
                onClick={() => alternarGrupo(entrada.rotulo)}
                title={!expandida ? entrada.rotulo : undefined}
                aria-expanded={aberto}
                className={`w-full ${classesItem(grupoAtivo && !expandida)}`}
              >
                <GrupoIcone size={20} className="shrink-0" />
                {expandida && (
                  <>
                    <span className="flex-1 truncate text-left">{entrada.rotulo}</span>
                    <ChevronDown
                      size={16}
                      className={`shrink-0 transition-transform duration-200 ${aberto ? 'rotate-180' : ''}`}
                    />
                  </>
                )}
              </button>

              {expandida && aberto && (
                <div className="mt-1 space-y-1 border-l border-slate-200 pl-4 dark:border-slate-800">
                  {filhosVisiveis.map((filho) => {
                    const Icone = filho.icone
                    return (
                      <NavLink
                        key={filho.caminho}
                        to={filho.caminho}
                        className={({ isActive }) => classesItem(isActive)}
                      >
                        <Icone size={18} className="shrink-0" />
                        <span className="truncate">{filho.rotulo}</span>
                      </NavLink>
                    )
                  })}
                </div>
              )}
            </div>
          )
        })}
      </nav>
    </aside>
  )
}
