import { useState, useRef, useEffect } from 'react'
import { Search, Factory, Moon, Sun, ChevronDown, LogOut } from 'lucide-react'
import { useAuth } from '@/context/AuthContext'
import { useTheme } from '@/context/ThemeContext'
import { usePlanta } from '@/context/PlantaContext'

export function TopBar() {
  const { usuario, sair } = useAuth()
  const { tema, alternarTema } = useTheme()
  const { plantaAtual, plantas, trocarPlanta } = usePlanta()

  const [menuPlantaAberto, setMenuPlantaAberto] = useState(false)
  const [menuUsuarioAberto, setMenuUsuarioAberto] = useState(false)
  const [busca, setBusca] = useState('')
  const refPlanta = useRef<HTMLDivElement>(null)
  const refUsuario = useRef<HTMLDivElement>(null)

  useEffect(() => {
    function aoClicarFora(evento: MouseEvent) {
      if (refPlanta.current && !refPlanta.current.contains(evento.target as Node)) {
        setMenuPlantaAberto(false)
      }
      if (refUsuario.current && !refUsuario.current.contains(evento.target as Node)) {
        setMenuUsuarioAberto(false)
      }
    }
    document.addEventListener('mousedown', aoClicarFora)
    return () => document.removeEventListener('mousedown', aoClicarFora)
  }, [])

  function aoSubmeterBusca(evento: React.FormEvent) {
    evento.preventDefault()
    // TODO: ligar na busca global real (endpoint /api/pesquisa) quando existir
    console.log('Pesquisar globalmente por:', busca)
  }

  return (
    <header className="flex h-16 items-center gap-4 border-b border-slate-200 bg-white px-4 dark:border-slate-800 dark:bg-slate-950">
      {/* Pesquisa global */}
      <form onSubmit={aoSubmeterBusca} className="relative flex-1 max-w-md">
        <Search
          size={18}
          className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
        />
        <input
          type="text"
          value={busca}
          onChange={(e) => setBusca(e.target.value)}
          placeholder="Pesquisar em todo o SWnet..."
          className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2 pl-10 pr-3 text-sm
            text-slate-700 placeholder:text-slate-400 focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
            dark:border-slate-800 dark:bg-slate-900 dark:text-slate-200"
        />
      </form>

      <div className="ml-auto flex items-center gap-2">
        {/* Troca de planta (fábrica) */}
        <div className="relative" ref={refPlanta}>
          <button
            type="button"
            onClick={() => setMenuPlantaAberto((v) => !v)}
            className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium
              text-slate-600 hover:bg-swnet-50 hover:text-swnet-700
              dark:border-slate-800 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            <Factory size={16} />
            {plantaAtual ? plantaAtual.nome : 'Selecionar planta'}
            <ChevronDown size={14} />
          </button>

          {menuPlantaAberto && (
            <div className="absolute right-0 z-20 mt-2 w-56 overflow-hidden rounded-lg border border-slate-200
              bg-white py-1 shadow-lg dark:border-slate-800 dark:bg-slate-900">
              <p className="px-3 py-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">
                Fábricas do grupo
              </p>
              {plantas.map((planta) => (
                <button
                  key={planta.id}
                  onClick={() => {
                    trocarPlanta(planta.id)
                    setMenuPlantaAberto(false)
                  }}
                  className={`flex w-full items-center justify-between px-3 py-2 text-sm hover:bg-swnet-50 dark:hover:bg-slate-800 ${
                    planta.id === plantaAtual?.id ? 'text-swnet-700 dark:text-swnet-300 font-medium' : 'text-slate-600 dark:text-slate-300'
                  }`}
                >
                  {planta.nome}
                  <span className="text-xs text-slate-400">{planta.codigo}</span>
                </button>
              ))}
            </div>
          )}
        </div>

        {/* Dark mode */}
        <button
          type="button"
          onClick={alternarTema}
          aria-label="Alternar tema"
          className="rounded-lg p-2 text-slate-500 hover:bg-swnet-50 hover:text-swnet-700
            dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-swnet-300"
        >
          {tema === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
        </button>

        {/* Usuário */}
        <div className="relative" ref={refUsuario}>
          <button
            type="button"
            onClick={() => setMenuUsuarioAberto((v) => !v)}
            className="flex items-center gap-2 rounded-lg px-2 py-1.5 hover:bg-swnet-50 dark:hover:bg-slate-800"
          >
            <span className="flex h-8 w-8 items-center justify-center rounded-full bg-swnet-600 text-sm font-semibold text-white">
              {usuario?.nomeCompleto.charAt(0).toUpperCase()}
            </span>
            <span className="hidden text-sm font-medium text-slate-700 sm:block dark:text-slate-200">
              {usuario?.nomeCompleto}
            </span>
          </button>

          {menuUsuarioAberto && (
            <div className="absolute right-0 z-20 mt-2 w-52 overflow-hidden rounded-lg border border-slate-200
              bg-white py-1 shadow-lg dark:border-slate-800 dark:bg-slate-900">
              <div className="px-3 py-2">
                <p className="text-sm font-medium text-slate-700 dark:text-slate-200">{usuario?.nomeCompleto}</p>
                <p className="text-xs text-slate-400">{usuario?.perfil}</p>
              </div>
              <button
                onClick={sair}
                className="flex w-full items-center gap-2 px-3 py-2 text-sm text-red-600 hover:bg-red-50 dark:hover:bg-red-950/40"
              >
                <LogOut size={15} /> Sair
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  )
}
