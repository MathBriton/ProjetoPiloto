import { useEffect, useRef, useState } from 'react'
import { Search, UserCircle2, Building2 } from 'lucide-react'
import { buscarUsuarios } from '@/services/acessosUsuariosApi'
import type { UsuarioBusca } from '@/pages/administrador/tipos'

interface BuscaUsuarioProps {
  aoSelecionar: (usuario: UsuarioBusca) => void
}

export function BuscaUsuario({ aoSelecionar }: BuscaUsuarioProps) {
  const [termo, setTermo] = useState('')
  const [resultados, setResultados] = useState<UsuarioBusca[]>([])
  const [dropdownAberto, setDropdownAberto] = useState(false)
  const refContainer = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (termo.trim().length === 0) {
      setResultados([])
      return
    }
    const timeout = setTimeout(() => {
      buscarUsuarios(termo).then((r) => {
        setResultados(r)
        setDropdownAberto(true)
      })
    }, 250)
    return () => clearTimeout(timeout)
  }, [termo])

  useEffect(() => {
    function aoClicarFora(evento: MouseEvent) {
      if (refContainer.current && !refContainer.current.contains(evento.target as Node)) {
        setDropdownAberto(false)
      }
    }
    document.addEventListener('mousedown', aoClicarFora)
    return () => document.removeEventListener('mousedown', aoClicarFora)
  }, [])

  function selecionar(usuario: UsuarioBusca) {
    setTermo('')
    setResultados([])
    setDropdownAberto(false)
    aoSelecionar(usuario)
  }

  return (
    <section className="mx-auto max-w-3xl" aria-label="Selecionar usuário">
      <div className="rounded-xl border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-900">
        <div className="rounded-t-xl border-b border-slate-200 px-4 py-3 text-sm font-semibold uppercase tracking-wide
          text-slate-600 dark:border-slate-800 dark:text-slate-300">
          Controle de Acessos
        </div>

        <div className="p-6">
          <div className="grid gap-4 sm:grid-cols-2">
            <div className="relative" ref={refContainer}>
              <label htmlFor="buscaUsuario" className="mb-1 flex items-center gap-1 text-sm font-medium text-slate-600 dark:text-slate-300">
                <UserCircle2 size={15} /> Usuário
              </label>
              <div className="relative">
                <Search size={16} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  id="buscaUsuario"
                  type="text"
                  value={termo}
                  onChange={(e) => setTermo(e.target.value)}
                  onFocus={() => resultados.length > 0 && setDropdownAberto(true)}
                  placeholder="Digite o nome ou login do usuário..."
                  autoComplete="off"
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm
                    focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                    dark:border-slate-700 dark:bg-slate-800 dark:text-white"
                />
              </div>

              {dropdownAberto && resultados.length > 0 && (
                <div className="absolute z-20 mt-1 w-full overflow-hidden rounded-lg border border-slate-200 bg-white
                  shadow-lg dark:border-slate-800 dark:bg-slate-900">
                  {resultados.map((usuario) => (
                    <button
                      key={`${usuario.matricula}-${usuario.filial}`}
                      onClick={() => selecionar(usuario)}
                      className="flex w-full items-center justify-between px-3 py-2 text-left text-sm hover:bg-swnet-50 dark:hover:bg-slate-800"
                    >
                      <span className="text-slate-700 dark:text-slate-200">
                        {usuario.nome} <span className="text-slate-400">— {usuario.login}</span>
                      </span>
                      <span className="text-xs text-slate-400">Filial {usuario.filial}</span>
                    </button>
                  ))}
                </div>
              )}

              {dropdownAberto && termo.trim() !== '' && resultados.length === 0 && (
                <div className="absolute z-20 mt-1 w-full rounded-lg border border-slate-200 bg-white px-3 py-2
                  text-sm text-slate-400 shadow-lg dark:border-slate-800 dark:bg-slate-900">
                  Nenhum usuário encontrado.
                </div>
              )}
            </div>

            <div>
              <label htmlFor="filialDisplay" className="mb-1 flex items-center gap-1 text-sm font-medium text-slate-600 dark:text-slate-300">
                <Building2 size={15} /> Filial
              </label>
              <input
                id="filialDisplay"
                type="text"
                readOnly
                value=""
                placeholder="Filial"
                className="w-full cursor-default rounded-lg border border-slate-200 bg-slate-100 py-2.5 px-3 text-sm
                  text-slate-500 dark:border-slate-700 dark:bg-slate-800/50 dark:text-slate-400"
              />
            </div>
          </div>

          <div className="mt-8 flex flex-col items-center gap-2 py-6 text-center text-slate-400">
            <UserCircle2 size={40} strokeWidth={1.2} />
            <p className="text-sm">Selecione um usuário acima para visualizar os dados e gerenciar a conta.</p>
          </div>
        </div>
      </div>
    </section>
  )
}
