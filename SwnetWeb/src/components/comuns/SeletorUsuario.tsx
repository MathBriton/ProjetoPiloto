import { useEffect, useRef, useState } from 'react'
import { Search, UserRound, X } from 'lucide-react'
import { listarUsuarios } from '@/services/administrador/usuariosSwnetApi'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'

/**
 * Autocomplete de usuário reaproveitado por Alterar Usuário, Ativar/Desativar
 * e Permissões Especiais — no legado cada tela embutia sua própria lista
 * (`listaUsuarios`) renderizada pelo ASP.
 */

interface SeletorUsuarioProps {
  rotulo?: string
  selecionado: UsuarioSwnet | null
  aoSelecionar: (usuario: UsuarioSwnet | null) => void
  apenasAtivos?: boolean
  /** Máximo de itens exibidos no dropdown (o legado mostra um contador) */
  limite?: number
}

export function SeletorUsuario({
  rotulo = 'Usuário',
  selecionado,
  aoSelecionar,
  apenasAtivos = false,
  limite = 8
}: SeletorUsuarioProps) {
  const [termo, setTermo] = useState('')
  const [resultados, setResultados] = useState<UsuarioSwnet[]>([])
  const [aberto, setAberto] = useState(false)
  const refContainer = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (selecionado) return

    const timeout = setTimeout(() => {
      listarUsuarios(termo, apenasAtivos).then(setResultados)
    }, 250)
    return () => clearTimeout(timeout)
  }, [termo, apenasAtivos, selecionado])

  useEffect(() => {
    function aoClicarFora(evento: MouseEvent) {
      if (refContainer.current && !refContainer.current.contains(evento.target as Node)) {
        setAberto(false)
      }
    }
    document.addEventListener('mousedown', aoClicarFora)
    return () => document.removeEventListener('mousedown', aoClicarFora)
  }, [])

  function selecionar(usuario: UsuarioSwnet) {
    setTermo('')
    setAberto(false)
    aoSelecionar(usuario)
  }

  function limpar() {
    setTermo('')
    setAberto(false)
    aoSelecionar(null)
  }

  const visiveis = resultados.slice(0, limite)
  const ocultos = resultados.length - visiveis.length

  if (selecionado) {
    return (
      <div>
        <span className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300">{rotulo}</span>
        <div
          className="flex items-center gap-3 rounded-lg border border-swnet-200 bg-swnet-50 px-3 py-2
            dark:border-swnet-900 dark:bg-swnet-950/40"
        >
          <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-swnet-600 text-white">
            <UserRound size={16} />
          </span>
          <span className="min-w-0 flex-1">
            <span className="block truncate text-sm font-semibold text-slate-800 dark:text-white">
              {selecionado.nome}
            </span>
            <span className="block truncate text-xs text-slate-500 dark:text-slate-400">
              {selecionado.login} &bull; Matrícula {selecionado.matricula}
              {!selecionado.ativo && ' • inativo'}
            </span>
          </span>
          <button
            type="button"
            onClick={limpar}
            aria-label="Trocar usuário"
            title="Trocar usuário"
            className="rounded-lg p-1.5 text-slate-400 hover:bg-white hover:text-slate-600
              dark:hover:bg-slate-800 dark:hover:text-slate-200"
          >
            <X size={16} />
          </button>
        </div>
      </div>
    )
  }

  return (
    <div className="relative" ref={refContainer}>
      <label htmlFor="seletorUsuario" className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300">
        {rotulo}
      </label>
      <div className="relative">
        <Search size={16} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          id="seletorUsuario"
          type="text"
          autoComplete="off"
          value={termo}
          onChange={(e) => {
            setTermo(e.target.value)
            setAberto(true)
          }}
          onFocus={() => setAberto(true)}
          placeholder="Digite o nome ou login do usuário..."
          className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm
            focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
            dark:border-slate-700 dark:bg-slate-800 dark:text-white"
        />
      </div>

      {aberto && (
        <div
          className="absolute z-30 mt-1 max-h-72 w-full overflow-y-auto rounded-lg border border-slate-200
            bg-white shadow-lg dark:border-slate-800 dark:bg-slate-900"
        >
          {visiveis.map((usuario) => (
            <button
              key={usuario.id}
              type="button"
              onClick={() => selecionar(usuario)}
              className="flex w-full items-center justify-between gap-3 px-3 py-2 text-left text-sm
                hover:bg-swnet-50 dark:hover:bg-slate-800"
            >
              <span className="min-w-0">
                <span className="block truncate text-slate-700 dark:text-slate-200">{usuario.nome}</span>
                <span className="block truncate text-xs text-slate-400">{usuario.login}</span>
              </span>
              <span
                className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-medium ${
                  usuario.ativo
                    ? 'bg-emerald-50 text-emerald-700 dark:bg-emerald-500/10 dark:text-emerald-400'
                    : 'bg-slate-100 text-slate-500 dark:bg-slate-800 dark:text-slate-400'
                }`}
              >
                {usuario.ativo ? 'Ativo' : 'Inativo'}
              </span>
            </button>
          ))}

          {ocultos > 0 && (
            <p className="border-t border-slate-100 px-3 py-1.5 text-center text-xs text-slate-400 dark:border-slate-800">
              + {ocultos} usuário(s). Refine a busca.
            </p>
          )}

          {visiveis.length === 0 && (
            <p className="px-3 py-2 text-sm text-slate-400">Nenhum usuário encontrado.</p>
          )}
        </div>
      )}
    </div>
  )
}
