import { useState, type FormEvent } from 'react'
import { Navigate, useNavigate } from 'react-router-dom'
import { Moon, Sun, Lock, User } from 'lucide-react'
import { useAuth } from '@/context/AuthContext'
import { useTheme } from '@/context/ThemeContext'

export function Login() {
  const { autenticado, carregando, entrar } = useAuth()
  const { tema, alternarTema } = useTheme()
  const navegar = useNavigate()

  const [login, setLogin] = useState('')
  const [senha, setSenha] = useState('')
  const [erro, setErro] = useState<string | null>(null)

  if (autenticado) return <Navigate to="/" replace />

  async function aoSubmeter(evento: FormEvent) {
    evento.preventDefault()
    setErro(null)

    const resultado = await entrar(login, senha)
    if (!resultado.sucesso) {
      setErro(resultado.mensagemErro ?? 'Não foi possível entrar.')
      return
    }
    navegar('/', { replace: true })
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-swnet-50 px-4 dark:bg-slate-950">
      <button
        type="button"
        onClick={alternarTema}
        aria-label="Alternar tema"
        className="absolute right-4 top-4 rounded-lg p-2 text-slate-500 hover:bg-white dark:text-slate-400 dark:hover:bg-slate-900"
      >
        {tema === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
      </button>

      <div className="w-full max-w-sm rounded-2xl border border-swnet-100 bg-white p-8 shadow-xl shadow-swnet-900/5
        dark:border-slate-800 dark:bg-slate-900">
        <div className="mb-8 text-center">
          <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-xl bg-swnet-600 text-xl font-bold text-white">
            S
          </div>
          <h1 className="text-xl font-bold text-slate-800 dark:text-white">SWnet</h1>
          <p className="text-sm text-slate-400">Entre com seu usuário para continuar</p>
        </div>

        <form onSubmit={aoSubmeter} className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">Usuário</label>
            <div className="relative">
              <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                value={login}
                onChange={(e) => setLogin(e.target.value)}
                type="text"
                autoComplete="username"
                placeholder="admin"
                className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm
                  focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                  dark:border-slate-700 dark:bg-slate-800 dark:text-white"
                required
              />
            </div>
          </div>

          <div>
            <label className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">Senha</label>
            <div className="relative">
              <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                value={senha}
                onChange={(e) => setSenha(e.target.value)}
                type="password"
                autoComplete="current-password"
                placeholder="••••••••"
                className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm
                  focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                  dark:border-slate-700 dark:bg-slate-800 dark:text-white"
                required
              />
            </div>
          </div>

          {erro && (
            <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-600 dark:bg-red-950/40 dark:text-red-400">
              {erro}
            </p>
          )}

          <button
            type="submit"
            disabled={carregando}
            className="w-full rounded-lg bg-swnet-600 py-2.5 text-sm font-semibold text-white transition-colors
              hover:bg-swnet-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {carregando ? 'Entrando...' : 'Entrar'}
          </button>
        </form>

        <p className="mt-6 rounded-lg bg-swnet-50 px-3 py-2 text-center text-xs text-swnet-700 dark:bg-slate-800 dark:text-swnet-300">
          Login de teste: <strong>admin</strong> / senha: <strong>Admin@123</strong>
        </p>
      </div>
    </div>
  )
}
