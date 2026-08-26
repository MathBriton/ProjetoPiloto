import { useEffect, useState } from 'react'
import { UserPlus, X } from 'lucide-react'
import { listarPerfis, listarPlantas, loginDisponivel } from '@/services/usuariosApi'
import type { NovoUsuario } from '@/pages/administrador/tipos'

interface FormularioUsuarioProps {
  aberto: boolean
  aoFechar: () => void
  aoSalvar: (dados: NovoUsuario) => Promise<void>
}

const ESTADO_INICIAL: NovoUsuario = {
  login: '',
  nomeCompleto: '',
  email: '',
  perfilId: 0,
  plantaIds: [],
  ativo: true
}

const classesCampo = `w-full rounded-lg border border-slate-200 bg-slate-50 px-3 py-2.5 text-sm
  focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
  dark:border-slate-700 dark:bg-slate-800 dark:text-white`

export function FormularioUsuario({ aberto, aoFechar, aoSalvar }: FormularioUsuarioProps) {
  const perfis = listarPerfis()
  const plantas = listarPlantas()

  const [dados, setDados] = useState<NovoUsuario>(ESTADO_INICIAL)
  const [erros, setErros] = useState<Partial<Record<keyof NovoUsuario, string>>>({})
  const [salvando, setSalvando] = useState(false)

  useEffect(() => {
    if (aberto) {
      setDados(ESTADO_INICIAL)
      setErros({})
    }
  }, [aberto])

  if (!aberto) return null

  function alterar<C extends keyof NovoUsuario>(campo: C, valor: NovoUsuario[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
    setErros((atual) => ({ ...atual, [campo]: undefined }))
  }

  function alternarPlanta(id: number) {
    setDados((atual) => ({
      ...atual,
      plantaIds: atual.plantaIds.includes(id)
        ? atual.plantaIds.filter((p) => p !== id)
        : [...atual.plantaIds, id]
    }))
    setErros((atual) => ({ ...atual, plantaIds: undefined }))
  }

  async function validar(): Promise<boolean> {
    const novosErros: Partial<Record<keyof NovoUsuario, string>> = {}

    if (dados.login.trim().length < 3) novosErros.login = 'Informe ao menos 3 caracteres.'
    else if (!/^[a-zA-Z0-9._-]+$/.test(dados.login.trim()))
      novosErros.login = 'Use apenas letras, números, ponto, hífen ou underline.'
    else if (!(await loginDisponivel(dados.login))) novosErros.login = 'Este login já está em uso.'

    if (dados.nomeCompleto.trim().length < 3) novosErros.nomeCompleto = 'Informe o nome completo.'
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(dados.email.trim())) novosErros.email = 'E-mail inválido.'
    if (!dados.perfilId) novosErros.perfilId = 'Selecione um perfil.'
    if (dados.plantaIds.length === 0) novosErros.plantaIds = 'Selecione ao menos uma planta.'

    setErros(novosErros)
    return Object.keys(novosErros).length === 0
  }

  async function enviar(evento: React.FormEvent) {
    evento.preventDefault()
    setSalvando(true)
    try {
      if (!(await validar())) return
      await aoSalvar({
        ...dados,
        login: dados.login.trim(),
        nomeCompleto: dados.nomeCompleto.trim(),
        email: dados.email.trim()
      })
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 px-4" role="dialog" aria-modal="true">
      <form
        onSubmit={enviar}
        className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-xl border border-slate-200 bg-white
          shadow-xl dark:border-slate-800 dark:bg-slate-900"
      >
        <div className="flex items-center gap-2 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
          <UserPlus size={17} className="text-swnet-600 dark:text-swnet-300" />
          <h2 className="text-sm font-semibold text-slate-800 dark:text-white">Novo usuário</h2>
          <button
            type="button"
            onClick={aoFechar}
            aria-label="Fechar"
            className="ml-auto text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
          >
            <X size={16} />
          </button>
        </div>

        <div className="space-y-4 px-4 py-4">
          <div className="grid gap-4 sm:grid-cols-2">
            <div>
              <label htmlFor="login" className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">
                Login
              </label>
              <input
                id="login"
                value={dados.login}
                onChange={(e) => alterar('login', e.target.value)}
                autoComplete="off"
                className={classesCampo}
              />
              {erros.login && <p className="mt-1 text-xs text-red-500">{erros.login}</p>}
            </div>

            <div>
              <label htmlFor="perfil" className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">
                Perfil
              </label>
              <select
                id="perfil"
                value={dados.perfilId}
                onChange={(e) => alterar('perfilId', Number(e.target.value))}
                className={classesCampo}
              >
                <option value={0}>Selecione...</option>
                {perfis.map((perfil) => (
                  <option key={perfil.id} value={perfil.id}>
                    {perfil.nome}
                  </option>
                ))}
              </select>
              {erros.perfilId && <p className="mt-1 text-xs text-red-500">{erros.perfilId}</p>}
            </div>
          </div>

          <div>
            <label htmlFor="nomeCompleto" className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">
              Nome completo
            </label>
            <input
              id="nomeCompleto"
              value={dados.nomeCompleto}
              onChange={(e) => alterar('nomeCompleto', e.target.value)}
              className={classesCampo}
            />
            {erros.nomeCompleto && <p className="mt-1 text-xs text-red-500">{erros.nomeCompleto}</p>}
          </div>

          <div>
            <label htmlFor="email" className="mb-1 block text-sm font-medium text-slate-600 dark:text-slate-300">
              E-mail
            </label>
            <input
              id="email"
              type="email"
              value={dados.email}
              onChange={(e) => alterar('email', e.target.value)}
              className={classesCampo}
            />
            {erros.email && <p className="mt-1 text-xs text-red-500">{erros.email}</p>}
          </div>

          <fieldset>
            <legend className="mb-1 text-sm font-medium text-slate-600 dark:text-slate-300">Plantas</legend>
            <div className="flex flex-wrap gap-2">
              {plantas.map((planta) => {
                const marcada = dados.plantaIds.includes(planta.id)
                return (
                  <button
                    key={planta.id}
                    type="button"
                    onClick={() => alternarPlanta(planta.id)}
                    aria-pressed={marcada}
                    className={`rounded-lg border px-3 py-1.5 text-xs font-medium transition-colors ${
                      marcada
                        ? 'border-swnet-600 bg-swnet-600 text-white'
                        : 'border-slate-200 text-slate-600 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800'
                    }`}
                  >
                    {planta.nome}
                  </button>
                )
              })}
            </div>
            {erros.plantaIds && <p className="mt-1 text-xs text-red-500">{erros.plantaIds}</p>}
          </fieldset>

          <label className="flex items-center gap-2 text-sm text-slate-600 dark:text-slate-300">
            <input
              type="checkbox"
              checked={dados.ativo}
              onChange={(e) => alterar('ativo', e.target.checked)}
              className="h-4 w-4 rounded border-slate-300 text-swnet-600 focus:ring-swnet-500"
            />
            Usuário ativo
          </label>

          <p className="rounded-lg bg-slate-50 px-3 py-2 text-xs text-slate-500 dark:bg-slate-800/50 dark:text-slate-400">
            A senha provisória será enviada por e-mail quando o backend de cadastro estiver disponível.
          </p>
        </div>

        <div className="flex justify-end gap-2 border-t border-slate-200 px-4 py-3 dark:border-slate-800">
          <button
            type="button"
            onClick={aoFechar}
            className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-600
              hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
          >
            Cancelar
          </button>
          <button
            type="submit"
            disabled={salvando}
            className="rounded-lg bg-swnet-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-swnet-700 disabled:opacity-50"
          >
            {salvando ? 'Salvando...' : 'Criar usuário'}
          </button>
        </div>
      </form>
    </div>
  )
}
