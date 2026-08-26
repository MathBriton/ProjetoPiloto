import type { ReactNode } from 'react'
import { Search, X, type LucideIcon } from 'lucide-react'

/** Peças pequenas repetidas nas telas do Administrador. */

export function TituloPagina({ titulo, acoes }: { titulo: string; acoes?: ReactNode }) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 px-1">
      <h1 className="text-xl font-bold text-slate-800 dark:text-white">{titulo}</h1>
      {acoes && <div className="flex flex-wrap items-center gap-2">{acoes}</div>}
    </div>
  )
}

type Tom = 'sucesso' | 'neutro' | 'alerta' | 'perigo' | 'info'

const TONS: Record<Tom, string> = {
  sucesso: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-500/10 dark:text-emerald-400',
  neutro: 'bg-slate-100 text-slate-500 dark:bg-slate-800 dark:text-slate-400',
  alerta: 'bg-amber-50 text-amber-700 dark:bg-amber-500/10 dark:text-amber-400',
  perigo: 'bg-red-50 text-red-700 dark:bg-red-500/10 dark:text-red-400',
  info: 'bg-swnet-50 text-swnet-700 dark:bg-swnet-500/10 dark:text-swnet-300'
}

export function Etiqueta({ texto, tom = 'neutro' }: { texto: string; tom?: Tom }) {
  return (
    <span className={`inline-block rounded-full px-2.5 py-0.5 text-xs font-medium ${TONS[tom]}`}>{texto}</span>
  )
}

interface CampoPesquisaProps {
  valor: string
  aoMudar: (valor: string) => void
  placeholder?: string
  larguraClasse?: string
  rotuloAcessivel?: string
}

export function CampoPesquisa({
  valor,
  aoMudar,
  placeholder = 'Pesquisar',
  larguraClasse = 'w-56',
  rotuloAcessivel = 'Pesquisar'
}: CampoPesquisaProps) {
  return (
    <div className="relative">
      <Search size={14} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
      <input
        type="text"
        value={valor}
        aria-label={rotuloAcessivel}
        onChange={(e) => aoMudar(e.target.value)}
        placeholder={placeholder}
        className={`${larguraClasse} rounded-lg border border-slate-200 bg-slate-50 py-1.5 pl-8 pr-7 text-xs
          focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
          dark:border-slate-700 dark:bg-slate-800 dark:text-white`}
      />
      {valor && (
        <button
          type="button"
          onClick={() => aoMudar('')}
          aria-label="Limpar pesquisa"
          className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
        >
          <X size={13} />
        </button>
      )}
    </div>
  )
}

interface EstadoVazioProps {
  icone: LucideIcon
  titulo: string
  descricao?: string
}

export function EstadoVazio({ icone: Icone, titulo, descricao }: EstadoVazioProps) {
  return (
    <div className="flex flex-col items-center gap-2 px-4 py-12 text-center text-slate-400">
      <Icone size={38} strokeWidth={1.2} />
      <p className="text-sm font-medium text-slate-500 dark:text-slate-400">{titulo}</p>
      {descricao && <p className="max-w-md text-xs">{descricao}</p>}
    </div>
  )
}

type VarianteBotao = 'primario' | 'sucesso' | 'perigo' | 'neutro'

const VARIANTES: Record<VarianteBotao, string> = {
  primario: 'bg-swnet-600 text-white hover:bg-swnet-700',
  sucesso: 'bg-emerald-600 text-white hover:bg-emerald-700',
  perigo: 'bg-red-600 text-white hover:bg-red-700',
  neutro:
    'border border-slate-200 text-slate-600 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800'
}

interface BotaoProps {
  children: ReactNode
  aoClicar?: () => void
  variante?: VarianteBotao
  icone?: LucideIcon
  desabilitado?: boolean
  tipo?: 'button' | 'submit'
  titulo?: string
}

export function Botao({
  children,
  aoClicar,
  variante = 'primario',
  icone: Icone,
  desabilitado,
  tipo = 'button',
  titulo
}: BotaoProps) {
  return (
    <button
      type={tipo}
      onClick={aoClicar}
      disabled={desabilitado}
      title={titulo}
      className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium
        transition-colors disabled:cursor-not-allowed disabled:opacity-50 ${VARIANTES[variante]}`}
    >
      {Icone && <Icone size={16} />}
      {children}
    </button>
  )
}

export function LinhaCarregando({ colunas, mensagem = 'Carregando...' }: { colunas: number; mensagem?: string }) {
  return (
    <tr>
      <td colSpan={colunas} className="px-4 py-10 text-center text-sm text-slate-400">
        {mensagem}
      </td>
    </tr>
  )
}
