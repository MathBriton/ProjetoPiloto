import type { InputHTMLAttributes, ReactNode, SelectHTMLAttributes } from 'react'

/**
 * Primitivas de formulário usadas pelas telas do Administrador.
 * Centralizam o visual (label + controle + erro) para as telas não repetirem
 * as mesmas ~6 classes do Tailwind em cada input.
 */

export const classesControle =
  `w-full rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-700
   focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
   disabled:cursor-not-allowed disabled:opacity-60
   dark:border-slate-700 dark:bg-slate-800 dark:text-white`

interface CampoProps {
  rotulo: string
  htmlFor?: string
  erro?: string
  ajuda?: string
  obrigatorio?: boolean
  className?: string
  children: ReactNode
}

export function Campo({ rotulo, htmlFor, erro, ajuda, obrigatorio, className = '', children }: CampoProps) {
  return (
    <div className={className}>
      <label
        htmlFor={htmlFor}
        className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300"
      >
        {rotulo}
        {obrigatorio && <span className="ml-0.5 text-red-500">*</span>}
      </label>
      {children}
      {erro && <p className="mt-1 text-xs font-medium text-red-600 dark:text-red-400">{erro}</p>}
      {!erro && ajuda && <p className="mt-1 text-xs text-slate-400">{ajuda}</p>}
    </div>
  )
}

type EntradaProps = InputHTMLAttributes<HTMLInputElement> & {
  rotulo: string
  erro?: string
  ajuda?: string
  larguraClasse?: string
}

export function Entrada({ rotulo, erro, ajuda, larguraClasse = '', ...props }: EntradaProps) {
  return (
    <Campo
      rotulo={rotulo}
      htmlFor={props.id}
      erro={erro}
      ajuda={ajuda}
      obrigatorio={props.required}
      className={larguraClasse}
    >
      <input
        {...props}
        className={`${classesControle} ${erro ? 'border-red-400 dark:border-red-500' : ''}`}
      />
    </Campo>
  )
}

type SelecaoProps = SelectHTMLAttributes<HTMLSelectElement> & {
  rotulo: string
  erro?: string
  ajuda?: string
  larguraClasse?: string
  children: ReactNode
}

export function Selecao({ rotulo, erro, ajuda, larguraClasse = '', children, ...props }: SelecaoProps) {
  return (
    <Campo
      rotulo={rotulo}
      htmlFor={props.id}
      erro={erro}
      ajuda={ajuda}
      obrigatorio={props.required}
      className={larguraClasse}
    >
      <select
        {...props}
        className={`${classesControle} ${erro ? 'border-red-400 dark:border-red-500' : ''}`}
      >
        {children}
      </select>
    </Campo>
  )
}

interface OpcoesRadioProps {
  rotulo: string
  nome: string
  valor: string
  opcoes: { valor: string; rotulo: string }[]
  aoMudar: (valor: string) => void
  desabilitado?: boolean
  larguraClasse?: string
}

export function OpcoesRadio({
  rotulo,
  nome,
  valor,
  opcoes,
  aoMudar,
  desabilitado,
  larguraClasse = ''
}: OpcoesRadioProps) {
  return (
    <Campo rotulo={rotulo} className={larguraClasse}>
      <div className="flex items-center gap-4 pt-1.5">
        {opcoes.map((opcao) => (
          <label
            key={opcao.valor}
            className={`flex items-center gap-1.5 text-sm text-slate-600 dark:text-slate-300 ${
              desabilitado ? 'cursor-not-allowed opacity-60' : 'cursor-pointer'
            }`}
          >
            <input
              type="radio"
              name={nome}
              value={opcao.valor}
              checked={valor === opcao.valor}
              disabled={desabilitado}
              onChange={(e) => aoMudar(e.target.value)}
              className="h-4 w-4 accent-swnet-600"
            />
            {opcao.rotulo}
          </label>
        ))}
      </div>
    </Campo>
  )
}

interface ChaveProps {
  rotulo: string
  marcado: boolean
  aoAlternar: (marcado: boolean) => void
  descricao?: string
  desabilitado?: boolean
}

/** Switch — os `form-switch` do Bootstrap nas telas de permissões */
export function Chave({ rotulo, marcado, aoAlternar, descricao, desabilitado }: ChaveProps) {
  return (
    <label
      className={`flex items-center gap-3 rounded-lg border border-slate-200 px-3 py-2.5
        dark:border-slate-700 ${
          desabilitado ? 'cursor-not-allowed opacity-60' : 'cursor-pointer hover:border-swnet-300 dark:hover:border-slate-600'
        }`}
    >
      <span
        className={`relative inline-flex h-5 w-9 shrink-0 items-center rounded-full transition-colors ${
          marcado ? 'bg-swnet-600' : 'bg-slate-300 dark:bg-slate-600'
        }`}
      >
        <input
          type="checkbox"
          checked={marcado}
          disabled={desabilitado}
          onChange={(e) => aoAlternar(e.target.checked)}
          className="peer sr-only"
        />
        <span
          className={`inline-block h-4 w-4 transform rounded-full bg-white shadow transition-transform ${
            marcado ? 'translate-x-4.5' : 'translate-x-0.5'
          }`}
        />
      </span>
      <span className="min-w-0">
        <span className="block text-sm font-medium text-slate-700 dark:text-slate-200">{rotulo}</span>
        {descricao && <span className="block text-xs text-slate-400">{descricao}</span>}
      </span>
    </label>
  )
}

interface OpcaoMultipla {
  valor: string
  rotulo: string
}

interface SelecaoMultiplaProps {
  rotulo: string
  opcoes: OpcaoMultipla[]
  selecionados: string[]
  aoMudar: (selecionados: string[]) => void
  desabilitado?: boolean
  colunas?: 2 | 3 | 4
  obrigatorio?: boolean
  erro?: string
}

/**
 * Substitui os `<select multiple>` (com select2) das telas legadas por uma
 * grade de checkboxes — mais legível e sem depender de biblioteca externa.
 */
export function SelecaoMultipla({
  rotulo,
  opcoes,
  selecionados,
  aoMudar,
  desabilitado,
  colunas = 3,
  obrigatorio,
  erro
}: SelecaoMultiplaProps) {
  const gradeColunas = { 2: 'sm:grid-cols-2', 3: 'sm:grid-cols-3', 4: 'sm:grid-cols-4' }[colunas]
  const todosMarcados = opcoes.length > 0 && selecionados.length === opcoes.length

  function alternar(valor: string) {
    aoMudar(
      selecionados.includes(valor)
        ? selecionados.filter((item) => item !== valor)
        : [...selecionados, valor]
    )
  }

  return (
    <Campo rotulo={rotulo} obrigatorio={obrigatorio} erro={erro}>
      <div className="mb-2 flex items-center justify-between">
        <span className="text-xs text-slate-400">{selecionados.length} selecionada(s)</span>
        <button
          type="button"
          disabled={desabilitado}
          onClick={() => aoMudar(todosMarcados ? [] : opcoes.map((o) => o.valor))}
          className="text-xs font-medium text-swnet-600 hover:underline disabled:opacity-50 dark:text-swnet-400"
        >
          {todosMarcados ? 'Limpar todas' : 'Marcar todas'}
        </button>
      </div>

      <div className={`grid gap-2 ${gradeColunas}`}>
        {opcoes.map((opcao) => (
          <label
            key={opcao.valor}
            className={`flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-xs
              text-slate-600 dark:border-slate-700 dark:text-slate-300 ${
                desabilitado ? 'cursor-not-allowed opacity-60' : 'cursor-pointer hover:border-swnet-300'
              }`}
          >
            <input
              type="checkbox"
              checked={selecionados.includes(opcao.valor)}
              disabled={desabilitado}
              onChange={() => alternar(opcao.valor)}
              className="h-4 w-4 shrink-0 accent-swnet-600"
            />
            <span className="truncate" title={opcao.rotulo}>
              {opcao.rotulo}
            </span>
          </label>
        ))}
      </div>
    </Campo>
  )
}
