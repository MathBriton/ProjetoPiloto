import type { ReactNode } from 'react'
import { ArrowDown, ArrowUp, ArrowUpDown } from 'lucide-react'

/**
 * Tabela de listagem usada pelas grades refatoradas dos ASP.
 *
 * Todas as telas de "Grade" do legado repetem a mesma estrutura: cabeçalho,
 * linhas zebradas, coluna de ações à direita e o aviso "Nenhum registro
 * encontrado" quando vazia.
 */

export interface ColunaTabela<T> {
  cabecalho: ReactNode
  conteudo: (item: T) => ReactNode
  /** Chave de ordenação — quando informada, o cabeçalho vira botão */
  ordenavelPor?: string
  className?: string
  alinhamento?: 'esquerda' | 'centro' | 'direita'
}

export interface OrdenacaoTabela {
  campo: string
  ascendente: boolean
}

interface TabelaProps<T> {
  colunas: ColunaTabela<T>[]
  itens: T[]
  chave: (item: T) => string | number
  carregando?: boolean
  mensagemVazia?: string
  mensagemCarregando?: string
  /** Renderiza a coluna de ações (editar/excluir) no fim de cada linha */
  acoes?: (item: T) => ReactNode
  rotuloAcoes?: string
  ordenacao?: OrdenacaoTabela
  aoOrdenar?: (campo: string) => void
}

const ALINHAMENTOS = {
  esquerda: 'text-left',
  centro: 'text-center',
  direita: 'text-right'
}

export function Tabela<T>({
  colunas,
  itens,
  chave,
  carregando = false,
  mensagemVazia = 'Nenhum registro encontrado',
  mensagemCarregando = 'Carregando...',
  acoes,
  rotuloAcoes = 'Ações',
  ordenacao,
  aoOrdenar
}: TabelaProps<T>) {
  const totalColunas = colunas.length + (acoes ? 1 : 0)

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-left text-sm">
        <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
          <tr>
            {colunas.map((coluna, indice) => {
              const alinhamento = ALINHAMENTOS[coluna.alinhamento ?? 'esquerda']
              const ordenavel = coluna.ordenavelPor && aoOrdenar
              const ativa = !!ordenacao && ordenacao.campo === coluna.ordenavelPor
              const Icone = !ativa ? ArrowUpDown : ordenacao!.ascendente ? ArrowUp : ArrowDown

              return (
                <th
                  key={indice}
                  className={`px-3 py-2 font-semibold ${alinhamento} ${coluna.className ?? ''}`}
                >
                  {ordenavel ? (
                    <button
                      type="button"
                      onClick={() => aoOrdenar(coluna.ordenavelPor!)}
                      title="Clique para ordenar"
                      className="inline-flex items-center gap-1 uppercase hover:text-slate-700 dark:hover:text-slate-200"
                    >
                      {coluna.cabecalho}
                      <Icone size={12} className={ativa ? 'text-swnet-600 dark:text-swnet-400' : ''} />
                    </button>
                  ) : (
                    coluna.cabecalho
                  )}
                </th>
              )
            })}
            {acoes && <th className="px-3 py-2 text-right font-semibold">{rotuloAcoes}</th>}
          </tr>
        </thead>

        <tbody>
          {carregando && (
            <tr>
              <td colSpan={totalColunas} className="px-3 py-10 text-center text-sm text-slate-400">
                {mensagemCarregando}
              </td>
            </tr>
          )}

          {!carregando && itens.length === 0 && (
            <tr>
              <td colSpan={totalColunas} className="px-3 py-10 text-center text-sm text-slate-400">
                {mensagemVazia}
              </td>
            </tr>
          )}

          {!carregando &&
            itens.map((item) => (
              <tr
                key={chave(item)}
                className="border-b border-slate-100 last:border-0 hover:bg-slate-50/60
                  dark:border-slate-800 dark:hover:bg-slate-800/40"
              >
                {colunas.map((coluna, indice) => (
                  <td
                    key={indice}
                    className={`px-3 py-2 text-slate-600 dark:text-slate-300 ${
                      ALINHAMENTOS[coluna.alinhamento ?? 'esquerda']
                    }`}
                  >
                    {coluna.conteudo(item)}
                  </td>
                ))}
                {acoes && (
                  <td className="px-3 py-2">
                    <div className="flex items-center justify-end gap-1">{acoes(item)}</div>
                  </td>
                )}
              </tr>
            ))}
        </tbody>
      </table>
    </div>
  )
}

interface BotaoAcaoProps {
  icone: React.ComponentType<{ size?: number }>
  titulo: string
  aoClicar: () => void
  tom?: 'neutro' | 'perigo'
}

/** Botãozinho de ícone das colunas de ação (editar / excluir) */
export function BotaoAcao({ icone: Icone, titulo, aoClicar, tom = 'neutro' }: BotaoAcaoProps) {
  const cores =
    tom === 'perigo'
      ? 'hover:bg-red-50 hover:text-red-600 dark:hover:bg-slate-800'
      : 'hover:bg-slate-100 hover:text-swnet-600 dark:hover:bg-slate-800'

  return (
    <button
      type="button"
      onClick={aoClicar}
      title={titulo}
      aria-label={titulo}
      className={`rounded-lg p-1.5 text-slate-400 transition-colors ${cores}`}
    >
      <Icone size={14} />
    </button>
  )
}

interface PaginacaoProps {
  pagina: number
  totalPaginas: number
  totalRegistros: number
  tamanho: number
  aoMudarPagina: (pagina: number) => void
  aoMudarTamanho: (tamanho: number) => void
}

/** Rodapé "Exibir N registros" + navegação, presente nas grades do Cavaco/Aparas */
export function Paginacao({
  pagina,
  totalPaginas,
  totalRegistros,
  tamanho,
  aoMudarPagina,
  aoMudarTamanho
}: PaginacaoProps) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-200 px-3 py-2 dark:border-slate-800">
      <label className="flex items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
        Exibir
        <select
          value={tamanho}
          onChange={(e) => aoMudarTamanho(Number(e.target.value))}
          className="rounded-lg border border-slate-200 bg-slate-50 px-2 py-1 text-xs
            dark:border-slate-700 dark:bg-slate-800 dark:text-white"
        >
          {[10, 25, 50, 100, 200].map((opcao) => (
            <option key={opcao} value={opcao}>
              {opcao}
            </option>
          ))}
        </select>
        registros
      </label>

      <span className="text-xs text-slate-400">
        {totalRegistros === 0
          ? 'Nenhum registro'
          : `Página ${pagina} de ${totalPaginas} — ${totalRegistros} registro(s)`}
      </span>

      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={() => aoMudarPagina(pagina - 1)}
          disabled={pagina <= 1}
          className="rounded-lg border border-slate-200 px-2.5 py-1 text-xs text-slate-600 disabled:opacity-40
            hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
        >
          Anterior
        </button>
        <button
          type="button"
          onClick={() => aoMudarPagina(pagina + 1)}
          disabled={pagina >= totalPaginas}
          className="rounded-lg border border-slate-200 px-2.5 py-1 text-xs text-slate-600 disabled:opacity-40
            hover:bg-slate-50 dark:border-slate-700 dark:text-slate-300 dark:hover:bg-slate-800"
        >
          Próxima
        </button>
      </div>
    </div>
  )
}
