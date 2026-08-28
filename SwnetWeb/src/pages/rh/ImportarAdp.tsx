import { useState } from 'react'
import { AlertTriangle, CheckCircle2, FileSpreadsheet, Upload } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import {
  PLANTAS_ADP,
  importarArquivoAdp,
  type ResultadoImportacao
} from '@/services/rh/rhApi'

/**
 * Refatoração de RH/Importar-Do-ADP
 * (SelecionaPlanta.asp + importADPv3.asp + as rotinas de processamento).
 *
 * O legado recebe dois CSVs — Funcionarios.csv e AcessosFuncionarios.csv —
 * e roda em seguida as rotinas de matrículas e jornadas.
 */

type TipoArquivo = 'funcionarios' | 'acessos'

const ARQUIVOS: { tipo: TipoArquivo; rotulo: string; esperado: string }[] = [
  { tipo: 'funcionarios', rotulo: 'Funcionários', esperado: 'Funcionarios.csv' },
  { tipo: 'acessos', rotulo: 'Acessos de Funcionários', esperado: 'AcessosFuncionarios.csv' }
]

export function ImportarAdp() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [planta, setPlanta] = useState('')
  const [arquivos, setArquivos] = useState<Record<TipoArquivo, string>>({
    funcionarios: '',
    acessos: ''
  })
  const [resultados, setResultados] = useState<Partial<Record<TipoArquivo, ResultadoImportacao>>>({})
  const [importando, setImportando] = useState<TipoArquivo | null>(null)

  async function importar(tipo: TipoArquivo) {
    setImportando(tipo)
    try {
      const resultado = await importarArquivoAdp(planta, tipo, arquivos[tipo])
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
      setResultados((atual) => ({ ...atual, [tipo]: resultado }))
    } finally {
      setImportando(null)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Importar do ADP" />

      <div className="mx-auto max-w-3xl space-y-4">
        <Painel titulo="Planta" icone={Upload}>
          <Selecao
            id="plantaAdp"
            rotulo="Selecione a planta da importação"
            required
            value={planta}
            onChange={(e) => setPlanta(e.target.value)}
            larguraClasse="max-w-sm"
          >
            <option value="">Selecione...</option>
            {PLANTAS_ADP.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.nome}
              </option>
            ))}
          </Selecao>
        </Painel>

        <Painel titulo="Funcionários e Acessos de Funcionários" icone={FileSpreadsheet}>
          {!planta && (
            <p className="mb-4 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800 dark:bg-amber-500/10 dark:text-amber-400">
              Escolha a planta antes de importar os arquivos.
            </p>
          )}

          {ARQUIVOS.map((arquivo) => {
            const resultado = resultados[arquivo.tipo]

            return (
              <Secao key={arquivo.tipo} titulo={arquivo.rotulo}>
                <div className="flex flex-wrap items-end gap-3">
                  <Entrada
                    id={`arquivo-${arquivo.tipo}`}
                    rotulo={`Arquivo ${arquivo.esperado}`}
                    placeholder={arquivo.esperado}
                    value={arquivos[arquivo.tipo]}
                    disabled={!planta || importando !== null}
                    ajuda="O upload real depende do backend; informe o nome do arquivo."
                    onChange={(e) =>
                      setArquivos((atual) => ({ ...atual, [arquivo.tipo]: e.target.value }))
                    }
                    larguraClasse="min-w-64 flex-1"
                  />
                  <Botao
                    icone={Upload}
                    aoClicar={() => importar(arquivo.tipo)}
                    desabilitado={!planta || importando !== null}
                  >
                    {importando === arquivo.tipo ? 'Importando...' : 'Importar'}
                  </Botao>
                </div>

                {resultado && (
                  <div
                    className={`mt-3 rounded-lg border p-3 ${
                      resultado.sucesso
                        ? 'border-emerald-200 bg-emerald-50 dark:border-emerald-900 dark:bg-emerald-500/10'
                        : 'border-red-200 bg-red-50 dark:border-red-900 dark:bg-red-500/10'
                    }`}
                  >
                    <p
                      className={`flex items-center gap-2 text-sm font-medium ${
                        resultado.sucesso
                          ? 'text-emerald-800 dark:text-emerald-300'
                          : 'text-red-800 dark:text-red-300'
                      }`}
                    >
                      {resultado.sucesso ? <CheckCircle2 size={15} /> : <AlertTriangle size={15} />}
                      {resultado.mensagem}
                    </p>

                    {resultado.sucesso && (
                      <p className="mt-1 text-xs text-emerald-700 dark:text-emerald-400">
                        {resultado.linhasImportadas} de {resultado.linhasLidas} linha(s) importada(s).
                      </p>
                    )}

                    {resultado.erros.length > 0 && (
                      <ul className="mt-2 grid gap-1">
                        {resultado.erros.map((mensagem) => (
                          <li
                            key={mensagem}
                            className="text-xs text-amber-700 dark:text-amber-400"
                          >
                            • {mensagem}
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                )}
              </Secao>
            )
          })}

          <p className="mt-2 text-xs text-slate-400">
            Após as duas importações, o legado ainda executa as rotinas de matrículas e jornadas — elas
            dependem do backend e não têm equivalente no frontend.
          </p>
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
