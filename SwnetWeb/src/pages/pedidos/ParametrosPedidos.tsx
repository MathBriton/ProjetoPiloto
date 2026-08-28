import { useCallback, useEffect, useState } from 'react'
import { Save, Settings2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Selecao, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import {
  PLANTAS_POR_REGIAO,
  listarParametros,
  salvarParametros,
  type ParametroPedido
} from '@/services/pedidos/pedidosApi'

/**
 * Refatoração de Pedidos/Parametros-Pedidos.
 *
 * Define as faixas mínima e máxima por tipo de produto usadas na crítica
 * da emissão de pedido. O legado começa pela seleção de planta
 * (SelecionaPlantaPedidoParametros.asp).
 */

const TODAS_PLANTAS = [...PLANTAS_POR_REGIAO.PIRAUBE, ...PLANTAS_POR_REGIAO.BRASIL]

export function ParametrosPedidos() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [planta, setPlanta] = useState('')
  const [parametros, setParametros] = useState<ParametroPedido[]>([])
  const [carregando, setCarregando] = useState(false)
  const [salvando, setSalvando] = useState(false)

  const carregar = useCallback(async () => {
    if (!planta) {
      setParametros([])
      return
    }

    setCarregando(true)
    try {
      setParametros(await listarParametros(planta))
    } finally {
      setCarregando(false)
    }
  }, [planta])

  useEffect(() => {
    carregar()
  }, [carregar])

  function mudar(tipoProduto: string, campo: 'minimo' | 'maximo', valor: string) {
    setParametros((atual) =>
      atual.map((item) => (item.tipoProduto === tipoProduto ? { ...item, [campo]: valor } : item))
    )
  }

  async function salvar() {
    setSalvando(true)
    try {
      const resultado = await salvarParametros(planta, parametros)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Parâmetros de Pedidos" />

      <div className="mx-auto max-w-3xl space-y-4">
        <Painel titulo="Planta" icone={Settings2}>
          <Selecao
            id="plantaParametros"
            rotulo="Selecione a planta"
            required
            value={planta}
            onChange={(e) => setPlanta(e.target.value)}
            larguraClasse="max-w-sm"
          >
            <option value="">Selecione...</option>
            {TODAS_PLANTAS.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.nome}
              </option>
            ))}
          </Selecao>
        </Painel>

        {planta && (
          <Painel
            titulo="Faixas por tipo de produto"
            icone={Settings2}
            acoes={
              <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando || carregando}>
                {salvando ? 'Salvando...' : 'Salvar'}
              </Botao>
            }
          >
            {carregando ? (
              <p className="py-8 text-center text-sm text-slate-400">Carregando parâmetros...</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-500 dark:border-slate-800 dark:text-slate-400">
                    <tr>
                      <th className="px-3 py-2 font-semibold">Tipo de Produto</th>
                      <th className="px-3 py-2 font-semibold">Mínimo</th>
                      <th className="px-3 py-2 font-semibold">Máximo</th>
                    </tr>
                  </thead>
                  <tbody>
                    {parametros.map((item) => (
                      <tr
                        key={item.tipoProduto}
                        className="border-b border-slate-100 last:border-0 dark:border-slate-800"
                      >
                        <td className="px-3 py-2 font-medium text-slate-700 dark:text-slate-200">
                          {item.tipoProduto}
                        </td>
                        <td className="px-3 py-2">
                          <input
                            type="number"
                            min="0"
                            value={item.minimo}
                            disabled={salvando}
                            aria-label={`Mínimo de ${item.tipoProduto}`}
                            onChange={(e) => mudar(item.tipoProduto, 'minimo', e.target.value)}
                            className={`${classesControle} max-w-40`}
                          />
                        </td>
                        <td className="px-3 py-2">
                          <input
                            type="number"
                            min="0"
                            value={item.maximo}
                            disabled={salvando}
                            aria-label={`Máximo de ${item.tipoProduto}`}
                            onChange={(e) => mudar(item.tipoProduto, 'maximo', e.target.value)}
                            className={`${classesControle} max-w-40`}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <p className="mt-3 text-xs text-slate-400">
              Estas faixas são aplicadas como crítica na emissão de pedido para a planta selecionada.
            </p>
          </Painel>
        )}
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
