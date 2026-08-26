import { useEffect, useState } from 'react'
import { Search, Users } from 'lucide-react'
import { Modal } from '@/components/comuns/Modal'
import { Botao } from '@/components/comuns/Comuns'
import { Entrada } from '@/components/comuns/Campos'
import { Tabela } from '@/components/comuns/Tabela'
import {
  buscarClientesFinanceiro,
  saldoDisponivel,
  type ClienteFinanceiro
} from '@/services/financeiro/financeiroApi'

/**
 * Modal "Buscar Cliente" — no legado aparece igual no alterar_risco.asp e no
 * TransferenciaCreditoFormulario.asp.
 */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

interface BuscaClienteProps {
  aberto: boolean
  planta: string
  aoFechar: () => void
  aoSelecionar: (cliente: ClienteFinanceiro) => void
  /** Mostra limite/saldo nas colunas — usado na transferência de crédito */
  mostrarLimites?: boolean
  titulo?: string
}

export function BuscaCliente({
  aberto,
  planta,
  aoFechar,
  aoSelecionar,
  mostrarLimites = false,
  titulo = 'Buscar Cliente'
}: BuscaClienteProps) {
  const [termo, setTermo] = useState('')
  const [resultados, setResultados] = useState<ClienteFinanceiro[]>([])
  const [buscando, setBuscando] = useState(false)

  useEffect(() => {
    if (!aberto) return

    setTermo('')
    setBuscando(true)
    buscarClientesFinanceiro(planta, '')
      .then(setResultados)
      .finally(() => setBuscando(false))
  }, [aberto, planta])

  async function pesquisar() {
    setBuscando(true)
    try {
      setResultados(await buscarClientesFinanceiro(planta, termo))
    } finally {
      setBuscando(false)
    }
  }

  return (
    <Modal aberto={aberto} titulo={titulo} icone={Users} largura="xl" aoFechar={aoFechar}>
      <div className="grid gap-4">
        <div className="flex items-end gap-2">
          <Entrada
            id="termoCliente"
            rotulo="Nome ou código do cliente"
            value={termo}
            disabled={buscando}
            onChange={(e) => setTermo(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault()
                pesquisar()
              }
            }}
            larguraClasse="flex-1"
          />
          <Botao icone={Search} aoClicar={pesquisar} desabilitado={buscando}>
            {buscando ? '...' : 'Buscar'}
          </Botao>
        </div>

        <Tabela
          itens={resultados}
          chave={(cliente) => `${cliente.codigo}-${cliente.loja}`}
          carregando={buscando}
          mensagemVazia="Nenhum cliente encontrado"
          colunas={[
            { cabecalho: 'Código', conteudo: (c) => c.codigo },
            { cabecalho: 'Loja', conteudo: (c) => c.loja },
            { cabecalho: 'Nome', conteudo: (c) => c.nome },
            { cabecalho: 'CNPJ', conteudo: (c) => c.cnpj },
            { cabecalho: 'Município', conteudo: (c) => c.municipio },
            { cabecalho: 'UF', conteudo: (c) => c.uf },
            ...(mostrarLimites
              ? ([
                  {
                    cabecalho: 'Limite Atual',
                    conteudo: (c: ClienteFinanceiro) => moeda.format(c.limiteAtual),
                    alinhamento: 'direita' as const
                  },
                  {
                    cabecalho: 'Em Aberto',
                    conteudo: (c: ClienteFinanceiro) => moeda.format(c.limiteAberto),
                    alinhamento: 'direita' as const
                  },
                  {
                    cabecalho: 'Saldo',
                    conteudo: (c: ClienteFinanceiro) => moeda.format(saldoDisponivel(c)),
                    alinhamento: 'direita' as const
                  }
                ])
              : [])
          ]}
          acoes={(cliente) => (
            <button
              type="button"
              onClick={() => {
                aoSelecionar(cliente)
                aoFechar()
              }}
              className="rounded-lg px-2 py-1 text-xs font-medium text-swnet-600 hover:bg-swnet-50
                dark:text-swnet-400 dark:hover:bg-slate-800"
            >
              Selecionar
            </button>
          )}
          rotuloAcoes=""
        />
      </div>
    </Modal>
  )
}
