import { useEffect, useState } from 'react'
import { AlertTriangle, ClipboardCheck, History, Save } from 'lucide-react'
import { Modal } from '@/components/comuns/Modal'
import { Botao, Etiqueta } from '@/components/comuns/Comuns'
import { Campo } from '@/components/comuns/Campos'
import { Tabela } from '@/components/comuns/Tabela'
import {
  MOTIVOS_BLOQUEIO,
  alertasDaLiberacao,
  ultimosPedidosDoCliente,
  type OrigemPedidos,
  type PedidoValidacao,
  type StatusSelecionavel,
  type StatusValidacao
} from '@/services/validacoes/validacoesApi'
import { formatarData } from '@/services/comuns/baseMock'

/**
 * Detalhe de validação do pedido — equivale ao ValidaPedido.asp: cabeçalho
 * do pedido, quadro de últimos pedidos do cliente e os radios de status.
 */

const numero = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 })
const decimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 4, maximumFractionDigits: 4 })

interface ModalValidacaoPedidoProps {
  pedido: PedidoValidacao | null
  origem: OrigemPedidos
  /** Opções de status oferecidas (papel tem menos que pedido normal) */
  opcoes: StatusSelecionavel[]
  salvando: boolean
  aoFechar: () => void
  aoConfirmar: (status: StatusValidacao) => void
}

export function ModalValidacaoPedido({
  pedido,
  origem,
  opcoes,
  salvando,
  aoFechar,
  aoConfirmar
}: ModalValidacaoPedidoProps) {
  const [status, setStatus] = useState<StatusValidacao>('PENDENTE')
  const [ultimos, setUltimos] = useState<PedidoValidacao[]>([])

  useEffect(() => {
    if (!pedido) return

    // O legado já marca "Bloqueado por preço" quando o preço caiu
    const sugestao: StatusSelecionavel = pedido.preco < pedido.ultimoPreco ? 'PR' : 'LI'
    setStatus(opcoes.includes(sugestao) ? sugestao : 'PENDENTE')

    ultimosPedidosDoCliente(origem, pedido.cliente, pedido.id).then(setUltimos)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pedido?.id])

  if (!pedido) return null

  const alertas = alertasDaLiberacao(pedido)
  const bloqueadoPorTes = pedido.tesLivreDebito && status === 'LI'

  return (
    <Modal
      aberto
      titulo={`Pedido ${pedido.pedido}`}
      icone={ClipboardCheck}
      largura="xl"
      aoFechar={aoFechar}
      rodape={
        <>
          <Botao variante="neutro" aoClicar={aoFechar} desabilitado={salvando}>
            Cancelar
          </Botao>
          <Botao
            variante="sucesso"
            icone={Save}
            aoClicar={() => aoConfirmar(status)}
            desabilitado={salvando || status === 'PENDENTE' || bloqueadoPorTes}
          >
            {salvando ? 'Gravando...' : 'Gravar validação'}
          </Botao>
        </>
      }
    >
      <div className="grid gap-4">
        <p className="text-sm text-slate-500 dark:text-slate-400">
          Emitido em {formatarData(pedido.emissao)} por <strong>{pedido.quemEmitiu}</strong>
        </p>

        {alertas.map((alerta) => (
          <p
            key={alerta.tipo}
            className={`flex items-start gap-2 rounded-lg px-3 py-2 text-sm ${
              alerta.tipo === 'tes'
                ? 'bg-red-50 text-red-700 dark:bg-red-500/10 dark:text-red-400'
                : 'bg-amber-50 text-amber-800 dark:bg-amber-500/10 dark:text-amber-400'
            }`}
          >
            <AlertTriangle size={15} className="mt-0.5 shrink-0" />
            {alerta.mensagem}
          </p>
        ))}

        <dl className="grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-4 dark:border-slate-700">
          {[
            ['Cliente', `${pedido.cliente}/${pedido.loja} — ${pedido.nome}`],
            ['Local de Entrega', pedido.localEntrega],
            ['Produto', `${pedido.produto} — ${pedido.descricaoProduto}`],
            ['FT', pedido.ft],
            ['Quantidade', numero.format(pedido.quantidade)],
            ['Preço', decimal.format(pedido.preco)],
            ['Preço anterior', pedido.ultimoPreco > 0 ? decimal.format(pedido.ultimoPreco) : '—'],
            ['Preço/Kg Net', decimal.format(pedido.precoKgNet)],
            ['M²', numero.format(pedido.m2)],
            ['Peso total', `${numero.format(pedido.pesoTotal)} kg`],
            ['Entrega', formatarData(pedido.entrega)],
            ['Cond. Pagamento', pedido.condicaoPagamento]
          ].map(([rotulo, valor]) => (
            <div key={rotulo}>
              <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
              <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                {valor}
              </dd>
            </div>
          ))}
        </dl>

        {pedido.observacao && (
          <p className="rounded-lg bg-slate-50 px-3 py-2 text-sm text-slate-600 dark:bg-slate-800 dark:text-slate-300">
            <strong>Obs.:</strong> {pedido.observacao}
          </p>
        )}

        <div>
          <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
            <History size={13} /> Últimos pedidos do cliente
          </p>
          <div className="rounded-lg border border-slate-200 dark:border-slate-700">
            <Tabela
              itens={ultimos}
              chave={(item) => item.id}
              mensagemVazia="Sem pedidos anteriores para este cliente"
              colunas={[
                { cabecalho: 'Pedido', conteudo: (p) => p.pedido },
                { cabecalho: 'Referência', conteudo: (p) => p.descricaoProduto },
                { cabecalho: 'Emissão', conteudo: (p) => formatarData(p.emissao) },
                { cabecalho: 'Entrega', conteudo: (p) => formatarData(p.entrega) },
                {
                  cabecalho: 'Quantidade',
                  conteudo: (p) => numero.format(p.quantidade),
                  alinhamento: 'direita'
                },
                { cabecalho: 'Preço', conteudo: (p) => decimal.format(p.preco), alinhamento: 'direita' },
                {
                  cabecalho: 'Preço/Kg',
                  conteudo: (p) => decimal.format(p.precoKg),
                  alinhamento: 'direita'
                },
                { cabecalho: 'Cond. Pagto.', conteudo: (p) => p.condicaoPagamento }
              ]}
            />
          </div>
        </div>

        <Campo rotulo="Validação" obrigatorio>
          <div className="grid gap-2 sm:grid-cols-2">
            {opcoes.map((opcao) => (
              <label
                key={opcao}
                className={`flex cursor-pointer items-center gap-2 rounded-lg border px-3 py-2 text-sm
                  ${
                    status === opcao
                      ? 'border-swnet-400 bg-swnet-50 text-swnet-800 dark:border-swnet-600 dark:bg-swnet-950/40 dark:text-swnet-300'
                      : 'border-slate-200 text-slate-600 hover:border-swnet-300 dark:border-slate-700 dark:text-slate-300'
                  }`}
              >
                <input
                  type="radio"
                  name="statusValidacao"
                  checked={status === opcao}
                  disabled={salvando}
                  onChange={() => setStatus(opcao)}
                  className="h-4 w-4 accent-swnet-600"
                />
                <span className="flex-1">{MOTIVOS_BLOQUEIO[opcao].rotulo}</span>
                {MOTIVOS_BLOQUEIO[opcao].codigo && (
                  <Etiqueta texto={MOTIVOS_BLOQUEIO[opcao].codigo} tom="neutro" />
                )}
              </label>
            ))}
          </div>
        </Campo>

        {bloqueadoPorTes && (
          <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
            Este pedido não pode ser liberado aqui: a TES exige validação do Gerente Comercial.
          </p>
        )}
      </div>
    </Modal>
  )
}
