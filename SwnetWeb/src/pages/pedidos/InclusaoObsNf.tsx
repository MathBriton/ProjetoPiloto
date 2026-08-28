import { useState } from 'react'
import { FileText, Save, Search } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import { buscarObservacaoNf, salvarObservacaoNf } from '@/services/pedidos/pedidosApi'

/**
 * Refatoração de Pedidos/Inclusao-OBS-NF (form_incluiobsnf.asp).
 *
 * Busca o pedido e grava a observação que sairá impressa na nota fiscal.
 */
export function InclusaoObsNf() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [pedido, setPedido] = useState('')
  const [nomeCliente, setNomeCliente] = useState('')
  const [observacao, setObservacao] = useState('')
  const [encontrado, setEncontrado] = useState(false)
  const [buscando, setBuscando] = useState(false)
  const [salvando, setSalvando] = useState(false)

  async function buscar() {
    if (!pedido.trim()) {
      mostrarToast('Informe o número do pedido.', 'warning')
      return
    }

    setBuscando(true)
    try {
      const resultado = await buscarObservacaoNf(pedido)
      setEncontrado(resultado.encontrado)
      setNomeCliente(resultado.nomeCliente)
      setObservacao(resultado.observacao)

      if (!resultado.encontrado) mostrarToast('Pedido não encontrado.', 'warning')
    } finally {
      setBuscando(false)
    }
  }

  async function salvar() {
    setSalvando(true)
    try {
      const resultado = await salvarObservacaoNf(pedido, observacao)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Inclusão de Observação em NF" />

      <div className="mx-auto max-w-2xl">
        <Painel titulo="Observação da Nota Fiscal" icone={FileText}>
          <form
            className="grid gap-4"
            onSubmit={(e) => {
              e.preventDefault()
              buscar()
            }}
          >
            <div className="flex items-end gap-2">
              <Entrada
                id="numeroPedidoNf"
                rotulo="Número do Pedido"
                required
                value={pedido}
                disabled={buscando || salvando}
                onChange={(e) => {
                  setPedido(e.target.value.replace(/\D/g, ''))
                  setEncontrado(false)
                }}
                larguraClasse="flex-1"
              />
              <Botao tipo="submit" icone={Search} desabilitado={buscando || salvando}>
                {buscando ? 'Buscando...' : 'Buscar'}
              </Botao>
            </div>
          </form>

          {encontrado && (
            <div className="mt-4 grid gap-4 border-t border-slate-200 pt-4 dark:border-slate-800">
              <Entrada id="clienteNf" rotulo="Cliente" value={nomeCliente} disabled />

              <Campo rotulo="Observação" htmlFor="observacaoNf">
                <textarea
                  id="observacaoNf"
                  rows={5}
                  value={observacao}
                  disabled={salvando}
                  placeholder="Texto que será impresso na nota fiscal..."
                  onChange={(e) => setObservacao(e.target.value)}
                  className={classesControle}
                />
              </Campo>

              <div className="flex justify-end">
                <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
                  {salvando ? 'Gravando...' : 'Gravar observação'}
                </Botao>
              </div>
            </div>
          )}
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
