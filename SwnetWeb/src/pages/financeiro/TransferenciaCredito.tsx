import { useState } from 'react'
import { ArrowRight, Coins, Search, Send } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { BuscaCliente } from '@/components/financeiro/BuscaCliente'
import {
  PLANTAS_PROTHEUS,
  TRANSFERENCIA_VAZIA,
  carregarCliente,
  saldoDisponivel,
  transferirCredito,
  validarTransferencia,
  type ClienteFinanceiro,
  type DadosTransferenciaCredito
} from '@/services/financeiro/financeiroApi'

/** Refatoração de Financeiro/Solicitar-Transferencia-De-Credito. */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

function ResumoCliente({ cliente, titulo }: { cliente: ClienteFinanceiro; titulo: string }) {
  return (
    <div className="rounded-lg border border-slate-200 p-3 dark:border-slate-700">
      <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
        {titulo}
      </p>
      <p className="truncate text-sm font-medium text-slate-800 dark:text-white" title={cliente.nome}>
        {cliente.nome}
      </p>
      <p className="text-xs text-slate-400">
        {cliente.codigo} / {cliente.loja} &bull; {cliente.municipio} - {cliente.uf}
      </p>
      <dl className="mt-2 grid grid-cols-3 gap-2 text-xs">
        {[
          ['Limite', moeda.format(cliente.limiteAtual)],
          ['Em aberto', moeda.format(cliente.limiteAberto)],
          ['Saldo', moeda.format(saldoDisponivel(cliente))]
        ].map(([rotulo, valor]) => (
          <div key={rotulo}>
            <dt className="text-slate-500 dark:text-slate-400">{rotulo}</dt>
            <dd className="font-medium text-slate-700 dark:text-slate-200">{valor}</dd>
          </div>
        ))}
      </dl>
    </div>
  )
}

export function TransferenciaCredito() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [dados, setDados] = useState<DadosTransferenciaCredito>(TRANSFERENCIA_VAZIA)
  const [origem, setOrigem] = useState<ClienteFinanceiro | null>(null)
  const [destino, setDestino] = useState<ClienteFinanceiro | null>(null)

  const [buscaAberta, setBuscaAberta] = useState<'origem' | 'destino' | null>(null)
  const [carregando, setCarregando] = useState(false)
  const [confirmando, setConfirmando] = useState(false)
  const [transferindo, setTransferindo] = useState(false)

  function mudar<C extends keyof DadosTransferenciaCredito>(campo: C, valor: string) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  async function buscarOrigem() {
    if (!dados.plantaOrigem || !dados.codClienteOrigem.trim()) {
      mostrarToast('Informe a planta e o cliente de origem.', 'warning')
      return
    }

    setCarregando(true)
    try {
      const encontrado = await carregarCliente(
        dados.plantaOrigem,
        dados.codClienteOrigem,
        dados.codLojaOrigem
      )
      if (!encontrado) {
        setOrigem(null)
        mostrarToast('Cliente de origem não encontrado.', 'warning')
        return
      }
      setOrigem(encontrado)
      setDados((atual) => ({ ...atual, codLojaOrigem: encontrado.loja }))
    } finally {
      setCarregando(false)
    }
  }

  function selecionarNaBusca(cliente: ClienteFinanceiro) {
    if (buscaAberta === 'origem') {
      setOrigem(cliente)
      setDados((atual) => ({
        ...atual,
        codClienteOrigem: cliente.codigo,
        codLojaOrigem: cliente.loja
      }))
    } else {
      setDestino(cliente)
      setDados((atual) => ({
        ...atual,
        codClienteDestino: cliente.codigo,
        codLojaDestino: cliente.loja
      }))
    }
  }

  async function validarEConfirmar() {
    const validacao = await validarTransferencia(dados)
    if (!validacao.sucesso) {
      mostrarToast(validacao.mensagem, 'warning')
      return
    }
    setConfirmando(true)
  }

  async function confirmar() {
    setConfirmando(false)
    setTransferindo(true)
    try {
      const resultado = await transferirCredito(dados)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setDados(TRANSFERENCIA_VAZIA)
        setOrigem(null)
        setDestino(null)
      }
    } finally {
      setTransferindo(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Transferência de Crédito" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Transferir Limite entre Clientes" icone={Coins}>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              validarEConfirmar()
            }}
          >
            <Secao titulo="Origem">
              <div className="grid items-end gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaOrigem"
                  rotulo="Planta de Origem"
                  required
                  value={dados.plantaOrigem}
                  disabled={transferindo}
                  onChange={(e) => {
                    mudar('plantaOrigem', e.target.value)
                    setOrigem(null)
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione a planta</option>
                  {PLANTAS_PROTHEUS.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Campo
                  rotulo="Cód. Cliente (Origem)"
                  htmlFor="codClienteOrigem"
                  obrigatorio
                  className="sm:col-span-3"
                >
                  <div className="flex">
                    <button
                      type="button"
                      onClick={() => setBuscaAberta('origem')}
                      disabled={!dados.plantaOrigem || transferindo}
                      aria-label="Pesquisar cliente de origem"
                      className="rounded-l-lg border border-r-0 border-slate-200 bg-slate-100 px-3 text-slate-500
                        hover:bg-slate-200 disabled:opacity-50 dark:border-slate-700 dark:bg-slate-800"
                    >
                      <Search size={15} />
                    </button>
                    <input
                      id="codClienteOrigem"
                      type="text"
                      value={dados.codClienteOrigem}
                      disabled={transferindo}
                      onChange={(e) => mudar('codClienteOrigem', e.target.value.trim())}
                      className={`${classesControle} rounded-l-none`}
                    />
                  </div>
                </Campo>

                <Entrada
                  id="codLojaOrigem"
                  rotulo="Cód. Loja (Origem)"
                  required
                  maxLength={2}
                  value={dados.codLojaOrigem}
                  disabled={transferindo}
                  onChange={(e) => mudar('codLojaOrigem', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-2"
                />

                <div className="sm:col-span-3">
                  <Botao icone={Search} aoClicar={buscarOrigem} desabilitado={carregando || transferindo}>
                    {carregando ? 'Buscando...' : 'Buscar'}
                  </Botao>
                </div>
              </div>
            </Secao>

            <Secao titulo="Destino">
              <div className="grid items-end gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaDestino"
                  rotulo="Planta de Destino"
                  required
                  value={dados.plantaDestino}
                  disabled={transferindo}
                  onChange={(e) => {
                    mudar('plantaDestino', e.target.value)
                    setDestino(null)
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione a planta</option>
                  {PLANTAS_PROTHEUS.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Campo rotulo="Cliente de Destino" obrigatorio className="sm:col-span-8">
                  <div className="flex">
                    <input
                      type="text"
                      readOnly
                      value={destino ? `${destino.codigo}/${destino.loja} — ${destino.nome}` : ''}
                      placeholder="Nenhum cliente selecionado"
                      className={`${classesControle} cursor-default rounded-r-none bg-slate-100 dark:bg-slate-800/50`}
                    />
                    <button
                      type="button"
                      onClick={() => setBuscaAberta('destino')}
                      disabled={!dados.plantaDestino || transferindo}
                      aria-label="Selecionar cliente de destino"
                      className="rounded-r-lg border border-l-0 border-slate-200 bg-swnet-600 px-3 text-white
                        hover:bg-swnet-700 disabled:opacity-50 dark:border-slate-700"
                    >
                      <Search size={15} />
                    </button>
                  </div>
                </Campo>
              </div>
            </Secao>

            {(origem || destino) && (
              <div className="mb-4 grid items-center gap-3 sm:grid-cols-[1fr_auto_1fr]">
                {origem ? (
                  <ResumoCliente cliente={origem} titulo="Origem" />
                ) : (
                  <div className="rounded-lg border border-dashed border-slate-200 p-3 text-center text-xs text-slate-400 dark:border-slate-700">
                    Selecione a origem
                  </div>
                )}

                <ArrowRight size={20} className="mx-auto shrink-0 text-slate-400" />

                {destino ? (
                  <ResumoCliente cliente={destino} titulo="Destino" />
                ) : (
                  <div className="rounded-lg border border-dashed border-slate-200 p-3 text-center text-xs text-slate-400 dark:border-slate-700">
                    Selecione o destino
                  </div>
                )}
              </div>
            )}

            <Secao titulo="Valor">
              <Entrada
                id="valorTransferencia"
                rotulo="Valor a transferir"
                required
                placeholder="0,00"
                value={dados.valorTransferencia}
                disabled={transferindo}
                ajuda={
                  origem
                    ? `Saldo disponível na origem: ${moeda.format(saldoDisponivel(origem))}`
                    : undefined
                }
                onChange={(e) => mudar('valorTransferencia', e.target.value)}
                larguraClasse="max-w-xs"
              />
            </Secao>

            <div className="mt-4 flex justify-end">
              <Botao tipo="submit" variante="sucesso" icone={Send} desabilitado={transferindo}>
                {transferindo ? 'Transferindo...' : 'Transferir crédito'}
              </Botao>
            </div>
          </form>
        </Painel>
      </div>

      <BuscaCliente
        aberto={buscaAberta !== null}
        planta={buscaAberta === 'origem' ? dados.plantaOrigem : dados.plantaDestino}
        titulo={buscaAberta === 'origem' ? 'Cliente de Origem' : 'Cliente de Destino'}
        mostrarLimites
        aoFechar={() => setBuscaAberta(null)}
        aoSelecionar={selecionarNaBusca}
      />

      <ConfirmDialog
        aberto={confirmando}
        titulo="Confirmar transferência"
        mensagem={
          origem && destino
            ? `Serão transferidos ${dados.valorTransferencia} de ${origem.nome} para ${destino.nome}. Deseja continuar?`
            : ''
        }
        rotuloConfirmar="Transferir"
        aoConfirmar={confirmar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
