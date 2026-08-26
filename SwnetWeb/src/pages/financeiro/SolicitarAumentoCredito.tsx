import { useState } from 'react'
import { Coins, Paperclip, Search, Send, TrendingUp, Users } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Chave, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import {
  PLANTAS_CREDITO,
  SOLICITACAO_CREDITO_VAZIA,
  carregarCliente,
  saldoDisponivel,
  solicitarAumentoCredito,
  type ClienteFinanceiro,
  type DadosSolicitacaoCredito
} from '@/services/financeiro/financeiroApi'

/** Refatoração de Financeiro/Solicitar-Aumento-De-Credito. */

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

/** Planta 13 (Florestal) troca os campos de preço pelos de tora/semente/muda */
const PLANTA_FLORESTAL = '13'

export function SolicitarAumentoCredito() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [dados, setDados] = useState<DadosSolicitacaoCredito>(SOLICITACAO_CREDITO_VAZIA)
  const [cliente, setCliente] = useState<ClienteFinanceiro | null>(null)
  const [carregando, setCarregando] = useState(false)
  const [enviando, setEnviando] = useState(false)

  const florestal = dados.planta === PLANTA_FLORESTAL

  function mudar<C extends keyof DadosSolicitacaoCredito>(campo: C, valor: DadosSolicitacaoCredito[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  async function buscarCliente() {
    if (!dados.planta || !dados.codCliente.trim()) {
      mostrarToast('Informe a planta e o código do cliente.', 'warning')
      return
    }

    setCarregando(true)
    try {
      const encontrado = await carregarCliente(dados.planta, dados.codCliente, dados.codLoja)
      if (!encontrado) {
        setCliente(null)
        mostrarToast('Cliente não encontrado nesta planta.', 'warning')
        return
      }
      setCliente(encontrado)
    } finally {
      setCarregando(false)
    }
  }

  async function enviar() {
    setEnviando(true)
    try {
      const resultado = await solicitarAumentoCredito(dados)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setDados(SOLICITACAO_CREDITO_VAZIA)
        setCliente(null)
      }
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Solicitar Aumento de Crédito" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Solicitação de Crédito" icone={Coins}>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              enviar()
            }}
          >
            <Secao titulo="Cliente" icone={Users}>
              <div className="grid items-end gap-4 sm:grid-cols-12">
                <Selecao
                  id="planta"
                  rotulo="Planta"
                  required
                  value={dados.planta}
                  disabled={enviando}
                  onChange={(e) => {
                    mudar('planta', e.target.value)
                    setCliente(null)
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione a planta</option>
                  {PLANTAS_CREDITO.map((item) => (
                    <option key={item.codigo} value={item.codigo}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="codCliente"
                  rotulo="Cód. Cliente"
                  required
                  value={dados.codCliente}
                  disabled={enviando}
                  onChange={(e) => mudar('codCliente', e.target.value.trim())}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="codLoja"
                  rotulo="Cód. Loja"
                  required
                  maxLength={2}
                  value={dados.codLoja}
                  disabled={enviando}
                  onChange={(e) => mudar('codLoja', e.target.value.replace(/\D/g, ''))}
                  larguraClasse="sm:col-span-2"
                />

                <div className="sm:col-span-3">
                  <Botao icone={Search} aoClicar={buscarCliente} desabilitado={carregando || enviando}>
                    {carregando ? 'Buscando...' : 'Buscar'}
                  </Botao>
                </div>
              </div>

              {cliente && (
                <dl className="mt-4 grid gap-3 rounded-lg border border-slate-200 p-3 sm:grid-cols-3 dark:border-slate-700">
                  {[
                    ['Nome do cliente', cliente.nome],
                    ['Limite Atual', moeda.format(cliente.limiteAtual)],
                    ['Limite em aberto', moeda.format(cliente.limiteAberto)],
                    ['Saldo', moeda.format(saldoDisponivel(cliente))],
                    ['Código Cliente Principal', cliente.codigoPrincipal],
                    ['Nome Cliente Principal', cliente.nomePrincipal]
                  ].map(([rotulo, valor]) => (
                    <div key={rotulo}>
                      <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                      <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                        {valor}
                      </dd>
                    </div>
                  ))}
                </dl>
              )}
            </Secao>

            <Secao titulo="Potencial de Vendas" icone={TrendingUp}>
              <div className="mb-4">
                <Chave
                  rotulo="Outras vendas (Resíduos/Aparas)"
                  descricao="Dispensa o envio de Contrato Social e Balanço/DRE."
                  marcado={dados.outrasVendas}
                  desabilitado={enviando}
                  aoAlternar={(v) => mudar('outrasVendas', v)}
                />
              </div>

              <div className="grid gap-4 sm:grid-cols-12">
                {florestal ? (
                  <>
                    <Entrada
                      id="valorToneladaTora"
                      rotulo="Valor por tonelada de tora"
                      value={dados.precoM2}
                      disabled={enviando}
                      onChange={(e) => mudar('precoM2', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                    <Entrada
                      id="valorUnitarioSemente"
                      rotulo="Valor unitário da semente"
                      value={dados.precoKg}
                      disabled={enviando}
                      onChange={(e) => mudar('precoKg', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                    <Entrada
                      id="valorUnitarioMuda"
                      rotulo="Valor unitário da muda"
                      value={dados.volumeM2}
                      disabled={enviando}
                      onChange={(e) => mudar('volumeM2', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                  </>
                ) : (
                  <>
                    <Entrada
                      id="precoM2"
                      rotulo="Preço médio por M²"
                      value={dados.precoM2}
                      disabled={enviando}
                      onChange={(e) => mudar('precoM2', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                    <Entrada
                      id="precoKg"
                      rotulo="Preço médio por Kg"
                      value={dados.precoKg}
                      disabled={enviando}
                      onChange={(e) => mudar('precoKg', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                    <Entrada
                      id="volumeM2"
                      rotulo="Potencial Volume MSM mensal"
                      required
                      value={dados.volumeM2}
                      disabled={enviando}
                      onChange={(e) => mudar('volumeM2', e.target.value)}
                      larguraClasse="sm:col-span-4"
                    />
                  </>
                )}

                <Entrada
                  id="volumeUnidades"
                  rotulo="Potencial Volume mensal (unidades)"
                  type="number"
                  min="0"
                  value={dados.volumeUnidades}
                  disabled={enviando}
                  ajuda="Opcional"
                  onChange={(e) => mudar('volumeUnidades', e.target.value)}
                  larguraClasse="sm:col-span-6"
                />
                <Entrada
                  id="volumeToneladas"
                  rotulo="Potencial Volume mensal (toneladas)"
                  type="number"
                  min="0"
                  value={dados.volumeToneladas}
                  disabled={enviando}
                  ajuda="Opcional"
                  onChange={(e) => mudar('volumeToneladas', e.target.value)}
                  larguraClasse="sm:col-span-6"
                />
              </div>
            </Secao>

            <Secao titulo="Condições" icone={Coins}>
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoMercado"
                  rotulo="Tipo de mercado"
                  required
                  value={dados.tipoMercado}
                  disabled={enviando}
                  onChange={(e) =>
                    mudar('tipoMercado', e.target.value as DadosSolicitacaoCredito['tipoMercado'])
                  }
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione o tipo de mercado</option>
                  <option value="NACIONAL">Nacional</option>
                  <option value="EXPORTACAO">Exportação</option>
                </Selecao>

                <Campo rotulo="Condição de pagamento" className="sm:col-span-4">
                  <input
                    type="text"
                    readOnly
                    value={cliente?.condicaoPagamento ?? ''}
                    placeholder="Definida pelo cadastro do cliente"
                    className={`${classesControle} cursor-default bg-slate-100 dark:bg-slate-800/50`}
                  />
                </Campo>

                <Entrada
                  id="limiteSolicitado"
                  rotulo="Limite Desejado"
                  required
                  placeholder="0,00"
                  value={dados.limiteSolicitado}
                  disabled={enviando}
                  onChange={(e) => mudar('limiteSolicitado', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />

                <Campo rotulo="Observação" htmlFor="observacaoVendedor" className="sm:col-span-12">
                  <textarea
                    id="observacaoVendedor"
                    rows={3}
                    value={dados.observacao}
                    disabled={enviando}
                    placeholder="Justificativa do aumento solicitado..."
                    onChange={(e) => mudar('observacao', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Documentos" icone={Paperclip}>
              {dados.outrasVendas && (
                <p className="mb-3 rounded-lg bg-swnet-50 px-3 py-2 text-xs text-swnet-700 dark:bg-swnet-950/40 dark:text-swnet-300">
                  Venda de resíduos/aparas: o envio de documentos é opcional.
                </p>
              )}

              <div className="grid gap-4 sm:grid-cols-3">
                <Entrada
                  id="contratoSocial"
                  rotulo="Contrato Social"
                  required={!dados.outrasVendas}
                  placeholder="Nome do arquivo"
                  value={dados.contratoSocial}
                  disabled={enviando}
                  onChange={(e) => mudar('contratoSocial', e.target.value)}
                />
                <Entrada
                  id="balancoDre"
                  rotulo="Balanço / DRE do exercício"
                  required={!dados.outrasVendas}
                  placeholder="Nome do arquivo"
                  value={dados.balancoDre}
                  disabled={enviando}
                  onChange={(e) => mudar('balancoDre', e.target.value)}
                />
                <Entrada
                  id="documentosComplementares"
                  rotulo="Documentos Complementares"
                  placeholder="Nome do arquivo"
                  value={dados.documentosComplementares}
                  disabled={enviando}
                  ajuda="Opcional"
                  onChange={(e) => mudar('documentosComplementares', e.target.value)}
                />
              </div>

              <p className="mt-2 text-xs text-slate-400">
                O upload real dos arquivos depende do backend; por ora informe o nome do documento.
              </p>
            </Secao>

            <div className="mt-4 flex justify-end">
              <Botao tipo="submit" variante="sucesso" icone={Send} desabilitado={enviando}>
                {enviando ? 'Enviando...' : 'Enviar solicitação'}
              </Botao>
            </div>
          </form>
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
