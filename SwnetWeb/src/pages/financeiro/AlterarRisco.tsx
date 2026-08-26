import { useEffect, useMemo, useState } from 'react'
import { Eraser, Search, Sliders } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { BuscaCliente } from '@/components/financeiro/BuscaCliente'
import {
  PLANTAS_PROTHEUS,
  carregarCliente,
  gravarRisco,
  opcoesRisco,
  type ClienteFinanceiro
} from '@/services/financeiro/financeiroApi'

/** Refatoração de Financeiro/Alterar-Risco (alterar_risco.asp + Repositorio). */

export function AlterarRisco() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [planta, setPlanta] = useState('')
  const [codigo, setCodigo] = useState('')
  const [loja, setLoja] = useState('')
  const [cliente, setCliente] = useState<ClienteFinanceiro | null>(null)
  const [novoRisco, setNovoRisco] = useState('')
  const [buscaAberta, setBuscaAberta] = useState(false)
  const [carregando, setCarregando] = useState(false)
  const [confirmando, setConfirmando] = useState(false)
  const [gravando, setGravando] = useState(false)

  /** As opções mudam conforme o risco atual e se a loja foi informada */
  const opcoes = useMemo(
    () => opcoesRisco(cliente?.risco ?? '', loja),
    [cliente?.risco, loja]
  )

  // Trocar de planta invalida o cliente carregado
  useEffect(() => {
    setCliente(null)
    setNovoRisco('')
  }, [planta])

  // Se a opção escolhida sumiu da lista, limpa a seleção
  useEffect(() => {
    if (novoRisco && !opcoes.some((opcao) => opcao.valor === novoRisco)) setNovoRisco('')
  }, [opcoes, novoRisco])

  async function filtrar() {
    if (!planta || !codigo.trim()) {
      mostrarToast('Informe a planta e o código do cliente.', 'warning')
      return
    }

    setCarregando(true)
    try {
      const encontrado = await carregarCliente(planta, codigo, loja)
      if (!encontrado) {
        setCliente(null)
        mostrarToast('Cliente não encontrado nesta planta.', 'warning')
        return
      }
      setCliente(encontrado)
      setNovoRisco('')
    } finally {
      setCarregando(false)
    }
  }

  function selecionarNaBusca(selecionado: ClienteFinanceiro) {
    setCodigo(selecionado.codigo)
    setLoja(selecionado.loja)
    setCliente(selecionado)
    setNovoRisco('')
  }

  function limpar() {
    setCodigo('')
    setLoja('')
    setCliente(null)
    setNovoRisco('')
  }

  async function confirmar() {
    if (!cliente) return

    setConfirmando(false)
    setGravando(true)
    try {
      const resultado = await gravarRisco(planta, cliente.codigo, loja, novoRisco)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        const atualizado = await carregarCliente(planta, cliente.codigo, loja)
        setCliente(atualizado)
        setNovoRisco('')
      }
    } finally {
      setGravando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Alterar Risco do Cliente" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Risco do Cliente" icone={Sliders}>
          <form
            className="grid gap-4 sm:grid-cols-12"
            onSubmit={(e) => {
              e.preventDefault()
              filtrar()
            }}
          >
            <Selecao
              id="Planta"
              rotulo="Planta"
              required
              value={planta}
              disabled={carregando || gravando}
              onChange={(e) => setPlanta(e.target.value)}
              larguraClasse="sm:col-span-4"
            >
              <option value="">Selecione...</option>
              {PLANTAS_PROTHEUS.map((item) => (
                <option key={item.codigo} value={item.codigo}>
                  {item.nome}
                </option>
              ))}
            </Selecao>

            <Campo rotulo="Código do Cliente" htmlFor="codcli" obrigatorio className="sm:col-span-5">
              <div className="flex">
                <button
                  type="button"
                  onClick={() => setBuscaAberta(true)}
                  disabled={!planta || carregando || gravando}
                  aria-label="Pesquisar cliente"
                  title="Pesquisar cliente"
                  className="rounded-l-lg border border-r-0 border-slate-200 bg-slate-100 px-3 text-slate-500
                    hover:bg-slate-200 disabled:opacity-50 dark:border-slate-700 dark:bg-slate-800"
                >
                  <Search size={15} />
                </button>
                <input
                  id="codcli"
                  type="text"
                  placeholder="Código..."
                  value={codigo}
                  disabled={carregando || gravando}
                  onChange={(e) => setCodigo(e.target.value.trim())}
                  className={`${classesControle} rounded-l-none`}
                />
              </div>
            </Campo>

            <Entrada
              id="codloja"
              rotulo="Loja"
              maxLength={2}
              placeholder="Nº..."
              value={loja}
              disabled={carregando || gravando}
              ajuda="Em branco altera todas as lojas"
              onChange={(e) => setLoja(e.target.value.replace(/\D/g, ''))}
              larguraClasse="sm:col-span-3"
            />

            <div className="flex gap-2 sm:col-span-12">
              <Botao tipo="submit" icone={Search} desabilitado={carregando || gravando}>
                {carregando ? 'Buscando...' : 'Filtrar'}
              </Botao>
              <Botao variante="neutro" icone={Eraser} aoClicar={limpar} desabilitado={gravando}>
                Limpar
              </Botao>
            </div>
          </form>

          {cliente && (
            <div className="mt-4 grid gap-4 border-t border-slate-200 pt-4 sm:grid-cols-12 dark:border-slate-800">
              <Campo rotulo="Razão Social" className="sm:col-span-5">
                <input
                  type="text"
                  readOnly
                  value={cliente.nome}
                  className={`${classesControle} cursor-default bg-slate-100 dark:bg-slate-800/50`}
                />
              </Campo>

              <Campo rotulo="Risco Atual" className="sm:col-span-3">
                <div className="flex h-[38px] items-center">
                  {cliente.risco ? (
                    <Etiqueta texto={`RISCO ${cliente.risco}`} tom={cliente.risco === 'A' ? 'sucesso' : 'alerta'} />
                  ) : (
                    <span className="text-sm text-slate-400">Sem risco cadastrado</span>
                  )}
                </div>
              </Campo>

              <Selecao
                id="cboRisco"
                rotulo="Alterar o Risco"
                value={novoRisco}
                disabled={gravando}
                onChange={(e) => setNovoRisco(e.target.value)}
                larguraClasse="sm:col-span-4"
              >
                <option value="">Selecione...</option>
                {opcoes.map((opcao) => (
                  <option key={opcao.valor} value={opcao.valor}>
                    {opcao.rotulo}
                  </option>
                ))}
              </Selecao>

              <div className="flex justify-end sm:col-span-12">
                <Botao
                  variante="sucesso"
                  aoClicar={() => setConfirmando(true)}
                  desabilitado={!novoRisco || gravando}
                >
                  {gravando ? 'Gravando...' : 'Alterar o Risco'}
                </Botao>
              </div>
            </div>
          )}
        </Painel>
      </div>

      <BuscaCliente
        aberto={buscaAberta}
        planta={planta}
        aoFechar={() => setBuscaAberta(false)}
        aoSelecionar={selecionarNaBusca}
      />

      <ConfirmDialog
        aberto={confirmando}
        titulo="Confirmar alteração"
        mensagem={
          cliente
            ? `O risco de ${cliente.nome} passará de "${cliente.risco || 'sem risco'}" para "${novoRisco}"${
                loja ? ` na loja ${loja}` : ' em todas as lojas'
              }. Deseja continuar?`
            : ''
        }
        rotuloConfirmar="Confirmar"
        aoConfirmar={confirmar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
