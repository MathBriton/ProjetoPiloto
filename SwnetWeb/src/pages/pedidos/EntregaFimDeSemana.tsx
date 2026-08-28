import { useCallback, useEffect, useState } from 'react'
import { Plus, Save, Trash2, Truck } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Chave, Entrada } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirEntregaFds,
  listarEntregaFds,
  salvarEntregaFds,
  type ClienteEntregaFds
} from '@/services/pedidos/pedidosApi'

/**
 * Refatoração de Pedidos/Definir-Entrega-No-Fim-De-Semana.
 *
 * Lista de clientes liberados para receber no fim de semana. O legado
 * valida o código contra a base de clientes antes de incluir.
 */
export function EntregaFimDeSemana() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [clientes, setClientes] = useState<ClienteEntregaFds[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{ codigo: string; ativo: boolean } | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<ClienteEntregaFds | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setClientes(await listarEntregaFds())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function salvar() {
    if (!edicao) return
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarEntregaFds(edicao.codigo, edicao.ativo)
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirEntregaFds(excluindo.codigo)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Entrega no Fim de Semana" />

      <div className="mx-auto max-w-3xl">
        <Painel
          titulo="Clientes liberados"
          icone={Truck}
          acoes={
            <Botao
              icone={Plus}
              aoClicar={() => {
                setErro('')
                setEdicao({ codigo: '', ativo: true })
              }}
            >
              Incluir cliente
            </Botao>
          }
        >
          <Tabela
            itens={clientes}
            chave={(item) => item.codigo}
            carregando={carregando}
            mensagemVazia="Nenhum cliente liberado para entrega no fim de semana"
            colunas={[
              { cabecalho: 'Código', conteudo: (c) => c.codigo, className: 'w-32' },
              { cabecalho: 'Nome', conteudo: (c) => c.nome },
              {
                cabecalho: 'Ativo',
                conteudo: (c) => (
                  <Etiqueta texto={c.ativo ? 'Sim' : 'Não'} tom={c.ativo ? 'sucesso' : 'neutro'} />
                )
              }
            ]}
            acoes={(item) => (
              <>
                <BotaoAcao
                  icone={Save}
                  titulo={item.ativo ? 'Desativar' : 'Ativar'}
                  aoClicar={async () => {
                    const resultado = await salvarEntregaFds(item.codigo, !item.ativo)
                    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
                    if (resultado.sucesso) await carregar()
                  }}
                />
                <BotaoAcao
                  icone={Trash2}
                  titulo="Remover da lista"
                  tom="perigo"
                  aoClicar={() => setExcluindo(item)}
                />
              </>
            )}
          />
        </Painel>
      </div>

      <Modal
        aberto={edicao !== null}
        titulo="Incluir cliente"
        icone={Truck}
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Cancelar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}
            <Entrada
              id="codigoClienteFds"
              rotulo="Código do cliente"
              required
              value={edicao.codigo}
              disabled={salvando}
              ajuda="O código é validado contra a base de clientes"
              onChange={(e) => setEdicao({ ...edicao, codigo: e.target.value.trim() })}
            />
            <Chave
              rotulo="Ativo"
              marcado={edicao.ativo}
              desabilitado={salvando}
              aoAlternar={(v) => setEdicao({ ...edicao, ativo: v })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Remover cliente"
        mensagem={excluindo ? `Remover ${excluindo.nome} da lista de entrega no fim de semana?` : ''}
        rotuloConfirmar="Remover"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
