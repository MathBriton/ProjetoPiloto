import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { excluirAprovador, listarAprovadores, salvarAprovador } from '@/services/cadastrosFnc/fncApi'
import { formatarData } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'
import type { AprovadorFnc } from './tipos'

/**
 * Refatoração de Cadastros-FNC/Cadastro-Aprovador-FNC.
 *
 * O formulário legado é só um select2 de usuário: nome e e-mail são copiados
 * da CRM_DTI no momento da inclusão, por isso não há edição — só incluir e
 * excluir.
 */
export function AprovadoresFnc() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [aprovadores, setAprovadores] = useState<AprovadorFnc[]>([])
  const [carregando, setCarregando] = useState(true)
  const [modalAberto, setModalAberto] = useState(false)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<AprovadorFnc | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setAprovadores(await listarAprovadores())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  function abrirModal() {
    setUsuario(null)
    setErro('')
    setModalAberto(true)
  }

  async function salvar() {
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarAprovador(usuario?.id ?? '')
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setModalAberto(false)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return

    const resultado = await excluirAprovador(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Aprovadores de FNC" />

      <Painel
        titulo="Aprovadores"
        icone={ClipboardCheck}
        acoes={
          <Botao icone={Plus} aoClicar={abrirModal}>
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={aprovadores}
          chave={(aprovador) => aprovador.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (a) => a.id, className: 'w-24' },
            { cabecalho: 'Nome', conteudo: (a) => a.nome },
            { cabecalho: 'E-mail', conteudo: (a) => a.email },
            { cabecalho: 'Data Inclusão', conteudo: (a) => formatarData(a.dataInclusao) }
          ]}
          acoes={(aprovador) => (
            <BotaoAcao
              icone={Trash2}
              titulo="Excluir aprovador"
              tom="perigo"
              aoClicar={() => setExcluindo(aprovador)}
            />
          )}
        />
      </Painel>

      <Modal
        aberto={modalAberto}
        titulo="Novo Aprovador de FNC"
        icone={ClipboardCheck}
        aoFechar={() => setModalAberto(false)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setModalAberto(false)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        <div className="grid gap-4">
          {erro && (
            <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
              {erro}
            </p>
          )}

          <SeletorUsuario selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />

          <p className="text-xs text-slate-400">
            Nome e e-mail são copiados do cadastro do usuário no momento da inclusão.
          </p>
        </div>
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir aprovador"
        mensagem={excluindo ? `Deseja realmente excluir o aprovador ${excluindo.nome}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
