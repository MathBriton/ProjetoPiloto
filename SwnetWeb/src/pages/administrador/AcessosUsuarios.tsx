import { useState } from 'react'
import { Save, Search, Copy, FileDown, X, UserRound, Wrench, BarChart3 } from 'lucide-react'
import { obterPermissoes, salvarPermissoes, importarPermissoes } from '@/services/acessosUsuariosApi'
import type { PermissoesUsuario, UsuarioBusca } from './tipos'
import { BuscaUsuario } from '@/components/acessosUsuarios/BuscaUsuario'
import { SecaoPermissoes } from '@/components/acessosUsuarios/SecaoPermissoes'
import { ModalImportarPermissoes } from '@/components/acessosUsuarios/ModalImportarPermissoes'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { useToast, ToastStack } from '@/components/Toast'

export function AcessosUsuarios() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [usuario, setUsuario] = useState<UsuarioBusca | null>(null)
  const [permissoes, setPermissoes] = useState<PermissoesUsuario | null>(null)
  const [carregando, setCarregando] = useState(false)
  const [salvando, setSalvando] = useState(false)
  const [filtro, setFiltro] = useState('')

  const [modalImportarAberto, setModalImportarAberto] = useState(false)
  const [confirmacao, setConfirmacao] = useState<{ titulo: string; mensagem: string; aoConfirmar: () => void } | null>(null)

  async function selecionarUsuario(usuarioSelecionado: UsuarioBusca) {
    setUsuario(usuarioSelecionado)
    setCarregando(true)
    try {
      const dados = await obterPermissoes(usuarioSelecionado.matricula, usuarioSelecionado.filial)
      setPermissoes(dados)
    } finally {
      setCarregando(false)
    }
  }

  function voltarParaBusca() {
    setUsuario(null)
    setPermissoes(null)
    setFiltro('')
  }

  function alternarItem(tipo: 'servicos' | 'consultas', codigo: string) {
    setPermissoes((atual) => {
      if (!atual) return atual
      return {
        ...atual,
        [tipo]: atual[tipo].map((item) => (item.codigo === codigo ? { ...item, habilitado: !item.habilitado } : item))
      }
    })
  }

  function alternarCategoria(tipo: 'servicos' | 'consultas', categoria: string, marcar: boolean) {
    setPermissoes((atual) => {
      if (!atual) return atual
      return {
        ...atual,
        [tipo]: atual[tipo].map((item) => (item.categoria === categoria ? { ...item, habilitado: marcar } : item))
      }
    })
  }

  async function salvar() {
    if (!usuario || !permissoes) return

    const totalMarcados =
      permissoes.servicos.filter((i) => i.habilitado).length + permissoes.consultas.filter((i) => i.habilitado).length

    const executarSalvamento = async () => {
      setSalvando(true)
      try {
        await salvarPermissoes(usuario.matricula, usuario.filial, permissoes.servicos, permissoes.consultas)
        mostrarToast('Permissões salvas com sucesso.', 'success')
      } finally {
        setSalvando(false)
        setConfirmacao(null)
      }
    }

    if (totalMarcados === 0) {
      setConfirmacao({
        titulo: 'Nenhuma permissão selecionada',
        mensagem: 'Nenhuma permissão foi marcada para este usuário. Deseja continuar mesmo assim?',
        aoConfirmar: executarSalvamento
      })
      return
    }

    await executarSalvamento()
  }

  function confirmarImportacao(usuarioOrigem: UsuarioBusca) {
    setModalImportarAberto(false)
    setConfirmacao({
      titulo: 'Importar permissões',
      mensagem: `As permissões atuais serão substituídas pelas de ${usuarioOrigem.nome}.`,
      aoConfirmar: async () => {
        if (!usuario) return
        setCarregando(true)
        try {
          const dados = await importarPermissoes(usuarioOrigem, usuario)
          setPermissoes(dados)
          mostrarToast(`Permissões importadas de ${usuarioOrigem.nome}.`, 'success')
        } finally {
          setCarregando(false)
          setConfirmacao(null)
        }
      }
    })
  }

  function gerarPdf() {
    // TODO: exportação em PDF ainda não implementada (sem lib de PDF no frontend por ora)
    mostrarToast('Exportação em PDF será implementada em breve.', 'info')
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between px-1">
        <h1 className="text-xl font-bold text-slate-800 dark:text-white">Acessos Usuários</h1>
      </div>

      {!usuario && <BuscaUsuario aoSelecionar={selecionarUsuario} />}

      {usuario && (
        <>
          <div className="flex flex-wrap items-center gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3
            shadow-sm dark:border-slate-800 dark:bg-slate-900">
            <div className="flex h-9 w-9 items-center justify-center rounded-full bg-slate-500 text-white">
              <UserRound size={18} />
            </div>
            <div className="min-w-0">
              <p className="truncate text-sm font-semibold text-slate-800 dark:text-white">{usuario.nome}</p>
              <p className="text-xs text-slate-400">
                {usuario.login} &bull; Filial {usuario.filial}
              </p>
            </div>

            <div className="ml-auto flex flex-wrap items-center gap-2">
              <button
                onClick={salvar}
                disabled={salvando || carregando}
                title="Salvar"
                aria-label="Salvar permissões"
                className="rounded-lg bg-emerald-600 p-2 text-white hover:bg-emerald-700 disabled:opacity-50"
              >
                <Save size={16} />
              </button>
              <button
                onClick={voltarParaBusca}
                title="Buscar"
                aria-label="Voltar para busca de usuário"
                className="rounded-lg bg-swnet-600 p-2 text-white hover:bg-swnet-700"
              >
                <Search size={16} />
              </button>
              <button
                onClick={() => setModalImportarAberto(true)}
                title="Importar Permissões"
                aria-label="Copiar permissões de outro usuário"
                className="rounded-lg bg-slate-500 p-2 text-white hover:bg-slate-600"
              >
                <Copy size={16} />
              </button>
              <button
                onClick={gerarPdf}
                title="Baixar PDF"
                aria-label="Gerar PDF"
                className="rounded-lg bg-red-500 p-2 text-white hover:bg-red-600"
              >
                <FileDown size={16} />
              </button>

              <div className="relative">
                <Search size={14} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  value={filtro}
                  onChange={(e) => setFiltro(e.target.value)}
                  placeholder="Pesquisar permissões"
                  className="w-48 rounded-lg border border-slate-200 bg-slate-50 py-1.5 pl-8 pr-7 text-xs
                    focus:border-swnet-500 focus:outline-none focus:ring-2 focus:ring-swnet-500/30
                    dark:border-slate-700 dark:bg-slate-800 dark:text-white"
                />
                {filtro && (
                  <button
                    onClick={() => setFiltro('')}
                    aria-label="Limpar pesquisa"
                    className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X size={13} />
                  </button>
                )}
              </div>
            </div>
          </div>

          {carregando && (
            <p className="py-10 text-center text-sm text-slate-400">Carregando permissões...</p>
          )}

          {!carregando && permissoes && (
            <div className="flex flex-col gap-4 lg:flex-row">
              <SecaoPermissoes
                titulo="Serviços"
                icone={Wrench}
                itens={permissoes.servicos}
                filtro={filtro}
                aoAlternarItem={(codigo) => alternarItem('servicos', codigo)}
                aoAlternarCategoria={(categoria, marcar) => alternarCategoria('servicos', categoria, marcar)}
              />
              <SecaoPermissoes
                titulo="Consultas"
                icone={BarChart3}
                itens={permissoes.consultas}
                filtro={filtro}
                aoAlternarItem={(codigo) => alternarItem('consultas', codigo)}
                aoAlternarCategoria={(categoria, marcar) => alternarCategoria('consultas', categoria, marcar)}
              />
            </div>
          )}
        </>
      )}

      <ModalImportarPermissoes
        aberto={modalImportarAberto}
        usuarioDestino={usuario}
        aoFechar={() => setModalImportarAberto(false)}
        aoConfirmar={confirmarImportacao}
      />

      <ConfirmDialog
        aberto={confirmacao !== null}
        titulo={confirmacao?.titulo ?? ''}
        mensagem={confirmacao?.mensagem ?? ''}
        aoConfirmar={() => confirmacao?.aoConfirmar()}
        aoCancelar={() => setConfirmacao(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
