import { useCallback, useEffect, useState } from 'react'
import { Building2, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { excluirSetor, listarSetores, salvarSetor } from '@/services/cadastrosFnc/fncApi'
import type { SetorFnc } from './tipos'

/** Refatoração de Cadastros-FNC/Cadastro-Setores-De-FNC (fncSetorGrade/Formulario/Repositorio). */
export function SetoresFnc() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [setores, setSetores] = useState<SetorFnc[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{ id?: number; descricao: string } | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<SetorFnc | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setSetores(await listarSetores())
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
      const resultado = await salvarSetor(edicao)
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

    const resultado = await excluirSetor(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Setores de FNC" />

      <Painel
        titulo="Setores"
        icone={Building2}
        acoes={
          <Botao icone={Plus} aoClicar={() => { setErro(''); setEdicao({ descricao: '' }) }}>
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={setores}
          chave={(setor) => setor.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (s) => s.id, className: 'w-20' },
            { cabecalho: 'Descrição', conteudo: (s) => s.descricao }
          ]}
          acoes={(setor) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar setor"
                aoClicar={() => { setErro(''); setEdicao({ id: setor.id, descricao: setor.descricao }) }}
              />
              <BotaoAcao icone={Trash2} titulo="Excluir setor" tom="perigo" aoClicar={() => setExcluindo(setor)} />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Setor de FNC' : 'Novo Setor de FNC'}
        icone={Building2}
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <>
            {erro && (
              <p className="mb-3 rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}
            <Entrada
              id="descricaoSetor"
              rotulo="Descrição"
              required
              autoFocus
              value={edicao.descricao}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, descricao: e.target.value })}
            />
          </>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir setor"
        mensagem={
          excluindo ? `Deseja realmente excluir o setor "${excluindo.descricao}"?` : ''
        }
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
