import { useCallback, useEffect, useState } from 'react'
import { Pencil, Plus, Save, Trash2, Wrench } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirDefeito,
  listarDefeitos,
  listarSetores,
  salvarDefeito
} from '@/services/cadastrosFnc/fncApi'
import type { DadosDefeitoFnc, DefeitoFnc, SetorFnc } from './tipos'

/** Refatoração de Cadastros-FNC/Cadastro-Defeitos-De-FNC. */

const DEFEITO_VAZIO: DadosDefeitoFnc = { tipo: '', idSetor: '', descricao: '' }

const ROTULO_TIPO = { ADM: 'Administrativo', PRO: 'Industrial' } as const

export function DefeitosFnc() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [defeitos, setDefeitos] = useState<DefeitoFnc[]>([])
  const [setores, setSetores] = useState<SetorFnc[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtro, setFiltro] = useState('')
  const [edicao, setEdicao] = useState<DadosDefeitoFnc | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<DefeitoFnc | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, listaSetores] = await Promise.all([listarDefeitos(), listarSetores()])
      setDefeitos(lista)
      setSetores(listaSetores)
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
      const resultado = await salvarDefeito(edicao)
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

    const resultado = await excluirDefeito(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  const busca = filtro.trim().toLowerCase()
  const visiveis = busca
    ? defeitos.filter(
        (d) =>
          d.descricao.toLowerCase().includes(busca) ||
          d.nomeSetor.toLowerCase().includes(busca) ||
          ROTULO_TIPO[d.tipo].toLowerCase().includes(busca)
      )
    : defeitos

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Defeitos de FNC" />

      <Painel
        titulo="Defeitos"
        icone={Wrench}
        acoes={
          <>
            <CampoPesquisa valor={filtro} aoMudar={setFiltro} placeholder="Pesquisar defeito..." />
            <Botao icone={Plus} aoClicar={() => { setErro(''); setEdicao(DEFEITO_VAZIO) }}>
              Novo
            </Botao>
          </>
        }
      >
        <Tabela
          itens={visiveis}
          chave={(defeito) => defeito.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (d) => d.id, className: 'w-20' },
            {
              cabecalho: 'Tipo',
              conteudo: (d) => (
                <Etiqueta texto={ROTULO_TIPO[d.tipo]} tom={d.tipo === 'ADM' ? 'info' : 'alerta'} />
              )
            },
            { cabecalho: 'Setor', conteudo: (d) => d.nomeSetor || '—' },
            { cabecalho: 'Descrição', conteudo: (d) => d.descricao }
          ]}
          acoes={(defeito) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar defeito"
                aoClicar={() => {
                  setErro('')
                  setEdicao({
                    id: defeito.id,
                    tipo: defeito.tipo,
                    idSetor: defeito.idSetor,
                    descricao: defeito.descricao
                  })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir defeito"
                tom="perigo"
                aoClicar={() => setExcluindo(defeito)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Defeito de FNC' : 'Novo Defeito de FNC'}
        icone={Wrench}
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
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <div className="grid gap-4 sm:grid-cols-2">
              <Selecao
                id="tipoDefeito"
                rotulo="Tipo"
                required
                value={edicao.tipo}
                disabled={salvando}
                onChange={(e) => setEdicao({ ...edicao, tipo: e.target.value as DadosDefeitoFnc['tipo'] })}
              >
                <option value="">Selecione</option>
                <option value="ADM">Administrativo</option>
                <option value="PRO">Industrial</option>
              </Selecao>

              <Selecao
                id="setorDefeito"
                rotulo="Setor"
                required
                value={edicao.idSetor}
                disabled={salvando}
                onChange={(e) =>
                  setEdicao({ ...edicao, idSetor: e.target.value ? Number(e.target.value) : '' })
                }
              >
                <option value="">Selecione um setor</option>
                {setores.map((setor) => (
                  <option key={setor.id} value={setor.id}>
                    {setor.descricao}
                  </option>
                ))}
              </Selecao>
            </div>

            <Entrada
              id="descricaoDefeito"
              rotulo="Descrição"
              required
              value={edicao.descricao}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, descricao: e.target.value })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir defeito"
        mensagem={excluindo ? `Deseja realmente excluir o defeito "${excluindo.descricao}"?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
