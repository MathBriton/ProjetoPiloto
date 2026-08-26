import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Selecao, SelecaoMultipla } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirAprovadorOrcamento,
  listarAprovadoresOrcamento,
  salvarAprovadorOrcamento
} from '@/services/cadastrosOrcamento/orcamentoApi'
import { EMPRESAS, codigoEmpresa, formatarData, nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'
import type { AprovadorOrcamento, DadosAprovadorOrcamento } from './tipos'

/** Refatoração de Cadastros-Orçamento/Cadastro-Aprovadores-De-Orcamento. */

const OPCOES_PLANTA = EMPRESAS.map((empresa) => ({
  valor: codigoEmpresa(empresa.id),
  rotulo: empresa.nome
}))

const APROVADOR_VAZIO: DadosAprovadorOrcamento = { idUsuario: '', nivel: '', plantas: [] }

const TOM_NIVEL = { '1': 'info', '2': 'alerta', '3': 'perigo' } as const

export function AprovadoresOrcamento() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [aprovadores, setAprovadores] = useState<AprovadorOrcamento[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtroPlanta, setFiltroPlanta] = useState('TODOS')

  const [edicao, setEdicao] = useState<DadosAprovadorOrcamento | null>(null)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<AprovadorOrcamento | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setAprovadores(await listarAprovadoresOrcamento(filtroPlanta))
    } finally {
      setCarregando(false)
    }
  }, [filtroPlanta])

  useEffect(() => {
    carregar()
  }, [carregar])

  function abrirCriacao() {
    setErro('')
    setUsuario(null)
    setEdicao(APROVADOR_VAZIO)
  }

  function abrirEdicao(aprovador: AprovadorOrcamento) {
    setErro('')
    setUsuario(usuariosSwnet.find((u) => u.id === aprovador.id) ?? null)
    setEdicao({
      id: aprovador.id,
      idUsuario: aprovador.id,
      nivel: aprovador.nivel,
      plantas: [...aprovador.plantas]
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarAprovadorOrcamento({ ...edicao, idUsuario: usuario?.id ?? '' })
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

    const resultado = await excluirAprovadorOrcamento(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Aprovadores de Orçamento" />

      <Painel titulo="Filtros">
        <Selecao
          id="plantaFiltro"
          rotulo="Planta"
          value={filtroPlanta}
          onChange={(e) => setFiltroPlanta(e.target.value)}
          larguraClasse="max-w-xs"
        >
          <option value="TODOS">Todas Plantas</option>
          {OPCOES_PLANTA.map((opcao) => (
            <option key={opcao.valor} value={opcao.valor}>
              {opcao.rotulo}
            </option>
          ))}
        </Selecao>
      </Painel>

      <Painel
        titulo="Aprovadores"
        icone={ClipboardCheck}
        acoes={
          <Botao icone={Plus} aoClicar={abrirCriacao}>
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
            {
              cabecalho: 'Nível',
              conteudo: (a) => <Etiqueta texto={`Nível ${a.nivel}`} tom={TOM_NIVEL[a.nivel]} />
            },
            {
              cabecalho: 'Plantas',
              conteudo: (a) => (
                <span className="text-xs">{a.plantas.map(nomeEmpresa).join(', ') || '—'}</span>
              )
            },
            { cabecalho: 'Data de Cadastro', conteudo: (a) => formatarData(a.dataInclusao) }
          ]}
          acoes={(aprovador) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Editar aprovador" aoClicar={() => abrirEdicao(aprovador)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir aprovador"
                tom="perigo"
                aoClicar={() => setExcluindo(aprovador)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Aprovador' : 'Novo Aprovador'}
        icone={ClipboardCheck}
        largura="lg"
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

            <SeletorUsuario rotulo="Aprovador" selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />

            <Selecao
              id="nivelAprovador"
              rotulo="Nível"
              required
              value={edicao.nivel}
              disabled={salvando}
              onChange={(e) =>
                setEdicao({ ...edicao, nivel: e.target.value as DadosAprovadorOrcamento['nivel'] })
              }
              larguraClasse="max-w-xs"
            >
              <option value="">Selecione</option>
              <option value="1">Nível 1</option>
              <option value="2">Nível 2</option>
              <option value="3">Nível 3</option>
            </Selecao>

            <SelecaoMultipla
              rotulo="Plantas"
              obrigatorio
              opcoes={OPCOES_PLANTA}
              selecionados={edicao.plantas}
              desabilitado={salvando}
              aoMudar={(plantas) => setEdicao({ ...edicao, plantas })}
            />
          </div>
        )}
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
