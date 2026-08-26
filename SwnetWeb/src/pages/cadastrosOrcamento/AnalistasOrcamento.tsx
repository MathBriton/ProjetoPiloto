import { useCallback, useEffect, useState } from 'react'
import { Pencil, Plus, Save, Trash2, Users } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Selecao, SelecaoMultipla } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirAnalista,
  listarAnalistas,
  salvarAnalista
} from '@/services/cadastrosOrcamento/orcamentoApi'
import { EMPRESAS, codigoEmpresa, formatarData, nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'
import type { AnalistaOrcamento, DadosAnalistaOrcamento } from './tipos'

/** Refatoração de Cadastros-Orçamento/Cadastro-Analistas-De-Orcamento. */

const OPCOES_PLANTA = EMPRESAS.map((empresa) => ({
  valor: codigoEmpresa(empresa.id),
  rotulo: empresa.nome
}))

const ANALISTA_VAZIO: DadosAnalistaOrcamento = { idUsuario: '', plantas: [] }

export function AnalistasOrcamento() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [analistas, setAnalistas] = useState<AnalistaOrcamento[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtroPlanta, setFiltroPlanta] = useState('TODOS')

  const [edicao, setEdicao] = useState<DadosAnalistaOrcamento | null>(null)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<AnalistaOrcamento | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setAnalistas(await listarAnalistas(filtroPlanta))
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
    setEdicao(ANALISTA_VAZIO)
  }

  function abrirEdicao(analista: AnalistaOrcamento) {
    setErro('')
    setUsuario(usuariosSwnet.find((u) => u.id === analista.id) ?? null)
    setEdicao({ id: analista.id, idUsuario: analista.id, plantas: [...analista.plantas] })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarAnalista({ ...edicao, idUsuario: usuario?.id ?? '' })
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

    const resultado = await excluirAnalista(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Analistas de Orçamento" />

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
        titulo="Analistas"
        icone={Users}
        acoes={
          <Botao icone={Plus} aoClicar={abrirCriacao}>
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={analistas}
          chave={(analista) => analista.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (a) => a.id, className: 'w-24' },
            { cabecalho: 'Nome', conteudo: (a) => a.nome },
            { cabecalho: 'E-mail', conteudo: (a) => a.email },
            {
              cabecalho: 'Plantas',
              conteudo: (a) => (
                <span className="text-xs">{a.plantas.map(nomeEmpresa).join(', ') || '—'}</span>
              )
            },
            { cabecalho: 'Data de Cadastro', conteudo: (a) => formatarData(a.dataInclusao) }
          ]}
          acoes={(analista) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Editar analista" aoClicar={() => abrirEdicao(analista)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir analista"
                tom="perigo"
                aoClicar={() => setExcluindo(analista)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Analista' : 'Novo Analista'}
        icone={Users}
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

            <SeletorUsuario rotulo="Analista" selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />

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
        titulo="Excluir analista"
        mensagem={excluindo ? `Deseja realmente excluir o analista ${excluindo.nome}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
