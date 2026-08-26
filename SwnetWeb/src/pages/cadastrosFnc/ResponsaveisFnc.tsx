import { useCallback, useEffect, useState } from 'react'
import { Contact, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Selecao, SelecaoMultipla } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirResponsavel,
  listarResponsaveis,
  listarSetores,
  salvarResponsavel
} from '@/services/cadastrosFnc/fncApi'
import { EMPRESAS, codigoEmpresa, formatarData, nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'
import type { DadosResponsavelFnc, ResponsavelFnc, SetorFnc } from './tipos'

/** Refatoração de Cadastros-FNC/Cadastro-Responsavel-FNC. */

const ROTULO_TIPO = { '1': 'Analista', '2': 'Responsável Setorial' } as const

const RESPONSAVEL_VAZIO: DadosResponsavelFnc = { idUsuario: '', idSetor: '', tipo: '', plantas: [] }

const OPCOES_PLANTA = EMPRESAS.map((empresa) => ({
  valor: codigoEmpresa(empresa.id),
  rotulo: empresa.nome
}))

export function ResponsaveisFnc() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [responsaveis, setResponsaveis] = useState<ResponsavelFnc[]>([])
  const [setores, setSetores] = useState<SetorFnc[]>([])
  const [carregando, setCarregando] = useState(true)

  const [filtroPlanta, setFiltroPlanta] = useState('TODOS')
  const [filtroSetor, setFiltroSetor] = useState('TODOS')

  const [edicao, setEdicao] = useState<DadosResponsavelFnc | null>(null)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<ResponsavelFnc | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, listaSetores] = await Promise.all([
        listarResponsaveis(filtroPlanta, filtroSetor),
        listarSetores()
      ])
      setResponsaveis(lista)
      setSetores(listaSetores)
    } finally {
      setCarregando(false)
    }
  }, [filtroPlanta, filtroSetor])

  useEffect(() => {
    carregar()
  }, [carregar])

  function abrirCriacao() {
    setErro('')
    setUsuario(null)
    setEdicao(RESPONSAVEL_VAZIO)
  }

  function abrirEdicao(responsavel: ResponsavelFnc) {
    setErro('')
    setUsuario(usuariosSwnet.find((u) => u.id === responsavel.id) ?? null)
    setEdicao({
      id: responsavel.id,
      idUsuario: responsavel.id,
      idSetor: responsavel.idSetor,
      tipo: responsavel.tipo,
      plantas: [...responsavel.plantas]
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarResponsavel({ ...edicao, idUsuario: usuario?.id ?? '' })
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

    const resultado = await excluirResponsavel(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Responsáveis de FNC" />

      <Painel titulo="Filtros">
        <div className="grid gap-4 sm:grid-cols-12">
          <Selecao
            id="plantaFiltro"
            rotulo="Planta"
            value={filtroPlanta}
            onChange={(e) => setFiltroPlanta(e.target.value)}
            larguraClasse="sm:col-span-4"
          >
            <option value="TODOS">Todas Plantas</option>
            {OPCOES_PLANTA.map((opcao) => (
              <option key={opcao.valor} value={opcao.valor}>
                {opcao.rotulo}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="setorFiltro"
            rotulo="Setor"
            value={filtroSetor}
            onChange={(e) => setFiltroSetor(e.target.value)}
            larguraClasse="sm:col-span-4"
          >
            <option value="TODOS">Todos Setores</option>
            {setores.map((setor) => (
              <option key={setor.id} value={setor.id}>
                {setor.descricao}
              </option>
            ))}
          </Selecao>
        </div>
      </Painel>

      <Painel
        titulo="Responsáveis"
        icone={Contact}
        acoes={
          <Botao icone={Plus} aoClicar={abrirCriacao}>
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={responsaveis}
          chave={(responsavel) => responsavel.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (r) => r.id, className: 'w-24' },
            { cabecalho: 'Setor', conteudo: (r) => r.nomeSetor || '—' },
            { cabecalho: 'Nome', conteudo: (r) => r.nome },
            { cabecalho: 'E-mail', conteudo: (r) => r.email },
            {
              cabecalho: 'Tipo',
              conteudo: (r) => (
                <Etiqueta texto={ROTULO_TIPO[r.tipo]} tom={r.tipo === '1' ? 'info' : 'alerta'} />
              )
            },
            {
              cabecalho: 'Plantas',
              conteudo: (r) => (
                <span className="text-xs">{r.plantas.map(nomeEmpresa).join(', ') || '—'}</span>
              )
            },
            { cabecalho: 'Data de Cadastro', conteudo: (r) => formatarData(r.dataInclusao) }
          ]}
          acoes={(responsavel) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Editar responsável" aoClicar={() => abrirEdicao(responsavel)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir responsável"
                tom="perigo"
                aoClicar={() => setExcluindo(responsavel)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Responsável de FNC' : 'Novo Responsável de FNC'}
        icone={Contact}
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

            <SeletorUsuario
              rotulo="Responsável"
              selecionado={usuario}
              aoSelecionar={setUsuario}
              apenasAtivos
            />

            <div className="grid gap-4 sm:grid-cols-2">
              <Selecao
                id="setorResponsavel"
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

              <Selecao
                id="tipoResponsavel"
                rotulo="Tipo"
                required
                value={edicao.tipo}
                disabled={salvando}
                onChange={(e) =>
                  setEdicao({ ...edicao, tipo: e.target.value as DadosResponsavelFnc['tipo'] })
                }
              >
                <option value="">Selecione</option>
                <option value="1">Analista</option>
                <option value="2">Responsável Setorial</option>
              </Selecao>
            </div>

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
        titulo="Excluir responsável"
        mensagem={excluindo ? `Deseja realmente excluir o responsável ${excluindo.nome}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
