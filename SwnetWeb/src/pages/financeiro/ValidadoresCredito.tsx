import { useCallback, useEffect, useState } from 'react'
import { ClipboardCheck, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao, SelecaoMultipla } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  NIVEIS_VALIDADOR,
  PLANTAS_CREDITO,
  excluirValidador,
  listarValidadores,
  rotuloNivel,
  salvarValidador,
  type DadosValidador,
  type ValidadorCredito
} from '@/services/financeiro/financeiroApi'
import { nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'

/** Refatoração de Financeiro/Cadastro-Validadores-De-Credito. */

const OPCOES_PLANTA = PLANTAS_CREDITO.map((planta) => ({
  valor: planta.codigo,
  rotulo: planta.nome
}))

const VALIDADOR_VAZIO: DadosValidador = {
  idUsuario: '',
  plantas: [],
  nivel: '',
  valorMinimo: '',
  valorMaximo: ''
}

const moeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

export function ValidadoresCredito() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [validadores, setValidadores] = useState<ValidadorCredito[]>([])
  const [carregando, setCarregando] = useState(true)
  const [termo, setTermo] = useState('')

  const [edicao, setEdicao] = useState<DadosValidador | null>(null)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<ValidadorCredito | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setValidadores(await listarValidadores(termo))
    } finally {
      setCarregando(false)
    }
  }, [termo])

  useEffect(() => {
    const atraso = setTimeout(carregar, 250)
    return () => clearTimeout(atraso)
  }, [carregar])

  function abrirCriacao() {
    setErro('')
    setUsuario(null)
    setEdicao(VALIDADOR_VAZIO)
  }

  function abrirEdicao(validador: ValidadorCredito) {
    setErro('')
    setUsuario(usuariosSwnet.find((u) => u.id === validador.codigo) ?? null)
    setEdicao({
      codigo: validador.codigo,
      idUsuario: validador.codigo,
      plantas: [...validador.plantas],
      nivel: validador.nivel,
      valorMinimo: String(validador.valorMinimo),
      valorMaximo: String(validador.valorMaximo)
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarValidador({ ...edicao, idUsuario: usuario?.id ?? '' })
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

    const resultado = await excluirValidador(excluindo.codigo)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Validadores de Crédito" />

      <Painel
        titulo="Validadores"
        icone={ClipboardCheck}
        acoes={
          <>
            <CampoPesquisa valor={termo} aoMudar={setTermo} placeholder="Buscar..." />
            <Botao icone={Plus} aoClicar={abrirCriacao}>
              Novo
            </Botao>
          </>
        }
      >
        <Tabela
          itens={validadores}
          chave={(validador) => validador.codigo}
          carregando={carregando}
          colunas={[
            { cabecalho: 'Código', conteudo: (v) => v.codigo, className: 'w-24' },
            { cabecalho: 'Nome', conteudo: (v) => v.nome },
            {
              cabecalho: 'Plantas',
              conteudo: (v) => (
                <span className="text-xs">{v.plantas.map(nomeEmpresa).join(', ') || '—'}</span>
              )
            },
            { cabecalho: 'Nível', conteudo: (v) => <Etiqueta texto={rotuloNivel(v.nivel)} tom="info" /> },
            {
              cabecalho: 'Mínimo',
              conteudo: (v) => moeda.format(v.valorMinimo),
              alinhamento: 'direita'
            },
            {
              cabecalho: 'Máximo',
              conteudo: (v) => moeda.format(v.valorMaximo),
              alinhamento: 'direita'
            }
          ]}
          acoes={(validador) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Alterar validador" aoClicar={() => abrirEdicao(validador)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Apagar validador"
                tom="perigo"
                aoClicar={() => setExcluindo(validador)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.codigo ? 'Alterar Validador' : 'Novo Validador'}
        icone={ClipboardCheck}
        largura="lg"
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

            <SeletorUsuario rotulo="Validador" selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />

            <div className="grid gap-4 sm:grid-cols-3">
              <Selecao
                id="cboValidadorNivel"
                rotulo="Nível"
                required
                value={edicao.nivel}
                disabled={salvando}
                onChange={(e) => setEdicao({ ...edicao, nivel: e.target.value })}
              >
                <option value="">Selecione...</option>
                {NIVEIS_VALIDADOR.map((nivel) => (
                  <option key={nivel.valor} value={nivel.valor}>
                    {nivel.rotulo}
                  </option>
                ))}
              </Selecao>

              <Entrada
                id="edtValidadorValorMinimo"
                rotulo="Valor Mínimo"
                type="number"
                min="0"
                step="0.01"
                required
                value={edicao.valorMinimo}
                disabled={salvando}
                onChange={(e) => setEdicao({ ...edicao, valorMinimo: e.target.value })}
              />
              <Entrada
                id="edtValidadorValorMaximo"
                rotulo="Valor Máximo"
                type="number"
                min="0"
                step="0.01"
                required
                value={edicao.valorMaximo}
                disabled={salvando}
                onChange={(e) => setEdicao({ ...edicao, valorMaximo: e.target.value })}
              />
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
        titulo="Apagar validador"
        mensagem={excluindo ? `Deseja realmente apagar o validador ${excluindo.nome}?` : ''}
        rotuloConfirmar="Apagar"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
