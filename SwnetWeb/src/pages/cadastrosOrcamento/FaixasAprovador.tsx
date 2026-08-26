import { useCallback, useEffect, useState } from 'react'
import { BadgePercent, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { excluirFaixa, listarFaixas, salvarFaixa } from '@/services/cadastrosOrcamento/orcamentoApi'
import { EMPRESAS, codigoEmpresa } from '@/services/comuns/baseMock'
import type { DadosFaixaAprovador, FaixaAprovador, TipoProduto } from './tipos'

/** Refatoração de Cadastros-Orçamento/Faixas-De-Aprovador-De-Orcamento. */

const OPCOES_PLANTA = EMPRESAS.map((empresa) => ({
  valor: codigoEmpresa(empresa.id),
  rotulo: empresa.nome
}))

const TIPOS_PRODUTO: { valor: TipoProduto; rotulo: string }[] = [
  { valor: 'CX', rotulo: 'Caixa' },
  { valor: 'CH', rotulo: 'Chapa' },
  { valor: 'PA', rotulo: 'Papel' },
  { valor: 'CJ', rotulo: 'Conjunto' }
]

const FAIXA_VAZIA: DadosFaixaAprovador = {
  planta: '',
  tipoProduto: '',
  valorMinKg: '',
  valorMaxKg: '',
  valorMinM2: '',
  valorMaxM2: '',
  quantidadeAprovadores: '1'
}

function rotuloTipo(tipo: TipoProduto) {
  return TIPOS_PRODUTO.find((t) => t.valor === tipo)?.rotulo ?? tipo
}

const moeda = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 })

export function FaixasAprovador() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [faixas, setFaixas] = useState<FaixaAprovador[]>([])
  const [carregando, setCarregando] = useState(true)
  const [filtroPlanta, setFiltroPlanta] = useState('TODOS')

  const [edicao, setEdicao] = useState<DadosFaixaAprovador | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<FaixaAprovador | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setFaixas(await listarFaixas(filtroPlanta))
    } finally {
      setCarregando(false)
    }
  }, [filtroPlanta])

  useEffect(() => {
    carregar()
  }, [carregar])

  function abrirEdicao(faixa: FaixaAprovador) {
    setErro('')
    setEdicao({
      id: faixa.id,
      planta: faixa.planta,
      tipoProduto: faixa.tipoProduto,
      valorMinKg: String(faixa.valorMinKg),
      valorMaxKg: String(faixa.valorMaxKg),
      valorMinM2: String(faixa.valorMinM2),
      valorMaxM2: String(faixa.valorMaxM2),
      quantidadeAprovadores: String(faixa.quantidadeAprovadores)
    })
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarFaixa(edicao)
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

    const resultado = await excluirFaixa(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  function mudar<C extends keyof DadosFaixaAprovador>(campo: C, valor: DadosFaixaAprovador[C]) {
    setEdicao((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Faixas de Aprovadores de Orçamento" />

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
        titulo="Faixas"
        icone={BadgePercent}
        acoes={
          <Botao icone={Plus} aoClicar={() => { setErro(''); setEdicao(FAIXA_VAZIA) }}>
            Nova
          </Botao>
        }
      >
        <Tabela
          itens={faixas}
          chave={(faixa) => faixa.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (f) => f.id, className: 'w-16' },
            { cabecalho: 'Planta', conteudo: (f) => f.nomePlanta },
            {
              cabecalho: 'Tipo Produto',
              conteudo: (f) => <Etiqueta texto={rotuloTipo(f.tipoProduto)} tom="info" />
            },
            { cabecalho: 'Mín. p/Kg', conteudo: (f) => moeda.format(f.valorMinKg), alinhamento: 'direita' },
            { cabecalho: 'Máx. p/Kg', conteudo: (f) => moeda.format(f.valorMaxKg), alinhamento: 'direita' },
            { cabecalho: 'Mín. p/M²', conteudo: (f) => moeda.format(f.valorMinM2), alinhamento: 'direita' },
            { cabecalho: 'Máx. p/M²', conteudo: (f) => moeda.format(f.valorMaxM2), alinhamento: 'direita' },
            {
              cabecalho: 'Qtd. Aprovadores',
              conteudo: (f) => f.quantidadeAprovadores,
              alinhamento: 'centro'
            }
          ]}
          acoes={(faixa) => (
            <>
              <BotaoAcao icone={Pencil} titulo="Editar faixa" aoClicar={() => abrirEdicao(faixa)} />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir faixa"
                tom="perigo"
                aoClicar={() => setExcluindo(faixa)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Faixa de Aprovador' : 'Nova Faixa de Aprovador'}
        icone={BadgePercent}
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

            <div className="grid gap-4 sm:grid-cols-2">
              <Selecao
                id="plantaFaixa"
                rotulo="Planta"
                required
                value={edicao.planta}
                disabled={salvando}
                onChange={(e) => mudar('planta', e.target.value)}
              >
                <option value="">Selecione uma planta</option>
                {OPCOES_PLANTA.map((opcao) => (
                  <option key={opcao.valor} value={opcao.valor}>
                    {opcao.rotulo}
                  </option>
                ))}
              </Selecao>

              <Selecao
                id="tipoProdutoFaixa"
                rotulo="Tipo de Produto"
                required
                value={edicao.tipoProduto}
                disabled={salvando}
                onChange={(e) => mudar('tipoProduto', e.target.value as DadosFaixaAprovador['tipoProduto'])}
              >
                <option value="">Selecione um tipo</option>
                {TIPOS_PRODUTO.map((tipo) => (
                  <option key={tipo.valor} value={tipo.valor}>
                    {tipo.rotulo}
                  </option>
                ))}
              </Selecao>
            </div>

            <div className="grid gap-4 sm:grid-cols-2">
              <Entrada
                id="valorMinKg"
                rotulo="Valor Mínimo Kg"
                type="number"
                step="0.0001"
                min="0"
                required
                value={edicao.valorMinKg}
                disabled={salvando}
                onChange={(e) => mudar('valorMinKg', e.target.value)}
              />
              <Entrada
                id="valorMaxKg"
                rotulo="Valor Máximo Kg"
                type="number"
                step="0.0001"
                min="0"
                required
                value={edicao.valorMaxKg}
                disabled={salvando}
                onChange={(e) => mudar('valorMaxKg', e.target.value)}
              />
              <Entrada
                id="valorMinM2"
                rotulo="Valor Mínimo M²"
                type="number"
                step="0.0001"
                min="0"
                required
                value={edicao.valorMinM2}
                disabled={salvando}
                onChange={(e) => mudar('valorMinM2', e.target.value)}
              />
              <Entrada
                id="valorMaxM2"
                rotulo="Valor Máximo M²"
                type="number"
                step="0.0001"
                min="0"
                required
                value={edicao.valorMaxM2}
                disabled={salvando}
                onChange={(e) => mudar('valorMaxM2', e.target.value)}
              />
            </div>

            <Entrada
              id="quantidadeAprovadores"
              rotulo="Qtd. Aprovadores"
              type="number"
              min="1"
              step="1"
              required
              value={edicao.quantidadeAprovadores}
              disabled={salvando}
              onChange={(e) => mudar('quantidadeAprovadores', e.target.value)}
              larguraClasse="max-w-xs"
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir faixa"
        mensagem={
          excluindo
            ? `Deseja realmente excluir a faixa de ${rotuloTipo(excluindo.tipoProduto)} da planta ${excluindo.nomePlanta}?`
            : ''
        }
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
