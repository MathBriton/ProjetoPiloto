import { useCallback, useEffect, useState } from 'react'
import { Package, Save, Search } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao } from '@/components/comuns/Campos'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  STATUS_PRODUTO,
  alterarStatusProdutos,
  descricaoStatusProduto,
  listarProdutos,
  type Produto
} from '@/services/desenvolvimento/desenvolvimentoApi'

/** Refatoração de Desenvolvimento/Alteracao-De-Produtos. */

const TOM_STATUS: Record<string, 'sucesso' | 'perigo' | 'info' | 'alerta' | 'neutro'> = {
  A: 'sucesso',
  C: 'perigo',
  P: 'info',
  V: 'info',
  W: 'sucesso',
  E: 'alerta'
}

export function AlteracaoProdutos() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [termo, setTermo] = useState('')
  const [termoAplicado, setTermoAplicado] = useState('')
  const [produtos, setProdutos] = useState<Produto[]>([])
  const [carregando, setCarregando] = useState(true)
  const [selecionados, setSelecionados] = useState<Set<string>>(new Set())
  const [novoStatus, setNovoStatus] = useState('')
  const [confirmando, setConfirmando] = useState(false)
  const [salvando, setSalvando] = useState(false)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setProdutos(await listarProdutos(termoAplicado))
    } finally {
      setCarregando(false)
    }
  }, [termoAplicado])

  useEffect(() => {
    carregar()
  }, [carregar])

  function alternar(codigo: string) {
    setSelecionados((atual) => {
      const novo = new Set(atual)
      if (novo.has(codigo)) novo.delete(codigo)
      else novo.add(codigo)
      return novo
    })
  }

  const todosMarcados = produtos.length > 0 && produtos.every((p) => selecionados.has(p.codigo))

  async function confirmar() {
    setConfirmando(false)
    setSalvando(true)
    try {
      const resultado = await alterarStatusProdutos([...selecionados], novoStatus)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setSelecionados(new Set())
        setNovoStatus('')
        await carregar()
      }
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Alterar Produtos" />

      <Painel titulo="Produtos" icone={Package}>
        <form
          className="mb-4 grid items-end gap-4 sm:grid-cols-12"
          onSubmit={(e) => {
            e.preventDefault()
            setTermoAplicado(termo)
          }}
        >
          <Entrada
            id="pesquisa"
            rotulo="Busque por código ou descrição"
            value={termo}
            disabled={salvando}
            onChange={(e) => setTermo(e.target.value)}
            larguraClasse="sm:col-span-5"
          />
          <div className="sm:col-span-2">
            <Botao tipo="submit" icone={Search} desabilitado={carregando || salvando}>
              Pesquisar
            </Botao>
          </div>

          <Selecao
            id="novoStatus"
            rotulo="Novo status"
            value={novoStatus}
            disabled={salvando}
            onChange={(e) => setNovoStatus(e.target.value)}
            larguraClasse="sm:col-span-3"
          >
            <option value="">Selecione...</option>
            {STATUS_PRODUTO.map((item) => (
              <option key={item.codigo} value={item.codigo}>
                {item.descricao} - {item.codigo}
              </option>
            ))}
          </Selecao>

          <div className="sm:col-span-2">
            <Botao
              variante="sucesso"
              icone={Save}
              aoClicar={() => setConfirmando(true)}
              desabilitado={selecionados.size === 0 || !novoStatus || salvando}
            >
              Alterar ({selecionados.size})
            </Botao>
          </div>
        </form>

        <Tabela
          itens={produtos}
          chave={(produto) => produto.codigo}
          carregando={carregando}
          mensagemVazia="Nenhum produto encontrado"
          colunas={[
            {
              cabecalho: (
                <input
                  type="checkbox"
                  checked={todosMarcados}
                  onChange={() =>
                    setSelecionados(todosMarcados ? new Set() : new Set(produtos.map((p) => p.codigo)))
                  }
                  aria-label="Selecionar todos os produtos"
                  className="h-4 w-4 accent-swnet-600"
                />
              ),
              conteudo: (p) => (
                <input
                  type="checkbox"
                  checked={selecionados.has(p.codigo)}
                  onChange={() => alternar(p.codigo)}
                  aria-label={`Selecionar ${p.codigo}`}
                  className="h-4 w-4 accent-swnet-600"
                />
              ),
              className: 'w-12'
            },
            { cabecalho: 'Cliente', conteudo: (p) => p.cliente },
            { cabecalho: 'Código', conteudo: (p) => p.codigo },
            { cabecalho: 'Descrição', conteudo: (p) => p.descricao },
            {
              cabecalho: 'Status',
              conteudo: (p) => (
                <Etiqueta texto={descricaoStatusProduto(p.status)} tom={TOM_STATUS[p.status] ?? 'neutro'} />
              )
            }
          ]}
        />
      </Painel>

      <ConfirmDialog
        aberto={confirmando}
        titulo="Alterar status dos produtos"
        mensagem={`${selecionados.size} produto(s) passarão para "${descricaoStatusProduto(novoStatus)}". Deseja continuar?`}
        rotuloConfirmar="Alterar"
        aoConfirmar={confirmar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
