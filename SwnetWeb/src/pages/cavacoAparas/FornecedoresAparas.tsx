import { useCallback, useEffect, useMemo, useState } from 'react'
import { Factory, Pencil, Plus, Save, Trash2 } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, CampoPesquisa, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Paginacao, Tabela, type OrdenacaoTabela } from '@/components/comuns/Tabela'
import { Entrada } from '@/components/comuns/Campos'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  excluirFornecedor,
  formatarCnpj,
  listarFornecedores,
  salvarFornecedor
} from '@/services/cavacoAparas/aparasApi'
import type { DadosFornecedorAparas, FornecedorAparas } from './tipos'

/** Refatoração de Cavaco-Aparas/Cadastro-De-Fornecedor-Aparas (grid_fornecedores.asp). */

const FORNECEDOR_VAZIO: DadosFornecedorAparas = { codigo: '', cnpj: '', nome: '' }

export function FornecedoresAparas() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [fornecedores, setFornecedores] = useState<FornecedorAparas[]>([])
  const [carregando, setCarregando] = useState(true)
  const [termo, setTermo] = useState('')
  const [ordenacao, setOrdenacao] = useState<OrdenacaoTabela>({ campo: 'NOME', ascendente: true })
  const [pagina, setPagina] = useState(1)
  const [tamanho, setTamanho] = useState(50)

  const [edicao, setEdicao] = useState<{ dados: DadosFornecedorAparas; editando: boolean } | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<FornecedorAparas | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setFornecedores(await listarFornecedores(termo))
    } finally {
      setCarregando(false)
    }
  }, [termo])

  useEffect(() => {
    const atraso = setTimeout(carregar, 250)
    return () => clearTimeout(atraso)
  }, [carregar])

  // Volta para a primeira página sempre que a busca ou o tamanho mudam
  useEffect(() => {
    setPagina(1)
  }, [termo, tamanho])

  const ordenados = useMemo(() => {
    const campo = ordenacao.campo.toLowerCase() as keyof FornecedorAparas
    return [...fornecedores].sort((a, b) => {
      const comparacao = String(a[campo]).localeCompare(String(b[campo]), 'pt-BR', { numeric: true })
      return ordenacao.ascendente ? comparacao : -comparacao
    })
  }, [fornecedores, ordenacao])

  const totalPaginas = Math.max(1, Math.ceil(ordenados.length / tamanho))
  const visiveis = ordenados.slice((pagina - 1) * tamanho, pagina * tamanho)

  function alternarOrdenacao(campo: string) {
    setOrdenacao((atual) =>
      atual.campo === campo ? { campo, ascendente: !atual.ascendente } : { campo, ascendente: true }
    )
  }

  async function salvar() {
    if (!edicao) return

    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarFornecedor(edicao.dados, edicao.editando)
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

    const resultado = await excluirFornecedor(excluindo.codigo)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  function mudar<C extends keyof DadosFornecedorAparas>(campo: C, valor: string) {
    setEdicao((atual) => (atual ? { ...atual, dados: { ...atual.dados, [campo]: valor } } : atual))
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Fornecedores de Aparas" />

      <Painel
        titulo="Lista de Fornecedores"
        icone={Factory}
        acoes={
          <>
            <CampoPesquisa
              valor={termo}
              aoMudar={setTermo}
              placeholder="Nome, CNPJ ou Código SAP..."
              larguraClasse="w-64"
            />
            <Botao
              icone={Plus}
              aoClicar={() => {
                setErro('')
                setEdicao({ dados: FORNECEDOR_VAZIO, editando: false })
              }}
            >
              Novo
            </Botao>
          </>
        }
      >
        <Tabela
          itens={visiveis}
          chave={(fornecedor) => fornecedor.codigo}
          carregando={carregando}
          ordenacao={ordenacao}
          aoOrdenar={alternarOrdenacao}
          colunas={[
            { cabecalho: 'Código SAP', conteudo: (f) => f.codigo, ordenavelPor: 'CODIGO', className: 'w-32' },
            { cabecalho: 'CNPJ', conteudo: (f) => formatarCnpj(f.cnpj), ordenavelPor: 'CNPJ', className: 'w-48' },
            { cabecalho: 'Nome', conteudo: (f) => f.nome, ordenavelPor: 'NOME' }
          ]}
          acoes={(fornecedor) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar fornecedor"
                aoClicar={() => {
                  setErro('')
                  setEdicao({ dados: { ...fornecedor }, editando: true })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Apagar fornecedor"
                tom="perigo"
                aoClicar={() => setExcluindo(fornecedor)}
              />
            </>
          )}
        />

        <Paginacao
          pagina={pagina}
          totalPaginas={totalPaginas}
          totalRegistros={ordenados.length}
          tamanho={tamanho}
          aoMudarPagina={setPagina}
          aoMudarTamanho={setTamanho}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.editando ? 'Editar Fornecedor' : 'Novo Fornecedor'}
        icone={Factory}
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

            <Entrada
              id="codigoFornecedor"
              rotulo="Código SAP"
              required
              maxLength={10}
              placeholder="Somente números"
              value={edicao.dados.codigo}
              // O código é a chave do registro: não muda na edição
              disabled={salvando || edicao.editando}
              onChange={(e) => mudar('codigo', e.target.value.replace(/\D/g, ''))}
            />

            <Entrada
              id="cnpjFornecedor"
              rotulo="CNPJ"
              required
              maxLength={14}
              placeholder="14 caracteres, somente números"
              value={edicao.dados.cnpj}
              disabled={salvando}
              ajuda={
                edicao.dados.cnpj.length === 14 ? formatarCnpj(edicao.dados.cnpj) : 'Faltam dígitos'
              }
              onChange={(e) => mudar('cnpj', e.target.value.replace(/\D/g, ''))}
            />

            <Entrada
              id="nomeFornecedor"
              rotulo="Razão Social"
              required
              value={edicao.dados.nome}
              disabled={salvando}
              onChange={(e) => mudar('nome', e.target.value)}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Apagar fornecedor"
        mensagem={excluindo ? `Deseja realmente apagar o fornecedor ${excluindo.nome}?` : ''}
        rotuloConfirmar="Apagar"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
