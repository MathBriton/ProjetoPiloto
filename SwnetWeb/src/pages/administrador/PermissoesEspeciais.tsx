import { useEffect, useState } from 'react'
import { RotateCcw, Save, ShieldCheck, UserSearch } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, EstadoVazio, TituloPagina } from '@/components/comuns/Comuns'
import { Chave, Selecao } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { ToastStack, useToast } from '@/components/Toast'
import {
  obterPermissoesEspeciais,
  salvarPermissoesEspeciais
} from '@/services/administrador/permissoesEspeciaisApi'
import type { PermissoesEspeciais as Permissoes, UsuarioSwnet } from './tipos'

/** Refatoração de Permissoes-Especiais/PermissaoUsuario.asp. */

/** Flags booleanas da CRM_DTI, na mesma ordem dos switches do legado */
const FLAGS: { campo: keyof Permissoes; rotulo: string }[] = [
  { campo: 'verTodasFncs', rotulo: 'Ver Todas FNCs' },
  { campo: 'alteraPo', rotulo: 'Altera PO' },
  { campo: 'alteraPa', rotulo: 'Altera PA' },
  { campo: 'ativaUsuario', rotulo: 'Ativa Usuário' },
  { campo: 'validaPreCadastroProduto', rotulo: 'Valida Pré-Cadastro Produto' },
  { campo: 'eTools', rotulo: 'E-Tools' },
  { campo: 'nivelLimiteCredito', rotulo: 'Nível Limite Crédito' },
  { campo: 'nivelControladoria', rotulo: 'Nível Controladoria' },
  { campo: 'nivelE2Open', rotulo: 'Nível E2OPEN' },
  { campo: 'nivelControlesInternos', rotulo: 'Nível Controles Internos' },
  { campo: 'nivelDiretorComercial', rotulo: 'Nível Diretor Comercial' },
  { campo: 'nivelFinanceiro', rotulo: 'Nível Financeiro' },
  { campo: 'nivelFiscal', rotulo: 'Nível Fiscal' }
]

export function PermissoesEspeciais() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [permissoes, setPermissoes] = useState<Permissoes | null>(null)
  const [carregando, setCarregando] = useState(false)
  const [salvando, setSalvando] = useState(false)

  useEffect(() => {
    if (!usuario) {
      setPermissoes(null)
      return
    }

    let cancelado = false
    setCarregando(true)
    obterPermissoesEspeciais(usuario.id)
      .then((dados) => {
        if (!cancelado) setPermissoes(dados)
      })
      .finally(() => {
        if (!cancelado) setCarregando(false)
      })

    return () => {
      cancelado = true
    }
  }, [usuario])

  function mudar<C extends keyof Permissoes>(campo: C, valor: Permissoes[C]) {
    setPermissoes((atual) => (atual ? { ...atual, [campo]: valor } : atual))
  }

  async function recarregar() {
    if (!usuario) return
    setCarregando(true)
    try {
      setPermissoes(await obterPermissoesEspeciais(usuario.id))
    } finally {
      setCarregando(false)
    }
  }

  async function salvar() {
    if (!permissoes) return

    // O legado exige Nível Altera Pedido e Permissão preenchidos (campos required)
    if (!permissoes.nivelAlteraPedido || !permissoes.permissao) {
      mostrarToast('Informe o Nível Altera Pedido e a Permissão.', 'warning')
      return
    }

    setSalvando(true)
    try {
      const resultado = await salvarPermissoesEspeciais(permissoes)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Permissões Especiais" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Permissões de Usuário" icone={ShieldCheck}>
          <div className="mb-4">
            <SeletorUsuario selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />
          </div>

          {!usuario && (
            <EstadoVazio
              icone={UserSearch}
              titulo="Selecione um usuário"
              descricao="As permissões especiais controlam níveis de aprovação e acessos por área."
            />
          )}

          {usuario && carregando && (
            <p className="py-10 text-center text-sm text-slate-400">Carregando permissões...</p>
          )}

          {permissoes && !carregando && (
            <form
              onSubmit={(e) => {
                e.preventDefault()
                salvar()
              }}
            >
              <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                <Selecao
                  id="nivelAlteraPedido"
                  rotulo="Nível Altera Pedido"
                  required
                  value={permissoes.nivelAlteraPedido}
                  disabled={salvando}
                  onChange={(e) => mudar('nivelAlteraPedido', e.target.value as Permissoes['nivelAlteraPedido'])}
                >
                  <option value="">Selecione</option>
                  <option value="0">0 - Nível 0</option>
                  <option value="14">14 - Nível 14</option>
                </Selecao>

                <Selecao
                  id="nivelSolicitacaoAlteracaoPedido"
                  rotulo="Nível Solic. Alteração Pedido"
                  value={permissoes.nivelSolicitacaoAlteracaoPedido}
                  disabled={salvando}
                  onChange={(e) =>
                    mudar(
                      'nivelSolicitacaoAlteracaoPedido',
                      e.target.value as Permissoes['nivelSolicitacaoAlteracaoPedido']
                    )
                  }
                >
                  <option value="">Nenhum</option>
                  <option value="V">Validador</option>
                  <option value="P">PCP</option>
                </Selecao>

                <Selecao
                  id="nivelSpd"
                  rotulo="Nível SPD"
                  value={permissoes.nivelSpd}
                  disabled={salvando}
                  onChange={(e) => mudar('nivelSpd', e.target.value as Permissoes['nivelSpd'])}
                >
                  <option value="">Nenhum</option>
                  <option value="C">Comercial</option>
                  <option value="D">Desenvolvimento</option>
                </Selecao>

                <Selecao
                  id="permissao"
                  rotulo="Permissão"
                  required
                  value={permissoes.permissao}
                  disabled={salvando}
                  onChange={(e) => mudar('permissao', e.target.value as Permissoes['permissao'])}
                >
                  <option value="">Selecione</option>
                  <option value="C">C - Validador</option>
                  <option value="G">G - Representante</option>
                  <option value="L">L - Usuário Interno</option>
                </Selecao>
              </div>

              <div className="mt-5 grid gap-2 sm:grid-cols-2">
                {FLAGS.map(({ campo, rotulo }) => (
                  <Chave
                    key={campo}
                    rotulo={rotulo}
                    marcado={permissoes[campo] as boolean}
                    desabilitado={salvando}
                    aoAlternar={(v) => mudar(campo, v as Permissoes[typeof campo])}
                  />
                ))}
              </div>

              <div className="mt-4 flex justify-end gap-2">
                <Botao variante="neutro" icone={RotateCcw} aoClicar={recarregar} desabilitado={salvando}>
                  Descartar
                </Botao>
                <Botao tipo="submit" variante="sucesso" icone={Save} desabilitado={salvando}>
                  {salvando ? 'Salvando...' : 'Salvar'}
                </Botao>
              </div>
            </form>
          )}
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
