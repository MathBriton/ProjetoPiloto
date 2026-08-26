import type { PermissoesEspeciais } from '@/pages/administrador/tipos'
import { atraso, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * PermissaoUsuario.asp / PermissaoUsuarioRepositorio.asp.
 * As permissões especiais são colunas da própria CRM_DTI; aqui ficam num
 * mapa em memória, criado sob demanda com os valores padrão.
 */

const permissoesPorUsuario = new Map<number, PermissoesEspeciais>()

function padrao(idUsuario: number, nome: string): PermissoesEspeciais {
  return {
    idUsuario,
    nome,
    nivelAlteraPedido: '',
    nivelSolicitacaoAlteracaoPedido: '',
    nivelSpd: '',
    permissao: '',
    verTodasFncs: false,
    alteraPo: false,
    alteraPa: false,
    ativaUsuario: false,
    validaPreCadastroProduto: false,
    eTools: false,
    nivelLimiteCredito: false,
    nivelControladoria: false,
    nivelE2Open: false,
    nivelControlesInternos: false,
    nivelFinanceiro: false,
    nivelDiretorComercial: false,
    nivelFiscal: false
  }
}

// Alguns usuários já vêm com permissões para a tela não abrir sempre vazia
permissoesPorUsuario.set(1001, {
  ...padrao(1001, 'ADMINISTRADOR (FICTÍCIO)'),
  nivelAlteraPedido: '14',
  permissao: 'L',
  verTodasFncs: true,
  alteraPo: true,
  alteraPa: true,
  ativaUsuario: true,
  eTools: true,
  nivelControlesInternos: true
})

permissoesPorUsuario.set(1004, {
  ...padrao(1004, 'CARLOS PEREIRA'),
  nivelAlteraPedido: '0',
  nivelSolicitacaoAlteracaoPedido: 'V',
  permissao: 'G',
  ativaUsuario: true,
  nivelLimiteCredito: true
})

export async function obterPermissoesEspeciais(idUsuario: number): Promise<PermissoesEspeciais> {
  const nome = usuariosSwnet.find((u) => u.id === idUsuario)?.nome ?? ''
  const existente = permissoesPorUsuario.get(idUsuario)
  return atraso(existente ? { ...existente, nome } : padrao(idUsuario, nome), 250)
}

export async function salvarPermissoesEspeciais(
  dados: PermissoesEspeciais
): Promise<{ sucesso: boolean; mensagem: string }> {
  if (!dados.idUsuario) {
    return atraso({ sucesso: false, mensagem: 'Usuário é obrigatório.' }, 200)
  }
  permissoesPorUsuario.set(dados.idUsuario, { ...dados })
  return atraso({ sucesso: true, mensagem: 'Permissões atualizadas com sucesso!' }, 450)
}
