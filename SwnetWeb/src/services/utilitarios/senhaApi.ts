import { atraso, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * Utilitarios/Alterar-Senha (TrocarSenha.asp).
 *
 * As regras abaixo são as mesmas funções VBScript do legado
 * (VerificaMaiuscula, VerificaSequencia, ValidarHistorico, ...), traduzidas
 * uma a uma. O histórico das últimas 24 senhas vive na tabela
 * HistoricoUsuario; aqui é um mapa em memória.
 */

/** Caracteres especiais aceitos — lista literal do VerificaCaracterEspecial */
const ESPECIAIS = '!@#$%¨&*()-+=§?/°|\\{}[]º_;¬.'

/** Sequências proibidas — VerificaSequencia */
const SEQUENCIAS_NUMERICAS = ['012', '123', '234', '345', '456', '567', '678', '789', '8910']
const SEQUENCIAS_ALFABETICAS = [
  'abc', 'bcd', 'cde', 'def', 'efg', 'fgh', 'ghi', 'hij', 'ijk', 'jkl', 'klm', 'lmn',
  'mno', 'nop', 'opq', 'pqr', 'qrs', 'rst', 'stu', 'tuv', 'uvw', 'vwx', 'wxy', 'xyz',
  'yza', 'zab'
]

/** O legado exige 24 caracteres para CRM_TIPO = '5' e 10 para os demais */
export function tamanhoMinimo(tipoUsuario: string) {
  return tipoUsuario === '5' ? 24 : 10
}

export interface RegraSenha {
  id: string
  texto: string
  atende: boolean
}

export interface ValidacaoSenha {
  regras: RegraSenha[]
  todasAtendidas: boolean
}

function temRepeticao(senha: string) {
  const minuscula = senha.toLowerCase()
  for (let i = 0; i + 2 < minuscula.length; i++) {
    if (minuscula[i] === minuscula[i + 1] && minuscula[i] === minuscula[i + 2]) return true
  }
  return false
}

function temSequencia(senha: string) {
  const minuscula = senha.toLowerCase()
  return (
    SEQUENCIAS_NUMERICAS.some((seq) => senha.includes(seq)) ||
    SEQUENCIAS_ALFABETICAS.some((seq) => minuscula.includes(seq))
  )
}

/**
 * Valida a senha nova contra todas as regras do TrocarSenha.asp.
 * O histórico é conferido à parte, no envio, porque depende do servidor.
 */
export function validarSenha(
  senha: string,
  senhaAtual: string,
  tipoUsuario: string
): ValidacaoSenha {
  const minimo = tamanhoMinimo(tipoUsuario)

  const regras: RegraSenha[] = [
    { id: 'tamanho', texto: `No mínimo ${minimo} caracteres.`, atende: senha.length >= minimo },
    {
      id: 'diferente',
      texto: 'Diferente da senha atual.',
      atende: senha.length > 0 && senha !== senhaAtual
    },
    { id: 'maiuscula', texto: 'Pelo menos um caractere maiúsculo.', atende: /[A-Z]/.test(senha) },
    { id: 'minuscula', texto: 'Pelo menos um caractere minúsculo.', atende: /[a-z]/.test(senha) },
    {
      id: 'especial',
      texto: 'Pelo menos um caractere especial.',
      atende: [...senha].some((caractere) => ESPECIAIS.includes(caractere))
    },
    { id: 'numero', texto: 'Pelo menos um número.', atende: /[0-9]/.test(senha) },
    {
      id: 'sequencia',
      texto: 'Sem caracteres sequenciais (ex.: 123, abc).',
      atende: senha.length > 0 && !temSequencia(senha)
    },
    {
      id: 'repeticao',
      texto: 'Sem repetições (ex.: 111, aaa).',
      atende: senha.length > 0 && !temRepeticao(senha)
    },
    {
      id: 'historico',
      texto: 'Não reutilize nenhuma das últimas 24 senhas.',
      // Conferido de fato no envio; aqui só sinaliza que a regra existe
      atende: senha.length > 0
    }
  ]

  return { regras, todasAtendidas: regras.every((regra) => regra.atende) }
}

/** Medidor de força de 0 a 4, igual ao calcularForca() do legado */
export function calcularForca(senha: string) {
  if (!senha) return 0

  let pontos = 0
  if (senha.length >= 10) pontos++
  if (senha.length >= 16) pontos++
  if (/[A-Z]/.test(senha) && /[a-z]/.test(senha)) pontos++
  if (/[0-9]/.test(senha) && /[^A-Za-z0-9]/.test(senha)) pontos++
  return pontos
}

export const ROTULOS_FORCA = ['Muito fraca', 'Fraca', 'Razoável', 'Boa', 'Forte']

/** HistoricoUsuario — últimas 24 senhas por usuário */
const historico = new Map<number, string[]>()
const LIMITE_HISTORICO = 24

/** Senha fictícia do usuário mockado, para o fluxo ficar testável */
const senhasAtuais = new Map<number, string>([[1001, 'Admin@123']])

export interface ResultadoTrocaSenha {
  sucesso: boolean
  mensagem: string
}

export async function trocarSenha(
  login: string,
  senhaAtual: string,
  senhaNova: string
): Promise<ResultadoTrocaSenha> {
  const usuario = usuariosSwnet.find((u) => u.login.toLowerCase() === login.trim().toLowerCase())

  if (!usuario) {
    return atraso({ sucesso: false, mensagem: 'Senha atual inválida.' }, 400)
  }
  // Ordem de checagem do legado: desativado, bloqueado, senha atual
  if (!usuario.ativo) {
    return atraso(
      { sucesso: false, mensagem: 'Este usuário está desativado. Entre em contato com o suporte.' },
      400
    )
  }
  if (usuario.bloqueado) {
    return atraso(
      { sucesso: false, mensagem: 'Usuário bloqueado. Entre em contato com o suporte.' },
      400
    )
  }

  const esperada = senhasAtuais.get(usuario.id)
  if (esperada !== undefined && senhaAtual !== esperada) {
    return atraso({ sucesso: false, mensagem: 'Senha atual inválida.' }, 400)
  }

  const { todasAtendidas, regras } = validarSenha(senhaNova, senhaAtual, usuario.tipo)
  if (!todasAtendidas) {
    const pendente = regras.find((regra) => !regra.atende)
    return atraso({ sucesso: false, mensagem: pendente?.texto ?? 'Senha inválida.' }, 300)
  }

  const anteriores = historico.get(usuario.id) ?? []
  if (anteriores.includes(senhaNova)) {
    return atraso(
      {
        sucesso: false,
        mensagem: 'A nova senha NÃO deve ter sido utilizada nas últimas 24 vezes que você a alterou.'
      },
      300
    )
  }

  historico.set(usuario.id, [senhaAtual, ...anteriores].slice(0, LIMITE_HISTORICO))
  senhasAtuais.set(usuario.id, senhaNova)

  return atraso({ sucesso: true, mensagem: 'Senha alterada com sucesso!' }, 600)
}

/** CRM_TIPO do usuário — define o tamanho mínimo exigido */
export function tipoDoUsuario(login: string) {
  return usuariosSwnet.find((u) => u.login.toLowerCase() === login.trim().toLowerCase())?.tipo ?? ''
}
