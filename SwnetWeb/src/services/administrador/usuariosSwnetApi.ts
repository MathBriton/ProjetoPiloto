import type {
  DadosAlteracaoUsuario,
  DadosCadastroUsuario,
  UsuarioSwnet
} from '@/pages/administrador/tipos'
import { atraso, gerarIdUsuario, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * CRUD de usuários do SWnet (CadUsuario.asp, AltUsuario.asp e ExcUsuario.asp).
 * Mock em memória — trocar por chamadas ao SwNet.Api quando os endpoints existirem.
 */

export interface ResultadoOperacao {
  sucesso: boolean
  mensagem: string
}

/** Lista para os autocompletes; por padrão inclui inativos (o legado busca todos). */
export async function listarUsuarios(termo = '', apenasAtivos = false): Promise<UsuarioSwnet[]> {
  const busca = termo.trim().toLowerCase()
  const resultado = usuariosSwnet
    .filter((u) => (apenasAtivos ? u.ativo : true))
    .filter(
      (u) =>
        !busca ||
        u.nome.toLowerCase().includes(busca) ||
        u.login.toLowerCase().includes(busca) ||
        u.matricula.includes(busca) ||
        u.email.toLowerCase().includes(busca)
    )
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'))

  return atraso(resultado.map((u) => ({ ...u })), 200)
}

export async function obterUsuario(id: number): Promise<UsuarioSwnet | undefined> {
  const encontrado = usuariosSwnet.find((u) => u.id === id)
  return atraso(encontrado ? { ...encontrado } : undefined, 150)
}

/**
 * Regras herdadas do CadUsuario.asp: matrícula + filial não podem se repetir
 * entre usuários ativos, e a matrícula tem 6 dígitos (exceto filiais 05 e 06).
 */
export function validarMatricula(matricula: string, empresa: string): string | null {
  const valor = matricula.trim()

  if (valor.length < 6 && empresa !== '05' && empresa !== '06') {
    return 'O campo matrícula deve conter 6 dígitos.'
  }
  if (usuariosSwnet.some((u) => u.ativo && u.matricula === valor && u.empresa === empresa)) {
    return 'Usuário já cadastrado para esta matrícula e empresa.'
  }
  return null
}

export function loginEmUso(login: string, ignorarId?: number) {
  const alvo = login.trim().toLowerCase()
  return usuariosSwnet.some((u) => u.login.toLowerCase() === alvo && u.id !== ignorarId)
}

export async function criarUsuario(dados: DadosCadastroUsuario): Promise<ResultadoOperacao> {
  const erroMatricula = validarMatricula(dados.matricula, dados.empresa)
  if (erroMatricula) return atraso({ sucesso: false, mensagem: erroMatricula }, 250)

  if (loginEmUso(dados.login)) {
    return atraso({ sucesso: false, mensagem: 'Já existe um usuário com este login.' }, 250)
  }

  const hoje = new Date().toISOString().slice(0, 10)

  usuariosSwnet.push({
    id: gerarIdUsuario(),
    matricula: dados.matricula.trim(),
    nome: dados.nome.trim().toUpperCase(),
    login: dados.login.trim(),
    userAd: `${dados.prefixoAd}\\${dados.login.trim()}`,
    email: dados.email.trim().toLowerCase(),
    sexo: dados.sexo,
    nascimento: dados.nascimento,
    ramal: dados.ramal.trim(),
    tipo: dados.tipo,
    empresa: dados.empresa,
    funcao: '',
    representante: dados.representante,
    bloqueado: false,
    ativo: true,
    dataCadastro: hoje,
    ultimoAcesso: null,
    dataBloqueio: null
  })

  return atraso(
    {
      sucesso: true,
      mensagem: dados.enviarEmail
        ? 'Usuário incluído com sucesso! E-mail enviado.'
        : 'Usuário incluído com sucesso!'
    },
    450
  )
}

export async function alterarUsuario(dados: DadosAlteracaoUsuario): Promise<ResultadoOperacao> {
  const indice = usuariosSwnet.findIndex((u) => u.id === dados.id)
  if (indice < 0) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 250)

  if (loginEmUso(dados.login, dados.id)) {
    return atraso({ sucesso: false, mensagem: 'Já existe outro usuário com este login.' }, 250)
  }

  const atual = usuariosSwnet[indice]

  usuariosSwnet[indice] = {
    ...atual,
    matricula: dados.matricula.trim(),
    nome: dados.nome.trim().toUpperCase(),
    login: dados.login.trim(),
    userAd: `${dados.prefixoAd}\\${dados.login.trim()}`,
    email: dados.email.trim().toLowerCase(),
    sexo: dados.sexo,
    nascimento: dados.nascimento,
    ramal: dados.ramal.trim(),
    tipo: dados.tipo,
    empresa: dados.empresa,
    representante: dados.representante,
    bloqueado: dados.bloqueado,
    // Desbloquear limpa a data de bloqueio, como no UPDATE do AltUsuario.asp
    dataBloqueio: dados.bloqueado ? (atual.dataBloqueio ?? new Date().toISOString().slice(0, 10)) : null
  }

  return atraso({ sucesso: true, mensagem: 'Usuário alterado com sucesso!' }, 450)
}

/**
 * ExcUsuario.asp: "desativar" marca CRM_D_E_L_E_T = '*' e "ativar" limpa a marca
 * junto com a DataBloqueio.
 */
export async function alternarSituacao(id: number): Promise<ResultadoOperacao & { usuario?: UsuarioSwnet }> {
  const indice = usuariosSwnet.findIndex((u) => u.id === id)
  if (indice < 0) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 250)

  const desativando = usuariosSwnet[indice].ativo

  usuariosSwnet[indice] = {
    ...usuariosSwnet[indice],
    ativo: !desativando,
    dataBloqueio: desativando ? new Date().toISOString().slice(0, 10) : null
  }

  return atraso(
    {
      sucesso: true,
      mensagem: desativando ? 'Usuário desativado com sucesso!' : 'Usuário ativado com sucesso!',
      usuario: { ...usuariosSwnet[indice] }
    },
    400
  )
}
