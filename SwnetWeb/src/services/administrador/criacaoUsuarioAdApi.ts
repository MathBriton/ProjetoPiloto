import { EMPRESAS, atraso, codigoEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * Criacao-Usuarios-AD.
 *
 * ATENÇÃO: a pasta legada só traz o `index.html` (o hub com os dois cards) —
 * as telas `creation_user.html` e `card.html` que ele linka NÃO estão no
 * material de refatoração. Os campos abaixo são uma reconstrução plausível
 * de uma solicitação de conta no AD e devem ser conferidos contra a tela
 * original antes de ligar no backend.
 */

export interface SolicitacaoUsuarioAd {
  nome: string
  sobrenome: string
  matricula: string
  dominio: 'GROUP' | 'WESTROCK'
  login: string
  email: string
  planta: string
  departamento: string
  cargo: string
  gestor: string
  dataAdmissao: string
  grupos: string[]
  observacoes: string
}

export const DEPARTAMENTOS = [
  'Comercial',
  'Controladoria',
  'Expedição',
  'Financeiro',
  'Fiscal',
  'Manutenção',
  'PCP',
  'Produção',
  'Qualidade (SGQ)',
  'Recursos Humanos',
  'Suprimentos',
  'TI'
]

export const GRUPOS_AD = [
  'SWNET_USUARIOS',
  'SWNET_ADMIN',
  'TRIMBOX_OPERACAO',
  'PROTHEUS_CONSULTA',
  'VPN_CORPORATIVA',
  'FILESERVER_LEITURA',
  'FILESERVER_ESCRITA',
  'IMPRESSAO_CORPORATIVA'
]

export const SOLICITACAO_VAZIA: SolicitacaoUsuarioAd = {
  nome: '',
  sobrenome: '',
  matricula: '',
  dominio: 'GROUP',
  login: '',
  email: '',
  planta: '',
  departamento: '',
  cargo: '',
  gestor: '',
  dataAdmissao: '',
  grupos: ['SWNET_USUARIOS'],
  observacoes: ''
}

/** Sugere `primeironome.sobrenome`, padrão usado nos logins da base */
export function sugerirLogin(nome: string, sobrenome: string) {
  const normalizar = (valor: string) =>
    valor
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z]/g, '')

  const primeiro = normalizar(nome)
  const ultimo = normalizar(sobrenome.split(/\s+/).pop() ?? '')
  if (!primeiro || !ultimo) return ''
  return `${primeiro}.${ultimo}`
}

export function sugerirEmail(login: string) {
  return login ? `${login}@smurfitwestrock.com.br` : ''
}

export function nomePlanta(codigo: string) {
  return EMPRESAS.find((e) => codigoEmpresa(e.id) === codigo)?.nome ?? ''
}

export interface FichaUsuario {
  nome: string
  login: string
  userAd: string
  email: string
  matricula: string
  planta: string
  ramal: string
  situacao: string
  dataCadastro: string
  ultimoAcesso: string | null
}

/** "Ficha de Usuário" — consulta os dados consolidados de uma conta */
export async function buscarFicha(termo: string): Promise<FichaUsuario[]> {
  const busca = termo.trim().toLowerCase()
  if (!busca) return atraso([], 150)

  const encontrados = usuariosSwnet
    .filter(
      (u) =>
        u.nome.toLowerCase().includes(busca) ||
        u.login.toLowerCase().includes(busca) ||
        u.email.toLowerCase().includes(busca) ||
        u.matricula.includes(busca)
    )
    .map<FichaUsuario>((u) => ({
      nome: u.nome,
      login: u.login,
      userAd: u.userAd,
      email: u.email,
      matricula: u.matricula,
      planta: nomePlanta(u.empresa),
      ramal: u.ramal,
      situacao: !u.ativo ? 'Inativo' : u.bloqueado ? 'Bloqueado' : 'Ativo',
      dataCadastro: u.dataCadastro,
      ultimoAcesso: u.ultimoAcesso
    }))

  return atraso(encontrados, 350)
}

export async function solicitarCriacao(
  dados: SolicitacaoUsuarioAd
): Promise<{ sucesso: boolean; mensagem: string; protocolo?: string }> {
  const obrigatorios: (keyof SolicitacaoUsuarioAd)[] = [
    'nome',
    'sobrenome',
    'login',
    'email',
    'planta',
    'departamento',
    'gestor'
  ]

  const faltando = obrigatorios.some((campo) => !String(dados[campo]).trim())
  if (faltando) {
    return atraso({ sucesso: false, mensagem: 'Preencha todos os campos obrigatórios.' }, 200)
  }

  if (dados.grupos.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos um grupo de acesso.' }, 200)
  }

  const protocolo = `AD-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`
  return atraso(
    { sucesso: true, mensagem: `Solicitação registrada com o protocolo ${protocolo}.`, protocolo },
    600
  )
}
