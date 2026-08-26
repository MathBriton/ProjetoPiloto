import type { OpcaoSimples, Sistema, UsuarioSwnet } from '@/pages/administrador/tipos'

/**
 * Base de dados fictícia compartilhada pelas telas do menu Administrador.
 *
 * Enquanto a SwNet.Api não expõe os endpoints, tudo vive em memória — os
 * dados se perdem ao recarregar a página. Nenhuma regra de negócio real
 * mora aqui: é só o suficiente para exercitar os fluxos no frontend.
 */

/** Empresas/filiais — mesma lista do RelatorioPermissaoUsuarios.asp */
export const EMPRESAS: OpcaoSimples[] = [
  { id: 0, nome: 'Pirapetinga' },
  { id: 4, nome: 'Uberaba' },
  { id: 5, nome: 'Bento' },
  { id: 6, nome: 'Maranguape' },
  { id: 7, nome: 'Saquarema' },
  { id: 8, nome: 'Araçatuba' },
  { id: 9, nome: 'Blumenau' },
  { id: 10, nome: 'Pacajus' },
  { id: 11, nome: 'Porto Feliz' },
  { id: 12, nome: 'Três Barras' },
  { id: 13, nome: 'Florestal' }
]

/** Código de empresa no formato do banco legado (EMP_CODIGO, 2 dígitos) */
export function codigoEmpresa(id: number) {
  return String(id).padStart(2, '0')
}

export function nomeEmpresa(codigo: string) {
  return EMPRESAS.find((e) => codigoEmpresa(e.id) === codigo)?.nome ?? codigo
}

/** Espelha a tabela TipoUsuario do legado */
export const TIPOS_USUARIO = [
  { codigo: 'I', nome: 'Interno' },
  { codigo: 'R', nome: 'Representante' },
  { codigo: 'C', nome: 'Cliente' },
  { codigo: 'F', nome: 'Fornecedor' },
  { codigo: 'T', nome: 'Terceiro' }
]

export function nomeTipoUsuario(codigo: string) {
  return TIPOS_USUARIO.find((t) => t.codigo === codigo)?.nome ?? codigo
}

/**
 * Sistemas varridos pelo Cancelamento/Revisão de usuários.
 * As chaves são as mesmas usadas pelo CancelamentoUsuarioRepositorio.asp.
 */
export const SISTEMAS: Sistema[] = [
  { id: 'swnet', nome: 'SWNET' },
  { id: 'benTrimbox', nome: 'TRIMBOX - Bento' },
  { id: 'forTrimbox', nome: 'TRIMBOX - Maranguape' },
  { id: 'pirTrimbox', nome: 'TRIMBOX - Pirapetinga' },
  { id: 'saqTrimbox', nome: 'TRIMBOX - Saquarema' },
  { id: 'ubeTrimbox', nome: 'TRIMBOX - Uberaba' },
  { id: 'benTrimpaper', nome: 'TRIMPAPER - Bento' },
  { id: 'pirTrimpaper', nome: 'TRIMPAPER - Pirapetinga' },
  { id: 'ubeTrimpaper', nome: 'TRIMPAPER - Uberaba' },
  { id: 'benDadosadv', nome: 'PROTHEUS - Bento' },
  { id: 'forDadosadv', nome: 'PROTHEUS - Maranguape' },
  { id: 'pirDadosadv', nome: 'PROTHEUS - Pirapetinga' },
  { id: 'saqDadosadv', nome: 'PROTHEUS - Saquarema' }
]

export function nomeSistema(id: string) {
  return SISTEMAS.find((s) => s.id === id)?.nome ?? id
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

export const usuariosSwnet: UsuarioSwnet[] = [
  {
    id: 1001,
    matricula: '000123',
    nome: 'ADMINISTRADOR (FICTÍCIO)',
    login: 'admin',
    userAd: 'GROUP\\admin',
    email: 'admin@smurfitwestrock.com.br',
    sexo: 'M',
    nascimento: '1985-03-12',
    ramal: '2010',
    tipo: 'I',
    empresa: '00',
    funcao: 'Analista de Sistemas',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(900),
    ultimoAcesso: iso(1),
    dataBloqueio: null
  },
  {
    id: 1002,
    matricula: '000456',
    nome: 'JOÃO SILVA',
    login: 'jsilva',
    userAd: 'GROUP\\jsilva',
    email: 'joao.silva@smurfitwestrock.com.br',
    sexo: 'M',
    nascimento: '1990-07-21',
    ramal: '2044',
    tipo: 'I',
    empresa: '00',
    funcao: 'Assistente de PCP',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(620),
    ultimoAcesso: iso(4),
    dataBloqueio: null
  },
  {
    id: 1003,
    matricula: '000789',
    nome: 'MARIA ALVES',
    login: 'malves',
    userAd: 'WESTROCK\\malves',
    email: 'maria.alves@smurfitwestrock.com.br',
    sexo: 'F',
    nascimento: '1988-11-02',
    ramal: '2071',
    tipo: 'I',
    empresa: '05',
    funcao: 'Supervisora Comercial',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(480),
    ultimoAcesso: iso(75),
    dataBloqueio: null
  },
  {
    id: 1004,
    matricula: '000321',
    nome: 'CARLOS PEREIRA',
    login: 'cpereira',
    userAd: 'GROUP\\cpereira',
    email: 'carlos.pereira@smurfitwestrock.com.br',
    sexo: 'M',
    nascimento: '1979-01-30',
    ramal: '2099',
    tipo: 'R',
    empresa: '04',
    funcao: 'Representante Comercial',
    representante: true,
    bloqueado: true,
    ativo: true,
    dataCadastro: iso(1200),
    ultimoAcesso: iso(190),
    dataBloqueio: iso(30)
  },
  {
    id: 1005,
    matricula: '000654',
    nome: 'FERNANDA SOUZA',
    login: 'fsouza',
    userAd: 'GROUP\\fsouza',
    email: 'fernanda.souza@smurfitwestrock.com.br',
    sexo: 'F',
    nascimento: '1993-05-17',
    ramal: '2033',
    tipo: 'I',
    empresa: '07',
    funcao: 'Analista Fiscal',
    representante: false,
    bloqueado: false,
    ativo: false,
    dataCadastro: iso(700),
    ultimoAcesso: iso(240),
    dataBloqueio: iso(120)
  },
  {
    id: 1006,
    matricula: '000987',
    nome: 'RICARDO NUNES',
    login: 'rnunes',
    userAd: 'WESTROCK\\rnunes',
    email: 'ricardo.nunes@smurfitwestrock.com.br',
    sexo: 'M',
    nascimento: '1982-09-09',
    ramal: '2088',
    tipo: 'T',
    empresa: '06',
    funcao: 'Consultor Externo',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(310),
    ultimoAcesso: null,
    dataBloqueio: null
  },
  {
    id: 1007,
    matricula: '001122',
    nome: 'PATRÍCIA LIMA',
    login: 'plima',
    userAd: 'GROUP\\plima',
    email: 'patricia.lima@smurfitwestrock.com.br',
    sexo: 'F',
    nascimento: '1995-12-25',
    ramal: '2012',
    tipo: 'I',
    empresa: '00',
    funcao: 'Analista de Controladoria',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(210),
    ultimoAcesso: iso(12),
    dataBloqueio: null
  },
  {
    id: 1008,
    matricula: '003344',
    nome: 'BRUNO CARVALHO',
    login: 'bcarvalho',
    userAd: 'GROUP\\bcarvalho',
    email: 'bruno.carvalho@smurfitwestrock.com.br',
    sexo: 'M',
    nascimento: '1991-04-08',
    ramal: '2055',
    tipo: 'I',
    empresa: '12',
    funcao: 'Coordenador de Expedição',
    representante: false,
    bloqueado: false,
    ativo: true,
    dataCadastro: iso(150),
    ultimoAcesso: iso(96),
    dataBloqueio: null
  }
]

/** Próximo CRM_R_E_C_N_O — no legado é identity do banco. */
let proximoId = 1009
export function gerarIdUsuario() {
  return proximoId++
}

export function atraso<T>(valor: T, ms = 300): Promise<T> {
  return new Promise((resolve) => setTimeout(() => resolve(valor), ms))
}

/** Dias corridos desde a data informada; sem acesso registrado, conta do cadastro. */
export function diasDesde(dataIso: string | null, fallbackIso: string) {
  return Math.floor((Date.now() - new Date(dataIso ?? fallbackIso).getTime()) / 86_400_000)
}

/** dd/mm/aaaa a partir de um ISO — o legado grava e exibe nesse formato. */
export function formatarData(dataIso: string | null) {
  if (!dataIso) return '—'
  const [ano, mes, dia] = dataIso.slice(0, 10).split('-')
  return `${dia}/${mes}/${ano}`
}
