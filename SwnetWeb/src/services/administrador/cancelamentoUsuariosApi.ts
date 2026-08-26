import type { AcaoCancelamento, ContaSistema, UsuarioInativo } from '@/pages/administrador/tipos'
import { SISTEMAS, atraso, diasDesde, nomeSistema, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * CancelamentoUsuario.asp / RevisaoUsuarios.asp.
 *
 * As duas telas varrem os mesmos bancos externos (SWNET, TRIMBOX, TRIMPAPER e
 * PROTHEUS por planta). Aqui as contas dos sistemas externos são derivadas dos
 * usuários SWnet: cada e-mail existe em alguns sistemas e não em outros,
 * reproduzindo as linhas "SEM REGISTRO" do legado.
 */

interface ContaExterna {
  id: number
  sistemaId: string
  email: string
  nome: string
  dataCadastro: string
  dataBloqueio: string | null
  ultimoAcesso: string | null
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

/** Contas fora do SWnet — em memória, para os fluxos de habilitar/desabilitar */
const contasExternas: ContaExterna[] = [
  { id: 5001, sistemaId: 'pirTrimbox', email: 'joao.silva@smurfitwestrock.com.br', nome: 'JOAO SILVA', dataCadastro: iso(600), dataBloqueio: null, ultimoAcesso: iso(9) },
  { id: 5002, sistemaId: 'pirTrimpaper', email: 'joao.silva@smurfitwestrock.com.br', nome: 'JOAO SILVA', dataCadastro: iso(590), dataBloqueio: null, ultimoAcesso: iso(120) },
  { id: 5003, sistemaId: 'pirDadosadv', email: 'joao.silva@smurfitwestrock.com.br', nome: 'JOAO SILVA', dataCadastro: iso(580), dataBloqueio: null, ultimoAcesso: iso(70) },
  { id: 5010, sistemaId: 'benTrimbox', email: 'maria.alves@smurfitwestrock.com.br', nome: 'MARIA ALVES', dataCadastro: iso(470), dataBloqueio: null, ultimoAcesso: iso(80) },
  { id: 5011, sistemaId: 'benTrimpaper', email: 'maria.alves@smurfitwestrock.com.br', nome: 'MARIA ALVES', dataCadastro: iso(460), dataBloqueio: iso(20), ultimoAcesso: iso(210) },
  { id: 5012, sistemaId: 'benDadosadv', email: 'maria.alves@smurfitwestrock.com.br', nome: 'MARIA ALVES', dataCadastro: iso(455), dataBloqueio: null, ultimoAcesso: iso(35) },
  { id: 5020, sistemaId: 'ubeTrimbox', email: 'carlos.pereira@smurfitwestrock.com.br', nome: 'CARLOS PEREIRA', dataCadastro: iso(1100), dataBloqueio: iso(30), ultimoAcesso: iso(195) },
  { id: 5021, sistemaId: 'ubeTrimpaper', email: 'carlos.pereira@smurfitwestrock.com.br', nome: 'CARLOS PEREIRA', dataCadastro: iso(1090), dataBloqueio: null, ultimoAcesso: iso(320) },
  { id: 5030, sistemaId: 'saqTrimbox', email: 'fernanda.souza@smurfitwestrock.com.br', nome: 'FERNANDA SOUZA', dataCadastro: iso(690), dataBloqueio: iso(120), ultimoAcesso: iso(250) },
  { id: 5031, sistemaId: 'saqDadosadv', email: 'fernanda.souza@smurfitwestrock.com.br', nome: 'FERNANDA SOUZA', dataCadastro: iso(680), dataBloqueio: null, ultimoAcesso: iso(260) },
  { id: 5040, sistemaId: 'forTrimbox', email: 'ricardo.nunes@smurfitwestrock.com.br', nome: 'RICARDO NUNES', dataCadastro: iso(300), dataBloqueio: null, ultimoAcesso: iso(140) },
  { id: 5041, sistemaId: 'forDadosadv', email: 'ricardo.nunes@smurfitwestrock.com.br', nome: 'RICARDO NUNES', dataCadastro: iso(295), dataBloqueio: null, ultimoAcesso: null },
  { id: 5050, sistemaId: 'pirTrimbox', email: 'patricia.lima@smurfitwestrock.com.br', nome: 'PATRICIA LIMA', dataCadastro: iso(200), dataBloqueio: null, ultimoAcesso: iso(15) },
  { id: 5051, sistemaId: 'pirDadosadv', email: 'patricia.lima@smurfitwestrock.com.br', nome: 'PATRICIA LIMA', dataCadastro: iso(195), dataBloqueio: null, ultimoAcesso: iso(18) },
  { id: 5060, sistemaId: 'benTrimbox', email: 'bruno.carvalho@smurfitwestrock.com.br', nome: 'BRUNO CARVALHO', dataCadastro: iso(140), dataBloqueio: null, ultimoAcesso: iso(101) },
  { id: 5061, sistemaId: 'benDadosadv', email: 'bruno.carvalho@smurfitwestrock.com.br', nome: 'BRUNO CARVALHO', dataCadastro: iso(138), dataBloqueio: null, ultimoAcesso: iso(99) }
]

/**
 * Etapa 1 do legado: pré-consulta em todos os bancos só para descobrir quais
 * e-mails existem em algum sistema.
 */
export async function pesquisarEmails(nome: string, email: string): Promise<string[]> {
  const buscaNome = nome.trim().toLowerCase()
  const buscaEmail = email.trim().toLowerCase()

  if (!buscaNome && !buscaEmail) return atraso([], 200)

  const encontrados = new Set<string>()

  const combina = (nomeConta: string, emailConta: string) => {
    // O ASP quebra o nome em palavras e exige todas elas (CriarCondicaoNome)
    const nomeOk = buscaNome
      ? buscaNome.split(/\s+/).every((palavra) => nomeConta.toLowerCase().includes(palavra))
      : true
    const emailOk = buscaEmail ? emailConta.toLowerCase().includes(buscaEmail) : true
    return nomeOk && emailOk
  }

  for (const usuario of usuariosSwnet) {
    if (combina(usuario.nome, usuario.email)) encontrados.add(usuario.email)
  }
  for (const conta of contasExternas) {
    if (combina(conta.nome, conta.email)) encontrados.add(conta.email)
  }

  return atraso([...encontrados].sort(), 500)
}

/** Etapa 2: para os e-mails escolhidos, monta o grid conta a conta */
export async function listarContasPorEmail(emails: string[]): Promise<ContaSistema[]> {
  const alvos = emails.map((e) => e.toLowerCase())
  const linhas: ContaSistema[] = []

  for (const sistema of SISTEMAS) {
    const contasDoSistema =
      sistema.id === 'swnet'
        ? usuariosSwnet
            .filter((u) => alvos.includes(u.email.toLowerCase()))
            .map<ContaSistema>((u) => ({
              id: u.id,
              sistemaId: 'swnet',
              sistema: sistema.nome,
              nome: u.nome,
              email: u.email,
              dataCadastro: u.dataCadastro,
              dataBloqueio: u.dataBloqueio,
              ativo: u.dataBloqueio === null,
              semRegistro: false
            }))
        : contasExternas
            .filter((c) => c.sistemaId === sistema.id && alvos.includes(c.email.toLowerCase()))
            .map<ContaSistema>((c) => ({
              id: c.id,
              sistemaId: c.sistemaId,
              sistema: sistema.nome,
              nome: c.nome,
              email: c.email,
              dataCadastro: c.dataCadastro,
              dataBloqueio: c.dataBloqueio,
              ativo: c.dataBloqueio === null,
              semRegistro: false
            }))

    if (contasDoSistema.length > 0) {
      linhas.push(...contasDoSistema)
    } else {
      linhas.push({
        id: 0,
        sistemaId: sistema.id,
        sistema: sistema.nome,
        nome: '',
        email: '',
        dataCadastro: null,
        dataBloqueio: null,
        ativo: false,
        semRegistro: true
      })
    }
  }

  return atraso(linhas, 600)
}

/**
 * Etapa 3: habilita/desabilita as contas marcadas. Como no repositório ASP,
 * só entram as que realmente mudam de estado (desabilitar exige
 * DataBloqueio IS NULL e habilitar exige DataBloqueio IS NOT NULL).
 */
export async function alterarContas(
  acao: AcaoCancelamento,
  contas: { id: number; sistemaId: string }[]
): Promise<{ sucesso: boolean; alteradas: ContaSistema[] }> {
  const desabilitando = acao === 'desabilitar'
  const alteradas: ContaSistema[] = []
  const hoje = new Date().toISOString().slice(0, 10)

  for (const alvo of contas) {
    if (alvo.sistemaId === 'swnet') {
      const usuario = usuariosSwnet.find((u) => u.id === alvo.id)
      if (!usuario) continue
      if (desabilitando === (usuario.dataBloqueio !== null)) continue

      usuario.dataBloqueio = desabilitando ? hoje : null
      alteradas.push({
        id: usuario.id,
        sistemaId: 'swnet',
        sistema: nomeSistema('swnet'),
        nome: usuario.nome,
        email: usuario.email,
        dataCadastro: usuario.dataCadastro,
        dataBloqueio: usuario.dataBloqueio,
        ativo: usuario.dataBloqueio === null,
        semRegistro: false
      })
      continue
    }

    const conta = contasExternas.find((c) => c.id === alvo.id && c.sistemaId === alvo.sistemaId)
    if (!conta) continue
    if (desabilitando === (conta.dataBloqueio !== null)) continue

    conta.dataBloqueio = desabilitando ? hoje : null
    alteradas.push({
      id: conta.id,
      sistemaId: conta.sistemaId,
      sistema: nomeSistema(conta.sistemaId),
      nome: conta.nome,
      email: conta.email,
      dataCadastro: conta.dataCadastro,
      dataBloqueio: conta.dataBloqueio,
      ativo: conta.dataBloqueio === null,
      semRegistro: false
    })
  }

  return atraso({ sucesso: true, alteradas }, 700)
}

/**
 * RevisaoUsuarios.asp: contas ativas ordenadas por tempo sem acesso.
 * `diasMinimos = 0` desliga o filtro ("Sem filtro" no combo do legado).
 */
export async function listarUsuariosInativos(
  diasMinimos: number,
  sistemaId: string
): Promise<UsuarioInativo[]> {
  const doSwnet: UsuarioInativo[] = usuariosSwnet
    .filter((u) => u.ativo && u.dataBloqueio === null)
    .map((u) => ({
      id: u.id,
      sistemaId: 'swnet',
      sistema: nomeSistema('swnet'),
      nome: u.nome,
      email: u.email,
      ultimoAcesso: u.ultimoAcesso,
      diasSemAcessar: diasDesde(u.ultimoAcesso, u.dataCadastro)
    }))

  const dosDemais: UsuarioInativo[] = contasExternas
    .filter((c) => c.dataBloqueio === null)
    .map((c) => ({
      id: c.id,
      sistemaId: c.sistemaId,
      sistema: nomeSistema(c.sistemaId),
      nome: c.nome,
      email: c.email,
      ultimoAcesso: c.ultimoAcesso,
      diasSemAcessar: diasDesde(c.ultimoAcesso, c.dataCadastro)
    }))

  const resultado = [...doSwnet, ...dosDemais]
    .filter((u) => sistemaId === 'TODOS' || u.sistemaId === sistemaId)
    .filter((u) => u.diasSemAcessar >= diasMinimos)
    .sort((a, b) => b.diasSemAcessar - a.diasSemAcessar)

  return atraso(resultado, 500)
}

/** Desabilita as contas selecionadas na revisão (mesma gravação do cancelamento) */
export async function desabilitarInativos(contas: { id: number; sistemaId: string }[]) {
  return alterarContas('desabilitar', contas)
}
