import type {
  AprovadorFnc,
  DadosDefeitoFnc,
  DadosResponsavelFnc,
  DadosSetorFnc,
  DefeitoFnc,
  ResponsavelFnc,
  ResultadoCrud,
  SetorFnc,
  TipoDefeito,
  TipoResponsavel
} from '@/pages/cadastrosFnc/tipos'
import { atraso, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * Cadastros-FNC — tabelas Setor, Defeito, Responsavel e AprovadorFNC
 * (banco BEN-CORPORATIVO no legado).
 *
 * Duas convenções do ASP preservadas aqui:
 * - o Id é gerado por MAX(Id)+1, porque as tabelas não têm identity;
 * - excluir é soft delete (FlagInativo = 1 / Excluido = 1), então o registro
 *   some da grade mas continua no array.
 */

interface RegistroInterno {
  inativo: boolean
}

interface SetorInterno extends SetorFnc, RegistroInterno {}
interface DefeitoInterno extends Omit<DefeitoFnc, 'nomeSetor'>, RegistroInterno {}
interface AprovadorInterno extends AprovadorFnc, RegistroInterno {}
interface ResponsavelInterno extends Omit<ResponsavelFnc, 'nomeSetor'>, RegistroInterno {}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const setores: SetorInterno[] = [
  { id: 1, descricao: 'Produção', inativo: false },
  { id: 2, descricao: 'Qualidade', inativo: false },
  { id: 3, descricao: 'Expedição', inativo: false },
  { id: 4, descricao: 'Comercial', inativo: false },
  { id: 5, descricao: 'Desenvolvimento', inativo: false },
  { id: 6, descricao: 'Manutenção', inativo: false },
  { id: 7, descricao: 'Setor Descontinuado', inativo: true }
]

const defeitos: DefeitoInterno[] = [
  { id: 1, tipo: 'PRO', idSetor: 1, descricao: 'Colagem deficiente', inativo: false },
  { id: 2, tipo: 'PRO', idSetor: 1, descricao: 'Impressão fora de registro', inativo: false },
  { id: 3, tipo: 'PRO', idSetor: 1, descricao: 'Corte irregular', inativo: false },
  { id: 4, tipo: 'PRO', idSetor: 2, descricao: 'Gramatura fora de especificação', inativo: false },
  { id: 5, tipo: 'PRO', idSetor: 2, descricao: 'Umidade acima do limite', inativo: false },
  { id: 6, tipo: 'PRO', idSetor: 3, descricao: 'Avaria no transporte', inativo: false },
  { id: 7, tipo: 'PRO', idSetor: 3, descricao: 'Paletização incorreta', inativo: false },
  { id: 8, tipo: 'ADM', idSetor: 4, descricao: 'Divergência de quantidade no pedido', inativo: false },
  { id: 9, tipo: 'ADM', idSetor: 4, descricao: 'Preço divergente da tabela', inativo: false },
  { id: 10, tipo: 'ADM', idSetor: 5, descricao: 'Especificação técnica incorreta', inativo: false }
]

const aprovadores: AprovadorInterno[] = [
  { id: 1003, nome: 'MARIA ALVES', email: 'maria.alves@smurfitwestrock.com.br', dataInclusao: iso(180), inativo: false },
  { id: 1007, nome: 'PATRÍCIA LIMA', email: 'patricia.lima@smurfitwestrock.com.br', dataInclusao: iso(90), inativo: false },
  { id: 1008, nome: 'BRUNO CARVALHO', email: 'bruno.carvalho@smurfitwestrock.com.br', dataInclusao: iso(45), inativo: false }
]

const responsaveis: ResponsavelInterno[] = [
  {
    id: 1002,
    nome: 'JOÃO SILVA',
    email: 'joao.silva@smurfitwestrock.com.br',
    tipo: '1',
    idSetor: 1,
    plantas: ['00', '04'],
    dataInclusao: iso(200),
    inativo: false
  },
  {
    id: 1003,
    nome: 'MARIA ALVES',
    email: 'maria.alves@smurfitwestrock.com.br',
    tipo: '2',
    idSetor: 2,
    plantas: ['05'],
    dataInclusao: iso(160),
    inativo: false
  },
  {
    id: 1008,
    nome: 'BRUNO CARVALHO',
    email: 'bruno.carvalho@smurfitwestrock.com.br',
    tipo: '2',
    idSetor: 3,
    plantas: ['12', '00'],
    dataInclusao: iso(60),
    inativo: false
  }
]

/** MAX(Id)+1, como o ObterProximoId() dos repositórios */
function proximoId(registros: { id: number }[]) {
  return Math.max(0, ...registros.map((r) => r.id)) + 1
}

function nomeSetor(idSetor: number) {
  return setores.find((s) => s.id === idSetor)?.descricao ?? ''
}

/* ------------------------------------------------------------------ *
 * Setores
 * ------------------------------------------------------------------ */

export async function listarSetores(): Promise<SetorFnc[]> {
  const ativos = setores
    .filter((s) => !s.inativo)
    .map(({ id, descricao }) => ({ id, descricao }))
    .sort((a, b) => a.descricao.localeCompare(b.descricao, 'pt-BR'))
  return atraso(ativos, 250)
}

export async function salvarSetor(dados: DadosSetorFnc): Promise<ResultadoCrud> {
  if (!dados.descricao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Descrição é obrigatória.' }, 150)
  }

  if (dados.id) {
    const setor = setores.find((s) => s.id === dados.id)
    if (!setor) return atraso({ sucesso: false, mensagem: 'Setor não encontrado.' }, 150)
    setor.descricao = dados.descricao.trim()
    return atraso({ sucesso: true, mensagem: 'Setor atualizado com sucesso!' }, 400)
  }

  setores.push({ id: proximoId(setores), descricao: dados.descricao.trim(), inativo: false })
  return atraso({ sucesso: true, mensagem: 'Setor criado com sucesso!' }, 400)
}

export async function excluirSetor(id: number): Promise<ResultadoCrud> {
  const setor = setores.find((s) => s.id === id)
  if (!setor) return atraso({ sucesso: false, mensagem: 'Setor não encontrado.' }, 150)

  setor.inativo = true
  return atraso({ sucesso: true, mensagem: 'Setor excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Defeitos
 * ------------------------------------------------------------------ */

export async function listarDefeitos(): Promise<DefeitoFnc[]> {
  const ativos = defeitos
    .filter((d) => !d.inativo)
    .map((d) => ({ ...d, nomeSetor: nomeSetor(d.idSetor) }))
  return atraso(ativos, 300)
}

export async function salvarDefeito(dados: DadosDefeitoFnc): Promise<ResultadoCrud> {
  if (!dados.tipo) return atraso({ sucesso: false, mensagem: 'Tipo é obrigatório.' }, 150)
  if (!dados.idSetor) return atraso({ sucesso: false, mensagem: 'Setor é obrigatório.' }, 150)
  if (!dados.descricao.trim()) {
    return atraso({ sucesso: false, mensagem: 'Descrição é obrigatória.' }, 150)
  }

  if (dados.id) {
    const defeito = defeitos.find((d) => d.id === dados.id)
    if (!defeito) return atraso({ sucesso: false, mensagem: 'Defeito não encontrado.' }, 150)

    defeito.tipo = dados.tipo as TipoDefeito
    defeito.idSetor = Number(dados.idSetor)
    defeito.descricao = dados.descricao.trim()
    return atraso({ sucesso: true, mensagem: 'Defeito atualizado com sucesso!' }, 400)
  }

  defeitos.push({
    id: proximoId(defeitos),
    tipo: dados.tipo as TipoDefeito,
    idSetor: Number(dados.idSetor),
    descricao: dados.descricao.trim(),
    inativo: false
  })
  return atraso({ sucesso: true, mensagem: 'Defeito criado com sucesso!' }, 400)
}

export async function excluirDefeito(id: number): Promise<ResultadoCrud> {
  const defeito = defeitos.find((d) => d.id === id)
  if (!defeito) return atraso({ sucesso: false, mensagem: 'Defeito não encontrado.' }, 150)

  defeito.inativo = true
  return atraso({ sucesso: true, mensagem: 'Defeito excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Aprovadores
 * ------------------------------------------------------------------ */

export async function listarAprovadores(): Promise<AprovadorFnc[]> {
  const ativos = aprovadores
    .filter((a) => !a.inativo)
    .map(({ inativo, ...resto }) => resto)
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'))
  return atraso(ativos, 300)
}

/** O legado copia nome e e-mail da CRM_DTI no momento da inclusão */
export async function salvarAprovador(idUsuario: number | ''): Promise<ResultadoCrud> {
  if (!idUsuario) return atraso({ sucesso: false, mensagem: 'Usuário é obrigatório.' }, 150)

  const usuario = usuariosSwnet.find((u) => u.id === Number(idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 150)

  const existente = aprovadores.find((a) => a.id === usuario.id)
  if (existente && !existente.inativo) {
    return atraso({ sucesso: false, mensagem: 'Este usuário já é aprovador de FNC.' }, 150)
  }

  if (existente) {
    existente.inativo = false
    existente.dataInclusao = new Date().toISOString().slice(0, 10)
  } else {
    aprovadores.push({
      id: usuario.id,
      nome: usuario.nome,
      email: usuario.email,
      dataInclusao: new Date().toISOString().slice(0, 10),
      inativo: false
    })
  }

  return atraso({ sucesso: true, mensagem: 'Aprovador criado com sucesso!' }, 400)
}

export async function excluirAprovador(id: number): Promise<ResultadoCrud> {
  const aprovador = aprovadores.find((a) => a.id === id)
  if (!aprovador) return atraso({ sucesso: false, mensagem: 'Aprovador não encontrado.' }, 150)

  aprovador.inativo = true
  return atraso({ sucesso: true, mensagem: 'Aprovador excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Responsáveis
 * ------------------------------------------------------------------ */

export async function listarResponsaveis(planta = 'TODOS', idSetor = 'TODOS'): Promise<ResponsavelFnc[]> {
  const ativos = responsaveis
    .filter((r) => !r.inativo)
    .filter((r) => planta === 'TODOS' || r.plantas.includes(planta))
    .filter((r) => idSetor === 'TODOS' || String(r.idSetor) === idSetor)
    .map(({ inativo, ...resto }) => ({ ...resto, nomeSetor: nomeSetor(resto.idSetor) }))
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'))

  return atraso(ativos, 350)
}

export async function salvarResponsavel(dados: DadosResponsavelFnc): Promise<ResultadoCrud> {
  if (!dados.idUsuario) return atraso({ sucesso: false, mensagem: 'Responsável é obrigatório.' }, 150)
  if (!dados.idSetor) return atraso({ sucesso: false, mensagem: 'Setor é obrigatório.' }, 150)
  if (!dados.tipo) return atraso({ sucesso: false, mensagem: 'Tipo é obrigatório.' }, 150)
  if (dados.plantas.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos uma planta.' }, 150)
  }

  const usuario = usuariosSwnet.find((u) => u.id === Number(dados.idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 150)

  const existente = responsaveis.find((r) => r.id === usuario.id && !r.inativo)

  // Edição mantém o registro; criação para um usuário que já é responsável
  // sobrescreve, já que a chave no legado é Id + Tipo + Email
  if (existente && (!dados.id || dados.id === existente.id)) {
    existente.tipo = dados.tipo as TipoResponsavel
    existente.idSetor = Number(dados.idSetor)
    existente.plantas = [...dados.plantas]
    return atraso({ sucesso: true, mensagem: 'Responsável atualizado com sucesso!' }, 400)
  }

  responsaveis.push({
    id: usuario.id,
    nome: usuario.nome,
    email: usuario.email,
    tipo: dados.tipo as TipoResponsavel,
    idSetor: Number(dados.idSetor),
    plantas: [...dados.plantas],
    dataInclusao: new Date().toISOString().slice(0, 10),
    inativo: false
  })

  return atraso({ sucesso: true, mensagem: 'Responsável criado com sucesso!' }, 400)
}

export async function excluirResponsavel(id: number): Promise<ResultadoCrud> {
  const responsavel = responsaveis.find((r) => r.id === id)
  if (!responsavel) return atraso({ sucesso: false, mensagem: 'Responsável não encontrado.' }, 150)

  responsavel.inativo = true
  return atraso({ sucesso: true, mensagem: 'Responsável excluído com sucesso!' }, 350)
}
