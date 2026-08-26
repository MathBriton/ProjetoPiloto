import type {
  AnalistaOrcamento,
  AprovadorOrcamento,
  DadosAnalistaOrcamento,
  DadosAprovadorOrcamento,
  DadosFaixaAprovador,
  FaixaAprovador,
  NivelAprovador,
  TipoProduto
} from '@/pages/cadastrosOrcamento/tipos'
import type { ResultadoCrud } from '@/pages/cadastrosFnc/tipos'
import { atraso, nomeEmpresa, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * Cadastros-Orçamento — ANALISTAS_ORCAMENTO, APROVADORES_ORCAMENTO e
 * FAIXA_APROVADORES.
 *
 * Mesmas convenções dos repositórios ASP: Id por MAX(Id)+1 e exclusão lógica.
 */

interface ComInativo {
  inativo: boolean
}

interface AnalistaInterno extends AnalistaOrcamento, ComInativo {}
interface AprovadorInterno extends AprovadorOrcamento, ComInativo {}
interface FaixaInterna extends Omit<FaixaAprovador, 'nomePlanta'>, ComInativo {}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const analistas: AnalistaInterno[] = [
  {
    id: 1002,
    nome: 'JOÃO SILVA',
    email: 'joao.silva@smurfitwestrock.com.br',
    plantas: ['00', '04'],
    dataInclusao: iso(250),
    inativo: false
  },
  {
    id: 1007,
    nome: 'PATRÍCIA LIMA',
    email: 'patricia.lima@smurfitwestrock.com.br',
    plantas: ['00'],
    dataInclusao: iso(120),
    inativo: false
  }
]

const aprovadores: AprovadorInterno[] = [
  {
    id: 1003,
    nome: 'MARIA ALVES',
    email: 'maria.alves@smurfitwestrock.com.br',
    nivel: '2',
    plantas: ['05'],
    dataInclusao: iso(300),
    inativo: false
  },
  {
    id: 1008,
    nome: 'BRUNO CARVALHO',
    email: 'bruno.carvalho@smurfitwestrock.com.br',
    nivel: '1',
    plantas: ['12', '00'],
    dataInclusao: iso(140),
    inativo: false
  },
  {
    id: 1001,
    nome: 'ADMINISTRADOR (FICTÍCIO)',
    email: 'admin@smurfitwestrock.com.br',
    nivel: '3',
    plantas: ['00', '04', '05', '06', '07'],
    dataInclusao: iso(400),
    inativo: false
  }
]

const faixas: FaixaInterna[] = [
  { id: 1, planta: '00', tipoProduto: 'CX', valorMinKg: 0, valorMaxKg: 4.5, valorMinM2: 0, valorMaxM2: 2.2, quantidadeAprovadores: 1, inativo: false },
  { id: 2, planta: '00', tipoProduto: 'CX', valorMinKg: 4.5, valorMaxKg: 9, valorMinM2: 2.2, valorMaxM2: 4.8, quantidadeAprovadores: 2, inativo: false },
  { id: 3, planta: '00', tipoProduto: 'CH', valorMinKg: 0, valorMaxKg: 3.8, valorMinM2: 0, valorMaxM2: 1.9, quantidadeAprovadores: 1, inativo: false },
  { id: 4, planta: '04', tipoProduto: 'CX', valorMinKg: 0, valorMaxKg: 5.2, valorMinM2: 0, valorMaxM2: 2.5, quantidadeAprovadores: 1, inativo: false },
  { id: 5, planta: '05', tipoProduto: 'PA', valorMinKg: 0, valorMaxKg: 2.9, valorMinM2: 0, valorMaxM2: 1.4, quantidadeAprovadores: 2, inativo: false },
  { id: 6, planta: '12', tipoProduto: 'CJ', valorMinKg: 0, valorMaxKg: 7.5, valorMinM2: 0, valorMaxM2: 3.6, quantidadeAprovadores: 3, inativo: false }
]

function proximoId(registros: { id: number }[]) {
  return Math.max(0, ...registros.map((r) => r.id)) + 1
}

/* ------------------------------------------------------------------ *
 * Analistas
 * ------------------------------------------------------------------ */

export async function listarAnalistas(planta = 'TODOS'): Promise<AnalistaOrcamento[]> {
  const ativos = analistas
    .filter((a) => !a.inativo)
    .filter((a) => planta === 'TODOS' || a.plantas.includes(planta))
    .map(({ inativo, ...resto }) => resto)
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'))

  return atraso(ativos, 300)
}

export async function salvarAnalista(dados: DadosAnalistaOrcamento): Promise<ResultadoCrud> {
  if (!dados.idUsuario) return atraso({ sucesso: false, mensagem: 'Analista é obrigatório.' }, 150)
  if (dados.plantas.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos uma planta.' }, 150)
  }

  const usuario = usuariosSwnet.find((u) => u.id === Number(dados.idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 150)

  const existente = analistas.find((a) => a.id === usuario.id && !a.inativo)
  if (existente && (!dados.id || dados.id === existente.id)) {
    existente.plantas = [...dados.plantas]
    return atraso({ sucesso: true, mensagem: 'Analista atualizado com sucesso!' }, 400)
  }

  analistas.push({
    id: usuario.id,
    nome: usuario.nome,
    email: usuario.email,
    plantas: [...dados.plantas],
    dataInclusao: new Date().toISOString().slice(0, 10),
    inativo: false
  })

  return atraso({ sucesso: true, mensagem: 'Analista criado com sucesso!' }, 400)
}

export async function excluirAnalista(id: number): Promise<ResultadoCrud> {
  const analista = analistas.find((a) => a.id === id)
  if (!analista) return atraso({ sucesso: false, mensagem: 'Analista não encontrado.' }, 150)

  analista.inativo = true
  return atraso({ sucesso: true, mensagem: 'Analista excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Aprovadores
 * ------------------------------------------------------------------ */

export async function listarAprovadoresOrcamento(planta = 'TODOS'): Promise<AprovadorOrcamento[]> {
  const ativos = aprovadores
    .filter((a) => !a.inativo)
    .filter((a) => planta === 'TODOS' || a.plantas.includes(planta))
    .map(({ inativo, ...resto }) => resto)
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'))

  return atraso(ativos, 300)
}

export async function salvarAprovadorOrcamento(dados: DadosAprovadorOrcamento): Promise<ResultadoCrud> {
  if (!dados.idUsuario) return atraso({ sucesso: false, mensagem: 'Aprovador é obrigatório.' }, 150)
  if (!dados.nivel) return atraso({ sucesso: false, mensagem: 'Nível é obrigatório.' }, 150)
  if (dados.plantas.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos uma planta.' }, 150)
  }

  const usuario = usuariosSwnet.find((u) => u.id === Number(dados.idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 150)

  const existente = aprovadores.find((a) => a.id === usuario.id && !a.inativo)
  if (existente && (!dados.id || dados.id === existente.id)) {
    existente.nivel = dados.nivel as NivelAprovador
    existente.plantas = [...dados.plantas]
    return atraso({ sucesso: true, mensagem: 'Aprovador atualizado com sucesso!' }, 400)
  }

  aprovadores.push({
    id: usuario.id,
    nome: usuario.nome,
    email: usuario.email,
    nivel: dados.nivel as NivelAprovador,
    plantas: [...dados.plantas],
    dataInclusao: new Date().toISOString().slice(0, 10),
    inativo: false
  })

  return atraso({ sucesso: true, mensagem: 'Aprovador criado com sucesso!' }, 400)
}

export async function excluirAprovadorOrcamento(id: number): Promise<ResultadoCrud> {
  const aprovador = aprovadores.find((a) => a.id === id)
  if (!aprovador) return atraso({ sucesso: false, mensagem: 'Aprovador não encontrado.' }, 150)

  aprovador.inativo = true
  return atraso({ sucesso: true, mensagem: 'Aprovador excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Faixas de aprovador
 * ------------------------------------------------------------------ */

export async function listarFaixas(planta = 'TODOS'): Promise<FaixaAprovador[]> {
  const ativas = faixas
    .filter((f) => !f.inativo)
    .filter((f) => planta === 'TODOS' || f.planta === planta)
    .map(({ inativo, ...resto }) => ({ ...resto, nomePlanta: nomeEmpresa(resto.planta) }))
    .sort((a, b) => a.planta.localeCompare(b.planta) || a.tipoProduto.localeCompare(b.tipoProduto))

  return atraso(ativas, 300)
}

export async function salvarFaixa(dados: DadosFaixaAprovador): Promise<ResultadoCrud> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 150)
  if (!dados.tipoProduto) return atraso({ sucesso: false, mensagem: 'Tipo de produto é obrigatório.' }, 150)

  const numeros = {
    valorMinKg: Number(dados.valorMinKg),
    valorMaxKg: Number(dados.valorMaxKg),
    valorMinM2: Number(dados.valorMinM2),
    valorMaxM2: Number(dados.valorMaxM2),
    quantidadeAprovadores: Number(dados.quantidadeAprovadores)
  }

  if (Object.values(numeros).some((valor) => Number.isNaN(valor))) {
    return atraso({ sucesso: false, mensagem: 'Preencha todos os valores numéricos.' }, 150)
  }
  if (numeros.valorMinKg > numeros.valorMaxKg) {
    return atraso({ sucesso: false, mensagem: 'O valor mínimo por Kg não pode ser maior que o máximo.' }, 150)
  }
  if (numeros.valorMinM2 > numeros.valorMaxM2) {
    return atraso({ sucesso: false, mensagem: 'O valor mínimo por M² não pode ser maior que o máximo.' }, 150)
  }
  if (numeros.quantidadeAprovadores < 1) {
    return atraso({ sucesso: false, mensagem: 'A quantidade de aprovadores deve ser ao menos 1.' }, 150)
  }

  if (dados.id) {
    const faixa = faixas.find((f) => f.id === dados.id)
    if (!faixa) return atraso({ sucesso: false, mensagem: 'Faixa não encontrada.' }, 150)

    Object.assign(faixa, {
      planta: dados.planta,
      tipoProduto: dados.tipoProduto as TipoProduto,
      ...numeros
    })
    return atraso({ sucesso: true, mensagem: 'Faixa atualizada com sucesso!' }, 400)
  }

  faixas.push({
    id: proximoId(faixas),
    planta: dados.planta,
    tipoProduto: dados.tipoProduto as TipoProduto,
    ...numeros,
    inativo: false
  })

  return atraso({ sucesso: true, mensagem: 'Faixa criada com sucesso!' }, 400)
}

export async function excluirFaixa(id: number): Promise<ResultadoCrud> {
  const faixa = faixas.find((f) => f.id === id)
  if (!faixa) return atraso({ sucesso: false, mensagem: 'Faixa não encontrada.' }, 150)

  faixa.inativo = true
  return atraso({ sucesso: true, mensagem: 'Faixa excluída com sucesso!' }, 350)
}
