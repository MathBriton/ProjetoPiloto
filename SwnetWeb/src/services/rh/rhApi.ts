import { atraso, usuariosSwnet } from '@/services/comuns/baseMock'

/**
 * RH — cadastros de apoio (Atendentes, Categorias, Departamentos,
 * Prioridades), o sistema de chamados e a importação do ADP.
 *
 * As três pastas de chamado do legado (Chamados-RH, Meus-Chamados-RH e
 * Novo-Chamado-RH) têm os mesmos 8 arquivos: é um só sistema com três
 * pontos de entrada.
 */

export interface ResultadoRh {
  sucesso: boolean
  mensagem: string
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

function agora() {
  return new Date().toISOString().slice(0, 16).replace('T', ' ')
}

/* ------------------------------------------------------------------ *
 * Cadastros de apoio
 * ------------------------------------------------------------------ */

export interface DepartamentoChamado {
  id: number
  planta: string
  nome: string
  dataCriacao: string
}

export interface CategoriaChamado {
  id: number
  planta: string
  nome: string
  tipo: string
  dataCriacao: string
}

/** PrioridadesChamado.NIVEL — o legado grava 'Media' com esse casing */
export const NIVEIS_PRIORIDADE = [
  { codigo: 'BAIXA', descricao: 'Baixa' },
  { codigo: 'Media', descricao: 'Média' },
  { codigo: 'ALTA', descricao: 'Alta' },
  { codigo: 'URGENTE', descricao: 'Urgente' }
]

export function descricaoNivel(codigo: string) {
  return NIVEIS_PRIORIDADE.find((n) => n.codigo === codigo)?.descricao ?? codigo
}

export interface PrioridadeChamado {
  id: number
  planta: string
  nivel: string
  prazoHoras: number
  dataCriacao: string
}

export interface AtendenteRh {
  id: number
  planta: string
  nome: string
  usuario: string
  email: string
  departamentos: number[]
  dataCriacao: string
}

interface ComExclusao {
  excluido: boolean
}

const departamentos: (DepartamentoChamado & ComExclusao)[] = [
  { id: 1, planta: '00', nome: 'Benefícios', dataCriacao: iso(200), excluido: false },
  { id: 2, planta: '00', nome: 'Folha de Pagamento', dataCriacao: iso(200), excluido: false },
  { id: 3, planta: '00', nome: 'Recrutamento e Seleção', dataCriacao: iso(180), excluido: false },
  { id: 4, planta: '05', nome: 'Segurança do Trabalho', dataCriacao: iso(150), excluido: false },
  { id: 5, planta: '04', nome: 'Treinamento', dataCriacao: iso(120), excluido: false }
]

const categorias: (CategoriaChamado & ComExclusao)[] = [
  { id: 1, planta: '00', nome: 'Vale-transporte', tipo: 'Benefícios', dataCriacao: iso(190), excluido: false },
  { id: 2, planta: '00', nome: 'Plano de saúde', tipo: 'Benefícios', dataCriacao: iso(190), excluido: false },
  { id: 3, planta: '00', nome: 'Holerite', tipo: 'Folha', dataCriacao: iso(180), excluido: false },
  { id: 4, planta: '00', nome: 'Férias', tipo: 'Folha', dataCriacao: iso(180), excluido: false },
  { id: 5, planta: '05', nome: 'EPI', tipo: 'Segurança', dataCriacao: iso(140), excluido: false },
  { id: 6, planta: '04', nome: 'Certificação', tipo: 'Treinamento', dataCriacao: iso(110), excluido: false }
]

const prioridades: (PrioridadeChamado & ComExclusao)[] = [
  { id: 1, planta: '00', nivel: 'BAIXA', prazoHoras: 72, dataCriacao: iso(200), excluido: false },
  { id: 2, planta: '00', nivel: 'Media', prazoHoras: 48, dataCriacao: iso(200), excluido: false },
  { id: 3, planta: '00', nivel: 'ALTA', prazoHoras: 24, dataCriacao: iso(200), excluido: false },
  { id: 4, planta: '00', nivel: 'URGENTE', prazoHoras: 4, dataCriacao: iso(200), excluido: false }
]

const atendentes: (AtendenteRh & ComExclusao)[] = [
  {
    id: 1007,
    planta: '00',
    nome: 'PATRÍCIA LIMA',
    usuario: 'plima',
    email: 'patricia.lima@smurfitwestrock.com.br',
    departamentos: [1, 2],
    dataCriacao: iso(150),
    excluido: false
  },
  {
    id: 1003,
    planta: '05',
    nome: 'MARIA ALVES',
    usuario: 'malves',
    email: 'maria.alves@smurfitwestrock.com.br',
    departamentos: [4],
    dataCriacao: iso(120),
    excluido: false
  },
  {
    id: 1008,
    planta: '00',
    nome: 'BRUNO CARVALHO',
    usuario: 'bcarvalho',
    email: 'bruno.carvalho@smurfitwestrock.com.br',
    departamentos: [1, 3],
    dataCriacao: iso(90),
    excluido: false
  }
]

function proximoId(lista: { id: number }[]) {
  return Math.max(0, ...lista.map((item) => item.id)) + 1
}

/* Departamentos */

export async function listarDepartamentos(planta = ''): Promise<DepartamentoChamado[]> {
  const resultado = departamentos
    .filter((d) => !d.excluido)
    .filter((d) => !planta || d.planta === planta)
    .map(({ excluido, ...resto }) => resto)
  return atraso(resultado, 300)
}

export async function salvarDepartamento(dados: {
  id?: number
  planta: string
  nome: string
}): Promise<ResultadoRh> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.nome.trim()) return atraso({ sucesso: false, mensagem: 'Nome é obrigatório.' }, 200)

  if (dados.id) {
    const registro = departamentos.find((d) => d.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Departamento não encontrado.' }, 200)
    Object.assign(registro, { planta: dados.planta, nome: dados.nome.trim() })
    return atraso({ sucesso: true, mensagem: 'Departamento atualizado com sucesso!' }, 400)
  }

  departamentos.push({
    id: proximoId(departamentos),
    planta: dados.planta,
    nome: dados.nome.trim(),
    dataCriacao: new Date().toISOString().slice(0, 10),
    excluido: false
  })
  return atraso({ sucesso: true, mensagem: 'Departamento criado com sucesso!' }, 400)
}

export async function excluirDepartamento(id: number): Promise<ResultadoRh> {
  const registro = departamentos.find((d) => d.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Departamento não encontrado.' }, 200)

  // Não deixa órfão: atendentes e chamados referenciam o departamento
  if (atendentes.some((a) => !a.excluido && a.departamentos.includes(id))) {
    return atraso(
      { sucesso: false, mensagem: 'Existem atendentes vinculados a este departamento.' },
      200
    )
  }

  registro.excluido = true
  return atraso({ sucesso: true, mensagem: 'Departamento excluído com sucesso!' }, 350)
}

/* Categorias */

export async function listarCategorias(planta = ''): Promise<CategoriaChamado[]> {
  const resultado = categorias
    .filter((c) => !c.excluido)
    .filter((c) => !planta || c.planta === planta)
    .map(({ excluido, ...resto }) => resto)
  return atraso(resultado, 300)
}

export async function salvarCategoria(dados: {
  id?: number
  planta: string
  nome: string
  tipo: string
}): Promise<ResultadoRh> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.nome.trim()) return atraso({ sucesso: false, mensagem: 'Nome é obrigatório.' }, 200)

  if (dados.id) {
    const registro = categorias.find((c) => c.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Categoria não encontrada.' }, 200)
    Object.assign(registro, { planta: dados.planta, nome: dados.nome.trim(), tipo: dados.tipo.trim() })
    return atraso({ sucesso: true, mensagem: 'Categoria atualizada com sucesso!' }, 400)
  }

  categorias.push({
    id: proximoId(categorias),
    planta: dados.planta,
    nome: dados.nome.trim(),
    tipo: dados.tipo.trim(),
    dataCriacao: new Date().toISOString().slice(0, 10),
    excluido: false
  })
  return atraso({ sucesso: true, mensagem: 'Categoria criada com sucesso!' }, 400)
}

export async function excluirCategoria(id: number): Promise<ResultadoRh> {
  const registro = categorias.find((c) => c.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Categoria não encontrada.' }, 200)

  registro.excluido = true
  return atraso({ sucesso: true, mensagem: 'Categoria excluída com sucesso!' }, 350)
}

/* Prioridades */

export async function listarPrioridades(planta = ''): Promise<PrioridadeChamado[]> {
  const resultado = prioridades
    .filter((p) => !p.excluido)
    .filter((p) => !planta || p.planta === planta)
    .map(({ excluido, ...resto }) => resto)
  return atraso(resultado, 300)
}

export async function salvarPrioridade(dados: {
  id?: number
  planta: string
  nivel: string
  prazoHoras: string
}): Promise<ResultadoRh> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.nivel) return atraso({ sucesso: false, mensagem: 'Nível é obrigatório.' }, 200)

  const prazo = Number(dados.prazoHoras)
  if (!prazo || Number.isNaN(prazo) || prazo <= 0) {
    return atraso({ sucesso: false, mensagem: 'Informe um prazo em horas válido.' }, 200)
  }

  const duplicada = prioridades.some(
    (p) => !p.excluido && p.planta === dados.planta && p.nivel === dados.nivel && p.id !== dados.id
  )
  if (duplicada) {
    return atraso({ sucesso: false, mensagem: 'Já existe este nível cadastrado para a planta.' }, 200)
  }

  if (dados.id) {
    const registro = prioridades.find((p) => p.id === dados.id)
    if (!registro) return atraso({ sucesso: false, mensagem: 'Prioridade não encontrada.' }, 200)
    Object.assign(registro, { planta: dados.planta, nivel: dados.nivel, prazoHoras: prazo })
    return atraso({ sucesso: true, mensagem: 'Prioridade atualizada com sucesso!' }, 400)
  }

  prioridades.push({
    id: proximoId(prioridades),
    planta: dados.planta,
    nivel: dados.nivel,
    prazoHoras: prazo,
    dataCriacao: new Date().toISOString().slice(0, 10),
    excluido: false
  })
  return atraso({ sucesso: true, mensagem: 'Prioridade criada com sucesso!' }, 400)
}

export async function excluirPrioridade(id: number): Promise<ResultadoRh> {
  const registro = prioridades.find((p) => p.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Prioridade não encontrada.' }, 200)

  registro.excluido = true
  return atraso({ sucesso: true, mensagem: 'Prioridade excluída com sucesso!' }, 350)
}

/* Atendentes */

export async function listarAtendentes(planta = ''): Promise<AtendenteRh[]> {
  const resultado = atendentes
    .filter((a) => !a.excluido)
    .filter((a) => !planta || a.planta === planta)
    .map(({ excluido, ...resto }) => resto)
  return atraso(resultado, 300)
}

export async function salvarAtendente(dados: {
  id?: number
  planta: string
  idUsuario: number | ''
  departamentos: number[]
}): Promise<ResultadoRh> {
  if (!dados.planta) return atraso({ sucesso: false, mensagem: 'Planta é obrigatória.' }, 200)
  if (!dados.idUsuario) return atraso({ sucesso: false, mensagem: 'Usuário é obrigatório.' }, 200)
  if (dados.departamentos.length === 0) {
    return atraso({ sucesso: false, mensagem: 'Selecione ao menos um departamento.' }, 200)
  }

  const usuario = usuariosSwnet.find((u) => u.id === Number(dados.idUsuario))
  if (!usuario) return atraso({ sucesso: false, mensagem: 'Usuário não encontrado.' }, 200)

  const existente = atendentes.find((a) => a.id === usuario.id && !a.excluido)
  if (existente && (!dados.id || dados.id === existente.id)) {
    Object.assign(existente, {
      planta: dados.planta,
      departamentos: [...dados.departamentos]
    })
    return atraso({ sucesso: true, mensagem: 'Atendente atualizado com sucesso!' }, 400)
  }

  atendentes.push({
    id: usuario.id,
    planta: dados.planta,
    nome: usuario.nome,
    usuario: usuario.login,
    email: usuario.email,
    departamentos: [...dados.departamentos],
    dataCriacao: new Date().toISOString().slice(0, 10),
    excluido: false
  })
  return atraso({ sucesso: true, mensagem: 'Atendente criado com sucesso!' }, 400)
}

export async function excluirAtendente(id: number): Promise<ResultadoRh> {
  const registro = atendentes.find((a) => a.id === id)
  if (!registro) return atraso({ sucesso: false, mensagem: 'Atendente não encontrado.' }, 200)

  registro.excluido = true
  return atraso({ sucesso: true, mensagem: 'Atendente excluído com sucesso!' }, 350)
}

/* ------------------------------------------------------------------ *
 * Chamados
 * ------------------------------------------------------------------ */

export const STATUS_CHAMADO = [
  { codigo: 'ABERTO', descricao: 'Aberto' },
  { codigo: 'AGUARDANDO_SOLICITANTE', descricao: 'Aguardando solicitante' },
  { codigo: 'REABERTO', descricao: 'Reaberto' },
  { codigo: 'FINALIZADO', descricao: 'Finalizado' },
  { codigo: 'CANCELADO', descricao: 'Cancelado' }
]

export function descricaoStatusChamado(codigo: string) {
  return STATUS_CHAMADO.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

export const TIPOS_CHAMADO = ['Dúvida', 'Solicitação', 'Reclamação', 'Sugestão']

/** Motivos de avaliação — o legado troca a lista conforme a nota */
export const MOTIVOS_POSITIVOS = [
  'Rapidez na resposta',
  'Atendimento cordial',
  'Problema resolvido',
  'Comunicação clara'
]

export const MOTIVOS_NEGATIVOS = [
  'Demora na resposta',
  'Problema não resolvido',
  'Falta de clareza',
  'Precisei reabrir o chamado'
]

export function motivosPorNota(nota: number) {
  return nota >= 4 ? MOTIVOS_POSITIVOS : MOTIVOS_NEGATIVOS
}

export interface MensagemChamado {
  autor: string
  ehAtendente: boolean
  dataHora: string
  texto: string
}

export interface AvaliacaoChamado {
  nota: number
  motivos: string[]
  comentario: string
  data: string
}

export interface Chamado {
  id: number
  planta: string
  assunto: string
  tipo: string
  descricao: string
  solicitante: string
  nomeSolicitante: string
  emailSolicitante: string
  telefone: string
  centroCusto: string
  idDepartamento: number
  idCategoria: number | null
  idPrioridade: number
  atendente: string
  status: string
  dataCriacao: string
  dataPrevista: string
  dataFinalizacao: string
  anexos: string[]
  mensagens: MensagemChamado[]
  avaliacao: AvaliacaoChamado | null
}

const ASSUNTOS = [
  'Dúvida sobre desconto no holerite',
  'Solicitação de segunda via do vale-transporte',
  'Agendamento de exame periódico',
  'Correção de banco de horas',
  'Solicitação de EPI',
  'Dúvida sobre férias proporcionais',
  'Reembolso de plano de saúde',
  'Certificado de treinamento NR-35'
]

const chamados: Chamado[] = Array.from({ length: 24 }, (_, indice) => {
  const status = STATUS_CHAMADO[indice % STATUS_CHAMADO.length].codigo
  const prioridade = prioridades[indice % prioridades.length]
  const finalizado = status === 'FINALIZADO'

  return {
    id: 5500 + indice,
    planta: ['00', '00', '04', '05'][indice % 4],
    assunto: ASSUNTOS[indice % ASSUNTOS.length],
    tipo: TIPOS_CHAMADO[indice % TIPOS_CHAMADO.length],
    descricao: 'Detalhamento da solicitação registrada pelo colaborador.',
    solicitante: ['jsilva', 'cpereira', 'rnunes', 'admin'][indice % 4],
    nomeSolicitante: ['JOÃO SILVA', 'CARLOS PEREIRA', 'RICARDO NUNES', 'ADMINISTRADOR'][indice % 4],
    emailSolicitante: 'colaborador@smurfitwestrock.com.br',
    telefone: '(32) 3222-1000',
    centroCusto: `CC${String(1000 + indice).slice(-4)}`,
    idDepartamento: departamentos[indice % departamentos.length].id,
    idCategoria: indice % 5 === 0 ? null : categorias[indice % categorias.length].id,
    idPrioridade: prioridade.id,
    atendente: indice % 3 === 0 ? '' : atendentes[indice % atendentes.length].usuario,
    status,
    dataCriacao: iso(indice + 1),
    dataPrevista: iso(indice + 1 - Math.ceil(prioridade.prazoHoras / 24)),
    dataFinalizacao: finalizado ? iso(indice) : '',
    anexos: indice % 4 === 0 ? ['comprovante.pdf'] : [],
    mensagens: [
      {
        autor: ['jsilva', 'cpereira', 'rnunes', 'admin'][indice % 4],
        ehAtendente: false,
        dataHora: `${iso(indice + 1)} 09:${String(10 + (indice % 40)).padStart(2, '0')}`,
        texto: 'Boa tarde, poderiam verificar por favor?'
      },
      ...(indice % 3 !== 0
        ? [
            {
              autor: atendentes[indice % atendentes.length].usuario,
              ehAtendente: true,
              dataHora: `${iso(indice)} 14:${String(10 + (indice % 40)).padStart(2, '0')}`,
              texto: 'Olá! Já estamos verificando com a área responsável.'
            }
          ]
        : [])
    ],
    avaliacao:
      finalizado && indice % 2 === 0
        ? {
            nota: 5,
            motivos: ['Rapidez na resposta', 'Problema resolvido'],
            comentario: 'Atendimento rápido, obrigado!',
            data: iso(indice)
          }
        : null
  }
})

export interface FiltrosChamado {
  planta: string
  status: string
  idDepartamento: string
  idPrioridade: string
  termo: string
}

export const FILTROS_CHAMADO_VAZIOS: FiltrosChamado = {
  planta: '',
  status: '',
  idDepartamento: '',
  idPrioridade: '',
  termo: ''
}

/** `escopo` reproduz os três pontos de entrada do legado */
export type EscopoChamados = 'FILA' | 'MEUS'

export async function listarChamados(
  escopo: EscopoChamados,
  filtros: FiltrosChamado,
  usuario: string
): Promise<Chamado[]> {
  const termo = filtros.termo.trim().toLowerCase()

  const resultado = chamados
    .filter((c) => (escopo === 'MEUS' ? c.solicitante === usuario : true))
    .filter((c) => !filtros.planta || c.planta === filtros.planta)
    .filter((c) => !filtros.status || c.status === filtros.status)
    .filter((c) => !filtros.idDepartamento || String(c.idDepartamento) === filtros.idDepartamento)
    .filter((c) => !filtros.idPrioridade || String(c.idPrioridade) === filtros.idPrioridade)
    .filter(
      (c) =>
        !termo ||
        c.assunto.toLowerCase().includes(termo) ||
        String(c.id).includes(termo) ||
        c.nomeSolicitante.toLowerCase().includes(termo)
    )
    .map((c) => ({ ...c }))
    .sort((a, b) => b.dataCriacao.localeCompare(a.dataCriacao))

  return atraso(resultado, 450)
}

export interface KpiChamados {
  abertos: number
  aguardando: number
  finalizados: number
  atrasados: number
  notaMedia: number
}

/** ObterDadosKPI — o painel de indicadores da fila do atendente */
export async function obterKpis(planta: string): Promise<KpiChamados> {
  const base = chamados.filter((c) => !planta || c.planta === planta)
  const hoje = new Date().toISOString().slice(0, 10)

  const avaliados = base.filter((c) => c.avaliacao)
  const notaMedia = avaliados.length
    ? avaliados.reduce((soma, c) => soma + (c.avaliacao?.nota ?? 0), 0) / avaliados.length
    : 0

  return atraso(
    {
      abertos: base.filter((c) => c.status === 'ABERTO' || c.status === 'REABERTO').length,
      aguardando: base.filter((c) => c.status === 'AGUARDANDO_SOLICITANTE').length,
      finalizados: base.filter((c) => c.status === 'FINALIZADO').length,
      // Atrasado = prazo vencido e ainda sem finalização
      atrasados: base.filter(
        (c) => c.status !== 'FINALIZADO' && c.status !== 'CANCELADO' && c.dataPrevista < hoje
      ).length,
      notaMedia: Number(notaMedia.toFixed(1))
    },
    350
  )
}

export interface DadosNovoChamado {
  planta: string
  tipo: string
  idCategoria: string
  nomeCompleto: string
  email: string
  telefone: string
  idDepartamento: string
  assunto: string
  centroCusto: string
  descricao: string
  anexos: string[]
}

export const NOVO_CHAMADO_VAZIO: DadosNovoChamado = {
  planta: '',
  tipo: '',
  idCategoria: '',
  nomeCompleto: '',
  email: '',
  telefone: '',
  idDepartamento: '',
  assunto: '',
  centroCusto: '',
  descricao: '',
  anexos: []
}

export async function criarChamado(
  dados: DadosNovoChamado,
  usuario: string
): Promise<ResultadoRh & { id?: number }> {
  const obrigatorios: [keyof DadosNovoChamado, string][] = [
    ['planta', 'Planta'],
    ['tipo', 'Tipo de Chamado'],
    ['nomeCompleto', 'Nome Completo'],
    ['email', 'E-mail'],
    ['idDepartamento', 'Departamento'],
    ['assunto', 'Assunto'],
    ['descricao', 'Descrição']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }

  // O prazo sai da prioridade padrão (Média) da planta
  const prioridade =
    prioridades.find((p) => !p.excluido && p.planta === dados.planta && p.nivel === 'Media') ??
    prioridades.find((p) => !p.excluido) ??
    prioridades[0]

  const id = proximoId(chamados)
  const hoje = new Date().toISOString().slice(0, 10)
  const previsto = new Date(Date.now() + prioridade.prazoHoras * 3_600_000)
    .toISOString()
    .slice(0, 10)

  chamados.push({
    id,
    planta: dados.planta,
    assunto: dados.assunto.trim(),
    tipo: dados.tipo,
    descricao: dados.descricao.trim(),
    solicitante: usuario,
    nomeSolicitante: dados.nomeCompleto.trim().toUpperCase(),
    emailSolicitante: dados.email.trim(),
    telefone: dados.telefone.trim(),
    centroCusto: dados.centroCusto.trim(),
    idDepartamento: Number(dados.idDepartamento),
    idCategoria: dados.idCategoria ? Number(dados.idCategoria) : null,
    idPrioridade: prioridade.id,
    atendente: '',
    status: 'ABERTO',
    dataCriacao: hoje,
    dataPrevista: previsto,
    dataFinalizacao: '',
    anexos: [...dados.anexos],
    mensagens: [],
    avaliacao: null
  })

  return atraso({ sucesso: true, mensagem: `Chamado ${id} aberto com sucesso!`, id }, 650)
}

export async function enviarMensagem(
  id: number,
  texto: string,
  autor: string,
  ehAtendente: boolean
): Promise<ResultadoRh> {
  if (!texto.trim()) return atraso({ sucesso: false, mensagem: 'Escreva a mensagem.' }, 200)

  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)
  if (chamado.status === 'FINALIZADO' || chamado.status === 'CANCELADO') {
    return atraso({ sucesso: false, mensagem: 'Chamado encerrado não aceita mensagens.' }, 200)
  }

  chamado.mensagens.push({ autor, ehAtendente, dataHora: agora(), texto: texto.trim() })

  // Resposta do atendente coloca a bola com o solicitante, e vice-versa
  chamado.status = ehAtendente ? 'AGUARDANDO_SOLICITANTE' : 'ABERTO'

  return atraso({ sucesso: true, mensagem: 'Mensagem enviada!' }, 450)
}

export async function atribuirAtendente(id: number, atendente: string): Promise<ResultadoRh> {
  if (!atendente) return atraso({ sucesso: false, mensagem: 'Selecione o atendente.' }, 200)

  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)

  chamado.atendente = atendente
  return atraso({ sucesso: true, mensagem: `Chamado atribuído a ${atendente}.` }, 400)
}

/** alterarStatus / alterarPrioridade / alterarCategoria / alterarDepartamento */
export async function alterarCampoChamado(
  id: number,
  campo: 'status' | 'idPrioridade' | 'idCategoria' | 'idDepartamento',
  valor: string
): Promise<ResultadoRh> {
  if (!valor) return atraso({ sucesso: false, mensagem: 'Informe o novo valor.' }, 200)

  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)

  if (campo === 'status') {
    chamado.status = valor
    if (valor === 'FINALIZADO') chamado.dataFinalizacao = new Date().toISOString().slice(0, 10)
  } else if (campo === 'idCategoria') {
    chamado.idCategoria = Number(valor)
  } else if (campo === 'idPrioridade') {
    chamado.idPrioridade = Number(valor)
    // Trocar a prioridade recalcula o prazo previsto
    const prioridade = prioridades.find((p) => p.id === Number(valor))
    if (prioridade) {
      chamado.dataPrevista = new Date(Date.now() + prioridade.prazoHoras * 3_600_000)
        .toISOString()
        .slice(0, 10)
    }
  } else {
    chamado.idDepartamento = Number(valor)
  }

  return atraso({ sucesso: true, mensagem: 'Chamado atualizado com sucesso!' }, 450)
}

export async function finalizarChamado(id: number, mensagem: string, autor: string): Promise<ResultadoRh> {
  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)
  if (chamado.status === 'FINALIZADO') {
    return atraso({ sucesso: false, mensagem: 'Chamado já finalizado.' }, 200)
  }

  if (mensagem.trim()) {
    chamado.mensagens.push({ autor, ehAtendente: true, dataHora: agora(), texto: mensagem.trim() })
  }

  chamado.status = 'FINALIZADO'
  chamado.dataFinalizacao = new Date().toISOString().slice(0, 10)

  return atraso({ sucesso: true, mensagem: 'Chamado finalizado! O solicitante poderá avaliar.' }, 550)
}

export async function reabrirChamado(id: number, motivo: string, autor: string): Promise<ResultadoRh> {
  if (!motivo.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe o motivo da reabertura.' }, 200)
  }

  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)
  if (chamado.status !== 'FINALIZADO') {
    return atraso({ sucesso: false, mensagem: 'Só é possível reabrir chamados finalizados.' }, 200)
  }

  chamado.mensagens.push({ autor, ehAtendente: false, dataHora: agora(), texto: motivo.trim() })
  chamado.status = 'REABERTO'
  chamado.dataFinalizacao = ''
  chamado.avaliacao = null

  return atraso({ sucesso: true, mensagem: 'Chamado reaberto.' }, 500)
}

export async function avaliarChamado(
  id: number,
  nota: number,
  motivos: string[],
  comentario: string
): Promise<ResultadoRh> {
  if (!nota) return atraso({ sucesso: false, mensagem: 'Selecione a nota.' }, 200)

  const chamado = chamados.find((c) => c.id === id)
  if (!chamado) return atraso({ sucesso: false, mensagem: 'Chamado não encontrado.' }, 200)
  if (chamado.status !== 'FINALIZADO') {
    return atraso({ sucesso: false, mensagem: 'Só é possível avaliar chamados finalizados.' }, 200)
  }

  chamado.avaliacao = {
    nota,
    motivos: [...motivos],
    comentario: comentario.trim(),
    data: new Date().toISOString().slice(0, 10)
  }

  return atraso(
    {
      sucesso: true,
      mensagem: nota >= 4 ? 'Excelente! Adoramos saber disso.' : 'Obrigado — vamos buscar melhorar.'
    },
    550
  )
}

/* ------------------------------------------------------------------ *
 * Importação do ADP
 * ------------------------------------------------------------------ */

/** Plantas atendidas pela importação (SelecionaPlanta.asp) */
export const PLANTAS_ADP = [
  { codigo: '00', nome: 'Pirapetinga' },
  { codigo: '04', nome: 'Uberaba' },
  { codigo: '05', nome: 'Bento Gonçalves' },
  { codigo: '06', nome: 'Maranguape' },
  { codigo: '07', nome: 'Saquarema' }
]

export interface ResultadoImportacao extends ResultadoRh {
  linhasLidas: number
  linhasImportadas: number
  erros: string[]
}

/**
 * O legado processa dois CSVs (Funcionarios.csv e AcessosFuncionarios.csv)
 * e ainda roda as rotinas de matrículas e jornadas. Aqui só simulamos o
 * retorno, já que o parsing depende do arquivo real.
 */
export async function importarArquivoAdp(
  planta: string,
  tipo: 'funcionarios' | 'acessos',
  nomeArquivo: string
): Promise<ResultadoImportacao> {
  if (!planta) {
    return atraso(
      { sucesso: false, mensagem: 'Selecione a planta.', linhasLidas: 0, linhasImportadas: 0, erros: [] },
      200
    )
  }
  if (!nomeArquivo.trim()) {
    return atraso(
      {
        sucesso: false,
        mensagem: 'Selecione o arquivo CSV.',
        linhasLidas: 0,
        linhasImportadas: 0,
        erros: []
      },
      200
    )
  }
  if (!nomeArquivo.toLowerCase().endsWith('.csv')) {
    return atraso(
      {
        sucesso: false,
        mensagem: 'O arquivo precisa ser .csv.',
        linhasLidas: 0,
        linhasImportadas: 0,
        erros: []
      },
      200
    )
  }

  const lidas = tipo === 'funcionarios' ? 428 : 1_136
  const erros = tipo === 'funcionarios' ? ['Linha 87: matrícula em branco.'] : []

  return atraso(
    {
      sucesso: true,
      mensagem: `Importação de ${tipo} concluída.`,
      linhasLidas: lidas,
      linhasImportadas: lidas - erros.length,
      erros
    },
    1_400
  )
}
