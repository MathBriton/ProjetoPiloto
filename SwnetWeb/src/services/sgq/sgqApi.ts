import { atraso } from '@/services/comuns/baseMock'

/**
 * SGQ — FNC Brasil (fluxo completo de não conformidade) e Ficha de QM
 * (devoluções).
 *
 * O nível do usuário na FNC vem da tabela Responsavel, a mesma do
 * Cadastros-FNC: Tipo 1 = Analista, Tipo 2 = Responsável Setorial.
 */

/* ------------------------------------------------------------------ *
 * FNC Brasil
 * ------------------------------------------------------------------ */

/** Situacao da FNC */
export const SITUACOES_FNC = [
  { codigo: '1', descricao: 'Registrada' },
  { codigo: '2', descricao: 'Em Análise' },
  { codigo: '3', descricao: 'Procede' },
  { codigo: '4', descricao: 'Não Procede' },
  { codigo: '5', descricao: 'Cancelada' }
]

/** Status do atendimento */
export const STATUS_FNC = [
  { codigo: '1', descricao: 'Aberta' },
  { codigo: '2', descricao: 'Respondida' },
  { codigo: '3', descricao: 'Encerrada' }
]

export const TIPOS_RECLAMACAO = ['NOTIFICACAO', 'FNC']

/** Defeito.Tipo — o mesmo do Cadastros-FNC */
export const TIPOS_DEFEITO = [
  { codigo: 'ADM', descricao: 'Administrativo' },
  { codigo: 'PRO', descricao: 'Industrial' }
]

export const DISPOSICOES = [
  { id: 1, codigo: 'D01', descricao: 'Substituição integral do material' },
  { id: 2, codigo: 'D02', descricao: 'Crédito ao cliente' },
  { id: 3, codigo: 'D03', descricao: 'Retrabalho na planta' },
  { id: 4, codigo: 'D04', descricao: 'Aceite com concessão' },
  { id: 5, codigo: 'D05', descricao: 'Devolução total' }
]

/**
 * Limites do coeficiente NQA. Abaixo do crítico a FNC é aprovada; entre
 * crítico e grave é "Crítico"; acima do grave é "Grave".
 */
export const LIMITES_NQA = { critico: 1.5, grave: 3 }

export type ClassificacaoNqa = 'APROVADO' | 'CRITICO' | 'GRAVE' | 'INDEFINIDO'

export interface ResultadoNqa {
  percentual: number
  classificacao: ClassificacaoNqa
  detalhe: string
}

/**
 * calcularCoeficienteNQA() do FNCFormularioAnalista.asp: só vale para
 * defeito industrial, e a quantidade reclamada não pode exceder o total
 * das notas vinculadas.
 */
export function calcularNqa(
  tipoDefeito: string,
  quantidadeReclamada: number,
  quantidadeTotalNotas: number
): ResultadoNqa {
  const indefinido: ResultadoNqa = { percentual: 0, classificacao: 'INDEFINIDO', detalhe: '' }

  if (tipoDefeito !== 'PRO') return indefinido
  if (
    !quantidadeTotalNotas ||
    quantidadeTotalNotas < 1 ||
    Number.isNaN(quantidadeReclamada) ||
    quantidadeReclamada < 0 ||
    quantidadeReclamada > quantidadeTotalNotas
  ) {
    return indefinido
  }

  const percentual = (quantidadeReclamada / quantidadeTotalNotas) * 100

  if (percentual < LIMITES_NQA.critico) {
    return {
      percentual,
      classificacao: 'APROVADO',
      detalhe: `Menor do que ${LIMITES_NQA.critico}%`
    }
  }
  if (percentual < LIMITES_NQA.grave) {
    return {
      percentual,
      classificacao: 'CRITICO',
      detalhe: `Maior ou igual a ${LIMITES_NQA.critico}%`
    }
  }
  return { percentual, classificacao: 'GRAVE', detalhe: `Maior ou igual a ${LIMITES_NQA.grave}%` }
}

export interface NotaVinculada {
  numero: string
  serie: string
  quantidade: number
  valor: number
}

export interface RespostaFnc {
  dataEnvio: string
  dataResposta: string
  usuarioEnvio: string
  usuarioResposta: string
  procede: boolean | null
  descricao: string
  justificativa: string
}

export interface Fnc {
  id: number
  dataInclusao: string
  codPlanta: string
  tipoReclamacao: string
  idSetor: number
  nomeSetor: string
  responsavel: string
  codCliente: string
  codLoja: string
  nomeCliente: string
  cidadeEstado: string
  vendedor: string
  dataOcorrencia: string
  tipoDefeito: string
  defeitoId: number
  descricaoDefeito: string
  disposicaoId: number
  descricaoDisposicao: string
  quantidadeReclamada: number
  valorReclamado: number
  contato: string
  email: string
  telefone: string
  aprovada: boolean
  situacao: string
  status: string
  notas: NotaVinculada[]
  respostas: RespostaFnc[]
  eficacia: string
}

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

const CLIENTES_SGQ = [
  { cod: '001234', loja: '01', nome: 'INDUSTRIA ALIMENTICIA BOA VISTA S.A.', cidade: 'JUIZ DE FORA/MG' },
  { cod: '002567', loja: '01', nome: 'BEBIDAS DO VALE LTDA', cidade: 'CATAGUASES/MG' },
  { cod: '003891', loja: '01', nome: 'FRIGORIFICO SERRA AZUL LTDA', cidade: 'BENTO GONCALVES/RS' },
  { cod: '004125', loja: '01', nome: 'HORTIFRUTI NORDESTE COMERCIO S.A.', cidade: 'MARANGUAPE/CE' }
]

const SETORES_SGQ = [
  { id: 1, nome: 'Produção' },
  { id: 2, nome: 'Qualidade' },
  { id: 3, nome: 'Expedição' },
  { id: 4, nome: 'Comercial' }
]

const DEFEITOS_SGQ = [
  { id: 1, tipo: 'PRO', descricao: 'Colagem deficiente' },
  { id: 2, tipo: 'PRO', descricao: 'Impressão fora de registro' },
  { id: 4, tipo: 'PRO', descricao: 'Gramatura fora de especificação' },
  { id: 8, tipo: 'ADM', descricao: 'Divergência de quantidade no pedido' },
  { id: 9, tipo: 'ADM', descricao: 'Preço divergente da tabela' }
]

export function listarSetoresSgq() {
  return SETORES_SGQ
}

export function listarDefeitosPorTipo(tipo: string) {
  return DEFEITOS_SGQ.filter((d) => !tipo || d.tipo === tipo)
}

const RESPONSAVEIS = ['jsilva', 'malves', 'bcarvalho']

const fncs: Fnc[] = Array.from({ length: 26 }, (_, indice) => {
  const cliente = CLIENTES_SGQ[indice % CLIENTES_SGQ.length]
  const setor = SETORES_SGQ[indice % SETORES_SGQ.length]
  const defeito = DEFEITOS_SGQ[indice % DEFEITOS_SGQ.length]
  const disposicao = DISPOSICOES[indice % DISPOSICOES.length]
  const situacao = SITUACOES_FNC[indice % SITUACOES_FNC.length]
  const quantidadeNota = 5_000 + ((indice * 431) % 20_000)
  const reclamada = Math.round(quantidadeNota * (0.005 + ((indice % 9) / 100)))

  return {
    id: 6000 + indice,
    dataInclusao: iso(indice + 1),
    codPlanta: ['00', '04', '05', '06', '07'][indice % 5],
    tipoReclamacao: TIPOS_RECLAMACAO[indice % 2],
    idSetor: setor.id,
    nomeSetor: setor.nome,
    responsavel: RESPONSAVEIS[indice % RESPONSAVEIS.length],
    codCliente: cliente.cod,
    codLoja: cliente.loja,
    nomeCliente: cliente.nome,
    cidadeEstado: cliente.cidade,
    vendedor: ['cpereira', 'rnunes'][indice % 2],
    dataOcorrencia: iso(indice + 4),
    tipoDefeito: defeito.tipo,
    defeitoId: defeito.id,
    descricaoDefeito: defeito.descricao,
    disposicaoId: disposicao.id,
    descricaoDisposicao: disposicao.descricao,
    quantidadeReclamada: reclamada,
    valorReclamado: Number((reclamada * 2.35).toFixed(2)),
    contato: 'Depto. de Qualidade',
    email: 'qualidade@cliente.com.br',
    telefone: '(32) 3222-1000',
    aprovada: indice % 3 === 0,
    situacao: situacao.codigo,
    status: STATUS_FNC[indice % STATUS_FNC.length].codigo,
    notas: [
      {
        numero: String(90_000 + indice * 3),
        serie: '1',
        quantidade: quantidadeNota,
        valor: Number((quantidadeNota * 2.35).toFixed(2))
      }
    ],
    respostas:
      indice % 3 === 1
        ? [
            {
              dataEnvio: iso(indice),
              dataResposta: iso(Math.max(0, indice - 2)),
              usuarioEnvio: 'plima',
              usuarioResposta: RESPONSAVEIS[indice % RESPONSAVEIS.length],
              procede: indice % 2 === 0,
              descricao: 'Análise da causa raiz concluída.',
              justificativa: 'Falha pontual no ajuste da coladeira, já corrigida.'
            }
          ]
        : [],
    eficacia: ''
  }
})

export interface FiltrosFnc {
  planta: string
  situacao: string
  status: string
  aprovada: string
  tipoReclamacao: string
  idFnc: string
  apenasMinhasPendencias: boolean
}

export const FILTROS_FNC_VAZIOS: FiltrosFnc = {
  planta: '',
  situacao: '',
  status: '',
  aprovada: '',
  tipoReclamacao: '',
  idFnc: '',
  apenasMinhasPendencias: false
}

export async function listarFncs(filtros: FiltrosFnc, usuario: string): Promise<Fnc[]> {
  const resultado = fncs
    .filter((f) => !filtros.planta || f.codPlanta === filtros.planta)
    .filter((f) => !filtros.situacao || f.situacao === filtros.situacao)
    .filter((f) => !filtros.status || f.status === filtros.status)
    .filter((f) => filtros.aprovada === '' || f.aprovada === (filtros.aprovada === '1'))
    .filter((f) => !filtros.tipoReclamacao || f.tipoReclamacao === filtros.tipoReclamacao)
    .filter((f) => !filtros.idFnc.trim() || String(f.id).includes(filtros.idFnc.trim()))
    .filter((f) => !filtros.apenasMinhasPendencias || (f.responsavel === usuario && f.status !== '3'))
    .map((f) => ({ ...f }))
    .sort((a, b) => b.dataInclusao.localeCompare(a.dataInclusao))

  return atraso(resultado, 450)
}

export async function obterFnc(id: number): Promise<Fnc | null> {
  const encontrada = fncs.find((f) => f.id === id)
  return atraso(encontrada ? { ...encontrada } : null, 300)
}

export interface DadosFnc {
  id?: number
  codPlanta: string
  tipoReclamacao: string
  idSetor: number | ''
  responsavel: string
  codCliente: string
  codLoja: string
  nomeCliente: string
  cidadeEstado: string
  vendedor: string
  dataOcorrencia: string
  tipoDefeito: string
  defeitoId: number | ''
  descricaoDefeito: string
  disposicaoId: number | ''
  descricaoDisposicao: string
  quantidadeReclamada: string
  valorReclamado: string
  contato: string
  email: string
  telefone: string
  notas: NotaVinculada[]
}

export const DADOS_FNC_VAZIOS: DadosFnc = {
  codPlanta: '',
  tipoReclamacao: '',
  idSetor: '',
  responsavel: '',
  codCliente: '',
  codLoja: '',
  nomeCliente: '',
  cidadeEstado: '',
  vendedor: '',
  dataOcorrencia: '',
  tipoDefeito: '',
  defeitoId: '',
  descricaoDefeito: '',
  disposicaoId: '',
  descricaoDisposicao: '',
  quantidadeReclamada: '',
  valorReclamado: '',
  contato: '',
  email: '',
  telefone: '',
  notas: []
}

export interface ResultadoSgq {
  sucesso: boolean
  mensagem: string
}

export async function salvarFnc(dados: DadosFnc, usuario: string): Promise<ResultadoSgq> {
  const obrigatorios: [keyof DadosFnc, string][] = [
    ['codPlanta', 'Planta'],
    ['tipoReclamacao', 'Tipo de Reclamação'],
    ['nomeCliente', 'Nome do Cliente'],
    ['codCliente', 'Cód. Cliente'],
    ['vendedor', 'Vendedor'],
    ['dataOcorrencia', 'Data de Ocorrência'],
    ['tipoDefeito', 'Tipo de Defeito'],
    ['descricaoDefeito', 'Descrição do Defeito']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }
  if (!dados.defeitoId) {
    return atraso({ sucesso: false, mensagem: 'Classificação do Defeito é obrigatória.' }, 200)
  }
  if (!dados.disposicaoId) {
    return atraso({ sucesso: false, mensagem: 'Código de Disposição é obrigatório.' }, 200)
  }

  const totalNotas = dados.notas.reduce((soma, nota) => soma + nota.quantidade, 0)
  const reclamada = Number(dados.quantidadeReclamada || 0)

  // A regra do NQA exige que a reclamação não passe do total das notas
  if (dados.tipoDefeito === 'PRO' && totalNotas > 0 && reclamada > totalNotas) {
    return atraso(
      { sucesso: false, mensagem: 'Quantidade reclamada maior que o total das notas vinculadas.' },
      200
    )
  }

  const setor = SETORES_SGQ.find((s) => s.id === Number(dados.idSetor))
  const disposicao = DISPOSICOES.find((d) => d.id === Number(dados.disposicaoId))

  const campos = {
    codPlanta: dados.codPlanta,
    tipoReclamacao: dados.tipoReclamacao,
    idSetor: Number(dados.idSetor || 0),
    nomeSetor: setor?.nome ?? '',
    responsavel: dados.responsavel,
    codCliente: dados.codCliente.trim(),
    codLoja: dados.codLoja.trim(),
    nomeCliente: dados.nomeCliente.trim(),
    cidadeEstado: dados.cidadeEstado.trim(),
    vendedor: dados.vendedor.trim(),
    dataOcorrencia: dados.dataOcorrencia,
    tipoDefeito: dados.tipoDefeito,
    defeitoId: Number(dados.defeitoId),
    descricaoDefeito: dados.descricaoDefeito.trim(),
    disposicaoId: Number(dados.disposicaoId),
    descricaoDisposicao: disposicao?.descricao ?? dados.descricaoDisposicao,
    quantidadeReclamada: reclamada,
    valorReclamado: Number(dados.valorReclamado || 0),
    contato: dados.contato.trim(),
    email: dados.email.trim(),
    telefone: dados.telefone.trim(),
    notas: [...dados.notas]
  }

  if (dados.id) {
    const fnc = fncs.find((f) => f.id === dados.id)
    if (!fnc) return atraso({ sucesso: false, mensagem: 'FNC não encontrada.' }, 200)

    Object.assign(fnc, campos)
    return atraso({ sucesso: true, mensagem: 'FNC atualizada com sucesso!' }, 500)
  }

  fncs.push({
    id: Math.max(0, ...fncs.map((f) => f.id)) + 1,
    dataInclusao: new Date().toISOString().slice(0, 10),
    aprovada: false,
    situacao: '1',
    status: '1',
    respostas: [],
    eficacia: '',
    ...campos,
    responsavel: campos.responsavel || usuario
  })

  return atraso({ sucesso: true, mensagem: 'FNC registrada com sucesso!' }, 550)
}

/** enviarParaResponsavel — passa a FNC de Registrada para Em Análise */
export async function enviarParaResponsavel(id: number, responsavel: string): Promise<ResultadoSgq> {
  const fnc = fncs.find((f) => f.id === id)
  if (!fnc) return atraso({ sucesso: false, mensagem: 'FNC não encontrada.' }, 200)
  if (!responsavel) {
    return atraso({ sucesso: false, mensagem: 'Selecione o responsável pelo setor.' }, 200)
  }

  fnc.responsavel = responsavel
  fnc.situacao = '2'
  fnc.status = '1'
  return atraso({ sucesso: true, mensagem: `FNC encaminhada para ${responsavel}.` }, 500)
}

/** responder — o responsável setorial informa se a reclamação procede */
export async function responderFnc(
  id: number,
  procede: boolean,
  descricao: string,
  justificativa: string,
  usuario: string
): Promise<ResultadoSgq> {
  const fnc = fncs.find((f) => f.id === id)
  if (!fnc) return atraso({ sucesso: false, mensagem: 'FNC não encontrada.' }, 200)
  if (!justificativa.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a justificativa da resposta.' }, 200)
  }

  fnc.respostas.push({
    dataEnvio: fnc.dataInclusao,
    dataResposta: new Date().toISOString().slice(0, 10),
    usuarioEnvio: fnc.responsavel,
    usuarioResposta: usuario,
    procede,
    descricao: descricao.trim(),
    justificativa: justificativa.trim()
  })

  fnc.situacao = procede ? '3' : '4'
  fnc.status = '2'
  return atraso({ sucesso: true, mensagem: 'Resposta registrada com sucesso!' }, 550)
}

/** avaliar — o analista fecha a FNC informando a eficácia */
export async function avaliarFnc(
  id: number,
  eficacia: string,
  aprovada: boolean
): Promise<ResultadoSgq> {
  const fnc = fncs.find((f) => f.id === id)
  if (!fnc) return atraso({ sucesso: false, mensagem: 'FNC não encontrada.' }, 200)
  if (!eficacia.trim()) {
    return atraso({ sucesso: false, mensagem: 'Informe a eficácia/observação da avaliação.' }, 200)
  }

  fnc.eficacia = eficacia.trim()
  fnc.aprovada = aprovada
  fnc.status = '3'
  return atraso({ sucesso: true, mensagem: 'Avaliação registrada. FNC encerrada.' }, 550)
}

export async function cancelarFnc(id: number): Promise<ResultadoSgq> {
  const fnc = fncs.find((f) => f.id === id)
  if (!fnc) return atraso({ sucesso: false, mensagem: 'FNC não encontrada.' }, 200)
  if (fnc.status === '3') {
    return atraso({ sucesso: false, mensagem: 'FNC encerrada não pode ser cancelada.' }, 200)
  }

  fnc.situacao = '5'
  fnc.status = '3'
  return atraso({ sucesso: true, mensagem: 'FNC cancelada com sucesso!' }, 450)
}

/* ------------------------------------------------------------------ *
 * Ficha de QM — devoluções
 * ------------------------------------------------------------------ */

export const STATUS_QM = [
  { codigo: 'AGUARDANDO_RETRABALHO', descricao: 'Aguardando Retrabalho' },
  { codigo: 'AGUARDANDO_FATURAMENTO', descricao: 'Aguardando Faturamento' },
  { codigo: 'RETRABALHO_FINALIZADO', descricao: 'Retrabalho Finalizado' },
  { codigo: 'REFATURADO_TOTAL', descricao: 'Refaturado Total' },
  { codigo: 'REFUGADO_TOTAL', descricao: 'Refugado Total' },
  { codigo: 'SIMBOLICO_REFATURADO', descricao: 'Simbólico Refaturado' },
  { codigo: 'SIMBOLICO_REFUGADO', descricao: 'Simbólico Refugado' }
]

export function descricaoStatusQm(codigo: string) {
  return STATUS_QM.find((s) => s.codigo === codigo)?.descricao ?? codigo
}

export interface FichaQm {
  id: number
  codPlanta: string
  nomeCliente: string
  idFnc: number | null
  status: string
  quantidadePaletes: number
  quantidadeRecebida: number
  quantidadeRefugada: number
  valor: number
  valorUnitarioMaterial: number
  nfDevolucao: string
  codigoVa: string
  codigoQm: string
  modoDefeito: string
  area: string
  ordemProducao: string
  dataEntradaDevolucao: string
  dataConclusao: string
  dataCriacao: string
}

const AREAS_QM = ['Onduladeira', 'Impressão', 'Colagem', 'Acabamento', 'Expedição']
const MODOS_DEFEITO = ['Empeno', 'Delaminação', 'Mancha de tinta', 'Corte irregular', 'Cola insuficiente']

interface FichaQmInterna extends FichaQm {
  excluido: boolean
}

const fichas: FichaQmInterna[] = Array.from({ length: 20 }, (_, indice) => {
  const cliente = CLIENTES_SGQ[indice % CLIENTES_SGQ.length]
  const recebida = 500 + ((indice * 173) % 4_000)
  const refugada = Math.round(recebida * (0.05 + (indice % 7) / 100))
  const valorUnitario = Number((1.8 + (indice % 12) / 10).toFixed(2))

  return {
    id: 7500 + indice,
    codPlanta: ['00', '04', '05', '06', '07'][indice % 5],
    nomeCliente: cliente.nome,
    idFnc: indice % 3 === 0 ? 6000 + indice : null,
    status: STATUS_QM[indice % STATUS_QM.length].codigo,
    quantidadePaletes: 2 + (indice % 12),
    quantidadeRecebida: recebida,
    quantidadeRefugada: refugada,
    valor: Number((recebida * valorUnitario).toFixed(2)),
    valorUnitarioMaterial: valorUnitario,
    nfDevolucao: String(55_000 + indice * 9),
    codigoVa: `VA${String(2000 + indice).slice(-4)}`,
    codigoQm: `QM${String(3000 + indice).slice(-4)}`,
    modoDefeito: MODOS_DEFEITO[indice % MODOS_DEFEITO.length],
    area: AREAS_QM[indice % AREAS_QM.length],
    ordemProducao: String(880_000 + indice * 17),
    dataEntradaDevolucao: iso(indice + 5),
    dataConclusao: indice % 4 === 0 ? '' : iso(indice),
    dataCriacao: iso(indice + 6),
    excluido: false
  }
})

export interface FiltrosQm {
  planta: string
  dataCriacaoDe: string
  dataCriacaoAte: string
  nomeCliente: string
  status: string
  id: string
  idFnc: string
  ordemProducao: string
  nfDevolucao: string
}

export const FILTROS_QM_VAZIOS: FiltrosQm = {
  planta: '',
  dataCriacaoDe: '',
  dataCriacaoAte: '',
  nomeCliente: '',
  status: '',
  id: '',
  idFnc: '',
  ordemProducao: '',
  nfDevolucao: ''
}

export async function listarFichasQm(filtros: FiltrosQm): Promise<FichaQm[]> {
  const cliente = filtros.nomeCliente.trim().toLowerCase()

  const resultado = fichas
    .filter((f) => !f.excluido)
    .filter((f) => !filtros.planta || f.codPlanta === filtros.planta)
    .filter((f) => !filtros.status || f.status === filtros.status)
    .filter((f) => !cliente || f.nomeCliente.toLowerCase().includes(cliente))
    .filter((f) => !filtros.id.trim() || String(f.id).includes(filtros.id.trim()))
    .filter((f) => !filtros.idFnc.trim() || String(f.idFnc ?? '').includes(filtros.idFnc.trim()))
    .filter((f) => !filtros.ordemProducao.trim() || f.ordemProducao.includes(filtros.ordemProducao.trim()))
    .filter((f) => !filtros.nfDevolucao.trim() || f.nfDevolucao.includes(filtros.nfDevolucao.trim()))
    .filter((f) => !filtros.dataCriacaoDe || f.dataCriacao >= filtros.dataCriacaoDe)
    .filter((f) => !filtros.dataCriacaoAte || f.dataCriacao <= filtros.dataCriacaoAte)
    .map(({ excluido, ...resto }) => resto)
    .sort((a, b) => b.dataCriacao.localeCompare(a.dataCriacao))

  return atraso(resultado, 400)
}

export interface DadosFichaQm {
  id?: number
  codPlanta: string
  idFnc: string
  status: string
  nomeCliente: string
  modoDefeito: string
  area: string
  ordemProducao: string
  dataEntradaDevolucao: string
  dataConclusao: string
  nfDevolucao: string
  codigoVa: string
  codigoQm: string
  quantidadePaletes: string
  quantidadeRecebida: string
  quantidadeRefugada: string
  valor: string
  valorUnitarioMaterial: string
}

export const DADOS_QM_VAZIOS: DadosFichaQm = {
  codPlanta: '',
  idFnc: '',
  status: '',
  nomeCliente: '',
  modoDefeito: '',
  area: '',
  ordemProducao: '',
  dataEntradaDevolucao: '',
  dataConclusao: '',
  nfDevolucao: '',
  codigoVa: '',
  codigoQm: '',
  quantidadePaletes: '',
  quantidadeRecebida: '',
  quantidadeRefugada: '',
  valor: '',
  valorUnitarioMaterial: ''
}

export async function salvarFichaQm(dados: DadosFichaQm): Promise<ResultadoSgq> {
  const obrigatorios: [keyof DadosFichaQm, string][] = [
    ['codPlanta', 'Planta'],
    ['status', 'Status'],
    ['dataEntradaDevolucao', 'Data Entrada Devolução'],
    ['dataConclusao', 'Data de Conclusão'],
    ['nfDevolucao', 'NF Devolução'],
    ['codigoVa', 'Código VA'],
    ['codigoQm', 'Código QM'],
    ['quantidadePaletes', 'Quantidade Paletes'],
    ['quantidadeRecebida', 'Quantidade Recebida'],
    ['quantidadeRefugada', 'Quantidade Refugada'],
    ['valor', 'Valor']
  ]

  for (const [campo, rotulo] of obrigatorios) {
    if (!String(dados[campo] ?? '').trim()) {
      return atraso({ sucesso: false, mensagem: `${rotulo} é obrigatório.` }, 200)
    }
  }

  const recebida = Number(dados.quantidadeRecebida)
  const refugada = Number(dados.quantidadeRefugada)
  if (refugada > recebida) {
    return atraso(
      { sucesso: false, mensagem: 'Quantidade refugada não pode ser maior que a recebida.' },
      200
    )
  }
  if (dados.dataConclusao < dados.dataEntradaDevolucao) {
    return atraso(
      { sucesso: false, mensagem: 'Data de conclusão não pode ser anterior à entrada da devolução.' },
      200
    )
  }

  const campos = {
    codPlanta: dados.codPlanta,
    idFnc: dados.idFnc.trim() ? Number(dados.idFnc) : null,
    status: dados.status,
    nomeCliente: dados.nomeCliente.trim(),
    modoDefeito: dados.modoDefeito.trim(),
    area: dados.area.trim(),
    ordemProducao: dados.ordemProducao.trim(),
    dataEntradaDevolucao: dados.dataEntradaDevolucao,
    dataConclusao: dados.dataConclusao,
    nfDevolucao: dados.nfDevolucao.trim(),
    codigoVa: dados.codigoVa.trim().toUpperCase(),
    codigoQm: dados.codigoQm.trim().toUpperCase(),
    quantidadePaletes: Number(dados.quantidadePaletes),
    quantidadeRecebida: recebida,
    quantidadeRefugada: refugada,
    valor: Number(dados.valor),
    valorUnitarioMaterial: Number(dados.valorUnitarioMaterial || 0)
  }

  if (dados.id) {
    const ficha = fichas.find((f) => f.id === dados.id)
    if (!ficha) return atraso({ sucesso: false, mensagem: 'Ficha não encontrada.' }, 200)

    Object.assign(ficha, campos)
    return atraso({ sucesso: true, mensagem: 'Ficha atualizada com sucesso!' }, 500)
  }

  fichas.push({
    id: Math.max(0, ...fichas.map((f) => f.id)) + 1,
    dataCriacao: new Date().toISOString().slice(0, 10),
    excluido: false,
    ...campos
  })

  return atraso({ sucesso: true, mensagem: 'Ficha criada com sucesso!' }, 500)
}

export async function excluirFichaQm(id: number): Promise<ResultadoSgq> {
  const ficha = fichas.find((f) => f.id === id)
  if (!ficha) return atraso({ sucesso: false, mensagem: 'Ficha não encontrada.' }, 200)

  ficha.excluido = true
  return atraso({ sucesso: true, mensagem: 'Ficha excluída com sucesso!' }, 400)
}
