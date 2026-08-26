import { atraso } from '@/services/comuns/baseMock'

/**
 * Compras/Classificar-Notas — VwNotasAnalises, MSAnalisesMaterial,
 * MSNotasAnalises e MSNotasConferidas (banco PIR-CORPORATIVO).
 *
 * Fluxo do legado: a lista traz as notas de insumo (classificadas ou não);
 * ao classificar, o usuário marca as análises aplicáveis, confirma
 * documentação e inspeção visual, e a nota passa a contar pontos.
 * Reclassificar apaga as análises anteriores e regrava.
 */

/** Filiais atendidas pela tela (lista literal do ListaNotasInsumo.asp) */
export const FILIAIS_NOTAS = [
  { codigo: '00', nome: 'Pirapetinga' },
  { codigo: '04', nome: 'Uberaba' },
  { codigo: '05', nome: 'Bento Gonçalves' },
  { codigo: '06', nome: 'Fortaleza' },
  { codigo: '07', nome: 'Saquarema' }
]

export function nomeFilial(codigo: string) {
  return FILIAIS_NOTAS.find((f) => f.codigo === codigo)?.nome ?? codigo
}

export interface NotaInsumo {
  idNota: number
  codigo: string
  razao: string
  notaFiscal: string
  serie: string
  item: string
  dataLancamento: string
  material: string
  descricao: string
  quantidade: number
  unidade: string
  filial: string
  /** Pontos obtidos nas análises marcadas */
  pontos: number
  /** Peso total possível */
  peso: number
  classificada: boolean
}

/** MSAnalisesMaterial — AN_TIPO 0 e 1 (o legado filtra AN_TIPO < 2) */
export interface AnaliseMaterial {
  codigo: string
  /** 0 = documentação, 1 = inspeção visual */
  tipo: 0 | 1
  descricao: string
  peso: number
}

export interface FiltrosNotas {
  filial: string
  razao: string
  notaFiscal: string
  serie: string
  /** '' = todas, 'S' = classificadas, 'N' = não classificadas */
  classificadas: string
}

export const ANALISES: AnaliseMaterial[] = [
  { codigo: 'D01', tipo: 0, descricao: 'Nota fiscal legível e sem rasuras', peso: 10 },
  { codigo: 'D02', tipo: 0, descricao: 'Dados do fornecedor conferem com o cadastro', peso: 10 },
  { codigo: 'D03', tipo: 0, descricao: 'Certificado de qualidade anexado', peso: 15 },
  { codigo: 'D04', tipo: 0, descricao: 'Ordem de compra referenciada na nota', peso: 10 },
  { codigo: 'D05', tipo: 0, descricao: 'Quantidade da nota confere com o pedido', peso: 15 },
  { codigo: 'V01', tipo: 1, descricao: 'Embalagem íntegra, sem avarias', peso: 10 },
  { codigo: 'V02', tipo: 1, descricao: 'Identificação do lote visível', peso: 10 },
  { codigo: 'V03', tipo: 1, descricao: 'Material sem sinais de umidade', peso: 10 },
  { codigo: 'V04', tipo: 1, descricao: 'Paletização adequada para o transporte', peso: 5 },
  { codigo: 'V05', tipo: 1, descricao: 'Ausência de contaminação por corpo estranho', peso: 5 }
]

const PESO_TOTAL = ANALISES.reduce((soma, analise) => soma + analise.peso, 0)

interface NotaInterna extends Omit<NotaInsumo, 'pontos' | 'peso' | 'classificada'> {
  /** Códigos das análises marcadas; vazio = não classificada */
  analisesMarcadas: string[]
  classificada: boolean
}

const FORNECEDORES = [
  { codigo: 'F00120', razao: 'PAPEL E CELULOSE ATLANTICO S.A.' },
  { codigo: 'F00234', razao: 'QUIMICA INDUSTRIAL MINAS LTDA' },
  { codigo: 'F00311', razao: 'TINTAS E VERNIZES SUL LTDA' },
  { codigo: 'F00405', razao: 'ADESIVOS TECNICOS BRASIL S.A.' },
  { codigo: 'F00512', razao: 'EMBALAGENS METALICAS NORDESTE LTDA' }
]

const MATERIAIS = [
  { material: '4000123', descricao: 'PAPEL KRAFTLINER 175G', unidade: 'KG' },
  { material: '4000456', descricao: 'AMIDO MODIFICADO PARA COLAGEM', unidade: 'KG' },
  { material: '4000789', descricao: 'TINTA FLEXOGRAFICA AZUL', unidade: 'LT' },
  { material: '4000912', descricao: 'ADESIVO HOT MELT', unidade: 'KG' },
  { material: '4001045', descricao: 'FITA DE ARQUEACAO POLIESTER', unidade: 'MT' }
]

function iso(dias: number) {
  return new Date(Date.now() - dias * 86_400_000).toISOString().slice(0, 10)
}

/** Massa determinística: 2 em cada 5 notas já vêm classificadas */
const notas: NotaInterna[] = Array.from({ length: 60 }, (_, indice) => {
  const fornecedor = FORNECEDORES[indice % FORNECEDORES.length]
  const material = MATERIAIS[indice % MATERIAIS.length]
  const jaClassificada = indice % 5 < 2

  return {
    idNota: 7000 + indice,
    codigo: fornecedor.codigo,
    razao: fornecedor.razao,
    notaFiscal: String(88_000 + indice * 7),
    serie: indice % 3 === 0 ? '1' : '2',
    item: String(1 + (indice % 4)).padStart(4, '0'),
    dataLancamento: iso(indice + 1),
    material: material.material,
    descricao: material.descricao,
    quantidade: 500 + ((indice * 137) % 4500),
    unidade: material.unidade,
    filial: FILIAIS_NOTAS[indice % FILIAIS_NOTAS.length].codigo,
    analisesMarcadas: jaClassificada
      ? ANALISES.filter((_, i) => (indice + i) % 3 !== 0).map((a) => a.codigo)
      : [],
    classificada: jaClassificada
  }
})

function paraLinha(nota: NotaInterna): NotaInsumo {
  const { analisesMarcadas, ...resto } = nota
  const pontos = ANALISES.filter((a) => analisesMarcadas.includes(a.codigo)).reduce(
    (soma, a) => soma + a.peso,
    0
  )
  return { ...resto, pontos, peso: PESO_TOTAL }
}

export async function listarNotas(filtros: FiltrosNotas): Promise<NotaInsumo[]> {
  const razao = filtros.razao.trim().toLowerCase()

  const resultado = notas
    .filter((n) => !filtros.filial || n.filial === filtros.filial)
    .filter((n) => !razao || n.razao.toLowerCase().includes(razao))
    .filter((n) => !filtros.notaFiscal.trim() || n.notaFiscal === filtros.notaFiscal.trim())
    .filter((n) => !filtros.serie.trim() || n.serie === filtros.serie.trim())
    .filter((n) =>
      filtros.classificadas === ''
        ? true
        : filtros.classificadas === 'S'
          ? n.classificada
          : !n.classificada
    )
    .sort((a, b) => b.dataLancamento.localeCompare(a.dataLancamento))
    .map(paraLinha)

  return atraso(resultado, 450)
}

export async function obterNota(idNota: number): Promise<{ nota: NotaInsumo; marcadas: string[] } | null> {
  const nota = notas.find((n) => n.idNota === idNota)
  if (!nota) return atraso(null, 200)

  return atraso({ nota: paraLinha(nota), marcadas: [...nota.analisesMarcadas] }, 250)
}

export async function listarAnalises(): Promise<AnaliseMaterial[]> {
  return atraso(ANALISES, 200)
}

export interface ResultadoClassificacao {
  sucesso: boolean
  mensagem: string
}

/**
 * Grava a classificação. No legado, salvar exige as duas confirmações
 * (documentação OK e inspeção visual OK) antes de inserir em
 * MSNotasAnalises/MSNotasConferidas.
 */
export async function classificarNota(
  idNota: number,
  marcadas: string[],
  documentacaoOk: boolean,
  visualOk: boolean
): Promise<ResultadoClassificacao> {
  if (!documentacaoOk || !visualOk) {
    return atraso(
      { sucesso: false, mensagem: 'Confirme a documentação e a inspeção visual antes de salvar.' },
      200
    )
  }

  const nota = notas.find((n) => n.idNota === idNota)
  if (!nota) return atraso({ sucesso: false, mensagem: 'Nota não encontrada.' }, 200)

  // Reclassificar sobrescreve as análises anteriores (DELETE + INSERT no ASP)
  nota.analisesMarcadas = [...marcadas]
  nota.classificada = true

  return atraso({ sucesso: true, mensagem: 'Nota classificada com sucesso!' }, 500)
}

/** Desfaz a classificação, liberando a nota para nova análise */
export async function reabrirNota(idNota: number): Promise<ResultadoClassificacao> {
  const nota = notas.find((n) => n.idNota === idNota)
  if (!nota) return atraso({ sucesso: false, mensagem: 'Nota não encontrada.' }, 200)

  nota.analisesMarcadas = []
  nota.classificada = false
  return atraso({ sucesso: true, mensagem: 'Nota liberada para reclassificação.' }, 400)
}

/**
 * ImportacaoFornecedores.asp / ImportacaoNotasFiscais.asp — no legado são
 * rotinas que puxam dados do Protheus. Aqui só simulam o retorno.
 */
export async function importarFornecedores(): Promise<ResultadoClassificacao> {
  return atraso({ sucesso: true, mensagem: `${FORNECEDORES.length} fornecedor(es) importado(s).` }, 900)
}

export async function importarNotas(): Promise<ResultadoClassificacao> {
  const naoClassificadas = notas.filter((n) => !n.classificada).length
  return atraso(
    { sucesso: true, mensagem: `Importação concluída. ${naoClassificadas} nota(s) aguardando classificação.` },
    1200
  )
}
