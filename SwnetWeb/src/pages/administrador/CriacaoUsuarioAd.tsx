import { useState } from 'react'
import { ArrowLeft, ContactRound, IdCard, Search, Send, UserPlus } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, EstadoVazio, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import {
  DEPARTAMENTOS,
  GRUPOS_AD,
  SOLICITACAO_VAZIA,
  buscarFicha,
  solicitarCriacao,
  sugerirEmail,
  sugerirLogin,
  type FichaUsuario,
  type SolicitacaoUsuarioAd
} from '@/services/administrador/criacaoUsuarioAdApi'
import { EMPRESAS, codigoEmpresa, formatarData } from '@/services/comuns/baseMock'

/**
 * Refatoração de Criacao-Usuarios-AD.
 *
 * O `index.html` legado é só um hub com dois cards ("Criação de Usuário" e
 * "Ficha de Usuário"). As duas telas que ele linka (creation_user.html e
 * card.html) NÃO vieram no material de refatoração — os campos do formulário
 * de criação e da ficha abaixo são uma reconstrução plausível e precisam ser
 * conferidos contra as telas originais.
 */

type Visao = 'hub' | 'criacao' | 'ficha'

export function CriacaoUsuarioAd() {
  const [visao, setVisao] = useState<Visao>('hub')

  return (
    <div className="space-y-4">
      <TituloPagina
        titulo="Criação de Usuários (AD)"
        acoes={
          visao !== 'hub' && (
            <Botao variante="neutro" icone={ArrowLeft} aoClicar={() => setVisao('hub')}>
              Voltar
            </Botao>
          )
        }
      />

      {visao === 'hub' && <Hub aoEscolher={setVisao} />}
      {visao === 'criacao' && <FormularioCriacao />}
      {visao === 'ficha' && <ConsultaFicha />}
    </div>
  )
}

function Hub({ aoEscolher }: { aoEscolher: (visao: Visao) => void }) {
  const opcoes = [
    {
      visao: 'criacao' as const,
      icone: UserPlus,
      titulo: 'Criação de Usuário',
      descricao: 'Solicitar uma nova conta no Active Directory.'
    },
    {
      visao: 'ficha' as const,
      icone: ContactRound,
      titulo: 'Ficha de Usuário',
      descricao: 'Consultar os dados consolidados de uma conta existente.'
    }
  ]

  return (
    <div className="mx-auto grid max-w-3xl gap-4 pt-6 sm:grid-cols-2">
      {opcoes.map(({ visao, icone: Icone, titulo, descricao }) => (
        <button
          key={visao}
          type="button"
          onClick={() => aoEscolher(visao)}
          className="flex flex-col items-center gap-3 rounded-xl border border-slate-200 bg-white px-6 py-10
            text-center shadow-sm transition-all hover:-translate-y-0.5 hover:border-swnet-300 hover:shadow-md
            dark:border-slate-800 dark:bg-slate-900 dark:hover:border-swnet-700"
        >
          <Icone size={38} strokeWidth={1.3} className="text-swnet-600 dark:text-swnet-400" />
          <span className="text-base font-semibold text-slate-800 dark:text-white">{titulo}</span>
          <span className="text-xs text-slate-500 dark:text-slate-400">{descricao}</span>
        </button>
      ))}
    </div>
  )
}

function FormularioCriacao() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const [dados, setDados] = useState<SolicitacaoUsuarioAd>(SOLICITACAO_VAZIA)
  const [enviando, setEnviando] = useState(false)

  function mudar<C extends keyof SolicitacaoUsuarioAd>(campo: C, valor: SolicitacaoUsuarioAd[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  /** Login e e-mail são derivados do nome enquanto não forem editados à mão */
  function mudarNome(campo: 'nome' | 'sobrenome', valor: string) {
    setDados((atual) => {
      const proximo = { ...atual, [campo]: valor }
      const loginSugeridoAtual = sugerirLogin(atual.nome, atual.sobrenome)
      const podeSugerir = !atual.login || atual.login === loginSugeridoAtual

      if (podeSugerir) {
        const novoLogin = sugerirLogin(proximo.nome, proximo.sobrenome)
        proximo.login = novoLogin
        if (!atual.email || atual.email === sugerirEmail(loginSugeridoAtual)) {
          proximo.email = sugerirEmail(novoLogin)
        }
      }
      return proximo
    })
  }

  function alternarGrupo(grupo: string) {
    setDados((atual) => ({
      ...atual,
      grupos: atual.grupos.includes(grupo)
        ? atual.grupos.filter((g) => g !== grupo)
        : [...atual.grupos, grupo]
    }))
  }

  async function enviar() {
    setEnviando(true)
    try {
      const resultado = await solicitarCriacao(dados)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
      if (resultado.sucesso) setDados(SOLICITACAO_VAZIA)
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div className="mx-auto max-w-4xl">
      <Painel titulo="Solicitação de Conta no AD" icone={UserPlus}>
        <form
          onSubmit={(e) => {
            e.preventDefault()
            enviar()
          }}
        >
          <Secao titulo="Identificação" icone={IdCard}>
            <div className="grid gap-4 sm:grid-cols-12">
              <Entrada
                id="adNome"
                rotulo="Nome"
                required
                value={dados.nome}
                disabled={enviando}
                onChange={(e) => mudarNome('nome', e.target.value)}
                larguraClasse="sm:col-span-4"
              />
              <Entrada
                id="adSobrenome"
                rotulo="Sobrenome"
                required
                value={dados.sobrenome}
                disabled={enviando}
                onChange={(e) => mudarNome('sobrenome', e.target.value)}
                larguraClasse="sm:col-span-5"
              />
              <Entrada
                id="adMatricula"
                rotulo="Matrícula"
                maxLength={6}
                value={dados.matricula}
                disabled={enviando}
                onChange={(e) => mudar('matricula', e.target.value.replace(/\D/g, ''))}
                larguraClasse="sm:col-span-3"
              />

              <Selecao
                id="adDominio"
                rotulo="Domínio"
                value={dados.dominio}
                disabled={enviando}
                onChange={(e) => mudar('dominio', e.target.value as SolicitacaoUsuarioAd['dominio'])}
                larguraClasse="sm:col-span-3"
              >
                <option value="GROUP">GROUP</option>
                <option value="WESTROCK">WESTROCK</option>
              </Selecao>
              <Entrada
                id="adLogin"
                rotulo="Login"
                required
                value={dados.login}
                disabled={enviando}
                ajuda="Sugerido a partir do nome; pode ser ajustado."
                onChange={(e) => mudar('login', e.target.value.trim().toLowerCase())}
                larguraClasse="sm:col-span-4"
              />
              <Entrada
                id="adEmail"
                rotulo="E-mail"
                type="email"
                required
                value={dados.email}
                disabled={enviando}
                onChange={(e) => mudar('email', e.target.value)}
                larguraClasse="sm:col-span-5"
              />
            </div>
          </Secao>

          <Secao titulo="Lotação" icone={ContactRound}>
            <div className="grid gap-4 sm:grid-cols-12">
              <Selecao
                id="adPlanta"
                rotulo="Planta"
                required
                value={dados.planta}
                disabled={enviando}
                onChange={(e) => mudar('planta', e.target.value)}
                larguraClasse="sm:col-span-4"
              >
                <option value="">Selecione...</option>
                {EMPRESAS.map((empresa) => (
                  <option key={empresa.id} value={codigoEmpresa(empresa.id)}>
                    {empresa.nome}
                  </option>
                ))}
              </Selecao>

              <Selecao
                id="adDepartamento"
                rotulo="Departamento"
                required
                value={dados.departamento}
                disabled={enviando}
                onChange={(e) => mudar('departamento', e.target.value)}
                larguraClasse="sm:col-span-4"
              >
                <option value="">Selecione...</option>
                {DEPARTAMENTOS.map((departamento) => (
                  <option key={departamento} value={departamento}>
                    {departamento}
                  </option>
                ))}
              </Selecao>

              <Entrada
                id="adCargo"
                rotulo="Cargo"
                value={dados.cargo}
                disabled={enviando}
                onChange={(e) => mudar('cargo', e.target.value)}
                larguraClasse="sm:col-span-4"
              />

              <Entrada
                id="adGestor"
                rotulo="Gestor responsável"
                required
                value={dados.gestor}
                disabled={enviando}
                onChange={(e) => mudar('gestor', e.target.value)}
                larguraClasse="sm:col-span-8"
              />
              <Entrada
                id="adAdmissao"
                rotulo="Data de admissão"
                type="date"
                value={dados.dataAdmissao}
                disabled={enviando}
                onChange={(e) => mudar('dataAdmissao', e.target.value)}
                larguraClasse="sm:col-span-4"
              />
            </div>
          </Secao>

          <Secao titulo="Acessos" icone={Send}>
            <Campo rotulo="Grupos do AD" obrigatorio>
              <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
                {GRUPOS_AD.map((grupo) => (
                  <label
                    key={grupo}
                    className="flex cursor-pointer items-center gap-2 rounded-lg border border-slate-200 px-3 py-2
                      text-xs text-slate-600 hover:border-swnet-300 dark:border-slate-700 dark:text-slate-300"
                  >
                    <input
                      type="checkbox"
                      checked={dados.grupos.includes(grupo)}
                      disabled={enviando}
                      onChange={() => alternarGrupo(grupo)}
                      className="h-4 w-4 accent-swnet-600"
                    />
                    <span className="truncate">{grupo}</span>
                  </label>
                ))}
              </div>
            </Campo>

            <div className="mt-4">
              <Campo rotulo="Observações" htmlFor="adObservacoes">
                <textarea
                  id="adObservacoes"
                  rows={3}
                  value={dados.observacoes}
                  disabled={enviando}
                  placeholder="Informações adicionais para a equipe de TI..."
                  onChange={(e) => mudar('observacoes', e.target.value)}
                  className={classesControle}
                />
              </Campo>
            </div>
          </Secao>

          <div className="mt-4 flex justify-end">
            <Botao tipo="submit" variante="sucesso" icone={Send} desabilitado={enviando}>
              {enviando ? 'Enviando...' : 'Enviar solicitação'}
            </Botao>
          </div>
        </form>
      </Painel>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

function ConsultaFicha() {
  const [termo, setTermo] = useState('')
  const [resultados, setResultados] = useState<FichaUsuario[]>([])
  const [buscando, setBuscando] = useState(false)
  const [buscaFeita, setBuscaFeita] = useState(false)

  async function buscar() {
    setBuscando(true)
    try {
      setResultados(await buscarFicha(termo))
      setBuscaFeita(true)
    } finally {
      setBuscando(false)
    }
  }

  return (
    <div className="mx-auto max-w-4xl space-y-4">
      <Painel titulo="Ficha de Usuário" icone={ContactRound}>
        <form
          className="flex flex-wrap items-end gap-3"
          onSubmit={(e) => {
            e.preventDefault()
            buscar()
          }}
        >
          <Entrada
            id="termoFicha"
            rotulo="Usuário"
            placeholder="Nome, login, matrícula ou e-mail..."
            value={termo}
            disabled={buscando}
            onChange={(e) => setTermo(e.target.value)}
            larguraClasse="min-w-64 flex-1"
          />
          <Botao tipo="submit" icone={Search} desabilitado={buscando}>
            {buscando ? 'Buscando...' : 'Buscar'}
          </Botao>
        </form>
      </Painel>

      {buscaFeita && resultados.length === 0 && !buscando && (
        <Painel titulo="Resultado">
          <EstadoVazio icone={ContactRound} titulo="Nenhuma conta encontrada" />
        </Painel>
      )}

      {resultados.map((ficha) => (
        <Painel
          key={ficha.login}
          titulo={ficha.nome}
          icone={ContactRound}
          acoes={
            <Etiqueta
              texto={ficha.situacao}
              tom={ficha.situacao === 'Ativo' ? 'sucesso' : ficha.situacao === 'Bloqueado' ? 'alerta' : 'neutro'}
            />
          }
        >
          <dl className="grid gap-4 sm:grid-cols-3">
            {[
              ['Login', ficha.login],
              ['Usuário AD', ficha.userAd],
              ['E-mail', ficha.email],
              ['Matrícula', ficha.matricula],
              ['Planta', ficha.planta],
              ['Ramal', ficha.ramal],
              ['Cadastro', formatarData(ficha.dataCadastro)],
              ['Último acesso', formatarData(ficha.ultimoAcesso)]
            ].map(([rotulo, valor]) => (
              <div key={rotulo}>
                <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
                <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
                  {valor || '—'}
                </dd>
              </div>
            ))}
          </dl>
        </Painel>
      ))}
    </div>
  )
}
