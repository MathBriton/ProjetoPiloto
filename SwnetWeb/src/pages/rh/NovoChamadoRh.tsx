import { useEffect, useState } from 'react'
import { Headset, Paperclip, Send } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, Entrada, Selecao, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  NOVO_CHAMADO_VAZIO,
  TIPOS_CHAMADO,
  criarChamado,
  listarCategorias,
  listarDepartamentos,
  type CategoriaChamado,
  type DadosNovoChamado,
  type DepartamentoChamado
} from '@/services/rh/rhApi'
import { PLANTAS_CREDITO } from '@/services/financeiro/financeiroApi'

/** Refatoração de RH/Novo-Chamado-RH (ChamadoFormulario.asp). */

export function NovoChamadoRh() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario } = useAuth()

  const [dados, setDados] = useState<DadosNovoChamado>({
    ...NOVO_CHAMADO_VAZIO,
    nomeCompleto: usuario?.nomeCompleto ?? ''
  })
  const [departamentos, setDepartamentos] = useState<DepartamentoChamado[]>([])
  const [categorias, setCategorias] = useState<CategoriaChamado[]>([])
  const [enviando, setEnviando] = useState(false)
  const [novoAnexo, setNovoAnexo] = useState('')

  // As listas de apoio são filtradas pela planta escolhida
  useEffect(() => {
    listarDepartamentos(dados.planta).then(setDepartamentos)
    listarCategorias(dados.planta).then(setCategorias)
  }, [dados.planta])

  function mudar<C extends keyof DadosNovoChamado>(campo: C, valor: DadosNovoChamado[C]) {
    setDados((atual) => ({ ...atual, [campo]: valor }))
  }

  function adicionarAnexo() {
    if (!novoAnexo.trim()) return
    mudar('anexos', [...dados.anexos, novoAnexo.trim()])
    setNovoAnexo('')
  }

  async function enviar() {
    setEnviando(true)
    try {
      const resultado = await criarChamado(dados, usuario?.login ?? '')
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setDados({ ...NOVO_CHAMADO_VAZIO, nomeCompleto: usuario?.nomeCompleto ?? '' })
      }
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Novo Chamado RH" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Abrir Chamado" icone={Headset}>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              enviar()
            }}
          >
            <Secao titulo="Solicitante">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="plantaChamado"
                  rotulo="Planta"
                  required
                  value={dados.planta}
                  disabled={enviando}
                  onChange={(e) => {
                    mudar('planta', e.target.value)
                    mudar('idDepartamento', '')
                    mudar('idCategoria', '')
                  }}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {PLANTAS_CREDITO.map((planta) => (
                    <option key={planta.codigo} value={planta.codigo}>
                      {planta.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="nomeCompleto"
                  rotulo="Nome Completo"
                  required
                  value={dados.nomeCompleto}
                  disabled={enviando}
                  onChange={(e) => mudar('nomeCompleto', e.target.value)}
                  larguraClasse="sm:col-span-8"
                />

                <Entrada
                  id="emailChamado"
                  rotulo="E-mail"
                  type="email"
                  required
                  value={dados.email}
                  disabled={enviando}
                  onChange={(e) => mudar('email', e.target.value)}
                  larguraClasse="sm:col-span-5"
                />
                <Entrada
                  id="telefoneChamado"
                  rotulo="Telefone"
                  value={dados.telefone}
                  disabled={enviando}
                  onChange={(e) => mudar('telefone', e.target.value)}
                  larguraClasse="sm:col-span-3"
                />
                <Entrada
                  id="centroCusto"
                  rotulo="Centro de Custo"
                  value={dados.centroCusto}
                  disabled={enviando}
                  ajuda="Opcional"
                  onChange={(e) => mudar('centroCusto', e.target.value)}
                  larguraClasse="sm:col-span-4"
                />
              </div>
            </Secao>

            <Secao titulo="Chamado">
              <div className="grid gap-4 sm:grid-cols-12">
                <Selecao
                  id="tipoChamado"
                  rotulo="Tipo de Chamado"
                  required
                  value={dados.tipo}
                  disabled={enviando}
                  onChange={(e) => mudar('tipo', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {TIPOS_CHAMADO.map((tipo) => (
                    <option key={tipo} value={tipo}>
                      {tipo}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="departamentoChamado"
                  rotulo="Departamento"
                  required
                  value={dados.idDepartamento}
                  disabled={enviando || !dados.planta}
                  onChange={(e) => mudar('idDepartamento', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {departamentos.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Selecao
                  id="categoriaChamado"
                  rotulo="Categoria"
                  value={dados.idCategoria}
                  disabled={enviando || !dados.planta}
                  ajuda="Opcional"
                  onChange={(e) => mudar('idCategoria', e.target.value)}
                  larguraClasse="sm:col-span-4"
                >
                  <option value="">Selecione...</option>
                  {categorias.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.nome}
                    </option>
                  ))}
                </Selecao>

                <Entrada
                  id="assuntoChamado"
                  rotulo="Assunto"
                  required
                  value={dados.assunto}
                  disabled={enviando}
                  onChange={(e) => mudar('assunto', e.target.value)}
                  larguraClasse="sm:col-span-12"
                />

                <Campo rotulo="Descrição" htmlFor="descricaoChamado" obrigatorio className="sm:col-span-12">
                  <textarea
                    id="descricaoChamado"
                    rows={5}
                    value={dados.descricao}
                    disabled={enviando}
                    placeholder="Descreva sua solicitação com o máximo de detalhes..."
                    onChange={(e) => mudar('descricao', e.target.value)}
                    className={classesControle}
                  />
                </Campo>
              </div>
            </Secao>

            <Secao titulo="Anexos" icone={Paperclip}>
              <div className="mb-3 flex items-end gap-2">
                <Entrada
                  id="novoAnexo"
                  rotulo="Nome do arquivo"
                  placeholder="ex.: holerite.pdf"
                  value={novoAnexo}
                  disabled={enviando}
                  onChange={(e) => setNovoAnexo(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault()
                      adicionarAnexo()
                    }
                  }}
                  larguraClasse="flex-1"
                />
                <Botao variante="neutro" aoClicar={adicionarAnexo} desabilitado={enviando}>
                  Adicionar
                </Botao>
              </div>

              {dados.anexos.length === 0 ? (
                <p className="rounded-lg border border-dashed border-slate-200 px-3 py-4 text-center text-xs text-slate-400 dark:border-slate-700">
                  Nenhum anexo. O upload real depende do backend.
                </p>
              ) : (
                <ul className="grid gap-1.5">
                  {dados.anexos.map((anexo, indice) => (
                    <li
                      key={indice}
                      className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 text-sm
                        text-slate-600 dark:border-slate-700 dark:text-slate-300"
                    >
                      <Paperclip size={14} className="shrink-0 text-slate-400" />
                      <span className="flex-1 truncate">{anexo}</span>
                      <button
                        type="button"
                        onClick={() => mudar('anexos', dados.anexos.filter((_, i) => i !== indice))}
                        disabled={enviando}
                        className="text-xs text-red-600 hover:underline dark:text-red-400"
                      >
                        remover
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </Secao>

            <div className="mt-4 flex justify-end">
              <Botao tipo="submit" variante="sucesso" icone={Send} desabilitado={enviando}>
                {enviando ? 'Abrindo...' : 'Abrir chamado'}
              </Botao>
            </div>
          </form>
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
