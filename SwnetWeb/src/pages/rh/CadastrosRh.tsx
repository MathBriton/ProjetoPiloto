import { useCallback, useEffect, useState } from 'react'
import { Building2, ListTree, Pencil, Plus, Save, Sliders, Trash2, Users } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, Etiqueta, TituloPagina } from '@/components/comuns/Comuns'
import { BotaoAcao, Tabela } from '@/components/comuns/Tabela'
import { Entrada, Selecao, SelecaoMultipla } from '@/components/comuns/Campos'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { Modal } from '@/components/comuns/Modal'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import {
  NIVEIS_PRIORIDADE,
  descricaoNivel,
  excluirAtendente,
  excluirCategoria,
  excluirDepartamento,
  excluirPrioridade,
  listarAtendentes,
  listarCategorias,
  listarDepartamentos,
  listarPrioridades,
  salvarAtendente,
  salvarCategoria,
  salvarDepartamento,
  salvarPrioridade,
  type AtendenteRh,
  type CategoriaChamado,
  type DepartamentoChamado,
  type PrioridadeChamado
} from '@/services/rh/rhApi'
import { PLANTAS_CREDITO, nomePlanta } from '@/services/financeiro/financeiroApi'
import { formatarData, usuariosSwnet } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from '@/pages/administrador/tipos'

/**
 * Refatoração dos quatro cadastros de apoio do RH:
 * Cadastro-Atendentes-RH, Categorias-De-Chamado,
 * Departamentos-De-Chamado-RH e Prioridades-De-Chamado-RH.
 *
 * No legado são quatro pastas com o mesmo trio Grade/Formulário/Repositório
 * mudando só os campos, por isso as quatro telas vivem neste arquivo.
 */

function SeletorPlanta({
  valor,
  aoMudar,
  id
}: {
  valor: string
  aoMudar: (valor: string) => void
  id: string
}) {
  return (
    <Selecao id={id} rotulo="Planta" required value={valor} onChange={(e) => aoMudar(e.target.value)}>
      <option value="">Selecione a planta</option>
      {PLANTAS_CREDITO.map((planta) => (
        <option key={planta.codigo} value={planta.codigo}>
          {planta.nome}
        </option>
      ))}
    </Selecao>
  )
}

/* ------------------------------------------------------------------ *
 * Departamentos
 * ------------------------------------------------------------------ */

export function DepartamentosChamado() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [itens, setItens] = useState<DepartamentoChamado[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{ id?: number; planta: string; nome: string } | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<DepartamentoChamado | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setItens(await listarDepartamentos())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function salvar() {
    if (!edicao) return
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarDepartamento(edicao)
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirDepartamento(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Departamentos de Chamado" />

      <Painel
        titulo="Departamentos"
        icone={Building2}
        acoes={
          <Botao
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ planta: '', nome: '' })
            }}
          >
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={itens}
          chave={(item) => item.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (d) => d.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (d) => nomePlanta(d.planta) },
            { cabecalho: 'Nome', conteudo: (d) => d.nome },
            { cabecalho: 'Data de Criação', conteudo: (d) => formatarData(d.dataCriacao) }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar departamento"
                aoClicar={() => {
                  setErro('')
                  setEdicao({ id: item.id, planta: item.planta, nome: item.nome })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir departamento"
                tom="perigo"
                aoClicar={() => setExcluindo(item)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Departamento' : 'Novo Departamento'}
        icone={Building2}
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}
            <SeletorPlanta
              id="plantaDepto"
              valor={edicao.planta}
              aoMudar={(v) => setEdicao({ ...edicao, planta: v })}
            />
            <Entrada
              id="nomeDepto"
              rotulo="Nome"
              required
              value={edicao.nome}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, nome: e.target.value })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir departamento"
        mensagem={excluindo ? `Deseja realmente excluir "${excluindo.nome}"?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

/* ------------------------------------------------------------------ *
 * Categorias
 * ------------------------------------------------------------------ */

export function CategoriasChamado() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [itens, setItens] = useState<CategoriaChamado[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{ id?: number; planta: string; nome: string; tipo: string } | null>(
    null
  )
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<CategoriaChamado | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setItens(await listarCategorias())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function salvar() {
    if (!edicao) return
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarCategoria(edicao)
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirCategoria(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Categorias de Chamado" />

      <Painel
        titulo="Categorias"
        icone={ListTree}
        acoes={
          <Botao
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ planta: '', nome: '', tipo: '' })
            }}
          >
            Nova
          </Botao>
        }
      >
        <Tabela
          itens={itens}
          chave={(item) => item.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (c) => c.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (c) => nomePlanta(c.planta) },
            { cabecalho: 'Nome', conteudo: (c) => c.nome },
            { cabecalho: 'Tipo', conteudo: (c) => c.tipo || '—' },
            { cabecalho: 'Data de Criação', conteudo: (c) => formatarData(c.dataCriacao) }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar categoria"
                aoClicar={() => {
                  setErro('')
                  setEdicao({ id: item.id, planta: item.planta, nome: item.nome, tipo: item.tipo })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir categoria"
                tom="perigo"
                aoClicar={() => setExcluindo(item)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Categoria' : 'Nova Categoria'}
        icone={ListTree}
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}
            <SeletorPlanta
              id="plantaCat"
              valor={edicao.planta}
              aoMudar={(v) => setEdicao({ ...edicao, planta: v })}
            />
            <Entrada
              id="nomeCat"
              rotulo="Nome"
              required
              value={edicao.nome}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, nome: e.target.value })}
            />
            <Entrada
              id="tipoCat"
              rotulo="Tipo"
              value={edicao.tipo}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, tipo: e.target.value })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir categoria"
        mensagem={excluindo ? `Deseja realmente excluir "${excluindo.nome}"?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

/* ------------------------------------------------------------------ *
 * Prioridades
 * ------------------------------------------------------------------ */

export function PrioridadesChamado() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [itens, setItens] = useState<PrioridadeChamado[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{
    id?: number
    planta: string
    nivel: string
    prazoHoras: string
  } | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<PrioridadeChamado | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      setItens(await listarPrioridades())
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  async function salvar() {
    if (!edicao) return
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarPrioridade(edicao)
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirPrioridade(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Prioridades de Chamado" />

      <Painel
        titulo="Prioridades"
        icone={Sliders}
        acoes={
          <Botao
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setEdicao({ planta: '', nivel: '', prazoHoras: '' })
            }}
          >
            Nova
          </Botao>
        }
      >
        <Tabela
          itens={itens}
          chave={(item) => item.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (p) => p.id, className: 'w-20' },
            { cabecalho: 'Planta', conteudo: (p) => nomePlanta(p.planta) },
            {
              cabecalho: 'Nível',
              conteudo: (p) => (
                <Etiqueta
                  texto={descricaoNivel(p.nivel)}
                  tom={
                    p.nivel === 'URGENTE'
                      ? 'perigo'
                      : p.nivel === 'ALTA'
                        ? 'alerta'
                        : p.nivel === 'Media'
                          ? 'info'
                          : 'neutro'
                  }
                />
              )
            },
            {
              cabecalho: 'Prazo SLA',
              conteudo: (p) => `${p.prazoHoras} h`,
              alinhamento: 'direita'
            },
            { cabecalho: 'Data de Criação', conteudo: (p) => formatarData(p.dataCriacao) }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar prioridade"
                aoClicar={() => {
                  setErro('')
                  setEdicao({
                    id: item.id,
                    planta: item.planta,
                    nivel: item.nivel,
                    prazoHoras: String(item.prazoHoras)
                  })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir prioridade"
                tom="perigo"
                aoClicar={() => setExcluindo(item)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Prioridade' : 'Nova Prioridade'}
        icone={Sliders}
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}
            <SeletorPlanta
              id="plantaPrio"
              valor={edicao.planta}
              aoMudar={(v) => setEdicao({ ...edicao, planta: v })}
            />
            <Selecao
              id="nivelPrio"
              rotulo="Nível"
              required
              value={edicao.nivel}
              disabled={salvando}
              onChange={(e) => setEdicao({ ...edicao, nivel: e.target.value })}
            >
              <option value="">Selecione o nível</option>
              {NIVEIS_PRIORIDADE.map((item) => (
                <option key={item.codigo} value={item.codigo}>
                  {item.descricao}
                </option>
              ))}
            </Selecao>
            <Entrada
              id="prazoHoras"
              rotulo="Prazo em horas"
              type="number"
              min="1"
              required
              value={edicao.prazoHoras}
              disabled={salvando}
              ajuda="Define a data prevista de atendimento"
              onChange={(e) => setEdicao({ ...edicao, prazoHoras: e.target.value })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir prioridade"
        mensagem={excluindo ? `Deseja realmente excluir a prioridade ${descricaoNivel(excluindo.nivel)}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}

/* ------------------------------------------------------------------ *
 * Atendentes
 * ------------------------------------------------------------------ */

export function AtendentesRh() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [itens, setItens] = useState<AtendenteRh[]>([])
  const [departamentos, setDepartamentos] = useState<DepartamentoChamado[]>([])
  const [carregando, setCarregando] = useState(true)
  const [edicao, setEdicao] = useState<{
    id?: number
    planta: string
    idUsuario: number | ''
    departamentos: number[]
  } | null>(null)
  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [salvando, setSalvando] = useState(false)
  const [erro, setErro] = useState('')
  const [excluindo, setExcluindo] = useState<AtendenteRh | null>(null)

  const carregar = useCallback(async () => {
    setCarregando(true)
    try {
      const [lista, listaDepartamentos] = await Promise.all([listarAtendentes(), listarDepartamentos()])
      setItens(lista)
      setDepartamentos(listaDepartamentos)
    } finally {
      setCarregando(false)
    }
  }, [])

  useEffect(() => {
    carregar()
  }, [carregar])

  function nomesDepartamentos(ids: number[]) {
    return ids.map((id) => departamentos.find((d) => d.id === id)?.nome ?? id).join(', ') || '—'
  }

  async function salvar() {
    if (!edicao) return
    setSalvando(true)
    setErro('')
    try {
      const resultado = await salvarAtendente({ ...edicao, idUsuario: usuario?.id ?? '' })
      if (!resultado.sucesso) {
        setErro(resultado.mensagem)
        return
      }
      setEdicao(null)
      mostrarToast(resultado.mensagem, 'success')
      await carregar()
    } finally {
      setSalvando(false)
    }
  }

  async function confirmarExclusao() {
    if (!excluindo) return
    const resultado = await excluirAtendente(excluindo.id)
    setExcluindo(null)
    mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')
    if (resultado.sucesso) await carregar()
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Atendentes de RH" />

      <Painel
        titulo="Atendentes"
        icone={Users}
        acoes={
          <Botao
            icone={Plus}
            aoClicar={() => {
              setErro('')
              setUsuario(null)
              setEdicao({ planta: '', idUsuario: '', departamentos: [] })
            }}
          >
            Novo
          </Botao>
        }
      >
        <Tabela
          itens={itens}
          chave={(item) => item.id}
          carregando={carregando}
          colunas={[
            { cabecalho: '#', conteudo: (a) => a.id, className: 'w-24' },
            { cabecalho: 'Planta', conteudo: (a) => nomePlanta(a.planta) },
            { cabecalho: 'Nome', conteudo: (a) => a.nome },
            { cabecalho: 'Usuário', conteudo: (a) => a.usuario },
            { cabecalho: 'E-mail', conteudo: (a) => a.email },
            {
              cabecalho: 'Departamentos',
              conteudo: (a) => <span className="text-xs">{nomesDepartamentos(a.departamentos)}</span>
            },
            { cabecalho: 'Data de Criação', conteudo: (a) => formatarData(a.dataCriacao) }
          ]}
          acoes={(item) => (
            <>
              <BotaoAcao
                icone={Pencil}
                titulo="Editar atendente"
                aoClicar={() => {
                  setErro('')
                  setUsuario(usuariosSwnet.find((u) => u.id === item.id) ?? null)
                  setEdicao({
                    id: item.id,
                    planta: item.planta,
                    idUsuario: item.id,
                    departamentos: [...item.departamentos]
                  })
                }}
              />
              <BotaoAcao
                icone={Trash2}
                titulo="Excluir atendente"
                tom="perigo"
                aoClicar={() => setExcluindo(item)}
              />
            </>
          )}
        />
      </Painel>

      <Modal
        aberto={edicao !== null}
        titulo={edicao?.id ? 'Editar Atendente' : 'Novo Atendente'}
        icone={Users}
        largura="lg"
        aoFechar={() => setEdicao(null)}
        rodape={
          <>
            <Botao variante="neutro" aoClicar={() => setEdicao(null)} desabilitado={salvando}>
              Voltar
            </Botao>
            <Botao variante="sucesso" icone={Save} aoClicar={salvar} desabilitado={salvando}>
              {salvando ? 'Salvando...' : 'Salvar'}
            </Botao>
          </>
        }
      >
        {edicao && (
          <div className="grid gap-4">
            {erro && (
              <p className="rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-500/10 dark:text-red-400">
                {erro}
              </p>
            )}

            <SeletorUsuario rotulo="Usuário" selecionado={usuario} aoSelecionar={setUsuario} apenasAtivos />

            <SeletorPlanta
              id="plantaAtendente"
              valor={edicao.planta}
              aoMudar={(v) => setEdicao({ ...edicao, planta: v })}
            />

            <SelecaoMultipla
              rotulo="Departamentos"
              obrigatorio
              colunas={2}
              opcoes={departamentos.map((d) => ({
                valor: String(d.id),
                rotulo: `${d.nome} (${nomePlanta(d.planta)})`
              }))}
              selecionados={edicao.departamentos.map(String)}
              desabilitado={salvando}
              aoMudar={(ids) => setEdicao({ ...edicao, departamentos: ids.map(Number) })}
            />
          </div>
        )}
      </Modal>

      <ConfirmDialog
        aberto={excluindo !== null}
        titulo="Excluir atendente"
        mensagem={excluindo ? `Deseja realmente excluir o atendente ${excluindo.nome}?` : ''}
        rotuloConfirmar="Excluir"
        aoConfirmar={confirmarExclusao}
        aoCancelar={() => setExcluindo(null)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
