import { useMemo, useState } from 'react'
import { Check, Eye, EyeOff, KeyRound, Save, X } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import { Campo, classesControle } from '@/components/comuns/Campos'
import { ToastStack, useToast } from '@/components/Toast'
import { useAuth } from '@/context/AuthContext'
import {
  ROTULOS_FORCA,
  calcularForca,
  tamanhoMinimo,
  tipoDoUsuario,
  trocarSenha,
  validarSenha
} from '@/services/utilitarios/senhaApi'

/** Refatoração de Utilitarios/Alterar-Senha (TrocarSenha.asp). */

const CORES_FORCA = ['bg-red-500', 'bg-red-500', 'bg-amber-500', 'bg-lime-500', 'bg-emerald-500']

interface CampoSenhaProps {
  id: string
  rotulo: string
  valor: string
  aoMudar: (valor: string) => void
  placeholder: string
  desabilitado?: boolean
  erro?: string
}

function CampoSenha({ id, rotulo, valor, aoMudar, placeholder, desabilitado, erro }: CampoSenhaProps) {
  const [visivel, setVisivel] = useState(false)

  return (
    <Campo rotulo={rotulo} htmlFor={id} obrigatorio erro={erro}>
      <div className="flex">
        <input
          id={id}
          type={visivel ? 'text' : 'password'}
          autoComplete={id === 'senhaAtual' ? 'current-password' : 'new-password'}
          value={valor}
          placeholder={placeholder}
          disabled={desabilitado}
          onChange={(e) => aoMudar(e.target.value)}
          className={`${classesControle} rounded-r-none ${erro ? 'border-red-400 dark:border-red-500' : ''}`}
        />
        <button
          type="button"
          onClick={() => setVisivel((v) => !v)}
          aria-label={visivel ? 'Ocultar senha' : 'Mostrar senha'}
          className="rounded-r-lg border border-l-0 border-slate-200 px-3 text-slate-500
            hover:bg-slate-50 dark:border-slate-700 dark:hover:bg-slate-800"
        >
          {visivel ? <EyeOff size={16} /> : <Eye size={16} />}
        </button>
      </div>
    </Campo>
  )
}

export function AlterarSenha() {
  const { toasts, mostrarToast, fecharToast } = useToast()
  const { usuario, sair } = useAuth()

  const login = usuario?.login ?? ''
  const tipoUsuario = useMemo(() => tipoDoUsuario(login), [login])
  const minimo = tamanhoMinimo(tipoUsuario)

  const [senhaAtual, setSenhaAtual] = useState('')
  const [senhaNova, setSenhaNova] = useState('')
  const [confirmacao, setConfirmacao] = useState('')
  const [salvando, setSalvando] = useState(false)

  const { regras, todasAtendidas } = validarSenha(senhaNova, senhaAtual, tipoUsuario)
  const forca = calcularForca(senhaNova)

  const confirmacaoDivergente = confirmacao.length > 0 && senhaNova !== confirmacao
  const podeSalvar = senhaAtual.length > 0 && todasAtendidas && !confirmacaoDivergente && confirmacao.length > 0

  async function salvar() {
    setSalvando(true)
    try {
      const resultado = await trocarSenha(login, senhaAtual, senhaNova)
      mostrarToast(resultado.mensagem, resultado.sucesso ? 'success' : 'danger')

      if (resultado.sucesso) {
        setSenhaAtual('')
        setSenhaNova('')
        setConfirmacao('')
      }
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Alterar Senha" />

      <div className="mx-auto max-w-2xl">
        <Painel titulo="Trocar Senha" icone={KeyRound}>
          <form
            className="grid gap-4"
            onSubmit={(e) => {
              e.preventDefault()
              if (podeSalvar) salvar()
            }}
          >
            <Campo rotulo="Usuário" htmlFor="usuarioSenha">
              <input
                id="usuarioSenha"
                type="text"
                readOnly
                value={login}
                className={`${classesControle} cursor-default bg-slate-100 dark:bg-slate-800/50`}
              />
            </Campo>

            <CampoSenha
              id="senhaAtual"
              rotulo="Senha Atual"
              valor={senhaAtual}
              aoMudar={setSenhaAtual}
              placeholder="Digite sua senha atual"
              desabilitado={salvando}
            />

            <div>
              <CampoSenha
                id="senhaNova"
                rotulo="Nova Senha"
                valor={senhaNova}
                aoMudar={setSenhaNova}
                placeholder="Digite a nova senha"
                desabilitado={salvando}
              />

              {senhaNova.length > 0 && (
                <div className="mt-2 flex items-center gap-2">
                  <div className="flex flex-1 gap-1">
                    {[0, 1, 2, 3].map((segmento) => (
                      <span
                        key={segmento}
                        className={`h-1.5 flex-1 rounded-full transition-colors ${
                          segmento < forca ? CORES_FORCA[forca] : 'bg-slate-200 dark:bg-slate-700'
                        }`}
                      />
                    ))}
                  </div>
                  <span className="w-24 text-right text-xs font-medium text-slate-500 dark:text-slate-400">
                    {ROTULOS_FORCA[forca]}
                  </span>
                </div>
              )}
            </div>

            <CampoSenha
              id="confirmacao"
              rotulo="Confirmar Nova Senha"
              valor={confirmacao}
              aoMudar={setConfirmacao}
              placeholder="Repita a nova senha"
              desabilitado={salvando}
              erro={confirmacaoDivergente ? 'As senhas não conferem.' : undefined}
            />

            <div className="rounded-lg border border-slate-200 p-3 dark:border-slate-700">
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                Regras de senha
              </p>
              <ul className="grid gap-1.5 sm:grid-cols-2">
                {regras.map((regra) => {
                  const neutra = senhaNova.length === 0
                  const Icone = regra.atende ? Check : X
                  return (
                    <li
                      key={regra.id}
                      className={`flex items-start gap-1.5 text-xs ${
                        neutra
                          ? 'text-slate-500 dark:text-slate-400'
                          : regra.atende
                            ? 'text-emerald-600 dark:text-emerald-400'
                            : 'text-red-600 dark:text-red-400'
                      }`}
                    >
                      <Icone size={13} className="mt-0.5 shrink-0" />
                      <span>{regra.texto}</span>
                    </li>
                  )
                })}
              </ul>
              {tipoUsuario === '5' && (
                <p className="mt-2 text-xs text-amber-600 dark:text-amber-400">
                  Seu perfil exige senha de {minimo} caracteres.
                </p>
              )}
            </div>

            <div className="flex justify-end gap-2">
              <Botao variante="neutro" aoClicar={sair} desabilitado={salvando}>
                Sair
              </Botao>
              <Botao tipo="submit" variante="sucesso" icone={Save} desabilitado={!podeSalvar || salvando}>
                {salvando ? 'Salvando...' : 'Salvar nova senha'}
              </Botao>
            </div>
          </form>
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
