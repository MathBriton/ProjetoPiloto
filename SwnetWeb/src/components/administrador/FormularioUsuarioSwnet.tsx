import { useEffect, useState } from 'react'
import { Building2, Eye, EyeOff, IdCard, KeyRound } from 'lucide-react'
import { Secao } from '@/components/comuns/Painel'
import { Chave, Entrada, OpcoesRadio, Selecao, classesControle } from '@/components/comuns/Campos'
import { EMPRESAS, TIPOS_USUARIO, codigoEmpresa } from '@/services/comuns/baseMock'
import type { PrefixoAd, Sexo, UsuarioSwnet } from '@/pages/administrador/tipos'

/**
 * Formulário compartilhado por CadUsuario.asp (modo "criar") e AltUsuario.asp
 * (modo "alterar"). As diferenças entre as duas telas legadas:
 *
 * - criar:   senha obrigatória, opção de enviar e-mail de boas-vindas
 * - alterar: senha opcional (em branco = não altera) e switch de bloqueio
 */

export interface ValoresFormularioUsuario {
  matricula: string
  nome: string
  sexo: Sexo
  nascimento: string
  email: string
  representante: boolean
  prefixoAd: PrefixoAd
  login: string
  senha: string
  confirmaSenha: string
  tipo: string
  empresa: string
  ramal: string
  enviarEmail: boolean
  bloqueado: boolean
}

export const VALORES_INICIAIS: ValoresFormularioUsuario = {
  matricula: '',
  nome: '',
  sexo: 'M',
  nascimento: '',
  email: '',
  representante: false,
  prefixoAd: 'GROUP',
  login: '',
  senha: '',
  confirmaSenha: '',
  tipo: '',
  empresa: '',
  ramal: '',
  enviarEmail: true,
  bloqueado: false
}

export function valoresDoUsuario(usuario: UsuarioSwnet): ValoresFormularioUsuario {
  const [prefixo] = usuario.userAd.split('\\')
  return {
    matricula: usuario.matricula,
    nome: usuario.nome,
    sexo: usuario.sexo,
    nascimento: usuario.nascimento,
    email: usuario.email,
    representante: usuario.representante,
    prefixoAd: prefixo === 'WESTROCK' ? 'WESTROCK' : 'GROUP',
    login: usuario.login,
    senha: '',
    confirmaSenha: '',
    tipo: usuario.tipo,
    empresa: usuario.empresa,
    ramal: usuario.ramal,
    enviarEmail: false,
    bloqueado: usuario.bloqueado
  }
}

export type ErrosFormularioUsuario = Partial<Record<keyof ValoresFormularioUsuario, string>>

/** Mesmas regras do JavaScript inline das duas telas ASP. */
export function validarFormulario(
  valores: ValoresFormularioUsuario,
  modo: 'criar' | 'alterar'
): ErrosFormularioUsuario {
  const erros: ErrosFormularioUsuario = {}

  if (!valores.matricula.trim()) erros.matricula = 'Informe a matrícula.'
  if (!valores.nome.trim()) erros.nome = 'Informe o nome completo.'
  if (!valores.email.trim()) erros.email = 'Informe o e-mail.'
  else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(valores.email.trim())) erros.email = 'E-mail inválido.'
  if (!valores.login.trim()) erros.login = 'Informe o login do usuário.'
  if (!valores.tipo) erros.tipo = 'Selecione o tipo.'
  if (!valores.empresa) erros.empresa = 'Selecione a empresa.'

  const senhaInformada = valores.senha.length > 0
  if (modo === 'criar' || senhaInformada) {
    if (valores.senha.length < 7) erros.senha = 'Mínimo de 7 caracteres.'
    if (valores.senha !== valores.confirmaSenha) erros.confirmaSenha = 'As senhas não conferem.'
  }

  return erros
}

interface FormularioUsuarioSwnetProps {
  modo: 'criar' | 'alterar'
  valores: ValoresFormularioUsuario
  erros: ErrosFormularioUsuario
  aoMudar: <C extends keyof ValoresFormularioUsuario>(campo: C, valor: ValoresFormularioUsuario[C]) => void
  desabilitado?: boolean
}

export function FormularioUsuarioSwnet({
  modo,
  valores,
  erros,
  aoMudar,
  desabilitado = false
}: FormularioUsuarioSwnetProps) {
  const [mostrarSenha, setMostrarSenha] = useState(false)
  const [mostrarConfirma, setMostrarConfirma] = useState(false)

  // O legado marca CRM_TIPO = 'R' junto com o flag de representante
  useEffect(() => {
    if (valores.representante && valores.tipo === '') aoMudar('tipo', 'R')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [valores.representante])

  const criando = modo === 'criar'

  return (
    <>
      <Secao titulo="Dados Pessoais" icone={IdCard}>
        <div className="grid gap-4 sm:grid-cols-12">
          <Entrada
            id="matricula"
            rotulo="Matrícula"
            required
            maxLength={6}
            placeholder="000000"
            value={valores.matricula}
            erro={erros.matricula}
            disabled={desabilitado}
            ajuda={criando ? '6 dígitos (exceto filiais Bento e Maranguape)' : undefined}
            onChange={(e) => aoMudar('matricula', e.target.value.replace(/\D/g, ''))}
            larguraClasse="sm:col-span-3"
          />
          <OpcoesRadio
            rotulo="Sexo"
            nome="sexo"
            valor={valores.sexo}
            desabilitado={desabilitado}
            opcoes={[
              { valor: 'M', rotulo: 'Masculino' },
              { valor: 'F', rotulo: 'Feminino' }
            ]}
            aoMudar={(v) => aoMudar('sexo', v as Sexo)}
            larguraClasse="sm:col-span-5"
          />
          <Entrada
            id="nascimento"
            rotulo="Nascimento"
            type="date"
            value={valores.nascimento}
            disabled={desabilitado}
            onChange={(e) => aoMudar('nascimento', e.target.value)}
            larguraClasse="sm:col-span-4"
          />

          <Entrada
            id="nome"
            rotulo="Nome"
            required
            placeholder="Nome completo"
            value={valores.nome}
            erro={erros.nome}
            disabled={desabilitado}
            onChange={(e) => aoMudar('nome', e.target.value)}
            larguraClasse="sm:col-span-12"
          />

          <Entrada
            id="email"
            rotulo="E-mail"
            type="email"
            required
            placeholder="exemplo@smurfitwestrock.com"
            value={valores.email}
            erro={erros.email}
            disabled={desabilitado}
            onChange={(e) => aoMudar('email', e.target.value)}
            larguraClasse="sm:col-span-8"
          />
          <OpcoesRadio
            rotulo="Representante"
            nome="representante"
            valor={valores.representante ? 'S' : 'N'}
            desabilitado={desabilitado}
            opcoes={[
              { valor: 'S', rotulo: 'Sim' },
              { valor: 'N', rotulo: 'Não' }
            ]}
            aoMudar={(v) => aoMudar('representante', v === 'S')}
            larguraClasse="sm:col-span-4"
          />
        </div>
      </Secao>

      <Secao titulo="Dados de Acesso" icone={KeyRound}>
        <div className="grid gap-4 sm:grid-cols-12">
          <Selecao
            id="prefixoAd"
            rotulo="Domínio AD"
            value={valores.prefixoAd}
            disabled={desabilitado}
            onChange={(e) => aoMudar('prefixoAd', e.target.value as PrefixoAd)}
            larguraClasse="sm:col-span-3"
          >
            <option value="GROUP">GROUP</option>
            <option value="WESTROCK">WESTROCK</option>
          </Selecao>

          <div className="sm:col-span-9">
            <label htmlFor="login" className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300">
              Usuário<span className="ml-0.5 text-red-500">*</span>
            </label>
            <div className="flex">
              <span
                className="inline-flex items-center rounded-l-lg border border-r-0 border-slate-200 bg-slate-100
                  px-3 text-sm font-semibold text-slate-500 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-400"
              >
                {valores.prefixoAd}\
              </span>
              <input
                id="login"
                type="text"
                maxLength={30}
                autoComplete="off"
                placeholder="Login do usuário..."
                value={valores.login}
                disabled={desabilitado}
                onChange={(e) => aoMudar('login', e.target.value.trim())}
                className={`${classesControle} rounded-l-none ${erros.login ? 'border-red-400 dark:border-red-500' : ''}`}
              />
            </div>
            {erros.login && <p className="mt-1 text-xs font-medium text-red-600 dark:text-red-400">{erros.login}</p>}
          </div>

          <div className="sm:col-span-6">
            <label htmlFor="senha" className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300">
              Senha{criando && <span className="ml-0.5 text-red-500">*</span>}
            </label>
            <div className="flex">
              <input
                id="senha"
                type={mostrarSenha ? 'text' : 'password'}
                autoComplete="new-password"
                placeholder={criando ? 'Nova senha' : 'Deixe em branco para não alterar'}
                value={valores.senha}
                disabled={desabilitado}
                onChange={(e) => aoMudar('senha', e.target.value)}
                className={`${classesControle} rounded-r-none ${erros.senha ? 'border-red-400 dark:border-red-500' : ''}`}
              />
              <button
                type="button"
                onClick={() => setMostrarSenha((v) => !v)}
                aria-label={mostrarSenha ? 'Ocultar senha' : 'Mostrar senha'}
                className="rounded-r-lg border border-l-0 border-slate-200 px-3 text-slate-500
                  hover:bg-slate-50 dark:border-slate-700 dark:hover:bg-slate-800"
              >
                {mostrarSenha ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
            {erros.senha ? (
              <p className="mt-1 text-xs font-medium text-red-600 dark:text-red-400">{erros.senha}</p>
            ) : (
              <p className="mt-1 text-xs text-slate-400">Mínimo de 7 caracteres.</p>
            )}
          </div>

          <div className="sm:col-span-6">
            <label
              htmlFor="confirmaSenha"
              className="mb-1 block text-xs font-semibold text-slate-600 dark:text-slate-300"
            >
              Confirmar Senha{criando && <span className="ml-0.5 text-red-500">*</span>}
            </label>
            <div className="flex">
              <input
                id="confirmaSenha"
                type={mostrarConfirma ? 'text' : 'password'}
                autoComplete="new-password"
                placeholder="Confirme a senha"
                value={valores.confirmaSenha}
                disabled={desabilitado}
                onChange={(e) => aoMudar('confirmaSenha', e.target.value)}
                className={`${classesControle} rounded-r-none ${
                  erros.confirmaSenha ? 'border-red-400 dark:border-red-500' : ''
                }`}
              />
              <button
                type="button"
                onClick={() => setMostrarConfirma((v) => !v)}
                aria-label={mostrarConfirma ? 'Ocultar senha' : 'Mostrar senha'}
                className="rounded-r-lg border border-l-0 border-slate-200 px-3 text-slate-500
                  hover:bg-slate-50 dark:border-slate-700 dark:hover:bg-slate-800"
              >
                {mostrarConfirma ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
            {erros.confirmaSenha ? (
              <p className="mt-1 text-xs font-medium text-red-600 dark:text-red-400">{erros.confirmaSenha}</p>
            ) : (
              valores.confirmaSenha.length > 0 &&
              valores.senha === valores.confirmaSenha && (
                <p className="mt-1 text-xs font-medium text-emerald-600 dark:text-emerald-400">Senhas conferem.</p>
              )
            )}
          </div>
        </div>
      </Secao>

      <Secao titulo="Vínculo" icone={Building2}>
        <div className="grid gap-4 sm:grid-cols-12">
          <Selecao
            id="tipo"
            rotulo="Tipo"
            required
            value={valores.tipo}
            erro={erros.tipo}
            disabled={desabilitado}
            onChange={(e) => aoMudar('tipo', e.target.value)}
            larguraClasse="sm:col-span-4"
          >
            <option value="">Selecione...</option>
            {TIPOS_USUARIO.map((tipo) => (
              <option key={tipo.codigo} value={tipo.codigo}>
                {tipo.nome}
              </option>
            ))}
          </Selecao>

          <Selecao
            id="empresa"
            rotulo="Empresa"
            required
            value={valores.empresa}
            erro={erros.empresa}
            disabled={desabilitado}
            onChange={(e) => aoMudar('empresa', e.target.value)}
            larguraClasse="sm:col-span-5"
          >
            <option value="">Selecione...</option>
            {EMPRESAS.map((empresa) => (
              <option key={empresa.id} value={codigoEmpresa(empresa.id)}>
                {codigoEmpresa(empresa.id)} — {empresa.nome}
              </option>
            ))}
          </Selecao>

          <Entrada
            id="ramal"
            rotulo="Ramal"
            maxLength={6}
            placeholder="Ramal"
            value={valores.ramal}
            disabled={desabilitado}
            onChange={(e) => aoMudar('ramal', e.target.value.replace(/\D/g, ''))}
            larguraClasse="sm:col-span-3"
          />

          <div className="sm:col-span-12">
            {criando ? (
              <Chave
                rotulo="Enviar e-mail de boas-vindas"
                descricao="Envia login e senha provisória para o e-mail informado."
                marcado={valores.enviarEmail}
                desabilitado={desabilitado}
                aoAlternar={(v) => aoMudar('enviarEmail', v)}
              />
            ) : (
              <Chave
                rotulo="Usuário bloqueado"
                descricao="Bloqueia o acesso sem desativar o cadastro."
                marcado={valores.bloqueado}
                desabilitado={desabilitado}
                aoAlternar={(v) => aoMudar('bloqueado', v)}
              />
            )}
          </div>
        </div>
      </Secao>
    </>
  )
}
