import { useState } from 'react'
import { FileDown, Power, PowerOff, UserSearch } from 'lucide-react'
import { Painel, Secao } from '@/components/comuns/Painel'
import { Botao, Etiqueta, EstadoVazio, TituloPagina } from '@/components/comuns/Comuns'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { alternarSituacao } from '@/services/administrador/usuariosSwnetApi'
import { formatarData, nomeEmpresa, nomeTipoUsuario } from '@/services/comuns/baseMock'
import type { UsuarioSwnet } from './tipos'

/**
 * Refatoração de Ativar-Desativar-Usuarios/ExcUsuario.asp.
 * A tela é somente leitura: mostra a ficha do usuário e alterna a situação
 * (CRM_D_E_L_E_T) mediante confirmação.
 */

function Dado({ rotulo, valor }: { rotulo: string; valor: string }) {
  return (
    <div>
      <dt className="text-xs font-semibold text-slate-500 dark:text-slate-400">{rotulo}</dt>
      <dd className="mt-0.5 truncate text-sm text-slate-700 dark:text-slate-200" title={valor}>
        {valor || '—'}
      </dd>
    </div>
  )
}

export function AtivarDesativarUsuarios() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [confirmando, setConfirmando] = useState(false)
  const [processando, setProcessando] = useState(false)

  async function confirmar() {
    if (!usuario) return

    setConfirmando(false)
    setProcessando(true)
    try {
      const resultado = await alternarSituacao(usuario.id)
      if (resultado.sucesso && resultado.usuario) {
        setUsuario(resultado.usuario)
        mostrarToast(resultado.mensagem, 'success')
      } else {
        mostrarToast(resultado.mensagem, 'danger')
      }
    } finally {
      setProcessando(false)
    }
  }

  function exportarPdf() {
    // TODO: o legado usa jsPDF; ainda não há lib de PDF no frontend
    mostrarToast('Exportação em PDF será implementada em breve.', 'info')
  }

  const desativando = usuario?.ativo ?? true

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Ativar / Desativar Usuários" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Situação do Usuário" icone={Power}>
          <div className="mb-4">
            <SeletorUsuario selecionado={usuario} aoSelecionar={setUsuario} />
          </div>

          {!usuario && (
            <EstadoVazio
              icone={UserSearch}
              titulo="Selecione um usuário"
              descricao="Busque pelo nome ou login para ver a ficha e ativar ou desativar o acesso."
            />
          )}

          {usuario && (
            <>
              <div className="mb-4 flex flex-wrap items-center gap-2">
                <Etiqueta texto={usuario.ativo ? 'Ativo' : 'Inativo'} tom={usuario.ativo ? 'sucesso' : 'neutro'} />
                {usuario.bloqueado && <Etiqueta texto="Bloqueado" tom="alerta" />}
                {usuario.dataBloqueio && (
                  <span className="text-xs text-slate-400">
                    Bloqueado em {formatarData(usuario.dataBloqueio)}
                  </span>
                )}
              </div>

              <Secao titulo="Dados Pessoais">
                <dl className="grid gap-4 sm:grid-cols-3">
                  <Dado rotulo="Matrícula" valor={usuario.matricula} />
                  <Dado rotulo="Nome" valor={usuario.nome} />
                  <Dado rotulo="E-mail" valor={usuario.email} />
                  <Dado rotulo="Nascimento" valor={formatarData(usuario.nascimento)} />
                  <Dado rotulo="Sexo" valor={usuario.sexo === 'M' ? 'Masculino' : 'Feminino'} />
                  <Dado rotulo="Último acesso" valor={formatarData(usuario.ultimoAcesso)} />
                </dl>
              </Secao>

              <Secao titulo="Dados de Acesso">
                <dl className="grid gap-4 sm:grid-cols-3">
                  <Dado rotulo="Usuário" valor={usuario.login} />
                  <Dado rotulo="Usuário AD" valor={usuario.userAd} />
                  <Dado rotulo="Data de cadastro" valor={formatarData(usuario.dataCadastro)} />
                </dl>
              </Secao>

              <Secao titulo="Vínculo">
                <dl className="grid gap-4 sm:grid-cols-3">
                  <Dado rotulo="Tipo" valor={nomeTipoUsuario(usuario.tipo)} />
                  <Dado rotulo="Empresa" valor={`${usuario.empresa} — ${nomeEmpresa(usuario.empresa)}`} />
                  <Dado rotulo="Ramal" valor={usuario.ramal} />
                </dl>
              </Secao>

              <div className="mt-4 flex justify-end gap-2">
                <Botao
                  variante="neutro"
                  icone={FileDown}
                  aoClicar={exportarPdf}
                  titulo="Exportar relatório em PDF"
                >
                  PDF
                </Botao>
                <Botao
                  variante={desativando ? 'perigo' : 'sucesso'}
                  icone={desativando ? PowerOff : Power}
                  aoClicar={() => setConfirmando(true)}
                  desabilitado={processando}
                >
                  {processando ? 'Processando...' : desativando ? 'Desativar' : 'Ativar'}
                </Botao>
              </div>
            </>
          )}
        </Painel>
      </div>

      <ConfirmDialog
        aberto={confirmando}
        titulo={desativando ? 'Desativar usuário' : 'Ativar usuário'}
        mensagem={
          usuario
            ? desativando
              ? `O usuário ${usuario.nome} perderá o acesso ao SWnet. Deseja continuar?`
              : `O acesso do usuário ${usuario.nome} será restabelecido. Deseja continuar?`
            : ''
        }
        rotuloConfirmar={desativando ? 'Desativar' : 'Ativar'}
        aoConfirmar={confirmar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
