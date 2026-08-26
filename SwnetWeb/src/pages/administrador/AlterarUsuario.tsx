import { useState } from 'react'
import { Save, UserCog, UserSearch } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, EstadoVazio, TituloPagina } from '@/components/comuns/Comuns'
import { SeletorUsuario } from '@/components/comuns/SeletorUsuario'
import {
  FormularioUsuarioSwnet,
  valoresDoUsuario,
  validarFormulario,
  type ErrosFormularioUsuario,
  type ValoresFormularioUsuario
} from '@/components/administrador/FormularioUsuarioSwnet'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ToastStack, useToast } from '@/components/Toast'
import { alterarUsuario } from '@/services/administrador/usuariosSwnetApi'
import type { UsuarioSwnet } from './tipos'

/** Refatoração de Alterar-Usuario-Swnet/AltUsuario.asp. */
export function AlterarUsuario() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [usuario, setUsuario] = useState<UsuarioSwnet | null>(null)
  const [valores, setValores] = useState<ValoresFormularioUsuario | null>(null)
  const [erros, setErros] = useState<ErrosFormularioUsuario>({})
  const [salvando, setSalvando] = useState(false)
  const [confirmando, setConfirmando] = useState(false)

  function selecionar(selecionado: UsuarioSwnet | null) {
    setUsuario(selecionado)
    setValores(selecionado ? valoresDoUsuario(selecionado) : null)
    setErros({})
  }

  function mudar<C extends keyof ValoresFormularioUsuario>(campo: C, valor: ValoresFormularioUsuario[C]) {
    setValores((atual) => (atual ? { ...atual, [campo]: valor } : atual))
    setErros((atual) => ({ ...atual, [campo]: undefined }))
  }

  function validarEConfirmar() {
    if (!valores) return

    const novosErros = validarFormulario(valores, 'alterar')
    setErros(novosErros)

    if (Object.values(novosErros).some(Boolean)) {
      mostrarToast('Verifique os campos destacados antes de salvar.', 'warning')
      return
    }
    setConfirmando(true)
  }

  async function salvar() {
    if (!usuario || !valores) return

    setConfirmando(false)
    setSalvando(true)
    try {
      const resultado = await alterarUsuario({
        id: usuario.id,
        matricula: valores.matricula,
        nome: valores.nome,
        sexo: valores.sexo,
        nascimento: valores.nascimento,
        email: valores.email,
        representante: valores.representante,
        prefixoAd: valores.prefixoAd,
        login: valores.login,
        senha: valores.senha,
        tipo: valores.tipo,
        empresa: valores.empresa,
        ramal: valores.ramal,
        bloqueado: valores.bloqueado
      })

      if (resultado.sucesso) {
        mostrarToast(resultado.mensagem, 'success')
        // Recarrega o formulário para limpar os campos de senha
        setValores((atual) => (atual ? { ...atual, senha: '', confirmaSenha: '' } : atual))
      } else {
        mostrarToast(resultado.mensagem, 'danger')
      }
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Alterar Usuário" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Alterar Usuário" icone={UserCog}>
          <div className="mb-4">
            <SeletorUsuario selecionado={usuario} aoSelecionar={selecionar} />
          </div>

          {!valores && (
            <EstadoVazio
              icone={UserSearch}
              titulo="Selecione um usuário para editar"
              descricao="Busque pelo nome ou login. Os dados cadastrais aparecem aqui para alteração."
            />
          )}

          {valores && (
            <form
              onSubmit={(e) => {
                e.preventDefault()
                validarEConfirmar()
              }}
            >
              <FormularioUsuarioSwnet
                modo="alterar"
                valores={valores}
                erros={erros}
                aoMudar={mudar}
                desabilitado={salvando}
              />

              <div className="mt-4 flex justify-end">
                <Botao tipo="submit" variante="sucesso" icone={Save} desabilitado={salvando}>
                  {salvando ? 'Salvando...' : 'Salvar alterações'}
                </Botao>
              </div>
            </form>
          )}
        </Painel>
      </div>

      <ConfirmDialog
        aberto={confirmando}
        titulo="Confirmar alteração"
        mensagem={
          valores?.senha
            ? `Os dados de ${usuario?.nome} serão atualizados e a senha será redefinida. Deseja continuar?`
            : `Os dados de ${usuario?.nome} serão atualizados. Deseja continuar?`
        }
        rotuloConfirmar="Salvar"
        aoConfirmar={salvar}
        aoCancelar={() => setConfirmando(false)}
      />

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
