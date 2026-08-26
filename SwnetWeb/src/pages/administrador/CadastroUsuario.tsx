import { useState } from 'react'
import { RotateCcw, Save, UserPlus } from 'lucide-react'
import { Painel } from '@/components/comuns/Painel'
import { Botao, TituloPagina } from '@/components/comuns/Comuns'
import {
  FormularioUsuarioSwnet,
  VALORES_INICIAIS,
  validarFormulario,
  type ErrosFormularioUsuario,
  type ValoresFormularioUsuario
} from '@/components/administrador/FormularioUsuarioSwnet'
import { ToastStack, useToast } from '@/components/Toast'
import { criarUsuario, validarMatricula } from '@/services/administrador/usuariosSwnetApi'

/** Refatoração de Cadastro-Usuario-SWNET/CadUsuario.asp. */
export function CadastroUsuario() {
  const { toasts, mostrarToast, fecharToast } = useToast()

  const [valores, setValores] = useState<ValoresFormularioUsuario>(VALORES_INICIAIS)
  const [erros, setErros] = useState<ErrosFormularioUsuario>({})
  const [salvando, setSalvando] = useState(false)

  function mudar<C extends keyof ValoresFormularioUsuario>(campo: C, valor: ValoresFormularioUsuario[C]) {
    setValores((atual) => ({ ...atual, [campo]: valor }))
    setErros((atual) => ({ ...atual, [campo]: undefined }))
  }

  function limpar() {
    setValores(VALORES_INICIAIS)
    setErros({})
  }

  async function salvar() {
    const novosErros = validarFormulario(valores, 'criar')

    // Matrícula duplicada / com menos de 6 dígitos é validada junto com o "banco"
    const erroMatricula =
      !novosErros.matricula && valores.empresa
        ? validarMatricula(valores.matricula, valores.empresa)
        : null
    if (erroMatricula) novosErros.matricula = erroMatricula

    setErros(novosErros)

    if (Object.values(novosErros).some(Boolean)) {
      mostrarToast('Verifique os campos destacados antes de salvar.', 'warning')
      return
    }

    setSalvando(true)
    try {
      const resultado = await criarUsuario({
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
        enviarEmail: valores.enviarEmail
      })

      if (resultado.sucesso) {
        mostrarToast(resultado.mensagem, 'success')
        limpar()
      } else {
        mostrarToast(resultado.mensagem, 'danger')
      }
    } finally {
      setSalvando(false)
    }
  }

  return (
    <div className="space-y-4">
      <TituloPagina titulo="Cadastro de Usuário" />

      <div className="mx-auto max-w-4xl">
        <Painel titulo="Incluir Usuário" icone={UserPlus}>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              salvar()
            }}
          >
            <FormularioUsuarioSwnet
              modo="criar"
              valores={valores}
              erros={erros}
              aoMudar={mudar}
              desabilitado={salvando}
            />

            <div className="mt-4 flex justify-end gap-2">
              <Botao variante="neutro" icone={RotateCcw} aoClicar={limpar} desabilitado={salvando}>
                Limpar
              </Botao>
              <Botao tipo="submit" variante="sucesso" icone={Save} desabilitado={salvando}>
                {salvando ? 'Salvando...' : 'Salvar'}
              </Botao>
            </div>
          </form>
        </Painel>
      </div>

      <ToastStack toasts={toasts} aoFechar={fecharToast} />
    </div>
  )
}
