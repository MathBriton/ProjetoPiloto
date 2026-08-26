import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom'
import { AuthProvider } from '@/context/AuthContext'
import { ThemeProvider } from '@/context/ThemeContext'
import { ProtectedRoute } from '@/components/ProtectedRoute'
import { AppShell } from '@/components/layout/AppShell'
import { ITENS_PENDENTES } from '@/config/menu'
import { Login } from '@/pages/Login'
import { Home } from '@/pages/Home'
import { EmConstrucao } from '@/pages/EmConstrucao'
import { Relatorios } from '@/pages/Relatorios'

import { AcessosUsuarios } from '@/pages/administrador/AcessosUsuarios'
import { Usuarios } from '@/pages/administrador/Usuarios'
import { AlterarUsuario } from '@/pages/administrador/AlterarUsuario'
import { AtivarDesativarUsuarios } from '@/pages/administrador/AtivarDesativarUsuarios'
import { CadastroUsuario } from '@/pages/administrador/CadastroUsuario'
import { CancelarUsuarios } from '@/pages/administrador/CancelarUsuarios'
import { CriacaoUsuarioAd } from '@/pages/administrador/CriacaoUsuarioAd'
import { GerenciarMenus } from '@/pages/administrador/GerenciarMenus'
import { PermissoesEspeciais } from '@/pages/administrador/PermissoesEspeciais'
import { RelatorioPermissoes } from '@/pages/administrador/RelatorioPermissoes'
import { RevisaoUsuarios } from '@/pages/administrador/RevisaoUsuarios'

import { AprovadoresFnc } from '@/pages/cadastrosFnc/AprovadoresFnc'
import { DefeitosFnc } from '@/pages/cadastrosFnc/DefeitosFnc'
import { ResponsaveisFnc } from '@/pages/cadastrosFnc/ResponsaveisFnc'
import { SetoresFnc } from '@/pages/cadastrosFnc/SetoresFnc'

import { AnalistasOrcamento } from '@/pages/cadastrosOrcamento/AnalistasOrcamento'
import { AprovadoresOrcamento } from '@/pages/cadastrosOrcamento/AprovadoresOrcamento'
import { FaixasAprovador } from '@/pages/cadastrosOrcamento/FaixasAprovador'

import { ControleRecebimento } from '@/pages/cavacoAparas/ControleRecebimento'
import { FornecedoresAparas } from '@/pages/cavacoAparas/FornecedoresAparas'
import { RelatorioRecebimento } from '@/pages/cavacoAparas/RelatorioRecebimento'

import { ClassificarNotas } from '@/pages/compras/ClassificarNotas'

import { AlterarRisco } from '@/pages/financeiro/AlterarRisco'
import { CadastroCanhoto } from '@/pages/financeiro/CadastroCanhoto'
import { SolicitarAumentoCredito } from '@/pages/financeiro/SolicitarAumentoCredito'
import { TransferenciaCredito } from '@/pages/financeiro/TransferenciaCredito'
import { ValidadoresCredito } from '@/pages/financeiro/ValidadoresCredito'
import { ChecklistTransporte } from '@/pages/expedicao/ChecklistTransporte'
import { SimuladorCargas } from '@/pages/expedicao/SimuladorCargas'
import { AlterarSenha } from '@/pages/utilitarios/AlterarSenha'

export function App() {
  return (
    <ThemeProvider>
      <AuthProvider>
        <BrowserRouter>
          <Routes>
            <Route path="/login" element={<Login />} />

            <Route element={<ProtectedRoute />}>
              <Route element={<AppShell />}>
                <Route path="/" element={<Home />} />

                {/* Administrador */}
                <Route path="/administrador/acessos-usuarios" element={<AcessosUsuarios />} />
                <Route path="/administrador/alterar-usuario" element={<AlterarUsuario />} />
                <Route path="/administrador/ativar-desativar" element={<AtivarDesativarUsuarios />} />
                <Route path="/administrador/cadastro-usuario" element={<CadastroUsuario />} />
                <Route path="/administrador/cancelar-usuarios" element={<CancelarUsuarios />} />
                <Route path="/administrador/criacao-usuario-ad" element={<CriacaoUsuarioAd />} />
                <Route path="/administrador/gerenciar-menus" element={<GerenciarMenus />} />
                <Route path="/administrador/permissoes-especiais" element={<PermissoesEspeciais />} />
                <Route path="/administrador/relatorio-permissoes" element={<RelatorioPermissoes />} />
                <Route path="/administrador/revisao-usuarios" element={<RevisaoUsuarios />} />
                {/* Listagem genérica do scaffold — fora do menu, mas ainda navegável */}
                <Route path="/administrador/usuarios" element={<Usuarios />} />

                {/* Cadastros FNC */}
                <Route path="/cadastros-fnc/aprovadores" element={<AprovadoresFnc />} />
                <Route path="/cadastros-fnc/defeitos" element={<DefeitosFnc />} />
                <Route path="/cadastros-fnc/responsaveis" element={<ResponsaveisFnc />} />
                <Route path="/cadastros-fnc/setores" element={<SetoresFnc />} />

                {/* Cadastros Orçamento */}
                <Route path="/cadastros-orcamento/analistas" element={<AnalistasOrcamento />} />
                <Route path="/cadastros-orcamento/aprovadores" element={<AprovadoresOrcamento />} />
                <Route path="/cadastros-orcamento/faixas-aprovador" element={<FaixasAprovador />} />

                {/* Cavaco / Aparas */}
                <Route path="/cavaco-aparas/fornecedores" element={<FornecedoresAparas />} />
                <Route path="/cavaco-aparas/controle-recebimento" element={<ControleRecebimento />} />
                <Route path="/cavaco-aparas/relatorio-recebimento" element={<RelatorioRecebimento />} />

                {/* Compras */}
                <Route path="/compras/classificar-notas" element={<ClassificarNotas />} />

                {/* Financeiro */}
                <Route path="/financeiro/alterar-risco" element={<AlterarRisco />} />
                <Route path="/financeiro/cadastro-canhoto" element={<CadastroCanhoto />} />
                <Route path="/financeiro/validadores-credito" element={<ValidadoresCredito />} />
                <Route path="/financeiro/aumento-credito" element={<SolicitarAumentoCredito />} />
                <Route path="/financeiro/transferencia-credito" element={<TransferenciaCredito />} />

                {/* Expedição */}
                <Route path="/expedicao/checklist-transporte" element={<ChecklistTransporte />} />
                <Route path="/expedicao/simulador-cargas" element={<SimuladorCargas />} />

                {/* Utilitários */}
                <Route path="/utilitarios/alterar-senha" element={<AlterarSenha />} />

                {/*
                  Menus cujo fonte ASP ainda não veio no material de refatoração:
                  a rota existe e abre o aviso de pendência, sem inventar a tela.
                */}
                {ITENS_PENDENTES.map((item) => (
                  <Route key={item.caminho} path={item.caminho} element={<EmConstrucao />} />
                ))}

                <Route path="/relatorios" element={<Relatorios />} />
              </Route>
            </Route>

            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </BrowserRouter>
      </AuthProvider>
    </ThemeProvider>
  )
}
