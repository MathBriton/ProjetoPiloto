import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom'
import { AuthProvider } from '@/context/AuthContext'
import { ThemeProvider } from '@/context/ThemeContext'
import { ProtectedRoute } from '@/components/ProtectedRoute'
import { AppShell } from '@/components/layout/AppShell'
import { Login } from '@/pages/Login'
import { Home } from '@/pages/Home'
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

import { AlterarDadosCliente } from '@/pages/cliente/AlterarDadosCliente'
import { AlterarStatusCliente } from '@/pages/cliente/AlterarStatusCliente'
import { CadastroUnificado } from '@/pages/cliente/CadastroUnificado'
import { TransferenciaExtensaoCliente } from '@/pages/cliente/TransferenciaExtensaoCliente'

import { AlteracaoProdutos } from '@/pages/desenvolvimento/AlteracaoProdutos'
import { AprovacaoComercialSpd } from '@/pages/desenvolvimento/AprovacaoComercialSpd'
import { LotePiloto } from '@/pages/desenvolvimento/LotePiloto'
import { SolicitacaoSpd } from '@/pages/desenvolvimento/SolicitacaoSpd'

import { AlterarRisco } from '@/pages/financeiro/AlterarRisco'
import { CadastroCanhoto } from '@/pages/financeiro/CadastroCanhoto'
import { SolicitarAumentoCredito } from '@/pages/financeiro/SolicitarAumentoCredito'
import { TransferenciaCredito } from '@/pages/financeiro/TransferenciaCredito'
import { ValidadoresCredito } from '@/pages/financeiro/ValidadoresCredito'
import { ChecklistTransporte } from '@/pages/expedicao/ChecklistTransporte'
import { SimuladorCargas } from '@/pages/expedicao/SimuladorCargas'
import { AlterarSenha } from '@/pages/utilitarios/AlterarSenha'

import { FichaQm } from '@/pages/sgq/FichaQm'
import { FncBrasil } from '@/pages/sgq/FncBrasil'

import { AlterarPedidosBrasil } from '@/pages/pedidos/AlterarPedidosBrasil'
import { AlterarPedidosPiraUbe } from '@/pages/pedidos/AlterarPedidosPiraUbe'
import { EntregaFimDeSemana } from '@/pages/pedidos/EntregaFimDeSemana'
import { InclusaoObsNf } from '@/pages/pedidos/InclusaoObsNf'
import { OrcamentoBrasil } from '@/pages/pedidos/OrcamentoBrasil'
import { ParametrosPedidos } from '@/pages/pedidos/ParametrosPedidos'
import { PedidosBrasil } from '@/pages/pedidos/PedidosBrasil'
import { PedidosDigitadosBrasil } from '@/pages/pedidos/PedidosDigitadosBrasil'
import { PedidosDigitadosPiraUbe } from '@/pages/pedidos/PedidosDigitadosPiraUbe'
import { PedidosPiraUbe } from '@/pages/pedidos/PedidosPiraUbe'

import { AtendentesRh, CategoriasChamado, DepartamentosChamado, PrioridadesChamado } from '@/pages/rh/CadastrosRh'
import { FilaChamadosRh, MeusChamadosRh } from '@/pages/rh/ChamadosRh'
import { ImportarAdp } from '@/pages/rh/ImportarAdp'
import { NovoChamadoRh } from '@/pages/rh/NovoChamadoRh'

import { SolicitacoesCredito } from '@/pages/validacoes/SolicitacoesCredito'
import { ValidacaoPapel } from '@/pages/validacoes/ValidacaoPapel'
import { ValidarPedido } from '@/pages/validacoes/ValidarPedido'
import { ValidarPedidoBrasil } from '@/pages/validacoes/ValidarPedidoBrasil'

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

                {/* Cliente */}
                <Route path="/cliente/alterar-dados" element={<AlterarDadosCliente />} />
                <Route path="/cliente/alterar-status" element={<AlterarStatusCliente />} />
                <Route path="/cliente/cadastro-unificado-v2" element={<CadastroUnificado />} />
                <Route path="/cliente/transferencia-extensao" element={<TransferenciaExtensaoCliente />} />

                {/* Desenvolvimento */}
                <Route path="/desenvolvimento/alteracao-produtos" element={<AlteracaoProdutos />} />
                <Route path="/desenvolvimento/aprovacao-comercial-spd" element={<AprovacaoComercialSpd />} />
                <Route path="/desenvolvimento/lote-piloto-v2" element={<LotePiloto />} />
                <Route path="/desenvolvimento/solicitacao-spd-v2" element={<SolicitacaoSpd />} />

                {/* Financeiro */}
                <Route path="/financeiro/alterar-risco" element={<AlterarRisco />} />
                <Route path="/financeiro/cadastro-canhoto" element={<CadastroCanhoto />} />
                <Route path="/financeiro/validadores-credito" element={<ValidadoresCredito />} />
                <Route path="/financeiro/aumento-credito" element={<SolicitarAumentoCredito />} />
                <Route path="/financeiro/transferencia-credito" element={<TransferenciaCredito />} />

                {/* Expedição */}
                <Route path="/expedicao/checklist-transporte" element={<ChecklistTransporte />} />
                <Route path="/expedicao/simulador-cargas" element={<SimuladorCargas />} />

                {/* Pedidos */}
                <Route path="/pedidos/alterar-brasil" element={<AlterarPedidosBrasil />} />
                <Route path="/pedidos/alterar-piraube" element={<AlterarPedidosPiraUbe />} />
                <Route path="/pedidos/entrega-fim-de-semana" element={<EntregaFimDeSemana />} />
                <Route path="/pedidos/inclusao-obs-nf" element={<InclusaoObsNf />} />
                <Route path="/pedidos/orcamento-brasil" element={<OrcamentoBrasil />} />
                <Route path="/pedidos/parametros" element={<ParametrosPedidos />} />
                <Route path="/pedidos/brasil-v2" element={<PedidosBrasil />} />
                <Route path="/pedidos/digitados-brasil" element={<PedidosDigitadosBrasil />} />
                <Route path="/pedidos/digitados-piraube" element={<PedidosDigitadosPiraUbe />} />
                <Route path="/pedidos/piraube" element={<PedidosPiraUbe />} />

                {/* RH */}
                <Route path="/rh/atendentes" element={<AtendentesRh />} />
                <Route path="/rh/categorias-chamado" element={<CategoriasChamado />} />
                <Route path="/rh/chamados" element={<FilaChamadosRh />} />
                <Route path="/rh/departamentos-chamado" element={<DepartamentosChamado />} />
                <Route path="/rh/importar-adp" element={<ImportarAdp />} />
                <Route path="/rh/meus-chamados" element={<MeusChamadosRh />} />
                <Route path="/rh/novo-chamado" element={<NovoChamadoRh />} />
                <Route path="/rh/prioridades-chamado" element={<PrioridadesChamado />} />

                {/* SGQ */}
                <Route path="/sgq/fnc-brasil" element={<FncBrasil />} />
                <Route path="/sgq/ficha-qm" element={<FichaQm />} />

                {/* Validações */}
                <Route path="/validacoes/papel" element={<ValidacaoPapel />} />
                <Route path="/validacoes/solicitacoes-credito" element={<SolicitacoesCredito />} />
                <Route path="/validacoes/validar-pedido" element={<ValidarPedido />} />
                <Route path="/validacoes/validar-pedido-brasil" element={<ValidarPedidoBrasil />} />

                {/* Utilitários */}
                <Route path="/utilitarios/alterar-senha" element={<AlterarSenha />} />


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
