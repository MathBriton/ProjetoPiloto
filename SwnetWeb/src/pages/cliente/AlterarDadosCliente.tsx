import { PaginaSolicitacaoCliente } from '@/components/cliente/PaginaSolicitacaoCliente'

/**
 * Refatoração de Cliente/Alterar-Dados-Do-Cliente.
 *
 * Fluxo: Em digitação -> Diretor Comercial -> Financeiro -> Controladoria
 * -> Finalização -> Finalizado.
 */
export function AlterarDadosCliente() {
  return <PaginaSolicitacaoCliente tipo="ALTERACAO" titulo="Alterar Dados do Cliente" />
}
