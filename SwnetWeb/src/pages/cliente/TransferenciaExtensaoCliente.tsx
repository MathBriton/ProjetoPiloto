import { PaginaSolicitacaoCliente } from '@/components/cliente/PaginaSolicitacaoCliente'

/**
 * Refatoração de Cliente/Transferencia-Extensao-De-Cliente.
 *
 * Fluxo: Em digitação -> Fiscal -> Diretor Comercial -> CDAG ->
 * Controladoria -> Finalizado. Exige documentação anexada antes do envio.
 */
export function TransferenciaExtensaoCliente() {
  return <PaginaSolicitacaoCliente tipo="EXTENSAO" titulo="Transferência / Extensão de Cliente" />
}
