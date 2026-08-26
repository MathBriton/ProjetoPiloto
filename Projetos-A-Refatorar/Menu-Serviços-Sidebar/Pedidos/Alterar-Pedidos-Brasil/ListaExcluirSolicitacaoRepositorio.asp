<!-- #include file = "../../cnn/cnn.asp" -->
<!-- #include file = "EmailExtratoPedido.asp" -->
<!-- #include file = "../../includes/variaveis.asp"-->
<%
' -----------------------------------------------------------------------------------------------------------------------------------------------------
' Empresa        :  Smurfit Kappa
' Sistema        :  SKnet
' Fonte          :  ListaExcluirSolicitacaoRepositorio.asp
' Funcionalidade :  Excluir a solicitação de alteração de pedido
' Parâmetro(s)   :  Pedido

' ATUALIZAÇÕES

' CHPL           :  371 - Inserçao de Fortaleza e Saquarema na rotina de alteraçao de pedidos.
' Analista       :  Jorge Queiroz
' Data           :  17/05/2024
' -----------------------------------------------------------------------------------------------------------------------------------------------------

session.lcid = 1046
idPedido = Request.QueryString("idPedido")
campoUsuario=  SESSION("Codigo")  'verificar 

'response.write "EXCLUSAO "
'response.end

SQL = "" &_
  "INSERT INTO dbo.HistoricoAlteracaoPedido " &_
    "SELECT p.Id " &_
      "," & campoUsuario & " as usuario " &_
      ",getdate() as DataOperacao " &_
      ",2 as Status " &_
    "FROM HistoricoSolicitacaoAlteracaoPedido p " &_
      "WHERE p.Id = " & idPedido 

Dados131.open(SQL)

SQL = "" &_
  "SELECT " &_
    "Id " &_
  "FROM HistoricoAlteracaoPedido " &_
  "WHERE Status = 2 " &_
    "AND Id = "& idPedido 
 
Dados132.open(SQL)

if not Dados132.eof then
  response.write "Registro descartado com sucesso!"  
else
  response.write "Nenhum registro foi descartado!"
end if

Dados132.Close()

response.end()
%>
