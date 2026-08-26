<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->

<%
'Option Explicit
Response.ContentType = "text/plain"

planta = trim(Request.form("planta"))

On Error Resume next

select case planta
case "00" 
  SET PIR_RBACESSO = Conecta( "PIR-RBACESSO", "EXEC sp_IntegraAdpMadisPirapetinga" )
  SET PIR_RBACESSO = NOTHING
case "04" 
  SET UBE_DMPACESSO = Conecta( "UBE-DMPACESSO", "EXEC sp_IntegraAdpMadisUberaba" )
  SET UBE_DMPACESSO = NOTHING
case "05" 
  SET BEN_MDACESSO = Conecta( "BEN-MDACESSO", "EXEC sp_IntegraAdpMadisBento" )
  SET BEN_MDACESSO = NOTHING
end select

On Error GoTo 0

Response.Write "<br>Importação Realizada com sucesso."
%>