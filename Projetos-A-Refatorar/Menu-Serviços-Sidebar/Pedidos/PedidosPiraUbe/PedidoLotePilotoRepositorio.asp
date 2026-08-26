<!--#include file="../../cnn2/cnn2.asp" -->
<%
'On Error Resume Next
if Trim(Request("paramCliente")) = "" or Trim(Request("paramLoja")) = "" or Trim(Request("paramProduto")) = "" or Trim(Request("paramPlanta")) = "" or Trim(Request("paramQuantidade")) = ""  then
       RESPONSE.WRITE "Faltam parametros para verificar o Lote Piloto!"
      'RESPONSE.END
else
    SQL = "SELECT CONCAT(LOTE_PILOTO_ENCONTRADO,'|', VALIDA_PEDIDO, '|', VALIDA_PALETE , '|', VALIDA_ONDULADEIRA , '|', VALIDA_IMPRESSORA , '|', "
    SQL = SQL & " VALIDA_EXPEDICAO , '|', VALIDA_CLIENTE, '|', VALIDA_DDP, '|', VALIDA_CLICHERIA, '|', TOTAL_PALETE , '|', MENSAGEM , '|', MENSAGEM_PEDIDO , '|', MENSAGEM_PALETE, '|', ATUALIZA_STATUS) as Dados "
    SQL = SQL & " FROM dbo.FUNC_VALIDA_LOTE_PILOTO('"&Trim(Request("paramCliente"))&"','"&Trim(Request("paramLoja"))&"','"&Trim(Request("paramProduto"))&"','"&Trim(Request("paramPlanta"))&"',"&Trim(Request("paramQuantidade"))&")"
   
    set Consulta = Conecta( "PIR-CORPORATIVO",SQL )
    if not isobject(Consulta) then
        response.end
    end if
    if (not Consulta.eof) then
        response.write Consulta("Dados")
    else
        RESPONSE.WRITE ""
    end if
    Consulta.Close
end if

On Error GoTo 0
'RESPONSE.END
%>