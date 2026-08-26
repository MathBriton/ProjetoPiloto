<%RESPONSE.BUFFER=false%>

<!--#Include File = "../../../corporativo/includes/variaveis.asp"-->
<!--#include file="../../../corporativo/cnn/cnn.asp" -->
<!--#include file="../../../corporativo/funcoes/funcoes.asp" -->

<%
call ValidaPermissaoRelatorios("../../inpanet/dptos/financeiro/consultar_risco",SESSION("CODUSER"),SESSION("FILIAL")) 
%>

<html>
  <head>
    <title>Consulta Clientes bloqueados para emiss&atilde;o de pedidos</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  </head>

  <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

    <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#0000CC">
      <tr>
        <td>
          <img src="images/barra_sup.gif" width="700" height="80" align="middle">
        </td>
      </tr>
    </table>

    <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#0000CC">
      <tr>
        <td>&nbsp; </td>
      </tr>
    </table>

    <table width="700" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#0000CC">
      <tr>
        <td>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="20" bgcolor="#CC9933"> 
                <div align="center">
                  <font face="Tahoma, verdana, arial" size="1"><b>CONSULTA DE CLIENTE COM RISCO &quot;E&quot; BLOQUEADO PARA EMISS&Atilde;O DE PEDIDOS</b></font>
                </div>
              </td>
            </tr>
            <tr bgcolor="#EAD5A8"> 
              <td height="20"> 
                <div align="center">
                  <font face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1></font>
                </div>
              </td>
            </tr>
            <tr> 
              <td>
                <img src="images/t.gif" width="50" height="10">
              </td>
            </tr>
          </table>
          <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCFF">
            <tr>             
              <td align="center" valign="top"> 
                <%

                  'MARRETADA PARA FUNCIONAR COM O CODIGO DO USUARIO QUE NAO � REPRESENTANTE
                  if nivel = 0 then
	                  codsiga = " "
                  end if

	                VarSQLCli = "Execute SP_CONSULTACLIENTESRISCOE [" & codsiga & "]"
	                Dados.Source = VarSQLCli
                  Dados.Open()

                  if Dados.eof then%>
                  
		                <script language="JavaScript">
		                  <!--  Begin
		                  function drawAlert () {
		                    alert ("N�o foram encontrados registros. \nTente novamente.");
		                  }
		                  drawAlert()
		                  // End -->
		                </script> 
    	              <%response.end

                  else	

	                  Response.Write "<TABLE BORDER='0'>"
                    Response.Write "<TR>"
                    if nivel <> 0 then 
                      Response.Write "<TD align='left'><font face=tahoma, arial size=3><div align=left><P><B>"& Dados("A1_VEND")&" - "& "</div></B></TD>" 
                    end if
                    if nivel <> 0 then 
                      Response.Write "<TD align='left'><font face=tahoma, arial size=3><div align=left><P><B>"& Dados("A3_NOME") & "</div></B></TD>"
                    end if

                    Response.Write "</TR>"
                    Response.Write "</TABLE>"
                    Response.Write "<TABLE BORDER='1' cellspacing='2' cellpadding='2' bordercolor='#FFCC99'>"
                    Response.Write "<TR>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>COD_CLI</DIV></TD>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>NOME_CLI</DIV></TD>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>LOJA</DIV></TD>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>CGC</DIV></TD>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>RISCO</DIV></TD>"
                    Response.Write "<TD bgcolor='#6699FF'><font face=tahoma, arial size=1><div align=center><B>DESCRI��O</DIV></TD>"

                    Response.Write "</TR>"

	                  do until Dados.EOF
	
                      'CRIANDO A DESCRI��O
                      If Dados("a1_risco") = "E" then
                        descri = "Bloqueado"
                      elseif  Dados("a1_risco") = "E" then
                        descri = "Bloqueado p/ Emiss�o"
                      else
                        descri = "Liberado"
                      end if
                      
                      Response.Write "<TR>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("A1_COD") & "</DIV></TD>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=left><P>" & Dados("A1_NOME") &"</DIV></TD>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("A1_LOJA") &"</DIV></TD>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=right><P>" & Dados("A1_CGC") &"</DIV></TD>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("A1_RISCO") & "</DIV></TD>"
                      Response.Write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & descri & "</DIV></TD>"
                      Response.Write "</TR>"

                      REGISTROS = REGISTROS + 1 
                      Dados.MoveNext

	                  Loop

                    Dados.close

	                end if
                %>

                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td>
                      <img src="../images/t.gif" width="50" height="10"><img src="file:///Y|/inpanet/dptos/images/t.gif" width="1" height="1">
                    </td>
                  </tr>
                  <tr bgcolor="#CC9933"> 
                    <td height="20"> 
                      <div align="center"> 
                        <strong>
                          <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%Response.Write("Foram encontrados " & registros & " clientes bloqueados.")%>&nbsp;
                          </font>
                        </strong>
                      </div>
                    </td>
                  </tr>
                </table>

              </td>
            </tr>
          </table>

          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>    
              <td>
                <img src="images/t.gif" width="50" height="10">
              </td>
            </tr>
          </table>

        </td>
      </tr>
    </table>

    <font face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="1"> </font>
    </font>

  </body>

</html>
