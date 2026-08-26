<%SESSION.LCID=1046%>

<html>
<head>
<title>Detalhes do Pedido de Papel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">

.barra_menu {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #FFFFFF; text-decoration: none}
.barra_menu:hover {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #99CCFF; text-decoration: none}

.asp {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; font-weight: normal; text-transform: uppercase; text-decoration: none; line-height: normal; font-variant: normal; color: #000000}
</style>
<link href="../Emissoes/style.css" rel="stylesheet" type="text/css">
</head>

<!--#include file="../../../../includes/formato.asp"--> 
<!--#include file="../../../../includes/variaveis.asp" -->
'<!--#include file="../../../../includes/conecta_banco_siga.inc" --> 'DADSOADV
<!--#include file="../../../../includes/atribui_nivel.inc" --> 
<!--#include file="../../../../cnn/cnn.asp" --> 'DADSOADV

<%
'.........................................................................................
'RECEBENDO PAR�METRO PARA GERAR DETALHES
NumPed = Request.QueryString("pedido")

'**************************************************'
'** Passando dados para a grava�ao do status     **'
'**************************************************'
session("pedido") = NumPed

Dados.Open( "EXEC SP_CONLIBERASZF [" & NumPed & "]" )

if Dados("ZF_TPFRETE") = "F" then
   vTipoFrete = "FOB"
else
   vTipoFrete = "CIF"
end if

vTES = Dados("ZF_TES")
vAvisoTES = ""
if vTES > 502 then
   vAvisoTES = "ATEN��O PARA ESTA TES!!"
end if
vAvisoTES2 = ""
if vTES = 580 or vTES = 581 or vTES = 715 or vTES = 716 then
   vAvisoTES2 = "TES PARA PEDIDO LIVRE DE DEBITO - ATEN��O"
end if

vProduto = Dados("ZF_PRODUTO")
Pr_Pedido = Dados("ZF_PRECO")
Ult_Preco = Dados("B1_UPRC")

Dados1.Open( "EXEC SP_CONULTIMASCOMPRAS [" & vProduto & "]" )

%>

<body bgcolor="#FFFFFF" leftmargin="25" topmargin="5" marginwidth="5" marginheight="5">
<form name="form" method="post" action="gravastatus_pa.asp">
  <table width="100%" height="60" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="38">&nbsp;</td>
      <td width="647"><font color="#3366FF" size="4" face="tahoma"><strong>Detalhes 
        do Pedido de Papel para Valida&ccedil;&atilde;o</strong></font></td>
    </tr>
  </table>
  <table width="704" border="1" cellspacing="0" cellpadding="2" bordercolor="#0000CC" align="center">
  <tr>
      
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#CC9933"> 
            <td height="20" bgcolor="#3366FF"> 
<div align="left"><font face="Tahoma, verdana, arial" size="1"><font color="#000000" size="2" face="tahoma, arial, verdana"><b>&nbsp;</b></font><font size="3" face="Arial, Helvetica, sans-serif"><font face="tahoma, arial" size="2"><b><font color="#FFFFFF">DETALHES 
                DOs PEDIDO</font></b> <font color="#FFFFFF"><%=NumPed%> <b>EMITIDO 
                EM</b> </font></font><font color="#FFFFFF" size="2" face="tahoma, arial"><%=(Dados("ZF_EMISSAO"))%> 
                <b>POR</b> <%=UCASE(Dados("ZF_DIGITAD"))%></font></font></font></div>
            </td>
          </tr>
        </table>
        
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#BFDFFF">
          <tr> 
          <td colspan="2" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="13%" height="25"><b><font size="1" face="Tahoma, verdana, arial">CLIENTE:</font></b></td>
                  <td width="86%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Dados("ZF_CLIENTE") & " - " & Dados("ZF_LOJACLI") & " - " & Dados("A1_NOME") & " - " & Dados("A1_MUN")%></font></td>
              </tr>
                <%if Trim(Dados("ZF_REMESSA")) <> "N" then %>
                <tr style="background-color:yellow">
                    <td width="13%" height="25" colspan="2"><b><font size="1" face="Tahoma, verdana, arial" color="red">LOCAL DE ENTREGA:</font>
                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=Dados("ZF_DESTINO") & " - " & Dados("ZF_LOJADES") & " - " & Dados("CLIENTEREMESSA") & " - " & Dados("MUNICIPIOREMESSA")%></b></font></td>
                </tr>
                <%end if %>
            </table>
          </td>
        </tr>
      </table>
        <% 
		'ATRIBUIDO O C�DIGO DO CLIENTE � VARI�VEL CODCLI
		CODCLI = Dados("ZF_CLIENTE") %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
          <td><img src="../../../images/spacer.gif" width="20" height="6"></td>
        </tr>
      </table>
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#BFDFFF">
          <tr> 
          <td width="50%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;PRODUTO:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Dados("ZF_PRODUTO")%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;REFER&Ecirc;NCIA:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Dados("B1_DESC")& " " &Dados("ZF_GRAMPAP")& " " &Dados("ZF_FORMPAP")& " " &Dados("ZF_DIAMPAP")& " " &Dados("ZF_TUBEPAP") %></font></td>
                </tr>
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;QUANTIDADE:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormatNumber(Dados("ZF_QUANT"),0)%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;PRE&Ccedil;O 
                    1:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%="R$ " & FormatNumber(Pr_Pedido,5)%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;ULT 
                    PRE&Ccedil;O:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%="R$ " & FormatNumber(Dados("B1_UPRC"),5)%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="27%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;ENTREGA:</font></b></td>
                  <td width="73%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Converte(Dados("ZF_ENT1")) & " - " & Converte(Dados("ZF_ENT2"))%></font></td>
                </tr>
              </table>
          </td>
          <td width="50%" valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;ULT 
                    COMPRA:</font></b></td>
                  <td width="70%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Converte(Dados("B1_UCOM")) %></font></td>
                </tr>
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;COND 
                    PGTO:</font></b></td>
                  <td width="70%">
                  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
				  <%=Dados("E4_DESCRI")%>
                  <% If trim(Dados("E4_TIPO")) = "9" Then %>
                  	<br><b>Vencimento: <%=Converte(Dados("ZF_DATA5"))%></b><br>
                  	<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000"><b>Aten��o: Data de Vencimento especificada pelo Vendedor</b></font>
                  <% End If %>
                  </font></td>
                </tr>
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;DESCONTO:</font></b></td>
                  <td width="70%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000"><%=Dados("ZF_DESC")& "%"%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;TRANSPORT.:</font></b></td>
                  <td width="70%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Dados("ZF_TRANSPA")%></font></td>
                </tr>
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;TES:</font></b></td>
                  <td width="70%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=vTES%>&nbsp;-&nbsp;<b style="color:#ff0000"><%=vAvisoTES%></b><br /><b style="color:#ff0000"><%=vAvisoTES2 %></b></font></td>
                </tr>
                <tr> 
                  <td height="25" width="30%"><b><font size="1" face="Tahoma, verdana, arial">&nbsp;FRETE:</font></b></td>
                  <td width="70%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=vTipoFrete %></font></td>
                </tr>
              </table>
          </td>
        </tr>
        <%if Trim(Dados("ZF_MENNOTA")) <> "" Then%>
        <tr bgcolor="yellow">
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial" color="red"><center><b>Aten��o - 1&ordf; Observa��o para Nota Fiscal</b></center></font></td>
        </tr>
        <tr>
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial"><%=Trim(Dados("ZF_MENNOTA")) %></font></td>
        </tr>
        <%end if %>
        <%if Trim(Dados("ZF_MENNOT2")) <> "" Then%>
        <tr bgcolor="yellow">
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial" color="red"><center><b>Aten��o - 2&ordf; Observa��o para Nota Fiscal</b></center></font></td>
        </tr>
        <tr>
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial"><%=Trim(Dados("ZF_MENNOT2")) %></font></td>
        </tr>
        <%end if %>
        <%if Trim(Dados("ZF_MENNOT3")) <> "" Then%>
        <tr bgcolor="yellow">
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial" color="red"><center><b>Aten��o - 3&ordf; Observa��o para Nota Fiscal</b></center></font></td>
        </tr>
        <tr>
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial"><%=Trim(Dados("ZF_MENNOT3")) %></font></td>
        </tr>
        <%end if %>
        <%if Trim(Dados("ZF_MENNOT4")) <> "" Then%>
        <tr bgcolor="yellow">
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial" color="red"><center><b>Aten��o - 4&ordf; Observa��o para Nota Fiscal</b></center></font></td>
        </tr>
        <tr>
            <td colspan="2"><font size="1" face="Tahoma, verdana, arial"><%=Trim(Dados("ZF_MENNOT4")) %></font></td>
        </tr>
        <%end if %>
      </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            
          <td><img src="../../../images/spacer.gif" width="20" height="12"></td>
          </tr>
          
        <tr bgcolor="#CC9933"> 
            <td bgcolor="#3366FF"> 
<div align="center"><font face="Tahoma, verdana, arial" size="1"><img src="../../../images/spacer.gif" width="20" height="12"><b><font color="#FFFFFF">&Uacute;LTIMOS 
                PEDIDOS </font></b></font></div>
            </td>
          </tr>
        </table>
        
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#BFDFFF">
          <tr align="center" valign="middle"> 
            <td colspan="2"> 
<%
if Dados1.EOF then

	Response.Write "</br>"
	Response.WRITE "<table width=450 border=0 cellpadding=0 cellspacing=0 bgcolor=#FF0000 align=center>"
	Response.WRITE "<tr>"
	Response.WRITE "<td align=center valign=top class=texto_atencao style='padding=10px'>ATEN&Ccedil;&Atilde;O<br>"
	Response.WRITE "<table width=100% border=0 cellspacing=0 cellpadding=0>"
	Response.WRITE "<tr>"
	Response.WRITE "<td align=center bgcolor=#FFFFFF style='padding=10px'>"
	Response.WRITE "<table width=100% border=0 cellspacing=0 cellpadding=0>"
	Response.WRITE "<tr>"
	Response.WRITE "<td style=padding=2px><img src=../../../images/alerta_val_user.gif width=31 height=31 align=middle></td>"
	Response.WRITE "<td class=texto_simples2>N�O H� PEDIDOS CADASTRADOS!<br></td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table>"                   
	Response.WRITE "</td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table> </td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table>"
	Response.Write "</br>"

else

  Dados1.MoveFirst
	Response.Write "<TABLE BORDER='1' cellspacing='2' cellpadding='2' bordercolor='#FFCC99'>"
  Response.Write "<TR>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>PEDIDO</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>REFER�NCIAS</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>ENTREGA</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>TES</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>QTD</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>COND PAGTO</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>PRE�O (2)</B></div></TD>"
  Response.Write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>%(1)/(2)</B></div></TD>"
  Response.Write "</TR>"

  Do Until Dados1.EOF
    Pr_Venda = Dados1("C6_PRCVEN")
    Response.Write "<TR>"
    Response.Write "<TD><font face=tahoma, arial size=0><div align=center><P>" & Dados1("C6_NUM") & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arial size=1><div align=center><P>" & Dados1("C6_DESCRI") & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arial size=1><div align=center><P>" & Converte(Dados1("C6_ENTREG")) & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arial size=1><div align=center><P>" & Dados1("C6_TES") & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arialV size=1><div align=center><P>" & FormatNumber(Dados1("C6_QTDVEN"),0) & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arial size=1><div align=center><P>" & Dados1("E4_DESCRI") & "</div></TD>"
    Response.Write "<TD><font face=tahoma, arial size=1><div align=center><P>" & FormatNumber(Pr_Venda,5) & "</div></TD>"

    Perc = PercPreco(Pr_Pedido,Pr_Venda)

    if Perc < 0 then
          Response.Write "<TD><font face=tahoma, arial size=1 color=#FF0000><div align=center><P>" & Perc & "%" & "</div></TD>"
    else
          Response.Write "<TD><font face=tahoma, arial size=1 color=#3333FF><div align=center><P>" & Perc & "%" & "</div></TD>"
    end if		

    Response.Write "</TR>"
    Dados1.MoveNext()

  Loop

  response.write "</table>"

end if

'Fechando os objetos
Dados.Close()
Dados1.Close()

%>
</td>
        </tr>
      </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            
          <td><img src="../../../images/spacer.gif" width="20" height="12"></td>
          </tr>
          
        <tr bgcolor="#CC9933"> 
            <td bgcolor="#3366FF"> 
<div align="center">
              <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr> 
                  <td width="23"><% if pr_pedido < ult_preco then%> 
                    <input type="radio" name="radiobutton" value="PR" checked>
                    <% else %> 
                    <input type="radio" name="radiobutton" value="PR">
                    <% end if %> </td>
                  <td width="276"><b><font color="#FFFFFF" size="2" face="tahoma, arial">Bloqueado 
                    por Pre&ccedil;o</font></b></td>
                  <td width="26"> 
                    <input type="radio" name="radiobutton" value="CP">
                  </td>
                  <td width="223"><b><font color="#FFFFFF" size="2" face="tahoma, arial">Bloqueado 
                    por Cond. Pagto</font></b></td>
                  <td width="24"><% if pr_pedido < ult_preco then%> 
                    <input type="radio" name="radiobutton" value="LI">
                    <% else %> 
                    <input type="radio" name="radiobutton" value="LI" checked>
                    <% end if %> </td>
                  <td width="178"><b><font color="#FFFFFF" size="2" face="tahoma, arial">Liberado</font></b></td>
                </tr>
              </table>
              
              </div>
            </td>
          </tr>
        </table>
        
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#BFDFFF"> 
<div align="center"> 
                <%
'..........................................................................................
' FUN��ES
'..........................................................................................
Function Converte(Data)
   Converte = mid(Data,7,2)& "/" & mid(Data,5,2)& "/" & mid(Data,1,4)
   End Function

Function PercPreco(PP,PV)
   if PV = 0 then
		PercPreco = 1
   else
		v_PercPreco = ((PP/PV)*100) - 100
        Pos = InStr(1,v_PerCPreco,",")
        PercPreco = Mid(v_PercPreco,1,Pos+2)
   end if
End Function

%>
                <input name="Submit" type="submit" class="textfield" value="Confirmar">
                <input name="CODCLI" type="hidden" id="CODCLI" value="<%=CODCLI%>">
              </div></td>
          </tr>
        </table>
        
    </td>
    </tr>
  </table>
  
<font 
 
         face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
</font></font> 
</form>
</body>
</html>
