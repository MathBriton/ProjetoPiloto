<%SESSION.LCID=1046%>
<!--#include file="../../../../includes/variaveis.asp" -->
'<!--#include file="../../../../includes/conecta_banco_siga.inc" -->
<!--#include file="../../../../includes/atribui_nivel.inc" -->
<!--#include file="../../../../cnn/cnn.asp" --> 'DADOSADV
<html>
<head>
<title>Grava&ccedil;&atilde;o da Valida&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">

.barra_menu {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #FFFFFF; text-decoration: none}
.barra_menu:hover {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #99CCFF; text-decoration: none}

.asp {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; font-weight: normal; text-transform: uppercase; text-decoration: none; line-height: normal; font-variant: normal; color: #000000}
</style>
</head>
<%
'**************************'
'**	Recuperando o radio	 **'
'**************************'
Status = request.form("radiobutton")
NumPed = session("pedido")

if Status = "PR" then
	CodBloq  = 17
    DescBloq = "Preco menor que do Pedido anterior"
    Mens = "PEDIDO " & NumPed & " BLOQUEADO POR PRE�O !!"
end if
if Status = "CP" then
	CodBloq  = 30
    DescBloq = "Condi��o de Pagamento" 
    Mens = "PEDIDO " & NumPed & " BLOQUEADO POR COND. PAGTO !!"
end if
if Status = "LI" then
	CodBloq  = " "
    DescBloq = " "
    Mens = "PEDIDO " & NumPed & " LIBERADO !!"
end if

'**************************************'
'***        Conectando ao BD        ***'
'**************************************'
'DataHora = Now
'Espaco = InStr(1,DataHora, " ")
'HW = Mid(DataHora,1,Espaco-1)
 
'if Len(HW) = 10 then
'    Hoje = Mid(HW,7,4) & Mid(HW,4,2) & Mid(HW,1,2) 
'else
'    Hoje = "20" & Mid(HW,7,2) & Mid(HW,4,2) & Mid(HW,1,2) 
'end if


'DoisPontos = InStr(1,DataHora,":")
'Hora = Mid(DataHora,DoisPontos-2,2) & Mid(DataHora,DoisPontos,6) 

'CHAMA DATA DA FUN��O INVERTEDATASIGA AAAAMMDD E HORA DO TIME LCID 1046
Hoje = Trim(InverteDataSiga(Date()))
Hora = Trim(Time())
'.......................................................................................................
'VERIFICANDO SE O CLIENTE EST� BLOQUEADO

'RECUPERANDO A C�DIGO DO CLIENTE
CODCLI=REQUEST.Form("CODCLI")
Dados.Open( "EXEC SP_RISCO ["&CodCli&"] " )

if NOT Dados.EOF then

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
	Response.WRITE "<td class=texto_simples2>Cliente bloqueado pelo Departamento Financeiro!</br>Motivo: Risco E!<br>Clique <a href=# onMouseDown='history.go(-1);'>aqui</a> para voltar.</td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table>"                   
	Response.WRITE "</td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table> </td>"
	Response.WRITE "</tr>"
	Response.WRITE "</table>"
	Response.Write "</br>"
	
	Dados.Close()
	Response.End

end if

'Fechando o objeto
Dados.Close()

'........................................................................................................

'VarSQLPED = "Execute SP_LIBERAPEDIDO ["&NumPed&"],["&Status&"],["&CodBloq&"],["&DescBloq&"],["&usr&"],["&Hoje&"],["&TRIM(Hora)&"]"
Dados.Open( "EXEC SP_LIBERAPEDIDO ["&NumPed&"],["&Status&"],["&CodBloq&"],["&DescBloq&"],["&USR&"],["&Hoje&"],["&Hora&"]" )

Dados.Close()

%>
<!--#include file="../../../../includes/conecta_banco.inc" -->
<!--#include file="../../../estatisticas.asp" -->

<script language="JavaScript">
location.href="pa.asp?param=1";
</script>

<body bgcolor="#FFFFFF" leftmargin="25" topmargin="5" marginwidth="5" marginheight="5">
<font face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
  </font></font> 
</body>
</html>
<%
Dim Ano, Mes, Dia
Function InverteDataSiga(InverteData)
  Ano = Year(InverteData)
  Mes = Month(InverteData)
  Dia = Day(InverteData)
  If Len(Ano) > 4 Then
    Ano = "20" & Ano
  End If		
  If Len(Mes) =1 Then 
    Mes = "0" & Mes
  End If
  If Len(Dia) = 1 Then
    Dia = "0" & Dia
  End If
  InverteDataSiga = Ano & Mes & Dia
End Function
%>