<!--#include file="../../../../cnn/cnn.asp" -->
<!--#include file="../../../estatisticas.asp" --> 
<!--#include file="../../../../includes/variaveis.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <title>Consulta a Clientes</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <!-- script para os layers dos bal�es de informa��o. -->
    <script language="Javascript">
    <!--
    function achaObjeto(n, d) { 
    var p,i,x; if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
    if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
    for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=achaObjeto(n,d.layers[i].document);
    if(!x && document.getElementById) x=document.getElementById(n); return x;
    }

    function SH_Layers() { //v3.0
    var i,p,v,obj,args=SH_Layers.arguments;
    for (i=0; i<(args.length-2); i+=3) if ((obj=achaObjeto(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
    }
    //-->
    </script>
    <!-- Fim do script que gera as camadas de informa��es -->
    <link href="../../Pedidos/Consultas/style.css" rel="stylesheet" type="text/css">

    <style type="text/css">
    /*
    .barra_menu {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #FFFFFF; text-decoration: none}
    .barra_menu:hover {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #99CCFF; text-decoration: none}
    */
    </style>
  </head>

  <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <table width="100%" height="60" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td width="38">&nbsp;</td>
        <td width="647">
          <font color="#3366FF" size="4" face="tahoma">
          <strong>Consulta a Clientes</strong></font> 
          <div id="cliente1" style=" padding=1px; position:absolute; width:520; height:185; z-index:1; left: 189px; top: 95px; background-color: #000000; layer-background-color: #3869E8; overflow: auto; border: 1px #000000; visibility: hidden;"> 
            <script language="Javascript1.3">
              function verifica2() {
                var themessage = "Campo de preechimento obrigat�rio:";
                /* Validando codigo do cliente em branco */
                if (document.busca.buscar.value=="") {
                  themessage = themessage + "\nDigite uma palavra chave. ";
                }
                if (themessage == "Campo de preechimento obrigat�rio:") {
                  document.busca.submit();
                  document.busca.buscar.value='';
                } else {
                  alert(themessage);
                  return false;
                }
              }
            </script>
            <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td valign="top" bgcolor="#BFDFFF">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center" bgcolor="#3366FF" class="barra_menu">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td width="96%" align="center" class="barra_menu">BUSCA DE CLIENTES </td>
                            <td width="4%" align="right">
                              <img src="../../Pedidos/images/close.gif" width="16" height="14" class="cursor"  onMouseDown="SH_Layers('cliente1','','hide')">
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#0099FF" style="padding=2px">
                    <form name="busca" method="post" action="../../Pedidos/Consultas/busca_cli.asp" target="ibusca">
                      <tr> 
                        <td width="30%" height="22" class="texto_simples">
                          Digite a palavra chave <input name="param" type="hidden" id="param4" value="1"> :
                        </td>
                        <td width="14%"> 
                          <input name="buscar" type="text" class="textfield3" id="buscar" size="10">
                        </td>
                        <td width="59%"> 
                          <input name="button" type='button' style='font:09pt Arial, Helvetica, sans-serif; border:1 solid #000000; cursor:hand; background:#0099FF; color: #FFFFFF; ' value="Buscar" onClick="verifica2();"> 
                        </td>
                      </tr>
                    </form>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"> 
                        <iframe name="ibusca" marginWidth=0 marginHeight=0 src="../../Pedidos/Consultas/busca_cli.asp" frameBorder="0" width="518px" scrolling="auto" height="140px" bordercolor="#000000"> </iframe> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </div>
        </td>
      </tr>
    </table>
<%
'VERIFICANDO O PAR�METRO PARA EXECUTAR A CONSULTA
if request("PARAMMAIN") <> 1 then

	'VERIFICANDO QUAL O TIPO DE ACESSO PARA MONTAR A BARRA DE FUN��ES
	select case CATCODIGO
	case 1 'Funcionario -----------------------------------------------------------------------------
%>
    <script language="Javascript">
      function verify() {
        var themessage = "Campo de preechimento obrigat�rio:\n........................................................";
        /* Validando c�digo do cliente em branco */
        if (document.form.codcli.value=="") {
          themessage = themessage + "\nC�digo do Cliente. ";
        }
        /* Validando 1� Intervalo de data em branco */
        if (document.form.lojacli.value=="") {
          themessage = themessage + "\nC�digo da Loja. ";
        }
        if (themessage == "Campo de preechimento obrigat�rio:\n........................................................") {
          document.form.submit();
        } else {
          alert(themessage);
          return false;
        }
      }
    </script>
  <%
    response.write "<table border=1 align=center cellpadding=0 cellspacing=1 bordercolor=#CCCCCC bgcolor=#3869E8 style='padding=2px'>"
    response.write "<form method=post action='conclientes.asp' name=form>"
    response.write "<tr>"
    response.write "<td height=43><font face=Tahoma,verdana,arial size=1 color=#FFFFFF><b>&nbsp;CLIENTE:</font></b></td>"
    response.write "<td><input name=codcli type=text class=textfield id=codcli size=7></td>"
    response.write "<td><font face=Tahoma,verdana,arial size=1 color=#FFFFFF><b>&nbsp;LOJA:</font></b></td>"
    response.write "<td><input name=lojacli type=text class=textfield id=lojacli size='7'></td>"
    response.write "<td><img src=images/lupa.jpg alt='Pesquisar o Cliente por Nome' width=24 height=23 border=0 class=cursor onMouseDown=SH_Layers('cliente1','','show');></td>"
    response.write "</tr>"
    response.write "<tr align=center>" 
    response.write "<td height=43 colspan=5>"
    response.write "<input name=Button type=Button class=textfield onClick='verify();' value='Pesquisar'>"
    response.write "<input name='parammain' type='hidden' id='parammain' value='1'>"
    response.write "</td>"
    response.write "</tr>"
    response.write "</form>"
    response.write "</table>"
    'Colocando o Foco no campo LojaCli%>
    <script language="Javascript">
    // Colocando o foco no campo CodCli
    document.form.codcli.focus();
    </script>
<%
	case 2 'Cliente ---------------------------------------------------------------------------------
%>
    <script language="Javascript">
      function verify() {
        var themessage = "Campo de preechimento obrigat�rio:\n........................................................";
        /* Validando c�digo do cliente em branco */
        if (document.form.codcli.value=="") {
          themessage = themessage + "\nC�digo do Cliente. ";
        }
        /* Validando 1� Intervalo de data em branco */
        if (document.form.lojacli.value=="") {
          themessage = themessage + "\nC�digo da Loja. ";
        }
        if (themessage == "Campo de preechimento obrigat�rio:\n........................................................") {
          document.form.submit();
        } else {
          alert(themessage);
          return false;
        }
      }
    </script>
<%
    response.write "<table border=1 align=center cellpadding=0 cellspacing=1 bordercolor=#CCCCCC bgcolor=#3869E8 style='padding=2px'>"
    response.write "<form method=post action='conclientes.asp' name=form>"
    response.write "<tr>"
    response.write "<td height=43><font face=Tahoma,verdana,arial size=1 color=#FFFFFF><b>&nbsp;CLIENTE:</font></b></td>"
    response.write "<td><input name=codcli type=text class=textfield id=codcli size=7 value="& CodUser & "></td>"
    response.write "<td><font face=Tahoma,verdana,arial size=1 color=#FFFFFF><b>&nbsp;LOJA:</font></b></td>"
    response.write "<td><input name=lojacli type=text class=textfield id=lojacli size='7'></td>"
    response.write "<td><img src=images/lupa.jpg alt='Pesquisar o Cliente por Nome' width=24 height=23 border=0 class=cursor onMouseDown=SH_Layers('cliente1','','show');></td>"
    response.write "</tr>"
    response.write "<tr align=center>" 
    response.write "<td height=43 colspan=5>"
    response.write "<input name=Button type=Button class=textfield onClick='verify();' value='Pesquisar'>"
    response.write "<input name='parammain' type='hidden' id='parammain' value='1'>"
    response.write "</td>"
    response.write "</tr>"
    response.write "</form>"
    response.write "</table>"
    'Colocando o Foco no campo LojaCli%>
    <script language="Javascript">
      // Colocando o foco no campo CodCli
      document.form.lojacli.focus();
    </script>
<%
	case 4 'Reprresentante --------------------------------------------------------------------------
%>
		<script language="Javascript">
			function verify() {
  			var themessage = "Campo de preechimento obrigatório:";
  			if (themessage == "Campo de preechimento obrigatório:") {
	    		document.form.submit();
			  }	else {
			    alert(themessage);
			    return false;
   			}
			}
			</script>	
<%
    response.write "<table border=1 cellspacing=1 cellpadding=0 bgcolor=#3869E8 bordercolor=#CCCCCC style=padding=2px align=center>"
    response.write "<form method=post action=conclientes.asp name=form>"
    response.write "<tr align=center>" 
    response.write "<td colspan=6><font color=#FFFFFF size=1 face=Verdana,Arial,Helvetica,sans-serif><strong>Selecione a melhor configura&ccedil;&atilde;o</strong></font></td>"
    response.write "</tr>"
    response.write "<tr bgcolor=#CCCCCC>" 
    response.write "<td><input type=radio name=pai value=1 checked></td>"
    response.write "<td colspan=5><font size=1 face=Verdana,Arial,Helvetica,sans-serif>C&oacute;digo do cliente</font></td>"
    response.write "</tr>"
    response.write "<tr bgcolor=#CCCCCC>"
    response.write "<td><input type=radio name=pai value=2></td>"
    response.write "<td colspan=5><font size=1 face=Verdana,Arial,Helvetica,sans-serif>Ordem Alfab&eacute;tica</font></td>"
    response.write "</tr>"
    response.write "<tr bgcolor=#CCCCCC>" 
    response.write "<td>"
    response.write "<input type=radio name=pai value=3></td>"
    response.write "<td colspan=5><font color=#000000 size=1 face=Verdana,Arial,Helvetica,sans-serif>Maior limite de cr&eacute;dito primeiro</font></td>"
    response.write "</tr>"
    response.write "<tr bgcolor=#CCCCCC>" 
    response.write "<td>"
    response.write "<input type=radio name=pai value=4></td>"
    response.write "<td colspan=5><font color=#000000 size=1 face=Verdana,Arial,Helvetica,sans-serif>Bloqueados e sem limite de cr�dito</font></td>"
    response.write "</tr>"
    response.write "</tr>"
    response.write "<tr bgcolor=#CCCCCC>"
    response.write "<td><select name=pagesize size=1 class=textfield>"
    response.write "<option value=5>5</option>"
    response.write "<option value=10>10</option>"
    response.write "<option value=15>15</option>"
    response.write "<option value=20 selected>20</option>"
    response.write "<option value=25>25</option>"
    response.write "<option value=30>30</option>"
    response.write "</select></td>"
    response.write "<td colspan=5><font size=1 face=Verdana,Arial,Helvetica,sans-serif>Clientes por p&aacute;gina</font></td>"
    response.write "</tr>"
    response.write "<tr>"
    response.write "<td colspan=6 align=center>"
    response.write "<input name=Button type=Button class=textfield onClick=verify(); value=Pesquisar>"
    response.write "<input name=parammain type=hidden id=parammain value=1></td>"
    response.write "</tr>"
    response.write "</form>"
    response.write "</table>"	
	end select

else

	'DECLARANDO VARI�VEIS DESTA P�GINA
	DIM CODREPRE, CODCLI, LOJACLI, ORDEM, RSCLI, VARSQLCLI, RISCO, LIMITECREDITO, ASP_DB, PAGINA
	DIM TAMANHOPAGINA, SQL, CONTADOR, MAXIMOPAGINA, MAXIMOREGISTROS, REGISTROS, CONSULTA
	DIM PAD, scriptNAME, COUNTER, REF
	
	'RECEBENDO OS PAR�METROS DA OUTRA P�GINA
	select case CATCODIGO
	case 1 'Funcionario -----------------------------------------------------------------------------

    CODREPRE = " "
    CODCLI = request("codcli")
    LOJACLI = request("LOJACLI")
    ORDEM = " "
    
    SQL = "EXECUTE SP_LISTA_CLIENTES_DETALHADO [" & CODREPRE & "],[" & CODCLI & "],[" & LOJACLI & "],[" & CATCODIGO & "],[" & ORDEM & "]"
    Dados.Open(SQL)

    if Dados.EOF then
      response.write "</br>"
      response.write "<table width=450 border=0 cellpadding=0 cellspacing=0 bgcolor=#FF0000 align=center>"
      response.write "<tr>"
      response.write "<td align=center valign=top class=texto_atencao style='padding=10px'>ATEN&Ccedil;&Atilde;O<br>"
      response.write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
      response.write "<tr>"
      response.write "<td align=center bgcolor=#FFFFFF style='padding=10px'>"
      response.write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
      response.write "<tr>"
      response.write "<td style=padding=2px><img src=../../../images/alerta_val_user.gif width=31 height=31 align=middle></td>"
      response.write "<td class=texto_simples2>Este cliente n�o foi encontrado em nossa base de dados<br>Clique <a href=# onMouseDown='history.go(-1);'>aqui</a> para voltar e tentar novamente.</td>"
      response.write "</tr>"
      response.write "</table>"                   
      response.write "</td>"
      response.write "</tr>"
      response.write "</table> </td>"
      response.write "</tr>"
      response.write "</table>"
      response.write "</br>"
      Dados.Close()
      Response.End
    end if

		'Passando Vari�veis para a outra p�gina
		Session("CodCli") = Dados("CodCli")
		Session("LojaCli") = Dados("LojaCLi")
		Session("CliNome") = Dados("CliNome")
		Session("CliNomReduz") = Dados("CliNomReduz")
		Session("TPFrete") = Dados("TPFrete")
		Session("Endereco") = Dados("Endereco")
		Session("Cidade") = Dados("Cidade")
		Session("UF") = Dados("UF")
		Session("Bairro") = Dados("Bairro")
		Session("Cep") = Dados("Cep")
		Session("Tel") = Dados("Tel")
		Session("Fax") = Dados("Fax")
		Session("Contato") = Dados("Contato")
		Session("Cnpj") = Dados("Cnpj")
		Session("CodRepre") = Dados("CodRepre")
		Session("Risco") = Dados("Risco")
		Session("LimiteCredito") = Dados("LimiteCredito")
		Session("Email") = Dados("email")
		Session("Site") = Dados("Site")
		Session("Pais") = Dados("Pais")
		Session("ELaudo") = Dados("ELaudo")
		Session("NomRepre") = Dados("NomRepre")
		Session("NomReduzRepre") = Dados("NomReduzRepre")
		Session("EndRepre") = Dados("EndRepre")
		Session("CidadeRepre") = Dados("CidadeRepre")
		Session("UfRepre") = Dados("UfRepre")
		Session("CnpjRepre") = Dados("CnpjRepre")
		Session("TelRepre") = Dados("TelRepre")
		Session("EmailRepre") = Dados("EmailRepre")
		Session("SiteRepre") = Dados("SiteRepre")

    'TABELA PRINCIPAL
    response.write "<table width=700 border=1 cellspacing=0 cellpadding=2 align=center bordercolor=#3869E8>"
    response.write "<tr>"
    response.write "<td>"

    'MONTANDO O CABE�ALHO COM O NOME DO REPRESENTANTE
    response.write "<TABLE BORDER='0' width=100% >"
    response.write "<TR>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>&nbsp;&nbsp;&nbsp;"& Dados("CODREPRE") & "</B></TD>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>"& Dados("NOMREPRE") & "</B></TD>"
    response.write "</TR>"
    response.write "</TABLE>"

    response.write "<TABLE BORDER='1' cellspacing='2' cellpadding='2' bordercolor='#FFCC99' width=100% >"
    response.write "<TR>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>CodCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LojaCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Nome</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Cidade</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Estado</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Risco</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LimiteCr�dito</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Detalhes</DIV></TD>"
    response.write "</TR>"

    response.write "<TR>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("CodCli") & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("LojaCli") & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CliNome") & "</TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CIdade") & "</TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("UF") & "</DIV></TD>"
    'Verificando o Risco para dar o nome
    if Dados("Risco") = "E" Then
      Risco = "Bloqueado"
    Elseif Dados("Risco") = "P" Then
      Risco = "Bloqueado p/ Novos Pedidos"
    Else
      Risco = "Liberado"
    End If
      response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Risco & "</DIV></TD>"
    'Verificando o Limite de Cr�dito para dar o nome
    if Dados("LimiteCredito") > 0 Then
      LimiteCredito = FormatNumber(Dados("LimiteCredito"),2)
    Else
      LimiteCredito = "Sem Limite"
    End If

    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & LimiteCredito & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P><a href=conclidet.asp?PCodCli="&TRIM(Dados("CodCli"))&"&PLojaCli="&TRIM(Dados("LojaCLi"))&">+ detalhes</a></td>"
    response.write "</TR>"

    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
    Dados.Close()
		
	case 2 'Cliente ---------------------------------------------------------------------------------

    CODREPRE = " "
    CODCLI = CODUSER
    LOJACLI = request("LojaCli")
    ORDEM = " "

    SQL = "EXECUTE SP_LISTA_CLIENTES_DETALHADO [" & CODREPRE & "],[" & CODCLI & "],[" & LOJACLI & "],[" & CATCODIGO & "],[" & ORDEM & "]"
    Dados.Open(SQL)

    'Passando Vari�veis para a outra p�gina
    Session("CodCli") = Dados("CodCli")
    Session("LojaCli") = Dados("LojaCLi")
    Session("CliNome") = Dados("CliNome")
    Session("CliNomReduz") = Dados("CliNomReduz")
    Session("TPFrete") = Dados("TPFrete")
    Session("Endereco") = Dados("Endereco")
    Session("Cidade") = Dados("Cidade")
    Session("UF") = Dados("UF")
    Session("Bairro") = Dados("Bairro")
    Session("Cep") = Dados("Cep")
    Session("Tel") = Dados("Tel")
    Session("Fax") = Dados("Fax")
    Session("Contato") = Dados("Contato")
    Session("Cnpj") = Dados("Cnpj")
    Session("CodRepre") = Dados("CodRepre")
    Session("Risco") = Dados("Risco")
    Session("LimiteCredito") = Dados("LimiteCredito")
    Session("Email") = Dados("email")
    Session("Site") = Dados("Site")
    Session("Pais") = Dados("Pais")&"'"
    Session("ELaudo") = Dados("ELaudo")
    Session("NomRepre") = Dados("NomRepre")
    Session("NomReduzRepre") = Dados("NomReduzRepre")
    Session("EndRepre") = Dados("EndRepre")
    Session("CidadeRepre") = Dados("CidadeRepre")			
    Session("UfRepre") = Dados("UfRepre")
    Session("CnpjRepre") = Dados("CnpjRepre")
    Session("TelRepre") = Dados("TelRepre")
    Session("EmailRepre") = Dados("EmailRepre")
    Session("SiteRepre") = Dados("SiteRepre")
    
    'TABELA PRINCIPAL
    response.write "<table width=700 border=1 cellspacing=0 cellpadding=2 align=center bordercolor=#3869E8>"
    response.write "<tr>"
    response.write "<td>"

    'MONTANDO O CABE�ALHO COM O NOME DO REPRESENTANTE
    response.write "<TABLE BORDER='0' width=100% >"
    response.write "<TR>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>&nbsp;&nbsp;&nbsp;"& Dados("CODREPRE") & "</B></TD>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>"& Dados("NOMREPRE") & "</B></TD>"
    response.write "</TR>"
    response.write "</TABLE>"

    response.write "<TABLE BORDER='1' cellspacing='2' cellpadding='2' bordercolor='#FFCC99' width=100% >"
    response.write "<TR>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>CodCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LojaCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Nome</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Cidade</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Estado</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Risco</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LimiteCr�dito</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Detalhes</DIV></TD>"
    response.write "</TR>"

    response.write "<TR>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("CodCli") & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("LojaCli") & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CliNome") & "</TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CIdade") & "</TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("UF") & "</DIV></TD>"

    'Verificando o Risco para dar o nome
    if Dados("Risco") = "E" then
      Risco = "Bloqueado"
    elseif Dados("Risco") = "P" then
      Risco = "Bloqueado p/ Novos Pedidos"
    else
      Risco = "Liberado"
    end if

    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Risco & "</DIV></TD>"

    'Verificando o Limite de Cr�dito para dar o nome
    if Dados("LimiteCredito") > 0 then
      LimiteCredito = FormatNumber(Dados("LimiteCredito"),2)
    else
      LimiteCredito = "Sem Limite"
    end if

    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & LimiteCredito & "</DIV></TD>"
    response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P><a href=conclidet.asp?PCodCli="&Dados("CodCli")&"&PLojaCli="&Dados("LojaCLi")&">+ Detalhes</a></td>"
    response.write "</TR>"

    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
    Dados.Close()

	case 4 'Representante ---------------------------------------------------------------------------

    CODREPRE = CODUSER
    CODCLI = " "
    LOJACLI = " "
    ORDEM = request("PAI")

    SQL = "EXECUTE SP_LISTA_CLIENTES_DETALHADO [" & CODREPRE & "],[" & CODCLI & "],[" & LOJACLI & "],[" & CATCODIGO & "],[" & ORDEM & "]"

    Pagina = request("QualPagina")
    if Pagina = "" then
      Pagina = 1
    end if

    Tamanhopagina = request("pagesize")
    if Tamanhopagina = "" then
      Tamanhopagina = 20
    end if

    SQL = request("sqlquery")
    if Sql = "" then
      Sql=VarSQLCLI
    end if

    Dados.Open(SQL)

    if Dados.Eof then
      response.write "nada"
      Response.end
    End If

    if not Dados.eof then
      contador = Dados.recordcount
      Dados.movefirst()
      Dados.PageSize = tamanhopagina
      maximopagina = cint(Dados.PageCount)
      maximoregistros = cint(Dados.PageSize)
      Dados.AbsolutePage = pagina
      registros = 0
      consulta="ok"
    end if

    'TABELA PRINCIPAL
    response.write "<table width=700 border=1 cellspacing=0 cellpadding=2 align=center bordercolor=#3869E8>"
    response.write "<tr>"
    response.write "<td>"

    'TABELA CONTENDO A QUANTIDADE DE P�GINAS E O N�MERO DE REGISTROS ENCONTRADOS			
    response.write "<TABLE BORDER='0' width=100% >"
    response.write "<TR bgcolor=#3869E8>"
    response.write "<TD height=23px><font face=verdana, arial size=1 color=#cccccc><P><B>&nbsp;&nbsp;&nbsp;P�gina: "& Pagina & " de " & Maximopagina & "</B></TD>"
    response.write "<TD><font face=verdana, arial size=1 color=#cccccc><P><B>&nbsp;&nbsp;&nbsp;Total de registros encontrados: " & contador & "</B></TD>"
    response.write "</TR>"
    response.write "</TABLE>"

    'MONTANDO O CABE�ALHO COM O NOME DO REPRESENTANTE
    response.write "<TABLE BORDER='0' width=100% >"
    response.write "<TR>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>&nbsp;&nbsp;&nbsp;"& Dados("CODREPRE") & "</B></TD>"
    response.write "<TD><font face=tahoma, arial size=4 color=#666666><P><B>"& Dados("NOMREPRE") & "</B></TD>"
    response.write "</TR>"
    response.write "</TABLE>"

    response.write "<TABLE BORDER='1' cellspacing='2' cellpadding='2' bordercolor='#FFCC99' width=100% >"
    response.write "<TR>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>CodCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LojaCli</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Nome</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Cidade</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Estado</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Risco</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>LimiteCr�dito</DIV></TD>"
    response.write "<TD bgcolor='#BFDFFF'><font face=tahoma, arial size=1><div align=center><B>Detalhes</DIV></TD>"
    response.write "</TR>"

    if Consulta = "ok" then

      Do Until Dados.EOF or Registros >= Maximoregistros

        response.write "<TR>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("CodCli") & "</DIV></TD>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("LojaCli") & "</DIV></TD>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CliNome") & "</TD>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P>" & Dados("CIdade") & "</TD>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Dados("UF") & "</DIV></TD>"
        'Verificando o Risco para dar o nome
        if Dados("Risco") = "E" Then
          Risco = "Bloqueado"
        else
          Risco = "Liberado"
        end if

        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & Risco & "</DIV></TD>"

        'Verificando o Limite de Cr�dito para dar o nome
        if Dados("LimiteCredito") > 0 then
          LimiteCredito = FormatNumber(Dados("LimiteCredito"),2)
        else
          LimiteCredito = "Sem Limite"
        end if
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><div align=center><P>" & LimiteCredito & "</DIV></TD>"
        response.write "<TD bgcolor='#FFFFFF'><font face=tahoma, arial size=1><P><a href=conclidet.asp?PCodCli="&TRIM(Dados("CodCli"))&"&PLojaCli="&TRIM(Dados("LojaCLi"))&">+ Detalhes</a></td>"
        response.write "</TR>"
        Dados.MoveNext
        REGISTROS = REGISTROS + 1 
      Loop

    end if

    'Criando os links das p�ginas
    'escrevendo a instru��o de pr�xima p�gina
    response.write "<table width=100% cellspacing='0' cellpadding='0' border='0'>"
    response.write "<tr bgcolor=#cccccc>"
    response.write "<td height=25px>"
    response.write "<font face = 'verdana', size = '1' color=#333333>&nbsp;&nbsp;&nbsp;Clique no n�mero da p�gina que deseja visualizar: " & "</font>"
    
    pad = "0"

    'armazena o nome da p�gina para criar o link para ela mesmo

    scriptname="conclientes.asp"
    'inicia o contador com 1 at� o total de p�ginas
    for counter = 1 to maximopagina

      if counter >= tamanhopagina then
        pad = ""
      end if

      'escreve os links para as p�ginas
      ref="<a href="&scriptname
      ref=ref&"?qualpagina="&counter
      ref=ref&"&parammain="&request("parammain")
      ref=ref&"&pagesize="&tamanhopagina
      ref=ref&"&sqlquery="&server.urlencode(sql)
      ref=ref&">"&pad&counter&"</a>"
      response.write "<font face = 'verdana', size = '1'>" &  ref & " " & "</font>"
      'se o resto da divis�o do contador por 20 for igual a zero vai para a pr�xima linha
      if counter mod tamanhopagina = 0 then
        response.write "<br>"
      end if

    Next

    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
    
    'CASO N�O HAJA REGISTRO, ATRIBUI O VALOR DE REGISTROS A 0 PARA MOSTRAR NA TELA.
    if contador = "" then
      contador = 0
    end if
    
    if not Dados.EOF then
      'Passando Variaveis para a outra pagina
      Session("CodCli")         = Dados("CodCli")
      Session("LojaCli")        = Dados("LojaCLi")
      Session("CliNome")        = Dados("CliNome")
      Session("CliNomReduz")    = Dados("CliNomReduz")
      Session("TPFrete")        = Dados("TPFrete")
      Session("Endereco")       = Dados("Endereco")
      Session("Cidade")         = Dados("Cidade")
      Session("UF")             = Dados("UF")
      Session("Bairro")         = Dados("Bairro")
      Session("Cep")            = Dados("Cep")
      Session("Tel")            = Dados("Tel")
      Session("Fax")            = Dados("Fax")
      Session("Contato")        = Dados("Contato")
      Session("Cnpj")           = Dados("Cnpj")
      Session("CodRepre")       = Dados("CodRepre")
      Session("Risco")          = Dados("Risco")
      Session("LimiteCredito")  = Dados("LimiteCredito")
      Session("Email")          = Dados("email")
      Session("Site")           = Dados("Site")
      Session("Pais")           = Dados("Pais")
      Session("ELaudo")         = Dados("ELaudo")
      Session("NomRepre")       = Dados("NomRepre")
      Session("NomReduzRepre")  = Dados("NomReduzRepre")
      Session("EndRepre")       = Dados("EndRepre")
      Session("CidadeRepre")    = Dados("CidadeRepre")
      Session("UfRepre")        = Dados("UfRepre")
      Session("CnpjRepre")      = Dados("CnpjRepre")
      Session("TelRepre")       = Dados("TelRepre")
      Session("EmailRepre")     = Dados("EmailRepre")
      Session("SiteRepre")      = Dados("SiteRepre")
    end if

    response.write "</td>"
    response.write "</tr>"
    response.write "</table>"
    Dados.Close
	end select

	'MONTANDO O LINK PARA MAIS DETALHES PASSANDO OS DADOS DESTA P�GINA
end if

'PRINTANDO UMA LINHA EM BRANCO COM TABELA
response.write "<table width=100% cellspacing='0' cellpadding='0' border='0'>"
response.write "<tr>"
response.write "<td>&nbsp;</td>"
response.write "</tr>"
response.write "</table>"

%> 
</body>
</html>
