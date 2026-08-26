<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <title>Pesquisa</title>
    <link href="../../../../../SisInfo/estilo.css" rel="stylesheet" type="text/css">
  </head>
  
  <!-- Formul�rio que abastace as campos da abertura do chamado -->
  <form name="abastece" method="post">
    <input name="Pedido" type="hidden" id="Pedido" value="">
  </form>

  <script language="javascript">

    function Atualiza() {
      //alert(document.selecao.selecionado.value)
      //window.self.opener.document.atualiza.pCodigo.value=document.selecao.selecionado.value;
	    //window.self.opener.document.form.Pedido.value=document.abastece.Pedido.value;
	    window.self.opener.document.getElementById('txtPedido').value=document.getElementById('Pedido').value;
      window.close();
    }

  </script>

  <body>
    <table width="495" height="125"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top" style="padding:10px">
<%
if request("param") <> 1 then
%>
          <script language="javascript">

            function Envia() {
	            if (document.form.chave.value=="") {
		            alert("Digite uma palavra chave para a pesquisa");
		            document.form.chave.focus();
		            return;
              }
              document.form.submit();
            }

          </script>

  	      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="400" valign="top">
                <form name="form" method="post" action="">
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="30">
                			  <select name="parametro" class="DescriForm" id="parametro" style="width:350px">
			                    <option value="1" selected>C&oacute;digo</option>
                          <option value="2">Refer&ecirc;ncia</option>
                          <option value="3">Pedido do Cliente</option>
                        </select>
                        <input name="param" type="hidden" id="param" value="1">
                        <input name="IdUsuario" type="hidden" id="IdUsuario" value="1">
                      </td>
                    </tr>
                    <tr>
                      <td height="30">
                        <input name="chave" type="text" class="DescriForm" id="chave" maxlength="30" style="width:334px">
                      </td>
                    </tr>
                  </table>
                </form>
              </td>
              <td valign="top" style="padding-left:10px">
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>
                      <span class="DescriForm">
                        <input name="Button" type="button" id="Submit" value="OK" onClick="Envia();" style="width:70px">
                      </span>
                    </td>
                  </tr>
                  <tr>
                    <td class="DescriForm" style="padding-top:5px">
                      <input name="Button" type="button" id="Submit" value="Fechar" onClick="window.close();" style="width:70px">
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
<%
else
  'Tabela com o resultado da pesquisa
  'Recuperando dados do formul�rio
  Parametro = Request("Parametro")

  Codigo  = iif( (Parametro = 1), Request("Chave"), 0 )
  Usuario = iif( (Parametro = 1), "", Ucase(Request("Chave")) )
   
  SQL = "EXEC dbo.SP_INPANET_HISTORICO_ALTERA_PEDIDO_BUSCA_PEDIDO " + replace(Parametro, "'", "''") + ",'" + replace(Codigo, "'", "''") + "','" + replace(Usuario, "'", "''") + "'"
  set Pesquisa = Conecta( "PIR-DADOSADV", SQL )

  if not isObject(Pesquisa) or (Pesquisa IS NOTHING) then
    response.end
  end if
%>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="400" valign="top">
                <div id="Conteudo" style="Z-INDEX: 10; LEFT: 0px; OVERFLOW: auto; WIDTH: 385px; CLIP: rect(0px 390px 115px 0px); TOP: 0px; HEIGHT: 110px">
                  <form name="selecao" method="post">
                    <table width="368px" border="0" cellspacing="1" cellpadding="1">
                      <tr>
                        <td>&nbsp;</td>
                        <td align="center" nowrap>Pedido</td>
                        <td align="center" nowrap>Referência</td>
                        <td align="center" nowrap>Produto</td>
                      </tr>
<% 
Dim a
a = 1
  while not Pesquisa.EOF
    'Rotina para alternar cor das colunas
    corLinha = iif( ((a Mod 2) > 0), "CabecalhoLinha1", "CabecalhoLinha1Muda" )
%>
                      <tr id="grid<%=Pesquisa("Codigo")%>" class="<%=corLinha %>">
                        <td align="center">
                          <input name="IdUsuario" type="hidden" id="IdUsuario" value="1">
                          <input name="Classe<%=Pesquisa("Codigo")%>" type="hidden" value="<%=Pesquisa("Codigo") %>">
                          <input name="selecionado" type="radio" value="<%=Pesquisa("Codigo") %>" onClick="document.abastece.Pedido.value='<%=trim(Pesquisa("Codigo"))%>';grid<%=Pesquisa("Codigo") %>.className='GridSelecionado';" onBlur="Retorna(this.value);">
                        </td>
                        <td align="center" nowrap>
                          <%=Pesquisa("Codigo")%>
                        </td>
                        <td nowrap>
                          <%=Pesquisa("Referencia")%>
                        </td>
                        <td nowrap>
                          <%=Pesquisa("Produto")%>
                        </td>
                      </tr>
<% 
    Pesquisa.MoveNext()
    a = a + 1
  wend
  Pesquisa.Close()
  set Pesquisa = Nothing
%>
                    </table>
                  </form>
                </div>
              </td>
              <td valign="top" style="padding-left:10px">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>
                      <span class="DescriForm">
                        <input name="Button" type="button" id="Button" value="OK" onClick="Atualiza();" style="width:70px">
                      </span>
                    </td>
                  </tr>
                  <tr>
                    <td class="DescriForm" style="padding-top:5px; padding-bottom:5px">
                      <input name="Button" type="button" id="Submit" value="Voltar" onClick="history.go(-1);" style="width:70px">
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <span class="DescriForm">
                        <input name="Button" type="button" id="Submit" value="Fechar" onClick="window.close();" style="width:70px">
                      </span>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
<%
end if
%>	
	      </td>
      </tr>
    </table>
  </body>
</html>