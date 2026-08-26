<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../../Funcoes/FuncoesAsp.asp" -->    
<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1046 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta charset="utf-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
        <title>SWnet - Solicitação de Alteração de Pedido</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
    
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" />
        <link 
            rel="stylesheet" 
            href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" 
            integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" 
            crossorigin="anonymous" 
        />

        <style type="text/css"> 
            .alert-vertical-center {
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                width: 50%; /* Largura opcional */
            }
        </style>
<!--        <link href="../consultas/estilo.css" rel="stylesheet" type="text/css"> -->
    </head>

    <script>
        /*----------------------------------------------------------------------------
        Formata��o para qualquer mascara
        -----------------------------------------------------------------------------*/
        function formatar(src, mask) {
            var i = src.value.length;
            var saida = mask.substring(0, 1);
            var texto = mask.substring(i)
            if (texto.substring(0, 1) != saida) {
                src.value += texto.substring(0, 1);
            }
        }
    </script>

    <script type="text/javascript">
        function CifFobPrecoUnitario(TipoFrete) {
            if ((TipoFrete == 'C') && ("<%=TpProdutoTipo %>" == "PA")) {
                unitario(document.getElementById('NovoFreteUnitario').value, document.getElementById('NovoFreteUnitario'))
                if (unitario(document.getElementById('NovoFreteUnitario').value, document.getElementById('NovoFreteUnitario')) <= 0) {
                    alert('O valor do frete deve ser maior que zero!');
                    return false;
                }
            }
            return true;
        }
        function CifFob() {
            if ((document.getElementById('NovoTPFrete').value == 'C') && ("<%=TpProdutoTipo %>" == "PA")) {
                var valor = document.getElementById('valorfrete').value;
                var nvalor = parseFloat(valor.replace(".", "").replace(",", "."));

                if (isNaN(nvalor)) {
                    alert('preencha o valor do frete!');
                    document.getElementById('valorfrete').value = "0";
                    return false;
                }

                if (nvalor <= 0) {
                    alert('O valor do frete deve ser maior que zero!');
                    document.getElementById('valorfrete').value = "0";
                    return false;
                }
            }
            return true;
        }
        function unitario(valor, campo) {
            var val = valor.replace(",", '.')
            val = parseFloat(val)
            if (isNaN(val)) {
                alert('Este campo é exclusivamente numérico')
                campo.value = "0"
            } else {
                if (valor != "") {
                    campo.value = val
                } else {
                    campo.value = "0"
                }
            }
        }
    </script>

    <script type="text/javascript">
        function validaData(campo, status) {
            if (campo.value != "") {
                var erro = 0;
                var Hoje = new Date();
                var HojeDia = Hoje.getDate();
                var HojeMes = Hoje.getMonth() + 1;
                var HojeAno = Hoje.getFullYear();
                if (HojeDia.toString().length == 1)
                    HojeDia = "0" + dia;
                if (HojeMes.toString().length == 1)
                    HojeMes = "0" + HojeMes;
                var Data = campo.value.split("/");
                var Dia, Mes, Ano;
                if (Data.length == 3) {
                    dia = Data[0];
                    mes = Data[1];
                    ano = Data[2];
                    resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano >= HojeAno && ano >= 1900));

                    if (!resultado) {
                        alert("Data inválida.");
                        campo.focus();
                        return false;
                    }
                    resultado2 = ((mes == 4) || (mes == 6) || (mes == 9) || (mes == 11));

                    if (resultado2) {
                        if (dia > 30) {
                            alert("Data inválida.");
                            campo.focus();
                            return false;
                        }
                    }
                    resultado3 = ((mes == 1) || (mes == 3) || (mes == 5) || (mes == 7) || (mes == 8) || (mes == 10) || (mes == 12));

                    if (resultado3) {
                        if (dia > 31) {
                            alert("Data inválida.");
                            campo.focus();
                            return false;
                        }
                    }

                    resultado4 = ((ano % 4 == 0) || (ano % 100 == 0) || (ano % 400 == 0)); //verifica se o ano � bissexto
                    if (resultado4) {
                        if ((mes == 2) && (dia > 29)) {
                            alert("Data inválida.");
                            campo.focus();
                            return false;
                        }
                    }
                    if (!resultado4) {
                        if ((mes == 2) && (dia > 28)) {
                            alert("Data inválida.");
                            campo.focus();
                            return false;
                        }
                    }
                }
                else {
                    alert("Data inválida.");
                    campo.focus();
                    return false;
                }
                DataHoje = HojeAno + "-" + HojeMes + "-" + HojeDia
                Data = new Date(DataHoje);

                DataParametro = ano + "-" + mes + "-" + dia;
                ParametroData = new Date(DataParametro)

                if ((Data > ParametroData) && (status != 'L')) {
                    alert("Data inválida. A data deve ser maior ou igual à data atual!");
                    campo.value = ''
                    campo.focus();

                    return false;
                }
                return true;
            }
        }
    </script>

    <script type="text/javascript">
        function verifica() {
            var themessage = "Campo de preechimento obrigat�rio:\n........................................................";
            if ('<%=TipoOperacao%>' == '1') {
                if (!CifFob()) {
                    return false
                }
            }
            if ('<%=TipoOperacao%>' == '14') {
                if (!CifFobPrecoUnitario(TipoFrete)) {
                    return false
                }
            }
            //alert if fields are empty and cancel form submit
            if (themessage == "Campo de preechimento obrigat�rio:\n........................................................") {
                document.form.submit();
            }
            else {
                alert(themessage);
                return false;
            }
        }

    </script>

    <script type="text/javascript" src="../../../../../Funcoes/ValDataJS.js"></script>

    <script type="text/javascript">
        function showDiv(id) {
            //document.all.cResponsavel.style.visibility = 'hidden'; 
            document.all[id].style.visibility = 'visible';
            document.all[id].focus();
        }

        function hideDiv(id) {
            document.all[id].style.visibility = 'hidden';
        }

        function ValorFrete(tipo) {
            if ((tipo == "F") && ("<%=TpProdutoTipo %>" == "PA")) {
                document.getElementById("valorfrete").value = '0';
                document.getElementById("valorfrete").disabled = true;
            }

            if ((tipo == "C") && ("<%=TpProdutoTipo %>" == "PA")) {
                alert('Insira o valor do frete');
                document.getElementById("valorfrete").disabled = false;
            }
        }
    </script>
<%
Pedido = trim(request("Pedido"))
item = "01"
idSolicitacao = trim(request("Id"))
TipoOperacao = trim(request("TPOperacoes"))

SQL = "" &_
    "SELECT " &_
        "C6_PRODUTO " &_
    "FROM SC6010 " &_
    "WHERE C6_NUM = '" & Pedido & "' " &_
        "AND D_E_L_E_T_ <> '*'"

SET TPProduto = Conecta("PIR-DADOSADV", SQL)

if not isObject(TPProduto) or (TPProduto IS NOTHING) then
    response.end
end if

if not TPProduto.EOF then
    select case Cint(left(TPProduto("C6_PRODUTO"),3))
    case 700 : TpProdutoTipo = "PA"
    case 720 : TpProdutoTipo= "CH"
    case else : TpProdutoTipo = "CX"
    end select
end if

'Recuperando os dados do formul�rio da p�gina anterior
'Consulta a lista de tipos de opera��es de altera��o do pedido

SQL = "SELECT Operacao FROM HELPDESK..InHistoricoTipoOperacao WHERE Id = " & TipoOperacao 

SET Operacoes = Conecta("PIR-HELPDESK", SQL)

if not isobject(Operacoes) or (Operacoes IS NOTHING) then
    response.end
end if

Operacoes_numRows = 0

'Consulta a lista de tipos de opera��es de altera��o do pedido

SQL = "{call dbo.SP_INPANET_HISTORICO_ALTERA_PEDIDO_BUSCA_DADOS_PEDIDO('" & Pedido & "', " & TipoOperacao & ")}"

SET TPOperacoes = Conecta("PIR-DADOSADV", SQL)

if not isobject(TPOperacoes) or (TPOperacoes IS NOTHING) then
    response.end
end if

'verificando se o pedido foi encontrado no db, pode estar liquidado ou cancelado
if TpProdutoTipo = "PA" then

    SQL = "" &_
        "SELECT " &_
            "COUNT(*) AS PEDIDO " &_
        "FROM SC6010 " &_
        "WHERE C6_PRODUTO >= '700000000' " &_
          "AND C6_PRODUTO <= '704999999' " &_
          "AND C6_NUM = '" & Pedido & "'"

    SET TPPedido = Conecta("PIR-DADOSADV", SQL)

    if not isobject(TPPedido) or (TPPedido IS NOTHING) then
        response.end
    end if

    if TPPedido.EOF or TPPedido("PEDIDO") < 1 then
        Aviso( "Alteração permitida APENAS para pedidos de PAPEL." )
        response.end
    end if

end if

if TPOperacoes.EOF then
    msg = "" &_
            "Este pedido está liquidado,<br><br>" &_
            "... ou foi cancelado,<br>" &_
            "... ou o código do produto do pedido não se encontra entre 700 e 769999999,<br>" &_
            "... ou a unidade de medida do pedido não é PA ( Produto Acabado ),<br>" &_
            "... ou está deletado na base do Protheus.<br>" &_
            "<br>" &_
            "Caso seja verificado que o pedido não consta como PA, por favor, verifique<br>" &_
            "também o cadastro do produto, para que os próximos pedidos entrem corretamente."
    Aviso( msg )
    response.end
end if

if (AlteraPO = "N") and (AlteraPapel = "N") then
    Aviso( "<b>Você não tem permissão para alterar o pedido!</b>" )
    response.end
end if

'Se o usu�rio n�o tem permiss�o para altera��o de pedidos de caixa informa
if (AlteraPO = "N") And trim(Left(TPOperacoes("Produto"),3)) <> "700" then
    Aviso( "Você não tem permissão para alteração de pedido de caixa e chapa!" )
    response.end
end if

'Se o usu�rio n�o tem permiss�o de altera��o de pedido de papel informa
if (AlteraPapel = "N") And trim(Left(TPOperacoes("Produto"),3)) = "700" then
    Aviso( "Você não tem permissão para alteração de pedido de papel!" )
    response.end
end if

if (trim(TpOperacoes("TipoFrete")) <> "C" and TipoOperacao = 14) then
    Aviso( "Para fazer este tipo de alteração, por favor altere o tipo de Frete de FOB para CIF." )
    response.end
end if
%>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

        <br><br>

        <div class="container-fluid" style="max-width: 1000px;">
            <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
                <div class="container-fluid" style="background-color: rgb(242, 242, 242);">

                    <br>

                    <div class="row">

                        <div class="col">
                            <div class="text-center">
                                <h3>Alteração do Pedido</h3>
                            </div>
                        </div>

                    </div>

                    <form id="form" name="form" method="post" action="AlteraPedido03.asp">

                        <input name="hPedido"           type="hidden" id="hPedido"              value="<%=TPOperacoes("Pedido")%>" />
                        <input name="hCliente"          type="hidden" id="hCliente"             value="<%=TPOperacoes("Cliente")%>" />
                        <input name="hLoja"             type="hidden" id="hLoja"                value="<%=TPOperacoes("Loja")%>" />
                        <input name="hProduto"          type="hidden" id="hProduto"             value="<%=TPOperacoes("Produto")%>"/>
                        <input name="hDescri"           type="hidden" id="hDescri"              value="<%=TPOperacoes("Descricao")%>"/>
                        <input name="hQuantidade"       type="hidden" id="hQuantidade"          value="<%=TPOperacoes("QtdVendida")%>"/>
                        <input name="hQuantidadePronta" type="hidden" id="hQuantidadePronta"    value="<%=TPOperacoes("QtdPronta")%>" />
                        <input name="hPreco"            type="hidden" id="hPreco"               value="<%=TPOperacoes("PrecoUnitario")%>" />
                        <input name="hTotal"            type="hidden" id="hTotal"               value="<%=TPOperacoes("Total")%>"/>
                        <input name="hEntrega1"         type="hidden" id="hEntrega1"            value="<%=TPOperacoes("Entrega1")%>" />
                        <input name="hEntrega2"         type="hidden" id="hEntrega2"            value="<%=TPOperacoes("Entrega2")%>" />
                        <input name="hEmissao"          type="hidden" id="hEmissao"             value="<%=TPOperacoes("Emissao")%>"/>
                        <input name="hTPFrete"          type="hidden" id="hTPFrete"             value="<%=TPOperacoes("TipoFrete")%>" />
                        <input name="hDataCadastro"     type="hidden" id="hDataCadastro"        value="<%=TPOperacoes("DataCadastro")%>"/>
                        <input name="hHoraCadastro"     type="hidden" id="hHoraCadastro"        value="<%=TPOperacoes("HoraCadastro")%>" />
                        <input name="hQuemCadastrou"    type="hidden" id="hQuemCadastrou"       value="<%=TPOperacoes("Digitado")%>"/>
                        <input name="hDataLiberado"     type="hidden" id="hDataLiberado"        value="<%=TPOperacoes("DataLiberado")%>"/>
                        <input name="hHoraLiberado"     type="hidden" id="hHoraLiberado"        value="<%=TPOperacoes("HoraLiberado")%>" />
                        <input name="hQuemLiberou"      type="hidden" id="hQuemLiberou"         value="<%=TPOperacoes("Liberado")%>"/>
                        <input name="hEstatus"          type="hidden" id="hEstatus"             value="<%=TPOperacoes("Estatus")%>"/>
                        <input name="hPedidoCliente"    type="hidden" id="hPedidoCliente"       value="<%=TPOperacoes("PedidoCliente")%>"/>
                        <input name="hCodCondPag"       type="hidden" id="hCodCondPag"          value="<%=TPOperacoes("CondPag")%>" />
                        <input name="hTipoOperacao"     type="hidden" id="hTipoOperacao"        value="<%=TipoOperacao%>" />
                        <input name="hNomeTipoOperacao" type="hidden" id="hNomeTipoOperacao"    value="<%=uCase(trim(Operacoes("operacao")))%>" /> 
                        <input name="NovaCondPag"       type="hidden" id="NovaCondPag"          value="<%=trim(request("para")) %>" />
                        <input name="IdSolicitacao"     type="hidden" id="IdSolicitacao"        value="<%=idSolicitacao %>" />

                        <div class="row">
                            <div class="col">
                                <div class="text-center">
                                    <h5>Dados do Pedido</h5>
                                </div>
                            </div>
                        </div>
                        <hr>

<%
SQL = "{call dbo.SP_HISTORICO_ALTERACOES_PEDIDOS_POR_PEDIDO('" & Pedido & "')}"

SET Historico = Conecta("PIR-HELPDESK", SQL)

if not isobject(Historico) or (Historico IS NOTHING) then
    response.end
end if

if not Historico.EOF then
%>
                        <div class="row">
                            <div class="col">
                                <div class="text-center">
                                    <b>Histórico de Alteração Deste Pedido</b>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col font-weight-bold"></div>
                            <div class="col font-weight-bold">Data</div>
                            <div class="col font-weight-bold">Operação</div>
                            <div class="col font-weight-bold">De</div>
                            <div class="col font-weight-bold">Para</div>
                        </div>
<% 
    while NOT Historico.EOF
%>
                        <div class="row">
                            <div class="col"><%=a%></div>
                            <div class="col"><%=AcertaData(Historico.Fields.Item("DataAlteracao").Value) & " " & (Historico.Fields.Item("HoraAlteracao").Value)%></div>
                            <div class="col"><%=(Historico.Fields.Item("Operacao").Value)%></div>
                            <div class="col"><%=(Historico.Fields.Item("De").Value)%></div>
                            <div class="col"><%=(Historico.Fields.Item("Para").Value)%></div>
                        </div>
<% 
        Historico.MoveNext()
    wend
  
    Historico.Close()
    Set Historico = Nothing
%>
                        <hr>
<%
end if
%>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Pedido:</div>
                            <div class="col"><%=trim(TPOperacoes("Pedido"))%></div>
                            <div class="col-2 font-weight-bold">Cliente:</div>
                            <div class="col"><%=trim(TPOperacoes("Cliente")) & " - " & trim(TPOperacoes("Loja"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Produto:</div>
                            <div class="col"><%=trim(TPOperacoes("Produto"))%></div>
                            <div class="col-2 font-weight-bold">Descrição:</div>
                            <div class="col"><%=trim(TPOperacoes("Descricao"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Quantidade:</div>
                            <div class="col"><%=trim(TPOperacoes("QtdVendida"))%></div>
                            <div class="col-2 font-weight-bold">Pronto:</div>
                            <div class="col"><%=trim(TPOperacoes("QtdPronta"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Qtde.Entregue:</div>
                            <div class="col"><%=trim(TPOperacoes("QtdEntregue"))%></div>
                            <div class="col-2 font-weight-bold"></div>
                            <div class="col"></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Preço unitário:</div>
                            <div class="col"><%=trim(TPOperacoes("PrecoUnitario"))%></div>
                            <div class="col-2 font-weight-bold">Total:</div>
                            <div class="col"><%=trim(TPOperacoes("Total"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Data Entrega 1:</div>
                            <div class="col"><%=trim(AcertaData(TPOperacoes("Entrega1")))%></div>
                            <div class="col-2 font-weight-bold">Data Entrega 2:</div>
                            <div class="col"><%=trim(AcertaData(TPOperacoes("Entrega2")))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Data Emissão</div>
                            <div class="col"><%=trim(AcertaData(TPOperacoes("Emissao")))%></div>
                            <div class="col-2 font-weight-bold">Tipo Frete:</div>
                            <div class="col"><%=trim(TPOperacoes("TipoFrete"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Data Cadastro:</div>
                            <div class="col"><%=AcertaData(TPOperacoes("DataCadastro")) & " " & trim(TPOperacoes("HoraCadastro"))%></div>
                            <div class="col-2 font-weight-bold">Quem Cadastrou:</div>
                            <div class="col"><%=trim(TPOperacoes("Digitado"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Data liberação:</div>
                            <div class="col"><%=trim(AcertaData(TPOperacoes("DataLiberado"))) & " " & trim(TPOperacoes("HoraLiberado"))%></div>
                            <div class="col-2 font-weight-bold">Quem Liberou</div>
                            <div class="col"><%=trim(TPOperacoes("Liberado"))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Status:</div>
                            <div class="col"><%=trim(TPOperacoes("Estatus"))%></div>
                            <div class="col-2 font-weight-bold">Pedido Cliente:</div>
                            <div class="col"><%=trim(trim(TPOperacoes("PedidoCliente")))%></div>
                        </div>
                        <div class="row">
                            <div class="col-2 font-weight-bold">Filial:</div>
                            <div class="col"><%=iif( trim(TPOperacoes("Filial")) = "" or trim(TPOperacoes("Filial")) = "I", "PIRAPETINGA", "UBERABA" )%></div>
                        </div>
                        <hr>
                        <div class="row">
                            <div class="col text-center font-weight-bold">Campos para alteração de <%=trim(ucase(request("NomeTPOperacoes")))%></div>
                        </div>
                        <br>
<%
                        select case TipoOperacao
                        case 1 'COLUNA D EALTERACAO DE TIPO DE FRETE'
%>       
                        <div class="row">
                            <div class="col-2 font-weight-bold"></div>De:</div>
                            <div class="col"><%=TPOperacoes("TipoFrete")%></div>
                            <div class="col-2 font-weight-bold">Para:</div>
                            <div class="col">
                                <input type="text" name="NovoTPFrete" class="form-control" id="NovoTPFrete" value="<%=trim(request("para")) %>" readonly="readonly" />
                                &nbsp;
<%
                            if TpProdutoTipo = "PA" then
%> 
                                <input type="text" name="valorfrete" id="valorfrete" class="form-control" value="0" size="50" />
                                <script type="text/javascript">ValorFrete(document.getElementById("NovoTPFrete").value)</script>
<%
                            end if 
%>
                            </div>
                        </div>
<%	 
                        case 5 'COLUNA D EALTERACAO DE DATA DE ENTREGA'
%>  	  
                        <div class="row">
                            <div class="col-2 font-weight-bold">De Entrega1:</div>
                            <div class="col">
                                <input name="Entrega1" id="Entrega1" type="date" class="form-control" size="10" maxlength="10" disabled="disabled" value="<%=ConverterData(TPOperacoes("Entrega1"),"AAAA-MM-DD") %>" />
                            </div>
                            <div class="col-2 font-weight-bold">Para Entrega1:</div>
                            <div class="col">
                                <input name="NovaEntrega1" id="NovaEntrega1" type="date" class="form-control" size="10" maxlength="10"  runat="server"  value="<%=ConverterData(trim(request("para")),"AAAA-MM-DD") %>" readonly="readonly" />
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-2 font-weight-bold">De Entrega2:</div>
                            <div class="col">
                                <input name="Entrega2" id="Entrega2" type="date" class="form-control" size="10" maxlength="10" disabled="disabled" value="<%=ConverterData(TPOperacoes("Entrega2"),"AAAA-MM-DD") %>" />
                            </div>
                            <div class="col-2 font-weight-bold">Para Entrega2:</div>
                            <div class="col">
                                <input name="NovaEntrega2" id="NovaEntrega2" type="date" class="form-control" size="10" maxlength="10"  runat="server" value="<%=ConverterData(trim(request("para2")),"AAAA-MM-DD") %>" readonly="readonly" />
                            </div>
                        </div>
<%
                        case 2 'COLUNA DE ALTERACAO DE PRECO UNITARIO'
%>
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="Preco" type="text" class="form-control" id="Preco" size="10" value="<%=TPOperacoes("PrecoUnitario")%>" disabled />
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input name="NovoPreco" type="text" class="form-control" id="NovoPreco" size="10" value="<%=trim(request("para")) %>" readonly="readonly" />
                            </div>
                        </div>
<%
	                    case 3,12 'COLUNA DE ALTERACAO DE QUANTIDADE'
%>
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="Quantidade" type="text" class="form-control" id="Quantidade" size="10" value="<%=TPOperacoes("QtdVendida")%>" disabled="disabled" maxlength="6"/>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input name="NovaQuantidade" type="text" class="form-control" id="NovaQuantidade" size="10" maxlength="6" value="<%=trim(request("para")) %>" readonly="readonly" />
                            </div>
                        </div>
<% 
                        case 4 'COLUNA DE ALTERACAO DE QUANTIDADE PRONTA'
%> 
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="QuantidadePronta" type="text" class="form-control" id="QuantidadePronta" size="10" value="<%=TPOperacoes("QtdPronta")%>" disabled="disabled" maxlength="6"/>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input name="NovaQuantidadePronta" type="text" class="form-control" id="NovaQuantidadePronta" size="10" maxlength="6" value="<%=trim(request("para")) %>" readonly="readonly" />
                            </div>
                        </div>
<%  	   
                        case 7 'COLUNA DE ALTERACAO DE STATUS (CANCELAR)'
%> 
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="Estatus" type="text" class="form-control" id="Estatus" size="10" value="<%=TPOperacoes("Estatus")%>" disabled="disabled"/>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input type="text" name="NovoEstatus" id="NovoEstatus" value="<%=trim(request("para")) %>" readonly="readonly" class="form-control"/>
                            </div>
                        </div>

<% 
	                    case 6 'COLUNA DE ALTERACAO DE CONDICOES DE PAGAMENTO'
%>  
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
<%
                            SQL = "" &_
                                "SELECT " &_
                                    "E4_CODIGO, " &_
                                    "E4_DESCRI " &_
                                "FROM ZZ7010 " &_
                                    "INNER JOIN SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                                "WHERE ZZ7_GRUPO = 'SE4' " &_
                                    "AND ZZ7010.D_E_L_E_T_ = '' " &_
                                    "AND E4_LIBVEN <> 'N' " &_
                                    "AND E4_MSBLQL <> '1' " &_
                                    "AND E4_CODIGO <> '090' " &_
                                "ORDER BY  E4_DESCRI"

                            SET CondPag = Conecta("PIR-DADOSADV", SQL)

                            if not isobject(CondPag) or (CondPag IS NOTHING) then
                                response.end
                            end if

                            CondPag_numRows = 0
%>
		                        <select name="CondPag" disabled class="form-control" id="CondPag">
<%
                            while Not CondPag.EOF
                                if trim(TPOperacoes("CondPag")) = trim(CondPag("E4_CODIGO")) then
%>
		                            <option value="<%=CondPag("E4_CODIGO")%>"><%=CondPag("E4_DESCRI")%></option>
<%
                                    varNomeCondPag = CondPag("E4_DESCRI")
                                end if
                                if trim(request("para")) = trim(CondPag("E4_CODIGO")) then
                                    nomeNovaCondicao = CondPag("E4_DESCRI")
                                end if
                                CondPag.MoveNext
                            wend
%>
                                </select> 
                                <input name="hNomeCondPag" type="hidden" id="hNomeCondPag" value="<%=varNomeCondPag%>" />                            
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input type="text" name="NomeNovaCondPag" id="NomeNovaCondPag" value="<%=nomeNovaCondicao %>" readonly="readonly" class="form-control"/>
                            </div>
                        </div>
<%
                        case 8 'COLUNA DE ALTERACAO DO PEDIDO DO CLIENTE'
%>
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="PedidoCliente" type="text" class="form-control" id="PedidoCliente" size="10" value="<%=TPOperacoes("PedidoCliente")%>" disabled="disabled"/>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input name="NovoPedidoCliente" type="text" class="form-control" id="NovoPedidoCliente" size="10" maxlength="15" value="<%=trim(request("para")) %>" readonly="readonly" />
                            </div>
                        </div>
<%
		                case 9,13 'COLUNA DE STATUS DO PEDIDO'
%>       
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <select name="Estatus2" class="form-control" id="Estatus2" disabled>
                                    <option value="<%=TPOperacoes("Estatus")%>" selected="selected"><%=TPOperacoes("Estatus")%></option>
                                </select>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input type="text" name="NovoEstatus2" id="NovoEstatus2" value="<%=trim(request("para")) %>" readonly="readonly" class="form-control"/>
                            </div>
                        </div>
<%
		                case 10 'COLUNA DE STATUS DO PEDIDO'
%>       
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
<%
		                    if trim(TPOperacoes("Filial")) = "" Or trim(TPOperacoes("Filial")) = "I" then
		   	                    Response.Write("<input type=""hidden"" name=""Filial"" value=""I"" />")
			                    Response.Write("PIRAPETINGA")
		                    elseif trim(TPOperacoes("Filial")) = "U" Then
		   	                    Response.Write("<input type=""hidden"" name=""Filial"" value=""U"" />")
		   	                    Response.Write("UBERABA")
		                    end if
%>                           
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input type="text" name="NovaFilial" id="NovaFilial" class="form-control" value="<%=trim(request("para")) %>" readonly="readonly" />
                            </div>
                        </div>
<%
                        case 11 'COLUNA DE ALTERACAO DE MOTIVO DE ATRAZO'
%> 
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <%Response.Write TPOperacoes("MtAtrazo") %>
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <textarea name="NovoMotivoAtrazo" class="form-control" id="NovoMotivoAtrazo" size="10" cols="45" rows="5" readonly="readonly"><%=trim(request("para")) %></textarea>
                            </div>
                        </div>
<%
                        case 14 'COLUNA DE ALTERACAO DE PRECO FRETE UNITARIO'
%>
                        <div class="row">
                            <div class="col font-weight-bold">De:</div>
                            <div class="col">
                                <input name="FreteUnitario" type="text" class="form-control" id="FreteUnitario" size="10" value="<%=TPOperacoes("Frete")%>" readonly="readonly" />
                            </div>
                            <div class="col font-weight-bold">Para:</div>
                            <div class="col">
                                <input name="NovoFreteUnitario" type="text" class="form-control" id="NovoFreteUnitario" value="<%=trim(request("para")) %>" readonly="readonly" size="10" />
                            </div>
                        </div>
<% 
	                    end select 
%> 
                        <hr>
                        <div class="row">
                            <div class="col d-flex justify-content-center">
                                <button class="btn btn-success" title="" onclick="verifica()">Alterar</button>
                                &nbsp;
                                <a href="#" class="btn btn-danger" onclick="history.go(-1);"><i class="fas fa-arrow-left"></i> Voltar</a>
                            </div>
                        </div>
                        <br>
                    </form>
                </div>
            </div>
        </div>
    </body>
<%
function Aviso( msg )

    mostraAviso = "" &_
        "<br><br>" &_
        "<div class='alert alert-danger alert-dismissible w-50 mx-auto alert-vertical-center' role='alert'>" &_
            "<a href='#' class='close' data-dismiss='alert' aria-label='close' onclick='window.history.back();'>&times;</a>" &_
            "<br>" &_
            "<h5>A V I S O !!!</h5>" &_
            "<hr>" &_
            "<p>" & msg & "</p>" &_
        "</div>"

    response.write mostraAviso
    'Aviso = NOTHING

end function
%>
</html>
