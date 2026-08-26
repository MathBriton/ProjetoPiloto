<!DOCTYPE html>
<%
'.........................................
'RECUPERANDO O PARAM DA GRAVAÇÃO DO STATUS
'.........................................
param=request.querystring("param")
%>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <title>Valida Pedido d Papel</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />

    <!--#include file="../../../../includes/variaveis.asp" -->
    <!--#include file="../../../../includes/atribui_nivel.inc" -->
    <!--#include file="../../../../cnn/cnn.asp" --> 'DADOSADV
    <!--#include file="../../../../funcoes/funcoes.asp" -->

    <%call ValidaPermissaoFormularios("sfa/pedidos/validacoes/pa",SESSION("CODUSER"),SESSION("FILIAL")) %>

    <%if param = 1 then%>
	    <script type="text/javascript">
	        function drawAlert () {
	            alert ("Operação realizada com sucesso.");
	        }
	        drawAlert()
	    </script> 
    <%end if%>

    <script type="text/javascript">
        function TesExcecao() {
            alert('Este pedido possui uma TES destinada a PEDIDOS LIVRES DE DEBITO.\nNestes casos somente o Gerente Comercial pode proceder com a validação.');
        }
        function confirma() {
            var agree = confirm("Confirma o envio de todos os pedidos liberados?");
            if (agree) {
                document.ValidaPedidos.submit();
            }
        }
    </script>

    <style type="text/css">
        body{
            margin-top:20px;
            font-size:95%;
        }
        .margemLegenda {
            margin-bottom:20px;
        }
    </style>

</head>
<body>
<%
    Dados.Source = "Execute SP_LIBERAPEDIDOSZF_PAPEL"    
    Dados.Open()
%>
    <div class="container-fluid">
        <div class="alert alert-secondary col-xs-10 col-sm-10 col-md-10 offset-1 text-center">
            <strong><h3>Validação de pedidos de Papel</h3></strong>
        </div>
        <div class="row">
            <div class="col-xs-10 col-sm-10 col-md-10 offset-1 text-center">
                <span class="badge bg-success">LEGENDA 1:</span>
                <span class="badge bg-secondary">LI - Pedidos Liberados</span>
                <span class="badge bg-secondary">PR - Bloqueado por Preço</span>
                <span class="badge bg-secondary">CP - Bloqueado por condição de pagamento</span>
            </div>
        </div>
        <div class="row margemLegenda">
            <div class="col-xs-10 col-sm-10 col-md-10 offset-1 text-center">
                <span class="badge bg-success">LEGENDA 2:</span>
                <span class="badge bg-danger">Bloqueados</span>
                <span class="badge bg-warning">Livres de débito</span>
                <span class="badge bg-success">Liberados</span>
                <span class="badge bg-info">Pendentes de liberação</span>
            </div>
        </div>
        <%if not Dados.eof then %>
        <form method="post" action="envia_pa.asp" name="ValidaPedidos">
            <div class="row">
                <div class="col-xs-10 col-sm-10 col-md-10 offset-1">
                    <table class="table table-condensed">
                        <thead>
                            <tr>
                                <th>Pedido</th>
                                <th>Cliente</th>
                                <th>Loja</th>
                                <th>Nome</th>
                                <th>FT</th>
                                <th>Produto</th>
                                <th>Quant.</th>
                                <th>Preço</th>
                                <th>Emissão</th>
                                <th>Entrega</th>
                                <th>Obs.</th>
                                <th>Status</th>
                                <th>Bloqueio</th>
                            </tr>              
                        </thead>
                        <tbody>
<% 
                        While not Dados.Eof 
                            select case Trim(Dados("ZF_ESTATUS"))
                                case "LI"
                                    classe = "success"
                                case "PR","DE","CP"
                                    classe = "danger"
                                case else
                                    classe = "info"
                            end select
                            if ((Trim(Dados("ZF_TES"))="580") or (Trim(Dados("ZF_TES"))="581") or (Trim(Dados("ZF_TES"))="715") or (Trim(Dados("ZF_TES"))="716")) then
                                classe = "warning"
                                tesExcecao = 1
                            else
                                tesExcecao = 0
                            end if   
%>
                            <tr class="<%=classe %>">
                                <%if tesExcecao = 1 AND USR <> "CARLOSS" AND USR <> "LPS" AND USR <> "romeu" then%>
                                <td class="text-center"><a href="#" onclick="TesExcecao()"><%=Dados("ZF_NUMERO") %></a></td>
                                <%else %>
                                <td class="text-center"><a href="detalha_pa.asp?pedido=<%=Trim(Dados("ZF_NUMERO")) %>"><%=Dados("ZF_NUMERO") %></a></td>
                                <%end if %>
                                <td><%=Dados("ZF_CLIENTE") %></td>
                                <td><%=Dados("ZF_LOJACLI") %></td>
                                <td><%=Trim(Dados("CLIENTE")) %></td>
                                <td class="text-center"><%=Trim(Dados("ZF_PRODUTO")) %></td>
                                <td><%=Trim(Dados("REF"))%></td>
                                <td class="text-right"><%=Dados("ZF_QUANT") %></td>
                                <td class="text-right"><%=FormatNumber(Dados("ZF_PRECO"),6) %></td>
                                <td class="text-center"><%=AcertaData(Dados("ZF_EMISSAO")) %></td>
                                <td class="text-center"><%=AcertaData(Dados("ZF_ENT1")) %></td>
                                <td><%=Trim(Dados("ZF_OBSVEND")) %></td>
                                <td class="text-center"><%=Dados("ZF_ESTATUS") %></td>
                                <td><%=Dados("ZF_DESCBL") %></td>
                            </tr>
<% 
                            Dados.MoveNext()
                        Wend
%>
                        </tbody>
                        <tfoot>
                            <tr>
                                <th colspan="13" class="pading-top">
                                    <button name="liberar" id="liberar" class="btn btn-success col-xs-12 col-sm-12 col-md-12" onclick="confirma();">
                                        LIBERAR PEDIDOS
                                    </button>
                                </th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </form>
        <%else %>
        <div class="alert alert-info col-xs-10 col-sm-10 col-md-10 offset-1">
            <strong>Info!</strong> Não foi encontrado nenhum pedido para liberação.
        </div>
        <%end if %>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>