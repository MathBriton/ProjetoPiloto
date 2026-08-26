<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../../cnn/cnn.asp" -->
<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<%
session.LCID = 1046 

%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Westrock</title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
    <script type="text/javascript">
        function excluir(pedido,planta) {
            var confirmacao = confirm("Deseja realmente Excluir o Pedido?");
            if(confirmacao == false){
               return 
            }
            document.getElementById('pedido').value = pedido;
            document.getElementById('hidPlanta').value=planta;
            PedidoDigitado.submit();
        }
    </script>
</head>
<%
Planta = request("Planta")

if trim(request("pedido")) <> "" then

        SQL = "UPDATE Pedido SET "
        SQL = SQL &   "Excluido = 1 "
        SQL = SQL &   ",QuemExcluiu = "&IdUsuario
        SQL = SQL &   ",DataExclusao = GETDATE() "
        SQL = SQL & "WHERE Isnull(Status,0) = 0 "
        SQL = SQL &   "AND Pedido = '"&trim(request("pedido"))&"' "

        if Nivel <> 0 then
            SQL = SQL & " AND QuemIncluiu = '"&IdUsuario&"' "
        end if

        select case Planta
               case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
               case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
               case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
        end select
      
        if vartype(consulta) = 0 or vartype(consulta) = 1 then 
          response.end
        end if

        SQL = "SELECT Pedido  "
        SQL = SQL & "FROM Pedido "
        SQL = SQL & " WHERE Pedido = '"&trim(request("pedido"))&"' "
        SQL = SQL & " AND Isnull(Excluido,0) = 1 "
        SQL = SQL & " AND QuemExcluiu = "&IdUsuario 
        'consulta.Open(SQL)
        select case Planta
               case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
               case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
               case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
        end select
        if vartype(consulta) = 0 or vartype(consulta) = 1 then 
          response.end
        end if

        if Not consulta.Eof then
            response.write "<script type=""text/javascript"">alert('Pedido "&trim(request("pedido"))&" excluido com sucesso.')</script>" 
        else
            response.write "<script type=""text/javascript"">alert('Pedido "&trim(request("pedido"))&" não pode ser excluído.')</script>" 
        end if

end if

SQL = "" &_
  "SELECT " &_
    "Pedido, " &_
    "Planta, " &_
    "Status, " &_
    "LotePiloto, " &_
    "ClienteRemessa, " &_
    "Produto, " &_
    "Referencia, " &_
    "Estrutura, " &_
    "Isnull(PrecoUnitario,0) PrecoUnitario, " &_
    "Quantidade, " &_
    "Observacao, " &_
    "Largura, " &_
    "Comprimento, " &_
    "Cliente, " &_
    "LojaCli, " &_
    "DataEmissao, " &_
    "PrimeiraDataEntrega, " &_
    "DescricaoBloqueio, " &_
    "NumeroOrdemCompra, " &_
    "ItemOrdemCompra " &_
  "FROM Pedido " &_
  "WHERE Isnull(Status,0) = 0 " &_
    "AND Isnull(Excluido,0) = 0 " &_
    "AND QuemIncluiu = '" & IdUsuario & "' " &_
  "ORDER BY Pedido DESC "

select case Planta
       case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
       case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
       case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
end select
      
if vartype(consulta) = 0 or vartype(consulta) = 1 then 
    response.end
end if

%>
<body>
    <div class="container-fluid">
        <div class="row">
            <div class="col-8 offset-2">
                <h3><strong>Pedidos Digitados</strong></h3>
            </div>
        </div>
        <div class="row">
            <div class="col-3 offset-2">
                <a href="selecionaPlanta.asp" class="btn btn-secondary"><i class="fas fa-plus"></i>&nbsp;Emitir novo pedido</a>
            </div>
        </div>
        <form name="PedidoDigitado" id="PedidoDigitado" action="PedidoDigitado.asp?Planta=<%=Planta%>" method="post">
            <input type="hidden" name="pedido" id="pedido" />
            <input type="hidden" name="hidPlanta" id="hidPlanta" />
            <div class="row" style="margin-top:20px;">
                <div class="col-8 offset-2">
                    <table class="table table-striped table-hover table-responsive-sm">
                        <thead class="table-dark">
                            <tr>
                                <th class="text-center">Pedido</th>
                                <th class="text-center">Cliente</th>
                                <th class="text-center">Nome</th>
                                <th class="text-center">Ordem de Compra</th>
                                <th class="text-center">FT</th>
                                <th class="text-center">Produto</th>
                                <th class="text-center">Quant.</th>
                                <th class="text-center">Preço</th>
                                <th class="text-center">Emissão</th>
                                <th class="text-center">Entrega</th>
                                <th class="text-center">Ação</th>
                            </tr>              
                        </thead>
                        <tbody>
<%
                            While not consulta.eof 
                                if Planta = "05" or Planta = "06" then 
                                    SQL = "SELECT NOME_CLIENTE As Razao FROM VW_CLIENTES_PEDIDO WHERE COD_CLIENTE = "&consulta("Cliente")&" "
                                else     
                                    SQL = "SELECT NOME_CLIENTE As Razao FROM VW_CLIENTES_PEDIDO WHERE COD_CLIENTE = "&consulta("Cliente")&" AND LOJA_CLIENTE = " & consulta("LojaCli") & " "
                                end if 
                                select case Planta
                                    case "05" set consultaCliente = Conecta("BEN-DADOSADV", SQL )
                                    case "06" set consultaCliente = Conecta("FOR-DADOSADV", SQL )
                                    case "07" set consultaCliente = Conecta("SAQ-DADOSADV", SQL )
                                end select
                                    
                                if vartype(consultaCliente) = 0 or vartype(consultaCliente) = 1 then 
                                    response.end
                                end if
                                razao = "" 
                                if not consultaCliente.eof then 
                                   razao = Trim(consultaCliente("Razao"))
                                end if 
%>
                            <tr class="<%=classe %>">
                                <td class="text-end"> <%=consulta("Pedido") %></td>
                                <td class="text-end"><%=consulta("Cliente") %></td>
                                <td><%=razao %></td>
                                <td class="text-end"><%=consulta("NumeroOrdemCompra") & "-" & consulta("ItemOrdemCompra") %></td>
                                <td class="text-center"><%=Trim(consulta("Produto")) %></td>
                                <td><%=Trim(consulta("Referencia"))%></td>
                                <td class="text-end"><%=consulta("Quantidade") %></td>
                                <td class="text-end"><%=FormatNumber(consulta("PrecoUnitario"),6) %></td>
                                <td class="text-center"><%=consulta("DataEmissao") %></td>
                                <td class="text-center"><%=consulta("PrimeiraDataEntrega") %></td>
                                <td class="text-center"><a href="#" onclick="excluir('<%=consulta("Pedido")%>','<%=consulta("Planta")%>')"><i class="far fa-trash-alt"></i></a> </td>
                            </tr>
<%
                            response.Flush() 
                            consulta.MoveNext() 
                        Wend 
                        SET consulta = nothing
%>
                        </tbody>
                    </table>
                </div>
            </div>
        </form>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>