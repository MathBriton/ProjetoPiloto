<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../../cnn/cnn.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<%
session.LCID = 1046 
'response.Cookies("PlantaDefinida") = request("Planta")
'Planta = request.cookies(PlantaDefinida.value)
'Planta = request("paramPlanta")
Planta = request("Planta")
select case Planta
       case "05" SubTitulo = "Bento Gonçalves"
       case "06" SubTitulo = "Fortaleza"
       case "07" SubTitulo = "Saquarema"
end select
'response.write "Planta " & Planta
%>

<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <title>Smurfit Kappa</title>
    <meta charset="utf-8"/>
    <meta http-equiv="refresh" content="30" />
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
  </head>
  <body>
    <div class="container">
      <div class="row">
        <div class="alert" role="alert">
          <center><h3>Validação de Pedidos <%=SubTitulo%></h3></center>
        </div>
      </div>
      <div class="row margemLegenda">
        <div class="col-10 offset-1 text-center">
          <span class="badge bg-dark">LEGENDA:</span>
          <span class="badge bg-danger">Bloqueados</span>
          <span class="badge bg-warning">Pagamento antecipado</span>
          <span class="badge bg-info">Pendentes de liberação</span>
          <span class="badge bg-secondary">Lote Piloto</span>
        </div>
      </div>
      <form name="ValidaPedidos" id="ValidaPedidos" action="ListaPedidosValidar.asp?Planta="+<%=axPlanta%> method="post">
        <div class="row" style="margin-top:20px;">
          <div class="col-10">
            <table class="table table-striped table-hover table-responsive-sm">
              <thead class="table-default">
                <tr>
                  <th class="text-center">#</th>
                  <th class="text-center">Pedido</th>
                  <th class="text-center">Cliente</th>
                  <th class="text-center">Nome</th>
                  <th class="text-center">FT</th>
                  <th class="text-center">Produto</th>
                  <th class="text-center">Tipo Produto</th>
                  <th class="text-center">Tipo Onda</th>
                  <th class="text-center">Quant.</th>
                  <th class="text-center">Preço Net</th>
                  <th class="text-center text-nowrap">Preço KG Net</th>
                  <th class="text-center text-nowrap">M2</th>
                  <th class="text-center">Emissão</th>
                  <th class="text-center">Entrega</th>
                  <th class="text-center">Status</th>
                  <th class="text-center">Bloqueio</th>
                  <th class="text-center">Quem Emitiu</th>
                </tr>              
              </thead>
              <tbody>
<%
if Planta = "05" then 
              SQL = "" & _
              "SELECT * " & _
              "FROM ( " & _
                "SELECT " & _
                  "p.Pedido " & _
                  ",ValorICMSUnitario " & _
                  ",ValorPISUnitario " & _
                  ",ValorCOFINSUnitario " & _
                  ",p.Planta  " & _
                  ",p.[Status] " & _
                  ",p.LotePiloto " & _
                  ",p.ClienteRemessa " & _
                  ",p.Produto " & _
                  ",p.Referencia " & _
                  ",ISNULL(IT.TIPO_PRODUTO, 'CONJUNTO') TIPO_PRODUTO " & _
                  ",p.Estrutura " & _
                  ",Isnull(p.PrecoUnitario,0) PrecoUnitario " & _
                  ",p.Quantidade " & _
                  ",p.Observacao " & _
                  ",p.Largura " & _
                  ",p.Comprimento " & _
                  ",p.Cliente " & _
                  ",p.DataEmissao " & _
                  ",p.PrimeiraDataEntrega " & _
                  ",p.DescricaoBloqueio " & _
                  ",p.StatusBloqueio " & _
                  ",p.CondicaoPagamento " & _
                  ",c.Razao_social_cliente As Razao " & _
                  ",pr.Situacao " & _
                  ",pr.Classe " & _
                  ",ISNULL(UI.CRM_NREDUZ, UI.CRM_LOGIN) AS QuemIncluiu, p.QuemIncluiu Id " & _
                  ",ISNULL(p.Quantidade * ISNULL(IT.AREA,CJ.AREA),0) Area_Vendida " & _
                "FROM Pedido p " & _
                  "INNER JOIN SOLPEN..interface_inpanet_clientes c ON c.Cliente = p.Cliente " & _
                  "LEFT JOIN SOLPEN..interface_inpanet_produtos pr ON p.Produto Collate SQL_Latin1_General_CP437_BIN = pr.Codigo " & _
                  "LEFT JOIN BRPIR1M001.HELPDESK.dbo.CRM_DTI UI ON ((QuemIncluiu = UI.CRM_COD AND UI.EMP_CODIGO = '05' ) OR QuemIncluiu = UI.CRM_R_E_C_N_O) AND CRM_D_E_L_E_T = '' " & _
                  "LEFT JOIN TRIMBOX..VW_BI_PRODUTOS_PEDIDO IT ON IT.CODIGO_ITEM = p.Produto COLLATE Latin1_General_CI_AS " & _
				          "OUTER APPLY(SELECT MAX(GRAMATURA) GRAMATURA, SUM(AREA) AREA FROM TRIMBOX..VW_BI_PRODUTOS_PEDIDO PP (NOLOCK) WHERE PP.CODIGO_CONJUNTO = p.Produto COLLATE Latin1_General_BIN) AS CJ " & _
                "WHERE Isnull(p.Status,0) = 0 " & _
                  "AND Isnull(p.Excluido,0) = 0 " & _
                " UNION " & _
                "SELECT " & _
                  "p.Pedido " & _
                  ",ValorICMSUnitario " & _
                  ",ValorPISUnitario " & _
                  ",ValorCOFINSUnitario" & _
                  ",p.Planta " & _
                  ",p.[Status] " & _
                  ",p.LotePiloto " & _
                  ",p.ClienteRemessa " & _
                  ",p.Produto " & _
                  ",p.Referencia " & _
                  ",ISNULL(IT.TIPO_PRODUTO, 'CONJUNTO') TIPO_PRODUTO " & _
                  ",p.Estrutura " & _
                  ",Isnull(p.PrecoUnitario,0) PrecoUnitario " & _
                  ",p.Quantidade " & _
                  ",p.Observacao " & _
                  ",p.Largura " & _
                  ",p.Comprimento " & _
                  ",p.Cliente " & _
                  ",p.DataEmissao " & _
                  ",p.PrimeiraDataEntrega " & _
                  ",p.DescricaoBloqueio " & _
                  ",p.StatusBloqueio " & _
                  ",p.CondicaoPagamento " & _
                  ",c.Razao_social_cliente As Razao " & _
                  ",pr.Situacao " & _
                  ",pr.Classe " & _
                  ",ISNULL(UI.CRM_NREDUZ, UI.CRM_LOGIN) AS QuemIncluiu, p.QuemIncluiu Id " & _ 
                  ",ISNULL(p.Quantidade * ISNULL(IT.AREA,CJ.AREA),0) Area_Vendida " & _
                "FROM Pedido p " & _
                  "INNER JOIN SOLPEN..interface_inpanet_clientes c ON c.Cliente = p.Cliente " & _
                  "LEFT JOIN SOLPEN..interface_inpanet_produtos pr ON p.Produto Collate SQL_Latin1_General_CP437_BIN = pr.Codigo " & _
                  "LEFT JOIN BRPIR1M001.HELPDESK.dbo.CRM_DTI UI ON ((QuemIncluiu = UI.CRM_COD AND UI.EMP_CODIGO = '05' ) OR QuemIncluiu = UI.CRM_R_E_C_N_O) AND CRM_D_E_L_E_T = ''" & _
                  "LEFT JOIN TRIMBOX..VW_BI_PRODUTOS_PEDIDO IT ON IT.CODIGO_ITEM = p.Produto COLLATE Latin1_General_CI_AS " & _
				          "OUTER APPLY(SELECT MAX(GRAMATURA) GRAMATURA, SUM(AREA) AREA FROM TRIMBOX..VW_BI_PRODUTOS_PEDIDO PP (NOLOCK) WHERE PP.CODIGO_CONJUNTO = p.Produto COLLATE Latin1_General_BIN) AS CJ " & _
                "WHERE Isnull(p.Status,0) = 2 " & _
                  "AND p.DataBloqueio > DATEADD(day,-1,GETDATE()) " & _
                  "AND Isnull(p.Excluido,0) = 0 ) AS TESTE " & _
                "ORDER BY isnull(Status,0), PrimeiraDataEntrega desc "

              Dados131.Open(SQL)
                           
              corDaLinha = ""

              while not Dados131.Eof 

                select case (cstr(Dados131("StatusBloqueio")))
                case "1" corDaLinha="table-warning"
                case "2" corDaLinha="table-danger"
                case "3" corDaLinha="table-danger"
                case "4" corDaLinha="table-danger"
                case "5" corDaLinha="table-danger"
                case "6" corDaLinha="table-danger"
                case else corDaLinha="table-info"
                end select 

                Fabricacao = iif( (Dados131("Planta") = "05"), "Bento gonçalves", "MATRIZ" )

                select case Trim(Dados131("Status"))
                case "1" classe = "table-success" : descricaoStatus = "Liberado"
                case "2" classe = "table-danger"  : descricaoStatus = "Bloqueado"
                case else classe = "table-info"   : descricaoStatus = "Pendente"
                end select

                if Trim(Dados131("LotePiloto")) = "S" then
                  corDaLinha = "table-secondary"
                  classe = "default"
                end if 

                if Trim(Dados131("CondicaoPagamento"))  = "354" then
                  classe = "table-warning"
                end if

                remessa = iif( (Trim(Dados131("ClienteRemessa")) <> ""), 1, 0 )
                estrutura = Trim(Dados131("Estrutura"))
                tipoOnda = iif( (replace(Mid(trim(estrutura),7,1),"0","") = ""), "SIMPLES", "DUPLA" )

                sql = "" & _
                  "SELECT " & _
                    "c.Composicao, " & _
                    "AreaBrutaPecaComRefilos as Area, " & _
                    "c.Gramatura, " & _
                    "Item as Produto, " & _
                    "c.Gramatura / 1000.000 * AreaBrutaPecaComRefilos as Peso " & _
                  "FROM vw_ConsultaFTsPelaInpanet ft " & _
                    "INNER JOIN vw_ConsultaComposicaoPelaInpaNet c ON ft.Composicao = c.Composicao " & _
                  "WHERE ft.Item = '"&trim(Dados131("Produto"))&"' and AreaBrutaPecaComRefilos > 0 "
                Dados92.Open(sql)

                if not Dados92.Eof then

                  composicao = Dados92("Composicao")
                  gramatura = Dados92("Gramatura")
                  peso = Dados92("Peso")
                  area = Dados92("Area")

                else

                  sql = "" & _
                    "SELECT " & _
                      "c.Composicao, " & _
                      "AreaBrutaPecaComRefilos as Area, " & _
                      "c.Gramatura, " & _
                      "Item as Produto, " & _
                      "c.Gramatura / 1000.000 * AreaBrutaPecaComRefilos as Peso " & _
                    "FROM vw_ConsultaFTsPelaInpanet ft " & _
                      "INNER JOIN vw_ConsultaComposicaoPelaInpaNet c ON ft.Composicao = c.Composicao " & _
                    "WHERE ft.Conjunto = '"&trim(Dados131("Produto"))&"' and AreaBrutaPecaComRefilos > 0 "
                  Dados93.Open(sql)

                  composicao  = ""
                  gramatura   = 0
                  peso        = 0
                  area        = 0

                  if not Dados93.Eof then

                    composicao  = Dados93("Composicao")
                    gramatura   = Dados93("Gramatura")
                    peso        = Dados93("Peso")
                    area        = Dados93("Area")

                  end if

                  Dados93.Close()

                end if

                Dados92.Close()

                if Trim(Dados131("Classe")) = "CH" and gramatura > 0 and Trim(Dados131("Estrutura")) <> "" then
                  area = FormatNumber((Dados131("Largura")/1000) * (Dados131("Comprimento")/1000),4)
                  peso = formatNumber(gramatura/1000*area,2)
                end if

                if Trim(Dados131("Classe")) = "PA" then
                  area = 0
                  peso = Dados131("Quantidade")
                end if

                precoUnitarioNet = Dados131("PrecoUnitario") - Dados131("ValorICMSUnitario") - Dados131("ValorPISUnitario") - Dados131("ValorCOFINSUnitario")

                if peso > 0 then
                  precoQuiloNet = precoUnitarioNet / peso
                else
                  precoQuiloNet = 0
                end if

                if area > 0 then
                  precoAreaNet = precoUnitarioNet / area
                else
                  precoAreaNet = 0
                end if
                set consulta = Dados131
%>
                <tr class="<%=corDaLinha %>">
                  <td>      
                    <a href="detalhesValidacao.asp?Pedido=<%=consulta("Pedido")%>" class="btn btn-outline-primary" title="Detalhes do pedido.">
                      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-arrow-return-left" viewBox="0 0 16 16">
                        <path fill-rule="evenodd" d="M14.5 1.5a.5.5 0 0 1 .5.5v4.8a2.5 2.5 0 0 1-2.5 2.5H2.707l3.347 3.346a.5.5 0 0 1-.708.708l-4.2-4.2a.5.5 0 0 1 0-.708l4-4a.5.5 0 1 1 .708.708L2.707 8.3H12.5A1.5 1.5 0 0 0 14 6.8V2a.5.5 0 0 1 .5-.5z"/>
                      </svg>
                    </a>
                  </td>
                  <td class="text-end"><%=consulta("Pedido")%></td>
                  <td class="text-end"><%=consulta("Cliente")%></td>
                  <td><%=Trim(consulta("Razao")) %></td>
                  <td class="text-center"><%=Trim(consulta("Produto")) %></td>
                  <td><%=Trim(consulta("Referencia"))%></td>
                  <td><%=Trim(consulta("TIPO_PRODUTO"))%></td>
                  <td><%=Trim(tipoOnda)%></td>
                  <td class="text-end"><%=consulta("Quantidade") %></td>
                  <td class="text-end"><%=FormatNumber(precoUnitarioNet,6) %></td>
                  <td class="text-end"><%=FormatNumber(precoQuiloNet,5) %> </td>
                  <td class="text-end"><%=FormatNumber(consulta("Area_Vendida"),5)%> </td>
                  <td class="text-center"><%=consulta("DataEmissao") %></td>
                  <td class="text-center"><%=consulta("PrimeiraDataEntrega") %></td>
                  <td class="text-center"><%=descricaoStatus %></td>
                  <td><%=consulta("DescricaoBloqueio") %></td>
                  <td><%=consulta("QuemIncluiu") %></td>
                </tr>
<%
                response.Flush() 
                consulta.MoveNext() 
              wend
end if 

if Planta = "06" then                 
          SQL =      " Select * from VW_CONSULTA_PEDIDOS_VALIDAR " 
          SQL = SQL & " ORDER BY isnull(Status,0), PrimeiraDataEntrega "
          Dados201.Open(SQL)

          corDaLinha = ""
          while not Dados201.Eof 

            select case (cstr(Dados201("StatusBloqueio")))
            case "1" corDaLinha="table-warning"
            case "2" corDaLinha="table-danger"
            case "3" corDaLinha="table-danger"
            case "4" corDaLinha="table-danger"
            case "5" corDaLinha="table-danger"
            case "6" corDaLinha="table-danger"
            case else corDaLinha="table-info"
            end select 

            Fabricacao = iif( (Dados201("Planta") = "06"), "Fortaleza", "MATRIZ" )
            axPlanta = Dados201("Planta")

            select case Trim(Dados201("Status"))
                   case "1" classe = "table-success" : descricaoStatus = "Liberado"
                   case "2" classe = "table-danger"  : descricaoStatus = "Bloqueado"
                   case else classe = "table-info"   : descricaoStatus = "Pendente"
            end select

            if Trim(Dados201("LotePiloto")) = "S" then
              corDaLinha = "table-secondary"
              classe = "default"
            end if 

            if Trim(Dados201("CondicaoPagamento"))  = "354" then
              classe = "table-warning"
            end if

            remessa = iif( (Trim(Dados201("ClienteRemessa")) <> ""), 1, 0 )
            estrutura = Trim(Dados201("Estrutura"))
            tipoOnda = iif( (replace(Mid(trim(estrutura),7,1),"0","") = ""), "SIMPLES", "DUPLA" )

            area = Dados201("Area_Unitaria")
            peso = Dados201("PesoLiquido")
            gramatura = Dados201("Gramatura")
            precoUnitarioNet = Dados201("Preco_Net") 
            precoQuiloNet = Dados201("Preco_KG_Net") 
            precoAreaNet = Dados201("Preco_M2_Net")
            AreaVendida = Dados201("Area_Vendida")
            set consulta = Dados201


%>
                <tr class="<%=corDaLinha %>">
                  <td>
                    <a href="detalhesValidacaoFortaleza.asp?Pedido=<%=consulta("Pedido")%>" class="btn btn-outline-primary" title="Detalhes do pedido.">
                      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-arrow-return-left" viewBox="0 0 16 16">
                        <path fill-rule="evenodd" d="M14.5 1.5a.5.5 0 0 1 .5.5v4.8a2.5 2.5 0 0 1-2.5 2.5H2.707l3.347 3.346a.5.5 0 0 1-.708.708l-4.2-4.2a.5.5 0 0 1 0-.708l4-4a.5.5 0 1 1 .708.708L2.707 8.3H12.5A1.5 1.5 0 0 0 14 6.8V2a.5.5 0 0 1 .5-.5z"/>
                      </svg>
                    </a>
                  </td>
                  <td class="text-end"><%=consulta("Pedido")%></td>
                  <td class="text-end"><%=consulta("Cliente")%></td>
                  <td><%=Trim(consulta("Razao")) %></td>
                  <td class="text-center"><%=Trim(consulta("Produto")) %></td>
                  <td><%=Trim(consulta("Referencia"))%></td>
                  <td><%=Trim(consulta("TIPO_PRODUTO"))%></td>
                  <td><%=Trim(tipoOnda)%></td>
                  <td class="text-end"><%=consulta("Quantidade") %></td>
                  <td class="text-end"><%=FormatNumber(precoUnitarioNet,6) %></td>
                  <td class="text-end"><%=FormatNumber(precoQuiloNet,5) %> </td>
                  <td class="text-end"><%=FormatNumber(AreaVendida,5) %> </td>
                  <td class="text-center"><%=consulta("DataEmissao") %></td>
                  <td class="text-center"><%=consulta("PrimeiraDataEntrega") %></td>
                  <td class="text-center"><%=descricaoStatus %></td>
                  <td><%=consulta("DescricaoBloqueio") %></td>
                  <td><%=consulta("QuemIncluiu") %></td>
                </tr>
<%
                response.Flush() 
                consulta.MoveNext() 
              wend
end if 




if Planta = "07" then                 
          SQL =      " Select * from VW_CONSULTA_PEDIDOS_VALIDAR " 
          SQL = SQL & " ORDER BY isnull(Status,0), PrimeiraDataEntrega "
          Dados181.Open(SQL)

          corDaLinha = ""
          while not Dados181.Eof 

            select case (cstr(Dados181("StatusBloqueio")))
            case "1" corDaLinha="table-warning"
            case "2" corDaLinha="table-danger"
            case "3" corDaLinha="table-danger"
            case "4" corDaLinha="table-danger"
            case "5" corDaLinha="table-danger"
            case "6" corDaLinha="table-danger"
            case else corDaLinha="table-info"
            end select 

            Fabricacao = iif( (Dados181("Planta") = "07"), "Saquarema", "MATRIZ" )
            axPlanta = Dados181("Planta")

            select case Trim(Dados181("Status"))
                   case "1" classe = "table-success" : descricaoStatus = "Liberado"
                   case "2" classe = "table-danger"  : descricaoStatus = "Bloqueado"
                   case else classe = "table-info"   : descricaoStatus = "Pendente"
            end select

            if Trim(Dados181("LotePiloto")) = "S" then
              corDaLinha = "table-secondary"
              classe = "default"
            end if 

            if Trim(Dados181("CondicaoPagamento"))  = "354" then
              classe = "table-warning"
            end if

            remessa = iif( (Trim(Dados181("ClienteRemessa")) <> ""), 1, 0 )
            estrutura = Trim(Dados181("Estrutura"))
            tipoOnda = iif( (replace(Mid(trim(estrutura),7,1),"0","") = ""), "SIMPLES", "DUPLA" )

            area = Dados181("Area_Unitaria")
            peso = Dados181("PesoLiquido")
            gramatura = Dados181("Gramatura")
            precoUnitarioNet = Dados181("Preco_Net") 
            precoQuiloNet = Dados181("Preco_KG_Net") 
            precoAreaNet = Dados181("Preco_M2_Net")
            AreaVendida = Dados181("Area_Vendida")
            set consulta = Dados181


%>
                <tr class="<%=corDaLinha %>">
                  <td>
                    <a href="detalhesValidacaoSaquarema.asp?Pedido=<%=consulta("Pedido")%>" class="btn btn-outline-primary" title="Detalhes do pedido.">
                      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-arrow-return-left" viewBox="0 0 16 16">
                        <path fill-rule="evenodd" d="M14.5 1.5a.5.5 0 0 1 .5.5v4.8a2.5 2.5 0 0 1-2.5 2.5H2.707l3.347 3.346a.5.5 0 0 1-.708.708l-4.2-4.2a.5.5 0 0 1 0-.708l4-4a.5.5 0 1 1 .708.708L2.707 8.3H12.5A1.5 1.5 0 0 0 14 6.8V2a.5.5 0 0 1 .5-.5z"/>
                      </svg>
                    </a>
                  </td>
                  <td class="text-end"><%=consulta("Pedido")%></td>
                  <td class="text-end"><%=consulta("Cliente")%></td>
                  <td><%=Trim(consulta("Razao")) %></td>
                  <td class="text-center"><%=Trim(consulta("Produto")) %></td>
                  <td><%=Trim(consulta("Referencia"))%></td>
                  <td><%=Trim(consulta("TIPO_PRODUTO"))%></td>
                  <td><%=Trim(tipoOnda)%></td>
                  <td class="text-end"><%=consulta("Quantidade") %></td>
                  <td class="text-end"><%=FormatNumber(precoUnitarioNet,6) %></td>
                  <td class="text-end"><%=FormatNumber(precoQuiloNet,5) %> </td>
                  <td class="text-end"><%=FormatNumber(AreaVendida,5) %> </td>
                  <td class="text-center"><%=consulta("DataEmissao") %></td>
                  <td class="text-center"><%=consulta("PrimeiraDataEntrega") %></td>
                  <td class="text-center"><%=descricaoStatus %></td>
                  <td><%=consulta("DescricaoBloqueio") %></td>
                  <td><%=consulta("QuemIncluiu") %></td>
                </tr>
<%
                response.Flush() 
                consulta.MoveNext() 
              wend
end if               
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