<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1046

%>
<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <!-- #include file="../../cnn/cnn.asp" -->
    <!-- #include file = "../../includes/variaveis.asp"-->
    <!-- #include file = "../../funcoes/funcoes.asp"-->
    <title>SKnet - Detalhes da validação do Pedido</title>
    <link href="css/bootstrap.min.css" rel="stylesheet" />
    <!-- <link href="css/cssPersonalizado/cadastroSetor.css" rel="stylesheet" /> -->
    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
    <script src="js/jquery-3.1.1.min.js"></script>
    <script src="js/jquery.mask.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js"></script> 
    <script>$(".readonly").keydown(function(e){e.preventDefault();});</script>
    <script src="js/bootstrap.min.js"></script>
    <script type="text/javascript">
            (function () {
            "use strict";
            window.addEventListener(
                "load",
                function () {
                var forms = document.getElementsByClassName("needs-validation");
                var validation = Array.prototype.filter.call(
                    forms,
                    function (form) {
                    form.addEventListener(
                        "submit",
                        function (event) {
                        if (form.checkValidity() === false) {
                            event.preventDefault();
                            event.stopPropagation();
                        }
                        form.classList.add("was-validated");
                        },
                        false
                    );
                    }
                );
                },
                false
            );
        })();

        function ValorSelecionado(valor)
        {
          switch (valor) {
            case 'P': hidBloqueadoPorPreco.value="P"; break;
            case 'D': hidBloqueadoPorData.value="D"; break;
            case 'L': hidBloqueadoLiberado.value="L"; break;
            case 'C': hidBloqueadoPorCondicoesPagamento.value="C"; break;                  
          }
        }
    </script>
    <style>
       .divider {
		  color: white;
		  background-color: #0D6EFD;
		  align-items: center;
		  justify-content: center;
		  display: flex;
		  margin-top: 25px;
		  border-radius: 3.5px;
      }
    </style>
  </head>
<%
Pedido = Request.QueryString("Pedido")

On Error GoTo 0
On Error Resume Next
' *** PEDIDO'

SQL =       " Select * from VW_CONSULTA_PEDIDOS_VALIDAR " 
SQL = SQL & " WHERE Pedido = '" & Pedido & "'" 
Dados181.Open(SQL)
if err.number <> 0 then
  call HandleError( err, GetFileName(), SQL )
end if

if not Dados181.eof then

    precoUnitarioNet = Dados181("Preco_Net")
    largura = Dados181("Largura")
    comprimento = Dados181("Comprimento")
    estrutura = Dados181("Estrutura")
    planta = Dados181("Planta")

    select case Dados181("Planta")
           case "00" Fabricacao="Pirapetinga" 
           case "04" descricaoPlanta="Uberaba" 
           case "05" descricaoPlanta="Bento Gonçalves" 
           case "06" descricaoPlanta="Fortaleza" 
           case "07" descricaoPlanta="Saquarema" 
    end select

    if Dados181("Planta") = "07" Then
       Fabricacao = "Saquarema"
    else
       Fabricacao = "MATRIZ"
    end If

    MarcarLotePiloto ="Não" 
    if trim(Dados181("LotePiloto")) = "1" Then
       MarcarLotePiloto ="Sim" 
    end if

    ' *** VALIDAÇÕES DO PRODUTO'
    gramatura =  Dados181("Gramatura")
    area =  Dados181("Area_Unitaria")
    peso = Dados181("PesoLiquido")
    precoQuiloNet =  Dados181("Preco_KG_Net")
    precoAreaNet = Dados181("Preco_M2_Net")
       

    ' *** VALIDA SITUACAO DO CLIENTE'
    if Dados181("Status_Cliente") <> "ATIVO" then
       response.write "<script type=""text/javascript"" charset=""utf-8"">"
       response.write "alert('O cliente esta em situação: "&Dados181("Status_Cliente")&".' );"
       response.write "location.href='ListaPedidosValidar.asp?Planta=07'"
       response.write "</script>"
       response.End
    end if

    ' *** VALIDA SITUACAO DO CLIENTE REMESSA'
    if Dados181("ClienteRemessa") <> "" then
         if  Dados181("Status_Cliente_Remessa") <> "ATIVO" then
                response.write "<script type=""text/javascript"" charset=""utf-8"">"
                response.write "alert('O cliente remessa esta em situação: "&Dados181("Status_Cliente_Remessa")&".' );"
                response.write "location.href='ListaPedidosValidar.asp?Planta=07'"
                response.write "</script>"
                response.End
        end if       
    end if
    ' *** '
    
     
    descricaotes = trim(Dados181("Tes"))
    if descricaotes = "" then 
        response.write "<script type=""text/javascript"" charset=""utf-8"">"
        response.write "alert('TES não encontrada.' );"
        response.write "location.href='ListaPedidosValidar.asp?Planta=07'"
        response.write "</script>"
        response.End
    end if
    ' ***'
    ChamaLimiteCreditoCliente Dados181("Cliente"), Dados181("Lojacli")

    '****** VALIDA CARGA MAQUINA
    ' Comentado conforme combinado entre Eduardo e Daniel. 07/07/2022
    'Call CargaMaquina(trim(Dados181("Produto")),largura,comprimento,estrutura,"","",Dados181("TipoFrete"),cliente,clienteRemessa ,Dados181("Quantidade"),invertedata(Dados181("PrimeiraDataEntrega")),pedido,planta)
    

else
    response.write "<script type=""text/javascript"" charset=""utf-8"">"
    response.write "alert('Pedido não encontrado.');"
    response.write "location.href='ListaPedidosValidar.asp?Planta=07'"
    response.write "</script>"
    response.end
end if
%>
    <body>
        <div class="container" style="margin-top: 10px">
            <form class="row g-2 needs-validation" id="detalhesValidacao" name="detalhesValidacao" method="post" action="detalhesValidacaoRepositorio.asp" novalidate>
                <input type="hidden" name="Planta" id="Planta" value="<%=planta %>" />
                <input type="hidden" name="Pedido" id="Pedido" value="<%=pedido%>">

                <div class="row">
                  <div class="col">
                    <center><h3 style="color: gray">Detalhes do Pedido</h3></center>
                  </div>
                  <hr>
                </div>
                
                <div class="row g-2">
                  <div class="col-2">
                    <label for="txtPedido" class="form-label">Número do Pedido</label>
                    <input type="text" class="form-control" name="txtPedido" id="txtPedido" value="<%=Dados181("Pedido")%>" disabled/>
                  </div>
                  <div class="col-1">
                  </div>
                  <div class="col-2">
                    <label for="txtEmissao" class="form-label">Emissão</label>
                    <input type="text" class="form-control" name="txtEmissao" id="txtEmissao" value="<%=Dados181("DataEmissao")%>" disabled/>
                  </div>
                  <div class="col-2">
                  </div>
                  <div class="col-2">
                    <label for="txtQuemEmitiu" class="form-label">Quem Emitiu</label>
                    <input type="text" class="form-control" name="txtQuemEmitiu" id="txtQuemEmitiu" value="<%=Dados181("QuemIncluiu")%>" disabled/>
                  </div>
                  <div class="col-1">
                  </div>
                  <div class="col-2">
                    <label for="txtOrdemDeCompra" class="form-label">Ordem de Compra</label>
                    <input type="text" class="form-control" name="txtOrdemDeCompra" id="txtOrdemDeCompra" value="<%=Dados181("NumeroOrdemCompra")%>" disabled/>
                  </div>
                </div>

                <div class="row g-2">
                    <div class="col-1">
                        <label for="txtClienteFatura" class="form-label">Codigo/Loja</label>
                        <input type="text" class="form-control" name="txtClienteFatura" id="txtClienteFatura" value="<%=Dados181("Cliente")%>-<%=Dados181("LojaCli")%>" disabled/>
                    </div>
                    <div class="col-5">
                        <label for="txtRazaoClienteFatura" class="form-label">Cliente Fatura</label>
                        <input type="text" class="form-control" name="txtRazaoClienteFatura" id="txtRazaoClienteFatura" value="<%=Dados181("Razao")%>" disabled/>
                    </div>
                    <div class="col-1">
                        <label for="txtClienteRemessa" class="form-label">Codigo/Loja</label>
                        <input type="text" class="form-control" name="txtClienteRemessa" id="txtClienteRemessa" value="<%=Dados181("ClienteRemessa")%>-<%=Dados181("LojaCliRemessa")%>" disabled/>
                    </div>
                    <div class="col-5">
                        <label for="txtRazaoClienteRemessa" class="form-label">Cliente Remessa</label>
                        <input type="text" class="form-control" name="txtRazaoClienteRemessa" id="txtRazaoClienteRemessa" value="<%=Dados181("Razao_Cliente_Remessa")%>" disabled/>
                    </div>
                </div>

                <div class="row g-2">
                    <div class="col-md-2">
                        <label for="txtProduto" class="form-label">Produto</label>
                        <input type="text" class="form-control" name="txtProduto" id="txtProduto" value="<%=Dados181("Produto")%>"  disabled/>
                    </div>
                    <div class="col-6">
                        <label for="txtReferencia" class="form-label">Referencia</label>
                        <input type="text" class="form-control" name="txtReferencia" id="txtReferencia" value="<%=Dados181("Referencia")%>"  disabled/>
                    </div>
                    <div class="col-md-2">
                        <label for="txtQuantidade" class="form-label">Quantidade</label>
                        <input type="text" class="form-control" name="txtQuantidade" id="txtQuantidade" value="<%=Dados181("Quantidade")%>"  disabled/>
                    </div>
                    <div class="col-2">
                        <label for="txtEntrega1" class="form-label">Entrega</label>
                        <input type="text" class="form-control" name="txtEntrega1" id="txtEntrega1" value="<%=Dados181("PrimeiraDataEntrega")%>" disabled/>
                    </div>
                </div>

                <div class="row g-2">
                    <div class="col-md-3">
                        <label for="txtLargura" class="form-label">Largura</label>
                        <input type="text" class="form-control" name="txtLargura" id="txtLargura" value="<%=Dados181("Largura")%>"  disabled/>
                    </div>
                    <div class="col-3">
                        <label for="txtComprimento" class="form-label">Comprimento</label>
                        <input type="text" class="form-control" name="txtComprimento" id="txtComprimento" value="<%=Dados181("Comprimento")%>"  disabled/>
                    </div>
                    <div class="col-2">
                        <label for="txtVincos" class="form-label">Vincos</label>
                        <input type="text" class="form-control" name="txtVincos" id="txtVincos" value="<%=Dados181("Vinco1")%>x<%=Dados181("Vinco2")%>x<%=Dados181("Vinco3")%>x<%=Dados181("Vinco4")%>x<%=Dados181("Vinco5")%>"  disabled/>
                    </div>
                    <div class="col-2">
                        <label for="txtDiametro" class="form-label">Diâmetro</label>
                        <input type="text" class="form-control" name="txtDiametro" id="txtDiametro" value="<%=Dados181("Diametro")%>"  disabled/>
                    </div>
                    <div class="col-2">
                        <label for="txtTubete" class="form-label">Tubete</label>
                        <input type="text" class="form-control" name="txtTubete" id="txtTubete" value="<%=Dados181("Tubete")%>"  disabled/>
                    </div>
                </div>
                <div class="row g-2">  
                    <div class="col-3">
                        <label for="txtComposicao" class="form-label">Composição</label >
                        <input type="text" class="form-control" name="txtComposicao" id="txtComposicao" value="<%=Dados181("Estrutura")%>" disabled/>
                    </div>
                    <div class="col-3">
                        <label for="txtGramatura" class="form-label">Gramatura</label >
                        <input type="text" class="form-control" name="txtGramatura" id="txtGramatura" value="<%=gramatura%>" disabled/>
                    </div>
                    <div class="col-3">
                        <label for="txtPeso" class="form-label">Área</label >
                        <input type="text" class="form-control" name="txtArea" id="txtArea" value="<%=FormatNumber(area,4)%>" disabled/>
                    </div>
                    <div class="col-3">
                        <label for="txtPeso" class="form-label">Peso</label >
                        <input type="text" class="form-control" name="txtPeso" id="txtPeso" value="<%=FormatNumber(peso,4)%>" disabled/>
                    </div>
                </div>
                <div class="row g-2">
                    <div class="col-md-3">
                        <label for="txtPreco" class="form-label">Preço Unitário</label>
                        <input type="text" class="form-control" name="txtPreco" id="txtPreco" value="<%=FormatNumber(Dados181("PrecoUnitario"),4)%>" disabled/>
                    </div>
                    <div class="col-md-3">
                        <label for="txtPreco" class="form-label">Preço Net</label>
                        <input type="text" class="form-control" name="txtPrecoNet" id="txtPrecoNet" value="<%=FormatNumber(precoUnitarioNet,4)%>" disabled/>
                    </div>
                    <div class="col-md-3">
                        <label for="txtPrecoKg" class="form-label">Preço Kg Net</label>
                        <input type="text" class="form-control" name="txtPrecoKg" id="txtPrecoKg" value="<%=FormatNumber(replace(precoQuiloNet,".",","),2)%>" readonly/>
                    </div>
                    <div class="col-md-3">
                        <label for="txtPrecoKg" class="form-label">Preço m² Net</label>
                        <input type="text" class="form-control" name="txtPrecoKg" id="txtPecoArea" value="<%=FormatNumber(replace(precoAreaNet,".",","),2)%>" readonly/>
                    </div>
                </div>
                <div class="row g-2">
                    <div class="col-3">
                        <label for="txtCodigoPagamento" class="form-label">Condição de Pagamento</label>
                        <input type="text" class="form-control" name="txtCodigoPagamento" id="txtCodigoPagamento" value="<%=Dados181("DescCondPagamento")%>" disabled/>
                    </div>
                    <div class="col-3">
                        <label for="txtTes" class="form-label">TES</label>
                        <input type="text" class="form-control" name="txtTes" id="txtTes" value="<%=descricaotes%>"  disabled/>
                        <div id="alertaTes" <% if Trim(MensagemAlertaTes) = "" then%>style="display: none;"<%end if %>>
                            <p><strong><%=MensagemAlertaTES%></strong></p>
                        </div>
                    </div>  
                    <div class="col-md-3">
                        <label for="txtLotePiloto" class="form-label">Lote Piloto</label>
                        <input type="text" class="form-control" name="txtLotePiloto" id="txtLotePiloto" value="<%=MarcarLotePiloto%>" disabled/>
                    </div>
                 </div> 
                <div class="row g-2">
                  <div class="col-md-12">
                    <label for="txtObservacao" class="form-label">Observação</label>
                    <input type="text" class="form-control" name="txtObservacao" id="txtObservacao" value="<%=Dados181("Observacao")%>" disabled/>
                  </div>
                </div>
  
                <div class="divider">Últimos pedidos validados</div>
                <div class="row g-2">
                  <p style="font-size:60%">
                    <strong>
                    (1) Preço pedido em avaliação<br />
                    (2) Preço Pedido histórico
                    </strong>
                  </p>
                  <table class="table table-bordered"> 
                    <thead>
                      <tr>
                        <th>Cliente</th>
                        <th>Cidade</th>
                        <th>Pedido</th>
                        <th>Referencia</th>
                        <th>Emissão</th>
                        <th>Entrega</th>
                        <th>TES</th>
                        <th>Qtde.</th>
                        <th>Preço</th>
                        <th>Preço/KG</th>
                        <th>Cond.Pagto.</th>
                        <th>%(2)/(1)</th>
                      </tr>
                    </thead>
                    <tbody>
<%
                    On Error GoTo 0
                    On Error Resume Next

                    SQL = "" &_
                      "SELECT TOP 10 * " &_
                      "FROM VW_CONSULTA_PEDIDOS_VALIDADOS " &_
                      "WHERE Cliente = '" & Dados181("Cliente") & "' " &_
                        "AND LojaCli = '" & Dados181("LojaCli")  & "' " &_
                        "AND Produto = '"& Dados181("Produto")  & "'" &_
                      "ORDER BY DataValidacao DESC " 
                  
                    Dados182.Open(SQL)

                    if err.number <> 0 then
                      call HandleError( err, GetFileName(), SQL )
                    end if 

                    do while not Dados182.EOF
                      gramaturaHistorico = Dados182("Gramatura")
                      pesoHistorico = Dados182("PesoLiquido")
                      areaHistorico = Dados182("Area_Unitaria")
                      Porcentagem=0
                      precoQuiloHistorico = Dados182("Preco_KG_Net")
                      AvaliacaoCusto = 0
                      if Dados182("PrecoUnitario") > 0 then
                          AvaliacaoCusto = (Dados182("PrecoUnitario")/Dados182("PrecoUnitario")-1)*100
                      end if

                      if AvaliacaoCusto > 0 then
                        classePedidoPreco = "table-danger text-danger"
                      else
                        classePedidoPreco = "table-success text-success"
                      end if 
%>
                      <tr>
                        <td><%=Dados182("Razao")%></td>
                        <td><%=Dados182("CIDADE")%></td>
                        <td><%=Dados182("Pedido")%></td>
                        <td><%=Dados182("Referencia")%></td>
                        <td><%=Dados182("DataEmissao")%></td>
                        <td><%=Dados182("PrimeiraDataEntrega")%></td>
                        <td class="text-end"><%=Dados182("Tes")%></td>
                        <td class="text-end"><%=Dados182("Quantidade")%></td>
                        <td class="text-end"><%=Dados182("PrecoUnitario")%></td>
                        <td class="text-end"><%=precoQuiloHistorico%></td>
                        <td class="text-end"><%=Dados182("DescCondPagamento")%></td>
                        <td class="text-end align-middle <%=classePedidoPreco %>"><strong><%=FormatNumber(AvaliacaoCusto,2)%></strong>%</td>
                      </tr>
<%
                      Dados182.MoveNext()
                    loop
                    Dados182.Close()
%>
                        </tbody>
                </table>
                </div>
  
                <div class="divider">Avaliação do pedido</div>
                <div class="row g-2">
                  <div class="col-4">
                    <label for="css">Selecione a ação desejada</label>
                    <select name="cboAvaliacao" id="cboAvaliacao" class="form-select" required="required" value="cboAvaliacao">
                      <option value="">Selecione...</option>
                      <option value="1">Liberado</option>
                      <option value="5">Bloqueado por condição de pagamento</option>
                      <option value="4">Bloqueado por data</option>
                      <option value="2">Bloqueado por preço</option>
                    </select>
                  </div>
                </div>
                <div class="col-12" style="margin-bottom:30px">
                  <div class="form-group">
                    <button type="submit" class="btn btn-primary">Confirmar</button>
                    <a href="ListaPedidosValidar.asp?Planta=07" class="btn btn-info">Voltar</a>
                  </div>
                </div>
            </form>
        </div>
<%
' -------------------------------------------------------------------------------------------------
function CargaMaquina(produto, largura, comprimento, estrutura, resinaExterna, resinaInterna, tipoFrete, CodCliente, clienteRemessa, quantidade, dataEntrega, pedido, planta)
  if resinaExterna="" then
    resinaExterna="NULL"
  end if   

  if resinaInterna="" then
    resinaInterna="NULL"
  end if 
    
  On Error GoTo 0
  On Error Resume Next

      '"," & Trim(largura) 
      '"," & Trim(comprimento)
      '",'" & Trim(estrutura) &"' "
      '"," & Trim(resinaExterna)
      '"," & Trim(resinaInterna)

  SQL = "" &_
    "SELECT " &_
      "DataDesejadaDisponivel, " &_
      "DataPossivel " &_
    "FROM dbo.fq_PreviaPedidos_InfoDataPossivel_Paema("  &_
      Trim(produto) &_
      ", NULL" &_
      ", NULL" &_
      ", NULL" &_
      ", NULL" &_
      ", NULL" &_
      ", " & Trim(tipoFrete) &_
      ", " & Trim(CodCliente) &_
      ", " & Trim(clienteRemessa) &_
      ", " & Trim(quantidade) &_
      ", '" & Trim(dataEntrega) & "' " &_
      ")"
  Dados91.open(SQL) 

  if err.number <> 0 then
    call HandleError( err, GetFileName(), SQL )
  end if     

  if Dados91.eof then
    DataDesejadaDisponivel =  0
    DataPossivel = ""
  else     
    if not Dados91("DataDesejadaDisponivel")  then
      DataDesejadaDisponivel =  0
      DataPossivel = Dados91("DataPossivel")
    else
      DataDesejadaDisponivel = 1
      DataPossivel = ""
    end if
  end if
  Dados91.close()

  if DataDesejadaDisponivel = 0 then
    response.write "<script type=""text/javascript"" charset=""utf-8"">"
    response.write "alert('Data de Entrega Indisponivel. Próxima Data Disponivel: "&DataPossivel&"\nO pedido será bloqueado por falta de disponibilidade de Carga Maquina.');"
    response.write "location.href='detalhesValidacaoRepositorio.asp?pedido="&pedido&"&cboAvaliacao=6&planta="&planta&"'"
    response.write "</script>"
    response.end        
  end if
end function
' -------------------------------------------------------------------------------------------------
function ChamaLimiteCreditoCliente(clienteLimiteCredito)
  byPass = 0
  matrizCliente = Array("76728","76946","75200","76927","76962","76527","76800","76930","75600","73910","74796","76722","76984","71242","2987","77020","77009","73693","76791","75353","77031")
  for each a in matrizCliente
    if Trim(a) = Trim(clienteLimiteCredito) then 
      byPass = 1
    end if
  next
  call LimiteCreditoCliente(byPass,clienteLimiteCredito)
End function
' -------------------------------------------------------------------------------------------------
function LimiteCreditoCliente(byPassaLimite,clienteLimiteCredito)
  
if byPassaLimite <> 1 then

    sql = "exec [spLimiteCredito] '"& Dados181("Cliente") &"', '"& Dados181("LojaCli") &"'"
    Dados161.source = sql 
    Dados161.open()

    sql = "Select Saldo From LimiteCredito Where CODCLI = '" & Dados181("Cliente") & "' and LOJACLI = '" & Dados181("LojaCli") &"'"
    Dados161.source = sql 
    Dados161.open()
   
    if err.number <> 0 then
      call HandleError( err, GetFileName(), SQL )
    end if 

    if not Dados161.eof then
      if Dados161("Saldo") < 0 then
        response.write "<script type=""text/javascript"" charset=""utf-8"">"
        response.write "alert('Limite de crédito indisponível para o cliente.\nO Pedido será bloqueado por falta de limite de crédito' );"
        response.write "location.href='detalhesValidacaoRepositorio.asp?pedido="&pedido&"&cboAvaliacao=3&planta="&planta&"'"
        response.write "</script>"
        response.End
      end if
    else
      response.write "<script type=""text/javascript"" charset=""utf-8"">"
      response.write "alert('Não foi possível verificar o limite de crédito do cliente.');"
      response.write "location.href='ListaPedidosValidar.asp?Planta=07'"
      response.write "</script>"
      response.end
    end if
    Dados161.close()
  end if
end function

%>
    </body>
</html>
