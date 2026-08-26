<!-- #include file = "../../cnn/cnn.asp" -->
<!-- #include file = "../../includes/variaveis.asp"-->
<!-- #include file = "../../funcoes/funcoes.asp"-->

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
    <title>SWnet - Detalhes da validação do Pedido</title>
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

        // ------------------------------------------------------------------------------------------
        function doOnLoad() {
          MostrarPagina();
        }  
        // ------------------------------------------------------------------------------------------
        function AdicionaTemporizador() {
          document.getElementById("loader").style.display = "block";
        }
        // ------------------------------------------------------------------------------------------
        function MostrarPagina() {
          document.getElementById("loader").style.display = "none";
        }
        // ------------------------------------------------------------------------------------------
        function ValidaForm() {
          document.getElementById('btnConfirma').disabled = true; 
          AdicionaTemporizador();
          detalhesValidacao.submit();
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

      #loader {
        position: absolute;
        left: 50%;
        top: 50%;
        z-index: 1;
        width: 120px;
        height: 120px;
        margin: -76px 0 0 -76px;
        border: 16px solid #f3f3f3;
        border-radius: 50%;
        border-top: 16px solid #3498db;
        -webkit-animation: spin 2s linear infinite;
        animation: spin 2s linear infinite;
      }
      
    </style>
  </head>
<%
Pedido = Request("Pedido")
PedidoValidar = request("Pedido")

On Error GoTo 0
On Error Resume Next
' *** PEDIDO'
SQL = "SELECT " &_
        "pd.*, " &_
        "ISNULL(UI.CRM_NREDUZ, UI.CRM_LOGIN) AS QuemIncluiu " & _
      "FROM Pedido pd " &_
        "LEFT JOIN BRPIR1M012TESTE.HELPDESK.dbo.CRM_DTI UI ON ((QuemIncluiu = UI.CRM_COD AND UI.EMP_CODIGO = '05' ) OR QuemIncluiu = UI.CRM_R_E_C_N_O) " & _
      "WHERE pd.Pedido = " & Pedido & " "
       
Dados132.Open(SQL)

if err.number <> 0 then
  call HandleError( err, GetFileName(), SQL )
end if

if not Dados132.eof then

    precoUnitarioNet = Dados132("PrecoUnitario")-Dados132("ValorICMSUnitario")-Dados132("ValorPISUnitario")-Dados132("ValorCOFINSUnitario")
    largura = Dados132("Largura")
    comprimento = Dados132("Comprimento")
    estrutura = Dados132("Estrutura")
    planta = Dados132("Planta")

    select case Dados132("Planta")
           case "00" Fabricacao="Pirapetinga" 
           case "04" descricaoPlanta="Uberaba" 
           case "05" descricaoPlanta="Bento Gonçalves" 
           case "06" descricaoPlanta="Fortaleza" 
    end select

    if Dados132("Planta") = "05" Then
       Fabricacao = "Bento Gonçalves"
    else
       Fabricacao = "MATRIZ"
    end If

    MarcarLotePiloto ="Não" 
    if trim(Dados132("LotePiloto")) = "1" Then
       MarcarLotePiloto ="Sim" 
    end if

    ' *** VALIDAÇÕES DO PRODUTO'
    peso = 0
    area = 0
    gramatura = 0
    tipoProduto = ""
    classePedido = ""

    On Error GoTo 0
    On Error Resume Next

    sqlProduto = "" &_
      "SELECT " &_
        "Codigo, " &_
        "Situacao, " &_
        "Classe, " &_
        "isnull(CodCliente,0) Cliente, " &_
        "isnull(area_liquida,0) Area, " &_
        "isnull(peso_liquido,0) Peso, " &_
        "Composicao, " &_
        "Gramatura, " &_
        "Largura, " &_
        "Comprimento " & _
      "FROM interface_inpanet_produtos " & _
      "WHERE Codigo = '" & trim(Dados132("Produto")) & "'"

    Dados101.Open(sqlProduto)

    if err.number <> 0 then

      call HandleError( err, GetFileName(), sqlProduto )

    end if

    if Dados101.EOF then

      response.write "<script type=""text/javascript"" charset=""utf-8"">"
      response.write "alert('Produto não encontrado');"
      response.write "location.href='ListaPedidosValidar.asp?Planta=05'"
      response.write "</script>"
      response.End

    else

      produtoCliente = Trim(Dados101("Cliente"))
      classePedido = Trim(Dados101("Classe"))

    end if

    Dados101.Close()

    On Error GoTo 0
    On Error Resume Next

    sql = "SELECT " & _
            "c.Composicao, " & _
            "AreaBrutaPecaComRefilos as Area, " & _
            "c.Gramatura, " & _
            "Item as Produto, " & _
            "c.Gramatura / 1000.000 * AreaBrutaPecaComRefilos as Peso " & _
          "FROM vw_ConsultaFTsPelaInpanet ft " & _
          "INNER JOIN vw_ConsultaComposicaoPelaInpaNet c on ft.Composicao = c.Composicao " & _
          "WHERE ft.Item = '"&trim(Dados132("Produto"))&"' AND AreaBrutaPecaComRefilos > 0 "
    Dados92.Open(sql)

    if err.number <> 0 then

      call HandleError( err, GetFileName(), SQL )

    end if

    if not Dados92.Eof then

        composicao = Dados92("Composicao")
        gramatura = Dados92("Gramatura")
        peso = Dados92("Peso")
        area = Dados92("Area")

    else

        On Error GoTo 0
        On Error Resume Next
        sql = "SELECT " & _
                "c.Composicao, " & _
                "AreaBrutaPecaComRefilos as Area, " & _
                "c.Gramatura, " & _
                "Item as Produto, " & _
                "c.Gramatura/1000.000*AreaBrutaPecaComRefilos as Peso " & _
              "FROM vw_ConsultaFTsPelaInpanet ft " & _
              "INNER JOIN vw_ConsultaComposicaoPelaInpaNet c ON ft.Composicao = c.Composicao " & _
              "WHERE ft.Conjunto = '"&trim(Dados132("Produto"))&"' AND AreaBrutaPecaComRefilos > 0 "
        Dados93.Open(sql)

        if err.number <> 0 then
          call HandleError( err, GetFileName(), SQL )
        end if

        if not Dados93.Eof then
            composicao = Dados93("Composicao")
            gramatura = Dados93("Gramatura")
            peso = Dados93("Peso")
            area = Dados93("Area")
        else
            composicao = ""
            gramatura = 0
            peso = 0
            area = 0
        end if
        Dados93.Close()

    End if
    Dados92.Close()
    
    if classePedido = "CH" and gramatura > 0 and Trim(Dados132("Estrutura")) <> "" then
        area = FormatNumber((Dados132("Largura")/1000) * (Dados132("Comprimento")/1000),4)
        peso = formatNumber(gramatura/1000*area,2)
    end if

    if classePedido = "PA" then
        area = 0
        peso = Dados132("Quantidade")
    end if

    if peso > 0 then
        precoQuiloNet = precoUnitarioNet/peso
    else
        precoQuiloNet = 0
    end if

    if area > 0 then
        precoAreaNet = precoUnitarioNet/area
    else
        precoAreaNet = 0
    end if

    On Error GoTo 0
    On Error Resume Next
    ' *** DESCRIÇÃO DA CONDICAO DE PAGAMENTO'
    DescricaoCondicaoPagamento = ""
    sql = "SELECT Descricao " & _
          "FROM interface_inpanet_condicao_pagamento " & _
          "WHERE Codigo = '" & trim(Dados132("CondicaoPagamento")) & "'"
    Dados101.Open(SQL)

    if err.number <> 0 then
      call HandleError( err, GetFileName(), SQL )
    end if

    if not Dados101.eof then
       DescricaoCondicaoPagamento = Dados101("Descricao")
    end if
    Dados101.Close()

    ' *** VALIDA SITUACAO DO CLIENTE'
    situacaoDoCliente = ""
    cliente = ""
    razaoCliente = ""

    On Error GoTo 0
    On Error Resume Next

    SQL = "SELECT " & _
            "Situacao, " & _
            "Cliente, " & _
            "Razao_social_cliente " & _
          "FROM interface_inpanet_clientes " & _
          "WHERE Cliente = '" & Dados132("Cliente") & "'"
    Dados101.Open(SQL)

    if err.number <> 0 then
      call HandleError( err, GetFileName(), SQL )
    end if    

    if not Dados101.eof then
        situacaoDoCliente = trim(Dados101("Situacao"))
        cliente = Dados101("Cliente")
        razaoCliente = Dados101("Razao_social_cliente")
    else
        response.write "<script type=""text/javascript"" charset=""utf-8"">"
        response.write "alert('Cliente não encontrado.');"
        response.write "location.href='ListaPedidosValidar.asp?Planta=05''"
        response.write "</script>"
        response.End
    end if
    Dados101.Close()

    if situacaoDoCliente <> "Ativo" then
       response.write "<script type=""text/javascript"" charset=""utf-8"">"
       response.write "alert('O cliente esta em situação: "&situacaoDoCliente&".' );"
       response.write "location.href='ListaPedidosValidar.asp?Planta=05''"
       response.write "</script>"
       response.End
    end if

    ' *** VALIDA SITUACAO DO CLIENTE REMESSA'
    situacaoDoClienteRemessa = ""
    clienteRemessa = 0
    razaoClienteRemessa = ""

    if Trim(Dados132("ClienteRemessa")) > 0 then

        On Error GoTo 0
        On Error Resume Next
        SQL = "SELECT " & _
                "Situacao, " & _
                "Cliente, " & _
                "Razao_social_cliente " & _
              "FROM interface_inpanet_clientes " & _
              "WHERE Cliente = '" & Dados132("ClienteRemessa") & "'"
        Dados101.Open(SQL)

        if err.number <> 0 then
          call HandleError( err, GetFileName(), SQL )
        end if 

        if not Dados101.eof then
            situacaoDoClienteRemessa = Trim(Dados101("Situacao"))
            clienteRemessa = Dados101("Cliente")
            razaoClienteRemessa = Trim(Dados101("Razao_social_cliente"))
        else
           response.write "<script type=""text/javascript"" charset=""utf-8"">"
           response.write "alert('Cliente remessa não encontrado.' );"
           response.write "location.href='ListaPedidosValidar.asp?Planta=05'"
           response.write "</script>"
           response.End
        end if
        Dados101.Close()
        if situacaoDoClienteRemessa <> "Ativo" then
           response.write "<script type=""text/javascript"" charset=""utf-8"">"
           response.write "alert('O cliente remessa esta em situação: "&situacaoDoClienteRemessa&".' );"
           response.write "location.href='ListaPedidosValidar.asp?Planta=05'"
           response.write "</script>"
           response.End
        end if
    end if
    ' *** '

    ' *** VALIDAÇÃO DAS TES'
    descricaotes = ""
    mensagemAlertaTes = ""
    descricaotes = ""

    On Error GoTo 0
    On Error Resume Next
    SQL = "SELECT " & _
            "Tipo_de_Operacao, " & _
            "Descricao " & _
          "FROM interface_inpanet_tes " & _
          "WHERE Codigo = " & Dados132("TESFatura")
    Dados101.Open(SQL)

    if err.number <> 0 then
      call HandleError( err, GetFileName(), SQL )
    end if 

    if not Dados101.eof then
        if trim(Dados101("Tipo_de_Operacao")) <> "V" then
            MensagemAlertaTES="Atenção para esta TES."
        end if
        descricaotes = trim(Dados101("Descricao"))
    else
        response.write "<script type=""text/javascript"" charset=""utf-8"">"
        response.write "alert('TES não encontrada.' );"
        response.write "location.href='ListaPedidosValidar.asp?Planta=05'"
        response.write "</script>"
        response.End
    end if
    Dados101.Close()
    ' ***'
    ChamaLimiteCreditoCliente Dados132("Cliente")

    '****** VALIDA CARGA MAQUINA
    ' Comentado conforme combinado entre Eduardo e Daniel. 07/07/2022
    if classePedido <> "PA" then
      'Call CargaMaquina(trim(Dados132("Produto")),largura,comprimento,estrutura,"","",Dados132("TipoFrete"),cliente,clienteRemessa ,Dados132("Quantidade"),invertedata(Dados132("PrimeiraDataEntrega")),pedido,planta)
    end if

else
    response.write "<script type=""text/javascript"" charset=""utf-8"">"
    response.write "alert('Pedido não encontrado.');"
    response.write "location.href='ListaPedidosValidar.asp?Planta=05'"
    response.write "</script>"
    response.end
end if
%>
  <body onload="doOnLoad()">

    <div class="container" style="margin-top: 10px">
      <form class="needs-validation" id="detalhesValidacao" name="detalhesValidacao" method="post" action="detalhesValidacaoRepositorio.asp" novalidate>
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
            <input type="text" class="form-control" name="txtPedido" id="txtPedido" value="<%=Dados132("Pedido")%>" disabled/>
          </div>
          <div class="col-1">
          </div>
          <div class="col-2">
            <label for="txtEmissao" class="form-label">Emissão</label>
            <input type="text" class="form-control" name="txtEmissao" id="txtEmissao" value="<%=Dados132("DataEmissao")%>" disabled/>
          </div>
          <div class="col-2">
          </div>
          <div class="col-2">
            <label for="txtQuemEmitiu" class="form-label">Quem Emitiu</label>
            <input type="text" class="form-control" name="txtQuemEmitiu" id="txtQuemEmitiu" value="<%=Dados132("QuemIncluiu")%>" disabled/>
          </div>
          <div class="col-1">
          </div>
          <div class="col-2">
            <label for="txtOrdemDeCompra" class="form-label">Ordem de Compra</label>
            <input type="text" class="form-control" name="txtOrdemDeCompra" id="txtOrdemDeCompra" value="<%=Dados132("NumeroOrdemCompra")%>" disabled/>
          </div>
        </div>

        <div class="row g-2">
          <div class="col-1">
            <label for="txtClienteFatura" class="form-label">Codigo</label>
            <input type="text" class="form-control" name="txtClienteFatura" id="txtClienteFatura" value="<%=cliente%>" disabled/>
          </div>
          <div class="col-5">
            <label for="txtRazaoClienteFatura" class="form-label">Cliente Fatura</label>
            <input type="text" class="form-control" name="txtRazaoClienteFatura" id="txtRazaoClienteFatura" value="<%=razaoCliente%>" disabled/>
          </div>
          <div class="col-1">
            <label for="txtClienteRemessa" class="form-label">Codigo</label>
            <input type="text" class="form-control" name="txtClienteRemessa" id="txtClienteRemessa" value="<%=clienteRemessa%>" disabled/>
          </div>
          <div class="col-5">
            <label for="txtRazaoClienteRemessa" class="form-label">Cliente Remessa</label>
            <input type="text" class="form-control" name="txtRazaoClienteRemessa" id="txtRazaoClienteRemessa" value="<%=razaoClienteRemessa%>" disabled/>
          </div>
        </div>

        <div class="row g-2">
          <div class="col-md-2">
            <label for="txtProduto" class="form-label">Produto</label>
            <input type="text" class="form-control" name="txtProduto" id="txtProduto" value="<%=Dados132("Produto")%>"  disabled/>
          </div>
          <div class="col-6">
            <label for="txtReferencia" class="form-label">Referencia</label>
            <input type="text" class="form-control" name="txtReferencia" id="txtReferencia" value="<%=Dados132("Referencia")%>"  disabled/>
          </div>
          <div class="col-md-2">
            <label for="txtQuantidade" class="form-label">Quantidade</label>
            <input type="text" class="form-control" name="txtQuantidade" id="txtQuantidade" value="<%=Dados132("Quantidade")%>"  disabled/>
          </div>
          <div class="col-2">
            <label for="txtEntrega1" class="form-label">Entrega</label>
            <input type="text" class="form-control" name="txtEntrega1" id="txtEntrega1" value="<%=Dados132("PrimeiraDataEntrega")%>" disabled/>
          </div>
        </div>

        <div class="row g-2">
          <div class="col-md-3">
            <label for="txtLargura" class="form-label">Largura</label>
            <input type="text" class="form-control" name="txtLargura" id="txtLargura" value="<%=Dados132("Largura")%>"  disabled/>
          </div>
          <div class="col-3">
            <label for="txtComprimento" class="form-label">Comprimento</label>
            <input type="text" class="form-control" name="txtComprimento" id="txtComprimento" value="<%=Dados132("Comprimento")%>"  disabled/>
          </div>
          <div class="col-2">
            <label for="txtVincos" class="form-label">Vincos</label>
            <input type="text" class="form-control" name="txtVincos" id="txtVincos" value="<%=Dados132("Vinco1")%>x<%=Dados132("Vinco2")%>x<%=Dados132("Vinco3")%>x<%=Dados132("Vinco4")%>x<%=Dados132("Vinco5")%>"  disabled/>
          </div>
          <div class="col-2">
            <label for="txtDiametro" class="form-label">Diâmetro</label>
            <input type="text" class="form-control" name="txtDiametro" id="txtDiametro" value="<%=Dados132("Diametro")%>"  disabled/>
          </div>
          <div class="col-2">
            <label for="txtTubete" class="form-label">Tubete</label>
            <input type="text" class="form-control" name="txtTubete" id="txtTubete" value="<%=Dados132("Tubete")%>"  disabled/>
          </div>
        </div>
        <div class="row g-2">  
          <div class="col-3">
            <label for="txtComposicao" class="form-label">Composição</label >
            <input type="text" class="form-control" name="txtComposicao" id="txtComposicao" value="<%=composicao%>" disabled/>
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
            <input type="text" class="form-control" name="txtPreco" id="txtPreco" value="<%=FormatNumber(Dados132("PrecoUnitario"),4)%>" disabled/>
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
            <input type="text" class="form-control" name="txtCodigoPagamento" id="txtCodigoPagamento" value="<%=DescricaoCondicaoPagamento%>" disabled/>
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
            <input type="text" class="form-control" name="txtObservacao" id="txtObservacao" value="<%=Dados132("Observacao")%>" disabled/>
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
          <div class="container-fluid">
            <table class="table table-bordered table-striped"> 
              <thead>
                <tr style="font-size:11px">
                  <th>Cliente</th>
                  <th>Cidade/UF</th>
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
          if Trim(Dados132("Estrutura")) = "" then

            On Error GoTo 0
            On Error Resume Next

            SQL = "" &_
            "SELECT top 10 " &_
              "p.Cliente " &_
              ", p.Razao_social_cliente " &_
              ", c.Nome_da_Cidade " &_
              ", c.UF " &_
              ", p.Pedido " &_
              ", Classe " &_
              ", p.Produto " &_
              ", p.Descricao as Referencia " &_
              ", Convert(varchar,p.Data_do_Pedido,101) as DataInclusao " &_
              ", p.DataEntrega as PrimeiraDataEntrega " &_
              ", p.TES as TESFatura " &_
              ", p.Qtde_Pedido as Quantidade " &_
              ", p.Preco_Unitario as PrecoUnitario " &_
              ", p.Peso as PesoLiquido " &_
              ", p.CondicaoPagamento " &_
              ", t.Descricao " &_
              ", p.Largura " &_
              ", p.Comprimento " &_
              ", p.Estrutura " &_
            "FROM interface_inpanet_pedidos p " &_
              "INNER JOIN interface_inpanet_tes t ON p.TES = t.codigo " &_
              "INNER JOIN interface_inpanet_clientes c ON p.Cliente = c.Cliente " &_
            "WHERE p.Produto = '" & Dados132("Produto") & "' " &_
              "AND p.Cliente  = '" & Dados132("Cliente") & "' " &_
            "ORDER BY DataInclusao DESC "
            Dados102.Open(SQL)

            if err.number <> 0 then
              call HandleError( err, GetFileName(), SQL )
            end if 

              do while not Dados102.EOF

                On Error GoTo 0
                On Error Resume Next

                SQL = "" &_
                  "SELECT " &_
                    "c.Composicao, " &_
                    "AreaBrutaPecaComRefilos as Area, " &_
                    "c.Gramatura, " &_
                    "Item as Produto, " &_
                    "c.Gramatura / 1000.000 * AreaBrutaPecaComRefilos as Peso " &_
                  "FROM vw_ConsultaFTsPelaInpanet ft " &_
                    "INNER JOIN vw_ConsultaComposicaoPelaInpaNet c ON ft.Composicao = c.Composicao " &_
                  "WHERE ft.Item = '"&trim(Dados102("Produto"))&"' " &_
                    "AND AreaBrutaPecaComRefilos > 0 "
                Dados92.Open(SQL)

                if err.number <> 0 then
                  call HandleError( err, GetFileName(), sql )
                end if 

                if not Dados92.Eof then
                    gramaturaHistorico = Dados92("Gramatura")
                    pesoHistorico = Dados92("Peso")
                    areaHistorico = Dados92("Area")
                else
                    gramaturaHistorico = 0
                    pesoHistorico = 0
                    areaHistorico = 0
                end if
                Dados92.Close()

                if Dados102("Classe") = "CH" and gramaturaHistorico > 0 and Trim(Dados102("Estrutura")) <> "" then
                  areaHistorico = FormatNumber((Dados102("Largura")/1000) * (Dados102("Comprimento")/1000),4)
                  pesoHistorico = formatNumber(gramatura/1000*area,2)
                end if

                if classePedido = "PA" then
                  areaHistorico = 0
                  pesoHistorico = Dados102("Quantidade")
                end if

                Porcentagem=0
                if Dados102("PrecoUnitario") > 0 and pesoHistorico > 0 then
                  precoQuiloHistorico = FormatNumber(Dados102("PrecoUnitario") / pesoHistorico,5)
                else
                  precoQuiloHistorico = 0
                end if

                AvaliacaoCusto = 0
                if Dados132("PrecoUnitario") > 0 then
                  AvaliacaoCusto = (Dados102("PrecoUnitario")/Dados132("PrecoUnitario")-1)*100
                end if

                if AvaliacaoCusto > 0 then
                  classePedidoPreco = "table-danger text-danger"
                else
                  classePedidoPreco = "table-success text-success"
                end if 
%>
                <tr style="font-size:10px">
                  <td><%=Dados102("Razao_social_cliente")%></td>
                  <td><%=Dados102("Nome_da_Cidade") & "/" & Dados102("UF") %></td>
                  <td><%=Dados102("Pedido")%></td>
                  <td><%=Dados102("Referencia")%></td>
                  <td><%=AcertaData(Dados102("DataInclusao"))%></td>
                  <td><%=AcertaData(Dados102("PrimeiraDataEntrega"))%></td>
                  <td class="text-end"><%=Dados102("Descricao")%></td>
                  <td class="text-end"><%=Dados102("Quantidade")%></td>
                  <td class="text-end"><%=Dados102("PrecoUnitario")%></td>
                  <td class="text-end"><%=precoQuiloHistorico%></td>
                  <td class="text-end"><%=Dados102("CondicaoPagamento")%></td>
                  <td class="text-end align-middle <%=classePedidoPreco %>"><strong><%=FormatNumber(AvaliacaoCusto,2)%></strong>%</td>
                </tr>
<%
                Dados102.MoveNext()
              loop
              Dados102.Close()
            end if
%>
              </tbody>
            </table>
          </div>
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
            <button id="btnConfirma" name="btnConfirma" class="btn btn-primary" onclick="ValidaForm()" >Confirmar</button>
            <a href="ListaPedidosValidar.asp?Planta=05" class="btn btn-info">Voltar</a>
          </div>
        </div>
    </form>

    <div id="loader"></div>

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

    On Error GoTo 0
    On Error Resume Next
    SQL = "EXEC [spLimiteCredito] "&clienteLimiteCredito&" "
    Dados101.open(SQL)

    if err.number <> 0 then
      call HandleError( err, GetFileName(), sql )
    end if 

    SQL = "" &_
      "SELECT " &_
        "Aberto, " &_
        "Pedido, " &_
        "PedidoLiberar, " &_
        "Limite, " &_
        "Saldo " &_
      "FROM LimiteCredito " &_
      "WHERE Cliente = " & clienteLimiteCredito & " "	
    Dados101.open(SQL)

    if err.number <> 0 then
      call HandleError( err, GetFileName(), SQL )
    end if 

    if not Dados101.eof then
      if Dados101("Saldo") < 0 then
        response.write "<script type=""text/javascript"" charset=""utf-8"">"
        response.write "alert('Limite de crédito indisponível para o cliente.\nO Pedido será bloqueado por falta de limite de crédito' );"
        response.write "location.href='detalhesValidacaoRepositorio.asp?pedido="&pedido&"&cboAvaliacao=3&planta="&planta&"'"
        response.write "</script>"
        response.End
      end if
    else
      response.write "<script type=""text/javascript"" charset=""utf-8"">"
      response.write "alert('Não foi possível verificar o limite de crédito do cliente.');"
      response.write "location.href='ListaPedidosValidar.asp'"
      response.write "</script>"
      response.end
    end if
    Dados101.close()
  end if
end function

%>
    </body>
</html>
