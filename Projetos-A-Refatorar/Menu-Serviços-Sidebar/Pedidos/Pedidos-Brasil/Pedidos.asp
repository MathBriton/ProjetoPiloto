<!--#include File="../../funcoes/funcoes.asp"-->
<!--#include File="../../cnn2/cnn2.asp" -->
<!--#Include File="../../includes/variaveis.asp"-->
<!--#include File="includes/asp/PedidoFuncoesAsp.asp"-->
<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1043

Usuario = SESSION("IdUsuario")
Planta = request("paramPlanta")
if Planta = "" then 
   Planta = SESSION("FILIAL")
end if 
descricaoPlanta = fnomePlanta(Planta)
%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Emissão de Pedidos</title>
    <link href="../../../bootstrap/bootstrap.min.css" rel="stylesheet" />

    <script src="includes/js/PedidoFuncoes.js?v=<%=Now()%>"></script>
    <script src="../../funcoes/funcoesJAVASCRIPT.js?id=2345652"></script>
    <script src="../../../jquery//jquery-3.1.1.min.js"></script>
    <script src="../../../jquery/jquery.mask.min.js"></script>
    <script src="../../../bootstrap/bootstrap.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js"></script> 
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.13.4/jquery.mask.min.js"></script>
    <script>$(".readonly").keydown(function(e){e.preventDefault();});</script>  

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
      @-webkit-keyframes spin {
        0% { -webkit-transform: rotate(0deg); }
        100% { -webkit-transform: rotate(360deg); }
      }
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
      .animate-bottom {
        position: relative;
        -webkit-animation-name: animatebottom;
        -webkit-animation-duration: 1s;
        animation-name: animatebottom;
        animation-duration: 1s
      }

      @-webkit-keyframes animatebottom {
        from { bottom:-100px; opacity:0 } 
        to { bottom:0px; opacity:1 }
      }

      @keyframes animatebottom { 
        from{ bottom:-100px; opacity:0 } 
        to{ bottom:0; opacity:1 }
      }
      #MostraMensagem {
        display: none;
        text-align: center;
      }

      .unselectable{
        background-color: #ddd;
        cursor: not-allowed;
      }

      .selectable{
        background-color: #4682B4;
        color: white;
      }    

      .form-group {
          margin-bottom: 12px;
      }

      .form-group label {
          display: block;
          font-weight: bold;
          margin-bottom: 4px;
      }

      .roteiro-item {
          margin-bottom: 1.5rem;
      }

      .roteiro-header {
          display: flex;
          align-items: center;
          gap: 0.75rem;
          padding: 0.75rem;
          background-color: #f8f9fa;
          border: 1px solid #dee2e6;
          border-radius: 0.25rem;
      }

      .roteiro-header-info {
          display: flex;
          align-items: center;
          gap: 1rem;
          flex-grow: 1;
          font-size: 0.875rem;
      }

      .roteiro-divider {
          width: 1px;
          height: 1rem;
          background-color: #dee2e6;
      }

      .roteiro-label {
          color: #6c757d;
      }

      .roteiro-value {
          margin-left: 0.5rem;
          font-weight: 600;
      }

      .roteiro-code {
          font-family: monospace;
      }

      .roteiro-table {
          margin-top: 0.75rem;
          border: 1px solid #dee2e6;
          border-radius: 0.25rem;
          overflow: hidden;
      }

      .roteiro-table thead {
          background-color: #f8f9fa;
      }

      .roteiro-table th {
          color: #6c757d;
          text-transform: uppercase;
          font-weight: 600;
          font-size: 0.75rem;
          letter-spacing: 0.05em;
          padding: 0.5rem 0.75rem;
      }

      .roteiro-table td {
          padding: 0.75rem;
          font-size: 0.875rem;
      }

      .roteiro-badge {
		display: flex;
		justify-content: center;
		align-items: center;
		height: 25px;
		width: 25px;
		border: 1px solid #dee2e6;
		border-radius: 50%;
		font-family: monospace;
		font-size: 0.75rem;
		font-weight: normal;
      }

      .roteiro-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 0.5rem 0.75rem;
          background-color: rgba(248, 249, 250, 0.5);
          border-top: 1px solid #dee2e6;
          font-size: 0.75rem;
          color: #6c757d;
      }
    </style>
  </head>

  <body onload="TipoFrete();ParametrosFiliais('<%=Planta%>');MostrarPagina()">
    <div class="container-fluid">
      <span id="MSGPeso" style = "display:none"> </span>
<% if trim(strMsg) <> "" then %>
      <div class="alert alert-success">
        <a href="#" class="close" data-dismiss="alert" aria-label="close" title="close">×</a>
        <%=strMsg%>      
      </div>
<% end if %>
		  <form class="row g-3 needs-validation" id="Pedidos" name="Pedidos" method="post" action="PedidosRepositorio.asp" novalidate>
        <input type="hidden" name="Planta" id="Planta" value="<%=Planta%>" />
        <input type="hidden" name="hidFilial" id="hidFilial" value="<%=Planta%>" />
        <input type="hidden" name="hidProdutoEstoque" id="hidProdutoEstoque" value="" />
        <input type="hidden" name="hidTipoItem" id="hidTipoItem" value="" />
        <input type="hidden" name="Tes" id="Tes" value="" />
        <input type="hidden" name="TesRemessa" id="TesRemessa" value="" />
        <input type="hidden" name="SubOperacaoTes" id="SubOperacaoTes" value= "" />
        <input type="hidden" name="TotalPedido" id="TotalPedido" value= "0" />
        <input type="hidden" name="listaPedido" id="listaPedido" value=""/>
        <input type="hidden" name="FaturaRemessa" id="FaturaRemessa" value= "0"/>       
        <input type="hidden" name="Generica" id="Generica" value="0"/>
        <input type="hidden" name="hidCondicoesFaturamento" id="hidCondicoesFaturamento" value="0">
        <input type="hidden" name="hidEnderecoEntrega" id="hidEnderecoEntrega" value="0" />
        <input type="hidden" name="hidEnderecoCobranca" id="hidEnderecoCobranca" value="0" />
        <input type="hidden" name="hidDiasTransporte" id="hidDiasTransporte" value="" />
        <input type="hidden" name="hidArea" id="hidArea" value="0" />
        <input type="hidden" name="hidAreaRefile" id="hidAreaRefile" value="0" />
        <input type="hidden" name="hidUsuario" id="hidUsuario" value="<%=Usuario%>" />
        <input type="hidden" name="hidTipoFrete" id="hidTipoFrete" value="" />
        <input type="hidden" name="hidCargaMaquina" id="hidCargaMaquina" value="" />
        <input type="hidden" name="hidLimiteDeCredito" id="hidLimiteDeCredito" value="" />
        <input type="hidden" name="hidFlagDiasUtil" id="hidFlagDiasUtil" value="" />
        <input type="hidden" name="hidLarguraMin" id="hidLarguraMin" value="" />
        <input type="hidden" name="hidLarguraMax" id="hidLarguraMax" value="" />
        <input type="hidden" name="hidComprimentoMin" id="hidComprimentoMin" value="" />
        <input type="hidden" name="hidComprimentoMax" id="hidComprimentoMax" value="" />
        <input type="hidden" name="hidDiasAcrescimoPA" id="hidDiasAcrescimoPA" value="" />
        <input type="hidden" name="hidDiasAcrescimoCX" id="hidDiasAcrescimoCX" value="" />
        <input type="hidden" name="hidDiasAcrescimoCH" id="hidDiasAcrescimoCH" value="" />
        <input type="hidden" name="hidFlagDiasViagem" id="hidFlagDiasViagem" value="" />
        <input type="hidden" name="hidFlagDiaUtil" id="hidFlagDiaUtil" value="" />
        <input type="hidden" name="hidBloqueiaMaquinaCaixa" id="hidBloqueiaMaquinaCaixa" value="" />
        <input type="hidden" name="hidBloqueiaMaquinaChapa" id="hidBloqueiaMaquinaChapa" value="" />
        <input type="hidden" name="hidBloqueiaMaquinaCorteVinco" id="hidBloqueiaMaquinaCorteVinco" value="" />
        <input type="hidden" name="hidBloqueiaMaletaDE" id="hidBloqueiaMaletaDE" value="" />
        <input type="hidden" name="hidBloqueiaMaletaATE" id="hidBloqueiaMaletaATE" value="" />
        <input type="hidden" name="hidBloqueiaCVDE" id="hidBloqueiaCVDE" value="" />
        <input type="hidden" name="hidBloqueiaCVATE" id="hidBloqueiaCVATE" value="" />
        <input type="hidden" name="hidBloqueiaCHDE" id="hidBloqueiaCHDE" value="" />
        <input type="hidden" name="hidBloqueiaCHATE" id="hidBloqueiaCHATE" value="" />
        <input type="hidden" name="hidMtsCVmin" id="hidMtsCVmin" value="" />
        <input type="hidden" name="hidMtsCVmax" id="hidMtsCVmax" value="" />
        <input type="hidden" name="hidMtsCXmin" id="hidMtsCXmin" value="" />
        <input type="hidden" name="hidMtsCXmax" id="hidMtsCXmax" value="" />
        <input type="hidden" name="hidMtsCHmin" id="hidMtsCHmin" value="" />
        <input type="hidden" name="hidMtsCHmax" id="hidMtsCHmax" value="" />
        <input type="hidden" name="hidPrecoKgCorteVinco" id="hidPrecoKgCorteVinco" value="" />
        <input type="hidden" name="hidPrecoKgMaleta" id="hidPrecoKgMaleta" value="" />
        <input type="hidden" name="hidPrecoKgChapa" id="hidPrecoKgChapa" value="" />
        <input type="hidden" name="hidPeso" id="hidPeso" value="0" />
        <input type="hidden" name="hidPesoMN" id="hidPesoMN" value="" />
        <input type="hidden" name="hidPesoCV" id="hidPesoCV" value="" />
        <input type="hidden" name="hidPesoCH" id="hidPesoCH" value="" />
        <input type="hidden" name="hidTipoProduto" id="hidTipoProduto" value="" />
        <input type="hidden" name="hidComprimentoProduto" id="hidComprimentoProduto" value="0" />
        <input type="hidden" name="hidLarguraProduto" id="hidLarguraProduto" value="0" />
        <input type="hidden" name="hidTemTesRemessa" id="hidTemTesRemessa" value="0"/>
        <input type="hidden" name="hidDataMaior" id="hidDataMaior" value="0"/>
        <input type="hidden" name="hidNivel" id="hidNivel" value="<%=Nivel%>"/>
        <input type="hidden" name="hidPoder3" id="hidPoder3" value="N"/>
        <input type="hidden" name="hidTolMaior" id="hidTolMaior" value="0"/>
        <input type="hidden" name="hidTesPoder3" id="hidTesPoder3" value=""/>
        <input type="hidden" name="id" id="id" value="<%=id%>">
        <input type="hidden" name="hidSaldoAparas" id="hidSaldoAparas" value="0"/>
        <input type="hidden" name="hidTotalAparas" id="hidTotalAparas" value="0"/>
        <input type="hidden" name="hidAparasItem" id="hidAparasItem" value="0"/>
        <input type="hidden" name="hidModalFrete"  id="hidModalFrete"  value=""/>
        <input type="hidden" name="hidComExterior" id="hidComExterior" value=""/>
        <input type="hidden" name="hidAtualizaLotePiloto" id="hidAtualizaLotePiloto" value=""/>
        <input type="hidden" name="hidAtualizaLotePilotoSQL" id="hidAtualizaLotePilotoSQL" value=""/>
        <input type="hidden" name="hidValidaLimiteCredito" id="hidValidaLimiteCredito" value=""/>
        <input type="hidden" name="hidAtualizaLotePilotoProduto" id="hidAtualizaLotePilotoProduto" value=""/>

        <h1>Cadastro de Pedidos - <%=descricaoPlanta%></h1>
        <div class="divider">DADOS DO CLIENTE</div>
        <div class="row g-3">
          <div class="col-md-2">
            <label for="codCliente" class="form-label">Código Cliente</label>
            <div class="input-group">
              <input type="text" class="form-control readonly" id="codCliente" name="codCliente" aria-describedby="inputGroupPrepend3" readonly />
              <span class="input-group-text" style="cursor: pointer;" onclick="PesquisaCliente()" >
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#0d6efd" class="bi bi-search" viewBox="0 0 16 16">
                  <path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1.007 1.007 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z"/>
                  </svg>
                </span>
            </div>
              <div class="invalid-feedback">
                Informe o código do cliente.
              </div>
          </div>
          <div class="col-md-1">
              <label for="codLoja" class="form-label">Loja</label>
              <input type="text" class="form-control readonly" id="codLoja" name="codLoja" aria-describedby="inputGroupPrepend"  readonly/>
              <div class="invalid-feedback">
                Informe o código da loja.
              </div>
          </div>
          <div class="col-md-7">
            <label for="txtNomeCliente" class="form-label">Nome Cliente</label>
            <input type="text" class="form-control readonly" id="txtNomeCliente" name="txtNomeCliente" aria-describedby="inputGroupPrepend"  readonly/>
            <div class="invalid-feedback">
              Informe o nome do cliente.
            </div>         
          </div>
          <div class="col-md-2">
            <button type="button" class="btn btn-primary" style="margin-top: 32px" onclick="VerificaCreditoCliente('<%=Planta%>','<%=Usuario%>')" id="btnLimiteCredito">
              Limite de Crédito
            </button>
          </div>
        </div>
      
       <div class="row g-3 align-items-start">
        <div class="col-md-3">
          <label for="cboTES" class="form-label">TES</label>
          <div class="input-group">
            <input type="text" class="form-control readonly" id="cboTES" name="cboTES" readonly>
            <span class="input-group-text" style="cursor: pointer;" onclick="PesquisaTes('TES')">
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#0d6efd" class="bi bi-search" viewBox="0 0 16 16">
                <path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1.007 1.007 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z"/>
              </svg>
            </span>
          </div>
        </div>

        <div class="col-md-2" id="divCodClienteRemessa" style="display: none;">
          <label for="codClienteRemessa" class="form-group">Cliente Remessa</label>
          <div class="input-group">
            <input type="text" class="form-control readonly" id="codClienteRemessa" name="codClienteRemessa" readonly>
            <span class="input-group-text" style="cursor: pointer; display:none;" onclick="PesquisaClienteRemessa()" id="LupaClienteRemessa">
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#0d6efd" class="bi bi-search" viewBox="0 0 16 16">
                <path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1.007 1.007 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z"/>
              </svg>
            </span>
          </div>
        </div>

        <div class="col-md-2" id="divCodLojaRemessa" style="display: none;">
          <label for="codLojaRemessa" class="form-group">Cod Loja Remessa</label>
            <input type="text" class="form-control readonly" id="codLojaRemessa" name="codLojaRemessa" readonly>
        </div>

        <div class="col-md-3" id="divNomeClienteRemessa" style="display: none;">
          <label for="txtNomeClienteRemessa" class="form-group">Nome Cliente Remessa</label>
          <input type="text" class="form-control readonly" id="txtNomeClienteRemessa" name="txtNomeClienteRemessa" readonly>
        </div>

        <div class="col-md-2" id="divSaldoAparas" style="display: none;">
          <label for="txtSaldoAparas" class="form-group">Saldo de Aparas</label>
          <input type="text" class="form-control readonly" id="txtSaldoAparas" name="txtSaldoAparas" value="<%=txtSaldoAparas%>" readonly>
        </div>
      </div>
    
        <div class="row g-2">
          <div class="divider">ITENS DO PEDIDO</div>
          <div class="col-md-2">
              <label for="cboTipoItem" class="form-label">Tipo do Item</label>
                <select id="cboTipoItem" name="cboTipoItem" class="form-select" onchange="HabilitaItemPorTipo(this.value)" required>
                  <option value="00" selected>Selecione...</option>
                  <option value="CX">CAIXA</option>
                  <option value="CH">CHAPA</option>
                  <option value="CJ">CONJUNTO</option>
                  <%if Planta = "05" then%> 
                      <option value="PA">PAPEL</option>
                  <%end if%>    
                </select>
          </div>
          <div class="col-md-1">
            
			<div style="display: flex; align-items: center; gap: 10px;">
				<label for="codProduto" class="form-label">Produto</label>
				<span style="cursor: pointer" class="form-label" onclick="BuscarRoteiroProduto()">
					<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#0d6efd" class="bi bi-info-circle" viewBox="0 0 16 16">
						<path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z"/>
						<path d="m8.93 6.588-2.29.287-.082.38.45.083c.294.07.352.176.288.469l-.738 3.468c-.194.897.105 1.319.808 1.319.545 0 1.178-.252 1.465-.598l.088-.416c-.2.176-.492.246-.686.246-.275 0-.375-.193-.304-.533L8.93 6.588zM9 4.5a1 1 0 1 1-2 0 1 1 0 0 1 2 0z"/>
					</svg>
				</span>
			</div>
      <div class="input-group">
				<input type="text" class="form-control" id="codProduto" name="codProduto" value="<%=codProduto%>" aria-describedby="inputGroupPrepend3" readonly />
				<span class="input-group-text" style="cursor: pointer;"  onclick="PesquisaProduto()" >
					<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="#0d6efd" class="bi bi-search" viewBox="0 0 16 16">
						<path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1.007 1.007 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z"/>
					</svg>
				</span>
				<div class="invalid-feedback">
					Informe o código do Produto.
				</div>
              </div>
          </div>
          <div class="col-md-3">
            <label for="txtReferencia" class="form-label">Referência</label>
            <input type="text" class="form-control" id="txtReferencia" name="txtReferencia" value="<%=txtReferencia%>"  aria-describedby="inputGroupPrepend" readonly  />
          </div>
          <div class="col-md-1">
            <label for="txtEstrutura" class="form-label">Estrutura</label>
            <input type="text" class="form-control" id="txtEstrutura" name="txtEstrutura"   aria-describedby="inputGroupPrepend" readonly>
          </div>  
         
          <div class="col-md-1">
            <label for="txtQuantidade" class="form-label">Quantidade</label>
            <input type="number" class="form-control" id="txtQuantidade" name="txtQuantidade" pattern="[0-9]+$" value="<%=txtQuantidade%>" onblur="valorTotal()" aria-describedby="inputGroupPrepend" required="required" />
          </div>
          <div class="col-md-1">
            <label for="txtValorUnitario" class="form-label">Unitario</label>
            <input type="text" class="form-control" id="txtValorUnitario" name="txtValorUnitario" value="<%=txtValorUnitario%>" onblur="VerificaValorUnitario(this.value); Unitario(this.value,this);OrcamentoPreco('<%=planta%>', this.value,this); valorTotal()" aria-describedby="inputGroupPrepend" required="required" />
          </div>
          <div class="col-md-1">
            <label for="txtValorTotal" class="form-label">Total</label>
            <input type="text" class="form-control money" id="txtValorTotal" name="txtValorTotal" value="<%=txtValorTotal%>" aria-describedby="inputGroupPrepend" disabled  />
          </div>
            <div class="col-md-2">
            <label for="txtAreaTotal" class="form-label">Area Total</label>
            <input type="text" class="form-control money" id="txtAreaTotal" name="txtAreaTotal" value="<%=txtAreaTotal%>" aria-describedby="inputGroupPrepend" disabled  />
          </div>
        </div> 

        <div class="row g-3">
          <div class="col-md-1">
            <label for="txtLargura" class="form-label">Largura</label>
            <input type="number" class="form-control" id="txtLargura" name="txtLargura" value="<%=txtLargura%>" aria-describedby="inputGroupPrepend" disabled required="required" />
          </div>

          <div class="col-md-1">
            <label for="txtComprimento" class="form-label">Comprim.</label>
            <input type="number" class="form-control " id="txtComprimento" name="txtComprimento" value="<%=txtComprimento%>" aria-describedby="inputGroupPrepend" disabled required="required" />
          </div>

          <div class="col-md-2">
            <label for="txtVinco1" class="form-label">Vinco 1</label>
            <input type="number" class="form-control " id="txtVinco1" name="txtVinco1" value="<%=txtVinco1%>" aria-describedby="inputGroupPrepend"  disabled />
          </div>
          
          <div class="col-md-2">
            <label for="txtVinco2" class="form-label">Vinco 2</label>
            <input type="number" class="form-control " id="txtVinco2" name="txtVinco2" value="<%=txtVinco2%>" aria-describedby="inputGroupPrepend"  disabled />
          </div>

          <div class="col-md-2">
            <label for="txtVinco3" class="form-label">Vinco 3</label>
            <input type="number" class="form-control " id="txtVinco3" name="txtVinco3" value="<%=txtVinco3%>" aria-describedby="inputGroupPrepend" disabled />
          </div>
          <div class="col-md-2">
            <label for="txtVinco4" class="form-label">Vinco 4</label>
            <input type="number" class="form-control " id="txtVinco4" name="txtVinco4" value="<%=txtVinco4%>" aria-describedby="inputGroupPrepend"  disabled />
          </div>
          <div class="col-md-2">
            <label for="txtVinco5" class="form-label">Vinco 5</label>
            <input type="number" class="form-control " id="txtVinco5" name="txtVinco5" value="<%=txtVinco5%>" aria-describedby="inputGroupPrepend" disabled />
          </div>
        </div>
<%
    SQL = "Select DISTINCT Gramatura from dbo.E001_GN060E82 Where Gramatura > 0 order by Gramatura --Gramatura"
    Set Consulta = Conecta("BEN-SOLPEN", SQL )
    if vartype(Consulta) = 0 or vartype(Consulta) = 1 then
       response.end
    end if
%>      
        <div class="row g-3">
            <div class="col-md-1">
              <label for="cboGramatura" class="form-label">Gramatura</label>
                <select id="cboGramatura" name="cboGramatura" class="form-select" disabled required>
                  <option value="0">Selecione...</option>
                  <%while not Consulta.EOF%>
                  <option value="<%=Consulta("Gramatura")%>"><%=Consulta("Gramatura")%></option>
<%
                  Consulta.MoveNext()
                  Wend
                  Consulta.Close()
%>
                </select>
            </div>
            <div class="col-md-1">
              <label for="txtDiametro" class="form-label">Diametro</label>
              <input type="number" class="form-control" id="txtDiametro" name="txtDiametro" aria-describedby="inputGroupPrepend" disabled required />
            </div>
            <div class="col-md-2">
              <label for="txtLarguraPapel" class="form-label">Largura Bob.</label>
              <input type="number" class="form-control" id="txtLarguraPapel" name="txtLarguraPapel" aria-describedby="inputGroupPrepend" disabled required  />
            </div>
            <div class="col-md-1">
                <label for="cboTubete" class="form-label">Tubete</label>
                <select id="cboTubete" name="cboTubete" class="form-select" disabled required>
                  <option value="0" selected>Selecione...</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                </select>
            </div>
            <div class="col-md-2">
              <label for="cboCondicaoFaturamento" class="form-label">Condição Pagamento</label>
              <div id="divCondicaoFaturamento"></div>
          </div>

          <div class="col-md-2">
            <label for="cboMoeda" class="form-label">Moeda</label>
              <select id="cboMoeda" name="cboMoeda" class="form-select" required>
                <option value="0" selected>REAL</option>
                <option value="4">DOLAR</option>
              </select>
          </div>
          <div class="col-md-2">
            <label for="cboTipoFrete" class="form-label">Frete</label>
            <div id="divCboTipoFrete">
            </div>
          </div>
      </div>

      <div class="row g-3">

        <div class="col-1" data-provide="datepicker">
          <label for="dtpDataEntrega" class="form-label">Data de Entrega</label>
          <input type="date" class="form-control" name="dtpDataEntrega" id="dtpDataEntrega" onblur="VerificaDataEntrega()" required="required" />
        </div>
        
            <div class="col-1" data-provide="datepicker">
              <label for="dtpDataPrevistaEntrega" class="form-label">Previsão Entrega</label>
              <input type="date" class="form-control" name="dtpDataPrevistaEntrega" id="dtpDataPrevistaEntrega" />
            </div>
        

        <div class="col-md-1">
          <label for="txtPedidoCompra" class="form-label">Pedido de Compra</label>
          <input type="text" class="form-control" id="txtPedidoCompra" name="txtPedidoCompra" value="<%=txtPedidoCompra%>" aria-describedby="inputGroupPrepend" />
        </div> 

        <div class="col-md-2">
          <label for="txtPedidoCompraItem" class="form-label">Item Pedido de Compra</label>
          <input type="number" class="form-control" id="txtPedidoCompraItem" name="txtPedidoCompraItem" value="<%=txtPedidoCompraItem%>" min="1" max="100" onblur="verificaPedidoCompra('<%=Planta%>')" aria-describedby="inputGroupPrepend">
        </div> 

<% if (Planta = "05") then %>    

        <div class="col-md-2">
            <label for="cboModalFrete" class="form-label">Modal de frete</label>
            <select id="cboModalFrete" name="cboModalFrete" class="form-select">
              <option value=""  selected>Selecione...</option>
              <option value="AEREO"     >AÉREO</option>
              <option value="MARITIMO"  >MARÍTIMO</option>
              <option value="RODOVIARIO">RODOVIÁRIO</option>
            </select>
        </div>

        <div class="col-md-2">
            <label for="cboComExterior" class="form-label">Incoterms de Comércio Exterior</label>
            <select id="cboComExterior" name="cboComExterior" class="form-select">
              <option value="" selected>Selecione...</option>
              <option value="EXW">EXW</option>
              <option value="FCA">FCA</option>
              <option value="FAS">FAS</option>
              <option value="FOB">FOB</option>
              <option value="CFR">CFR</option>
              <option value="CIF">CIF</option>
              <option value="DPU">DPU</option>
              <option value="CPT">CPT</option>
              <option value="DAP">DAP</option>
              <option value="CIP">CIP</option>
              <option value="DDP">DDP</option>
            </select>
        </div>

<% end if %>

        <div class="col-md-3">
          <label for="" class="form-label">Lote piloto</label>
          <div class="input-group-prepend">
            <input type="checkbox" id="chkMarcarLotePiloto" name="chkMarcarLotePiloto">
            <label for="chkMarcarLotePiloto">Sim</label><br>
          </div>
        </div>  
    </div>
    <div class="row g-3">
      <div class="col-md-12">
        <label for="txtObservacao" class="form-label">Observação</label>
        <input type="text" class="form-control" id="txtObservacao" name="txtObservacao" value="<%=txtObservacao%>" aria-describedby="inputGroupPrepend" maxlength="100" />
      </div>
    </div>
    <div class="row g-3">    
      <div class="col-sm-1">
            <button type="button" name="botaoInserir" id="botaoInserir" style="width:100%" class="btn btn-primary" onclick="VerificaDados('<%=Planta%>')">Inserir</button>
      </div>            
      <div class="col-sm-1">
            <button type="button" class="btn btn-primary" id="gravar"  style="width:100%" onclick="GravarPedido()">Concluir Pedido</button>
      </div>            
      <div class="col-md-8">            
            <button type="button" class="btn btn-warning" style="display:none" id="mensagemLotePiloto">Esse pedido possui item com Lote Piloto, não poderá inserir mais itens. Grave o pedido!</button>
      </div>
    </div>
    <div name="temporizador" class="row g-3">
        <div id="loader"></div>
        <div style="display:none;" id="MostraMensagem" class="animate-bottom"></div>
    </div>
       <div class="divider">ITENS LANÇADOS NO PEDIDO</div>

          <table class="table table-bordered table-sm" id="gradePedido" name="gradePedido">
              <thead>
                  <tr>
                      <th>#</th>
                      <th>Produto</th>
                      <th>Referência</th>
                      <th>Qtd. Pedido</th>
                      <th>Área</th>
                      <th>Preço Unitário</th>
                      <th>Valor Total</th>
                      <th>O.C. Cliente</th>
                      <th>Seq. OC</th>
                      <th>Data Entrega</th>
                  </tr>
              </thead>

              <tbody>
              </tbody>
              <tfoot style="display:table-footer-group;">
                  <tr style="background-color:#dcdcdc;font-weight:bold;">
                      <td colspan="4" style="text-align:right;">TOTAIS:</td>
                      <td id="lblTotalArea" style="text-align:right;">0,00</td>
                      <td></td>
                      <td id="lblTotalValor" style="text-align:right;">0,00</td>
                      <td colspan="3"></td>
                  </tr>
              </tfoot>
              
          </table>
      </form>

      <div class="modal fade" id="ModalLimiteCredito" tabindex="-1" aria-labelledby="ModalLimiteCreditoLabel" aria-hidden="true">
		<div class="modal-dialog">
			<div class="modal-content">
				<div class="modal-header">
					<h5 class="modal-title">Limite de crédito</h5>
					<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
				</div>
				<div class="modal-body">
					<table class="table table-condensed">
						<thead>
						<tr>
							<th>Limite</th>
							<th>Aberto</th>
							<th>Pedidos em Carteira</th>
							<th>Pedidos a liberar</th>
							<th>Saldo</th>
						</tr>
						</thead>
					<tbody>
						<tr>
							<td id="Limite">&nbsp;</td>
							<td id="Aberto">&nbsp;</td>
							<td id="Carteira">&nbsp;</td>
							<td id="Liberar">&nbsp;</td>
							<td id="Saldo">&nbsp;</td>
						</tr>
					</tbody>
					</table>
				</div>
					<div class="modal-footer">
					<button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
				</div>	
			</div>
		</div>
      </div>

	<div class="modal fade" id="ModalRoteiroProduto" tabindex="-1" aria-labelledby="ModalRoteiroProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="ModalRoteiroProdutoLabel">Roteiro de Produção</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="conteudoRoteiro">
                        <div id="dadosRoteiro">
                        </div>
                    </div>
                    <div id="mensagemRoteiro" style="display:none;">
                        <div class="text-center py-4">
                            <div class="spinner-border text-primary" role="status">
                                <span class="visually-hidden">Carregando...</span>
                            </div>
                            <p class="mt-3 text-muted">Carregando informações...</p>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

	<template id="templateRoteiroItem">
		<div class="roteiro-item">
			<div class="roteiro-header">
				<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-secondary">
					<path d="M11 21.73a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73z"></path>
					<path d="M12 22V12"></path>
					<path d="m3.3 7 7.703 4.734a2 2 0 0 0 1.994 0L20.7 7"></path>
					<path d="m7.5 4.27 9 5.15"></path>
				</svg>
				<div class="roteiro-header-info">
					<div>
						<span class="roteiro-label">Referência:</span>
						<span class="roteiro-value data-referencia">-</span>
					</div>
					<div class="roteiro-divider"></div>
					<div>
						<span class="roteiro-label">Código:</span>
						<span class="roteiro-value roteiro-code data-codigo">-</span>
					</div>
				</div>
			</div>
			
			<div class="roteiro-table">
				<table class="table table-hover mb-0">
					<thead>
						<tr>
							<th style="width: 80px;">Etapa</th>
							<th style="width: 150px;">Código</th>
							<th>Máquina</th>
							<th style="width: 120px;">Status</th>
						</tr>
					</thead>
					<tbody class="data-maquinas">
					</tbody>
				</table>
				<div class="roteiro-footer">
					<span class="data-total-etapas">Total de etapas: 0</span>
				</div>
			</div>
		</div>
	</template>

	<template id="templateMaquinaRow">
		<tr>
			<td>
				<span class="roteiro-badge data-etapa">1</span>
			</td>
			<td>
				<code class="px-2 py-1 rounded bg-light data-codigo-maq">-</code>
			</td>
			<td>
				<span class="fw-medium data-nome-maq">-</span>
			</td>
			<td>
				<span class="data-status-maq">-</span>
			</td>
		</tr>
	</template>

    </div>
    <div>
    </div>
     <div class="modal fade" id="ModalSaldoAparas" tabindex="-1" aria-labelledby="ModalSaldoAparasLabel" aria-hidden="true">
      <div class="modal-dialog">
      <div class="modal-content">
      <div class="modal-header">
          <h5 class="modal-title">Saldo de Aparas</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
      </div>
      <div class="modal-body">
      <table class="table table-condensed">
      <thead>
      <tr>
          <th>Saldo de Aparas</th>
      </tr>
      </thead>
      <tbody>
      <tr>
          <td id="250">&nbsp;</td>
         
      </tr>
      </tbody>
      </table>
      </div>
      <div class="modal-footer">
      <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
      </div>
  </body>

  <script type="text/javascript">
      $(document).ready(function(){
        $('.money').mask("#.##0,00000", {reverse: true});
        $('txtValorUnitario').mask("#.##0,00000", {reverse: true});
        $('txtValorUnitario').on('blur').mask("#.##0,00000", {reverse: true});
      });
  
      $("#cboTipoItem").change(function () {      
          $("#txtLargura").attr("disabled", true);   
          $("#txtComprimento").attr("disabled", true);   
          $("#txtVinco1").attr("disabled", true);            
          $("#txtVinco2").attr("disabled", true);            
          $("#txtVinco3").attr("disabled", true);            
          $("#txtVinco4").attr("disabled", true);            
          $("#txtVinco5").attr("disabled", true);    
          $("#cboGramatura").attr("disabled", true);  
          $("#txtDiametro").attr("disabled", true);            
          $("#cboTubete").attr("disabled", true);   
          $("#txtLarguraPapel").attr("disabled", true);   

        if ($("#cboTipoItem").val() == "CH") {
          $("#txtLargura").attr("disabled", false);   
          $("#txtComprimento").attr("disabled", false);  
          $("#txtVinco1").attr("disabled", false);            
          $("#txtVinco2").attr("disabled", false);            
          $("#txtVinco3").attr("disabled", false);            
          $("#txtVinco4").attr("disabled", false);            
          $("#txtVinco5").attr("disabled", false);     
        }   
        else  if ($("#cboTipoItem").val() == "PA") {
          $("#cboGramatura").attr("disabled", false);  
          $("#txtDiametro").attr("disabled", false);            
          $("#cboTubete").attr("disabled", false);   
          $("#txtLarguraPapel").attr("disabled", false);   
        }   
      });

     var arrayPedido=[];
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
</script>
</html>