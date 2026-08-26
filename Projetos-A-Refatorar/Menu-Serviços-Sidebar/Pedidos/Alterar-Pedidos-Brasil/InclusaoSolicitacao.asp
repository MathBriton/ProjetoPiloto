<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../../cnn/cnn.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<%session.LCID = 1046 %>
<%
'<----- Expirando a página para não armazenar em cache
Response.Buffer=true
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires=Now()-1
Response.AddHeader "Cache-Control", "must-revalidate"
Response.AddHeader "Cache-Control", "no-cache"
Response.Addheader "Last-Modified:", Now()

Titulo="Solicitação de Alteração de Pedidos"
TituloBotao="Confirmar Solicitação"
%>

<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <title>Smurfit Kappa</title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
    <style>
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

          /* Add animation to "page content" */
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
    </style>
    <script type="text/javascript">
    var myVar;
    function AdicionaTemporizador() {
             document.getElementById("loader").style.display = "block";
             document.getElementById("MostraMensagem").style.display = "none";

             myVar = setTimeout(MostrarPagina, 3000);
    }

    function MostrarPagina() {
             document.getElementById("MostraMensagem").innerHTML="<h2>Pedido Localizado</h2>";
             if (document.getElementById("hidDados").value==""){
                 document.getElementById("MostraMensagem").innerHTML="<h2>Pedido não localizado</h2>";
             }

            document.getElementById("loader").style.display = "none";
            if(document.getElementById("txtPedido").value==""){
              document.getElementById("MostraMensagem").style.display = "none";
            }
            else{
              document.getElementById("MostraMensagem").style.display = "block";
            }
    }  

   function GravarSolicitacao(){
           
            if(document.getElementById("txtDe").value == ""){
              alert("Informação De não foi informada");
            }
            else if(document.getElementById("txtPara").value == ""){
                 alert("Informação Para não foi informada");
            }  
            else if(document.getElementById("txtPara").value == 0){
                 alert("Informação Para não foi informada");
            } 
            else if(true){
              var resposta = false
              var respostaStatus = false
              var itemSelecionado=document.getElementById("cboTipoAlteracao").value;
              switch (itemSelecionado) {
                     case '1': //Tipo de Frete
                          if(document.getElementById("txtDe").value == document.getElementById("txtPara").value){
                            resposta = true
                          }  
                          break;
                     case '2': //Preço
                          if(document.getElementById("txtDe").value == document.getElementById("txtPara").value.replace(".",",")){
                            resposta = true
                          } 
                          break;
                     case '3': //Quantidade
                          if(document.getElementById("txtDe").value == document.getElementById("txtPara").value.replace(",","").replace(".","")){
                            resposta = true
                          } 
                          break;
                     case '5': //Data de Entrega
                         if(document.getElementById("txtDe").value == document.getElementById("txtPara").value.split("-").reverse().join("/")){
                            resposta = true
                          }                             
                          break;
                     case '6': //Condiçoes de Pagamento
                          if(document.getElementById("hidIDCondicaoPagamento").value == document.getElementById("txtPara").value){
                            resposta = true
                          }  
                          break;
                     case '7': //Cancelar
                          if(document.getElementById("txtDe").value == document.getElementById("txtPara").value){
                            resposta = true
                          }  
                          break;
                     case '8': //Pedido do Cliente
                          if(document.getElementById("txtDe").value.toUpperCase() == document.getElementById("txtPara").value.toUpperCase()){
                            resposta = true
                          }  
                          break;
                     case '9': //Estatus
                          if(document.getElementById("txtDe").value.substring(0,1) == document.getElementById("txtPara").value){
                            resposta = true
                          }  
                          else if(document.getElementById("txtDe").value=="Atendido" && document.getElementById("txtPara").value=="Cancelado"){
                            respostaStatus = true
                            resposta = true
                          }
                          break;
                     case '14': //Preço Frete
                          if(document.getElementById("txtDe").value == document.getElementById("txtPara").value.replace(".",",")){
                            resposta = true
                          }  
                          break;               
               }
               if(resposta){
                    
                    if(respostaStatus){
                      alert("Pedido Atendido não pode ser Cancelado");
                    }
                    alert("Informações De e Para são idênticas.");
               }
               else{
                 Pedidos.submit()
            }

            }
    }  

    function VerificaPedidoDigitado(){
            if(document.getElementById("txtPedido").value != document.getElementById("hidNroPedido").value){
              document.getElementById("txtItem").value = ""
              document.getElementById("cboTipoAlteracao").text = ""
              document.getElementById("txtDescricaoProduto").value = ""
              document.getElementById("txtDe").value = ""
              if(document.getElementById("txtPara") != null){ 
                  document.getElementById("txtPara").value = ""
              }
              document.getElementById("txtObservacao").value = ""
            }  
    }

    function selecionaCampoTipoFrete(){
             var arrayDados=document.getElementById("hidDados").value.split("|");
             if (arrayDados[12]=="EX" && arrayDados[11]=="PA" && arrayDados[5]=="FOB" && document.getElementById("txtPara").value=="CIF"){
                var txt = document.createElement("input");
                txt.type="number";
                txt.classList.add("form-control");
                txt.name = "txtValorFrete";
                txt.id = "txtValorFrete";

                document.getElementById("tagValorFrete").appendChild(txt);
                document.getElementById("tagValorFrete").style="displayblock:";
             }
    }

    function selecionaEstatus(){
             if(document.getElementById("txtDe").value=="Atendido" && document.getElementById("txtPara").value=="C"){
                document.getElementById("btnConfirmar").style.visibility = "block";
                alert("Não pode proceder a troca dos Estatus");
             }
             else{
                 document.getElementById("btnConfirmar").style.visibility = "none";
                 document.getElementById("hidEstatus").value = document.getElementById("txtPara").value
              }
    }

    function VerificaStatusPedido(){
             if(document.getElementById("txtDe").value=="Atendido" && document.getElementById("txtPara").value=="Cancelado"){
                document.getElementById("btnConfirmar").style.visibility = "block";
                alert("Pedido Atendido não pode ser Cancelado");
             }
             else{
                 document.getElementById("btnConfirmar").style.visibility = "none";
                 document.getElementById("hidEstatus").value = document.getElementById("txtPara").value
              }
    }

    function criaEstatus(){
             var valoresDaOpcao = ["Atendido", "Normal"];

             var cbo = document.createElement("select");
             cbo.classList.add("form-select");
             cbo.name = "txtPara";
             cbo.id = "txtPara";
             cbo.setAttribute("onchange", "selecionaEstatus()");

             var option = document.createElement("option");
             option.value = "0";
             option.text = "Selecione...";
             cbo.appendChild(option);
             for (const opcao of valoresDaOpcao){
                   var option = document.createElement("option");
                   option.value = opcao.substring(0,1);
                   option.text = opcao;
                   cbo.appendChild(option);
              }

             document.getElementById("tagDinamica").appendChild(cbo);
    } 

    function criaTipoFrete(){
             var valoresDaOpcao = ["CIF", "FOB"];

             var cbo = document.createElement("select");
             cbo.classList.add("form-select");
             cbo.name = "txtPara";
             cbo.id = "txtPara";
             cbo.setAttribute("onchange", "selecionaCampoTipoFrete()");

             var option = document.createElement("option");
             option.value = "0";
             option.text = "Selecione...";
             cbo.appendChild(option);
             for (const opcao of valoresDaOpcao){
                   var option = document.createElement("option");
                   option.value = opcao;
                   option.text = opcao;
                   cbo.appendChild(option);
              }

             document.getElementById("tagDinamica").appendChild(cbo);
    }

   function criaCondicaoPagamento(){
             if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp=new XMLHttpRequest();
             }
             else {// code for IE6, IE5
                   xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
             }

             xmlhttp.onreadystatechange=function(){
               if (this.readyState==4 && this.status==200){
                   document.getElementById("hidCondicaoPagamento").value=xmlhttp.responseText;
                   console.log(document.getElementById("hidCondicaoPagamento").value);
                   
                   var condicoes=document.getElementById("hidCondicaoPagamento").value.split("|");
                   var cbo = document.createElement("select");
                   cbo.classList.add("form-select");
                   cbo.name = "txtPara";
                   cbo.id = "txtPara";

                   var contador=0;
                   var option = document.createElement("option");
                   option.value = "0";
                   option.text = "Selecione...";
                   cbo.appendChild(option);
                   for (const opcao of condicoes){
                        var quebra=opcao.split("-");

                        contador++; 
                        var option = document.createElement("option");
                        option.value = quebra[0];
                        option.text = quebra[1];
                        cbo.appendChild(option);
                   }

                   document.getElementById("tagDinamica").appendChild(cbo);
               }
             }

            xmlhttp.open("GET","ObterCondicaoPagamento.asp",true);
            xmlhttp.send();
    }

    function criaNumeros(){
             var txt = document.createElement("input");
             txt.type="number";
             txt.classList.add("form-control");
             txt.name = "txtPara";
             txt.id = "txtPara";

             document.getElementById("tagDinamica").appendChild(txt);
    }

    function criaTexto(){
             var txt = document.createElement("input");
             txt.type="text";
             txt.classList.add("form-control");
             txt.name = "txtPara";
             txt.id = "txtPara";

             document.getElementById("tagDinamica").appendChild(txt);
    }
    
    function criaCancelar(){
             var txt = document.createElement("input");
             txt.type="text";
             txt.classList.add("form-control");
             txt.name = "txtPara";
             txt.id = "txtPara";
             txt.readOnly="true";
             
             document.getElementById("tagDinamica").appendChild(txt);
             document.getElementById("txtPara").value='Cancelado';
    }

    function criaData(){
             var txt = document.createElement("input");
             txt.type="date";
             txt.classList.add("form-control");
             txt.name = "txtPara";
             txt.id = "txtPara";

             document.getElementById("tagDinamica").appendChild(txt);
    }
    function SelecaoNaCombo(){
             if (document.getElementById("hidDados").value==''){
               //alert("Pedido não foi localizado");
               return;
             }

             var arrayDados=document.getElementById("hidDados").value.split("|");
             console.log(arrayDados);

             document.getElementById("hidCliente").value=arrayDados[0];
             document.getElementById("hidNomeCliente").value=arrayDados[6];
             document.getElementById("hidIdSolicitacao").value=document.getElementById("cboTipoAlteracao").value;
             document.getElementById("hidTipoOperacaoId").value=document.getElementById("cboTipoAlteracao").value;
             document.getElementById("hidNroPedido").value=arrayDados[10];
             document.getElementById("txtDescricaoProduto").value=arrayDados[13];

             if (document.getElementById("txtPara")!=null){
                 const elemento = document.getElementById("txtPara");
                 elemento.remove();
             }

             console.log(document.getElementById("hidNroPedido").value);
             var itemSelecionado=document.getElementById("cboTipoAlteracao").value;
             switch (itemSelecionado) {
                     case '1': //Tipo de Frete
                         document.getElementById("txtDe").value=arrayDados[5];
                         criaTipoFrete();
                          break;
                     case '2': //Preço
                          document.getElementById("txtDe").value=arrayDados[9];
                          criaNumeros();
                          break;
                     case '3': //Quantidade
                          document.getElementById("txtDe").value=arrayDados[1];
                          criaNumeros();
                          break;
                     case '4': //Quantidade Pronto
                          document.getElementById("txtDe").value=arrayDados[2];
                          criaNumeros();
                          break;
                     case '5': //Data de Entrega
                          document.getElementById("txtDe").value= arrayDados[4];
                          criaData();                             
                          break;
                     case '6': //Condiçoes de Pagamento
                          document.getElementById("hidIDCondicaoPagamento").value=arrayDados[16];
                          document.getElementById("txtDe").value=arrayDados[7];
                          criaCondicaoPagamento();
                          break;
                     case '7': //Cancelar
                          document.getElementById("txtDe").value=arrayDados[14];
                          criaCancelar();
                          break;
                     case '8': //Pedido do Cliente
                          document.getElementById("txtDe").value=arrayDados[8];
                          criaTexto();
                          break;
                     case '9': //Estatus
                          document.getElementById("txtDe").value=arrayDados[14];
                          criaEstatus();
                          break;
                     case '14': //Preço Frete
                          document.getElementById("txtDe").value=arrayDados[15];
                          criaNumeros();                          
                          break;               
               }
    }

    function ValidaOPedido(nroPedido, nroItem){
             AdicionaTemporizador();

             console.log("Pedido: " + nroPedido);
             console.log("Item: " + nroItem);
             document.getElementById("hidDados").value="";
             if (nroPedido=="" && nroPedido==="undefined"){
                document.getElementById("cboTipoAlteracao").disabled=true;
                document.getElementById("txtItem").disabled=true;
                document.getElementById("txtObservacao").disabled=true;
                document.getElementById("txtPedido").innerHTML="";
                return;
             }

             if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp=new XMLHttpRequest();
             }
             else {// code for IE6, IE5
                  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
             }

             var flag=true
             xmlhttp.onreadystatechange=function(){
             if (this.readyState==4 && this.status==200){
                if (xmlhttp.responseText==="" || xmlhttp.responseText==="undefined"){
                   document.getElementById("cboTipoAlteracao").disabled=true;
                   document.getElementById("txtItem").disabled=true;
                   document.getElementById("txtObservacao").disabled=true;
                }
                else{
                    document.getElementById("cboTipoAlteracao").disabled=false;
                    document.getElementById("txtItem").disabled=false;
                    document.getElementById("txtObservacao").disabled=false;
                }

                document.getElementById("hidDados").value=xmlhttp.responseText;
                SelecaoNaCombo();
                console.log(document.getElementById("hidDados").value);
              }
              else{
                document.getElementById("cboTipoAlteracao").disabled=true;
                document.getElementById("txtItem").disabled=true;
                document.getElementById("txtObservacao").disabled=true;
                document.getElementById("hidDados").value=""
                
                return
              }
             }

             console.log(nroPedido);
             xmlhttp.open("GET","AlteracaoPedidoObterPedido.asp?nroPedido=" + nroPedido + "," + nroItem ,true);
             xmlhttp.send();
      } 

    function ObterCondicoesPagamento(){
             if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp=new XMLHttpRequest();
             }
             else {// code for IE6, IE5
                   xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
             }

             xmlhttp.onreadystatechange=function(){
               if (this.readyState==4 && this.status==200){
                   document.getElementById("hidCondicaoPagamento").value=xmlhttp.responseText;
                   console.log(document.getElementById("hidCondicaoPagamento").value);
               }
             }

            xmlhttp.open("GET","ObterCondicaoPagamento.asp",true);
            xmlhttp.send();
      } 
 </script>
  </head>
  <body onload="MostrarPagina()">
    <div class="container">
    <form class="row g-3 needs-validation" id="Pedidos" name="Pedidos" method="post" action="AlteracaoPedidoRepositorio.asp?tipo=SAP|0" novalidate>
    <input type="hidden" name="hidIDCondicaoPagamento" id="hidIDCondicaoPagamento" value=""/>
    <input type="hidden" name="hidCondicaoPagamento" id="hidCondicaoPagamento" value=""/>
    <input type="hidden" name="hidDados" id="hidDados" value=""/>
    <input type="hidden" name="hidCliente" id="hidCliente" value=""/>
    <input type="hidden" name="hidNomeCliente" id="hidNomeCliente" value=""/>
    <input type="hidden" name="hidTipoOperacaoId" id="hidTipoOperacaoId" value=""/>
    <input type="hidden" name="hidIdSolicitacao" id="hidIdSolicitacao" value=""/>
    <input type="hidden" name="hidNroPedido" id="hidNroPedido" value=""/>
    <input type="hidden" name="hidEstatus" id="hidEstatus" value=""/>
    <div>
        <p>Neste formulário você irá solicitar a alteração de um pedido ao departamento comercial. O pedido só será alterado com a solicitação feita por esta página, não sendo permitido por exemplo a solicitação de alteração por e-mail ou telefone. <br />
          Preencha os campos abaixo corretamente e clique no botão<strong> Solicitar Alteração</strong>.</p>
    </div>
      <div class="row">
          <div class="alert alert-dark" role="alert">
              <h3><strong><%=Titulo%></strong></h3>
       </div>

       <div class="row g-4">
        <div class="col-md-2">
          <label for="Pedido" class="form-label">Pedido</label>
          <input type="text" class="form-control readonly" id="txtPedido" name="txtPedido" aria-describedby="inputGroupPrepend3" onchange="VerificaPedidoDigitado()" />
          <div class="invalid-feedback">
            Informe o Pedido.
          </div>
        </div>
        <div class="col-md-2">
          <label for="txtItem" class="form-label">Item</label>
          <input type="number" class="form-control" id="txtItem" name="txtItem"  aria-describedby="inputGroupPrepend"  />
        </div> 
        
        <div class="col-md-4">
          <label for="cboTipoAlteracao" class="form-label">Tipo de Alteração</label>
          <select id="cboTipoAlteracao" name="cboTipoAlteracao" class="form-select">
             <option value="0" selected>Selecione...</option>
          <%
                  SQL = "SELECT "
            SQL = SQL & "   Id, "
            SQL = SQL & "   Operacao "
            SQL = SQL & "FROM InHistoricoTipoOperacao "
            SQL = SQL & "WHERE Id<=9 OR Id >13"

            Dados23.Open(SQL)

            do while not Dados23.EOF

              if Dados23("Operacao") <> "Quantidade Pronto"  then
          %>
             
            <option value="<%=Dados23("Id")%>"><%=Dados23("Operacao")%></option>
          <%
              end if

              Dados23.MoveNext()

            loop

            Dados23.Close()
            
          %>
          </select>
        </div> 
        <div class="col-md-2">
                <label for="btnPesquisar" name="LblPesquisar" id="LblPesquisar" class="form-label">Clique para Pesquisar</label> 
                <input type="button" class="btn btn-primary" name="btnPesquisar" id="btnPesquisar" value="Pesquisar" onclick="ValidaOPedido(txtPedido.value,txtItem.value)"/>
        </div>  
        <div id="loader"></div>
       </div> 

       <div class="row g-3">
        <div class="col-md-6">
          <label for="txtDescricaoProduto" class="form-label">Descrição do Produto</label>
          <input type="text" class="form-control" id="txtDescricaoProduto" name="txtDescricaoProduto" aria-describedby="inputGroupPrepend" readonly />
        </div> 

          <div class="col-md-2">
            <label for="txtDe" class="form-label">De</label>
            <input type="text" class="form-control" id="txtDe" name="txtDe" aria-describedby="inputGroupPrepend" readonly />
          </div> 

          <div class="col-md-3" id="tagDinamica" name="tagDinamica">
               <label for="txtPara" class="form-label">Para</label>
          </div>

          <div class="col-md-3" id="tagValorFrete" name="tagValorFrete" style="display:none">
               <label for="txtValorFrete" class="form-label">Valor Frete</label>
          </div>

        <div class="row g-2">
          <div class="col-md-10">
           <label for="txtObservacao" class="form-label">Observação</label>
           <input type="text" class="form-control" id="txtObservacao" name="txtObservacao" aria-describedby="inputGroupPrepend" disabled />
         </div>
        </div>

        <div style="display:none;" id="MostraMensagem" class="animate-bottom">

        </div>

        <div class="row g-3">
          <div class="col-12">
            <div class="form-group">
              <button type="button" id="btnConfirmar" name="btnConfirmar" class="btn btn-primary" onclick="GravarSolicitacao()"><%=TituloBotao%></button>
              <a href="ListaSolicitacaoAlteracaoPedido.asp" class="btn btn-primary">Voltar</a>
            </div>
          </div>     
        </div> 
      </form>
    </div>
  </body>   
</html>