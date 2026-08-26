<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->
<%session.LCID = 1046 %>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <title>Smurfit Kappa</title>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
<!--
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
-->
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" />
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

  <script type="text/javascript">
    // --------------------------------------------------------------------------------------------
    function Ajax() {
      var xmlHttp = null;
      if (window.XMLHttpRequest) {
          xmlHttp = new XMLHttpRequest(); // code for IE7+, Firefox, Chrome, Opera, Safari
      } else if (window.ActiveXObject) {
          xmlHttp = new ActiveXObject("Microsoft.XMLHTTP"); // code for IE6, IE5
      } else {
          alert("Seu navegador não suporta AJAX!");
      }
      return xmlHttp;
    }
    // --------------------------------------------------------------------------------------------
    function DescartarSolicitacao(url) {
      var urlFixada='ListaExcluirSolicitacaoRepositorio.asp?idPedido=' + url
      if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp = new XMLHttpRequest();
      } else {// code for IE6, IE5
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
      }
      xmlhttp.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
          alert(xmlhttp.responseText);
          selecaoSolicitar();
        }
      }
      xmlhttp.open("GET", urlFixada, true);
      xmlhttp.send();
    }
    // --------------------------------------------------------------------------------------------       
    function selecaoSolicitar(){

      document.getElementById("TitleSelect").value = "Solicitações a Realizar";

      document.getElementById("grade").innerHTML = "";

      if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp=new XMLHttpRequest();
      } else {// code for IE6, IE5
        xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
      }

      xmlhttp.onreadystatechange=function(){
        if (this.readyState==4 && this.status==200){
          document.getElementById("grade").innerHTML = xmlhttp.responseText;
        }
      }

      xmlhttp.open("GET","ListaObterSolicitacao.asp?selecao=Solicitar",true);
      xmlhttp.send();
    }
    // --------------------------------------------------------------------------------------------
    function selecaoRealizadas(){
      document.getElementById("grade").innerHTML = "";
      if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp=new XMLHttpRequest();
      } else {// code for IE6, IE5
        xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
      }
      xmlhttp.onreadystatechange=function(){
        if (this.readyState==4 && this.status==200){
          //console.log(xmlhttp.responseText)
          document.getElementById("grade").innerHTML = xmlhttp.responseText;
        }
      }
      xmlhttp.open("GET","ListaObterSolicitacao.asp?selecao=Realizada",true);
      xmlhttp.send();
    }  
    // --------------------------------------------------------------------------------------------
    function selecaoDescartadas(){
      document.getElementById("grade").innerHTML = "";
      if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp=new XMLHttpRequest();
      } else {// code for IE6, IE5
        xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
      }
      xmlhttp.onreadystatechange=function(){
        if (this.readyState==4 && this.status==200){
          console.log(xmlhttp.responseText)
          document.getElementById("grade").innerHTML = xmlhttp.responseText;
        }
        //document.getElementById("grade").innerHTML = "";
      }
      xmlhttp.open("GET","ListaObterSolicitacao.asp?selecao=Descartadas",true);
      xmlhttp.send();
    } 
  </script>

  <style type="text/css">
      body{
          font-size:70% !important;
      }
      table{ font-size: 80% !important; }
      label{ font-size: 80% !important; }

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

      .tableFixHead          { 
        overflow: auto; 
        height: 65vh; 
      }

      .tableFixHead thead th { 
        position: sticky; /* Fixa o cabeçalho */
        top: 0; /* Fixa o cabeçalho no topo do contêiner */
        z-index: 1; /* Garante que o cabeçalho fique acima do conteúdo */
        background-color: #c0c0c0; /* Cor de fundo do cabeçalho */
        color: black; /* Cor do fonte do cabeçalho */
        font-weight: bold; /* Estilo do fonte no cabeçalho */
        font-size:10px;
        border: 1px solid #FFF;
        text-align:center;
      }

      /* Just common table stuff. Really. */
      table  { 
        border-collapse: collapse; 
        width: 100%; 
      }
      /*th, td { padding: 8px 16px; }*/
      th     { 
        background-color: #666666; 
        border-bottom: 0px;         
      }

      /* Para a segunda linha do cabeçalho */
      thead tr:nth-child(2) th {
        position: sticky; /* Fixa a segunda linha do cabeçalho */
        top: 24px; /* Ajuste para a altura da primeira linha do cabeçalho */
        z-index: 1; /* Garante que a segunda linha do cabeçalho fique acima do conteúdo */
      }

      .btn-mini {
        height: auto;
        line-height: 10px;
        padding-right: 2px;
      }

    </style>  
</head>
<%
SQL = "" &_
  "SELECT " &_
    "b.CRM_COD " &_
  "FROM LINKS a " &_
    "INNER JOIN MONTAMENU b ON a.LIN_COD = b.LIN_COD " &_
  "WHERE LIN_URL = '../comercial/bento/ListaSolicitacaoAlteracaoPedido' " &_
    "AND ISNULL(b.MON_D_E_L_E_T,'') = '' " &_
    "AND ISNULL(a.LIN_D_E_L_E_T,'') = '' " &_
    "AND b.CRM_COD = '" & CODUSER & "' " &_
    "AND b.FILIAL = '" & UNIDADE_FILIAL & "' "

SET PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )

if not isObject(PIR_HELPDESK) or (PIR_HELPDESK IS NOTHING) then
  response.end
end if

colunas = 3

if not PIR_HELPDESK.eof then

    alteraPedido = 0
    colunas = 3

end if

PIR_HELPDESK.Close()
SET PIR_HELPDESK = NOTHING

%>
  <body onload= "selecaoSolicitar()" >
    <div class="container">

      <div class="row">

        <div class="col">
          <div class="text-center">
            <h3>Solicitação de Alteração de Pedidos</h3>
          </div>
        </div>

        <div class="col-1" align="left">
          <a href="../../../manuais/manual_solicitacao_de_alteracao_pedido_bento.pdf" target="_blank" title="Manual" class="btn btn-info btn-sm">
            <i class="fas fa-question-circle"></i>
          </a>
        </div>

      </div>

    <div class="container-fluid">
      <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
        <div class="container-fluid" style="background-color: rgb(242, 242, 242);">

          <div class="row">
           
            <div class="col-md-<%=colunas %>">
              <a href="InclusaoSolicitacao.asp" class="btn btn-secondary"><i class="fas fa-plus"></i>&nbsp;Nova Solicitação</a>
            </div>
          
            <div class="col-md-<%=colunas %>">
              <input type="button" class="btn btn-primary" name="btnSolicitacaoARealizar" id="btnSolicitacaoARealizar" value="Solicitações a Realizar" onclick="selecaoSolicitar()"/>
            </div>

            <div class="col-md-<%=colunas %>">
              <input type="button" class="btn btn-success" name="btnSolicitacaoARealizar" id="btnSolicitacaoARealizar" value="Solicitações Realizadas" onclick="selecaoRealizadas()"/>
            </div>

            <div class="col-md-<%=colunas %>">
              <input type="button" class="btn btn-info" name="btnSolicitacaoARealizar" id="btnSolicitacaoARealizar" value="Solicitações Descartadas" onclick="selecaoDescartadas()"/>
            </div>

          </div>

          <div class="row">
           
            <div class="col">
              <h5 id="TitleSelect" name="TitleSelect"></h5>
            </div>

          </div>

          <form name="SolicitarAlteracao" id="SolicitarAlteracao" action="ListaPedidosValidar.asp" method="post">

            <input type="hidden" name="botaoClicado" id="botaoClicado" value="<%=botaoClicado%>">

            <div class="row" style="padding-top: 15px;">

              <div class="col-12" id="grade" name="grade">
              </div>

            </div>

          </form>

        </div>
      </div>
    </div>

  </body>

</html>