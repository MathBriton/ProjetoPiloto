<!DOCTYPE html>
<!--#Include File = "../../cnn/cnn.asp"-->
<!--#Include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->

<%
Session.LCID = 1046 
Server.ScriptTimeout = 1000

if request("excel") = 1 then

  response.ContentType="application/vnd.ms-excel"
  response.AddHeader "content-disposition", "inline; filename=carteira.xls"

end if

%>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Relatório de Carteira</title>   
<%
    call ValidaPermissaoRelatorios("../comercial/relatorio/carteira",SESSION("CODUSER"),SESSION("FILIAL"))  
%>

    <link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>
    
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />


<!-- PADRAO JORGE.QUEIROZ -----------------------------------------------------------------------------------------------------
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <script src="https://cdn.jsdelivr.net/npm/jquery@3.5.0/dist/jquery.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
-->

    <script src="../../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
<!--    
    <script src="css/4.3.1/jquery.min.js"></script>
    <script src="css/4.3.1/popper.min.js"></script>
    <script src="js/bootstrap.min.js"></script>
-->

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>

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

      #imprimiAqui {
        display: none;
        text-align: center;
      }

      label { font-weight: bold; }

      .margimBody { margin-top:50px; }

    </style>

    <script type="text/javascript">

      // ------------------------------------------------------------------------------------------
      function imprimir(){
        window.print();
      }
      // ------------------------------------------------------------------------------------------
      function MostrarPagina() {
        document.getElementById("loader").style.display = "none";
        if(document.getElementById("flagTela").value==""){
            document.getElementById("imprimiAqui").style.display = "none";
        } else {
          document.getElementById("imprimiAqui").style.display = "block";
        }
      }     
      // ------------------------------------------------------------------------------------------
      function AdicionaTemporizador() {
        document.getElementById("loader").style.display = "block";
        document.getElementById("imprimiAqui").style.display = "none";
        myVar = setTimeout(MostrarPagina, 3000);
      }
      // ------------------------------------------------------------------------------------------
      function TratacboTipoPapelao(objeto){
                 var tipoProduto =objeto.value; 
                 if(tipoProduto=="1"){
                         //document.getElementById('cbotipoPapelao').setAttribute("readonly", "true");
                         document.getElementById('cbotipoPapelao').setAttribute("disabled", "true");
                 }else{
                         document.getElementById('cbotipoPapelao').removeAttribute("disabled", "");
                 }
        }
      // ------------------------------------------------------------------------------------------
      function somaComOQueTem(oQueTemNoArray, novoValor) {
        if(oQueTemNoArray===undefined)
          oQueTemNoArray=0;
        return oQueTemNoArray + novoValor;
      }
      // ------------------------------------------------------------------------------------------
      function myURL() {

        var urlProtocol = window.location.protocol;
        var urlHost = window.location.hostname;
        var urlFolder = "/corporativo/comercial/relatorio/";

        console.log( urlProtocol + "//" + urlHost + urlFolder );

        return urlProtocol + "//" + urlHost + urlFolder;

      }      
      // ------------------------------------------------------------------------------------------
      function geraRelatorio(){

        cliente         = document.getElementById("cliente").value;
        razao           = document.getElementById("razao").value;
        representante   = document.getElementById("Representante").value;
        tipoProduto     = parseFloat(document.getElementById("cbotipoProduto").value);
        tipoPapelao     = parseFloat(document.getElementById("cbotipoPapelao").value);
        cboPlanta       = document.getElementById("filial").value;
        dataIni         = document.getElementById("DataIniMes").value + document.getElementById("DataIniAno").value;
        cboClienteGrupo = document.getElementById("cboClienteGrupo").value;
        comOuSemImposto = document.getElementById("cboImpostos").value;
        cboValores      = document.getElementById("cboValores").value;

        if(cboPlanta=="0"){
          alert("Código da planta não foi informado!")
          return
        }

        if (dataIni==""){
          alert("Data inicial vazia!")
          return                
        }
        
        if(comOuSemImposto=="")
          comOuSemImposto="0"

        AdicionaTemporizador();
        debugger;
        var urlRelatorio= myURL() + "ResumoDeVendaCarteira_Repositorio.asp" +
                                      "?Filial=" + cboPlanta + 
                                      "&DataIni=" + dataIni + 
                                      "&cboClienteGrupo=" + cboClienteGrupo.trim() + 
                                      "&Representante=" + representante.trim() + 
                                      "&Cliente=" + cliente.trim() +
                                      "&Razao=" + razao.trim() +
                                      "&comOuSemImposto=" + comOuSemImposto + 
                                      "&tipoProduto=" + tipoProduto +
                                      "&tipoPapelao=" + tipoPapelao + 
                                      "&valores=" + cboValores + 
                                      "&retirahrf=0" +
                                      "&flagExecutaProc=1";

        window.location.href = urlRelatorio
      }
      // ------------------------------------------------------------------------------------------
      function imprima() {
        print();
      }
      // ------------------------------------------------------------------------------------------
      function doOnLoad() {
/*
        myCalendar = new dhtmlXCalendarObject(["DataIni"]);
        myCalendar.setDate(myCalendar.getDate());
        myCalendar.loadUserLanguage("br");
*/        
      }
      // --------------------------------------------------------------------------------------------
      function pesquisaCliente( obj ) {

        // console.log(obj.value.trim());
        // console.log(form.codcli.value);
        
        if (document.carteira.filial.value != '0') {

          if(obj.value.trim() != '' && obj.value.trim() != '') { 

            lnk = '../../ajax/validacliente_ajax.asp?planta=' + document.carteira.filial.value + '&cliente=' + obj.value + '&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Grupo';

            console.log(lnk);

            ObjAjax(lnk);

          } else {

            obj.value = ''; 
            //document.getElementById('codcli').value = ''; 
            //document.getElementById('razao').value = '' 
          }

        } else {

          alert('Selecione a filial!');

        }

      }    
      // --------------------------------------------------------------------------------------------
      function ConsultaCliente()
      {
        if (document.carteira.filial.value != '0') {

          lnk = '../../buscaCLIENTECorporativo.asp?id1=cliente&id2=lojacli&id3=razao&id6=' + document.carteira.filial.value;
          janela( lnk, 'Consulta Cliente', 'width=650,height=400,scrollbars=1' ); 
         
        } else {

          alert('Selecione a filial!');

        }         
      }      
    // --------------------------------------------------------------------------------------------
      function objReturn(obj) {

        if (obj.erro == 'ErroLogof') {

          alert('Favor se conectar novamente ao SWnet.\nSeu logon EXPIROU.');

          if (obj.tipo == 'Grupo') {

            document.getElementById('cliente').value = ''
            document.getElementById('lojacli').value = ''
            document.getElementById('razao').value = ''

          }

          return false;

        }

        if (obj.erro == 'ErroNEncontrado') {

          if (obj.tipo == 'Grupo') {

              document.getElementById('cliente').value = ''
              document.getElementById('lojacli').value = ''
              document.getElementById('razao').value = ''

              alert('CLIENTE NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente Bloqueado.\nCliente não pertence a sua carteira.\nCliente não Cadastrado.');

          }

          return false;

        }

        if (obj.erro == 'ErroCliente') {

          alert('Favor preencher os campos Cliente e Loja.');

          if (obj.tipo == 'Grupo') {

            document.getElementById('cliente').value = ''
            document.getElementById('lojacli').value = ''
            document.getElementById('razao').value = ''

          }

          return false;

        }

        // console.log(obj.tipo);
        // console.log(obj.val1);

        if (obj.tipo == 'Grupo') {

          // console.log(obj.val1);

          carteira.razao.value = obj.val1;

        }

      }         
    </script>

    <style>
      .margimBody {
        margin-top:10px;
      }
    </style>
  </head>

  <body class="body margimBody" onload="MostrarPagina()">

    <div class="container-fluid">
      <div class="panel panel-primary">
        <div class="panel-heading">
          <h3 class="text-center">Resumo de Vendas Carteira</h3>
        </div>

        <div class="panel-body">

          <form name="carteira" id="carteira" action="Carteira.asp" method="post" role="form" >

            <input type="hidden" name="flagTela" id="flagTela" value=""/>

            <div class="row">

              <div class="col-sm-2">

                <label for="filial" class="form-label">Filial</label>

                <select name="filial" id="filial" class="form-control">
                  <option value="0" >Selecione...</option>
                  <option value="00">Pirapetinga</option>
                  <option value="04">Uberaba</option>
                  <option value="05">Bento Gonçalves</option>
                  <option value="06">Maranguape</option>
                  <option value="07">Saquarema</option>
                  <option value="10">ESHOP</option>
                </select>

              </div>

              <div class="col-sm-1">

                <label for="DataIniMes" class="form-label">Mês</label>

                <select name="DataIniMes" id="DataIniMes" class="form-control">
                <option value="01" <% if month(now) = 1  then %> selected <% end if %> >Jan</option>
                <option value="02" <% if month(now) = 2  then %> selected <% end if %> >Fev</option>
                <option value="03" <% if month(now) = 3  then %> selected <% end if %> >Mar</option>
                <option value="04" <% if month(now) = 4  then %> selected <% end if %> >Abr</option>
                <option value="05" <% if month(now) = 5  then %> selected <% end if %> >Mai</option>
                <option value="06" <% if month(now) = 6  then %> selected <% end if %> >Jun</option>
                <option value="07" <% if month(now) = 7  then %> selected <% end if %> >Jul</option>
                <option value="08" <% if month(now) = 8  then %> selected <% end if %> >Ago</option>
                <option value="09" <% if month(now) = 9  then %> selected <% end if %> >Set</option>
                <option value="10" <% if month(now) = 10 then %> selected <% end if %> >Out</option>
                <option value="11" <% if month(now) = 11 then %> selected <% end if %> >Nov</option>
                <option value="12" <% if month(now) = 12 then %> selected <% end if %> >Dez</option>
                </select>     

              </div>              
              
              <div class="col-sm-1">

                <label for="DataIniAno" class="form-label">Ano</label>

                <select name="DataIniAno" id="DataIniAno" class="form-control">
                </select>     

                <script type="text/javascript">

                  var max = new Date().getFullYear(),
                  min = max - 4,
                  select = document.getElementById('DataIniAno');

                  for (var i = max; i >= min ; i--) {                   
                    var opc = document.createElement('option');
                    opc.value = i;
                    opc.innerHTML = i;
                    select.appendChild(opc);
                  }

                </script>

              </div>
<!--
              <div class="col-sm-2">

                <label for="cliente" class="form-label">Cliente Especifico</label>
                <input type="text" name="cliente" id="cliente" class="form-control" maxlength="6" />

              </div>
-->
              <div class="col-sm-2">

                <label for="cliente" class="form-label">Cliente</label>
                <div class="input-group">

                  <input type="text" name="cliente" id="cliente" class="form-control" value="<%=trim(request("cliente"))%>"  maxlength="6" onblur="pesquisaCliente( this )" style="width:170px" />
                  &nbsp;
                  <button type="button" class="btn btn-primary" onclick="ConsultaCliente()">
                    <i class="fa fa-search"></i>
                  </button>

                </div>

              </div>

              <div class="col-sm-3">
                <label for="razao" class="form-label">Nome</label>
                <input name="razao" id="razao" type="text" class="form-control" style="width:100%" disabled="disabled" value="<%=trim(request("razao"))%>" />
                <input type="hidden" name="lojacli" id="lojacli"/>
              </div>

              <div class="col-sm-2">

                <label for="cboClienteGrupo" class="form-label">Clientes Grupo</label>
                <select name="cboClienteGrupo" id="cboClienteGrupo" class="form-control">
                  <option value="T">Todos</option>
                  <option value="S">Somente Grupo</option>
                  <option value="N">Somente Clientes</option>
                </select> 

              </div>

            </div>

            <div class="row">

              <div class="col-sm-2">
                <label for="Representante" class="form-label">Representante</label>
                <input type="text" name="Representante" id="Representante" class="form-control" maxlength="6" />
              </div>

              <div class="col-sm-2">
                <label for="cbotipoProduto" class="form-label">Tipo de Produto</label>
                <select name="cbotipoProduto" id="cbotipoProduto" onchange="TratacboTipoPapelao(this)" class="form-control">
                <option value="0">Ambos</option>
                <option value="1">Papel</option>
                <option value="2">Papelão</option>
                </select>     
              </div>
              <div class="col-sm-2">
                <label for="cbotipoPapelao" class="form-label">Tipo de Papelão</label>
                <select name="cbotipoPapelao" id="cbotipoPapelao" class="form-control">
                    <option value="0">Ambos</option>
                    <option value="1" >Caixa</option>
                    <option value="2" >Chapa</option>
                    <option value="3" >Corte Vinco</option>
                    <option value="4" >Maleta</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="cboValores" class="form-label">Valores</label>
                <select name="cboValores" id="cboValores" class="form-control">
                <option value="0">Todos</option>
                <option value="1">Só Peso</option>
                <option value="2">Só M²</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="cboImpostos" class="form-label">Impostos</label>
                <select name="cboImpostos" id="cboImpostos" class="form-control">
                <option value="0">Com IPI</option>
                <option value="1">Sem IPI</option>
                <option value="2">Sem Imposto (ICM, PIS, COF)</option>
                </select>     
              </div>

              <div class="col-mD-3">

                <br>     
                &nbsp;
                <button type="button" onclick="geraRelatorio()" class="btn btn-success">Gerar</button>
                &nbsp;
                <button type="button" onclick="geraRelatorio()" class="btn btn-warning" style="display:none">Voltar</button>
                &nbsp;
                <a href="../../manuais/manual_resumo_venda_carteira.pdf" target="_blank" title="Manual" class="btn btn-info">
                  <i class="fas fa-question-circle"></i>
                </a>
                
            </div> 

          </form>
        </div>

        <div id="loader"></div>

        <div style="display:none;" class="animate-bottom" id="imprimiAqui" name="imprimiAqui"></div>

      </div>
    </div>

  </body>
</html>