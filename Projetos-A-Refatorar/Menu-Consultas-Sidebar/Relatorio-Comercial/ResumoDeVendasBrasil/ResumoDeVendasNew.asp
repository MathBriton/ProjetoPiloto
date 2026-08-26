<!DOCTYPE html>
<!--#include file = "../../cnn/cnn.asp"-->
<!--#include file = "../../funcoes/funcoes.asp"-->
<!--#include file = "../../includes/variaveis.asp"-->
<%
Session.LCID = 1046 
Server.ScriptTimeout = 1000
%>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet - Relatório de Resumo de Vendas</title>   
<%
    'call ValidaPermissaoRelatorios("../comercial/relatorio/carteira",SESSION("CODUSER"),SESSION("FILIAL"))  
%>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.3/umd/popper.min.js" integrity="sha384-vFJXuSJphROIrBnz7yo7oB41mKfc8JzQZiCq4NCceLEaO4IHwicKwpJf9c9IpFgh" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.2/js/bootstrap.min.js" integrity="sha384-alpBpkh1PFOepccYVYDB4do5UnbKysX5WZXm3XxPqe5iKTfUKjNkCk9SaVuEZflJ" crossorigin="anonymous"></script>

    <link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
<!--
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
-->
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
    <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
    <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
    
<!--
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />
-->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
    <script src="../../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
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

      /* Estilo para esconder elementos na impressão */
      @media print {
          @page {
              size: landscape; /* Define a orientação da página como paisagem */
          }
          body {
              margin: 0; /* Remove margens para melhor aproveitamento do espaço */
          }
          .nao-imprimir {
              display: none; /* Esconde elementos com a classe 'nao-imprimir' */
          }
      }

    </style>

    <script type="text/javascript">

      // ------------------------------------------------------------------------------------------
      function imprimir(){
        window.print();
      }
      //-------------------------------------------------------------------------------------------
      function MostrarPagina() {
        document.getElementById("loader").style.display = "none";
        if(document.getElementById("flagTela").value==""){
            document.getElementById("imprimiAqui").style.display = "none";
        } else {
          document.getElementById("imprimiAqui").style.display = "block";
        }
      }     
      //-------------------------------------------------------------------------------------------
      function AdicionaTemporizador() {
        document.getElementById("loader").style.display = "block";
        document.getElementById("imprimiAqui").style.display = "none";
        myVar = setTimeout(MostrarPagina, 3000);
      }
      //-------------------------------------------------------------------------------------------
      function TratacboTipoPapelao(objeto){
        var tipoProduto =objeto.value; 
        if(tipoProduto=="1"){
          //document.getElementById('cbotipoPapelao').setAttribute("readonly", "true");
          document.getElementById('cbotipoPapelao').setAttribute("disabled", "true");
        }else{
          document.getElementById('cbotipoPapelao').removeAttribute("disabled", "");
        }
      }
      //-------------------------------------------------------------------------------------------
      function trataCampoSoma(valor){
        console.log(valor);
        if(valor=="0")
          return "0";
        return parseFloat(valor.replace(".","").replace(",","."));
      }
      //-------------------------------------------------------------------------------------------
      function trataValoresTotalizadores(valor){
        if(valor=="0")
          return "0";
        return parseFloat(valor).toLocaleString('pt-BR', { minimumFractionDigits: 2 } );
      }        
      //-------------------------------------------------------------------------------------------       
      function formataCampo(valor){
        if(valor==0)
          return 0;        
        return valor.toFixed(2).replace(".",",");
      }
      //-------------------------------------------------------------------------------------------
      function somaComOQueTem(oQueTemNoArray, novoValor)
      {
        if(oQueTemNoArray===undefined)
          oQueTemNoArray=0;

        return oQueTemNoArray + novoValor   
      }
      //-------------------------------------------------------------------------------------------
      function gerarRelatorioAnalitico(codigoCliente) {

        codcli = document.getElementById("codcli").value
        representante = document.getElementById("Representante").value 
        tipoProdutoParaORelatorio = parseFloat(document.getElementById("cbotipoProduto").value)
        tipoProduto = parseFloat(document.getElementById("cbotipoProduto").value)
        tipoPapelao = parseFloat(document.getElementById("cbotipoPapelao").value)
        cboPlanta=document.getElementById("filial").value
        dataIni=document.getElementById("DataIni").value
        dataFim=document.getElementById("DataFim").value
        cboClienteGrupo=document.getElementById("cboClienteGrupo").value
        comOuSemImposto=document.getElementById("cboImpostos").value
        cboValores=document.getElementById("cboValores").value
        document.getElementById("flagTela").value="1"
        
        const element = document.getElementById("tabelaDeImpressao")
        element.remove();
        //document.getElementById("btnVoltar").style.display = " inline-block";

        //Tipo produto
        const enumTipoProdutoAmbos=0;
        const enumTipoProdutoPapel=1;
        const enumTipoProdutoPapelao=2;

        //Valores
        const enumValoresTodos=0
        const enumValoresSoPeso=1
        const enumValoresSoMetro=2

        const enumPeso=1;
        const enumArea=2;
        const enumValor=3;

        if(cboPlanta=="0"){
          alert("Código da planta não foi informado!")
          return
        }

        if (dataIni==""){
          alert("Data inicial vazia!")
          return                
        }

        if (dataFim==""){
          alert("Data final vazia!")
          return                
        }

        if(comOuSemImposto=="")
          comOuSemImposto="0"

        retiraOPeso=0;

        if (parseFloat(cboValores)!=enumValoresTodos) {
          //quando for somente peso, retira a area
          switch (parseFloat(cboValores)) {
          case enumValoresSoPeso:
            tipoProduto=enumTipoProdutoPapel;    
            break;
          case enumValoresSoMetro:
            if(tipoProduto==enumTipoProdutoPapel){
              alert("Campo tipo de produto e Valores estão divergentes!")
              return                
            } else
                retiraOPeso=1;  //igual a 1 eh para retirar o peso
            break;
          }
        }

        if (tipoProduto!=enumTipoProdutoAmbos) {

          switch (parseFloat(tipoProduto)) {
          case enumTipoProdutoPapel:

            if(parseFloat(cboValores)==enumValoresSoPeso) 
              tipoProduto=enumTipoProdutoPapel;    
            else if(parseFloat(cboValores)==enumValoresSoMetro) {
              alert("Campo tipo de produto e Valores estão divergentes!")
              return
            } else
              cboValores = "1";
            break;

          case enumTipoProdutoPapelao:

            if(parseFloat(cboValores)==enumValoresSoPeso)
              tipoProduto=enumTipoProdutoPapel;  
            else if(parseFloat(cboValores)==enumValoresSoMetro)
              retiraOPeso = 1;  
            else
              cboValores = "1";
            break;

          }
        }

        //Gera as datas do relatorio
        if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
          xmlhttp=new XMLHttpRequest();
        } else {// code for IE6, IE5
          xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
        }
                
        xmlhttp.onreadystatechange=function() {

          var quantidadeDeColuna=4;

          if (this.readyState==4 && this.status==200) {

            var resultadoDatas=xmlhttp.responseText.split('_')[0]
            var resultadoDados=xmlhttp.responseText.split('_')[1]
            var arrayDatas=resultadoDatas.split(';');

            console.log("DATA: ",resultadoDatas)
            console.log("DADOS: ",resultadoDados)

            if(retiraOPeso==1) {

              quantidadeDeColuna=3

            } else {

              if(tipoProduto==enumTipoProdutoPapel)
                quantidadeDeColuna=3

            }

            tabelaCriada = "<div id='tabelaContainer'>";
            tabelaCriada = tabelaCriada + '<table class="table table-bordered table-striped" id="tabelaResumo">';
            tabelaCriada = tabelaCriada + '<thead>';
            tabelaCriada = tabelaCriada + '<tr class="table-active">';
            tabelaCriada = tabelaCriada + '<th colspan="3" class="ltb" style="text-align:center;background-color:#FFF"></th>';

            //Acrescenta os meses ao array de campos
            for (var coluna = 0; coluna < arrayDatas.length; coluna++) {

              tabelaCriada = tabelaCriada + `<th colspan="${quantidadeDeColuna}" class="ltb" style="text-align:center;background-color:#EEE">${arrayDatas[coluna]}</th>`;

            }

            //Array que controla a soma do rodape
            var arrayValoresRodape=Array(arrayDatas.length*(quantidadeDeColuna));

            tabelaCriada=tabelaCriada + '<th colspan="4" class="ltb" style="text-align:center;background-color:#CCC">TOTAL</th>';
            tabelaCriada=tabelaCriada + '</tr>'; 
            tabelaCriada=tabelaCriada + '<tr>';
            tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Representante</th>';
            tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Código</th>';
            tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Loja</th>';   
                  
            //Acrescenta os meses ao array de campos       
            for (var coluna = 0; coluna < arrayDatas.length; coluna++) {
              
              if (arrayDatas[coluna] != "") {

                if(retiraOPeso!=1)
                  tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Peso</th>';
                
                if(tipoProduto!=enumTipoProdutoPapel)
                  tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">M2</th>';

                tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Valor</th>';
                tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Prc</th>';

              }
            }

            if(retiraOPeso!=1)
              tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Peso</th>';

            if(tipoProduto!=enumTipoProdutoPapel) 
              tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">M2</th>';

            tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Valor</th>';
            tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Prc</th>'; 
            tabelaCriada=tabelaCriada + '</tr>'; 
            tabelaCriada=tabelaCriada + '</thead>'; 

            var cadaRegistro=resultadoDados.split(";");

            for (var cada = 0; cada < cadaRegistro.length; cada++) {

              tabelaCriada=tabelaCriada + '<tr>';
              var campos=cadaRegistro[cada].split("|");
              var totalPeso=0;
              var totalValor=0;
              var totalArea=0;
              var totalLinhaPeso=0;
              var totalLinhaValor=0;
              var totalLinhaArea=0;
              var proximo=0;     
              //Loop para cada registro
                      
              for (var tdColuna = 0; tdColuna < campos.length; tdColuna++) { 

                if (tdColuna==3) {

                  //loop de meses do filtro do relatorio 
                  var quantidadeDeMeses=resultadoDatas.split(';');
                  var quebraPorMes=campos[tdColuna].split(":");
                  proximo=0;
                  posicaoDoArrayDeValoresRodape=0;

                  for (var _qtdeMeses = 0; _qtdeMeses < quantidadeDeMeses.length; _qtdeMeses++) {

                    //separação dos dados do mes
                    if (proximo<quebraPorMes.length)
                      var quebraPorBarra=quebraPorMes[proximo].split("/")

                    if (quantidadeDeMeses[_qtdeMeses]==quebraPorBarra[0]) {
                      
                      if (retiraOPeso!=1) {
                        // PESO -------------------------------------------------------------------
                        tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumPeso]}</td>`;
                        totalPeso+=trataCampoSoma(quebraPorBarra[enumPeso]);
                        totalLinhaPeso+=totalPeso;

                      }

                      if (tipoProduto!=enumTipoProdutoPapel) {
                        // AREA -------------------------------------------------------------------
                        tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumArea]}</td>`;
                        totalArea+=trataCampoSoma(quebraPorBarra[enumArea]);
                        totalLinhaArea+=totalArea;

                      }

                      // VALOR --------------------------------------------------------------------
                      tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumValor]}</td>`;
                      totalValor+=trataCampoSoma(quebraPorBarra[enumValor]);
                      totalLinhaValor+=totalValor;
                  
                      // PRECO --------------------------------------------------------------------
                      var calculaPRC = trataCampoSoma(quebraPorBarra[enumValor]) / (parseInt(cboValores) != enumValoresSoMetro?trataCampoSoma(quebraPorBarra[enumPeso]):trataCampoSoma(quebraPorBarra[enumArea]))
                          
                      tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(calculaPRC)}</td>`;
                        
                      if(retiraOPeso!=1){

                        arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalPeso);
                        posicaoDoArrayDeValoresRodape++

                      }

                      if(tipoProduto!=enumTipoProdutoPapel){

                        arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalArea);
                        posicaoDoArrayDeValoresRodape++

                      }

                      arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalValor);
                      posicaoDoArrayDeValoresRodape++
                      arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalValor/(totalPeso==0?1:totalPeso));
                      posicaoDoArrayDeValoresRodape++

                      proximo++
                      totalPeso=0;
                      totalArea=0;
                      totalValor=0;

                    } else {

                      if(retiraOPeso!=1)
                        tabelaCriada=tabelaCriada + '<td class="ltb" style="text-align:right;">0,000</td>';

                      if(tipoProduto!=enumTipoProdutoPapel)
                        tabelaCriada=tabelaCriada + '<td class="ltb" style="text-align:right;">0,000</td>';

                      posicaoDoArrayDeValoresRodape++;
                      arrayValoresRodape[posicaoDoArrayDeValoresRodape]=0;
                      posicaoDoArrayDeValoresRodape++;
                      arrayValoresRodape[posicaoDoArrayDeValoresRodape]=0;                                                
                      tabelaCriada=tabelaCriada + '<td class="ltb" style="text-align:right;">0,000</td>';
                      tabelaCriada=tabelaCriada + '<td class="ltb" style="text-align:right;">0,00</td>';

                    }

                  } 

                  // TOTAL PESO -------------------------------------------------------------------
                  if(retiraOPeso!=1)                                        
                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaPeso.toString())}</td>`;

                  // TOTAL AREA -------------------------------------------------------------------
                  if(tipoProduto!=enumTipoProdutoPapel)
                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaArea.toString())}</td>`;

                  // TOTAL VALOR ------------------------------------------------------------------
                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaValor.toString())}</td>`;

                  // TOTAL PRECO ------------------------------------------------------------------
                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${(totalLinhaValor/(parseInt(cboValores)!=enumValoresSoMetro?totalLinhaPeso:totalLinhaArea)).toFixed(2)}</td>`;

                  totalPeso=0;
                  totalValor=0;
                  totalArea=0;
                  posicaoDoArrayDeValoresRodape=0;
                  totalLinhaPeso=0;
                  totalLinhaValor=0;
                  totalLinhaArea=0;

                } else {

                  // REPRESENTANTE ----------------------------------------------------------------
                  if (tdColuna==0)
                    tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:left;">${campos[tdColuna]}</td>`;

                  // GRUPO ------------------------------------------------------------------------
                  if (tdColuna==1)
                    tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:center;">${campos[tdColuna]}</td>`;

                  // CLIENTE ----------------------------------------------------------------------
                  if (tdColuna==2)
                    tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:left;"><a href="#" onclick="gerarRelatorioAnalitico('${campos[tdColuna]}')"><spam>${campos[tdColuna]}</spam></a></td>`;

                }

              }

              tabelaCriada = tabelaCriada + '</tr>';

            }

            //Linha de total
            tabelaCriada = tabelaCriada + '<tr>';
            tabelaCriada = tabelaCriada + '<td colspan="3" class="ltb"></td>';
            contaPassos  = 0

            totalPRCPeso = 0
            totalPRCPreco= 0
            totalPRCArea = 0
            totalGeralRodapePeso = 0
            totalGeralRodapeArea = 0
            totalGeralRodapeValor= 0
            totalGeralRodapePr   = 0

            if (quantidadeDeColuna==4) {

              for (var _totalPorData = 0; _totalPorData < arrayValoresRodape.length; _totalPorData++) {

                if (contaPassos==quantidadeDeColuna-1) {

                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalPRCPreco/totalPRCPeso)}</td>`;
                  totalPRCPeso=0
                  totalPRCPreco=0

                } else {

                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(arrayValoresRodape[_totalPorData].toString())}</td>`;

                } 

                if (contaPassos==0) {

                  contaPassos++;
                  totalPRCPeso+=parseFloat(arrayValoresRodape[_totalPorData]);
                  totalGeralRodapePeso+=parseFloat(arrayValoresRodape[_totalPorData]);

                } else if(contaPassos==1) {

                  contaPassos++;
                  totalPRCArea+=parseFloat(arrayValoresRodape[_totalPorData]);
                  totalGeralRodapeArea+=parseFloat(arrayValoresRodape[_totalPorData]);

                } else if(contaPassos==2) {

                  contaPassos++;
                  totalPRCPreco+=parseFloat(arrayValoresRodape[_totalPorData]);
                  totalGeralRodapeValor+=parseFloat(arrayValoresRodape[_totalPorData]);

                } else if ( contaPassos < (quantidadeDeColuna-1) )

                  contaPassos++;

                else

                  contaPassos=0;

              }

            } else {
              // 3 colunas

              for (var _totalPorData = 0; _totalPorData < arrayValoresRodape.length; _totalPorData++) {

                if(contaPassos==quantidadeDeColuna-1) {
                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo((totalPRCPreco/(parseInt(cboValores)!=enumValoresSoMetro?totalPRCPeso:totalPRCArea)))}</td>`;
                  totalPRCArea=0
                  totalPRCPeso=0
                  totalPRCPreco=0
                  contaPassos++;
                } else {
                  tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(arrayValoresRodape[_totalPorData])}</td>`;
                } 

                if(contaPassos==0){

                  contaPassos++;

                  if(parseInt(cboValores)==enumValoresSoPeso){
                    totalPRCPeso+=parseFloat(arrayValoresRodape[_totalPorData]);
                    totalGeralRodapePeso+=parseFloat(arrayValoresRodape[_totalPorData]);
                  } else{ 
                    totalPRCArea+=parseFloat(arrayValoresRodape[_totalPorData]);
                    totalGeralRodapeArea+=parseFloat(arrayValoresRodape[_totalPorData]);
                  }

                } else if(contaPassos==1) {

                  contaPassos++;
                  totalPRCPreco+=parseFloat(arrayValoresRodape[_totalPorData]);
                  totalGeralRodapeValor+=parseFloat(arrayValoresRodape[_totalPorData]);

                } else if(contaPassos<(quantidadeDeColuna-1))

                  contaPassos++;

                else

                  contaPassos=0;

              }

            }

            if(retiraOPeso!=1)
              tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapePeso.toString())}</td>`;

            if(tipoProduto!=enumTipoProdutoPapel) 
              tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapeArea.toString())}</td>`;

            tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapeValor.toString())}</td>`;
            tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalGeralRodapeValor/(parseInt(cboValores)!=enumValoresSoMetro?totalGeralRodapePeso:totalGeralRodapeArea))}</td>`;
            tabelaCriada = tabelaCriada + '</tr>';
            tabelaCriada = tabelaCriada + '</table>';
            tabelaCriada=tabelaCriada + '</div>';

            tabelaCriada = tabelaCriada + "<div style='align:center'>";
            tabelaCriada = tabelaCriada + "<button type='button' onclick='geraRelatorio()' class='btn btn-danger' title='Voltar' style='margin-top: 27px;'>";
            tabelaCriada = tabelaCriada + "<i class='fas fa-arrow-left'></i>";
            tabelaCriada = tabelaCriada + "</button>";
            tabelaCriada = tabelaCriada + "</div>";

            document.getElementById('imprimiAqui').innerHTML= tabelaCriada;
            MostrarPagina();
          }

        }

        AdicionaTemporizador();

        var urlRelatorio = "ResumoDeVenda_RepositorioAnalitico.asp" +
                            "?Filial=" + cboPlanta +
                            "&DataIni=" + dataIni +
                            "&DataFim=" + dataFim +
                            "&cboClienteGrupo=" + cboClienteGrupo.trim() +
                            "&Representante=" + representante.trim() +
                            "&Cliente=" + codcli.trim() +
                            "&comOuSemImposto=" + comOuSemImposto +
                            "&codigoCliente=" + codigoCliente +
                            "&tipoProduto=" + tipoProdutoParaORelatorio +
                            "&tipoPapelao=" + tipoPapelao;

        console.log(urlRelatorio);

        xmlhttp.open("GET",urlRelatorio,true);
        xmlhttp.send();
        console.log("geraRelatorioAnalistico");

      }
      //-------------------------------------------------------------------------------------------
      function geraRelatorio(){

        const div = document.getElementById('imprimiAqui');
        div.innerHTML = '';

        codcli = document.getElementById("codcli").value;
        representante = document.getElementById("Representante").value;
        tipoProdutoParaORelatorio = parseFloat(document.getElementById("cbotipoProduto").value);
        tipoProduto = parseFloat(document.getElementById("cbotipoProduto").value);
        tipoPapelao = parseFloat(document.getElementById("cbotipoPapelao").value);
        cboPlanta=document.getElementById("filial").value;
        dataIni=document.getElementById("DataIni").value;
        dataFim=document.getElementById("DataFim").value;
        cboClienteGrupo=document.getElementById("cboClienteGrupo").value;
        comOuSemImposto=document.getElementById("cboImpostos").value;
        cboValores=document.getElementById("cboValores").value;
        document.getElementById("flagTela").value="1";

        document.getElementById("imprimiAqui").removeAttribute("tabelaDeImpressao");
        //document.getElementById("btnVoltar").style.display ="none";

        //Tipo produto
        const enumTipoProdutoAmbos=0;
        const enumTipoProdutoPapel=1;
        const enumTipoProdutoPapelao=2;

        //Valores
        const enumValoresTodos=0
        const enumValoresSoPeso=1
        const enumValoresSoMetro=2

        const enumPeso=1;
        const enumArea=2;
        const enumValor=3;

        if (cboPlanta=="0") {
          alert("Código da planta não foi informado!")
          return
        }

        if (dataIni=="") {
          alert("Data inicial vazia!")
          return                
        }

        if (dataFim==""){
          alert("Data final vazia!")
          return                
        }

        if (comOuSemImposto=="")
          comOuSemImposto="0"
  
        retiraOPeso=0;
        
        if ( parseFloat(cboValores)!=enumValoresTodos ) {

          // quando for somente peso, retira a area            
          switch (parseFloat(cboValores)) {
          case enumValoresSoPeso:

            tipoProduto = enumTipoProdutoPapel;
            break;

          case enumValoresSoMetro:

            if ( tipoProduto == enumTipoProdutoPapel ) {
              alert("Campo tipo de produto e Valores estão divergentes!")
              return                
            } else
              retiraOPeso=1;  //igual a 1 eh para retirar o peso
            break;

          }

        }

        if (tipoProduto!=enumTipoProdutoAmbos) {

          switch (parseFloat(tipoProduto)) {
          case enumTipoProdutoPapel:

            if(parseFloat(cboValores)==enumValoresSoPeso) 
              tipoProduto=enumTipoProdutoPapel;    
            else if(parseFloat(cboValores)==enumValoresSoMetro) {
              alert("Campo tipo de produto e Valores estão divergentes!")
              return
            } else
              cboValores = "1";
            break;

          case enumTipoProdutoPapelao:

            if(parseFloat(cboValores)==enumValoresSoPeso)
              tipoProduto=enumTipoProdutoPapel;  
            else if(parseFloat(cboValores)==enumValoresSoMetro)
              retiraOPeso=1;  
            else
              cboValores = "1";
            break;

          }
        }

        //Gera as datas do relatorio
        if (window.XMLHttpRequest){// code for IE7+, Firefox, Chrome, Opera, Safari
          xmlhttp = new XMLHttpRequest();
        } else {// code for IE6, IE5
          xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
          
        xmlhttp.onreadystatechange=function() {

          var quantidadeDeColuna=4;

          if (this.readyState == 4 && this.status == 200) {

            if (xmlhttp.responseText == "") {
              const element = document.getElementsByTagName("tabelaDeImpressao");
              if (element.length > 0) {
                element.remove();
              }
              alert("Dados não encontrados nos parametros informados!");
              return;                    
            }

            console.log('xmlhttp.responseText',xmlhttp.responseText)

            var resultadoDatas=xmlhttp.responseText.split('_')[0];
            var resultadoDados=xmlhttp.responseText.split('_')[1];
            var arrayDatas=resultadoDatas.split(';');
            var flagTemColunasVazias=false;

            //console.log("DATA: ",resultadoDatas)
            //console.log("DADOS: ",resultadoDados)

            if (retiraOPeso == 1 ) {
              quantidadeDeColuna = 3;
            } else {
              if ( tipoProduto == enumTipoProdutoPapel )
                quantidadeDeColuna = 3;
            }

            tabelaCriada = "<table class='table table-bordered table-striped' id='tabelaDeImpressao'>";
            tabelaCriada = tabelaCriada + '<thead style="position: sticky; inset-block-start: 0; background: white;">';
            tabelaCriada = tabelaCriada + '<tr class="table-active">';
            tabelaCriada = tabelaCriada + '<th colspan="3" class="ltb" style="background-color:#FFF"></th>';

            //Acrescenta os meses ao array de campos
            for (var coluna = 0; coluna < arrayDatas.length; coluna++) {
              tabelaCriada = tabelaCriada + `<th colspan="${quantidadeDeColuna}" class="ltb" style="text-align:center;background-color:#EEE">${arrayDatas[coluna]}</th>`;
            }

            //Array que controla a soma do rodape
            var arrayValoresRodape=Array(arrayDatas.length*(quantidadeDeColuna));

            tabelaCriada = tabelaCriada + '<th colspan="4" class="ltb" style="text-align:center;background-color:#CCC">TOTAL</th>';

            tabelaCriada = tabelaCriada + '</tr>'; 
            tabelaCriada = tabelaCriada + '<tr>';
            tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Representante</th>';
            tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Grupo</th>';
            tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Cliente</th>';   
                    
            //Acrescenta os meses ao array de campos       
            for (var coluna = 0; coluna < arrayDatas.length; coluna++) {

              if (arrayDatas[coluna]!="") {

                if (retiraOPeso!=1)
                  tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Peso</th>';

                if(tipoProduto!=enumTipoProdutoPapel)
                  tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">M2</th>';

                tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Valor</th>';
                tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Prc</th>';

              }

            }

            if(retiraOPeso!=1)
              tabelaCriada=tabelaCriada + '<th class="ltb" style="text-align:center;">Peso</th>';

            if (tipoProduto != enumTipoProdutoPapel) 

              tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">M2</th>';
              tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Valor</th>';
              tabelaCriada = tabelaCriada + '<th class="ltb" style="text-align:center;">Prc</th>'; 
              tabelaCriada = tabelaCriada + '</tr>'; 
              tabelaCriada = tabelaCriada + '</thead>'; 

              var cadaRegistro = resultadoDados.split(";");

              for (var cada = 0; cada < cadaRegistro.length; cada++) {

                tabelaCriada = tabelaCriada + '<tr>';

                var campos=cadaRegistro[cada].split("|");
                var totalPeso=0;
                var totalValor=0;
                var totalArea=0;
                var totalLinhaPeso=0;
                var totalLinhaValor=0;
                var totalLinhaArea=0;
                var calculaPorPesoPorArea=0;
                var proximo=0;     

                //Loop para cada registro

                for (var tdColuna = 0; tdColuna < campos.length; tdColuna++) { 

                  if (tdColuna==3) {

                    //loop de meses do filtro do relatorio 
                    var quantidadeDeMeses=resultadoDatas.split(';');
                    var quebraPorMes=campos[tdColuna].split(":");
                    proximo=0;
                    posicaoDoArrayDeValoresRodape=0;

                    for (var _qtdeMeses = 0; _qtdeMeses < quantidadeDeMeses.length; _qtdeMeses++) {

                      //separação dos dados do mes
                      if(proximo<quebraPorMes.length)
                        var quebraPorBarra=quebraPorMes[proximo].split("/")

                      if (quantidadeDeMeses[_qtdeMeses]==quebraPorBarra[0]) {

                        if (retiraOPeso!=1) {
                          // PESO -----------------------------------------------------------------
                          tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumPeso]}</td>`;
                          totalPeso+=trataCampoSoma(quebraPorBarra[enumPeso]);
                          totalLinhaPeso+=totalPeso;

                        }

                        if (tipoProduto!=enumTipoProdutoPapel) {
                          // AREA -----------------------------------------------------------------
                          tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumArea]}</td>`;
                          totalArea+=trataCampoSoma(quebraPorBarra[enumArea]);
                          totalLinhaArea+=totalArea;

                        }

                        // VALOR ------------------------------------------------------------------
                        tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${quebraPorBarra[enumValor]}</td>`;
                        totalValor += trataCampoSoma(quebraPorBarra[enumValor]);
                        totalLinhaValor+=totalValor;
                    
                        calculaPorPesoPorArea=trataCampoSoma(parseInt(cboValores)!=enumValoresSoMetro?quebraPorBarra[enumPeso]:quebraPorBarra[enumArea]);

                        if (calculaPorPesoPorArea==0)
                          calculaPorPesoPorArea=1;

                        // TOTAL AREA -------------------------------------------------------------------
                        var calculaPRC = trataCampoSoma(quebraPorBarra[enumValor]) / calculaPorPesoPorArea;                           
                        tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(calculaPRC)}</td>`;
                        
                        if (retiraOPeso!=1) {

                          arrayValoresRodape[posicaoDoArrayDeValoresRodape] = somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalPeso);
                          posicaoDoArrayDeValoresRodape++

                        }

                        if(tipoProduto!=enumTipoProdutoPapel){

                          arrayValoresRodape[posicaoDoArrayDeValoresRodape] = somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalArea);
                          posicaoDoArrayDeValoresRodape++

                        }

                        arrayValoresRodape[posicaoDoArrayDeValoresRodape] = somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalValor);
                        posicaoDoArrayDeValoresRodape++
                        arrayValoresRodape[posicaoDoArrayDeValoresRodape] = somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],totalValor/(totalPeso==0?1:totalPeso));
                        posicaoDoArrayDeValoresRodape++

                        proximo++
                        totalPeso=0;
                        totalArea=0;
                        totalValor=0;

                      } else {

                        if (retiraOPeso!=1) {

                          tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">0,000</td>`;
                          posicaoDoArrayDeValoresRodape++
                          arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],0);

                        }
                          
                        if (tipoProduto!=enumTipoProdutoPapel) {

                          tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">0,000</td>`;
                          posicaoDoArrayDeValoresRodape++
                          arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],0);

                        }
                        
                        flagTemColunasVazias=true;
                        posicaoDoArrayDeValoresRodape++
                        arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],0);
                        tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">0,000</td>`;

                        posicaoDoArrayDeValoresRodape++
                        arrayValoresRodape[posicaoDoArrayDeValoresRodape]=somaComOQueTem(arrayValoresRodape[posicaoDoArrayDeValoresRodape],0);
                        tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">0,00</td>`;

                      }

                    }
                        
                    // TOTAL PESO -----------------------------------------------------------------
                    if(retiraOPeso!=1)                                        
                      tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaPeso.toString())}</td>`;

                    // TOTAL AREA -------------------------------------------------------------------
                    if(tipoProduto!=enumTipoProdutoPapel)
                      tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaArea.toString())}</td>`;

                    calculaPorPesoPorArea=parseInt(cboValores)!=enumValoresSoMetro?totalLinhaPeso:totalLinhaArea;

                    if(calculaPorPesoPorArea==0)
                      calculaPorPesoPorArea=1
                      
                    // TOTAL VALOR ----------------------------------------------------------------
                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalLinhaValor.toString())}</td>`;

                    // TOTAL PRECO ----------------------------------------------------------------
                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalLinhaValor/calculaPorPesoPorArea)}</td>`;

                    totalPeso=0;
                    totalValor=0;
                    totalArea=0;
                    posicaoDoArrayDeValoresRodape=0;
                    totalLinhaPeso=0;
                    totalLinhaValor=0;
                    totalLinhaArea=0;
                  } else {

                    // REPRESENTANTE --------------------------------------------------------------
                    if (tdColuna==0)
                      tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:left;">${campos[tdColuna]}</td>`;

                    // GRUPO ----------------------------------------------------------------------
                    if (tdColuna==1)
                      tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:center;">${campos[tdColuna]}</td>`;

                    // CLIENTE --------------------------------------------------------------------
                    if(tdColuna==2)
                        tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:left;"><a href="#" onclick="gerarRelatorioAnalitico('${campos[tdColuna-1]}')"><spam>${campos[tdColuna]}</spam></a></td>`;  
                  }
                }
                tabelaCriada=tabelaCriada + '</tr>';
              }

              //Linha de total
              tabelaCriada=tabelaCriada + '<tr>';
              tabelaCriada=tabelaCriada + '<td colspan="3" class="ltb"></td>';
              contaPassos=0

              totalPRCPeso=0
              totalPRCPreco=0
              totalPRCArea=0
              totalGeralRodapePeso=0
              totalGeralRodapeArea=0
              totalGeralRodapeValor=0
              totalGeralRodapePr=0

              var arrayValido=arrayValoresRodape.length
              
              if(flagTemColunasVazias)
                  arrayValido=arrayValoresRodape.length-1
              
              if (quantidadeDeColuna==4) {

                for (var totalPorData = 0; totalPorData < arrayValido; totalPorData++) {

                  if (contaPassos == quantidadeDeColuna-1) {

                    tabelaCriada  = tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalPRCPreco/totalPRCPeso)}</td>`;
                    totalPRCPeso  = 0
                    totalPRCPreco = 0

                  } else {

                    tabelaCriada = tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(arrayValoresRodape[totalPorData])}</td>`;

                  } 

                  if (contaPassos==0) {

                    contaPassos++;
                    totalPRCPeso+=parseFloat(arrayValoresRodape[totalPorData]);
                    totalGeralRodapePeso+=parseFloat(arrayValoresRodape[totalPorData]);

                  } else if (contaPassos==1) {

                    contaPassos++;
                    totalPRCArea+=parseFloat(arrayValoresRodape[totalPorData]);
                    totalGeralRodapeArea+=parseFloat(arrayValoresRodape[totalPorData]);

                  } else if (contaPassos==2) {

                    contaPassos++;
                    totalPRCPreco+=parseFloat(arrayValoresRodape[totalPorData]);
                    totalGeralRodapeValor+=parseFloat(arrayValoresRodape[totalPorData]);

                  } else if (contaPassos<(quantidadeDeColuna-1))
                    contaPassos++;
                  else
                    contaPassos=0;
                }
              } else {
                // 3 colunas

                for (var _totalPorData = 0; _totalPorData < arrayValoresRodape.length; _totalPorData++) {

                  if (contaPassos == quantidadeDeColuna-1) {

                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalPRCPreco/((parseInt(cboValores)!=enumValoresSoMetro)?totalPRCPeso:totalPRCArea))}</td>`;
                    totalPRCArea=0
                    totalPRCPeso=0
                    totalPRCPreco=0
                    contaPassos++;
                  } else {
                    tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(arrayValoresRodape[_totalPorData])}</td>`;
                  }    

                  if (contaPassos==0) {

                    contaPassos++;

                    if ( parseInt(cboValores) == enumValoresSoPeso ) {

                      totalPRCPeso+=parseFloat(arrayValoresRodape[_totalPorData]);
                      totalGeralRodapePeso+=parseFloat(arrayValoresRodape[_totalPorData]);

                    } else { 

                      totalPRCArea+=parseFloat(arrayValoresRodape[_totalPorData]);
                      totalGeralRodapeArea+=parseFloat(arrayValoresRodape[_totalPorData]);

                    }

                  } else if(contaPassos==1) {

                    contaPassos++;
                    totalPRCPreco+=parseFloat(arrayValoresRodape[_totalPorData]);
                    totalGeralRodapeValor+=parseFloat(arrayValoresRodape[_totalPorData]);

                  } else if(contaPassos<(quantidadeDeColuna-1))
                    contaPassos++;
                  else
                    contaPassos=0;
                }

              }

              if(retiraOPeso!=1)
                tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapePeso.toString())}</td>`;

              if(tipoProduto!=enumTipoProdutoPapel) 
                tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapeArea.toString())}</td>`;

              tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${trataValoresTotalizadores(totalGeralRodapeValor.toString())}</td>`;
              tabelaCriada=tabelaCriada + `<td class="ltb" style="text-align:right;">${formataCampo(totalGeralRodapeValor/(parseInt(cboValores)!=enumValoresSoMetro?totalGeralRodapePeso:totalGeralRodapeArea))}</td>`;
              tabelaCriada=tabelaCriada + '</tr>';
              tabelaCriada=tabelaCriada + '</table>';

              document.getElementById('imprimiAqui').innerHTML = tabelaCriada;
              MostrarPagina();

          }

        }

        AdicionaTemporizador();

        var urlRelatorio="ResumoDeVenda_Repositorio.asp?Filial=" + cboPlanta + 
                                                      "&DataIni=" + dataIni + 
                                                      "&DataFim=" + dataFim + 
                                                      "&cboClienteGrupo=" + cboClienteGrupo.trim() + 
                                                      "&Representante=" + representante.trim() + 
                                                      "&Cliente=" + codcli.trim() +
                                                      "&comOuSemImposto=" + comOuSemImposto + 
                                                      "&tipoProduto=" + tipoProdutoParaORelatorio +
                                                      "&tipoPapelao=" + tipoPapelao 

        console.log( urlRelatorio );
        xmlhttp.open( "GET", urlRelatorio, true );
        xmlhttp.send();
        console.log( "geraRelatorio" );

      }
      // ------------------------------------------------------------------------------------------
      function imprima() {
        print();
      }
      // ------------------------------------------------------------------------------------------
      function doOnLoad() {
        myCalendar = new dhtmlXCalendarObject(["DataIni", "DataFim"]);
        myCalendar.setDate(myCalendar.getDate());
        myCalendar.loadUserLanguage("br");
      }
      // ------------------------------------------------------------------------------------------
      function ConsultaCliente()
      {
        if (document.ResumoDeVendas.filial.value != '0') {

          lnk = '../../buscaCLIENTECorporativo.asp?id1=codcli&id2=lojacli&id3=razao&id6=' + document.ResumoDeVendas.filial.value;
          janela( lnk, 'Consulta Cliente', 'width=650,height=400,scrollbars=1' ); 
          
        } else {

          alert('Selecione a Filial !');

        }
      }
      // --------------------------------------------------------------------------------------------
      function pesquisaCliente( obj ) {
        if (document.ResumoDeVendas.filial.value != '0') {
          if(obj.value.trim() != '' && obj.value.trim() != '') { 
            lnk = '../../ajax/validacliente_ajax.asp?planta=' + document.ResumoDeVendas.filial.value + '&cliente=' + obj.value + '&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Grupo';
            ObjAjax(lnk);
          } else {
            obj.value = ''; 
          }
        } else {
          alert('Selecione a Filial !');
        }
      }
      // --------------------------------------------------------------------------------------------
      function objReturn(obj) {

        if (obj.erro == 'ErroLogof') {

          alert('Favor se conectar novamente ao SWnet.\nSeu logon EXPIROU.');

          if (obj.tipo == 'Grupo') {

            document.getElementById('codcli').value = ''
            document.getElementById('lojacli').value = ''
            document.getElementById('razao').value = ''

          }

          return false;

        }

        if (obj.erro == 'ErroNEncontrado') {

          if (obj.tipo == 'Grupo') {

              document.getElementById('codcli').value = ''
              document.getElementById('lojacli').value = ''
              document.getElementById('razao').value = ''

              alert('CLIENTE NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente Bloqueado.\nCliente não pertence a sua carteira.\nCliente não Cadastrado.');

          }

          return false;

        }

        if (obj.erro == 'ErroCliente') {

          alert('Favor preencher os campos Cliente e Loja.');

          if (obj.tipo == 'Grupo') {

            document.getElementById('codcli').value = ''
            document.getElementById('lojacli').value = ''
            document.getElementById('razao').value = ''

          }

          return false;

        }

        // console.log(obj.tipo);
        // console.log(obj.val1);

        if (obj.tipo == 'Grupo') {

          // console.log(obj.val1);

          ResumoDeVendas.razao.value = obj.val1;

        }

      }      
      // ------------------------------------------------------------------------------------------
      function abrirManual() {
          // Substitua pelo caminho do seu arquivo PDF
          const url = '../../manuais/manual_resumo_venda.pdf ';
          window.open(url, '_blank');
      }
      // ------------------------------------------------------------------------------------------
      function geraExcel(){
        document.getElementById("excel").value = '1';
        ResumoDeVendas.submit();
      }
    </script>

    <style>
      .margimBody { margin-top:10px; }
      #tabelaContainer {
        overflow: auto; 
        max-height: 600px
      }
      #tabelaResumo thead {
        position: sticky; 
        inset-block-start: 0; 
        background: white;
      }
      table {
        font-size: 1.2rem;
      }
    </style>

  </head>

  <body id="ResumoDeVendasBrasil" class="body margimBody" onload="MostrarPagina()">
    <div class="container-fluid" id="divTabela">
      <div class="panel panel-primary">
        <div class="panel-heading"><h3 class="text-center">Resumo de Vendas</h3></div>
        <div class="panel-body">
        <% if request("excel") <> 1 then %>
          <form name="ResumoDeVendas" id="ResumoDeVendas" action="ResumoDeVendasNew.asp" method="post" role="form" >

            <input type="hidden" name="excel" id="excel" value=""/>
            <input type="hidden" name="flagTela" id="flagTela" value=""/>

            <div class="row">
                        
              <div class="col-sm-2">

                <label for="filial">Filial:</label>
                <select name="filial" id="filial" class="form-control">
                  <option value="0">Selecione...</option>
                  <option value="00">Pirapetinga</option>
                  <option value="04">Uberaba</option>
                  <option value="05">Bento Gonçalves</option>
                  <option value="06">Fortaleza</option>
                  <option value="03">Maranguape</option>
                  <option value="07">Saquarema</option>
                  <option value="10">ESHOP</option>
                </select>

              </div>
                            
              <div class="col-sm-6">

                <label for="cliente">Cliente Especifico:</label>
                <div class="input-group-prepend">
                  <input type="number" name="codcli" id="codcli" value="<%=trim(request("codcli"))%>" onblur="pesquisaCliente( this )" class="form-control" style="width:100px" />
                  &nbsp;
                  <button type="button" class="btn btn-primary nao-imprimir" onclick="ConsultaCliente()">
                    <i class="fa fa-search"></i>
                  </button>
                  &nbsp;
                  <input name="razao" id="razao" type="text" class="form-control" readonly value="<%=trim(request("razao"))%>" style="width:100%" />
                  <input type="hidden" name="lojacli" id="lojacli"/>

                </div>

              </div>

            </div>  

            <div class="row">

              <div class="col-sm-2">
                <label for="DataIni">Data Inicial:</label>
                <!-- <input type="text" name="DataIni" id="DataIni" class="form-control" maxlength="10" onkeypress="formatar(this, '##/##/####');" placeholder="Inicial" />&nbsp; -->
                <input type="date" name="DataIni" id="DataIni"  maxlength="10" class="form-control" placeholder="Data Inicial"/>                  
              </div>

              <div class="col-sm-2">
                <label for="DataFim">Data Final:</label>
                <!-- <input type="text" name="DataFim" id="DataFim" class="form-control" maxlength="10" onkeypress="formatar(this, '##/##/####');" placeholder="Final"/> -->
                <input type="date" name="DataFim" id="DataFim" maxlength="10" class="form-control" placeholder="Data Final"/>
              </div>

              <div class="col-sm-2">
                <label for="cboClienteGrupo">Clientes Grupo :</label>
                <select name="cboClienteGrupo" id="cboClienteGrupo" class="form-control">
                  <option value="T"> Todos</option>
                  <option value="S" >Somente Grupo</option>
                  <option value="N" >Somente Clientes</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="Representante">Representante:</label>
                <input type="text" name="Representante" id="Representante" class="form-control" maxlength="6" />
              </div>

            </div>  

            <div class="row">

              <div class="col-sm-2">
                <label for="cbotipoProduto">Tipo de Produto</label>
                <select name="cbotipoProduto" id="cbotipoProduto" onchange="TratacboTipoPapelao(this)" class="form-control">
                  <option value="0">Ambos</option>
                  <option value="1">Papel</option>
                  <option value="2">Papelão</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="cbotipoPapelao">Tipo de Papelão</label>
                <select name="cbotipoPapelao" id="cbotipoPapelao" class="form-control">
                  <option value="0">Ambos</option>
                  <option value="1" >Caixa</option>
                  <option value="2" >Chapa</option>
                  <option value="3" >Corte e Vinco</option>
                  <option value="4" >Maleta</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="cboValores">Valores</label>
                <select name="cboValores" id="cboValores" class="form-control">
                  <option value="0">Todos</option>
                  <option value="1">Só Peso</option>
                  <option value="2">Só Mts2</option>
                </select>     
              </div>

              <div class="col-sm-2">
                <label for="cboImpostos">Impostos</label>
                <select name="cboImpostos" id="cboImpostos" class="form-control">
                  <option value="0">Com IPI</option>
                  <option value="1">Sem IPI</option>
                  <option value="2">Sem Imposto (ICM, PIS, COF)</option>
                </select>     
              </div>

              <div class="col-md-3 nao-imprimir">
                <button type="button" onclick="geraRelatorio()" class="btn btn-warning" title="Gerar" style="margin-top: 27px;">
                  <i style="font-size:24px" class="fa">&#xf0f6;</i>                                
                </button>
                <button type="button" onclick="imprimir()"      class="btn btn-primary" title="Imprimir" style="margin-top: 27px;" >
                  <i style="font-size:24px" class="fa">&#xf02f;</i>
                </button>
                <button type="button" class="btn btn-success" id="btnExcel" title="Exportar para Excel" style="margin-top: 27px;" >
                  <i style="font-size:24px" class="fa">&#xf1c3;</i>
                </button>
                <button type="button" onClick="abrirManual()" class="btn btn-info" title="Manual" style="margin-top: 27px;" >
                  <i style="font-size:24px" class="fa">&#xf29c;</i>
                </button>
              </div>
                        
            </div>

          </form>
        <% end if %>
        </div>
      </div>
      <div id="loader"></div>

      <div class="row">                     
        <div class="col text-right">
          <p>Data/Hora: <%=right("00" & Day(Now), 2) & "/" & right("00" & Month(Now), 2) & "/" & Year(Now) & " " %><span id="relogio"></span></p>
        </div>
      </div>

      <div style="display:none;" class="animate-bottom" id="imprimiAqui" name="imprimiAqui" onload="MostrarPagina()"></div>
    </div>

    <script>

      // ------------------------------------------------------------------------------------------
      document.addEventListener('DOMContentLoaded', function() {

        $(document).ready(function () {

          $("#btnExcel").click(function (e) {

            // Obtém a tabela
            var tabela = document.getElementById('imprimiAqui');

            // Verifica se a tabela existe
            if (!tabela) {
                alert("Tabela não encontrada!");
                return;                
            } else {

              console.log( 'imprimiAqui: ' + tabela );

              // Verifica se existem linhas na tabela
              if (tabela.children.length === 0) {
                  alert("Por favor, primeiro filtre o relatório!");
                  return;                
              }

              var linhas = tabela.getElementsByTagName('tr'); // Obtém todas as linhas da tabela

              for (var i = 0; i < linhas.length; i++) {

                  var celulas = linhas[i].getElementsByTagName('td'); // Obtém as células da linha

                  if (celulas.length > 0) { // Verifica se a linha tem células
                      for (var j = 0; j < celulas.length; j++) {
                        if (!isNaN(celulas[j].innerText.replace(',', '.'))) {
                          celulas[j].innerText = celulas[j].innerText.replace(',', '.');
                        }
                      }
                  }
              }

              const dataHoraAtual = obterDataHoraAtual();
              NomeDaPlanilha = dataHoraAtual.data + "_" + dataHoraAtual.hora;

              // Cria uma nova planilha
              var wb = XLSX.utils.table_to_book(tabela, { sheet: NomeDaPlanilha });

              NomeDoArquivo = 'Relatorio-Resumo-de-Vendas-Brasil_' + dataHoraAtual.data + "_" + dataHoraAtual.hora + '.xlsx';

              // Exporta a planilha para um arquivo Excel
              XLSX.writeFile(wb, NomeDoArquivo);

            }

          });

        });      

      });      
      // ------------------------------------------------------------------------------------------
      function obterDataHoraAtual() {
          const agora = new Date();

          // Formatar a data no formato YYYYMMDD
          const ano = agora.getFullYear();
          const mes = String(agora.getMonth() + 1).padStart(2, '0'); // Meses começam em 0
          const dia = String(agora.getDate()).padStart(2, '0');

          const dataFormatada = `${ano}${mes}${dia}`;

          // Formatar a hora no formato HH:MM:SS
          const horas = String(agora.getHours()).padStart(2, '0');
          const minutos = String(agora.getMinutes()).padStart(2, '0');
          const segundos = String(agora.getSeconds()).padStart(2, '0');

          const horaFormatada = `${horas}${minutos}${segundos}`;

          return {
              data: dataFormatada,
              hora: horaFormatada
          };
      }
    </script>


    <script type="text/javascript">

      function adicionarZero(numero) {
        return numero < 10 ? '0' + numero : numero.toString();
      }

      function startTime() {
        const today = new Date();
        let h = today.getHours();
        let m = today.getMinutes();
        let s = today.getSeconds();
        // Adiciona um zero à esquerda para minutos e segundos menores que 10
        m = checkTime(m);
        s = checkTime(s);
        // Atualiza o conteúdo do div com a hora atual
        const span = document.getElementById('relogio');
        // Gera um novo conteúdo (por exemplo, a hora atual)
        const newContent = adicionarZero(h) + ":" + m + ":" + s;
        // Atualiza o conteúdo da tag <span>
        span.innerHTML = newContent;
        // Chama a função a cada segundo
        setTimeout(startTime, 1000);
      }

      function checkTime(i) {
        return (i < 10) ? "0" + i : i; // Adiciona um zero à esquerda
      }

      // Inicia o relógio
      startTime();

    </script>


  </body>
</html>