<!--#include File = "../../cnn2/cnn2.asp"-->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file = "../../includes/variaveis.asp" -->

<!DOCTYPE html>
<%
Session.LCID = 1046 

if request("acao") = 3 then

  response.ContentType="application/vnd.ms-excel"
  response.AddHeader "content-disposition", "inline; filename=faturamento.xls"

end if
%>
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>SWnet - Mapa de Faturamento</title>  
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
    
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    
    <!-- Select2 CSS -->
    <link href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" rel="stylesheet" />
    
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

      @media print {
        .nao-imprimir {
          display: none;
        }
      }

      @media print {
        @page {
          size: landscape;
        }
      }      

    </style>

    <script charset="utf-8" type="text/javascript">
    
      // ------------------------------------------------------------------------------------------
      function Acao(aAcao) {
        document.getElementById('acao').value = aAcao; 
      }

      function imprima() {
        print();
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
      function verifica() {

        var msg = "Campo de preechimento obrigatório:";
        var valido = 1;

        dtini  = document.form.DataIni.value;
        dtfin  = document.form.DataFim.value; 
        acao   = document.form.acao.value;
       
        if (dtini == "") {
          valido = 0;
          msg = msg + "\nDigite a data inicial. ";
        }

        if (dtfin=="") {
          valido = 0;
          msg = msg + "\nDigite a data final. ";
        }

        if (Date.parse(dtini) > Date.parse(dtfin)) {
          valido = 0;
          msg = msg + "\nA data inicial deve ser menor que a data final! ";
        }

        if (valido == 1) {

          document.form.param.value = 1;

          if (document.form.acao.value != "3") {
            AdicionaTemporizador();
          }

          document.form.submit();

        } else {

          document.form.acao.value = 0;
          document.form.param.value = 0;
          alert(msg);
          return false;

        }

      }  
      // ------------------------------------------------------------------------------------------
      function limparFiltro() {
        document.form.DataIni.value  = "";
        document.form.DataFim.value  = "";
        document.form.Serie.value    = "";
        document.form.param.value    = "0";
        document.form.acao.value     = "0";
      }

    </script>

    <style>
      .margem {
        margin-top:50px;
      }
      .CabecalhoTabela {
        background-color:#CCC;
        text-align: center;
      }
      .Cellspacing {
        border-spacing:5px;
      }
      .Fonte {
        font-size:11px;
      }
    </style>

  </head>

  <body onload="MostrarPagina();" onafterprint="MostrarPagina();">

    <div class="text-center">
      <h3 style="color: gray">Mapa de Faturamento</h3>
    </div>

<%
if request("acao") <> 3 then
%>    
    <!-- FILTRO ------------------------------------------------------------------------------- -->
    <div class="container-fluid">
      <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);" class="nao-imprimir">
        <div class="container-fluid" style="background-color: rgb(242, 242, 242);">

          <form id="form" name="form" method="post" action="">

            <input id="param" name="param" type="hidden" value="1" />
            <input id="acao" name="acao" type="hidden" />

            <div class="row">

              <div class="col-2">
                <label for="consulta">Consultar por</label>
                <select name="consulta" id="consulta" class="form-control">
                  <option value="1" <%if Trim(Request("consulta")) = "1" then%> selected="selected"<%end if %>>Data de Emissão</option>
                  <option value="2" <%if Trim(Request("consulta")) = "2" then%> selected="selected"<%end if %>>Data de Saída</option>
                </select>
              </div>

              <div class="col-2">
                <label for="DataIn">Data Inicial</label>
                <input name="DataIni" id="DataIni" type="date" size="10" value="<%=Trim(Request("DataIni")) %>" class="form-control" placeholder="Inicial" />
              </div>

              <div class="col-2">
                <label for="DataIn">Data Final</label>
                <input name="DataFim" id="DataFim" type="date" size="10" value="<%=Trim(Request("DataFim")) %>" class="form-control" placeholder="Final" />
              </div>

              <div class="col-4">
                <label for="Serie">Filial</label>
                <select name="Serie" id="Serie" multiple="multiple" class="js-example-basic-multiple js-states form-control">
                  <option value="001" <%if InStr(1,Trim(Request("Serie")),"001") > 0 then%> selected="selected"<%end if%>> 00 - MATRIZ</option>
                  <option value="003" <%if InStr(1,Trim(Request("Serie")),"003") > 0 then%> selected="selected"<%end if%>> 04 - UBERABA</option>
                  <option value="002" <%if InStr(1,Trim(Request("Serie")),"002") > 0 then%> selected="selected"<%end if%>> 05 - MATRIZ CD</option>
                  <option value="004" <%if InStr(1,Trim(Request("Serie")),"004") > 0 then%> selected="selected"<%end if%>> 06 - UBERABA CD</option>
                </select>
              </div>

              <div class="col">
                <label>Ação</label>
                <div class="input-group-prepend">
                  <button class="btn btn-primary" title="Filtrar"             onclick="Acao('1'); verifica();"><i class="fas fa-filter"></i></button>
                  &nbsp;
                  <button class="btn btn-info"    title="Imprimir"            onclick="Acao('2'); imprima();"><i class="fas fa-print"></i></button>
                  &nbsp;
                  <button class="btn btn-success" title="Exportar para Excel" onclick="Acao('3'); verifica();" ><i class="far fa-file-excel"></i></button>
                  &nbsp;
                  <button class="btn btn-danger"  title="Limpar o filtro"     onclick="limparFiltro();"><i class="fas fa-eraser"></i></button>
                </div>
              </div>

            </div>

            <br>

          </form>
        </div>
      </div>

<%
end if
%>
      <div id="loader"></div>

      <br>

      <div class="row">
        <div class="col-lg-12">
          <table class="table table-striped table-condensed Fonte">
            <thead>
              <tr class="CabecalhoTabela" style="align:center">
                <th></th>
                <th colspan="28"><center>(A) Diferença entre Data de Saída e Data Prevista de Entrega / (B) Diferença entre Data de Saída e Data de Emissão da NF / (C) Diferença entre Data Prevista de Entrega e Data de Emissão da NF</center></th>
                <th>Data/Hora: <%=right("00" & Day(Now), 2) & "/" & right("00" & Month(Now), 2) & "/" & Year(Now) & " " & Time %></th>
              </tr>
              <tr class="CabecalhoTabela">
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th>Tipo</th>
                <th></th>
                <th>Data</th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
                <th>Valor</th>
                <th>Frete</th>
                <th>Frete</th>
                <th>Tipo</th>
                <th></th>
                <th>Data</th>
                <th>Data</th>
                <th>Local</th>
                <th></th>
                <th></th>
                <th></th>
                <th></th>
              </tr>
              <tr class="CabecalhoTabela">
                <th>Filial</th>
                <th>Placa</th>
                <th>Pedido</th>
                <th>Cliente</th>
                <th>Referência</th>
                <th>Estrutura</th>
                <th>Tipo</th>
                <th>Papel</th>
                <th>NF - Série</th>
                <th>Emissão NF</th>
                <th>Hora</th>
                <th>Quat.</th>
                <th>Peso</th>
                <th>Área</th>
                <th>Valor</th>
                <th>IPI</th>
                <th>PIS</th>
                <th>Cofins</th>
                <th>Líquido.</th>
                <th>Total NF</th>
                <th>Prop./item</th>
                <th>Frete</th>
                <th>Encarregado</th>
                <th>Prev.Entrega</th>
                <th>Saída</th>
                <th>Entrega</th>
                <th>UF</th>
                <th>(A)</th>
                <th>(B)</th>
                <th>(C)</th>
              </tr>
            </thead>
            <tbody>
<%
          if Trim(Request("DataIni")) <> "" and Trim(Request("DataFim")) <> "" then

            TotalNF = 0
            TotalIpi = 0
            TotalPis = 0
            TotalCofins = 0
            TotalSemImpostos = 0                             
            TotalFrete = 0
            TotalPeso = 0
            TotalArea = 0
            x = 0

          Emissao = iif( Trim(Request("consulta")) = "1", "D2_EMISSAO", "F2_DTSAIDA" )

          if len(Trim(Request("Serie"))) > 2 then

            Serie = Split(Trim(Request("Serie")),",")
            SerieSql = ""

            for i=0 to UBound(Serie)
              if Trim(Serie(i)) <> "" then
                SerieSql = SerieSql + "'"&Trim(Serie(i))&"',"
              end if
            next

          else

            SerieSql = "'"&Trim(Request("Serie"))&"',"

          end if

          if SerieSql <> "" then

            SerieSql = Mid(SerieSql,1,len(SerieSql)-1)

          end if

          SQL = "" &_
            "SELECT	" &_
              "F2_VEND1, " &_
              "F2_EMISSAO, " &_
              "CASE B1_DESCNF	" &_
                "WHEN '900' THEN 'PAPEL' " &_
                "WHEN '949' THEN 'PAPEL' " &_
                "WHEN '901' THEN 'MALETA NORMAL' " &_
                "WHEN '902' THEN 'CORTE E VINCO' " &_
                "WHEN '903' THEN 'CHAPA' " &_
              "ELSE " &_
                "B1_DESCNF " &_
              "END AS TIPO, " &_
              "CASE B1_TIPOPAP " &_
                "WHEN '1' THEN 'Papel Maq. 2' " &_
                "WHEN '2' THEN 'Papel Capa Test' " &_
                "WHEN '3' THEN 'Ponta Bobina' " &_
                "WHEN '4' THEN 'Papel Miolo' " &_
                "WHEN '5' THEN 'Papel Test Line' " &_
                "WHEN '6' THEN 'Papel W. T. L.' " &_
                "WHEN '7' THEN 'Papel KRAFT LINER' " &_
              "END AS TIPOPAPEL, " &_
              "CASE WHEN D2_COD BETWEEN '700000000' AND '704999999' THEN " &_
                "D2_TOTAL - D2_VALICM - D2_VALIMP5 - D2_VALIMP6 + D2_VALFRE " &_
              "ELSE " &_
                "D2_TOTAL - D2_VALICM - D2_VALIMP5 - D2_VALIMP6 " &_
              "END AS SIMPOSTOS, " &_
              "CASE WHEN D2_COD BETWEEN '700000000' AND '704999999' THEN (" &_
                "CASE WHEN D2_TOTAL <> D2_VALIPI THEN " &_
                  "D2_TOTAL " &_
                "ELSE 0 END) + D2_VALIPI + D2_VALFRE " &_
              "ELSE (" &_
                "CASE WHEN D2_TOTAL <> D2_VALIPI THEN " &_
                  "D2_TOTAL " &_
                "ELSE 0 END) + D2_VALIPI " &_
              "END AS D2_TOTAL, " &_
              "CASE WHEN D2_COD >= '120000000' AND D2_COD <= '120999999' THEN " &_
                "0 " &_
              "ELSE " &_
                "CASE WHEN D2_COD BETWEEN '700000000' AND '704999999' THEN " &_
                  "D2_QUANT " &_
                "ELSE " &_
                  "SD2010.D2_KGAUD " &_
                "END " &_
              "END AS D2_PCARGA, " &_
              "CASE WHEN ISNULL(F2_PBRUTO,0) <> 0 THEN " &_
                "(( CASE WHEN D2_COD >= '120000000' AND D2_COD <= '120999999' THEN " &_
                      "0 " &_
                    "ELSE " &_
                      "D2_PCARGA " &_
                    "END) / F2_PBRUTO ) * F2_VLRFRT " &_
              "ELSE " &_
                "F2_VLRFRT " &_
              "END AS VALORFRETE, " &_
              "CASE WHEN LTRIM(RTRIM(C6_ENTREG2)) <> '' THEN " &_
                "C6_ENTREG2 " &_
              "ELSE " &_
                "C6_ENTREG " &_
              "END AS C6_ENTREG, " &_
              "CASE WHEN D2_COD BETWEEN '700000000' AND '704999999' THEN " &_
                "'' " &_
              "ELSE " &_
                "ISNULL( " &_
                  "CASE " &_
                  "WHEN SUBSTRING(D2_COD,1,3) IN ('750','760') THEN " &_
                    "ISNULL((" &_
                      "SELECT TOP 1 " &_
                        "B5_ESTRUTU " &_
                      "FROM SG1010 SG1 " &_
                        "INNER JOIN SB5010 SB5 ON B5_COD = G1_COMP AND ISNULL(SG1.D_E_L_E_T_,'') = '' " &_
                      "WHERE SB5.D_E_L_E_T_ <> '*' " &_
                        "AND (SUBSTRING(B5_ESTRUTU,1,3) = 'WTL' OR SUBSTRING(B5_ESTRUTU,1,2) = 'TK') " &_
                        "AND G1_COD = D2_COD),'OUTRAS') " &_
                  "WHEN SUBSTRING(D2_COD,1,3) = '720' THEN " &_
                    "ISNULL((" &_
                      "SELECT TOP 1 " &_
                        "ISNULL(ISNULL(B5_ESTRUTU,B1_ESTRUCH),'OUTRAS') " &_
                      "FROM SB1010 SB1 " &_
                        "LEFT JOIN SB5010 SB5 ON B1_COD = B5_COD " &_
                                            "AND ISNULL(SB5.D_E_L_E_T_,'') = '' " &_
                                            "AND (SUBSTRING(ISNULL(B5_ESTRUTU,''),1,3) = 'WTL' OR SUBSTRING(ISNULL(B5_ESTRUTU,''),1,2) = 'TK') " &_
                      "WHERE SB1010.D_E_L_E_T_ <> '*' " &_
                        "AND (SUBSTRING(ISNULL(B1_ESTRUCH,''),1,3) = 'WTL' OR SUBSTRING(ISNULL(B1_ESTRUCH,''),1,2) = 'TK') " &_
                        "AND B1_COD = D2_COD),'OUTRAS') " &_
                  "ELSE " &_
                    "ISNULL((" &_
                      "SELECT TOP 1 " &_
                        "B5_ESTRUTU " &_
                      "FROM SB1010 SB1 " &_
                        "INNER JOIN SB5010 SB5 ON B1_COD = B5_COD AND ISNULL(SB5.D_E_L_E_T_,'') = '' " &_
                      "WHERE SB1.D_E_L_E_T_ <> '*' " &_
                        "AND (SUBSTRING(B5_ESTRUTU,1,3) = 'WTL' OR SUBSTRING(B5_ESTRUTU,1,2) = 'TK') " &_
                        "AND B1_COD = D2_COD),'OUTRAS') " &_
                  "END, " &_
                "'OUTRAS') " &_
              "END AS ESTRUTURA, " &_
              "( SELECT " &_
                  "ISNULL(A4_NOME,'TRANSP. NÃO ENCONTRADO') " &_
                "FROM SA4010 " &_
                "WHERE A4_COD = F2_TRAN " &_
                  "AND SA4010.D_E_L_E_T_ <> '*') AS A4_NOME, " &_
              "F2_EST, " &_
              "SF2010.F2_PLACFRT, " &_
              "F2_PLACFRT, " &_
              "D2_VALBRUT, " &_
              "F2_VLRFRT, " &_
              "F2_VALBRUT, " &_
              "D2_QUANT, " &_
              "D2_CF, " &_
              "D2_PEDIDO, " &_
              "D2_DOC, " &_
              "D2_SERIE, " &_
              "D2_EMISSAO, " &_
              "D2_HORA, " &_
              "D2_ENCA, " &_
              "A1_NREDUZ, " &_
              "F2_TRAN AS A4_COD, " &_
              "A1_MUN, " &_
              "A1_EST, " &_
              "A1_TIPO, " &_
              "F2_TPFRETE, " &_
              "B1_COD, " &_
              "B1_DESC, " &_
              "B1_DESCNF, " &_
              "C6_ITEM, " &_
              "C6_LARG, " &_
              "C6_COMP, " &_
              "B1_PESO, " &_
              "A1_COD + '-' + A1_LOJA AS CODIGO, " &_
              "C5_OBSVEND, " &_
              "D2_AREAAUD, " &_
              "F2_DTSAIDA, " &_
              "D2_FILIAL, " &_
              "D2_VALIMP5, " &_
              "D2_VALIMP6, " &_
              "D2_VALIPI " &_
            "FROM SF2010 " &_
              "INNER JOIN SA1010 ON F2_CLIENTE = A1_COD AND F2_LOJA = A1_LOJA " &_
              "INNER JOIN (" &_
                "SELECT " &_
                  "D2_IPI, " &_
                  "D2_DOC, " &_
                  "D2_EMISSAO, " &_
                  "D2_COD, " &_
                  "D2_QUANT, " &_
                  "D2_VALIPI, " &_
                  "D2_VALIMP6, " &_
                  "D2_VALIMP5, " &_
                  "D2_VALICM, " &_
                  "D2_FILIAL, " &_
                  "D2_HORA, " &_
                  "D2_ENCA, " &_
                  "D2_AREAAUD, " &_
                  "D2_PEDIDO, " &_
                  "D2_VALBRUT, " &_
                  "CASE WHEN D2_TOTAL <> D2_VALIPI THEN " &_
                    "D2_TOTAL " &_
                  "ELSE " &_
                    "0 " &_
                  "END AS D2_TOTAL, " &_
                  "D2_VALFRE, " &_
                  "D2_AREAVEN, " &_
                  "D2_PCARGA, " &_
                  "D2_KGVEN, " &_
                  "D2_CLIENTE, " &_
                  "D2_LOJA, " &_
                  "D2_CF, " &_
                  "D2_SERIE, " &_
                  "D2_KGAUD, " &_
                  "D2_TIPO, " &_
                  "D_E_L_E_T_ " &_
                "FROM SD2010) AS SD2010 ON SF2010.F2_DOC = SD2010.D2_DOC " &_
                                      "AND SF2010.F2_SERIE = SD2010.D2_SERIE " &_
                                      "AND SF2010.F2_FILIAL = SD2010.D2_FILIAL " &_
              "INNER JOIN SB1010 ON D2_COD = B1_COD " &_
              "INNER JOIN SC5010 ON D2_PEDIDO = C5_NUM AND C5_CLIENTE = D2_CLIENTE AND C5_LOJACLI = D2_LOJA " &_
              "INNER JOIN SC6010 ON C5_NUM = C6_NUM AND C5_CLIENTE = C6_CLI AND C5_LOJACLI =  C6_LOJA " &_
            "WHERE C6_ITEM = '01' " &_
              "AND D2_EMISSAO >= '20091225' " &_
              "AND " & Emissao & " BETWEEN '" & replace(trim(request("DataIni")),"-","") & "' AND '" & replace(trim(request("DataFim")),"-","") & "' " &_
              "AND (D2_CLIENTE <> '9099') " &_
              "AND ((D2_COD >= '700000000' AND D2_COD < '705000000' ) OR D2_COD > '705999999') " &_
              "AND D2_CF IN ('5101','5102', '5103','5105', '5122','5124','5922','6101','6102', '6103','6105','6107', '6109', '6118', '6122','6124','7101','6922') " &_
              "AND D2_TIPO IN ('N','C','P') " &_
              "AND B1_DESCNF IN ('900','901','902','903','906', '949', '950') " &_
              iif( Trim(Request("Serie")) <> "", "AND D2_SERIE IN (" & SerieSql & ") ", "" ) &_
              "AND SF2010.D_E_L_E_T_ <> '*' " &_
              "AND SD2010.D_E_L_E_T_ <> '*' " &_
              "AND SA1010.D_E_L_E_T_ <> '*' " &_
              "AND SC5010.D_E_L_E_T_ <> '*' " &_
              "AND SB1010.D_E_L_E_T_ <> '*' " &_
              "AND SC6010.D_E_L_E_T_ <> '*' " &_	
            "ORDER BY " &_
              "D2_FILIAL, " &_
              "F2_EMISSAO ASC, " &_
              "SF2010.F2_PLACFRT ASC, " &_
              "D2_DOC ASC "

            'response.write SQL            

            SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

            if not isObject( PIR_DADOSADV ) Then
              response.end
            end if

            while NOT PIR_DADOSADV.Eof

              select case PIR_DADOSADV("D2_FILIAL")
              case "00" Filial = "PIRAPETINGA"
              case "04" Filial = "UBERABA"
              case "05" Filial = "PIRAPETINGA CD"
              case "06" Filial = "UBERABA CD"
              end select

              Peso = iif( Left(Trim(PIR_DADOSADV("B1_COD")),1) = "7", PIR_DADOSADV("D2_PCARGA"), 0 )

              select case PIR_DADOSADV("F2_TPFRETE")
              case "C"  TipoFrete = "CIF"  
              case "F"  TipoFrete = "FOB"
              case else TipoFrete = "INDEFINIDO"
              end select
%>
              <tr>   
                <td class="text-nowrap"><%=Filial %></td>
                <td class="text-nowrap text-center"><%=Right(PIR_DADOSADV("F2_PLACFRT"),3) & " " & Left(PIR_DADOSADV("F2_PLACFRT"),4)%></td>
                <td class="text-nowrap text-center"><%=PIR_DADOSADV("D2_PEDIDO")%></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("CODIGO") & " " & Left(PIR_DADOSADV("A1_NREDUZ"),25)%></td>
                <td class="text-nowrap"><%=Left(PIR_DADOSADV("B1_DESC"),25)%></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("ESTRUTURA") %></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("TIPO")%></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("TIPOPAPEL")%></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("D2_DOC")%> - <%=PIR_DADOSADV("D2_SERIE")%></td>
                <td class="text-nowrap text-center"><%=AcertaData(PIR_DADOSADV("D2_EMISSAO"))%></td>
                <td class="text-nowrap text-center"><%=PIR_DADOSADV("D2_HORA")%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("D2_QUANT"),0)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(Peso,2)%></td>
                <td class="text-nowrap text-right"><%=formatNumber(PIR_DADOSADV("D2_AREAAUD"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("D2_TOTAL"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("D2_VALIPI"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("D2_VALIMP6"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("D2_VALIMP5"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("SIMPOSTOS"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("F2_VLRFRT"),2)%></td>
                <td class="text-nowrap text-right"><%=FormatNumber(PIR_DADOSADV("VALORFRETE"),3)%></td>
                <td class="text-nowrap text-center"><%=TipoFrete%></td>
                <td class="text-nowrap"><%=Left(PIR_DADOSADV("D2_ENCA"),12)%></td>
                <td class="text-nowrap text-center"><%=AcertaData(PIR_DADOSADV("C6_ENTREG"))%></td>
                <td class="text-nowrap text-center"><%=AcertaData(PIR_DADOSADV("F2_DTSAIDA"))%></td>
                <td class="text-nowrap"><%=Left(PIR_DADOSADV("A1_MUN"),12)%></td>
                <td class="text-nowrap"><%=PIR_DADOSADV("A1_EST")%></td> 

              <% if Trim(PIR_DADOSADV("F2_DTSAIDA")) <> "" and Trim(PIR_DADOSADV("C6_ENTREG")) <> "" then %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap"><%=datediff("d",AcertaData(PIR_DADOSADV("F2_DTSAIDA")),AcertaData(PIR_DADOSADV("C6_ENTREG")))%></td>
              <% else %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap">-</td>
              <% end if %>

              <% if Trim(PIR_DADOSADV("F2_DTSAIDA")) <> "" and Trim(PIR_DADOSADV("D2_EMISSAO")) <> "" then  %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap"><%=datediff("d",AcertaData(PIR_DADOSADV("F2_DTSAIDA")),AcertaData(PIR_DADOSADV("D2_EMISSAO")))%></td>
              <% else %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap">-</td>
              <% end if %>

              <% if Trim(PIR_DADOSADV("D2_EMISSAO")) <> "" and Trim(PIR_DADOSADV("C6_ENTREG")) <> "" then  %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap"><%=datediff("d",AcertaData(PIR_DADOSADV("D2_EMISSAO")),AcertaData(PIR_DADOSADV("C6_ENTREG")))%></td>
              <% else %>
                <td style="padding-left:2px; padding-right:2px;text-align:right;white-space:nowrap">-</td>
              <% end if %>
              </tr> 
<%
              TotalNF = TotalNF + PIR_DADOSADV("D2_TOTAL")
              TotalIpi = TotalIpi + PIR_DADOSADV("D2_VALIPI")
              TotalPis = TotalPis + PIR_DADOSADV("D2_VALIMP6")
              TotalCofins = TotalCofins + PIR_DADOSADV("D2_VALIMP5")
              TotalSemImpostos = TotalSemImpostos + PIR_DADOSADV("SIMPOSTOS")                               
              TotalFrete = TotalFrete + PIR_DADOSADV("VALORFRETE")
              TotalPeso = TotalPeso + Peso
              TotalArea = TotalArea + PIR_DADOSADV("D2_AREAAUD")
              x = x + 1

              Response.Flush()
              PIR_DADOSADV.MoveNext()

            wend

            PIR_DADOSADV.Close()
            SET PIR_DADOSADV = NOTHING

          end if
%>
            </tbody>
            <tfoot>
              <tr>
                <th Colspan="12">Total para <%=x %> registro(s)</th>
                <th class="text-right" title="Total Peso"><%=formatnumber(TotalPeso,0) %></th>
                <th class="text-right" title="Total Área"><%=formatnumber(TotalArea,0) %></th>
                <th class="text-right" title="Total Valor"><%=formatnumber(TotalNF,0) %></th>
                <th class="text-right" title="Total IPI"><%=formatnumber(TotalIpi,0) %></th>
                <th class="text-right" title="Total Pis"><%=formatnumber(TotalPis,0) %></th>
                <th class="text-right" title="Total Cofins"><%=formatnumber(TotalCofins,0) %></th>
                <th class="text-right" title="Total Sem Impostos"><%=formatnumber(TotalSemImpostos,0) %></th>
                <th>&nbsp;</th>
                <th class="text-right" title="Total Frete"><%=formatnumber(TotalFrete,0) %></th>
                <th colspan="9">&nbsp;</th>
              </tr>
              <tr>
                <th colspan="30">Fim do relatório</th>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>

    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.3.1.min.js"></script>
    <!-- Bootstrap JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>
    <!-- Select2 JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"></script>

    <script>
        $(document).ready(function() {
            $('#Serie').select2({
                placeholder: "Selecione uma ou mais opções",
                allowClear: true
            });
        });
    </script>

<%
' -------------------------------------------------------------------------------------------------
function DataHoraAtual(pDate, pHour)

  Dim dd, mm, yy, hh, nn, ss
  Dim datevalue, timevalue, dtsnow, dtsvalue
  on error resume next

  if len(pDate) > 10 then
    dtsnow = pDate 
  else
    dtsnow = pDate & " " & pHour
  end if

  dd = Right("00" & Day(dtsnow), 2)
  mm = Right("00" & Month(dtsnow), 2)
  yy = Year(dtsnow)

  hh = Right("00" & Hour(dtsnow), 2)
  nn = Right("00" & Minute(dtsnow), 2)
  ss = Right("00" & Second(dtsnow), 2)

  datevalue = dd & "/" & mm & "/" & yy
  timevalue = hh & ":" & nn & ":" & ss
  dtsvalue = datevalue & " " & timevalue

  if err.number <> 0 then
    datetime = "01/01/1900 00:00:00.000"
		'response.write err.number & "<br>"
		'response.write err.description & "<br>"
  else
    datetime = dtsvalue
	end if

  on error goto 0

end function
%>
  </body>
</html>