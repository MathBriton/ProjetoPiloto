<!--#include file="../../../cnn2.1/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp" -->
<%
Server.ScriptTimeout=6000
Session.LCID = 1046

'response.Write request("excel")

if trim(request("excel") = "on") then
  response.ContentType="application/vnd.ms-excel"
  response.AddHeader "content-disposition", "inline; filename=carregamento.xls"
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html lang="pt-br">
  <head>

    <meta charset="utf-8" />
    <title>SWnet - Mapa de Carregamento</title>

    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

    <script src="../../../funcoes/funcoesJAVASCRIPT.js"></script>

    <script type="text/JavaScript">

      // ------------------------------------------------------------------------------------------
      function doOnLoad() {
        OcultaLoader();
      }  
      // ------------------------------------------------------------------------------------------
      function MostraLoader() {
        document.getElementById("loader").style.display = "block";
      }
      // ------------------------------------------------------------------------------------------
      function OcultaLoader() {
        document.getElementById("loader").style.display = "none";
      }
      // ------------------------------------------------------------------------------------------
      function verifica() {

        var tMsg = '';
        var vOk = 1;

        if (document.getElementById("dataini").value=="") {
          tMsg = "\nDigite a data inicial. ";
          vOk = 0;
        }

        if (document.getElementById("datafin").value=="") {
          tMsg = "\nDigite a data final. ";
          vOk = 0;
        }

        if (vOk == 1) {

          document.form.filtra.value  = "1";  
          document.form.excel.value = 0;

          console.log('excel : ' + document.form.excel.value );
          console.log('filtra: ' + document.form.filtra.value );

          MostraLoader();

          document.form.submit();

        } else {

          alert( tMsg );

        }
      }
      // ------------------------------------------------------------------------------------------
      function exportarPdf() {
        url = "MapaCarregamento01_pdf.asp" + 
              "?DataIn=" + document.form.DataIn.value + 
              "&filial=" + document.form.cboFilial.value;

        console.log(url);

        window.open(url);
      }
      // ------------------------------------------------------------------------------------------
      function limparFiltro(){
        document.form.DataIn.value = '';
        document.form.cboFilial.value = '00'
      }

    </script>

    <style type="text/css">
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

      .tableFixHead          { overflow: auto; height: 78vh; }
      .tableFixHead thead th { 
        position: sticky; 
        top: 0; 
        background-color: #c0c0c0; 
        color: black; 
        font-weight: bold; 
        }

      /* Just common table stuff. Really. */
      table  { width: 100%; }
      /*th, td { padding: 8px 16px; }*/

      #pedido {
        text-transform:uppercase
      }

      @media print {
          @page {
              size: landscape; /* Define a orientação paisagem */
          }
          body {
              margin: 0; /* Remove margens para melhor aproveitamento do espaço */
          }
      }  

      /* Ocultar div quando imprimir */
      @media print {
          .no-print {
              display: none !important;
          }
      }

    </style>

  </head>

  <body onload="doOnLoad()">

    <div class="row">

      <div class="col-1">
        &nbsp;
      </div>

      <div class="col text-center">
        <h3>Mapa de Carregamento</h3>
      </div>

      <div class="col-1" style="align: right">
          <a href="../../../corporativo/manuais/manual_mapa_carregamento_excel.pdf" target="_blank" title="Manual" class="btn btn-info btn-sm">
            <i class="fas fa-question-circle"></i>
          </a>  
      </div>
    </div>

    <div class="container-fluid">
      <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
        <div class="container-fluid" style="background-color: rgb(242, 242, 242);">

          <!-- FILTRO ------------------------------------------------------------------------- -->

          <form id="form" name="form" method="post">

            <input type="hidden" name="filtra" id="filtra" value="0" />
            <input type="hidden" name="excel"  id="excel"  value="0" />
            <input type="hidden" name="pdf"    id="pdf"    value="0" />

            <div class="row">

              <div class="col">
                &nbsp;
              </div>

              <div class="col-2">
                <label for="tipo_consulta" class="form-label">Tipo de consulta</label>
                <select name="tipo_consulta" id="tipo_consulta" class="form-control">
                  <option value="1"<%if trim(request("tipo_consulta")) = "1" then %> selected="selected"<%end if %>>Data de Emissão</option>
                  <option value="2"<%if trim(request("tipo_consulta")) = "2" then %> selected="selected"<%end if %>>Data de Saída</option>
                </select>
              </div>

              <div class="col-2">
                <label for="dataini" class="form-label">Data de</label>
                <input name="dataini" id="dataini" type="date" class="form-control" value="<%=request("dataini")%>">
              </div>

              <div class="col-2">
                <label for="datafin" class="form-label">Data até</label>
                <input name="datafin" id="datafin" type="date" class="form-control" value="<%=request("datafin")%>">
              </div>

              <div class="col-2">
                <label for="cboFilial" class="form-label">Filial</label>
                <select name="cboFilial" id="cboFilial" class="form-control">
                  <option value="00" <% if trim(request("cboFilial")) = "00" then %> selected="selected" <% end if %> >PIRAPETINGA</option>
                  <option value="04" <% if trim(request("cboFilial")) = "04" then %> selected="selected" <% end if %> >UBERABA</option>
                  <option value="05" <% if trim(request("cboFilial")) = "05" then %> selected="selected" <% end if %> >BENTO GONÇALVES</option>
                  <option value="06" <% if trim(request("cboFilial")) = "06" then %> selected="selected" <% end if %> >MARANGUAPE</option>
                  <option value="07" <% if trim(request("cboFilial")) = "07" then %> selected="selected" <% end if %> >CONSOLIDADO</option>
                </select>
              </div>

              <div class="col-1 no-print">
                <label for="btnFiltrar" class="form-label">&nbsp;</label>
                <div class="input-group-prepend">
                  <button name="btnFiltrar" class="btn btn-success" title="Filtrar" onClick="verifica()"><i style="font-size: 22px" class="fas fa-filter"></i></button>
                  &nbsp;
                  <button name="btnClear" class="btn btn-danger" title="Limpar o filtro" onClick="limparFiltro()"><i style="font-size: 22px" class="fas fa-eraser"></i></button>
                  &nbsp;
                  <button name="btnExpPDF" class="btn btn-primary" title="Exportar para PDF" onClick="exportarPdf()"><i style="font-size: 22px" class="fas fa-print"></i></button>
                  &nbsp;
                  <button name="btnExpExcel" class="btn btn-outline-success" title="Exportar para Excel" onClick="exportarExcel()"><i class="far fa-file-excel"></i></i></button>
                </div>
              </div>

              <div class="col">
                &nbsp;
              </div>

            </div>

		      </form>

          <!-- FIM DO FILTRO ------------------------------------------------------------------ -->

          <div id="loader"></div>

          <div class="table-responsive tableFixHead">

            <table class="table table-bordered table-sm table-responsive-sm table-hover table-striped">
              <thead class="border border-light">

                <tr style="font-size:10px;" class="border border-light text-center">
                  <td colspan="26">Data de referência: <%=Data%> &nbsp; até &nbsp<%=datafim %></td>
                </tr>

                <tr >
                  <th>Filial</th>
                  <th>Placa</th>
                  <th>Pedido</th>
                  <th>Cliente</th>
                  <th>Referência</th>
                  <th>Tipo</th>
                  <th>Tipo Papel</th>
                  <th>NF/Serie</th>
                  <th>Emiss.NF.</th>
                  <th>Hora</th>
                  <th>Qtde.</th>
                  <th>Peso</th>
                  <th>Área</th>
                  <th>Valor</th>
                  <th>Vlr.Líquido</th>
                  <th>Frete</th>
                  <th>Frete_Prop.</th>
                  <th>Tipo_Frete</th>
                  <th>Encarreg</th>
                  <th>Dt.Entrega</th>
                  <th>Dt.Saída</th>
                  <th>LocalEnt</th>
                  <th>UF</th>
                  <th>Dif.Dt.Saída.Dt.Entr.</th>
                  <th>Dif.Dt.Saída.Dt.Emissão NF</th>
                  <th>Dif.Dt.Entrega.Dt.Emissão NF</th>
                </tr>

              </thead>

              <tbody>
<% 
if request("filtra") = 1 then

  Dim BuscaNotas

  Dim Param, Num, Bg, Placa, SubTotal, SubTotal2, Total, Total2, TotalPapel, TotalMaleta, TotalChapa, TotalCorteV, Frota
  Dim FrotaTerceirizada, FrotaPropria

  if request("Filial") = "07" then
    laco = 3
  else
    laco = Cint(request("Filial"))
  end if

  Dim vRepre01
  Dim vNOTAIN
  Dim vNOTAFIM
  Dim vRepre02

  Filial = ""

  SQL = "" &_
    "SELECT " &_
      "MIN(Z6_NFINI) AS NOTAIN, " &_
      "MAX(Z6_NFFIM) AS NOTAFIM, " &_
      "MIN(Z6_NFINIB) AS NOTAIN2, " &_
      "MAX(Z6_NFFINB) AS NOTAFIM2 " &_
    "FROM SZ6010 " &_
    "WHERE Z6_DATA BETWEEN '" & replace(request("dataini"), "-", "") & "' AND '" & replace(request("datafiN"), "-", "") & "' " &_
      "AND Z6_FILIAL = '" & request("cboFilial") & "' " &_
      "AND Z6_NFINI <> '000000' " &_
      "AND D_E_L_E_T_ <> '*' "

  'RESPONSE.WRITE SQL & "<BR>"

  SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(PIR_DADOSADV) ) or (PIR_DADOSADV IS NOTHING) then
    response.end
  end if

  if PIR_DADOSADV.EOF then

    response.Write("SZ6 - Não há fechamento diário do faturamento na data solicitada<br><br><a href='javascript:history.back(1);'>Clique aqui para voltar</a>")
    response.End()

  end if

  vNOTAIN = iif( PIR_DADOSADV("NOTAIN") <> "", PIR_DADOSADV("NOTAIN"), "" )
  vNOTAFIM = iif( PIR_DADOSADV("NOTAFIM") <> "", PIR_DADOSADV("NOTAFIM"), "" )

  vNOTAIN2 = iif(isnull(PIR_DADOSADV("NOTAIN2")),".",PIR_DADOSADV("NOTAIN2"))
  vNOTAFIM2 = iif(isnull(PIR_DADOSADV("NOTAFIM2")),".",PIR_DADOSADV("NOTAFIM2")) 

  'Preparando os par�metros de ranger de representantes para enviar para a procedure
  'Se o n�vel for 0, envia todos os representantes, se for 1 ou 2, envia o pr�prio c�digo

  Repre01 = iif( Nivel = 0, "000", CodUser )
  Repre02 = iif( Nivel = 0, "ZZZ", CodUser )

  vRepre01 = ""
  vRepre02 = ""
  vRepre01 = Repre01
  vRepre02 = Repre02

  PIR_DADOSADV.Close()
  SET PIR_DADOSADV = NOTHING

  SQL = "EXEC dbo.SP_MAPACARREGAMENTO '" &_
    replace(vNOTAIN, "'", "''") & "','" &_
    replace(vNOTAFIM, "'", "''") & "','" &_
    replace(vRepre01, "'", "''") & "','" &_
    replace(vRepre02, "'", "''") & "', '" &_
    replace(vNOTAIN2, "'", "''") & "','" &_
    replace(vNOTAFIM2, "'", "''") & "','" &_
    replace(request("dataini"), "-", "") & "','" &_
    replace(request("datafin"), "-", "") & "', " &_
    "'', '" &_
    trim(request("tipo_consulta")) & "'"

  'RESPONSE.WRITE SQL & "<BR>"

  SET PIR_DADOSADV2 = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(PIR_DADOSADV2) ) or (PIR_DADOSADV2 IS NOTHING) then
    response.end
  end if

  tpFrete = ""

  while NOT PIR_DADOSADV2.EOF 

    varPeso = iif( Left(Trim(PIR_DADOSADV2("B1_COD")),1) = "7", PIR_DADOSADV2("D2_PCARGA"), 0 )

    select case PIR_DADOSADV2("F2_TPFRETE") 
    case "C" : tpFrete = "CIF"
    case "F" : tpFrete = "FOB"
    case else: tpFrete = "INDEFINIDO"
    end select   
%>
                <tr class="texto_simples">
                  <td><%=Filial %></td>
                  <td><%=Right((PIR_DADOSADV2("F2_PLACFRT")),3) & " " & Left((PIR_DADOSADV2("F2_PLACFRT")),4)%></td>
                  <td><%=(PIR_DADOSADV2("D2_PEDIDO"))%></td>
                  <td><%=PIR_DADOSADV2("CODIGO") & " " & Left((PIR_DADOSADV2("A1_NREDUZ")),25)%></td>
                  <td><%=Left((PIR_DADOSADV2("B1_DESC")),25)%></td>
                  <td><%=PIR_DADOSADV2("TIPO")%></td>
                  <td><%=PIR_DADOSADV2("TIPOPAPEL")%></td>
                  <td><%=(PIR_DADOSADV2("D2_DOC"))%> - <%=(PIR_DADOSADV2("D2_SERIE"))%></td>
                  <td><%=Right((PIR_DADOSADV2("D2_EMISSAO")),2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),5,2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),3,2)%></td>
                  <td><%=(PIR_DADOSADV2("D2_HORA"))%></td>
                  <td><%=FormatNumber((PIR_DADOSADV2("D2_QUANT")),0)%></td>
                  <td><%=varPeso%></td>
                  <td><%=formatNumber ((PIR_DADOSADV2("D2_AREAAUD")),2)%></td>
                  <td><%=FormatNumber(PIR_DADOSADV2("D2_VALBRUT"),2)%></td>
                  <td><%=FormatNumber(PIR_DADOSADV2("SIMPOSTOS"),2)%></td>
                  <td><%=FormatNumber(PIR_DADOSADV2("F2_VLRFRT"),2)%></td>
                  <td><%=FormatNumber(PIR_DADOSADV2("VALORFRETE"),3)%></td>
                  <td><%=tpFrete%></td>
                  <td><%=Left((PIR_DADOSADV2("D2_ENCA")),12)%></td>
                  <td><%=Right((PIR_DADOSADV2("C6_ENTREG")),2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),5,2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),3,2)%></td>
                  <td><%=Right((PIR_DADOSADV2("F2_DTSAIDA")),2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),5,2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),3,2)%></td>
                  <td><%=Left((PIR_DADOSADV2("A1_MUN")),12)%></td>
                  <td><%=PIR_DADOSADV2("A1_EST")%></td>
<%
    if Trim(PIR_DADOSADV2("F2_DTSAIDA")) <> "" and Trim(PIR_DADOSADV2("C6_ENTREG")) <> "" then  
%>
                  <td><%=datediff("d",cdate(Right((PIR_DADOSADV2("F2_DTSAIDA")),2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),5,2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),3,2)), cdate(Right((PIR_DADOSADV2("C6_ENTREG")),2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),5,2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),3,2)))%></td>
<%
    else 
%>
                  <td>-</td>
<%
    end if

    if Trim(PIR_DADOSADV2("F2_DTSAIDA")) <> "" and Trim(PIR_DADOSADV2("D2_EMISSAO")) <> "" then
%>
                  <td><%=datediff("d", cdate(Right((PIR_DADOSADV2("F2_DTSAIDA")),2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),5,2) & "/" & Mid((PIR_DADOSADV2("F2_DTSAIDA")),3,2)), cdate(Right((PIR_DADOSADV2("D2_EMISSAO")),2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),5,2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),3,2)))%></td>
<%
    else
%>
                  <td>-</td>
<%
    end if

    if Trim(PIR_DADOSADV2("D2_EMISSAO")) <> "" and Trim(PIR_DADOSADV2("C6_ENTREG")) <> "" then
%>
                  <td><%=datediff("d",cdate(Right((PIR_DADOSADV2("D2_EMISSAO")),2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),5,2) & "/" & Mid((PIR_DADOSADV2("D2_EMISSAO")),3,2)), cdate(Right((PIR_DADOSADV2("C6_ENTREG")),2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),5,2) & "/" & Mid((PIR_DADOSADV2("C6_ENTREG")),3,2)))%></td>
<%
    else 
%>
                  <td>-</td>
<%
    end if
%>
                </tr>

<% 
    response.flush()
    PIR_DADOSADV2.MoveNext()
    
  wend

  PIR_DADOSADV2.Close()
  SET PIR_DADOSADV2 = NOTHING

end if
%>
        </div>
      </div>
    </div>         
  </body>
</html>
