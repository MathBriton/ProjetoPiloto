<!DOCTYPE html>
    <!--#Include File = "../cnn2/cnn2.asp"-->
    <!--#Include File = "../funcoes/funcoes.asp"-->

<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Ficha Técnica de Produto</title>
    <!-- BOOTSTRAP -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
    <link rel="stylesheet" type="text/css" href="../../../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
    <!-- BOOTSTRAP -->
    <script charset="utf-8" src="../funcoes/funcoesJAVASCRIPT.js"></script>

    <script type="text/javascript">
      $(document).ready(function() {
        window.print();
        return false;
      });
    </script>

    <style type="text/css">
      body {
        font-size:90%;
        margin:0px;       
      }
      img {
        width:147px;
        height:46px;
      }
      h4 {
        display:inline-block;
      }
      table {
        margin-top: 0px;
        margin-bottom: 0px;
      }
      .BordaInferior {
        border-bottom: 1px solid #ddd;
      }
      .BordaDireita {
        border-right: 1px solid #ddd;
      }
      .BordaEsquerda {
        border-left: 1px solid #ddd;
      }
      .BordaSuperior {
        border-top: 1px solid #ddd;
      }
      .MargemBotton {
        margin-bottom:10px;
      }
      .border {
        border:1px solid black;
      }
      .CabecalhoDestaque {
        background-color:#ccc;
      }
      .MargemEsquerda {
        margin-left:15px;
      }
      .MargemDireita {
        margin-right:30px;
      }
    </style>

    <style type="text/css" media="print">
      @media print {
        @page {
          margin: 0.5cm;
        }
        body {
          font-size:98.5%;
          margin:0px;       
        }
        img {
          margin-right:42px;
        }
        #printable {
          visibility:collapse;
        }
      }
    </style>
    <% session.LCID = 1046 %>
  </head>
<%
SQL = "" &_
  "SELECT " &_
    "B5_COD, " &_
    "B5_COMPR, " &_
    "B5_LARG, " &_
    "B5_ALTU, " &_
    "B5_REF, " &_
    "B5_VEND, " &_
    "B5_NOMVEN, " &_
    "B5_CLI, " &_
    "B5_LOJA, " &_
    "B5_NOMECLI, " &_
    "B5_TOLMA, " &_
    "B5_TOLME, " &_
    "B5_ESTRUTU, " &_
    "B5_ESTILO, " &_
    "B5_SAP, " &_
    "B5_LARGCH, " &_
    "B5_COMPCH, " &_
    "B5_BLANKL, " &_
    "B5_BLANKC, " &_
    "B5_AREA, " &_
    "B5_PESOB, " &_
    "B5_PESOL, " &_
    "B5_ARRL, " &_
    "B5_ARRC, " &_
    "B5_PCSCON, " &_
    "B5_TRATAM, " &_
    "B5_NUMFACA, " &_
    "B5_NUMCLI, " &_
    "B5_FECHAM, " &_
    "B5_COR1, " &_
    "B5_COR2, " &_
    "B5_COR3, " &_
    "B5_COR4, " &_
    "B5_AMARRAD, " &_
    "B5_QTDPCT, " &_
    "B5_FITLAR, " &_
    "B5_FITCOM, " &_
    "B5_PALETE, " &_
    "B5_FACOP, " &_
    "B5_CANTON, " &_
    "B5_FILME, " &_
    "B5_TAMPA, " &_
    "B5_ESPELHO, " &_
    "B5_TPPALE, " &_
    "B5_DESCPAL, " &_
    "B5_FITASL, " &_
    "B5_FICASC, " &_
    "B5_PLLAST, " &_
    "B5_ARRPA, " &_
    "B5_PCSPAL, " &_
    "B5_VINCO1, " &_
    "B5_VINCO2, " &_
    "B5_VINCO3, " &_
    "B5_VINCO4, " &_
    "B5_VINCO5, " &_
    "B5_VINCO6, " &_
    "B5_VINCO7, " &_
    "B5_VINCI1, " &_
    "B5_VINCI2, " &_
    "B5_VINCI3, " &_
    "B5_VINCI4, " &_
    "B5_VINCI5, " &_
    "B5_MAQ1, " &_
    "B5_MAQ2, " &_
    "B5_MAQ3, " &_
    "B5_MAQ4, " &_
    "B5_MAQ5, " &_
    "B5_GRAM, " &_
    "B5_MULLEN, " &_
    "B5_ESMAG, " &_
    "B5_COLUNA, " &_
    "B5_ESPES2, " &_
    "B5_COBINT, " &_
    "B5_COBEXT, " &_
    "B5_COMPRE, " &_
    "B5_VINCACE, " &_
    "B5_CORTACE, " &_
    "B5_OBS1, " &_
    "B5_OBS2, " &_
    "B5_OBS3, " &_
    "B5_DTRECSA, " &_
    "B5_USUARIO, " &_
    "B5_DTCADFT, " &_
    "B5_DTAPRFT, " &_
    "B5_RESPONS, " &_
    "B5_QTDAMO, " &_
    "B5_DTDESP, " &_
    "B5_OBSDDP, " &_
    "B5_AREALI, " &_
    "B5_OBS4, " &_
    "B5_OBS5, " &_
    "B1_STATUS, " &_
    "B1_DESCNF, " &_
    "B1_CODPCP, " &_
    "B1_UPRC," &_
    "B1_INPUBE, " &_
    "Z1_COLUGER, " &_
    "Z1_MULLGER, " &_
    "Z1_ESMAGER, " &_
    "Z1_ESPEGER, " &_
    "Z1_GESTGER, " &_
    "Z1_COBBIGE, " &_
    "Z1_COBBEGE, " &_
    "Z2_DESC, " &_
    "B5_VISCOR1, " &_
    "B5_VISCOR2, " &_
    "B5_VISCOR3 " &_
  "FROM SB5010 SB5010 WITH(NOLOCK) " &_
    "INNER JOIN SZ1010 SZ1010 WITH(NOLOCK) ON B5_ESTRUTU = Z1_ESTRUT AND SZ1010.D_E_L_E_T_ = '' " &_
    "INNER JOIN SZ2010 SZ2010 WITH(NOLOCK) ON B5_ESTILO = Z2_CODIGO AND SZ2010.D_E_L_E_T_ = '' " &_
    "INNER JOIN SB1010 SB1010 WITH(NOLOCK) ON B5_COD = B1_COD AND SB1010.D_E_L_E_T_ = '' " &_
  "WHERE SB5010.D_E_L_E_T_ = '' " &_
    "AND B5_COD = '" & trim(Request("ft")) & "' "

SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

%>
  <body>
    <form action="FichaTecnica.asp" method="post" name="FichaTecnica" id="FichaTecnica">
      <input type="hidden" name="imprimir" id="imprimir" />
    </form>
    <div class="container-fluid body-content margem">
      <div class="row">
        <div class="col-md-9">
          <h4 class="MargemEsquerda">
            <strong>Ficha Técnica de Produto</strong>
          </h4>
          <img src="../../../CORPORATIVO/images/sw_logo_40h.png" class="pull-right"/>
        </div>
      </div>
<%      
    if not PIR_DADOSADV.Eof then 
    
      DescricaoFilial = iif( PIR_DADOSADV("B1_INPUBE") = "U", "UBERABA", "PIRAPETINGA" )
%>
      <div class="row">
        <div class="col-md-10">
          <table>
            <tr>
              <td style="vertical-align:top">
                <div class="col-md-12">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="4" class="text-center" style="border-top:2px solid #ddd">DADOS DO PRODUTO</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td class="BordaEsquerda" style="border-top:1px solid black"><strong>Cliente:</strong></td>
                        <td colspan="3" class="BordaDireita text-nowrap"><%=PIR_DADOSADV("B5_NOMECLI")%>&nbsp;&nbsp;<%=PIR_DADOSADV("B5_CLI") %>&nbsp;<%=PIR_DADOSADV("B5_LOJA") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Representante:</strong></td>
                        <td><%=PIR_DADOSADV("B5_NOMVEN")%>&nbsp;<%=PIR_DADOSADV("B5_VEND") %></td>
                        <td><strong>Preço:</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B1_UPRC")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Referência:</strong></td>
                        <td colspan="3" class="BordaDireita" ><%=PIR_DADOSADV("B5_REF")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Tol + :</strong></td>
                        <td><%=PIR_DADOSADV("B5_TOLMA")%></td>
                        <td><strong>Tol - :</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_TOLME")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Medida Int.:</strong></td>
                        <td><%=PIR_DADOSADV("B5_COMPR")%></td>
                        <td><%=PIR_DADOSADV("B5_LARG")%></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_ALTU")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Estilo:</strong></td>
                        <td><%=PIR_DADOSADV("B5_ESTILO")%>&nbsp;<%=PIR_DADOSADV("Z2_DESC") %></td>
                        <td><strong>NF:</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B1_DESCNF")%></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Peças/Conj.:</strong></td>
                        <td><%=PIR_DADOSADV("B5_PCSCON")%></td>
                        <td><strong>Fábrica.:</strong></td>
                        <td class="BordaDireita"><%=DescricaoFilial%></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div class="col-md-12">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="7" class="text-center" style="border-top:2px solid #ddd">ACABAMENTO</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td class="BordaEsquerda"><strong>Pacote:</strong></td>
                        <td><%=PIR_DADOSADV("B5_AMARRAD") %>&nbsp;-&nbsp;<%=PIR_DADOSADV("B5_QTDPCT") %></td>
                        <td><strong>Paletizado:</strong></td>
                        <td><%=PIR_DADOSADV("B5_PALETE")%></td>
                        <td ><strong>Fitas(LxC):</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_FITASL")%>&nbsp;x&nbsp;<%=PIR_DADOSADV("B5_FICASC")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda text-nowrap"><strong>Palete: Tipo</strong></td>
                        <td colspan="3" class="text-nowrap"><%=PIR_DADOSADV("B5_TPPALE")%>&nbsp;<%=PIR_DADOSADV("B5_DESCPAL")%></td>
                        <td><strong>Lastro</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_PLLAST")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda text-nowrap"><strong>Palete Real - Comp.:</strong></td>
                        <td class="text-nowrap"><%=PIR_DADOSADV("B5_VISCOR1")%></td>
                        <td><strong>Larg.:</strong></td>
                        <td class="text-nowrap"><%=PIR_DADOSADV("B5_VISCOR2")%></td>
                        <td><strong>Alt.</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_VISCOR3")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda text-nowrap"><strong>Palete: Altura</strong></td>
                        <td><%=PIR_DADOSADV("B5_ARRPA")%></td>
                        <td><strong>Quantidade</strong></td>
                        <td colspan="3"  class="BordaDireita"><%=PIR_DADOSADV("B5_PCSPAL")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Fitilhos(LxC):</strong></td>
                        <td  ><%=PIR_DADOSADV("B5_FITLAR")%>&nbsp;x&nbsp;<%=PIR_DADOSADV("B5_FITCOM")%></td>
                        <td><strong>Cantoneira:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=PIR_DADOSADV("B5_CANTON")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda text-nowrap"><strong>Face Oposta:</strong></td>
                        <td><%=PIR_DADOSADV("B5_FACOP")%></td>
                        <td><strong>Tampa:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=PIR_DADOSADV("B5_TAMPA")%></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Espelho:</strong></td>
                        <td><%=PIR_DADOSADV("B5_ESPELHO")%></td>
                        <td><strong>Filme:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=PIR_DADOSADV("B5_FILME")%></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div class="col-md-12">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="7" class="text-center" style="border-top:2px solid #ddd">IMPRESSORA</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td colspan="5" class="BordaDireita BordaEsquerda text-center"><strong>Máquinas</strong></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">1</td>
                        <td>2</td>
                        <td>3</td>
                        <td>4</td>
                        <td class="BordaDireita">5</td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><%=PIR_DADOSADV("B5_MAQ1")%></td>
                        <td><%=PIR_DADOSADV("B5_MAQ2")%></td>
                        <td><%=PIR_DADOSADV("B5_MAQ3")%></td>
                        <td><%=PIR_DADOSADV("B5_MAQ4")%></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_MAQ5")%></td>
                      </tr>
                      <tr>
                        <td colspan="5" class="BordaDireita BordaEsquerda text-center"><strong>Vinco</strong></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">1</td>
                        <td>2</td>
                        <td>3</td>
                        <td>4</td>
                        <td class="BordaDireita">5</td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><%=PIR_DADOSADV("B5_VINCI1")%></td>
                        <td><%=PIR_DADOSADV("B5_VINCI2")%></td>
                        <td><%=PIR_DADOSADV("B5_VINCI3")%></td>
                        <td><%=PIR_DADOSADV("B5_VINCI4")%></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_VINCI5")%></td>
                      </tr>
                      <tr>
                        <td colspan="5" class="BordaDireita BordaEsquerda text-center"><strong>Cor</strong></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">1</td>
                        <td>2</td>
                        <td>3</td>
                        <td colspan="2" class="BordaDireita">4</td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><%=PIR_DADOSADV("B5_COR1")%></td>
                        <td><%=PIR_DADOSADV("B5_COR2")%></td>
                        <td><%=PIR_DADOSADV("B5_COR3")%></td>
                        <td colspan="2" class="BordaDireita"><%=PIR_DADOSADV("B5_COR4")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Fechamento:</strong></td>
                        <td colspan="5" class="BordaDireita" ><%=PIR_DADOSADV("B5_FECHAM")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Cliche:</strong></td>
                        <td><%=PIR_DADOSADV("B5_NUMCLI")%></td>
                        <td><strong>Arquivo:</strong></td>
                        <td colspan="2" class="BordaDireita"><%=Mid(Trim(PIR_DADOSADV("B1_CODPCP")),1,len(Trim(PIR_DADOSADV("B1_CODPCP")))-6) %>-<%=Right(Trim(PIR_DADOSADV("B1_CODPCP")),6) %></td>
                      </tr>

                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Matriz:</strong></td>
                        <td><%=PIR_DADOSADV("B5_NUMFACA") %></td>
                        <td><strong>Cliche/Matriz:</strong></td>
                        <td colspan="2" class="BordaDireita"></td>
                      </tr>
                    </tbody>
                  </table>
                </div>  
              </td>
              <td style="vertical-align:top">
                <div class="col-md-10 text-center MargemBotton">
                  <h2><strong><%=Left(PIR_DADOSADV("B5_COD"),3) %>-<%=Right(Trim(PIR_DADOSADV("B5_COD")),6) %></strong></h2>
                  <label class="control-label">Impressao:</label>
                  <%=Date() %>&nbsp;<%=Time() %>
                </div>
                <div class="col-md-10">
                  <table class="table table-condensed table-striped">
                    <tbody>
                      <tr>
                        <td class="BordaEsquerda"><strong>SAP:</strong></td>
                        <td><%=PIR_DADOSADV("B5_SAP") %></td>
                        <td><strong>Status:</strong></td>
                        <td><%=PIR_DADOSADV("B1_STATUS") %></td>
                        <td><strong>Amostra:</strong></td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_QTDAMO") %></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Recebimento:</strong></td>
                        <td><%=AcertaData(PIR_DADOSADV("B5_DTRECSA")) %></td>
                        <td><strong>Despacho:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=AcertaData(PIR_DADOSADV("B5_DTDESP")) %></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div class="col-md-10">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="7" class="text-center" style="border-top:2px solid #ddd">ONDULADEIRA</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td class="BordaEsquerda"><strong>Estrutura:</strong></td>
                        <td><%=PIR_DADOSADV("B5_ESTRUTU")%></td>
                        <td colspan="2"><strong>Tratamento:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=PIR_DADOSADV("B5_TRATAM") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Chapa(LxC):</strong></td>
                        <td><%=PIR_DADOSADV("B5_LARGCH")%>&nbsp;x&nbsp;<%=PIR_DADOSADV("B5_COMPCH")%></td>
                        <td><strong>Arranjo(LxC):</strong></td>
                        <td colspan="4" class="BordaDireita" ><%=PIR_DADOSADV("B5_ARRL")%>&nbsp;x&nbsp;<%=PIR_DADOSADV("B5_ARRC")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Área:</strong></td>
                        <td><%=PIR_DADOSADV("B5_AREA")%></td>
                        <td><strong>Área Liq.:</strong></td>
                        <td colspan="4" class="BordaDireita"><%=PIR_DADOSADV("B5_AREALI") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Blank(LxC):</strong></td>
                        <td><%=PIR_DADOSADV("B5_BLANKL")%>&nbsp;x&nbsp;<%=PIR_DADOSADV("B5_BLANKC")%></td>
                        <td><strong>Area Limpa:</strong></td>
                        <td colspan="4" class="BordaDireita"><%=PIR_DADOSADV("B5_AREA")%></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Peso Bruto:</strong></td>
                        <td><%=PIR_DADOSADV("B5_PESOB")%></td>
                        <td colspan="2"><strong>Peso Liq.:</strong></td>
                        <td colspan="3" class="BordaDireita"><%=PIR_DADOSADV("B5_PESOL")%></td>
                      </tr>
                      <tr>
                        <td colspan="7" class="BordaEsquerda BordaDireita text-center"><strong>Vinco</strong></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda text-center"><strong>1</strong></td>
                        <td class="text-center"><strong>2</strong></td>
                        <td class="text-center"><strong>3</strong></td>
                        <td class="text-center"><strong>4</strong></td>
                        <td class="text-center"><strong>5</strong></td>
                        <td class="text-center"><strong>6</strong></td>
                        <td class="BordaDireita text-center"><strong>7</strong></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda text-center"><%=PIR_DADOSADV("B5_VINCO1") %></td>
                        <td class="text-center"><%=PIR_DADOSADV("B5_VINCO2") %></td>
                        <td class="text-center"><%=PIR_DADOSADV("B5_VINCO3") %></td>
                        <td class="text-center"><%=PIR_DADOSADV("B5_VINCO4") %></td>
                        <td class="text-center"><%=PIR_DADOSADV("B5_VINCO5") %></td>
                        <td class="text-center"><%=PIR_DADOSADV("B5_VINCO6") %></td>
                        <td class="BordaDireita text-center"><%=PIR_DADOSADV("B5_VINCO7") %></td>
                      </tr>
                    </tbody>
                  </table>
                </div> 
                <div class="col-md-10">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="3" class="text-center" style="border-top:2px solid #ddd">ESPECIFICAÇÕES</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td class="BordaEsquerda">Descrição</td>
                        <td class="text-center">SMURFIT WESTROCK <%=DescricaoFilial%></td>
                        <td class="BordaDireita text-center">Cliente</td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Gramatura</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_GESTGER") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_GRAM") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Coluna[min]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_COLUGER") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_COLUNA") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Mullen[min]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_MULLGER") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_MULLEN") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Esmagamento[min]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_ESMAGER") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_ESMAG") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Espessura[min]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_ESPEGER") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_ESPES2") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Cobb Int.[max]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_COBBIGE") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_COBINT") %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda">Cobb Ext.[max]</td>
                        <td class="text-right"><%=PIR_DADOSADV("Z1_COBBEGE") %></td>
                        <td class="BordaDireita text-right"><%=PIR_DADOSADV("B5_COBEXT") %></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda">Compressão</td>
                        <td colspan="2" class="BordaDireita text-right"><%=PIR_DADOSADV("B5_COMPRE") %></td>
                      </tr>
                    </tbody>
                  </table>
                </div>  
                <div class="col-md-10">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th colspan="4" class="text-center" style="border-top:2px solid #ddd">ACESSÓRIO</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda">Vinco:</td>
                        <td><%=PIR_DADOSADV("B5_VINCACE") %></td>
                        <td>Cortes:</td>
                        <td class="BordaDireita"><%=PIR_DADOSADV("B5_CORTACE") %></td>
                      </tr>
                    </tbody>
                  </table>
                </div>                                                      
              </td>
            </tr>
            <tr>
              <td colspan="2" style="vertical-align:top">
                <div class="col-md-11">
                  <table class="table table-condensed table-striped">
                    <thead>
                      <tr class="BordaSuperior">
                        <td class="BordaEsquerda"><strong>Obs.PCP:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBS1")) %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Obs.OND:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBS2")) %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Obs.IMP1:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBS3")) %></td>
                      </tr>
                      <tr>
                        <td class="BordaEsquerda"><strong>Obs.IMP2:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBS4")) %></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Obs.Int.DDP:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBSDDP")) %></td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="BordaEsquerda"><strong>Obs.Expedição:</strong></td>
                        <td class="BordaDireita"><%=Trim(PIR_DADOSADV("B5_OBS5")) %></td>
                      </tr>
                    </thead>
                  </table>
                </div>
              </td>
            </tr>
            <tr>
              <td colspan="2" style="vertical-align:top">
                <div class="col-md-11 MargemBotton">
                  <table class="table table-bordered table-condensed table-striped">
                    <thead>
                      <tr class="CabecalhoDestaque">
                        <th class="text-center">Dados Entrada</th>
                        <th class="text-center">Análise Crítica</th>
                        <th class="text-center">Conferido</th>
                        <th colspan="2" class="text-center">Aprovação</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td rowspan="3" style="vertical-align:middle" class="text-center"><%=AcertaData(PIR_DADOSADV("B5_DTRECSA")) %></td>
                        <td></td>
                        <td><%=PIR_DADOSADV("B5_USUARIO") %></td>
                        <td><%=PIR_DADOSADV("B5_RESPONS") %></td>
                        <td><%=AcertaData(PIR_DADOSADV("B5_DTAPRFT"))%></td>
                      </tr>
                      <tr>
                        <td colspan="2" class="text-center">Data</td>
                        <td rowspan="2" class="text-right">
                          <em>
                            Critério de aceitação (NQA)<br>
                            Critico 1,5%<br>
                            Grave 4%
                          </em>
                        </td>
                        <td rowspan="3" class="text-right">
                          <em>
                            -RGQ309rev04<br>
                            Data original 03/09/2013<br>
                            Data revisão 27/08/2024
                          </em>
                        </td>
                      </tr>
                      <tr class="BordaInferior">
                        <td class="text-center"></td>
                        <td class="text-center"><%=AcertaData(PIR_DADOSADV("B5_DTCADFT")) %></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </td>
            </tr>
          </table>
        </div>
      </div>

      <div>
        <div class="col-md-10 text-center" id="printable">
              
        </div>
      </div>
        
    <% else %>
      <div class="row">
        <div class="col-md-10">
          <div class="alert alert-warning">
            <h4><strong>Alerta!</strong> Ficha Técnica Não Encontrada.</h4> <a href="#" onclick="history.go(-1)">Voltar</a>
          </div>
        </div>
      </div>
    <% end if %>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>

  </body>
</html>