<%
Response.Buffer   = True
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID      = 1046
%>
<!-- #include file = "../../../cnn2/cnn2.asp" -->
<!-- #include file = "../../../funcoes/funcoes.asp" -->
<!-- #include file = "../../../includes/formato.asp" -->
<%
Function Converte(Data)
  Converte = Mid(Data,7,2) & "/" & Mid(Data,5,2) & "/" & Mid(Data,1,4)
End Function

Dim NumFt, rsFT, SQL

NumFt = Request.QueryString("FT")

If Request("filial") = "00" Or Request("filial") = "04" Then

  SQL = "EXECUTE SP_DETALHEFT [" & Request.QueryString("FT") & "]"
  Set rsFT = Conecta("PIR-DADOSADV", SQL)

  If Not IsObject(rsFT) Or (rsFT Is Nothing) Then Response.End
  If rsFT.EOF Then
%>
    <script>alert("NÃO HÁ DETALHES PARA ESSA FT"); history.go(-1);</script>
<%
    Response.End
  End If

Else

  SQL = "" &_
    "SELECT " &_
      "ITE.Item                     AS B1_COD, " &_
      "DataCriacao                  AS DataCriacao, " &_
      "Referencia                   AS B1_DESC, " &_
      "LarguraInterna               AS B5_LARG, " &_
      "ISNULL(ComprimentoInterno,0) AS B5_COMPR, " &_
      "ISNULL(AlturaInterna,0)      AS B5_ALTU, " &_
      "0                            AS area_liquida, " &_
      "Composicao                   AS B5_ESTRUTU, " &_
      "''                           AS situacao, " &_
      "PecasPorConjunto             AS PecasPorConjunto, " &_
      "ITE.Largura                  AS Largura, " &_
      "ITE.Comprimento              AS Comprimento, " &_
      "ITE.MultLarg                 AS MultLarg, " &_
      "ITE.MultComp                 AS MultComp, " &_
      "0                            AS AreaLiquidaPeca, " &_
      "ISNULL(AreaBrutaPeca,0)      AS AreaBrutaPeca, " &_
      "0                            AS LargPeca, " &_
      "0                            AS CompPeca, " &_
      "0                            AS peso_liquido, " &_
      "0                            AS PesoBruto, " &_
      "0                            AS PesoCaixa, " &_
      "ISNULL(VincoComp1,0)         AS VincoComp1, " &_
      "ISNULL(VincoComp2,0)         AS VincoComp2, " &_
      "ISNULL(VincoComp3,0)         AS VincoComp3, " &_
      "ISNULL(VincoComp4,0)         AS VincoComp4, " &_
      "ISNULL(VincoComp5,0)         AS VincoComp5, " &_
      "ISNULL(PecasPorPacote,0)     AS PecasPorPacote, " &_
      "0                            AS LarguraPalete, " &_
      "0                            AS ComprimentoPalete, " &_
      "''                           AS TipoPalete, " &_
      "0                            AS AlturaPaleteFechado, " &_
      "0                            AS UnidadesPorPalete, " &_
      "ISNULL(FitilhosLargPacote,0) AS FitilhosLargPacote, " &_
      "ISNULL(FitilhosCompPacote,0) AS FitilhosCompPacote, " &_
      "FacesOpostas                 AS FacesOpostas, " &_
      "Cantoneiras                  AS Cantoneiras, " &_
      "Espelho                      AS Espelho, " &_
      "Filme                        AS Filme, " &_
      "ROT.Maquina1                 AS Maquina1, " &_
      "ROT.Maquina2                 AS Maquina2, " &_
      "ROT.Maquina3                 AS Maquina3, " &_
      "ROT.Maquina4                 AS Maquina4, " &_
      "ROT.Maquina5                 AS Maquina5, " &_
      "ROT.Maquina6                 AS Maquina6, " &_
      "ISNULL(VincoLarg1,0)         AS VincoLarg1, " &_
      "ISNULL(VincoLarg2,0)         AS VincoLarg2, " &_
      "ISNULL(VincoLarg3,0)         AS VincoLarg3, " &_
      "ISNULL(VincoLarg4,0)         AS VincoLarg4, " &_
      "ISNULL(VincoLarg5,0)         AS VincoLarg5, " &_
      "ISNULL(Cor1,'')              AS Cor1, " &_
      "ISNULL(Cor2,'')              AS Cor2, " &_
      "ISNULL(Cor3,'')              AS Cor3, " &_
      "ISNULL(Cor4,'')              AS Cor4, " &_
      "ISNULL(Cor5,'')              AS Cor5, " &_
      "Fechamento                   AS Fechamento, " &_
      "ISNULL(ColunaMinimo,0)       AS ColunaMinimo, " &_
      "ISNULL(MullenMinimo,0)       AS MullenMinimo, " &_
      "0                            AS Compressao, " &_
      "ISNULL(Espessura,0)          AS Espessura, " &_
      "ISNULL(CobbIntMaximo,0)      AS CobbIntMaximo, " &_
      "ISNULL(CobbExtMaximo,0)      AS CobbExtMaximo, " &_
      "''                           AS LastroPalete, " &_
      "''                           AS TipoObs1, " &_
      "ISNULL(Obs1,'')              AS Obs1, " &_
      "''                           AS TipoObs2, " &_
      "ISNULL(Obs2,'')              AS Obs2, " &_
      "''                           AS TipoObs3, " &_
      "ISNULL(Obs3,'')              AS Obs3, " &_
      "''                           AS Observacao, " &_
      "''                           AS Repreentante, " &_
      "TipoABNT                     AS Estilo, " &_
      "TolMais                      AS TolMais, " &_
      "TolMenos                     AS TolMenos, " &_
      "TextoEstadoFT_Detec, " &_
      "'UN' UM, " &_
      "'' TES, " &_
      "0  IPI, " &_
      "0  PRECO, " &_
      "'' ULTIMACOMPRA, " &_
      "'' VENDEDOR, " &_
      "'' IDCliente, " &_
      "'' B5_SAP, " &_
      "'' Tratamento, " &_
      "'' Cliche, " &_
      "'' ValorMatriz, " &_
      "'' ValorCliche, " &_
      "Paletizado, " &_
      "'' Tampa, " &_
      "ITE.IDPalete, " &_
      "PacotesAltura, " &_
      "PecasPorPalete, " &_
      "FitasLargPalete, " &_
      "FitasCompPalete, " &_
      "'' Gramatura, " &_
      "CrushMinimo, " &_
      "'' ObsVen, " &_
      "'' DataRecebimentoSAP, " &_
      "'' DataAprovacaoFT, " &_
      "'' ResponsavelAprovacaoFT, " &_
      "'' ResponsavelCadastroFT, " &_
      "'' VincoAce, " &_
      "'' CorteAce, " &_
      "'' DescricaoPalete, " &_
      "'' TipoProduto, " &_
      "'' OBS, " &_
      "'' COMPRE, " &_
      "'' GramaturaComposicao " &_
    "FROM Itens ITE " &_
      "LEFT JOIN dbo.Roteiro ROT ON ITE.Item     = ROT.Item " &_
      "LEFT JOIN dbo.Paletes PAL ON ITE.IDPalete = PAL.IDPalete " &_
    "WHERE ITE.Item = '" & Request.QueryString("FT") & "'"

  Select Case Request("filial")
    Case "05" : Set rsFT = Conecta("BEN-TRIMBOX", SQL)
    Case "06" : Set rsFT = Conecta("FOR-TRIMBOX", SQL)
    Case "07" : Set rsFT = Conecta("SAQ-TRIMBOX", SQL)
  End Select

  If Not IsObject(rsFT) Or (rsFT Is Nothing) Then Response.End
  If rsFT.EOF Then
%>
    <script>alert("NÃO HÁ DETALHES PARA ESSA FT"); history.go(-1);</script>
<%
    Response.End
  End If

End If

Dim sStatus, sBadgeCss
Select Case rsFT("B1_STATUS")
  Case "P" : sStatus = "Projeto"    : sBadgeCss = "bg-secondary"
  Case "V" : sStatus = "Verificado" : sBadgeCss = "bg-primary"
  Case "C" : sStatus = "Cancelado"  : sBadgeCss = "bg-danger"
  Case "W" : sStatus = "Validado"   : sBadgeCss = "bg-success"
  Case Else : sStatus = rsFT("B1_STATUS") : sBadgeCss = "bg-light text-primary"
End Select
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SWnet — Detalhes da FT</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
      font-size: .82rem;
      color: #212529;
    }

    .ft-sheet {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 8px;
      box-shadow: 0 2px 12px rgba(0,0,0,.09);
      overflow: hidden;
    }

    .ft-sheet-header {
      background-color: #fff;
      color: #212529;
      padding: .85rem 1.25rem;
      border-bottom: 2px solid #dee2e6;
      display: flex;
      align-items: center;
      gap: .75rem;
    }
    .ft-sheet-header .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }
    .ft-sheet-header .ft-ref {
      font-size: .78rem;
      color: #0004ff;
      font-style: italic;
    }
    .ft-sheet-header .ft-code {
      font-size: .95rem;
      font-weight: 700;
      letter-spacing: .06em;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 5px;
      padding: .25rem .65rem;
      color: #212529;
    }

    .ft-sheet-body {
      padding: 1rem 1.25rem 1.4rem;
      display: flex;
      flex-direction: column;
      gap: 1rem;
    }

    .ft-section {
      border: 1px solid #c8d6e5;
      border-radius: 6px;
      overflow: hidden;
    }

    .ft-section-title {
      background-color: #e8f0fe;
      border-bottom: 1px solid #c8d6e5;
      padding: .42rem .85rem;
      font-size: .7rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .09em;
      color: #3a5a8a;
    }

    .ft-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
    }
    .ft-grid.cols-4 { grid-template-columns: repeat(4, 1fr); }
    .ft-grid.cols-1 { grid-template-columns: 1fr; }

    .ft-field {
      padding: .42rem .75rem;
      border-right: 1px solid #e9ecef;
      border-bottom: 1px solid #e9ecef;
    }
    .ft-field:last-child,
    .ft-field.ft-last  { border-bottom: none; }
    .ft-field.span-2   { grid-column: span 2; border-right: none; }
    .ft-field.span-3   { grid-column: span 3; border-right: none; }
    .ft-field.span-4   { grid-column: span 4; border-right: none; }
    .ft-grid.cols-1 .ft-field { border-right: none; }

    .ft-label {
      display: block;
      font-size: .67rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: #6c757d;
      margin-bottom: .1rem;
      white-space: nowrap;
    }
    .ft-value {
      display: block;
      font-size: .82rem;
      font-weight: 500;
      color: #212529;
      min-height: 1.1rem;
    }
    .ft-value.empty { color: #ced4da; }

    .ft-sign-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
    }
    .ft-sign-col { border-right: 1px solid #e9ecef; }
    .ft-sign-col:last-child { border-right: none; }

    .ft-sign-head {
      background-color: #f7f5f5d7;
      color: #000000;
      font-size: .69rem;
      font-weight: 700;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: .05em;
      padding: .38rem .5rem;
    }
    .ft-sign-cell {
      border-top: 1px solid #e9ecef;
      padding: .38rem .65rem;
    }
    .ft-sign-cell-label {
      font-size: .64rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .04em;
      color: #6c757d;
    }
    .ft-sign-cell-value {
      font-size: .8rem;
      font-weight: 500;
      color: #212529;
      min-height: 1rem;
    }
    .ft-sign-blank {
      border-top: 1px solid #e9ecef;
      height: 44px;
    }

    .ft-sheet-footer {
      border-top: 1px solid #dee2e6;
      background-color: #f8f9fa;
      padding: .6rem 1.25rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      font-size: .75rem;
      color: #ff0303;
    }

    @media print {
      @page {
        size: A4 portrait;
        margin: 5mm;
      }
      html, body {
        background-color: #fff !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      .no-print { display: none !important; }

      .container-fluid {
        max-width: 100% !important;
        padding: 0 !important;
        margin: 0 !important;
      }

      .ft-sheet {
        box-shadow: none !important;
        border-radius: 0 !important;
        border: none !important;
        font-size: .56rem !important;
      }

      .ft-sheet-header {
        padding: .25rem .5rem !important;
        border-bottom: 1px solid #dee2e6 !important;
      }
      .ft-sheet-header .page-title { font-size: .78rem !important; }
      .ft-sheet-header .ft-ref     { font-size: .58rem !important; }
      .ft-sheet-header .ft-code    { font-size: .68rem !important; padding: .1rem .3rem !important; }

      .ft-sheet-body {
        padding: .25rem .5rem .4rem !important;
        gap: .22rem !important;
      }

      .ft-section { border-radius: 2px !important; }

      .ft-section-title {
        padding: .15rem .45rem !important;
        font-size: .5rem !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }

      .ft-field {
        padding: .13rem .35rem !important;
      }

      .ft-label {
        font-size: .48rem !important;
        margin-bottom: .03rem !important;
      }

      .ft-value {
        font-size: .58rem !important;
        min-height: .65rem !important;
      }

      .ft-sign-head {
        font-size: .5rem !important;
        padding: .15rem .25rem !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .ft-sign-cell       { padding: .13rem .35rem !important; }
      .ft-sign-cell-label { font-size: .46rem !important; }
      .ft-sign-cell-value { font-size: .57rem !important; min-height: .6rem !important; }
      .ft-sign-blank      { height: 16px !important; }

      .ft-sheet-footer { display: none !important; }
    }
  </style>
</head>
<body>
  
  <div class="container-fluid py-3 px-3 px-lg-4" style="max-width: 860px;">
    <div class="ft-sheet">
      
      <div class="ft-sheet-header">
        <button class="btn btn-sm btn-light border no-print" onclick="history.go(-1)" title="Voltar">
          <i class="bi bi-arrow-left"></i>
        </button>
    <i class="bi bi-file-earmark-ruled fs-4 flex-shrink-0 text-secondary"></i>
    <div>
      <div class="page-title">Ficha Técnica</div>
      <div class="ft-ref"><%=rsFT("B1_DESC")%></div>
    </div>
    <div class="ms-auto d-flex align-items-center gap-2 flex-shrink-0">
      <span class="badge <%=sBadgeCss%> px-3 py-2" style="font-size:.8rem;"><%=sStatus%></span>
      <span class="ft-code"><%=rsFT("B1_COD")%></span>
    </div>
  </div>

  <div class="ft-sheet-body">

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-info-circle me-1"></i>&emsp;Identificação</div>
      <div class="ft-grid cols-4">
        <div class="ft-field ft-last">
          <span class="ft-label">Emissão</span>
          <span class="ft-value"><%=Date()%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">SAP</span>
          <span class="ft-value"><%=rsFT("B5_SAP")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Data Receb.</span>
          <span class="ft-value"><%=Converte(rsFT("B5_DTRECSA"))%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Qtd. Amostra</span>
          <span class="ft-value empty">—</span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-box-seam me-1"></i>&emsp;Dados do Produto</div>
      <div class="ft-grid cols-4">
        <div class="ft-field span-2">
          <span class="ft-label">Cliente</span>
          <span class="ft-value"><%=rsFT("B5_CLI") & " — " & rsFT("B5_NOMECLI")%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Referência</span>
          <span class="ft-value"><%=rsFT("B1_DESC")%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Representante</span>
          <span class="ft-value"><%=rsFT("B5_NOMVEN") & " — " & rsFT("B5_VEND")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Loja</span>
          <span class="ft-value"><%=rsFT("B5_LOJA")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Estilo</span>
          <span class="ft-value"><%=rsFT("B5_ESTILO")%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Medidas Internas (C x L x A)</span>
          <span class="ft-value"><%=rsFT("B5_COMPR") & " x " & rsFT("B5_LARG") & " x " & rsFT("B5_ALTU")%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Preço</span>
          <span class="ft-value">R$ <%=FormatNumber(rsFT("B1_UPRC"), 2)%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Tolerância (+)</span>
          <span class="ft-value"><%=rsFT("B5_TOLMA")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Tolerância (−)</span>
          <span class="ft-value"><%=rsFT("B5_TOLME")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Peças / Conjunto</span>
          <span class="ft-value"><%=rsFT("B5_PCSCON")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Despacho</span>
          <span class="ft-value empty">—</span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-layers me-1"></i>&emsp;Onduladeira</div>
      <div class="ft-grid cols-4">
        <div class="ft-field span-2">
          <span class="ft-label">Estrutura</span>
          <span class="ft-value"><%=rsFT("B5_ESTRUTU")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Área Limpa</span>
          <span class="ft-value"><%=vAreaLimpa%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Área Bruta</span>
          <span class="ft-value"><%=rsFT("B5_AREA")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Chapa (L x C)</span>
          <span class="ft-value"><%=IIf(Request("filial")="00" Or Request("filial")="04", rsFT("B5_LARGCH") & " x " & rsFT("B5_COMPCH"), rsFT("B5_LARG") & " x " & rsFT("B5_COMPR"))%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Arranjo (L x C)</span>
          <span class="ft-value"><%=rsFT("B5_ARRL") & " x " & rsFT("B5_ARRC")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Blank (L x C)</span>
          <span class="ft-value"><%=rsFT("B5_BLANKL") & " x " & rsFT("B5_BLANKC")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Peso Bruto</span>
          <span class="ft-value"><%=rsFT("B5_PESOB")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Peso Líquido</span>
          <span class="ft-value"><%=rsFT("B5_PESOL")%></span>
        </div>
        <div class="ft-field span-4 ft-last">
          <span class="ft-label">Vincos</span>
          <span class="ft-value"><%=rsFT("B5_VINCO1") & " x " & rsFT("B5_VINCO2") & " x " & rsFT("B5_VINCO3") & " x " & rsFT("B5_VINCO4") & " x " & rsFT("B5_VINCO5") & " x " & rsFT("B5_VINCO6") & " x " & rsFT("B5_VINCO7")%></span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-box me-1"></i>&emsp;Acabamento</div>
      <div class="ft-grid cols-4">
        <div class="ft-field">
          <span class="ft-label">Pacote</span>
          <span class="ft-value"><%=rsFT("B5_AMARRAD") & " — " & rsFT("B5_QTDPCT")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Fitilhos (L x C)</span>
          <span class="ft-value"><%=rsFT("B5_FITLAR") & " x " & rsFT("B5_FITCOM")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Fitas (L x C)</span>
          <span class="ft-value"><%=rsFT("B5_FITASL") & " x " & rsFT("B5_FICASC")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Faces Opostas</span>
          <span class="ft-value"><%=rsFT("B5_FACOP")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Paletizado</span>
          <span class="ft-value"><%=rsFT("B5_PALETE")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Tipo Palete</span>
          <span class="ft-value"><%=rsFT("B5_TPPALE") & " — " & rsFT("B5_DESCPAL")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Lastro Palete</span>
          <span class="ft-value"><%=rsFT("B5_PLLAST")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Altura Palete</span>
          <span class="ft-value"><%=rsFT("B5_ARRPA")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Qtd. p/ Palete</span>
          <span class="ft-value"><%=rsFT("B5_PCSPAL")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Cantoneiras</span>
          <span class="ft-value"><%=rsFT("B5_CANTON")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Espelho</span>
          <span class="ft-value"><%=rsFT("B5_ESPELHO")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Tampa</span>
          <span class="ft-value"><%=rsFT("B5_TAMPA")%></span>
        </div>
        <div class="ft-field span-4">
          <span class="ft-label">Filme</span>
          <span class="ft-value"><%=rsFT("B5_FILME")%></span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-printer me-1"></i>&emsp;Impressora</div>
      <div class="ft-grid cols-4">
        <div class="ft-field span-4">
          <span class="ft-label">Seq. Máquinas</span>
          <span class="ft-value"><%=rsFT("B5_MAQ1") & "  " & rsFT("B5_MAQ2") & "  " & rsFT("B5_MAQ3") & "  " & rsFT("B5_MAQ4") & "  " & rsFT("B5_MAQ5") & "  " & rsFT("B5_MAQ6")%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Vincos (impress.)</span>
          <span class="ft-value"><%=FormatNumber(rsFT("B5_VINCI1"),0) & " x " & FormatNumber(rsFT("B5_VINCI2"),0) & " x " & FormatNumber(rsFT("B5_VINCI3"),0) & " x " & FormatNumber(rsFT("B5_VINCI4"),0) & " x " & FormatNumber(rsFT("B5_VINCI5"),0)%></span>
        </div>
        <div class="ft-field span-2">
          <span class="ft-label">Cores</span>
          <span class="ft-value"><%=rsFT("B5_COR1") & "  " & rsFT("B5_COR2") & "  " & rsFT("B5_COR3") & "  " & rsFT("B5_COR4")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Fechamento</span>
          <span class="ft-value"><%=rsFT("B5_FECHAM")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Clichê</span>
          <span class="ft-value"><%=rsFT("B5_NUMCLI")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Valor Clichê</span>
          <span class="ft-value"><%=FormatNumber(rsFT("B5_VLMATRI"), 2)%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Matriz</span>
          <span class="ft-value"><%=rsFT("B5_NUMFACA")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Valor Matriz</span>
          <span class="ft-value"><%=FormatNumber(rsFT("B5_VLCLICH"), 2)%></span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-clipboard-check me-1"></i>&emsp;Especificações do Cliente</div>
      <div class="ft-grid cols-4">
        <div class="ft-field">
          <span class="ft-label">Coluna [mín]</span>
          <span class="ft-value"><%=rsFT("B5_COLUNA")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Mullen [mín]</span>
          <span class="ft-value"><%=rsFT("B5_MULLEN")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Esmag. [mín]</span>
          <span class="ft-value"><%=rsFT("B5_ESMAG")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Espessura [mín]</span>
          <span class="ft-value"><%=rsFT("B5_ESPESS")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Cobb Int. [máx]</span>
          <span class="ft-value"><%=rsFT("B5_COBINT")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Cobb Ext. [máx]</span>
          <span class="ft-value"><%=rsFT("B5_COBEXT")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Vincos Acess.</span>
          <span class="ft-value"><%=rsFT("B5_VINCACE")%></span>
        </div>
        <div class="ft-field ft-last">
          <span class="ft-label">Corte Acess.</span>
          <span class="ft-value"><%=rsFT("B5_CORTACE")%></span>
        </div>
        <div class="ft-field span-4">
          <span class="ft-label">Compressão</span>
          <span class="ft-value"><%=rsFT("B5_COMPRE")%></span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-chat-left-text me-1"></i>&emsp;Observações</div>
      <div class="ft-grid cols-1">
        <div class="ft-field">
          <span class="ft-label">Obs. Vendas</span>
          <span class="ft-value"><%=rsFT("B5_OBSVEN")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Obs. Técnica 1</span>
          <span class="ft-value"><%=rsFT("B5_OBS1")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Obs. Técnica 2</span>
          <span class="ft-value"><%=rsFT("B5_OBS2")%></span>
        </div>
        <div class="ft-field">
          <span class="ft-label">Obs. Técnica 3</span>
          <span class="ft-value"><%=rsFT("B5_OBS3")%></span>
        </div>
      </div>
    </div>

    <div class="ft-section">
      <div class="ft-section-title"><i class="bi bi-pen me-1"></i>&emsp;Assinaturas</div>
      <div class="ft-sign-grid">
        <div class="ft-sign-col">
          <div class="ft-sign-head">Análise Crítica</div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Responsável</div>
            <div class="ft-sign-cell-value empty">—</div>
          </div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Data</div>
            <div class="ft-sign-cell-value empty">—</div>
          </div>
          <div class="ft-sign-blank"></div>
        </div>
        <div class="ft-sign-col">
          <div class="ft-sign-head">Dados Entrada</div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Responsável</div>
            <div class="ft-sign-cell-value empty">—</div>
          </div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Data</div>
            <div class="ft-sign-cell-value"><%=Converte(rsFT("B5_DTCADFT"))%></div>
          </div>
          <div class="ft-sign-blank"></div>
        </div>
        <div class="ft-sign-col">
          <div class="ft-sign-head">Digitação</div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Responsável</div>
            <div class="ft-sign-cell-value"><%=rsFT("B5_USUARIO")%></div>
          </div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Data</div>
            <div class="ft-sign-cell-value"><%=Converte(rsFT("B5_DTCADFT"))%></div>
          </div>
          <div class="ft-sign-blank"></div>
        </div>
        <div class="ft-sign-col">
          <div class="ft-sign-head">Verificação</div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Responsável</div>
            <div class="ft-sign-cell-value"><%=rsFT("B5_RESPONS")%></div>
          </div>
          <div class="ft-sign-cell">
            <div class="ft-sign-cell-label">Data</div>
            <div class="ft-sign-cell-value"><%=Converte(rsFT("B5_DTAPRFT"))%></div>
          </div>
          <div class="ft-sign-blank"></div>
        </div>
      </div>
    </div>
  </div>

  <div class="ft-sheet-footer no-print">
    <span>
      <i class="bi bi-calendar3 me-1"></i>
      Emitido em <%=Right("0" & Day(Now()), 2) & "/" & Right("0" & Month(Now()), 2) & "/" & Year(Now())%>
      às <%=Hour(Now()) & ":" & Right("0" & Minute(Now()), 2)%>
    </span>
    <div class="d-flex gap-2">
      <button class="btn btn-sm btn-danger" onclick="window.print()">
        <i class="bi bi-printer me-1"></i>Imprimir
      </button>
    </div>
  </div>

</div>
</div>

<%
rsFT.Close()
Set rsFT = Nothing
%>

<script>
  try { sessionStorage.setItem('ft_origem_ativa', '1'); } catch (_) {}
</script>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
</body>
</html>
