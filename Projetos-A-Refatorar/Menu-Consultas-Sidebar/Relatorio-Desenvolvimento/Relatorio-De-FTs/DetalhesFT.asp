<%
Response.Buffer   = True
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID       = 1046
%>
<!-- #include file = "../../../cnn2/cnn2.asp"-->
<!-- #include file = "../../../includes/variaveis.asp" -->
 <!DOCTYPE html>
 <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>SWnet - Detalhe da FT</title>

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
            gap: .75rem
        }
        .ft-sheet-header .page-title {
            font-size: 1.15rem;
            font-weight: 700;
            color: 212529;
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
        .ft-grid.cols-4 {
             grid-template-columns: repeat(4, 1fr);
        }
        .ft-grid.cols-1 {
            grid-template-columns: 1fr;
        }

        .ft-field {
            padding: .42rem .75rem;
            border-right: 1px solid #e9ecef;
            border-bottom: 1px solid #e9ecef;
        }
        .ft-field:last-child,
        .ft-field.ft-last { border-bottom: none; }
        .ft-field.span-2 { grid-column: span 2; border-right: none; }
        .ft-field.span-4 { grid-column: span 4; border-right: none; }
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
        .ft-value.empty { 
            color: #ced4da;
        }

        @media print {
            @page {
                size: A4;
                margin: 10mm;
            }
            body { 
                background-color: #fff !important;
                font-size: .75rem;
            }
            .no-print {
                display: none !important;
            }
            .ft-sheet {
                box-shadow: none;
            }
            .ft-section-title {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }

    </style>
</head>
<body>

    <div class="container-fluid py-3 px-3 px-lg-4" style="max-width:900px;">
        <%
        If request("planta") = "00" Or request("planta") = "04" Then

        SQL = "" &_
    "SELECT " &_
      "t.DataCriacao, " &_
      "p.codigo, " &_
      "p.referencia, " &_
      "t.LarguraInterna, " &_
      "t.ComprimentoInterno, " &_
      "t.Alturainterna, " &_
      "p.area_liquida, " &_
      "p.composicao, " &_
      "p.situacao, " &_
      "cast(t.CodRepresentante as varchar(10)) + ' - ' + t.representante AS Representante, " &_
      "CodSK + ' - ' + DescTipoABNT AS Estilo, " &_
      "TolMais, " &_
      "TolMenos, " &_
      "PecasPorConjunto, " &_
      "t.Largura, " &_
      "t.Comprimento, " &_
      "MultLarg, " &_
      "MultComp, " &_
      "isnull(AreaLiquidaPeca,0) AS AreaLiquidaPeca, " &_
      "isnull(AreaBrutaPeca,0)  AS AreaBrutaPeca, " &_
      "isnull(LargPeca,0)       AS LargPeca, " &_
      "isnull(CompPeca,0)       AS CompPeca, " &_
      "isnull(peso_liquido,0)   AS peso_liquido, " &_
      "isnull(PesoCaixa,0)      AS PesoCaixa, " &_
      "isnull(VincoComp1,0)     AS VincoComp1, " &_
      "isnull(VincoComp2,0)     AS VincoComp2, " &_
      "isnull(VincoComp3,0)     AS VincoComp3, " &_
      "isnull(VincoComp4,0)     AS VincoComp4, " &_
      "isnull(VincoComp5,0)     AS VincoComp5, " &_
      "isnull(PecasPorPacote,0) AS PecasPorPacote, " &_
      "LarguraPalete, " &_
      "ComprimentoPalete, " &_
      "TipoPalete, " &_
      "AlturaPaleteFechado, " &_
      "UnidadesPorPalete, " &_
      "isnull(FitilhosLargPacote,0) AS FitilhosLargPacote, " &_
      "isnull(FitilhosCompPacote,0) AS FitilhosCompPacote, " &_
      "FacesOpostas, " &_
      "Cantoneiras, " &_
      "Espelho, " &_
      "Filme, " &_
      "isnull(Maquina1,'') AS Maquina1, " &_
      "isnull(Maquina2,'') AS Maquina2, " &_
      "isnull(Maquina3,'') AS Maquina3, " &_
      "isnull(Maquina4,'') AS Maquina4, " &_
      "isnull(Maquina5,'') AS Maquina5, " &_
      "isnull(VincoLarg1,0) AS VincoLarg1, " &_
      "isnull(VincoLarg2,0) AS VincoLarg2, " &_
      "isnull(VincoLarg3,0) AS VincoLarg3, " &_
      "isnull(VincoLarg4,0) AS VincoLarg4, " &_
      "isnull(VincoLarg5,0) AS VincoLarg5, " &_
      "isnull(Cor1,'') AS Cor1, " &_
      "isnull(Cor2,'') AS Cor2, " &_
      "isnull(Cor3,'') AS Cor3, " &_
      "isnull(Cor4,'') AS Cor4, " &_
      "isnull(Cor5,'') AS Cor5, " &_
      "Fechamento, " &_
      "isnull(ColunaMinimo,0)  AS ColunaMinimo, " &_
      "isnull(MullenMinimo,0)  AS MullenMinimo, " &_
      "isnull(Compressao,0)    AS Compressao, " &_
      "isnull(Espessura,0)     AS Espessura, " &_
      "isnull(CobbIntMaximo,0) AS CobbIntMaximo, " &_
      "isnull(CobbExtMaximo,0) AS CobbExtMaximo, " &_
      "isnull(PacotesLargura,0) * isnull(PacotesComprimento,0) AS Lastroalete, " &_
      "isnull(TipoObs1,'') AS TipoObs1, " &_
      "isnull(Obs1,'')     AS Obs1, " &_
      "isnull(TipoObs2,'') AS TipoObs2, " &_
      "isnull(Obs2,'')     AS Obs2, " &_
      "isnull(TipoObs3,'') AS TipoObs3, " &_
      "isnull(Obs3,'')     AS Obs3, " &_
      "isnull(TipoObs4,'') AS TipoObs4, " &_
      "isnull(Obs4,'')     AS Obs4, " &_
      "isnull(cast(Versao as varchar(9)),'') + ' - ' + ObsVersao AS Observacao, " &_
      "'' AS IDCLIENTE, '' AS CLIENTE, '' AS CodFI " &_
    "FROM dbo.interface_inpanet_produtos p " &_
      "LEFT JOIN trimboxteste..vw_ConsultaFTs t ON p.Codigo = t.Item " &_
    "WHERE p.codigo = '" & Trim(Request("id")) & "'"

Else

  SQL = "" &_
    "SELECT " &_
      "ITE.Item                     AS codigo, " &_
      "ITE.DataCriacao              AS DataCriacao, " &_
      "TextoEstadoFT_Detec          AS situacao, " &_
      "Referencia                   AS referencia, " &_
      "ITE.IDCliente                AS IDCLIENTE, " &_
      "CLI.Cliente                  AS CLIENTE, " &_
      "ComprimentoInterno, " &_
      "LarguraInterna, " &_
      "AlturaInterna, " &_
      "ITE.TolMais, " &_
      "ITE.TolMenos, " &_
      "Composicao                   AS composicao, " &_
      "TipoABNT                     AS Estilo, " &_
      "ITE.Largura                  AS Largura, " &_
      "ITE.Comprimento              AS Comprimento, " &_
      "ITE.MultLarg                 AS MultLarg, " &_
      "ITE.MultComp                 AS MultComp, " &_
      "AreaBrutaPeca, " &_
      "AreaLiquidaPeca, " &_
      "LargPeca, " &_
      "CompPeca, " &_
      "0                            AS peso_liquido, " &_
      "0                            AS PesoCaixa, " &_
      "PecasPorConjunto, " &_
      "0                            AS PecasPorPacote, " &_
      "Faca                         AS CodFI, " &_
      "Fechamento, " &_
      "Cor1, Cor2, Cor3, Cor4, Cor5, " &_
      "PAL.Largura                  AS LarguraPalete, " &_
      "PAL.Comprimento              AS ComprimentoPalete, " &_
      "PAL.Tipo                     AS TipoPalete, " &_
      "AlturaPaleteFechado, " &_
      "UnidadesPorPalete, " &_
      "FitilhosLargPacote, " &_
      "FitilhosCompPacote, " &_
      "ITE.FacesOpostas, " &_
      "ITE.Cantoneiras, " &_
      "ITE.Espelho, " &_
      "ITE.Filme, " &_
      "'' AS Lastroalete, " &_
      "ROT.Maquina1, ROT.Maquina2, ROT.Maquina3, ROT.Maquina4, ROT.Maquina5, " &_
      "VincoLarg1, VincoLarg2, VincoLarg3, VincoLarg4, VincoLarg5, " &_
      "VincoComp1, VincoComp2, VincoComp3, VincoComp4, VincoComp5, " &_
      "MullenMinimo, " &_
      "ColunaMinimo, " &_
      "CobbIntMaximo, " &_
      "CobbExtMaximo, " &_
      "Espessura, " &_
      "Compressao, " &_
      "ITE.TipoObs1, ITE.Obs1, " &_
      "ITE.TipoObs2, ITE.Obs2, " &_
      "ITE.TipoObs3, ITE.Obs3, " &_
      "'' AS TipoObs4, '' AS Obs4, " &_
      "'' AS Observacao, " &_
      "CLI.CodRepresentante         AS CodRepresentante, " &_
      "REP.Nome                     AS Representante " &_
    "FROM dbo.Itens ITE " &_
      "LEFT JOIN dbo.Roteiro  ROT ON ITE.Item     = ROT.Item " &_
      "LEFT JOIN dbo.Paletes  PAL ON ITE.IDPalete = PAL.IDPalete " &_
      "LEFT JOIN dbo.Clientes CLI ON ITE.IDCliente = CLI.IDCliente " &_
      "LEFT JOIN dbo.Represen REP ON CLI.CodRepresentante = REP.Codigo " &_
    "WHERE ITE.Item = '" & Trim(Request("id")) & "'"

End If

Select Case Request("planta")
    Case "00", "04" : Set CONSULTA = Conecta("PIR-CORPORATIVO", SQL)
    Case "05"       : Set CONSULTA = Conecta("BEN-TRIMBOX", SQL)
    Case "06"       : Set CONSULTA = Conecta("FOR-TRIMBOX", SQL)
    Case "07"       : Set CONSULTA = Conecta("SAQ-TRIMBOX", SQL)
End Select

If Not IsObject(CONSULTA) Or (CONSULTA Is Nothing) Then
    Response.End
End If

If Not CONSULTA.EOF Then
%>

<div class="ft-sheet">

    <%
    Dim sRef
    sRef = Trim(CONSULTA("referencia"))
    %>

    <div class="ft-sheet-header">
        <i class="bi bi-file-earmark-ruled fs-4 flex-shrink-0 text-secondary"></i>
        <div>
            <div class="page-title">Detalhe da FT</div>
            <div class="text-muted" style="font-size:.78rem;font-style:italic;"><%=sRef%></div>
        </div>
        <div class="ms-auto flex-shrink-0">
            <span class="ft-code"><%=Trim(CONSULTA("codigo"))%></span>
        </div>
    </div>

    <div class="ft-sheet-body">

        <div class="ft-section">
            <div class="ft-section-title">
                <i class="bi bi-info-circle me-1"></i>Identificação
            </div>
        <div class="ft-grid cols-4">
            <div class="ft-field ft-last">
                <span class="ft-label">Emissão</span>
                <span class="ft-value"><%=Trim(CONSULTA("DataCriacao"))%></span>
            </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Status</span>
            <span class="ft-value"><%=Trim(CONSULTA("situacao"))%></span>
        </div>
        <div class="ft-field span-2 ft-last">
            <span class="ft-label">Cliente</span>
            <span class="ft-value"><%=Trim(CONSULTA("IDCLIENTE")) & IIF(Trim(CONSULTA("CLIENTE")) <> "", " - " & Trim(CONSULTA("CLIENTE")), "")%></span>
        </div>
    </div>
</div>

        <div class="ft-section">
            <div class="ft-section-title">
                <i class="bi bi-box-seam me-1"></i>Dados do Produto
            </div>
        <div class="ft-grid cols-4">
            <div class="ft-field span-2">
                <span class="ft-label">Representante</span>
                <span class="ft-value"><%=Trim(Consulta("CodRepresentante")) & IIf(Trim(CONSULTA("REPRESENTANTE")) <> "", " - " & Trim(CONSULTA("Representante")), "")%></span>
            </div>
        <div class="ft-field">
            <span class="ft-label">Estilo</span>
            <span class="ft-value"><%=Trim(CONSULTA("Estilo"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Medidas Internas (C x L x A)</span>
            <span class="ft-value">
                <%
                Dim sMed
                sMed = ""
                If CInt(0 & Trim(CONSULTA("ComprimentoInterno"))) > 0 Then sMed = sMed & Trim(CONSULTA("ComprimentoInterno"))
                If CInt(0 & Trim(CONSULTA("LarguraInterna")))     > 0 Then sMed = sMed & " x " & Trim(CONSULTA("LarguraInterna"))
                If CInt(0 & Trim(CONSULTA("Alturainterna")))      > 0 Then sMed = sMed & " x " & Trim(CONSULTA("Alturainterna"))
                If sMed = "" Then sMed = "-"
                Response.Write Trim(sMed)
                %>
            </span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Tolerância (+)</span>
            <span class="ft-value"><%=Trim(CONSULTA("TolMais"))%></span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Tolerância (-)</span>
            <span class="ft-value"><%=Trim(CONSULTA("TolMenos"))%></span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Peças / Conjunto</span>
            <span class="ft-value"><%=Trim(CONSULTA("PecasPorConjunto"))%></span>
        </div>
    </div>
</div>

        <div class="ft-section">
            <div class="ft-section-title"><i class="bi bi-layers me-1"></i>Onduladeira</div>
            <div class="ft-grid cols-4">
                <div class="ft-field span-2">
                    <span class="ft-label">Estrutura</span>
                    <span class="ft-value"><%=Trim(CONSULTA("composicao"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Chapa (L x C)</span>
                    <span class="ft-value"><%=Trim(CONSULTA("Largura"))%> x <%=Trim(CONSULTA("Comprimento"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Arranjo (L x C)</span>
                    <span class="ft-value"><%=Trim(CONSULTA("MultLarg"))%> x <%=Trim(CONSULTA("MultComp"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Área Bruta</span>
                    <span class="ft-value"><%=Trim(CONSULTA("AreaBrutaPeca"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Área Limpa</span>
                    <span class="ft-value"><%=Trim(CONSULTA("AreaLiquidaPeca"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Blank (L x C)</span>
                    <span class="ft-value"><%=Trim(Consulta("LargPeca"))%> x <%=Trim(CONSULTA("CompPeca"))%></span>
                </div>
                <div class="ft-field">
                    <span class="ft-label">Peso Bruto</span>
                    <span class="ft-value"><%=Trim(CONSULTA("peso_liquido"))%></span>
                </div>
                <div class="ft-field span-4 ft-last">
                    <span class="ft-label">Vincos (Ondul.)</span>
                    <span class="ft-value"><%=Trim(CONSULTA("VincoComp1"))%> x <%=Trim(CONSULTA("VincoComp2"))%> x <%=Trim(CONSULTA("VincoComp3"))%> x <%=Trim(CONSULTA("VincoComp4"))%> x <%=Trim(CONSULTA("VincoComp5"))%></span>
                </div>
            </div>
        </div>

        <div class="ft-section">
            <div class="ft-section-title">
                <i class="bi bi-box me-1"></i>Acabamento
            </div>
        <div class="ft-grid cols-4">
            <div class="ft-field">
                <span class="ft-label">Pacote</span>
                <span class="ft-value"><%=Trim(CONSULTA("PecasPorPacote"))%></span>
            </div>
        <div class="ft-field">
            <span class="ft-label">Fitilhos (L x C)</span>
            <span class="ft-value"><%=Trim(CONSULTA("FitilhosLargPacote"))%> x <%=Trim(CONSULTA("FitilhosCompPacote"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Faces Opostas</span>
            <span class="ft-value"><%=Trim(CONSULTA("FacesOpostas"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Cantoneiras</span>
            <span class="ft-value"><%=Trim(CONSULTA("Cantoneiras"))%></span>
        </div>
        <div class ="ft-field">
            <span class="ft-label">Tipo Palete</span>
            <span class="ft-value"><%=Trim(CONSULTA("TipoPalete"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Paletizado (L x C)</span>
            <span class="ft-value"><%=Trim(CONSULTA("LarguraPalete"))%> x <%=Trim(CONSULTA("ComprimentoPalete"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Lastro do Palete</span>
            <span class="ft-value"><%=Trim(CONSULTA("Lastroalete"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Qtd. por Palete</span>
            <span class="ft-value"><%=Trim(CONSULTA("Espelho"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Espelho</span>
            <span class="ft-value"><%=Trim(Consulta("Espelho"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Tampa</span>
            <span class="ft-value empty">-</span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Filme</span>
            <span class="ft-value"><%=Trim(CONSULTA("Filme"))%></span>
        </div>
    </div>
</div>

        <div class="ft-section">
            <div class="ft-section-title">
                <i class="bi bi-printer me-1"></i>Impressora
            </div>
        <div class="ft-grid cols-4">
            <div class="ft-field span-4">
                <span class="ft-label">Sequência de Máquinas</span>
                <span class="ft-value">
                    <%
                    Dim sMaq
                    sMaq = ""
                    If Trim(CONSULTA("Maquina1")) <> "" Then sMaq = sMaq & Trim(CONSULTA("Maquina1"))
                    If Trim(CONSULTA("Maquina2")) <> "" Then sMaq = sMaq & " " & Trim(CONSULTA("Maquina2"))
                    If Trim(CONSULTA("Maquina3")) <> "" Then sMaq = sMaq & " " & Trim(CONSULTA("Maquina3"))
                    If Trim(CONSULTA("Maquina4")) <> "" Then sMaq = sMaq & " " & Trim(CONSULTA("Maquina4"))
                    If Trim(CONSULTA("Maquina5")) <> "" Then sMaq = sMaq & " " & Trim(CONSULTA("Maquina5"))
                    If sMaq = "" Then sMaq = "-"
                    Response.Write sMaq
                    %>
                </span>
            </div>
            <div class="ft-field span-2">
                <span class="ft-label">Vincos (Impress.)</span>
                <span class="ft-value">
                    <%
                    Dim sVL
                    sVL = ""
                    If CONSULTA("VincoLarg1") > 0 Then sVL = sVL & Trim(CONSULTA("VincoLarg1"))
                    If CONSULTA("VincoLarg2") > 0 Then sVL = sVL & " x " & Trim(CONSULTA("VincoLarg2"))
                    If CONSULTA("VincoLarg3") > 0 Then sVL = sVL & " x " & Trim(CONSULTA("VincoLarg3"))
                    If CONSULTA("VincoLarg4") > 0 Then sVL = sVL & " x " & Trim(CONSULTA("VincoLarg4"))
                    If CONSULTA("VincoLarg5") > 0 Then sVL = sVL & " x " & Trim(CONSULTA("VincoLarg5"))
                    If sVL = "" Then sVL = "-"
                    Response.Write sVL
                    %>
                </span>
            </div>
            <div class="ft-field span-2">
                <span class="ft-label">Cores</span>
                <span class="ft-value">
                    <%
                    Dim sCor
                    sCor = ""
                    If Trim(CONSULTA("Cor1")) <> "" Then sCor = sCor & Trim(CONSULTA("Cor1"))
                    If Trim(CONSULTA("Cor2")) <> "" Then sCor = sCor & " - " & Trim(CONSULTA("Cor2"))
                    If Trim(CONSULTA("Cor3")) <> "" Then sCor = sCor & " - " & Trim(CONSULTA("Cor3"))
                    If Trim(CONSULTA("Cor4")) <> "" Then sCor = sCor & " - " & Trim(CONSULTA("Cor4"))
                    If Trim(CONSULTA("Cor5")) <> "" Then sCor = sCor & " - " & Trim(CONSULTA("Cor5"))
                    If sCor = "" Then sCor = "-"
                    Response.Write sCor
                    %>
                </span>
            </div>
            <div class="ft-field ft-last">
                <span class="ft-label">Fechamento</span>
                <span class="ft-value"><%=Trim(CONSULTA("Fechamento"))%></span>
            </div>
            <div class="ft-field ft-last">
                <span class="ft-label">Clichê</span>
                <span class="ft-value"><%=Trim(CONSULTA("CodFI"))%></span>
            </div>
            <div class="ft-field ft-last">
                <span class="ft-label">Matriz</span>
                <span class="ft-value empty">-</span>
            </div>
            <div class="ft-field ft-last">
                <span class="ft-label">Valor Matriz</span>
                <span class="ft-value empty">-</span>
            </div>
        </div>
    </div>

    <div class="ft-section">
        <div class="ft-section-title">
            <i class="bi bi-clipboard-check me-1"></i>Especificações do Cliente
        </div>
    <div class="ft-grid cols-4">
        <div class="ft-field">
            <span class="ft-label">Coluna [mín]</span>
            <span class="ft-value"><%=Trim(CONSULTA("ColunaMinimo"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Mullen [min]</span>
            <span class="ft-value"><%=Trim(CONSULTA("MullenMinimo"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Espessura[min]</span>
            <span class="ft-value"><%=Trim(CONSULTA("Espessura"))%></span>
        </div>
        <div class="ft-field">
            <span class="ft-label">Compressão</span>
            <span class="ft-value"><%=Trim(CONSULTA("Compressao"))%></span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Cobb Int. [máx]</span>
            <span class="ft-value"><%=Trim(CONSULTA("CobbIntMaximo"))%></span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Vinco Acessórios</span>
            <span class="ft-value empty">-</span>
        </div>
        <div class="ft-field ft-last">
            <span class="ft-label">Esmagamento [min]</span>
            <span class="ft-value empty">-</span>
        </div>
    </div>
</div>

    <div class="ft-section">
        <div class="ft-section-title">
            <i class="bi bi-chat-left-text me-1"></i>Observações
        </div>
        <div class="ft-grid cols-1">
            <div class="ft-field">
                <span class="ft-label">Observações de Versão</span>
                <span class="ft-value"><%=Trim(CONSULTA("Observacao"))%></span>
            </div>
            <div class="ft-field">
                <span class="ft-label">Obs. Técnica <%=Trim(CONSULTA("TipoObs1"))%></span>
                <span class="ft-value"><%=Trim(CONSULTA("Obs1"))%></span>
            </div>
            <div class="ft-field">
                <span class="ft-label">Obs. Técnica <%=Trim(CONSULTA("TipoObs2"))%></span>
                <span class="ft-value"><%=Trim(CONSULTA("Obs2"))%></span>
            </div>
            <div class="ft-field ft-last">
                <span class="ft-label">Obs. Técnica <%=Trim(CONSULTA("TipoObs3"))%></span>
                <span class="ft-value"><%=Trim(CONSULTA("Obs3"))%></span>
            </div>
        </div>
    </div>

</div>

<div class="ft-sheet-footer no-print d-flex align-items-center justify-content-between px-3 py-2 border-top bg-light" style="font-size:.75rem;">
    <span class="text-muted">
        <i class="bi bi-calendar3 me-1"></i>
        Emitido em <%=Right("0" & Day(Now()), 2) & "/" & Right("0" & Month(Now()), 2) & "/" & Year(Now())%>
    </span>
    <div class="d-flex gap-2">
        <button class="btn btn-sm btn-light border" onclick="location.href='RelatorioFT.asp'" title="Voltar">
            <i class="bi bi-arrow-left"></i>
        </button>
        <button class="btn btn-sm btn-danger" onclick="window.print()">
            <i class="bi bi-printer me-1"></i>Imprimir
        </button>
    </div>
</div>

</div>

<%
Else
    Response.Write "<script>alert('NÃO EXISTE DETALHES PARA ESSA FT!'); location.href='RelatorioFT.asp'</script>"
End If
CONSULTA.Close()
Set CONSULTA = Nothing
%>

</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
</body>
</html>