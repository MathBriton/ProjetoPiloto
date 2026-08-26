<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%Response.Buffer=True%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
Response.Charset     = "utf-8"
Response.CodePage    = 65001
Session.LCID         = 1046

' ══════════════════════════════════════════════════════════════════════════════
'  DISPATCHER — endpoints JSON
'  Acionado quando a querystring contém ?acao=
' ══════════════════════════════════════════════════════════════════════════════
Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.Clear
    Response.Expires     = 0
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "listar"       : resultado = Listar()
        Case "listarHist"   : resultado = ListarHistorico()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}"
    End Select

    Response.Clear
    Response.Write resultado
    Response.End()

End If

' ══════════════════════════════════════════════════════════════════════════════
'  FUNÇÃO: Listar
'  Retorna os itens analíticos por fornecedor para o período/filial informados.
' ══════════════════════════════════════════════════════════════════════════════
Function Listar()

    Dim dataIni, dataFim, filialParam, fornecedor, dbServer
    Dim filialSQL, SQL, rs
    Dim json, primeiroItem

    dataIni    = Trim(Request("DataIni"))
    dataFim    = Trim(Request("DataFim"))
    filialParam = Trim(Request("Filial"))
    fornecedor  = Trim(Request("Fornecedor"))

    ' ── Validações básicas ──────────────────────────────────────────────────
    If dataIni = "" Or dataFim = "" Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Informe a Data Inicial e a Data Final.""}"
        Exit Function
    End If

    ' ── Mapeamento Filial → cláusula SQL ────────────────────────────────────
    filialSQL = ""
    Select Case filialParam
        Case "P"  : filialSQL = "('00')"
        Case "U"  : filialSQL = "('04')"
        Case "B"  : filialSQL = "('05')"
        Case "F"  : filialSQL = "('06')"
        Case "S"  : filialSQL = "('07')"
    End Select

    ' ── Mapeamento Filial → servidor (ambas em PIR) ──────────────────────────
    dbServer = "PIR-CORPORATIVO"

    ' ── Monta SQL ───────────────────────────────────────────────────────────
    Dim dataIniBD, dataFimBD
    dataIniBD = Replace(dataIni, "-", "")
    dataFimBD = Replace(dataFim, "-", "")

    SQL = "SELECT vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE, vw_MSNotasAnalisadas_Fornecedores.CO_LOJA, "
    SQL = SQL & "SUBSTRING(Razao,1,20) + '...' A2_NREDUZ, Razao A2_NOME, "
    SQL = SQL & "SUM(ISNULL(vw_MSNotasAnalisadas_Fornecedores.PONTOS,0)) AS SOMA_PONTOS, "
    SQL = SQL & "SUM(ISNULL(vw_MSNotasAnalisadas_Fornecedores.PESO,0)) AS SOMA_ENCONTRADA, "
    SQL = SQL & "ISNULL((SELECT NOTA FROM vw_MSIQF WHERE FORNECEDOR = CO_FORNECE collate SQL_Latin1_General_CP437_BIN AND MES = Substring(CO_DATCONF,1,6) collate SQL_Latin1_General_CP437_BIN),100) IQF, "
    SQL = SQL & "SUM(isnull(vw_MSNotasAnalisadas_Fornecedores.CO_TOTAL,0)) AS TOTAL, "
    SQL = SQL & "COUNT(DISTINCT vw_MSNotasAnalisadas_Fornecedores.CO_DOC) AS NFS, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_DOC, vw_MSNotasAnalisadas_Fornecedores.CO_SERIE, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_ITEM, vw_MSNotasAnalisadas_Fornecedores.CO_COD, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_DESCRI, "
    SQL = SQL & "isnull(vw_MSNotasAnalisadas_Fornecedores.CO_TOTAL,0) as CO_TOTAL, "
    SQL = SQL & "CONVERT(SMALLDATETIME, vw_MSNotasAnalisadas_Fornecedores.CO_DTDIGIT) AS CO_DTDIGIT, "
    SQL = SQL & "CONVERT(SMALLDATETIME, vw_MSNotasAnalisadas_Fornecedores.CO_DATPRF) AS CO_DATPRF, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_PEDIDO, vw_MSNotasAnalisadas_Fornecedores.CO_DATCONF, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_HORACONF, vw_MSNotasAnalisadas_Fornecedores.CO_USERCONF, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_QUANT, vw_MSNotasAnalisadas_Fornecedores.CO_UM, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_VUNIT "
    SQL = SQL & "FROM vw_MSNotasAnalisadas_Fornecedores "
    SQL = SQL & "LEFT JOIN FornecedoresInsumos "
    SQL = SQL & "ON vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE COLLATE SQL_Latin1_General_CP437_BIN = Codigo "
    SQL = SQL & "WHERE vw_MSNotasAnalisadas_Fornecedores.CO_DATCONF BETWEEN '" & dataIniBD & "' AND '" & dataFimBD & "' "
    If filialSQL <> "" Then
        SQL = SQL & "AND vw_MSNotasAnalisadas_Fornecedores.CO_FILIAL IN " & filialSQL & " "
    End If
    SQL = SQL & "GROUP BY vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE, vw_MSNotasAnalisadas_Fornecedores.CO_LOJA, "
    SQL = SQL & "SUBSTRING(Razao,1,25), Razao, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_DOC, vw_MSNotasAnalisadas_Fornecedores.CO_SERIE, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_ITEM, vw_MSNotasAnalisadas_Fornecedores.CO_COD, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_DESCRI, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_TOTAL, "
    SQL = SQL & "CONVERT(SMALLDATETIME, vw_MSNotasAnalisadas_Fornecedores.CO_DTDIGIT), "
    SQL = SQL & "CONVERT(SMALLDATETIME, vw_MSNotasAnalisadas_Fornecedores.CO_DATPRF), "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_PEDIDO, vw_MSNotasAnalisadas_Fornecedores.CO_DATCONF, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_HORACONF, vw_MSNotasAnalisadas_Fornecedores.CO_USERCONF, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_QUANT, vw_MSNotasAnalisadas_Fornecedores.CO_UM, "
    SQL = SQL & "vw_MSNotasAnalisadas_Fornecedores.CO_VUNIT "
    If fornecedor <> "" Then
        SQL = SQL & "HAVING vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE = '" & Replace(fornecedor, "'", "''") & "' "
    End If
    SQL = SQL & "ORDER BY CO_FORNECE, CO_LOJA "

    ' ── Executa — ADODB direto (isola JSON de HTML de erro do Conecta) ────────
    Dim iqfConn, strConn
    strConn = BuscaStringConexao(UCase(dbServer))

    If Trim(strConn) = "" Then
        Listar = "{""sucesso"": false, ""mensagem"": ""String de conexão não encontrada para: " & EscapeJSON(dbServer) & """}"
        Exit Function
    End If

    On Error Resume Next
    Set iqfConn = Server.CreateObject("ADODB.Connection")
    iqfConn.CommandTimeout = 300
    iqfConn.Open strConn
    If Err.Number <> 0 Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro de conexão (" & EscapeJSON(dbServer) & "): " & EscapeJSON(Err.Description) & """}"
        Err.Clear : Set iqfConn = Nothing
        Exit Function
    End If

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open SQL, iqfConn
    If Err.Number <> 0 Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro na consulta: " & EscapeJSON(Err.Description) & """}"
        Err.Clear
        rs.Close : Set rs = Nothing
        iqfConn.Close : Set iqfConn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close()
        Set rs = Nothing
        Listar = "{""sucesso"": true, ""mensagem"": ""Nenhum registro encontrado."", ""dados"": []}"
        Exit Function
    End If

    ' ── Serializa ───────────────────────────────────────────────────────────
    primeiroItem = True
    json = "["

    Do While Not rs.EOF
        If Not primeiroItem Then json = json & ","  End If

        Dim somaPontos, somaEncontrada, coTotal, iqfVal
        If IsNull(rs("SOMA_PONTOS"))     Then somaPontos     = 0 Else somaPontos     = CDbl(rs("SOMA_PONTOS"))
        If IsNull(rs("SOMA_ENCONTRADA")) Then somaEncontrada = 0 Else somaEncontrada = CDbl(rs("SOMA_ENCONTRADA"))
        If IsNull(rs("CO_TOTAL"))        Then coTotal        = 0 Else coTotal        = CDbl(rs("CO_TOTAL"))
        If IsNull(rs("IQF"))             Then iqfVal         = 0 Else iqfVal         = CDbl(rs("IQF"))

        json = json & "{"
        json = json & """coFornece"":"   & """" & EscapeJSON(rs("CO_FORNECE") & "") & ""","
        json = json & """coLoja"":"      & """" & EscapeJSON(rs("CO_LOJA") & "")    & ""","
        json = json & """nReduz"":"      & """" & EscapeJSON(rs("A2_NREDUZ") & "")  & ""","
        json = json & """nome"":"        & """" & EscapeJSON(rs("A2_NOME") & "")    & ""","
        json = json & """coDoc"":"       & """" & EscapeJSON(rs("CO_DOC") & "")     & ""","
        json = json & """coSerie"":"     & """" & EscapeJSON(rs("CO_SERIE") & "")   & ""","
        json = json & """coItem"":"      & """" & EscapeJSON(rs("CO_ITEM") & "")    & ""","
        json = json & """coCod"":"       & """" & EscapeJSON(rs("CO_COD") & "")     & ""","
        json = json & """coDescri"":"    & """" & EscapeJSON(rs("CO_DESCRI") & "")  & ""","
        json = json & """coQuant"":"     & """" & EscapeJSON(rs("CO_QUANT") & "")   & ""","
        json = json & """coUm"":"        & """" & EscapeJSON(rs("CO_UM") & "")      & ""","
        json = json & """coDatConf"":"   & """" & EscapeJSON(rs("CO_DATCONF") & "") & ""","
        json = json & """coHoraConf"":"  & """" & EscapeJSON(rs("CO_HORACONF") & "") & ""","
        json = json & """coUserConf"":"  & """" & EscapeJSON(rs("CO_USERCONF") & "") & ""","
        json = json & """somaPontos"":"  & Replace(CStr(somaPontos),     ",", ".") & ","
        json = json & """somaEncontrada"":" & Replace(CStr(somaEncontrada), ",", ".") & ","
        json = json & """coTotal"":"     & Replace(CStr(coTotal),         ",", ".") & ","
        json = json & """iqf"":"         & Replace(CStr(iqfVal),           ",", ".")
        json = json & "}"

        primeiroItem = False
        rs.MoveNext
    Loop

    json = json & "]"

    rs.Close()
    Set rs = Nothing
    iqfConn.Close()
    Set iqfConn = Nothing

    Listar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & json & "}"

End Function

' ══════════════════════════════════════════════════════════════════════════════
'  FUNÇÃO: ListarHistorico
'  Retorna histórico IQF mensal para um único fornecedor/loja.
' ══════════════════════════════════════════════════════════════════════════════
Function ListarHistorico()

    Dim fornecedorLoja, dataIni, dataFim, filialParam, dbServer
    Dim dataIniMes, dataFimMes, sqlHist, rsHist
    Dim json, primeiroItem

    fornecedorLoja = Trim(Request("FornecedorLoja"))
    dataIni        = Trim(Request("DataIni"))
    dataFim        = Trim(Request("DataFim"))
    filialParam    = Trim(Request("Filial"))

    If fornecedorLoja = "" Or dataIni = "" Or dataFim = "" Then
        ListarHistorico = "{""sucesso"": false, ""mensagem"": ""Parâmetros obrigatórios ausentes.""}"
        Exit Function
    End If

    ' ── Servidor (ambas as filiais em PIR) ──────────────────────────────────
    dbServer = "PIR-CORPORATIVO"

    dataIniMes = Replace(Left(dataIni, 7), "-", "")
    dataFimMes = Replace(Left(dataFim, 7), "-", "")

    Dim fornSeguro
    fornSeguro = Replace(Left(fornecedorLoja, 6), "'", "''")

    sqlHist = "SELECT MES, MIN(NOTA) IQF, MAX(QTD) QTD FROM ( "
    sqlHist = sqlHist & "(SELECT DISTINCT SUBSTRING(CO_DATCONF,1,6) MES, 100 NOTA, COUNT(CO_DOC) QTD "
    sqlHist = sqlHist & " FROM vw_MSNotasAnalisadas_Fornecedores "
    sqlHist = sqlHist & " WHERE CO_FORNECE = SUBSTRING('" & fornSeguro & "',1,6) "
    sqlHist = sqlHist & " GROUP BY SUBSTRING(CO_DATCONF,1,6)) UNION "
    sqlHist = sqlHist & "(SELECT MES, NOTA, 0 QTD FROM vw_MSIQF "
    sqlHist = sqlHist & " WHERE FORNECEDOR = SUBSTRING('" & fornSeguro & "',1,6))) X "
    sqlHist = sqlHist & "WHERE MES BETWEEN SUBSTRING('" & dataIniMes & "',1,6) AND SUBSTRING('" & dataFimMes & "',1,6) "
    sqlHist = sqlHist & "GROUP BY MES ORDER BY MES "

    ' ── Executa — ADODB direto ──────────────────────────────────────────────
    Dim histConn, strConnH
    strConnH = BuscaStringConexao(UCase(dbServer))

    If Trim(strConnH) = "" Then
        ListarHistorico = "{""sucesso"": false, ""mensagem"": ""String de conexão não encontrada para: " & EscapeJSON(dbServer) & """}"
        Exit Function
    End If

    On Error Resume Next
    Set histConn = Server.CreateObject("ADODB.Connection")
    histConn.CommandTimeout = 300
    histConn.Open strConnH
    If Err.Number <> 0 Then
        ListarHistorico = "{""sucesso"": false, ""mensagem"": ""Erro de conexão (hist): " & EscapeJSON(Err.Description) & """}"
        Err.Clear : Set histConn = Nothing
        Exit Function
    End If

    Set rsHist = Server.CreateObject("ADODB.Recordset")
    rsHist.CursorLocation = 3
    rsHist.Open sqlHist, histConn
    If Err.Number <> 0 Then
        ListarHistorico = "{""sucesso"": false, ""mensagem"": ""Erro na consulta de histórico: " & EscapeJSON(Err.Description) & """}"
        Err.Clear
        rsHist.Close : Set rsHist = Nothing
        histConn.Close : Set histConn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    primeiroItem = True
    json = "["

    Do While Not rsHist.EOF
        If Not primeiroItem Then json = json & ","  End If
        json = json & "{"
        json = json & """mes"":"  & """" & EscapeJSON(rsHist("MES") & "")  & ""","
        json = json & """iqf"":"  & Replace(CStr(CDbl("0" & rsHist("IQF") & "")), ",", ".") & ","
        json = json & """qtd"":"  & Replace(CStr(CDbl("0" & rsHist("QTD") & "")), ",", ".")
        json = json & "}"
        primeiroItem = False
        rsHist.MoveNext
    Loop

    json = json & "]"

    rsHist.Close()
    Set rsHist = Nothing
    histConn.Close()
    Set histConn = Nothing

    ListarHistorico = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & json & "}"

End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>IQF - Analítico</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>
    
    <script>
        window.SWNET_CODUSER = "<%=Replace(CStr(Session("CODUSER")), Chr(34), "")%>";
        window.SWNET_NOME    = "<%=Replace(CStr(Session("NOME")),    Chr(34), "")%>";
    </script>

    <style>
        body { background-color: #f0f0f0; min-height: 100vh; }

        /* ── Card de Filtros ── */
        .filter-card {
            background-color: #fff;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
        }
        .filter-card-header {
            background-color: #fff;
            color: #212529;
            padding: .6rem 1rem;
            border-radius: 6px 6px 0 0;
            border-bottom: 1px solid #dee2e6;
            font-weight: 600;
            font-size: .9rem;
        }
        .filter-card-body { padding: .75rem 1.25rem 1rem; }
        .filter-label {
            font-size: .82rem;
            font-weight: 600;
            color: #6c757d;
            margin-bottom: .25rem;
        }

        /* ── Card de Tabela ── */
        .table-card {
            background-color: #fff;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            position: relative;
            min-height: 160px;
        }

        /* ── Toolbar ── */
        .table-toolbar { background-color: #fff; border-radius: 6px 6px 0 0; }
        .search-box { position: relative; max-width: 200px; }
        .search-box .form-control { padding-right: 76px; font-size: .82rem; }
        .search-box .search-hotkey {
            position: absolute; right: 8px; top: 50%;
            transform: translateY(-50%);
            pointer-events: none; user-select: none;
            display: flex; gap: 2px;
        }
        .search-box .search-hotkey kbd {
            font-size: .65rem; padding: .1rem .3rem;
            background-color: #f8f9fa; border: 1px solid #dee2e6;
            border-radius: 3px; color: #6c757d;
        }

        /* ── Placeholder ── */
        .table-placeholder {
            display: flex; flex-direction: column;
            align-items: center; justify-content: center;
            padding: 3.5rem 1rem; color: #6c757d; gap: .75rem;
        }
        .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
        .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

        /* ── Tabela ── */
        .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
        .table-wrapper table { min-width: max-content; }
        .table thead th {
            position: sticky; top: 0;
            background-color: #fff; color: #212529;
            font-size: .82rem; font-weight: 700;
            text-align: center; white-space: nowrap;
            z-index: 2; border-color: #dee2e6;
            padding: .55rem .75rem;
        }
        .table tbody td {
            font-size: .82rem; text-align: center;
            white-space: nowrap; color: #212529;
            vertical-align: middle; padding: .45rem .75rem;
        }
        .table tbody td.text-start { text-align: left !important; }

        /* ── Linhas de agrupamento ── */
        .tr-fornecedor-header th {
            background-color: #4a4a4a; color: #fff;
            font-size: .82rem; padding: .45rem .75rem; white-space: nowrap;
        }
        .tr-fornecedor-footer th {
            background-color: #d4d4d4; color: #212529;
            font-size: .82rem; padding: .45rem .75rem; white-space: nowrap;
        }
        .tr-iqf-historico th {
            background-color: #f8f9fa; color: #495057;
            font-size: .78rem; padding: .4rem .75rem; font-style: italic;
        }

        /* ── Rodapé da tabela ── */
        .table-footer {
            display: flex; align-items: center;
            justify-content: space-between;
            padding: .6rem 1rem; border-top: 1px solid #dee2e6;
            flex-wrap: wrap; gap: .5rem;
        }
        .table-footer-left  { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
        .table-footer-center { flex: 1; display: flex; justify-content: center; }
        .table-footer-right  { min-width: 140px; text-align: right; }
        .info-registros { font-size: .8rem; color: #6c757d; }
        .select-limite  { width: auto; }

        /* ── Paginação ── */
        #paginacao .page-link { color: #6c757d; border-color: #dee2e6; }
        #paginacao .page-item.active .page-link {
            background-color: #6c757d; border-color: #6c757d; color: #fff;
        }
        #paginacao .page-link:hover { background-color: #e9ecef; color: #495057; }
        #paginacao .page-item.disabled .page-link { color: #adb5bd; }

        /* ── Overlay de carregamento ── */
        .overlay-loading {
            position: absolute; inset: 0;
            min-height: 120px;
            background: rgba(255,255,255,.85);
            display: flex; flex-direction: column;
            align-items: center; justify-content: center;
            z-index: 10; gap: .75rem; border-radius: 6px;
        }
        .overlay-loading .spinner-border { width: 2rem; height: 2rem; color: #6c757d; }
        .overlay-loading span { font-size: .85rem; font-weight: 600; color: #495057; }

        /* ── Toast ── */
        .toast-container-custom {
            position: fixed; top: 1rem; right: 1rem;
            z-index: 1100; display: flex; flex-direction: column; gap: .5rem;
        }
        .toast-custom {
            min-width: 280px; max-width: 360px;
            border-radius: 10px; overflow: hidden;
            box-shadow: 0 4px 18px rgba(0,0,0,.13);
            background: #fff; border: 1px solid #dee2e6;
            animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
        }
        @keyframes toastSlideIn {
            from { opacity: 0; transform: translateX(40px); }
            to   { opacity: 1; transform: translateX(0); }
        }
        .toast-custom .toast-progress-bar {
            animation: toastProgressShrink 4.5s linear forwards;
        }
        @keyframes toastProgressShrink {
            from { width: 100%; } to { width: 0%; }
        }
    </style>
</head>
<body>
    <div class="container-fluid py-3 px-3 px-lg-4">

        <!-- Cabeçalho da página -->
        <div class="d-flex align-items-center mb-3">
            <span class="page-title" style="font-size:1.15rem;font-weight:700;color:#212529;">
                IQF Analítico — Índice de Qualificação de Fornecimento
            </span>
        </div>

        <!-- Card de Filtros -->
        <div class="filter-card mb-3">
            <div class="filter-card-header">
                <i class="bi bi-funnel me-2"></i>Filtros
            </div>
            <div class="filter-card-body">
                <div class="row g-2 align-items-end flex-wrap">

                    <div class="col-12 col-sm-6 col-md-auto" style="min-width:160px;">
                        <label for="Filial" class="form-label filter-label">Filial</label>
                        <select id="Filial" class="form-select form-select-sm">
                            <option value="">Todas as filiais</option>
                            <option value="P">Pirapetinga</option>
                            <option value="U">Uberaba</option>
                            <option value="B">Bento</option>
                            <option value="F">Fortaleza</option>
                            <option value="S">Saquarema</option>
                        </select>
                    </div>

                    <div class="col-auto">
                        <label for="DataIni" class="form-label filter-label">Data Inicial</label>
                        <input type="date" id="DataIni" class="form-control form-control-sm" style="width:150px;" />
                    </div>

                    <div class="col-auto">
                        <label for="DataFim" class="form-label filter-label">Data Final</label>
                        <input type="date" id="DataFim" class="form-control form-control-sm" style="width:150px;" />
                    </div>

                    <div class="col-12 col-sm-6 col-md-auto" style="min-width:130px;">
                        <label for="Fornecedor" class="form-label filter-label">Fornecedor</label>
                        <input type="text" id="Fornecedor"
                               class="form-control form-control-sm"
                               placeholder="Código" />
                    </div>

                    <div class="col-12 col-md-auto ms-md-auto">
                        <div class="d-flex gap-2 justify-content-end">
                            <button type="button" id="btnLimpar" class="btn btn-light btn-sm">
                                <i class="bi bi-eraser me-2"></i>Limpar
                            </button>
                            <button type="button" id="btnFiltrar" class="btn btn-primary btn-sm">
                                <i class="bi bi-funnel-fill me-2"></i>Filtrar
                            </button>
                        </div>
                    </div>

                </div>
            </div>
        </div>

        <!-- Card de Tabela -->
        <div class="table-card">

            <!-- Overlay de carregamento -->
            <div id="overlayLoading" class="overlay-loading d-none">
                <div class="spinner-border" role="status">
                    <span class="visually-hidden">Carregando...</span>
                </div>
                <span>Carregando...</span>
            </div>

            <!-- Toolbar -->
            <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
                <div class="d-flex gap-2">
                    <button id="btnExportCsv" type="button" class="btn btn-success btn-sm">
                        <i class="bi bi-file-earmark-excel me-2"></i>Exportar Excel
                    </button>
                    <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm">
                        <i class="bi bi-file-earmark-pdf me-2"></i>Exportar PDF
                    </button>
                </div>
                <div class="search-box">
                    <input type="text" id="searchBox" class="form-control form-control-sm"
                           placeholder="Buscar..." autocomplete="off" />
                    <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
                </div>
            </div>

            <!-- Conteúdo da tabela -->
            <div id="tabelaWrapper">
                <!-- Placeholder inicial -->
                <div id="placeholderInicial" class="table-placeholder">
                    <i class="bi bi-clipboard-data"></i>
                    <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
                </div>
            </div>

        </div><!-- /.table-card -->

    </div><!-- /.container-fluid -->

    <!-- Toast container -->
    <div class="toast-container-custom" id="toastContainer"></div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <script>
    "use strict";

    const IqfAnalitico = (function () {

        /* ── Constantes ── */
        const REPOSITORIO  = 'IqfAnalitico.asp';
        const LOGO_PATH    = 'Logo/logo-smurfit1.png';
        const TOAST_DELAY  = 4500;

        const TOAST_ICON = Object.freeze({
            sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
            erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
            aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
            info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' },
        });

        /* ── Estado ── */
        let _dadosCompletos   = [];   // todos os itens brutos da API
        let _todasLinhas      = [];   // <tr> de dados (excl. agrupamento)
        let _linhasFiltradas  = [];
        let _paginaAtual      = 1;
        let _totalPaginas     = 1;
        let _limite           = 50;

        /* ── Utilitários DOM ── */
        const $el = id => document.getElementById(id);

        /* ── Toast ── */
        function exibirToast(tipo, mensagem) {
            const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
            const id  = 'toast_' + Date.now();
            const html = `
                <div id="${id}" class="toast-custom">
                    <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
                        <i class="bi ${cfg.icon} fs-5" style="color:${cfg.cor};"></i>
                        <span class="fw-semibold flex-grow-1" style="font-size:.88rem;color:${cfg.cor};">${mensagem}</span>
                        <button style="background:none;border:none;cursor:pointer;font-size:1rem;color:#6c757d;"
                                onclick="this.closest('.toast-custom').remove()">&times;</button>
                    </div>
                    <div style="height:3px;background:#e9ecef;margin:0 .75rem .75rem;">
                        <div class="toast-progress-bar" style="height:100%;background:${cfg.barra};border-radius:2px;"></div>
                    </div>
                </div>`;
            $('#toastContainer').prepend(html);
            setTimeout(() => { const el = $el(id); if (el) el.remove(); }, TOAST_DELAY + 300);
        }

        /* ── Paginação ── */
        function _renderPaginacao() {
            const $ul = $('#paginacao').empty();
            if (_totalPaginas <= 1) return;

            const mkItem = (label, page, disabled, active) => {
                const $li = $('<li>').addClass('page-item');
                if (disabled) $li.addClass('disabled');
                if (active)   $li.addClass('active');
                const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
                if (!disabled && !active) {
                    $a.on('click', e => { e.preventDefault(); _irParaPagina(page); });
                }
                return $li.append($a);
            };

            $ul.append(mkItem('&laquo;', 1,                _paginaAtual === 1,            false));
            $ul.append(mkItem('&lsaquo;', _paginaAtual - 1, _paginaAtual === 1,            false));

            const inicio = Math.max(1, _paginaAtual - 2);
            const fim    = Math.min(_totalPaginas, _paginaAtual + 2);
            for (let p = inicio; p <= fim; p++) {
                $ul.append(mkItem(p, p, false, p === _paginaAtual));
            }

            $ul.append(mkItem('&rsaquo;', _paginaAtual + 1, _paginaAtual === _totalPaginas, false));
            $ul.append(mkItem('&raquo;', _totalPaginas,     _paginaAtual === _totalPaginas, false));
        }

        function _aplicarPagina() {
            const inicio = (_paginaAtual - 1) * _limite;
            const fim    = inicio + _limite;

            _todasLinhas.forEach((tr, i) => {
                tr.style.display = (i >= inicio && i < fim) ? '' : 'none';
            });

            _renderPaginacao();
            $('#infoRegistros').text(`Total: ${_linhasFiltradas.length} registro(s)`);
        }

        function _irParaPagina(p) {
            _paginaAtual = Math.max(1, Math.min(_totalPaginas, p));
            _aplicarPagina();
        }

        /* ── Filtro dinâmico ── */
        function _filtrar(termo) {
            const t = termo.toLowerCase().trim();
            const tbody = document.querySelector('#tabelaPrincipal tbody');
            if (!tbody) return;

            const todasTrs = Array.from(tbody.querySelectorAll('tr'));
            _linhasFiltradas = [];

            todasTrs.forEach(tr => {
                const isAgrupamento = tr.classList.contains('tr-fornecedor-header')
                                   || tr.classList.contains('tr-fornecedor-footer')
                                   || tr.classList.contains('tr-iqf-historico');
                if (isAgrupamento) { tr.style.display = ''; return; }

                const visivel = t === '' || tr.textContent.toLowerCase().includes(t);
                tr.style.display = visivel ? '' : 'none';
                if (visivel) _linhasFiltradas.push(tr);
            });

            _todasLinhas  = _linhasFiltradas.slice();
            _totalPaginas = Math.max(1, Math.ceil(_todasLinhas.length / _limite));
            _paginaAtual  = 1;
            _aplicarPagina();
        }

        function _initPaginacao() {
            const tbody = document.querySelector('#tabelaPrincipal tbody');
            if (!tbody) return;

            _todasLinhas = Array.from(tbody.querySelectorAll('tr')).filter(tr =>
                !tr.classList.contains('tr-fornecedor-header') &&
                !tr.classList.contains('tr-fornecedor-footer') &&
                !tr.classList.contains('tr-iqf-historico')
            );
            _linhasFiltradas = _todasLinhas.slice();
            _totalPaginas    = Math.max(1, Math.ceil(_todasLinhas.length / _limite));
            _paginaAtual     = 1;
            _aplicarPagina();
        }

        /* ── Formatação de data DD/MM/YYYY HH:MM ── */
        function _fmtData(yyyymmdd) {
            if (!yyyymmdd || yyyymmdd.length < 8) return yyyymmdd || '';
            return yyyymmdd.slice(6, 8) + '/' + yyyymmdd.slice(4, 6) + '/' + yyyymmdd.slice(0, 4);
        }

        /* ── Render tabela a partir dos dados JSON ── */
        function _renderTabela(dados) {
            if (!dados || dados.length === 0) {
                $('#tabelaWrapper').html(
                    '<div class="table-placeholder">' +
                    '<i class="bi bi-clipboard-data"></i>' +
                    '<p>Nenhum registro encontrado para o período selecionado.</p>' +
                    '</div>'
                );
                return;
            }

            /* Agrupa por fornecedor+loja */
            const grupos = [];
            let grupoAtual = null;

            dados.forEach(item => {
                const chave = item.coFornece.trim() + item.coLoja.trim();
                if (!grupoAtual || grupoAtual.chave !== chave) {
                    grupoAtual = { chave, itens: [], meta: item };
                    grupos.push(grupoAtual);
                }
                grupoAtual.itens.push(item);
            });

            /* Monta HTML */
            const frag = document.createDocumentFragment();

            /* wrapper + tabela */
            const wrapper = document.createElement('div');
            wrapper.className = 'table-wrapper';

            const table = document.createElement('table');
            table.id        = 'tabelaPrincipal';
            table.className = 'table table-bordered table-striped table-hover table-sm mb-0';

            /* thead */
            table.innerHTML = `
                <thead>
                    <tr>
                        <th>NF</th><th>Série</th><th>Item</th><th>Produto</th>
                        <th>Descrição</th><th>Quant.</th><th>UM</th>
                        <th>Conferência</th><th>Conferente</th>
                        <th>Pontos</th><th>Peso</th>
                    </tr>
                </thead>`;

            const tbody = document.createElement('tbody');

            grupos.forEach(grupo => {
                /* Cabeçalho do fornecedor */
                const trH = document.createElement('tr');
                trH.className = 'tr-fornecedor-header';
                trH.innerHTML = `<th colspan="11">${grupo.meta.coFornece.trim()}-${grupo.meta.coLoja.trim()} &mdash; ${grupo.meta.nReduz} &ndash; ${grupo.meta.nome}</th>`;
                tbody.appendChild(trH);

                let totalValor = 0;
                let totalPontos = 0;
                let iqfValor  = grupo.meta.iqf;

                /* Linhas de itens */
                grupo.itens.forEach(item => {
                    totalValor  += parseFloat(item.coTotal)  || 0;
                    totalPontos += parseFloat(item.somaPontos) || 0;
                    iqfValor     = item.iqf;

                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td class="text-center">${item.coDoc}</td>
                        <td class="text-center">${item.coSerie}</td>
                        <td class="text-center">${item.coItem}</td>
                        <td class="text-center">${item.coCod}</td>
                        <td class="text-start">${item.coDescri}</td>
                        <td class="text-center">${item.coQuant}</td>
                        <td class="text-center">${item.coUm}</td>
                        <td class="text-center">${_fmtData(item.coDatConf)}</td>
                        <td class="text-center">${item.coUserConf}</td>
                        <td class="text-center">${item.somaPontos}</td>
                        <td class="text-center">${item.somaEncontrada}</td>`;
                    tbody.appendChild(tr);
                });

                /* Rodapé do fornecedor */
                const iqfCellId = 'iqf_' + grupo.chave;
                const trF = document.createElement('tr');
                trF.className = 'tr-fornecedor-footer';
                trF.innerHTML = `
                    <th colspan="3">Notas Fiscais Conferidas: ${grupo.itens.length}</th>
                    <th colspan="2" class="text-end">Valor Conferido:</th>
                    <th class="text-end">${totalValor.toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2})}</th>
                    <th colspan="2" class="text-end">Pontuação não conforme:</th>
                    <th class="text-end">${totalPontos.toLocaleString('pt-BR', {maximumFractionDigits:0})}</th>
                    <th class="text-end">IQF Médio:</th>
                    <th class="text-end" id="${iqfCellId}">${Math.round(parseFloat(iqfValor))}</th>`;
                tbody.appendChild(trF);

                /* Linha histórico — placeholder enquanto carrega */
                const trHist = document.createElement('tr');
                trHist.className = 'tr-iqf-historico';
                trHist.dataset.chave = grupo.chave;
                trHist.innerHTML = `<th colspan="11" class="text-end" id="hist_${grupo.chave}">Histórico IQF Mensal: carregando...</th>`;
                tbody.appendChild(trHist);
            });

            table.appendChild(tbody);
            wrapper.appendChild(table);

            /* Rodapé paginação */
            const footer = document.createElement('div');
            footer.className = 'table-footer';
            footer.innerHTML = `
                <div class="table-footer-left">
                    <label for="selectLimite" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">Linhas por página:</label>
                    <select id="selectLimite" class="form-select form-select-sm select-limite">
                        <option value="10">10</option>
                        <option value="25">25</option>
                        <option value="50" selected>50</option>
                        <option value="100">100</option>
                        <option value="250">250</option>
                        <option value="500">500</option>
                    </select>
                </div>
                <div class="table-footer-center">
                    <nav aria-label="Navegação de páginas">
                        <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
                    </nav>
                </div>
                <div class="table-footer-right">
                    <span id="infoRegistros" class="info-registros"></span>
                </div>`;

            frag.appendChild(wrapper);
            frag.appendChild(footer);

            $('#tabelaWrapper').empty().append(frag);

            /* Paginação */
            _initPaginacao();

            /* Page-size */
            $('#selectLimite').on('change', function () {
                _limite       = parseInt(this.value, 10);
                _paginaAtual  = 1;
                _totalPaginas = Math.max(1, Math.ceil(_todasLinhas.length / _limite));
                _aplicarPagina();
            });

            /* Carrega históricos em paralelo */
            const filial  = $('#Filial').val();
            const dataIni = $('#DataIni').val();
            const dataFim = $('#DataFim').val();

            grupos.forEach(grupo => {
                $.ajax({
                    url      : REPOSITORIO,
                    type     : 'GET',
                    dataType : 'json',
                    data     : {
                        acao           : 'listarHist',
                        FornecedorLoja : grupo.chave,
                        DataIni        : dataIni,
                        DataFim        : dataFim,
                        Filial         : filial,
                    },
                    success  : function (resp) {
                        const el = document.getElementById('hist_' + grupo.chave);
                        if (!el) return;
                        if (!resp.sucesso || !resp.dados || resp.dados.length === 0) {
                            el.textContent = 'Histórico IQF Mensal: sem dados';
                            return;
                        }

                        let texto = 'Histórico IQF Mensal: ';
                        let numTot = 0, somaIqf = 0;

                        resp.dados.forEach(h => {
                            const mes = h.mes.slice(4, 6) + '/' + h.mes.slice(0, 4);
                            texto += mes + ': ' + Math.round(parseFloat(h.iqf)) + '  ';
                            numTot  += parseFloat(h.qtd)  || 0;
                            somaIqf += (parseFloat(h.iqf) || 0) * (parseFloat(h.qtd) || 0);
                        });

                        el.textContent = texto.trim();

                        /* Atualiza IQF médio ponderado no rodapé */
                        if (numTot > 0) {
                            const iqfCel = document.getElementById('iqf_' + grupo.chave);
                            if (iqfCel) iqfCel.textContent = Math.round(somaIqf / numTot);
                        }
                    },
                    error    : function () {
                        const el = document.getElementById('hist_' + grupo.chave);
                        if (el) el.textContent = 'Histórico IQF Mensal: erro ao carregar';
                    },
                });
            });
        }

        /* ── Exportação CSV ── */
        function _exportarCSV(nomeArquivo, cabecalhos, linhas) {
            const esc = v => {
                const s = v == null ? '' : String(v);
                return (s.includes(';') || s.includes('"') || s.includes('\n'))
                    ? `"${s.replace(/"/g, '""')}"` : s;
            };
            const rows = [
                cabecalhos.map(esc).join(';'),
                ...linhas.map(row => row.map(esc).join(';'))
            ];
            const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href     = url;
            a.download = nomeArquivo.endsWith('.csv') ? nomeArquivo : nomeArquivo + '.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        function _coletarDadosTabela() {
            const cabecalhos = [];
            document.querySelectorAll('#tabelaPrincipal thead th').forEach(th => {
                cabecalhos.push(th.textContent.trim());
            });
            const linhas = [];
            document.querySelectorAll('#tabelaPrincipal tbody tr').forEach(tr => {
                if (tr.style.display === 'none') return;
                if (tr.classList.contains('tr-fornecedor-header') ||
                    tr.classList.contains('tr-fornecedor-footer') ||
                    tr.classList.contains('tr-iqf-historico')) return;
                const celulas = [];
                tr.querySelectorAll('td').forEach(td => celulas.push(td.textContent.trim()));
                if (celulas.length > 0) linhas.push(celulas);
            });
            return { cabecalhos, linhas };
        }

        function _onExportCsv() {
            const { cabecalhos, linhas } = _coletarDadosTabela();
            if (linhas.length === 0) { exibirToast('aviso', 'Nenhum dado disponível para exportação.'); return; }
            _exportarCSV('IQF_Analitico.csv', cabecalhos, linhas);
        }

        /* ── Coleta dados estruturados para PDF (com linhas de agrupamento) ── */
        function _coletarDadosPdf() {
            if (!_dadosCompletos || _dadosCompletos.length === 0) return { cabecalhos: [], linhas: [], estilosLinhas: [] };

            const cabecalhos = ['Fornecedor', 'NF', 'Série', 'Item', 'Produto', 'Descrição', 'Quant.', 'UM', 'Conferência', 'Conferente', 'Pontos', 'Peso'];

            const linhas        = [];
            const estilosLinhas = [];

            /* Agrupa igual ao _renderTabela */
            const grupos = [];
            let grupoAtual = null;
            _dadosCompletos.forEach(item => {
                const chave = item.coFornece.trim() + item.coLoja.trim();
                if (!grupoAtual || grupoAtual.chave !== chave) {
                    grupoAtual = { chave, itens: [], meta: item };
                    grupos.push(grupoAtual);
                }
                grupoAtual.itens.push(item);
            });

            grupos.forEach(grupo => {
                /* Linha de cabeçalho do fornecedor */
                const labelForn = grupo.meta.coFornece + '-' + grupo.meta.coLoja + '  \u2014  ' + grupo.meta.nome;
                const linhForn = [labelForn, '', '', '', '', '', '', '', '', '', '', ''];
                linhForn._isFornecedor = true;
                linhas.push(linhForn);
                estilosLinhas.push({ fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'left' });

                /* Linhas de itens */
                grupo.itens.forEach(item => {
                    linhas.push([
                        grupo.meta.coFornece.trim() + '-' + grupo.meta.coLoja.trim() + ' \u2014 ' + grupo.meta.nome,
                        item.coDoc,
                        item.coSerie,
                        item.coItem,
                        item.coCod,
                        item.coDescri,
                        item.coQuant,
                        item.coUm,
                        _fmtData(item.coDatConf),
                        item.coUserConf,
                        item.somaPontos,
                        item.somaEncontrada,
                    ]);
                    estilosLinhas.push(null);
                });
            });

            return { cabecalhos, linhas, estilosLinhas };
        }

        /* ── Exportação PDF ── */
        function _onExportPdf() {
            const { cabecalhos, linhas, estilosLinhas } = _coletarDadosPdf();
            if (linhas.length === 0) { exibirToast('aviso', 'Nenhum dado disponível para exportação.'); return; }

            const { jsPDF } = window.jspdf;
            const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
            const pageW = doc.internal.pageSize.getWidth();

            const img  = new Image();
            img.src    = LOGO_PATH;
            img.onload = function () {
                const canvas  = document.createElement('canvas');
                canvas.width  = 900;
                canvas.height = Math.round(900 * (img.height / img.width));
                canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
                const logoData = canvas.toDataURL('image/png');
                doc.addImage(logoData, 'PNG', 14, 5, 38, 13);

                doc.setFontSize(14);
                doc.text('IQF Analítico — Índice de Qualificação de Fornecimento', pageW / 2, 24, { align: 'center' });
                doc.setFontSize(8);
                doc.text('Smurfit Westrock Brasil', pageW / 2, 30, { align: 'center' });

                const _emitidoPor = (window.SWNET_CODUSER || '') + ' - ' + (window.SWNET_NOME || '');
                const _agora      = new Date();
                const _dataHora   = _agora.toLocaleDateString('pt-BR') + ' ' +
                                    _agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
                doc.setFontSize(7);
                doc.text('Emitido por: ' + _emitidoPor, pageW - 14, 10, { align: 'right' });
                doc.text('Emitido em: ' + _dataHora,    pageW - 14, 15, { align: 'right' });

                doc.autoTable({
                    startY     : 40,
                    theme      : 'striped',
                    head       : [cabecalhos],
                    body       : linhas,
                    styles     : { fontSize: 4, cellPadding: 1.2 },
                    headStyles : { fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: 'bold' },
                    alternateRowStyles: { fillColor: [248, 248, 248] },
                    didParseCell: function (data) {
                        const raw = data.row.raw;
                        if (Array.isArray(raw) && raw._isFornecedor) {
                            data.cell.styles.fillColor = [74, 74, 74];
                            data.cell.styles.textColor = [255, 255, 255];
                            data.cell.styles.fontStyle = 'bold';
                            data.cell.styles.halign    = 'left';
                        }
                    },
                });
                doc.save('IQF_Analitico.pdf');
            };
            img.onerror = function () {
                doc.setFontSize(14);
                doc.text('IQF Analítico', pageW / 2, 14, { align: 'center' });
                const _emitidoPorErr = (window.SWNET_CODUSER || '') + ' - ' + (window.SWNET_NOME || '');
                const _agoraErr      = new Date();
                const _dataHoraErr   = _agoraErr.toLocaleDateString('pt-BR') + ' ' +
                                       _agoraErr.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
                doc.setFontSize(7);
                doc.text('Emitido por: ' + _emitidoPorErr, 14, 10);
                doc.text('Emitido em: ' + _dataHoraErr, 14, 15);
                doc.autoTable({
                    startY     : 26,
                    theme      : 'striped',
                    head       : [cabecalhos],
                    body       : linhas,
                    styles     : { fontSize: 4, cellPadding: 1.2 },
                    headStyles : { fillColor: [74, 74, 74], textColor: [255, 255, 255] },
                    alternateRowStyles: { fillColor: [248, 248, 248] },
                    didParseCell: function (data) {
                        const raw = data.row.raw;
                        if (Array.isArray(raw) && raw._isFornecedor) {
                            data.cell.styles.fillColor = [74, 74, 74];
                            data.cell.styles.textColor = [255, 255, 255];
                            data.cell.styles.fontStyle = 'bold';
                            data.cell.styles.halign    = 'left';
                        }
                    },
                });
                doc.save('IQF_Analitico.pdf');
            };
        }

        /* ── Filtrar (chama endpoint JSON) ── */
        function _onFiltrar() {
            const filial  = $('#Filial').val();
            const dataIni = $('#DataIni').val();
            const dataFim = $('#DataFim').val();
            const fornec  = $('#Fornecedor').val();

            if (!dataIni || !dataFim) {
                exibirToast('aviso', 'Informe a Data Inicial e a Data Final.');
                return;
            }
            if (dataIni > dataFim) {
                exibirToast('aviso', 'A Data Inicial não pode ser maior que a Data Final.');
                return;
            }

            $el('overlayLoading').classList.remove('d-none');
            $el('tabelaWrapper').innerHTML = '';

            $.ajax({
                url      : REPOSITORIO,
                type     : 'GET',
                dataType : 'json',
                data     : { acao: 'listar', Filial: filial, DataIni: dataIni, DataFim: dataFim, Fornecedor: fornec },
                success  : function (resp) {
                    $el('overlayLoading').classList.add('d-none');
                    if (!resp.sucesso) {
                        exibirToast('erro', resp.mensagem || 'Erro desconhecido.');
                        $('#tabelaWrapper').html(
                            '<div class="table-placeholder">' +
                            '<i class="bi bi-exclamation-triangle text-danger"></i>' +
                            '<p class="text-danger">' + (resp.mensagem || 'Erro') + '</p>' +
                            '</div>'
                        );
                        return;
                    }
                    _dadosCompletos = resp.dados || [];
                    _renderTabela(_dadosCompletos);
                    if (_dadosCompletos.length > 0) {
                        exibirToast('sucesso', _dadosCompletos.length + ' registro(s) carregado(s).');
                    }
                },
                error    : function (xhr) {
                    $el('overlayLoading').classList.add('d-none');
                    exibirToast('erro', 'Falha na requisição. Verifique a conexão.');
                    $('#tabelaWrapper').html(
                        '<div class="table-placeholder">' +
                        '<i class="bi bi-exclamation-triangle text-danger"></i>' +
                        '<p class="text-danger">Falha ao comunicar com o servidor.</p>' +
                        '</div>'
                    );
                },
            });
        }

        function _onLimpar() {
            $('#Filial').val('');
            $('#DataIni').val('');
            $('#DataFim').val('');
            $('#Fornecedor').val('');
        }

        /* ── Init ── */
        function init() {
            $('#btnFiltrar').on('click', _onFiltrar);
            $('#btnLimpar').on('click', _onLimpar);
            $('#btnExportCsv').on('click', _onExportCsv);
            $('#btnExportPdf').on('click', _onExportPdf);

            let _debounce;
            $('#searchBox').on('input', function () {
                clearTimeout(_debounce);
                _debounce = setTimeout(() => _filtrar(this.value), 250);
            });

            $(document).on('keydown', e => {
                if (e.ctrlKey && e.key === 'k') {
                    e.preventDefault();
                    $('#searchBox').trigger('focus');
                }
            });
        }

        return Object.freeze({ init });
    })();

    $(function () { IqfAnalitico.init(); });
    </script>
</body>
</html>
