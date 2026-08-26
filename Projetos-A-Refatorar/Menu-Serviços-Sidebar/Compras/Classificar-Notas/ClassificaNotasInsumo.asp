<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = True %>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID      = 1046

' ---------------------------------------------------------------------------
' EscapeJSON local — funcoes.asp tem bug: deleta aspas em vez de escapar
' ---------------------------------------------------------------------------
Function EscapeJSONLocal(str)
    If IsNull(str) Then str = ""
    Dim s
    s = CStr(str)
    s = Replace(s, Chr(92), Chr(92) & Chr(92))
    s = Replace(s, Chr(34), Chr(92) & Chr(34))
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr,   "\n")
    s = Replace(s, vbLf,   "\n")
    s = Replace(s, vbTab,  "\t")
    EscapeJSONLocal = s
End Function

Function JsonStr(v)
    JsonStr = Chr(34) & EscapeJSONLocal(v & "") & Chr(34)
End Function

Function JsonNum(v)
    If IsNull(v) Or v = "" Then
        JsonNum = "0"
    Else
        JsonNum = Replace(CStr(CDbl(v)), ",", ".")
    End If
End Function

' ---------------------------------------------------------------------------
' ConectaLocalRS — abre Recordset sem Response.Write em caso de erro
' Retorna Nothing em falha; popula sErroCNN com a mensagem
' ---------------------------------------------------------------------------
Dim sErroCNN
sErroCNN = ""

Function ConectaLocalRS(nomeCnx, sql)
    Const adOpenStatic   = 3
    Const adUseClient    = 3
    Const adLockReadOnly = 1

    sErroCNN = ""
    Set ConectaLocalRS = Nothing

    Dim strCnx
    strCnx = BuscaStringConexao(UCase(nomeCnx))
    If Trim(strCnx) = "" Then
        sErroCNN = "String de conexao nao encontrada: " & nomeCnx
        Exit Function
    End If

    On Error Resume Next

    Dim oCnn
    Set oCnn = Server.CreateObject("ADODB.Connection")
    oCnn.CommandTimeout = 300
    oCnn.Open strCnx

    If Err.Number <> 0 Then
        sErroCNN = "Falha ao conectar (" & nomeCnx & "): " & Err.Description
        Err.Clear : Set oCnn = Nothing : Exit Function
    End If

    Dim oRs
    Set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.CursorLocation = adUseClient
    oRs.Open sql, oCnn, adOpenStatic, adLockReadOnly

    If Err.Number <> 0 Then
        sErroCNN = "Erro na query (" & nomeCnx & "): " & Err.Description
        Err.Clear
        On Error Resume Next
        oRs.Close : oCnn.Close
        Set oRs = Nothing : Set oCnn = Nothing
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0
    Set oRs.ActiveConnection = Nothing
    oCnn.Close
    Set oCnn = Nothing
    Set ConectaLocalRS = oRs
End Function

' ---------------------------------------------------------------------------
' ConectaLocalExec — executa DML sem retornar Recordset (seguro para transacoes)
' ---------------------------------------------------------------------------
Function ConectaLocalExec(nomeCnx, sql)
    sErroCNN = ""
    ConectaLocalExec = False

    Dim strCnx
    strCnx = BuscaStringConexao(UCase(nomeCnx))
    If Trim(strCnx) = "" Then
        sErroCNN = "String de conexao nao encontrada: " & nomeCnx
        Exit Function
    End If

    On Error Resume Next

    Dim oCnn
    Set oCnn = Server.CreateObject("ADODB.Connection")
    oCnn.CommandTimeout = 300
    oCnn.Open strCnx

    If Err.Number <> 0 Then
        sErroCNN = "Falha ao conectar (" & nomeCnx & "): " & Err.Description
        Err.Clear : Set oCnn = Nothing : Exit Function
    End If

    oCnn.Execute sql

    If Err.Number <> 0 Then
        sErroCNN = "Erro ao executar DML (" & nomeCnx & "): " & Err.Description
        Err.Clear : oCnn.Close : Set oCnn = Nothing : Exit Function
    End If

    On Error GoTo 0
    oCnn.Close
    Set oCnn = Nothing
    ConectaLocalExec = True
End Function

' ---------------------------------------------------------------------------
' Dispatcher — responde como JSON se acao informada
' ---------------------------------------------------------------------------
Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then
    Response.Clear()
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "obterNota"     : resultado = ObterNota()
        Case "obterAnalises" : resultado = ObterAnalises()
        Case "classificar"   : resultado = Classificar()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Acao invalida: " & EscapeJSONLocal(acao) & """}"
    End Select

    Response.Clear()
    Response.Write resultado
    Response.End()
End If

' ---------------------------------------------------------------------------
' FUNÇÃO: ObterNota
' ---------------------------------------------------------------------------
Function ObterNota()
    Dim idNota, SQL, rs, json

    idNota = Trim(Request("id") & "")

    If idNota = "" Or Not IsNumeric(idNota) Then
        ObterNota = "{""sucesso"": false, ""mensagem"": ""Parametro id invalido ou nao informado""}"
        Exit Function
    End If

    SQL = "SELECT IdNota, Codigo, Razao, NotaFiscal, Serie, Item, " & _
          "       DataLancamento, Material, Descricao, Quantidade, UM, " & _
          "       Filial, PONTOS, PESO, ENCONTRADO " & _
          "FROM VwNotasAnalises " & _
          "WHERE IdNota = " & CLng(idNota) & " AND PONTOS IS NULL"

    Set rs = ConectaLocalRS("PIR-CORPORATIVO", SQL)

    If IsNull(rs) Or Not IsObject(rs) Then
        ObterNota = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar nota: " & EscapeJSONLocal(sErroCNN) & """}"
        Exit Function
    End If

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        ObterNota = "{""sucesso"": false, ""mensagem"": ""Nota nao encontrada ou ja classificada""}"
        Exit Function
    End If

    Dim dataLanc
    dataLanc = ""
    If Not IsNull(rs("DataLancamento")) And rs("DataLancamento") <> "" Then
        On Error Resume Next
        Dim dtVal
        dtVal = CDate(rs("DataLancamento"))
        If Err.Number = 0 Then
            dataLanc = Right("0" & Day(dtVal),   2) & "/" & _
                       Right("0" & Month(dtVal), 2) & "/" & _
                       Year(dtVal)
        Else
            ' Fallback: campo chega como string yyyymmdd (formato Protheus)
            Err.Clear
            Dim sRaw
            sRaw = CStr(rs("DataLancamento"))
            If Len(sRaw) = 8 And IsNumeric(sRaw) Then
                dataLanc = Mid(sRaw,7,2) & "/" & Mid(sRaw,5,2) & "/" & Left(sRaw,4)
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    json = "{"
    json = json & """sucesso"": true,"
    json = json & """dados"": {"
    json = json & """idNota"": "          & CLng(rs("IdNota"))          & ","
    json = json & """codigo"": "          & JsonStr(rs("Codigo"))        & ","
    json = json & """razao"": "           & JsonStr(rs("Razao"))         & ","
    json = json & """notaFiscal"": "      & JsonStr(rs("NotaFiscal"))    & ","
    json = json & """serie"": "           & JsonStr(rs("Serie"))         & ","
    json = json & """item"": "            & JsonStr(rs("Item"))          & ","
    json = json & """dataLancamento"": "  & JsonStr(dataLanc)            & ","
    json = json & """material"": "        & JsonStr(rs("Material"))      & ","
    json = json & """descricao"": "       & JsonStr(rs("Descricao"))     & ","
    json = json & """quantidade"": "      & JsonNum(rs("Quantidade"))    & ","
    json = json & """um"": "              & JsonStr(rs("UM"))            & ","
    json = json & """filial"": "          & JsonStr(rs("Filial"))
    json = json & "}}"

    rs.Close() : Set rs = Nothing
    ObterNota = json
End Function

' ---------------------------------------------------------------------------
' FUNÇÃO: ObterAnalises
' ---------------------------------------------------------------------------
Function ObterAnalises()
    Dim SQL, rs, json, primeiroItem

    SQL = "SELECT AN_COD, AN_TIPO, AN_DESCRI, AN_PESO " & _
          "FROM dbo.MSAnalisesMaterial " & _
          "WHERE AN_TIPO < 2 " & _
          "ORDER BY AN_TIPO, AN_COD"

    Set rs = ConectaLocalRS("PIR-CORPORATIVO", SQL)

    If IsNull(rs) Or Not IsObject(rs) Then
        ObterAnalises = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar analises: " & EscapeJSONLocal(sErroCNN) & """}"
        Exit Function
    End If

    primeiroItem = True
    json = "{""sucesso"": true, ""dados"": ["

    Do While Not rs.EOF
        If Not primeiroItem Then json = json & ","
        json = json & "{"
        json = json & """cod"": "    & JsonStr(rs("AN_COD"))   & ","
        json = json & """tipo"": "   & JsonNum(rs("AN_TIPO"))  & ","
        json = json & """descri"": " & JsonStr(rs("AN_DESCRI")) & ","
        json = json & """peso"": "   & JsonNum(rs("AN_PESO"))
        json = json & "}"
        primeiroItem = False
        rs.MoveNext()
    Loop

    json = json & "]}"
    rs.Close() : Set rs = Nothing
    ObterAnalises = json
End Function

' ---------------------------------------------------------------------------
' FUNÇÃO: Classificar
' Usa ConectaLocalExec para o DML — evita o Response.Write do cnn2.asp
' ---------------------------------------------------------------------------
Function Classificar()
    Dim codigo, nota, serie, item, material, descricao, quantidade, um
    Dim documentaisJSON, visuaisJSON

    codigo         = Trim(Request("codigo")      & "")
    nota           = Trim(Request("nota")        & "")
    serie          = Trim(Request("serie")       & "")
    item           = Trim(Request("item")        & "")
    material       = Trim(Request("material")    & "")
    descricao      = Trim(Request("descricao")   & "")
    quantidade     = Trim(Request("quantidade")  & "")
    um             = Trim(Request("um")          & "")
    documentaisJSON = Trim(Request("documentais") & "")
    visuaisJSON     = Trim(Request("visuais")     & "")

    If codigo = "" Or nota = "" Or item = "" Then
        Classificar = "{""sucesso"": false, ""mensagem"": ""Parametros obrigatorios nao informados (codigo, nota, item)""}"
        Exit Function
    End If

    If documentaisJSON = "" And visuaisJSON = "" Then
        Classificar = "{""sucesso"": false, ""mensagem"": ""Selecione pelo menos uma opcao de analise documental e uma de analise visual""}"
        Exit Function
    End If

    ' Busca pesos das analises
    Dim rsAn, SQL
    SQL = "SELECT AN_COD, AN_TIPO, AN_PESO FROM dbo.MSAnalisesMaterial WHERE AN_TIPO < 2"
    Set rsAn = ConectaLocalRS("PIR-CORPORATIVO", SQL)

    If IsNull(rsAn) Or Not IsObject(rsAn) Then
        Classificar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar analises: " & EscapeJSONLocal(sErroCNN) & """}"
        Exit Function
    End If

    Dim sqlInserts, pesoDocumental, pesoVisual
    sqlInserts     = ""
    pesoDocumental = 0
    pesoVisual     = 0

    Dim dataHoje, horaAgora
    dataHoje  = InverteData(Date())
    horaAgora = Mid(CStr(Time()), 1, 8)

    Do While Not rsAn.EOF
        Dim anCod, anTipo, anPeso
        anCod  = Trim(rsAn("AN_COD")  & "")
        anTipo = CStr(rsAn("AN_TIPO") & "")
        anPeso = 0
        On Error Resume Next
        anPeso = CInt(rsAn("AN_PESO"))
        On Error GoTo 0

        Dim marcadoDoc, marcadoVis
        marcadoDoc = (InStr(documentaisJSON, Chr(34) & anCod & Chr(34)) > 0)
        marcadoVis = (InStr(visuaisJSON,     Chr(34) & anCod & Chr(34)) > 0)

        If anTipo = "0" And marcadoDoc Then
            sqlInserts = sqlInserts & _
                "INSERT INTO MSNotasAnalises(NA_FORNECE,NA_LOJA,NA_DOC,NA_SERIE,NA_ITEM,NA_AN_COD,D_E_L_E_T_,R_E_C_D_E_L_,R_E_C_N_O_) " & _
                "VALUES('" & Replace(codigo, "'", "''") & "','','" & _
                Replace(nota,  "'", "''") & "','" & _
                Replace(serie, "'", "''") & "','" & _
                Replace(item,  "'", "''") & "','" & _
                Replace(anCod, "'", "''") & "','',0,(SELECT ISNULL(MAX(R_E_C_N_O_),0) FROM MSNotasAnalises)+1) "
            pesoDocumental = pesoDocumental + anPeso
        End If

        If anTipo = "1" And marcadoVis Then
            sqlInserts = sqlInserts & _
                "INSERT INTO MSNotasAnalises(NA_FORNECE,NA_LOJA,NA_DOC,NA_SERIE,NA_ITEM,NA_AN_COD,D_E_L_E_T_,R_E_C_D_E_L_,R_E_C_N_O_) " & _
                "VALUES('" & Replace(codigo, "'", "''") & "','','" & _
                Replace(nota,  "'", "''") & "','" & _
                Replace(serie, "'", "''") & "','" & _
                Replace(item,  "'", "''") & "','" & _
                Replace(anCod, "'", "''") & "','',0,(SELECT ISNULL(MAX(R_E_C_N_O_),0) FROM MSNotasAnalises)+1) "
            pesoVisual = pesoVisual + anPeso
        End If

        rsAn.MoveNext()
    Loop
    rsAn.Close() : Set rsAn = Nothing

    Dim quantNum
    quantNum = Replace(Replace(quantidade, ".", ""), ",", ".")
    If Not IsNumeric(quantNum) Then quantNum = "0"

    Dim sqlTrans
    sqlTrans = "BEGIN TRANSACTION " & _
               sqlInserts & _
               "INSERT INTO MSNotasConferidas(" & _
                 "CO_FORNECE,CO_DOC,CO_SERIE,CO_ITEM,CO_COD,CO_DESCRI," & _
                 "CO_QUANT,CO_UM,CO_DATCONF,CO_HORACONF,CO_USERCONF," & _
                 "CO_PONTOS,CO_PEDIDO,CO_ITEMPC,CO_ESP1,CO_ESP2,CO_ESP3,CO_ESP4,R_E_C_N_O_) " & _
               "VALUES(" & _
                 "'" & Replace(codigo,    "'", "''") & "'," & _
                 "'" & Replace(nota,      "'", "''") & "'," & _
                 "'" & Replace(serie,     "'", "''") & "'," & _
                 "'" & Replace(item,      "'", "''") & "'," & _
                 "'" & Replace(material,  "'", "''") & "'," & _
                 "'" & Replace(descricao, "'", "''") & "'," & _
                 quantNum & "," & _
                 "'" & Replace(um,        "'", "''") & "'," & _
                 "'" & dataHoje  & "'," & _
                 "'" & horaAgora & "'," & _
                 "'" & Replace(Session("USR") & "", "'", "''") & "'," & _
                 (pesoDocumental + pesoVisual) & "," & _
                 "'','','','','','', (SELECT ISNULL(MAX(R_E_C_N_O_),0) FROM MSNotasConferidas)+1) " & _
               "IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

    ' USA ConectaLocalExec — nunca usa Conecta() do cnn2 para DML
    Dim okTrans
    okTrans = ConectaLocalExec("PIR-CORPORATIVO", sqlTrans)

    If Not okTrans Then
        Classificar = "{""sucesso"": false, ""mensagem"": ""Erro ao executar classificacao: " & EscapeJSONLocal(sErroCNN) & """}"
        Exit Function
    End If

    ' Confirma gravacao
    Dim sqlVer, rsVer
    sqlVer = "SELECT COUNT(*) AS EXISTE FROM MSNotasConferidas " & _
             "WHERE CO_USERCONF = '" & Replace(Session("USR") & "", "'", "''") & "' " & _
             "  AND CO_DATCONF  = '" & dataHoje  & "' " & _
             "  AND CO_HORACONF = '" & horaAgora & "'"

    Set rsVer = ConectaLocalRS("PIR-CORPORATIVO", sqlVer)

    If IsNull(rsVer) Or Not IsObject(rsVer) Then
        Classificar = "{""sucesso"": false, ""mensagem"": ""Classificacao executada mas nao foi possivel confirmar o registro""}"
        Exit Function
    End If

    Dim existeReg
    existeReg = 0
    If Not rsVer.EOF Then
        On Error Resume Next
        existeReg = CInt(rsVer("EXISTE") & "")
        On Error GoTo 0
    End If
    rsVer.Close() : Set rsVer = Nothing

    If existeReg > 0 Then
        Classificar = "{""sucesso"": true, ""mensagem"": ""Registro classificado com sucesso!""}"
    Else
        Classificar = "{""sucesso"": false, ""mensagem"": ""Houve um erro e nao foi possivel realizar a classificacao deste item""}"
    End If
End Function

' ---------------------------------------------------------------------------
' Prepara valores para injeção no HTML
' ---------------------------------------------------------------------------
Dim idNotaHtml, parametroHtml
idNotaHtml    = 0
parametroHtml = Trim(Request("parametro") & "")

If Trim(Request("Id") & "") <> "" And IsNumeric(Trim(Request("Id") & "")) Then
    idNotaHtml = CLng(Trim(Request("Id") & ""))
End If

If parametroHtml = "" Then
    parametroHtml = Trim(Request("filial") & "") & ":" & Trim(Request("classificadas") & "")
End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Classificação de Nota Fiscal</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
        /* ── Layout ── */
        body        { background-color: #f0f0f0; min-height: 100vh; }
        .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; margin: 0; }

        /* ── Filter card (dados da NF) ── */
        .filter-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
        .filter-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
        .filter-card-body   { padding: 1rem 1.25rem 1.25rem; }
        .filter-label       { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; display: block; }

        /* ── Table card (item da nota) ── */
        .table-card    { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
        .table-wrapper { overflow-x: auto; }
        .table thead th { background-color: #fff; color: #212529; font-size: .82rem; font-weight: 700; text-align: center; white-space: nowrap; border-color: #dee2e6; padding: .55rem .75rem; }
        .table tbody td { font-size: .82rem; color: #212529; vertical-align: middle; padding: .45rem .75rem; }

        /* ── Análise cards ── */
        .analise-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); height: 100%; }
        .analise-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
        .analise-card-body   { padding: 1rem 1.25rem; }
        .analise-card-body .form-check              { padding-top: .3rem; padding-bottom: .3rem; border-bottom: 1px solid #f0f0f0; }
        .analise-card-body .form-check:last-child   { border-bottom: none; }
        .analise-card-body .form-check-label        { font-size: .85rem; cursor: pointer; }

        /* ── Overlay (sobre a área de análises) ── */
        .analise-area       { position: relative; min-height: 120px; }
        .overlay-analises   { position: absolute; inset: 0; background: rgba(255,255,255,.88); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 10; gap: .75rem; border-radius: 6px; }
        .overlay-analises .spinner-border { width: 2rem; height: 2rem; color: #6c757d; }
        .overlay-analises span            { font-size: .85rem; color: #495057; font-weight: 600; }

        /* ── Toast ── */
        .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
        .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
        @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
        .toast-custom .toast-progress-bar { height: 3px; animation: toastProgressShrink 4.5s linear forwards; }
        @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }
    </style>
</head>
<body>

<!-- Toast container -->
<div class="toast-container-custom" id="toastContainer"></div>

<div class="container-fluid py-3 px-3 px-lg-4">

  <!-- Cabeçalho -->
  <div class="d-flex align-items-center gap-3 mb-3 flex-wrap">
    <button type="button" class="btn btn-light btn-sm" onclick="history.go(-1)" title="Voltar">
      <i class="bi bi-arrow-left me-2"></i>Voltar
    </button>
    <span class="page-title">Classificação de Nota Fiscal</span>
  </div>

  <!-- Área principal -->
  <div id="areaConteudo">

    <!-- Placeholder inicial -->
    <div id="placeholderCarregando" class="d-flex flex-column align-items-center justify-content-center py-5 text-muted gap-3">
      <div class="spinner-border" style="width:2.5rem;height:2.5rem;color:#6c757d;" role="status">
        <span class="visually-hidden">Carregando...</span>
      </div>
      <span style="font-size:.9rem;font-weight:600;">Carregando dados...</span>
    </div>

    <!-- Alerta de não encontrado -->
    <div id="erroCarregamento" class="alert alert-warning d-flex align-items-center gap-2 d-none" role="alert">
      <i class="bi bi-exclamation-triangle-fill me-2"></i>
      <span>
        <strong>Atenção!</strong>
        <span id="erroMsg">Item de nota fiscal não encontrado.</span>
        <a href="#" class="alert-link ms-1" onclick="history.go(-1)">Voltar</a>
      </span>
    </div>

    <!-- Formulário (oculto até carregar via AJAX) -->
    <div id="formWrapper" class="d-none">

        <!-- Dados da Nota Fiscal (igual ao filter-card do Reclassifica) -->
        <div class="filter-card mb-3">
            <div class="filter-card-header">
                <i class="bi bi-file-earmark-text me-2"></i>Dados da Nota Fiscal
            </div>
            <div class="filter-card-body">
                <div class="row g-3 align-items-end">
                    <div class="col-12 col-md-2">
                        <label class="filter-label" for="fCodigo">Código</label>
                        <input type="text" id="fCodigo" class="form-control form-control-sm" readonly />
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="filter-label" for="fRazao">Razão Social</label>
                        <input type="text" id="fRazao" class="form-control form-control-sm" readonly />
                    </div>
                    <div class="col-6 col-md-2">
                        <label class="filter-label" for="fNota">Nota Fiscal</label>
                        <input type="text" id="fNota" class="form-control form-control-sm" readonly />
                    </div>
                    <div class="col-6 col-md-1">
                        <label class="filter-label" for="fSerie">Série</label>
                        <input type="text" id="fSerie" class="form-control form-control-sm" readonly />
                    </div>
                    <div class="col-6 col-md-1">
                        <label class="filter-label" for="fItem">Item</label>
                        <input type="text" id="fItem" class="form-control form-control-sm" readonly />
                    </div>
                    <div class="col-6 col-md-1">
                        <button type="button" id="btnSalvar" class="btn btn-success btn-sm w-100">
                            Classificar
                        </button>
                    </div>
                </div>

                <!-- Tabela do item (dentro do filter-card, igual ao Reclassifica) -->
                <div class="table-card mt-3">
                    <div class="table-wrapper">
                        <table class="table table-bordered table-striped table-hover table-sm mb-0">
                            <thead>
                                <tr>
                                    <th class="text-center">Lançamento</th>
                                    <th class="text-center">Material</th>
                                    <th class="text-center" style="width:55%">Descrição</th>
                                    <th class="text-center">Quantidade</th>
                                    <th class="text-center">UM</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td id="tdDataLanc" class="text-center align-middle"></td>
                                    <td id="tdMaterial" class="text-center align-middle"></td>
                                    <td id="tdDescricao" class="text-center align-middle"></td>
                                    <td id="tdQuantidade" class="text-center align-middle"></td>
                                    <td id="tdUm" class="text-center align-middle"></td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- Análises (igual ao Reclassifica: analise-card com header separado) -->
        <div class="analise-area">
            <div id="overlayAnalises" class="overlay-analises d-none">
                <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
                <span>Classificando...</span>
            </div>

            <div class="row g-3 mb-4">
                <div class="col-12 col-lg-6">
                    <div class="analise-card">
                        <div class="analise-card-header">
                            <i class="bi bi-file-earmark-check me-2"></i>Análise Documental
                        </div>
                        <div class="analise-card-body">
                            <div class="form-check">
                                <input type="checkbox" id="chkDocOk" class="form-check-input" />
                                <label class="form-check-label text-success fw-semibold" for="chkDocOk">
                                    Nenhuma inconformidade
                                </label>
                            </div>
                            <div id="listaDocumental"></div>
                        </div>
                    </div>
                </div>
                <div class="col-12 col-lg-6">
                    <div class="analise-card">
                        <div class="analise-card-header">
                            <i class="bi bi-eye me-2"></i>Análise Visual
                        </div>
                        <div class="analise-card-body">
                            <div class="form-check">
                                <input type="checkbox" id="chkVisOk" class="form-check-input" />
                                <label class="form-check-label text-success fw-semibold" for="chkVisOk">
                                    Nenhuma inconformidade
                                </label>
                            </div>
                            <div id="listaVisual"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

    </div><!-- /formWrapper -->
  </div><!-- /areaConteudo -->
</div><!-- /container-fluid -->

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
const ClassificaNotasInsumo = (function () {
    'use strict';

    // ── Constantes ──────────────────────────────────────────
    const REPOSITORIO     = 'ClassificaNotasInsumo.asp';
    const ID_NOTA         = <%=idNotaHtml%>;
    const PARAMETRO_VOLTA = '<%=Replace(parametroHtml, "'", "\'")%>';
    const TOAST_DELAY     = 4500;

    const TOAST_ICON = {
        sucesso: { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
        erro:    { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
        aviso:   { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
        info:    { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' }
    };

    // ── Estado ──────────────────────────────────────────────
    let _nota     = {};
    let _analises = [];

    // ── Helper ──────────────────────────────────────────────
    const $el = id => document.getElementById(id);

    // ── Toast ───────────────────────────────────────────────
    function exibirToast(tipo, mensagem) {
        const t  = TOAST_ICON[tipo] || TOAST_ICON.info;
        const el = document.createElement('div');
        el.className = 'toast-custom';
        el.innerHTML =
            '<div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">' +
                '<i class="bi ' + t.icon + '" style="color:' + t.cor + ';font-size:1.1rem;"></i>' +
                '<span style="font-weight:600;font-size:.85rem;color:' + t.cor + ';">' + mensagem + '</span>' +
                '<button onclick="this.closest(\'.toast-custom\').remove()" ' +
                        'style="margin-left:auto;background:none;border:none;font-size:1rem;color:#adb5bd;cursor:pointer;">&times;</button>' +
            '</div>' +
            '<div style="height:3px;background:#e9ecef;">' +
                '<div class="toast-progress-bar" style="height:100%;background:' + t.barra + ';"></div>' +
            '</div>';
        $el('toastContainer').appendChild(el);
        setTimeout(() => el.remove(), TOAST_DELAY);
    }

    // ── Overlay ─────────────────────────────────────────────
    function showLoading(v) {
        const el = $el('overlayAnalises');
        if (!el) return;
        v ? el.classList.remove('d-none') : el.classList.add('d-none');
    }

    // ── Erro de carregamento ────────────────────────────────
    function mostrarErro(msg) {
        $el('placeholderCarregando').classList.add('d-none');
        $el('erroMsg').textContent = msg || 'Item não encontrado.';
        $el('erroCarregamento').classList.remove('d-none');
        $el('formWrapper').classList.add('d-none');
    }

    // ── Carregamento inicial em paralelo ────────────────────
    function init() {
        if (!ID_NOTA) { mostrarErro('Id da nota não informado.'); return; }

        showLoading(true);

        const reqNota = $.ajax({
            url: REPOSITORIO, method: 'GET', dataType: 'json',
            data: { acao: 'obterNota', id: ID_NOTA }
        });

        const reqAnalises = $.ajax({
            url: REPOSITORIO, method: 'GET', dataType: 'json',
            data: { acao: 'obterAnalises' }
        });

        $.when(reqNota, reqAnalises)
            .done(function (resNota, resAnalises) {
                showLoading(false);
                $el('placeholderCarregando').classList.add('d-none');
                const rn = resNota[0];
                const ra = resAnalises[0];

                if (!rn.sucesso) { mostrarErro(rn.mensagem); return; }
                if (!ra.sucesso) { mostrarErro(ra.mensagem); return; }

                _nota     = rn.dados;
                _analises = ra.dados;

                _preencherCabecalho();
                _preencherAnalises();
                $el('formWrapper').classList.remove('d-none');
            })
            .fail(function () {
                showLoading(false);
                mostrarErro('Erro de comunicação com o servidor.');
            });
    }

    // ── Preenchimento do formulário ─────────────────────────
    function _preencherCabecalho() {
        $el('fCodigo').value  = _nota.codigo      || '';
        $el('fRazao').value   = _nota.razao        || '';
        $el('fNota').value    = _nota.notaFiscal   || '';
        $el('fSerie').value   = _nota.serie        || '';
        $el('fItem').value    = _nota.item         || '';

        $el('tdDataLanc').textContent  = _nota.dataLancamento || '';
        $el('tdMaterial').textContent  = _nota.material       || '';
        $el('tdDescricao').textContent = _nota.descricao      || '';
        $el('tdUm').textContent        = _nota.um             || '';

        const qtd = parseFloat(_nota.quantidade) || 0;
        $el('tdQuantidade').textContent = qtd.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function _preencherAnalises() {
        const divDoc = $el('listaDocumental');
        const divVis = $el('listaVisual');
        divDoc.innerHTML = '';
        divVis.innerHTML = '';

        const fragDoc = document.createDocumentFragment();
        const fragVis = document.createDocumentFragment();

        _analises.forEach(a => {
            const wrap = document.createElement('div');
            wrap.className = 'form-check';

            const chk = document.createElement('input');
            chk.type          = 'checkbox';
            chk.className     = 'form-check-input analise-item';
            chk.value         = a.cod;
            chk.dataset.tipo  = a.tipo;
            chk.dataset.peso  = a.peso;

            const lbl = document.createElement('label');
            lbl.className   = 'form-check-label';
            lbl.textContent = a.descri + '.';

            wrap.appendChild(chk);
            wrap.appendChild(lbl);

            parseInt(a.tipo) === 0 ? fragDoc.appendChild(wrap) : fragVis.appendChild(wrap);
        });

        divDoc.appendChild(fragDoc);
        divVis.appendChild(fragVis);
        _bindControles();
    }

    // ── Lógica de controle dos checkboxes ──────────────────
    function _bindControles() {
        $el('chkDocOk').addEventListener('change', function () {
            document.querySelectorAll(".analise-item[data-tipo='0']").forEach(c => {
                c.disabled = this.checked;
                if (this.checked) c.checked = false;
            });
        });

        $el('chkVisOk').addEventListener('change', function () {
            document.querySelectorAll(".analise-item[data-tipo='1']").forEach(c => {
                c.disabled = this.checked;
                if (this.checked) c.checked = false;
            });
        });

        document.querySelectorAll(".analise-item[data-tipo='0']").forEach(c => {
            c.addEventListener('change', () => {
                const algum = Array.from(document.querySelectorAll(".analise-item[data-tipo='0']")).some(function(x) { return x.checked; });
                $el('chkDocOk').disabled = algum;
                if (algum) $el('chkDocOk').checked = false;
            });
        });

        document.querySelectorAll(".analise-item[data-tipo='1']").forEach(c => {
            c.addEventListener('change', () => {
                const algum = Array.from(document.querySelectorAll(".analise-item[data-tipo='1']")).some(function(x) { return x.checked; });
                $el('chkVisOk').disabled = algum;
                if (algum) $el('chkVisOk').checked = false;
            });
        });
    }

    // ── Coleta selecionados ─────────────────────────────────
    function _coletarSelecionados(tipo) {
        return Array.from(document.querySelectorAll(".analise-item[data-tipo='" + tipo + "']"))
            .filter(c => c.checked).map(c => c.value);
    }

    function _validarSelecao() {
        const docOk   = $el('chkDocOk').checked;
        const visOk   = $el('chkVisOk').checked;
        const docs    = _coletarSelecionados(0);
        const visuais = _coletarSelecionados(1);

        if (!docOk && docs.length === 0) {
            exibirToast('aviso', "Selecione pelo menos um item em Análise Documental (ou marque 'Nenhuma inconformidade').");
            return false;
        }
        if (!visOk && visuais.length === 0) {
            exibirToast('aviso', "Selecione pelo menos um item em Análise Visual (ou marque 'Nenhuma inconformidade').");
            return false;
        }
        return true;
    }

    // ── Submissão ───────────────────────────────────────────
    function _salvar() {
        if (!_validarSelecao()) return;

        const docOk   = $el('chkDocOk').checked;
        const visOk   = $el('chkVisOk').checked;
        const docs    = docOk ? [] : _coletarSelecionados(0);
        const visuais = visOk ? [] : _coletarSelecionados(1);

        showLoading(true);
        $el('btnSalvar').disabled = true;

        $.ajax({
            url:      REPOSITORIO,
            method:   'POST',
            dataType: 'json',
            data: {
                acao:        'classificar',
                codigo:      _nota.codigo,
                nota:        _nota.notaFiscal,
                serie:       _nota.serie,
                item:        _nota.item,
                material:    _nota.material,
                descricao:   _nota.descricao,
                quantidade:  _nota.quantidade,
                um:          _nota.um,
                documentais: JSON.stringify(docs),
                visuais:     JSON.stringify(visuais)
            },
            success: function (res) {
                if (res.sucesso) {
                    exibirToast('sucesso', res.mensagem);
                    if (PARAMETRO_VOLTA && PARAMETRO_VOLTA !== ':') {
                        const partes = PARAMETRO_VOLTA.split(':');
                        setTimeout(() => {
                            window.location.href = 'ListaNotasInsumo.asp?refresh=1&filial=' +
                                encodeURIComponent(partes[0] || '') +
                                '&classificadas=' + encodeURIComponent(partes[1] || '');
                        }, 1200);
                    }
                } else {
                    exibirToast('erro', res.mensagem || 'Erro ao classificar.');
                    $el('btnSalvar').disabled = false;
                }
            },
            error: function (xhr) {
                exibirToast('erro', 'Erro de comunicação com o servidor.');
                console.error(xhr.responseText);
                $el('btnSalvar').disabled = false;
            },
            complete: function () {
                showLoading(false);
            }
        });
    }

    // ── Bind eventos ────────────────────────────────────────
    function _bindEventos() {
        $el('btnSalvar').addEventListener('click', _salvar);
    }

    // ── Init público ────────────────────────────────────────
    function start() {
        _bindEventos();
        init();
    }

    return Object.freeze({ init: start });
}());

$(function () { ClassificaNotasInsumo.init(); });
</script>
</body>
</html>
