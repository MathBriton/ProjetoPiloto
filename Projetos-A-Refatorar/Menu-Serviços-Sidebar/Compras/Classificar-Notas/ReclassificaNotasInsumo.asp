<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Expires     = 0
Session.LCID         = 1046

' ============================================================
' HELPER: EscapeJSONLocal
' Escape correto, independente do EscapeJSON do funcoes.asp
' (que apaga aspas em vez de escapá-las).
' ============================================================
Function EscapeJSONLocal(val)
    Dim s, i, c, code, resultado
    If IsNull(val) Then
        EscapeJSONLocal = ""
        Exit Function
    End If
    s = CStr(val)
    s = Replace(s, Chr(92), Chr(92) & Chr(92))
    s = Replace(s, Chr(34), Chr(92) & Chr(34))
    s = Replace(s, vbCrLf, Chr(92) & "n")
    s = Replace(s, vbCr,   Chr(92) & "n")
    s = Replace(s, vbLf,   Chr(92) & "n")
    s = Replace(s, vbTab,  Chr(92) & "t")
    resultado = ""
    For i = 1 To Len(s)
        c    = Mid(s, i, 1)
        code = AscW(c)
        If code >= 0 And code <= 7 Then
            resultado = resultado & "\u00" & Right("0" & Hex(code), 2)
        ElseIf code = 8 Then
            resultado = resultado & "\b"
        ElseIf code = 11 Then
            resultado = resultado & "\u000B"
        ElseIf code = 12 Then
            resultado = resultado & "\f"
        ElseIf code >= 14 And code <= 31 Then
            resultado = resultado & "\u00" & Right("0" & Hex(code), 2)
        Else
            resultado = resultado & c
        End If
    Next
    EscapeJSONLocal = resultado
End Function

' ============================================================
' HELPERS DE CONEXÃO (cnn2-safe, sem Response.Write em erro)
' ============================================================
Dim sErroConectaLocal
sErroConectaLocal = ""

Function ConectaLocal(nomeCnx, sql)
    Const adOpenStatic   = 3
    Const adUseClient    = 3
    Const adLockReadOnly = 1

    sErroConectaLocal = ""
    Set ConectaLocal   = Nothing

    Dim strCnx
    strCnx = BuscaStringConexao(UCase(nomeCnx))
    If Trim(strCnx) = "" Then
        sErroConectaLocal = "String de conexão não encontrada: " & nomeCnx
        Exit Function
    End If

    On Error Resume Next

    Dim oCnn
    Set oCnn = Server.CreateObject("ADODB.Connection")
    oCnn.CommandTimeout = 300
    oCnn.Open strCnx

    If Err.Number <> 0 Then
        sErroConectaLocal = "Falha ao conectar (" & nomeCnx & "): " & Err.Description
        Err.Clear
        Set oCnn = Nothing
        Exit Function
    End If

    Dim oRs
    Set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.CursorLocation = adUseClient
    oRs.Open sql, oCnn, adOpenStatic, adLockReadOnly

    If Err.Number <> 0 Then
        sErroConectaLocal = "Erro na query (" & nomeCnx & "): " & Err.Description
        Err.Clear
        oRs.Close
        oCnn.Close
        Set oRs  = Nothing
        Set oCnn = Nothing
        Exit Function
    End If

    On Error GoTo 0

    Set oRs.ActiveConnection = Nothing
    oCnn.Close
    Set oCnn = Nothing
    Set ConectaLocal = oRs
End Function

Function ConectaLocalExec(nomeCnx, sql)
    sErroConectaLocal = ""
    ConectaLocalExec  = False

    Dim strCnx
    strCnx = BuscaStringConexao(UCase(nomeCnx))
    If Trim(strCnx) = "" Then
        sErroConectaLocal = "String de conexão não encontrada: " & nomeCnx
        Exit Function
    End If

    On Error Resume Next

    Dim oCnn
    Set oCnn = Server.CreateObject("ADODB.Connection")
    oCnn.CommandTimeout = 300
    oCnn.Open strCnx

    If Err.Number <> 0 Then
        sErroConectaLocal = "Falha ao conectar (" & nomeCnx & "): " & Err.Description
        Err.Clear
        Set oCnn = Nothing
        Exit Function
    End If

    oCnn.Execute sql

    If Err.Number <> 0 Then
        sErroConectaLocal = "Erro ao executar DML (" & nomeCnx & "): " & Err.Description
        Err.Clear
        oCnn.Close
        Set oCnn = Nothing
        Exit Function
    End If

    On Error GoTo 0
    oCnn.Close
    Set oCnn = Nothing
    ConectaLocalExec = True
End Function

' ============================================================
' DISPATCHER — JSON API
' ============================================================
Dim acao
acao = Trim(Request("acao") & "")

' Filial passada pela URL (ex: &filial=04 para Uberaba)
' Usado para rotear para a conexão correta de cada planta.
Dim filialParam
filialParam = Trim(Request("filial") & "")

' Gravação centralizada em PIR-CORPORATIVO para todas as filiais,
' igual ao ClassificaNotasInsumo.asp. A filial é usada apenas como
' filtro nas queries, não para rotear a conexão.
Function ConexaoPorFilial(cod)
    ConexaoPorFilial = "PIR-CORPORATIVO"
End Function

If acao <> "" Then
    Response.Clear()
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "carregarDados" : resultado = CarregarDados()
        Case "salvar"        : resultado = Salvar()
        Case "listarAnalises": resultado = ListarAnalises()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSONLocal(acao) & """}"
    End Select

    Response.Write resultado
    Response.End()
End If

' ============================================================
' FUNÇÃO: CarregarDados
' Retorna dados da nota + análises para montar o formulário
' ============================================================
Function CarregarDados()
    Dim idNota
    idNota = Trim(Request("id") & "")

    If idNota = "" Or Not IsNumeric(idNota) Then
        CarregarDados = "{""sucesso"": false, ""mensagem"": ""Identificador inválido.""}"
        Exit Function
    End If

    Dim SQL
    SQL = "SELECT IdNota, Codigo, Razao, NotaFiscal, Serie, Item, DataLancamento, " & _
          "Material, Descricao, Quantidade, UM, Filial, PONTOS, PESO, ENCONTRADO " & _
          "FROM VwNotasAnalises " & _
          "WHERE IdNota = " & CLng(idNota) & " AND CO_PONTOS IS NOT NULL"

    Dim rsNota
    Set rsNota = ConectaLocal("PIR-CORPORATIVO", SQL)

    If IsNull(rsNota) Or Not IsObject(rsNota) Then
        CarregarDados = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar nota: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    If rsNota.EOF Then
        rsNota.Close()
        Set rsNota = Nothing
        CarregarDados = "{""sucesso"": false, ""mensagem"": ""Item de nota fiscal não encontrado.""}"
        Exit Function
    End If

    Dim sCodigo, sRazao, sNota, sSerie, sItem, sDataLanc, sMaterial, sDescricao, sQuantidade, sUM
    sCodigo    = rsNota("Codigo")          & ""
    sRazao     = rsNota("Razao")           & ""
    sNota      = rsNota("NotaFiscal")      & ""
    sSerie     = rsNota("Serie")           & ""
    sItem      = rsNota("Item")            & ""
    sDataLanc  = rsNota("DataLancamento")  & ""
    sMaterial  = rsNota("Material")        & ""
    sDescricao = rsNota("Descricao")       & ""
    sQuantidade= rsNota("Quantidade")      & ""
    sUM        = rsNota("UM")              & ""
    rsNota.Close()
    Set rsNota = Nothing

    ' Formata data para exibição
    Dim sDataFmt
    sDataFmt = ""
    If sDataLanc <> "" Then sDataFmt = AcertaData(sDataLanc)

    ' Carrega análises com flag de seleção atual
    Dim sqlAnalises
    sqlAnalises = "SELECT AN_COD, AN_TIPO, AN_DESCRI, AN_PESO, " & _
                  "ISNULL(NA_AN_COD, '') AS CODIGO " & _
                  "FROM MSAnalisesMaterial " & _
                  "LEFT JOIN (" & _
                  "    SELECT NA_AN_COD FROM MSNotasAnalises " & _
                  "    WHERE NA_FORNECE = '" & Replace(sCodigo, "'", "''") & "'" & _
                  "    AND NA_DOC    = '" & Replace(sNota,   "'", "''") & "'" & _
                  "    AND NA_SERIE  = '" & Replace(sSerie,  "'", "''") & "'" & _
                  "    AND NA_ITEM   = '" & Replace(sItem,   "'", "''") & "'" & _
                  ") AS ANALISES ON NA_AN_COD = AN_COD " & _
                  "WHERE AN_TIPO < 2 " & _
                  "ORDER BY AN_TIPO, AN_COD"

    Dim rsAnalises
    Set rsAnalises = ConectaLocal("PIR-CORPORATIVO", sqlAnalises)

    If IsNull(rsAnalises) Or Not IsObject(rsAnalises) Then
        CarregarDados = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar análises: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    ' Serializa análises
    Dim analItems()
    Dim nAnal
    nAnal = 0
    ReDim analItems(0)

    Do While Not rsAnalises.EOF
        Dim analObj
        analObj = "{"
        analObj = analObj & """cod"":"     & Chr(34) & EscapeJSONLocal(rsAnalises("AN_COD")  & "") & Chr(34) & ","
        analObj = analObj & """tipo"":"    & IIf(IsNull(rsAnalises("AN_TIPO")), 0, CInt(rsAnalises("AN_TIPO")))  & ","
        analObj = analObj & """descri"":"  & Chr(34) & EscapeJSONLocal(rsAnalises("AN_DESCRI") & "") & Chr(34) & ","
        analObj = analObj & """peso"":"    & IIf(IsNull(rsAnalises("AN_PESO")), 0, CInt(rsAnalises("AN_PESO"))) & ","
        analObj = analObj & """marcado"":" & LCase(CStr(Trim(rsAnalises("CODIGO") & "") <> ""))
        analObj = analObj & "}"

        ReDim Preserve analItems(nAnal)
        analItems(nAnal) = analObj
        nAnal = nAnal + 1
        rsAnalises.MoveNext()
    Loop
    rsAnalises.Close()
    Set rsAnalises = Nothing

    Dim jsonAnalises
    If nAnal = 0 Then
        jsonAnalises = "[]"
    Else
        jsonAnalises = "[" & Join(analItems, ",") & "]"
    End If

    ' Monta objeto nota
    Dim jsonNota
    jsonNota = "{"
    jsonNota = jsonNota & """codigo"":"      & Chr(34) & EscapeJSONLocal(sCodigo)    & Chr(34) & ","
    jsonNota = jsonNota & """razao"":"       & Chr(34) & EscapeJSONLocal(sRazao)     & Chr(34) & ","
    jsonNota = jsonNota & """nota"":"        & Chr(34) & EscapeJSONLocal(sNota)      & Chr(34) & ","
    jsonNota = jsonNota & """serie"":"       & Chr(34) & EscapeJSONLocal(sSerie)     & Chr(34) & ","
    jsonNota = jsonNota & """item"":"        & Chr(34) & EscapeJSONLocal(sItem)      & Chr(34) & ","
    jsonNota = jsonNota & """dataLanc"":"    & Chr(34) & EscapeJSONLocal(sDataFmt)   & Chr(34) & ","
    jsonNota = jsonNota & """material"":"    & Chr(34) & EscapeJSONLocal(sMaterial)  & Chr(34) & ","
    jsonNota = jsonNota & """descricao"":"   & Chr(34) & EscapeJSONLocal(sDescricao) & Chr(34) & ","
    jsonNota = jsonNota & """quantidade"":"  & Chr(34) & EscapeJSONLocal(sQuantidade)& Chr(34) & ","
    jsonNota = jsonNota & """um"":"          & Chr(34) & EscapeJSONLocal(sUM)        & Chr(34)
    jsonNota = jsonNota & "}"

    CarregarDados = "{""sucesso"": true, ""mensagem"": """", ""nota"": " & jsonNota & ", ""analises"": " & jsonAnalises & "}"
End Function

' ============================================================
' FUNÇÃO: ListarAnalises
' Retorna apenas os AN_COD de cada tipo (para geração de
' funções JS de controle de checkboxes)
' ============================================================
Function ListarAnalises()
    Dim SQL
    SQL = "SELECT AN_COD, AN_TIPO FROM dbo.MSAnalisesMaterial WHERE AN_TIPO < 2 ORDER BY AN_TIPO, AN_COD"

    Dim rs
    Set rs = ConectaLocal("PIR-CORPORATIVO", SQL)

    If IsNull(rs) Or Not IsObject(rs) Then
        ListarAnalises = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar análises: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    Dim docItems(), visItems()
    Dim nDoc, nVis
    nDoc = 0 : nVis = 0
    ReDim docItems(0) : ReDim visItems(0)

    Do While Not rs.EOF
        Dim cod
        cod = Chr(34) & EscapeJSONLocal(rs("AN_COD") & "") & Chr(34)
        If CInt(rs("AN_TIPO") & "0") = 0 Then
            ReDim Preserve docItems(nDoc)
            docItems(nDoc) = cod
            nDoc = nDoc + 1
        Else
            ReDim Preserve visItems(nVis)
            visItems(nVis) = cod
            nVis = nVis + 1
        End If
        rs.MoveNext()
    Loop
    rs.Close()
    Set rs = Nothing

    Dim jsonDoc, jsonVis
    If nDoc = 0 Then jsonDoc = "[]" Else jsonDoc = "[" & Join(docItems, ",") & "]"
    If nVis = 0 Then jsonVis = "[]" Else jsonVis = "[" & Join(visItems, ",") & "]"

    ListarAnalises = "{""sucesso"": true, ""documental"": " & jsonDoc & ", ""visual"": " & jsonVis & "}"
End Function

' ============================================================
' FUNÇÃO: Salvar
' Reclassifica a nota: apaga MSNotasAnalises, reinsere com
' as análises marcadas e atualiza MSNotasConferidas.
' ============================================================
Function Salvar()
    Dim codigo, nota, serie, item, filialSalvar
    codigo       = Trim(Request("codigo") & "")
    nota         = Trim(Request("nota")   & "")
    serie        = Trim(Request("serie")  & "")
    item         = Trim(Request("item")   & "")
    ' filial vem do POST; fallback para o filialParam da querystring
    filialSalvar = Trim(Request("filial") & "")
    If filialSalvar = "" Then filialSalvar = filialParam

    If codigo = "" Or nota = "" Or serie = "" Or item = "" Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Parâmetros obrigatórios ausentes.""}"
        Exit Function
    End If

    ' Conexão de gravação: usa a base da filial da nota
    Dim cnxGravacao
    cnxGravacao = ConexaoPorFilial(filialSalvar)

    ' Carrega todas as análises disponíveis (tipo 0 e 1)
    ' MSAnalisesMaterial é centralizado no PIR-CORPORATIVO
    Dim sqlAnal
    sqlAnal = "SELECT AN_COD, AN_TIPO, AN_PESO FROM dbo.MSAnalisesMaterial WHERE AN_TIPO < 2"

    Dim rsAnal
    Set rsAnal = ConectaLocal("PIR-CORPORATIVO", sqlAnal)

    If IsNull(rsAnal) Or Not IsObject(rsAnal) Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar análises: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    Dim pesoDocumental, pesoVisual
    pesoDocumental = 0
    pesoVisual     = 0

    ' Monta INSERT das análises marcadas
    Dim insertParts()
    Dim nInserts
    nInserts = 0
    ReDim insertParts(0)

    Do While Not rsAnal.EOF
        Dim anCod, anTipo, anPeso
        anCod  = rsAnal("AN_COD") & ""
        anTipo = IIf(IsNull(rsAnal("AN_TIPO")), 0, CInt(rsAnal("AN_TIPO")))
        anPeso = IIf(IsNull(rsAnal("AN_PESO")), 0, CInt(rsAnal("AN_PESO")))

        Dim paramNome, marcado
        If anTipo = 0 Then
            paramNome = "documental" & anCod
        Else
            paramNome = "visual" & anCod
        End If
        marcado = Trim(Request(paramNome) & "")

        If marcado <> "" Then
            Dim ins
            ins = "INSERT INTO MSNotasAnalises(NA_FORNECE,NA_LOJA,NA_DOC,NA_SERIE,NA_ITEM,NA_AN_COD,D_E_L_E_T_,R_E_C_D_E_L_,R_E_C_N_O_) "
            ins = ins & "VALUES("
            ins = ins & "'" & Replace(codigo, "'", "''") & "',"
            ins = ins & "'',"
            ins = ins & "'" & Replace(nota,   "'", "''") & "',"
            ins = ins & "'" & Replace(serie,  "'", "''") & "',"
            ins = ins & "'" & Replace(item,   "'", "''") & "',"
            ins = ins & "'" & Replace(anCod,  "'", "''") & "',"
            ins = ins & "'',0,"
            ins = ins & "(SELECT MAX(R_E_C_N_O_) FROM MSNotasAnalises)+1)"

            ReDim Preserve insertParts(nInserts)
            insertParts(nInserts) = ins
            nInserts = nInserts + 1

            If anTipo = 0 Then
                pesoDocumental = pesoDocumental + anPeso
            Else
                pesoVisual = pesoVisual + anPeso
            End If
        End If
        rsAnal.MoveNext()
    Loop
    rsAnal.Close()
    Set rsAnal = Nothing

    Dim usr, dataHj, horaHj
    usr    = Session("IdUsuario") & ""
    dataHj = InverteData(Date())
    horaHj = Mid(CStr(Time()), 1, 8)

    ' Monta transação
    Dim SQL
    SQL = "BEGIN TRANSACTION "
    SQL = SQL & "DELETE FROM MSNotasAnalises " & _
                "WHERE NA_FORNECE = '" & Replace(codigo, "'", "''") & "' " & _
                "AND NA_DOC    = '" & Replace(nota,   "'", "''") & "' " & _
                "AND NA_SERIE  = '" & Replace(serie,  "'", "''") & "' " & _
                "AND NA_ITEM   = '" & Replace(item,   "'", "''") & "' "

    If nInserts > 0 Then
        SQL = SQL & Join(insertParts, " ")
    End If

    SQL = SQL & "UPDATE MSNotasConferidas SET " & _
                "CO_PONTOS = " & (pesoDocumental + pesoVisual) & ", " & _
                "CO_DATALT = '" & dataHj & "', " & _
                "CO_HORALT = '" & horaHj & "', " & _
                "CO_USERALT = '" & Replace(usr, "'", "''") & "' " & _
                "WHERE CO_FORNECE = '" & Replace(codigo, "'", "''") & "' " & _
                "AND CO_DOC    = '" & Replace(nota,   "'", "''") & "' " & _
                "AND CO_SERIE  = '" & Replace(serie,  "'", "''") & "' " & _
                "AND CO_ITEM   = '" & Replace(item,   "'", "''") & "' "
    SQL = SQL & "IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

    Dim okTrans
    okTrans = ConectaLocalExec(cnxGravacao, SQL)

    If Not okTrans Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro ao reclassificar: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    ' Verifica se o UPDATE foi confirmado
    Dim sqlVerif
    sqlVerif = "SELECT COUNT(*) AS EXISTE FROM MSNotasConferidas " & _
               "WHERE CO_USERALT = '" & Replace(usr, "'", "''") & "' " & _
               "AND CO_DATALT = '" & dataHj & "' " & _
               "AND CO_HORALT = '" & horaHj & "'"

    Dim rsVerif
    Set rsVerif = ConectaLocal(cnxGravacao, sqlVerif)

    If IsNull(rsVerif) Or Not IsObject(rsVerif) Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar reclassificação: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    Dim existe
    existe = CInt(rsVerif("EXISTE") & "0")
    rsVerif.Close()
    Set rsVerif = Nothing

    If existe > 0 Then
        Salvar = "{""sucesso"": true, ""mensagem"": ""Registro reclassificado com sucesso!""}"
    Else
        Salvar = "{""sucesso"": false, ""mensagem"": ""Houve um erro e não foi possível realizar a reclassificação deste item.""}"
    End If
End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Reclassificação de Nota Fiscal</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-header { display: flex; align-items: center; gap: .75rem; margin-bottom: 1.25rem; }
    .page-title  { font-size: 1.15rem; font-weight: 700; color: #212529; margin: 0; }

    .filter-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .filter-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
    .filter-card-body   { padding: 1rem 1.25rem 1.25rem; }
    .filter-label       { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    .analise-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); height: 100%; }
    .analise-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
    .analise-card-body   { padding: 1rem 1.25rem; }
    .analise-card-body .form-check                    { padding-top: .3rem; padding-bottom: .3rem; border-bottom: 1px solid #f0f0f0; }
    .analise-card-body .form-check:last-child         { border-bottom: none; }
    .analise-card-body .form-check-label              { font-size: .85rem; cursor: pointer; }

    .table-card    { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .table-wrapper { overflow-x: auto; }
    .table thead th { background-color: #fff; color: #212529; font-size: .82rem; font-weight: 700; text-align: center; white-space: nowrap; border-color: #dee2e6; padding: .55rem .75rem; }
    .table tbody td { font-size: .82rem; color: #212529; vertical-align: middle; padding: .45rem .75rem; }

    /* Loading overlay sobre os cards de análise */
    .analise-area       { position: relative; min-height: 120px; }
    .overlay-analises   { position: absolute; inset: 0; background: rgba(255,255,255,.88); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 10; gap: .75rem; border-radius: 6px; }
    .overlay-analises .spinner-border { width: 2rem; height: 2rem; color: #6c757d; }
    .overlay-analises span            { font-size: .85rem; color: #495057; font-weight: 600; }

    .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
    .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
    @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
    .toast-custom .toast-progress-bar { height: 3px; animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }
  </style>
</head>
<body>
<div class="container-fluid py-3 px-3 px-lg-4">

  <!-- Toast container -->
  <div class="toast-container-custom" id="toastContainer"></div>

  <!-- Cabeçalho -->
  <div class="page-header">
    <button type="button" class="btn btn-light btn-sm" onclick="history.go(-1)" title="Voltar">
      <i class="bi bi-arrow-left me-2"></i>Voltar
    </button>
    <span class="page-title">Reclassificação de Nota Fiscal</span>
  </div>

  <!-- Área principal (preenchida via JS após carregarDados) -->
  <div id="areaConteudo">

    <!-- Placeholder inicial -->
    <div id="placeholderCarregando" class="d-flex flex-column align-items-center justify-content-center py-5 text-muted gap-3">
      <div class="spinner-border" style="width:2.5rem;height:2.5rem;color:#6c757d;" role="status">
        <span class="visually-hidden">Carregando...</span>
      </div>
      <span style="font-size:.9rem;font-weight:600;">Carregando dados...</span>
    </div>

    <!-- Alerta de não encontrado (oculto inicialmente) -->
    <div id="alertaNaoEncontrado" class="alert alert-warning d-flex align-items-center gap-2 d-none" role="alert">
      <i class="bi bi-exclamation-triangle-fill me-2"></i>
      <span><strong>Atenção!</strong> Item de nota fiscal não encontrado.
        <a href="#" onclick="history.go(-1)" class="alert-link ms-1">Voltar</a>
      </span>
    </div>

    <!-- Formulário (oculto até carregar) -->
    <div id="areaFormulario" class="d-none">

      <!-- Dados da NF -->
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

            <div class="col-6 col-md-2">
              <button type="button" id="btnSalvar" class="btn btn-success btn-sm">
                Salvar
              </button>
            </div>

          </div>

          <!-- Tabela do item -->
          <div class="table-card mt-3">
            <div class="table-wrapper">
              <table class="table table-bordered table-striped table-hover table-sm mb-0">
                <thead>
                  <tr>
                    <th>Lançamento</th>
                    <th class="text-center">Material</th>
                    <th class="text-center" style="width:55%">Descrição</th>
                    <th class="text-center">Quantidade</th>
                    <th class="text-center">UM</th>
                  </tr>
                </thead>
                <tbody id="tabelaItemBody"></tbody>
              </table>
            </div>
          </div>

        </div>
      </div>

      <!-- Análises -->
      <div class="analise-area">
        <div id="overlayAnalises" class="overlay-analises d-none">
          <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
          <span>Salvando...</span>
        </div>

        <div class="row g-3">

          <!-- Análise Documental -->
          <div class="col-12 col-lg-6">
            <div class="analise-card">
              <div class="analise-card-header">
                <i class="bi bi-file-earmark-check me-2"></i>Análise Documental
              </div>
              <div class="analise-card-body" id="listaDocumental"></div>
            </div>
          </div>

          <!-- Análise Visual -->
          <div class="col-12 col-lg-6">
            <div class="analise-card">
              <div class="analise-card-header">
                <i class="bi bi-eye me-2"></i>Análise Visual
              </div>
              <div class="analise-card-body" id="listaVisual"></div>
            </div>
          </div>

        </div>
      </div>

    </div><!-- /areaFormulario -->
  </div><!-- /areaConteudo -->

</div><!-- /container-fluid -->

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
"use strict";

const ReclassificaNotasInsumo = (function () {

    // ── Constantes ──────────────────────────────────────────
    const REPOSITORIO = 'ReclassificaNotasInsumo.asp';
    const TOAST_DELAY = 4500;
    const TOAST_ICON  = Object.freeze({
        sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754' },
        erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545' },
        aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404' },
        info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d' }
    });

    // ── Estado ──────────────────────────────────────────────
    let _nota     = null;   // objeto nota carregado
    let _analises = [];     // array de análises carregadas

    // ── Helpers ─────────────────────────────────────────────
    const $el = id => document.getElementById(id);

    function _getIdFromUrl() {
        const params = new URLSearchParams(window.location.search);
        return params.get('Id') || params.get('id') || '';
    }

    function _getFilialFromUrl() {
        const params = new URLSearchParams(window.location.search);
        return params.get('filial') || '';
    }

    // ── Toast ────────────────────────────────────────────────
    function exibirToast(tipo, mensagem) {
        const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
        const el  = document.createElement('div');
        el.className = 'toast-custom';
        el.innerHTML =
            '<div class="d-flex align-items-center px-3 py-2 border-bottom gap-2">' +
                '<i class="bi ' + cfg.icon + '" style="color:' + cfg.cor + ';font-size:1.1rem"></i>' +
                '<strong class="flex-grow-1" style="font-size:.88rem;color:' + cfg.cor + '">' +
                    tipo.charAt(0).toUpperCase() + tipo.slice(1) +
                '</strong>' +
                '<button type="button" style="background:none;border:none;font-size:1rem;cursor:pointer;color:#6c757d" ' +
                        'onclick="this.closest(\'.toast-custom\').remove()">&times;</button>' +
            '</div>' +
            '<div class="px-3 py-2" style="font-size:.85rem;color:#212529">' + mensagem + '</div>' +
            '<div class="toast-progress-bar" style="background:' + cfg.cor + '"></div>';
        $el('toastContainer').appendChild(el);
        setTimeout(function () { el.remove(); }, TOAST_DELAY);
    }

    // ── Formata quantidade ───────────────────────────────────
    function _fmtQtd(v) {
        const n = parseFloat(v);
        if (isNaN(n)) return v || '';
        return n.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    }

    // ── Carrega dados da nota + análises ─────────────────────
    function _carregar() {
        const id = _getIdFromUrl();
        if (!id) {
            $el('placeholderCarregando').classList.add('d-none');
            $el('alertaNaoEncontrado').classList.remove('d-none');
            return;
        }

        $.ajax({
            url      : REPOSITORIO,
            method   : 'GET',
            dataType : 'json',
            data     : { acao: 'carregarDados', id: id },
            success  : function (resp) {
                $el('placeholderCarregando').classList.add('d-none');
                if (!resp.sucesso) {
                    $el('alertaNaoEncontrado').classList.remove('d-none');
                    return;
                }
                _nota     = resp.nota;
                _analises = resp.analises || [];
                _preencherFormulario();
                _renderAnalises();
                $el('areaFormulario').classList.remove('d-none');
            },
            error    : function () {
                $el('placeholderCarregando').classList.add('d-none');
                $el('alertaNaoEncontrado').classList.remove('d-none');
            }
        });
    }

    // ── Preenche campos da NF ────────────────────────────────
    function _preencherFormulario() {
        $el('fCodigo').value = _nota.codigo;
        $el('fRazao').value  = _nota.razao;
        $el('fNota').value   = _nota.nota;
        $el('fSerie').value  = _nota.serie;
        $el('fItem').value   = _nota.item;

        // Linha da tabela
        const tbody = $el('tabelaItemBody');
        tbody.innerHTML =
            '<tr>' +
            '<td class="text-center">' + (_nota.dataLanc || '') + '</td>' +
            '<td class="text-center">' + (_nota.material || '') + '</td>' +
            '<td class="text-center">' + (_nota.descricao || '') + '</td>' +
            '<td class="text-center">' + _fmtQtd(_nota.quantidade) + '</td>' +
            '<td class="text-center">' + (_nota.um || '') + '</td>' +
            '</tr>';
    }

    // ── Renderiza checkboxes de análises ─────────────────────
    function _renderAnalises() {
        const docCont = $el('listaDocumental');
        const visCont = $el('listaVisual');

        // Verifica se já há algum item marcado em cada grupo
        const temDocMarcado = _analises.filter(a => a.tipo === 0 && a.marcado).length > 0;
        const temVisMarcado = _analises.filter(a => a.tipo === 1 && a.marcado).length > 0;

        // Helper para montar um grupo
        function _buildGrupo(container, tipo, temMarcado) {
            container.innerHTML = '';

            // Checkbox "Nenhuma inconformidade"
            const wrapNenhum = document.createElement('div');
            wrapNenhum.className = 'form-check';
            const chkNenhum = document.createElement('input');
            chkNenhum.type      = 'checkbox';
            chkNenhum.className = 'form-check-input';
            chkNenhum.id        = (tipo === 0 ? 'documental' : 'visual');
            chkNenhum.checked   = !temMarcado;
            // "Nenhuma inconformidade" fica desabilitado apenas quando há itens marcados
            chkNenhum.disabled  = temMarcado;
            chkNenhum.addEventListener('change', function () { _controlaGrupo(tipo); });
            const lblNenhum = document.createElement('label');
            lblNenhum.className  = 'form-check-label text-success';
            lblNenhum.htmlFor    = chkNenhum.id;
            lblNenhum.textContent = 'Nenhuma inconformidade';
            wrapNenhum.appendChild(chkNenhum);
            wrapNenhum.appendChild(lblNenhum);
            container.appendChild(wrapNenhum);

            // Checkboxes das análises
            _analises.filter(a => a.tipo === tipo).forEach(function (a) {
                const wrap = document.createElement('div');
                wrap.className = 'form-check';
                const chk = document.createElement('input');
                chk.type      = 'checkbox';
                chk.className = 'form-check-input';
                chk.id        = (tipo === 0 ? 'documental' : 'visual') + a.cod;
                chk.name      = chk.id;
                chk.value     = '1';
                chk.checked   = a.marcado;
                // Itens ficam desabilitados apenas quando "Nenhuma inconformidade" está marcado
                chk.disabled  = !temMarcado;
                chk.addEventListener('change', function () { _controlaGrupoItems(tipo); });
                const lbl = document.createElement('label');
                lbl.className  = 'form-check-label';
                lbl.htmlFor    = chk.id;
                lbl.textContent = a.descri;
                wrap.appendChild(chk);
                wrap.appendChild(lbl);
                container.appendChild(wrap);
            });
        }

        _buildGrupo(docCont, 0, temDocMarcado);
        _buildGrupo(visCont, 1, temVisMarcado);
    }

    // ── Controle de checkboxes ───────────────────────────────
    // Ao marcar "Nenhuma inconformidade": desabilita e desmarca todos os itens
    // Ao desmarcar "Nenhuma inconformidade": habilita os itens (usuário escolhe quais marcar)
    function _controlaGrupo(tipo) {
        const prefixo = tipo === 0 ? 'documental' : 'visual';
        const chkNenhum = $el(prefixo);
        const nenhummarcado = chkNenhum.checked;
        _analises.filter(a => a.tipo === tipo).forEach(function (a) {
            const chk = $el(prefixo + a.cod);
            if (chk) {
                chk.disabled = nenhummarcado;
                if (nenhummarcado) chk.checked = false;
            }
        });
    }

    // Ao clicar em item individual:
    // - se algum marcado → desabilita e desmarca "Nenhuma inconformidade"
    // - se todos desmarcados → habilita "Nenhuma inconformidade" (sem marcar automaticamente)
    function _controlaGrupoItems(tipo) {
        const prefixo = tipo === 0 ? 'documental' : 'visual';
        const temMarcado = _analises.filter(a => a.tipo === tipo).some(function (a) {
            const chk = $el(prefixo + a.cod);
            return chk && chk.checked;
        });
        const chkNenhum = $el(prefixo);
        if (chkNenhum) {
            chkNenhum.disabled = temMarcado;
            if (temMarcado) chkNenhum.checked = false;
        }
    }

    // ── Coleta parâmetros para envio ─────────────────────────
    function _coletarParams() {
        const params = {
            acao    : 'salvar',
            codigo  : _nota.codigo,
            nota    : _nota.nota,
            serie   : _nota.serie,
            item    : _nota.item,
            filial  : _getFilialFromUrl()   // necessário para rotear para a base correta
        };

        // "Nenhuma inconformidade"
        const chkNenhumDoc = $el('documental');
        const chkNenhumVis = $el('visual');
        if (chkNenhumDoc && chkNenhumDoc.checked) params['documental'] = '1';
        if (chkNenhumVis && chkNenhumVis.checked) params['visual']     = '1';

        // Itens individuais
        _analises.forEach(function (a) {
            const prefixo = a.tipo === 0 ? 'documental' : 'visual';
            const chk = $el(prefixo + a.cod);
            if (chk && chk.checked) params[prefixo + a.cod] = '1';
        });

        return params;
    }

    // ── Salvar reclassificação ───────────────────────────────
    function _salvar() {
        const params = _coletarParams();

        $.ajax({
            url        : REPOSITORIO,
            method     : 'POST',
            dataType   : 'json',
            data       : params,
            beforeSend : function () {
                $el('overlayAnalises').classList.remove('d-none');
                $el('btnSalvar').disabled = true;
            },
            success    : function (resp) {
                if (resp.sucesso) {
                    exibirToast('sucesso', resp.mensagem);
                    setTimeout(function () {
                        const filial = _getFilialFromUrl();
                        location.href = 'ListaNotasInsumo.asp?refresh=1' +
                                        (filial ? '&filial=' + encodeURIComponent(filial) : '');
                    }, 1200);
                } else {
                    exibirToast('erro', resp.mensagem || 'Erro ao salvar.');
                }
            },
            error      : function (xhr) {
                exibirToast('erro', 'Erro ao salvar reclassificação.');
                console.error(xhr.responseText);
            },
            complete   : function () {
                $el('overlayAnalises').classList.add('d-none');
                $el('btnSalvar').disabled = false;
            }
        });
    }

    // ── Init ─────────────────────────────────────────────────
    function init() {
        $('#btnSalvar').on('click', function () { _salvar(); });
        _carregar();
    }

    return Object.freeze({ init });

})();

$(function () { ReclassificaNotasInsumo.init(); });
</script>
</body>
</html>
