<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = True %>
<!--#Include File="../cnn2/cnn2.asp"-->
<%
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Response.Expires  = 0
Session.LCID      = 1046

' ============================================================
' HELPER: EscapeJSONLocal — escape correto, independente do funcoes.asp
' O EscapeJSON do funcoes.asp apaga aspas em vez de escapá-las,
' quebrando o JSON em campos com " (ex.: TUBETE 4"X2,51M)
' ============================================================
Function EscapeJSONLocal(val)
    Dim s, i, c, code
    If IsNull(val) Then
        EscapeJSONLocal = ""
        Exit Function
    End If
    s = CStr(val)
    ' Escapes obrigatórios primeiro (ordem importa: \ antes de ")
    s = Replace(s, Chr(92), Chr(92) & Chr(92))
    s = Replace(s, Chr(34), Chr(92) & Chr(34))
    ' Sequências nomeadas
    s = Replace(s, vbCrLf, Chr(92) & "n")
    s = Replace(s, vbCr,   Chr(92) & "n")
    s = Replace(s, vbLf,   Chr(92) & "n")
    s = Replace(s, vbTab,  Chr(92) & "t")
    ' Demais caracteres de controle Chr(0-31) exceto os já tratados acima
    ' (0x08=BS, 0x0C=FF — menos comuns mas invalidam JSON)
    Dim resultado
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
' HELPER: ConectaLocal
' Abre ADODB diretamente usando BuscaStringConexao() do cnn2,
' sem Response.Write / Response.End em caso de erro.
' Retorna Nothing em falha; popula sErroConectaLocal com mensagem.
' ============================================================
Dim sErroConectaLocal
sErroConectaLocal = ""

Function ConectaLocal(nomeCnx, sql)
    Const adOpenStatic     = 3
    Const adUseClient      = 3
    Const adLockReadOnly   = 1

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

    ' Desacopla recordset da conexão (cursor client-side)
    Set oRs.ActiveConnection = Nothing
    oCnn.Close
    Set oCnn = Nothing

    Set ConectaLocal = oRs
End Function

' ============================================================
' HELPER: ConectaLocalExec — executa DML sem retornar Recordset
' ============================================================
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
' DISPATCHER — endpoints JSON
' ============================================================
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.Clear()
    Response.ContentType = "application/json"

    Select Case acao
        Case "listar"  : resultado = Listar()
        Case "excluir" : resultado = Excluir()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSONLocal(acao) & """}"
    End Select

    Response.Clear()
    Response.Write resultado
    Response.End()

End If

' ============================================================
' FUNÇÃO: Listar
' Usa Join() sobre array de itens para evitar truncamento por
' concatenação O(n²) em strings longas do VBScript
' ============================================================
Function Listar()
    Dim filialF, razaoF, notaFiscalF, serieF, classificadasF
    filialF        = Trim(Request("filial"))
    razaoF         = Trim(Request("razao"))
    notaFiscalF    = Trim(Request("notaFiscal"))
    serieF         = Trim(Request("serie"))
    classificadasF = Trim(Request("classificadas"))

    ' VwNotasAnalises existe apenas no PIR-CORPORATIVO (dados de todas as filiais)
    Dim strConexao
    strConexao = "PIR-CORPORATIVO"

    SQL = "SELECT "
    SQL = SQL &   "IdNota, "
    SQL = SQL &   "Codigo, "
    SQL = SQL &   "Razao, "
    SQL = SQL &   "NotaFiscal, "
    SQL = SQL &   "Serie, "
    SQL = SQL &   "Item, "
    SQL = SQL &   "DataLancamento, "
    SQL = SQL &   "Material, "
    SQL = SQL &   "Descricao, "
    SQL = SQL &   "Quantidade, "
    SQL = SQL &   "Filial, "
    SQL = SQL &   "PONTOS, "
    SQL = SQL &   "PESO, "
    SQL = SQL &   "ENCONTRADO, "
    SQL = SQL &   "CO_PONTOS "
    SQL = SQL & "FROM "
    SQL = SQL &   "VwNotasAnalises "
    SQL = SQL & "WHERE "
    SQL = SQL &   "1=1 "

    If filialF <> "" Then
        SQL = SQL & "AND Filial = '" & Replace(filialF, "'", "''") & "' "
    End If
    If razaoF <> "" Then
        SQL = SQL & "AND Razao LIKE UPPER('%" & Replace(razaoF, "'", "''") & "%') "
    End If
    If notaFiscalF <> "" Then
        SQL = SQL & "AND NotaFiscal = '" & Replace(notaFiscalF, "'", "''") & "' "
    End If
    If serieF <> "" Then
        SQL = SQL & "AND Serie = '" & Replace(serieF, "'", "''") & "' "
    End If
    If classificadasF <> "" Then
        If classificadasF = "S" Then
            SQL = SQL & "AND CO_PONTOS IS NOT NULL "
        Else
            SQL = SQL & "AND CO_PONTOS IS NULL "
        End If
    End If

    SQL = SQL & "ORDER BY Razao, NotaFiscal, Serie, Item"

    Set rsLista = ConectaLocal(strConexao, SQL)
    If IsNull(rsLista) Or Not IsObject(rsLista) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar notas: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    ' --- Usa array + Join para evitar truncamento em grandes volumes ---
    Dim itens()
    Dim nItens
    nItens = 0
    ReDim itens(0)

    Do While Not rsLista.EOF
        Dim dtLanc
        dtLanc = rsLista("DataLancamento") & ""
        If dtLanc <> "" Then
            dtLanc = AcertaData(rsLista("DataLancamento"))
        End If

        Dim obj
        obj = "{"
        obj = obj & """idNota"":"        & CLng(rsLista("IdNota"))                                        & ","
        obj = obj & """codigo"":"        & Chr(34) & EscapeJSONLocal(rsLista("Codigo")      & "") & Chr(34)   & ","
        obj = obj & """razao"":"         & Chr(34) & EscapeJSONLocal(rsLista("Razao")        & "") & Chr(34)  & ","
        obj = obj & """notaFiscal"":"    & Chr(34) & EscapeJSONLocal(rsLista("NotaFiscal")   & "") & Chr(34)  & ","
        obj = obj & """serie"":"         & Chr(34) & EscapeJSONLocal(rsLista("Serie")        & "") & Chr(34)  & ","
        obj = obj & """item"":"          & Chr(34) & EscapeJSONLocal(rsLista("Item")         & "") & Chr(34)  & ","
        obj = obj & """dataLancamento"":" & Chr(34) & EscapeJSONLocal(dtLanc)                      & Chr(34)  & ","
        obj = obj & """material"":"      & Chr(34) & EscapeJSONLocal(rsLista("Material")     & "") & Chr(34)  & ","
        obj = obj & """descricao"":"     & Chr(34) & EscapeJSONLocal(rsLista("Descricao")    & "") & Chr(34)  & ","
        obj = obj & """quantidade"":"    & Chr(34) & EscapeJSONLocal(rsLista("Quantidade")   & "") & Chr(34)  & ","
        obj = obj & """filial"":"        & Chr(34) & EscapeJSONLocal(rsLista("Filial")        & "") & Chr(34) & ","
        obj = obj & """pontos"":"        & Chr(34) & EscapeJSONLocal(rsLista("PONTOS")        & "") & Chr(34) & ","
        obj = obj & """peso"":"          & Chr(34) & EscapeJSONLocal(rsLista("PESO")          & "") & Chr(34) & ","
        obj = obj & """encontrado"":"    & Chr(34) & EscapeJSONLocal(rsLista("ENCONTRADO")    & "") & Chr(34) & ","
        If IsNull(rsLista("CO_PONTOS")) Then
            obj = obj & """coPontos"":null,""classificado"":false"
        Else
            obj = obj & """coPontos"":"  & Chr(34) & EscapeJSONLocal(rsLista("CO_PONTOS") & "") & Chr(34)     & ",""classificado"":true"
        End If
        obj = obj & "}"

        ReDim Preserve itens(nItens)
        itens(nItens) = obj
        nItens = nItens + 1

        rsLista.MoveNext
    Loop

    rsLista.Close()
    Set rsLista = Nothing

    Dim jsonArray
    If nItens = 0 Then
        jsonArray = "[]"
    Else
        jsonArray = "[" & Join(itens, ",") & "]"
    End If

    Listar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & jsonArray & "}"
End Function

' ============================================================
' FUNÇÃO: Excluir
' ============================================================
Function Excluir()
    Dim idNota
    idNota = Trim(Request("idNota") & "")

    If idNota = "" Or Not IsNumeric(idNota) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Identificador inválido.""}"
        Exit Function
    End If

    Dim usr, data, hora
    usr  = Session("IdUsuario")
    data = InverteData(Date())
    hora = Mid(CStr(Time()), 1, 8)

    ' Busca dados da nota no PIR-CORPORATIVO (view centralizada)
    Dim strConexaoExcl
    strConexaoExcl = "PIR-CORPORATIVO"

    SQL = "SELECT IdNota, Codigo, NotaFiscal, Serie, Item, Filial "
    SQL = SQL & "FROM VwNotasAnalises "
    SQL = SQL & "WHERE IdNota = " & idNota

    Dim rsBusca
    Set rsBusca = ConectaLocal(strConexaoExcl, SQL)

    If IsNull(rsBusca) Or Not IsObject(rsBusca) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao localizar nota: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    If rsBusca.EOF Then
        rsBusca.Close()
        Set rsBusca = Nothing
        Excluir = "{""sucesso"": false, ""mensagem"": ""Não foi possível encontrar a classificação para o item selecionado.""}"
        Exit Function
    End If

    Dim codForn, numNota, serieNota, itemNota
    codForn   = rsBusca("Codigo")     & ""
    numNota   = rsBusca("NotaFiscal") & ""
    serieNota = rsBusca("Serie")      & ""
    itemNota  = rsBusca("Item")       & ""
    rsBusca.Close()
    Set rsBusca = Nothing

    ' Transação DELETE + soft-delete UPDATE
    SQL = "BEGIN TRANSACTION "
    SQL = SQL &   "DELETE FROM MSNotasAnalises "
    SQL = SQL &   "WHERE NA_FORNECE = '" & Replace(codForn,   "'","''") & "' "
    SQL = SQL &   "AND   NA_DOC     = '" & Replace(numNota,   "'","''") & "' "
    SQL = SQL &   "AND   NA_SERIE   = '" & Replace(serieNota, "'","''") & "' "
    SQL = SQL &   "AND   NA_ITEM    = '" & Replace(itemNota,  "'","''") & "' "
    SQL = SQL &   "UPDATE MSNotasConferidas SET "
    SQL = SQL &     "D_E_L_E_T_ = '*', "
    SQL = SQL &     "CO_DATALT  = '" & data & "', "
    SQL = SQL &     "CO_HORALT  = '" & hora & "', "
    SQL = SQL &     "CO_USERALT = '" & Replace(usr, "'","''") & "' "
    SQL = SQL &   "WHERE CO_FORNECE = '" & Replace(codForn,   "'","''") & "' "
    SQL = SQL &   "AND   CO_DOC     = '" & Replace(numNota,   "'","''") & "' "
    SQL = SQL &   "AND   CO_SERIE   = '" & Replace(serieNota, "'","''") & "' "
    SQL = SQL &   "AND   CO_ITEM    = '" & Replace(itemNota,  "'","''") & "' "
    SQL = SQL & "IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

    Dim okTrans
    okTrans = ConectaLocalExec(strConexaoExcl, SQL)
    If Not okTrans Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao executar exclusão: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    ' Verifica soft-delete
    SQL = "SELECT COUNT(*) AS EXISTE FROM MSNotasConferidas "
    SQL = SQL & "WHERE D_E_L_E_T_ = '*' "
    SQL = SQL & "AND CO_DATALT  = '" & data & "' "
    SQL = SQL & "AND CO_HORALT  = '" & hora & "' "
    SQL = SQL & "AND CO_USERALT = '" & Replace(usr, "'","''") & "'"

    Dim rsVerifica
    Set rsVerifica = ConectaLocal(strConexaoExcl, SQL)
    If IsNull(rsVerifica) Or Not IsObject(rsVerifica) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar exclusão: " & EscapeJSONLocal(sErroConectaLocal) & """}"
        Exit Function
    End If

    Dim existeExclusao
    existeExclusao = CInt(rsVerifica("EXISTE") & "")
    rsVerifica.Close()
    Set rsVerifica = Nothing

    If existeExclusao > 0 Then
        Excluir = "{""sucesso"": true, ""mensagem"": ""Registro excluído com sucesso!""}"
    Else
        Excluir = "{""sucesso"": false, ""mensagem"": ""Houve um erro e não foi possível realizar a exclusão deste item.""}"
    End If
End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Lista de Notas Fiscais de Insumo</title>

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />

    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
        /* ── Layout ── */
        body { background-color: #f0f0f0; min-height: 100vh; }
        .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

        /* ── Filter card ── */
        .filter-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
        .filter-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
        .filter-card-body   { padding: 1rem 1.25rem 1.25rem; }
        .filter-label       { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

        /* ── Table card ── */
        .table-card { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }

        /* ── Toolbar ── */
        .table-toolbar { background-color: #fff; border-radius: 6px 6px 0 0; }

        /* ── Search box ── */
        .search-box { position: relative; max-width: 200px; }
        .search-box .form-control { padding-right: 76px; font-size: .82rem; }
        .search-box .search-hotkey { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); pointer-events: none; user-select: none; display: flex; gap: 2px; }
        .search-box .search-hotkey kbd { font-size: .65rem; padding: .1rem .3rem; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 3px; color: #6c757d; }

        /* ── Table ── */
        .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
        .table-wrapper table { min-width: max-content; }
        .table thead th { position: sticky; top: 0; background-color: #fff; color: #212529; font-size: .82rem; font-weight: 700; text-align: center; white-space: nowrap; z-index: 2; border-color: #dee2e6; padding: .55rem .75rem; }
        .table tbody td { font-size: .82rem; text-align: center; white-space: nowrap; color: #212529; vertical-align: middle; padding: .45rem .75rem; }
        .table tbody td.td-razao,
        .table tbody td.td-descricao { text-align: left; }

        /* ── Placeholder ── */
        .table-placeholder { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 3.5rem 1rem; color: #6c757d; gap: .75rem; }
        .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
        .table-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

        /* ── Footer ── */
        .table-footer        { display: flex; align-items: center; justify-content: space-between; padding: .6rem 1rem; border-top: 1px solid #dee2e6; flex-wrap: wrap; gap: .5rem; }
        .table-footer-left   { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
        .table-footer-center { flex: 1; display: flex; justify-content: center; }
        .table-footer-right  { min-width: 140px; text-align: right; }
        .info-registros      { font-size: .8rem; color: #6c757d; }
        .select-limite       { width: auto; }

        /* ── Paginação ── */
        #paginacao .page-link                   { color: #6c757d; border-color: #dee2e6; }
        #paginacao .page-item.active .page-link { background-color: #6c757d; border-color: #6c757d; color: #fff; }
        #paginacao .page-link:hover             { background-color: #e9ecef; color: #495057; }
        #paginacao .page-item.disabled .page-link { color: #adb5bd; }

        /* ── Overlay — dentro da .table-data-area (position:relative) ── */
        .table-data-area { position: relative; min-height: 120px; }
        .overlay-loading { position: absolute; inset: 0; background: rgba(255,255,255,.85); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 10; gap: 1rem; border-radius: 0 0 6px 6px; }
        .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
        .overlay-loading span            { font-size: .9rem; font-weight: 600; color: #495057; }

        /* ── Toast ── */
        .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
        .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
        @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
        .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
        @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

        /* ── Linha classificada ── */
        tr.linha-classificada td { background-color: #d1e7dd !important; }
    </style>
</head>
<body>
<div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center justify-content-between mb-3">
        <span class="page-title">Lista de Notas Fiscais de Insumo</span>
        <div class="d-flex gap-2">
            <button id="btnImportarFornecedor" type="button" class="btn btn-outline-primary btn-sm"
                    title="Importar Fornecedores">
                <i class="bi bi-people me-2"></i>Fornecedores
            </button>
            <button id="btnImportarNotas" type="button" class="btn btn-outline-primary btn-sm"
                    title="Importar Notas Fiscais">
                <i class="bi bi-upc-scan me-2"></i>Notas
            </button>
        </div>
    </div>

    <!-- Card de Filtros -->
    <div class="filter-card mb-3">
        <div class="filter-card-header">
            <i class="bi bi-funnel me-2"></i>Filtros
        </div>
        <div class="filter-card-body">
            <div class="row g-3 align-items-end">

                <div class="col-12 col-sm-6 col-lg-1">
                    <label class="filter-label" for="filtroFilial">Filial</label>
                    <select id="filtroFilial" class="form-select form-select-sm">
                        <option value="">Todas</option>
                        <option value="00">Pirapetinga</option>
                        <option value="04">Uberaba</option>
                        <option value="05">Bento Gonçalves</option>
                        <option value="06">Fortaleza</option>
                        <option value="07">Saquarema</option>
                    </select>
                </div>

                <div class="col-12 col-sm-6 col-lg-3">
                    <label class="filter-label" for="filtroRazao">Razão Social</label>
                    <input type="text" id="filtroRazao" class="form-control form-control-sm"
                           placeholder="Buscar por razão..." autocomplete="off" />
                </div>

                <div class="col-12 col-sm-4 col-lg-2">
                    <label class="filter-label" for="filtroNotaFiscal">Nota Fiscal</label>
                    <input type="text" id="filtroNotaFiscal" class="form-control form-control-sm"
                           placeholder="Nº da nota" autocomplete="off" />
                </div>

                <div class="col-12 col-sm-4 col-lg-1">
                    <label class="filter-label" for="filtroSerie">Série</label>
                    <input type="text" id="filtroSerie" class="form-control form-control-sm"
                           placeholder="Série" autocomplete="off" />
                </div>

                <div class="col-12 col-sm-4 col-lg-1">
                    <label class="filter-label" for="filtroClassificadas">Tipo</label>
                    <select id="filtroClassificadas" class="form-select form-select-sm">
                        <option value="">Todas</option>
                        <option value="S">Classificadas</option>
                        <option value="N" selected>Não Classificadas</option>
                    </select>
                </div>

                <!-- Botões alinhados à direita usando ms-auto -->
                <div class="col-12">
                    <div class="d-flex gap-2 justify-content-end">
                        <button id="btnLimpar" type="button" class="btn btn-light btn-sm">
                            <i class="bi bi-eraser me-2"></i>Limpar
                        </button>
                        <button id="btnFiltrar" type="button" class="btn btn-primary btn-sm">
                            <i class="bi bi-funnel-fill me-2"></i>Filtrar
                        </button>
                    </div>
                </div>

            </div>
        </div>
    </div>

    <!-- Card de Tabela -->
    <div class="table-card">

        <!-- Toolbar -->
        <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
            <div class="d-flex gap-2">
                <button id="btnExportCsv" type="button" class="btn btn-success btn-sm">
                    <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
                </button>
                <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm">
                    <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
                </button>
            </div>
            <div class="search-box">
                <input type="text" id="searchBox" class="form-control form-control-sm"
                       placeholder="Buscar..." autocomplete="off" />
                <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
            </div>
        </div>

        <!-- Área de dados: position:relative contém o overlay -->
        <div class="table-data-area">

            <!-- Overlay de carregamento — posicionado dentro da área de dados -->
            <div id="overlayLoading" class="overlay-loading d-none">
                <div class="spinner-border" role="status">
                    <span class="visually-hidden">Carregando...</span>
                </div>
                <span>Carregando...</span>
            </div>

            <!-- Placeholder inicial -->
            <div id="tablePlaceholder" class="table-placeholder">
                <i class="bi bi-clipboard-data"></i>
                <p>Carregando notas não classificadas...</p>
            </div>

            <!-- Wrapper da tabela -->
            <div id="tableWrapper" class="table-wrapper d-none">
                <table class="table table-bordered table-striped table-hover table-sm mb-0">
                    <thead>
                        <tr>
                            <th>Planta</th>
                            <th>Cód. Forn.</th>
                            <th>Razão Social</th>
                            <th>Nota Fiscal</th>
                            <th>Série</th>
                            <th>Item</th>
                            <th>Data Lanç.</th>
                            <th>Material</th>
                            <th>Descrição</th>
                            <th>Quantidade</th>
                            <th>Pontos</th>
                            <th>Peso</th>
                            <th>Ação</th>
                        </tr>
                    </thead>
                    <tbody id="tabelaBody"></tbody>
                </table>
            </div>

        </div><!-- /table-data-area -->

        <!-- Rodapé -->
        <div id="tableFooter" class="table-footer d-none">
            <div class="table-footer-left">
                <label for="selectLimite" class="form-label mb-0 text-muted"
                       style="font-size:.8rem;white-space:nowrap;">Linhas por página:</label>
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
            </div>
        </div>

    </div><!-- /table-card -->
</div><!-- /container-fluid -->

<!-- Modal de confirmação de exclusão -->
<div class="modal fade" id="modalExcluir" tabindex="-1" aria-labelledby="modalExcluirLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="modalExcluirLabel">
                    <i class="bi bi-exclamation-triangle-fill text-warning me-2"></i>Confirmar Exclusão
                </h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
            </div>
            <div class="modal-body">
                Confirma a <strong>remoção da classificação</strong> desta nota fiscal?
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancelar</button>
                <button type="button" id="btnConfirmarExclusao" class="btn btn-danger btn-sm">
                    Excluir
                </button>
            </div>
        </div>
    </div>
</div>

<!-- Toast container -->
<div class="toast-container-custom" id="toastContainer"></div>

<!-- Bootstrap JS -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>

<!-- jQuery -->
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<!-- jsPDF + autoTable -->
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

<script>
const ListaNotasInsumo = (function () {

    // ── Constantes ─────────────────────────────────────────
    const REPOSITORIO = 'ListaNotasInsumo.asp';
    const LOGO_PATH   = '/corporativo/suprimentos/Logo/logo-smurfit1.png';
    const TOAST_DELAY = 4500;
    const TOAST_ICON  = {
        sucesso: { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
        erro:    { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
        aviso:   { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
        info:    { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' }
    };

    const FILIAL_NOME = {
        '00': 'Pirapetinga',
        '04': 'Uberaba',
        '05': 'Bento Gonçalves',
        '06': 'Fortaleza',
        '07': 'Saquarema'
    };

    const CABECALHOS_CSV = [
        'Planta','Cód. Fornecedor','Razão Social','Nota Fiscal','Série','Item',
        'Data Lançamento','Material','Descrição','Quantidade','Pontos','Peso','Classificado'
    ];

    // ── Cache (sessionStorage) ─────────────────────────────
    const CACHE_KEY = 'ListaNotasInsumo_cache';

    const _Cache = {
        salvar: function () {
            try {
                const estado = {
                    filtros: {
                        filial:        $el('filtroFilial').val(),
                        razao:         $el('filtroRazao').val(),
                        notaFiscal:    $el('filtroNotaFiscal').val(),
                        serie:         $el('filtroSerie').val(),
                        classificadas: $el('filtroClassificadas').val()
                    },
                    searchTermo:   $el('searchBox').val(),
                    limiteLinhas:  limiteLinhas,
                    paginaAtual:   paginaAtual,
                    dadosCompletos: dadosCompletos
                };
                sessionStorage.setItem(CACHE_KEY, JSON.stringify(estado));
            } catch (e) { /* sessionStorage cheio ou indisponível — ignora */ }
        },
        carregar: function () {
            try {
                const raw = sessionStorage.getItem(CACHE_KEY);
                return raw ? JSON.parse(raw) : null;
            } catch (e) { return null; }
        },
        limpar: function () {
            try { sessionStorage.removeItem(CACHE_KEY); } catch (e) {}
        },
        marcarSaida: function (destino) {
            try { sessionStorage.setItem(CACHE_KEY + '_saida', destino); } catch (e) {}
        },
        validarRetorno: function () {
            try {
                const origens = [
                    'ImportacaoFornecedores.asp',
                    'ImportacaoNotasFiscais.asp',
                    'ClassificaNotasinsumo.asp',
                    'ReclassificaNotasinsumo.asp'
                ];
                const saida  = sessionStorage.getItem(CACHE_KEY + '_saida') || '';
                const valido = origens.some(function (o) { return saida.indexOf(o) !== -1; });
                sessionStorage.removeItem(CACHE_KEY + '_saida');
                if (!valido) { sessionStorage.removeItem(CACHE_KEY); }
            } catch (e) {}
        }
    };

    // ── Estado ─────────────────────────────────────────────
    let dadosCompletos = [];
    let dadosFiltrados = [];
    let paginaAtual    = 1;
    let totalPaginas   = 1;
    let totalRegistros = 0;
    let limiteLinhas   = 50;
    let idNotaPendente = null;
    let modalExcluir   = null;
    let searchDebounce = null;

    // ── Helper $ ───────────────────────────────────────────
    const $el = id => $('#' + id);

    // ── Toast ──────────────────────────────────────────────
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
        document.getElementById('toastContainer').appendChild(el);
        setTimeout(function () { el.remove(); }, TOAST_DELAY);
    }

    // ── Helpers ────────────────────────────────────────────
    function nomeFilial(cod) {
        return FILIAL_NOME[String(cod)] || cod || '';
    }

    // ── Carregar dados ─────────────────────────────────────
    function _carregar() {
        $.ajax({
            url: REPOSITORIO,
            method: 'GET',
            dataType: 'json',
            data: {
                acao:          'listar',
                filial:        $el('filtroFilial').val(),
                razao:         $el('filtroRazao').val(),
                notaFiscal:    $el('filtroNotaFiscal').val(),
                serie:         $el('filtroSerie').val(),
                classificadas: $el('filtroClassificadas').val()
            },
            beforeSend: function () {
                document.querySelector('.table-data-area').style.minHeight = '220px';
                $el('overlayLoading').removeClass('d-none');
                $el('tablePlaceholder').addClass('d-none');
                $el('tableWrapper').addClass('d-none');
                $el('tableFooter').addClass('d-none');
            },
            success: function (resp) {
                if (!resp.sucesso) {
                    exibirToast('erro', resp.mensagem || 'Erro ao carregar dados.');
                    $el('tablePlaceholder').removeClass('d-none').find('p').html(
                        'Selecione os filtros desejados e clique em <strong>Filtrar</strong> para exibir os resultados.');
                    return;
                }
                dadosCompletos = resp.dados || [];
                dadosFiltrados = dadosCompletos.slice();
                totalRegistros = dadosFiltrados.length;
                limiteLinhas   = parseInt($el('selectLimite').val(), 10) || 50;
                totalPaginas   = Math.max(1, Math.ceil(totalRegistros / limiteLinhas));
                paginaAtual    = 1;
                _Cache.salvar();

                if (totalRegistros === 0) {
                    $el('tablePlaceholder').removeClass('d-none').find('p').html(
                        'Nenhum registro encontrado para os filtros informados.');
                } else {
                    $el('tableWrapper').removeClass('d-none');
                    $el('tableFooter').removeClass('d-none');
                    _renderTabela();
                    _renderPaginacao(paginaAtual, totalPaginas);
                    $el('infoRegistros').text('Total: ' + totalRegistros + ' registro(s)');
                }
            },
            error: function (xhr) {
                exibirToast('erro', 'Erro ao carregar dados.');
                console.error(xhr.responseText);
                $el('tablePlaceholder').removeClass('d-none').find('p').html(
                    'Selecione os filtros desejados e clique em <strong>Filtrar</strong> para exibir os resultados.');
            },
            complete: function () {
                $el('overlayLoading').addClass('d-none');
                document.querySelector('.table-data-area').style.minHeight = '';
            }
        });
    }

    // ── Render tabela ──────────────────────────────────────
    function _renderTabela() {
        const inicio = (paginaAtual - 1) * limiteLinhas;
        const fatia  = dadosFiltrados.slice(inicio, inicio + limiteLinhas);
        const $tbody = $el('tabelaBody').empty();

        if (fatia.length === 0) {
            $tbody.append('<tr><td colspan="13" class="text-center text-muted py-3">Nenhum registro nesta página.</td></tr>');
            return;
        }

        fatia.forEach(function (row) {
            const planta  = nomeFilial(row.filial);
            const qtdFmt  = row.quantidade !== '' ? parseFloat(row.quantidade).toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}) : '';
            const classTr = row.classificado ? 'linha-classificada' : '';

            let btnAcoes = '';
            if (!row.classificado) {
                btnAcoes = '<button type="button" class="btn btn-success btn-sm btn-classificar"' +
                           ' data-id="' + row.idNota + '" data-filial="' + row.filial + '" title="Classificar">' +
                           '<i class="bi bi-check2-square"></i></button>';
            } else {
                btnAcoes = '<button type="button" class="btn btn-primary btn-sm btn-reclassificar me-1"' +
                           ' data-id="' + row.idNota + '" data-filial="' + row.filial + '" title="Reclassificar">' +
                           '<i class="bi bi-arrow-repeat"></i></button>' +
                           '<button type="button" class="btn btn-danger btn-sm btn-excluir"' +
                           ' data-id="' + row.idNota + '" title="Remover Classificação">' +
                           '<i class="bi bi-trash"></i></button>';
            }

            $tbody.append(
                '<tr class="' + classTr + '">' +
                '<td>' + planta + '</td>' +
                '<td>' + row.codigo + '</td>' +
                '<td class="td-razao">' + row.razao + '</td>' +
                '<td>' + row.notaFiscal + '</td>' +
                '<td>' + row.serie + '</td>' +
                '<td>' + row.item + '</td>' +
                '<td>' + row.dataLancamento + '</td>' +
                '<td>' + row.material + '</td>' +
                '<td class="td-descricao">' + row.descricao + '</td>' +
                '<td>' + qtdFmt + '</td>' +
                '<td>' + row.pontos + '</td>' +
                '<td>' + row.peso + '</td>' +
                '<td class="text-nowrap">' + btnAcoes + '</td>' +
                '</tr>'
            );
        });
    }

    // ── Paginação ──────────────────────────────────────────
    function _renderPaginacao(atual, total) {
        const $ul = $el('paginacao').empty();
        if (total <= 1) return;

        const mkItem = function (label, page, disabled, active) {
            const $li = $('<li>').addClass('page-item');
            if (disabled) $li.addClass('disabled');
            if (active)   $li.addClass('active');
            const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
            if (!disabled && !active) {
                $a.on('click', function (e) {
                    e.preventDefault();
                    paginaAtual = page;
                    _renderTabela();
                    _renderPaginacao(paginaAtual, totalPaginas);
                    _Cache.salvar();
                });
            }
            return $li.append($a);
        };

        $ul.append(mkItem('&laquo;',  1,            atual === 1,     false));
        $ul.append(mkItem('&lsaquo;', atual - 1,    atual === 1,     false));
        const inicio = Math.max(1, atual - 2);
        const fim    = Math.min(total, atual + 2);
        for (let p = inicio; p <= fim; p++) {
            $ul.append(mkItem(p, p, false, p === atual));
        }
        $ul.append(mkItem('&rsaquo;', atual + 1,    atual === total, false));
        $ul.append(mkItem('&raquo;',  total,         atual === total, false));
    }

    // ── Search em memória ──────────────────────────────────
    function _filtrarMemoria(termo) {
        termo = termo.toLowerCase().trim();
        dadosFiltrados = termo === ''
            ? dadosCompletos.slice()
            : dadosCompletos.filter(function (r) {
                return (r.razao + ' ' + r.notaFiscal + ' ' + r.material + ' ' +
                        r.descricao + ' ' + r.codigo + ' ' + nomeFilial(r.filial))
                        .toLowerCase().indexOf(termo) !== -1;
            });
        totalRegistros = dadosFiltrados.length;
        totalPaginas   = Math.max(1, Math.ceil(totalRegistros / limiteLinhas));
        paginaAtual    = 1;
        _renderTabela();
        _renderPaginacao(paginaAtual, totalPaginas);
        $el('infoRegistros').text('Total: ' + totalRegistros + ' registro(s)');
    }

    // ── Excluir ────────────────────────────────────────────
    function _confirmarExclusao(idNota) {
        idNotaPendente = idNota;
        modalExcluir.show();
    }

    function _executarExclusao() {
        if (!idNotaPendente) return;
        const id = idNotaPendente;
        modalExcluir.hide();

        $.ajax({
            url: REPOSITORIO,
            method: 'POST',
            dataType: 'json',
            data: { acao: 'excluir', idNota: id },
            beforeSend: function () {
                $el('overlayLoading').removeClass('d-none');
            },
            success: function (resp) {
                if (resp.sucesso) {
                    exibirToast('sucesso', resp.mensagem);
                    _carregar();
                } else {
                    exibirToast('erro', resp.mensagem || 'Erro ao excluir.');
                }
            },
            error: function (xhr) {
                exibirToast('erro', 'Erro ao excluir classificação.');
                console.error(xhr.responseText);
            },
            complete: function () {
                $el('overlayLoading').addClass('d-none');
                idNotaPendente = null;
            }
        });
    }

    // ── Exportar CSV ───────────────────────────────────────
    function _exportarCSV() {
        if (dadosFiltrados.length === 0) { exibirToast('aviso', 'Não há dados para exportar.'); return; }

        const esc = function (v) {
            const s = v == null ? '' : String(v);
            return (s.indexOf(';') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
                ? '"' + s.replace(/"/g, '""') + '"' : s;
        };
        const linhas = dadosFiltrados.map(function (r) {
            return [
                nomeFilial(r.filial), r.codigo, r.razao, r.notaFiscal, r.serie, r.item,
                r.dataLancamento, r.material, r.descricao,
                r.quantidade !== '' ? parseFloat(r.quantidade).toFixed(2).replace('.', ',') : '',
                r.pontos, r.peso, r.classificado ? 'Sim' : 'Não'
            ].map(esc).join(';');
        });
        const conteudo = [CABECALHOS_CSV.map(esc).join(';')].concat(linhas).join('\r\n');
        const blob = new Blob(['\uFEFF' + conteudo], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href = url; a.download = 'ListaNotasInsumo.csv';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ── Exportar PDF ───────────────────────────────────────
    function _exportarPDF() {
        if (dadosFiltrados.length === 0) { exibirToast('aviso', 'Não há dados para exportar.'); return; }

        const img = new Image();
        img.onload  = function () { _gerarPDF(img); };
        img.onerror = function () { _gerarPDF(null); };
        img.src = LOGO_PATH;
    }

    function _gerarPDF(imgLogo) {
        const { jsPDF } = window.jspdf;
        const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
        const pageW = doc.internal.pageSize.getWidth();

        // ── Logo (canvas 900x315, posição padrão SWnet) ───
        if (imgLogo) {
            const canvas  = document.createElement('canvas');
            canvas.width  = 900;
            canvas.height = 315;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(imgLogo, 0, 0, 900, 315);
            const dataUrl = canvas.toDataURL('image/png');
            doc.addImage(dataUrl, 'PNG', 14, 5, 38, 13);
        }

        // ── Título e data ─────────────────────────────────
        doc.setFontSize(13);
        doc.text('Lista de Notas Fiscais de Insumo', pageW / 2, 14, { align: 'center' });
        doc.setFontSize(7);
        doc.text('Emitido em: ' + new Date().toLocaleString('pt-BR'), pageW - 10, 8, { align: 'right' });

        // ── Tabela ────────────────────────────────────────
        const head = [['Planta','Cód. Forn.','Razão Social','Nota Fiscal','Série','Item',
                        'Data Lanç.','Material','Descrição','Qtd','Pontos','Peso','Classif.']];
        const body = dadosFiltrados.map(function (r) {
            return [
                nomeFilial(r.filial), r.codigo, r.razao, r.notaFiscal, r.serie, r.item,
                r.dataLancamento, r.material, r.descricao,
                r.quantidade !== '' ? parseFloat(r.quantidade).toFixed(2).replace('.', ',') : '',
                r.pontos, r.peso, r.classificado ? 'Sim' : 'Não'
            ];
        });

        doc.autoTable({
            startY: 22, theme: 'grid', head: head, body: body,
            headStyles:         { fillColor: [74,74,74], textColor: [255,255,255], fontSize: 4.5, cellPadding: 1.5 },
            alternateRowStyles: { fillColor: [245,245,245] },
            bodyStyles:         { fontSize: 4.5, cellPadding: 1.5, lineColor: [180,180,180], lineWidth: 0.2 }
        });
        doc.save('ListaNotasInsumo.pdf');
    }

    // ── Init ───────────────────────────────────────────────
    function init() {
        modalExcluir = new bootstrap.Modal(document.getElementById('modalExcluir'));

        // ── Restaurar estado cacheado ──────────────────────
        (function _restaurarCache() {
            _Cache.validarRetorno();

            // Voltou de Classificar/Reclassificar com ?refresh=1:
            // invalida os dados stale e rebusca com os filtros anteriores.
            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.get('refresh') === '1') {
                const estado = _Cache.carregar();
                _Cache.limpar();
                if (estado && estado.filtros) {
                    const f = estado.filtros;
                    const filialUrl = urlParams.get('filial') || '';
                    $el('filtroFilial').val(filialUrl || f.filial || '');
                    $el('filtroRazao').val(f.razao || '');
                    $el('filtroNotaFiscal').val(f.notaFiscal || '');
                    $el('filtroSerie').val(f.serie || '');
                    $el('filtroClassificadas').val(
                        urlParams.get('classificadas') || f.classificadas || ''
                    );
                    if (estado.searchTermo) $el('searchBox').val(estado.searchTermo);
                    if (estado.limiteLinhas) {
                        limiteLinhas = estado.limiteLinhas;
                        $el('selectLimite').val(String(limiteLinhas));
                    }
                } else {
                    if (urlParams.get('filial'))        $el('filtroFilial').val(urlParams.get('filial'));
                    if (urlParams.get('classificadas')) $el('filtroClassificadas').val(urlParams.get('classificadas'));
                }
                // Remove o ?refresh=1 da URL sem recarregar
                try { window.history.replaceState(null, '', window.location.pathname); } catch (e) {}
                // Rebusca automaticamente se há algum filtro ativo
                const temFiltro = $el('filtroFilial').val()        ||
                                  $el('filtroRazao').val()         ||
                                  $el('filtroNotaFiscal').val()    ||
                                  $el('filtroSerie').val()         ||
                                  $el('filtroClassificadas').val();
                if (temFiltro) { _carregar(); }
                return;
            }

            const estado = _Cache.carregar();
            if (!estado || !estado.dadosCompletos || estado.dadosCompletos.length === 0) {
                // Nenhum cache válido — carrega automaticamente com o filtro padrão (Não Classificadas)
                _carregar();
                return;
            }

            // Preenche filtros
            const f = estado.filtros || {};
            $el('filtroFilial').val(f.filial || '');
            $el('filtroRazao').val(f.razao || '');
            $el('filtroNotaFiscal').val(f.notaFiscal || '');
            $el('filtroSerie').val(f.serie || '');
            $el('filtroClassificadas').val(f.classificadas || '');

            // Restaura estado de dados
            dadosCompletos = estado.dadosCompletos;
            limiteLinhas   = estado.limiteLinhas || 50;
            paginaAtual    = estado.paginaAtual  || 1;

            $el('selectLimite').val(String(limiteLinhas));

            // Aplica search se havia termo
            const termo = (estado.searchTermo || '').toLowerCase().trim();
            if (termo) {
                $el('searchBox').val(estado.searchTermo);
                dadosFiltrados = dadosCompletos.filter(function (r) {
                    return (r.razao + ' ' + r.notaFiscal + ' ' + r.material + ' ' +
                            r.descricao + ' ' + r.codigo + ' ' + nomeFilial(r.filial))
                            .toLowerCase().indexOf(termo) !== -1;
                });
            } else {
                dadosFiltrados = dadosCompletos.slice();
            }

            totalRegistros = dadosFiltrados.length;
            totalPaginas   = Math.max(1, Math.ceil(totalRegistros / limiteLinhas));
            if (paginaAtual > totalPaginas) paginaAtual = 1;

            $el('tablePlaceholder').addClass('d-none');
            $el('tableWrapper').removeClass('d-none');
            $el('tableFooter').removeClass('d-none');
            _renderTabela();
            _renderPaginacao(paginaAtual, totalPaginas);
            $el('infoRegistros').text('Total: ' + totalRegistros + ' registro(s)');
        }());

        $el('btnFiltrar').on('click', function () { _carregar(); });

        $el('btnLimpar').on('click', function () {
            _Cache.limpar();
            $el('filtroFilial').val('');
            $el('filtroRazao').val('');
            $el('filtroNotaFiscal').val('');
            $el('filtroSerie').val('');
            $el('filtroClassificadas').val('N');
            $el('searchBox').val('');
            dadosCompletos = []; dadosFiltrados = [];
            $el('tabelaBody').empty();
            $el('tableWrapper').addClass('d-none');
            $el('tableFooter').addClass('d-none');
            $el('tablePlaceholder').removeClass('d-none').find('p').html('Carregando notas não classificadas...');
            $el('paginacao').empty();
            $el('infoRegistros').text('');
            _carregar();
        });

        $el('btnImportarFornecedor').on('click', function () {
            _Cache.marcarSaida('ImportacaoFornecedores.asp');
            location.href = 'ImportacaoFornecedores.asp';
        });
        $el('btnImportarNotas').on('click', function () {
            _Cache.marcarSaida('ImportacaoNotasFiscais.asp');
            location.href = 'ImportacaoNotasFiscais.asp';
        });

        $el('tabelaBody').on('click', '.btn-classificar', function () {
            const id      = $(this).data('id');
            const filial  = $(this).data('filial');
            const classif = $el('filtroClassificadas').val();
            const href    = 'ClassificaNotasinsumo.asp?Id=' + id +
                            '&tipo=1&filial=' + encodeURIComponent(filial) +
                            '&classificadas=' + encodeURIComponent(classif);
            _Cache.marcarSaida('ClassificaNotasinsumo.asp');
            location.href = href;
        });

        $el('tabelaBody').on('click', '.btn-reclassificar', function () {
            const filial = $(this).data('filial');
            _Cache.marcarSaida('ReclassificaNotasinsumo.asp');
            location.href = 'ReclassificaNotasinsumo.asp?Id=' + $(this).data('id') +
                            '&tipo=2&filial=' + encodeURIComponent(filial);
        });

        $el('tabelaBody').on('click', '.btn-excluir', function () {
            _confirmarExclusao($(this).data('id'));
        });

        $el('btnConfirmarExclusao').on('click', function () { _executarExclusao(); });

        $el('selectLimite').on('change', function () {
            limiteLinhas = parseInt($(this).val(), 10) || 50;
            totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / limiteLinhas));
            paginaAtual  = 1;
            _renderTabela();
            _renderPaginacao(paginaAtual, totalPaginas);
            _Cache.salvar();
        });

        $el('searchBox').on('input', function () {
            clearTimeout(searchDebounce);
            const termo = $(this).val();
            searchDebounce = setTimeout(function () { _filtrarMemoria(termo); _Cache.salvar(); }, 250);
        });

        $(document).on('keydown', function (e) {
            if (e.ctrlKey && e.key === 'k') { e.preventDefault(); $el('searchBox').trigger('focus'); }
        });

        $el('btnExportCsv').on('click', function () { _exportarCSV(); });
        $el('btnExportPdf').on('click', function () { _exportarPDF(); });
    }

    return Object.freeze({ init });

})();

$(function () { ListaNotasInsumo.init(); });
</script>
</body>
</html>
