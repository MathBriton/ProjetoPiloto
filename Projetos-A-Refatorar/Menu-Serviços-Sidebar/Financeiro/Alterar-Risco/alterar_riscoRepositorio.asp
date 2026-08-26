<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/corporativo/cnn2/cnn2.asp"-->
<!--#include virtual="/corporativo/funcoes/funcoes.asp"-->
<%
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Buffer      = True
Response.Expires     = 0
Session.LCID         = 1046
Response.ContentType = "application/json"

' ── Contexto de sessão ────────────────────────────────────────────────────────
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id",    Trim(Session("CODIGO"))
usuarioLogado.Add "login", Trim(Session("USR"))
usuarioLogado.Add "nome",  Trim(Session("NOME"))
usuarioLogado.Add "nivel", Trim(Session("NIVEL"))

' ── Verificação de acesso ─────────────────────────────────────────────────────
Dim sNivel : sNivel = usuarioLogado("nivel")
If sNivel = "" Then sNivel = "99"

If CInt(sNivel) > 0 Then
  Response.Write("{""sucesso"":false,""mensagem"":""Acesso negado.""}")
  Response.End()
End If

' ── Dispatcher ────────────────────────────────────────────────────────────────
Dim acao : acao = Trim(Request("acao"))

If acao = "" Then
  Response.Write("{""sucesso"":false,""mensagem"":""Acao nao informada.""}")
  Response.End()
End If

Dim resultado

Select Case LCase(acao)
  Case "buscarclientes"  : resultado = BuscarClientes()
  Case "carregarcliente" : resultado = CarregarCliente()
  Case "gravarrisco"     : resultado = GravarRisco()
  Case Else
    resultado = "{""sucesso"":false,""mensagem"":""Acao invalida: " & EscapeJSON(acao) & """}"
End Select

Response.Write(resultado)

' =============================================================================
' Helper: normaliza codigo de cliente para CHAR(6) com zeros a esquerda
' =============================================================================
Function PadCodCliente(sCod)
  If IsNumeric(sCod) Then
    PadCodCliente = Right("000000" & CStr(CLng(sCod)), 6)
  Else
    PadCodCliente = sCod
  End If
End Function

' =============================================================================
' Helper: mapeia codigo de planta para CIDADE na tabela BI_CLIENTES (SPO-BI)
' =============================================================================
Function PlantaToCidade(sPlanta)
  Select Case sPlanta
    Case "00" : PlantaToCidade = "PIRAPETINGA"
    Case "04" : PlantaToCidade = "UBERABA"
    Case "05" : PlantaToCidade = "BENTO"
    Case "06" : PlantaToCidade = "MARANGUAPE"
    Case "07" : PlantaToCidade = "SAQUAREMA"
    Case Else  : PlantaToCidade = ""
  End Select
End Function

' =============================================================================
' Helper: mapeia codigo de planta para nome legivel
' =============================================================================
Function NomePlanta(sPlanta)
  Select Case sPlanta
    Case "00" : NomePlanta = "Pirapetinga"
    Case "04" : NomePlanta = "Uberaba"
    Case "05" : NomePlanta = "Bento Goncalves"
    Case "06" : NomePlanta = "Maranguape"
    Case "07" : NomePlanta = "Saquarema"
    Case Else  : NomePlanta = "Planta " & sPlanta
  End Select
End Function

' =============================================================================
' Helper: mapeia codigo de planta para DSN do Protheus
' =============================================================================
Function PlantaToDSN(sPlanta)
  Select Case sPlanta
    Case "00" : PlantaToDSN = "PIR-DADOSADV"
    Case "04" : PlantaToDSN = "PIR-DADOSADV"
    Case "05" : PlantaToDSN = "BEN-DADOSADV"
    Case "06" : PlantaToDSN = "FOR-DADOSADV"
    Case "07" : PlantaToDSN = "SAQ-DADOSADV"
    Case Else  : PlantaToDSN = ""
  End Select
End Function

' =============================================================================
' FUNCTION: BuscarClientes  (GET)
'
' Usa a tabela centralizada BI_CLIENTES via SPO-BI, identica ao buscaBICLIENTES.asp.
'
' Bento (05) e Maranguape (06): o usuario pode buscar por CODIGO_ERP (codigo externo),
' mas COD retornado e sempre COD_CLIENTE — unico campo que existe no SA1010 (A1_COD).
' Demais plantas: busca e retorno por COD_CLIENTE.
'
' Retorna: { sucesso, mensagem, dados: [ { cod, codext, loja, nome, cnpj, municipio, estado } ] }
' =============================================================================
Function BuscarClientes()

  Dim sPlanta    : sPlanta    = Trim(Request("Planta"))
  Dim sParametro : sParametro = Trim(Request("Parametro"))

  If sPlanta = "" Or sParametro = "" Then
    BuscarClientes = "{""sucesso"":false,""mensagem"":""Parametros insuficientes.""}"
    Exit Function
  End If

  Dim sCidade : sCidade = PlantaToCidade(sPlanta)
  If sCidade = "" Then
    BuscarClientes = "{""sucesso"":false,""mensagem"":""Planta desconhecida.""}"
    Exit Function
  End If

  ' Sanitiza o parametro para uso em LIKE
  Dim sSafe  : sSafe  = Replace(Replace(sParametro, "'", "''"), "%", "[%]")
  Dim sUpper : sUpper = UCase(sSafe)

  ' Bento e Maranguape: busca tambem por CODIGO_ERP (codigo externo).
  ' COD retornado e sempre COD_CLIENTE para garantir o match com A1_COD no SA1010.
  Dim sFiltroExtra : sFiltroExtra = ""
  If sPlanta = "05" Or sPlanta = "06" Then
    sFiltroExtra = " OR (CODIGO_ERP = '" & sSafe & "')"
  End If

  Dim sSQL
  sSQL = _
    "SELECT TOP 100 " & _
      "COD_CLIENTE       AS COD, " & _
      "CODIGO_ERP        AS CODEXT, " & _
      "LOJA_CLIENTE      AS LOJA, " & _
      "NOME_CLIENTE      AS NOME, " & _
      "CNPJ_CLIENTE      AS CNPJ, " & _
      "MUNICIPIO_CLIENTE AS MUNICIPIO, " & _
      "ESTADO_CLIENTE    AS ESTADO " & _
    "FROM BI_CLIENTES " & _
    "WHERE CIDADE = '" & sCidade & "' " & _
      "AND STATUS = 'ATIVO' " & _
      "AND (" & _
        "(FANTASIA_CLIENTE LIKE '%" & sUpper & "%')" & _
        " OR (NOME_CLIENTE LIKE '%" & sUpper & "%')" & _
        " OR (COD_CLIENTE = '" & sSafe & "')" & _
        sFiltroExtra & _
      ") " & _
    "ORDER BY NOME_CLIENTE, LOJA_CLIENTE"

  Dim rs
  On Error Resume Next
  Set rs = Conecta("SPO-BI", sSQL)

  If Err.Number <> 0 Or Not IsObject(rs) Then
    Dim sErrB : sErrB = EscapeJSON(Err.Description)
    Err.Clear
    On Error GoTo 0
    BuscarClientes = "{""sucesso"":false,""mensagem"":""Erro ao conectar SPO-BI: " & sErrB & """}"
    Exit Function
  End If
  On Error GoTo 0

  Dim sJson     : sJson     = "{""sucesso"":true,""mensagem"":"""",""dados"":["
  Dim bPrimeiro : bPrimeiro = True

  Do While Not rs.EOF
    If Not bPrimeiro Then sJson = sJson & ","
    bPrimeiro = False

    sJson = sJson & "{"
    sJson = sJson & """cod"":"      & """" & EscapeJSON(Trim(rs("COD")       & "")) & ""","
    sJson = sJson & """codext"":"   & """" & EscapeJSON(Trim(rs("CODEXT")    & "")) & ""","
    sJson = sJson & """loja"":"     & """" & EscapeJSON(Trim(rs("LOJA")      & "")) & ""","
    sJson = sJson & """nome"":"     & """" & EscapeJSON(Trim(rs("NOME")      & "")) & ""","
    sJson = sJson & """cnpj"":"     & """" & EscapeJSON(Trim(rs("CNPJ")      & "")) & ""","
    sJson = sJson & """municipio"":" & """" & EscapeJSON(Trim(rs("MUNICIPIO") & "")) & ""","
    sJson = sJson & """estado"":"   & """" & EscapeJSON(Trim(rs("ESTADO")    & "")) & """"
    sJson = sJson & "}"

    rs.MoveNext()
  Loop

  rs.Close()
  Set rs = Nothing

  sJson = sJson & "]}"

  BuscarClientes = sJson

End Function

' =============================================================================
' FUNCTION: CarregarCliente  (GET)
'
' Busca nome e risco atual no SA1010 do DSN correspondente a planta.
' Todas as plantas usam A1_COD + A1_LOJA como chave — inclusive Bento e Maranguape.
' A1_FILIAL e vazio nos bancos de Bento e Maranguape, portanto sem filtro de filial.
'
' Retorna: { sucesso, mensagem, dados: { nomeCliente, riscoAtual } }
' =============================================================================
Function CarregarCliente()

  Dim sPlanta  : sPlanta  = Trim(Request("Planta"))
  Dim sCliente : sCliente = Left(Trim(Request("Cliente")), 6)
  Dim sLoja    : sLoja    = Left(Trim(Request("Loja")), 2)

  If sPlanta = "" Or sCliente = "" Then
    CarregarCliente = "{""sucesso"":false,""mensagem"":""Parametros insuficientes.""}"
    Exit Function
  End If

  Dim sDSN : sDSN = PlantaToDSN(sPlanta)
  If sDSN = "" Then
    CarregarCliente = "{""sucesso"":false,""mensagem"":""Planta desconhecida.""}"
    Exit Function
  End If

  Dim sCodParam, sSQL

  ' ── Todas as plantas: A1_COD + A1_LOJA no SA1010 ────────────────────────────
  ' Bento (05) e Maranguape (06): A1_FILIAL e vazio nesses bancos — sem filtro.
  ' O DSN dedicado ja garante o isolamento correto por planta.
  sCodParam = sCliente
  Dim sWhereLoja : sWhereLoja = ""
  If sLoja <> "" Then
    sWhereLoja = " AND RTRIM(A1_LOJA) = '" & Replace(sLoja, "'", "''") & "'"
  End If
  sSQL = _
    "SELECT TOP 1 " & _
      "RTRIM(A1_NOME)         AS NomeCliente, " & _
      "UPPER(RTRIM(A1_RISCO)) AS RiscoAtual " & _
    "FROM SA1010 " & _
    "WHERE D_E_L_E_T_ <> '*' " & _
      "AND RTRIM(A1_COD) = '" & Replace(sCodParam, "'", "''") & "'" & _
      sWhereLoja

  Dim rs
  On Error Resume Next
  Set rs = Conecta(sDSN, sSQL)

  If Err.Number <> 0 Or Not IsObject(rs) Then
    Dim sErrC : sErrC = EscapeJSON(Err.Description)
    Err.Clear
    On Error GoTo 0
    CarregarCliente = "{""sucesso"":false,""mensagem"":""Erro ao conectar " & sDSN & ": " & sErrC & """}"
    Exit Function
  End If
  On Error GoTo 0

  If rs.EOF Then
    rs.Close()
    Set rs = Nothing
    CarregarCliente = "{""sucesso"":false,""mensagem"":""Cliente nao encontrado.""}"
    Exit Function
  End If

  Dim sNome  : sNome  = EscapeJSON(Trim(rs("NomeCliente") & ""))
  Dim sRisco : sRisco = EscapeJSON(Trim(rs("RiscoAtual")  & ""))

  rs.Close()
  Set rs = Nothing

  CarregarCliente = _
    "{""sucesso"":true,""mensagem"":"""",""dados"":{" & _
      """nomeCliente"":""" & sNome  & """," & _
      """riscoAtual"":"""  & sRisco & """" & _
    "}}"

End Function

' =============================================================================
' FUNCTION: GravarRisco  (POST)
'
' Atualiza A1_RISCO no SA1010 da planta.
' Todas as plantas usam A1_COD + A1_LOJA como chave.
' A1_FILIAL e vazio nos bancos de Bento e Maranguape — sem filtro de filial.
'
' Retorna: { sucesso, mensagem }
' =============================================================================
Function GravarRisco()

  If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    GravarRisco = "{""sucesso"":false,""mensagem"":""Metodo nao permitido.""}"
    Exit Function
  End If

  Dim sPlanta        : sPlanta        = Trim(Request.Form("Planta"))
  Dim sCliente       : sCliente       = Left(Trim(Request.Form("Cliente")),       6)
  Dim sLoja          : sLoja          = Left(Trim(Request.Form("Loja")),          2)
  Dim sRisco         : sRisco         = Left(Trim(Request.Form("Risco")),         1)
  Dim sRiscoAnterior : sRiscoAnterior = Left(Trim(Request.Form("RiscoAnterior")), 10)

  If sPlanta = "" Or sCliente = "" Or sRisco = "" Then
    GravarRisco = "{""sucesso"":false,""mensagem"":""Parametros insuficientes.""}"
    Exit Function
  End If

  Dim sDSN : sDSN = PlantaToDSN(sPlanta)
  If sDSN = "" Then
    GravarRisco = "{""sucesso"":false,""mensagem"":""Planta desconhecida.""}"
    Exit Function
  End If

  ' ── Todas as plantas: A1_COD + A1_LOJA no SA1010 ────────────────────────────
  ' O DSN dedicado por planta ja garante o isolamento correto.
  Dim sCodParam : sCodParam = sCliente
  Dim sSQL
  sSQL = _
    "UPDATE SA1010 SET A1_RISCO = '" & Replace(sRisco, "'", "''") & "' " & _
    "WHERE D_E_L_E_T_ <> '*' " & _
      "AND RTRIM(A1_COD) = '" & Replace(sCodParam, "'", "''") & "'"
  If sLoja <> "" Then
    sSQL = sSQL & " AND RTRIM(A1_LOJA) = '" & Replace(sLoja, "'", "''") & "'"
  End If

  ' ── UPDATE/DML via ADODB.Connection.Execute (DML nao retorna Recordset) ────
  Dim oConn
  Set oConn = Server.CreateObject("ADODB.Connection")

  On Error Resume Next
  oConn.Open BuscaStringConexao(sDSN)
  If Err.Number <> 0 Then
    Dim sErrUpd : sErrUpd = EscapeJSON(Err.Description)
    Err.Clear : On Error GoTo 0
    Set oConn = Nothing
    GravarRisco = "{""sucesso"":false,""mensagem"":""Erro ao conectar " & sDSN & ": " & sErrUpd & """}"
    Exit Function
  End If
  On Error GoTo 0

  On Error Resume Next
  oConn.Execute(sSQL)
  If Err.Number <> 0 Then
    Dim sErrExec : sErrExec = EscapeJSON(Err.Description)
    Err.Clear : On Error GoTo 0
    oConn.Close : Set oConn = Nothing
    GravarRisco = "{""sucesso"":false,""mensagem"":""Erro no UPDATE " & sDSN & ": " & sErrExec & """}"
    Exit Function
  End If
  On Error GoTo 0


  oConn.Close
  Set oConn = Nothing
  ' ── Log de auditoria em PIR-CORPORATIVO ─────────────────────────────────────
  Dim sLogSQL
  sLogSQL = _
    "INSERT INTO AlterarRiscoLog (" & _
      "Planta, Cliente, Loja, RiscoAnterior, RiscoNovo, Usuario, Data" & _
    ") VALUES (" & _
      "'" & Replace(sPlanta,              "'", "''") & "', " & _
      "'" & Replace(sCodParam,            "'", "''") & "', " & _
      "'" & Replace(sLoja,                "'", "''") & "', " & _
      "'" & Replace(sRiscoAnterior,       "'", "''") & "', " & _
      "'" & Replace(sRisco,               "'", "''") & "', " & _
      "'" & Replace(usuarioLogado("id"),  "'", "''") & "', " & _
      "GETDATE()" & _
    ")"

  Dim rsLog
  On Error Resume Next
  Set rsLog = Conecta("PIR-CORPORATIVO", sLogSQL)
  If Err.Number = 0 And IsObject(rsLog) Then
    rsLog.Close()
    Set rsLog = Nothing
  End If
  Err.Clear
  On Error GoTo 0

  ' ── E-mail ao representante (nao-bloqueante) ─────────────────────────────────
  On Error Resume Next
  Call EnviarEmailRepresentante(sPlanta, sCodParam, sLoja, sRisco, sRiscoAnterior)
  Err.Clear
  On Error GoTo 0

  GravarRisco = "{""sucesso"":true,""mensagem"":""Risco atualizado com sucesso.""}"

End Function

' =============================================================================
' SUB: EnviarEmailRepresentante
' =============================================================================
Sub EnviarEmailRepresentante(sPlanta, sCliente, sLoja, sNovoRisco, sRiscoAnterior)

  Dim sDSN : sDSN = PlantaToDSN(sPlanta)

  Dim sFilial
  Select Case sPlanta
    Case "00" : sFilial = "01"
    Case "04" : sFilial = "04"
    Case "05" : sFilial = "05"
    Case "06" : sFilial = "06"
    Case "07" : sFilial = "07"
    Case Else  : sFilial = "01"
  End Select

  ' Todas as plantas usam A1_COD como chave de juncao
  Dim sWhereCliente
  sWhereCliente = "RTRIM(SA1.A1_COD) = '" & Replace(sCliente, "'", "''") & "'"

  Dim sWhereLoja : sWhereLoja = ""
  If sLoja <> "" Then sWhereLoja = " AND RTRIM(SA1.A1_LOJA) = '" & Replace(sLoja, "'", "''") & "'"

  Dim sSQL
  sSQL = _
    "SELECT TOP 1 " & _
      "RTRIM(SA1.A1_NOME)  AS NomeCliente, " & _
      "RTRIM(SA3.A3_EMAIL) AS EmailRepresentante, " & _
      "RTRIM(SA3.A3_NOME)  AS NomeRepresentante " & _
    "FROM SA1010 SA1 " & _
    "INNER JOIN SA3010 SA3 " & _
      "ON RTRIM(SA3.A3_COD) = RTRIM(SA1.A1_VEND) " & _
      "AND SA3.D_E_L_E_T_ <> '*' " & _
    "WHERE SA1.D_E_L_E_T_ <> '*' " & _
      "AND RTRIM(SA1.A1_FILIAL) = '" & sFilial & "' " & _
      "AND " & sWhereCliente & _
      sWhereLoja

  Dim rs
  On Error Resume Next
  Set rs = Conecta(sDSN, sSQL)
  If Err.Number <> 0 Or Not IsObject(rs) Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0

  If rs.EOF Then
    rs.Close() : Set rs = Nothing
    Exit Sub
  End If

  Dim sNomeCli  : sNomeCli  = Trim(rs("NomeCliente")        & "")
  Dim sEmailRep : sEmailRep = Trim(rs("EmailRepresentante") & "")
  Dim sNomeRep  : sNomeRep  = Trim(rs("NomeRepresentante")  & "")

  rs.Close() : Set rs = Nothing

  Dim sPara    : sPara    = sEmailRep
  Dim sAssunto : sAssunto = "SWnet - Alteracao de Risco de Cliente"

  Dim sMensagem
  sMensagem = _
    "<!doctype html><html><head><meta charset='utf-8'>" & _
    "<style>" & _
      "body{font-family:Arial,sans-serif;font-size:14px;color:#333;margin:0;padding:0}" & _
      "p{margin:0 0 10px}" & _
      ".wrap{max-width:560px;margin:24px auto;background:#fff;border:1px solid #dee2e6;border-radius:6px;overflow:hidden}" & _
      ".hdr{background:#4a4a4a;color:#fff;padding:14px 20px;font-size:15px;font-weight:700}" & _
      ".bdy{padding:20px 24px}" & _
      ".cred{background:#f5f5f5;border-left:4px solid #4a4a4a;padding:12px 16px;border-radius:4px;margin:14px 0}" & _
      ".cred p{margin:4px 0;font-size:13px}" & _
      ".lbl{font-weight:600;color:#555;display:inline-block;width:130px}" & _
      ".badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:700}" & _
      ".b-ant{background:#fff3cd;color:#856404;border:1px solid #ffc107}" & _
      ".b-novo{background:#d1e7dd;color:#0a3622;border:1px solid #198754}" & _
      ".ftr{background:#f8f9fa;border-top:1px solid #dee2e6;padding:10px 24px;font-size:11px;color:#888}" & _
    "</style></head><body>" & _
    "<div class='wrap'>" & _
      "<div class='hdr'>Alteracao de Risco &mdash; SWnet</div>" & _
      "<div class='bdy'>" & _
        "<p>Ola, <strong>" & sNomeRep & "</strong>.</p>" & _
        "<p>O risco de um cliente sob sua responsabilidade foi alterado no <strong>SWnet &mdash; Smurfit Westrock Brasil</strong>.</p>" & _
        "<div class='cred'>" & _
          "<p><span class='lbl'>Cliente:</span> " & sNomeCli & " (" & sCliente & ")</p>" & _
          "<p><span class='lbl'>Loja:</span> " & IIf(sLoja = "", "Todas", sLoja) & "</p>" & _
          "<p><span class='lbl'>Planta:</span> " & NomePlanta(sPlanta) & "</p>" & _
          "<p><span class='lbl'>Risco Anterior:</span> <span class='badge b-ant'>" & sRiscoAnterior & "</span></p>" & _
          "<p><span class='lbl'>Novo Risco:</span> <span class='badge b-novo'>" & sNovoRisco & "</span></p>" & _
          "<p><span class='lbl'>Data/Hora:</span> " & FormatDateTime(Now(), 1) & " " & FormatDateTime(Now(), 3) & "</p>" & _
          "<p><span class='lbl'>Operador:</span> " & usuarioLogado("login") & "</p>" & _
        "</div>" & _
        "<p>Em caso de duvidas, entre em contato com a equipe comercial.</p>" & _
        "<p>At.te,<br><strong>Equipe de Desenvolvimento &mdash; Smurfit Westrock Brasil</strong></p>" & _
      "</div>" & _
      "<div class='ftr'>Esta mensagem foi gerada automaticamente pelo SWnet. Nao responda este e-mail.</div>" & _
    "</div></body></html>"

  If sPara <> "" Then
    Call EnviaEmail(sPara, "", "", "", "noreply.brasil@smurfitwestrock.com.br", sMensagem, sAssunto)
  End If

End Sub

%>
