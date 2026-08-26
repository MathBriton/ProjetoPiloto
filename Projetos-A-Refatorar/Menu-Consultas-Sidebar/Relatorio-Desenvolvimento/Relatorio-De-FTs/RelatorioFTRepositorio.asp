<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp"-->
<!--#include file="../../../funcoes/funcoes.asp"-->
<!--#include file="../../../includes/variaveis.asp"-->
<%
Response.Buffer       = True
Response.Charset      = "UTF-8"
Response.CodePage     = 65001
Response.Expires      = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
Response.ContentType = "application/json"

' ============================================================
' CONSTANTES
' ============================================================

Const PAGE_SIZE_SVR_DEFAULT = 1000

Dim nivelUsr, codsiga
nivelUsr = CInt(0 & Trim(NIVEL))
codsiga  = Trim(CODUSER)

' ============================================================
' HELPERS LOCAIS
' ============================================================

' JStr: substitui EscapeJSON com escape correto para JSON.
' Necessário porque EscapeJSON em funcoes.asp opera sobre
' variável local errada e não escapa aspas/barras invertidas.
Function JStr(str)
  If IsNull(str) Then str = ""
  Dim result, i, c, asc_c
  result = CStr(str)
  result = Replace(result, "\",  "\\")
  result = Replace(result, """", "\""")
  result = Replace(result, vbCrLf, "\n")
  result = Replace(result, vbCr,   "\n")
  result = Replace(result, vbLf,   "\n")
  result = Replace(result, vbTab,  "\t")
  Dim out : out = ""
  For i = 1 To Len(result)
    c     = Mid(result, i, 1)
    asc_c = Asc(c)
    If asc_c >= 0 And asc_c < 32 Then
      out = out & "\u00" & Right("0" & Hex(asc_c), 2)
    Else
      out = out & c
    End If
  Next
  JStr = out
End Function

' JNum: serializa número para JSON (vírgula → ponto, NULL → "0")
Function JNum(v)
  If IsNull(v) Then
    JNum = "0"
  ElseIf Trim(CStr(v)) = "" Then
    JNum = "0"
  Else
    JNum = Replace(CStr(v), ",", ".")
  End If
End Function

Function JBool(v)
  If v Then JBool = "true" Else JBool = "false"
End Function

Function FmtData(s)
  If IsNull(s) Then
    FmtData = ""
    Exit Function
  End If
  s = CStr(s)
  If Len(s) >= 8 Then
    FmtData = Mid(s,7,2) & "/" & Mid(s,5,2) & "/" & Mid(s,1,4)
  Else
    FmtData = s
  End If
End Function

' ============================================================
' HELPERS DE IMAGEM FT
' ============================================================

' ObterConexao: retorna ADODB.Connection aberta para o DSN informado.
' Usa BuscaStringConexao() — mesmo padrão do Conecta() em cnn2.asp
' (Provider=...;Data Source=...;Initial Catalog=...;User ID=...;Password=...).
' Necessário para associar a ADODB.Command com CommandTimeout explícito,
' contornando o timeout padrão de 300s do Conecta() em queries que precisam
' de controle granular por operação.
Function ObterConexao(dsnName)
  Dim conn, connStr
  connStr = BuscaStringConexao(UCase(dsnName))
  Set conn = Server.CreateObject("ADODB.Connection")
  conn.ConnectionTimeout = 15
  conn.CommandTimeout    = 120
  conn.Open connStr
  Set ObterConexao = conn
End Function

Function ImgFTExiste(codigo)
  Dim objFSO
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  ImgFTExiste = objFSO.FileExists("c:\Inetpub\wwwroot\ft\" & codigo & ".jpg") _
             Or objFSO.FileExists("c:\Inetpub\wwwroot\ft\" & codigo & ".png")
  Set objFSO = Nothing
End Function

' ============================================================
' ROTEAMENTO
' ============================================================

Dim acao
acao = Trim(Request("acao"))

If acao = "" Then
  Response.Write "{""sucesso"":false,""mensagem"":""Ação não informada""}"
  Response.End
End If

Dim resultado
Select Case acao
  Case "obterCombos"    : resultado = ObterCombos()
  Case "validarCliente" : resultado = ValidarCliente()
  Case "obterFTs"       : resultado = ObterFTs()
  Case Else
    resultado = "{""sucesso"":false,""mensagem"":""Ação inválida: " & JStr(acao) & """}"
End Select

Response.Clear
Response.Write resultado
Response.End

' ============================================================
' AÇÃO: obterCombos
' Retorna opções de Grupo e Tipo (base BI)
' ============================================================

Function ObterCombos()
  Dim rs, json, sep

  json = "{""sucesso"":true,""grupos"":["
  sep  = ""

  On Error Resume Next
  Set rs = Conecta("SPO-BI", _
    "SELECT GRUPO_PRODUTO FROM BI_PRODUTOS " & _
    "GROUP BY GRUPO_PRODUTO ORDER BY GRUPO_PRODUTO")
  If Err.Number <> 0 Then
    ObterCombos = "{""sucesso"":false,""mensagem"":""Erro ao carregar grupos""}"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsObject(rs) And Not (rs Is Nothing) Then
    Do While Not rs.EOF
      json = json & sep & """" & JStr(Trim(rs("GRUPO_PRODUTO"))) & """"
      sep  = ","
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
  End If

  json = json & "],""tipos"":["
  sep  = ""

  On Error Resume Next
  Set rs = Conecta("SPO-BI", _
    "SELECT TIPO_PRODUTO FROM BI_PRODUTOS " & _
    "WHERE TIPO_PRODUTO IS NOT NULL " & _
    "GROUP BY TIPO_PRODUTO ORDER BY TIPO_PRODUTO")
  If Err.Number <> 0 Then
    ObterCombos = "{""sucesso"":false,""mensagem"":""Erro ao carregar tipos""}"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsObject(rs) And Not (rs Is Nothing) Then
    Do While Not rs.EOF
      json = json & sep & """" & JStr(Trim(rs("TIPO_PRODUTO"))) & """"
      sep  = ","
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
  End If

  json = json & "]}"
  ObterCombos = json
End Function

' ============================================================
' AÇÃO: validarCliente
' Valida se o representante (nível 1) tem permissão sobre o cliente
' ============================================================

Function ValidarCliente()
  Dim pTermoCli, SQL, rs, temPermissao
  pTermoCli    = Trim(Request("termoCli"))
  temPermissao = True

  ' Validação de representante só faz sentido com código numérico exato
  If nivelUsr = 1 And pTermoCli <> "" And IsNumeric(pTermoCli) Then
    SQL = "EXEC SP_VALIDACLIENTE2 [" & nivelUsr & "],[" & codsiga & "],[" & pTermoCli & "]"

    On Error Resume Next
    Set rs = Conecta("PIR-DADOSADV", SQL)
    If Err.Number <> 0 Or Not IsObject(rs) Or (rs Is Nothing) Or rs.EOF Then
      temPermissao = False
    End If
    On Error GoTo 0

    If IsObject(rs) And Not (rs Is Nothing) Then
      rs.Close
      Set rs = Nothing
    End If
  End If

  ValidarCliente = "{""sucesso"":true,""permitido"":" & JBool(temPermissao) & "}"
End Function

' ============================================================
' AÇÃO: obterFTs
' ============================================================

Function ObterFTs()
  Dim pPlanta, pStatus, pGrupo, pTipo, pTermoCli, pLojaCli, nPagCan, nPagSize
  Dim bIsPirUbe, bIsBenForSaq, bIsCancelado

  pPlanta    = Trim(Request("cboPlanta"))
  pStatus    = Trim(Request("cboStatus"))
  pGrupo     = Trim(Request("cboGrupo"))
  pTipo      = Trim(Request("cboTipo"))
  pTermoCli  = Trim(Request("termoCli"))
  pLojaCli   = Trim(Request("lojacli"))
  nPagCan    = CLng(0 & Trim(Request("pagCan")))
  nPagSize   = CLng(0 & Trim(Request("pagSize")))
  If nPagCan  < 1     Then nPagCan  = 1
  If nPagSize < 10    Then nPagSize = PAGE_SIZE_SVR_DEFAULT
  If nPagSize > 99999 Then nPagSize = 99999

  If pPlanta = "" Then
    ObterFTs = "{""sucesso"":false,""mensagem"":""Filial não informada""}"
    Exit Function
  End If

  bIsPirUbe    = (pPlanta = "00" Or pPlanta = "04")
  bIsBenForSaq = (pPlanta = "05" Or pPlanta = "06" Or pPlanta = "07")
  bIsCancelado = (bIsPirUbe And pStatus = "C")

  If bIsPirUbe Then
    ObterFTs = MontarJsonPirUbe(pPlanta, pStatus, pTermoCli, pLojaCli, nPagCan, bIsCancelado, nPagSize)
  ElseIf bIsBenForSaq Then
    ObterFTs = MontarJsonTrimbox(pPlanta, pStatus, pGrupo, pTipo, pTermoCli, pLojaCli, nPagCan, nPagSize)
  Else
    ObterFTs = "{""sucesso"":false,""mensagem"":""Filial inválida""}"
  End If
End Function

' ============================================================
' PIR / UBE — Pirapetinga (00) e Uberaba (04)
' ============================================================

Function MontarJsonPirUbe(pPlanta, pStatus, pTermoCli, pLojaCli, nPagCan, bIsCancelado, nPagSize)
  Dim whereNivel, whereCliPIR, whereInpube, whereStatusPIR
  Dim SQLFT, rs, json
  Dim nTotalCan, nOffset, bHasMoreCan, nMaxPagCan
  Dim sCodCli, sNomeCli
  Dim bVerNotCPE, bVerPE
  Dim jsonSecLiberadas, jsonSecPE

  nTotalCan   = 0
  nOffset     = (nPagCan - 1) * nPagSize
  bHasMoreCan = False
  nMaxPagCan  = 1
  sCodCli     = ""
  sNomeCli    = ""
  bVerNotCPE  = False
  bVerPE      = False

  ' IIf não é usado com Recordsets abertos — apenas para montar strings WHERE
  If nivelUsr = 1 Then
    whereNivel = "AND SB5010.B5_VEND = '" & codsiga & "' "
  Else
    whereNivel = ""
  End If

  ' Filtro de cliente: código exato (numérico OU alfanumérico sem espaço, ex: ZB3425)
  ' OU LIKE na fantasia quando o termo contém espaço (nome/razão social)
  If pTermoCli <> "" Then
    Dim pTermoEscPIR
    pTermoEscPIR = Replace(pTermoCli, "'", "''")
    If InStr(pTermoCli, " ") = 0 Then
      whereCliPIR = "AND RTRIM(SB5010.B5_CLI) = '" & pTermoEscPIR & "' "
      If pLojaCli <> "" Then
        whereCliPIR = whereCliPIR & "AND RTRIM(SB5010.B5_LOJA) = '" & pLojaCli & "' "
      End If
    Else
      whereCliPIR = "AND UPPER(SB5010.B5_NOMECLI) LIKE '%" & UCase(pTermoEscPIR) & "%' "
    End If
  Else
    whereCliPIR = ""
  End If

  If pPlanta = "00" Then
    whereInpube = "'I'"
  Else
    whereInpube = "'U'"
  End If

  If pStatus <> "" Then
    whereStatusPIR = "AND SB1010.B1_STATUS = '" & pStatus & "' "
  Else
    whereStatusPIR = ""
  End If

  ' Determina cláusula de status para COUNT e consulta principal
  Dim whereFiltroPIR
  If bIsCancelado Then
    whereFiltroPIR = "AND SB1010.B1_STATUS = 'C' "
  ElseIf whereStatusPIR <> "" Then
    whereFiltroPIR = whereStatusPIR
  Else
    whereFiltroPIR = "AND SB1010.B1_STATUS NOT IN ('P','E','C') "
  End If

  ' ── Contagem server-side ──────────────────────────────────
  Dim SQLContaPIR, rsContaPIR, nTotalPIR
  nTotalPIR = 0

  SQLContaPIR = _
    "SELECT COUNT(*) AS TOTAL " & _
    "FROM SB1010 with(nolock) " & _
      "INNER JOIN SB5010 with(nolock) ON SB1010.B1_COD = SB5010.B5_COD " & _
    "WHERE 1=1 " & _
      whereNivel & whereCliPIR & _
      "AND SB1010.B1_INPUBE = " & whereInpube & " " & _
      "AND SB1010.D_E_L_E_T_ <> '*' " & _
      "AND SB5010.D_E_L_E_T_ <> '*' " & _
      whereFiltroPIR & _
      "AND SUBSTRING(SB1010.B1_COD,1,2) <> '77' "

  On Error Resume Next
  Set rsContaPIR = Conecta("PIR-DADOSADV", SQLContaPIR)
  If Err.Number <> 0 Then
    MontarJsonPirUbe = "{""sucesso"":false,""mensagem"":""Erro ao contar registros""}"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsObject(rsContaPIR) And Not (rsContaPIR Is Nothing) Then
    nTotalPIR = CLng(0 & rsContaPIR("TOTAL"))
    rsContaPIR.Close
    Set rsContaPIR = Nothing
  End If

  nTotalCan  = nTotalPIR
  If nTotalCan > (nPagCan * nPagSize) Then bHasMoreCan = True
  nMaxPagCan = CLng(((nTotalCan - 1) \ nPagSize) + 1)
  If nMaxPagCan < 1 Then nMaxPagCan = 1

  ' ── Consulta principal paginada via ROW_NUMBER ────────────
  SQLFT = _
    "SELECT * FROM (" & _
      "SELECT " & _
        "SB5010.B5_CLI, SB5010.B5_LOJA, SB5010.B5_NOMECLI, " & _
        "SB5010.B5_VEND, SB5010.B5_NOMVEN, SB1010.B1_COD, SB1010.B1_DESC, " & _
        "SB1010.B1_STATUS, SB1010.B1_TS, SB1010.B1_UPRC, SB1010.B1_UCOM, " & _
        "SB5010.B5_COMPR, SB5010.B5_LARG, SB5010.B5_ALTU, " & _
        "(SB5010.B5_AREA * B5_PCSCON) AS B5_AREA, " & _
        "SB5010.B5_ESTRUTU, SB5010.B5_COR1, SB5010.B5_COR2, " & _
        "SB5010.B5_COR3, SB5010.B5_COR4, SB5010.B5_SAP, SB1010.B1_INPUBE, " & _
        "ROW_NUMBER() OVER (ORDER BY SB1010.B1_COD) AS RN " & _
      "FROM SB1010 with(nolock) " & _
        "INNER JOIN SB5010 with(nolock) ON SB1010.B1_COD = SB5010.B5_COD " & _
      "WHERE 1=1 " & _
        whereNivel & whereCliPIR & _
        "AND SB1010.B1_INPUBE = " & whereInpube & " " & _
        "AND SB1010.D_E_L_E_T_ <> '*' " & _
        "AND SB5010.D_E_L_E_T_ <> '*' " & _
        whereFiltroPIR & _
        "AND SUBSTRING(SB1010.B1_COD,1,2) <> '77' " & _
    ") X " & _
    "WHERE RN > " & nOffset & " AND RN <= " & (nOffset + nPagSize) & " " & _
    "ORDER BY RN"

  ' ── Executa consulta principal ────────────────────────────
  Dim secLibRegs, sepReg, titSecLib, sepSec
  Dim secPERegsInline, separadorPEMostrado

  secLibRegs          = ""
  titSecLib           = ""
  secPERegsInline     = ""
  separadorPEMostrado = False

  If bIsCancelado Then
    titSecLib = "FTs Canceladas"
  Else
    titSecLib = "FTs Liberadas DDP"
  End If

  On Error Resume Next
  Set rs = Conecta("PIR-DADOSADV", SQLFT)
  If Err.Number <> 0 Then
    MontarJsonPirUbe = "{""sucesso"":false,""mensagem"":""Erro ao consultar FTs""}"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsObject(rs) And Not (rs Is Nothing) Then
    Do While Not rs.EOF
      Dim cod, regJson, fab, ucom, inpube, status
      cod    = Trim(rs("B1_COD"))
      inpube = Trim(rs("B1_INPUBE"))
      status = Trim(rs("B1_STATUS"))

      If inpube = "U" Then
        fab = "UBERABA"
      Else
        fab = "MATRIZ"
      End If

      ucom    = FmtData(rs("B1_UCOM"))
      regJson = MontarRegPirUbe(rs, cod, fab, ucom, inpube)

      ' Quando não está em modo cancelado e aparece status "E",
      ' quebra a partir dali para a seção PE (mantém comportamento legado)
      If Not bIsCancelado And Not separadorPEMostrado And status = "E" Then
        separadorPEMostrado = True
      End If

      If separadorPEMostrado Then
        secPERegsInline = secPERegsInline & SepInline(secPERegsInline) & regJson
      Else
        secLibRegs = secLibRegs & SepInline(secLibRegs) & regJson
      End If

      rs.MoveNext
    Loop
    bVerNotCPE = (secLibRegs <> "" Or secPERegsInline <> "")
    rs.Close
    Set rs = Nothing
  End If

  ' ── Nome do cliente (cabeçalho) ───────────────────────────
  ' Resolve se o termo digitado é um código (sem espaço) ou busca por nome (com espaço)
  Dim pCodCliExato
  pCodCliExato = ""
  If pTermoCli <> "" Then
    If InStr(pTermoCli, " ") = 0 Then
      pCodCliExato = pTermoCli
    End If
  End If

  If pTermoCli <> "" And bVerNotCPE Then
    Dim SQLCli, rsCli
    If pCodCliExato <> "" Then
      SQLCli = "SELECT A1_NOME, A1_COD FROM SA1010 WHERE RTRIM(A1_COD) = '" & pCodCliExato & "' "
    Else
      SQLCli = "SELECT TOP 1 A1_NOME, A1_COD FROM SA1010 " & _
               "WHERE UPPER(A1_NOME) LIKE '%" & UCase(Replace(pTermoCli, "'", "''")) & "%' "
    End If

    On Error Resume Next
    Set rsCli = Conecta("PIR-DADOSADV", SQLCli)
    On Error GoTo 0

    If IsObject(rsCli) And Not (rsCli Is Nothing) And Not rsCli.EOF Then
      sCodCli  = Trim(rsCli("A1_COD"))
      sNomeCli = Trim(rsCli("A1_NOME"))
    End If
    If IsObject(rsCli) And Not (rsCli Is Nothing) Then
      rsCli.Close
      Set rsCli = Nothing
    End If
  End If

  ' ── FTs em Projeto ou Espera (SP_CONSULTAFT3) ─────────────
  ' Só consultada quando há código exato, não está em modo
  ' cancelado e não há filtro de status específico
  Dim secPERegs
  secPERegs = ""

  If pCodCliExato <> "" And Not bIsCancelado And pStatus = "" Then
    Dim SQLFT3, rs3
    SQLFT3 = "EXECUTE SP_CONSULTAFT3 [" & pCodCliExato & "]"

    On Error Resume Next
    Set rs3 = Conecta("PIR-DADOSADV", SQLFT3)
    On Error GoTo 0

    If IsObject(rs3) And Not (rs3 Is Nothing) Then
      Do While Not rs3.EOF
        Dim cod3, fab3, ucom3, inpube3, reg3, loja3
        cod3    = Trim(rs3("B1_COD"))
        inpube3 = Trim(rs3("B1_INPUBE"))
        loja3   = Trim(rs3("B5_LOJA"))

        ' Se foi informada uma loja específica, descarta registros de outras lojas
        If pLojaCli = "" Or loja3 = pLojaCli Then
          If inpube3 = "U" Then
            fab3 = "UBERABA"
          Else
            fab3 = "MATRIZ"
          End If

          ucom3     = FmtData(rs3("B1_UCOM"))
          reg3      = MontarRegPirUbe(rs3, cod3, fab3, ucom3, inpube3)
          secPERegs = secPERegs & SepInline(secPERegs) & reg3
          bVerPE    = True
        End If

        rs3.MoveNext
      Loop
      rs3.Close
      Set rs3 = Nothing
    End If
  End If

  ' ── Consolida registros PE (inline + vindos da SP) ────────
  Dim secPEConsolidado
  secPEConsolidado = secPERegsInline
  If secPERegs <> "" Then
    secPEConsolidado = secPEConsolidado & SepInline(secPEConsolidado) & secPERegs
  End If

  ' ── Monta array de seções ─────────────────────────────────
  Dim secoesJson
  secoesJson = "["
  sepSec     = ""

  If secLibRegs <> "" Then
    secoesJson = secoesJson & sepSec & _
      "{""titulo"":""" & JStr(titSecLib) & """,""registros"":[" & secLibRegs & "]}"
    sepSec = ","
  End If

  If secPEConsolidado <> "" Then
    secoesJson = secoesJson & sepSec & _
      "{""titulo"":""FTs em Projeto ou Espera – Intervenção DDP"",""registros"":[" & secPEConsolidado & "]}"
  End If

  secoesJson = secoesJson & "]"

  ' ── Monta envelope final ──────────────────────────────────
  Dim json2
  json2 = "{""sucesso"":true," & _
    """tipo"":""PIRUBE""," & _
    """nivel"":"            & nivelUsr            & "," & _
    """filial"":"""         & JStr(pPlanta)  & """," & _
    """cliente"":{""codigo"":""" & JStr(sCodCli)  & """,""nome"":""" & JStr(sNomeCli) & """}," & _
    """paginacao"":{" & _
      """pagina"":"     & nPagCan           & "," & _
      """totalCan"":"   & nTotalCan         & "," & _
      """maxPag"":"     & nMaxPagCan        & "," & _
      """hasMore"":"    & JBool(bHasMoreCan) & "," & _
      """pageSize"":"   & nPagSize     & _
    "}," & _
    """cancelado"":"   & JBool(bIsCancelado) & "," & _
    """secoes"":"      & secoesJson          & _
    "}"

  MontarJsonPirUbe = json2
End Function

' ── Monta 1 registro JSON da consulta SB1010/SB5010 ou SP_CONSULTAFT3 ──
Function MontarRegPirUbe(rs, cod, fab, ucom, inpube)
  Dim imgExiste, urlImg, filialFolha

  imgExiste = ImgFTExiste(cod)
  urlImg    = "getImgFT.asp?cod=" & cod

  If inpube = "I" Then
    filialFolha = "00"
  Else
    filialFolha = "04"
  End If

  MontarRegPirUbe = "{" & _
    """codigo"":"""      & JStr(cod)                        & """," & _
    """imgExiste"":"     & JBool(imgExiste)                       & "," & _
    """urlImg"":"""      & JStr(urlImg)                     & """," & _
    """filial"":"""      & JStr(filialFolha)                & """," & _
    """descricao"":"""   & JStr(rs("B1_DESC"))              & """," & _
    """fabricacao"":""" & JStr(fab)                        & """," & _
    """preco"":"         & JNum(rs("B1_UPRC"))                    & "," & _
    """entrega"":"""     & JStr(ucom)                       & """," & _
    """status"":"""      & JStr(rs("B1_STATUS"))            & """," & _
    """compr"":"""       & JStr(rs("B5_COMPR"))             & """," & _
    """larg"":"""        & JStr(rs("B5_LARG"))              & """," & _
    """altu"":"""        & JStr(rs("B5_ALTU"))              & """," & _
    """area"":"          & JNum(rs("B5_AREA"))                    & "," & _
    """estrutura"":""" & JStr(rs("B5_ESTRUTU"))           & """," & _
    """cor1"":"""        & JStr(rs("B5_COR1"))              & """," & _
    """cor2"":"""        & JStr(rs("B5_COR2"))              & """," & _
    """cor3"":"""        & JStr(rs("B5_COR3"))              & """," & _
    """cor4"":"""        & JStr(rs("B5_COR4"))              & """," & _
    """sap"":"""         & JStr(rs("B5_SAP"))               & """," & _
    """codcli"":"""      & JStr(rs("B5_CLI"))               & """," & _
    """lojacli"":"""     & JStr(rs("B5_LOJA"))              & """," & _
    """fantasia"":""" & JStr(rs("B5_NOMECLI"))           & """" & _
    "}"
End Function

' Retorna vírgula se o buffer já tem conteúdo, senão string vazia
Function SepInline(buffer)
  If buffer = "" Then SepInline = "" Else SepInline = ","
End Function

' ============================================================
' TRIMBOX — Bento (05), Fortaleza/Maranguape (06), Saquarema (07)
'
' OTIMIZAÇÕES APLICADAS:
'   1. COLLATE movido para o lado menor dos JOINs (FT1/SA1010),
'      preservando uso de índice na coluna da view (CODIGO_ITEM,
'      CODIGO_CLIENTE, LOJA_CLIENTE).
'   2. CTE BASE pré-filtra a view antes dos JOINs cross-database,
'      reduzindo drasticamente o conjunto de linhas enviado aos
'      linked servers CORPORATIVO e DADOSADV.
'   3. COUNT separado (apenas na página 1) + OFFSET/FETCH na query de
'      dados. A CTE PAGED com ROW_NUMBER() OVER (ORDER BY CODIGO_ITEM)
'      forçava materialização total de ~56k registros antes de retornar
'      qualquer linha, causando timeout sem filtro de cliente. OFFSET/FETCH
'      permite early-exit após nPagSize rows.
'   4. ImgFTExiste removido do loop (N chamadas ao filesystem);
'      imgExiste retorna false e o frontend faz lazy-load via 404.
' ============================================================

Function MontarJsonTrimbox(pPlanta, pStatus, pGrupo, pTipo, pTermoCli, pLojaCli, nPagCan, nPagSize)
  Dim colFT1, whereCidade, whereCliTri, whereStatusTri, whereGrupoTri, whereTipoTri
  Dim SQLTRI, rs, regs, dsn
  Dim nTotalTri, nOffsetTri, bHasMoreTri, nMaxPagTri
  Dim whereNivelTri
  Dim sCliCodTri, sCliLojaTri, sCliNomeTri
  Dim bNeedSA1010, joinSA1010, colFantasia

  nTotalTri   = 0
  nOffsetTri  = (nPagCan - 1) * nPagSize
  bHasMoreTri = False
  nMaxPagTri  = 1

  colFT1         = ""
  whereCidade    = ""
  whereCliTri    = ""
  whereStatusTri = ""
  whereGrupoTri  = ""
  whereTipoTri   = ""
  whereNivelTri  = ""

  sCliCodTri  = ""
  sCliLojaTri = ""
  sCliNomeTri = ""

  ' ── Coluna de join com SKFT1 varia por planta ─────────────
  ' COLLATE aplicado no lado FT1 (tabela menor / linked server),
  ' preservando o índice sobre CODIGO_ITEM na view.
  Select Case pPlanta
    Case "05" : colFT1 = "FT1_FTNOVABEN"
    Case "06" : colFT1 = "FT1_FTNOVAFOR"
    Case "07" : colFT1 = "FT1_FTNOVASAQ"
  End Select

  ' ── Filtro de cidade ──────────────────────────────────────
  If fNomePlanta(pPlanta) = "FORTALEZA" Then
    whereCidade = "AND CIDADE IN ('FORTALEZA', 'MARANGUAPE') "
  Else
    whereCidade = "AND CIDADE = '" & fNomePlanta(pPlanta) & "' "
  End If

  ' ── Filtro de representante (nível 1) ─────────────────────
  ' Aplicado na CTE BASE (antes dos JOINs), sobre A1_VEND da SA1010.
  ' Nota: este filtro ainda exige o JOIN com SA1010 mesmo na base;
  ' se A1_VEND existir na view, mova o filtro para evitar o JOIN antecipado.
  If nivelUsr = 1 Then
    whereNivelTri = "AND D.A1_VEND = '" & Trim(Session("CODUSER")) & "' "
  End If

  ' ── Demais filtros ────────────────────────────────────────
  If pStatus <> "" Then whereStatusTri = "AND STATUS = '"         & pStatus & "' "
  If pGrupo  <> "" Then whereGrupoTri  = "AND GRUPO_PRODUTO = '"  & pGrupo  & "' "
  If pTipo   <> "" Then whereTipoTri   = "AND TIPO_PRODUTO = '"   & pTipo   & "' "

  If pTermoCli <> "" Then
    Dim pTermoEscTri
    pTermoEscTri = Replace(pTermoCli, "'", "''")
    If InStr(pTermoCli, " ") = 0 Then
      ' Código exato (numérico OU alfanumérico sem espaço, ex: ZB3425)
      whereCliTri = "AND RTRIM(CODIGO_CLIENTE) = '" & pTermoEscTri & "' "
      If pLojaCli <> "" Then
        whereCliTri = whereCliTri & "AND LOJA_CLIENTE = '" & pLojaCli & "' "
      End If
    Else
      ' Busca por fantasia (A1_NREDUZ) via LIKE — requer JOIN com SA1010
      whereCliTri = "AND UPPER(D.A1_NREDUZ) LIKE '%" & UCase(pTermoEscTri) & "%' "
    End If
  End If

  Select Case pPlanta
    Case "05" : dsn = "BEN-TRIMBOX"
    Case "06" : dsn = "FOR-TRIMBOX"
    Case "07" : dsn = "SAQ-TRIMBOX"
  End Select

  ' ── Decide quais JOINs cross-linked-server são necessários ───────────────
  '
  ' SKFT1 (CORPORATIVO): fornece COD_SAP. Só é útil quando o usuário filtrou
  '   por cliente específico — na listagem geral o campo aparece vazio e o
  '   custo do join sobre toda a tabela remota causa timeout em Bento.
  '   Critério: qualquer filtro de cliente informado (código ou fantasia).
  '
  ' SA1010 (DADOSADV): fornece FANTASIA_CLIENTE e A1_VEND (filtro nível 1).
  '   Critério: representante (nível 1) OU busca por fantasia (termo com espaço).
  '
  ' Quando ambos estão ausentes a CTE JOINED é eliminada e BASE vai direto
  ' para DEDUP, tornando a query 100% local — sem nenhum roundtrip remoto.

  bNeedSA1010 = (nivelUsr = 1) Or (pTermoCli <> "" And InStr(pTermoCli, " ") > 0)
  Dim bNeedSKFT1, joinSKFT1, colSAP
  bNeedSKFT1 = (pTermoCli <> "")

  If bNeedSKFT1 Then
    joinSKFT1 = _
      "LEFT JOIN CORPORATIVO..SKFT1 FT1 " & _
        "ON B.CODIGO_ITEM = FT1." & colFT1 & " COLLATE SQL_Latin1_General_CP1_CI_AS "
    colSAP = "FT1.FT1_ID AS COD_SAP, "
  Else
    joinSKFT1 = ""
    colSAP    = "'' AS COD_SAP, "
  End If

  If bNeedSA1010 Then
    joinSA1010 = _
      "LEFT JOIN DADOSADV..SA1010 D " & _
        "ON  B.CODIGO_CLIENTE = D.A1_COD  COLLATE SQL_Latin1_General_CP1_CI_AS " & _
        "AND B.LOJA_CLIENTE   = D.A1_LOJA COLLATE SQL_Latin1_General_CP1_CI_AS " & _
        "AND D.D_E_L_E_T_ <> '*' "
    colFantasia = "D.A1_NREDUZ AS FANTASIA_CLIENTE "
  Else
    joinSA1010  = ""
    colFantasia = "'' AS FANTASIA_CLIENTE "
  End If

  ' ── Resolve fantasia do cliente via BI_CLIENTES (SPO-BI) ──────────────────
  ' SA1010 nao existe no DSN TRIMBOX. A fonte correta e a tabela BI_CLIENTES
  ' no SPO-BI, que e a mesma usada pelo popup buscaBICLIENTES.asp.
  ' Para Bento (05) e Maranguape (06) o codigo do cliente fica em CODIGO_ERP;
  ' para Saquarema (07) fica em COD_CLIENTE — mesma logica do popup.
  Dim fCidadeTri
  Select Case pPlanta
    Case "05" : fCidadeTri = "BENTO"
    Case "06" : fCidadeTri = "MARANGUAPE"
    Case "07" : fCidadeTri = "SAQUAREMA"
  End Select

  If pTermoCli <> "" Then
    Dim SQLCliTri, rsCliTri, pTermoEscCli, colCodigoTri
    pTermoEscCli = Replace(pTermoCli, "'", "''")

    ' Bento e Maranguape: codigo ERP; Saquarema: COD_CLIENTE padrao
    If pPlanta = "05" Or pPlanta = "06" Then
      colCodigoTri = "CODIGO_ERP"
    Else
      colCodigoTri = "COD_CLIENTE"
    End If

    If InStr(pTermoCli, " ") = 0 Then
      ' Codigo exato
      SQLCliTri = "SELECT TOP 1 CODIGO_ERP, COD_CLIENTE, LOJA_CLIENTE, FANTASIA_CLIENTE " & _
                  "FROM BI_CLIENTES " & _
                  "WHERE " & colCodigoTri & " = '" & pTermoEscCli & "' " & _
                  "AND CIDADE = '" & fCidadeTri & "' "
      If pLojaCli <> "" Then
        SQLCliTri = SQLCliTri & "AND LOJA_CLIENTE = '" & Replace(pLojaCli, "'", "''") & "' "
      End If
    Else
      ' Busca por fantasia/nome
      SQLCliTri = "SELECT TOP 1 CODIGO_ERP, COD_CLIENTE, LOJA_CLIENTE, FANTASIA_CLIENTE " & _
                  "FROM BI_CLIENTES " & _
                  "WHERE (UPPER(FANTASIA_CLIENTE) LIKE '%" & UCase(pTermoEscCli) & "%' " & _
                  "OR UPPER(NOME_CLIENTE) LIKE '%" & UCase(pTermoEscCli) & "%') " & _
                  "AND CIDADE = '" & fCidadeTri & "' "
    End If

    On Error Resume Next
    Set rsCliTri = Conecta("SPO-BI", SQLCliTri)
    On Error GoTo 0
    If IsObject(rsCliTri) And Not (rsCliTri Is Nothing) And Not rsCliTri.EOF Then
      ' Codigo exposto: CODIGO_ERP para BEN/MAR, COD_CLIENTE para SAQ
      If pPlanta = "05" Or pPlanta = "06" Then
        sCliCodTri = Trim(rsCliTri("CODIGO_ERP"))
      Else
        sCliCodTri = Trim(rsCliTri("COD_CLIENTE"))
      End If
      sCliLojaTri = Trim(rsCliTri("LOJA_CLIENTE"))
      sCliNomeTri = Trim(rsCliTri("FANTASIA_CLIENTE"))
    End If
    If IsObject(rsCliTri) And Not (rsCliTri Is Nothing) Then
      rsCliTri.Close
      Set rsCliTri = Nothing
    End If
  End If

  ' ── Duas queries independentes: COUNT leve + página de dados ─────────────
  '
  ' PROBLEMA ANTERIOR: COUNT(*) OVER() numa CTE com 56k registros força o
  ' SQL Server a materializar TODOS os rows antes de retornar qualquer linha,
  ' mesmo com WHERE RN <= 100 na camada exterior. Isso causava o timeout.
  '
  ' SOLUÇÃO: query de COUNT separada, executada apenas na página 1 (ou quando
  ' o frontend pedir explicitamente). A query principal usa TOP + ROW_NUMBER
  ' para retornar apenas os N registros da página sem tocar no restante.
  '
  ' Query de COUNT: filtra apenas a view local (sem linked servers) com
  ' DISTINCT sobre CODIGO_ITEM — replica a lógica de dedup sem window function.
  '
  ' Query de dados: CTE BASE → (JOINED opcional) → paginação via ROW_NUMBER.
  ' Sem COUNT(*) OVER(), o otimizador pode fazer early-exit após N rows.

  ' ── 1. COUNT: somente na primeira página ──────────────────────────────────
  ' Nas páginas seguintes o frontend já conhece o total — não reexecutamos.
  If nPagCan = 1 Then
    Dim SQLCount, rsCount, cmdCount
    Dim whereCliCnt

    ' Reescreve filtro de cliente usando colunas da view (sem alias D.)
    ' para que o COUNT fique 100% local (sem linked server).
    whereCliCnt = ""
    If whereCliTri <> "" And Not bNeedSA1010 Then
      whereCliCnt = Replace(whereCliTri, "RTRIM(CODIGO_CLIENTE)", "RTRIM(P.CODIGO_CLIENTE)")
      whereCliCnt = Replace(whereCliCnt, "LOJA_CLIENTE", "P.LOJA_CLIENTE")
    End If

    ' COUNT usa SELECT DISTINCT direto — sem CTE — para minimizar trabalho
    ' do otimizador. OPTION(RECOMPILE) força plano específico para este
    ' conjunto de filtros, evitando reuso de plano ruim de consulta anterior.
    SQLCount = _
      "SELECT COUNT(DISTINCT P.CODIGO_ITEM) AS TOTAL " & _
      "FROM VW_BI_PRODUTOS_PEDIDO P " & _
      "WHERE P.CODIGO_ITEM IS NOT NULL " & _
        "AND P.CODIGO_ITEM <> '' " & _
        whereCidade & whereStatusTri & whereGrupoTri & whereTipoTri & whereCliCnt & _
      "OPTION (RECOMPILE)"

    On Error Resume Next
    Set cmdCount = Server.CreateObject("ADODB.Command")
    cmdCount.CommandTimeout = 120
    Set cmdCount.ActiveConnection = ObterConexao(dsn)
    cmdCount.CommandText = SQLCount
    cmdCount.CommandType = 1
    Set rsCount = cmdCount.Execute()
    If Err.Number <> 0 Then
      MontarJsonTrimbox = "{""sucesso"":false,""mensagem"":""Erro ao contar registros Trimbox""}"
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0

    If IsObject(rsCount) And Not (rsCount Is Nothing) Then
      nTotalTri = CLng(0 & rsCount("TOTAL"))
      rsCount.Close
      Set rsCount = Nothing
    End If
    If IsObject(cmdCount.ActiveConnection) Then
      If cmdCount.ActiveConnection.State = 1 Then cmdCount.ActiveConnection.Close
      Set cmdCount.ActiveConnection = Nothing
    End If
    Set cmdCount = Nothing
  Else
    ' Páginas 2+ : o frontend deve passar o total já conhecido via totalCan.
    ' Se não passou, mantemos 0 e o frontend trata como "total desconhecido".
    nTotalTri = CLng(0 & Trim(Request("totalCan")))
  End If

  ' ── Calcula paginação ─────────────────────────────────────────────────────
  If nTotalTri > (nPagCan * nPagSize) Then bHasMoreTri = True
  nMaxPagTri = CLng(((nTotalTri - 1) \ nPagSize) + 1)
  If nMaxPagTri < 1 Then nMaxPagTri = 1

  ' ── 2. Query de dados: somente os N registros da página ───────────────────
  '
  ' Dois caminhos gerados dinamicamente:
  '
  '   SEM linked servers (bNeedSKFT1=False, bNeedSA1010=False):
  '     Estratégia MIN-based dedup: em vez de ROW_NUMBER() OVER (PARTITION BY
  '     CODIGO_ITEM) que força materialização total de BASE, usamos
  '     MIN(LOJA_CLIENTE) para deduplicate. Isso permite que o otimizador
  '     aplique o OFFSET/FETCH antes de processar todas as linhas.
  '     COD_SAP e FANTASIA_CLIENTE retornam vazios.
  '
  '   COM linked servers (qualquer JOIN ativo):
  '     Mantém ROW_NUMBER (necessário para dedup com colunas de linked server).
  '     COD_SAP e/ou FANTASIA_CLIENTE preenchidos conforme necessário.
  '
  ' OPTION (MAXDOP 1) removido no caminho local: só prejudica performance
  ' sem linked servers. Mantido somente quando há linked server para evitar
  ' problemas de paralelismo em remote queries.

  Dim cmdData
  Dim optMaxdop

  If bNeedSKFT1 Or bNeedSA1010 Then
    ' ── Caminho COM linked servers: mantém ROW_NUMBER + MAXDOP 1 ──────────
    ctaJoined = _
      "JOINED AS ( " & _
        "SELECT " & _
          "B.CODIGO_ITEM, B.REFERENCIA_ITEM, B.CIDADE, B.GRUPO_PRODUTO, B.TIPO_PRODUTO, " & _
          "B.PRECO, B.STATUS, B.COMPRIMENTO, B.LARGURA, B.ALTURA, B.AREA, B.ESTRUTURA, " & _
          "B.CODIGO_CLIENTE, B.LOJA_CLIENTE, " & _
          colSAP & colFantasia & _
        "FROM BASE B " & _
        joinSKFT1 & _
        joinSA1010 & _
        "WHERE 1=1 " & whereNivelTri & whereCliTri & _
      "), "
    fromDedup = "FROM JOINED "
    optMaxdop  = "OPTION (MAXDOP 1)"

    SQLTRI = _
      ";WITH BASE AS ( " & _
        "SELECT " & _
          "P.CODIGO_ITEM, P.REFERENCIA_ITEM, P.CIDADE, P.GRUPO_PRODUTO, P.TIPO_PRODUTO, " & _
          "P.PRECO, P.STATUS, P.COMPRIMENTO, P.LARGURA, P.ALTURA, P.AREA, P.ESTRUTURA, " & _
          "P.CODIGO_CLIENTE, P.LOJA_CLIENTE " & _
        "FROM VW_BI_PRODUTOS_PEDIDO P " & _
        "WHERE P.CODIGO_ITEM IS NOT NULL " & _
          "AND P.CODIGO_ITEM <> '' " & _
          whereCidade & whereStatusTri & whereGrupoTri & whereTipoTri & _
      "), " & _
      ctaJoined & _
      "DEDUP AS ( " & _
        "SELECT *, " & _
          "ROW_NUMBER() OVER (PARTITION BY CODIGO_ITEM ORDER BY LOJA_CLIENTE) AS RN_DEDUP " & _
        fromDedup & _
      ") " & _
      "SELECT " & _
        "CODIGO_ITEM, REFERENCIA_ITEM, CIDADE, GRUPO_PRODUTO, TIPO_PRODUTO, " & _
        "PRECO, STATUS, COMPRIMENTO, LARGURA, ALTURA, AREA, ESTRUTURA, " & _
        "COD_SAP, CODIGO_CLIENTE, LOJA_CLIENTE, FANTASIA_CLIENTE " & _
      "FROM DEDUP " & _
      "WHERE RN_DEDUP = 1 " & _
      "ORDER BY CODIGO_ITEM " & _
      "OFFSET " & nOffsetTri & " ROWS FETCH NEXT " & nPagSize & " ROWS ONLY " & _
      optMaxdop
  Else
    ' ── Caminho LOCAL (sem linked servers): MIN-based dedup ───────────────
    ' Agrupa por CODIGO_ITEM e pega MIN(LOJA_CLIENTE) para dedup.
    ' Sem ROW_NUMBER() global — o otimizador pode usar índice e early-exit.
    ' Sem MAXDOP 1 — paralelismo local é benéfico aqui.
    ' OPTION(RECOMPILE) força plano específico para este conjunto de filtros.
    '
    ' Nota: colunas sem group-by (REFERENCIA_ITEM, PRECO, etc.) usam MIN()
    ' que é correto pois o dedup garante que o registro selecionado é sempre
    ' o de menor LOJA_CLIENTE — comportamento idêntico ao ROW_NUMBER anterior.

    Dim whereCliLocal
    whereCliLocal = ""
    If whereCliTri <> "" Then
      whereCliLocal = Replace(whereCliTri, "RTRIM(CODIGO_CLIENTE)", "RTRIM(P.CODIGO_CLIENTE)")
      whereCliLocal = Replace(whereCliLocal, "LOJA_CLIENTE", "P.LOJA_CLIENTE")
    End If

    SQLTRI = _
      "SELECT " & _
        "CODIGO_ITEM, REFERENCIA_ITEM, CIDADE, GRUPO_PRODUTO, TIPO_PRODUTO, " & _
        "PRECO, STATUS, COMPRIMENTO, LARGURA, ALTURA, AREA, ESTRUTURA, " & _
        "CODIGO_CLIENTE, LOJA_CLIENTE, COD_SAP, FANTASIA_CLIENTE " & _
      "FROM ( " & _
        "SELECT " & _
          "MIN(CODIGO_ITEM)      AS CODIGO_ITEM, " & _
          "MIN(REFERENCIA_ITEM)  AS REFERENCIA_ITEM, " & _
          "MIN(CIDADE)           AS CIDADE, " & _
          "MIN(GRUPO_PRODUTO)    AS GRUPO_PRODUTO, " & _
          "MIN(TIPO_PRODUTO)     AS TIPO_PRODUTO, " & _
          "MIN(PRECO)            AS PRECO, " & _
          "MIN(STATUS)           AS STATUS, " & _
          "MIN(COMPRIMENTO)      AS COMPRIMENTO, " & _
          "MIN(LARGURA)          AS LARGURA, " & _
          "MIN(ALTURA)           AS ALTURA, " & _
          "MIN(AREA)             AS AREA, " & _
          "MIN(ESTRUTURA)        AS ESTRUTURA, " & _
          "MIN(CODIGO_CLIENTE)   AS CODIGO_CLIENTE, " & _
          "MIN(LOJA_CLIENTE)     AS LOJA_CLIENTE, " & _
          "''                    AS COD_SAP, " & _
          "''                    AS FANTASIA_CLIENTE " & _
        "FROM VW_BI_PRODUTOS_PEDIDO P " & _
        "WHERE P.CODIGO_ITEM IS NOT NULL " & _
          "AND P.CODIGO_ITEM <> '' " & _
          whereCidade & whereStatusTri & whereGrupoTri & whereTipoTri & whereCliLocal & _
        "GROUP BY P.CODIGO_ITEM " & _
      ") AS DEDUP " & _
      "ORDER BY CODIGO_ITEM " & _
      "OFFSET " & nOffsetTri & " ROWS FETCH NEXT " & nPagSize & " ROWS ONLY " & _
      "OPTION (RECOMPILE)"
  End If

  ' Executa com timeout explícito de 120s via ADODB.Command
  regs = ""

  On Error Resume Next
  Set cmdData = Server.CreateObject("ADODB.Command")
  cmdData.CommandTimeout = 120
  Set cmdData.ActiveConnection = ObterConexao(dsn)
  cmdData.CommandText = SQLTRI
  cmdData.CommandType = 1
  Set rs = cmdData.Execute()
  If Err.Number <> 0 Then
    MontarJsonTrimbox = "{""sucesso"":false,""mensagem"":""Erro ao consultar FTs Trimbox""}"
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsObject(rs) And Not (rs Is Nothing) Then
    Do While Not rs.EOF
      Dim codTri, urlTri, reg

      codTri = Trim(rs("CODIGO_ITEM"))
      urlTri = "getImgFT.asp?cod=" & codTri

      ' imgExiste = False: evita N chamadas ao filesystem no loop.
      ' O frontend deve tentar carregar a imagem e exibir placeholder em caso de 404.
      reg = "{" & _
        """codigo"":"""      & JStr(codTri)                         & """," & _
        """imgExiste"":false,"                                                    & _
        """urlImg"":"""      & JStr(urlTri)                         & """," & _
        """referencia"":""" & JStr(rs("REFERENCIA_ITEM"))          & """," & _
        """cidade"":""" & JStr(fNomePlanta(pPlanta))           & """," & _
        """grupo"":""" & JStr(rs("GRUPO_PRODUTO"))            & """," & _
        """tipo"":""" & JStr(rs("TIPO_PRODUTO"))             & """," & _
        """preco"":"         & JNum(rs("PRECO"))                          & "," & _
        """status"":""" & JStr(rs("STATUS"))                   & """," & _
        """compr"":""" & JStr(rs("COMPRIMENTO"))              & """," & _
        """larg"":""" & JStr(rs("LARGURA"))                  & """," & _
        """altu"":""" & JStr(rs("ALTURA"))                   & """," & _
        """area"":"          & JNum(rs("AREA"))                           & "," & _
        """estrutura"":""" & JStr(rs("ESTRUTURA"))                & """," & _
        """sap"":""" & JStr(rs("COD_SAP"))                  & """," & _
        """codcli"":""" & JStr(rs("CODIGO_CLIENTE"))           & """," & _
        """lojacli"":""" & JStr(rs("LOJA_CLIENTE"))             & """," & _
        """fantasia"":""" & JStr(rs("FANTASIA_CLIENTE"))         & """" & _
        "}"

      regs = regs & SepInline(regs) & reg
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
  End If

  ' Libera o Command e fecha a conexão aberta por ObterConexao()
  If IsObject(cmdData) Then
    If IsObject(cmdData.ActiveConnection) Then
      If cmdData.ActiveConnection.State = 1 Then cmdData.ActiveConnection.Close
      Set cmdData.ActiveConnection = Nothing
    End If
    Set cmdData = Nothing
  End If

  Dim jsonT
  jsonT = "{""sucesso"":true," & _
    """tipo"":""TRIMBOX""," & _
    """nivel"":"            & nivelUsr                       & "," & _
    """filial"":""" & JStr(pPlanta)             & """," & _
    """cidade"":""" & JStr(fNomePlanta(pPlanta)) & """," & _
    """cliente"":{""codigo"":""" & JStr(sCliCodTri) & """,""nome"":""" & JStr(sCliNomeTri) & """,""loja"":""" & JStr(sCliLojaTri) & """}," & _
    """paginacao"":{" & _
      """pagina"":"     & nPagCan             & "," & _
      """totalCan"":"   & nTotalTri           & "," & _
      """maxPag"":"     & nMaxPagTri          & "," & _
      """hasMore"":"    & JBool(bHasMoreTri)  & "," & _
      """pageSize"":"   & nPagSize       & _
    "}," & _
    """cancelado"":false," & _
    """secoes"":[{""titulo"":"" "",""registros"":[" & regs & "]}]" & _
    "}"

  MontarJsonTrimbox = jsonT
End Function
%>
