<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.ContentType = "application/json"
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------
Function FormatISO(dt)
  If IsDate(dt) Then
    FormatISO = Right("0" & Day(dt),    2) & "/" & _
                Right("0" & Month(dt),  2) & "/" & _
                Year(dt)                   & " " & _
                Right("0" & Hour(dt),   2) & ":" & _
                Right("0" & Minute(dt), 2) & ":" & _
                Right("0" & Second(dt), 2)
  Else
    FormatISO = dt
  End If
End Function

Function JsonStr(s)
  JsonStr = Replace(Replace(CStr(s), "\", "\\"), """", "\""")
End Function

Function JsonErro(msg)
  JsonErro = "{""erro"":""" & JsonStr(msg) & """}"
End Function

' ---------------------------------------------------------------------------
' Roteamento
' ---------------------------------------------------------------------------
Dim acao : acao = Trim(Request("acao"))

Select Case UCase(acao)

  Case "FILTRAR"
    Call AcaoFiltrar()

  Case Else
    Response.Status = "400 Bad Request"
    Response.Write JsonErro("Ação inválida.")

End Select

Response.End

' ---------------------------------------------------------------------------
' FILTRAR
' ---------------------------------------------------------------------------
Sub AcaoFiltrar()
  Dim dtIni, dtFim, identidade
  dtIni      = Replace(Trim(Request("datainicial")), "-", "")
  dtFim      = Replace(Trim(Request("datafinal")),   "-", "")
  identidade = Trim(Request("identidade"))

  If dtIni = "" Or dtFim = "" Then
    Response.Status = "422 Unprocessable Entity"
    Response.Write JsonErro("Preencha o período.")
    Exit Sub
  End If

  Dim filtroIdent : filtroIdent = ""
  If identidade <> "" Then
    Dim identEsc : identEsc = Replace(identidade, "'", "''")
    filtroIdent = "AND (CAST(L.CRED_NUMERO AS NVARCHAR(50)) LIKE '%" & identEsc & "%' " & _
                  "OR CAST(L.PES_NUMERO AS NVARCHAR(50)) LIKE '%" & identEsc & "%' " & _
                  "OR P.PES_NOME LIKE '%" & identEsc & "%') "
  End If

  Dim sql
  sql = "" & _
    "SELECT " & _
      "L.CRED_NUMERO, " & _
      "L.PES_NUMERO, " & _
      "ISNULL(P.PES_NOME, '') AS PES_NOME, " & _
      "L.EQPI_NUMERO, " & _
      "ISNULL(E.EQPI_DESCRICAO, '') AS EQPI_DESCRICAO, " & _
      "L.MOV_DATAHORA " & _
    "FROM LOG_CREDENCIAL L " & _
      "LEFT JOIN PESSOAS     P ON L.PES_NUMERO  = P.PES_NUMERO " & _
      "LEFT JOIN EQUIPAMENTOS E ON L.EQPI_NUMERO = E.EQPI_NUMERO " & _
    "WHERE L.GRP_NUMERO    = 2 " & _
      "AND L.LGCRTI_NUMERO IN (10, 25, 12, 27, 21) " & _
      "AND L.MOV_DATAHORA  > '20140205 23:59:59.000' " & _
      "AND CONVERT(NVARCHAR(8), L.MOV_DATAHORA, 112) BETWEEN '" & dtIni & "' AND '" & dtFim & "' " & _
      filtroIdent & _
    "ORDER BY P.PES_NOME, L.MOV_DATAHORA"

  Dim rs : Set rs = Conecta("PIR-RBACESSO", sql)

  If Not IsObject(rs) Or (rs Is Nothing) Then
    Response.Status = "500 Internal Server Error"
    Response.Write JsonErro("Falha ao conectar ao banco de dados.")
    Exit Sub
  End If

  Const CHUNK  = 200
  Dim primeiro : primeiro = True
  Dim buf      : buf      = ""
  Dim contador : contador = 0

  Dim jCred : jCred = """credencial"":"
  Dim jMatr : jMatr = """matricula"":"
  Dim jNome : jNome = """nome"":"
  Dim jEqpi : jEqpi = """equipamento"":"
  Dim jDesc : jDesc = """descricao"":"
  Dim jData : jData = """datahora"":"

  Response.Write "["

  Do While Not rs.EOF
    If Not primeiro Then buf = buf & ","
    buf = buf & "{" & _
      jCred & """" & JsonStr(rs("CRED_NUMERO")    & "") & """," & _
      jMatr & """" & JsonStr(rs("PES_NUMERO")     & "") & """," & _
      jNome & """" & JsonStr(rs("PES_NOME")       & "") & """," & _
      jEqpi & """" & JsonStr(rs("EQPI_NUMERO")    & "") & """," & _
      jDesc & """" & JsonStr(rs("EQPI_DESCRICAO") & "") & """," & _
      jData & """" & FormatISO(rs("MOV_DATAHORA"))       & """"  & _
    "}"
    primeiro = False
    contador = contador + 1

    If contador Mod CHUNK = 0 Then
      Response.Write buf
      Response.Flush
      buf = ""
    End If

    rs.MoveNext()
  Loop

  rs.Close()
  Set rs = Nothing

  If buf <> "" Then Response.Write buf
  Response.Write "]"
End Sub
%>
