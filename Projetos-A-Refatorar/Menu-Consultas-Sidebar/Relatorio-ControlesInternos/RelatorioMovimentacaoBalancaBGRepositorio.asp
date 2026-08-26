<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.ContentType = "application/json"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Session.LCID = 1046

Function JsonStr(v)
  If IsNull(v) Then
    JsonStr = "null"
  Else
    Dim s
    s = CStr(v)
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(9),  "\t")
    JsonStr = """" & s & """"
  End If
End Function

Function JsonNum(v)
  If IsNull(v) Or CStr(v) = "" Then
    JsonNum = "null"
  Else
    JsonNum = CStr(CDbl(v))
  End If
End Function

Dim acao
acao = Trim(Request("acao"))

If acao <> "FILTRAR" Then
  Response.Write "[]"
  Response.End
End If

Dim di, df
di = Trim(Request("datainicial"))
df = Trim(Request("datafinal"))

If Not IsDate(di) Or Not IsDate(df) Then
  Response.Write "{""erro"":""Datas inválidas.""}"
  Response.End
End If

If CDate(di) > CDate(df) Then
  Response.Write "{""erro"":""A data inicial deve ser menor ou igual à data final.""}"
  Response.End
End If

Dim diffDias
diffDias = DateDiff("d", CDate(di), CDate(df))
If diffDias > 14 Then
  Response.Write "{""erro"":""O período deve ser de no máximo 15 dias.""}"
  Response.End
End If

Dim diFormatada, dfFormatada
diFormatada = formataDataYYYYMMDD(di)
dfFormatada = formataDataYYYYMMDD(df)

If Not (Len(diFormatada) = 8 And IsNumeric(diFormatada)) Or Not (Len(dfFormatada) = 8 And IsNumeric(dfFormatada)) Then
  Response.Write "{""erro"":""Formato de data inválido.""}"
  Response.End
End If

Dim SQL
SQL = "" &_
  "SELECT " &_
    "TICKET, " &_
    "DATA_CRIACAO_TICKET, " &_
    "PLACA_VEICULO, " &_
    "PLACA_DA_CARRETA, " &_
    "TRANSPORTADORA, " &_
    "COD_EMISSOR, " &_
    "NOME_EMISSOR, " &_
    "COD_ITEM, " &_
    "DESCRICAO_ITEM, " &_
    "CAMPO_ADICIONAL, " &_
    "NOTA_FISCAL, " &_
    "PESO_INICIAL, " &_
    "DATA_PESO_INICIAL, " &_
    "HORA_PESO_INICIAL, " &_
    "PESO_FINAL, " &_
    "DATA_PESO_FINAL, " &_
    "HORA_PESO_FINAL, " &_
    "PESO_LIQUIDO " &_
  "FROM dbo.vw_consulta_ticket_bento " &_
  "WHERE CONVERT(VARCHAR(10), DATA_CRIACAO_TICKET, 112) BETWEEN '" & diFormatada & "' AND '" & dfFormatada & "' " &_
  "ORDER BY TICKET DESC"

Dim rs
Set rs = Conecta("BEN-GUARDIANTI", SQL)

Dim arrRows
arrRows = ""
Dim bFirst
bFirst = True

Do While Not rs.EOF
  If Not bFirst Then arrRows = arrRows & ","
  arrRows = arrRows & "{" &_
    """ticket"":"       & JsonStr(rs("TICKET"))               & "," &_
    """dataCriacao"":"  & JsonStr(rs("DATA_CRIACAO_TICKET"))   & "," &_
    """placaVei"":"     & JsonStr(rs("PLACA_VEICULO"))         & "," &_
    """placaCar"":"     & JsonStr(rs("PLACA_DA_CARRETA"))      & "," &_
    """transp"":"       & JsonStr(rs("TRANSPORTADORA"))        & "," &_
    """codEmi"":"       & JsonStr(rs("COD_EMISSOR"))           & "," &_
    """nomeEmi"":"      & JsonStr(rs("NOME_EMISSOR"))          & "," &_
    """codItem"":"      & JsonStr(rs("COD_ITEM"))              & "," &_
    """descrItem"":"    & JsonStr(rs("DESCRICAO_ITEM"))        & "," &_
    """campoAdic"":"    & JsonStr(rs("CAMPO_ADICIONAL"))       & "," &_
    """nf"":"           & JsonStr(rs("NOTA_FISCAL"))           & "," &_
    """pesoIni"":"      & JsonNum(rs("PESO_INICIAL"))          & "," &_
    """dataPesoIni"":"  & JsonStr(rs("DATA_PESO_INICIAL"))     & "," &_
    """horaPesoIni"":"  & JsonStr(rs("HORA_PESO_INICIAL"))     & "," &_
    """pesoFin"":"      & JsonNum(rs("PESO_FINAL"))            & "," &_
    """dataPesoFin"":"  & JsonStr(rs("DATA_PESO_FINAL"))       & "," &_
    """horaPesoFin"":"  & JsonStr(rs("HORA_PESO_FINAL"))       & "," &_
    """pesoLiq"":"      & JsonNum(rs("PESO_LIQUIDO"))          &_
  "}"
  bFirst = False
  rs.MoveNext()
Loop

rs.Close()
Set rs = Nothing

Response.Write "[" & arrRows & "]"
%>
