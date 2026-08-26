<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.ContentType = "application/json"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Pragma",        "no-cache"
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

If acao = "CLIENTES" Then
  Dim sqlCli, rsCli, arrCli, bFirstCli, qParam, whereQ
  qParam  = Trim(Request("q"))
  whereQ  = ""
  If Len(qParam) >= 2 Then
    Dim qSafe : qSafe = Replace(qParam, "'", "''")
    whereQ = "AND (CodigoCliente LIKE '%" & UCase(qSafe) & "%' OR UPPER(NomeCliente) LIKE '%" & UCase(qSafe) & "%') "
  End If
  sqlCli = "" &_
    "SELECT TOP 50 Cliente FROM (" &_
      "SELECT DISTINCT CONCAT('[', CodigoCliente, '] ', NomeCliente) AS Cliente, NomeCliente " &_
      "FROM CANHOTOS " &_
      "WHERE D_E_L_E_T_ <> '*' " &_
      whereQ &_
    ") AS sub " &_
    "ORDER BY NomeCliente"
  Set rsCli  = Conecta("PIR-CORPORATIVO", sqlCli)
  arrCli     = ""
  bFirstCli  = True
  Do While Not rsCli.EOF
    If Not bFirstCli Then arrCli = arrCli & ","
    arrCli    = arrCli & JsonStr(rsCli("Cliente"))
    bFirstCli = False
    rsCli.MoveNext()
  Loop
  rsCli.Close()
  Set rsCli = Nothing
  Response.Write "[" & arrCli & "]"
  Response.End
End If

If acao <> "FILTRAR" Then
  Response.Write "[]"
  Response.End
End If

Dim di, df
di = Trim(Request("dtini"))
df = Trim(Request("dtfin"))

If Not IsDate(di) Or Not IsDate(df) Then
  Response.Write "{""erro"":""Datas inválidas.""}"
  Response.End
End If

If CDate(di) > CDate(df) Then
  Response.Write "{""erro"":""A data inicial deve ser menor ou igual à data final.""}"
  Response.End
End If

Dim diFormatada, dfFormatada
diFormatada = formataDataYYYYMMDD(di)
dfFormatada = formataDataYYYYMMDD(df)

If Not (Len(diFormatada) = 8 And IsNumeric(diFormatada)) Or _
   Not (Len(dfFormatada) = 8 And IsNumeric(dfFormatada)) Then
  Response.Write "{""erro"":""Formato de data inválido.""}"
  Response.End
End If

Dim FilSql, CrgSql, VeiSql
FilSql = ""
CrgSql = ""
VeiSql = ""

Dim filialParam, cargaParam, veiculoParam
filialParam  = Trim(Request("filial"))
cargaParam   = Trim(Request("carga"))
veiculoParam = Trim(Request("veiculo"))

If Len(filialParam) >= 1 Then
  Dim Fil, fi
  Fil = Split(filialParam, ",")
  For fi = 0 To UBound(Fil)
    If Trim(Fil(fi)) <> "" Then FilSql = FilSql & "'" & Trim(Fil(fi)) & "',"
  Next
  If FilSql <> "" Then FilSql = Mid(FilSql, 1, Len(FilSql) - 1)
End If

If Len(cargaParam) >= 1 Then
  Dim Crg, ci
  Crg = Split(cargaParam, ",")
  For ci = 0 To UBound(Crg)
    If Trim(Crg(ci)) <> "" Then CrgSql = CrgSql & "'" & Trim(Crg(ci)) & "',"
  Next
  If CrgSql <> "" Then CrgSql = Mid(CrgSql, 1, Len(CrgSql) - 1)
End If

If Len(veiculoParam) >= 1 Then
  Dim Vei, vi
  Vei = Split(veiculoParam, ",")
  For vi = 0 To UBound(Vei)
    If Trim(Vei(vi)) <> "" Then VeiSql = VeiSql & "'" & Trim(Vei(vi)) & "',"
  Next
  If VeiSql <> "" Then VeiSql = Mid(VeiSql, 1, Len(VeiSql) - 1)
End If

Dim serieParam, nfParam, clienteParam
serieParam   = Trim(Request("serie"))
nfParam      = Trim(Request("nf"))
clienteParam = Trim(Request("cliente"))

Dim PltSql
PltSql = ""
If Len(serieParam) >= 1 Then
  Dim Plt, pi
  Plt = Split(serieParam, ",")
  For pi = 0 To UBound(Plt)
    If Trim(Plt(pi)) <> "" Then PltSql = PltSql & "'" & Trim(Plt(pi)) & "',"
  Next
  If PltSql <> "" Then PltSql = Mid(PltSql, 1, Len(PltSql) - 1)
End If

Dim whereData, whereNf, whereSerie, whereFilial, whereCarga, whereVeiculo, whereCliente
whereData    = "AND CONVERT(VARCHAR(8), A.DATA, 112) BETWEEN '" & diFormatada & "' AND '" & dfFormatada & "' "
whereNf      = ""
whereSerie   = ""
whereFilial  = ""
whereCarga   = ""
whereVeiculo = ""
whereCliente = ""

If nfParam      <> "" Then whereNf      = "AND A.NF = '" & Replace(nfParam, "'", "''") & "' "
If PltSql       <> "" Then whereSerie   = "AND A.SERIE IN (" & PltSql & ") "
If FilSql       <> "" Then whereFilial  = "AND A.Filial IN (" & FilSql & ") "
If CrgSql       <> "" Then whereCarga   = "AND Carga IN (" & CrgSql & ") "
If VeiSql       <> "" Then whereVeiculo = "AND Veiculo IN (" & VeiSql & ") "
If clienteParam <> "" Then whereCliente = "AND CodigoCliente = '" & Replace(clienteParam, "'", "''") & "' "

Dim SQL
SQL = "" &_
  "SELECT " &_
    "A.NF, " &_
    "A.DATA, " &_
    "A.SERIE, " &_
    "A.USERDIGIT, " &_
    "A.OBS, " &_
    "ISNULL(A.Carga,0) AS Carga, " &_
    "ISNULL(A.Veiculo,0) AS Veiculo, " &_
    "A.Filial, " &_
    "CONCAT('[', A.CodigoCliente, '] ', A.NomeCliente) AS Cliente " &_
  "FROM CANHOTOS A " &_
  "WHERE A.D_E_L_E_T_ <> '*' " &_
    whereData &_
    whereNf &_
    whereSerie &_
    whereFilial &_
    whereCarga &_
    whereVeiculo &_
    whereCliente &_
  "ORDER BY A.DATA DESC, A.NF, A.SERIE"

Dim rs
Set rs = Conecta("PIR-CORPORATIVO", SQL)

Dim arrRows, bFirst
arrRows = ""
bFirst  = True

Do While Not rs.EOF

  Dim labelCarga, labelVeiculo, labelFilial
  labelCarga   = ""
  labelVeiculo = ""
  labelFilial  = ""

  Select Case CStr(rs("Carga"))
    Case "1" labelCarga = "Muito Insatisfeito"
    Case "2" labelCarga = "Insatisfeito"
    Case "3" labelCarga = "Indiferente"
    Case "4" labelCarga = "Satisfeito"
    Case "5" labelCarga = "Muito Satisfeito"
    Case Else labelCarga = "Não respondido"
  End Select

  Select Case CStr(rs("Veiculo"))
    Case "1" labelVeiculo = "Muito Insatisfeito"
    Case "2" labelVeiculo = "Insatisfeito"
    Case "3" labelVeiculo = "Indiferente"
    Case "4" labelVeiculo = "Satisfeito"
    Case "5" labelVeiculo = "Muito Satisfeito"
    Case Else labelVeiculo = "Não respondido"
  End Select

  Select Case CStr(rs("Filial"))
    Case "0" labelFilial = "Pirapetinga"
    Case "4" labelFilial = "Uberaba"
    Case "5" labelFilial = "Bento Gonçalves"
    Case "6" labelFilial = "Fortaleza"
    Case "7" labelFilial = "Saquarema"
    Case Else labelFilial = ""
  End Select

  If Not bFirst Then arrRows = arrRows & ","
  arrRows = arrRows & "{" &_
    """nf"":"      & JsonStr(rs("NF"))                          & "," &_
    """serie"":"   & JsonStr(rs("SERIE"))                       & "," &_
    """cliente"":"  & JsonStr(rs("Cliente"))                   & "," &_
    """filial"":"  & JsonStr(labelFilial)                       & "," &_
    """data"":"    & JsonStr(acertadata(CStr(rs("DATA"))))      & "," &_
    """carga"":"   & JsonStr(labelCarga)                        & "," &_
    """veiculo"":"  & JsonStr(labelVeiculo)                    & "," &_
    """usuario"":"  & JsonStr(rs("USERDIGIT"))                 & "," &_
    """obs"":"     & JsonStr(rs("OBS"))                         &_
  "}"
  bFirst = False
  rs.MoveNext()
Loop

rs.Close()
Set rs = Nothing

Response.Write "[" & arrRows & "]"
%>
