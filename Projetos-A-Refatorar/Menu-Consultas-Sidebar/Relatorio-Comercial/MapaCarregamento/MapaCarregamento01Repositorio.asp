<%@ CodePage=65001 Language="VBScript" %>
<%
  Response.CodePage    = 65001
  Response.Charset     = "utf-8"
  Response.ContentType = "application/json"
  Response.CacheControl = "no-cache"
%>
<!--#include file="../../../cnn2.1/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp" -->

<%
' ═══════════════════════════════════════════════════════════
' MapaCarregamento01Repositorio.asp
' Endpoint JSON do Mapa de Carregamento.
'
' GET ?DataIn=AAAA-MM-DD&cboFilial=00
'
' Resposta de sucesso:
' {
'   "ok": true,
'   "registros": [
'     {
'       "placa":        "ABC1D23",
'       "transportadora": "INPA TRANSPORTES L",
'       "tel":          "32 9999-9999",
'       "pedido":       "000123",
'       "cliente":      "00001 NOME REDUZIDO",
'       "pedCompra":    "PC001-01",
'       "referencia":   "DESCRIÇÃO DO PRODUTO",
'       "nfSerie":      "000456 - 1",
'       "dataEmissao":  "30/07/2025",
'       "hora":         "08:30",
'       "qtd":          10,
'       "peso":         150.0,
'       "area":         2.5,
'       "encarregado":  "JOAO SILVA",
'       "dtEntrega":    "31/07/2025",
'       "localEntrega": "SAO PAULO",
'       "tipoProduto":  "PA"
'     }
'   ],
'   "resumo": {
'     "totalRegistros": 42,
'     "peso":  { "total": 5000, "PA": 1000, "ML": 500, "CV": 800, "CH": 2700 },
'     "area":  { "total": 120,  "PA": 30,   "ML": 10,  "CV": 20,  "CH": 60   },
'     "frota": { "total": 5, "propria": 2, "terceirizada": 3 }
'   }
' }
'
' Resposta de erro:
' { "ok": false, "erro": "mensagem" }
' ═══════════════════════════════════════════════════════════

' ─── Helpers ────────────────────────────────────────────────

Function NzTxt(v)
  If IsNull(v) Or IsEmpty(v) Then NzTxt = "" Else NzTxt = CStr(v & "")
End Function

Function ParseNum(v)
  Dim s
  If IsNull(v) Or IsEmpty(v) Then ParseNum = 0 : Exit Function
  s = Trim(CStr(v & ""))
  If s = "" Then ParseNum = 0 : Exit Function
  s = Replace(s, Chr(160), "") : s = Replace(s, " ", "")
  If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
    s = Replace(s, ".", "") : s = Replace(s, ",", ".")
  ElseIf InStr(s, ",") > 0 Then
    s = Replace(s, ",", ".")
  ElseIf InStr(s, ".") > 0 Then
    s = Replace(s, ".", "")
  End If
  On Error Resume Next
  ParseNum = CDbl(s)
  If Err.Number <> 0 Then ParseNum = 0 : Err.Clear
  On Error GoTo 0
End Function

' Escapa aspas duplas e caracteres de controle para JSON seguro
Function JsonStr(v)
  Dim s : s = NzTxt(v)
  s = Replace(s, "\",  "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, Chr(13), "\r")
  s = Replace(s, Chr(10), "\n")
  s = Replace(s, Chr(9),  "\t")
  JsonStr = s
End Function

' Formata data no formato Protheus AAAAMMDD para DD/MM/AAAA
Function FormatarDataProtheus(v)
  Dim s : s = NzTxt(v)
  If Len(s) < 8 Then FormatarDataProtheus = "" : Exit Function
  FormatarDataProtheus = Right(s, 2) & "/" & Mid(s, 5, 2) & "/" & Left(s, 4)
End Function

' ─── Validação da entrada ───────────────────────────────────

Dim paramDataIn  : paramDataIn  = Trim(Request.QueryString("DataIn")    & "")
Dim paramFilial  : paramFilial  = Trim(Request.QueryString("cboFilial")  & "")

If paramDataIn = "" Then
  Response.Write "{""ok"":false,""erro"":""Par\u00e2metro DataIn obrigat\u00f3rio.""}"
  Response.End
End If

If paramFilial = "" Then paramFilial = "00"

' DataIn chega como AAAA-MM-DD do input[type=date]; converte para AAAAMMDD (formato Protheus)
Dim rsDATA : rsDATA = Replace(paramDataIn, "-", "")

' ─── Consulta 1 — range de notas (SZ6) ─────────────────────

Dim SQL_SZ6
SQL_SZ6 = "" &_
  "SELECT Z6_NFINI AS NOTAIN, Z6_NFFIM AS NOTAFIM, " &_
         "Z6_NFINIB AS NOTAIN2, Z6_NFFINB AS NOTAFIM2 " &_
  "FROM SZ6010 " &_
  "WHERE Z6_DATA   = '" & Replace(rsDATA,     "'", "''") & "' " &_
  "  AND Z6_FILIAL = '" & Replace(paramFilial, "'", "''") & "' " &_
  "  AND D_E_L_E_T_ <> '*'"

Dim rsSZ6 : Set rsSZ6 = Conecta("PIR-DADOSADV", SQL_SZ6)
If (Not IsObject(rsSZ6)) Or (rsSZ6 Is Nothing) Then
  Response.Write "{""ok"":false,""erro"":""Falha na conex\u00e3o com o banco de dados.""}"
  Response.End
End If

If rsSZ6.EOF Then
  rsSZ6.Close : Set rsSZ6 = Nothing
  Response.Write "{""ok"":false,""erro"":""SZ6 - N\u00e3o h\u00e1 fechamento di\u00e1rio do faturamento na data solicitada.""}"
  Response.End
End If

Dim rsNOTAIN, rsNOTAFIM, rsNOTAIN2, rsNOTAFIM2
rsNOTAIN   = Trim(NzTxt(rsSZ6("NOTAIN")))
rsNOTAFIM  = Trim(NzTxt(rsSZ6("NOTAFIM")))
rsNOTAIN2  = Trim(NzTxt(rsSZ6("NOTAIN2")))
rsNOTAFIM2 = Trim(NzTxt(rsSZ6("NOTAFIM2")))
rsSZ6.Close : Set rsSZ6 = Nothing

' ─── Consulta 2 — SP principal ──────────────────────────────

Dim rsRepre01, rsRepre02
rsRepre01 = iif(Nivel = 0, "000", CodUser)
rsRepre02 = iif(Nivel = 0, "ZZZ", CodUser)

Dim SQL_SP
SQL_SP = "EXEC dbo.SP_CONMAPACARREGAMENTO '" & _
          Replace(rsNOTAIN,   "'", "''") & "','" & _
          Replace(rsNOTAFIM,  "'", "''") & "','" & _
          Replace(rsRepre01,  "'", "''") & "','" & _
          Replace(rsRepre02,  "'", "''") & "','" & _
          Replace(rsNOTAIN2,  "'", "''") & "','" & _
          Replace(rsNOTAFIM2, "'", "''") & "','" & _
          Replace(rsDATA,     "'", "''") & "','" & _
          Replace(paramFilial,"'", "''") & "'"

Dim rsSP : Set rsSP = Conecta("PIR-DADOSADV", SQL_SP)
If (Not IsObject(rsSP)) Or (rsSP Is Nothing) Then
  Response.Write "{""ok"":false,""erro"":""Falha ao executar a procedure.""}"
  Response.End
End If

If rsSP.EOF Then
  rsSP.Close : Set rsSP = Nothing
  Response.Write "{""ok"":true,""registros"":[],""resumo"":{""totalRegistros"":0,""peso"":{""total"":0,""PA"":0,""ML"":0,""CV"":0,""CH"":0},""area"":{""total"":0,""PA"":0,""ML"":0,""CV"":0,""CH"":0},""frota"":{""total"":0,""propria"":0,""terceirizada"":0}}}"
  Response.End
End If

' ─── Serialização ───────────────────────────────────────────

Dim Total, Total2
Dim TotalPA, TotalML, TotalCH, TotalCV
Dim TotalPAm2, TotalMLm2, TotalCHm2, TotalCVm2
Dim Frota, FrotaPropria, FrotaTerceirizada, Num

Total = 0 : Total2 = 0
TotalPA = 0 : TotalML = 0 : TotalCH = 0 : TotalCV = 0
TotalPAm2 = 0 : TotalMLm2 = 0 : TotalCHm2 = 0 : TotalCVm2 = 0
Frota = 0 : FrotaPropria = 0 : FrotaTerceirizada = 0
Num = 0

Dim primeiroItem : primeiroItem = True
Dim placaAnterior : placaAnterior = ""

Response.Write "{""ok"":true,""registros"":["

Do While Not rsSP.EOF

  Dim vPlaca, vNome, vTel
  vPlaca = Trim(NzTxt(rsSP("F2_PLACFRT")))
  If vPlaca = "" Then vPlaca = "(SEM PLACA)"
  vNome = Left(Trim(NzTxt(rsSP("A4_NOME"))), 20)
  vTel  = Trim(NzTxt(rsSP("A4_TEL")))

  ' Contagem de frota — apenas uma vez por placa
  If vPlaca <> placaAnterior Then
    If Left(Trim(NzTxt(rsSP("A4_NOME"))), 4) = "INPA" Then
      FrotaPropria = FrotaPropria + 1
    Else
      FrotaTerceirizada = FrotaTerceirizada + 1
    End If
    Frota = Frota + 1
    placaAnterior = vPlaca
  End If

  ' Cálculo de peso e área
  Dim vQtd, vPesoSP, vAreaSP, vB1Peso, vLarg, vComp
  Dim vPesoItem, vAreaItem

  vQtd    = ParseNum(rsSP("D2_QUANT"))
  vPesoSP = ParseNum(rsSP("D2_PCARGA"))
  vAreaSP = ParseNum(rsSP("D2_AREAAUD"))
  vB1Peso = ParseNum(rsSP("B1_PESO"))
  vLarg   = ParseNum(rsSP("C6_LARG"))
  vComp   = ParseNum(rsSP("C6_COMP"))

  If vPesoSP > 0 Then
    vPesoItem = vPesoSP
  ElseIf vB1Peso > 0 And vQtd > 0 Then
    vPesoItem = vB1Peso * vQtd
  Else
    vPesoItem = 0
  End If

  If vAreaSP > 0 Then
    vAreaItem = vAreaSP
  ElseIf vLarg > 0 And vComp > 0 And vQtd > 0 Then
    vAreaItem = (vLarg / 1000) * (vComp / 1000) * vQtd
  Else
    vAreaItem = 0
  End If

  ' Acumuladores de resumo
  Total  = Total  + vPesoItem
  Total2 = Total2 + vAreaItem

  Dim vDescNF : vDescNF = Trim(NzTxt(rsSP("B1_DESCNF")))
  Dim vTipoProduto : vTipoProduto = ""
  Select Case vDescNF
    Case "900" : TotalPA   = TotalPA   + vPesoItem : TotalPAm2 = TotalPAm2 + vAreaItem : vTipoProduto = "PA"
    Case "901" : TotalML   = TotalML   + vPesoItem : TotalMLm2 = TotalMLm2 + vAreaItem : vTipoProduto = "ML"
    Case "902" : TotalCV   = TotalCV   + vPesoItem : TotalCVm2 = TotalCVm2 + vAreaItem : vTipoProduto = "CV"
    Case "903" : TotalCH   = TotalCH   + vPesoItem : TotalCHm2 = TotalCHm2 + vAreaItem : vTipoProduto = "CH"
  End Select

  Num = Num + 1

  ' Separador de itens JSON
  If Not primeiroItem Then Response.Write ","
  primeiroItem = False

  ' Serializa o registro
  Response.Write "{" & _
    """placa"":"         & """" & JsonStr(vPlaca) & """," & _
    """transportadora"":""" & JsonStr(vNome)   & """," & _
    """tel"":"           & """" & JsonStr(vTel)  & """," & _
    """pedido"":"        & """" & JsonStr(NzTxt(rsSP("D2_PEDIDO")))                                    & """," & _
    """cliente"":"       & """" & JsonStr(NzTxt(rsSP("CODIGO")) & " " & Left(NzTxt(rsSP("A1_NREDUZ")),25)) & """," & _
    """pedCompra"":"     & """" & JsonStr(Trim(NzTxt(rsSP("C6_NUMPCOM"))) & "-" & NzTxt(rsSP("C6_ITEMPC"))) & """," & _
    """referencia"":"    & """" & JsonStr(Left(NzTxt(rsSP("B1_DESC")), 50))                            & """," & _
    """nfSerie"":"       & """" & JsonStr(NzTxt(rsSP("D2_DOC")) & " - " & NzTxt(rsSP("D2_SERIE")))    & """," & _
    """dataEmissao"":"   & """" & JsonStr(FormatarDataProtheus(NzTxt(rsSP("D2_EMISSAO"))))             & """," & _
    """hora"":"          & """" & JsonStr(NzTxt(rsSP("D2_HORA")))                                      & """," & _
    """qtd"":"           & CStr(CLng(vQtd))                                                            & ","  & _
    """peso"":"          & Replace(CStr(CDbl(vPesoItem)), ",", ".")                                    & ","  & _
    """area"":"          & Replace(CStr(CDbl(vAreaItem)), ",", ".")                                    & ","  & _
    """encarregado"":"   & """" & JsonStr(Left(NzTxt(rsSP("D2_ENCA")), 12))                           & """," & _
    """dtEntrega"":"     & """" & JsonStr(FormatarDataProtheus(NzTxt(rsSP("C6_ENTREG"))))              & """," & _
    """localEntrega"":"  & """" & JsonStr(Left(NzTxt(rsSP("C5_MUN")), 12))                            & """," & _
    """tipoProduto"":"   & """" & JsonStr(vTipoProduto)                                                & """" & _
  "}"

  rsSP.MoveNext
Loop

rsSP.Close : Set rsSP = Nothing

' ─── Resumo ─────────────────────────────────────────────────

Response.Write "],""resumo"":{" & _
  """totalRegistros"":"  & CStr(Num)                                                             & "," & _
  """peso"":{"           & _
    """total"":"  & Replace(CStr(CDbl(Total)),    ",", ".") & "," & _
    """PA"":"     & Replace(CStr(CDbl(TotalPA)),  ",", ".") & "," & _
    """ML"":"     & Replace(CStr(CDbl(TotalML)),  ",", ".") & "," & _
    """CV"":"     & Replace(CStr(CDbl(TotalCV)),  ",", ".") & "," & _
    """CH"":"     & Replace(CStr(CDbl(TotalCH)),  ",", ".") & _
  "}," & _
  """area"":{"           & _
    """total"":"  & Replace(CStr(CDbl(Total2)),    ",", ".") & "," & _
    """PA"":"     & Replace(CStr(CDbl(TotalPAm2)), ",", ".") & "," & _
    """ML"":"     & Replace(CStr(CDbl(TotalMLm2)), ",", ".") & "," & _
    """CV"":"     & Replace(CStr(CDbl(TotalCVm2)), ",", ".") & "," & _
    """CH"":"     & Replace(CStr(CDbl(TotalCHm2)), ",", ".") & _
  "}," & _
  """frota"":{"          & _
    """total"":"         & CStr(Frota)             & "," & _
    """propria"":"       & CStr(FrotaPropria)       & "," & _
    """terceirizada"":"  & CStr(FrotaTerceirizada)  & _
  "}" & _
"}}"
%>
