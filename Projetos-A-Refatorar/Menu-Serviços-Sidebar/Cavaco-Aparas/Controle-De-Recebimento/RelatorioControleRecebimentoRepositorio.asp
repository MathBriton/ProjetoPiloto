<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer   = True
Response.Charset  = "UTF-8"
Response.CodePage = 65001
%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<%
Response.Clear
Response.ContentType  = "application/json"
Response.Expires      = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

Dim CODIGOHELPDESK
CODIGOHELPDESK = Trim(Session("CODIGO"))

Dim acao
acao = Trim(Request("acao"))

If acao = "" Then
  Response.Write "{""sucesso"":false,""mensagem"":""Acao nao informada""}"
  Response.End()
End If

Dim resultado
Select Case acao
  Case "listar"       : resultado = Listar()
  Case "fornecedores" : resultado = ListarFornecedores()
  Case "salvar"       : resultado = Salvar()
  Case "alterar"      : resultado = Alterar()
  Case "apagar"       : resultado = Apagar()
  Case Else           : resultado = "{""sucesso"":false,""mensagem"":""Acao invalida""}"
End Select

Response.Write resultado
Response.End()

Function JsonEscape(str)
  If IsNull(str) Or IsEmpty(str) Then JsonEscape = "" : Exit Function
  Dim r : r = CStr(str)
  r = Replace(r, "\",  "\\")
  r = Replace(r, """", "\""")
  r = Replace(r, vbCrLf, "\n")
  r = Replace(r, vbCr,   "\n")
  r = Replace(r, vbLf,   "\n")
  r = Replace(r, vbTab,  "\t")
  JsonEscape = Trim(r)
End Function

Function FmtData(raw)
  If IsNull(raw) Or IsEmpty(raw) Then FmtData = "" : Exit Function
  Dim s : s = Trim(CStr(raw))
  If s = "" Or s = "0" Then FmtData = "" : Exit Function
  If Len(s) = 8 Then
    FmtData = Mid(s,7,2) & "/" & Mid(s,5,2) & "/" & Left(s,4)
  ElseIf Len(s) >= 10 Then
    FmtData = Mid(s,9,2) & "/" & Mid(s,6,2) & "/" & Left(s,4)
  Else
    FmtData = s
  End If
End Function

Function FmtNum(v)
  If IsNull(v) Or IsEmpty(v) Then FmtNum = "0" : Exit Function
  Dim sV : sV = CStr(v)
  sV = Replace(sV, ",", ".")
  FmtNum = sV
End Function

Function DataParaSQL(s)
  s = Trim(s)
  If s = "" Then DataParaSQL = "" : Exit Function
  DataParaSQL = Left(s,4) & Mid(s,6,2) & Right(s,2)
End Function

Function LoginUser()
  Dim sqlU, rsU
  sqlU = "SELECT CRM_LOGIN FROM CRM_DTI WHERE CRM_D_E_L_E_T = '' AND CRM_R_E_C_N_O = '" & CODIGOHELPDESK & "'"
  Set rsU = Conecta("PIR-HELPDESK", sqlU)
  If IsObject(rsU) Then
    If Not rsU.EOF Then
      LoginUser = Trim(rsU("CRM_LOGIN"))
    End If
    rsU.Close : Set rsU = Nothing
  End If
End Function

Function Listar()
  Dim planta, fornecedor, dtIni, dtFim, SQL

  planta     = Trim(Request("planta"))
  fornecedor = Trim(Request("fornecedor"))
  dtIni      = Trim(Request("datainicial"))
  dtFim      = Trim(Request("datafinal"))

  If dtIni = "" Or dtFim = "" Then
    Listar = "{""sucesso"":false,""mensagem"":""Datas obrigatorias""}"
    Exit Function
  End If

  SQL = "SELECT " & _
          "FILIAL, " & _
          "NF, " & _
          "SERIE, " & _
          "DATA_NF, " & _
          "DATA_RECEBE, " & _
          "HORA_RECEBE, " & _
          "CNPJ_FORNECEDOR, " & _
          "NOME_FORNECEDOR, " & _
          "PRODUTO, " & _
          "PESO_NF, " & _
          "PESO_BALANCA, " & _
          "PERC_UMIDADE, " & _
          "PERC_IMPUREZA, " & _
          "PERC_PROIBITIVO, " & _
          "PERC_OUTROS, " & _
          "LOCAL_ARMAZENA, " & _
          "NUM_FARDOS, " & _
          "FARDOS_ANALIZADOS, " & _
          "PLACA, " & _
          "NOME_MOTORISTA, " & _
          "OBS, " & _
          "QUEM_INCLUIU " & _
        "FROM CONTROLE_RECEBIMENTO " & _
        "WHERE DATA_EXCLUSAO IS NULL " & _
        "AND DATA_RECEBE BETWEEN '" & DataParaSQL(dtIni) & "' AND '" & DataParaSQL(dtFim) & "' "

  If planta <> "" Then
    SQL = SQL & "AND FILIAL = '" & Replace(planta, "'", "''") & "' "
  End If

  If fornecedor <> "" Then
    SQL = SQL & "AND CNPJ_FORNECEDOR = '" & Replace(fornecedor, "'", "''") & "' "
  End If

  SQL = SQL & "ORDER BY DATA_RECEBE + HORA_RECEBE + NF + SERIE DESC"

  Dim RS
  Set RS = Conecta("PIR-CORPORATIVO", SQL)

  If Not IsObject(RS) Then
    Listar = "{""sucesso"":false,""mensagem"":""Erro ao consultar banco de dados""}"
    Exit Function
  End If

  Dim json, primeiro
  json     = "["
  primeiro = True

  Do While Not RS.EOF
    If Not primeiro Then json = json & ","
    json = json & "{" & _
      """filial"":"           & """" & JsonEscape(RS("FILIAL"))            & """," & _
      """nf"":"               & """" & JsonEscape(RS("NF"))                & """," & _
      """serie"":"            & """" & JsonEscape(RS("SERIE"))             & """," & _
      """dataNf"":"           & """" & JsonEscape(FmtData(RS("DATA_NF")))  & """," & _
      """dataRecebe"":"       & """" & JsonEscape(FmtData(RS("DATA_RECEBE"))) & """," & _
      """horaRecebe"":"       & """" & JsonEscape(RS("HORA_RECEBE"))       & """," & _
      """cnpj"":"             & """" & JsonEscape(RS("CNPJ_FORNECEDOR"))   & """," & _
      """fornecedor"":"       & """" & JsonEscape(RS("NOME_FORNECEDOR"))   & """," & _
      """produto"":"          & """" & JsonEscape(RS("PRODUTO"))           & """," & _
      """pesoNf"":"           & """" & FmtNum(RS("PESO_NF"))               & """," & _
      """pesoBalanca"":"      & """" & FmtNum(RS("PESO_BALANCA"))          & """," & _
      """percUmidade"":"      & """" & FmtNum(RS("PERC_UMIDADE"))          & """," & _
      """percImpureza"":"     & """" & FmtNum(RS("PERC_IMPUREZA"))         & """," & _
      """percProibitivo"":"   & """" & FmtNum(RS("PERC_PROIBITIVO"))       & """," & _
      """percOutros"":"       & """" & FmtNum(RS("PERC_OUTROS"))           & """," & _
      """local"":"            & """" & JsonEscape(RS("LOCAL_ARMAZENA"))    & """," & _
      """numFardos"":"        & """" & JsonEscape(RS("NUM_FARDOS"))        & """," & _
      """fardosAnalizados"":" & """" & JsonEscape(RS("FARDOS_ANALIZADOS")) & """," & _
      """placa"":"            & """" & JsonEscape(RS("PLACA"))             & """," & _
      """motorista"":"        & """" & JsonEscape(RS("NOME_MOTORISTA"))    & """," & _
      """obs"":"              & """" & JsonEscape(RS("OBS"))               & """," & _
      """quemIncluiu"":"      & """" & JsonEscape(RS("QUEM_INCLUIU"))      & """" & _
    "}"
    primeiro = False
    RS.MoveNext()
  Loop

  RS.Close() : Set RS = Nothing
  json = json & "]"

  Listar = "{""sucesso"":true,""dados"":" & json & "}"
End Function

Function ListarFornecedores()
  Dim SQL : SQL = "SELECT DISTINCT CNPJ_FORNECEDOR, NOME_FORNECEDOR " & _
                  "FROM CONTROLE_RECEBIMENTO " & _
                  "WHERE DATA_EXCLUSAO IS NULL " & _
                  "ORDER BY NOME_FORNECEDOR"

  Dim RS
  Set RS = Conecta("PIR-CORPORATIVO", SQL)

  If Not IsObject(RS) Then
    ListarFornecedores = "{""sucesso"":false,""mensagem"":""Erro ao carregar fornecedores""}"
    Exit Function
  End If

  Dim json, primeiro
  json     = "["
  primeiro = True

  Do While Not RS.EOF
    If Not primeiro Then json = json & ","
    json = json & "{" & _
      """cnpj"":""" & JsonEscape(RS("CNPJ_FORNECEDOR")) & """," & _
      """nome"":""" & JsonEscape(RS("NOME_FORNECEDOR"))  & """" & _
    "}"
    primeiro = False
    RS.MoveNext()
  Loop

  RS.Close() : Set RS = Nothing
  json = json & "]"

  ListarFornecedores = "{""sucesso"":true,""dados"":" & json & "}"
End Function

Function Salvar()
  Dim edtFILIAL, edtNF, edtSERIE, edtDATA_NF, edtDATA_RECEBE, edtHORA_RECEBE
  Dim edtCNPJ, edtNOME, edtPRODUTO, edtPESO_NF, edtPESO_BAL
  Dim edtUMID, edtIMPUR, edtPROIB, edtOUT, edtLOCAL, edtFARDOS, edtFARD_AN
  Dim edtPLACA, edtMOTOR, edtOBS

  edtFILIAL     = Trim(Request("edtFILIAL"))
  edtNF         = Trim(Request("edtNF"))
  edtSERIE      = Trim(Request("edtSERIE"))
  edtDATA_NF    = Trim(Request("edtDATA_NF"))
  edtDATA_RECEBE = Trim(Request("edtDATA_RECEBE"))
  edtHORA_RECEBE = Trim(Request("edtHORA_RECEBE"))
  edtCNPJ       = Trim(Request("edtCNPJ_FORNECEDOR"))
  edtNOME       = Trim(Request("edtNOME_FORNECEDOR"))
  edtPRODUTO    = Trim(Request("edtPRODUTO"))
  edtPESO_NF    = Replace(Trim(Request("edtPESO_NF")),      ",", ".")
  edtPESO_BAL   = Replace(Trim(Request("edtPESO_BALANCA")), ",", ".")
  edtUMID       = Replace(Trim(Request("edtPERC_UMIDADE")),    ",", ".")
  edtIMPUR      = Replace(Trim(Request("edtPERC_IMPUREZA")),   ",", ".")
  edtPROIB      = Replace(Trim(Request("edtPERC_PROIBITIVO")), ",", ".")
  edtOUT        = Replace(Trim(Request("edtPERC_OUTROS")),     ",", ".")
  edtLOCAL      = Trim(Request("edtLOCAL_ARMAZENA"))
  edtFARDOS     = Trim(Request("edtNUM_FARDOS"))
  edtFARD_AN    = Trim(Request("edtFARDOS_ANALIZADOS"))
  edtPLACA      = Trim(Request("edtPLACA"))
  edtMOTOR      = Trim(Request("edtNOME_MOTORISTA"))
  edtOBS        = Trim(Request("edtOBS"))

  If edtNF = "" Or edtSERIE = "" Then
    Salvar = "{""sucesso"":false,""mensagem"":""NF e Serie sao obrigatorios""}"
    Exit Function
  End If

  Dim sqlChk, rsChk
  sqlChk = "SELECT NF FROM CONTROLE_RECEBIMENTO " & _
           "WHERE NF = '" & Replace(edtNF,"'","''") & "' " & _
           "AND SERIE = '" & Replace(edtSERIE,"'","''") & "' " & _
           "AND FILIAL = '" & Replace(edtFILIAL,"'","''") & "' " & _
           "AND CNPJ_FORNECEDOR = '" & Replace(edtCNPJ,"'","''") & "' " & _
           "AND DATA_EXCLUSAO IS NULL"
  Set rsChk = Conecta("PIR-CORPORATIVO", sqlChk)
  If IsObject(rsChk) Then
    If Not rsChk.EOF Then
      rsChk.Close : Set rsChk = Nothing
      Salvar = "{""sucesso"":false,""mensagem"":""Nota Fiscal ja cadastrada para este fornecedor""}"
      Exit Function
    End If
    rsChk.Close : Set rsChk = Nothing
  End If

  Dim agora, dtInc, hrInc, login
  agora = Now()
  dtInc = Year(agora) & Right("0" & Month(agora),2) & Right("0" & Day(agora),2)
  hrInc = Right("0" & Hour(agora),2) & ":" & Right("0" & Minute(agora),2)
  login = LoginUser()

  Dim sqlIns
  sqlIns = "BEGIN TRANSACTION TRANS; " & _
    "INSERT INTO CONTROLE_RECEBIMENTO (" & _
      "FILIAL,NF,SERIE,DATA_NF,DATA_RECEBE,HORA_RECEBE," & _
      "CNPJ_FORNECEDOR,NOME_FORNECEDOR,PRODUTO," & _
      "PESO_NF,PESO_BALANCA,PERC_UMIDADE,PERC_IMPUREZA,PERC_PROIBITIVO,PERC_OUTROS," & _
      "LOCAL_ARMAZENA,NUM_FARDOS,FARDOS_ANALIZADOS,PLACA,NOME_MOTORISTA,OBS," & _
      "DATA_INCLUSAO,HORA_INCLUSAO,QUEM_INCLUIU" & _
    ") VALUES (" & _
      "'" & Replace(edtFILIAL,"'","''") & "'," & _
      "'" & Replace(edtNF,"'","''") & "'," & _
      "'" & Replace(edtSERIE,"'","''") & "'," & _
      "'" & DataParaSQL(edtDATA_NF) & "'," & _
      "'" & DataParaSQL(edtDATA_RECEBE) & "'," & _
      "'" & Replace(edtHORA_RECEBE,"'","''") & "'," & _
      "'" & Replace(edtCNPJ,"'","''") & "'," & _
      "'" & Replace(edtNOME,"'","''") & "'," & _
      "'" & Replace(edtPRODUTO,"'","''") & "'," & _
      "" & edtPESO_NF & "," & _
      "" & edtPESO_BAL & "," & _
      "" & edtUMID & "," & _
      "" & edtIMPUR & "," & _
      "" & edtPROIB & "," & _
      "" & edtOUT & "," & _
      "'" & Replace(edtLOCAL,"'","''") & "'," & _
      "" & IIf(edtFARDOS = "", "0", edtFARDOS) & "," & _
      "" & IIf(edtFARD_AN = "", "0", edtFARD_AN) & "," & _
      "'" & Replace(edtPLACA,"'","''") & "'," & _
      "'" & Replace(edtMOTOR,"'","''") & "'," & _
      "'" & Replace(edtOBS,"'","''") & "'," & _
      "'" & dtInc & "'," & _
      "'" & hrInc & "'," & _
      "'" & Replace(login,"'","''") & "'" & _
    "); " & _
    "IF @@ERROR <> 0 ROLLBACK TRANSACTION TRANS ELSE COMMIT TRANSACTION TRANS"

  Dim cn : Set cn = Server.CreateObject("ADODB.Connection")
  cn.Open BuscaStringConexao("PIR-CORPORATIVO")
  On Error Resume Next
  Err.Clear
  cn.Execute sqlIns
  Dim numErr : numErr = Err.Number
  Dim descErr : descErr = Err.Description
  On Error GoTo 0
  cn.Close : Set cn = Nothing

  If numErr <> 0 Then
    Salvar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(descErr) & """}"
  Else
    Salvar = "{""sucesso"":true,""mensagem"":""Registro incluido com sucesso""}"
  End If
End Function

Function Alterar()
  Dim edtFILIAL, edtNF, edtSERIE, edtDATA_NF, edtDATA_RECEBE, edtHORA_RECEBE
  Dim edtCNPJ, edtNOME, edtPRODUTO, edtPESO_NF, edtPESO_BAL
  Dim edtUMID, edtIMPUR, edtPROIB, edtOUT, edtLOCAL, edtFARDOS, edtFARD_AN
  Dim edtPLACA, edtMOTOR, edtOBS

  edtFILIAL      = Trim(Request("edtFILIAL"))
  edtNF          = Trim(Request("edtNF"))
  edtSERIE       = Trim(Request("edtSERIE"))
  edtDATA_NF     = Trim(Request("edtDATA_NF"))
  edtDATA_RECEBE = Trim(Request("edtDATA_RECEBE"))
  edtHORA_RECEBE = Trim(Request("edtHORA_RECEBE"))
  edtCNPJ        = Trim(Request("edtCNPJ_FORNECEDOR"))
  edtNOME        = Trim(Request("edtNOME_FORNECEDOR"))
  edtPRODUTO     = Trim(Request("edtPRODUTO"))
  edtPESO_NF     = Replace(Trim(Request("edtPESO_NF")),      ",", ".")
  edtPESO_BAL    = Replace(Trim(Request("edtPESO_BALANCA")), ",", ".")
  edtUMID        = Replace(Trim(Request("edtPERC_UMIDADE")),    ",", ".")
  edtIMPUR       = Replace(Trim(Request("edtPERC_IMPUREZA")),   ",", ".")
  edtPROIB       = Replace(Trim(Request("edtPERC_PROIBITIVO")), ",", ".")
  edtOUT         = Replace(Trim(Request("edtPERC_OUTROS")),     ",", ".")
  edtLOCAL       = Trim(Request("edtLOCAL_ARMAZENA"))
  edtFARDOS      = Trim(Request("edtNUM_FARDOS"))
  edtFARD_AN     = Trim(Request("edtFARDOS_ANALIZADOS"))
  edtPLACA       = Trim(Request("edtPLACA"))
  edtMOTOR       = Trim(Request("edtNOME_MOTORISTA"))
  edtOBS         = Trim(Request("edtOBS"))

  If edtNF = "" Or edtSERIE = "" Then
    Alterar = "{""sucesso"":false,""mensagem"":""NF e Serie sao obrigatorios""}"
    Exit Function
  End If

  Dim agora, dtAlt, hrAlt, login
  agora = Now()
  dtAlt = Year(agora) & Right("0" & Month(agora),2) & Right("0" & Day(agora),2)
  hrAlt = Right("0" & Hour(agora),2) & ":" & Right("0" & Minute(agora),2)
  login = LoginUser()

  Dim sqlAlt
  sqlAlt = "BEGIN TRANSACTION TRANS; " & _
    "UPDATE CONTROLE_RECEBIMENTO SET " & _
      "DATA_NF         = '" & DataParaSQL(edtDATA_NF) & "'," & _
      "DATA_RECEBE     = '" & DataParaSQL(edtDATA_RECEBE) & "'," & _
      "HORA_RECEBE     = '" & Replace(edtHORA_RECEBE,"'","''") & "'," & _
      "CNPJ_FORNECEDOR = '" & Replace(edtCNPJ,"'","''") & "'," & _
      "NOME_FORNECEDOR = '" & Replace(edtNOME,"'","''") & "'," & _
      "PRODUTO         = '" & Replace(edtPRODUTO,"'","''") & "'," & _
      "PESO_NF         = "  & edtPESO_NF & "," & _
      "PESO_BALANCA    = "  & edtPESO_BAL & "," & _
      "PERC_UMIDADE    = "  & edtUMID & "," & _
      "PERC_IMPUREZA   = "  & edtIMPUR & "," & _
      "PERC_PROIBITIVO = "  & edtPROIB & "," & _
      "PERC_OUTROS     = "  & edtOUT & "," & _
      "LOCAL_ARMAZENA  = '" & Replace(edtLOCAL,"'","''") & "'," & _
      "NUM_FARDOS      = "  & IIf(edtFARDOS = "", "0", edtFARDOS) & "," & _
      "FARDOS_ANALIZADOS = " & IIf(edtFARD_AN = "", "0", edtFARD_AN) & "," & _
      "PLACA           = '" & Replace(edtPLACA,"'","''") & "'," & _
      "NOME_MOTORISTA  = '" & Replace(edtMOTOR,"'","''") & "'," & _
      "OBS             = '" & Replace(edtOBS,"'","''") & "'," & _
      "DATA_ALTERACAO  = '" & dtAlt & "'," & _
      "HORA_ALTERACAO  = '" & hrAlt & "'," & _
      "QUEM_ALTEROU    = '" & Replace(login,"'","''") & "' " & _
    "WHERE NF = '" & Replace(edtNF,"'","''") & "' " & _
    "AND SERIE = '" & Replace(edtSERIE,"'","''") & "' " & _
    "AND FILIAL = '" & Replace(edtFILIAL,"'","''") & "'; " & _
    "IF @@ERROR <> 0 ROLLBACK TRANSACTION TRANS ELSE COMMIT TRANSACTION TRANS"

  Dim cn : Set cn = Server.CreateObject("ADODB.Connection")
  cn.Open BuscaStringConexao("PIR-CORPORATIVO")
  On Error Resume Next
  Err.Clear
  cn.Execute sqlAlt
  Dim numErr : numErr = Err.Number
  Dim descErr : descErr = Err.Description
  On Error GoTo 0
  cn.Close : Set cn = Nothing

  If numErr <> 0 Then
    Alterar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(descErr) & """}"
  Else
    Alterar = "{""sucesso"":true,""mensagem"":""Registro alterado com sucesso""}"
  End If
End Function

Function Apagar()
  Dim edtFILIAL, edtNF, edtSERIE
  edtFILIAL = Trim(Request("edtFILIAL"))
  edtNF     = Trim(Request("edtNF"))
  edtSERIE  = Trim(Request("edtSERIE"))

  If edtNF = "" Or edtSERIE = "" Then
    Apagar = "{""sucesso"":false,""mensagem"":""NF e Serie sao obrigatorios""}"
    Exit Function
  End If

  Dim agora, dtExc, hrExc, login
  agora = Now()
  dtExc = Year(agora) & Right("0" & Month(agora),2) & Right("0" & Day(agora),2)
  hrExc = Right("0" & Hour(agora),2) & ":" & Right("0" & Minute(agora),2)
  login = LoginUser()

  Dim sqlDel
  sqlDel = "BEGIN TRANSACTION TRANS; " & _
    "UPDATE CONTROLE_RECEBIMENTO SET " & _
      "DATA_EXCLUSAO = '" & dtExc & "'," & _
      "HORA_EXCLUSAO = '" & hrExc & "'," & _
      "QUEM_EXCLUIU  = '" & Replace(login,"'","''") & "' " & _
    "WHERE NF = '" & Replace(edtNF,"'","''") & "' " & _
    "AND SERIE = '" & Replace(edtSERIE,"'","''") & "' " & _
    "AND FILIAL = '" & Replace(edtFILIAL,"'","''") & "'; " & _
    "IF @@ERROR <> 0 ROLLBACK TRANSACTION TRANS ELSE COMMIT TRANSACTION TRANS"

  Dim cn : Set cn = Server.CreateObject("ADODB.Connection")
  cn.Open BuscaStringConexao("PIR-CORPORATIVO")
  On Error Resume Next
  Err.Clear
  cn.Execute sqlDel
  Dim numErr : numErr = Err.Number
  Dim descErr : descErr = Err.Description
  On Error GoTo 0
  cn.Close : Set cn = Nothing

  If numErr <> 0 Then
    Apagar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(descErr) & """}"
  Else
    Apagar = "{""sucesso"":true,""mensagem"":""Registro apagado com sucesso""}"
  End If
End Function
%>
