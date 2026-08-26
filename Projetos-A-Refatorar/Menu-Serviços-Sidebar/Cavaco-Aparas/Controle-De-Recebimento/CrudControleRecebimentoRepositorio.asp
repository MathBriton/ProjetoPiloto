<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.ContentType = "application/json"
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!--#include file="../../cnn2/cnn2.asp"-->
<%
Dim USR, UNIDADE_FILIAL
USR            = Trim(Session("USR"))
UNIDADE_FILIAL = Trim(Session("FILIAL"))

If USR = "" Or UNIDADE_FILIAL = "" Then
    Response.Clear
    Response.Write "{""sucesso"":false,""mensagem"":""Sessão expirada""}"
    Response.End()
End If

Dim acao, resultado
acao = Trim(Request("acao"))

If acao = "" Then
    Response.Write "{""sucesso"":false,""mensagem"":""Ação não informada""}"
    Response.End()
End If

Select Case acao
    Case "listar"               : resultado = Listar()
    Case "nivel"                : resultado = ObterNivel()
    Case "buscar_registro"      : resultado = BuscarRegistro()
    Case "salvar"               : resultado = Salvar()
    Case "apagar"               : resultado = Apagar()
    Case "cadastrar_fornecedor" : resultado = CadastrarFornecedor()
    Case "buscar_fornecedores"  : resultado = BuscarFornecedores()
    Case "verificar_nf"         : resultado = VerificarNF()
    Case Else                   : resultado = "{""sucesso"":false,""mensagem"":""Ação inválida""}"
End Select

Response.Write resultado

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

Function SafeNum(v)
    If IsNull(v) Or IsEmpty(v) Then SafeNum = "0" : Exit Function
    Dim s : s = Trim(CStr(v))
    If s = "" Then SafeNum = "0" : Exit Function
    SafeNum = s
End Function

Function SafeNumDec(v, decimais)
    If IsNull(v) Or IsEmpty(v) Then
        SafeNumDec = "0." & String(decimais, "0")
        Exit Function
    End If
    Dim n
    On Error Resume Next
    n = CDbl(v)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        SafeNumDec = "0." & String(decimais, "0")
        Exit Function
    End If
    On Error GoTo 0
    SafeNumDec = Replace(FormatNumber(n, decimais, -1, 0, 0), ",", ".")
End Function

Function FormatarData8(data)
    If IsNull(data) Or IsEmpty(data) Then FormatarData8 = "" : Exit Function
    Dim s : s = Trim(CStr(data))
    If s = "" Or Len(s) < 8 Then FormatarData8 = "" : Exit Function
    FormatarData8 = Right("00" & Mid(s,7,2), 2) & "/" & Right("00" & Mid(s,5,2), 2) & "/" & Left(s,4)
End Function

Function NormalizaDecimal(v)
    If IsNull(v) Or IsEmpty(v) Then NormalizaDecimal = "0" : Exit Function
    Dim s : s = Trim(CStr(v))
    If s = "" Then NormalizaDecimal = "0" : Exit Function
    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")
    If s = "" Or s = "." Then NormalizaDecimal = "0" : Exit Function
    NormalizaDecimal = s
End Function

Function CNumero(v)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v)) = "" Then CNumero = CDbl(0) : Exit Function
    Dim s : s = Replace(Trim(CStr(v)), ".", ",")
    On Error Resume Next
    CNumero = CDbl(s)
    If Err.Number <> 0 Then Err.Clear : CNumero = CDbl(0)
    On Error GoTo 0
End Function

Function SafeInt(v)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v)) = "" Then SafeInt = 0 : Exit Function
    If IsNumeric(v) Then SafeInt = CInt(v) Else SafeInt = 0
End Function

Function ObterNivel()
    Dim sqlN, rsN, nivel
    sqlN = "SELECT NivelControladoria AS NIVEL FROM CRM_DTI WHERE CRM_LOGIN = '" & USR & "' AND CRM_D_E_L_E_T = ''"

    On Error Resume Next
    Response.Clear
    Set rsN = Conecta("PIR-HELPDESK", sqlN)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ObterNivel = "{""sucesso"":false,""mensagem"":""Erro de conexão""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsN) Then
        ObterNivel = "{""sucesso"":false,""mensagem"":""Erro ao obter nível""}"
        Exit Function
    End If

    nivel = 0
    If Not rsN.EOF Then
        If rsN("NIVEL") Then nivel = 1
    End If
    rsN.Close() : Set rsN = Nothing
    ObterNivel = "{""sucesso"":true,""nivel"":" & nivel & "}"
End Function

Function Listar()
    Dim sqlL, rsL, qtd, filtroNF, filtroForn, filtroProd, filtroDataIni, filtroDataFim
    Dim primeiroItem, jsonArr

    qtd           = Trim(Request("qtd"))
    filtroNF      = Trim(Request("nf"))
    filtroForn    = Trim(Request("fornecedor"))
    filtroProd    = Trim(Request("produto"))
    filtroDataIni = Trim(Request("dataini"))
    filtroDataFim = Trim(Request("datafim"))

    If qtd = "" Or Not IsNumeric(qtd) Then qtd = 50

    sqlL = "SELECT TOP " & CInt(qtd) & " " & _
           "FILIAL, NF, SERIE, DATA_NF, DATA_RECEBE, HORA_RECEBE, " & _
           "CNPJ_FORNECEDOR, NOME_FORNECEDOR, PRODUTO, " & _
           "PESO_NF, PESO_BALANCA, PERC_UMIDADE, PERC_IMPUREZA, " & _
           "PERC_PROIBITIVO, PERC_OUTROS, LOCAL_ARMAZENA, " & _
           "NUM_FARDOS, FARDOS_ANALIZADOS, PLACA, NOME_MOTORISTA, OBS, DATA_INCLUSAO " & _
           "FROM CONTROLE_RECEBIMENTO " & _
           "WHERE FILIAL = '" & UNIDADE_FILIAL & "' " & _
           "AND DATA_EXCLUSAO IS NULL"

    If filtroNF <> "" Then
        sqlL = sqlL & " AND NF LIKE '%" & filtroNF & "%'"
    End If
    If filtroForn <> "" Then
        sqlL = sqlL & " AND (CNPJ_FORNECEDOR LIKE '%" & filtroForn & "%' OR NOME_FORNECEDOR LIKE '%" & filtroForn & "%')"
    End If
    If filtroProd <> "" Then
        sqlL = sqlL & " AND PRODUTO LIKE '%" & filtroProd & "%'"
    End If
    If filtroDataIni <> "" Then
        sqlL = sqlL & " AND DATA_INCLUSAO >= '" & Replace(filtroDataIni, "-", "") & "'"
    End If
    If filtroDataFim <> "" Then
        sqlL = sqlL & " AND DATA_INCLUSAO <= '" & Replace(filtroDataFim, "-", "") & "'"
    End If

    sqlL = sqlL & " ORDER BY DATA_RECEBE DESC, HORA_RECEBE DESC, NF DESC, SERIE DESC"

    On Error Resume Next
    Response.Clear
    Set rsL = Conecta("PIR-CORPORATIVO", sqlL)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Listar = "{""sucesso"":false,""mensagem"":""Erro de conexão""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsL) Then
        Listar = "{""sucesso"":false,""mensagem"":""Erro ao consultar registros""}"
        Exit Function
    End If

    jsonArr      = "["
    primeiroItem = True
    Do While Not rsL.EOF
        If Not primeiroItem Then jsonArr = jsonArr & ","
        jsonArr = jsonArr & "{" & _
            """filial"":"""      & JsonEscape(rsL("FILIAL"))            & """," & _
            """nf"":"""          & JsonEscape(rsL("NF"))                & """," & _
            """serie"":"""       & JsonEscape(rsL("SERIE"))             & """," & _
            """dataNf"":"""      & FormatarData8(rsL("DATA_NF"))        & """," & _
            """dataRecebe"":"""  & FormatarData8(rsL("DATA_RECEBE"))    & """," & _
            """horaRecebe"":"""  & JsonEscape(rsL("HORA_RECEBE"))       & """," & _
            """cnpj"":"""        & JsonEscape(rsL("CNPJ_FORNECEDOR"))   & """," & _
            """nomeForn"":"""    & JsonEscape(rsL("NOME_FORNECEDOR"))   & """," & _
            """produto"":"""     & JsonEscape(rsL("PRODUTO"))           & """," & _
            """pesoNf"":"""      & SafeNumDec(rsL("PESO_NF"), 3)        & """," & _
            """pesoBal"":"""     & SafeNumDec(rsL("PESO_BALANCA"), 3)   & """," & _
            """umidade"":"""     & SafeNumDec(rsL("PERC_UMIDADE"), 3)   & """," & _
            """impureza"":"""    & SafeNumDec(rsL("PERC_IMPUREZA"), 3)  & """," & _
            """proibitivo"":"""  & SafeNumDec(rsL("PERC_PROIBITIVO"), 3)& """," & _
            """outros"":"""      & SafeNumDec(rsL("PERC_OUTROS"), 3)    & """," & _
            """local"":"""       & JsonEscape(rsL("LOCAL_ARMAZENA"))    & """," & _
            """numFardos"":"""   & SafeNum(rsL("NUM_FARDOS"))           & """," & _
            """fardosAnal"":"""  & SafeNum(rsL("FARDOS_ANALIZADOS"))    & """," & _
            """placa"":"""       & JsonEscape(rsL("PLACA"))             & """," & _
            """motorista"":"""   & JsonEscape(rsL("NOME_MOTORISTA"))    & """," & _
            """obs"":"""         & JsonEscape(rsL("OBS"))               & """," & _
            """dataInclusao"":"""& FormatarData8(rsL("DATA_INCLUSAO"))  & """" & _
            "}"
        primeiroItem = False
        rsL.MoveNext()
    Loop
    jsonArr = jsonArr & "]"

    rsL.Close() : Set rsL = Nothing
    Listar = "{""sucesso"":true,""dados"":" & jsonArr & "}"
End Function

Function BuscarRegistro()
    Dim sqlB, rsB
    Dim vNF, vSERIE
    vNF    = Trim(Request("nf"))
    vSERIE = Trim(Request("serie"))

    sqlB = "SELECT TOP 1 FILIAL, NF, SERIE, DATA_NF, DATA_RECEBE, HORA_RECEBE, " & _
           "CNPJ_FORNECEDOR, NOME_FORNECEDOR, PRODUTO, PESO_NF, PESO_BALANCA, " & _
           "PERC_UMIDADE, PERC_IMPUREZA, PERC_PROIBITIVO, PERC_OUTROS, " & _
           "LOCAL_ARMAZENA, NUM_FARDOS, FARDOS_ANALIZADOS, PLACA, NOME_MOTORISTA, OBS " & _
           "FROM CONTROLE_RECEBIMENTO " & _
           "WHERE ISNULL(DATA_EXCLUSAO, 0) = 0 " & _
           "AND FILIAL = '" & UNIDADE_FILIAL & "' " & _
           "AND NF = '" & vNF & "' " & _
           "AND SERIE = '" & vSERIE & "'"

    On Error Resume Next
    Response.Clear
    Set rsB = Conecta("PIR-CORPORATIVO", sqlB)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        BuscarRegistro = "{""sucesso"":false,""mensagem"":""Erro de conexão""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsB) Then
        BuscarRegistro = "{""sucesso"":false,""mensagem"":""Registro não encontrado""}"
        Exit Function
    End If
    If rsB.EOF Then
        rsB.Close() : Set rsB = Nothing
        BuscarRegistro = "{""sucesso"":false,""mensagem"":""Registro não encontrado""}"
        Exit Function
    End If

    Dim dataFimNF, dataFimRec
    dataFimNF  = ""
    dataFimRec = ""
    If Not IsNull(rsB("DATA_NF")) And Not IsEmpty(rsB("DATA_NF")) Then
        If Len(Trim(CStr(rsB("DATA_NF")))) >= 8 Then
            dataFimNF = Left(rsB("DATA_NF"),4) & "-" & Mid(rsB("DATA_NF"),5,2) & "-" & Mid(rsB("DATA_NF"),7,2)
        End If
    End If
    If Not IsNull(rsB("DATA_RECEBE")) And Not IsEmpty(rsB("DATA_RECEBE")) Then
        If Len(Trim(CStr(rsB("DATA_RECEBE")))) >= 8 Then
            dataFimRec = Left(rsB("DATA_RECEBE"),4) & "-" & Mid(rsB("DATA_RECEBE"),5,2) & "-" & Mid(rsB("DATA_RECEBE"),7,2)
        End If
    End If

    BuscarRegistro = "{""sucesso"":true,""dado"":{" & _
        """filial"":"""     & JsonEscape(rsB("FILIAL"))              & """," & _
        """nf"":"""         & JsonEscape(rsB("NF"))                  & """," & _
        """serie"":"""      & JsonEscape(rsB("SERIE"))               & """," & _
        """dataNf"":"""     & dataFimNF                              & """," & _
        """dataRecebe"":""" & dataFimRec                             & """," & _
        """horaRecebe"":""" & JsonEscape(rsB("HORA_RECEBE"))         & """," & _
        """cnpj"":"""       & JsonEscape(rsB("CNPJ_FORNECEDOR"))     & """," & _
        """nomeForn"":"""   & JsonEscape(rsB("NOME_FORNECEDOR"))     & """," & _
        """produto"":"""    & JsonEscape(rsB("PRODUTO"))             & """," & _
        """pesoNf"":"""     & SafeNumDec(rsB("PESO_NF"), 3)          & """," & _
        """pesoBal"":"""    & SafeNumDec(rsB("PESO_BALANCA"), 3)     & """," & _
        """umidade"":"""    & SafeNumDec(rsB("PERC_UMIDADE"), 3)     & """," & _
        """impureza"":"""   & SafeNumDec(rsB("PERC_IMPUREZA"), 3)    & """," & _
        """proibitivo"":""" & SafeNumDec(rsB("PERC_PROIBITIVO"), 3)  & """," & _
        """outros"":"""     & SafeNumDec(rsB("PERC_OUTROS"), 3)      & """," & _
        """local"":"""      & JsonEscape(rsB("LOCAL_ARMAZENA"))      & """," & _
        """numFardos"":"""  & SafeNum(rsB("NUM_FARDOS"))             & """," & _
        """fardosAnal"":""" & SafeNum(rsB("FARDOS_ANALIZADOS"))      & """," & _
        """placa"":"""      & JsonEscape(rsB("PLACA"))               & """," & _
        """motorista"":"""  & JsonEscape(rsB("NOME_MOTORISTA"))      & """," & _
        """obs"":"""        & JsonEscape(rsB("OBS"))                 & """" & _
        "}}"

    rsB.Close() : Set rsB = Nothing
End Function

Function VerificarNF()
    Dim vNF, vSERIE, vCNPJ, sqlV, rsV

    vNF    = Trim(Request("nf"))
    vSERIE = Trim(Request("serie"))
    vCNPJ  = Trim(Request("cnpj"))

    If vNF = "" Or vSERIE = "" Or vCNPJ = "" Then
        VerificarNF = "{""sucesso"":false,""mensagem"":""Parâmetros insuficientes""}"
        Exit Function
    End If

    sqlV = "SELECT TOP 1 NF FROM CONTROLE_RECEBIMENTO " & _
           "WHERE FILIAL = '" & UNIDADE_FILIAL & "' " & _
           "AND NF = '" & vNF & "' " & _
           "AND SERIE = '" & vSERIE & "' " & _
           "AND CNPJ_FORNECEDOR = '" & vCNPJ & "' " & _
           "AND DATA_EXCLUSAO IS NULL"

    On Error Resume Next
    Set rsV = Conecta("PIR-CORPORATIVO", sqlV)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        VerificarNF = "{""sucesso"":false,""mensagem"":""Erro de conexão""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsV) Then
        VerificarNF = "{""sucesso"":false,""mensagem"":""Erro ao verificar NF""}"
        Exit Function
    End If

    Dim existe : existe = Not rsV.EOF
    rsV.Close() : Set rsV = Nothing

    If existe Then
        VerificarNF = "{""sucesso"":true,""existe"":true}"
    Else
        VerificarNF = "{""sucesso"":true,""existe"":false}"
    End If
End Function

Function Salvar()
    Dim vAcao, oConn, oCmd
    Dim vNF, vSERIE, vDATA_NF, vDATA_RECEBE, vHORA_RECEBE
    Dim vCNPJ, vNOME_FORN, vPRODUTO, vPESO_NF, vPESO_BAL
    Dim vUMIDADE, vIMPUREZA, vPROIBITIVO, vOUTROS
    Dim vLOCAL, vNUM_FARDOS, vFARDOS, vPLACA, vMOTORISTA, vOBS
    Dim errDesc

    vAcao        = Trim(Request("tipoAcao"))
    vNF          = Left(Trim(Request("nf")), 6)
    vSERIE       = Trim(Request("serie"))
    vDATA_NF     = Trim(Request("dataNf"))
    vDATA_RECEBE = Trim(Request("dataRecebe"))
    vHORA_RECEBE = Trim(Request("horaRecebe"))
    If Len(vHORA_RECEBE) > 5 Then vHORA_RECEBE = Left(vHORA_RECEBE, 5)
    If Len(UNIDADE_FILIAL) > 2 Then UNIDADE_FILIAL = Left(UNIDADE_FILIAL, 2)
    If Len(USR) > 20 Then USR = Left(USR, 20)
    vCNPJ        = Trim(Request("cnpj"))
    vNOME_FORN   = Trim(Request("nomeForn"))
    vPRODUTO     = Trim(Request("produto"))
    vPESO_NF     = Trim(Request("pesoNf"))
    vPESO_BAL    = Trim(Request("pesoBal"))
    vUMIDADE     = Trim(Request("umidade"))
    vIMPUREZA    = Trim(Request("impureza"))
    vPROIBITIVO  = Trim(Request("proibitivo"))
    vOUTROS      = Trim(Request("outros"))
    vLOCAL       = Trim(Request("local"))
    vNUM_FARDOS  = Trim(Request("numFardos"))
    vFARDOS      = Trim(Request("fardosAnal"))
    vPLACA       = Trim(Request("placa"))
    vMOTORISTA   = Trim(Request("motorista"))
    vOBS         = Trim(Request("obs"))

    If vNF = "" Or vSERIE = "" Or vCNPJ = "" Or vPRODUTO = "" Or vMOTORISTA = "" Or vPLACA = "" Then
        Salvar = "{""sucesso"":false,""mensagem"":""Campos obrigatórios não preenchidos""}"
        Exit Function
    End If

    Dim vDataNFNum, vDataRecNum
    vDataNFNum  = Replace(vDATA_NF, "-", "")
    vDataRecNum = Replace(vDATA_RECEBE, "-", "")

    vPESO_NF  = NormalizaDecimal(vPESO_NF)
    vPESO_BAL = NormalizaDecimal(vPESO_BAL)

    Dim numErrConn, errDescConn
    On Error Resume Next
    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open BuscaStringConexao("PIR-CORPORATIVO")
    numErrConn  = Err.Number
    errDescConn = Err.Description
    On Error GoTo 0
    If numErrConn <> 0 Then
        Set oConn = Nothing
        Salvar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(errDescConn) & """}"
        Exit Function
    End If

    Set oCmd  = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandType = 1

    If vAcao = "INS" Then
        oCmd.CommandText = "INSERT INTO CONTROLE_RECEBIMENTO " & _
            "(FILIAL, NF, SERIE, DATA_NF, DATA_RECEBE, HORA_RECEBE, " & _
            "CNPJ_FORNECEDOR, NOME_FORNECEDOR, PRODUTO, " & _
            "PESO_NF, PESO_BALANCA, PERC_UMIDADE, PERC_IMPUREZA, " & _
            "PERC_PROIBITIVO, PERC_OUTROS, LOCAL_ARMAZENA, " & _
            "NUM_FARDOS, FARDOS_ANALIZADOS, PLACA, NOME_MOTORISTA, OBS, QUEM_INCLUIU) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        oCmd.Parameters.Append oCmd.CreateParameter("@FILIAL",      200, 1, 10,   UNIDADE_FILIAL)
        oCmd.Parameters.Append oCmd.CreateParameter("@NF",          200, 1, 20,  vNF)
        oCmd.Parameters.Append oCmd.CreateParameter("@SERIE",       200, 1, 3,   vSERIE)
        oCmd.Parameters.Append oCmd.CreateParameter("@DATA_NF",     200, 1, 8,   vDataNFNum)
        oCmd.Parameters.Append oCmd.CreateParameter("@DATA_RECEBE", 200, 1, 8,   vDataRecNum)
        oCmd.Parameters.Append oCmd.CreateParameter("@HORA_RECEBE", 200, 1, 5,   vHORA_RECEBE)
        oCmd.Parameters.Append oCmd.CreateParameter("@CNPJ",        200, 1, 14,  vCNPJ)
        oCmd.Parameters.Append oCmd.CreateParameter("@NOME_FORN",   200, 1, 300, vNOME_FORN)
        oCmd.Parameters.Append oCmd.CreateParameter("@PRODUTO",     200, 1, 100, vPRODUTO)
        oCmd.Parameters.Append oCmd.CreateParameter("@PESO_NF",     5,   1, ,    CNumero(vPESO_NF))
        oCmd.Parameters.Append oCmd.CreateParameter("@PESO_BAL",    5,   1, ,    CNumero(vPESO_BAL))
        oCmd.Parameters.Append oCmd.CreateParameter("@UMIDADE",     5,   1, ,    CNumero(vUMIDADE))
        oCmd.Parameters.Append oCmd.CreateParameter("@IMPUREZA",    5,   1, ,    CNumero(vIMPUREZA))
        oCmd.Parameters.Append oCmd.CreateParameter("@PROIBITIVO",  5,   1, ,    CNumero(vPROIBITIVO))
        oCmd.Parameters.Append oCmd.CreateParameter("@OUTROS",      5,   1, ,    CNumero(vOUTROS))
        oCmd.Parameters.Append oCmd.CreateParameter("@LOCAL",       200, 1, 10,  vLOCAL)
        oCmd.Parameters.Append oCmd.CreateParameter("@NUM_FARDOS",  3,   1, ,    SafeInt(vNUM_FARDOS))
        oCmd.Parameters.Append oCmd.CreateParameter("@FARDOS",      3,   1, ,    SafeInt(vFARDOS))
        oCmd.Parameters.Append oCmd.CreateParameter("@PLACA",       200, 1, 10,  vPLACA)
        oCmd.Parameters.Append oCmd.CreateParameter("@MOTORISTA",   200, 1, 300, vMOTORISTA)
        oCmd.Parameters.Append oCmd.CreateParameter("@OBS",         200, 1, 300, vOBS)
        oCmd.Parameters.Append oCmd.CreateParameter("@QUEM",        200, 1, 60,  USR)
    ElseIf vAcao = "ALT" Then
        oCmd.CommandText = "UPDATE CONTROLE_RECEBIMENTO SET " & _
            "DATA_NF=?, DATA_RECEBE=?, HORA_RECEBE=?, " & _
            "CNPJ_FORNECEDOR=?, NOME_FORNECEDOR=?, PRODUTO=?, " & _
            "PESO_NF=?, PESO_BALANCA=?, PERC_UMIDADE=?, PERC_IMPUREZA=?, " & _
            "PERC_PROIBITIVO=?, PERC_OUTROS=?, LOCAL_ARMAZENA=?, " & _
            "NUM_FARDOS=?, FARDOS_ANALIZADOS=?, PLACA=?, NOME_MOTORISTA=?, OBS=? " & _
            "WHERE FILIAL=? AND NF=? AND SERIE=? AND DATA_EXCLUSAO IS NULL"
        oCmd.Parameters.Append oCmd.CreateParameter("@DATA_NF",     200, 1, 8,   vDataNFNum)
        oCmd.Parameters.Append oCmd.CreateParameter("@DATA_RECEBE", 200, 1, 8,   vDataRecNum)
        oCmd.Parameters.Append oCmd.CreateParameter("@HORA",        200, 1, 5,   vHORA_RECEBE)
        oCmd.Parameters.Append oCmd.CreateParameter("@CNPJ",        200, 1, 14,  vCNPJ)
        oCmd.Parameters.Append oCmd.CreateParameter("@NOME_FORN",   200, 1, 300, vNOME_FORN)
        oCmd.Parameters.Append oCmd.CreateParameter("@PRODUTO",     200, 1, 100, vPRODUTO)
        oCmd.Parameters.Append oCmd.CreateParameter("@PESO_NF",     5,   1, ,    CNumero(vPESO_NF))
        oCmd.Parameters.Append oCmd.CreateParameter("@PESO_BAL",    5,   1, ,    CNumero(vPESO_BAL))
        oCmd.Parameters.Append oCmd.CreateParameter("@UMIDADE",     5,   1, ,    CNumero(vUMIDADE))
        oCmd.Parameters.Append oCmd.CreateParameter("@IMPUREZA",    5,   1, ,    CNumero(vIMPUREZA))
        oCmd.Parameters.Append oCmd.CreateParameter("@PROIBITIVO",  5,   1, ,    CNumero(vPROIBITIVO))
        oCmd.Parameters.Append oCmd.CreateParameter("@OUTROS",      5,   1, ,    CNumero(vOUTROS))
        oCmd.Parameters.Append oCmd.CreateParameter("@LOCAL",       200, 1, 10,  vLOCAL)
        oCmd.Parameters.Append oCmd.CreateParameter("@NUM_FARDOS",  3,   1, ,    SafeInt(vNUM_FARDOS))
        oCmd.Parameters.Append oCmd.CreateParameter("@FARDOS",      3,   1, ,    SafeInt(vFARDOS))
        oCmd.Parameters.Append oCmd.CreateParameter("@PLACA",       200, 1, 10,  vPLACA)
        oCmd.Parameters.Append oCmd.CreateParameter("@MOTORISTA",   200, 1, 300, vMOTORISTA)
        oCmd.Parameters.Append oCmd.CreateParameter("@OBS",         200, 1, 300, vOBS)
        oCmd.Parameters.Append oCmd.CreateParameter("@FILIAL",      200, 1, 10,   UNIDADE_FILIAL)
        oCmd.Parameters.Append oCmd.CreateParameter("@NF",          200, 1, 20,  vNF)
        oCmd.Parameters.Append oCmd.CreateParameter("@SERIE",       200, 1, 3,   vSERIE)
    ElseIf vAcao = "APB" Then
        oCmd.CommandText = "UPDATE CONTROLE_RECEBIMENTO SET PESO_BALANCA=? " & _
                           "WHERE FILIAL=? AND NF=? AND SERIE=? AND DATA_EXCLUSAO IS NULL"
        oCmd.Parameters.Append oCmd.CreateParameter("@PESO_BAL", 5,   1, ,  CNumero(vPESO_BAL))
        oCmd.Parameters.Append oCmd.CreateParameter("@FILIAL",   200, 1, 10,  UNIDADE_FILIAL)
        oCmd.Parameters.Append oCmd.CreateParameter("@NF",       200, 1, 20, vNF)
        oCmd.Parameters.Append oCmd.CreateParameter("@SERIE",    200, 1, 3,  vSERIE)
    Else
        Set oCmd = Nothing : oConn.Close() : Set oConn = Nothing
        Salvar = "{""sucesso"":false,""mensagem"":""Tipo de ação desconhecido""}"
        Exit Function
    End If

    On Error Resume Next
    oCmd.Execute()
    errDesc = Err.Description
    Dim numErr : numErr = Err.Number
    On Error GoTo 0

    Set oCmd = Nothing : oConn.Close() : Set oConn = Nothing

    If numErr <> 0 Then
        Salvar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(errDesc) & """}"
    Else
        Dim msgSalvar
        Select Case vAcao
        Case "INS" : msgSalvar = "Registro incluído com sucesso!"
        Case "ALT" : msgSalvar = "Registro alterado com sucesso!"
        Case "APB" : msgSalvar = "Peso da balança atualizado com sucesso!"
        End Select
        Salvar = "{""sucesso"":true,""mensagem"":""" & msgSalvar & """}"
    End If
End Function

Function Apagar()
    Dim vNF, vSERIE, oConn, oCmd, errDesc, numErr
    vNF    = Trim(Request("nf"))
    vSERIE = Trim(Request("serie"))

    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open BuscaStringConexao("PIR-CORPORATIVO")
    Set oCmd  = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandType = 1
    oCmd.CommandText = "UPDATE CONTROLE_RECEBIMENTO SET DATA_EXCLUSAO = GETDATE() " & _
                       "WHERE FILIAL=? AND NF=? AND SERIE=? AND DATA_EXCLUSAO IS NULL"
    oCmd.Parameters.Append oCmd.CreateParameter("@FILIAL", 200, 1, 10,  UNIDADE_FILIAL)
    oCmd.Parameters.Append oCmd.CreateParameter("@NF",     200, 1, 20, vNF)
    oCmd.Parameters.Append oCmd.CreateParameter("@SERIE",  200, 1, 3,  vSERIE)

    On Error Resume Next
    oCmd.Execute()
    errDesc = Err.Description
    numErr  = Err.Number
    On Error GoTo 0

    Set oCmd = Nothing : oConn.Close() : Set oConn = Nothing

    If numErr <> 0 Then
        Apagar = "{""sucesso"":false,""mensagem"":""" & JsonEscape(errDesc) & """}"
    Else
        Apagar = "{""sucesso"":true,""mensagem"":""Registro apagado com sucesso!""}"
    End If
End Function

Function CadastrarFornecedor()
    Dim sCnpj, sCodigo, sNome
    Dim oConn, oCmd, rsV, numErr, errDesc

    sCnpj   = Trim(Request("cnpj"))
    sCodigo = Trim(Request("codigo"))
    sNome   = Trim(Request("nome"))

    If sCnpj = "" Or sCodigo = "" Or sNome = "" Then
        CadastrarFornecedor = "{""sucesso"":false,""mensagem"":""Campos obrigatórios não informados""}"
        Exit Function
    End If

    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open BuscaStringConexao("PIR-CORPORATIVO")
    Set oCmd  = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandType = 1
    oCmd.CommandText = "SELECT CNPJ FROM RECEBIMENTO_FORNECEDORES WHERE CNPJ = ?"
    oCmd.Parameters.Append oCmd.CreateParameter("@CNPJ", 200, 1, 14, sCnpj)

    On Error Resume Next
    Set rsV = oCmd.Execute()
    numErr  = Err.Number
    On Error GoTo 0

    If numErr <> 0 Then
        oConn.Close() : Set oConn = Nothing
        CadastrarFornecedor = "{""sucesso"":false,""mensagem"":""Erro ao verificar CNPJ""}"
        Exit Function
    End If

    If Not rsV.EOF Then
        rsV.Close() : Set rsV = Nothing
        Set oCmd = Nothing : oConn.Close() : Set oConn = Nothing
        CadastrarFornecedor = "{""sucesso"":false,""mensagem"":""Este CNPJ já está cadastrado""}"
        Exit Function
    End If
    rsV.Close() : Set rsV = Nothing

    Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandType = 1
    oCmd.CommandText = "INSERT INTO RECEBIMENTO_FORNECEDORES (CODIGO, NOME, CNPJ) VALUES (?,?,?)"
    oCmd.Parameters.Append oCmd.CreateParameter("@CODIGO", 200, 1, 20,  sCodigo)
    oCmd.Parameters.Append oCmd.CreateParameter("@NOME",   200, 1, 300, sNome)
    oCmd.Parameters.Append oCmd.CreateParameter("@CNPJ",   200, 1, 14,  sCnpj)

    On Error Resume Next
    oCmd.Execute()
    errDesc = Err.Description
    numErr  = Err.Number
    On Error GoTo 0

    Set oCmd = Nothing : oConn.Close() : Set oConn = Nothing

    If numErr <> 0 Then
        CadastrarFornecedor = "{""sucesso"":false,""mensagem"":""" & JsonEscape(errDesc) & """}"
    Else
        CadastrarFornecedor = "{""sucesso"":true,""mensagem"":""Fornecedor cadastrado com sucesso!""}"
    End If
End Function

Function BuscarFornecedores()
    Dim sParametro, oConn, oCmd, rsF, jsonArr, primeiroItem, numErr
    sParametro = Trim(Request("parametro"))

    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open BuscaStringConexao("PIR-CORPORATIVO")
    Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn
    oCmd.CommandType = 1

    If sParametro <> "" Then
        oCmd.CommandText = "SELECT CODIGO, NOME, CNPJ FROM RECEBIMENTO_FORNECEDORES " & _
                           "WHERE UPPER(NOME) LIKE ? OR CNPJ = ? " & _
                           "ORDER BY NOME"
        oCmd.Parameters.Append oCmd.CreateParameter("@NOME", 200, 1, 302, "%" & UCase(sParametro) & "%")
        oCmd.Parameters.Append oCmd.CreateParameter("@CNPJ", 200, 1, 14,  sParametro)
    Else
        oCmd.CommandText = "SELECT CODIGO, NOME, CNPJ FROM RECEBIMENTO_FORNECEDORES ORDER BY NOME"
    End If

    On Error Resume Next
    Set rsF = oCmd.Execute()
    numErr  = Err.Number
    On Error GoTo 0

    If numErr <> 0 Or Not IsObject(rsF) Then
        oConn.Close() : Set oConn = Nothing
        BuscarFornecedores = "{""sucesso"":false,""mensagem"":""Erro ao buscar fornecedores""}"
        Exit Function
    End If

    jsonArr      = "["
    primeiroItem = True
    Do While Not rsF.EOF
        If Not primeiroItem Then jsonArr = jsonArr & ","
        jsonArr = jsonArr & "{""cnpj"":"""    & JsonEscape(rsF("CNPJ"))    & """," & _
                            """codigo"":"""   & JsonEscape(rsF("CODIGO"))  & """," & _
                            """nome"":"""     & JsonEscape(rsF("NOME"))    & """}"
        primeiroItem = False
        rsF.MoveNext()
    Loop
    jsonArr = jsonArr & "]"
    rsF.Close() : Set rsF = Nothing
    oConn.Close() : Set oConn = Nothing

    BuscarFornecedores = "{""sucesso"":true,""dados"":" & jsonArr & "}"
End Function
%>
