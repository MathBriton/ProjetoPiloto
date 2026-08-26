<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
' =================================================================================================
' BLOCO DE API — executado antes de qualquer HTML
' =================================================================================================
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046

Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.Charset     = "UTF-8"
    Response.CodePage    = 65001
    Response.ContentType = "application/json"

    Set usuarioLogado = CreateObject("Scripting.Dictionary")
    usuarioLogado.Add "id",    Session("CODIGO")
    usuarioLogado.Add "login", Session("USR")
    usuarioLogado.Add "nome",  Session("NOME")

    Dim resultado

    Select Case acao
        Case "listar"            : resultado = Listar()
        Case "mudarStatus"       : resultado = MudarStatus()
        Case "excluir"           : resultado = Excluir()
        Case "obterRegistro"     : resultado = ObterRegistro()
        Case "obterTiposVeiculo" : resultado = ObterTiposVeiculo()
        Case "salvarAlteracao"   : resultado = SalvarAlteracao()
        Case "salvar"            : resultado = Salvar()
        Case "buscarPlaca"       : resultado = BuscarPlaca()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}"
    End Select

    Response.Write resultado
    Response.End()

End If

' =================================================================================================
' FUNÇÕES DA API
' =================================================================================================

Function Listar()
    Dim parStatus, parValor, parLinhas, SQL
    Dim parDataDe, parDataAte
    Dim rs, jsonArray, primeiroItem

    parStatus  = Trim(Request("parStatus"))
    parValor   = Trim(Request("parValor"))
    parLinhas  = Trim(Request("edtLinhas"))
    parDataDe  = Trim(Request("parDataDe"))
    parDataAte = Trim(Request("parDataAte"))

    ' "0" é o valor do option "Selecione..." no front-end — trata como sem filtro
    If parStatus = "0" Then parStatus = ""

    Dim statusValidos
    statusValidos = Array("", "'H'", "'P'", "'E'", "'F'", "NP", "NM", "DO", "PV", "DP")
    If Not inArray(parStatus, statusValidos) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Filtro de status inválido""}"
        Exit Function
    End If

    Dim topClause
    If parLinhas = "TODOS" Or parLinhas = "" Then
        topClause = ""
    Else
        If Not IsNumeric(parLinhas) Then parLinhas = "50"
        topClause = "TOP " & CLng(parLinhas) & " "
    End If

    SQL = "SELECT " & topClause & _
          "T1.ID " & _
          ",UPPER(CASE WHEN LEN(T1.NOMEPROPRIETARIO)>50 THEN CONCAT(SUBSTRING(T1.NOMEPROPRIETARIO,0,30),'...') ELSE T1.NOMEPROPRIETARIO END) AS NOMEPROPRIETARIO " & _
          ",UPPER(CASE WHEN LEN(T1.NOMEMOTORISTA)>20 THEN CONCAT(SUBSTRING(T1.NOMEMOTORISTA,0,20),'...') ELSE T1.NOMEMOTORISTA END) AS NOMEMOTORISTA " & _
          ",T1.DOCUMENTO " & _
          ",T1.PLACAVEICULO " & _
          ",T1.STATUS " & _
          ",T1.DESTINOCARGA " & _
          ",CASE T1.STATUS WHEN 'FINALIZADO' THEN 1 ELSE 0 END AS FINALIZADO " & _
          ",CONVERT(VARCHAR,CAST(T1.DATAPREENCHIMENTO AS DATE),103) AS DATAPREENCHIMENTO " & _
          "FROM CHECKLIST_TRANSPORTES T1 " & _
          "WHERE T1.ATIVO = 1 "

    If parStatus = "'H'" Then SQL = SQL & "AND T1.STATUS = 'PATIO HJ' "
    If parStatus = "'P'" Then SQL = SQL & "AND T1.STATUS = 'PORTARIA' "
    If parStatus = "'E'" Then SQL = SQL & "AND T1.STATUS = 'EXPEDIÇÃO' "
    If parStatus = "'F'" Then SQL = SQL & "AND T1.STATUS = 'FINALIZADO' "

    Dim parValorEsc
    parValorEsc = Replace(parValor, "'", "''")
    If parStatus = "NP" Then SQL = SQL & "AND UPPER(T1.NOMEPROPRIETARIO) LIKE UPPER('%" & parValorEsc & "%') "
    If parStatus = "NM" Then SQL = SQL & "AND UPPER(T1.NOMEMOTORISTA) LIKE UPPER('%" & parValorEsc & "%') "
    If parStatus = "DO" Then SQL = SQL & "AND UPPER(T1.DOCUMENTO) LIKE UPPER('%" & parValorEsc & "%') "
    If parStatus = "PV" Then SQL = SQL & "AND UPPER(T1.PLACAVEICULO) LIKE UPPER('%" & parValorEsc & "%') "
    If parStatus = "DP" Then
        ' Filtro por range de data (yyyy-MM-dd vindo do input type="date")
        Dim dpDe, dpAte
        dpDe  = Replace(parDataDe,  "'", "''")
        dpAte = Replace(parDataAte, "'", "''")
        If dpDe <> "" And dpAte <> "" Then
            SQL = SQL & "AND CAST(T1.DATAPREENCHIMENTO AS DATE) BETWEEN '" & dpDe & "' AND '" & dpAte & "' "
        ElseIf dpDe <> "" Then
            SQL = SQL & "AND CAST(T1.DATAPREENCHIMENTO AS DATE) >= '" & dpDe & "' "
        ElseIf dpAte <> "" Then
            SQL = SQL & "AND CAST(T1.DATAPREENCHIMENTO AS DATE) <= '" & dpAte & "' "
        End If
    End If

    SQL = SQL & "ORDER BY T1.DATAPREENCHIMENTO DESC"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar os dados: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    primeiroItem = True
    jsonArray    = "["

    Do While Not rs.EOF
        If Not primeiroItem Then jsonArray = jsonArray & ","
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """id"": "          & (rs("ID") & "")                        & ", "
        jsonArray = jsonArray & """proprietario"": """ & EscapeJSON(rs("NOMEPROPRIETARIO") & "") & """, "
        jsonArray = jsonArray & """motorista"": """    & EscapeJSON(rs("NOMEMOTORISTA") & "")    & """, "
        jsonArray = jsonArray & """documento"": """    & EscapeJSON(rs("DOCUMENTO") & "")        & """, "
        jsonArray = jsonArray & """placa"": """        & EscapeJSON(rs("PLACAVEICULO") & "")     & """, "
        jsonArray = jsonArray & """status"": """       & EscapeJSON(rs("STATUS") & "")           & """, "
        jsonArray = jsonArray & """destino"": """      & EscapeJSON(rs("DESTINOCARGA") & "")     & """, "
        jsonArray = jsonArray & """data"": """         & EscapeJSON(rs("DATAPREENCHIMENTO") & "") & """, "
        jsonArray = jsonArray & """finalizado"": "     & (rs("FINALIZADO") & "")
        jsonArray = jsonArray & "}"
        primeiroItem = False
        rs.MoveNext()
    Loop

    jsonArray = jsonArray & "]"
    rs.Close()
    Set rs = Nothing

    Listar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & jsonArray & "}"
End Function

Function ObterRegistro()
    Dim parID, SQL, rs, rsItens
    Dim json, jsonItens, primeiroItem

    parID = Trim(Request("parID"))

    If Not IsNumeric(parID) Or parID = "" Then
        ObterRegistro = "{""sucesso"": false, ""mensagem"": ""ID inválido""}"
        Exit Function
    End If

    SQL = "SELECT T1.ID " & _
          ",T1.NOMEPROPRIETARIO " & _
          ",T1.NOMEMOTORISTA " & _
          ",T1.TELEFONE " & _
          ",T1.DOCUMENTO " & _
          ",T1.PLACAVEICULO " & _
          ",CONVERT(VARCHAR,CAST(T1.DATAPREENCHIMENTO AS DATE),103) AS DATAPREENCHIMENTO " & _
          ",T1.ID_TIPOVEICULO " & _
          ",T1.QTDEEIXO " & _
          ",CONVERT(VARCHAR(5),CAST(T1.HORARIOINICIO AS datetime),108) AS HORARIOINICIO " & _
          ",CONVERT(VARCHAR(5),CAST(T1.HORARIOTERMINO AS datetime),108) AS HORARIOTERMINO " & _
          ",T1.DESTINOCARGA " & _
          ",T1.ID_TIPOTRANSPORTADORA " & _
          ",CASE T1.MANUTENCAOVEICULO WHEN 1 THEN '1' ELSE '0' END AS MANUTENCAOVEICULO " & _
          ",T1.CONFERENTE " & _
          ",T1.RESPONSAVELPORTARIA " & _
          ",replace(CAST(T1.COMPRIMENTO AS varchar),',','.') AS COMPRIMENTO " & _
          ",replace(CAST(T1.ALTURA AS varchar),',','.') AS ALTURA " & _
          ",replace(CAST(T1.LARGURA AS varchar),',','.') AS LARGURA " & _
          ",T1.STATUS " & _
          "FROM CHECKLIST_TRANSPORTES T1 " & _
          "WHERE T1.ID = " & CLng(parID) & " AND T1.ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        ObterRegistro = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar registro""}"
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        ObterRegistro = "{""sucesso"": false, ""mensagem"": ""Registro não encontrado""}"
        Exit Function
    End If

    json = "{"
    json = json & """id"": "                  & CLng(parID)                                           & ", "
    json = json & """nomeProprietario"": """  & EscapeJSON(rs("NOMEPROPRIETARIO") & "")              & """, "
    json = json & """nomeMoto"": """          & EscapeJSON(rs("NOMEMOTORISTA") & "")                  & """, "
    json = json & """telefone"": """          & EscapeJSON(rs("TELEFONE") & "")                       & """, "
    json = json & """documento"": """         & EscapeJSON(rs("DOCUMENTO") & "")                      & """, "
    json = json & """placa"": """             & EscapeJSON(rs("PLACAVEICULO") & "")                   & """, "
    json = json & """dataPreenchimento"": """ & EscapeJSON(rs("DATAPREENCHIMENTO") & "")              & """, "
    json = json & """idTipoVeiculo"": "       & (rs("ID_TIPOVEICULO") & "")                           & ", "
    json = json & """qtdeEixo"": """          & EscapeJSON(rs("QTDEEIXO") & "")                       & """, "
    json = json & """horarioInicio"": """     & EscapeJSON(rs("HORARIOINICIO") & "")                  & """, "
    json = json & """horarioTermino"": """    & EscapeJSON(rs("HORARIOTERMINO") & "")                 & """, "
    json = json & """destinoCarga"": """      & EscapeJSON(rs("DESTINOCARGA") & "")                   & """, "
    json = json & """idTipoTransportadora"": " & (rs("ID_TIPOTRANSPORTADORA") & "")                   & ", "
    json = json & """manutencao"": """        & EscapeJSON(rs("MANUTENCAOVEICULO") & "")              & """, "
    json = json & """conferente"": """        & EscapeJSON(rs("CONFERENTE") & "")                     & """, "
    json = json & """responsavelPortaria"": """ & EscapeJSON(rs("RESPONSAVELPORTARIA") & "")          & """, "
    json = json & """comprimento"": """       & EscapeJSON(rs("COMPRIMENTO") & "")                    & """, "
    json = json & """altura"": """            & EscapeJSON(rs("ALTURA") & "")                         & """, "
    json = json & """largura"": """           & EscapeJSON(rs("LARGURA") & "")                        & """, "
    json = json & """status"": """            & EscapeJSON(rs("STATUS") & "")                         & """"

    rs.Close() : Set rs = Nothing

    Dim sqlItens
    sqlItens = "SELECT ID, DESCRICAO, ITENSAVALIADOS_SIM FROM (" & _
               " SELECT T1.ID, T1.DESCRICAO, 0 AS ITENSAVALIADOS_SIM" & _
               " FROM CHECKLIST_LISTAITENS T1" & _
               " WHERE T1.ATIVO = 1" & _
               " AND T1.ID NOT IN (SELECT ID_CHEKLIST_ITENSAVALIADOS FROM CHECKLIST_TRANSPORTESITENS WHERE ID_CHEKLIST_TRANSPORTE = " & CLng(parID) & ")" & _
               " UNION" & _
               " SELECT T1.ID, T1.DESCRICAO, 1 AS ITENSAVALIADOS_SIM" & _
               " FROM CHECKLIST_LISTAITENS T1" & _
               " LEFT JOIN CHECKLIST_TRANSPORTESITENS T2 ON T2.ID_CHEKLIST_ITENSAVALIADOS = T1.ID" & _
               " WHERE T1.ATIVO = 1" & _
               " AND T2.ID_CHEKLIST_TRANSPORTE = " & CLng(parID) & _
               ") AS Q ORDER BY ID"

    On Error Resume Next
    Set rsItens = Conecta("PIR-CORPORATIVO", sqlItens)
    If Err.Number <> 0 Or Not IsObject(rsItens) Then
        ObterRegistro = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar itens""}"
        Exit Function
    End If
    On Error GoTo 0

    jsonItens    = "["
    primeiroItem = True

    Do While Not rsItens.EOF
        If Not primeiroItem Then jsonItens = jsonItens & ","
        jsonItens = jsonItens & "{"
        jsonItens = jsonItens & """id"": "      & (rsItens("ID") & "")                      & ", "
        jsonItens = jsonItens & """descricao"": """ & EscapeJSON(rsItens("DESCRICAO") & "") & """, "
        jsonItens = jsonItens & """sim"": "     & (rsItens("ITENSAVALIADOS_SIM") & "")
        jsonItens = jsonItens & "}"
        primeiroItem = False
        rsItens.MoveNext()
    Loop

    jsonItens = jsonItens & "]"
    rsItens.Close() : Set rsItens = Nothing

    json = json & ", ""itens"": " & jsonItens & "}"

    ObterRegistro = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & json & "}"
End Function

Function ObterTiposVeiculo()
    Dim SQL, rs, jsonArray, primeiroItem

    SQL = "SELECT ID, DESCRICAO FROM CHECKLIST_TIPOVEICULO WHERE ATIVO = 1 ORDER BY DESCRICAO"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        ObterTiposVeiculo = "{""sucesso"": false, ""mensagem"": ""Erro ao carregar tipos de veículo""}"
        Exit Function
    End If
    On Error GoTo 0

    jsonArray    = "["
    primeiroItem = True

    Do While Not rs.EOF
        If Not primeiroItem Then jsonArray = jsonArray & ","
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """id"": "       & (rs("ID") & "")                     & ", "
        jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rs("DESCRICAO") & "") & """"
        jsonArray = jsonArray & "}"
        primeiroItem = False
        rs.MoveNext()
    Loop

    jsonArray = jsonArray & "]"
    rs.Close() : Set rs = Nothing

    ObterTiposVeiculo = "{""sucesso"": true, ""dados"": " & jsonArray & "}"
End Function

Function SalvarAlteracao()
    Dim parID, rs, SQL

    parID = Trim(Request("parID"))

    If Not IsNumeric(parID) Or parID = "" Then
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""ID inválido""}"
        Exit Function
    End If

    Dim sqlVerifica
    sqlVerifica = "SELECT STATUS FROM CHECKLIST_TRANSPORTES WHERE ID = " & CLng(parID) & " AND ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", sqlVerifica)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar registro""}"
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Registro não encontrado""}"
        Exit Function
    End If

    If rs("STATUS") = "FINALIZADO" Then
        rs.Close() : Set rs = Nothing
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Registro finalizado não pode ser alterado""}"
        Exit Function
    End If

    rs.Close() : Set rs = Nothing

    Dim fNomeProprietario, fNomeMotorista, fTelefone, fDocumento, fPlaca
    Dim fDataPreenchimento, fIdTipoVeiculo, fQtdeEixo, fHorarioInicio, fHorarioTermino
    Dim fDestinoCarga, fIdTipoTransportadora, fManutencao, fConferente, fResponsavelPortaria
    Dim fComprimento, fAltura, fLargura, fStatus, fItensAvaliados

    fNomeProprietario    = Replace(Trim(Request("nomeProprietario")),    "'", "''")
    fNomeMotorista       = Replace(Trim(Request("nomeMotorista")),       "'", "''")
    fTelefone            = Replace(Trim(Request("telefone")),            "'", "''")
    fDocumento           = Replace(Trim(Request("documento")),           "'", "''")
    fPlaca               = Replace(Trim(Request("placa")),               "'", "''")
    fDataPreenchimento   = Replace(Trim(Request("dataPreenchimento")),   "'", "''")
    fDestinoCarga        = Replace(Trim(Request("destinoCarga")),        "'", "''")
    fConferente          = Replace(Trim(Request("conferente")),          "'", "''")
    fResponsavelPortaria = Replace(Trim(Request("responsavelPortaria")), "'", "''")
    fStatus              = Replace(Trim(Request("status")),              "'", "''")
    fItensAvaliados      = Trim(Request("itensAvaliados"))

    If Not IsNumeric(Trim(Request("idTipoVeiculo")))       Then SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Tipo de veículo inválido""}"     : Exit Function
    If Not IsNumeric(Trim(Request("idTipoTransportadora"))) Then SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Tipo de transportadora inválido""}" : Exit Function

    fIdTipoVeiculo        = CLng(Trim(Request("idTipoVeiculo")))
    fIdTipoTransportadora = CLng(Trim(Request("idTipoTransportadora")))
    fQtdeEixo             = CLng(Trim(Request("qtdeEixo")))

    Dim fManutBit
    If Trim(Request("manutencao")) = "1" Then fManutBit = 1 Else fManutBit = 0

    Dim dataParts, dataSQL
    dataParts = Split(fDataPreenchimento, "/")
    If UBound(dataParts) = 2 Then
        dataSQL = dataParts(2) & "-" & dataParts(1) & "-" & dataParts(0)
    Else
        dataSQL = fDataPreenchimento
    End If

    Dim statusPermitidos
    statusPermitidos = Array("PATIO HJ", "PORTARIA", "EXPEDIÇÃO", "FINALIZADO")
    If Not inArray(fStatus, statusPermitidos) Then
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Status inválido""}"
        Exit Function
    End If

    SQL = "UPDATE CHECKLIST_TRANSPORTES SET " & _
          " NOMEPROPRIETARIO    = '" & fNomeProprietario    & "'" & _
          ",NOMEMOTORISTA       = '" & fNomeMotorista       & "'" & _
          ",TELEFONE            = '" & fTelefone            & "'" & _
          ",DOCUMENTO           = '" & fDocumento           & "'" & _
          ",PLACAVEICULO        = '" & fPlaca               & "'" & _
          ",DATAPREENCHIMENTO   = '" & dataSQL              & "'" & _
          ",ID_TIPOVEICULO      = "  & fIdTipoVeiculo               & _
          ",QTDEEIXO            = "  & fQtdeEixo                    & _
          ",HORARIOINICIO       = '" & Trim(Request("horarioInicio"))  & "'" & _
          ",HORARIOTERMINO      = '" & Trim(Request("horarioTermino")) & "'" & _
          ",DESTINOCARGA        = '" & fDestinoCarga        & "'" & _
          ",ID_TIPOTRANSPORTADORA = " & fIdTipoTransportadora         & _
          ",MANUTENCAOVEICULO   = "  & fManutBit                     & _
          ",CONFERENTE          = '" & fConferente          & "'" & _
          ",RESPONSAVELPORTARIA = '" & fResponsavelPortaria & "'" & _
          ",COMPRIMENTO         = "  & Replace(Trim(Request("comprimento")), ",", ".") & _
          ",ALTURA              = "  & Replace(Trim(Request("altura")),      ",", ".") & _
          ",LARGURA             = "  & Replace(Trim(Request("largura")),     ",", ".") & _
          ",STATUS              = '" & fStatus              & "'" & _
          " WHERE ID = " & CLng(parID) & " AND ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    Dim errNumSA : errNumSA = Err.Number
    Dim errDescSA : errDescSA = Err.Description
    On Error GoTo 0

    If errNumSA <> 0 Then
        SalvarAlteracao = "{""sucesso"": false, ""mensagem"": ""Erro ao salvar: " & EscapeJSON(errDescSA) & """}"
        Exit Function
    End If

    On Error Resume Next
    If IsObject(rs) Then rs.Close() : Set rs = Nothing
    On Error GoTo 0

    Dim sqlDel
    sqlDel = "DELETE FROM CHECKLIST_TRANSPORTESITENS WHERE ID_CHEKLIST_TRANSPORTE = " & CLng(parID)

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", sqlDel)
    On Error GoTo 0
    On Error Resume Next
    If IsObject(rs) Then rs.Close() : Set rs = Nothing
    On Error GoTo 0

    If fItensAvaliados <> "" Then
        Dim arrItens, itemID
        arrItens = Split(fItensAvaliados, "|")
        Dim i
        For i = 0 To UBound(arrItens)
            itemID = Trim(arrItens(i))
            If IsNumeric(itemID) Then
            If CLng(itemID) > 0 Then
                Dim sqlIns
                sqlIns = "INSERT INTO CHECKLIST_TRANSPORTESITENS (ID_CHEKLIST_TRANSPORTE, ID_CHEKLIST_ITENSAVALIADOS)" & _
                         " VALUES (" & CLng(parID) & ", " & CLng(itemID) & ")"
                On Error Resume Next
                Set rs = Conecta("PIR-CORPORATIVO", sqlIns)
                On Error GoTo 0
                On Error Resume Next
                If IsObject(rs) Then rs.Close() : Set rs = Nothing
                On Error GoTo 0
            End If
            End If
        Next
    End If

    SalvarAlteracao = "{""sucesso"": true, ""mensagem"": ""Registro alterado com sucesso""}"
End Function

Function BuscarPlaca()
    Dim placa, SQL, rs, errNumBP, jsonBP

    placa = Replace(Trim(Request("placa")), "'", "''")
    If placa = "" Then
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Placa não informada.""}"
        Exit Function
    End If

    SQL = "SELECT TOP 1 NOMEPROPRIETARIO, TELEFONE, DOCUMENTO, NOMEMOTORISTA, " & _
          "ID_TIPOVEICULO, QTDEEIXO, " & _
          "replace(CAST(ALTURA     AS varchar),',','.') AS ALTURA, " & _
          "replace(CAST(COMPRIMENTO AS varchar),',','.') AS COMPRIMENTO, " & _
          "replace(CAST(LARGURA    AS varchar),',','.') AS LARGURA " & _
          "FROM CHECKLIST_TRANSPORTES " & _
          "WHERE PLACAVEICULO = '" & placa & "' AND ATIVO = 1 ORDER BY DATACADASTRO DESC"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    errNumBP = Err.Number
    On Error GoTo 0

    If errNumBP <> 0 Or Not IsObject(rs) Then
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar placa.""}"
        Exit Function
    End If

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Placa não encontrada.""}"
        Exit Function
    End If

    jsonBP = "{"
    jsonBP = jsonBP & """sucesso"": true,"
    jsonBP = jsonBP & """dados"": {"
    jsonBP = jsonBP & """nomeProprietario"": """ & EscapeJSON(rs("NOMEPROPRIETARIO") & "") & ""","
    jsonBP = jsonBP & """telefone"": """         & EscapeJSON(rs("TELEFONE")         & "") & ""","
    jsonBP = jsonBP & """cnpjCpf"": """          & EscapeJSON(rs("DOCUMENTO")        & "") & ""","
    jsonBP = jsonBP & """nomeMotorista"": """     & EscapeJSON(rs("NOMEMOTORISTA")    & "") & ""","
    jsonBP = jsonBP & """idTipoVeiculo"": "       & CLng(rs("ID_TIPOVEICULO") & "0")        & ","
    jsonBP = jsonBP & """quantidadeEixo"": "      & CLng(rs("QTDEEIXO")       & "0")        & ","
    jsonBP = jsonBP & """altura"": """            & EscapeJSON(rs("ALTURA")           & "") & ""","
    jsonBP = jsonBP & """comprimento"": """       & EscapeJSON(rs("COMPRIMENTO")      & "") & ""","
    jsonBP = jsonBP & """largura"": """           & EscapeJSON(rs("LARGURA")          & "") & """"
    jsonBP = jsonBP & "}}"

    rs.Close() : Set rs = Nothing
    BuscarPlaca = jsonBP
End Function

Function Salvar()
    Dim vNomeProprietario, vNomeMotorista, vTelefone, vDocumento, vPlaca
    Dim vDataPreenchimento, vIdTipoVeiculo, vQtdeEixo
    Dim vDestinoCarga, vTipoTransporte, vConferente, vResponsavelPortaria
    Dim vStatus, vItensAvaliados
    Dim vManutBit, vDataPartes, vDataFormatada, vStatusOk
    Dim vHorarioInicio, vHorarioTermino, vHorarioInicioSql, vHorarioTerminoSql
    Dim vSqlInsert, vRec, vErrNum, vErrDesc
    Dim vSqlNovoId, vRecNovoId, vNovoId
    Dim vArrItens, vItemId, vIdx, vSqlItem

    vNomeProprietario    = Replace(Trim(Request("nomeProprietario")),    "'", "''")
    vNomeMotorista       = Replace(Trim(Request("nomeMotorista")),       "'", "''")
    vTelefone            = Replace(Trim(Request("telefone")),            "'", "''")
    vDocumento           = Replace(Trim(Request("cnpjCpf")),             "'", "''")
    vPlaca               = Replace(Trim(Request("placa")),               "'", "''")
    vDataPreenchimento   = Replace(Trim(Request("dataPreenchimento")),   "'", "''")
    vDestinoCarga        = Replace(Trim(Request("destinoCarga")),        "'", "''")
    vConferente          = Replace(Trim(Request("nomeConferente")),      "'", "''")
    vResponsavelPortaria = Replace(Trim(Request("portariaResponsavel")), "'", "''")
    vStatus              = Replace(Trim(Request("status")),              "'", "''")
    vItensAvaliados      = Trim(Request("itensAvaliados"))

    If Not IsNumeric(Trim(Request("idTipoVeiculo")))  Then Salvar = "{""sucesso"": false, ""mensagem"": ""Tipo de veículo inválido.""}"    : Exit Function
    If Not IsNumeric(Trim(Request("tipoTransporte"))) Then Salvar = "{""sucesso"": false, ""mensagem"": ""Tipo de transporte inválido.""}"  : Exit Function
    If Not IsNumeric(Trim(Request("quantidadeEixo"))) Then Salvar = "{""sucesso"": false, ""mensagem"": ""Quantidade de eixos inválida.""}" : Exit Function

    vIdTipoVeiculo  = CLng(Trim(Request("idTipoVeiculo")))
    vTipoTransporte = CLng(Trim(Request("tipoTransporte")))
    vQtdeEixo       = CLng(Trim(Request("quantidadeEixo")))

    If Trim(Request("manutencao")) = "1" Then vManutBit = 1 Else vManutBit = 0

    ' Converte dd/MM/yyyy para yyyy-MM-dd
    vDataPartes = Split(vDataPreenchimento, "/")
    If UBound(vDataPartes) = 2 Then
        vDataFormatada = vDataPartes(2) & "-" & vDataPartes(1) & "-" & vDataPartes(0)
    Else
        vDataFormatada = vDataPreenchimento
    End If

    vStatusOk = Array("PATIO HJ", "PORTARIA", "EXPEDIÇÃO", "FINALIZADO")
    If Not inArray(vStatus, vStatusOk) Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Status inválido.""}"
        Exit Function
    End If

    vHorarioInicio  = Trim(Request("horarioInicio"))
    vHorarioTermino = Trim(Request("horarioTermino"))

    ' Formata HH:MM ou passa NULL quando o valor nao e um horario valido
    If Len(vHorarioInicio) >= 3 And InStr(vHorarioInicio, ":") > 0 Then
        vHorarioInicioSql  = "'" & Replace(vHorarioInicio,  "'", "''") & "'"
    ElseIf IsNumeric(vHorarioInicio) And Len(vHorarioInicio) = 4 Then
        vHorarioInicioSql  = "'" & Left(vHorarioInicio, 2) & ":" & Right(vHorarioInicio, 2) & "'"
    Else
        vHorarioInicioSql  = "NULL"
    End If

    If Len(vHorarioTermino) >= 3 And InStr(vHorarioTermino, ":") > 0 Then
        vHorarioTerminoSql = "'" & Replace(vHorarioTermino, "'", "''") & "'"
    ElseIf IsNumeric(vHorarioTermino) And Len(vHorarioTermino) = 4 Then
        vHorarioTerminoSql = "'" & Left(vHorarioTermino, 2) & ":" & Right(vHorarioTermino, 2) & "'"
    Else
        vHorarioTerminoSql = "NULL"
    End If

    vSqlInsert = "INSERT INTO CHECKLIST_TRANSPORTES " & _
                 "(NOMEPROPRIETARIO, NOMEMOTORISTA, TELEFONE, DOCUMENTO, PLACAVEICULO, " & _
                 " DATAPREENCHIMENTO, ID_TIPOVEICULO, QTDEEIXO, HORARIOINICIO, HORARIOTERMINO, " & _
                 " DESTINOCARGA, ID_TIPOTRANSPORTADORA, MANUTENCAOVEICULO, CONFERENTE, " & _
                 " RESPONSAVELPORTARIA, COMPRIMENTO, ALTURA, LARGURA, STATUS, ATIVO, DATACADASTRO) " & _
                 "VALUES (" & _
                 "'" & vNomeProprietario    & "'," & _
                 "'" & vNomeMotorista       & "'," & _
                 "'" & vTelefone            & "'," & _
                 "'" & vDocumento           & "'," & _
                 "'" & vPlaca               & "'," & _
                 "'" & vDataFormatada       & "'," & _
                      vIdTipoVeiculo        & ","  & _
                      vQtdeEixo             & ","  & _
                      vHorarioInicioSql     & ","  & _
                      vHorarioTerminoSql    & ","  & _
                 "'" & vDestinoCarga        & "'," & _
                      vTipoTransporte       & ","  & _
                      vManutBit             & ","  & _
                 "'" & vConferente          & "'," & _
                 "'" & vResponsavelPortaria & "'," & _
                      Replace(Trim(Request("comprimento")), ",", ".") & "," & _
                      Replace(Trim(Request("altura")),      ",", ".") & "," & _
                      Replace(Trim(Request("largura")),     ",", ".") & "," & _
                 "'" & vStatus              & "'," & _
                 "1, GETDATE())"

    On Error Resume Next
    Set vRec = Conecta("PIR-CORPORATIVO", vSqlInsert)
    vErrNum  = Err.Number
    vErrDesc = Err.Description
    On Error GoTo 0

    If vErrNum <> 0 Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro ao gravar checklist: " & EscapeJSON(vErrDesc) & """}"
        Exit Function
    End If

    On Error Resume Next
    If IsObject(vRec) Then vRec.Close() : Set vRec = Nothing
    On Error GoTo 0

    ' Recupera o ID recem inserido
    vSqlNovoId = "SELECT TOP 1 ID FROM CHECKLIST_TRANSPORTES WHERE ATIVO = 1 ORDER BY DATACADASTRO DESC, ID DESC"
    On Error Resume Next
    Set vRecNovoId = Conecta("PIR-CORPORATIVO", vSqlNovoId)
    On Error GoTo 0

    vNovoId = 0
    If Not (vRecNovoId Is Nothing) Then
        If Not vRecNovoId.EOF Then vNovoId = CLng(vRecNovoId("ID") & "0")
        vRecNovoId.Close() : Set vRecNovoId = Nothing
    End If

    ' Grava itens avaliados
    If vItensAvaliados <> "" And vNovoId > 0 Then
        vArrItens = Split(vItensAvaliados, ",")
        For vIdx = 0 To UBound(vArrItens)
            vItemId = Trim(vArrItens(vIdx))
            If IsNumeric(vItemId) And CLng(vItemId) > 0 Then
                vSqlItem = "INSERT INTO CHECKLIST_TRANSPORTESITENS (ID_CHEKLIST_TRANSPORTE, ID_CHEKLIST_ITENSAVALIADOS)" & _
                           " VALUES (" & vNovoId & ", " & CLng(vItemId) & ")"
                On Error Resume Next
                Set vRec = Conecta("PIR-CORPORATIVO", vSqlItem)
                On Error GoTo 0
                On Error Resume Next
                If IsObject(vRec) Then vRec.Close() : Set vRec = Nothing
                On Error GoTo 0
            End If
        Next
    End If

    Salvar = "{""sucesso"": true, ""mensagem"": ""CheckList gravado com sucesso.""}"
End Function

Function MudarStatus()
    Dim parID, parStatus, novoStatus, SQL, rs

    parID     = Trim(Request("parID"))
    parStatus = Trim(Request("parStatus"))

    If Not IsNumeric(parID) Or parID = "" Then
        MudarStatus = "{""sucesso"": false, ""mensagem"": ""ID inválido""}"
        Exit Function
    End If

    Select Case UCase(parStatus)
        Case "PAT" : novoStatus = "PATIO HJ"
        Case "POR" : novoStatus = "PORTARIA"
        Case "EXP" : novoStatus = "EXPEDIÇÃO"
        Case "FIN" : novoStatus = "FINALIZADO"
        Case Else
            MudarStatus = "{""sucesso"": false, ""mensagem"": ""Status inválido""}"
            Exit Function
    End Select

    Dim sqlVerifica
    sqlVerifica = "SELECT STATUS FROM CHECKLIST_TRANSPORTES WHERE ID = " & CLng(parID) & " AND ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", sqlVerifica)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        MudarStatus = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar registro""}"
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        MudarStatus = "{""sucesso"": false, ""mensagem"": ""Registro não encontrado""}"
        Exit Function
    End If

    If rs("STATUS") = "FINALIZADO" Then
        rs.Close() : Set rs = Nothing
        MudarStatus = "{""sucesso"": false, ""mensagem"": ""Registro finalizado não pode ser alterado""}"
        Exit Function
    End If

    rs.Close() : Set rs = Nothing

    Dim novoStatusEsc
    novoStatusEsc = Replace(novoStatus, "'", "''")

    SQL = "UPDATE CHECKLIST_TRANSPORTES " & _
          "SET STATUS = '" & novoStatusEsc & "' " & _
          "WHERE ID = " & CLng(parID) & " AND ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    Dim errNumMS : errNumMS = Err.Number
    Dim errDescMS : errDescMS = Err.Description
    On Error GoTo 0

    If errNumMS <> 0 Then
        MudarStatus = "{""sucesso"": false, ""mensagem"": ""Erro ao atualizar status: " & EscapeJSON(errDescMS) & """}"
        Exit Function
    End If

    On Error Resume Next
    If IsObject(rs) Then rs.Close() : Set rs = Nothing
    On Error GoTo 0

    MudarStatus = "{""sucesso"": true, ""mensagem"": ""Status atualizado para " & EscapeJSON(novoStatus) & """}"
End Function

Function Excluir()
    Dim parID, SQL, rs

    parID = Trim(Request("parID"))

    If Not IsNumeric(parID) Or parID = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""ID inválido""}"
        Exit Function
    End If

    Dim sqlVerifica
    sqlVerifica = "SELECT STATUS FROM CHECKLIST_TRANSPORTES WHERE ID = " & CLng(parID) & " AND ATIVO = 1"

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", sqlVerifica)
    If Err.Number <> 0 Or Not IsObject(rs) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar registro""}"
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close() : Set rs = Nothing
        Excluir = "{""sucesso"": false, ""mensagem"": ""Registro não encontrado""}"
        Exit Function
    End If

    If rs("STATUS") = "FINALIZADO" Then
        rs.Close() : Set rs = Nothing
        Excluir = "{""sucesso"": false, ""mensagem"": ""Registros finalizados não podem ser excluídos""}"
        Exit Function
    End If

    rs.Close() : Set rs = Nothing

    SQL = "UPDATE CHECKLIST_TRANSPORTES SET ATIVO = 0 WHERE ID = " & CLng(parID)

    On Error Resume Next
    Set rs = Conecta("PIR-CORPORATIVO", SQL)
    Dim errNumEx : errNumEx = Err.Number
    Dim errDescEx : errDescEx = Err.Description
    On Error GoTo 0

    If errNumEx <> 0 Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir: " & EscapeJSON(errDescEx) & """}"
        Exit Function
    End If

    On Error Resume Next
    If IsObject(rs) Then rs.Close() : Set rs = Nothing
    On Error GoTo 0

    Excluir = "{""sucesso"": true, ""mensagem"": ""Registro excluído com sucesso""}"
End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — CheckList Transportes</title>

  <link href="css/bootstrap.min.css" rel="stylesheet" />
  <link href="css/bootstrap-icons.min.css" rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    /* Filter card */
    .filter-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .filter-card-header {
      background-color: #fff;
      color: #212529;
      padding: .6rem 1rem;
      border-radius: 6px 6px 0 0;
      border-bottom: 1px solid #dee2e6;
      font-weight: 600;
      font-size: .9rem;
    }
    .filter-card-body { padding: 1rem 1.25rem 1.25rem; }
    .filter-card label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }
    .input-data { width: 155px; }

    .btn-status-group .btn { font-size: .78rem; }

    /* Table card */
    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      position: relative;
    }
    .table-toolbar { background-color: #fff; border-radius: 6px 6px 0 0; }

    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
    .table-wrapper table { min-width: max-content; }
    .table thead th {
      position: sticky; top: 0; background-color: #fff; color: #212529;
      font-size: .82rem; font-weight: 700; text-align: center;
      white-space: nowrap; z-index: 2; border-color: #dee2e6;
      padding: .55rem .75rem;
    }
    .table tbody td {
      font-size: .82rem; text-align: center; white-space: nowrap;
      color: #212529; vertical-align: middle; padding: .45rem .75rem;
    }

    .table-footer {
      display: flex; align-items: center; justify-content: space-between;
      padding: .6rem 1rem; border-top: 1px solid #dee2e6;
      flex-wrap: wrap; gap: .5rem;
    }
    .table-footer-left  { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
    .table-footer-center{ flex: 1; display: flex; justify-content: center; }
    .table-footer-right { min-width: 140px; text-align: right; }
    .info-registros { font-size: .8rem; color: #6c757d; }
    .select-limite  { width: auto; }

    .overlay-loading {
      position: absolute; inset: 0;
      background: rgba(255,255,255,.80);
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; z-index: 10; gap: 1rem; border-radius: 6px;
    }
    .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
    .overlay-loading span { font-size: .9rem; font-weight: 600; color: #495057; }

    .table-placeholder {
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; padding: 3.5rem 1rem; color: #6c757d; gap: .75rem;
    }
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

    /* ── Toast canônico (igual ao CheckListTransporte.asp) ── */
    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: 1100; display: flex; flex-direction: column; gap: .5rem;
    }
    .toast-custom {
      min-width: 280px; max-width: 360px;
      border-radius: 10px; overflow: hidden;
      box-shadow: 0 4px 18px rgba(0,0,0,.13);
      background: #fff; border: 1px solid #dee2e6;
      animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
    }
    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(40px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-custom .toast-progress-bar {
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; } to { width: 0%; }
    }

    /* Status badges */
    .badge-status {
      font-size: .72rem; font-weight: 700; padding: .3em .65em; border-radius: 4px;
      display: inline-block;
    }
    .badge-patio     { background-color: #cff4fc; color: #055160; }
    .badge-portaria  { background-color: #fff3cd; color: #664d03; }
    .badge-expedicao { background-color: #cfe2ff; color: #084298; }
    .badge-finalizado{ background-color: #d1e7dd; color: #0a3622; }
    .badge-outros    { background-color: #e2e3e5; color: #2b2d2f; }

    /* Seções dos modais */
    .form-section {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      margin-bottom: 1rem;
    }
    .form-section-header {
      background-color: #cfcfcf;
      color: #050000;
      padding: .5rem 1rem;
      border-radius: 6px 6px 0 0;
      font-weight: 600;
      font-size: .85rem;
      display: flex;
      align-items: center;
      gap: .5rem;
    }
    .form-section-body { padding: 1rem 1.25rem; }

    .modal .form-label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    .table-itens thead th {
      background-color: #909090; color: #fff;
      font-size: .8rem; font-weight: 600;
      text-align: center; padding: .45rem .75rem;
    }
    .table-itens tbody td {
      font-size: .82rem; vertical-align: middle; padding: .35rem .75rem;
    }

    .modal-overlay-loading {
      position: absolute; inset: 0;
      background: rgba(255,255,255,.85);
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; z-index: 10; gap: .75rem; border-radius: 6px;
    }

    .radio-opcao { cursor: pointer; display: flex; align-items: center; gap: .4rem; }
    .radio-opcao input[type="radio"] { cursor: pointer; }

    .notas-lista { font-size: .82rem; color: #495057; line-height: 1.6; }
    .notas-lista li { margin-bottom: .3rem; }

    /* Skeleton do modal Novo */
    .skeleton-bar {
      height: 14px; border-radius: 4px;
      background: linear-gradient(90deg, #e9ecef 25%, #dee2e6 50%, #e9ecef 75%);
      background-size: 200% 100%;
      animation: shimmer 1.2s infinite;
    }
    @keyframes shimmer { to { background-position: -200% 0; } }
  </style>
</head>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center justify-content-between mb-3">
      <span class="page-title">CheckList Transportes</span>
      <button type="button" class="btn btn-outline-secondary btn-sm"
              onclick="window.location.href='CheckListTransporteIndicador.asp'">
        <i class="bi bi-bar-chart-line me-2"></i>Indicadores
      </button>
    </div>

    <!-- Card de Filtros -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">

        <div class="d-flex align-items-end flex-wrap gap-3">
          <div>
            <label for="cboMudaStatus" class="form-label">Pesquisar por</label>
            <select id="cboMudaStatus" class="form-select form-select-sm"
                    onchange="toggleCampoPesquisa()">
              <option value="0">Selecione...</option>
              <option value="NP">Nome do Proprietário</option>
              <option value="NM">Nome do Motorista</option>
              <option value="PV">Placa do Veículo</option>
              <option value="DP">Data de Preenchimento</option>
              <option value="IN">Indicadores</option>
            </select>
          </div>

          <div id="wrapperPesquisa" style="display:none;">
            <label for="txtPesquisa" class="form-label">Valor</label>
            <input type="text" id="txtPesquisa"
                   class="form-control form-control-sm" autocomplete="off"
                   placeholder="Digite para filtrar..." style="min-width:200px;" />
          </div>

          <div id="wrapperDataDe" style="display:none;">
            <label for="parDataDe" class="form-label">De</label>
            <input type="date" id="parDataDe"
                   class="form-control form-control-sm input-data" />
          </div>
          <div id="wrapperDataAte" style="display:none;">
            <label for="parDataAte" class="form-label">Até</label>
            <input type="date" id="parDataAte"
                   class="form-control form-control-sm input-data" />
          </div>

          <div id="wrapperIndicadores" style="display:none;">
            <label class="form-label">Data Inicial</label>
            <input type="date" id="parDataInicial"
                   class="form-control form-control-sm input-data" />
          </div>
          <div id="wrapperIndicadores2" style="display:none;">
            <label class="form-label">Data Final</label>
            <input type="date" id="parDataFinal"
                   class="form-control form-control-sm input-data" />
          </div>
        </div>

        <div class="d-flex align-items-center flex-wrap gap-2 mt-3 justify-content-between">
          <div class="d-flex flex-wrap gap-2 btn-status-group">
            <button type="button" id="btnPatio"      class="btn btn-sm btn-light" onclick="filtrarStatus('H')">
              <i class="bi bi-truck me-2"></i>Pátio
            </button>
            <button type="button" id="btnPortaria"   class="btn btn-sm btn-light" onclick="filtrarStatus('P')">
              <i class="bi bi-door-open me-2"></i>Portaria
            </button>
            <button type="button" id="btnExpedicao"  class="btn btn-sm btn-light" onclick="filtrarStatus('E')">
              <i class="bi bi-box-seam me-2"></i>Expedição
            </button>
            <button type="button" id="btnFinalizado" class="btn btn-sm btn-light" onclick="filtrarStatus('F')">
              <i class="bi bi-check-circle me-2"></i>Finalizado
            </button>
            <button type="button" id="btnTodos"      class="btn btn-sm btn-light" onclick="filtrarStatus('T')">
              <i class="bi bi-list me-2"></i>Todos
            </button>
          </div>
          <div class="d-flex gap-2">
            <button type="button" class="btn btn-light btn-sm" onclick="limparFiltros()">
              <i class="bi bi-eraser me-2"></i>Limpar
            </button>
            <button type="button" class="btn btn-primary btn-sm" onclick="pesquisar()">
              <i class="bi bi-funnel-fill me-2"></i>Filtrar
            </button>
          </div>
        </div>

      </div>
    </div>

    <!-- Card de Tabela -->
    <div class="table-card">

      <div id="overlayLoading" class="overlay-loading">
        <div class="spinner-border" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          
          <button type="button" class="btn btn-primary btn-sm"
                  onclick="abrirModalNovo()"
                  title="Cadastrar novo CheckList">
            <i class="bi bi-plus-lg me-2"></i>Novo CheckList
          </button>
        </div>
        <div>
          <button id="btnExportCsv" type="button" class="btn btn-success btn-sm">
            <i class="bi bi-file-earmark-excel me-2"></i>Exportar Excel
          </button>
        </div>
      </div>

      <div class="table-wrapper" id="wrapperTabela"></div>

      <div class="table-footer">
        <div class="table-footer-left">
          <label for="edtLinhas" class="form-label mb-0 text-muted"
                 style="font-size:.8rem;white-space:nowrap;">Linhas por página:</label>
          <select id="edtLinhas" class="form-select form-select-sm select-limite"
                  onchange="carregarDados()">
            <option value="10">10</option>
            <option value="20">20</option>
            <option value="50" selected>50</option>
            <option value="100">100</option>
            <option value="TODOS">Todos</option>
          </select>
        </div>
        <div class="table-footer-center"></div>
        <div class="table-footer-right">
          <span id="infoRegistros" class="info-registros"></span>
        </div>
      </div>

    </div><!-- /.table-card -->

  </div><!-- /.container-fluid -->

  <!-- ============================================================
       MODAL — NOVO CHECKLIST
       Consome os endpoints de CheckListTransporte.asp via AJAX,
       sem trocar de página (sem window.location.href).
  ============================================================ -->
  <div class="modal fade" id="modalNovo" tabindex="-1"
       aria-labelledby="modalNovoLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-scrollable">
      <div class="modal-content" style="position:relative;">

        <div id="modalNovoOverlay" class="modal-overlay-loading">
          <div class="spinner-border text-secondary" role="status">
            <span class="visually-hidden">Carregando...</span>
          </div>
          <span style="font-size:.9rem;font-weight:600;color:#495057;">Carregando formulário...</span>
        </div>

        <div class="modal-header border-bottom">
          <h6 class="modal-title fw-bold" id="modalNovoLabel">
            <i class="bi bi-plus-circle me-2 text-primary"></i>Novo CheckList de Transporte
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>

        <div class="modal-body p-3">
          <div id="modalNovoCorpo">
            <!-- preenchido via JS após obterDadosIniciais -->
          </div>
        </div>

        <div class="modal-footer border-top gap-2">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">
            Cancelar
          </button>
          <button type="button" class="btn btn-primary btn-sm" id="btnGravarNovo">
            Gravar
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ============================================================
       MODAL — ALTERAR
  ============================================================ -->
  <div class="modal fade" id="modalAlterar" tabindex="-1"
       aria-labelledby="modalAlterarLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-scrollable">
      <div class="modal-content" style="position:relative;">

        <div id="modalAlterarOverlay" class="modal-overlay-loading d-none">
          <div class="spinner-border text-secondary" role="status">
            <span class="visually-hidden">Carregando...</span>
          </div>
          <span style="font-size:.9rem;font-weight:600;color:#495057;">Carregando...</span>
        </div>

        <div class="modal-header border-bottom">
          <h6 class="modal-title fw-bold" id="modalAlterarLabel">
            <i class="bi bi-pencil me-2 text-primary"></i>Alterar CheckList — <span id="lblAlterarId"></span>
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>

        <div class="modal-body p-3">
          <div id="modalAlterarCorpo"></div>
        </div>

        <div class="modal-footer border-top gap-2">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">
            Cancelar
          </button>
          <button type="button" class="btn btn-primary btn-sm" id="btnSalvarAlteracao">
            Salvar Alteração
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ============================================================
       MODAL — VISUALIZAR
  ============================================================ -->
  <div class="modal fade" id="modalVisualizar" tabindex="-1"
       aria-labelledby="modalVisualizarLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-scrollable">
      <div class="modal-content" style="position:relative;">

        <div id="modalVisualizarOverlay" class="modal-overlay-loading d-none">
          <div class="spinner-border text-secondary" role="status">
            <span class="visually-hidden">Carregando...</span>
          </div>
          <span style="font-size:.9rem;font-weight:600;color:#495057;">Carregando...</span>
        </div>

        <div class="modal-header border-bottom">
          <h6 class="modal-title fw-bold" id="modalVisualizarLabel">
            <i class="bi bi-eye me-2 text-primary"></i>Visualizar CheckList — <span id="lblVisualizarId"></span>
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>

        <div class="modal-body p-3">
          <div id="modalVisualizarCorpo"></div>
        </div>

        <div class="modal-footer border-top">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">
            Fechar
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- Modal de confirmação de exclusão -->
  <div class="modal fade" id="modalExcluir" tabindex="-1"
       aria-labelledby="modalExcluirLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header border-bottom-0 pb-0">
          <h6 class="modal-title fw-bold" id="modalExcluirLabel">
            <i class="bi bi-exclamation-triangle-fill text-danger me-2"></i>Confirmar Exclusão
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body py-3">
          <p class="mb-1" style="font-size:.9rem;">Deseja realmente excluir o registro <strong id="modalExcluirId"></strong>?</p>
          <p class="mb-0 text-muted" style="font-size:.8rem;">Esta ação não poderá ser desfeita.</p>
        </div>
        <div class="modal-footer border-top-0 pt-0 gap-2">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancelar</button>
          <button type="button" class="btn btn-danger btn-sm" id="btnConfirmarExclusao">Excluir</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Toast container -->
  <div id="toastContainer" class="toast-container-custom"></div>

 <script src="js/jquery.min.js"></script>
 <script src="js/bootstrap.bundle.min.js"></script>


  <script>
    'use strict';

    /* ─── Constantes ─── */
    const REPOSITORIO      = 'CheckListTransporteGrade.asp';
    const ENDPOINT_FORM    = 'CheckListTransporte.asp';   // API do formulário de cadastro
    const TOAST_DELAY      = 4500;
    /* ── Toast canônico ── */
    const TOAST_ICON = {
      sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754' },
      erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545' },
      aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404' },
      info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d' }
    };

    const estado = { parStatus: '', parValor: '', parDataDe: '', parDataAte: '', edtLinhas: '50' };
    let cacheTiposVeiculo = null;

    /* ─── Toast ─── */
    function exibirToast(tipo, mensagem) {
      const t      = TOAST_ICON[tipo] || TOAST_ICON.info;
      const barCor = tipo === 'aviso' ? '#ffc107' : t.cor;
      const el = $(`
        <div class="toast-custom p-3">
          <div class="d-flex align-items-start gap-2">
            <i class="bi ${t.icon} mt-1" style="color:${t.cor};font-size:1.1rem;"></i>
            <div class="flex-grow-1">
              <div style="font-size:.85rem;font-weight:600;color:${t.cor};">${mensagem}</div>
            </div>
            <button class="btn-close btn-close-sm ms-auto" style="font-size:.7rem;"
                    onclick="this.closest('.toast-custom').remove()"></button>
          </div>
          <div class="mt-2" style="height:3px;background:#e9ecef;border-radius:2px;">
            <div class="toast-progress-bar"
                 style="height:100%;background:${barCor};border-radius:2px;"></div>
          </div>
        </div>`);
      $('#toastContainer').append(el);
      setTimeout(() => el.remove(), TOAST_DELAY);
    }

    /* ─── Overlay grade ─── */
    function mostrarOverlay() { $('#overlayLoading').removeClass('d-none'); }
    function ocultarOverlay() { $('#overlayLoading').addClass('d-none');    }

    /* ─── Badges de status ─── */
    function badgeClass(status) {
      switch (status) {
        case 'PATIO HJ'  : return 'badge-patio';
        case 'PORTARIA'  : return 'badge-portaria';
        case 'EXPEDIÇÃO' : return 'badge-expedicao';
        case 'FINALIZADO': return 'badge-finalizado';
        default          : return 'badge-outros';
      }
    }

    /* ─── Renderiza tabela ─── */
    function renderizarTabela(dados) {
      const wrapper = $('#wrapperTabela');

      if (!dados || dados.length === 0) {
        wrapper.html(`
          <div class="table-placeholder">
            <i class="bi bi-clipboard-data"></i>
            <p>Nenhum registro encontrado para os filtros aplicados.</p>
          </div>`);
        $('#infoRegistros').text('');
        return;
      }

      let html = `
        <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaChecklist">
          <thead>
            <tr>
              <th>#</th>
              <th>Proprietário</th>
              <th>Motorista</th>
              <th>Placa</th>
              <th>Destino</th>
              <th>Data</th>
              <th>Status</th>
              <th colspan="2">Operação</th>
              <th colspan="4">Mudar Status</th>
            </tr>
          </thead>
          <tbody>`;

      dados.forEach(function(r) {
        const disabled = r.finalizado == 1 ? 'disabled' : '';
        const btnOp    = r.finalizado == 1
          ? `<button type="button" class="btn btn-primary btn-sm" title="Visualizar"
                     onclick="abrirVisualizar(${r.id})"><i class="bi bi-eye"></i></button>`
          : `<button type="button" class="btn btn-warning btn-sm" title="Alterar"
                     onclick="abrirAlterar(${r.id})"><i class="bi bi-pencil"></i></button>`;

        html += `
            <tr>
              <td>${r.id}</td>
              <td>${r.proprietario}</td>
              <td>${r.motorista}</td>
              <td>${r.placa}</td>
              <td>${r.destino}</td>
              <td>${r.data}</td>
              <td><span class="badge-status ${badgeClass(r.status)}">${r.status}</span></td>
              <td colspan="2">
                <div class="d-flex gap-1 justify-content-center">
                  ${btnOp}
                  <button type="button" class="btn btn-danger btn-sm" title="Excluir"
                          ${disabled} onclick="abrirModalExcluir(${r.id})">
                    <i class="bi bi-trash"></i>
                  </button>
                </div>
              </td>
              <td colspan="4">
                <div class="d-flex gap-1 justify-content-center flex-wrap">
                  <button type="button" class="btn btn-info btn-sm" title="Mover para Pátio"
                          ${disabled} onclick="mudarStatus(${r.id}, 'PAT')">
                    <i class="bi bi-truck"></i>
                  </button>
                  <button type="button" class="btn btn-warning btn-sm" title="Mover para Portaria"
                          ${disabled} onclick="mudarStatus(${r.id}, 'POR')">
                    <i class="bi bi-door-open"></i>
                  </button>
                  <button type="button" class="btn btn-primary btn-sm" title="Mover para Expedição"
                          ${disabled} onclick="mudarStatus(${r.id}, 'EXP')">
                    <i class="bi bi-box-seam"></i>
                  </button>
                  <button type="button" class="btn btn-success btn-sm" title="Finalizar"
                          ${disabled} onclick="mudarStatus(${r.id}, 'FIN')">
                    <i class="bi bi-check-circle"></i>
                  </button>
                </div>
              </td>
            </tr>`;
      });

      html += `</tbody></table>`;
      wrapper.html(html);
      $('#infoRegistros').text('Total: ' + dados.length + ' registro(s)');
    }

    /* ─── Carregar dados da grade ─── */
    function carregarDados() {
      estado.edtLinhas = $('#edtLinhas').val();
      mostrarOverlay();

      $.ajax({
        url      : REPOSITORIO,
        method   : 'GET',
        dataType : 'json',
        data     : {
          acao       : 'listar',
          parStatus  : estado.parStatus,
          parValor   : estado.parValor,
          parDataDe  : estado.parDataDe,
          parDataAte : estado.parDataAte,
          edtLinhas  : estado.edtLinhas
        },
        success: function(resp) {
          ocultarOverlay();
          if (resp.sucesso) {
            renderizarTabela(resp.dados);
          } else {
            exibirToast('erro', resp.mensagem || 'Erro ao carregar dados.');
            renderizarTabela([]);
          }
        },
        error: function() {
          ocultarOverlay();
          exibirToast('erro', 'Falha na comunicação com o servidor.');
          renderizarTabela([]);
        }
      });
    }

    /* ─── Carregar tipos de veículo (com cache) ─── */
    function carregarTiposVeiculo(callback) {
      if (cacheTiposVeiculo) { callback(cacheTiposVeiculo); return; }
      $.ajax({
        url      : REPOSITORIO,
        method   : 'GET',
        dataType : 'json',
        data     : { acao: 'obterTiposVeiculo' },
        success: function(resp) {
          if (resp.sucesso) { cacheTiposVeiculo = resp.dados; callback(resp.dados); }
        }
      });
    }

    /* ─── Helpers de formulário ─── */
    function secao(icon, titulo, corpo) {
      return `<div class="form-section">
                <div class="form-section-header"><i class="bi ${icon} me-2"></i>${titulo}</div>
                <div class="form-section-body">${corpo}</div>
              </div>`;
    }

    function htmlItens(itens, readonly) {
      const dis = readonly ? 'disabled' : '';
      let rows = '';
      itens.forEach(function(it) {
        const checkedSim = it.sim == 1 ? 'checked' : '';
        const checkedNao = it.sim != 1 ? 'checked' : '';
        rows += `<tr>
          <td class="text-start">${it.descricao}</td>
          <td class="text-center"><input type="radio" class="form-check-input item-radio"
              name="item_${it.id}" value="${it.id}" ${checkedSim} ${dis} /></td>
          <td class="text-center"><input type="radio" class="form-check-input"
              name="item_${it.id}" value="0" ${checkedNao} ${dis} /></td>
        </tr>`;
      });
      return `<div class="table-responsive">
        <table class="table table-bordered table-sm mb-0 table-itens">
          <thead><tr>
            <th class="text-start" style="width:70%">Item Avaliado</th>
            <th>Atende</th><th>Não Atende</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table></div>`;
    }

    function htmlSelectTipoVeiculo(tipos, idSelecionado, readonly) {
      const dis = readonly ? 'disabled' : '';
      let opts = '<option value="0">Selecione...</option>';
      tipos.forEach(function(t) {
        const sel = t.id == idSelecionado ? 'selected' : '';
        opts += `<option value="${t.id}" ${sel}>${t.descricao}</option>`;
      });
      return `<select id="cboTipoVeiculo" class="form-select form-select-sm" ${dis}>${opts}</select>`;
    }

    /* ─── Máscara de horário ─── */
    function aplicarMascaraHorario(sel) {
      $(sel).off('input').on('input', function() {
        let v = $(this).val().replace(/\D/g, '').substring(0, 4);
        if (v.length >= 3) v = v.substring(0, 2) + ':' + v.substring(2);
        $(this).val(v);
      });
    }

    /* ============================================================
       MODAL — NOVO CHECKLIST
       Consome CheckListTransporte.asp:obterDadosIniciais,
       buscarPlaca e salvar — tudo via AJAX, sem trocar de página.
    ============================================================ */
    function abrirModalNovo() {
      $('#modalNovoCorpo').html('');
      $('#modalNovoOverlay').removeClass('d-none');

      const modal = new bootstrap.Modal(document.getElementById('modalNovo'));
      modal.show();

      $.ajax({
        url      : ENDPOINT_FORM,
        method   : 'GET',
        dataType : 'json',
        data     : { acao: 'obterDadosIniciais' },
        success: function(resp) {
          $('#modalNovoOverlay').addClass('d-none');

          if (!resp.sucesso) {
            exibirToast('erro', resp.mensagem || 'Erro ao carregar formulário.');
            return;
          }

          /* Monta itens avaliados sem <template> (disponível inline aqui) */
          let rowsItens = '';
          resp.dados.itensAvaliados.forEach(function(it) {
            rowsItens += `<tr>
              <td class="text-start">${it.descricao}</td>
              <td class="text-center">
                <input type="radio" class="form-check-input novo-item-sim"
                       name="novoitem_${it.id}" value="${it.id}" checked />
              </td>
              <td class="text-center">
                <input type="radio" class="form-check-input"
                       name="novoitem_${it.id}" value="0" />
              </td>
            </tr>`;
          });

          /* Monta select tipos de veículo */
          let optsTV = '<option value="0">Selecione...</option>';
          resp.dados.tiposVeiculo.forEach(function(tv) {
            optsTV += `<option value="${tv.id}">${tv.descricao}</option>`;
          });

          const hoje = new Date();
          const dataHoje = String(hoje.getDate()).padStart(2,'0') + '/'
                         + String(hoje.getMonth()+1).padStart(2,'0') + '/'
                         + hoje.getFullYear();

          $('#modalNovoCorpo').html(`
            ${secao('bi-truck','Dados do Veículo / Proprietário',`
              <div class="row g-3">
                <div class="col-md-6">
                  <label class="form-label">Nome do Proprietário <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_nomeProprietario" maxlength="150" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Telefone</label>
                  <input type="text" class="form-control form-control-sm" id="nov_telefone" maxlength="15" />
                </div>
              </div>
              <div class="row g-3 mt-1">
                <div class="col-md-6">
                  <label class="form-label">Nome do Motorista <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_nomeMotorista" maxlength="150" />
                </div>
              </div>
              <div class="row g-3 mt-1">
                <div class="col-md-3">
                  <label class="form-label">CPF / CNPJ <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_cnpjCpf" maxlength="18" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Placa do Veículo <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_placa" maxlength="8" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Data de Preenchimento <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_data"
                         value="${dataHoje}" maxlength="10" placeholder="DD/MM/AAAA" readonly />
                </div>
              </div>
              <div class="row g-3 mt-1">
                <div class="col-md-5">
                  <label class="form-label">Tipo de Veículo <span class="text-danger">*</span></label>
                  <select id="nov_tipoVeiculo" class="form-select form-select-sm">
                    ${optsTV}
                  </select>
                </div>
                <div class="col-md-2">
                  <label class="form-label">Qtd. Eixos <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_qtdeEixo" maxlength="5" />
                </div>
              </div>
              <div class="row g-3 mt-1">
                <div class="col-md-3">
                  <label class="form-label">Horário Início Carregamento <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_horarioInicio" maxlength="5" placeholder="HH:MM" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Horário Término Carregamento</label>
                  <input type="text" class="form-control form-control-sm" id="nov_horarioTermino" maxlength="5" placeholder="HH:MM" />
                </div>
              </div>`)}
            ${secao('bi-signpost-split','Tipo de Transporte',`
              <div class="d-flex flex-wrap gap-4">
                <label class="radio-opcao">
                  <input type="radio" name="nov_rdTipo" value="1" /> Smurfit Westrock
                </label>
                <label class="radio-opcao">
                  <input type="radio" name="nov_rdTipo" value="2" checked /> Transportadora
                </label>
                <label class="radio-opcao">
                  <input type="radio" name="nov_rdTipo" value="3" /> Freteiros
                </label>
              </div>
              <div class="row g-3 mt-2">
                <div class="col-md-6">
                  <label class="form-label">Destino da Carga <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_destinoCarga" />
                </div>
              </div>`)}
            ${secao('bi-list-check','Itens Avaliados',`
              <div class="table-responsive">
                <table class="table table-bordered table-sm mb-0 table-itens">
                  <thead><tr>
                    <th class="text-start" style="width:70%">Item Avaliado</th>
                    <th>Atende</th><th>Não Atende</th>
                  </tr></thead>
                  <tbody>${rowsItens}</tbody>
                </table>
              </div>`)}
            ${secao('bi-rulers','Dimensões do Veículo',`
              <div class="row g-3">
                <div class="col-md-3">
                  <label class="form-label">Altura (m) <span class="text-danger">*</span></label>
                  <input type="number" class="form-control form-control-sm" id="nov_altura" step=".01" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Comprimento (m) <span class="text-danger">*</span></label>
                  <input type="number" class="form-control form-control-sm" id="nov_comprimento" step=".01" />
                </div>
                <div class="col-md-3">
                  <label class="form-label">Largura (m) <span class="text-danger">*</span></label>
                  <input type="number" class="form-control form-control-sm" id="nov_largura" step=".01" />
                </div>
              </div>
              <div class="mt-3">
                <label class="form-label d-block">Manutenção do veículo em dia? <span class="text-danger">*</span></label>
                <div class="d-flex gap-4">
                  <label class="radio-opcao">
                    <input type="radio" name="nov_rdManut" value="1" checked /> <strong>Sim</strong>
                  </label>
                  <label class="radio-opcao">
                    <input type="radio" name="nov_rdManut" value="0" /> <strong>Não</strong>
                  </label>
                </div>
              </div>`)}
            ${secao('bi-person-check','Conferência',`
              <div class="row g-3">
                <div class="col-md-6">
                  <label class="form-label">Nome do Conferente <span class="text-danger">*</span></label>
                  <input type="text" class="form-control form-control-sm" id="nov_conferente" />
                </div>
                <div class="col-md-6">
                  <label class="form-label">Responsável na Portaria</label>
                  <input type="text" class="form-control form-control-sm" id="nov_portariaResponsavel" />
                </div>
              </div>
              <div class="row g-3 mt-1">
                <div class="col-md-3">
                  <label class="form-label">Status <span class="text-danger">*</span></label>
                  <select id="nov_status" class="form-select form-select-sm">
                    <option value="0">Selecione...</option>
                    <option value="PATIO HJ">Pátio HJ</option>
                    <option value="PORTARIA">Portaria</option>
                    <option value="EXPEDIÇÃO">Expedição</option>
                    <option value="FINALIZADO">Finalizado</option>
                  </select>
                </div>
              </div>`)}
            ${secao('bi-info-circle','Notas',`
              <ol class="notas-lista mb-0">
                <li>O não cumprimento a qualquer um dos itens de avaliação exclui o motorista de carregar o caminhão.</li>
                <li>Será impossibilitado de viajar o motorista que apresentar sintomas de embriaguez ou estado físico alterado.</li>
                <li>Também será excluído do cadastro o motorista que causar qualquer tipo de problemas com os clientes da Smurfit Westrock.</li>
              </ol>
              <p class="mt-2 mb-0 text-muted" style="font-size:.75rem;"><br><strong>RGQ105rev12</strong><br>Ed. Original: 26/09/2013<br>Revisão: 01/06/2026</p>`)}`
          );

          /* Máscara horário */
          aplicarMascaraHorario('#nov_horarioInicio');
          aplicarMascaraHorario('#nov_horarioTermino');

          /* Auto-preenchimento pela placa */
          $('#nov_placa').off('blur').on('blur', function() {
            const placa = $(this).val().trim();
            if (!placa) return;
            $.ajax({
              url      : ENDPOINT_FORM,
              method   : 'GET',
              dataType : 'json',
              data     : { acao: 'buscarPlaca', placa: placa },
              success: function(r) {
                if (!r.sucesso) { exibirToast('aviso', r.mensagem || 'Placa não encontrada.'); return; }
                const d = r.dados;
                $('#nov_nomeProprietario').val(d.nomeProprietario);
                $('#nov_telefone').val(d.telefone);
                $('#nov_cnpjCpf').val(d.cnpjCpf);
                $('#nov_nomeMotorista').val(d.nomeMotorista);
                $('#nov_tipoVeiculo').val(d.idTipoVeiculo);
                $('#nov_qtdeEixo').val(d.quantidadeEixo);
                $('#nov_altura').val(d.altura);
                $('#nov_comprimento').val(d.comprimento);
                $('#nov_largura').val(d.largura);
              }
            });
          });
        },
        error: function() {
          $('#modalNovoOverlay').addClass('d-none');
          exibirToast('erro', 'Não foi possível carregar o formulário.');
        }
      });

      /* Vincula o botão Gravar */
      $('#btnGravarNovo').off('click').on('click', function() { gravarNovo(); });
    }

    /* ─── Helpers de validação inline ─── */
    function limparValidacao(prefixo) {
      const sels = prefixo === 'nov'
        ? ['#nov_nomeProprietario','#nov_nomeMotorista','#nov_cnpjCpf','#nov_placa',
           '#nov_tipoVeiculo','#nov_qtdeEixo','#nov_horarioInicio','#nov_destinoCarga',
           '#nov_altura','#nov_comprimento','#nov_largura','#nov_conferente','#nov_status']
        : ['#alt_nomeProprietario','#alt_nomeMotorista','#alt_documento','#alt_placa',
           '#alt_dataPreenchimento','#cboTipoVeiculo','#alt_qtdeEixo','#alt_horarioInicio',
           '#alt_destinoCarga','#alt_altura','#alt_comprimento','#alt_largura',
           '#alt_conferente','#alt_status'];
      sels.forEach(function(s) {
        $(s).removeClass('is-invalid');
        $(s).siblings('.invalid-feedback').remove();
      });
    }

    function marcarInvalido(sel, msg) {
      $(sel).addClass('is-invalid');
      if (!$(sel).siblings('.invalid-feedback').length) {
        $(sel).after('<div class="invalid-feedback">' + msg + '</div>');
      }
    }

    function validarFormulario(prefixo) {
      limparValidacao(prefixo);
      let ok = true;

      function req(sel, msg) {
        const v = $(sel).val();
        if (!v || v === '0' || !v.trim()) { marcarInvalido(sel, msg); ok = false; }
      }

      if (prefixo === 'nov') {
        req('#nov_nomeProprietario', 'Informe o nome do proprietário.');
        req('#nov_nomeMotorista',    'Informe o nome do motorista.');
        req('#nov_cnpjCpf',         'Informe o CPF/CNPJ.');
        req('#nov_placa',           'Informe a placa do veículo.');
        req('#nov_tipoVeiculo',     'Selecione o tipo de veículo.');
        req('#nov_qtdeEixo',        'Informe a quantidade de eixos.');
        req('#nov_horarioInicio',   'Informe o horário de início.');
        req('#nov_destinoCarga',    'Informe o destino da carga.');
        req('#nov_altura',          'Informe a altura do veículo.');
        req('#nov_comprimento',     'Informe o comprimento do veículo.');
        req('#nov_largura',         'Informe a largura do veículo.');
        req('#nov_conferente',      'Informe o nome do conferente.');
        req('#nov_status',          'Selecione o status.');
      } else {
        req('#alt_nomeProprietario',  'Informe o nome do proprietário.');
        req('#alt_nomeMotorista',     'Informe o nome do motorista.');
        req('#alt_documento',         'Informe o CPF/CNPJ.');
        req('#alt_placa',             'Informe a placa do veículo.');
        req('#alt_dataPreenchimento', 'Informe a data de preenchimento.');
        req('#cboTipoVeiculo',        'Selecione o tipo de veículo.');
        req('#alt_qtdeEixo',          'Informe a quantidade de eixos.');
        req('#alt_horarioInicio',     'Informe o horário de início.');
        req('#alt_destinoCarga',      'Informe o destino da carga.');
        req('#alt_altura',            'Informe a altura do veículo.');
        req('#alt_comprimento',       'Informe o comprimento do veículo.');
        req('#alt_largura',           'Informe a largura do veículo.');
        req('#alt_conferente',        'Informe o nome do conferente.');
        req('#alt_status',            'Selecione o status.');

        if (!$('input[name="alt_tipoTransporte"]:checked').val()) {
          exibirToast('aviso', 'Selecione o tipo de transporte.');
          ok = false;
        }
        if ($('input[name="alt_manutencao"]:checked').val() === undefined) {
          exibirToast('aviso', 'Informe se a manutenção do veículo está em dia.');
          ok = false;
        }
      }

      if (!ok) {
        // Rola até o primeiro campo inválido
        const primeiro = $('.is-invalid').first();
        if (primeiro.length) {
          primeiro[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }
      return ok;
    }

    /* ─── Gravar novo checklist via AJAX ─── */
    function gravarNovo() {
      if (!validarFormulario('nov')) return;

      /* Coleta itens que atendem */
      let itensAvaliados = '';
      $('#modalNovoCorpo input.novo-item-sim:checked').each(function() {
        const v = $(this).val();
        if (v && v !== '0') itensAvaliados += (itensAvaliados ? ',' : '') + v;
      });

      $('#modalNovoOverlay').removeClass('d-none');

      $.ajax({
        url      : ENDPOINT_FORM,
        method   : 'POST',
        dataType : 'json',
        data     : {
          acao               : 'salvar',
          nomeProprietario   : $('#nov_nomeProprietario').val().trim(),
          telefone           : $('#nov_telefone').val().trim(),
          nomeMotorista      : $('#nov_nomeMotorista').val().trim(),
          cnpjCpf            : $('#nov_cnpjCpf').val().trim(),
          placa              : $('#nov_placa').val().trim(),
          dataPreenchimento  : $('#nov_data').val().trim(),
          idTipoVeiculo      : $('#nov_tipoVeiculo').val(),
          quantidadeEixo     : $('#nov_qtdeEixo').val().trim(),
          horarioInicio      : $('#nov_horarioInicio').val().trim(),
          horarioTermino     : $('#nov_horarioTermino').val().trim(),
          tipoTransporte     : $('input[name="nov_rdTipo"]:checked').val(),
          destinoCarga       : $('#nov_destinoCarga').val().trim(),
          altura             : $('#nov_altura').val().trim(),
          comprimento        : $('#nov_comprimento').val().trim(),
          largura            : $('#nov_largura').val().trim(),
          manutencao         : $('input[name="nov_rdManut"]:checked').val(),
          nomeConferente     : $('#nov_conferente').val().trim(),
          portariaResponsavel: $('#nov_portariaResponsavel').val().trim(),
          status             : $('#nov_status').val(),
          itensAvaliados     : itensAvaliados
        },
        success: function(resp) {
          $('#modalNovoOverlay').addClass('d-none');
          if (resp.sucesso) {
            bootstrap.Modal.getInstance(document.getElementById('modalNovo')).hide();
            exibirToast('sucesso', resp.mensagem || 'CheckList gravado com sucesso.');
            carregarDados();
          } else {
            exibirToast('erro', resp.mensagem || 'Erro ao gravar.');
          }
        },
        error: function() {
          $('#modalNovoOverlay').addClass('d-none');
          exibirToast('erro', 'Falha na comunicação com o servidor.');
        }
      });
    }

    /* ─── Alterar ─── */
    function montarFormularioAlterar(d, tipos) {
      const selTipo = htmlSelectTipoVeiculo(tipos, d.idTipoVeiculo, false);

      const secVeiculo = secao('bi-truck', 'Dados do Veículo / Proprietário', `
        <div class="row g-3">
          <div class="col-md-6">
            <label class="form-label">Nome do Proprietário <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_nomeProprietario"
                   value="${d.nomeProprietario}" maxlength="150" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Telefone</label>
            <input type="text" class="form-control form-control-sm" id="alt_telefone"
                   value="${d.telefone}" maxlength="15" />
          </div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-6">
            <label class="form-label">Nome do Motorista <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_nomeMotorista"
                   value="${d.nomeMoto}" maxlength="150" />
          </div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">
            <label class="form-label">CPF / CNPJ <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_documento"
                   value="${d.documento}" maxlength="18" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Placa do Veículo <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_placa"
                   value="${d.placa}" maxlength="8" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Data de Preenchimento <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_dataPreenchimento"
                   value="${d.dataPreenchimento}" maxlength="10" placeholder="DD/MM/AAAA" />
          </div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-5">
            <label class="form-label">Tipo de Veículo <span class="text-danger">*</span></label>
            ${selTipo}
          </div>
          <div class="col-md-2">
            <label class="form-label">Qtd. Eixos <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_qtdeEixo"
                   value="${d.qtdeEixo}" maxlength="5" />
          </div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">
            <label class="form-label">Horário Início Carregamento <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_horarioInicio"
                   value="${d.horarioInicio}" maxlength="5" placeholder="HH:MM" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Horário Término Carregamento</label>
            <input type="text" class="form-control form-control-sm" id="alt_horarioTermino"
                   value="${d.horarioTermino}" maxlength="5" placeholder="HH:MM" />
          </div>
        </div>`);

      const checkedS = d.idTipoTransportadora == 1 ? 'checked' : '';
      const checkedT = d.idTipoTransportadora == 2 ? 'checked' : '';
      const checkedF = d.idTipoTransportadora == 3 ? 'checked' : '';
      const secTransporte = secao('bi-signpost-split', 'Tipo de Transporte', `
        <div class="d-flex flex-wrap gap-4">
          <label class="radio-opcao">
            <input type="radio" id="alt_tipoS" name="alt_tipoTransporte" value="1" ${checkedS} />
            Smurfit Westrock
          </label>
          <label class="radio-opcao">
            <input type="radio" id="alt_tipoT" name="alt_tipoTransporte" value="2" ${checkedT} />
            Transportadora
          </label>
          <label class="radio-opcao">
            <input type="radio" id="alt_tipoF" name="alt_tipoTransporte" value="3" ${checkedF} />
            Freteiros
          </label>
        </div>
        <div class="row g-3 mt-2">
          <div class="col-md-6">
            <label class="form-label">Destino da Carga <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_destinoCarga"
                   value="${d.destinoCarga}" />
          </div>
        </div>`);

      const secItens = secao('bi-list-check', 'Itens Avaliados', htmlItens(d.itens, false));

      const checkedMS = d.manutencao == 1 ? 'checked' : '';
      const checkedMN = d.manutencao != 1 ? 'checked' : '';
      const secDimensoes = secao('bi-rulers', 'Dimensões do Veículo', `
        <div class="row g-3">
          <div class="col-md-3">
            <label class="form-label">Altura (m) <span class="text-danger">*</span></label>
            <input type="number" class="form-control form-control-sm" id="alt_altura"
                   value="${d.altura}" step=".01" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Comprimento (m) <span class="text-danger">*</span></label>
            <input type="number" class="form-control form-control-sm" id="alt_comprimento"
                   value="${d.comprimento}" step=".01" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Largura (m) <span class="text-danger">*</span></label>
            <input type="number" class="form-control form-control-sm" id="alt_largura"
                   value="${d.largura}" step=".01" />
          </div>
        </div>
        <div class="mt-3">
          <label class="form-label d-block">Manutenção do veículo em dia? <span class="text-danger">*</span></label>
          <div class="d-flex gap-4">
            <label class="radio-opcao">
              <input type="radio" id="alt_manutSim" name="alt_manutencao" value="1" ${checkedMS} />
              <strong>Sim</strong>
            </label>
            <label class="radio-opcao">
              <input type="radio" id="alt_manutNao" name="alt_manutencao" value="0" ${checkedMN} />
              <strong>Não</strong>
            </label>
          </div>
        </div>`);

      const secConferencia = secao('bi-person-check', 'Conferência', `
        <div class="row g-3">
          <div class="col-md-6">
            <label class="form-label">Nome do Conferente <span class="text-danger">*</span></label>
            <input type="text" class="form-control form-control-sm" id="alt_conferente"
                   value="${d.conferente}" />
          </div>
          <div class="col-md-6">
            <label class="form-label">Responsável na Portaria</label>
            <input type="text" class="form-control form-control-sm" id="alt_responsavelPortaria"
                   value="${d.responsavelPortaria}" />
          </div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">
            <label class="form-label">Status <span class="text-danger">*</span></label>
            <select id="alt_status" class="form-select form-select-sm">
              <option value="0">Selecione...</option>
              <option value="PATIO HJ"  ${d.status==='PATIO HJ'   ? 'selected':''}>Pátio HJ</option>
              <option value="PORTARIA"  ${d.status==='PORTARIA'   ? 'selected':''}>Portaria</option>
              <option value="EXPEDIÇÃO" ${d.status==='EXPEDIÇÃO'  ? 'selected':''}>Expedição</option>
              <option value="FINALIZADO"${d.status==='FINALIZADO' ? 'selected':''}>Finalizado</option>
            </select>
          </div>
        </div>`);

      const secNotas = secao('bi-info-circle', 'Notas', `
        <ol class="notas-lista mb-0">
          <li>O não cumprimento a qualquer um dos itens de avaliação exclui o motorista de carregar o caminhão.</li>
          <li>Será impossibilitado de viajar o motorista que apresentar sintomas de embriaguez ou estado físico alterado.</li>
          <li>Também será excluído do cadastro o motorista que causar qualquer tipo de problemas com os clientes da Smurfit Westrock.</li>
        </ol>
        <p class="mt-2 mb-0 text-muted" style="font-size:.75rem;">RGQ105rev12</strong><br>Ed. Original: 26/09/2013<br>Revisão: 01/06/2026</p>`);

      return secVeiculo + secTransporte + secItens + secDimensoes + secConferencia + secNotas;
    }

    function montarFormularioVisualizar(d, tipos) {
      const selTipo = htmlSelectTipoVeiculo(tipos, d.idTipoVeiculo, true);

      function campo(label, valor) {
        return `<div class="mb-1">
          <label class="form-label">${label}</label>
          <input type="text" class="form-control form-control-sm" value="${valor}" disabled readonly />
        </div>`;
      }

      const secVeiculo = secao('bi-truck', 'Dados do Veículo / Proprietário', `
        <div class="row g-3">
          <div class="col-md-6">${campo('Nome do Proprietário', d.nomeProprietario)}</div>
          <div class="col-md-3">${campo('Telefone', d.telefone)}</div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-6">${campo('Nome do Motorista', d.nomeMoto)}</div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">${campo('CPF / CNPJ', d.documento)}</div>
          <div class="col-md-3">${campo('Placa do Veículo', d.placa)}</div>
          <div class="col-md-3">${campo('Data de Preenchimento', d.dataPreenchimento)}</div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-5">
            <label class="form-label">Tipo de Veículo</label>
            ${selTipo}
          </div>
          <div class="col-md-2">${campo('Qtd. Eixos', d.qtdeEixo)}</div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">${campo('Horário Início', d.horarioInicio)}</div>
          <div class="col-md-3">${campo('Horário Término', d.horarioTermino)}</div>
        </div>`);

      const cS = d.idTipoTransportadora == 1 ? 'checked' : '';
      const cT = d.idTipoTransportadora == 2 ? 'checked' : '';
      const cF = d.idTipoTransportadora == 3 ? 'checked' : '';
      const secTransporte = secao('bi-signpost-split', 'Tipo de Transporte', `
        <div class="d-flex flex-wrap gap-4">
          <label class="radio-opcao">
            <input type="radio" name="viz_tipo" value="1" ${cS} disabled /> Smurfit Westrock
          </label>
          <label class="radio-opcao">
            <input type="radio" name="viz_tipo" value="2" ${cT} disabled /> Transportadora
          </label>
          <label class="radio-opcao">
            <input type="radio" name="viz_tipo" value="3" ${cF} disabled /> Freteiros
          </label>
        </div>
        <div class="row g-3 mt-2">
          <div class="col-md-6">${campo('Destino da Carga', d.destinoCarga)}</div>
        </div>`);

      const secItens = secao('bi-list-check', 'Itens Avaliados', htmlItens(d.itens, true));

      const mS = d.manutencao == 1 ? 'checked' : '';
      const mN = d.manutencao != 1 ? 'checked' : '';
      const secDimensoes = secao('bi-rulers', 'Dimensões do Veículo', `
        <div class="row g-3">
          <div class="col-md-3">${campo('Altura (m)', d.altura)}</div>
          <div class="col-md-3">${campo('Comprimento (m)', d.comprimento)}</div>
          <div class="col-md-3">${campo('Largura (m)', d.largura)}</div>
        </div>
        <div class="mt-3">
          <label class="form-label d-block">Manutenção do veículo em dia?</label>
          <div class="d-flex gap-4">
            <label class="radio-opcao">
              <input type="radio" name="viz_manut" value="1" ${mS} disabled /> <strong>Sim</strong>
            </label>
            <label class="radio-opcao">
              <input type="radio" name="viz_manut" value="0" ${mN} disabled /> <strong>Não</strong>
            </label>
          </div>
        </div>`);

      const statusOpts = [
        {v: 'PATIO HJ',   l: 'Pátio HJ'},
        {v: 'PORTARIA',   l: 'Portaria'},
        {v: 'EXPEDIÇÃO',  l: 'Expedição'},
        {v: 'FINALIZADO', l: 'Finalizado'}
      ].map(o => `<option value="${o.v}" ${d.status === o.v ? 'selected':''}>${o.l}</option>`).join('');

      const secConferencia = secao('bi-person-check', 'Conferência', `
        <div class="row g-3">
          <div class="col-md-6">${campo('Nome do Conferente', d.conferente)}</div>
          <div class="col-md-6">${campo('Responsável na Portaria', d.responsavelPortaria)}</div>
        </div>
        <div class="row g-3 mt-1">
          <div class="col-md-3">
            <label class="form-label">Status</label>
            <select class="form-select form-select-sm" disabled>
              <option value="0">Selecione...</option>
              ${statusOpts}
            </select>
          </div>
        </div>`);

      const secNotas = secao('bi-info-circle', 'Notas', `
        <ol class="notas-lista mb-0">
          <li>O não cumprimento a qualquer um dos itens de avaliação exclui o motorista de carregar o caminhão.</li>
          <li>Será impossibilitado de viajar o motorista que apresentar sintomas de embriaguez ou estado físico alterado.</li>
          <li>Também será excluído do cadastro o motorista que causar qualquer tipo de problemas com os clientes da Smurfit Westrock.</li>
        </ol>
        <p class="mt-2 mb-0 text-muted" style="font-size:.75rem;">RGQ105rev12</strong><br>Ed. Original: 26/09/2013<br>Revisão: 01/06/2026</p>`);

      return secVeiculo + secTransporte + secItens + secDimensoes + secConferencia + secNotas;
    }

    function abrirAlterar(id) {
      $('#lblAlterarId').text('#' + id);
      $('#modalAlterarCorpo').html('');
      $('#modalAlterarOverlay').removeClass('d-none');

      const modal = new bootstrap.Modal(document.getElementById('modalAlterar'));
      modal.show();

      const promDados = $.ajax({
        url: REPOSITORIO, method: 'GET', dataType: 'json',
        data: { acao: 'obterRegistro', parID: id }
      });

      carregarTiposVeiculo(function(tipos) {
        promDados.done(function(resp) {
          $('#modalAlterarOverlay').addClass('d-none');
          if (!resp.sucesso) {
            exibirToast('erro', resp.mensagem || 'Erro ao carregar registro.');
            return;
          }
          $('#modalAlterarCorpo').html(montarFormularioAlterar(resp.dados, tipos));
          aplicarMascaraHorario('#alt_horarioInicio');
          aplicarMascaraHorario('#alt_horarioTermino');
        }).fail(function() {
          $('#modalAlterarOverlay').addClass('d-none');
          exibirToast('erro', 'Falha ao carregar dados do registro.');
        });
      });

      $('#btnSalvarAlteracao').off('click').on('click', function() {
        salvarAlteracao(id);
      });
    }

    function abrirVisualizar(id) {
      $('#lblVisualizarId').text('#' + id);
      $('#modalVisualizarCorpo').html('');
      $('#modalVisualizarOverlay').removeClass('d-none');

      const modal = new bootstrap.Modal(document.getElementById('modalVisualizar'));
      modal.show();

      const promDados = $.ajax({
        url: REPOSITORIO, method: 'GET', dataType: 'json',
        data: { acao: 'obterRegistro', parID: id }
      });

      carregarTiposVeiculo(function(tipos) {
        promDados.done(function(resp) {
          $('#modalVisualizarOverlay').addClass('d-none');
          if (!resp.sucesso) {
            exibirToast('erro', resp.mensagem || 'Erro ao carregar registro.');
            return;
          }
          $('#modalVisualizarCorpo').html(montarFormularioVisualizar(resp.dados, tipos));
        }).fail(function() {
          $('#modalVisualizarOverlay').addClass('d-none');
          exibirToast('erro', 'Falha ao carregar dados do registro.');
        });
      });
    }

    function salvarAlteracao(id) {
      if (!validarFormulario('alt')) return;

      const tipoTransp = $('input[name="alt_tipoTransporte"]:checked').val();
      const manutencao = $('input[name="alt_manutencao"]:checked').val();

      let itensAvaliados = '';
      $('.item-radio:checked').each(function() {
        const v = $(this).val();
        if (v && v !== '0') itensAvaliados += v + '|';
      });

      $('#modalAlterarOverlay').removeClass('d-none');

      $.ajax({
        url      : REPOSITORIO,
        method   : 'POST',
        dataType : 'json',
        data     : {
          acao                : 'salvarAlteracao',
          parID               : id,
          nomeProprietario    : $('#alt_nomeProprietario').val(),
          nomeMotorista       : $('#alt_nomeMotorista').val(),
          telefone            : $('#alt_telefone').val(),
          documento           : $('#alt_documento').val(),
          placa               : $('#alt_placa').val(),
          dataPreenchimento   : $('#alt_dataPreenchimento').val(),
          idTipoVeiculo       : $('#cboTipoVeiculo').val(),
          qtdeEixo            : $('#alt_qtdeEixo').val(),
          horarioInicio       : $('#alt_horarioInicio').val(),
          horarioTermino      : $('#alt_horarioTermino').val(),
          destinoCarga        : $('#alt_destinoCarga').val(),
          idTipoTransportadora: tipoTransp,
          manutencao          : manutencao,
          conferente          : $('#alt_conferente').val(),
          responsavelPortaria : $('#alt_responsavelPortaria').val(),
          comprimento         : $('#alt_comprimento').val(),
          altura              : $('#alt_altura').val(),
          largura             : $('#alt_largura').val(),
          status              : $('#alt_status').val(),
          itensAvaliados      : itensAvaliados
        },
        success: function(resp) {
          $('#modalAlterarOverlay').addClass('d-none');
          if (resp.sucesso) {
            bootstrap.Modal.getInstance(document.getElementById('modalAlterar')).hide();
            exibirToast('sucesso', resp.mensagem);
            carregarDados();
          } else {
            exibirToast('erro', resp.mensagem || 'Erro ao salvar alteração.');
          }
        },
        error: function() {
          $('#modalAlterarOverlay').addClass('d-none');
          exibirToast('erro', 'Falha na comunicação com o servidor.');
        }
      });
    }

    function mudarStatus(id, novoStatus) {
      mostrarOverlay();
      $.ajax({
        url      : REPOSITORIO,
        method   : 'POST',
        dataType : 'json',
        data     : { acao: 'mudarStatus', parID: id, parStatus: novoStatus },
        success: function(resp) {
          if (resp.sucesso) {
            exibirToast('sucesso', resp.mensagem);
            carregarDados();
          } else {
            ocultarOverlay();
            exibirToast('erro', resp.mensagem || 'Erro ao atualizar status.');
          }
        },
        error: function() {
          ocultarOverlay();
          exibirToast('erro', 'Falha na comunicação com o servidor.');
        }
      });
    }

    function abrirModalExcluir(id) {
      $('#modalExcluirId').text('#' + id);
      $('#btnConfirmarExclusao').off('click').on('click', function() {
        const modal = bootstrap.Modal.getInstance(document.getElementById('modalExcluir'));
        modal.hide();
        mostrarOverlay();
        $.ajax({
          url      : REPOSITORIO,
          method   : 'POST',
          dataType : 'json',
          data     : { acao: 'excluir', parID: id },
          success: function(resp) {
            if (resp.sucesso) {
              exibirToast('sucesso', resp.mensagem);
              carregarDados();
            } else {
              ocultarOverlay();
              exibirToast('erro', resp.mensagem || 'Erro ao excluir.');
            }
          },
          error: function() {
            ocultarOverlay();
            exibirToast('erro', 'Falha na comunicação com o servidor.');
          }
        });
      });
      new bootstrap.Modal(document.getElementById('modalExcluir')).show();
    }

    /* ─── Filtros ─── */
    function toggleCampoPesquisa() {
      const val = $('#cboMudaStatus').val();
      const isTexto = (val !== '0' && val !== 'IN' && val !== 'DP' && val !== '');
      const isData  = (val === 'DP');
      $('#wrapperPesquisa').toggle(isTexto);
      $('#wrapperDataDe, #wrapperDataAte').toggle(isData);
      $('#wrapperIndicadores, #wrapperIndicadores2').toggle(val === 'IN');
    }

    function filtrarStatus(s) {
      const map = { H: "'H'", P: "'P'", E: "'E'", F: "'F'", T: '' };
      estado.parStatus = map[s] !== undefined ? map[s] : '';
      estado.parValor  = '';
      posicionarBotaoStatus();
      carregarDados();
    }

    function pesquisar() {
      const tipo = $('#cboMudaStatus').val();
      if (tipo === 'IN') {
        const ini = $('#parDataInicial').val();
        if (ini !== '') {
          window.location.href = 'CheckListTransporteIndicador.asp?parDataInicial='
            + ini + '&parDataFinal=' + $('#parDataFinal').val();
        }
        return;
      }
      estado.parStatus  = tipo;
      if (tipo === 'DP') {
        estado.parValor   = '';
        estado.parDataDe  = $('#parDataDe').val();
        estado.parDataAte = $('#parDataAte').val();
      } else {
        estado.parValor   = $('#txtPesquisa').val();
        estado.parDataDe  = '';
        estado.parDataAte = '';
      }
      carregarDados();
    }

    function limparFiltros() {
      estado.parStatus  = '';
      estado.parValor   = '';
      estado.parDataDe  = '';
      estado.parDataAte = '';
      $('#cboMudaStatus').val('0');
      $('#txtPesquisa').val('');
      $('#parDataDe').val('');
      $('#parDataAte').val('');
      toggleCampoPesquisa();
      posicionarBotaoStatus();
      carregarDados();
    }

    function posicionarBotaoStatus() {
      const map = { "'H'": 'btnPatio', "'P'": 'btnPortaria', "'E'": 'btnExpedicao', "'F'": 'btnFinalizado' };
      $('#btnPatio, #btnPortaria, #btnExpedicao, #btnFinalizado, #btnTodos')
        .removeClass('btn-primary').addClass('btn-light');
      const alvo = map[estado.parStatus];
      if (alvo) {
        $('#' + alvo).removeClass('btn-light').addClass('btn-primary');
      } else {
        $('#btnTodos').removeClass('btn-light').addClass('btn-primary');
      }
    }

    /* ─── Export CSV ─── */
    $('#btnExportCsv').on('click', function() {
      const cab = [];
      $('#tabelaChecklist thead th').each(function() { cab.push($(this).text().trim()); });
      const linhas = [];
      $('#tabelaChecklist tbody tr').each(function() {
        const row = [];
        $(this).find('td').each(function() { row.push($(this).text().trim()); });
        linhas.push(row);
      });
      if (linhas.length === 0) { exibirToast('aviso', 'Nenhum dado para exportar.'); return; }
      exportarCSV('checklist_transportes', cab, linhas);
    });

    function exportarCSV(nomeArquivo, cabecalhos, linhas) {
      const esc = v => {
        const s = v == null ? '' : String(v);
        return (s.includes(';') || s.includes('"') || s.includes('\n'))
          ? '"' + s.replace(/"/g, '""') + '"' : s;
      };
      const rows = [
        cabecalhos.map(esc).join(';'),
        ...linhas.map(row => row.map(esc).join(';'))
      ];
      const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = nomeArquivo.endsWith('.csv') ? nomeArquivo : nomeArquivo + '.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    /* ─── Init ─── */
    $(function() {
      toggleCampoPesquisa();
      posicionarBotaoStatus();
      carregarDados();
    });
  </script>
</body>
</html>
