<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../funcoes/funcoes.asp"-->
<!--#Include File="../../../includes/variaveis.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

' Variáveis globais
Dim acao
Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listar"
        resultado = Listar()
    Case "criar"
        resultado = Criar()
    Case "excluir"
        resultado = Excluir()
    Case "obterPorId"
        resultado = ObterPorId()
    Case "listarUsuarios"
        resultado = ListarUsuarios()
    Case "listarSetores"
        resultado = ListarSetores()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    SQL = "SELECT a.Id as IdUsuario, a.Nome, a.Email, " 
    SQL = SQL & " MIN(a.DataInclusao) AS DataInclusao"
    SQL = SQL & " FROM AprovadorFNC a"
    SQL = SQL & " WHERE (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " GROUP BY a.Id, a.Nome, a.Email"
    SQL = SQL & " ORDER BY a.Nome" 

    On Error Resume Next
    SET rsListar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de observações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """idUsuario"": """ & EscapeJSON(rsListar("IdUsuario") & "") & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsListar("Nome") & "") & """, "
            jsonArray = jsonArray & """email"": """ & EscapeJSON(rsListar("Email") & "") & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao") & "")& """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsListar.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsListar.Close()
    SET rsListar = NOTHING
    
    Listar = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarUsuarios()
    SQL = "SELECT "
    SQL = SQL & "CRM_R_E_C_N_O as Id, "
    SQL = SQL & "CRM_NOME as Nome "
    SQL = SQL & "FROM CRM_DTI "
    SQL = SQL & "WHERE CRM_D_E_L_E_T <> '*' "
    SQL = SQL & "ORDER BY CRM_NOME"

    On Error Resume Next
    SET rsListarUsuarios = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListarUsuarios) Then
        ListarUsuarios = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de usuários: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListarUsuarios.EOF Then
        Do While Not rsListarUsuarios.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListarUsuarios("Id") & "") & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsListarUsuarios("Nome") & "") & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsListarUsuarios.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsListarUsuarios.Close()
    SET rsListarUsuarios = NOTHING
    
    ListarUsuarios = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarSetores()
    SQL = "SELECT Id, Descricao FROM Setor"

    On Error Resume Next
    SET rsSetores = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSetores) Then
        ListarSetores = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de setores: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSetores.EOF Then
        Do While Not rsSetores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsSetores("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsSetores("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSetores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSetores.Close()
    SET rsSetores = NOTHING
    
    ListarSetores = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterPorId()    
    idUsuario = Request("idUsuario")

    If idUsuario = "" Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If
    
    SQL = "SELECT a.Id as IdUsuario, a.Nome, a.Email, " 
    SQL = SQL & " MIN(a.DataInclusao) AS DataInclusao"
    SQL = SQL & " FROM AprovadorFNC a"
    SQL = SQL & " WHERE (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " AND Id = " & idUsuario
    SQL = SQL & " GROUP BY a.Id, a.Nome, a.Email"
    SQL = SQL & " ORDER BY a.Nome" 
    
    On Error Resume Next
    SET rsObter = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar observação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """idUsuario"": """ & EscapeJSON(rsObter("IdUsuario") & "") & """, "
    jsonObjeto = jsonObjeto & """nome"": """ & EscapeJSON(rsObter("Nome") & "") & """, "
    jsonObjeto = jsonObjeto & """email"": """ & EscapeJSON(rsObter("Email") & "") & """, "
    jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao") & "") & """"
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Observação encontrada"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()    
    idUsuario = Request("idUsuario")

    If idUsuario = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Id de usuário não informado""}"
        Exit Function
    End If

    ' Busca dados do aprovador
    SQL = "SELECT "
    SQL = SQL & "CRM_NOME as Nome, "
    SQL = SQL & "CRM_EMAILPRIMARIO as Email "
    SQL = SQL & "FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idUsuario

    SET rsUsuario = Conecta("PIR-HELPDESK", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUsuario) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao dados de usuário: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsUsuario.EOF Then
        nome = rsUsuario("Nome")
        email = rsUsuario("Email")
    End If

    rsUsuario.Close()
    SET rsUsuario = NOTHING

    SQL = "INSERT INTO AprovadorFNC ("
    SQL = SQL & "   Planta, Id, Nome, Email, Nivel, "
    SQL = SQL & "   AprovadorTipo, IdSetor, QuemIncluiu, "
    SQL = SQL & "   DataInclusao, Excluido"
    SQL = SQL & ") VALUES (" 
    SQL = SQL & "   '', "
    SQL = SQL & "   '" & idUsuario & "', "
    SQL = SQL & "   '" & nome & "', "
    SQL = SQL & "   '" & email & "', "
    SQL = SQL & "   '', "
    SQL = SQL & "   '', "
    SQL = SQL & "   '', "
    SQL = SQL & "   '" & usuarioLogado("id") & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   NULL"
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao Criar aprovador: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("criacao", idUsuario)
    Criar = "{""sucesso"": true, ""mensagem"": ""Aprovador criado com sucesso!""}"
End Function

Function Excluir()
    idUsuario = Request("idUsuario")
    
    If idUsuario = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    SQL = "UPDATE AprovadorFNC SET "
    SQL = SQL & "Excluido = 1, "
    SQL = SQL & "QuemExcluiu = " & usuarioLogado("id") & ", "
    SQL = SQL & "DataExcluiu = GETDATE() "
    SQL = SQL & "WHERE Id = " & idUsuario 
    SQL = SQL & " AND (Excluido = 0 OR Excluido IS NULL)"
    
    On Error Resume Next
    SET rsExcluir = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir aprovador: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsExcluir.Close()
    SET rsExcluir = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("exclusao", idUsuario)
    Excluir = "{""sucesso"": true, ""mensagem"": ""Aprovador excluído com sucesso!""}"
End Function


Sub RegistrarLog(acao, idUsuario)
    Dim SQL, rsLog
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de aprovador de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou o aprovador de id " & idUsuario
        Case "exclusao"
            acaoLog = "Exclusão de aprovador de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu o aprovador de id " & idUsuario
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Aprovadores de FNC', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET rsLog = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsLog) Then
        rsLog.Close()
        SET rsLog = NOTHING
    End If
    On Error Goto 0
End Sub
%>