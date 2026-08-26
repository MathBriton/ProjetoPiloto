<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

' Variáveis globais
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Select Case acao
    Case "listar"
        resultado = Listar()
    Case "criar"
        resultado = Criar()
    Case "editar"
        resultado = Editar()
    Case "excluir"
        resultado = Excluir()
    Case "obterPorId"
        resultado = ObterPorId()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    codPlanta = Request("codPlanta")

    SQL = "SELECT Id, CodPlanta, Nome, DataCriacao, FlagInativo "
    SQL = SQL & "FROM DepartamentosChamado  "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND FlagInativo = 0"

    If codPlanta <> "" Then
        SQL = SQL & "AND CodPlanta = '" & codPlanta & "'"
    End If

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de departamentos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id") & "") & """, "
            jsonArray = jsonArray & """codPlanta"": """ & EscapeJSON(rsListar("CodPlanta") & "") & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsListar("Nome") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """, "
            jsonArray = jsonArray & """flagInativo"": """ & EscapeJSON(rsListar("FlagInativo") & "") & """"
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

Function ObterPorId()    
    id = Request("id")

    If id = "" Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If
    
    SQL = "SELECT Id, CodPlanta, Nome, DataCriacao, FlagInativo "
    SQL = SQL & "FROM DepartamentosChamado  "
    SQL = SQL & "WHERE Id = " & id
    SQL = SQL & "AND FlagInativo = 0"
    
    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados do departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Departamento nao encontrado""}"
        Exit Function
    End If
    
    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id") & "") & """, "
    jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta") & "") & """, "
    jsonObjeto = jsonObjeto & """nome"": """ & EscapeJSON(rsObter("Nome") & "") & """, "
    jsonObjeto = jsonObjeto & """dataCriacao"": """ & EscapeJSON(rsObter("DataCriacao") & "") & """, "
    jsonObjeto = jsonObjeto & """flagInativo"": """ & EscapeJSON(rsObter("FlagInativo") & "") & """"
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Departamento encontrado"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    codPlanta = Request("codPlanta")
    nome = Request("nome")

    SQL = "INSERT INTO DepartamentosChamado (CodPlanta, Nome, DataCriacao, FlagInativo) "
    SQL = SQL & "OUTPUT INSERTED.Id "
    SQL = SQL & "VALUES ( "
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "'" & nome & "', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "0 "
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    id = rsCriar("Id")
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Departamento criado com sucesso!""}"
End Function

Function Editar()
    id = Request("id")
    nome = Request("nome")
    
    If id = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""ID de departamento não informado""}"
        Exit Function
    End If
    
    ' Dados formulário
    SQL = "UPDATE DepartamentosChamado SET "
    SQL = SQL & "Nome = '" & nome & "' "
    SQL = SQL & "WHERE Id = " & id

    On Error Resume Next
    SET rsEditar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("edicao", id)
    Editar = "{""sucesso"": true, ""mensagem"": ""Departamento editado com sucesso!""}"
End Function

Function Excluir()
    id = Request("id")
    
    If id = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    SQL = "UPDATE DepartamentosChamado SET FlagInativo = 1 WHERE Id = " & id
    
    On Error Resume Next
    SET rsExcluir = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    On Error Goto 0
    
    Call RegistrarLog("exclusao", id)
    Excluir = "{""sucesso"": true, ""mensagem"": ""Departamento excluído com sucesso!""}"
End Function


Sub RegistrarLog(acao, id)
    Dim SQL, rsLog
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Departamento de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou o departamento de id " & id
        Case "edicao"
            acaoLog = "Edição de Departamento de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou o departamento de id " & id
        Case "exclusao"
            acaoLog = "Exclusão de Departamento de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu o departamento de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Departamentos de Chamado', "
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