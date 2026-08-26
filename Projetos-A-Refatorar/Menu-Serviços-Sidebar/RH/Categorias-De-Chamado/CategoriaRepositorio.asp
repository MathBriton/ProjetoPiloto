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
    idTipo = Request("idTipo")
    codPlanta = Request("codPlanta")

    SQL = "SELECT c.Id, c.CodPlanta, c.Nome, t.Nome as Tipo, c.IdTipo, c.DataCriacao "
    SQL = SQL & "FROM CategoriasChamado c "
    SQL = SQL & "LEFT JOIN TiposChamado t on t.Id = c.IdTipo "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.FlagInativo = 0"

    If idTipo <> "" Then 
        SQL = SQL & " AND c.IdTipo = " & idTipo
    End If

    If codPlanta <> "" Then 
        SQL = SQL & " AND c.CodPlanta = '" & codPlanta & "'"
    End If

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de categorias: " & EscapeJSON(Err.Description) & """}"
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
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsListar("Nome") & "") & """, "
            jsonArray = jsonArray & """codPlanta"": """ & EscapeJSON(rsListar("CodPlanta") & "") & """, "
            jsonArray = jsonArray & """idtipo"": """ & EscapeJSON(rsListar("Idtipo") & "") & """, "
            jsonArray = jsonArray & """tipo"": """ & EscapeJSON(rsListar("Tipo") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """"
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
    
    SQL = "SELECT c.Id, c.CodPlanta, c.Nome, t.Nome as Tipo, c.IdTipo, c.DataCriacao "
    SQL = SQL & "FROM CategoriasChamado c "
    SQL = SQL & "LEFT JOIN TiposChamado t on t.Id = c.IdTipo "
    SQL = SQL & "WHERE c.Id = " & id
    SQL = SQL & "AND c.FlagInativo = 0"
    
    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados da categoria: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Categoria nao encontrada""}"
        Exit Function
    End If
    
    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id") & "") & """, "
    jsonObjeto = jsonObjeto & """idTipo"": """ & EscapeJSON(rsObter("IdTipo") & "") & """, "
    jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta") & "") & """, "
    jsonObjeto = jsonObjeto & """tipo"": """ & EscapeJSON(rsObter("Tipo") & "") & """, "
    jsonObjeto = jsonObjeto & """nome"": """ & EscapeJSON(rsObter("Nome") & "") & """, "
    jsonObjeto = jsonObjeto & """dataCriacao"": """ & EscapeJSON(rsObter("DataCriacao") & "") & """"
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Categoria encontrada"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    nome = Request("nome")
    idTipo = Request("idTipo")
    codPlanta = Request("codPlanta")

    SQL = "INSERT INTO CategoriasChamado (Nome, CodPlanta, IdTipo, DataCriacao, FlagInativo) "
    SQL = SQL & "OUTPUT INSERTED.Id "
    SQL = SQL & "VALUES ( "
    SQL = SQL & "'" & nome & "', "
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "'" & idTipo & "', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "0 "
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar categoria: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    id = rsCriar("Id")
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Categoria criada com sucesso!""}"
End Function

Function Editar()
    id = Request("id")
    codPlanta = Request("codPlanta")
    idTipo = Request("idTipo")
    nome = Request("nome")
    
    If id = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""ID de categoria não informada""}"
        Exit Function
    End If
    
    ' Dados formulário
    SQL = "UPDATE CategoriasChamado SET "
    SQL = SQL & "Nome = '" & nome & "', "
    SQL = SQL & "IdTipo = '" & idTipo & "', "
    SQL = SQL & "CodPlanta = '" & codPlanta & "' "
    SQL = SQL & "WHERE Id = " & id

    On Error Resume Next
    SET rsEditar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar categoria: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("edicao", id)
    Editar = "{""sucesso"": true, ""mensagem"": ""Categoria editado com sucesso!""}"
End Function

Function Excluir()
    id = Request("id")
    
    If id = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    SQL = "UPDATE CategoriasChamado SET FlagInativo = 1 WHERE Id = " & id
    
    On Error Resume Next
    SET rsExcluir = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir categoria: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    On Error Goto 0
    
    Call RegistrarLog("exclusao", id)
    Excluir = "{""sucesso"": true, ""mensagem"": ""Categoria excluída com sucesso!""}"
End Function


Sub RegistrarLog(acao, id)
    Dim SQL, rsLog
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Categoria de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a categoria de id " & id
        Case "edicao"
            acaoLog = "Edição de Categoria de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou a categoria de id " & id
        Case "exclusao"
            acaoLog = "Exclusão de Categoria de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu a categoria de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Categorias de Chamado', "
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