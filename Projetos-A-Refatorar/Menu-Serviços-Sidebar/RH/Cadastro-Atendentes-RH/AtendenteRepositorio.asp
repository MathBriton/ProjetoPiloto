<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
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
    idDepartamento = Request("idDepartamento")
    codPlanta = Request("codPlanta")

    SQL = "SELECT a.Id, a.CodPlanta, a.IdUsuario, a.Nome, a.Email, a.DataCriacao, "
    SQL = SQL & "( "
    SQL = SQL & "    SELECT d.Id, d.Nome "
    SQL = SQL & "    FROM DepartamentosChamado d "
    SQL = SQL & "    INNER JOIN AtendenteDepartamentoChamado ad on ad.IdAtendente = a.Id AND ad.IdDepartamento = d.Id "
    SQL = SQL & "    WHERE d.FlagInativo = 0 "
    SQL = SQL & "    FOR JSON PATH "
    SQL = SQL & ") AS Departamentos "
    SQL = SQL & "FROM AtendentesChamado a " 
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND a.FlagInativo = 0 "

    If idDepartamento <> "" Then
        SQL = SQL & "AND EXISTS ( "
        SQL = SQL & "    SELECT TOP 1 a.Id "
        SQL = SQL & "    FROM AtendenteDepartamentoChamado ad "
        SQL = SQL & "    WHERE ad.IdAtendente = a.Id "
        SQL = SQL & "    AND ad.IdDepartamento = " & idDepartamento
        SQL = SQL & ") "
    End If

    If codPlanta <> "" Then
        SQL = SQL & "AND a.CodPlanta = " & codPlanta
    End If

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de atendentes: " & EscapeJSON(Err.Description) & """}"
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
            jsonArray = jsonArray & """idUsuario"": """ & EscapeJSON(rsListar("IdUsuario") & "") & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsListar("Nome") & "") & """, "
            jsonArray = jsonArray & """email"": """ & EscapeJSON(rsListar("Email") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """, "
            jsonArray = jsonArray & """departamentos"": " & Iif(IsNull(rsListar("Departamentos")), "[]", rsListar("Departamentos")) & ""
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

    SQL = "SELECT a.Id, a.CodPlanta, a.IdUsuario, a.Nome, a.Email, a.DataCriacao, "
    SQL = SQL & "( "
    SQL = SQL & "    SELECT d.Id, d.Nome "
    SQL = SQL & "    FROM DepartamentosChamado d "
    SQL = SQL & "    INNER JOIN AtendenteDepartamentoChamado ad on ad.IdAtendente = a.Id AND ad.IdDepartamento = d.Id "
    SQL = SQL & "    WHERE d.FlagInativo = 0"
    SQL = SQL & "    FOR JSON PATH "
    SQL = SQL & ") AS Departamentos "
    SQL = SQL & "FROM AtendentesChamado a " 
    SQL = SQL & "WHERE a.Id = " & id
    SQL = SQL & "AND a.FlagInativo = 0 "

    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados do atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Atendente nao encontrado""}"
        Exit Function
    End If
    
    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id") & "") & """, "
    jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta") & "") & """, "
    jsonObjeto = jsonObjeto & """idUsuario"": """ & EscapeJSON(rsObter("IdUsuario") & "") & """, "
    jsonObjeto = jsonObjeto & """nome"": """ & EscapeJSON(rsObter("Nome") & "") & """, "
    jsonObjeto = jsonObjeto & """email"": """ & EscapeJSON(rsObter("Email") & "") & """, "
    jsonObjeto = jsonObjeto & """dataCriacao"": """ & EscapeJSON(rsObter("DataCriacao") & "") & """, "
    jsonObjeto = jsonObjeto & """departamentos"": " & Iif(IsNull(rsObter("Departamentos")), "[]", rsObter("Departamentos")) & ""
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Atendente encontrado"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    idUsuario = Request("idUsuario")
    codPlanta = Request("codPlanta")
    nome = Request("nome")
    email = Request("email")

    departamentosJSON = Request("idsDepartamento")

    ' Insere registro de atendente
    SQL = SQL & "INSERT INTO AtendentesChamado (CodPlanta, IdUsuario, Nome, Email, DataCriacao, FlagInativo) "
    SQL = SQL & "OUTPUT INSERTED.Id "
    SQL = SQL & "VALUES ( "
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "'" & idUsuario & "', "
    SQL = SQL & "'" & nome & "', "
    SQL = SQL & "'" & email & "', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "0 "
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    idAtendente = rsCriar("Id")

    rsCriar.Close()
    SET rsCriar = NOTHING

    ' Monta insert para tabela associativa
    If departamentosJSON <> "" Then
        departamentosJSON = Replace(departamentosJSON, "[", "")
        departamentosJSON = Replace(departamentosJSON, "]", "")
        departamentosJSON = Replace(departamentosJSON, """", "")
        listaDepartamentos = Split(departamentosJSON, ",")
    End If

    clausulasInsert = ""
    For j = LBound(listaDepartamentos) To UBound(listaDepartamentos)
        idDepartamento = Trim(listaDepartamentos(j))

        If clausulasInsert <> "" Then
            clausulasInsert = clausulasInsert & ", "
        End If
        
        clausulasInsert = clausulasInsert & "( "
        clausulasInsert = clausulasInsert & "'" & idAtendente & "', "
        clausulasInsert = clausulasInsert & "'" & idDepartamento & "' "
        clausulasInsert = clausulasInsert & ")"
    Next

    ' Insere Registros
    SQL = "INSERT INTO AtendenteDepartamentoChamado (IdAtendente, IdDepartamento) "
    SQL = SQL & "VALUES " & clausulasInsert

    On Error Resume Next
    SET rsCriarDepartamento = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriarDepartamento) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Atendente criado com sucesso!""}"
End Function

Function Editar()
    id = Request("id")
    codPlanta = Request("codPlanta")
    departamentosJSON = Request("idsDepartamento")
    
    If id = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""ID de atendente não informado""}"
        Exit Function
    End If
   
    ' Monta insert para tabela associativa
    If departamentosJSON <> "" Then
        departamentosJSON = Replace(departamentosJSON, "[", "")
        departamentosJSON = Replace(departamentosJSON, "]", "")
        departamentosJSON = Replace(departamentosJSON, """", "")
        listaDepartamentos = Split(departamentosJSON, ",")
    End If

    clausulasInsert = ""
    For j = LBound(listaDepartamentos) To UBound(listaDepartamentos)
        idDepartamento = Trim(listaDepartamentos(j))

        If clausulasInsert <> "" Then
            clausulasInsert = clausulasInsert & ", "
        End If
        
        clausulasInsert = clausulasInsert & "( "
        clausulasInsert = clausulasInsert & "'" & id & "', "
        clausulasInsert = clausulasInsert & "'" & idDepartamento & "' "
        clausulasInsert = clausulasInsert & ")"
    Next

    ' Insere Registros
    SQL = "UPDATE AtendentesChamado SET"
    SQL = SQL & "   CodPlanta = '" & codPlanta & "' "
    SQL = SQL & "WHERE Id = '" & id & "'; "

    SQL = SQL & "DELETE FROM AtendenteDepartamentoChamado WHERE IdAtendente = '" & id & "'; "

    SQL = SQL & "INSERT INTO AtendenteDepartamentoChamado (IdAtendente, IdDepartamento) "
    SQL = SQL & "VALUES " & clausulasInsert

    On Error Resume Next
    SET rsEditar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("edicao", id)
    Editar = "{""sucesso"": true, ""mensagem"": ""Atendente editado com sucesso!""}"
End Function

Function Excluir()
    id = Request("id")
    
    If id = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "DELETE FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & id
    SQL = SQL & "UPDATE AtendentesChamado SET FlagInativo = 1 WHERE Id = " & id

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
    
    On Error Resume Next
    SET rsExcluir = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call RegistrarLog("exclusao", id)
    Excluir = "{""sucesso"": true, ""mensagem"": ""Atendente excluído com sucesso!""}"
End Function


Sub RegistrarLog(acao, id)
    Dim SQL, rsLog
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Atendente de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou o atendente de id " & id
        Case "edicao"
            acaoLog = "Edição de Atendente de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou o atendente de id " & id
        Case "exclusao"
            acaoLog = "Exclusão de Atendente de Chamado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu o atendente de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Atendentes de Chamado', "
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