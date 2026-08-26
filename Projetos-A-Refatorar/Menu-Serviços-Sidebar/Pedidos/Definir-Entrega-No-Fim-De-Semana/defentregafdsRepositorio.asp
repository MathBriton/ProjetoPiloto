<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../cnn2/conexoes.asp"-->
<!--#Include File="../../includes/variaveis.asp"-->
<%
'===================================================================================================
' FONTE    : defentregafdsRepositorio.asp
' DESCRICAO: Endpoint JSON (API) do modulo "Define Entrega no Final de Semana".
'            Substitui repo_defentregafds.asp + validaClienteAjax.asp (escopo deste modulo).
'            Consumido via AJAX por defentregafdsFormulario.asp.
'===================================================================================================
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Buffer      = True
Response.Expires     = 0
Session.LCID         = 1046
Response.ContentType = "application/json"

'----------------------------------------------------------------------------------------------
' Contexto de sessao (usuario logado)
'----------------------------------------------------------------------------------------------
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id",    Session("CODIGO")
usuarioLogado.Add "login", Session("CODUSER")
usuarioLogado.Add "nome",  Session("NOME")

'----------------------------------------------------------------------------------------------
' EscapeJSONLocal — EscapeJSON de funcoes.asp possui bug conhecido, nao escapa corretamente.
'----------------------------------------------------------------------------------------------
Function EscapeJSONLocal(valor)
    Dim resultado
    resultado = valor & ""
    resultado = Replace(resultado, "\", "\\")
    resultado = Replace(resultado, """", "\""")
    resultado = Replace(resultado, vbCrLf, "\n")
    resultado = Replace(resultado, vbCr, "\n")
    resultado = Replace(resultado, vbLf, "\n")
    resultado = Replace(resultado, vbTab, "\t")
    EscapeJSONLocal = resultado
End Function

'----------------------------------------------------------------------------------------------
' Dispatcher
'----------------------------------------------------------------------------------------------
acao = Trim(Request("acao"))

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Select Case acao
    Case "listar"        : resultado = Listar()
    Case "obter"          : resultado = Obter()
    Case "validarCliente" : resultado = ValidarCliente()
    Case "criar"          : resultado = Criar()
    Case "alterar"        : resultado = Alterar()
    Case "excluir"        : resultado = Excluir()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSONLocal(acao) & """}"
End Select

Response.Write(resultado)
Response.End()

'====================================================================================================
' Listar — grid de clientes cadastrados (nao excluidos)
'====================================================================================================
Function Listar()

    SQL = "" &_
        "SELECT " &_
            "DEF.COD_CLIENTE, " &_
            "CLI.A1_NOME AS NOME_CLIENTE, " &_
            "DEF.ATIVO " &_
        "FROM DEF_ENTREGA_FDS DEF " &_
        "OUTER APPLY ( " &_
            "SELECT TOP 1 A1_NOME " &_
            "FROM " & cnx(skCnx("PIR-DADOSADV")).getDataBase() & ".dbo.SA1010 SA1 " &_
            "WHERE DEF.COD_CLIENTE = SA1.A1_COD COLLATE SQL_Latin1_General_CP437_BIN " &_
        ") CLI " &_
        "WHERE (DEF.QUEM_EXCLUIU = 0 OR DEF.QUEM_EXCLUIU = '') " &_
        "ORDER BY CASE DEF.ATIVO WHEN 1 THEN 'A' ELSE 'B' END, CLI.A1_NOME"

    On Error Resume Next
    SET rs = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    Do While Not rs.EOF
        If Not primeiroItem Then jsonArray = jsonArray & ","

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & EscapeJSONLocal(rs("COD_CLIENTE") & "")  & """, "
        jsonArray = jsonArray & """nome"": """   & EscapeJSONLocal(rs("NOME_CLIENTE") & "") & """, "
        jsonArray = jsonArray & """ativo"": "    & iif(rs("ATIVO"), "true", "false")
        jsonArray = jsonArray & "}"

        primeiroItem = False
        rs.MoveNext
    Loop

    jsonArray = jsonArray & "]"

    rs.Close()
    SET rs = NOTHING

    Listar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & jsonArray & "}"

End Function

'====================================================================================================
' Obter — dados de um unico registro, usado para pre-popular o modal de editar/visualizar
'====================================================================================================
Function Obter()

    codigo = Trim(Request("codigo"))

    If codigo = "" Then
        Obter = "{""sucesso"": false, ""mensagem"": ""Código não informado""}"
        Exit Function
    End If

    SQL = "" &_
        "SELECT " &_
            "DEF.COD_CLIENTE, " &_
            "CLI.A1_NOME AS NOME_CLIENTE, " &_
            "DEF.ATIVO " &_
        "FROM DEF_ENTREGA_FDS DEF " &_
        "OUTER APPLY ( " &_
            "SELECT TOP 1 A1_NOME " &_
            "FROM " & cnx(skCnx("PIR-DADOSADV")).getDataBase() & ".dbo.SA1010 SA1 " &_
            "WHERE DEF.COD_CLIENTE = SA1.A1_COD COLLATE SQL_Latin1_General_CP437_BIN " &_
        ") CLI " &_
        "WHERE DEF.COD_CLIENTE = '" & Replace(codigo, "'", "''") & "'"

    On Error Resume Next
    SET rs = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        Obter = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rs.EOF Then
        rs.Close()
        SET rs = NOTHING
        Obter = "{""sucesso"": false, ""mensagem"": ""Registro não encontrado""}"
        Exit Function
    End If

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """codigo"": """ & EscapeJSONLocal(rs("COD_CLIENTE") & "")  & """, "
    jsonObjeto = jsonObjeto & """nome"": """   & EscapeJSONLocal(rs("NOME_CLIENTE") & "") & """, "
    jsonObjeto = jsonObjeto & """ativo"": "    & iif(rs("ATIVO"), "true", "false")
    jsonObjeto = jsonObjeto & "}"

    rs.Close()
    SET rs = NOTHING

    Obter = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & jsonObjeto & "}"

End Function

'====================================================================================================
' ValidarCliente — busca o nome do cliente em SA1010 a partir do codigo digitado (modal de criar)
'====================================================================================================
Function ValidarCliente()

    codigo = Trim(Request("codigo"))

    If codigo = "" Then
        ValidarCliente = "{""sucesso"": false, ""mensagem"": ""Código não informado""}"
        Exit Function
    End If

    SQL = "" &_
        "SELECT TOP 1 A1_NOME " &_
        "FROM " & cnx(skCnx("PIR-DADOSADV")).getDataBase() & ".dbo.SA1010 " &_
        "WHERE A1_COD = '" & Replace(codigo, "'", "''") & "' COLLATE SQL_Latin1_General_CP437_BIN " &_
            "AND D_E_L_E_T_ <> '*'"

    On Error Resume Next
    SET rs = Conecta("PIR-DADOSADV", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        ValidarCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rs.EOF Then
        rs.Close()
        SET rs = NOTHING
        ValidarCliente = "{""sucesso"": false, ""mensagem"": ""Cliente não encontrado.""}"
        Exit Function
    End If

    nomeCliente = rs("A1_NOME") & ""

    rs.Close()
    SET rs = NOTHING

    ValidarCliente = "{""sucesso"": true, ""mensagem"": """", ""dados"": {""codigo"": """ & EscapeJSONLocal(codigo) & """, ""nome"": """ & EscapeJSONLocal(nomeCliente) & """}}"

End Function

'====================================================================================================
' Criar — inclui novo registro de cliente com entrega liberada/bloqueada no fim de semana
'====================================================================================================
Function Criar()

    codigo = Trim(Request("codcli"))
    ativo  = IIf(Trim(Request("ativo")) = "1", "1", "0")

    If codigo = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Informe o código do cliente""}"
        Exit Function
    End If

    codigoSQL = Replace(codigo, "'", "''")

    SQL = "" &_
        "SELECT COUNT(*) AS EXISTE " &_
        "FROM DEF_ENTREGA_FDS " &_
        "WHERE COD_CLIENTE = '" & codigoSQL & "' " &_
            "AND (QUEM_EXCLUIU = 0 OR QUEM_EXCLUIU = '') "

    On Error Resume Next
    SET rs = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jaExiste = (Not rs.EOF) And (rs("EXISTE") > 0)
    rs.Close()
    SET rs = NOTHING

    If jaExiste Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Cliente já definido!""}"
        Exit Function
    End If

    SQL = "" &_
        "BEGIN TRANSACTION TRANS; " &_
            "INSERT INTO DEF_ENTREGA_FDS ( " &_
                "COD_CLIENTE, ATIVO, QUEM_INCLUIU, DATA_INCLUSAO, " &_
                "QUEM_ALTEROU, DATA_ALTERACAO, QUEM_EXCLUIU, DATA_EXCLUSAO " &_
            ") VALUES ( " &_
                "'" & codigoSQL & "', " &_
                "'" & ativo & "', " &_
                "'" & usuarioLogado("login") & "', " &_
                "GETDATE(), " &_
                "'', NULL, '', NULL " &_
            "); " &_
        "IF @@ERROR <> 0 " &_
            "ROLLBACK TRANSACTION TRANS " &_
        "ELSE " &_
            "COMMIT TRANSACTION TRANS "

    On Error Resume Next
    SET rsIns = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsIns) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao incluir: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    SET rsIns = NOTHING

    Criar = "{""sucesso"": true, ""mensagem"": ""Cliente incluído com sucesso!""}"

End Function

'====================================================================================================
' Alterar — atualiza o flag ATIVO de um registro existente
'====================================================================================================
Function Alterar()

    codigo = Trim(Request("codcli"))
    ativo  = IIf(Trim(Request("ativo")) = "1", "1", "0")

    If codigo = "" Then
        Alterar = "{""sucesso"": false, ""mensagem"": ""Informe o código do cliente""}"
        Exit Function
    End If

    codigoSQL = Replace(codigo, "'", "''")

    SQL = "" &_
        "BEGIN TRANSACTION TRANS; " &_
            "UPDATE DEF_ENTREGA_FDS SET " &_
                "ATIVO = '" & ativo & "', " &_
                "DATA_ALTERACAO = GETDATE(), " &_
                "QUEM_ALTEROU = '" & usuarioLogado("login") & "' " &_
            "WHERE COD_CLIENTE = '" & codigoSQL & "'; " &_
        "IF @@ERROR <> 0 " &_
            "ROLLBACK TRANSACTION TRANS " &_
        "ELSE " &_
            "COMMIT TRANSACTION TRANS "

    On Error Resume Next
    SET rs = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        Alterar = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    SET rs = NOTHING

    Alterar = "{""sucesso"": true, ""mensagem"": ""Cliente alterado com sucesso!""}"

End Function

'====================================================================================================
' Excluir — marca o registro como excluido (soft delete), acionado pelo modal Bootstrap de confirmacao
'====================================================================================================
Function Excluir()

    codigo = Trim(Request("codcli"))

    If codigo = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Informe o código do cliente""}"
        Exit Function
    End If

    codigoSQL = Replace(codigo, "'", "''")

    SQL = "" &_
        "BEGIN TRANSACTION TRANS; " &_
            "UPDATE DEF_ENTREGA_FDS SET " &_
                "DATA_EXCLUSAO = GETDATE(), " &_
                "QUEM_EXCLUIU = '" & usuarioLogado("login") & "' " &_
            "WHERE COD_CLIENTE = '" & codigoSQL & "'; " &_
        "IF @@ERROR <> 0 " &_
            "ROLLBACK TRANSACTION TRANS " &_
        "ELSE " &_
            "COMMIT TRANSACTION TRANS "

    On Error Resume Next
    SET rs = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rs) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao apagar: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    SET rs = NOTHING

    Excluir = "{""sucesso"": true, ""mensagem"": ""Registro apagado com sucesso!""}"

End Function
%>
