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
Dim idUsuarioLogado, loginUsuarioLogado
Dim idSetor, descricao
Dim resultado, mensagem

idUsuarioLogado = Session("CODIGO")
loginUsuarioLogado = Session("USR")

acao = Request("acao")
idSetor = Request("idSetor")
descricao = Request("descricao")

resultado = false
mensagem = ""

if acao = "" Then
    mensagem = "Ação não informada"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

if idSetor = "" Then
    mensagem = "Id de setor não informado"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

Select case LCase(acao)
    case "criacao"
        Call Criar()
    case "edicao"
        Call Editar()
    case "exclusao"
        Call Excluir()
    case Else
        mensagem = "Ação inválida: " & acao
        resultado = false
End Select


if resultado Then
    Response.Write("{""sucesso"": true, ""mensagem"": """ & mensagem & """}")
else
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
end if


'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------

Sub Criar()
    Dim SQL, BEN_CORPORATIVO

    idNovoRegistro = ObterProximoId()
    
    SQL = "INSERT INTO Setor (Id, Descricao, QuemIncluiu, DataInclusao) VALUES ("
    SQL = SQL & "'" & idNovoRegistro & "', "
    SQL = SQL & "'" & descricao & "', "
    SQL = SQL & "'" & idUsuarioLogado & "', "
    SQL = SQL & "GETDATE())"

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao criar setor: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Setor criado com sucesso!"

    Call RegistrarLog()
End Sub


Sub Editar()
    Dim SQL, BEN_CORPORATIVO
    
    SQL = "UPDATE Setor SET "
    SQL = SQL & "Descricao = " & "'" & descricao & "'" & " "
    SQL = SQL & "WHERE Id = " & idSetor 
    
    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao editar setor: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Setor atualizado com sucesso!"

    Call RegistrarLog()
End Sub


Sub Excluir()
    Dim SQL, BEN_CORPORATIVO

    SQL = "UPDATE Setor SET FlagInativo = 1 WHERE Id = " & idSetor 
    
    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir setor: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Setor excluído com sucesso!"
    
    Call RegistrarLog()
End Sub


'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------


'Função para obter proximo Id (tabela não tem autoincrement)'
Function ObterProximoId()
    Dim idNovoRegistro, SQL, BEN_CORPORATIVO
    idNovoRegistro = 0

    SQL = "SELECT ISNULL(MAX(Id),0) + 1 as ProximoId FROM Setor"

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number = 0 AND isObject(BEN_CORPORATIVO) then
        if Not BEN_CORPORATIVO.EOF Then
            idNovoRegistro = BEN_CORPORATIVO("ProximoId")
        end if
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    end if
    
    On Error Goto 0

    ObterProximoId = idNovoRegistro
End Function

Sub RegistrarLog()
    Dim SQL, LOG_CONN
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de setor de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " criou o setor de FNC de id " & idSetor
        Case "edicao"
            acaoLog = "Edição de setor de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " editou o setor de FNC de id " & idSetor
        Case "exclusao"
            acaoLog = "Exclusão de setor de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " excluiu o setor de FNC de id " & idSetor
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuarioLogado & ", "
    SQL = SQL & "'Setores de FNC', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET LOG_CONN = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number = 0 AND isObject(LOG_CONN) Then
        LOG_CONN.Close()
        SET LOG_CONN = NOTHING
    end if
    On Error Goto 0
End Sub

%>
