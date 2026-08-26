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
Dim idDefeito, idSetor, tipo, descricao
Dim resultado, mensagem

idUsuarioLogado = Session("CODIGO")
loginUsuarioLogado = Session("USR")

acao = Request("acao")
idDefeito = Request("idDefeito")
idSetor = Request("idSetor")
tipo = Request("tipo")
descricao = Request("descricao")

resultado = false
mensagem = ""

if acao = "" Then
    mensagem = "Ação não informada"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

if idDefeito = "" Then
    mensagem = "Id de defeito não informado"
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
    
    SQL = "INSERT INTO Defeito (Id, Tipo, Descricao, IdSetor, DataInclusao) VALUES ("
    SQL = SQL & "'" & idNovoRegistro & "', "
    SQL = SQL & "'" & tipo & "', "
    SQL = SQL & "'" & descricao & "', "
    SQL = SQL & "'" & idSetor & "', "
    SQL = SQL & "GETDATE())"

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao criar defeito: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Defeito criado com sucesso!"

    Call RegistrarLog()
End Sub


Sub Editar()
    Dim SQL, BEN_CORPORATIVO
    
    SQL = "UPDATE Defeito SET "
    SQL = SQL & "Tipo = " & "'" & tipo & "'" & ", "
    SQL = SQL & "Descricao = " & "'" & descricao & "'" & ", "
    SQL = SQL & "IdSetor = " & "'" & idSetor & "'" & " "
    SQL = SQL & "WHERE Id = " & idDefeito 
    
    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao editar defeito: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Defeito atualizado com sucesso!"

    Call RegistrarLog()
End Sub


Sub Excluir()
    Dim SQL, BEN_CORPORATIVO

    SQL = "UPDATE Defeito SET FlagInativo = 1 WHERE Id = " & idDefeito 
    
    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(BEN_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir defeito: " & Err.Description
        Exit Sub
    end if
    
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Defeito excluído com sucesso!"
    
    Call RegistrarLog()
End Sub


'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------


'Função para obter proximo Id (tabela não tem autoincrement)'
Function ObterProximoId()
    Dim idNovoRegistro, SQL, BEN_CORPORATIVO
    idNovoRegistro = 0

    SQL = "SELECT ISNULL(MAX(Id),0) + 1 as ProximoId FROM Defeito"

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
            acaoLog = "Criação de defeito de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " criou o defeito de FNC de id " & idDefeito
        Case "edicao"
            acaoLog = "Edição de defeito de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " editou o defeito de FNC de id " & idDefeito
        Case "exclusao"
            acaoLog = "Exclusão de defeito de FNC"
            descricaoLog = "Usuário " & loginUsuarioLogado & " excluiu o defeito de FNC de id " & idDefeito
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuarioLogado & ", "
    SQL = SQL & "'Defeitos de FNC', "
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
