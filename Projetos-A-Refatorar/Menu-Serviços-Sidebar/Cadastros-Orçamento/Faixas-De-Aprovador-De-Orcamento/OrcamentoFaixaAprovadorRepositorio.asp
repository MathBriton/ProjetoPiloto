<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

' Variáveis globais
Dim idUsuarioLogado, loginUsuarioLogado
Dim idFaixa, planta, tipoProduto, valorMinKg, valorMaxKg, valorMinM2, valorMaxM2, quantidade
Dim resultado, mensagem

idUsuarioLogado = Session("CODIGO")
loginUsuarioLogado = Session("USR")

acao = Request("acao")
idFaixa = Request("idFaixa")
planta = Request("planta")
tipoProduto = Request("tipoProduto")
valorMinKg = Request("valorMinKg")
valorMaxKg = Request("valorMaxKg")
valorMinM2 = Request("valorMinM2")
valorMaxM2 = Request("valorMaxM2")
quantidade = Request("quantidade")

resultado = false
mensagem = ""

if acao = "" Then
    mensagem = "Ação não informada"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

if idFaixa = "" Then
    mensagem = "Id de faixa não informado"
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
    Dim SQL, PIR_CORPORATIVO
    
    SQL = "INSERT INTO FAIXA_APROVADORES (PLANTA, TIPO_PRODUTO, VALOR_MIN_KG, VALOR_MAX_KG, VALOR_MIN_M2, VALOR_MAX_M2, QUANTIDADE) VALUES ("
    SQL = SQL & "'" & planta & "', "
    SQL = SQL & "'" & tipoProduto & "', "
    SQL = SQL & valorMinKg & ", "
    SQL = SQL & valorMaxKg & ", "
    SQL = SQL & valorMinM2 & ", "
    SQL = SQL & valorMaxM2 & ", "
    SQL = SQL & quantidade & ")"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao criar faixa: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Faixa atualizada com sucesso!"
End Sub


Sub Editar()
    Dim SQL, PIR_CORPORATIVO
    
    SQL = "UPDATE FAIXA_APROVADORES SET "
    SQL = SQL & "PLANTA = '" & planta & "', "
    SQL = SQL & "TIPO_PRODUTO = '" & tipoProduto & "', "
    SQL = SQL & "VALOR_MIN_KG = " & valorMinKg & ", "
    SQL = SQL & "VALOR_MAX_KG = " & valorMaxKg & ", "
    SQL = SQL & "VALOR_MIN_M2 = " & valorMinM2 & ", "
    SQL = SQL & "VALOR_MAX_M2 = " & valorMaxM2 & ", "
    SQL = SQL & "QUANTIDADE = " & valorMaxM2 & " "
    SQL = SQL & "WHERE ID = " & idFaixa 
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao editar faixa: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Faixa atualizada com sucesso!"
End Sub


Sub Excluir()
    Dim SQL, PIR_CORPORATIVO

    SQL = "DELETE FAIXA_APROVADORES WHERE ID = " & idFaixa 
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir faixa: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Faixa excluída com sucesso!"
    
    Call RegistrarLog()
End Sub


'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------

Sub RegistrarLog()
    Dim SQL, LOG_CONN, PIR_HELPDESK
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de faixa de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " criou a faixa de aprovador de id" & idFaixa
        Case "edicao"
            acaoLog = "Edição de faixa de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " editou a faixa de aprovador de id" & idFaixa
        Case "exclusao"
            acaoLog = "Exclusão de faixa de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " excluiu a faixa de aprovador de id" & idFaixa
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuarioLogado & ", "
    SQL = SQL & "'Faixas de Aprovadores de Orçamento', "
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
