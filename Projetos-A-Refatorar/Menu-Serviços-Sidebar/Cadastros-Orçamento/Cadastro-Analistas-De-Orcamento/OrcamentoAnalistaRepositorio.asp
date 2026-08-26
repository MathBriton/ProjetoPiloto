<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

' Variáveis globais
Dim idUsuarioLogado, loginUsuarioLogado
Dim parametroAcao, idAnalista, plantas
Dim resultado, mensagem
Dim plantasString, listaPlantas

idUsuarioLogado = Session("CODIGO")
loginUsuarioLogado = Session("USR")

acao = Request("acao")
idAnalista = Request("idAnalista")
plantas = Request("plantas")

'Converte json pra array'
if (plantas <> "") THEN
    plantasString = Replace(plantas, "[", "")
    plantasString = Replace(plantasString, "]", "")
    plantasString = Replace(plantasString, """", "")
    listaPlantas = Split(plantasString, ",")
end if

resultado = false
mensagem = ""

if acao = "" Then
    mensagem = "Ação não informada"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

if idAnalista = "" Then
    mensagem = "Id de analista não informado"
    Response.Write("{""sucesso"": false, ""mensagem"": """ & mensagem & """}")
    Response.End()
end if

Select case LCase(acao)
    case "criacao"
        Call Processar()
    case "edicao"
        Call Processar()
    case "exclusao"
        Call ExcluirTodos()
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
'          MÉTODOS PARA LIDAR COM AÇÃO
'----------------------------------------------

Sub Processar() 
    if idAnalista = "" Then
        resultado = false
        mensagem = "ID do analista não informado"
        Exit Sub
    end if

    Dim listaInclusao(), listaEdicao()
    Dim countInclusao, countEdicao
    Dim planta, i
    countInclusao = -1
    countEdicao = -1

    ' Monta listas de inclusão e edição
    For i = LBound(listaPlantas) To UBound(listaPlantas)
        planta = Trim(listaPlantas(i))
        
        if planta <> "" Then
            Dim jaExiste
            jaExiste = ExisteRegistro(planta)

            if jaExiste Then
                countEdicao = countEdicao + 1
                ReDim Preserve listaEdicao(countEdicao)
                listaEdicao(countEdicao) = planta
            else
                countInclusao = countInclusao + 1
                ReDim Preserve listaInclusao(countInclusao)
                listaInclusao(countInclusao) = planta
            end if
        end if
    Next

    Call Excluir()
    if Not resultado Then
        Exit Sub
    end if

    if countEdicao > -1 Then
        Call Editar(listaEdicao)
        if Not resultado Then
            Exit Sub
        end if
    end if

    if countInclusao > -1 Then
        Call Criar(listaInclusao)
        if Not resultado Then
            Exit Sub
        end if
    end if
    
    Call RegistrarLog()
End Sub


Sub ExcluirTodos()
    Dim SQL, PIR_CORPORATIVO
    
    SQL = "UPDATE ANALISTAS_ORCAMENTO SET "
    SQL = SQL & "EXCLUIDO = 1, "
    SQL = SQL & "QUEM_EXCLUIU = " & idUsuarioLogado & ", "
    SQL = SQL & "DATA_EXCLUSAO = GETDATE() "
    SQL = SQL & "WHERE CODIGO_USUARIO = " & idAnalista 
    SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir analista: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Analista excluído com sucesso!"
    
    Call RegistrarLog()
End Sub


'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------

Sub Criar(listaInclusao)
    Dim SQL, PIR_CORPORATIVO, PIR_HELPDESK, j
    Dim nomeAnalista, email
    Dim valuesSQL
    
    ' Busca dados do analista
    SQL = "SELECT CRM_NOME as NOME, CRM_EMAILPRIMARIO as EMAIL FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAnalista
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAnalista = Replace(PIR_HELPDESK("NOME"), "'", "''")
        email = Replace(PIR_HELPDESK("EMAIL"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    else
        resultado = false
        mensagem = "Erro ao buscar dados do analista: " & Err.Description
        Exit Sub
    end if
    On Error Goto 0
    
    ' Montando clausulas para inserte em massa
    SQL = "INSERT INTO ANALISTAS_ORCAMENTO (PLANTA, CODIGO_USUARIO, NOME_USUARIO, EMAIL, QUEM_INCLUIU, DATA_INCLUSAO, EXCLUIDO) VALUES "
    
    clausulasInsert = ""
    For j = LBound(listaInclusao) To UBound(listaInclusao)
        if clausulasInsert <> "" Then
            clausulasInsert = clausulasInsert & ", "
        end if
        
        clausulasInsert = clausulasInsert & "('" & listaInclusao(j) & "', "
        clausulasInsert = clausulasInsert & idAnalista & ", "
        clausulasInsert = clausulasInsert & "'" & nomeAnalista & "', "
        clausulasInsert = clausulasInsert & "'" & email & "', "
        clausulasInsert = clausulasInsert & idUsuarioLogado & ", "
        clausulasInsert = clausulasInsert & "GETDATE(), NULL)"
    Next
    
    SQL = SQL & clausulasInsert
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao criar analista: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Analista atualizado com sucesso!"
End Sub


Sub Editar(listaEdicao)
    Dim SQL, PIR_CORPORATIVO, PIR_HELPDESK, j
    Dim nomeAnalista, email
    Dim updateStatements
    
    if idAnalista = "" Then
        resultado = false
        mensagem = "ID do analista não informado"
        Exit Sub
    end if
    
    ' Busca dados atualizados do analista
    SQL = "SELECT CRM_NOME as NOME, CRM_EMAILPRIMARIO as EMAIL FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAnalista
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAnalista = Replace(PIR_HELPDESK("NOME"), "'", "''")
        email = Replace(PIR_HELPDESK("EMAIL"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    else
        resultado = false
        mensagem = "Erro ao buscar dados do analista: " & Err.Description
        Exit Sub
    end if
    On Error Goto 0
    
    ' Atualiza cada registro da lista de edição
    For j = LBound(listaEdicao) To UBound(listaEdicao)
        SQL = "UPDATE ANALISTAS_ORCAMENTO SET "
        SQL = SQL & "NOME_USUARIO = '" & nomeAnalista & "', "
        SQL = SQL & "EMAIL = '" & email & "', "
        SQL = SQL & "QUEM_ALTEROU = " & idUsuarioLogado & ", "
        SQL = SQL & "DATA_ALTERACAO = GETDATE() "
        SQL = SQL & "WHERE CODIGO_USUARIO = " & idAnalista 
        SQL = SQL & " AND PLANTA = '" & listaEdicao(j) & "' "
        SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"
        
        On Error Resume Next
        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
        if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
            resultado = false
            mensagem = "Erro ao editar analista: " & Err.Description
            Exit Sub
        end if
        
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
        On Error Goto 0
    Next
    
    resultado = true
    mensagem = "Analista atualizado com sucesso!"
End Sub


Sub Excluir()
    Dim SQL, PIR_CORPORATIVO, j
    Dim clausulaNotIn
    
    clausulaNotIn = ""
    For j = LBound(listaPlantas) To UBound(listaPlantas)
        Dim plantaTemp
        plantaTemp = Trim(listaPlantas(j))
        if plantaTemp <> "" Then
            if clausulaNotIn <> "" Then
                clausulaNotIn = clausulaNotIn & ","
            end if
            clausulaNotIn = clausulaNotIn & "'" & plantaTemp & "'"
        end if
    Next
    
    SQL = "UPDATE ANALISTAS_ORCAMENTO SET "
    SQL = SQL & "EXCLUIDO = 1, "
    SQL = SQL & "QUEM_EXCLUIU = " & idUsuarioLogado & ", "
    SQL = SQL & "DATA_EXCLUSAO = GETDATE() "
    SQL = SQL & "WHERE CODIGO_USUARIO = " & idAnalista 
    SQL = SQL & " AND PLANTA NOT IN (" & clausulaNotIn & ") "
    SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir registros: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
End Sub

'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------

Function ExisteRegistro(planta)
    Dim registroExiste, SQL, PIR_CORPORATIVO
    registroExiste = false

    SQL = "SELECT TOP 1 CODIGO_USUARIO FROM ANALISTAS_ORCAMENTO WHERE CODIGO_USUARIO = " & idAnalista & " AND PLANTA = '" & planta & "' AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number = 0 AND isObject(PIR_CORPORATIVO) then
        if Not PIR_CORPORATIVO.EOF Then
            registroExiste = true
        end if
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    end if
    
    On Error Goto 0

    ExisteRegistro = registroExiste
End Function


Sub RegistrarLog()
    Dim SQL, LOG_CONN, PIR_HELPDESK
    Dim descricaoLog, acaoLog, nomeAnalista
    nomeAnalista = ""
    
    ' Busca o nome do analista para incluir no log
    SQL = "SELECT CRM_NOME as NOME FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAnalista
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAnalista = Replace(PIR_HELPDESK("NOME"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    end if
    
    On Error Goto 0
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de analista de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " criou o analista " & nomeAnalista
        Case "edicao"
            acaoLog = "Edição de analista de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " editou o analista " & nomeAnalista
        Case "exclusao"
            acaoLog = "Exclusão de analista de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " excluiu o analista " & nomeAnalista
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, IdUsuarioAlvo, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuarioLogado & ", "
    SQL = SQL & idAnalista & ", "
    SQL = SQL & "'Analistas de Orçamento', "
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
