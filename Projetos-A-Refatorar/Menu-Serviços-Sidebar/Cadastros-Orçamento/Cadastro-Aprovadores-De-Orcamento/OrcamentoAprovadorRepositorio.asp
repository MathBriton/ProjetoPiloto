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
Dim parametroAcao, idAprovador, plantas
Dim resultado, mensagem
Dim plantasString, listaPlantas

idUsuarioLogado = Session("CODIGO")
loginUsuarioLogado = Session("USR")

acao = Request("acao")
idAprovador = Request("idAprovador")
nivel = Request("nivel")
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

if idAprovador = "" Then
    mensagem = "Id de aprovador não informado"
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
    if idAprovador = "" Then
        resultado = false
        mensagem = "ID do aprovador não informado"
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

    'Exclui todos registros que não estiverem presentes na lista de plantas (wherenotin)
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
    
    SQL = "UPDATE APROVADORES_ORCAMENTO SET "
    SQL = SQL & "EXCLUIDO = 1, "
    SQL = SQL & "QUEM_EXCLUIU = " & idUsuarioLogado & ", "
    SQL = SQL & "DATA_EXCLUSAO = GETDATE() "
    SQL = SQL & "WHERE ID_APROVADOR = " & idAprovador 
    SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao excluir aprovador: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Aprovador excluído com sucesso!"
    
    Call RegistrarLog()
End Sub


'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------

Sub Criar(listaInclusao)
    Dim SQL, PIR_CORPORATIVO, PIR_HELPDESK, j
    Dim nomeAprovador, email
    Dim valuesSQL
    
    ' Busca dados do aprovador
    SQL = "SELECT CRM_NOME as NOME, CRM_EMAILPRIMARIO as EMAIL FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAprovador
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAprovador = Replace(PIR_HELPDESK("NOME"), "'", "''")
        email = Replace(PIR_HELPDESK("EMAIL"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    else
        resultado = false
        mensagem = "Erro ao buscar dados do aprovador: " & Err.Description
        Exit Sub
    end if
    On Error Goto 0
    
    ' Montando clausulas para inserte em massa
    SQL = "INSERT INTO APROVADORES_ORCAMENTO (PLANTA, ID_APROVADOR, NOME_APROVADOR, EMAIL, NIVEL, QUEM_INCLUIU, DATA_INCLUSAO, EXCLUIDO) VALUES "
    
    clausulasInsert = ""
    For j = LBound(listaInclusao) To UBound(listaInclusao)
        if clausulasInsert <> "" Then
            clausulasInsert = clausulasInsert & ", "
        end if
        
        clausulasInsert = clausulasInsert & "('" & listaInclusao(j) & "', "
        clausulasInsert = clausulasInsert & idAprovador & ", "
        clausulasInsert = clausulasInsert & "'" & nomeAprovador & "', "
        clausulasInsert = clausulasInsert & "'" & email & "', "
        clausulasInsert = clausulasInsert & nivel & ", "
        clausulasInsert = clausulasInsert & idUsuarioLogado & ", "
        clausulasInsert = clausulasInsert & "GETDATE(), NULL)"
    Next
    
    SQL = SQL & clausulasInsert
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
        resultado = false
        mensagem = "Erro ao criar aprovadores: " & Err.Description
        Exit Sub
    end if
    
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0
    
    resultado = true
    mensagem = "Aprovador atualizado com sucesso!"
End Sub


Sub Editar(listaEdicao)
    Dim SQL, PIR_CORPORATIVO, PIR_HELPDESK, j
    Dim nomeAprovador, email
    Dim updateStatements
    
    if idAprovador = "" Then
        resultado = false
        mensagem = "ID do aprovador não informado"
        Exit Sub
    end if
    
    ' Busca dados atualizados do aprovador
    SQL = "SELECT CRM_NOME as NOME, CRM_EMAILPRIMARIO as EMAIL FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAprovador
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAprovador = Replace(PIR_HELPDESK("NOME"), "'", "''")
        email = Replace(PIR_HELPDESK("EMAIL"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    else
        resultado = false
        mensagem = "Erro ao buscar dados do aprovador: " & Err.Description
        Exit Sub
    end if
    On Error Goto 0
    
    ' Atualiza cada registro da lista de edição
    For j = LBound(listaEdicao) To UBound(listaEdicao)
        SQL = "UPDATE APROVADORES_ORCAMENTO SET "
        SQL = SQL & "NOME_APROVADOR = '" & nomeAprovador & "', "
        SQL = SQL & "EMAIL = '" & email & "', "
        SQL = SQL & "NIVEL = " & nivel & ", "
        SQL = SQL & "QUEM_ALTEROU = " & idUsuarioLogado & ", "
        SQL = SQL & "DATA_ALTERACAO = GETDATE() "
        SQL = SQL & "WHERE ID_APROVADOR = " & idAprovador 
        SQL = SQL & " AND PLANTA = '" & listaEdicao(j) & "' "
        SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"
        
        On Error Resume Next
        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
        if Err.Number <> 0 OR Not isObject(PIR_CORPORATIVO) then
            resultado = false
            mensagem = "Erro ao editar aprovador: " & Err.Description
            Exit Sub
        end if
        
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
        On Error Goto 0
    Next
    
    resultado = true
    mensagem = "Aprovador atualizado com sucesso!"
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
    
    SQL = "UPDATE APROVADORES_ORCAMENTO SET "
    SQL = SQL & "EXCLUIDO = 1, "
    SQL = SQL & "QUEM_EXCLUIU = " & idUsuarioLogado & ", "
    SQL = SQL & "DATA_EXCLUSAO = GETDATE() "
    SQL = SQL & "WHERE ID_APROVADOR = " & idAprovador 
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

    SQL = "SELECT TOP 1 ID_APROVADOR FROM APROVADORES_ORCAMENTO WHERE ID_APROVADOR = " & idAprovador & " AND PLANTA = '" & planta & "' AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"

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
    Dim descricaoLog, acaoLog, nomeAprovador
    nomeAprovador = ""
    
    ' Busca o nome do aprovador para incluir no log
    SQL = "SELECT CRM_NOME as NOME FROM CRM_DTI WHERE CRM_R_E_C_N_O = " & idAprovador
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    if Err.Number = 0 AND isObject(PIR_HELPDESK) AND Not PIR_HELPDESK.EOF then
        nomeAprovador = Replace(PIR_HELPDESK("NOME"), "'", "''")
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    end if
    
    On Error Goto 0
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " criou o aprovador " & nomeAprovador
        Case "edicao"
            acaoLog = "Edição de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " editou o aprovador " & nomeAprovador
        Case "exclusao"
            acaoLog = "Exclusão de aprovador de orçamento"
            descricaoLog = "Usuário " & loginUsuarioLogado & " excluiu o aprovador " & nomeAprovador
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, IdUsuarioAlvo, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuarioLogado & ", "
    SQL = SQL & idAprovador & ", "
    SQL = SQL & "'Aprovadores de Orçamento', "
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
