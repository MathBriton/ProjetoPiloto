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

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")

resultado = AtualizarPermissoesUsuario()
Response.Write(resultado)

Function AtualizarPermissoesUsuario()
   ' Dados Formulario
    idUsuario = Request("idUsuario")
    nome = Request("nome")
    verTodasFncs = Request("verTodasFncs")
    alteraPo = Request("alteraPo")
    alteraPa = Request("alteraPa")
    ativaUsuario = Request("ativaUsuario")
    nivelAlteraPedido = Request("nivelAlteraPedido")
    permissao = Request("permissao")
    validaPreCadastroProduto = iif(Request("validaPreCadastroProduto") <> "", Request("validaPreCadastroProduto"), "NULL")
    eTools = iif(Request("eTools") <> "", Request("eTools"), "NULL")
    nivelLimiteCredito = iif(Request("nivelLimiteCredito") <> "", Request("nivelLimiteCredito"), "NULL")
    nivelControladoria = iif(Request("nivelControladoria") <> "", Request("nivelControladoria"), "NULL")
    nivelE2Open = iif(Request("nivelE2Open") <> "", Request("nivelE2Open"), "NULL")
    nivelControlesInternos = iif(Request("nivelControlesInternos") <> "", Request("nivelControlesInternos"), "NULL")
    nivelFinanceiro = iif(Request("nivelFinanceiro") <> "", Request("nivelFinanceiro"), "NULL")
    nivelDiretorComercial = iif(Request("nivelDiretorComercial") <> "", Request("nivelDiretorComercial"), "NULL")
    nivelFiscal = iif(Request("nivelFiscal") <> "", Request("nivelFiscal"), "NULL")
    nivelSPD = iif(Request("nivelSPD") <> "", Request("nivelSPD"), "NULL")
    nivelSolicitacaoAlteracaoPedido = iif(Request("nivelSolicitacaoAlteracaoPedido") <> "", Request("nivelSolicitacaoAlteracaoPedido"), "NULL")

    If idUsuario = "" Then
        AtualizarPermissoesUsuario = "{""sucesso"": false, ""mensagem"": ""Usuário é obrigatório.""}"
        Exit Function
    End If

    SQL = "UPDATE CRM_DTI SET"
    SQL = SQL & " CRM_PERMISSAO = " & "'" & permissao & "',"
    SQL = SQL & " CRM_VERTODASFNCS = " & "'" & verTodasFncs & "',"
    SQL = SQL & " CRM_ALTERAPAPEL = " & "'" &alteraPa & "',"
    SQL = SQL & " CRM_ALTERAPO = " & "'" & alteraPo & "',"
    SQL = SQL & " AtivaUsuario = " & "'" & ativaUsuario & "',"
    SQL = SQL & " CRM_NIVELALTERAPEDIDO = " & nivelAlteraPedido & ","
    SQL = SQL & " ValidaPreCadastroProduto = " & validaPreCadastroProduto & ","
    SQL = SQL & " ETools = " & eTools & ","
    SQL = SQL & " NivelLimiteCredito = " & nivelLimiteCredito & ","
    SQL = SQL & " NivelControladoria = " & nivelControladoria & ","
    SQL = SQL & " NivelE2OPEN = " & nivelE2Open & ","
    SQL = SQL & " NivelControlesInternos = " & nivelControlesInternos & ","
    SQL = SQL & " NivelFinanceiro = " & nivelFinanceiro & ","
    SQL = SQL & " NivelDiretorComercial = " & nivelDiretorComercial & ","
    SQL = SQL & " NivelFiscal = " & nivelFiscal & ","
    SQL = SQL & " NivelSPD = " & "'" & nivelSPD & "', "
    SQL = SQL & " NivelSolicitacaoAlteracaoPedido = " & "'" & nivelSolicitacaoAlteracaoPedido & "'"
    SQL = SQL & " WHERE CRM_R_E_C_N_O = " & idUsuario

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)

    If Err.Number <> 0 OR Not IsObject(PIR_HELPDESK) Then
        AtualizarPermissoesUsuario = "{""sucesso"": false, ""mensagem"": ""Erro ao criar atualizar permissões: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    PIR_HELPDESK.Close()
    SET PIR_HELPDESK = NOTHING
    On Error Goto 0
    
    Call RegistrarLog(idUsuario, nome)
    AtualizarPermissoesUsuario = "{""sucesso"": true, ""mensagem"": ""Permissões atualizadas com sucesso!""}"
End Function

Sub RegistrarLog(idUsuario, nomeUsuario)
    Dim SQL, PIR_CORPORATIVO
    Dim descricaoLog, acaoLog
    
    acaoLog = "Atualização de Permissões de Usuário"
    descricaoLog = "Usuário " & usuarioLogado("login") & " atualizou as permissões para o usuário " & nomeUsuario & ""
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, IdUsuarioAlvo, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & idUsuario & ", "
    SQL = SQL & "'Permissões de usuário', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub
%>