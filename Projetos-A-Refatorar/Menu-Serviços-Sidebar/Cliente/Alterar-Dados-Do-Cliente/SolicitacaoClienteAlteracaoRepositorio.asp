<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<!--#Include File="../includes/variaveis.asp"-->
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
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "nivel", Session("NIVEL")
usuarioLogado.Add "cod", Session("CODUSER")
usuarioLogado.Add "planta", SESSION("FILIAL")

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
    Case "cancelar"
        resultado = Cancelar()
    Case "obterPorId"
        resultado = ObterPorId()
    Case "obterNivelUsuario"
        resultado = ObterNivelUsuario()
    Case "enviarParaDiretorComercial"
        resultado = EnviarParaDiretorComercial()
    Case "enviarParaFinanceiro"
        resultado = EnviarParaFinanceiro()
    Case "enviarParaControladoria"
        resultado = enviarParaControladoria()
    Case "avaliacaoDiretorComercial"
        resultado = AvaliacaoDiretorComercial()
    Case "avaliacaoFinanceiro"
        resultado = AvaliacaoFinanceiro()
    Case "avaliacaoControladoria"
        resultado = AvaliacaoControladoria()
    Case "finalizar"
        resultado = Finalizar()
    Case "listarCamposAlteracaoCliente"
        resultado = ListarCamposAlteracaoCliente()
    Case "listarStatus"
        resultado = ListarStatus()
    Case "listarClientes"
        resultado = ListarClientes()
    Case "listarHistorico"
        resultado = ListarHistorico()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select


' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    id = Request("filtroId")
    dataInclusaoDe = Request("dataInclusaoDe")
    dataInclusaoAte = Request("dataInclusaoAte")
    codigoPlanta = Request("codigoPlanta")
    status = Request("status")
    codigoCliente = Request("codigoCliente")
    tipoSolicitacao = Request("tipoSolicitacao")
    flagApenasPendentes = Request("flagApenasPendentes")

    nivelUsuario = ""
    respostaNivelUsuario = ObterNivelUsuario()
    Set respostaObj = JsonToDictionary(respostaNivelUsuario)
    
    If Not respostaObj("sucesso") Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário""}"
        Exit Function
    End If

    Set dadosUsuario = respostaObj("dados")
    nivelUsuario = dadosUsuario("nivel")

    If nivelUsuario = "" Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Nível de usuário inválido""}"
        Exit Function
    End If

    SQL = "SELECT TOP 100 "
    SQL = SQL & "   s.Id, s.Status, s.CodigoPlanta, s.TipoPessoa, s.IdUsuarioInclusao,"
    SQL = SQL & "   s.CodigoCliente, s.CodigoLoja, s.RazaoSocial, s.CpfCnpj, s.TipoSolicitacao, "
    SQL = SQL & "   FORMAT(s.DataInclusao , 'MM/dd/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "   CASE WHEN (                                                              "
	SQL = SQL & "   	SELECT COUNT(*) FROM SolicitacaoClienteAlteracaoDetalhamento d "
	SQL = SQL & "   	LEFT JOIN SolicitacaoClienteAlteracaoCampos c on d.IdCampo = c.Id "
	SQL = SQL & "   	WHERE c.CodigoCampo = 'CondicaoPagamento' "
	SQL = SQL & "   	AND d.IdSolicitacaoAlteracao = s.Id "
	SQL = SQL & "   ) > 0 THEN 1 ELSE 0 END AS PossuiCondicaoPagamento, "
    SQL = SQL & "   CASE WHEN (                                                              "
    SQL = SQL & "   	SELECT COUNT(*) FROM SolicitacaoClienteAlteracaoDetalhamento d "
	SQL = SQL & "   	LEFT JOIN SolicitacaoClienteAlteracaoCampos c on d.IdCampo = c.Id "
	SQL = SQL & "   	WHERE c.CodigoCampo = 'CodigoVendedor' "
	SQL = SQL & "   	AND d.IdSolicitacaoAlteracao = s.Id "
	SQL = SQL & "   ) > 0 THEN 1 ELSE 0 END AS PossuiCodigoVendedor "
    SQL = SQL & "FROM SolicitacaoClienteAlteracao s "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI  c ON c.CRM_R_E_C_N_O = s.IdUsuarioInclusao  "
    SQL = SQL & "WHERE s.FlagInativo = 0 "

    If nivelUsuario = "SOLICITANTE" then
        SQL = SQL & "AND s.IdUsuarioInclusao = '" & usuarioLogado("id") & "' "
    End If

    If flagApenasPendentes = "1" Then
        If nivelUsuario = "SOLICITANTE" then
            SQL = SQL & "AND (s.Status = 'EM_DIGITACAO')"
        ElseIf nivelUsuario = "DIRETOR_COMERCIAL" then
            SQL = SQL & "AND s.Status = 'PENDENTE_DIRETOR_COMERCIAL' "  
        ElseIf nivelUsuario = "CDAG" then
            SQL = SQL & "AND (s.Status = 'PENDENTE_FINALIZACAO')"
        ElseIf nivelUsuario = "FINANCEIRO" then
            SQL = SQL & "AND s.Status = 'PENDENTE_FINANCEIRO' "
        ElseIf nivelUsuario = "CONTROLADORIA" then
            SQL = SQL & "AND s.Status = 'PENDENTE_CONTROLADORIA' "          
        End If   
    End If   
    
    If tipoSolicitacao <> "" Then
        SQL = SQL & " AND s.TipoSolicitacao = '" & tipoSolicitacao & "'"
    End If

    If codigoPlanta <> "" Then
        SQL = SQL & " AND s.CodigoPlanta = '" & codigoPlanta & "'"
    End If

    If codigoCliente <> "" Then
        SQL = SQL & " AND s.CodigoCliente = '" & codigoCliente & "'"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND s.DataInclusao >= '" & dataInclusaoDe & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND s.DataInclusao < DATEADD(DAY, 1, '" & dataInclusaoAte & "')"
    End If

    SQL = SQL & " ORDER BY s.Id Desc"

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de Cadastros: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsListar("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """codigoPlanta"": """ & EscapeJSON(rsListar("CodigoPlanta")) & """, "
            jsonArray = jsonArray & """tipoPessoa"": """ & EscapeJSON(rsListar("TipoPessoa")) & """, "
            jsonArray = jsonArray & """tipoSolicitacao"": """ & EscapeJSON(rsListar("TipoSolicitacao")) & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & EscapeJSON(rsListar("CodigoCliente")) & """, "
            jsonArray = jsonArray & """codigoLoja"": """ & EscapeJSON(rsListar("CodigoLoja")) & """, "
            jsonArray = jsonArray & """razaoSocial"": """ & EscapeJSON(rsListar("RazaoSocial")) & """, "
            jsonArray = jsonArray & """cpfCnpj"": """ & EscapeJSON(rsListar("CpfCnpj")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsListar("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsListar("LoginUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """possuiCondicaoPagamento"": """ & EscapeJSON(rsListar("PossuiCondicaoPagamento")) & """, "
            jsonArray = jsonArray & """possuiCodigoVendedor"": """ & EscapeJSON(rsListar("PossuiCodigoVendedor")) & """ "
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
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado.""}"
        Exit Function
    End If
 
    SQL = "SELECT TOP 1 " 
    SQL = SQL & "   s.Id, s.Status, s.CodigoPlanta, s.TipoPessoa,  s.TipoSolicitacao, s.Justificativa, "
    SQL = SQL & "   s.CodigoCliente, s.CodigoLoja, s.RazaoSocial, s.CpfCnpj, s.IdUsuarioInclusao,"
    SQL = SQL & "   FORMAT(s.DataInclusao , 'MM/dd/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao "
    SQL = SQL & "FROM SolicitacaoClienteAlteracao s "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c ON c.CRM_R_E_C_N_O = s.IdUsuarioInclusao  "
    SQL = SQL & "WHERE s.FlagInativo = 0 "
    SQL = SQL & "AND s.Id = '" & id & "'"

    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de solicitações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Solicitação não encontrada.""}"
        Exit Function
    End If

    ' Obtem lista de campos
    SQL = "SELECT "
    SQL = SQL & "Id, IdSolicitacaoAlteracao, IdCampo, ValorAtual, ValorNovo "
    SQL = SQL & "FROM SolicitacaoClienteAlteracaoDetalhamento "
    SQL = SQL & "WHERE IdSolicitacaoAlteracao = " & id & " "
    SQL = SQL & "AND FlagInativo = 0 "
    SQL = SQL & "ORDER BY DataInclusao Desc"

    On Error Resume Next
    SET rsCampos = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCampos) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de campos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    camposJSON = "["

    If Not rsCampos.EOF Then
        Do While Not rsCampos.EOF
            If Not primeiroItem Then camposJSON = camposJSON & "," End If  

            camposJSON = camposJSON & "{"
            camposJSON = camposJSON & """id"": """ & EscapeJSON(rsCampos("Id")) & """, "
            camposJSON = camposJSON & """idSolicitacaoAlteracao"": """ & EscapeJSON(rsCampos("IdSolicitacaoAlteracao")) & """, "
            camposJSON = camposJSON & """idCampo"": """ & EscapeJSON(rsCampos("IdCampo")) & """, "
            camposJSON = camposJSON & """valorAtual"": """ & EscapeJSON(rsCampos("ValorAtual")) & """, "
            camposJSON = camposJSON & """valorNovo"": """ & EscapeJSON(rsCampos("ValorNovo")) & """ "
            camposJSON = camposJSON & "}"
            
            primeiroItem = False
            rsCampos.MoveNext
        Loop
    End If

    camposJSON = camposJSON & "]"

    rsCampos.Close()
    SET rsCampos = NOTHING
    
    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id")) & """, "
    jsonObjeto = jsonObjeto & """idUsuarioInclusao"": """ & EscapeJSON(rsObter("IdUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
    jsonObjeto = jsonObjeto & """codigoPlanta"": """ & EscapeJSON(rsObter("CodigoPlanta")) & """, "
    jsonObjeto = jsonObjeto & """tipoPessoa"": """ & EscapeJSON(rsObter("TipoPessoa")) & """, "
    jsonObjeto = jsonObjeto & """tipoSolicitacao"": """ & EscapeJSON(rsObter("TipoSolicitacao")) & """, "
    jsonObjeto = jsonObjeto & """codigoCliente"": """ & EscapeJSON(rsObter("CodigoCliente")) & """, "
    jsonObjeto = jsonObjeto & """codigoLoja"": """ & EscapeJSON(rsObter("CodigoLoja")) & """, "
    jsonObjeto = jsonObjeto & """razaoSocial"": """ & EscapeJSON(rsObter("RazaoSocial")) & """, "
    jsonObjeto = jsonObjeto & """cpfCnpj"": """ & EscapeJSON(rsObter("CpfCnpj")) & """, "
    jsonObjeto = jsonObjeto & """justificativa"": """ & EscapeJSON(rsObter("Justificativa")) & """, "
    jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao")) & """, "
    jsonObjeto = jsonObjeto & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObter("NomeUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """loginUsuarioInclusao"": """ & EscapeJSON(rsObter("LoginUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """campos"": " & camposJSON & ""
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("tipoPessoa") = Request("tipoPessoa")
    Dados("tipoSolicitacao") = Request("tipoSolicitacao")
    Dados("codigoCliente") = Request("codigoCliente")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("razaoSocial") = Request("razaoSocial")
    Dados("cpfCnpj") = Request("cpfCnpj")
    Dados("justificativa") = Request("justificativa")
    Dados("detalhamentoJSON") = Request("detalhamento")
    Dados("status") = "EM_DIGITACAO"

    'jaExiste = VerificarSolicitacaoPendente(Dados("cnpj"))
    'If jaExiste Then
        'Criar = "{""sucesso"": false, ""mensagem"": ""Já uma solicitação pendente para este cliente""}"
        'Exit Function
    'End If
    
    SQL = "BEGIN TRANSACTION; "
    SQL = SQL & "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @DadosSolicitacao TABLE (Id int); "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracao ("
    SQL = SQL & "   IdUsuarioInclusao, TipoSolicitacao, Status, CodigoPlanta, "
    SQL = SQL & "   CodigoLoja, RazaoSocial, CpfCnpj, TipoPessoa, CodigoCliente, Justificativa "
    SQL = SQL & ") OUTPUT INSERTED.Id INTO @DadosSolicitacao VALUES ("
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & Dados("tipoSolicitacao") & "', "
    SQL = SQL & "'" & Dados("status") & "', "
    SQL = SQL & "'" & Dados("codigoPlanta") & "', "
    SQL = SQL & "'" & Dados("codigoLoja") & "', "
    SQL = SQL & "'" & Dados("razaoSocial") & "', "
    SQL = SQL & "'" & Dados("cpfCnpj") & "', "
    SQL = SQL & "'" & Dados("tipoPessoa") & "', "
    SQL = SQL & "'" & Dados("codigoCliente") & "', "
    SQL = SQL & "'" & Dados("justificativa") & "' "
    SQL = SQL & "); "

    Dados("detalhamentoJSON") = Replace(Dados("detalhamentoJSON"), "[", "")
    Dados("detalhamentoJSON") = Replace(Dados("detalhamentoJSON"), "]", "")
    detalhamentoArray = Split(Dados("detalhamentoJSON"), "},{")
    
    valoresDetalhamento = ""
    For i = 0 To UBound(detalhamentoArray)
        item = detalhamentoArray(i)
        
        idCampo = ExtrairValorJSON(item, "idCampo")
        valorAtual = ExtrairValorJSON(item, "valorAtual")
        valorNovo = ExtrairValorJSON(item, "valorNovo")

        If i > 0 Then 
            valoresDetalhamento = valoresDetalhamento & ","
        End If
        
        valoresDetalhamento = valoresDetalhamento & "("
        valoresDetalhamento = valoresDetalhamento & "(SELECT Id FROM @DadosSolicitacao), "
        valoresDetalhamento = valoresDetalhamento & "'" & idCampo & "', "
        valoresDetalhamento = valoresDetalhamento & "'" & valorAtual & "', "
        valoresDetalhamento = valoresDetalhamento & "'" & valorNovo & "' "
        valoresDetalhamento = valoresDetalhamento & ")"
    Next

    If valoresDetalhamento <> "" Then
        SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoDetalhamento ("
        SQL = SQL & "   IdSolicitacaoAlteracao, IdCampo, ValorAtual, ValorNovo "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresDetalhamento & "; "
    End If

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "(SELECT Id FROM @DadosSolicitacao), "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'Criação', "
    SQL = SQL & "'' "
    SQL = SQL & "); "

    SQL = SQL & "SELECT Id FROM @DadosSolicitacao; "
    
    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL &   "ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL &   "COMMIT TRANSACTION "

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar Solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    id = rsCriar("Id")

    rsCriar.Close()
    SET rsCriar = NOTHING

    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Solicitação criada com sucesso!"", ""id"": " & id & "}"
End Function

Function Editar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("detalhamentoJSON") = Request("detalhamento")
    Dados("justificativa") = Request("justificativa")

    valoresDetalhamento = ""
    If Dados("detalhamentoJSON") <> "" Then
        Dados("detalhamentoJSON") = Replace(Dados("detalhamentoJSON"), "[", "")
        Dados("detalhamentoJSON") = Replace(Dados("detalhamentoJSON"), "]", "")

        detalhamentoArray = Split(Dados("detalhamentoJSON"), "},{")
        For i = 0 To UBound(detalhamentoArray)
            item = detalhamentoArray(i)
            
            idCampo = ExtrairValorJSON(item, "idCampo")
            valorAtual = ExtrairValorJSON(item, "valorAtual")
            valorNovo = ExtrairValorJSON(item, "valorNovo")

            If i > 0 Then 
                valoresDetalhamento = valoresDetalhamento & ","
            End If
            
            valoresDetalhamento = valoresDetalhamento & "("
            valoresDetalhamento = valoresDetalhamento & "'" & Dados("id") & "', "
            valoresDetalhamento = valoresDetalhamento & "'" & idCampo & "', "
            valoresDetalhamento = valoresDetalhamento & "'" & valorAtual & "', "
            valoresDetalhamento = valoresDetalhamento & "'" & valorNovo & "' "
            valoresDetalhamento = valoresDetalhamento & ")"
        Next
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "Justificativa = '" & Dados("justificativa") & "'; "

    If valoresDetalhamento <> "" Then
        SQL = SQL & "DELETE FROM SolicitacaoClienteAlteracaoDetalhamento WHERE IdSolicitacaoAlteracao = '" & Dados("id") & "'; "

        SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoDetalhamento ("
        SQL = SQL & "   IdSolicitacaoAlteracao, IdCampo, ValorAtual, ValorNovo "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresDetalhamento & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    Set rsEditar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar Solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Call RegistrarLog("edicao", Dados("id"))
    Editar = "{""sucesso"": true, ""mensagem"": ""Solicitação atualizada com sucesso!""}"
End Function

Function EnviarParaDiretorComercial()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        EnviarParaDiretorComercial = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    novoStatus = "PENDENTE_DIRETOR_COMERCIAL"
    acaoHistorico = "Envio para Diretor Comercial"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaDiretorComercial = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailEnvioParaNivel(Dados, "DIRETOR_COMERCIAL", "Diretor Comercial")
    Call RegistrarLog("envio_diretor_comercial",  Dados("id"))
    EnviarParaDiretorComercial = "{""sucesso"": true, ""mensagem"": ""Solicitação avaliada com sucesso!""}"
End Function

Function EnviarParaFinanceiro()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        EnviarParaFinanceiro = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    novoStatus = "PENDENTE_FINANCEIRO"
    acaoHistorico = "Envio para Financeiro"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaFinanceiro = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailEnvioParaNivel(Dados, "FINANCEIRO", "Financeiro")
    Call RegistrarLog("envio_financeiro",  Dados("id"))
    EnviarParaFinanceiro = "{""sucesso"": true, ""mensagem"": ""Solicitação avaliada com sucesso!""}"
End Function

Function EnviarParaControladoria()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        EnviarParaControladoria = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    novoStatus = "PENDENTE_CONTROLADORIA"
    acaoHistorico = "Envio para Controladoria"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaControladoria = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailEnvioParaNivel(Dados, "CONTROLADORIA", "Controladoria")
    Call RegistrarLog("envio_controladoria",  Dados("id"))
    EnviarParaControladoria = "{""sucesso"": true, ""mensagem"": ""Solicitação avaliada com sucesso!""}"
End Function

Function AvaliacaoDiretorComercial()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("flagAprovar") = Request("flagAprovar")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        AvaliacaoDiretorComercial = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End If
    
    If Dados("flagAprovar") = 1 Then
        novoStatus = "PENDENTE_FINANCEIRO"
        acaoHistorico = "Aprovação Diretor Comercial"
    Else
        novoStatus = "CANCELADO"
        acaoHistorico = "Rejeição Diretor Comercial"
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        AvaliacaoDiretorComercial = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Dados("flagAprovar") = 1 Then 
        Call EnviarEmailEnvioParaNivel(Dados, "CDAG", "CDAG")
        Call RegistrarLog("aprovacao_diretor_comercial",  Dados("id"))
    Else
        Call EnviarEmailSolicitacaoCancelada(Dados)
        Call RegistrarLog("reprovacao_diretor_comercial",  Dados("id"))
    End If

    AvaliacaoDiretorComercial = "{""sucesso"": true, ""mensagem"": ""Solicitação enviada com sucesso!""}"
End Function

Function AvaliacaoFinanceiro()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")
    Dados("flagAprovar") = Request("flagAprovar")

    If Dados("id") = "" Then
        AvaliacaoFinanceiro = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    If Dados("flagAprovar") = 1 Then
        novoStatus = "PENDENTE_CONTROLADORIA"
        acaoHistorico = "Aprovação Financeiro"
    Else
        novoStatus = "CANCELADO"
        acaoHistorico = "Rejeição Financeiro"
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        AvaliacaoFinanceiro = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
      
    If Dados("flagAprovar") = 1 Then   
        Call EnviarEmailEnvioParaNivel(Dados, "CONTROLADORIA", "Controladoria")
        Call RegistrarLog("aprovacao_financeiro",  Dados("id"))
    Else
        Call EnviarEmailSolicitacaoCancelada(Dados)
        Call RegistrarLog("reprovacao_financeiro",  Dados("id"))
    End If

    AvaliacaoFinanceiro = "{""sucesso"": true, ""mensagem"": ""Solicitação avaliada com sucesso!""}"
End Function

Function AvaliacaoControladoria()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")
    Dados("flagAprovar") = Request("flagAprovar")

    If Dados("id") = "" Then
        AvaliacaoControladoria = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    If Dados("flagAprovar") = 1 Then
        novoStatus = "PENDENTE_FINALIZACAO"
        acaoHistorico = "Aprovação Controladoria"
    Else
        novoStatus = "CANCELADO"
        acaoHistorico = "Rejeição Controladoria"
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & acaoHistorico & "', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        AvaliacaoControladoria = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
      
    If Dados("flagAprovar") = 1 Then   
        Call EnviarEmailEnvioParaNivel(Dados, "CONTROLADORIA", "Controladoria")
        Call RegistrarLog("aprovacao_controladoria",  Dados("id"))
    Else
        Call EnviarEmailSolicitacaoCancelada(Dados)
        Call RegistrarLog("reprovacao_controladoria",  Dados("id"))
    End If

    AvaliacaoControladoria = "{""sucesso"": true, ""mensagem"": ""Solicitação avaliada com sucesso!""}"
End Function

Function Finalizar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "   Status = 'FINALIZADO' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteAlteracaoHistorico ("
    SQL = SQL & "   IdSolicitacaoAlteracao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'Finalização', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
            
    Call EnviarEmailSolicitacaoFinalizada(Dados)
    Call RegistrarLog("finalizacao",  Dados("id"))
    Finalizar = "{""sucesso"": true, ""mensagem"": ""Solicitação finalizada com sucesso!""}"
End Function

Function Cancelar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    
    SQL = "UPDATE SolicitacaoClienteAlteracao SET "
    SQL = SQL & "Status = 'CANCELADO' "
    SQL = SQL & "WHERE Id = " & Dados("id")
  
    On Error Resume Next
    SET rsCancelar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCancelar) Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Erro ao cancelar: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailSolicitacaoCancelada(Dados)
    Call RegistrarLog("cancelamento", Dados("id"))
    Cancelar = "{""sucesso"": true, ""mensagem"": ""Solicitação cancelada com sucesso!""}"
End Function

Function ListarClientes()
    codigoPlanta = request("codigoPlanta")
    codCliente = request("codCliente")
    codigoLoja = request("codigoLoja")
    cpfCnpj = request("cpfCnpj")
    busca = Iif(request("busca") = "", "", trim(request("busca")))
    
    If codigoPlanta = "" Then
        ListarClientes = "{""sucesso"": false, ""mensagem"": ""Nenhuma planta informada.""}"
        Exit Function
    End If

    colunaCodigo = Iif(codigoPlanta = "05" OR codigoPlanta = "06", "s.A1_CDSIADN", "s.A1_COD")

    SQL = "SELECT "
    SQL = SQL &     colunaCodigo & " AS CodigoCliente, "
    SQL = SQL & "   s.A1_LOJA AS CodigoLoja, "
    SQL = SQL & "   s.A1_NOME AS RazaoSocial,"
    SQL = SQL & "   s.A1_PESSOA AS TipoPessoa, "
    SQL = SQL & "   s.A1_CGC AS CpfCnpj, "
    SQL = SQL & "   s.A1_VEND AS CodigoVendedor,"
    SQL = SQL & "   s2.A3_NOME AS NomeVendedor, "
    SQL = SQL & "   s.A1_MUN AS Municipio, "
	SQL = SQL & "   s.A1_EST AS UF, "
    SQL = SQL & "   s.A1_INSCR as InscricaoEstadual "
    SQL = SQL & "FROM SA1010 s "
    SQL = SQL & "LEFT JOIN SA3010 s2 on s2.A3_COD = s.A1_VEND AND s2.D_E_L_E_T_ = '' "
    SQL = SQL & "WHERE s.D_E_L_E_T_ = '' "

    If busca <> "" Then
        SQL = SQL & "AND ("
        SQL = SQL & "   A1_NOME like '%" & Ucase(busca) & "%' "
        SQL = SQL & "   OR " & colunaCodigo & " LIKE '%" & busca & "%' "
        SQL = SQL & ")"
    End If

    If codCliente <> "" Then
        SQL = SQL & "AND " & colunaCodigo & " = '" & codCliente & "' "
    End If

    If codigoLoja <> "" Then
        SQL = SQL & "AND s.A1_LOJA ='" & codigoLoja & "' "
    End If

    If cpfCnpj <> "" Then
        SQL = SQL & "AND s.A1_CGC ='" & cpfCnpj & "' "
    End If

    If usuarioLogado("nivel") <> 0 then
        SQL = SQL & "AND s.A1_VEND = '" & usuarioLogado("cod") & "' "
    End If

    SQL = SQL & " ORDER BY s.A1_NOME "

    On Error Resume Next
    SET rsClientes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsClientes) Then
        ListarClientes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsClientes.EOF Then
        Do While Not rsClientes.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """codigoCliente"": """ & EscapeJSON(rsClientes("CodigoCliente")) & """, "
            dadosJSON = dadosJSON & """codigoLoja"": """ & EscapeJSON(rsClientes("CodigoLoja")) & """, "
            dadosJSON = dadosJSON & """razaoSocial"": """ & EscapeJSON(rsClientes("RazaoSocial")) & """, "
            dadosJSON = dadosJSON & """tipoPessoa"": """ & EscapeJSON(rsClientes("TipoPessoa")) & """, "
            dadosJSON = dadosJSON & """municipio"": """ & EscapeJSON(rsClientes("Municipio")) & """, "
            dadosJSON = dadosJSON & """uf"": """ & EscapeJSON(rsClientes("UF")) & """, "
            dadosJSON = dadosJSON & """inscricaoEstadual"": """ & EscapeJSON(rsClientes("InscricaoEstadual")) & """, "
            dadosJSON = dadosJSON & """cpfCnpj"": """ & EscapeJSON(rsClientes("CpfCnpj")) & """, "
            dadosJSON = dadosJSON & """codigoVendedor"": """ & EscapeJSON(rsClientes("CodigoVendedor")) & """, "
            dadosJSON = dadosJSON & """nomeVendedor"": """ & EscapeJSON(rsClientes("NomeVendedor")) & """ "
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsClientes.MoveNext()
        Loop
    End If

    dadosJSON = dadosJSON & "]"
    rsClientes.Close()
    SET rsClientes = NOTHING

    ListarClientes = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ListarCamposAlteracaoCliente()
    codigoPlanta = request("codigoPlanta")
    codigoCliente = request("codigoCliente")
    codigoLoja = request("codigoLoja")
    
    If codigoPlanta = "" Then
        ListarCamposAlteracaoCliente = "{""sucesso"": false, ""mensagem"": ""Nenhuma planta informada.""}"
        Exit Function
    End If

    If codigoCliente = "" Then
        ListarCamposAlteracaoCliente = "{""sucesso"": false, ""mensagem"": ""Nenhum cliente informado.""}"
        Exit Function
    End If

    If codigoLoja = "" Then
        ListarCamposAlteracaoCliente = "{""sucesso"": false, ""mensagem"": ""Nenhuma loja informada.""}"
        Exit Function
    End If

    Set retornoCliente = ObterDadosClienteAlteracao(codigoPlanta, codigoCliente, codigoLoja)
    If Not retornoCliente("sucesso") Then
        ListarCamposAlteracaoCliente = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON(retornoCliente("mensagem")) & """}"
        Exit Function
    End If
    Set dadosCliente = retornoCliente("dados")

    SQL = "SELECT "
    SQL = SQL & "   Id as IdCampo, " 
    SQL = SQL & "   CodigoCampo, " 
    SQL = SQL & "   NomeExibicao, " 
    SQL = SQL & "   GrupoCampo, " 
    SQL = SQL & "   TipoDado, " 
    SQL = SQL & "   Ordem " 
    SQL = SQL & "FROM SolicitacaoClienteAlteracaoCampos s "
    SQL = SQL & "WHERE s.FlagInativo = 0"
    SQL = SQL & " ORDER BY s.Ordem "

    On Error Resume Next
    SET rsCampos = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCampos) Then
        ListarCamposAlteracaoCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsCampos.EOF Then
        Do While Not rsCampos.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If
            
            valorAtual = ObterValorAtualCampo(dadosCliente, rsCampos("CodigoCampo"))
            opcoes = "[]"

            If rsCampos("TipoDado") = "LISTA" Then
                opcoes = ObterOpcoesCampoAlteracaoCliente(codigoPlanta, rsCampos("CodigoCampo"), rsCampos("GrupoCampo"))
            End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """idCampo"": """ & EscapeJSON(rsCampos("IdCampo")) & """, "
            dadosJSON = dadosJSON & """codigoCampo"": """ & EscapeJSON(rsCampos("CodigoCampo")) & """, "
            dadosJSON = dadosJSON & """nomeExibicao"": """ & EscapeJSON(rsCampos("NomeExibicao")) & """, "
            dadosJSON = dadosJSON & """grupoCampo"": """ & EscapeJSON(rsCampos("GrupoCampo")) & """, "
            dadosJSON = dadosJSON & """tipoDado"": """ & EscapeJSON(rsCampos("TipoDado")) & """, "
            dadosJSON = dadosJSON & """ordem"": """ & EscapeJSON(rsCampos("Ordem")) & """, "
            dadosJSON = dadosJSON & """valorAtual"": """ & EscapeJSON(valorAtual) & """, "
            dadosJSON = dadosJSON & """opcoes"": " & opcoes & ""
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsCampos.MoveNext()
        Loop
    End If

    rsCampos.Close()
    SET rsCampos = NOTHING
    Set dadosCliente = NOTHING
    Set retornoCliente = NOTHING

    dadosJSON = dadosJSON & "]"
    ListarCamposAlteracaoCliente = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ObterDadosClienteAlteracao(codigoPlanta, codCliente, codigoLoja)
    Set retorno = Server.CreateObject("Scripting.Dictionary")
    retorno("sucesso") = False
    retorno("mensagem") = ""
    
    Set dados = Server.CreateObject("Scripting.Dictionary")
    Set retorno("dados") = dados

    conexaoProtheus = ObterConexaoProtheus(codigoPlanta)
    If conexaoProtheus = "" Then
        retorno("mensagem") = "Planta inválida."
        Set ObterDadosClienteAlteracao = retorno
        Exit Function
    End If

    colunaCodigo = Iif(codigoPlanta = "05" OR codigoPlanta = "06", "s.A1_CDSIADN", "s.A1_COD")
    colunaDiasViagem = Iif(codigoPlanta = "04", "s.A1_DIASVIU", "s.A1_DIASVIA")

    SQL = "SELECT TOP 1 "
    SQL = SQL & "   s.A1_LOJA AS CodigoLoja, "
    SQL = SQL & "   s.A1_NOME AS RazaoSocial, "
    SQL = SQL & "   s.A1_NREDUZ AS NomeFantasia, "
    SQL = SQL & "   s.A1_PESSOA AS TipoPessoa, "
    SQL = SQL & "   s.A1_CGC AS CpfCnpj, "
    SQL = SQL & "   s.A1_VEND AS CodigoVendedor, "
    SQL = SQL & "   s.A1_MUN AS Municipio, "
    SQL = SQL & "   s.A1_EST AS UF, "
    SQL = SQL & "   s.A1_INSCR AS InscricaoEstadual, "
    SQL = SQL & "   s.A1_TEL AS Telefone, "
    SQL = SQL & "   s.A1_EMAIL AS Email, "
    SQL = SQL & "   COALESCE(s.A1_CONTRIB, 0) AS ContribuinteIcms, "
    SQL = SQL & "   s.A1_COND AS CondicaoPagamento, "
    SQL = SQL & "   'Informação Indisponível' AS FormaPagamento, "
    SQL = SQL & "   s.A1_CEP AS Cep, "
    SQL = SQL & "   s.A1_COMPLEM AS Complemento, "
    SQL = SQL & "   s.A1_END AS Endereco, "
    SQL = SQL & "   s.A1_BAIRRO AS Bairro, "
    SQL = SQL & "   s.A1_EST AS Estado, "
    SQL = SQL & "   s.A1_CEPC AS CepCobranca, "
    SQL = SQL & "   'Informação Indisponível' AS ComplementoCobranca, "
    SQL = SQL & "   s.A1_ENDCOB AS EnderecoCobranca, "
    SQL = SQL & "   s.A1_BAIRROC AS BairroCobranca, "
    SQL = SQL & "   s.A1_MUNC AS MunicipioCobranca, "
    SQL = SQL & "   s.A1_ESTC AS EstadoCobranca, "
    SQL = SQL & "   s.A1_CEPE AS CepEntrega, "
    SQL = SQL & "   s.A1_COMPENT AS ComplementoEntrega, "
    SQL = SQL & "   s.A1_ENDENT AS EnderecoEntrega, "
    SQL = SQL & "   s.A1_BAIRROE AS BairroEntrega, "
    SQL = SQL & "   s.A1_MUNE AS MunicipioEntrega, "
    SQL = SQL & "   s.A1_ESTE AS EstadoEntrega, "
    SQL = SQL & "   s.A1_TIPO AS TipoCliente, "
    SQL = SQL & "   s.A1_CODSEG AS CodigoSegmento, "
    SQL = SQL & "   CASE WHEN A1_CLIGRP = 'S' THEN 1 ELSE 0 END AS ClienteGrupo, "
    SQL = SQL & "   s.A1_TPFRET AS TipoFrete, "
    SQL = SQL & "   s.A1_NATUREZ AS Natureza, "
    SQL = SQL & "   s.A1_CODPAIS AS PaisBacen, "
    SQL = SQL & "   s.A1_NFFAT AS ClienteRemessa, "
    SQL = SQL & "   CASE WHEN s.A1_IENCONT = 1 THEN 1 ELSE 0 END AS DestacaIe, "
    SQL = SQL & "   s.A1_TOLMAIS AS ToleranciaMaisPed, "
    SQL = SQL & "   s.A1_TOLMENO AS ToleranciaMenosPed, "
    SQL = SQL & "   'Informação Indisponível' AS ToleranciaQuantidadeMaior, "
    SQL = SQL & "   'Informação Indisponível' AS ToleranciaQuantidadeMenor, "

    If codigoPlanta = "05" OR codigoPlanta = "06" Then
        SQL = SQL & "   s.A1_TOLDTMA AS ToleranciaDiasMaior, "
        SQL = SQL & "   s.A1_TOLDTME AS ToleranciaDiasMenor, "
    Else
        SQL = SQL & "   'Informação Indisponível' AS ToleranciaDiasMaior, "
        SQL = SQL & "   'Informação Indisponível' AS ToleranciaDiasMenor, "
    End If

    SQL = SQL & "   " & colunaDiasViagem & " AS DiasViagem, "

    If codigoPlanta = "00" OR codigoPlanta = "04" Then
        SQL = SQL & "   s.A1_DIASENT AS DiasEntrega, "
        SQL = SQL & "   s.A1_DTEMAIO AS DiasEntregaMaior, "
        SQL = SQL & "   s.A1_DTEMENO AS DiasEntregaMenor "
    Else
        SQL = SQL & "   'Informação Indisponível' AS DiasEntrega, "
        SQL = SQL & "   'Informação Indisponível' AS DiasEntregaMaior, "
        SQL = SQL & "   'Informação Indisponível' AS DiasEntregaMenor "
    End If

    SQL = SQL & "FROM SA1010 s "
    SQL = SQL & "WHERE s.D_E_L_E_T_ = '' "
    SQL = SQL & "AND " & colunaCodigo & " = '" & Replace(codCliente, "'", "''") & "' "
    SQL = SQL & "AND s.A1_LOJA = '" & Replace(codigoLoja, "'", "''") & "' "

    On Error Resume Next
    SET rsCliente = Conecta(conexaoProtheus, SQL)
    If Err.Number <> 0 OR Not IsObject(rsCliente) Then
        retorno("mensagem") = "Erro ao buscar dados do cliente: " & EscapeJSON(Err.Description)
        Set ObterDadosClienteAlteracao = retorno
        Exit Function
    End If
    On Error Goto 0

    If rsCliente.EOF Then
        retorno("mensagem") = "Cliente não encontrado."
        rsCliente.Close()
        SET rsCliente = NOTHING
        Set ObterDadosClienteAlteracao = retorno
        Exit Function
    End If

    For Each campo In rsCliente.Fields
        dados(campo.Name) = campo.Value
    Next

    rsCliente.Close()
    SET rsCliente = NOTHING

    retorno("sucesso") = True
    retorno("mensagem") = "Dados do cliente obtidos com sucesso."
    Set ObterDadosClienteAlteracao = retorno
End Function

Function ObterValorAtualCampo(dadosCliente, codigoCampo)
    codigoCampo = Trim(codigoCampo & "")
    valor = ""

    If IsObject(dadosCliente) Then
        If dadosCliente.Exists(codigoCampo) Then
            If IsNull(dadosCliente(codigoCampo)) Then
                valor = ""
            Else
                valor = dadosCliente(codigoCampo) & ""
            End If
        End If
    End If

    ObterValorAtualCampo = valor
End Function

Function ObterOpcoesCampoAlteracaoCliente(codigoPlanta, codigoCampo, grupoCampo)
    chave = UCase(Trim(codigoCampo & ""))
    If chave = "" Then chave = UCase(Trim(grupoCampo & ""))

    Select Case chave
        Case "CONDICAOPAGAMENTO"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesCondicaoPagamento(codigoPlanta)
        Case "FORMAPAGAMENTO"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesFormaPagamento(codigoPlanta)
        Case "TIPOCLIENTE"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesTipoCliente(codigoPlanta)
        Case "TIPOFRETE"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesTipoFrete(codigoPlanta)
        Case "CODIGOSEGMENTO"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesSegmento(codigoPlanta)
        Case "CODIGOVENDEDOR"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesVendedor(codigoPlanta)
        Case "NATUREZA"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesNatureza(codigoPlanta)
        Case "PAISBACEN"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesPaisBacen(codigoPlanta)
        Case "TOLERANCIAQUANTIDADEMAIOR"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaQuantidade(codigoPlanta)
        Case "TOLERANCIAQUANTIDADEMENOR"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaQuantidade(codigoPlanta)
        Case "TOLERANCIADIASMAIOR"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaDiasMaior(codigoPlanta)
        Case "TOLERANCIADIASMENOR"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaDiasMenor(codigoPlanta)
        Case "TOLERANCIAMAISPED"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaQuantidade(codigoPlanta)
        Case "TOLERANCIAMENOSPED"
            ObterOpcoesCampoAlteracaoCliente = ObterOpcoesToleranciaQuantidade(codigoPlanta)
        Case Else
            ObterOpcoesCampoAlteracaoCliente = "[]"
    End Select
End Function

Function ObterOpcoesCondicaoPagamento(codigoPlanta)
    SQL = "SELECT " 
    SQL = SQL & "    E4_CODIGO AS Codigo,"
    SQL = SQL & "    E4_DESCRI AS Descricao "
    SQL = SQL & "FROM SE4010"

    On Error Resume Next
    SET rsCondicoes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCondicoes) Then
        ObterOpcoesCondicaoPagamento = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCondicoes.EOF Then
        Do While Not rsCondicoes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsCondicoes("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsCondicoes("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCondicoes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCondicoes.Close()
    SET rsCondicoes = NOTHING

    ObterOpcoesCondicaoPagamento = jsonArray
End Function

Function ObterOpcoesFormaPagamento(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "D" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Boleto" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "A" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Ordem de Pagamento" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ObterOpcoesFormaPagamento = jsonArray
End Function

Function ObterOpcoesTipoCliente(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "PRODUTOR RURAL" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "R" & """, "
    jsonArray = jsonArray & """descricao"": """ & "REVENDEDOR" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CONS. FINAL" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "S" & """, "
    jsonArray = jsonArray & """descricao"": """ & "SOLIDARIO" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "E" & """, "
    jsonArray = jsonArray & """descricao"": """ & "EXPORTAÇÃO" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ObterOpcoesTipoCliente = jsonArray
End Function

Function ObterOpcoesSegmento(codigoPlanta)
    SQL = "SELECT "
    SQL = SQL & "   AOV_CODSEG AS Codigo, "
    SQL = SQL & "   AOV_DESSEG AS Descricao "
    SQL = SQL & "FROM AOV010 WHERE AOV_MSBLQL = '2' AND ISNULL(AOV010.D_E_L_E_T_,'') = '' "
    SQL = SQL & "ORDER BY AOV_DESSEG "

    On Error Resume Next
    SET rsSegmentos = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSegmentos) Then
        ObterOpcoesSegmento = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSegmentos.EOF Then
        Do While Not rsSegmentos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsSegmentos("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsSegmentos("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSegmentos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSegmentos.Close()
    SET rsSegmentos = NOTHING

    ObterOpcoesSegmento = jsonArray
End Function

Function ObterOpcoesVendedor(codigoPlanta)
    SQL = "SELECT "
    SQL = SQL & "   A3_COD AS Codigo, "
    SQL = SQL & "   Concat(A3_COD, '-', RTRIM(LTRIM(A3_NOME))) AS Descricao "
    SQL = SQL & "FROM SA3010 WHERE D_E_L_E_T_ = '' AND A3_MSBLQL <> 1 "

    If usuarioLogado("nivel") = 1 then
        SQL = SQL & "AND A3_COD = '" & usuarioLogado("cod") & "' "
    End If
    
    SQL = SQL & "ORDER BY A3_NOME "

    On Error Resume Next
    SET rsVendedores = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsVendedores) Then
        ObterOpcoesVendedor = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsVendedores.EOF Then
        Do While Not rsVendedores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsVendedores("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsVendedores("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsVendedores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsVendedores.Close()
    SET rsVendedores = NOTHING

    ObterOpcoesVendedor = jsonArray
End Function

Function ObterOpcoesTipoFrete(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CIF" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "F" & """, "
    jsonArray = jsonArray & """descricao"": """ & "FOB" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"

    ObterOpcoesTipoFrete = jsonArray
End Function

Function ObterOpcoesNatureza(codigoPlanta)
    jsonArray = "["

    If codigoPlanta = "00" OR codigoPlanta = "04" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """valor"": """ & "2010101" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010101 - VENDAS DE PAPEIS" & """"
        jsonArray = jsonArray & "}"
        jsonArray = jsonArray & ","

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """valor"": """ & "2010102" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010102 - VENDAS DE CHAPAS" & """"
        jsonArray = jsonArray & "}"
        jsonArray = jsonArray & ","

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """valor"": """ & "2010103" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010103 - VENDA CAIXAS MALETA NORMAL" & """"
        jsonArray = jsonArray & "}"
    ElseIf codigoPlanta = "05" OR codigoPlanta = "06" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """valor"": """ & "20.10.100" & """, "
        jsonArray = jsonArray & """descricao"": """ & "20.10.100" & """"
        jsonArray = jsonArray & "}"
    ElseIf codigoPlanta = "07" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """valor"": """ & "05.01.05" & """, "
        jsonArray = jsonArray & """descricao"": """ & "05.01.05" & """"
        jsonArray = jsonArray & "}"
    End If

    jsonArray = jsonArray & "]"
    ObterOpcoesNatureza = jsonArray
End Function

Function ObterOpcoesPaisBacen(codigoPlanta)
    SQL = "SELECT " 
    SQL = SQL & "    CCH_CODIGO AS Codigo,"
    SQL = SQL & "    CCH_PAIS AS Descricao "
    SQL = SQL & "FROM CCH010"

    On Error Resume Next
    SET rsPaises = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsPaises) Then
        ObterOpcoesPaisBacen = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsPaises.EOF Then
        Do While Not rsPaises.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsPaises("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsPaises("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsPaises.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsPaises.Close()
    SET rsPaises = NOTHING

    ObterOpcoesPaisBacen = jsonArray
End Function

Function ObterOpcoesToleranciaQuantidade(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "0" & """, "
    jsonArray = jsonArray & """descricao"": """ & "0" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "5" & """, "
    jsonArray = jsonArray & """descricao"": """ & "5" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "10" & """, "
    jsonArray = jsonArray & """descricao"": """ & "10" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"

    ObterOpcoesToleranciaQuantidade = jsonArray
End Function

Function ObterOpcoesToleranciaDiasMaior(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "0" & """, "
    jsonArray = jsonArray & """descricao"": """ & "0" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "1" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","
    
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "2" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"

    ObterOpcoesToleranciaDiasMaior = jsonArray
End Function

Function ObterOpcoesToleranciaDiasMenor(codigoPlanta)
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "0" & """, "
    jsonArray = jsonArray & """descricao"": """ & "0" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "-1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "-1" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","
    
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "-2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "-2" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "-3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "-3" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """valor"": """ & "-4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "-4" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"

    ObterOpcoesToleranciaDiasMenor = jsonArray
End Function


Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "EM_DIGITACAO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Em digitação" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_DIRETOR_COMERCIAL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Diretor Comercial" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_FINANCEIRO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Financeiro" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_CONTROLADORIA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Controladoria" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_FINALIZACAO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Finalizacao" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "FINALIZADO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Finalizado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CANCELADO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Cancelado" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposSolicitacao()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "DESBLOQUEIO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Desbloqueio Simples" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "DESBLOQUEIO_COM_ALTERACAO_DADOS" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Desbloqueio com Alteração de Dados" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "BLOQUEIO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Bloqueio Simples" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "ALTERACAO_DADOS" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Alteração de Dados" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposSolicitacao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistorico()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")

    SQL = "SELECT "
    SQL = SQL & "   h.Id, h.IdSolicitacaoAlteracao, h.IdUsuarioInclusao, h.Acao, "
    SQL = SQL & "   FORMAT(h.DataInclusao , 'dd/MM/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), h.Observacao), 'json') AS Observacao "
    SQL = SQL & "FROM SolicitacaoClienteAlteracaoHistorico h "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c ON c.CRM_R_E_C_N_O = h.IdUsuarioInclusao "
    SQL = SQL & "WHERE h.FlagInativo = 0 "
    SQL = SQL & "AND h.IdSolicitacaoAlteracao = '" & Dados("id") & "' "
    SQL = SQL & "ORDER BY h.Id"

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistorico = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsHistorico("Id")) & """, "
            jsonArray = jsonArray & """idSolicitacaoAlteracao"": """ & EscapeJSON(rsHistorico("IdSolicitacaoAlteracao")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsHistorico("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsHistorico("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsHistorico("LoginUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """acao"": """ & EscapeJSON(rsHistorico("Acao")) & """, "
            jsonArray = jsonArray & """observacao"": """ & EscapeJSON(rsHistorico("Observacao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistorico = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

' Auxiliares
Function ObterNivelUsuario()
    nivelUsuario = ""
    
    SQL = "SELECT "
    SQL = SQL & "   ISNULL(NivelControladoria, 0) AS NivelControladoria, "
    SQL = SQL & "   ISNULL(NivelDiretorComercial, 0) AS NivelDiretorComercial, "
    SQL = SQL & "   ISNULL(NivelE2OPEN, 0) AS NivelCDAG, "
    SQL = SQL & "   ISNULL(NivelFinanceiro, 0) AS NivelFinanceiro "
    SQL = SQL & "FROM CRM_DTI WHERE CRM_D_E_L_E_T <> '*' "
    SQL = SQL & "AND CRM_R_E_C_N_O = " & usuarioLogado("id")
    
    SET rsPermissao = Conecta("PIR-HELPDESK", SQL)
    
    ' Se der erro finaliza fluxo
    If Err.Number <> 0 OR Not IsObject(rsPermissao) Then
        ObterNivelUsuario = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsPermissao.EOF Then
        If rsPermissao("NivelDiretorComercial") Then
            nivelUsuario = "DIRETOR_COMERCIAL"
        ElseIf rsPermissao("NivelControladoria") Then
            nivelUsuario = "CONTROLADORIA"
        ElseIf rsPermissao("NivelCDAG") Then
            nivelUsuario = "CDAG"
        ElseIf rsPermissao("NivelFinanceiro") Then
            nivelUsuario = "FINANCEIRO"
        End If
    End If

    rsPermissao.Close()
    SET rsPermissao = NOTHING

    ' Se não tiver nenhum nível consideramos que é vendedor
    If nivelUsuario = "" Then
        nivelUsuario = "SOLICITANTE"
    End If
    
    dadosJSON = "{""nivel"": """ & nivelUsuario & """}"
    ObterNivelUsuario = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ObterConexaoProtheus(codigoPlanta)
    Select Case codigoPlanta
        Case "00": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "04": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "05": ObterConexaoProtheus = "BEN-DADOSADV"
        Case "06": ObterConexaoProtheus = "FOR-DADOSADV"
        Case "07": ObterConexaoProtheus = "SAQ-DADOSADV"
        Case Else: ObterConexaoProtheus = ""
    End Select
End Function

Function AdicionarEmailLista(listaEmails, email)
    email = Trim(email & "")

    If email <> "" Then
        If InStr(1, ";" & LCase(listaEmails) & ";", ";" & LCase(email) & ";", vbTextCompare) = 0 Then
            If listaEmails <> "" Then listaEmails = listaEmails & ";"
            listaEmails = listaEmails & email
        End If
    End If

    AdicionarEmailLista = listaEmails
End Function

Function ObterDadosEmailSolicitacao(id)
    Set dadosEmail = Server.CreateObject("Scripting.Dictionary")
    dadosEmail("id") = id
    dadosEmail("razaoSocial") = ""
    dadosEmail("documento") = ""
    dadosEmail("codigoPlanta") = ""
    dadosEmail("nomeSolicitante") = ""
    dadosEmail("emailSolicitante") = ""

    SQL = "SELECT TOP 1 "
    SQL = SQL & "   a.Id, "
    SQL = SQL & "   a.RazaoSocial, "
    SQL = SQL & "   a.CpfCnpj, "
    SQL = SQL & "   a.CodigoPlanta, "
    SQL = SQL & "   ISNULL(t.CRM_NOME, a.IdUsuarioInclusao) AS NomeSolicitante, "
    SQL = SQL & "   ISNULL(t.CRM_EMAILPRIMARIO, '') AS EmailSolicitante "
    SQL = SQL & "FROM SolicitacaoClienteAlteracao a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t ON t.CRM_R_E_C_N_O = a.IdUsuarioInclusao "
    SQL = SQL & "WHERE a.Id = '" & Replace(id, "'", "''") & "' "

    On Error Resume Next
    SET rsDadosEmail = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(rsDadosEmail) Then
        If Not rsDadosEmail.EOF Then
            dadosEmail("id") = rsDadosEmail("Id") & ""
            dadosEmail("razaoSocial") = rsDadosEmail("RazaoSocial") & ""
            dadosEmail("documento") = rsDadosEmail("CpfCnpj") & ""
            dadosEmail("codigoPlanta") = rsDadosEmail("CodigoPlanta") & ""
            dadosEmail("nomeSolicitante") = rsDadosEmail("NomeSolicitante") & ""
            dadosEmail("emailSolicitante") = rsDadosEmail("EmailSolicitante") & ""
        End If

        rsDadosEmail.Close()
        SET rsDadosEmail = NOTHING
    End If
    On Error Goto 0

    Set ObterDadosEmailSolicitacao = dadosEmail
End Function

Function montarHTMLEmail(dadosEmail, titulo, texto, observacao)
    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>" & EscapeHTML(texto) & "</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Id:</strong> " & EscapeHTML(dadosEmail("id")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Razão Social:</strong> " & EscapeHTML(dadosEmail("razaoSocial")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>CNPJ/CPF:</strong> " & EscapeHTML(dadosEmail("documento")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & EscapeHTML(fNomePlanta(dadosEmail("codigoPlanta"))) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Solicitante:</strong> " & EscapeHTML(dadosEmail("nomeSolicitante")) & "</div>"

    If Trim(observacao & "") <> "" Then
        mensagem = mensagem & "<br>"
        mensagem = mensagem & "<div class=""info-item""><strong>Observação:</strong> " & EscapeHTML(observacao) & "</div>"
    End If

    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    montarHTMLEmail = mensagem
End Function

Sub EnviarEmailEnvioParaNivel(Dados, nivelDestino, nomeNivel)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = ObterEmailsUsuariosNivel(nivelDestino)
    If destinatarios = "" Then Exit Sub

    texto = "Uma solicitação de alteração de dados de cliente foi enviada para análise do nível " & nomeNivel & "."
    titulo = "Solicitação enviada para " & nomeNivel

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Alteracao de Dados de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailSolicitacaoFinalizada(Dados)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = AdicionarEmailLista("", dadosEmail("emailSolicitante"))
    If destinatarios = "" Then Exit Sub

    texto = "Uma solicitação de alteração de dados de cliente foi finalizada"
    titulo = "Alteração Dados de Cliente Finalizada"    

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Alteracao de Dados de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailSolicitacaoCancelada(Dados)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = AdicionarEmailLista("", dadosEmail("emailSolicitante"))
    If destinatarios = "" Then Exit Sub

    texto = "Uma solicitação de alteração de dados de cliente foi cancelada"
    titulo = "Alteração Dados de Cliente Cancelada"    

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Alteracao de Dados de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Function ObterEmailsUsuariosNivel(nivel)
    emails = ""
    colunaNivel = ""

    Select Case UCase(nivel)
        Case "CONTROLADORIA"
            colunaNivel = "NivelControladoria"
        Case "CDAG"
            colunaNivel = "NivelE2OPEN"
        Case "DIRETOR_COMERCIAL"
            colunaNivel = "NivelDiretorComercial"
        Case "FINANCEIRO"
            colunaNivel = "NivelFinanceiro"
    End Select

    If colunaNivel = "" Then
        ObterEmailsUsuariosNivel = ""
        Exit Function
    End If

    SQL = "SELECT DISTINCT CRM_EMAILPRIMARIO AS Email "
    SQL = SQL & "FROM CRM_DTI "
    SQL = SQL & "WHERE CRM_D_E_L_E_T <> '*' "
    SQL = SQL & "AND ISNULL(" & colunaNivel & ", 0) = 1 "
    SQL = SQL & "AND ISNULL(CRM_EMAILPRIMARIO, '') <> '' "

    On Error Resume Next
    SET rsEmailsNivel = Conecta("PIR-HELPDESK", SQL)

    If Err.Number = 0 AND IsObject(rsEmailsNivel) Then
        While Not rsEmailsNivel.EOF
            emails = AdicionarEmailLista(emails, rsEmailsNivel("Email") & "")
            rsEmailsNivel.MoveNext()
        Wend

        rsEmailsNivel.Close()
        SET rsEmailsNivel = NOTHING
    End If

    On Error Goto 0

    emails = "thiago.pereira@smurfitwestrock.com.br"
    ObterEmailsUsuariosNivel = emails
End Function

Sub RegistrarLog(acao, id)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Solicitação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a solicitação de id " & id
        Case "edicao"
            acaoLog = "Edição de Solicitação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou a solicitação de id " & id
        Case "envio_diretor_comercial"
            acaoLog = "Envio de Solicitação para Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " a solicitação de id " & id & " para o diretor comercial"
        Case "envio_controladoria"
            acaoLog = "Envio de Solicitação para Controladoria"
            descricaoLog = "Usuário " & usuarioLogado("login") & " a solicitação de id " & id & " para a controladoria"
        Case "envio_financeiro"
            acaoLog = "Envio de Solicitação para Financeiro"
            descricaoLog = "Usuário " & usuarioLogado("login") & " a solicitação de id " & id & " para o financeiro"
        Case "aprovacao_financeiro"
            acaoLog = "Aprovação de Solicitação pelo Financeiro"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_financeiro"
            acaoLog = "Reprovação de Solicitação pelo Financeiro"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a solicitação de id " & id
        Case "aprovacao_diretor_comercial"
            acaoLog = "Aprovação de Solicitação pelo Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_diretor_comercial"
            acaoLog = "Reprovação de Solicitação pelo Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a solicitação de id " & id
        Case "aprovacao_controladoria"
            acaoLog = "Aprovação de Solicitação pelo Controladoria"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_controladoria"
            acaoLog = "Reprovação de Solicitação pelo Controladoria"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a solicitação de id " & id
        Case "finalizacao"
            acaoLog = "Finalização de Solicitação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " finalizou a solicitação de id " & id
        Case "cancelamento"
            acaoLog = "Cancelamento de Solicitação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " cancelou a solicitação de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Solicitação de Alteracao de Dados de Cliente', "
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