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
    Case "enviarParaFiscal"
        resultado = EnviarParaFiscal()
    Case "enviarParaDiretorComercial"
        resultado = EnviarParaDiretorComercial()
    Case "avaliacaoDiretorComercial"
        resultado = AvaliacaoDiretorComercial()
    Case "avaliacaoCDAG"
        resultado = AvaliacaoCDAG()
    Case "avaliacaoFinanceiro"
        resultado = AvaliacaoFinanceiro()
    Case "avaliacaoControladoria"
        resultado = AvaliacaoControladoria()
    Case "finalizar"
        resultado = Finalizar()
    Case "listarTiposContribuinte"
        resultado = ListarTiposContribuinte()
    Case "listarStatus"
        resultado = ListarStatus()
    Case "listarVendedores"
        resultado = ListarVendedores()
    Case "listarClientes"
        resultado = ListarClientes()
    Case "listarNaturezas"
        resultado = ListarNaturezas()
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
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    codigoPlanta = Request("filtroFilial")
    status = Request("filtroStatus")
    codCliente = Request("filtroCodCliente")
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
    SQL = SQL & "   s.Id, s.Status, s.CodigoPlantaOrigem, s.CodigoPlantaDestino, s.TipoPessoa, "
    SQL = SQL & "   s.CodigoCliente, s.CodigoLoja, s.RazaoSocial, s.CpfCnpj, s.Tes, "
    SQL = SQL & "   s.Natureza, s.CodigoVendedor, s.NomeVendedor, s.ToleranciaQuantidade, "
    SQL = SQL & "   CASE WHEN s.IcmsDiferido = 1 THEN 1 ELSE 0 END AS IcmsDiferido, "
    SQL = SQL & "   s.ToleranciaDias, s.IdUsuarioInclusao, s.ContribuinteIcms, s.InscricaoEstadual, "
    SQL = SQL & "   FORMAT(s.DataInclusao , 'MM/dd/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao "
    SQL = SQL & "FROM SolicitacaoClienteExtensao s "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI  c ON c.CRM_R_E_C_N_O = s.IdUsuarioInclusao  "
    SQL = SQL & "WHERE s.FlagInativo = 0 "

    If nivelUsuario = "SOLICITANTE" then
        SQL = SQL & "AND s.IdUsuarioInclusao = '" & usuarioLogado("id") & "' "
    End If

    If flagApenasPendentes = "1" Then
        If nivelUsuario = "SOLICITANTE" then
            SQL = SQL & "AND (s.Status = 'EM_DIGITACAO' OR s.Status = 'DEVOLVIDO')"
        ElseIf nivelUsuario = "FISCAL" then
            SQL = SQL & "AND s.Status = 'PENDENTE_FISCAL' "   
        ElseIf nivelUsuario = "DIRETOR_COMERCIAL" then
            SQL = SQL & "AND s.Status = 'PENDENTE_DIRETOR_COMERCIAL' "  
        ElseIf nivelUsuario = "CDAG" then
            SQL = SQL & "AND (s.Status = 'PENDENTE_AVALIACAO_DOCUMENTACAO' OR s.Status = 'PENDENTE_FINALIZACAO')"
        ElseIf nivelUsuario = "FINANCEIRO" then
            SQL = SQL & "AND s.Status = 'PENDENTE_FINANCEIRO' "
        ElseIf nivelUsuario = "CONTROLADORIA" then
            SQL = SQL & "AND s.Status = 'PENDENTE_CONTROLADORIA' "          
        End If   
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
            jsonArray = jsonArray & """codigoPlantaOrigem"": """ & EscapeJSON(rsListar("CodigoPlantaOrigem")) & """, "
            jsonArray = jsonArray & """codigoPlantaDestino"": """ & EscapeJSON(rsListar("CodigoPlantaDestino")) & """, "
            jsonArray = jsonArray & """tipoPessoa"": """ & EscapeJSON(rsListar("TipoPessoa")) & """, "
            jsonArray = jsonArray & """inscricaoEstadual"": """ & EscapeJSON(rsListar("InscricaoEstadual")) & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & EscapeJSON(rsListar("CodigoCliente")) & """, "
            jsonArray = jsonArray & """codigoLoja"": """ & EscapeJSON(rsListar("CodigoLoja")) & """, "
            jsonArray = jsonArray & """razaoSocial"": """ & EscapeJSON(rsListar("RazaoSocial")) & """, "
            jsonArray = jsonArray & """cpfCnpj"": """ & EscapeJSON(rsListar("CpfCnpj")) & """, "
            jsonArray = jsonArray & """tes"": """ & EscapeJSON(rsListar("Tes")) & """, "
            jsonArray = jsonArray & """icmsDiferido"": """ & EscapeJSON(rsListar("IcmsDiferido")) & """, "
            jsonArray = jsonArray & """contribuinteIcms"": """ & EscapeJSON(rsListar("ContribuinteIcms")) & """, "
            jsonArray = jsonArray & """natureza"": """ & EscapeJSON(rsListar("Natureza")) & """, "
            jsonArray = jsonArray & """codigoVendedor"": """ & EscapeJSON(rsListar("CodigoVendedor")) & """, "
            jsonArray = jsonArray & """nomeVendedor"": """ & EscapeJSON(rsListar("NomeVendedor")) & """, "
            jsonArray = jsonArray & """toleranciaQuantidade"": """ & EscapeJSON(rsListar("ToleranciaQuantidade")) & """, "
            jsonArray = jsonArray & """toleranciaDias"": """ & EscapeJSON(rsListar("ToleranciaDias")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsListar("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsListar("LoginUsuarioInclusao")) & """ "
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
    SQL = SQL & "   s.Id, s.Status, s.CodigoPlantaOrigem, s.CodigoPlantaDestino, s.TipoPessoa, "
    SQL = SQL & "   s.CodigoCliente, s.CodigoLoja, s.RazaoSocial, s.CpfCnpj, s.Tes, "
    SQL = SQL & "   s.Natureza, s.CodigoVendedor, s.NomeVendedor, s.ToleranciaQuantidade, "
    SQL = SQL & "   s.ToleranciaDias, s.IdUsuarioInclusao, s.InscricaoEstadual, "
    SQL = SQL & "   FORMAT(s.DataInclusao , 'MM/dd/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   CASE WHEN s.IcmsDiferido = 1 THEN 1 ELSE 0 END AS IcmsDiferido, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao, s.ContribuinteIcms "
    SQL = SQL & "FROM SolicitacaoClienteExtensao s "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c ON c.CRM_R_E_C_N_O = s.IdUsuarioInclusao  "
    SQL = SQL & "WHERE s.FlagInativo = 0 "
    SQL = SQL & "AND s.Id = '" & id & "'"

    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de cadastros: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Solicitação não encontrada.""}"
        Exit Function
    End If
    
    ' Obtem lista de anexos
    SQL = "SELECT "
    SQL = SQL & "   s.Id, s.IdSolicitacaoExtensao, s.Nome, s.Caminho, s.IdUsuarioInclusao, s.DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "   CASE WHEN s.FlagInvalido = 1 THEN 1 ELSE 0 END AS FlagInvalido, s.MotivoInvalido "
    SQL = SQL & "FROM SolicitacaoClienteExtensaoDocumento s "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c ON c.CRM_R_E_C_N_O = s.IdUsuarioInclusao  "
    SQL = SQL & "WHERE s.IdSolicitacaoExtensao = " & id & " "
    SQL = SQL & "AND s.FlagInativo = 0 "
    SQL = SQL & "ORDER BY s.DataInclusao Desc"

    On Error Resume Next
    SET rsAnexos = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAnexos) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de anexos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    anexosJSON = "["

    If Not rsAnexos.EOF Then
        Do While Not rsAnexos.EOF
            If Not primeiroItem Then anexosJSON = anexosJSON & "," End If  

            anexosJSON = anexosJSON & "{"
            anexosJSON = anexosJSON & """id"": """ & EscapeJSON(rsAnexos("Id")) & """, "
            anexosJSON = anexosJSON & """idSolicitacaoExtensao"": """ & EscapeJSON(rsAnexos("IdSolicitacaoExtensao")) & """, "
            anexosJSON = anexosJSON & """nome"": """ & EscapeJSON(rsAnexos("Nome")) & """, "
            anexosJSON = anexosJSON & """idUsuarioCriacao"": """ & EscapeJSON(rsAnexos("IdUsuarioInclusao")) & """, "
            anexosJSON = anexosJSON & """loginUsuarioInclusao"": """ & EscapeJSON(rsAnexos("LoginUsuarioInclusao")) & """, "
            anexosJSON = anexosJSON & """caminho"": """ & EscapeJSON(rsAnexos("Caminho")) & """, "
            anexosJSON = anexosJSON & """flagInvalido"": """ & EscapeJSON(rsAnexos("FlagInvalido")) & """, "
            anexosJSON = anexosJSON & """motivoInvalido"": """ & EscapeJSON(rsAnexos("MotivoInvalido")) & """, "
            anexosJSON = anexosJSON & """dataInclusao"": """ & EscapeJSON(rsAnexos("DataInclusao")) & """ "
            anexosJSON = anexosJSON & "}"
            
            primeiroItem = False
            rsAnexos.MoveNext
        Loop
    End If

    anexosJSON = anexosJSON & "]"

    rsAnexos.Close()
    SET rsAnexos = NOTHING

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id")) & """, "
    jsonObjeto = jsonObjeto & """idUsuarioInclusao"": """ & EscapeJSON(rsObter("IdUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
    jsonObjeto = jsonObjeto & """codigoPlantaOrigem"": """ & EscapeJSON(rsObter("CodigoPlantaOrigem")) & """, "
    jsonObjeto = jsonObjeto & """codigoPlantaDestino"": """ & EscapeJSON(rsObter("CodigoPlantaDestino")) & """, "
    jsonObjeto = jsonObjeto & """tipoPessoa"": """ & EscapeJSON(rsObter("TipoPessoa")) & """, "
    jsonObjeto = jsonObjeto & """inscricaoEstadual"": """ & EscapeJSON(rsObter("InscricaoEstadual")) & """, "
    jsonObjeto = jsonObjeto & """codigoCliente"": """ & EscapeJSON(rsObter("CodigoCliente")) & """, "
    jsonObjeto = jsonObjeto & """codigoLoja"": """ & EscapeJSON(rsObter("CodigoLoja")) & """, "
    jsonObjeto = jsonObjeto & """razaoSocial"": """ & EscapeJSON(rsObter("RazaoSocial")) & """, "
    jsonObjeto = jsonObjeto & """cpfCnpj"": """ & EscapeJSON(rsObter("CpfCnpj")) & """, "
    jsonObjeto = jsonObjeto & """tes"": """ & EscapeJSON(rsObter("Tes")) & """, "
    jsonObjeto = jsonObjeto & """icmsDiferido"": """ & EscapeJSON(rsObter("IcmsDiferido")) & """, "
    jsonObjeto = jsonObjeto & """contribuinteIcms"": """ & EscapeJSON(rsObter("ContribuinteIcms")) & """, "
    jsonObjeto = jsonObjeto & """natureza"": """ & EscapeJSON(rsObter("Natureza")) & """, "
    jsonObjeto = jsonObjeto & """codigoVendedor"": """ & EscapeJSON(rsObter("CodigoVendedor")) & """, "
    jsonObjeto = jsonObjeto & """nomeVendedor"": """ & EscapeJSON(rsObter("NomeVendedor")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaQuantidade"": """ & EscapeJSON(rsObter("ToleranciaQuantidade")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaDias"": """ & EscapeJSON(rsObter("ToleranciaDias")) & """, "
    jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao")) & """, "
    jsonObjeto = jsonObjeto & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObter("NomeUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """loginUsuarioInclusao"": """ & EscapeJSON(rsObter("LoginUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """anexos"": " & anexosJSON & ""
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigoPlantaOrigem") = Request("codigoPlantaOrigem")
    Dados("codigoPlantaDestino") = Request("codigoPlantaDestino")
    Dados("tipoPessoa") = Request("tipoPessoa")
    Dados("inscricaoEstadual") = Request("inscricaoEstadual")
    Dados("codigoCliente") = Request("codigoCliente")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("razaoSocial") = Request("razaoSocial")
    Dados("cpfCnpj") = Request("cpfCnpj")
    Dados("codigoVendedor") = Request("codigoVendedor")
    Dados("nomeVendedor") = Request("nomeVendedor")
    Dados("tes") = Request("tes")
    Dados("natureza") = Request("natureza")
    Dados("toleranciaQuantidade") = Request("toleranciaQuantidade")
    Dados("toleranciaDias") = Request("toleranciaDias")
    Dados("icmsDiferido") = Request("icmsDiferido")
    Dados("contribuinteIcms") = Request("contribuinteIcms")
    Dados("anexosJSON") = Request("anexos")
    Dados("status") = "EM_DIGITACAO"

    'jaExiste = VerificarCnpjExiste(Dados("cnpj"))
    'If jaExiste Then
        'Criar = "{""sucesso"": false, ""mensagem"": ""Já existe um cliente com esse CNPJ na planta de destino""}"
        'Exit Function
    'End If
    
    SQL = "BEGIN TRANSACTION; "
    SQL = SQL & "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @DadosSolicitacao TABLE (Id int); "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensao ("
    SQL = SQL & "   IdUsuarioInclusao, Status, CodigoPlantaOrigem, CodigoPlantaDestino, TipoPessoa, CodigoCliente,"
    SQL = SQL & "   CodigoLoja, RazaoSocial, CpfCnpj, Tes, Natureza, CodigoVendedor, NomeVendedor, "
    SQL = SQL & "   ToleranciaQuantidade, ToleranciaDias, IcmsDiferido, ContribuinteIcms, InscricaoEstadual "
    SQL = SQL & ") OUTPUT INSERTED.Id INTO @DadosSolicitacao VALUES ("
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & Dados("status") & "', "
    SQL = SQL & "'" & Dados("codigoPlantaOrigem") & "', "
    SQL = SQL & "'" & Dados("codigoPlantaDestino") & "', "
    SQL = SQL & "'" & Dados("tipoPessoa") & "', "
    SQL = SQL & "'" & Dados("codigoCliente") & "', "
    SQL = SQL & "'" & Dados("codigoLoja") & "', "
    SQL = SQL & "'" & Dados("razaoSocial") & "', "
    SQL = SQL & "'" & Dados("cpfCnpj") & "', "
    SQL = SQL & "'" & Dados("tes") & "', "
    SQL = SQL & "'" & Dados("natureza") & "', "
    SQL = SQL & "'" & Dados("codigoVendedor") & "', "
    SQL = SQL & "'" & Dados("nomeVendedor") & "', "
    SQL = SQL & "'" & Dados("toleranciaQuantidade") & "', "
    SQL = SQL & "'" & Dados("toleranciaDias") & "', "
    SQL = SQL & "'" & Dados("icmsDiferido") & "', "
    SQL = SQL & "'" & Dados("contribuinteIcms") & "', "
    SQL = SQL & "'" & Dados("inscricaoEstadual") & "' "
    SQL = SQL & "); "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "(SELECT Id FROM @DadosSolicitacao), "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'Criação', "
    SQL = SQL & "'' "
    SQL = SQL & "); "

    ' Anexos (se houver)
    Dados("anexosJSON") = Replace(Dados("anexosJSON"), "[", "")
    Dados("anexosJSON") = Replace(Dados("anexosJSON"), "]", "")
    anexosArray = Split(Dados("anexosJSON"), "},{")
    
    valoresAnexos = ""
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")

        If i > 0 Then 
            valoresAnexos = valoresAnexos & ","
        End If
        
        valoresAnexos = valoresAnexos & "("
        valoresAnexos = valoresAnexos & "(SELECT Id FROM @DadosSolicitacao), "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("id") & "'"
        valoresAnexos = valoresAnexos & ")"
    Next

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoDocumento ("
        SQL = SQL & "   IdSolicitacaoExtensao, Nome, Caminho, IdUsuarioInclusao "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresAnexos & "; "
    End If

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
    Dados("tes") = Request("tes")
    Dados("natureza") = Request("natureza")
    Dados("toleranciaQuantidade") = Iif(Request("toleranciaQuantidade") = "", 0, Request("toleranciaQuantidade"))
    Dados("toleranciaDias") = Iif(Request("toleranciaDias") = "", 0, Request("toleranciaDias"))
    Dados("icmsDiferido") = Request("icmsDiferido")
    Dados("contribuinteIcms") = Request("contribuinteIcms")
    Dados("anexosJSON") = Request("anexos")

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Tes = '" & Dados("tes") & "', "
    SQL = SQL & "   Natureza = '" & Dados("natureza") & "', "
    SQL = SQL & "   ToleranciaQuantidade = '" & Dados("toleranciaQuantidade") & "', "
    SQL = SQL & "   ToleranciaDias = '" & Dados("toleranciaDias") & "', "
    SQL = SQL & "   ContribuinteIcms = '" & Dados("contribuinteIcms") & "', "
    SQL = SQL & "   IcmsDiferido = '" & Dados("icmsDiferido") & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & MontarSQLAnexosEdicao(Dados("id"), Dados("anexosJSON"))

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

Function EnviarParaFiscal()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")

    If Dados("id") = "" Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End If

    ' Verifica se existem documentos inválido (reenvio para fiscal)
    SQL = "SELECT 1 FROM SolicitacaoClienteExtensaoDocumento "
    SQL = SQL & "WHERE IdSolicitacaoExtensao = '" & Dados("id") & "' "
    SQL = SQL & "AND FlagInativo = 0 "
    SQL = SQL & "AND FlagInvalido = 1"

    On Error Resume Next
    SET rsVerificar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsVerificar) Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If NOT rsVerificar.EOF Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Existem documentos pendentes, verifique a documentação antes de enviar.""}"
        Exit Function
    End If

    novoStatus = "PENDENTE_FISCAL"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'Envio para Fiscal', "
    SQL = SQL & "'" & Dados("observacao") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailEnvioParaNivel(Dados, "FISCAL", "Fiscal")
    Call RegistrarLog("envio_fiscal",  Dados("id"))
    EnviarParaFiscal = "{""sucesso"": true, ""mensagem"": ""Solicitação enviada com sucesso!""}"
End Function

Function EnviarParaDiretorComercial()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("observacao") = Request("observacao")
    Dados("tes") = Request("tes")

    If Dados("id") = "" Then
        EnviarParaDiretorComercial = "{""sucesso"": false, ""mensagem"": ""Nenhum id informado""}"
        Exit Function
    End IF

    novoStatus = "PENDENTE_DIRETOR_COMERCIAL"
    acaoHistorico = "Envio para Diretor Comercial"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "', "
    SQL = SQL & "   Tes = '" & Dados("tes") & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
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
        novoStatus = "PENDENTE_AVALIACAO_DOCUMENTACAO"
        acaoHistorico = "Aprovação Diretor Comercial"
    Else
        novoStatus = "CANCELADO"
        acaoHistorico = "Rejeição Diretor Comercial"
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
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

Function AvaliacaoCDAG()
    Set Dados = Server.CreateObject("Scripting.Dictionary")

    Dados("id") = Trim(Request("id") & "")
    Dados("observacao") = Request("observacao")
    Dados("flagAprovar") = Trim(Request("flagAprovar") & "")
    Dados("anexosJSON") = Request("anexos")

    observacaoTratada = Replace(Trim(Dados("observacao") & ""), "'", "''")

    If Dados("id") = "" Then
        AvaliacaoCDAG = "{""sucesso"": false, ""mensagem"": ""Solicitação não informada.""}"
        Exit Function
    End If

    If Dados("flagAprovar") = "" Then
        AvaliacaoCDAG = "{""sucesso"": false, ""mensagem"": ""Status da avaliação não informado.""}"
        Exit Function
    End If

    If Dados("anexosJSON") = "" Then
        AvaliacaoCDAG = "{""sucesso"": false, ""mensagem"": ""Nenhum documento enviado""}"
        Exit Function
    End If

    flagAprovar = CInt(Dados("flagAprovar"))

    If flagAprovar = 1 Then
        novoStatus = "PENDENTE_CONTROLADORIA"
        acaoHistorico = "Aprovação de Documentação"
    Else
        novoStatus = "DEVOLVIDO"
        acaoHistorico = "Reprovação de Documentação"
    End If

    SQL = ""
    SQL = SQL & "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "    IdSolicitacaoExtensao, "
    SQL = SQL & "    IdUsuarioInclusao, "
    SQL = SQL & "    Acao, "
    SQL = SQL & "    Observacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "    '" & Dados("id") & "', "
    SQL = SQL & "    '" & usuarioLogado("id") & "', "
    SQL = SQL & "    '" & acaoHistorico & "', "
    SQL = SQL & "    '" & observacaoTratada & "' "
    SQL = SQL & "); "

    ' Avaliação individual dos anexos/documentos
    anexosJSON = Dados("anexosJSON")
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")

    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)

        idDocumento = Trim(ExtrairValorJSON(anexo, "id") & "")
        statusDocumento = UCase(Trim(ExtrairValorJSON(anexo, "status") & ""))
        flagInvalido = Trim(ExtrairValorJSON(anexo, "flagInvalido") & "")
        motivoInvalidacao = Trim(ExtrairValorJSON(anexo, "motivoInvalidacao") & "")
        motivoInvalidacao = Replace(motivoInvalidacao, "'", "''")

        If flagInvalido = "" Then
            If statusDocumento = "INVALIDO" Then
                flagInvalido = "1"
            Else
                flagInvalido = "0"
            End If
        End If

        If CInt(flagInvalido) = 1 And motivoInvalidacao = "" Then
            AvaliacaoCDAG = "{""sucesso"": false, ""mensagem"": ""Informe o motivo de invalidação para todos os documentos inválidos.""}"
            Exit Function
        End If

        SQL = SQL & "UPDATE SolicitacaoClienteExtensaoDocumento SET "
        SQL = SQL & "    FlagInvalido = '" & CInt(flagInvalido) & "', "
        SQL = SQL & "    MotivoInvalido = '" & motivoInvalidacao & "', "
        SQL = SQL & "    DataAvaliacao = GETDATE() "
        SQL = SQL & "WHERE Id = '" & idDocumento & "' "
        SQL = SQL & "  AND IdSolicitacaoExtensao = '" & Dados("id") & "'; "
    Next

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    Set rsAvaliacaoCDAG = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsAvaliacaoCDAG) Then
        AvaliacaoCDAG = "{""sucesso"": false, ""mensagem"": ""Erro ao avaliar documentação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    If flagAprovar = 1 Then
        Call EnviarEmailEnvioParaNivel(Dados, "CONTROLADORIA", "Controladoria")
        Call RegistrarLog("aprovacao_documentacao", Dados("id"))
    Else
        Call EnviarEmailSolicitacaoDevolvida(Dados)
        Call RegistrarLog("reprovacao_documentacao", Dados("id"))
    End If

    AvaliacaoCDAG = "{""sucesso"": true, ""mensagem"": ""Documentação avaliada com sucesso!""}"
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

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
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

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = '" & novoStatus & "' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
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

    SQL = SQL & "UPDATE SolicitacaoClienteExtensao SET "
    SQL = SQL & "   Status = 'FINALIZADO' "
    SQL = SQL & "WHERE Id = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO SolicitacaoClienteExtensaoHistorico ("
    SQL = SQL & "   IdSolicitacaoExtensao, IdUsuarioInclusao, Acao, Observacao"
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
    
    SQL = "UPDATE SolicitacaoClienteExtensao SET "
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

Function ListarVendedores()
    codigoPlanta = Request("codigoPlanta")
    If codigoPlanta = "" Then
        ListarVendedores = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
    End If

    SQL = "SELECT A3_COD AS Codigo, Concat(A3_COD,'-',RTRIM(LTRIM(A3_NOME))) AS Descricao FROM SA3010 WHERE D_E_L_E_T_ = '' AND A3_MSBLQL <> 1 "

    If usuarioLogado("nivel") = 1 then
        SQL = SQL & "AND A3_COD = '" & coduser & "' "
    End If
    
    SQL = SQL & "ORDER BY A3_NOME "

    On Error Resume Next
    SET rsVendedores = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsVendedores) Then
        ListarVendedores = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar vendedores: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsVendedores.EOF Then
        Do While Not rsVendedores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsVendedores("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsVendedores("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsVendedores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsVendedores.Close()
    SET rsVendedores = NOTHING
    
    ListarVendedores = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarClientes()
    codigoPlanta = request("codigoPlanta")
    codCliente = request("codCliente")
    codLoja = request("codLoja")
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
    SQL = SQL & "WHERE s.D_E_L_E_T_ = ''"

    If busca <> "" Then
        SQL = SQL & " AND ("
        SQL = SQL & "   A1_NOME like '%" & Ucase(busca) & "%' "
        SQL = SQL & "   OR " & colunaCodigo & " LIKE '%" & busca & "%' "
        SQL = SQL & ")"
    End If

    If codCliente <> "" Then
        SQL = SQL & "AND " & colunaCodigo & " = '" & codCliente & "' "
    End If

    If codLoja <> "" Then
        SQL = SQL & "AND s.A1_LOJA ='" & codLoja & "' "
    End If

    If cpfCnpj <> "" Then
        SQL = SQL & "AND s.A1_CGC ='" & cpfCnpj & "' "
    End If

    If usuarioLogado("nivel") <> 0 then
        SQL = SQL & " AND s.A1_VEND = '" & usuarioLogado("cod") & "' "
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

Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "EM_DIGITACAO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Em digitação" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_FISCAL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Fiscal" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_DIRETOR_COMERCIAL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Diretor Comercial" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_CDAG" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente CDAG" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PENDENTE_CONTROLADORIA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pendente Controladoria" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposContribuinte()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CI" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Contribuinte Isento" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Contribuinte de ICMS" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "NA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pessoa Física" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "NC" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Não contribuinte" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PR" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Produtor Rural" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposContribuinte = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarNaturezas()
    codigoPlanta = Request("codigoPlanta")
    If codigoPlanta = "" Then
        ListarSegmentos = "{""sucesso"": false, ""mensagem"": ""Nenhuma Planta Informada""}"
    End If

    jsonArray = "["

    If codigoPlanta = "00" OR codigoPlanta = "04" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "2010101" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010101 - VENDAS DE PAPEIS" & """"
        jsonArray = jsonArray & "}"
        jsonArray = jsonArray & ","

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "2010102" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010102 - VENDAS DE CHAPAS" & """"
        jsonArray = jsonArray & "}"
        jsonArray = jsonArray & ","

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "2010103" & """, "
        jsonArray = jsonArray & """descricao"": """ & "2010103 - VENDA CAIXAS MALETA NORMAL" & """"
        jsonArray = jsonArray & "}"
    ElseIf codigoPlanta = "05" OR codigoPlanta = "06" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "20.10.100" & """, "
        jsonArray = jsonArray & """descricao"": """ & "20.10.100" & """"
        jsonArray = jsonArray & "}"
    ElseIf codigoPlanta = "07" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "05.01.05" & """, "
        jsonArray = jsonArray & """descricao"": """ & "05.01.05" & """"
        jsonArray = jsonArray & "}"
    End If

    jsonArray = jsonArray & "]"
    
    ListarNaturezas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistorico()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")

    SQL = "SELECT "
    SQL = SQL & "   h.Id, h.IdSolicitacaoExtensao, h.IdUsuarioInclusao, h.Acao, "
    SQL = SQL & "   FORMAT(h.DataInclusao , 'dd/MM/yyyy HH:mm:ss') AS DataInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, c.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), h.Observacao), 'json') AS Observacao "
    SQL = SQL & "FROM SolicitacaoClienteExtensaoHistorico h "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c ON c.CRM_R_E_C_N_O = h.IdUsuarioInclusao "
    SQL = SQL & "WHERE h.FlagInativo = 0 "
    SQL = SQL & "AND h.IdSolicitacaoExtensao = '" & Dados("id") & "' "
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
            jsonArray = jsonArray & """idSolicitacaoExtensao"": """ & EscapeJSON(rsHistorico("IdSolicitacaoExtensao")) & """, "
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
    SQL = SQL & "   ISNULL(NivelFinanceiro, 0) AS NivelFinanceiro, "
    SQL = SQL & "   ISNULL(NivelFiscal, 0) AS NivelFiscal "
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
        ElseIf rsPermissao("NivelFiscal") Then
            nivelUsuario = "FISCAL"
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

Function MontarSQLAnexosEdicao(idSolicitacao, anexosJSON)
    SQLAnexos = ""

    If Trim(anexosJSON & "") = "" Then
        MontarSQLAnexosEdicao = ""
        Exit Function
    End If

    If Trim(anexosJSON & "") = "[]" Then
        MontarSQLAnexosEdicao = ""
        Exit Function
    End If

    anexosJSON = Trim(anexosJSON)
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")

    If Trim(anexosJSON) = "" Then
        MontarSQLAnexosEdicao = ""
        Exit Function
    End If

    anexosArray = Split(anexosJSON, "},{")

    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)

        action = LCase(Trim(ExtrairValorJSON(anexo, "action")))
        originalId = Trim(ExtrairValorJSON(anexo, "originalId"))
        nome = Trim(ExtrairValorJSON(anexo, "name"))
        caminho = Trim(ExtrairValorJSON(anexo, "uploadedName"))

        Select Case action
            Case "added"
                If nome <> "" And caminho <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInserirAnexo(idSolicitacao, nome, caminho)
                End If
            Case "removed"
                If originalId <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInativarAnexo(idSolicitacao, originalId)
                End If
            
            Case "replaced"
                If originalId <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInativarAnexo(idSolicitacao, originalId)
                End If
                If nome <> "" And caminho <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInserirAnexo(idSolicitacao, nome, caminho)
                End If
            Case "unchanged"
                ' Não faz nada.
            Case Else
                ' Ação desconhecida.
                ' Neste primeiro momento, ignora para não quebrar a edição.
                ' Se preferir, podemos tratar como erro depois.
        End Select
    Next

    MontarSQLAnexosEdicao = SQLAnexos
End Function

Function MontarSQLInserirAnexo(idSolicitacao, nome, caminho)
    SQLInsert = ""
    SQLInsert = SQLInsert & "INSERT INTO SolicitacaoClienteExtensaoDocumento ("
    SQLInsert = SQLInsert & "   IdSolicitacaoExtensao, "
    SQLInsert = SQLInsert & "   Nome, "
    SQLInsert = SQLInsert & "   Caminho, "
    SQLInsert = SQLInsert & "   IdUsuarioInclusao, "
    SQLInsert = SQLInsert & "   FlagInativo, "
    SQLInsert = SQLInsert & "   DataInclusao "
    SQLInsert = SQLInsert & ") VALUES ("
    SQLInsert = SQLInsert & "   '" & idSolicitacao & "', "
    SQLInsert = SQLInsert & "   '" & nome & "', "
    SQLInsert = SQLInsert & "   '" & caminho & "', "
    SQLInsert = SQLInsert & "   '" & usuarioLogado("id") & "', "
    SQLInsert = SQLInsert & "   0, "
    SQLInsert = SQLInsert & "   GETDATE() "
    SQLInsert = SQLInsert & "); "

    MontarSQLInserirAnexo = SQLInsert
End Function

Function MontarSQLInativarAnexo(idSolicitacao, idAnexo)
    SQLUpdate = ""
    SQLUpdate = SQLUpdate & "UPDATE SolicitacaoClienteExtensaoDocumento SET "
    SQLUpdate = SQLUpdate & "   FlagInativo = 1 "
    SQLUpdate = SQLUpdate & "WHERE Id = '" & idAnexo & "' "
    SQLUpdate = SQLUpdate & "  AND IdSolicitacaoExtensao = '" & idSolicitacao & "' "
    SQLUpdate = SQLUpdate & "  AND FlagInativo = 0; "

    MontarSQLInativarAnexo = SQLUpdate
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
    dadosEmail("codigoPlantaOrigem") = ""
    dadosEmail("codigoPlantaDestino") = ""
    dadosEmail("nomeSolicitante") = ""
    dadosEmail("emailSolicitante") = ""

    SQL = "SELECT TOP 1 "
    SQL = SQL & "   a.Id, "
    SQL = SQL & "   a.RazaoSocial, "
    SQL = SQL & "   a.CpfCnpj, "
    SQL = SQL & "   a.CodigoPlantaOrigem, "
    SQL = SQL & "   a.CodigoPlantaDestino, "
    SQL = SQL & "   ISNULL(t.CRM_NOME, a.IdUsuarioInclusao) AS NomeSolicitante, "
    SQL = SQL & "   ISNULL(t.CRM_EMAILPRIMARIO, '') AS EmailSolicitante "
    SQL = SQL & "FROM SolicitacaoClienteExtensao a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t ON t.CRM_R_E_C_N_O = a.IdUsuarioInclusao "
    SQL = SQL & "WHERE a.Id = '" & Replace(id, "'", "''") & "' "

    On Error Resume Next
    SET rsDadosEmail = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(rsDadosEmail) Then
        If Not rsDadosEmail.EOF Then
            dadosEmail("id") = rsDadosEmail("Id") & ""
            dadosEmail("razaoSocial") = rsDadosEmail("RazaoSocial") & ""
            dadosEmail("documento") = rsDadosEmail("CpfCnpj") & ""
            dadosEmail("codigoPlantaOrigem") = rsDadosEmail("CodigoPlantaOrigem") & ""
            dadosEmail("codigoPlantaDestino") = rsDadosEmail("CodigoPlantaDestino") & ""
            dadosEmail("nomeSolicitante") = rsDadosEmail("NomeSolicitante") & ""
            dadosEmail("emailSolicitante") = rsDadosEmail("EmailSolicitante") & ""
            'dadosEmail("emailSolicitante") = "thiago.pereira@smurfitwestrock.com.br"
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
    mensagem = mensagem & "<div class=""info-item""><strong>Planta Origem:</strong> " & EscapeHTML(fNomePlanta(dadosEmail("codigoPlantaOrigem"))) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta Destino:</strong> " & EscapeHTML(fNomePlanta(dadosEmail("codigoPlantaDestino"))) & "</div>"
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

    texto = "Uma solicitação de extensão de cliente foi enviada para análise do nível " & nomeNivel & "."
    titulo = "Solicitação enviada para " & nomeNivel

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Extensão de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailSolicitacaoFinalizada(Dados)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = AdicionarEmailLista("", dadosEmail("emailSolicitante"))
    If destinatarios = "" Then Exit Sub

    texto = "Uma solicitação de extensão de cliente foi finalizada"
    titulo = "Extensão de Cliente Finalizada"    

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Extensão de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailSolicitacaoDevolvida(Dados)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = AdicionarEmailLista("", dadosEmail("emailSolicitante"))
    If destinatarios = "" Then Exit Sub

    texto = "A Documentação enviada foi reprovada. Pedimos que revise a documentação antes de realizar um novo envio"
    titulo = "Extensão de Cliente Devolvida"    

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Extensão de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailSolicitacaoCancelada(Dados)
    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    destinatarios = AdicionarEmailLista("", dadosEmail("emailSolicitante"))
    If destinatarios = "" Then Exit Sub

    texto = "Uma solicitação de extensão de cliente foi cancelada"
    titulo = "Extensão de Cliente Cancelada"    

    Set dadosEmail = ObterDadosEmailSolicitacao(Dados("id"))
    assunto = "[Solicitação de Extensão de Cliente] " & titulo & " - " & dadosEmail("id")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Function ObterEmailsUsuariosNivel(nivel)
    emails = ""
    colunaNivel = ""

    Select Case UCase(nivel)
        Case "FISCAL"
            colunaNivel = "NivelFiscal"
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
        Case "envio_fiscal"
            acaoLog = "Envio de Solicitação para Fiscal"
            descricaoLog = "Usuário " & usuarioLogado("login") & " a solicitação de id " & id & " para o fiscal"
        Case "envio_diretor_comercial"
            acaoLog = "Envio de Solicitação para Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " a solicitação de id " & id & " para o diretor comercial"
        Case "aprovacao_diretor_comercial"
            acaoLog = "Aprovação de Solicitação pelo Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_diretor_comercial"
            acaoLog = "Reprovação de Solicitação pelo Diretor Comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a solicitação de id " & id
        Case "aprovacao_documentacao"
            acaoLog = "Aprovação de Documentação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a documentação da solicitação de id " & id
        Case "reprovacao_documentacao"
            acaoLog = "Reprovação de Documentação"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a documentação da solicitação de id " & id
        Case "aprovacao_financeiro"
            acaoLog = "Aprovação de Solicitação pelo Financeiro"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_financeiro"
            acaoLog = "Reprovação de Solicitação pelo Financeiro"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reprovou a solicitação de id " & id
        Case "aprovacao_controladoria"
            acaoLog = "Aprovação de Solicitação pela Controladoria"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a solicitação de id " & id 
        Case "reprovacao_controladoria"
            acaoLog = "Reprovação de Solicitação pela Controladoria"
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
    SQL = SQL & "'Solicitação de Extensão de Cliente', "
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