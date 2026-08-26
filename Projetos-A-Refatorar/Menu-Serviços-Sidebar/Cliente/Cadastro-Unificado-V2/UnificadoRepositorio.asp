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
    Case "excluir"
        resultado = Excluir()
    Case "anexarDocumentos"
        resultado = AnexarDocumentos()
    Case "enviarParaFiscal"
        resultado = EnviarParaFiscal()
    Case "enviarParaControladoria"
        resultado = EnviarParaControladoria()
    Case "enviarParaCDAG"
        resultado = EnviarParaCDAG()
    Case "enviarParaControlesInternos"
        resultado = EnviarParaControlesInternos()
    Case "finalizar"
        resultado = Finalizar()
    Case "cancelar"
        resultado = Cancelar()
    Case "devolver"
        resultado = Devolver()
    Case "obterPorCodigo"
        resultado = ObterPorCodigo()
    Case "obterNivelUsuario"
        resultado = ObterNivelUsuario()
    Case "listarHistoricoVistos"
        resultado = ListarHistoricoVistos()
    Case "listarEstados"
        resultado = ListarEstados()
    Case "listarCidades"
        resultado = ListarCidades()
    Case "listarSegmentos"
        resultado = ListarSegmentos()
    Case "listarContasConciliacao"
        resultado = ListarContasConciliacao()
    Case "listarTiposIndustria"
        resultado = ListarTiposIndustria()
    Case "listarCondicoesPagamento"
        resultado = ListarCondicoesPagamento()
    Case "listarNaturezas"
        resultado = ListarNaturezas()
    Case "listarTiposCliente"
        resultado = ListarTiposCliente()
    Case "listarVendedores"
        resultado = ListarVendedores()
    Case "listarGruposConta"
        resultado = ListarGruposConta()
    Case "listarTiposContribuinte"
        resultado = ListarTiposContribuinte()
    Case "listarTiposDeclarante"
        resultado = ListarTiposDeclarante()
    Case "listarMetodosPagamento"
        resultado = ListarMetodosPagamento()
    Case "listarTiposFrete"
        resultado = ListarTiposFrete()
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
    SQL = SQL & "   Case WHEN a.PreCadastro = 'P' THEN 'P' + RIGHT('000000' + a.Codigo, 6) ELSE RTRIM(LTRIM(a.Codigo)) END as Codigo, "
    SQL = SQL & "   a.Codigo as CodigoOriginal, CONVERT(char(10), CONVERT(date, a.DataEmissao, 112), 23) AS DataInclusao, "
    SQL = SQL & "   a.RazaoSocial, a.NomeFantasia, a.StatusCli AS Status, Cnpj, Cpf,"
    SQL = SQL & "   a.QuemIncluiu, a.Filial as CodigoPlanta, a.CodLoja as CodigoLoja, "
    SQL = SQL & "   a.VolumeVendam2 as VolumeVendaMensalM2, a.PrecoEstm2 as PrecoVendaNetM2, "
    SQL = SQL & "   a.PrecoEstTon as precoVendaNetKg, a.Eshop, ISNULL(a.PreCadastro, '') as PreCadastro, a.LimiteCredito, "
    SQL = SQL & "   a.CodPro as CodigoProtheus, t2.CRM_NOME AS NomeUsuarioInclusao, t2.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "   CASE WHEN Cnpj <> '' THEN 'J' ELSE 'F' END AS TipoPessoa, "
    SQL = SQL & "   ("
    SQL = SQL & "       SELECT DATEDIFF(DAY, MIN(DataAlt), MAX(DataAlt)) "
    SQL = SQL & "       FROM UnificadoHistoricoCliente b WHERE b.CodigoCad = a.Codigo"
    SQL = SQL & "   ) LeadTime "
    SQL = SQL & "FROM UnificadoCliente a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t2 ON t2.CRM_LOGIN = a.QuemIncluiu "
    SQL = SQL & "WHERE Del = 0 "

    If nivelUsuario = "SOLICITANTE" then
        SQL = SQL & "AND QuemIncluiuCod = '" & usuarioLogado("id") & "' "
    End If

    If flagApenasPendentes = "1" Then
        SQL = SQL & "AND StatusCli <> 'X' "

        If nivelUsuario = "SOLICITANTE" then
            SQL = SQL & "AND StatusCli <> 'F'"
        ElseIf nivelUsuario = "CDAG" then
            SQL = SQL & "AND (StatusCli = 'D' OR StatusCli = 'A' Or StatusCli = 'E') "
        ElseIf nivelUsuario = "CONTROLES_INTERNOS" then
            SQL = SQL & "AND StatusCli = 'I' "
        ElseIf nivelUsuario = "CONTROLADORIA" then
            SQL = SQL & "AND StatusCli = 'C' "
        ElseIf nivelUsuario = "FISCAL" then
            SQL = SQL & "AND StatusCli = 'S' "               
        End If   
    End If   

    If codCliente <> "" then
        SQL = SQL & "AND Codigo = '" & codCliente & "' "
    End If
   
    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND DataEmissao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND DataEmissao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If status <> "" and status <> "Z" then
        SQL = SQL & "AND StatusCli = '" & status & "' "
    End If

    If codigoPlanta <> "" then
        SQL = SQL & "AND Filial = '" & codigoPlanta & "' "
    End If   

    If status = "Z" then
        SQL = SQL & "AND PreCadastro = 'P' "
    End If

    SQL = SQL & " ORDER BY Codigo Desc"

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
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsListar("Codigo")) & """, "
            jsonArray = jsonArray & """tipoPessoa"": """ & EscapeJSON(rsListar("TipoPessoa")) & """, "
            jsonArray = jsonArray & """cnpj"": """ & EscapeJSON(rsListar("Cnpj")) & """, "
            jsonArray = jsonArray & """cpf"": """ & EscapeJSON(rsListar("Cpf")) & """, "
            jsonArray = jsonArray & """codigoOriginal"": """ & EscapeJSON(rsListar("CodigoOriginal")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao")) & """, "
            jsonArray = jsonArray & """razaoSocial"": """ & EscapeJSON(rsListar("RazaoSocial")) & """, "
            jsonArray = jsonArray & """nomeFantasia"": """ & EscapeJSON(rsListar("NomeFantasia")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """quemIncluiu"": """ & EscapeJSON(rsListar("QuemIncluiu")) & """, "
            jsonArray = jsonArray & """codigoPlanta"": """ & EscapeJSON(rsListar("CodigoPlanta")) & """, "
            jsonArray = jsonArray & """codigoLoja"": """ & EscapeJSON(rsListar("CodigoLoja")) & """, "
            jsonArray = jsonArray & """volumeVendaMensalM2"": """ & EscapeJSON(rsListar("VolumeVendaMensalM2")) & """, "
            jsonArray = jsonArray & """precoVendaNetM2"": """ & EscapeJSON(rsListar("PrecoVendaNetM2")) & """, "
            jsonArray = jsonArray & """precoVendaNetKg"": """ & EscapeJSON(rsListar("PrecoVendaNetKg")) & """, "
            jsonArray = jsonArray & """eshop"": """ & EscapeJSON(rsListar("Eshop")) & """, "
            jsonArray = jsonArray & """preCadastro"": """ & EscapeJSON(rsListar("PreCadastro")) & """, "
            jsonArray = jsonArray & """limiteCredito"": """ & EscapeJSON(rsListar("LimiteCredito")) & """, "
            jsonArray = jsonArray & """codigoProtheus"": """ & EscapeJSON(rsListar("CodigoProtheus")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsListar("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsListar("LoginUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """leadTime"": """ & EscapeJSON(rsListar("LeadTime")) & """"
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

Function ObterPorCodigo()
    codigo = Request("codigo")
    If codigo = "" Then
        ObterPorCodigo = "{""sucesso"": false, ""mensagem"": ""Nenhum código informado.""}"
        Exit Function
    End If
 
    SQL = "SELECT TOP 1 "
    SQL = SQL & "   CASE WHEN a.PreCadastro = 'P' THEN 'P' + RIGHT('000000' + a.Codigo, 6) ELSE RTRIM(LTRIM(a.Codigo)) END as Codigo, "
    SQL = SQL & "   CASE WHEN a.PreCadastro = 'P' THEN 5 ELSE 1 END AS TipoSolicitacao, "
    SQL = SQL & "   CONVERT(char(10), CONVERT(date, a.DataEmissao, 112), 23) AS DataInclusao, "
    SQL = SQL & "   CASE WHEN Cnpj <> '' THEN 'J' ELSE 'F' END AS TipoPessoa, "
    SQL = SQL & "   a.Cnpj, "
    SQL = SQL & "   a.Cpf, "
    SQL = SQL & "   a.GrupoContas, "
    SQL = SQL & "   a.CodigoHfm, "
    SQL = SQL & "   a.Codigo as CodigoOriginal, "
    SQL = SQL & "   a.RazaoSocial, "
    SQL = SQL & "   a.NomeFantasia, "
    SQL = SQL & "   a.StatusCli AS Status, "
    SQL = SQL & "   a.Rua, "
    SQL = SQL & "   a.Numero, "
    SQL = SQL & "   a.Bairro, "
    SQL = SQL & "   a.Cep, "
    SQL = SQL & "   a.Estado, "
    SQL = SQL & "   a.Cidade, "
    SQL = SQL & "   a.Pais, "
    SQL = SQL & "   a.QuemIncluiu, "
    SQL = SQL & "   a.Filial as CodigoPlanta, "
    SQL = SQL & "   a.CodLoja as CodigoLoja, "
    SQL = SQL & "   a.VolumeVendam2 as VolumeVendaMensalM2, "
    SQL = SQL & "   a.PrecoEstm2 as PrecoVendaNetM2, "
    SQL = SQL & "   a.PrecoEstTon as precoVendaNetKg, "
    SQL = SQL & "   a.Eshop, "
    SQL = SQL & "   a.InscricaoEstadual, "
    SQL = SQL & "   a.InscricaoMunicipal, "
    SQL = SQL & "   ISNULL(a.PreCadastro, '') as PreCadastro, "
    SQL = SQL & "   a.KmEndereco as KmDestino, "
    SQL = SQL & "   a.Idioma, "
    SQL = SQL & "   a.Telefone, "
    SQL = SQL & "   a.EmailXml, "
    SQL = SQL & "   a.EmailPagamento as EmailOrdemPagamento, "
    SQL = SQL & "   a.Declarante as DeclarantePisCofins, "
    SQL = SQL & "   a.Contribuinte as ContribuinteIcms, "
    SQL = SQL & "   a.IcmsDiferido, "
    SQL = SQL & "   a.percentualDiferimento as PercentualDiferimento, "
    SQL = SQL & "   a.isencaoIPI as IsencaoIpi, "
    SQL = SQL & "   a.TipoIndustria, "
    SQL = SQL & "   a.CodigoSuframa, "
    SQL = SQL & "   a.ContaConciliacao, "
    SQL = SQL & "   a.ChaveOrdenacao, "
    SQL = SQL & "   a.CondicaoPagamento, "
    SQL = SQL & "   a.MetodoPagamento, "
    SQL = SQL & "   a.CodigoJuros, "
    SQL = SQL & "   a.LimiteCredito, "
    SQL = SQL & "   a.Seguimento as Segmento, "
    SQL = SQL & "   a.CodSap as CodigoSap, "
    SQL = SQL & "   a.Natureza, "
    SQL = SQL & "   a.Tipo as TipoCliente, "
    SQL = SQL & "   a.TipoFrete, "
    SQL = SQL & "   a.Tes, "
    SQL = SQL & "   a.CodigoVendedorPro as CodigoVendedor, "
    SQL = SQL & "   a.CodPro as CodigoProtheus, "
    SQL = SQL & "   a.DiasViagem, "
    SQL = SQL & "   a.DiasEntregaMaior as ToleranciaDiasMaior, "
    SQL = SQL & "   a.DiasEntregaMenor as ToleranciaDiasMenor, "
    SQL = SQL & "   CASE WHEN a.RuaCobPro <> '' THEN 1 ELSE 0 END as EnderecoEntregaDiferente, "
    SQL = SQL & "   a.RuaEntPro as RuaEntrega, "
    SQL = SQL & "   a.NumeroEntPro as NumeroEntrega, "
    SQL = SQL & "   a.BairroEntPro as BairroEntrega, "
    SQL = SQL & "   a.CepEntPro as CepEntrega, "
    SQL = SQL & "   a.EstadoEntPro as EstadoEntrega, "
    SQL = SQL & "   a.CidadeEntPro as CidadeEntrega, "
    SQL = SQL & "   CASE WHEN a.RuaCobPro <> '' THEN 1 ELSE 0 END as EnderecoCobrancaDiferente, "
    SQL = SQL & "   a.RuaCobPro as RuaCobranca, "
    SQL = SQL & "   a.NumeroCobPro as NumeroCobranca, "
    SQL = SQL & "   a.BairroCobPro as BairroCobranca, "
    SQL = SQL & "   a.CepCobPro as CepCobranca, "
    SQL = SQL & "   a.EstadoCobPro as EstadoCobranca, "
    SQL = SQL & "   a.CidadeCobPro as CidadeCobranca, "
    SQL = SQL & "   a.TQtdeMaior as ToleranciaQuantidadeMaior, "
    SQL = SQL & "   a.TQtdeMenor as ToleranciaQuantidadeMenor, "
    SQL = SQL & "   a.ObsEmissor as ObservacaoEmissor, "
    SQL = SQL & "   a.ObsCedag as ObservacaoCdag, "
    SQL = SQL & "   t2.CRM_NOME AS NomeUsuarioInclusao, "
    SQL = SQL & "   t2.CRM_LOGIN AS LoginUsuarioInclusao, "
    SQL = SQL & "   ("
    SQL = SQL & "       SELECT DATEDIFF(DAY, MIN(DataAlt), MAX(DataAlt)) "
    SQL = SQL & "       FROM UnificadoHistoricoCliente b WHERE b.CodigoCad = a.Codigo"
    SQL = SQL & "   ) LeadTime "
    SQL = SQL & "FROM UnificadoCliente a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t2 ON t2.CRM_LOGIN = a.QuemIncluiu "
    SQL = SQL & "WHERE Del = 0 "
    SQL = SQL & "AND Codigo = '" & codigo & "'"

    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorCodigo = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de cadastros: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorCodigo = "{""sucesso"": false, ""mensagem"": ""Cadastro não encontrado.""}"
        Exit Function
    End If

    
    ' Obtem lista de anexos
    SQL = "SELECT idUnificadoClienteDocumento AS Id, idUnificadoCliente, Filial, Nome, Usuario, Caminho, DataDaOperacao "
    SQL = SQL & "FROM UnificadoClienteDocumento "
    SQL = SQL & "WHERE idUnificadoCliente = " & codigo & " "
    SQL = SQL & "ORDER BY DataDaOperacao Desc"

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
            caminho = ""

            ' Tratativa para arquivos antigos
            If rsAnexos("Caminho") = "Clientes_Docs" Then
                caminho = rsAnexos("Nome")
            Else
                caminho = rsAnexos("Caminho")
            End If
            
            anexosJSON = anexosJSON & "{"
            anexosJSON = anexosJSON & """id"": """ & EscapeJSON(rsAnexos("Id")) & """, "
            anexosJSON = anexosJSON & """idUnificadoCliente"": """ & EscapeJSON(rsAnexos("idUnificadoCliente")) & """, "
            anexosJSON = anexosJSON & """filial"": """ & EscapeJSON(rsAnexos("Filial")) & """, "
            anexosJSON = anexosJSON & """nome"": """ & EscapeJSON(rsAnexos("Nome")) & """, "
            anexosJSON = anexosJSON & """usuario"": """ & EscapeJSON(rsAnexos("Usuario")) & """, "
            anexosJSON = anexosJSON & """caminho"": """ & EscapeJSON(caminho) & """, "
            anexosJSON = anexosJSON & """dataCriacao"": """ & EscapeJSON(rsAnexos("DataDaOperacao")) & """ "
            anexosJSON = anexosJSON & "}"
            
            primeiroItem = False
            rsAnexos.MoveNext
        Loop
    End If

    anexosJSON = anexosJSON & "]"

    rsAnexos.Close()
    SET rsAnexos = NOTHING

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """codigo"": """ & EscapeJSON(rsObter("Codigo")) & """, "
    jsonObjeto = jsonObjeto & """tipoSolicitacao"": """ & EscapeJSON(rsObter("TipoSolicitacao")) & """, "
    jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao")) & """, "
    jsonObjeto = jsonObjeto & """tipoPessoa"": """ & EscapeJSON(rsObter("TipoPessoa")) & """, "
    jsonObjeto = jsonObjeto & """cnpj"": """ & EscapeJSON(rsObter("Cnpj")) & """, "
    jsonObjeto = jsonObjeto & """cpf"": """ & EscapeJSON(rsObter("Cpf")) & """, "
    jsonObjeto = jsonObjeto & """grupoContas"": """ & EscapeJSON(rsObter("GrupoContas")) & """, "
    jsonObjeto = jsonObjeto & """codigoHfm"": """ & EscapeJSON(rsObter("CodigoHfm")) & """, "
    jsonObjeto = jsonObjeto & """codigoOriginal"": """ & EscapeJSON(rsObter("CodigoOriginal")) & """, "
    jsonObjeto = jsonObjeto & """razaoSocial"": """ & EscapeJSON(rsObter("RazaoSocial")) & """, "
    jsonObjeto = jsonObjeto & """nomeFantasia"": """ & EscapeJSON(rsObter("NomeFantasia")) & """, "
    jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
    jsonObjeto = jsonObjeto & """rua"": """ & EscapeJSON(rsObter("Rua")) & """, "
    jsonObjeto = jsonObjeto & """numero"": """ & EscapeJSON(rsObter("Numero")) & """, "
    jsonObjeto = jsonObjeto & """bairro"": """ & EscapeJSON(rsObter("Bairro")) & """, "
    jsonObjeto = jsonObjeto & """cep"": """ & EscapeJSON(rsObter("Cep")) & """, "
    jsonObjeto = jsonObjeto & """estado"": """ & EscapeJSON(rsObter("Estado")) & """, "
    jsonObjeto = jsonObjeto & """cidade"": """ & EscapeJSON(rsObter("Cidade")) & """, "
    jsonObjeto = jsonObjeto & """pais"": """ & EscapeJSON(rsObter("Pais")) & """, "
    jsonObjeto = jsonObjeto & """enderecoEntregaDiferente"": """ & EscapeJSON(rsObter("EnderecoEntregaDiferente")) & """, "
    jsonObjeto = jsonObjeto & """enderecoCobrancaDiferente"": """ & EscapeJSON(rsObter("EnderecoCobrancaDiferente")) & """, "
    jsonObjeto = jsonObjeto & """quemIncluiu"": """ & EscapeJSON(rsObter("QuemIncluiu")) & """, "
    jsonObjeto = jsonObjeto & """codigoPlanta"": """ & EscapeJSON(rsObter("CodigoPlanta")) & """, "
    jsonObjeto = jsonObjeto & """codigoLoja"": """ & EscapeJSON(rsObter("CodigoLoja")) & """, "
    jsonObjeto = jsonObjeto & """volumeVendaMensalM2"": """ & EscapeJSON(rsObter("VolumeVendaMensalM2")) & """, "
    jsonObjeto = jsonObjeto & """precoVendaNetM2"": """ & EscapeJSON(rsObter("PrecoVendaNetM2")) & """, "
    jsonObjeto = jsonObjeto & """precoVendaNetKg"": """ & EscapeJSON(rsObter("PrecoVendaNetKg")) & """, "
    jsonObjeto = jsonObjeto & """eshop"": """ & EscapeJSON(Iif(rsObter("Eshop"), 1, 0)) & """, "
    jsonObjeto = jsonObjeto & """inscricaoEstadual"": """ & EscapeJSON(rsObter("InscricaoEstadual")) & """, "
    jsonObjeto = jsonObjeto & """inscricaoMunicipal"": """ & EscapeJSON(rsObter("InscricaoMunicipal")) & """, "
    jsonObjeto = jsonObjeto & """preCadastro"": """ & EscapeJSON(rsObter("PreCadastro")) & """, "
    jsonObjeto = jsonObjeto & """limiteCredito"": """ & EscapeJSON(rsObter("LimiteCredito")) & """, "
    jsonObjeto = jsonObjeto & """kmDestino"": """ & EscapeJSON(rsObter("KmDestino")) & """, "
    jsonObjeto = jsonObjeto & """idioma"": """ & EscapeJSON(rsObter("Idioma")) & """, "
    jsonObjeto = jsonObjeto & """telefone"": """ & EscapeJSON(rsObter("Telefone")) & """, "
    jsonObjeto = jsonObjeto & """emailXml"": """ & EscapeJSON(rsObter("EmailXml")) & """, "
    jsonObjeto = jsonObjeto & """emailOrdemPagamento"": """ & EscapeJSON(rsObter("EmailOrdemPagamento")) & """, "
    jsonObjeto = jsonObjeto & """declarantePisCofins"": """ & EscapeJSON(rsObter("DeclarantePisCofins")) & """, "
    jsonObjeto = jsonObjeto & """contribuinteIcms"": """ & EscapeJSON(rsObter("ContribuinteIcms")) & """, "
    jsonObjeto = jsonObjeto & """icmsDiferido"": """ & EscapeJSON(rsObter("IcmsDiferido")) & """, "
    jsonObjeto = jsonObjeto & """percentualDiferimento"": """ & EscapeJSON(rsObter("PercentualDiferimento")) & """, "
    jsonObjeto = jsonObjeto & """isencaoIpi"": """ & EscapeJSON(rsObter("IsencaoIpi")) & """, "
    jsonObjeto = jsonObjeto & """tipoIndustria"": """ & EscapeJSON(rsObter("TipoIndustria")) & """, "
    jsonObjeto = jsonObjeto & """codigoSuframa"": """ & EscapeJSON(rsObter("CodigoSuframa")) & """, "
    jsonObjeto = jsonObjeto & """contaConciliacao"": """ & EscapeJSON(rsObter("ContaConciliacao")) & """, "
    jsonObjeto = jsonObjeto & """chaveOrdenacao"": """ & EscapeJSON(rsObter("ChaveOrdenacao")) & """, "
    jsonObjeto = jsonObjeto & """condicaoPagamento"": """ & EscapeJSON(rsObter("CondicaoPagamento")) & """, "
    jsonObjeto = jsonObjeto & """metodoPagamento"": """ & EscapeJSON(rsObter("MetodoPagamento")) & """, "
    jsonObjeto = jsonObjeto & """codigoJuros"": """ & EscapeJSON(rsObter("CodigoJuros")) & """, "
    jsonObjeto = jsonObjeto & """segmento"": """ & EscapeJSON(rsObter("Segmento")) & """, "
    jsonObjeto = jsonObjeto & """codigoSap"": """ & EscapeJSON(rsObter("CodigoSap")) & """, "
    jsonObjeto = jsonObjeto & """natureza"": """ & EscapeJSON(rsObter("Natureza")) & """, "
    jsonObjeto = jsonObjeto & """tipoCliente"": """ & EscapeJSON(rsObter("TipoCliente")) & """, "
    jsonObjeto = jsonObjeto & """tipoFrete"": """ & EscapeJSON(rsObter("TipoFrete")) & """, "
    jsonObjeto = jsonObjeto & """tes"": """ & EscapeJSON(rsObter("Tes")) & """, "
    jsonObjeto = jsonObjeto & """codigoVendedor"": """ & EscapeJSON(rsObter("CodigoVendedor")) & """, "
    jsonObjeto = jsonObjeto & """codigoProtheus"": """ & EscapeJSON(rsObter("CodigoProtheus")) & """, "
    jsonObjeto = jsonObjeto & """diasViagem"": """ & EscapeJSON(rsObter("DiasViagem")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaDiasMaior"": """ & EscapeJSON(rsObter("ToleranciaDiasMaior")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaDiasMenor"": """ & EscapeJSON(rsObter("ToleranciaDiasMenor")) & """, "
    jsonObjeto = jsonObjeto & """ruaEntrega"": """ & EscapeJSON(rsObter("RuaEntrega")) & """, "
    jsonObjeto = jsonObjeto & """numeroEntrega"": """ & EscapeJSON(rsObter("NumeroEntrega")) & """, "
    jsonObjeto = jsonObjeto & """bairroEntrega"": """ & EscapeJSON(rsObter("BairroEntrega")) & """, "
    jsonObjeto = jsonObjeto & """cepEntrega"": """ & EscapeJSON(rsObter("CepEntrega")) & """, "
    jsonObjeto = jsonObjeto & """estadoEntrega"": """ & EscapeJSON(rsObter("EstadoEntrega")) & """, "
    jsonObjeto = jsonObjeto & """cidadeEntrega"": """ & EscapeJSON(rsObter("CidadeEntrega")) & """, "
    jsonObjeto = jsonObjeto & """ruaCobranca"": """ & EscapeJSON(rsObter("RuaCobranca")) & """, "
    jsonObjeto = jsonObjeto & """numeroCobranca"": """ & EscapeJSON(rsObter("NumeroCobranca")) & """, "
    jsonObjeto = jsonObjeto & """bairroCobranca"": """ & EscapeJSON(rsObter("BairroCobranca")) & """, "
    jsonObjeto = jsonObjeto & """cepCobranca"": """ & EscapeJSON(rsObter("CepCobranca")) & """, "
    jsonObjeto = jsonObjeto & """estadoCobranca"": """ & EscapeJSON(rsObter("EstadoCobranca")) & """, "
    jsonObjeto = jsonObjeto & """cidadeCobranca"": """ & EscapeJSON(rsObter("CidadeCobranca")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaQuantidadeMaior"": """ & EscapeJSON(rsObter("ToleranciaQuantidadeMaior")) & """, "
    jsonObjeto = jsonObjeto & """toleranciaQuantidadeMenor"": """ & EscapeJSON(rsObter("ToleranciaQuantidadeMenor")) & """, "
    jsonObjeto = jsonObjeto & """observacaoEmissor"": """ & EscapeJSON(rsObter("ObservacaoEmissor")) & """, "
    jsonObjeto = jsonObjeto & """observacaoCdag"": """ & EscapeJSON(rsObter("ObservacaoCdag")) & """, "
    jsonObjeto = jsonObjeto & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObter("NomeUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """loginUsuarioInclusao"": """ & EscapeJSON(rsObter("LoginUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """leadTime"": """ & EscapeJSON(rsObter("LeadTime")) & """, "
    jsonObjeto = jsonObjeto & """anexos"": " & anexosJSON & ""
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorCodigo = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("codigoHfm") = Request("codigoHfm")
    Dados("tipoSolicitacao") = Request("tipoSolicitacao")
    Dados("eshop") = Request("eshop")
    Dados("grupoContas") = Request("grupoContas")
    Dados("tipoPessoa") = Request("tipoPessoa")
    Dados("razaoSocial") = Request("razaoSocial")
    Dados("nomeFantasia") = Request("nomeFantasia")
    Dados("cnpj") = Request("cnpj")
    Dados("cpf") = Request("cpf")
    Dados("inscricaoEstadual") = Request("inscricaoEstadual")
    Dados("inscricaoMunicipal") = Request("inscricaoMunicipal")
    Dados("rua") = Request("rua")
    Dados("numero") = Request("numero")
    Dados("bairro") = Request("bairro")
    Dados("cep") = Request("cep")
    Dados("cidade") = Request("cidade")
    Dados("estado") = Request("estado")
    Dados("pais") = Request("pais")
    Dados("kmDestino") = Request("kmDestino")
    Dados("idioma") = Request("idioma")
    Dados("telefone") = Request("telefone")
    Dados("emailXml") = Request("emailXml")
    Dados("emailOrdemPagamento") = Request("emailOrdemPagamento")
    Dados("declarantePisCofins") = Request("declarantePisCofins")
    Dados("contribuinteIcms") = Request("contribuinteIcms")
    Dados("icmsDiferido") = Request("icmsDiferido")
    Dados("percentualDiferimento") = Request("percentualDiferimento")
    Dados("isencaoIpi") = Request("isencaoIpi")
    Dados("tipoIndustria") = Request("tipoIndustria")
    Dados("codigoSuframa") = Request("codigoSuframa")
    Dados("contaConciliacao") = Request("contaConciliacao")
    Dados("chaveOrdenacao") = Request("chaveOrdenacao")
    Dados("condicaoPagamento") = Request("condicaoPagamento")
    Dados("metodoPagamento") = Request("metodoPagamento")
    Dados("codigoJuros") = Request("codigoJuros")
    Dados("limiteCredito") = Request("limiteCredito")
    Dados("segmento") = Request("segmento")
    Dados("volumeVendaMensalM2") = Request("volumeVendaMensalM2")
    Dados("precoVendaNetM2") = Request("precoVendaNetM2")
    Dados("precoVendaNetKg") = Request("precoVendaNetKg")
    Dados("codigoSap") = Request("codigoSap")
    Dados("natureza") = Request("natureza")
    Dados("tipoCliente") = Request("tipoCliente")
    Dados("tipoFrete") = Request("tipoFrete")
    Dados("tes") = Request("tes")
    Dados("codigoVendedor") = Request("codigoVendedor")
    Dados("diasViagem") = Request("diasViagem")
    Dados("toleranciaDiasMaior") = Request("toleranciaDiasMaior")
    Dados("toleranciaDiasMenor") = Request("toleranciaDiasMenor")
    Dados("enderecoEntregaDiferente") = Request("enderecoEntregaDiferente")
    Dados("ruaEntrega") = Request("ruaEntrega")
    Dados("numeroEntrega") = Request("numeroEntrega")
    Dados("bairroEntrega") = Request("bairroEntrega")
    Dados("cepEntrega") = Request("cepEntrega")
    Dados("estadoEntrega") = Request("estadoEntrega")
    Dados("cidadeEntrega") = Request("cidadeEntrega")
    Dados("kmEntrega") = Request("kmEntrega")
    Dados("enderecoCobrancaDiferente") = Request("enderecoCobrancaDiferente")
    Dados("ruaCobranca") = Request("ruaCobranca")
    Dados("numeroCobranca") = Request("numeroCobranca")
    Dados("bairroCobranca") = Request("bairroCobranca")
    Dados("cepCobranca") = Request("cepCobranca")
    Dados("estadoCobranca") = Request("estadoCobranca")
    Dados("cidadeCobranca") = Request("cidadeCobranca")
    Dados("toleranciaQuantidadeMaior") = Request("toleranciaQuantidadeMaior")
    Dados("toleranciaQuantidadeMenor") = Request("toleranciaQuantidadeMenor")
    Dados("codigoProtheus") = Request("codigoProtheus")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("observacaoEmissor") = Request("observacaoEmissor")
    Dados("observacaoCdag") = Request("observacaoCdag")
    Dados("anexosJSON") = Request("anexos")

    jaExiste = VerificarCpfCnpjExiste(Iif(Dados("tipoPessoa") = "F", Request("cpf"), Request("cnpj")))
    If jaExiste Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Já existe um cliente com esse CNPJ""}"
        Exit Function
    End If

    correspondencia = ObterCorrespondencia(Dados("codigoPlanta"))
    codigo = GerarCodigoCadastro(codigoPlanta)
    If codigo = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao obter próximo id""}"
        Exit Function
    End If
    
    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "INSERT INTO UnificadoCliente ("
    SQL = SQL & "Codigo, "
    SQL = SQL & "GrupoContas, "
    SQL = SQL & "Empresa, "
    SQL = SQL & "RazaoSocial, "
    SQL = SQL & "NomeFantasia, "
    SQL = SQL & "Rua, "
    SQL = SQL & "Numero, "
    SQL = SQL & "Bairro, "
    SQL = SQL & "Cep, "
    SQL = SQL & "Cidade, "
    SQL = SQL & "Estado, "
    SQL = SQL & "Pais, "
    SQL = SQL & "Idioma, "
    SQL = SQL & "Telefone, "
    SQL = SQL & "Fax, "
    SQL = SQL & "EmailXml, "
    SQL = SQL & "EmailPagamento, "
    SQL = SQL & "CodigoVendedor, "
    SQL = SQL & "CodigoHfm, "
    SQL = SQL & "Cpf, "
    SQL = SQL & "Cnpj, "
    SQL = SQL & "InscricaoEstadual, "
    SQL = SQL & "InscricaoMunicipal, "
    SQL = SQL & "Declarante, "
    SQL = SQL & "Contribuinte, "
    SQL = SQL & "NaturezaLegal, "
    SQL = SQL & "TipoIndustria, "
    SQL = SQL & "CodigoSuframa, "
    SQL = SQL & "ContaConciliacao, "
    SQL = SQL & "ContaBanco, "
    SQL = SQL & "ChaveOrdenacao, "
    SQL = SQL & "CondicaoPagamento, "
    SQL = SQL & "MetodoPagamento, "
    SQL = SQL & "CodigoJuros, "
    SQL = SQL & "CodigoAntigo, "
    SQL = SQL & "Correspondencia, "
    SQL = SQL & "LimiteCredito, "
    SQL = SQL & "VolumeVendam2, "
    SQL = SQL & "PrecoEstm2, "
    SQL = SQL & "PrecoEstTon, "
    SQL = SQL & "LimiteCredSug, "
    SQL = SQL & "PrazoPag, "
    SQL = SQL & "DtVenda, "
    SQL = SQL & "Seguimento, "
    SQL = SQL & "ResponsavelAprovacao, "
    SQL = SQL & "Natureza, "
    SQL = SQL & "Tipo, "
    SQL = SQL & "TipoFrete, "
    SQL = SQL & "CodigoVendedorPro, "
    SQL = SQL & "CodigoClasProd, "
    SQL = SQL & "DiasViagem, "
    SQL = SQL & "DiasEntregaMaior, "
    SQL = SQL & "DiasEntregaMenor, "
    SQL = SQL & "RuaEntPro, "
    SQL = SQL & "NumeroEntPro, "
    SQL = SQL & "BairroEntPro, "
    SQL = SQL & "CepEntPro, "
    SQL = SQL & "CidadeEntPro, "
    SQL = SQL & "EstadoEntPro, "
    SQL = SQL & "RuaCobPro, "
    SQL = SQL & "NumeroCobPro, "
    SQL = SQL & "BairroCobPro, "
    SQL = SQL & "CepCobPro, "
    SQL = SQL & "CidadeCobPro, "
    SQL = SQL & "EstadoCobPro, "
    SQL = SQL & "ContatoCompras, "
    SQL = SQL & "Departamento, "
    SQL = SQL & "TelefoneSia, "
    SQL = SQL & "Email, "
    SQL = SQL & "PrecoPadrao, "
    SQL = SQL & "TransacaoPedidos, "
    SQL = SQL & "AtividadePincipal, "
    SQL = SQL & "GrupoEmpresa, "
    SQL = SQL & "AreaVendas, "
    SQL = SQL & "NomeRepresentante, "
    SQL = SQL & "EmailRepresentante, "
    SQL = SQL & "TelefoneRep, "
    SQL = SQL & "SegmentacaoMercado, "
    SQL = SQL & "AvisaEmbarque, "
    SQL = SQL & "ContEmbalagens, "
    SQL = SQL & "RegraEmbarque, "
    SQL = SQL & "DiasTransporte, "
    SQL = SQL & "TolEntregaMenos, "
    SQL = SQL & "TolEntregaMais, "
    SQL = SQL & "JanTolMenos, "
    SQL = SQL & "JanTolMais, "
    SQL = SQL & "EmiteCertificado, "
    SQL = SQL & "Frete, "
    SQL = SQL & "DesprezarTarifa, "
    SQL = SQL & "DistRomaneio, "
    SQL = SQL & "Del, "
    SQL = SQL & "QuemIncluiu, "
    SQL = SQL & "DataEmissao, "
    SQL = SQL & "HoraEmissao, "
    SQL = SQL & "StatusCli, "
    SQL = SQL & "QuemIncluiuCod, "
    SQL = SQL & "ObsEmissor, "
    SQL = SQL & "ObsCedag, "
    SQL = SQL & "CodSap, "
    SQL = SQL & "CodPro, "
    SQL = SQL & "CodSia, "
    SQL = SQL & "Filial, "
    SQL = SQL & "Tes, "
    SQL = SQL & "Eshop, "
    SQL = SQL & "TQtdeMaior, "
    SQL = SQL & "TQtdeMenor, "
    SQL = SQL & "PreCadastro, "
    SQL = SQL & "CodLoja, "
    SQL = SQL & "KmEndereco, "
    SQL = SQL & "KmEnderecoCobranca, "
    SQL = SQL & "isencaoIPI, "
    SQL = SQL & "IcmsDiferido, "
    SQL = SQL & "percentualDiferimento "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & codigo & "', "
    SQL = SQL & "'" & Dados("grupoContas") & "', "
    SQL = SQL & "'" & correspondencia & "', "
    SQL = SQL & "'" & ucase(Dados("razaoSocial")) & "', "
    SQL = SQL & "'" & Left(Dados("nomeFantasia"), 20) & "', "
    SQL = SQL & "'" & ucase(Dados("rua")) & "', "
    SQL = SQL & "'" & Dados("numero") & "', "
    SQL = SQL & "'" & ucase(Dados("bairro")) & "', "
    SQL = SQL & "'" & Dados("cep") & "', "
    SQL = SQL & "'" & ucase(Trim(Dados("cidade"))) & "', "
    SQL = SQL & "'" & Dados("estado") & "', "
    SQL = SQL & "'" & Dados("pais") & "', "
    SQL = SQL & "'PT-BR', "
    SQL = SQL & "'" & Dados("telefone") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("emailXml") & "', "
    SQL = SQL & "'" & Dados("emailOrdemPagamento") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("codigoHfm") & "', "
    SQL = SQL & "'" & Dados("cpf") & "', "
    SQL = SQL & "'" & Dados("cnpj") & "', "
    SQL = SQL & "'" & Dados("inscricaoEstadual") & "', "
    SQL = SQL & "'" & Dados("inscricaoMunicipal") & "', "
    SQL = SQL & "'" & Dados("declarantePisCofins") & "', "
    SQL = SQL & "'" & Dados("contribuinteIcms") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("tipoIndustria") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("contaConciliacao")  & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("chaveOrdenacao") & "', "
    SQL = SQL & "'" & Dados("condicaoPagamento") & "', "
    SQL = SQL & "'" & Dados("metodoPagamento") & "', "
    SQL = SQL & "'" & Dados("codigoJuros") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & correspondencia & "', "
    SQL = SQL & "'" & Dados("limiteCredito") & "', "
    SQL = SQL & "'" & Dados("volumeVendaMensalM2") & "', "
    SQL = SQL & "'" & Dados("precoVendaNetM2") & "', "
    SQL = SQL & "'" & Dados("precoVendaNetKg") & "', "
    SQL = SQL & "'" & Dados("limiteCredSug") & "', " 'TODO: Verificar 
    SQL = SQL & "'" & Dados("prazoPag") & "', "
    SQL = SQL & "'" & Dados("DtVenda") & "', " 'TODO: Verificar 
    SQL = SQL & "'" & Dados("segmento") & "', "
    SQL = SQL & "'" & Dados("responsavelAprovacao") & "', "
    SQL = SQL & "'" & Dados("natureza") & "', "
    SQL = SQL & "'" & Dados("tipoCliente") & "', "
    SQL = SQL & "'" & Dados("tipoFrete") & "', "
    SQL = SQL & "'" & Dados("codigoVendedor") & "', " 
    SQL = SQL & "'" & Dados("codigoClasProd") & "', " 'TODO: Verificar 
    SQL = SQL & "'" & Dados("diasViagem") & "', "
    SQL = SQL & "'" & Dados("toleranciaDiasMaior") & "', "
    SQL = SQL & "'" & Dados("toleranciaDiasMenor") & "', "
    SQL = SQL & "'" & Dados("ruaEntrega") & "', "
    SQL = SQL & "'" & Dados("numeroEntrega") & "', "
    SQL = SQL & "'" & Dados("bairroEntrega") & "', "
    SQL = SQL & "'" & Dados("cepEntrega") & "', "
    SQL = SQL & "'" & Dados("cidadeEntrega") & "', "
    SQL = SQL & "'" & Dados("estadoEntrega") & "', "
    SQL = SQL & "'" & Dados("ruaCobranca") & "', "
    SQL = SQL & "'" & Dados("numeroCobranca") & "', "
    SQL = SQL & "'" & Dados("bairroCobranca") & "', "
    SQL = SQL & "'" & Dados("cepCobranca") & "', "
    SQL = SQL & "'" & Dados("cidadeCobranca") & "', "
    SQL = SQL & "'" & Dados("estadoCobranca") & "', "
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'', " 
    SQL = SQL & "'0', " 'DEL
    SQL = SQL & "'" & usuarioLogado("login") & "', "
    SQL = SQL & "CONVERT(char(8), GETDATE(), 112), "
    SQL = SQL & "CONVERT(char(8), GETDATE(), 108), "
    SQL = SQL & "'P', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "'" & Dados("observacaoEmissor") & "', "
    SQL = SQL & "'" & Dados("observacaoCdag") & "', "
    SQL = SQL & "'" & Dados("codigoSap") & "', "
    SQL = SQL & "'" & Dados("codigoProtheus") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Dados("codigoPlanta") & "', "
    SQL = SQL & "'" & Dados("tes") & "', "
    SQL = SQL & "'" & Dados("eshop") & "', "
    SQL = SQL & "" & Iif(Dados("toleranciaQuantidadeMaior") = "", "NULL", Dados("toleranciaQuantidadeMaior")) & ", "
    SQL = SQL & "" & Iif(Dados("toleranciaQuantidadeMaior") = "", "NULL", Dados("toleranciaQuantidadeMaior")) & ", "
    SQL = SQL & "'" & iif(Dados("tipoSolicitacao") = "5", "P", "") & "', "
    SQL = SQL & "'" & Dados("codigoLoja") & "', " 
    SQL = SQL & "'" & Dados("kmDestino") & "', " 
    SQL = SQL & "'" & Dados("kmEntrega") & "', " 
    SQL = SQL & "'" & Dados("isencaoIpi") & "', "
    SQL = SQL & "'" & Dados("icmsDiferido") & "', " 
    SQL = SQL & ""  & iif(Dados("percentualDiferimento") = "", 0, Dados("percentualDiferimento")) 
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
        valoresAnexos = valoresAnexos & "'" & codigo & "', "
        valoresAnexos = valoresAnexos & "'" & Dados("codigoPlanta") & "', "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO UnificadoClienteDocumento ("
        SQL = SQL & "   idUnificadoCliente, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresAnexos & "; "
    End If
    
    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL &   "ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL &   "COMMIT TRANSACTION "

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar Cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("criacao", codigo)
    Criar = "{""sucesso"": true, ""mensagem"": ""Cadastro criado com sucesso!"", ""codigo"": " & codigo & "}"
End Function

Function Editar()   
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("flagUnificarCliente") = Request("flagUnificarCliente")
    Dados("codigo") = Request("codigo")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("codigoHfm") = Request("codigoHfm")
    Dados("tipoSolicitacao") = Request("tipoSolicitacao")
    Dados("eshop") = Request("eshop")
    Dados("grupoContas") = Request("grupoContas")
    Dados("tipoPessoa") = Request("tipoPessoa")
    Dados("razaoSocial") = Request("razaoSocial")
    Dados("nomeFantasia") = Request("nomeFantasia")
    Dados("cnpj") = Request("cnpj")
    Dados("cpf") = Request("cpf")
    Dados("inscricaoEstadual") = Request("inscricaoEstadual")
    Dados("inscricaoMunicipal") = Request("inscricaoMunicipal")
    Dados("rua") = Request("rua")
    Dados("numero") = Request("numero")
    Dados("bairro") = Request("bairro")
    Dados("cep") = Request("cep")
    Dados("cidade") = Request("cidade")
    Dados("estado") = Request("estado")
    Dados("pais") = Request("pais")
    Dados("kmDestino") = Request("kmDestino")
    Dados("idioma") = Request("idioma")
    Dados("telefone") = Request("telefone")
    Dados("emailXml") = Request("emailXml")
    Dados("emailOrdemPagamento") = Request("emailOrdemPagamento")
    Dados("declarantePisCofins") = Request("declarantePisCofins")
    Dados("contribuinteIcms") = Request("contribuinteIcms")
    Dados("icmsDiferido") = Request("icmsDiferido")
    Dados("percentualDiferimento") = Request("percentualDiferimento")
    Dados("isencaoIpi") = Request("isencaoIpi")
    Dados("tipoIndustria") = Request("tipoIndustria")
    Dados("codigoSuframa") = Request("codigoSuframa")
    Dados("contaConciliacao") = Request("contaConciliacao")
    Dados("chaveOrdenacao") = Request("chaveOrdenacao")
    Dados("condicaoPagamento") = Request("condicaoPagamento")
    Dados("metodoPagamento") = Request("metodoPagamento")
    Dados("codigoJuros") = Request("codigoJuros")
    Dados("limiteCredito") = Request("limiteCredito")
    Dados("segmento") = Request("segmento")
    Dados("volumeVendaMensalM2") = Request("volumeVendaMensalM2")
    Dados("precoVendaNetM2") = Request("precoVendaNetM2")
    Dados("precoVendaNetKg") = Request("precoVendaNetKg")
    Dados("codigoSap") = Request("codigoSap")
    Dados("natureza") = Request("natureza")
    Dados("tipoCliente") = Request("tipoCliente")
    Dados("tipoFrete") = Request("tipoFrete")
    Dados("tes") = Request("tes")
    Dados("codigoVendedor") = Request("codigoVendedor")
    Dados("diasViagem") = Request("diasViagem")
    Dados("toleranciaDiasMaior") = Request("toleranciaDiasMaior")
    Dados("toleranciaDiasMenor") = Request("toleranciaDiasMenor")
    Dados("enderecoEntregaDiferente") = Request("enderecoEntregaDiferente")
    Dados("ruaEntrega") = Request("ruaEntrega")
    Dados("numeroEntrega") = Request("numeroEntrega")
    Dados("bairroEntrega") = Request("bairroEntrega")
    Dados("cepEntrega") = Request("cepEntrega")
    Dados("estadoEntrega") = Request("estadoEntrega")
    Dados("cidadeEntrega") = Request("cidadeEntrega")
    Dados("kmEntrega") = Request("kmEntrega")
    Dados("enderecoCobrancaDiferente") = Request("enderecoCobrancaDiferente")
    Dados("ruaCobranca") = Request("ruaCobranca")
    Dados("numeroCobranca") = Request("numeroCobranca")
    Dados("bairroCobranca") = Request("bairroCobranca")
    Dados("cepCobranca") = Request("cepCobranca")
    Dados("estadoCobranca") = Request("estadoCobranca")
    Dados("cidadeCobranca") = Request("cidadeCobranca")
    Dados("toleranciaQuantidadeMaior") = Request("toleranciaQuantidadeMaior")
    Dados("toleranciaQuantidadeMenor") = Request("toleranciaQuantidadeMenor")
    Dados("codigoProtheus") = Request("codigoProtheus")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("observacaoEmissor") = Request("observacaoEmissor")
    Dados("observacaoCdag") = Request("observacaoCdag")
    Dados("anexosJSON") = Request("anexos")

    correspondencia = ObterCorrespondencia(Dados("codigoPlanta"))

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE UnificadoCliente SET "
    SQL = SQL & "   GrupoContas = '" & Dados("grupoContas") & "', "
    SQL = SQL & "   Empresa = '" & correspondencia & "', "
    SQL = SQL & "   Filial = '" & Dados("codigoPlanta") & "', "
    SQL = SQL & "   RazaoSocial = '" & Dados("razaoSocial") & "', "
    SQL = SQL & "   NomeFantasia = '" & Dados("nomeFantasia") & "', "
    SQL = SQL & "   Rua = '" & Dados("rua") & "', "
    SQL = SQL & "   Numero = '" & Dados("numero") & "', "
    SQL = SQL & "   Bairro = '" & Dados("bairro") & "', "
    SQL = SQL & "   Cep = '" & Dados("cep") & "', "
    SQL = SQL & "   Cidade = '" & Dados("cidade") & "', "
    SQL = SQL & "   Estado = '" & Dados("estado") & "', "
    SQL = SQL & "   Pais = '" & Dados("pais") & "', "
    SQL = SQL & "   Telefone = '" & Dados("telefone") & "', "
    SQL = SQL & "   EmailXml = '" & Dados("emailXml") & "', "
    SQL = SQL & "   EmailPagamento = '" & Dados("emailOrdemPagamento") & "', "
    SQL = SQL & "   CodigoHfm = '" & Dados("codigoHfm") & "', "
    SQL = SQL & "   Cpf = '" & Dados("cpf") & "', "
    SQL = SQL & "   Cnpj = '" & Dados("cnpj") & "', "
    SQL = SQL & "   InscricaoEstadual = '" & Dados("inscricaoEstadual") & "', "
    SQL = SQL & "   InscricaoMunicipal = '" & Dados("inscricaoMunicipal") & "', "
    SQL = SQL & "   Declarante = '" & Dados("declarantePisCofins") & "', "
    SQL = SQL & "   Contribuinte = '" & Dados("contribuinteIcms") & "', "
    SQL = SQL & "   TipoIndustria = '" & Dados("tipoIndustria") & "', "
    SQL = SQL & "   CodigoSuframa = '" & Dados("codigoSuframa") & "', "
    SQL = SQL & "   ContaConciliacao = '" & Dados("contaConciliacao") & "', "
    SQL = SQL & "   ChaveOrdenacao = '" & Dados("chaveOrdenacao") & "', "
    SQL = SQL & "   CondicaoPagamento = '" & Dados("condicaoPagamento") & "', "
    SQL = SQL & "   MetodoPagamento = '" & Dados("metodoPagamento") & "', "
    SQL = SQL & "   CodigoJuros = '" & Dados("codigoJuros") & "', "
    SQL = SQL & "   Correspondencia = '" & correspondencia & "', "
    SQL = SQL & "   LimiteCredito = '" & Dados("limiteCredito") & "', "
    SQL = SQL & "   VolumeVendam2 = '" & Dados("volumeVendaMensalM2") & "', "
    SQL = SQL & "   PrecoEstm2 = '" & Dados("precoVendaNetM2") & "', "
    SQL = SQL & "   PrecoEstTon = '" & Dados("precoVendaNetKg") & "', "
    SQL = SQL & "   Seguimento = '" & Dados("segmento") & "', "
    SQL = SQL & "   Natureza = '" & Dados("natureza") & "', "
    SQL = SQL & "   Tipo = '" & Dados("tipoCliente") & "', "
    SQL = SQL & "   TipoFrete = '" & Dados("tipoFrete") & "', "
    SQL = SQL & "   CodigoVendedorPro = '" & Dados("codigoVendedor") & "', "
    SQL = SQL & "   DiasViagem = '" & Dados("diasViagem") & "', "
    SQL = SQL & "   DiasEntregaMaior = " & Dados("toleranciaDiasMaior") & ", "
    SQL = SQL & "   DiasEntregaMenor = " & Dados("toleranciaDiasMenor") & ", "
    SQL = SQL & "   RuaEntPro = '" & Dados("ruaEntrega") & "', "
    SQL = SQL & "   NumeroEntPro = '" & Dados("numeroEntrega") & "', "
    SQL = SQL & "   BairroEntPro = '" & Dados("bairroEntrega") & "', "
    SQL = SQL & "   CepEntPro = '" & Dados("cepEntrega") & "', "
    SQL = SQL & "   CidadeEntPro = '" & Dados("cidadeEntrega") & "', "
    SQL = SQL & "   EstadoEntPro = '" & Dados("estadoEntrega") & "', "
    SQL = SQL & "   RuaCobPro = '" & Dados("ruaCobranca") & "', "
    SQL = SQL & "   NumeroCobPro = '" & Dados("numeroCobranca") & "', "
    SQL = SQL & "   BairroCobPro = '" & Dados("bairroCobranca") & "', "
    SQL = SQL & "   CepCobPro = '" & Dados("cepCobranca") & "', "
    SQL = SQL & "   CidadeCobPro = '" & Dados("cidadeCobranca") & "', "
    SQL = SQL & "   EstadoCobPro = '" & Dados("estadoCobranca") & "', "
    SQL = SQL & "   TolEntregaMenos = '" & Dados("toleranciaDiasMenor") & "', "
    SQL = SQL & "   TolEntregaMais = '" & Dados("toleranciaDiasMaior") & "', "
    SQL = SQL & "   ObsEmissor = '" & Dados("ObsEmissor") & "', "
    SQL = SQL & "   ObsCedag = '" & Dados("ObsCedag") & "', "
    SQL = SQL & "   CodSap = '" & Dados("codigoSap") & "', "
    SQL = SQL & "   CodPro = '" & Dados("codigoProtheus") & "', "
    SQL = SQL & "   Tes = '" & Dados("tes") & "', "
    SQL = SQL & "   Eshop = '" & Dados("eshop") & "', "
    SQL = SQL & "   TQtdeMaior = " & Iif(Dados("toleranciaQuantidadeMaior") = "", "NULL", Dados("toleranciaQuantidadeMaior")) & ", "
    SQL = SQL & "   TQtdeMenor = " & Iif(Dados("toleranciaQuantidadeMenor") = "", "NULL", Dados("toleranciaQuantidadeMenor")) & ", "

    If Dados("flagUnificarCliente") = "1" Then
        'Se marcado para unificar atualizar status de pre-cadastro
        SQL = SQL & "   PreCadastro = 'C'," 
    End If

    SQL = SQL & "   CodLoja = '" & Dados("codigoLoja") & "', "
    SQL = SQL & "   KmEndereco = '" & Dados("kmDestino") & "', "
    SQL = SQL & "   isencaoIPI = " & Dados("isencaoIpi") & ", " 
    SQL = SQL & "   IcmsDiferido = '" & Dados("icmsDiferido") & "', " 
    SQL = SQL & "   percentualDiferimento = " & iif(Dados("percentualDiferimento") = "", 0, Dados("percentualDiferimento"))  & " " 
    SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; "

    SQL = SQL & MontarSQLAnexosEdicao(Dados("codigo"), Dados("codigoPlanta"), Dados("anexosJSON"))
    
    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsEditar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar Cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("edicao",  Dados("codigo"))
    Editar = "{""sucesso"": true, ""mensagem"": ""Cadastro atualizado com sucesso!""}"
End Function

Function MontarSQLAnexosEdicao(codigo, codigoPlanta, anexosJSON)
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
                    SQLAnexos = SQLAnexos & MontarSQLInserirAnexo(codigo, codigoPlanta, nome, caminho)
                End If
            Case "removed"
                If originalId <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInativarAnexo(codigo, originalId)
                End If
            Case "replaced"
                If originalId <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInativarAnexo(codigo, originalId)
                End If
                If nome <> "" And caminho <> "" Then
                    SQLAnexos = SQLAnexos & MontarSQLInserirAnexo(codigo, codigoPlanta, nome, caminho)
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

Function MontarSQLInserirAnexo(codigo, codigoPlanta, nome, caminho)
    SQLInsert = ""
    SQLInsert = SQLInsert & "INSERT INTO UnificadoClienteDocumento ("
    SQLInsert = SQLInsert & "   idUnificadoCliente, "
    SQLInsert = SQLInsert & "   Filial, "
    SQLInsert = SQLInsert & "   Nome, "
    SQLInsert = SQLInsert & "   Caminho, "
    SQLInsert = SQLInsert & "   Usuario, "
    SQLInsert = SQLInsert & "   DataDaOperacao "
    SQLInsert = SQLInsert & ") VALUES ("
    SQLInsert = SQLInsert & "   '" & codigo & "', "
    SQLInsert = SQLInsert & "   '" & codigoPlanta & "', "
    SQLInsert = SQLInsert & "   '" & nome & "', "
    SQLInsert = SQLInsert & "   '" & caminho & "', "
    SQLInsert = SQLInsert & "   '" & usuarioLogado("login") & "', "
    SQLInsert = SQLInsert & "   GETDATE() "
    SQLInsert = SQLInsert & "); "

    MontarSQLInserirAnexo = SQLInsert
End Function

Function MontarSQLInativarAnexo(codigo, idAnexo)
    SQLUpdate = ""
    SQLUpdate = SQLUpdate & "DELETE FROM UnificadoClienteDocumento  "
    SQLUpdate = SQLUpdate & "WHERE idUnificadoClienteDocumento = '" & idAnexo & "' "
    SQLUpdate = SQLUpdate & "  AND idUnificadoCliente = '" & codigo & "' "

    MontarSQLInativarAnexo = SQLUpdate
End Function

Function EnviarParaFiscal()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("observacao") = Request("observacao")

    If Dados("codigo") = "" Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Nenhum codigo enviado""}"
        Exit Function
    End IF

    novoStatus = "S"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE UnificadoCliente SET "
    SQL = SQL & "   StatusCli = '" & novoStatus & "' "
    SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; "

    SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
    SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
    SQL = SQL & "   '" & Dados("codigo") & "', "
    SQL = SQL & "   '" & Dados("observacao") & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
    SQL = SQL & "   '" & usuarioLogado("login") & "'"
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaFiscal = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCadastroParaNivel(Dados, "FISCAL", "Fiscal")
    Call RegistrarLog("envio_fiscal",  Dados("codigo"))
    EnviarParaFiscal = "{""sucesso"": true, ""mensagem"": ""Cadastro enviado com sucesso!""}"
End Function

Function EnviarParaControladoria()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("declarantePisCofins") = Request("declarantePisCofins")
    Dados("contribuinteIcms") = Request("contribuinteIcms")
    Dados("icmsDiferido") = Request("icmsDiferido")
    Dados("percentualDiferimento") = Request("percentualDiferimento")
    Dados("isencaoIpi") = Request("isencaoIpi")
    Dados("tipoCliente") = Request("tipoCliente")
    Dados("tes") = Request("tes")
    Dados("observacao") = Request("observacao")

    If Dados("codigo") = "" Then
        EnviarParaControladoria = "{""sucesso"": false, ""mensagem"": ""Nenhum codigo enviado""}"
        Exit Function
    End IF

    novoStatus = "C"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE UnificadoCliente SET "
    SQL = SQL & "   StatusCli = '" & novoStatus & "', "
    SQL = SQL & "   Declarante = '" & Dados("declarantePisCofins") & "', "
    SQL = SQL & "   Contribuinte = '" & Dados("contribuinteIcms") & "', "
    SQL = SQL & "   IcmsDiferido = '" & Dados("icmsDiferido") & "', "
    SQL = SQL & "   percentualDiferimento = '" & Dados("percentualDiferimento") & "', "
    SQL = SQL & "   isencaoIPI = '" & Dados("isencaoIpi") & "', "
    SQL = SQL & "   Tipo = '" & Dados("tipoCliente") & "', "
    SQL = SQL & "   Tes = '" & Dados("tes") & "' "
    SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; "

    SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
    SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
    SQL = SQL & "   '" & Dados("codigo") & "', "
    SQL = SQL & "   '" & Dados("observacao") & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
    SQL = SQL & "   '" & usuarioLogado("login") & "'"
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaControladoria = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCadastroParaNivel(Dados, "CONTROLADORIA", "Controladoria")
    Call RegistrarLog("envio_controladoria", Dados("codigo"))
    EnviarParaControladoria = "{""sucesso"": true, ""mensagem"": ""Cadastro enviado com sucesso!""}"
End Function

Function EnviarParaCDAG()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("observacao") = Request("observacao")

    If Dados("codigo") = "" Then
        EnviarParaCDAG = "{""sucesso"": false, ""mensagem"": ""Nenhum codigo enviado""}"
        Exit Function
    End IF

    novoStatus = "E"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE UnificadoCliente SET "
    SQL = SQL & "   StatusCli = '" & novoStatus & "' "
    SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; "

    SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
    SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
    SQL = SQL & "   '" & Dados("codigo") & "', "
    SQL = SQL & "   '" & Dados("observacao") & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
    SQL = SQL & "   '" & usuarioLogado("login") & "'"
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaCDAG = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCadastroParaNivel(Dados, "CDAG", "CDAG")
    Call RegistrarLog("envio_cdag", Dados("codigo"))
    EnviarParaCDAG = "{""sucesso"": true, ""mensagem"": ""Cadastro enviado com sucesso!""}"
End Function

Function EnviarParaControlesInternos()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("observacao") = Request("observacao")

    If Dados("codigo") = "" Then
        EnviarParaControlesInternos = "{""sucesso"": false, ""mensagem"": ""Nenhum codigo enviado""}"
        Exit Function
    End IF

    novoStatus = "I"

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE UnificadoCliente SET "
    SQL = SQL & "   StatusCli = '" & novoStatus & "' "
    SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; "

    SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
    SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
    SQL = SQL & "   '" & Dados("codigo") & "', "
    SQL = SQL & "   '" & Dados("observacao") & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
    SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
    SQL = SQL & "   '" & usuarioLogado("login") & "'"
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        EnviarParaControlesInternos = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCadastroParaNivel(Dados, "CONTROLES_INTERNOS", "Controles Internos")
    Call RegistrarLog("envio_controles_internos", Dados("codigo"))
    EnviarParaControlesInternos = "{""sucesso"": true, ""mensagem"": ""Cadastro enviado com sucesso!""}"
End Function

Function Finalizar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("codigoProtheus") = Request("codigoProtheus")
    Dados("codigoSap") = Request("codigoSap")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("limiteCredito") = Request("limiteCredito")
    Dados("volumeVendaMensalM2") = Request("volumeVendaMensalM2")
    Dados("precoVendaNetM2") = Request("precoVendaNetM2")
    Dados("precoVendaNetKg") = Request("precoVendaNetKg")
    Dados("razaoSocial") = Request("razaoSocial")
    Dados("cpfCnpj") = Request("cpfCnpj")
    Dados("observacao") = Request("observacao")

    If Dados("codigo") = "" Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Nenhum codigo enviado""}"
        Exit Function
    End IF

    nomeColuna = Iif(Dados("codigoPlanta") = "00" OR Dados("codigoPlanta") = "04" OR Dados("codigoPlanta") = "07", "A1_COD", "A1_CDSIADN")

    SQL = "SELECT COUNT(*) as existe FROM SA1010 WHERE " & nomeColuna &  "= '" & Dados("codigoProtheus") & "'"

    On Error Resume Next
    SET rsVerificacao = Conecta(ObterConexaoProtheus(Dados("codigoPlanta")), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsVerificacao) Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar cliente""}"
        Exit Function
    End If
    On Error Goto 0

    If rsVerificacao("existe") = 0 Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Cliente nao encontrado. Certifique-se de que o cliente foi criado no Protheus""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    If Dados("limiteCredito") <= 0 Then
        'X - Foi aprovado com valor zerado
        SQL = SQL & "UPDATE UnificadoCliente SET "
        SQL = SQL & "   PreCadastro = 'X', "
        SQL = SQL & "   StatusCli = 'F', "
        SQL = SQL & "   CodSap = '" & Dados("codigoSap") & "', "
        SQL = SQL & "   CodPro = '" & Dados("codigoProtheus") & "', "
        SQL = SQL & "   CodLoja = '" & Dados("codigoLoja") & "' "
        SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; " 

        SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
        SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
        SQL = SQL & ") VALUES ("
        SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
        SQL = SQL & "   '" & Dados("codigo") & "', "
        SQL = SQL & "   '" & Dados("observacao") & "', "
        SQL = SQL & "   'F', "
        SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
        SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
        SQL = SQL & "   '" & usuarioLogado("login") & "'"
        SQL = SQL & "); "
    Else       
        'V - Aprovado com o valor informado
        SQL = SQL & "UPDATE UnificadoCliente SET "
        SQL = SQL & "PreCadastro = 'V', "
        SQL = SQL & "   StatusCli = 'F', "
        SQL = SQL & "   CodSap = '" & Dados("codigoSap") & "', "
        SQL = SQL & "   CodPro = '" & Dados("codigoProtheus") & "', "
        SQL = SQL & "   CodLoja = '" & Dados("codigoLoja") & "' "
        SQL = SQL & "WHERE Codigo = '" & Dados("codigo") & "'; " 

        SQL = SQL & "INSERT INTO UnificadoHistoricoCliente ( "
        SQL = SQL & "   Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem "
        SQL = SQL & ") VALUES ("
        SQL = SQL & "   (SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)), 0) + 1 FROM UnificadoHistoricoCliente), "
        SQL = SQL & "   '" & Dados("codigo") & "', "
        SQL = SQL & "   '" & Dados("observacao") & "', "
        SQL = SQL & "   'F', "
        SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 112), "
        SQL = SQL & "   CONVERT(CHAR(8), GETDATE(), 108), "
        SQL = SQL & "   '" & usuarioLogado("login") & "'"
        SQL = SQL & "); "

        SQL = SQL & "INSERT INTO SOLCREDITOCLIENTE ( "
        SQL = SQL & "   NUM, LOGININPANET, SOLICITANTE, COD, CODLOJA, LCATUAL, LCSOLICIT, "
        SQL = SQL & "   LCNOVO, DTSOLICIT, STATUS, QUEMVAL, DTVALID, LCVALID, Planta, "
        SQL = SQL & "   VolumeM2, PrecoM2, PrecoKg, StatusEnviado, NomeCliente, CodClientePrincipal,  "
        SQL = SQL & "   CNPJCPF "
        SQL = SQL & ") VALUES (" 
        SQL = SQL & "   (SELECT MAX(NUM)+1 FROM SOLCREDITOCLIENTE), " 
        SQL = SQL & "   '" & usuarioLogado("login") & "', " 
        SQL = SQL & "   '" & usuarioLogado("nome") & "', " 
        SQL = SQL & "   '" & Dados("codigoProtheus") & "', " 
        SQL = SQL & "   '" & Dados("codigoLoja") & "', " 
        SQL = SQL & "   0, " 
        SQL = SQL & "   " & Dados("limiteCredito") & ", "
        SQL = SQL & "   " & Dados("limiteCredito") & ", " 
        SQL = SQL & "   FORMAT(GETDATE(), 'yyyyMMdd'), " 
        SQL = SQL & "   'A', " 
        SQL = SQL & "   '', " 
        SQL = SQL & "   '', " 
        SQL = SQL & "   '', " 
        SQL = SQL & "   '" & Dados("codigoPlanta") & "', " 
        SQL = SQL & "   " & Dados("volumeVendaMensalM2") & ", "
        SQL = SQL & "   " & Dados("precoVendaNetM2") & ", "
        SQL = SQL & "   " & Dados("precoVendaNetKg")  & ", "
        SQL = SQL & "   NULL, " 
        SQL = SQL & "   '" & Dados("razaoSocial") & "', "
        SQL = SQL & "   '" & Dados("codigoProtheus") & "', " 
        SQL = SQL & "   '" & Dados("cpfCnpj") & "' " 
        SQL = SQL & "); "
    End If
    
    SQL = SQL & " IF @@ERROR <> 0 "
    SQL = SQL &   "ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL &   "COMMIT TRANSACTION "

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdate) Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Erro aprovar limite: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    SQL = "UPDATE SKFT1 SET" 
    SQL = SQL & " FT1_PRECADASTRO = '', "
    SQL = SQL & " FT1_B5CLI = '" & Dados("codigoProtheus") & "',"
    SQL = SQL & " FT1_OBSERV = FT1_OBSERV + ' ALTERADO PROSPECT DE: " & Dados("codigo") & ", PARA: " & Dados("codigoProtheus") & "'" 
    SQL = SQL & " WHERE FT1_PRECADASTRO ='P' AND FT1_B5CLI = '" & Dados("codigoSap") & "' AND FT1_FILIAL = '" & Dados("codigoPlanta") & "'"

    On Error Resume Next
    SET rsUpdateSAP = Conecta(ObterConexaoCorporativo(Dados("codigoPlanta")), SQL)

    If Err.Number <> 0 OR Not IsObject(rsUpdateSAP) Then
        Finalizar = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar cadastro: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCadastroFinalizado(Dados)
    Call RegistrarLog("finalizacao", Dados("codigo"))
    Finalizar = "{""sucesso"": true, ""mensagem"": ""Cadastro aprovado com sucesso!""}"
End Function

Function Cancelar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    
    SQL = "UPDATE UnificadoCliente SET "
    SQL = SQL & "StatusCli = 'X' "
    SQL = SQL & "WHERE Codigo = " & Dados("codigo")
  
    On Error Resume Next
    SET rsCancelar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCancelar) Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Erro ao cancelar: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailCancelamento(Dados)
    Call RegistrarLog("cancelamento", Dados("codigo"))
    Cancelar = "{""sucesso"": true, ""mensagem"": ""Cadastro cancelado com sucesso!""}"
End Function

Function Devolver()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    
    SQL = "UPDATE UnificadoCliente SET "
    SQL = SQL & "StatusCli = 'R' "
    SQL = SQL & "WHERE Codigo = " & Dados("codigo")
  
    On Error Resume Next
    SET rsDevolver = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsDevolver) Then
        Devolver = "{""sucesso"": false, ""mensagem"": ""Erro ao devolver: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call EnviarEmailDevolucao(Dados)
    Call RegistrarLog("devolucao_requisitante", Dados("codigo"))
    Devolver = "{""sucesso"": true, ""mensagem"": ""Cadastro devolvido com sucesso!""}"
End Function

Function Excluir()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")

    SQL = "UPDATE UnificadoCliente SET "
    SQL = SQL &   "Del = 1, "
    SQL = SQL &   "QuemExcluiu = '" & usuarioLogado("login") & "', "
    SQL = SQL &   "DataExclusao = GETDATE() "
    SQL = SQL & "WHERE Codigo = '" & codigo & "' "
  
    On Error Resume Next
    SET rsExcluir = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao Excluir: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("exclusao", Dados("codigo"))
    Excluir = "{""sucesso"": true, ""mensagem"": ""Cadastro excluído com sucesso!""}"
End Function

Function AnexarDocumentos()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("anexosJSON") = Request("anexos")

    ' Monta SQL Anexos
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
        valoresAnexos = valoresAnexos & "'" & Dados("codigo") & "', "
        valoresAnexos = valoresAnexos & "'" & Dados("codigoPlanta") & "', "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    SQL = "INSERT INTO UnificadoClienteDocumento ("
    SQL = SQL & "   idUnificadoCliente, Filial, Nome, Caminho, Usuario, DataDaOperacao "
    SQL = SQL & ") VALUES "
    SQL = SQL & valoresAnexos & "; "
  
    On Error Resume Next
    SET rsAnexar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAnexar) Then
        AnexarDocumentos = "{""sucesso"": false, ""mensagem"": ""Erro ao Anexar documentos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("anexar_documentos", Dados("codigo"))
    AnexarDocumentos = "{""sucesso"": true, ""mensagem"": ""Documentos anexados com sucesso!""}"
End Function

Function ListarHistoricoVistos()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigo") = Request("codigo")

    SQL = "WITH Historico AS "
    SQL = SQL & "( "
    SQL = SQL & "    SELECT "
    SQL = SQL & "        Codigo, "
    SQL = SQL & "        CodigoCad, "
    SQL = SQL & "        Obs, "
    SQL = SQL & "        HStatus, "
    SQL = SQL & "        CAST(CONVERT(date, DataAlt, 112) AS datetime) "
    SQL = SQL & "            + CAST(Hora AS datetime) AS DataInclusao, "
    SQL = SQL & "        CRM_NOME, "
    SQL = SQL & "        LAG( "
    SQL = SQL & "            CAST(CONVERT(date, DataAlt, 112) AS datetime) "
    SQL = SQL & "                + CAST(Hora AS datetime) "
    SQL = SQL & "        ) OVER "
    SQL = SQL & "        ( "
    SQL = SQL & "            PARTITION BY CodigoCad "
    SQL = SQL & "            ORDER BY CAST(Codigo AS bigint) "
    SQL = SQL & "        ) AS DataInclusaoAnterior "
    SQL = SQL & "    FROM UnificadoHistoricoCliente "
    SQL = SQL & "    LEFT JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = Quem "
    SQL = SQL & "   WHERE CodigoCad = '" & Dados("codigo") & "' "
    SQL = SQL & ") "
    SQL = SQL & "SELECT "
    SQL = SQL & "    Codigo AS Id, "
    SQL = SQL & "    CodigoCad, "
    SQL = SQL & "    Obs AS Observacao, "
    SQL = SQL & "    HStatus AS Status, "
    SQL = SQL & "    DataInclusao, "
    SQL = SQL & "    CRM_NOME AS LoginUsuarioInclusao, "
    SQL = SQL & "    CASE "
    SQL = SQL & "        WHEN DataInclusaoAnterior IS NULL THEN NULL "
    SQL = SQL & "        ELSE "
    SQL = SQL & "            DATEDIFF(DAY, DataInclusaoAnterior, DataInclusao) "
    SQL = SQL & "            - DATEDIFF(WEEK, DataInclusaoAnterior, DataInclusao) * 2 "
    SQL = SQL & "    END AS Leadtime "
    SQL = SQL & "FROM Historico "
    SQL = SQL & "ORDER BY CAST(Codigo AS bigint); "

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistoricoVistos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
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
            jsonArray = jsonArray & """codigoCad"": """ & EscapeJSON(rsHistorico("CodigoCad")) & """, "
            jsonArray = jsonArray & """observacao"": """ & EscapeJSON(rsHistorico("Observacao")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsHistorico("Status")) & """, "
            jsonArray = jsonArray & """leadtime"": """ & EscapeJSON(rsHistorico("Leadtime")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsHistorico("LoginUsuarioInclusao")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistoricoVistos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarCidades()
    codigoEstado = Request("codigoEstado")
    If codigoEstado = "" Then
        ListarCidades = "{""sucesso"": false, ""mensagem"": ""Nenhum estado informado.""}"
        Exit Function
    End If
    
    SQL = "SELECT DISTINCT NOME as Nome, CODMUN as Codigo FROM UnificadoCidades WHERE ESTADO = '" & codigoEstado & "' ORDER BY NOME"

    On Error Resume Next
    SET rsCidades = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCidades) Then
        ListarCidades = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCidades.EOF Then
        Do While Not rsCidades.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsCidades("Codigo")) & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsCidades("Nome")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCidades.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCidades.Close()
    SET rsCidades = NOTHING
    
    ListarCidades = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarCondicoesPagamento()
    SQL = "SELECT Codigo, Descricao FROM UnificadoCondPag ORDER BY Codigo"

    On Error Resume Next
    SET rsCondicoes = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCondicoes) Then
        ListarCondicoesPagamento = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar condições de pagamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCondicoes.EOF Then
        Do While Not rsCondicoes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsCondicoes("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsCondicoes("Descricao")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCondicoes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCondicoes.Close()
    SET rsCondicoes = NOTHING
    
    ListarCondicoesPagamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarSegmentos()
    SQL = "SELECT AOV_CODSEG AS Codigo, AOV_DESSEG AS Descricao FROM AOV010 WHERE AOV_MSBLQL = '2' AND ISNULL(AOV010.D_E_L_E_T_,'') = '' "
    SQL = SQL & "ORDER BY AOV_DESSEG "

    On Error Resume Next
    SET rsSegmentos = Conecta("PIR-DADOSADV", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSegmentos) Then
        ListarSegmentos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSegmentos.EOF Then
        Do While Not rsSegmentos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsSegmentos("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsSegmentos("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSegmentos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSegmentos.Close()
    SET rsSegmentos = NOTHING
    
    ListarSegmentos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposIndustria()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CN" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Construtora" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CR" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Comércio" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "IN" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Indústria" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PE" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Preponderantemente Exportador" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Revendedor Atacadista" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RV" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Revendedor Varejista" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposIndustria = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarContasConciliacao()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "41100000" & """, "
    jsonArray = jsonArray & """descricao"": """ & "41100000 - Clientes terceiros nacional" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "41100100" & """, "
    jsonArray = jsonArray & """descricao"": """ & "41100100 - Clientes terceiros internacional" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "94100000" & """, "
    jsonArray = jsonArray & """descricao"": """ & "94100000 - Clientes intraco nacional - operação dentro da mesma entidade legal" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "41200100" & """, "
    jsonArray = jsonArray & """descricao"": """ & "41200100 - Clientes interco internacional" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "41200000" & """, "
    jsonArray = jsonArray & """descricao"": """ & "41200000 - Clientes interco nacional - operação entre entidades legais distintas" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarContasConciliacao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarEstados()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AC" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Acre" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Alagoas" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AP" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Amapá" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AM" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Amazonas" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "BA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Bahia" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CE" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Ceará" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "DF" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Distrito Federal" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "ES" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Espírito Santo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "GO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Goías" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "MA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Maranhão" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "MT" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Mato Grosso" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "MS" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Mato Grosso do Sul" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "MG" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Minas Gerais" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PA" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pará" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PB" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Paraíba" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PR" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Paraná" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PE" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pernambuco" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "PI" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Piauí" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RJ" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Rio de Janeiro" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RN" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Rio Grande do Norte" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RS" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Rio Grande do Sul" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Rondônia" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RR" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Roraima" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "SC" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Santa Catarina" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "SP" & """, "
    jsonArray = jsonArray & """descricao"": """ & "São Paulo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "SE" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sergipe" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "TO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Tocantins" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarEstados = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
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

Function ListarTiposCliente()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "PRODUTOR RURAL" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "R" & """, "
    jsonArray = jsonArray & """descricao"": """ & "REVENDEDOR" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CONS. FINAL" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "S" & """, "
    jsonArray = jsonArray & """descricao"": """ & "SOLIDARIO" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "E" & """, "
    jsonArray = jsonArray & """descricao"": """ & "EXPORTAÇÃO" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposCliente = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
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
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsVendedores("Codigo")) & """, "
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

Function ListarGruposConta()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "Z003N" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CLIENTES NACIONAIS" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "Z003I" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CLIENTES INTERNACIONAIS" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarGruposConta = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
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

Function ListarTiposDeclarante()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CM" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Regime Cumulativo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "NC" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Regime Não Cumulativo" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposDeclarante = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarMetodosPagamento()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "D" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Boleto" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "A" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Ordem de Pagamento" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarMetodosPagamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposFrete()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CIF" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "F" & """, "
    jsonArray = jsonArray & """descricao"": """ & "FOB" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposFrete = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterNivelUsuario()
    nivelUsuario = ""
    
    SQL = "SELECT "
    SQL = SQL & "   ISNULL(NivelControlesInternos, 0) AS NivelControlesInternos, "
    SQL = SQL & "   ISNULL(NivelControladoria, 0) AS NivelControladoria, "
    SQL = SQL & "   ISNULL(NivelE2OPEN, 0) AS NivelCDAG, "
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
        If rsPermissao("NivelControlesInternos") Then
            nivelUsuario = "CONTROLES_INTERNOS"
        ElseIf rsPermissao("NivelControladoria") Then
            nivelUsuario = "CONTROLADORIA"
        ElseIf rsPermissao("NivelCDAG") Then
            nivelUsuario = "CDAG"
        ElseIf rsPermissao("NivelFiscal") Then
            nivelUsuario = "FISCAL"
        End If
    End If

    rsPermissao.Close()
    SET rsPermissao = NOTHING

    ' Se não tiver nenhum nível consideramos que é solicitante
    If nivelUsuario = "" Then
        nivelUsuario = "SOLICITANTE"
    End If
    
    dadosJSON = "{""nivel"": """ & nivelUsuario & """}"
    ObterNivelUsuario = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function UsuarioPodeValidarEtapa(dadosOrcamento, idUsuario)
    temPermissao = False

    SQL = ""

    On Error Resume Next
    SET rsPermissao = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsPermissao) Then
        UsuarioPodeValidarEtapa = temPermissao
        Exit Function
    End If

    If Not rsPermissao.EOF Then
        temPermissao = True
    End If
    
    rsPermissao.Close()
    SET rsPermissao = NOTHING
    On Error Goto 0

    UsuarioPodeValidarEtapa = temPermissao
End Function

Function ObterConexaoCorporativo(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "04": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "05": ObterConexaoCorporativo = "BEN-CORPORATIVO"
        Case "06": ObterConexaoCorporativo = "FOR-CORPORATIVO"
        Case "07": ObterConexaoCorporativo = "SAQ-CORPORATIVO"
        Case Else: ObterConexaoCorporativo = ""
    End Select
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

Function ObterCorrespondencia(codigoPlanta)
    Select Case codigoPlanta
        Case "00": ObterCorrespondencia = "BR02"
        Case "04": ObterCorrespondencia = "BR02"
        Case "05": ObterCorrespondencia = "BR03"
        Case "06": ObterCorrespondencia = "BR04"
        Case "07": ObterCorrespondencia = "BR07"
        Case Else: ObterCorrespondencia = ""
    End Select
End Function

Function GerarCodigoCadastro(codPlanta)
    codigo = ""

    SQL = "SELECT (ISNULL(MAX(CAST(Codigo AS bigint)), 999) + 1) AS Codigo FROM UnificadoCliente  "

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            codigo = CStr(PIR_CORPORATIVO("Codigo"))
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    GerarCodigoCadastro = codigo
End Function

' Auxiliares
Function VerificarCpfCnpjExiste(cpfCnpj)
    flagExiste = False
    
    SQL = "SELECT COUNT(*) as qtde FROM UnificadoCliente WHERE ISNULL(Del, 0) = 0 AND (Cnpj = '" & cpfCnpj & "' OR Cpf = '" & cpfCnpj & "') AND StatusCli <> 'X'"

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        VerificarCpfCnpjExiste = flagExiste
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.EOF AND rsExiste("qtde") > 0 Then
        flagExiste = True
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    
    VerificarCpfCnpjExiste = flagExiste
End Function

Function EscapeHTML(valor)
    If IsNull(valor) Then valor = ""
    EscapeHTML = Server.HTMLEncode(CStr(valor))
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

Function ObterDadosEmailCadastro(codigo)
    Set dadosEmail = Server.CreateObject("Scripting.Dictionary")
    dadosEmail("codigo") = codigo
    dadosEmail("razaoSocial") = ""
    dadosEmail("nomeFantasia") = ""
    dadosEmail("documento") = ""
    dadosEmail("codigoPlanta") = ""
    dadosEmail("nomeSolicitante") = ""
    dadosEmail("emailSolicitante") = ""

    SQL = "SELECT TOP 1 "
    SQL = SQL & "   a.Codigo, "
    SQL = SQL & "   a.RazaoSocial, "
    SQL = SQL & "   a.NomeFantasia, "
    SQL = SQL & "   ISNULL(NULLIF(a.Cnpj, ''), a.Cpf) AS Documento, "
    SQL = SQL & "   a.Filial AS CodigoPlanta, "
    SQL = SQL & "   ISNULL(t.CRM_NOME, a.QuemIncluiu) AS NomeSolicitante, "
    SQL = SQL & "   ISNULL(t.CRM_EMAILPRIMARIO, '') AS EmailSolicitante "
    SQL = SQL & "FROM UnificadoCliente a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t ON t.CRM_LOGIN = a.QuemIncluiu "
    SQL = SQL & "WHERE a.Codigo = '" & Replace(codigo, "'", "''") & "' "

    On Error Resume Next
    SET rsDadosEmail = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(rsDadosEmail) Then
        If Not rsDadosEmail.EOF Then
            dadosEmail("codigo") = rsDadosEmail("Codigo") & ""
            dadosEmail("razaoSocial") = rsDadosEmail("RazaoSocial") & ""
            dadosEmail("nomeFantasia") = rsDadosEmail("NomeFantasia") & ""
            dadosEmail("documento") = rsDadosEmail("Documento") & ""
            dadosEmail("codigoPlanta") = rsDadosEmail("CodigoPlanta") & ""
            dadosEmail("nomeSolicitante") = rsDadosEmail("NomeSolicitante") & ""
            dadosEmail("emailSolicitante") = rsDadosEmail("EmailSolicitante") & ""
            'dadosEmail("emailSolicitante") = "thiago.pereira@smurfitwestrock.com.br"
        End If

        rsDadosEmail.Close()
        SET rsDadosEmail = NOTHING
    End If

    On Error Goto 0

    Set ObterDadosEmailCadastro = dadosEmail
End Function

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
        Case "CONTROLES_INTERNOS"
            colunaNivel = "NivelControlesInternos"
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

    'emails = "thiago.pereira@smurfitwestrock.com.br"
    ObterEmailsUsuariosNivel = emails
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
    mensagem = mensagem & "<div class=""info-item""><strong>Código:</strong> " & EscapeHTML(dadosEmail("codigo")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Razão Social:</strong> " & EscapeHTML(dadosEmail("razaoSocial")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Nome Fantasia:</strong> " & EscapeHTML(dadosEmail("nomeFantasia")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>CNPJ/CPF:</strong> " & EscapeHTML(dadosEmail("documento")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & EscapeHTML(dadosEmail("codigoPlanta")) & "</div>"
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

Sub EnviarEmailCadastroParaNivel(Dados, nivelDestino, nomeNivel)
    destinatarios = ObterEmailsUsuariosNivel(nivelDestino)
    If destinatarios = "" Then Exit Sub

    texto = "Um cadastro de cliente foi enviado para análise do nível " & nomeNivel & "."
    titulo = "Cadastro enviado para " & nomeNivel

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailCadastroFinalizado(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi finalizado."
    titulo = "Cadastro finalizado"    

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailDevolucao(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi devolvido ao solicitante para ajustes."
    titulo = "Cadastro devolvido ao solicitante"

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailCancelamento(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi cancelado."
    titulo = "Cadastro cancelado"

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub


Sub RegistrarLog(acao, codigo)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Cadastro Unificado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou o cadastro de codigo " & codigo
        Case "edicao"
            acaoLog = "Edição de Cadastro Unificado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou o cadastro de codigo " & codigo
        Case "cancelamento"
            acaoLog = "Cancelamento de Cadastro Unificado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " cancelou o cadastro de codigo " & codigo 
        Case "exclusao"
            acaoLog = "Exclusão de Cadastro Unificado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu o cadastro de codigo " & codigo 
        Case "envio_fiscal"
            acaoLog = "Envio para Fiscal"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou o cadastro de codigo " & codigo & " para o Fiscal"
        Case "envio_controladoria"
            acaoLog = "Envio para Controladoria"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou o cadastro de codigo " & codigo & " para a Controladoria"
        Case "envio_cdag"
            acaoLog = "Envio para CDAG"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou o cadastro de codigo " & codigo & " para o CDAG"
        Case "envio_controles_internos"
            acaoLog = "Envio para Controles Internos"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou o cadastro de codigo " & codigo & " para o Controles Internos"
        Case "finalizacao"
            acaoLog = "Finalização de Cadastro Unificado"
            descricaoLog = "Usuário " & usuarioLogado("login") & " finalizou o cadastro de codigo " & codigo 
        Case "devolucao_requisitante"
            acaoLog = "Devolucao para Requisitante"
            descricaoLog = "Usuário " & usuarioLogado("login") & " devolveu o cadastro de codigo " & codigo & " para o Requisitante"
        Case "anexar_documentos"
            acaoLog = "Anexar Documentos"
            descricaoLog = "Usuário " & usuarioLogado("login") & " anexou documentos ao cadastro de codigo " & codigo
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Cadastro Unificado', "
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