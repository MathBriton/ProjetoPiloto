<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<!--#include file="./util/Util.asp" -->
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
usuarioLogado.Add "email", Session("EMAIL")

' Variáveis globais de limites de aprovação
limiteMaximoAprovacaoComercial = 180000
limiteMaximoAprovacaoCoordenador = 1000000
limiteMaximoAprovacaoGerenteTesouraria = 3000000
limiteMaximoAprovacaoGerentePlanta = 5000000

Dim acao
acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

If usuarioLogado("id") = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Sua sessão expirou, realize login novamente para continuar""}")
    Response.End()
End If

Dim resultado
Select case acao
    Case "criarSolicitacao"
        resultado = CriarSolicitacao()
    Case "validarSolicitacao"
        resultado = ValidarSolicitacao()
    Case "obterCliente"
        resultado = ObterCliente(Request("planta"), Request("codCliente"), Request("codLoja"))
    Case "obterValidador"
        resultado = ObterValidador(usuarioLogado("id"))
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


'--------------------------------------
'           FUNÇÕES PRINCIPAIS
'--------------------------------------

Function CriarSolicitacao()
    Dim planta, codCliente, codLoja, limiteSolicitado, anexosJSON
    Dim volumeM2, precoM2, precoKg, condicaoPagamento
    
    ' Dados Formulario
    planta = Request("planta")
    codLoja = Request("codLoja")
    tipoMercado = Request("tipoMercado")
    codCliente = Request("codCliente")
    limiteSolicitado = Request("limiteSolicitado")
    condicaoPagamento = Request("condicaoPagamento")
    observacaoVendedor = Request("observacaoVendedor")
    anexosJSON = Request("anexos")

    precoM2 = iif (Request("precoM2") <> "", Request("precoM2"), 0)
    precoKg = iif (Request("precoKg") <> "", Request("precoKg"), 0)
    volumeM2 = iif (Request("volumeM2") <> "", Request("volumeM2"), 0)
    volumeUnidades = iif (Request("volumeUnidades") <> "", Request("volumeUnidades"), 0)
    volumeToneladas = iif (Request("volumeToneladas") <> "", Request("volumeToneladas"), 0)
    valorToneladaTora = iif (Request("valorToneladaTora") <> "", Request("valorToneladaTora"), 0)
    valorUnitarioSemente = iif (Request("valorUnitarioSemente") <> "", Request("valorUnitarioSemente"), 0)
    valorUnitarioMuda = iif (Request("valorUnitarioMuda") <> "", Request("valorUnitarioMuda"), 0)

    ' Validações
    Dim respostaCliente, clienteObj
    respostaCliente = ObterCliente(planta, codCliente, codLoja)
    Set clienteObj = JsonToDictionary(respostaCliente)

    If Not clienteObj("sucesso") Then
        CriarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Cliente inválido""}"
        Exit Function
    End If

    Dim dadosCliente
    Set dadosCliente = clienteObj("dados")
    
    If CDbl(limiteSolicitado) < CDbl(dadosCliente("aberto")) Then
        CriarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Limite solicitado não pode ser menor que o valor em aberto. Valor em aberto atual: " & FormatarMoeda(dadosCliente("aberto"), 2) & """}"
        Exit Function
    End If
    
    If VerificarSolicitacaoPendente(planta, codCliente, codLoja) Then
        CriarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Já existe uma solicitação de aumento pendente para este cliente.""}"
        Exit Function
    End If
    
    Dim idNovaSolicitacao
    idNovaSolicitacao = GerarProximoIdSolicitacao()
    If idNovaSolicitacao = "" Then
        CriarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Erro ao obter próximo id""}"
        Exit Function
    End If
    
    Dim SQL, PIR_CORPORATIVO, limiteAtual
    limiteAtual = Replace(CStr(dadosCliente("limite")), ",", ".")
        
    SQL = "INSERT INTO SOLCREDITOCLIENTE "
    SQL = SQL & "(NUM, LOGININPANET, SOLICITANTE, NomeCliente, ObservacaoVendedor, QUEMVAL, COD, CodClientePrincipal, CODLOJA, Planta, LCATUAL, LCSOLICIT, LCNOVO, LCVALID, FaturamentoEstimadoPositivo, DTSOLICIT, DTVALID, STATUS, VolumeM2, PrecoM2, PrecoKg, CondicaoPagamento, ClienteGrupo, CNPJCPF, TipoMercado, ValorToneladaTora, ValorUnitarioSemente, ValorUnitarioMuda, VolumeUnidades, VolumeToneladas) VALUES ("
    SQL = SQL & "'" & idNovaSolicitacao & "', "
    SQL = SQL & "'" & usuarioLogado("login") & "', "
    SQL = SQL & "'" & usuarioLogado("nome") & "', "
    SQL = SQL & "'" & dadosCliente("nomeCliente") & "', "
    SQL = SQL & "'" & observacaoVendedor & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & codCliente & "', "
    SQL = SQL & "'" & dadosCliente("codClientePrincipal") & "', "
    SQL = SQL & "'" & codLoja & "', "
    SQL = SQL & "'" & planta & "', "
    SQL = SQL & limiteAtual & ", "
    SQL = SQL & limiteSolicitado & ", "
    SQL = SQL & limiteSolicitado & ", "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "'', "
    SQL = SQL & "'" & ObterCodStatus("NOVO") & "', "
    SQL = SQL & volumeM2 & ", "
    SQL = SQL & precoM2 & ", "
    SQL = SQL & precoKg & ", "
    SQL = SQL & "'" & condicaoPagamento & "', "
    SQL = SQL & "'" & dadosCliente("clienteGrupo") & "', "
    SQL = SQL & "'" & dadosCliente("cnpj") & "', "
    SQL = SQL & "'" & tipoMercado & "', "
    SQL = SQL & valorToneladaTora & ", "
    SQL = SQL & valorUnitarioSemente & ", "
    SQL = SQL & valorUnitarioMuda & ", "
    SQL = SQL & volumeUnidades & ", "
    SQL = SQL & volumeToneladas
    SQL = SQL & ")"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(PIR_CORPORATIVO) Then
        CriarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Erro ao criar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0

    If anexosJSON <> "" Then
        Call SalvarAnexosSolicitacao(idNovaSolicitacao, anexosJSON)
    End If

    Call RegistrarLog("criacao", idNovaSolicitacao, 0)
    Call EnviarEmailSolicitacaoCriada(dadosCliente("nomeCliente"), codCliente, planta, ObterNomePlanta(planta), dadosCliente("limite"), limiteSolicitado)
    CriarSolicitacao = "{""sucesso"": true, ""mensagem"": ""Solicitação criada com sucesso!""}"
End Function


Function ValidarSolicitacao()
    ' ***** Dados recebidos
    idSolicitacao = Request("idSolicitacao")
    planta = Request("planta")
    flagRejeitar = Request("flagRejeitar")
    limiteValidado = Request("limiteValidado")
    observacao = Request("observacao")
    scoreSerasa = Request("scoreSerasa")
    limiteSugeridoSerasa = Request("limiteSugeridoSerasa")
    anexosJSON = Request("anexos")
    classeRiscoSerasa = Request("classeRiscoSerasa")
    balancoAno = Request("balancoAno")

    faturamentoEstimadoPositivo = iif(Request("faturamentoEstimadoPositivo") <> "", Request("faturamentoEstimadoPositivo"), 0) 
    valorSeguradora = iif(Request("valorSeguradora") <> "", Request("valorSeguradora"), 0)
    limiteSugeridoSW = iif(Request("limiteSugeridoSW") <> "", Request("limiteSugeridoSW"), 0)
    scoreSW = iif(Request("scoreSW") <> "", Request("scoreSW"), 0) 
    classeRiscoSW = iif(Request("classeRiscoSW") <> "", Request("classeRiscoSW"), "") 

    respostaSolicitacao = ObterSolicitacao(idSolicitacao, planta)
    Set solicitacaoObj = JsonToDictionary(respostaSolicitacao)
    
    If Not solicitacaoObj("sucesso") Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")
    
    ' ***** Verifica permissões
    respostaValidador = ObterValidador(usuarioLogado("id"))
    Set validadorObj = JsonToDictionary(respostaValidador)
    
    If Not validadorObj("sucesso") Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Validador inválido""}"
        Exit Function
    End If
    
    Set validador = validadorObj("dados")
    nivelValidador = ObterNivelValidadorTraduzido(validador("codNivel"))

    If nivelValidador = "" Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Nível de validador inválido""}"
        Exit Function
    End If
    
    statusAtual = ObterStatusTraduzido(solicitacao("status"))

    If Not UsuarioPodeValidarEtapa(statusAtual, nivelValidador) Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Você não tem permissão para validar essa solicitação""}"
        Exit Function
    End If

    If trim(solicitacao("nomeAnalista")) = trim(usuarioLogado("nome")) Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Usuário analista não pode aprovar""}"
        Exit Function
    End If

    ' Gera o SQL de atualização baseado na etapa atual
    ' Só usamos esse SQL no final do processo caso validação tenha ocorrido com sucesso no protheus
    ' Fazemos isso pois se houver um problema do lado do protheus não realizamos atualização do nosso lado
    Select Case nivelValidador
        ' Nível analista participa das duas primeiras etapas: análise e validação
        Case "ANALISTA"
            If statusAtual = "NOVO" Then
                Set resultadoValidacao = AnaliseAnalista(idSolicitacao, flagRejeitar, faturamentoEstimadoPositivo, valorSeguradora, limiteValidado, observacao, anexosJSON, scoreSerasa, limiteSugeridoSerasa, limiteSugeridoSW, scoreSW, classeRiscoSW, classeRiscoSerasa, balancoAno)
            Else 
                Set resultadoValidacao = ValidacaoAnalista(idSolicitacao, flagRejeitar, limiteValidado, observacao, solicitacao("clienteGrupo"))
            End If        
        Case "COORDENADOR"
            Set resultadoValidacao = ValidacaoCoordenador(idSolicitacao, flagRejeitar, limiteValidado, observacao)
        Case "GERENTE_TESOURARIA"
            Set resultadoValidacao = ValidacaoGerenteTesouraria(idSolicitacao, flagRejeitar, limiteValidado, observacao)
        Case "CONTROLLER"
            Set resultadoValidacao = ValidacaoController(idSolicitacao, flagRejeitar, limiteValidado, observacao)
        Case "GERENTE_PLANTA"
            Set resultadoValidacao = ValidacaoGerentePlanta(idSolicitacao, flagRejeitar, limiteValidado, observacao)
        Case "CFO"
            Set resultadoValidacao = ValidacaoCFO(idSolicitacao, flagRejeitar, limiteValidado, observacao)
        Case Else
            ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Nível de validador inválido""}"
            Exit Function
    End Select

    novoStatus = resultadoValidacao("novoStatus")

    If novoStatus = "APROVADO" OR novoStatus = "REJEITADO" Then
        flagFinalizado = 1
    Else
        flagFinalizado = 0
    End If

    ' Se novo status for APROVADO atualizamos o limite do cliente e enviamos email
    If novoStatus = "APROVADO" Then
        flagAtualizou = AtualizarLimiteCliente(idSolicitacao, solicitacao("planta"), solicitacao("codCliente"), solicitacao("codClientePrincipal"), solicitacao("codLoja"), solicitacao("nomeCliente"), limiteValidado, solicitacao("limiteAtual"), solicitacao("faturamentoEstimadoPositivo"), solicitacao("scoreSerasa"), solicitacao("limiteSugeridoSerasa"), solicitacao("limiteSugeridoSW"), solicitacao("scoreSW"), solicitacao("classeRiscoSW"), solicitacao("classeRiscoSerasa"), solicitacao("balancoAno"), solicitacao("clienteGrupo"), solicitacao("volumeM2"), solicitacao("precoM2"), solicitacao("precoKg"), solicitacao("condicaoPagamento"), solicitacao("cnpjCpf"), solicitacao("tipoMercado"), solicitacao("valorSeguradora"))

        If flagAtualizou = 0 Then
            ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Ocorreu um erro ao atualizar o limite do cliente, tente novamente""}"
            Exit Function
        End If
        
        Call EnviarEmailSolicitacaoAprovada(solicitacao("loginSolicitante"), solicitacao("nomeSolicitante"), solicitacao("nomeAnalista"), solicitacao("nomeCliente"), solicitacao("codCliente"), solicitacao("planta"), ObterNomePlanta(solicitacao("planta")), solicitacao("limiteAtual"), solicitacao("limiteSolicitado"), limiteValidado)
    End If

    ' Se novo status for REJEITADO enviamos email de rejeição
    If novoStatus = "REJEITADO" Then
        Call EnviarEmailSolicitacaoRejeitada(solicitacao("loginSolicitante"), solicitacao("nomeSolicitante"), solicitacao("nomeCliente"), solicitacao("codCliente"), solicitacao("planta"), ObterNomePlanta(solicitacao("planta")), solicitacao("limiteAtual"), solicitacao("limiteSolicitado"), observacao)
    End If

    ' Executa SQL de atualização
    SQL = resultadoValidacao("SQL")

    On Error Resume Next
    SET rsAtualizacao = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAtualizacao) Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Erro ao criar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    ' Salva anexos se tiver sido passado
    If anexosJSON <> "" Then
        Call SalvarAnexosSolicitacao(idSolicitacao, anexosJSON)
    End If

    ' Se não tiver finalizado (nem aprovado, nem rejeitado) continua Fluxo e envia email para proximos validadores
    If flagFinalizado = 0 Then
        Call EnviarEmailSolicitacaoValidada(solicitacao("loginSolicitante"), solicitacao("nomeSolicitante"), solicitacao("nomeAnalista"), novoStatus, solicitacao("nomeCliente"), solicitacao("codCliente"), solicitacao("planta"), ObterNomePlanta(solicitacao("planta")), solicitacao("limiteAtual"), limiteValidado)
    End If

    Call RegistrarLog("validacao", idSolicitacao, flagRejeitar)
    ValidarSolicitacao = "{""sucesso"": true, ""mensagem"": ""Solicitação atualizada com sucesso!"", ""flagFinalizado"": " & flagFinalizado & "}"
End Function


Public Function ObterCliente(plantaParam, codClienteParam, codLojaParam)
    If plantaParam = "" Then
        ObterCliente = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    If codClienteParam = "" Then
        ObterCliente = "{""sucesso"": false, ""mensagem"": ""Código de cliente não informado""}"
        Exit Function
    End If

    If plantaParam = "07" AND codLojaParam = "" Then
        ObterCliente = "{""sucesso"": false, ""mensagem"": ""Código de loja não informado""}"
        Exit Function
    End If
    
    limite = 0
    saldo = 0
    aberto = 0
    condicaoPagamento = ""
    nomeCliente = ""
    clienteEncontrado = False
    
    conexao = ObterConexao(plantaParam)
    
    ' Query para pegar nome do cliente
    If conexao = "SPO-BI" Then
        SQL = SQL & "WITH Cliente as ( "
        SQL = SQL & "   SELECT DISTINCT "
        SQL = SQL & "       b1.NOME_CLIENTE as Nome, "
        SQL = SQL & "       b1.CNPJ, "
        SQL = SQL & "       b1.CONDICAO_PAGAMENTO as CondicaoPagamento, "
        SQL = SQL & "       b1.CODIGO_SAP_PRINCIPAL as CodClientePrincipal, "
        SQL = SQL & "       '01' as CodLoja, "
        SQL = SQL & "       '0' as ClienteGrupo "
        SQL = SQL & "   FROM BI_CLIENTES_WRK b1 "
        SQL = SQL & "   WHERE CODIGO_SAP = '" & codClienteParam & "' "
        SQL = SQL & ")"
        SQL = SQL & "SELECT Cliente.*, b2.LIMITE_CREDITO as Limite, b2.NOME_CLIENTE AS NomeClientePrincipal FROM Cliente "
        SQL = SQL & "LEFT JOIN BI_CLIENTES_WRK b2 on b2.CODIGO_SAP = cliente.CodClientePrincipal"
    Else
        SQL = "SELECT Nome, DescricaoCondPag as CondicaoPagamento, CnpjCpf as CNPJ, Loja as CodLoja, "
        SQL = SQL & "CASE WHEN ClienteGrupo = 'S' THEN '1' ELSE '0' END AS ClienteGrupo, '' as CodClientePrincipal, '' as NomeClientePrincipal "
        SQL = SQL & "FROM vwCliente WHERE Codigo = '" & codClienteParam & "' "
        
        If codLojaParam <> "" Then
            SQL = SQL & "AND Loja = '" & codLojaParam & "'"
        End If
    End If

    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexao, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        If Not CLIENTE_CONEXAO.EOF Then
            limite = iif(Not IsNull(CLIENTE_CONEXAO("Limite")), CLIENTE_CONEXAO("Limite"), 0)
            condicaoPagamento = iif(Not IsNull(CLIENTE_CONEXAO("CondicaoPagamento")), CLIENTE_CONEXAO("CondicaoPagamento"), "")
            cnpj = iif(Not IsNull(CLIENTE_CONEXAO("CNPJ")), CLIENTE_CONEXAO("CNPJ"), "")
            codLoja = iif(Not IsNull(CLIENTE_CONEXAO("CodLoja")), CLIENTE_CONEXAO("CodLoja"), "")
            nomeCliente = CLIENTE_CONEXAO("Nome")
            clienteGrupo = CLIENTE_CONEXAO("ClienteGrupo")
            codClientePrincipal = CLIENTE_CONEXAO("CodClientePrincipal")
            nomeClientePrincipal = CLIENTE_CONEXAO("NomeClientePrincipal")
            clienteEncontrado = True
        End If

        CLIENTE_CONEXAO.Close()
        SET CLIENTE_CONEXAO = NOTHING
    End If
    On Error Goto 0
    
    If Not clienteEncontrado Then
        ObterCliente = "{""sucesso"": false, ""mensagem"": ""Cliente não encontrado!""}"
        Exit Function
    End If
    
    ' Query para buscar dados de saldo (exceto Westrock que não tem estes dados)
    If conexao <> "SPO-BI" Then
        SQL = "SELECT * FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA("
        SQL = SQL & "'" & codClienteParam & "'"

        If codLojaParam <> "" AND plantaParam = "07" Then 
            SQL = SQL & ", '" & codLojaParam & "'"
        End If

        SQL = SQL & ")"
        
        On Error Resume Next
        Err.Clear
        SET SALDO_CONEXAO = Conecta(conexao, SQL)
        
        If Err.Number = 0 AND IsObject(SALDO_CONEXAO) Then
            If Not SALDO_CONEXAO.EOF Then
                limite = SALDO_CONEXAO("Limite")
                saldo = SALDO_CONEXAO("Saldo")
                aberto = limite - saldo
            Else
                SALDO_CONEXAO.Close()
                SET SALDO_CONEXAO = NOTHING
                ObterCliente = "{""sucesso"": false, ""mensagem"": ""Não foi possível obter os dados de saldo do cliente!""}"
                Exit Function
            End If
            SALDO_CONEXAO.Close()
            SET SALDO_CONEXAO = NOTHING
        Else
            If IsObject(SALDO_CONEXAO) Then
                SALDO_CONEXAO.Close()
                SET SALDO_CONEXAO = NOTHING
            End If
            ObterCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao obter dados de saldo do cliente!""}"
            Exit Function
        End If
        On Error Goto 0
    End If
    
    dadosJson = "{"
    dadosJson = dadosJson & """nomeCliente"":""" & EscapeJSON(nomeCliente) & """"
    dadosJson = dadosJson & ",""codClientePrincipal"":""" & EscapeJSON(codClientePrincipal) & """"
    dadosJson = dadosJson & ",""nomeClientePrincipal"":""" & EscapeJSON(nomeClientePrincipal) & """"
    dadosJson = dadosJson & ",""codLoja"":""" & EscapeJSON(codLoja) & """"
    dadosJson = dadosJson & ",""cnpj"":""" & EscapeJSON(cnpj) & """"
    dadosJson = dadosJson & ",""clienteGrupo"":""" & EscapeJSON(clienteGrupo) & """"
    dadosJson = dadosJson & ",""condicaoPagamento"":""" & EscapeJSON(condicaoPagamento) & """"
    dadosJson = dadosJson & ",""limite"":" & Replace(CStr(limite), ",", ".")
    dadosJson = dadosJson & ",""aberto"":" & Replace(CStr(aberto), ",", ".")
    dadosJson = dadosJson & ",""saldo"":" & Replace(CStr(saldo), ",", ".")
    dadosJson = dadosJson & "}"
    
    ObterCliente = "{""sucesso"": true, ""mensagem"": ""Dados obtidos com sucesso!"", ""dados"": " & dadosJson & "}"
End Function

Public Function ObterValidador(idUsuarioParam)
    Dim SQL, PIR_CORPORATIVO
    Dim idValidador, nome, codNivel, minimo, maximo
    Dim validadorEncontrado

    If idUsuarioParam = "" Then
        ObterValidador = "{""sucesso"": false, ""mensagem"": ""Id de usuário não informado""}"
        Exit Function
    End If
    
    validadorEncontrado = False
    
    SQL = "SELECT DISTINCT "
    SQL = SQL & " CAST(vc.IdValidador AS INT) AS IdValidador,"
    SQL = SQL & " u.CRM_NOME AS Nome,"
    SQL = SQL & " CAST(vc.Nivel AS NVARCHAR(10)) as CodNivel,"
    SQL = SQL & " MIN(vc.Minimo) AS Minimo,"
    SQL = SQL & " MAX(vc.Maximo) AS Maximo"
    SQL = SQL & " FROM ValidadoresCredito vc"
    SQL = SQL & " LEFT JOIN HELPDESK..CRM_DTI u ON vc.IdValidador = u.CRM_R_E_C_N_O"
    SQL = SQL & " WHERE vc.IdValidador = " & idUsuarioParam
    SQL = SQL & " AND DataExclusao is NULL" 
    SQL = SQL & " GROUP BY vc.IdValidador, u.CRM_NOME, vc.Nivel"
    SQL = SQL & " ORDER BY Nome"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            idValidador = PIR_CORPORATIVO("IdValidador")
            nome = PIR_CORPORATIVO("Nome")
            codNivel = PIR_CORPORATIVO("CodNivel")
            minimo = PIR_CORPORATIVO("Minimo")
            maximo = PIR_CORPORATIVO("Maximo")
            validadorEncontrado = True
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
    
    If Not validadorEncontrado Then
        ObterValidador = "{""sucesso"": false, ""mensagem"": ""Validador não encontrado!""}"
        Exit Function
    End If
    
    Dim dadosJson
    dadosJson = "{"
    dadosJson = dadosJson & """idValidador"":" & idValidador
    dadosJson = dadosJson & ",""nome"":""" & EscapeJSON(nome) & """"
    dadosJson = dadosJson & ",""codNivel"":" & codNivel
    dadosJson = dadosJson & ",""minimo"":" & Replace(CStr(minimo), ",", ".")
    dadosJson = dadosJson & ",""maximo"":" & Replace(CStr(maximo), ",", ".")
    dadosJson = dadosJson & "}"
    
    ObterValidador = "{""sucesso"": true, ""mensagem"": ""Validador obtido com sucesso!"", ""dados"": " & dadosJson & "}"
End Function


'--------------------------------------
'      FUNÇÕES DE VALIDAÇÃO POR NÍVEL
'--------------------------------------


Function AnaliseAnalista(idSolicitacao, flagRejeitar, faturamentoEstimadoPositivo, valorSeguradora, limiteValidado, observacao, anexosJSON, scoreSerasa, limiteSugeridoSerasa, limiteSugeridoSW, scoreSW, classeRiscoSW, classeRiscoSerasa, balancoAno )
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""
        
    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        novoStatus = "ANALISE_FINALIZADA"
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel1 = " & limiteValidado & ", "
        SQL = SQL & "ScoreSerasa = " & scoreSerasa & ", "
        SQL = SQL & "ValorSeguradora = " & valorSeguradora & ", "
        SQL = SQL & "LimiteSugeridoSerasa = " & limiteSugeridoSerasa & ", "
        SQL = SQL & "LimiteSugeridoSW = " & limiteSugeridoSW & ", "
        SQL = SQL & "ScoreSW = " & scoreSW & ", "
        SQL = SQL & "FaturamentoEstimadoPositivo = " & faturamentoEstimadoPositivo & ", "
        SQL = SQL & "ClasseRiscoSW = '" & classeRiscoSW & "', "
        SQL = SQL & "ClasseRiscoSerasa = '" & classeRiscoSerasa & "', "
        SQL = SQL & "BalancoAno = " & balancoAno & ", "
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoPrimeiroNivel = GETDATE(), "
    SQL = SQL & "AprovadorPrimeiroNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao1 = '" & Replace(observacao, "'", "''") &  "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao

    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set AnaliseAnalista = dados
End Function

Function ValidacaoAnalista(idSolicitacao, flagRejeitar, limiteValidado, observacao, clienteGrupo)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""

    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel2 = " & limiteValidado & ", "
        
        ' Se está dentro do limite do analista marca como aprovado e termina fluxo, senão continua aprovações
        If (CDbl(limiteValidado) <= CDbl(limiteMaximoAprovacaoComercial)) OR (clienteGrupo = "1") Then
            novoStatus = "APROVADO"
        Else
            novoStatus = "VALIDADO_ANALISTA"
        End If
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoSegundoNivel = GETDATE(), "
    SQL = SQL & "AprovadorSegundoNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao2 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoAnalista = dados
End Function

Function ValidacaoCoordenador(idSolicitacao, flagRejeitar, limiteValidado, observacao)   
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""

    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel3 = " & limiteValidado & ", "
        
        ' Se está dentro do limite da controladoria marca como aprovado e termina fluxo, senão continua aprovações
        If CDbl(limiteValidado) <= CDbl(limiteMaximoAprovacaoCoordenador) Then
            novoStatus = "APROVADO"
        Else
            novoStatus = "VALIDADO_COORDENADOR"
        End If
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoTerceiroNivel = GETDATE(), "
    SQL = SQL & "AprovadorTerceiroNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao3 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoCoordenador = dados
End Function

Function ValidacaoGerenteTesouraria(idSolicitacao, flagRejeitar, limiteValidado, observacao)    
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""

    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel4 = " & limiteValidado & ", "
        
        ' Se está dentro do limite do geremte de tesouraria marca como aprovado e termina fluxo, senão continua aprovações
        If CDbl(limiteValidado) <= CDbl(limiteMaximoAprovacaoGerenteTesouraria) Then
            novoStatus = "APROVADO"
        Else
            novoStatus = "VALIDADO_GERENTE_TESOURARIA"
        End If
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoQuartoNivel = GETDATE(), "
    SQL = SQL & "AprovadorQuartoNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao4 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoGerenteTesouraria = dados
End Function

Function ValidacaoController(idSolicitacao, flagRejeitar, limiteValidado, observacao)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""
    
    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel5 = " & limiteValidado & ", "
        novoStatus = "VALIDADO_CONTROLLER"
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoQuintoNivel = GETDATE(), "
    SQL = SQL & "AprovadorQuintoNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao5 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoController = dados
End Function

Function ValidacaoGerentePlanta(idSolicitacao, flagRejeitar, limiteValidado, observacao)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""

    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel6 = " & limiteValidado & ", "

        ' Se está dentro do limite do gerente de planta marca como aprovado e termina fluxo, senão continua aprovações
        If CDbl(limiteValidado) <= CDbl(limiteMaximoAprovacaoGerentePlanta) Then
            novoStatus = "APROVADO"
        Else
            novoStatus = "VALIDADO_GERENTE_PLANTA"
        End If
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoSextoNivel = GETDATE(), "
    SQL = SQL & "AprovadorSextoNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao6 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoGerentePlanta = dados
End Function

Function ValidacaoCFO(idSolicitacao, flagRejeitar, limiteValidado, observacao)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("SQL") = ""
    dados("novoStatus") = ""
    
    SQL = "UPDATE SOLCREDITOCLIENTE SET "
    
    If flagRejeitar <> "" AND flagRejeitar = 1 Then
        novoStatus = "REJEITADO"
    Else
        novoStatus = "APROVADO"
        SQL = SQL & "LCVALID = " & limiteValidado & ", "
        SQL = SQL & "LimiteValidadoNivel7 = " & limiteValidado & ", "
    End If
    
    SQL = SQL & "STATUS = '" & ObterCodStatus(novoStatus) & "', "
    SQL = SQL & "QUEMVAL = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "DTVALID = FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "DataAprovacaoSetimoNivel = GETDATE(), "
    SQL = SQL & "AprovadorSetimoNivel = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Observacao7 = '" & Replace(observacao, "'", "''") & "' "
    SQL = SQL & "WHERE NUM = " & idSolicitacao
    
    dados("SQL") = SQL
    dados("novoStatus") = novoStatus
    Set ValidacaoCFO = dados
End Function


'--------------------------------------
'         FUNÇÕES AUXILIARES
'--------------------------------------
Function UsuarioPodeValidarEtapa(status, nivelUsuario)
    Dim temPermissao
    temPermissao = True
 
    ' Se já finalizou não pode aprovar
    If status = "APROVADO" OR status = "REJEITADO" Then
        temPermissao = False
    End If
 
    If (status = "NOVO" OR status = "ANALISE_FINALIZADA") AND nivelUsuario <> "ANALISTA" Then
        temPermissao = False
    End If
 
    If status = "VALIDADO_ANALISTA" AND nivelUsuario <> "COORDENADOR" Then
        temPermissao = False
    End If
 
    If status = "VALIDADO_COORDENADOR" AND nivelUsuario <> "GERENTE_TESOURARIA" Then
        temPermissao = False
    End If
   
    If status = "VALIDADO_GERENTE_TESOURARIA" AND nivelUsuario <> "CONTROLLER" Then
        temPermissao = False
    End If
 
    If status = "VALIDADO_CONTROLLER" AND nivelUsuario <> "GERENTE_PLANTA" Then
        temPermissao = False
    End If
 
    If status = "VALIDADO_GERENTE_PLANTA" AND nivelUsuario <> "CFO" Then
        temPermissao = False
    End If
 
    UsuarioPodeValidarEtapa = temPermissao
End Function

Function VerificarSolicitacaoPendente(planta, codCliente, codLoja)
    Dim SQL, PIR_CORPORATIVO, existe
    existe = False

    SQL = "SELECT NUM FROM SOLCREDITOCLIENTE WHERE COD = '" & codCliente & "'"
    SQL = SQL & " AND Planta = '" & planta & "'"
    SQL = SQL & " AND STATUS NOT IN('" & ObterCodStatus("APROVADO") & "', '" & ObterCodStatus("REJEITADO") & "')"

    If codLoja <> "" AND planta = "07" Then
        SQL = SQL & " AND CODLOJA = '" & codLoja & "'"
    End If

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            existe = True
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    VerificarSolicitacaoPendente = existe
End Function

Function GerarProximoIdSolicitacao()
    Dim SQL, PIR_CORPORATIVO, proximoId
    proximoId = ""

    SQL = "SELECT ISNULL(MAX(CAST(NUM AS INT)), 0) + 1 AS ProximoId FROM SOLCREDITOCLIENTE"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            proximoId = CStr(PIR_CORPORATIVO("ProximoId"))
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    GerarProximoIdSolicitacao = proximoId
End Function

Function ObterSolicitacao(idSolicitacao, planta)
    SQL = "SELECT DISTINCT NUM as IdSolicitacao, " 
    SQL = SQL & "COD as CodCliente, "
    SQL = SQL & "ISNULL(CodClientePrincipal, '') AS CodClientePrincipal, "
    SQL = SQL & "CODLOJA as CodLoja, "
    SQL = SQL & "Planta as CodPlanta, "
    SQL = SQL & "LCATUAL as LimiteAtual, "
    SQL = SQL & "LCSOLICIT as LimiteSolicitado, "
    SQL = SQL & "LCNOVO as LimiteNovo, "
    SQL = SQL & "LCVALID as LimiteValidado, "
    SQL = SQL & "SOLICITANTE as NomeSolicitante, "
    SQL = SQL & "LOGININPANET as LoginSolicitante, "
    SQL = SQL & "AprovadorPrimeiroNivel as NomeAnalista, "
    SQL = SQL & "STATUS as Status, "
    SQL = SQL & "ClienteGrupo, "
    SQL = SQL & "VolumeM2, "
    SQL = SQL & "PrecoM2, "
    SQL = SQL & "PrecoKg, "
    SQL = SQL & "CondicaoPagamento, "
    SQL = SQL & "FaturamentoEstimadoPositivo, "
    SQL = SQL & "ScoreSerasa, "
    SQL = SQL & "LimiteSugeridoSerasa, "
    SQL = SQL & "LimiteSugeridoSW, "
    SQL = SQL & "ScoreSW, "
    SQL = SQL & "ClasseRiscoSW, "
    SQL = SQL & "ClasseRiscoSerasa, "
    SQL = SQL & "BalancoAno, "
    SQL = SQL & "ClienteGrupo, "
    SQL = SQL & "NomeCliente, "
    SQL = SQL & "TipoMercado, "
    SQL = SQL & "CNPJCPF, "
    SQL = SQL & "ValorSeguradora "
    SQL = SQL & "FROM SOLCREDITOCLIENTE WHERE NUM = " & idSolicitacao

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            planta = PIR_CORPORATIVO("CodPlanta")
            nomeCliente = PIR_CORPORATIVO("NomeCliente") 
            nomeSolicitante = PIR_CORPORATIVO("NomeSolicitante")
            loginSolicitante = PIR_CORPORATIVO("LoginSolicitante")
            nomeAnalista = PIR_CORPORATIVO("NomeAnalista")
            codCliente = PIR_CORPORATIVO("CodCliente")
            codClientePrincipal = PIR_CORPORATIVO("CodClientePrincipal")
            codLoja = PIR_CORPORATIVO("CodLoja")
            status = PIR_CORPORATIVO("Status")
            limiteAtual = PIR_CORPORATIVO("LimiteAtual")
            limiteSolicitado = PIR_CORPORATIVO("LimiteSolicitado")
            limiteValidado = PIR_CORPORATIVO("LimiteValidado")
            tipoMercado = PIR_CORPORATIVO("TipoMercado")
            condicaoPagamento = iif (IsNull(PIR_CORPORATIVO("CondicaoPagamento")), "", PIR_CORPORATIVO("CondicaoPagamento"))
            classeRiscoSW = iif (IsNull(PIR_CORPORATIVO("ClasseRiscoSW")), "", PIR_CORPORATIVO("ClasseRiscoSW"))
            classeRiscoSerasa = iif (IsNull(PIR_CORPORATIVO("ClasseRiscoSerasa")), "", PIR_CORPORATIVO("ClasseRiscoSerasa"))
            volumeM2 = iif (IsNull(PIR_CORPORATIVO("VolumeM2")), 0, PIR_CORPORATIVO("VolumeM2"))
            precoM2 = iif (IsNull(PIR_CORPORATIVO("PrecoM2")), 0, PIR_CORPORATIVO("PrecoM2"))
            precoKg = iif (IsNull(PIR_CORPORATIVO("PrecoKg")), 0, PIR_CORPORATIVO("PrecoKg"))
            faturamentoEstimadoPositivo = iif (IsNull(PIR_CORPORATIVO("FaturamentoEstimadoPositivo")), 0, PIR_CORPORATIVO("FaturamentoEstimadoPositivo"))
            scoreSerasa = iif (IsNull(PIR_CORPORATIVO("ScoreSerasa")), 0, PIR_CORPORATIVO("ScoreSerasa"))
            limiteSugeridoSerasa = iif (IsNull(PIR_CORPORATIVO("LimiteSugeridoSerasa")), 0, PIR_CORPORATIVO("LimiteSugeridoSerasa"))
            limiteSugeridoSW = iif (IsNull(PIR_CORPORATIVO("LimiteSugeridoSW")), 0, PIR_CORPORATIVO("LimiteSugeridoSW"))
            scoreSW = iif (IsNull(PIR_CORPORATIVO("ScoreSW")), 0, PIR_CORPORATIVO("ScoreSW"))
            balancoAno = iif (IsNull(PIR_CORPORATIVO("BalancoAno")), 0, PIR_CORPORATIVO("BalancoAno"))
            clienteGrupo = iif (IsNull(PIR_CORPORATIVO("ClienteGrupo")), 0, PIR_CORPORATIVO("ClienteGrupo"))
            cnpjCpf = iif (IsNull(PIR_CORPORATIVO("CNPJCPF")), "", PIR_CORPORATIVO("CNPJCPF"))
            valorSeguradora = PIR_CORPORATIVO("ValorSeguradora")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    Dim dadosJson
    dadosJson = "{"
    dadosJson = dadosJson & """planta"":""" & planta & """"
    dadosJson = dadosJson & ",""nomeCliente"":""" & nomeCliente & """"
    dadosJson = dadosJson & ",""clienteGrupo"":""" & clienteGrupo & """"
    dadosJson = dadosJson & ",""nomeSolicitante"":""" & nomeSolicitante & """"
    dadosJson = dadosJson & ",""loginSolicitante"":""" & loginSolicitante & """"
    dadosJson = dadosJson & ",""nomeAnalista"":""" & nomeAnalista & """"
    dadosJson = dadosJson & ",""codCliente"":""" & codCliente & """"
    dadosJson = dadosJson & ",""codClientePrincipal"":""" & codClientePrincipal & """"
    dadosJson = dadosJson & ",""codLoja"":""" & codLoja & """"
    dadosJson = dadosJson & ",""status"":""" & status & """"
    dadosJson = dadosJson & ",""tipoMercado"":""" & tipoMercado & """"
    dadosJson = dadosJson & ",""limiteAtual"":" & Replace(CStr(limiteAtual), ",", ".")
    dadosJson = dadosJson & ",""limiteSolicitado"":" & Replace(CStr(limiteSolicitado), ",", ".")
    dadosJson = dadosJson & ",""limiteValidado"":" & Replace(CStr(limiteValidado), ",", ".")
    dadosJson = dadosJson & ",""volumeM2"":" & Replace(CStr(volumeM2), ",", ".")
    dadosJson = dadosJson & ",""precoM2"":" & Replace(CStr(precoM2), ",", ".")
    dadosJson = dadosJson & ",""precoKg"":" & Replace(CStr(precoKg), ",", ".")
    dadosJson = dadosJson & ",""faturamentoEstimadoPositivo"":" & Replace(CStr(faturamentoEstimadoPositivo), ",", ".")
    dadosJson = dadosJson & ",""limiteSugeridoSerasa"":" & Replace(CStr(limiteSugeridoSerasa), ",", ".")
    dadosJson = dadosJson & ",""limiteSugeridoSW"":" & Replace(CStr(limiteSugeridoSW), ",", ".")
    dadosJson = dadosJson & ",""condicaoPagamento"":""" & condicaoPagamento & """"
    dadosJson = dadosJson & ",""scoreSerasa"":" & scoreSerasa
    dadosJson = dadosJson & ",""scoreSW"":" & scoreSW
    dadosJson = dadosJson & ",""classeRiscoSW"":""" & classeRiscoSW & """"
    dadosJson = dadosJson & ",""classeRiscoSerasa"":""" & classeRiscoSerasa & """"
    dadosJson = dadosJson & ",""balancoAno"":" & balancoAno
    dadosJson = dadosJson & ",""cnpjCpf"":""" & cnpjCpf & """"
    dadosJson = dadosJson & ",""valorSeguradora"":""" & valorSeguradora & """"
    dadosJson = dadosJson & "}"
    
    ObterSolicitacao = "{""sucesso"": true, ""mensagem"": ""Dados obtidos com sucesso!"", ""dados"": " & dadosJson & "}"
End Function

Function AtualizarLimiteCliente(idSolicitacao, planta, codCliente, codClientePrincipal, codLoja, nomeCliente, limiteValidado, limiteAtual, faturamentoEstimadoPositivo, scoreSerasa, limiteSugeridoSerasa, limiteSugeridoSW, scoreSW, classeRiscoSW, classeRiscoSerasa, balancoAno, clienteGrupo, volumeM2, precoM2, precoKg, condicaoPagamento, CNPJ, tipoMercado, valorSeguradora)
    conexao = ObterConexao(planta)

    If CNPJ <> "" AND not IsNull(CNPJ) Then
        raizCNPJ = Left(CNPJ, 8)
    Else
        raizCNPJ = ""
    End If

    SQL = ""
    Select Case conexao
          Case "PIR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = " & limiteValidado & " WHERE A1_COD = '" & codCliente & "'"
        Case "BEN-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = " & limiteValidado & " WHERE A1_CDSIADN = '" & codCliente & "'"
        Case "FOR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = " & limiteValidado & " WHERE A1_CDSIADN = '" & codCliente & "'"
        Case "SAQ-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = " & limiteValidado & " WHERE A1_COD = '" & codCliente & "' AND A1_LOJA = '" & codLoja & "'"
        Case "SPO-BI"
            nomeTabela = ""
            If Session("Ambiente") = "HOMOLOGACAO" Then
                nomeTabela = "OUT_LC_WRK_TESTE"
            Else
                nomeTabela = "OUT_LC_WRK"
            End if

            SQL = "INSERT INTO " & nomeTabela & " ( "
            SQL = SQL & "ID_SOLICITACAO, CODIGO_SAP, CODIGO_SAP_PRINCIPAL, NOME_CLIENTE, LIMITE_CREDITO_ATUAL, LIMITE_CREDITO_APROVADO, "
            SQL = SQL & "ESTADO, VOLUME_M2, PRECO_M2, PRECO_KG, CONDICAO_PAGAMENTO, "
            SQL = SQL & "FATURAMENTO_ESTIMADO_POSITIVO, SCORE_SERASA, SCORE_SW, LIMITE_SUGERIDO_SERASA, "
            SQL = SQL & "LIMITE_SUGERIDO_SW, CLASSE_RISCO_SW, CLASSE_RISCO_SERASA, BALANCO_ANO, RAIZ_CNPJ, TIPO_MERCADO, VALOR_SEGURADORA "
            SQL = SQL & ") VALUES ("
            SQL = SQL & "'" & idSolicitacao & "', "
            SQL = SQL & "'" & codCliente & "', "
            SQL = SQL & "'" & codClientePrincipal & "', "
            SQL = SQL & "'" & nomeCliente & "', "
            SQL = SQL & Replace(limiteAtual, ",", ".") & ", "
            SQL = SQL & Replace(limiteValidado, ",", ".") & ", "
            SQL = SQL & "0, "
            SQL = SQL & Replace(volumeM2, ",", ".") & ", "
            SQL = SQL & Replace(precoM2, ",", ".") & ", "
            SQL = SQL & Replace(precoKg, ",", ".") & ", "
            SQL = SQL & "'" & condicaoPagamento & "', "
            SQL = SQL & Replace(faturamentoEstimadoPositivo, ",", ".") & ", "
            SQL = SQL & Replace(scoreSerasa, ",", ".") & ", "
            SQL = SQL & Replace(scoreSW, ",", ".") & ", "
            SQL = SQL & Replace(limiteSugeridoSerasa, ",", ".") & ", "
            SQL = SQL & Replace(limiteSugeridoSW, ",", ".") & ", "
            SQL = SQL & "'" & classeRiscoSW & "', "
            SQL = SQL & "'" & classeRiscoSerasa & "', "
            SQL = SQL & "'" & balancoAno & "', "
            SQL = SQL & "'" & raizCNPJ & "', "
            SQL = SQL & "'" & tipoMercado & "', "
            SQL = SQL & "'" & valorSeguradora & "' "
            SQL = SQL & ")"
        Case Else
            SQL = ""
    End Select

    On Error Resume Next
    SET rsAtualizar = Conecta(conexao, SQL)

    If Err.Number <> 0 OR Not IsObject(rsAtualizar) Then
        AtualizarLimiteCliente = 0
        Exit Function
    End If
    On Error Goto 0
    
    AtualizarLimiteCliente = 1
End Function

Sub SalvarAnexosSolicitacao(idSolicitacao, anexosJSON)
    Dim SQL, PIR_CORPORATIVO

    On Error Goto 0

    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    
    Dim anexosArray, anexo, i
    anexosArray = Split(anexosJSON, "},{")
    
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
        
        SQL = "INSERT INTO SolicitacaoCreditoDocumento "
        SQL = SQL & "(IdSolicitacaoCredito, Nome, Caminho, DataCriacao) VALUES ("
        SQL = SQL & idSolicitacao & ", "
        SQL = SQL & "'" & nome & "', "
        SQL = SQL & "'" & caminho & "', "
        SQL = SQL & "GETDATE()"
        SQL = SQL & ")"
        
        On Error Resume Next
        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
        If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
            PIR_CORPORATIVO.Close()
            SET PIR_CORPORATIVO = NOTHING
        End If
        On Error Goto 0
    Next
End Sub

Sub RegistrarLog(acao, idSolicitacao, flagRejeitar)
    Dim SQL, LOG_CONN
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de solicitação de aumento de crédito"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a solicitação de aumento de crédito de id " & idSolicitacao
        Case "validacao"
            If flagRejeitar = 1 Then
                acaoLog = "Rejeição de solicitação de aumento de crédito"
                descricaoLog = "Usuário " & usuarioLogado("login") & " rejeitou a solicitação de aumento de crédito de id " & idSolicitacao
            Else
                acaoLog = "Validação de solicitação de aumento de crédito"
                descricaoLog = "Usuário " & usuarioLogado("login") & " validou a solicitação de aumento de crédito de id " & idSolicitacao
            End If
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Solicitação de Aumento de Crédito', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET LOG_CONN = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(LOG_CONN) Then
        LOG_CONN.Close()
        SET LOG_CONN = NOTHING
    End If
    On Error Goto 0
End Sub

Function ObterCodStatus(status) 
    Dim codStatus 
    codStatus = ""

    Select Case UCase(status)
        Case "NOVO"
            codStatus = "A"
        Case "REJEITADO"
            codStatus = "B"
        Case "VALIDADO_ANALISTA"
            codStatus = "C"
        Case "ANALISE_FINALIZADA"
            codStatus = "D"
        Case "VALIDADO_COORDENADOR"
            codStatus = "E"
        Case "VALIDADO_GERENTE_TESOURARIA"
            codStatus = "F"
        Case "VALIDADO_CONTROLLER"
            codStatus = "G"
        Case "VALIDADO_GERENTE_PLANTA"
            codStatus = "H"
        Case "APROVADO"
            codStatus = "V"
        Case Else
            codStatus = ""
    End Select
    
    ObterCodStatus = codStatus
End Function

Function ObterStatusTraduzido(status) 
    Dim codStatus 
    codStatus = ""

    Select Case UCase(status)
        Case "A"
            codStatus = "NOVO"
        Case "B"
            codStatus = "REJEITADO"
        Case "C"
            codStatus = "VALIDADO_ANALISTA"
        Case "D"
            codStatus = "ANALISE_FINALIZADA"
        Case "E"
            codStatus = "VALIDADO_COORDENADOR"
        Case "F"
            codStatus = "VALIDADO_GERENTE_TESOURARIA"
        Case "G"
            codStatus = "VALIDADO_CONTROLLER"
        Case "H"
            codStatus = "VALIDADO_GERENTE_PLANTA"
        Case "V"
            codStatus = "APROVADO"
        Case Else
            codStatus = ""
    End Select
    
    ObterStatusTraduzido = codStatus
End Function

Function ObterEmailsValidadoresProximaEtapa(statusAtual, codPlanta, loginSolicitante, nomeAnalista)
    emails = ""
    proximoNivel = ""

    Select Case UCase(statusAtual)
        Case "NOVO"
            proximoNivel = ObterCodNivel("ANALISTA")
        Case "ANALISE_FINALIZADA"
            proximoNivel = ObterCodNivel("ANALISTA")
        Case "VALIDADO_ANALISTA"
            proximoNivel = ObterCodNivel("COORDENADOR")
        Case "VALIDADO_COORDENADOR"
            proximoNivel = ObterCodNivel("GERENTE_TESOURARIA")
        Case "VALIDADO_GERENTE_TESOURARIA"
            proximoNivel = ObterCodNivel("CONTROLLER")
        Case "VALIDADO_CONTROLLER"
            proximoNivel = ObterCodNivel("GERENTE_PLANTA")
        Case "VALIDADO_GERENTE_PLANTA"
            proximoNivel = ObterCodNivel("CFO")
        Case Else
            proximoNivel = ""
    End Select

    SQL = "SELECT DISTINCT "
    SQL = SQL & " CAST(vc.IdValidador AS INT) AS IdValidador,"
    SQL = SQL & " u.CRM_NOME AS Nome,"
    SQL = SQL & " u.CRM_EMAILPRIMARIO AS Email,"
    SQL = SQL & " CAST(vc.Nivel AS NVARCHAR(10)) as CodNivel"
    SQL = SQL & " FROM ValidadoresCredito vc"
    SQL = SQL & " LEFT JOIN HELPDESK..CRM_DTI u ON vc.IdValidador = u.CRM_R_E_C_N_O"
    SQL = SQL & " WHERE vc.DataExclusao is NULL"
    SQL = SQL & " AND vc.Planta = '" & codPlanta & "' "
    SQL = SQL & " AND u.CRM_LOGIN != '" & loginSolicitante & "' "
    SQL = SQL & " AND u.CRM_NOME != '" & nomeAnalista & "' "
    SQL = SQL & " AND vc.Nivel = '" & proximoNivel & "' "
    SQL = SQL & " GROUP BY vc.IdValidador, u.CRM_NOME, u.CRM_EMAILPRIMARIO, vc.Nivel"
    SQL = SQL & " ORDER BY Nome"

    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    While NOT PIR_CORPORATIVO.EOF
        emails = emails & ";" & PIR_CORPORATIVO("Email")
        PIR_CORPORATIVO.MoveNext()
    Wend
    If isObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If

    ObterEmailsValidadoresProximaEtapa = emails
End Function

Function EnviarEmailSolicitacaoCriada(nomeCliente, codCliente, codPlanta, nomePlanta, limiteAnterior, limiteSolicitado)
    Dim SQL, PIR_HELPDESK, emailSolicitante

    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_LOGIN = '" & usuarioLogado("login") & "'"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If Not PIR_HELPDESK.EOF Then
            emailSolicitante = PIR_HELPDESK("email")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    Dim corpoHTML
    
    corpoHTML = ""
    corpoHTML = corpoHTML & "<!DOCTYPE html>"
    corpoHTML = corpoHTML & "<html>"
    corpoHTML = corpoHTML & "<head>"
    corpoHTML = corpoHTML & "<meta charset=""UTF-8"">"
    corpoHTML = corpoHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    corpoHTML = corpoHTML & "<style>"
    corpoHTML = corpoHTML & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    corpoHTML = corpoHTML & ".content { margin-bottom: 25px; }"
    corpoHTML = corpoHTML & ".info-item { margin-bottom: 0px; }"
    corpoHTML = corpoHTML & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    corpoHTML = corpoHTML & "</style>"
    corpoHTML = corpoHTML & "</head>"
    corpoHTML = corpoHTML & "<body>"
    corpoHTML = corpoHTML & "<div class=""content"">"
    corpoHTML = corpoHTML & "<p>Olá,</p>"
    corpoHTML = corpoHTML & "<p>Sua solicitação de alteração de limite de crédito foi recebida pela equipe de Crédito e está pendente de análise:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Solicitante:</strong> " & usuarioLogado("nome") & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código do Cliente:</strong> " & codCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta:</strong> " & nomePlanta & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior: </strong>" & FormatarMoeda(limiteAnterior) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Solicitado: </strong>" & FormatarMoeda(limiteSolicitado) & "</div>"
    corpoHTML = corpoHTML & "<p>Você receberá uma notificação quando o fluxo de aprovações for finalizado.</p>"
    corpoHTML = corpoHTML & "<p>Agradecemos e, em caso de dúvidas, entre em contato com o departamento de Crédito.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"

    If emailSolicitante <> "" Then
        CALL EnviaEmail(emailSolicitante, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Solicitação de Crédito Enviada")
    End If
    
    EnviarEmailSolicitacaoCriada = True
End Function

Function EnviarEmailSolicitacaoRejeitada(loginSolicitante, nomeSolicitante, nomeCliente, codCliente, codPlanta, nomePlanta, limiteAnterior, limiteSolicitado, observacao)
    Dim SQL, PIR_HELPDESK, emailSolicitante

    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_LOGIN = '" & loginSolicitante & "'"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If Not PIR_HELPDESK.EOF Then
            emailSolicitante = PIR_HELPDESK("email")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    Dim corpoHTML
    
    corpoHTML = ""
    corpoHTML = corpoHTML & "<!DOCTYPE html>"
    corpoHTML = corpoHTML & "<html>"
    corpoHTML = corpoHTML & "<head>"
    corpoHTML = corpoHTML & "<meta charset=""UTF-8"">"
    corpoHTML = corpoHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    corpoHTML = corpoHTML & "<style>"
    corpoHTML = corpoHTML & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    corpoHTML = corpoHTML & ".content { margin-bottom: 25px; }"
    corpoHTML = corpoHTML & ".info-item { margin-bottom: 0px; }"
    corpoHTML = corpoHTML & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    corpoHTML = corpoHTML & "</style>"
    corpoHTML = corpoHTML & "</head>"
    corpoHTML = corpoHTML & "<body>"
    corpoHTML = corpoHTML & "<div class=""content"">"
    corpoHTML = corpoHTML & "<p>Olá,</p>"
    corpoHTML = corpoHTML & "<p>A solicitação de alteração de limite de crédito foi <strong>rejeitada</strong>:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Solicitante:</strong> " & nomeSolicitante & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código do Cliente:</strong> " & codCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta:</strong> " & nomePlanta & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior: </strong>" & FormatarMoeda(limiteAnterior) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Solicitado: </strong>" & FormatarMoeda(limiteSolicitado) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Motivo:</strong> " & observacao & "</div>"
    corpoHTML = corpoHTML & "<p>Para mais informações, entre em contato com a equipe de suporte.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"
    
    If emailSolicitante <> "" Then
        CALL EnviaEmail(emailSolicitante, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Solicitação de Crédito Rejeitada")
    End If

    EnviarEmailSolicitacaoRejeitada = True
End Function

Function EnviarEmailSolicitacaoValidada(loginSolicitante, nomeSolicitante, nomeAnalista, statusAtual, nomeCliente, codCliente, codPlanta, nomePlanta, limiteAnterior, limiteValidado)
    Dim SQL, PIR_HELPDESK, emailSolicitante, emailDestinatarios

    emailDestinatarios = ObterEmailsValidadoresProximaEtapa(statusAtual, codPlanta, loginSolicitante, nomeAnalista)

    If emailDestinatarios <> "" Then
        Dim corpoHTML
        
        corpoHTML = ""
        corpoHTML = corpoHTML & "<!DOCTYPE html>"
        corpoHTML = corpoHTML & "<html>"
        corpoHTML = corpoHTML & "<head>"
        corpoHTML = corpoHTML & "<meta charset=""UTF-8"">"
        corpoHTML = corpoHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
        corpoHTML = corpoHTML & "<style>"
        corpoHTML = corpoHTML & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
        corpoHTML = corpoHTML & ".content { margin-bottom: 25px; }"
        corpoHTML = corpoHTML & ".info-item { margin-bottom: 0px; }"
        corpoHTML = corpoHTML & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
        corpoHTML = corpoHTML & "</style>"
        corpoHTML = corpoHTML & "</head>"
        corpoHTML = corpoHTML & "<body>"
        corpoHTML = corpoHTML & "<div class=""content"">"
        corpoHTML = corpoHTML & "<p>Olá,</p>"
        corpoHTML = corpoHTML & "<p>Existe uma solicitação pendente de validação:</p>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Solicitante:</strong> " & nomeSolicitante & "</div>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código do Cliente:</strong> " & codCliente & "</div>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta:</strong> " & nomePlanta & "</div>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior: </strong>" & FormatarMoeda(limiteAnterior) & "</div>"
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Solicitado: </strong>" & FormatarMoeda(limiteValidado) & "</div>"
        corpoHTML = corpoHTML & "</div>"
        corpoHTML = corpoHTML & "<div class=""footer"">"
        corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
        corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
        corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
        corpoHTML = corpoHTML & "</div>"
        corpoHTML = corpoHTML & "</body>"
        corpoHTML = corpoHTML & "</html>"
        
        CALL EnviaEmail(emailDestinatarios, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Aprovação de Solicitação de Crédito")
    End If
    
    EnviarEmailSolicitacaoValidada = True
End Function

Function EnviarEmailSolicitacaoAprovada(loginSolicitante, nomeSolicitante, nomeAnalista, nomeCliente, codCliente, codPlanta, nomePlanta, limiteAnterior, limiteSolicitado, limiteValidado)
    Dim SQL, PIR_HELPDESK, emailSolicitante, emailAnalista, emailDestinatarios

    ' Busca o email do solicitante
    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_LOGIN = '" & loginSolicitante & "'"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If Not PIR_HELPDESK.EOF Then
            emailSolicitante = PIR_HELPDESK("email")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    ' Busca o email do analista
    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_NOME = '" & nomeAnalista & "'"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If Not PIR_HELPDESK.EOF Then
            emailAnalista = PIR_HELPDESK("email")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    emailDestinatarios = emailSolicitante & ";" & emailAnalista

    Dim corpoHTML
    
    corpoHTML = ""
    corpoHTML = corpoHTML & "<!DOCTYPE html>"
    corpoHTML = corpoHTML & "<html>"
    corpoHTML = corpoHTML & "<head>"
    corpoHTML = corpoHTML & "<meta charset=""UTF-8"">"
    corpoHTML = corpoHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    corpoHTML = corpoHTML & "<style>"
    corpoHTML = corpoHTML & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    corpoHTML = corpoHTML & ".content { margin-bottom: 25px; }"
    corpoHTML = corpoHTML & ".info-item { margin-bottom: 0px; }"
    corpoHTML = corpoHTML & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    corpoHTML = corpoHTML & "</style>"
    corpoHTML = corpoHTML & "</head>"
    corpoHTML = corpoHTML & "<body>"
    corpoHTML = corpoHTML & "<div class=""content"">"
    corpoHTML = corpoHTML & "<p>Olá,</p>"
    corpoHTML = corpoHTML & "<p>O fluxo de aprovações do limite de crédito foi finalizado, e o seguinte limite foi aprovado:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Solicitante:</strong> " & nomeSolicitante & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código do Cliente:</strong> " & codCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta:</strong> " & nomePlanta & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior: </strong>" & FormatarMoeda(limiteAnterior) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Solicitado: </strong>" & FormatarMoeda(limiteSolicitado) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Aprovado: </strong>" & FormatarMoeda(limiteValidado) & "</div>"
    corpoHTML = corpoHTML & "<p>Agradecemos e, em caso de dúvidas, entre em contato com o departamento de Crédito.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"

    If emailDestinatarios <> "" Then
        CALL EnviaEmail(emailDestinatarios, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Solicitação de Crédito finalizada")
    End If
    
    EnviarEmailSolicitacaoAprovada = True
End Function
%>