<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../../cnn2/cnn2.asp" -->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../util/Util.asp" -->
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

Select Case acao
    Case "criarTransferencia"
        resultado = CriarTransferencia()
    Case "validarTransferencia"
        resultado = ValidarTransferencia()
    Case "obterDadosTransferencia"
        resultado = ObterDadosTransferencia()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


'--------------------------------------
'           FUNÇÕES PRINCIPAIS
'--------------------------------------

Function CriarTransferencia()
    Dim plantaOrigem, codClienteOrigem, codLojaOrigem, valorTransferencia
    Dim plantaDestino, codClienteDestino, codLojaDestino

    ' **** Dados Formulario
    valorTransferencia = Request("valorTransferencia")
    plantaOrigem = Request("plantaOrigem")
    codClienteOrigem = Request("codClienteOrigem")
    codLojaOrigem = Request("codLojaOrigem")
    plantaDestino = Request("plantaDestino")
    codClienteDestino = Request("codClienteDestino")
    codLojaDestino = Request("codLojaDestino")

    ' Validações    
    Dim respostaClienteOrigem, clienteOrigemObj
    Dim respostaClienteDestino, clienteDestinoObj

    ' Valida cliente origem
    respostaClienteOrigem = ObterCliente(plantaOrigem, codClienteOrigem, codLojaOrigem)
    Set clienteOrigemObj = JsonToDictionary(respostaClienteOrigem)

    If Not clienteOrigemObj("sucesso") Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Cliente inválido""}"
        Exit Function
    End If

    'If (plantaOrigem = "00" AND plantaDestino = "04") OR (plantaOrigem = "04" AND plantaDestino = "00") Then
        'CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""A transferência de crédito não é permitida entre essas plantas""}"
        'Exit Function
    'End If

    Dim dadosClienteOrigem
    Set dadosClienteOrigem = clienteOrigemObj("dados")

    valorTransferencia = Replace(CStr(valorTransferencia), ",", ".")
    saldoClienteOrigem = Replace(CStr(dadosClienteOrigem("saldo")), ",", ".")

    If CDbl(valorTransferencia) > CDbl(saldoClienteOrigem) Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Limite solicitado não pode ser maior do que o valor de saldo. Valor de saldo atual: " & FormatarMoeda(dadosClienteOrigem("saldo")) & """}"
        Exit Function
    End If
    
    If VerificarTransferenciaPendente(plantaOrigem, codClienteOrigem, codLojaOrigem) Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Já existe uma transferência crédito pendente para este cliente.""}"
        Exit Function
    End If

    ' Valida cliente destino
    respostaClienteDestino = ObterCliente(plantaDestino, codClienteDestino, codLojaDestino)
    Set clienteDestinoObj = JsonToDictionary(respostaClienteDestino)

    If Not clienteDestinoObj("sucesso") Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Cliente inválido""}"
        Exit Function
    End If

    Dim dadosClienteDestino, limiteAtualDestino, limiteAtualOrigem
    Set dadosClienteDestino = clienteDestinoObj("dados")

    limiteAtualOrigem = Replace(CStr(dadosClienteOrigem("limite")), ",", ".")
    limiteAtualDestino = Replace(CStr(dadosClienteDestino("limite")), ",", ".")

    SQL = "INSERT INTO TransferenciaCredito "
    SQL = SQL & "(NomeCliente, NomeSolicitante, NomeAprovador, Status, LimiteAtualOrigem, LimiteAtualDestino, ValorTransferencia, PlantaOrigem, CodClienteOrigem, CodLojaOrigem, PlantaDestino, CodClienteDestino, CodLojaDestino, Observacao, FlagInativo, CNPJCPF, DataCriacao) VALUES ("
    SQL = SQL & "'" & dadosClienteOrigem("nomeCliente") & "', "
    SQL = SQL & "'" & usuarioLogado("nome") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'PENDENTE', "
    SQL = SQL & limiteAtualOrigem & ", "
    SQL = SQL & limiteAtualDestino & ", "
    SQL = SQL & valorTransferencia & ", "
    SQL = SQL & "'" & plantaOrigem & "', "
    SQL = SQL & "'" & codClienteOrigem & "', "
    SQL = SQL & "'" & codLojaOrigem & "', "
    SQL = SQL & "'" & plantaDestino & "', "
    SQL = SQL & "'" & codClienteDestino & "', "
    SQL = SQL & "'" & codLojaDestino & "', "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "'" & dadosClienteOrigem("cnpj") & "', "
    SQL = SQL & "GETDATE())"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(PIR_CORPORATIVO) Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Erro ao criar transferência: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0

    SQL = "SELECT MAX(Id) as ultimoID FROM TransferenciaCredito WHERE NomeSolicitante = '" & usuarioLogado("nome") & "'"
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If Not PIR_CORPORATIVO.EOF Then
        idNovaSolicitacao = PIR_CORPORATIVO("ultimoID")
    End If
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING

    Call RegistrarLog("criacao", idNovaSolicitacao, 0)
    Call EnviarEmailTransferenciaCriada(dadosClienteOrigem("nomeCliente"), plantaOrigem, codClienteOrigem, codLojaOrigem, plantaDestino, codClienteDestino, codLojaDestino, valorTransferencia, usuarioLogado("login"))
    CriarTransferencia = "{""sucesso"": true, ""mensagem"": ""Transferência criada com sucesso!""}"
End Function


Function ValidarTransferencia()
    Dim idTransferencia, flagRejeitar, observacao

    ' ***** Dados recebidos
    idTransferencia = Request("idTransferencia")
    flagRejeitar = Request("flagRejeitar")
    observacao = Request("observacao")

    Dim respostaTransferencia, transferenciaObj
    respostaTransferencia = ObterTransferencia(idTransferencia)

    Set transferenciaObj = JsonToDictionary(respostaTransferencia)
    
    If Not transferenciaObj("sucesso") Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Transferência inválida""}"
        Exit Function
    End If
    
    Dim transferencia
    Set transferencia = transferenciaObj("dados")

    If transferencia("status") <> "PENDENTE" Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""A transferência já foi validada anteriormente""}"
        Exit Function
    End If
    
    ' ***** Verifica permissões
    Dim respostaValidador, validadorObj
    respostaValidador = ObterValidador(usuarioLogado("id"))

    Set validadorObj = JsonToDictionary(respostaValidador)
    
    If Not validadorObj("sucesso") Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Validador inválido""}"
        Exit Function
    End If
    
    Dim validador
    Set validador = validadorObj("dados")

    Dim nivelValidador
    nivelValidador = ObterNivelValidadorTraduzido(validador("codNivel"))

    If nivelValidador = "" Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Nível de validador inválido""}"
        Exit Function
    End If

    If nivelValidador <> "ANALISTA" Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Você não tem permissão para validar essa transferência""}"
        Exit Function
    End If

    If trim(transferencia("nomeSolicitante")) = trim(usuarioLogado("nome")) Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Usuário solicitante não pode aprovar""}"
        Exit Function
    End If

    if (transferencia("valorTransferencia") < validador("minimo")) Then
        ValidarTransferencia = "{""sucesso"": false, ""mensagem"": ""Limite validado fora do intervalo permitido para o validador""}"
        Exit Function
    End If

    ' Atualiza Registro da Transferência
    SQL = "UPDATE TransferenciaCredito SET "
    SQL = SQL & "NomeAprovador = '" & usuarioLogado("nome") & "', "
    SQL = SQL & "Status = " & iif(flagRejeitar = 1, "'REJEITADO'", "'APROVADO'") & ", "
    SQL = SQL & "Observacao = '" & observacao & "', "
    SQL = SQL & "DataValidacao = GETDATE() "
    SQL = SQL & "WHERE Id = " & idTransferencia

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(PIR_CORPORATIVO) Then
        CriarTransferencia = "{""sucesso"": false, ""mensagem"": ""Erro ao criar transferência: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
    On Error Goto 0

    If flagRejeitar = 0 Then
        Call AtualizarLimiteCliente(transferencia("plantaOrigem"), transferencia("plantaDestino"), transferencia("codClienteOrigem"), transferencia("codClienteDestino"), transferencia("codLojaOrigem"), transferencia("codLojaDestino"), transferencia("valorTransferencia")) 
        Call EnviarEmailTransferenciaAprovada(transferencia("nomeCliente"), transferencia("plantaOrigem"), transferencia("codClienteOrigem"), transferencia("codLojaOrigem"), transferencia("plantaDestino"), transferencia("codClienteDestino"), transferencia("codLojaDestino"), transferencia("limiteAtualOrigem"), transferencia("limiteAtualDestino"), transferencia("valorTransferencia"), usuarioLogado("login"))
    End If

    If flagRejeitar = 1 Then
        Call EnviarEmailTransferenciaRejeitada(transferencia("nomeCliente"), transferencia("plantaOrigem"), transferencia("codClienteOrigem"), transferencia("codLojaOrigem"), transferencia("plantaDestino"), transferencia("codClienteDestino"), transferencia("codLojaDestino"), transferencia("valorTransferencia"), observacao, usuarioLogado("login"))
    End If

    Call RegistrarLog("validacao", idTransferencia, flagRejeitar)
    ValidarTransferencia = "{""sucesso"": true, ""mensagem"": ""Transferência atualizada com sucesso!""}"
End Function


Public Function ObterCliente(plantaParam, codClienteParam, codLojaParam)
    Dim SQL, CLIENTE_CONEXAO, SALDO_CONEXAO
    Dim limite, saldo, aberto, nomeCliente, cnpj
    Dim clienteEncontrado, conexao

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
    nomeCliente = ""
    clienteEncontrado = False
    
    conexao = ObterConexao(plantaParam)
    SQL = "SELECT Nome, DescricaoCondPag as CondicaoPagamento, CnpjCpf as CNPJ "
    SQL = SQL & "FROM vwCliente WHERE Codigo = '" & codClienteParam & "' "
    
    If codLojaParam <> "" Then
        SQL = SQL & "AND Loja = '" & codLojaParam & "'"
    End If

    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexao, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        If Not CLIENTE_CONEXAO.EOF Then
            limite = iif(Not IsNull(CLIENTE_CONEXAO("Limite")), CLIENTE_CONEXAO("Limite"), 0)
            cnpj = iif(Not IsNull(CLIENTE_CONEXAO("CNPJ")), CLIENTE_CONEXAO("CNPJ"), "")
            nomeCliente = CLIENTE_CONEXAO("Nome")
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
    
    ' Query para buscar dados de saldo 
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
            aberto = SALDO_CONEXAO("Aberto")
            saldo = SALDO_CONEXAO("Saldo")
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
    
    Dim dadosJson
    dadosJson = "{"
    dadosJson = dadosJson & """nomeCliente"":""" & EscapeJSON(nomeCliente) & """"
    dadosJson = dadosJson & ",""cnpj"":""" & EscapeJSON(cnpj) & """"
    dadosJson = dadosJson & ",""limite"":" & Replace(CStr(limite), ",", ".")
    dadosJson = dadosJson & ",""aberto"":" & Replace(CStr(aberto), ",", ".")
    dadosJson = dadosJson & ",""saldo"":" & Replace(CStr(saldo), ",", ".")
    dadosJson = dadosJson & "}"
    
    ObterCliente = "{""sucesso"": true, ""mensagem"": ""Dados obtidos com sucesso!"", ""dados"": " & dadosJson & "}"
End Function

Public Function ObterDadosTransferencia()
    Dim plantaOrigem, plantaDestino, codClienteOrigem, codLojaParam
    Dim limiteOrigem, saldoOrigem, abertoOrigem
    Dim nomeCliente, cnpjCpf, conexaoOrigem, conexaoDestino
    Dim clienteOrigemEncontrado
    Dim clientesDestinoJson, primeiroCliente

    plantaOrigem = Request("plantaOrigem")
    plantaDestino = Request("plantaDestino")
    codClienteOrigem = Request("codClienteOrigem")
    codLojaOrigem = Iif(Request("codLojaOrigem") = "", "",  Right("0" & Request("codLojaOrigem"), 2)) 
    
    If plantaOrigem = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Planta de origem não informada""}"
        Exit Function
    End If
    
    If plantaDestino = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Planta de destino não informada""}"
        Exit Function
    End If
    
    If codClienteOrigem = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Código de cliente não informado""}"
        Exit Function
    End If
    
    If codLojaOrigem = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Código de loja não informado""}"
        Exit Function
    End If
    
    limiteOrigem = 0
    saldoOrigem = 0
    abertoOrigem = 0
    nomeCliente = ""
    CNPJ = ""
    baseCNPJ = ""
    clienteOrigemEncontrado = False
    clientesDestinoJson = ""
    primeiroCliente = True
    
    ' Buscando os dados na planta de origem
    conexaoOrigem = ObterConexao(plantaOrigem)

    SQL = "SELECT Nome, CnpjCpf FROM vwCliente WHERE Codigo = '" & codClienteOrigem & "'"
    SQL = SQL & " AND Loja = '" & codLojaOrigem & "'"
    
    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexaoOrigem, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        If Not CLIENTE_CONEXAO.EOF Then
            limiteOrigem = CLIENTE_CONEXAO("Limite")
            nomeCliente = CLIENTE_CONEXAO("Nome")
            CNPJ = CLIENTE_CONEXAO("CnpjCpf")
            baseCNPJ = Left(CLIENTE_CONEXAO("CnpjCpf"), 8)
            clienteOrigemEncontrado = True
        End If
        CLIENTE_CONEXAO.Close()
        SET CLIENTE_CONEXAO = NOTHING
    End If
    On Error Goto 0
    
    If Not clienteOrigemEncontrado Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Cliente não encontrado na planta de origem!""}"
        Exit Function
    End If
    
    If baseCNPJ = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""CNPJ/CPF não encontrado para o cliente na origem!""}"
        Exit Function
    End If
    
    SQL = "SELECT * FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA("
    SQL = SQL & "'" & codClienteOrigem & "'"
    
    If codLojaOrigem <> "" AND plantaOrigem = "07" Then 
        SQL = SQL & ", '" & codLojaOrigem & "'"
    End If
    
    SQL = SQL & ")"
    
    On Error Resume Next
    Err.Clear
    SET SALDO_CONEXAO = Conecta(conexaoOrigem, SQL)
    
    If Err.Number = 0 AND IsObject(SALDO_CONEXAO) Then
        If Not SALDO_CONEXAO.EOF Then
            limiteOrigem = SALDO_CONEXAO("Limite")
            abertoOrigem = SALDO_CONEXAO("Aberto")
            saldoOrigem = SALDO_CONEXAO("Saldo")
        Else
            SALDO_CONEXAO.Close()
            SET SALDO_CONEXAO = NOTHING
            ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Não foi possível obter os dados de saldo do cliente na origem!""}"
            Exit Function
        End If
        SALDO_CONEXAO.Close()
        SET SALDO_CONEXAO = NOTHING
    Else
        If IsObject(SALDO_CONEXAO) Then
            SALDO_CONEXAO.Close()
            SET SALDO_CONEXAO = NOTHING
        End If
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Erro ao obter dados de saldo do cliente na origem!""}"
        Exit Function
    End If
    On Error Goto 0
    
    ' Buscando todos os clientes de destino com mesmo CNPJ
    conexaoDestino = ObterConexao(plantaDestino)
    SQL = "SELECT Nome, Loja, Codigo, CnpjCpf, Estado FROM vwCliente WHERE CnpjCpf LIKE '%" & baseCNPJ & "%'"

    If plantaOrigem = plantaDestino OR (plantaOrigem = "00" AND plantaDestino = "04") Then
        SQL = SQL & " AND CnpjCpf <> '" & CNPJ & "'"
    End If

    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexaoDestino, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        If Not CLIENTE_CONEXAO.EOF Then
            Do While Not CLIENTE_CONEXAO.EOF
                limiteDestino = 0
                abertoDestino = 0
                saldoDestino = 0
                nomeClienteDestino = CLIENTE_CONEXAO("Nome")
                codClienteDestino = CLIENTE_CONEXAO("Codigo")                
                codLojaDestino = CLIENTE_CONEXAO("Loja")
                cnpjClienteDestino = CLIENTE_CONEXAO("CnpjCpf")
                estadoClienteDestino = CLIENTE_CONEXAO("Estado")
                
                SQL = "SELECT * FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA("
                SQL = SQL & "'" & codClienteDestino & "'"
                
                If codLojaDestino <> "" AND plantaDestino = "07" Then 
                    SQL = SQL & ", '" & codLojaDestino & "'"
                End If
                
                SQL = SQL & ")"
                
                On Error Resume Next
                Err.Clear
                SET SALDO_CONEXAO = Conecta(conexaoDestino, SQL)
                
                If Err.Number = 0 AND IsObject(SALDO_CONEXAO) Then
                    If Not SALDO_CONEXAO.EOF Then
                        limiteDestino = SALDO_CONEXAO("Limite")
                        abertoDestino = SALDO_CONEXAO("Aberto")
                        saldoDestino = SALDO_CONEXAO("Saldo")
                    End If
                    SALDO_CONEXAO.Close()
                    SET SALDO_CONEXAO = NOTHING
                Else
                    If IsObject(SALDO_CONEXAO) Then
                        SALDO_CONEXAO.Close()
                        SET SALDO_CONEXAO = NOTHING
                    End If
                End If
                On Error Goto 0
                
                If Not primeiroCliente Then
                    clientesDestinoJson = clientesDestinoJson & ","
                End If
                
                clientesDestinoJson = clientesDestinoJson & "{"

                clientesDestinoJson = clientesDestinoJson & """limiteAtual"":" & Replace(CStr(limiteDestino), ",", ".")
                clientesDestinoJson = clientesDestinoJson & ",""aberto"":" & Replace(CStr(abertoDestino), ",", ".")
                clientesDestinoJson = clientesDestinoJson & ",""saldo"":" & Replace(CStr(saldoDestino), ",", ".")
                clientesDestinoJson = clientesDestinoJson & ",""codigo"":""" & EscapeJSON(codClienteDestino) & """"
                clientesDestinoJson = clientesDestinoJson & ",""loja"":""" & EscapeJSON(codLojaDestino) & """"
                clientesDestinoJson = clientesDestinoJson & ",""nome"":""" & EscapeJSON(nomeClienteDestino) & """"
                clientesDestinoJson = clientesDestinoJson & ",""cnpj"":""" & EscapeJSON(cnpjClienteDestino) & """"
                clientesDestinoJson = clientesDestinoJson & ",""estado"":""" & EscapeJSON(estadoClienteDestino) & """"
                clientesDestinoJson = clientesDestinoJson & "}"
                
                primeiroCliente = False
                CLIENTE_CONEXAO.MoveNext
            Loop
        End If
        CLIENTE_CONEXAO.Close()
        SET CLIENTE_CONEXAO = NOTHING
    End If
    On Error Goto 0
    
    If clientesDestinoJson = "" Then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Cliente não existe na planta de destino!""}"
        Exit Function
    End If
    
    Dim dadosJson
    dadosJson = "{"
    dadosJson = dadosJson & """nomeCliente"":""" & EscapeJSON(nomeCliente) & """"
    dadosJson = dadosJson & ",""limiteAtual"":" & Replace(CStr(limiteOrigem), ",", ".")
    dadosJson = dadosJson & ",""aberto"":" & Replace(CStr(abertoOrigem), ",", ".")
    dadosJson = dadosJson & ",""saldo"":" & Replace(CStr(saldoOrigem), ",", ".")
    dadosJson = dadosJson & ",""clientesDestino"":[" & clientesDestinoJson & "]"
    dadosJson = dadosJson & "}"
    
    ObterDadosTransferencia = "{""sucesso"": true, ""mensagem"": ""Dados de transferência obtidos com sucesso!"", ""dados"": " & dadosJson & "}"
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
'         FUNÇÕES AUXILIARES
'--------------------------------------

Function VerificarTransferenciaPendente(plantaOrigem, codClienteOrigem, codLojaOrigem)
    Dim SQL, PIR_CORPORATIVO, existe
    existe = False

    SQL = "SELECT Id FROM TransferenciaCredito WHERE CodClienteOrigem = '" & codClienteOrigem & "'"
    SQL = SQL & " AND PlantaOrigem = '" & plantaOrigem & "'"
    SQL = SQL & " AND Status = 'PENDENTE'"

    If codLojaOrigem <> "" Then
        SQL = SQL & " AND CodClienteOrigem = '" & codLojaOrigem & "'"
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

    VerificarTransferenciaPendente = existe
End Function


Function ObterTransferencia(idTransferencia)
    Dim SQL, PIR_CORPORATIVO
    Dim nomeCliente, nomeSolicitante, nomeAprovador, status
    Dim limiteAtual, valorTransferencia, plantaOrigem, codClienteOrigem, codLojaOrigem
    Dim plantaDestino, codClienteDestino, codLojaDestino, observacao, dataCriacao

    SQL = "SELECT "
    SQL = SQL & "Id, NomeCliente, NomeSolicitante, NomeAprovador, "
    SQL = SQL & "Status, LimiteAtualOrigem, LimiteAtualDestino, ValorTransferencia, PlantaOrigem, "
    SQL = SQL & "CodClienteOrigem, CodLojaOrigem, PlantaDestino, CodClienteDestino, "
    SQL = SQL & "CodLojaDestino, Observacao, FlagInativo, "
    SQL = SQL & "CASE "
    SQL = SQL & "    WHEN ISDATE(DataCriacao) = 1 "
    SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, DataCriacao, 120), 126) "
    SQL = SQL & "    ELSE NULL "
    SQL = SQL & "END AS DataCriacao "
    SQL = SQL & "FROM TransferenciaCredito "
    SQL = SQL & "WHERE FlagInativo = 0 "
    SQL = SQL & "AND Id = " & idTransferencia

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            nomeCliente = PIR_CORPORATIVO("NomeCliente")
            nomeSolicitante = PIR_CORPORATIVO("NomeSolicitante")
            nomeAprovador = PIR_CORPORATIVO("NomeAprovador")
            status = PIR_CORPORATIVO("Status")
            limiteAtualOrigem = PIR_CORPORATIVO("LimiteAtualOrigem")
            limiteAtualDestino = PIR_CORPORATIVO("LimiteAtualDestino")
            valorTransferencia = PIR_CORPORATIVO("ValorTransferencia")
            plantaOrigem = PIR_CORPORATIVO("PlantaOrigem")
            codClienteOrigem = PIR_CORPORATIVO("CodClienteOrigem")
            codLojaOrigem = PIR_CORPORATIVO("CodLojaOrigem")
            plantaDestino = PIR_CORPORATIVO("PlantaDestino")
            codClienteDestino = PIR_CORPORATIVO("CodClienteDestino")
            codLojaDestino = PIR_CORPORATIVO("CodLojaDestino")
            observacao = PIR_CORPORATIVO("Observacao")    
            dataCriacao = PIR_CORPORATIVO("DataCriacao")  
        End If    
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    Dim dadosJson
    dadosJson = "{"
    dadosJson = dadosJson & """nomeCliente"":""" & nomeCliente & """"
    dadosJson = dadosJson & ",""nomeSolicitante"":""" & nomeSolicitante & """"
    dadosJson = dadosJson & ",""nomeAprovador"":""" & nomeAprovador & """"
    dadosJson = dadosJson & ",""status"":""" & status & """"
    dadosJson = dadosJson & ",""limiteAtualOrigem"":" & Replace(CStr(limiteAtualOrigem), ",", ".")
    dadosJson = dadosJson & ",""limiteAtualDestino"":" & Replace(CStr(limiteAtualDestino), ",", ".")
    dadosJson = dadosJson & ",""valorTransferencia"":" & Replace(CStr(valorTransferencia), ",", ".")
    dadosJson = dadosJson & ",""plantaOrigem"":""" & plantaOrigem & """"
    dadosJson = dadosJson & ",""codClienteOrigem"":""" & codClienteOrigem & """"
    dadosJson = dadosJson & ",""codLojaOrigem"":""" & codLojaOrigem & """"
    dadosJson = dadosJson & ",""plantaDestino"":""" & plantaDestino & """"
    dadosJson = dadosJson & ",""codClienteDestino"":""" & codClienteDestino & """"
    dadosJson = dadosJson & ",""codLojaDestino"":""" & codLojaDestino & """"
    dadosJson = dadosJson & ",""observacao"":""" & observacao & """"
    dadosJson = dadosJson & ",""dataCriacao"":""" & dataCriacao & """"
    dadosJson = dadosJson & "}"
    
    ObterTransferencia = "{""sucesso"": true, ""mensagem"": ""Dados obtidos com sucesso!"", ""dados"": " & dadosJson & "}"
End Function

Sub AtualizarLimiteCliente(plantaOrigem, plantaDestino, codClienteOrigem, codClienteDestino, codLojaOrigem, codLojaDestino, valorTransferencia)
    Dim SQL, CLIENTE_CONEXAO, conexao
    valorTransferencia = Replace(CStr(valorTransferencia), ",", ".")

    ' Atualiza Limite do Cliente de Origem (reduzindo o limite)
    conexao = ObterConexao(plantaOrigem)
    
    SQL = ""
    Select Case conexao
        Case "PIR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC - " & valorTransferencia & " WHERE A1_COD = '" & codClienteOrigem & "'"
        Case "BEN-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC - " & valorTransferencia & " WHERE A1_CDSIADN = '" & codClienteOrigem & "'"
        Case "FOR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC - " & valorTransferencia & " WHERE A1_CDSIADN = '" & codClienteOrigem & "'"
        Case "SAQ-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC - " & valorTransferencia & " WHERE A1_COD = '" & codClienteOrigem & "' AND A1_LOJA = '" & codLojaOrigem & "'"
        Case Else
            SQL = ""
    End Select
    
    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexao, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        CLIENTE_CONEXAO.Close()
        SET CLIENTE_CONEXAO = NOTHING
    End If
    On Error Goto 0
    
    ' Atualiza Limite do Cliente na planta de destino (somando o limite)
    conexao = ObterConexao(plantaDestino)

    SQL = ""
    Select Case conexao
        Case "PIR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC + " & valorTransferencia & " WHERE A1_COD = '" & codClienteDestino & "'"
        Case "BEN-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC + " & valorTransferencia & " WHERE A1_CDSIADN = '" & codClienteDestino & "'"
        Case "FOR-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC + " & valorTransferencia & " WHERE A1_CDSIADN = '" & codClienteDestino & "'"
        Case "SAQ-DADOSADV"
            SQL = "UPDATE SA1010 SET A1_LC = A1_LC + " & valorTransferencia & " WHERE A1_COD = '" & codClienteDestino & "' AND A1_LOJA = '" & codLojaDestino & "'"
        Case Else
            SQL = ""
    End Select
    
    On Error Resume Next
    SET CLIENTE_CONEXAO = Conecta(conexao, SQL)
    
    If Err.Number = 0 AND IsObject(CLIENTE_CONEXAO) Then
        CLIENTE_CONEXAO.Close()
        SET CLIENTE_CONEXAO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub RegistrarLog(acao, idTransferencia, flagRejeitar)
    Dim SQL, LOG_CONN
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de transferência de crédito"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a transferência de crédito de id " & idTransferencia
        Case "validacao"
            If flagRejeitar = 1 Then
                acaoLog = "Rejeição de transferência de crédito"
                descricaoLog = "Usuário " & usuarioLogado("login") & " rejeitou a transferência de crédito de id " & idTransferencia
            Else
                acaoLog = "Validação de transferência de crédito"
                descricaoLog = "Usuário " & usuarioLogado("login") & " validou a transferência de crédito de id " & idTransferencia
            End If
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Transferência de Crédito', "
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

Function ObterEmailSolicitante(loginSolicitante)
    Dim SQL, PIR_HELPDESK, emailRetorno
    emailRetorno = ""

    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_LOGIN = '" & loginSolicitante & "'"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If Not PIR_HELPDESK.EOF Then
            emailRetorno = PIR_HELPDESK("email")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0
    
    ObterEmailSolicitante = emailRetorno
End Function


Function EnviarEmailTransferenciaCriada(nomeCliente, plantaOrigem, codClienteOrigem, codLojaOrigem, plantaDestino, codClienteDestino, codLojaDestino, valorTransferencia, loginSolicitante)
    Dim emailSolicitante, corpoHTML

    emailSolicitante = ObterEmailSolicitante(loginSolicitante)
    
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
    corpoHTML = corpoHTML & "<p>Uma solicitação de transferência de limite de crédito foi recebida pela equipe de suporte e está pendente de aprovação:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Solicitante:</strong> " & usuarioLogado("nome") & "</div>"
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Origem:</strong> " & ObterNomePlanta(plantaOrigem) & " - Cliente: " & codClienteOrigem
    
    If codLojaOrigem <> "" Then
        corpoHTML = corpoHTML & " / Loja: " & codLojaOrigem
    End If
    
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Destino:</strong> " & ObterNomePlanta(plantaDestino) & " - Cliente: " & codClienteDestino
    
    If codLojaDestino <> "" Then
        corpoHTML = corpoHTML & " / Loja: " & codLojaDestino
    End If
    
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Valor de transferência:</strong> " & FormatarMoeda(valorTransferencia) & "</div>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"
    
    CALL EnviaEmail(emailSolicitante, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Transferência de Crédito Criada")
    EnviarEmailTransferenciaCriada = True
End Function

Function EnviarEmailTransferenciaAprovada(nomeCliente, plantaOrigem, codClienteOrigem, codLojaOrigem, plantaDestino, codClienteDestino, codLojaDestino, limiteAtualOrigem, limiteAtualDestino, valorTransferencia, loginSolicitante)
    Dim emailSolicitante, corpoHTML
    Dim novoLimiteOrigem, novoLimiteDestino

    emailSolicitante = ObterEmailSolicitante(loginSolicitante)
    
    novoLimiteOrigem = CDbl(limiteAtualOrigem) - CDbl(valorTransferencia)
    novoLimiteDestino = CDbl(limiteAtualDestino) + CDbl(valorTransferencia)
    
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
    corpoHTML = corpoHTML & "<p>A transferência de crédito foi <strong>aprovada</strong> com sucesso:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Aprovador:</strong> " & usuarioLogado("nome") & "</div>"
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Origem:</strong> " & ObterNomePlanta(plantaOrigem) & " - Cliente: " & codClienteOrigem
    
    If codLojaOrigem <> "" Then
        corpoHTML = corpoHTML & " / Loja: " & codLojaOrigem
    End If
    
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior:</strong> " & FormatarMoeda(limiteAtualOrigem) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Novo Limite:</strong> " & FormatarMoeda(novoLimiteOrigem) & "</div>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Destino:</strong> " & ObterNomePlanta(plantaDestino) & " - Cliente: " & codClienteDestino
    
    If codLojaDestino <> "" Then
        corpoHTML = corpoHTML & " / Loja: " & codLojaDestino
    End If
    
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Limite Anterior:</strong> " & FormatarMoeda(limiteAtualDestino) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Novo Limite:</strong> " & FormatarMoeda(novoLimiteDestino) & "</div>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Valor Transferido:</strong> " & FormatarMoeda(valorTransferencia) & "</div>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"
    
    CALL EnviaEmail(emailSolicitante, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Transferência de Crédito Aprovada")
    EnviarEmailTransferenciaAprovada = True
End Function

Function EnviarEmailTransferenciaRejeitada(nomeCliente, plantaOrigem, codClienteOrigem, codLojaOrigem, plantaDestino, codClienteDestino, codLojaDestino, valorTransferencia, observacao, loginSolicitante)
    Dim emailSolicitante, corpoHTML

    emailSolicitante = ObterEmailSolicitante(loginSolicitante)

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
    corpoHTML = corpoHTML & ".warning { background-color: #ffebee; padding: 10px; border-radius: 5px; margin: 15px 0; }"
    corpoHTML = corpoHTML & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    corpoHTML = corpoHTML & "</style>"
    corpoHTML = corpoHTML & "</head>"
    corpoHTML = corpoHTML & "<body>"
    corpoHTML = corpoHTML & "<div class=""content"">"
    corpoHTML = corpoHTML & "<p>Olá,</p>"
    corpoHTML = corpoHTML & "<p>A solicitação de transferência de crédito foi <strong>rejeitada</strong>:</p>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Cliente:</strong> " & nomeCliente & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Aprovador:</strong> " & usuarioLogado("nome") & "</div>"
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta Origem:</strong> " & ObterNomePlanta(plantaOrigem) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código Cliente Origem:</strong> " & codClienteOrigem & "</div>"
    
    If codLojaOrigem <> "" Then
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código Loja Origem:</strong> " & codLojaOrigem & "</div>"
    End If
    
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Planta Destino:</strong> " & ObterNomePlanta(plantaDestino) & "</div>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código Cliente Destino:</strong> " & codClienteDestino & "</div>"
    
    If codLojaDestino <> "" Then
        corpoHTML = corpoHTML & "<div class=""info-item""><strong>Código Loja Destino:</strong> " & codLojaDestino & "</div>"
    End If
    
    corpoHTML = corpoHTML & "<br>"
    corpoHTML = corpoHTML & "<div class=""info-item""><strong>Valor Solicitado:</strong> " & FormatarMoeda(valorTransferencia) & "</div>"
    
    corpoHTML = corpoHTML & "<p>Para mais informações, entre em contato com a equipe de suporte.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "<div class=""footer"">"
    corpoHTML = corpoHTML & "<p><strong>Smurfit Westrock</strong></p>"
    corpoHTML = corpoHTML & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    corpoHTML = corpoHTML & "<p>Este é um email automático, por favor não responda.</p>"
    corpoHTML = corpoHTML & "</div>"
    corpoHTML = corpoHTML & "</body>"
    corpoHTML = corpoHTML & "</html>"
    
    CALL EnviaEmail(emailSolicitante, "", "", "", "Sistema de Automatização Smurfit Westrock <inpanet.inpa@smurfitkappa.com.br>", corpoHTML, "Transferência de Crédito Rejeitada")
    EnviarEmailTransferenciaRejeitada = True
End Function
%>