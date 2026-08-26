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
Dim acao
Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "planta", SESSION("FILIAL")

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listar"
        resultado = Listar()
    Case "criar"
        resultado = Criar()
    Case "editar"
        resultado = Editar()
    Case "excluir"
        resultado = Excluir()
    Case "obterPorId"
        resultado = ObterPorId()
    Case "listarStatus"
        resultado = ListarStatus()
    Case "obterDadosUsuario"
        resultado = ObterDadosUsuario()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    id = Request("id")
    idFNC = Request("idFNC")
    codPlanta = Request("codPlanta")
    dataCriacaoDe = Request("dataCriacaoDe")
    dataCriacaoAte = Request("dataCriacaoAte")
    status = Request("status")
    nomeCliente = Request("nomeCliente")
    ordemProducao = Request("ordemProducao")
    NFDevolucao = Request("NFDevolucao")

    setorUsuario = ""
    plantasUsuario = ""

    respostaSetorUsuario = ObterDadosUsuario()
    Set respostaObj = JsonToDictionary(respostaSetorUsuario)
    
    If Not respostaObj("sucesso") Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário""}"
        Exit Function
    End If

    Set dadosUsuario = respostaObj("dados")
    setorUsuario = dadosUsuario("setor")
    plantasVisiveis = dadosUsuario("plantas")

    If plantasVisiveis = "" Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Nível de usuário inválido""}"
        Exit Function
    End If

    'MONTA clausula planta
    clausulaPlanta = ""
    primeiroItem = True
    arrayPlantas = Split(plantasVisiveis, ",")

    clausulaPlanta = clausulaPlanta & " AND CodPlanta in ("
    For Each item In arrayPlantas
        If Not primeiroItem Then 
            clausulaPlanta = clausulaPlanta & "," 
        End If

        clausulaPlanta = clausulaPlanta & "'" & trim(item) & "'"
        primeiroItem = False
    Next

    clausulaPlanta = clausulaPlanta & ")"
       
    SQL = "SELECT "
    SQL = SQL & "   Id, "
    SQL = SQL & "   Status, "
    SQL = SQL & "   CodPlanta, "
    SQL = SQL & "   IdFNC, "
    SQL = SQL & "   QuantidadeRecebida, "
    SQL = SQL & "   QuantidadeRefugada, "
    SQL = SQL & "   NFDevolucao, "
    SQL = SQL & "   ISNULL(Valor, 0) as Valor, "
    SQL = SQL & "   ISNULL(ValorUnitarioMaterial, 0)as ValorUnitarioMaterial, "
    SQL = SQL & "   CodigoVA, "
    SQL = SQL & "   CodigoQM, "
    SQL = SQL & "   FlagInativo, "
    SQL = SQL & "   DataEntradaDevolucao, "
    SQL = SQL & "   DataCriacao, "
    SQL = SQL & "   DataConclusao, "
    SQL = SQL & "   QuantidadePaletes, "
    SQL = SQL & "   NomeCliente, "
    SQL = SQL & "   ModoDefeito, "
    SQL = SQL & "   Area, "
    SQL = SQL & "   OrdemProducao "
    SQL = SQL & "FROM FichaQM "
    SQL = SQL & "WHERE FlagInativo = 0 OR FlagInativo IS NULL "
    SQL = SQL & clausulaPlanta

    If id <> "" Then 
        SQL = SQL & " AND Id = '" & id & "'"
    Else 
        If codPlanta <> "" Then
            SQL = SQL & " AND CodPlanta = '" & codPlanta & "'"
        End If

        If idFNC <> "" Then
            SQL = SQL & " AND IdFNC = " & idFNC
        End If
        
        If dataCriacaoDe <> "" Then
            SQL = SQL & " AND DataCriacao >= '" & Replace(dataCriacaoDe, "-", "") & "'"
        End If

        If dataCriacaoAte <> "" Then
            SQL = SQL & " AND DataCriacao < DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "')"
        End If

        If status <> "" Then
            SQL = SQL & " AND Status = '" & status & "' "
        End If

        If nomeCliente <> "" Then
            SQL = SQL & " AND NomeCliente = '" & nomeCliente & "' "
        End If

        If ordemProducao <> "" Then
            SQL = SQL & " AND OrdemProducao = '" & ordemProducao & "' "
        End If

        If NFDevolucao <> "" Then
            SQL = SQL & " AND NFDevolucao = '" & NFDevolucao & "' "
        End If
    End If

    SQL = SQL & "Order by Id DESC"

    On Error Resume Next
    SET rsListar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de fichas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id") & "") & """, "
            jsonArray = jsonArray & """codPlanta"": """ & EscapeJSON(rsListar("CodPlanta") & "") & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status") & "") & """, "
            jsonArray = jsonArray & """idFNC"": """ & EscapeJSON(rsListar("IdFNC") & "") & """, "
            jsonArray = jsonArray & """quantidadeRecebida"": """ & EscapeJSON(rsListar("QuantidadeRecebida") & "") & """, "
            jsonArray = jsonArray & """quantidadeRefugada"": """ & EscapeJSON(rsListar("QuantidadeRefugada") & "") & """, "
            jsonArray = jsonArray & """NFDevolucao"": """ & EscapeJSON(rsListar("NFDevolucao") & "") & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(Replace(CStr(rsListar("Valor")), ",", ".")) & """, "
            jsonArray = jsonArray & """valorUnitarioMaterial"": """ & EscapeJSON(Replace(CStr(rsListar("ValorUnitarioMaterial")), ",", ".")) & """, "
            jsonArray = jsonArray & """codigoVA"": """ & EscapeJSON(rsListar("CodigoVA") & "") & """, "
            jsonArray = jsonArray & """codigoQM"": """ & EscapeJSON(rsListar("CodigoQM") & "") & """, "
            jsonArray = jsonArray & """dataEntradaDevolucao"": """ & EscapeJSON(rsListar("DataEntradaDevolucao") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """, "
            jsonArray = jsonArray & """dataConclusao"": """ & EscapeJSON(rsListar("DataConclusao") & "") & """, "
            jsonArray = jsonArray & """quantidadePaletes"": """ & EscapeJSON(rsListar("QuantidadePaletes") & "") & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsListar("NomeCliente") & "") & """, "
            jsonArray = jsonArray & """modoDefeito"": """ & EscapeJSON(rsListar("ModoDefeito") & "") & """, "
            jsonArray = jsonArray & """area"": """ & EscapeJSON(rsListar("Area") & "") & """, "
            jsonArray = jsonArray & """ordemProducao"": """ & EscapeJSON(rsListar("OrdemProducao") & "") & """ "
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
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If
    
    SQL = "SELECT TOP 1 "
    SQL = SQL & "   Id, "
    SQL = SQL & "   Status, "
    SQL = SQL & "   CodPlanta, "
    SQL = SQL & "   IdFNC, "
    SQL = SQL & "   QuantidadeRecebida, "
    SQL = SQL & "   QuantidadeRefugada, "
    SQL = SQL & "   NFDevolucao, "
    SQL = SQL & "   ISNULL(Valor, 0) as Valor, "
    SQL = SQL & "   ISNULL(ValorUnitarioMaterial, 0)as ValorUnitarioMaterial, "
    SQL = SQL & "   CodigoVA, "
    SQL = SQL & "   CodigoQM, "
    SQL = SQL & "   FlagInativo, "
    SQL = SQL & "   DataEntradaDevolucao, "
    SQL = SQL & "   FORMAT(DataCriacao,'yyyy-MM-dd') as DataCriacao, "
    SQL = SQL & "   FORMAT(DataConclusao,'yyyy-MM-dd') as DataConclusao, "
    SQL = SQL & "   QuantidadePaletes, "
    SQL = SQL & "   NomeCliente, "
    SQL = SQL & "   ModoDefeito, "
    SQL = SQL & "   Area, "
    SQL = SQL & "   OrdemProducao "
    SQL = SQL & "FROM FichaQM "
    SQL = SQL & "WHERE Id = '" & id & "' "
    SQL = SQL & "AND (FlagInativo = 0 OR FlagInativo IS NULL)"
    
    On Error Resume Next
    SET rsObter = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar ficha: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Ficha não encontrada.""}"
        Exit Function
    End If
    
    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id") & "") & """, "
    jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta") & "") & """, "
    jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status") & "") & """, "
    jsonObjeto = jsonObjeto & """idFNC"": """ & EscapeJSON(rsObter("IdFNC") & "") & """, "
    jsonObjeto = jsonObjeto & """quantidadeRecebida"": """ & EscapeJSON(rsObter("QuantidadeRecebida") & "") & """, "
    jsonObjeto = jsonObjeto & """quantidadeRefugada"": """ & EscapeJSON(rsObter("QuantidadeRefugada") & "") & """, "
    jsonObjeto = jsonObjeto & """NFDevolucao"": """ & EscapeJSON(rsObter("NFDevolucao") & "") & """, "
    jsonObjeto = jsonObjeto & """valor"": """ & EscapeJSON(Replace(CStr(rsObter("Valor")), ",", ".")) & """, "
    jsonObjeto = jsonObjeto & """valorUnitarioMaterial"": """ & EscapeJSON(Replace(CStr(rsObter("ValorUnitarioMaterial")), ",", ".")) & """, "
    jsonObjeto = jsonObjeto & """codigoVA"": """ & EscapeJSON(rsObter("CodigoVA") & "") & """, "
    jsonObjeto = jsonObjeto & """codigoQM"": """ & EscapeJSON(rsObter("CodigoQM") & "") & """, "
    jsonObjeto = jsonObjeto & """dataEntradaDevolucao"": """ & EscapeJSON(formataDataUS(rsObter("DataEntradaDevolucao")) & "") & """, "
    jsonObjeto = jsonObjeto & """dataCriacao"": """ & EscapeJSON(rsObter("DataCriacao") & "") & """, "
    jsonObjeto = jsonObjeto & """dataConclusao"": """ & EscapeJSON(rsObter("DataConclusao") & "") & """, "
    jsonObjeto = jsonObjeto & """quantidadePaletes"": """ & EscapeJSON(rsObter("QuantidadePaletes") & "") & """, "
    jsonObjeto = jsonObjeto & """nomeCliente"": """ & EscapeJSON(rsObter("NomeCliente") & "") & """, "
    jsonObjeto = jsonObjeto & """modoDefeito"": """ & EscapeJSON(rsObter("ModoDefeito") & "") & """, "
    jsonObjeto = jsonObjeto & """area"": """ & EscapeJSON(rsObter("Area") & "") & """, "
    jsonObjeto = jsonObjeto & """ordemProducao"": """ & EscapeJSON(rsObter("OrdemProducao") & "") & """ "
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Ficha encontrada"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()    
    idFNC = Request("idFNC")
    codPlanta = Request("codPlanta")
    quantidadePaletes =  Iif(Request("quantidadePaletes") = "", 0, Request("quantidadePaletes"))
    quantidadeRecebida =  Iif(Request("quantidadeRecebida") = "", 0, Request("quantidadeRecebida"))
    quantidadeRefugada = Iif(Request("quantidadeRefugada") = "", 0, Request("quantidadeRefugada"))
    valor = Iif(Request("valor") = "", 0, Request("valor"))
    valorUnitarioMaterial = Iif(Request("valorUnitarioMaterial") = "", 0, Request("valorUnitarioMaterial"))
    NFDevolucao = Request("NFDevolucao")
    codigoVA = Request("codigoVA")
    codigoQM = Request("codigoQM")
    dataEntradaDevolucao = Request("dataEntradaDevolucao")
    dataConclusao = Iif(Request("dataConclusao") <> "", "'" & Request("dataConclusao") & "'", "NULL")
    status = Request("status")
    nomeCliente = Request("nomeCliente")
    modoDefeito = Request("modoDefeito")
    area = Request("area")
    ordemProducao = Request("ordemProducao")

    SQL = "INSERT INTO FichaQM ("
    SQL = SQL & "   IdFNC, CodPlanta, Status, QuantidadeRecebida, QuantidadeRefugada, "
    SQL = SQL & "   NFDevolucao, Valor, ValorUnitarioMaterial, CodigoVA, CodigoQM, DataEntradaDevolucao, "
    SQL = SQL & "   QuantidadePaletes, NomeCliente, ModoDefeito, Area, OrdemProducao, DataConclusao "
    SQL = SQL & ") "
    SQL = SQL & "OUTPUT INSERTED.Id "
    SQL = SQL & "VALUES ( "
    SQL = SQL & "'" & idFNC & "', "
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "'" & status & "', "
    SQL = SQL & "'" & quantidadeRecebida & "', "
    SQL = SQL & "'" & quantidadeRefugada & "', "
    SQL = SQL & "'" & NFDevolucao & "', "
    SQL = SQL & "'" & Replace(valor, ",", ".") & "', "
    SQL = SQL & "'" & Replace(valorUnitarioMaterial, ",", ".") & "', "
    SQL = SQL & "'" & codigoVA & "', "
    SQL = SQL & "'" & codigoQM & "', "
    SQL = SQL & "'" & dataEntradaDevolucao & "',"
    SQL = SQL & "'" & quantidadePaletes & "',"
    SQL = SQL & "'" & nomeCliente & "',"
    SQL = SQL & "'" & modoDefeito & "',"
    SQL = SQL & "'" & area & "',"
    SQL = SQL & "'" & ordemProducao & "',"
    SQL = SQL & "" & dataConclusao & ""
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar ficha: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    id = rsCriar("Id")
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Ficha criada com sucesso!""}"
End Function

Function Editar()
    id = Request("id")

    If id = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""ID da ficha não informado""}"
        Exit Function
    End If

    codPlanta = Request("codPlanta")
    idFNC = Request("idFNC")
    quantidadeRecebida =  Iif(Request("quantidadeRecebida") = "", 0, Request("quantidadeRecebida"))
    quantidadeRefugada = Iif(Request("quantidadeRefugada") = "", 0, Request("quantidadeRefugada"))
    valor = Iif(Request("valor") = "", 0, Request("valor"))
    valorUnitarioMaterial = Iif(Request("valorUnitarioMaterial") = "", 0, Request("valorUnitarioMaterial"))
    NFDevolucao = Request("NFDevolucao")
    codigoVA = Request("codigoVA")
    codigoQM = Request("codigoQM")
    status = Request("status")
    dataEntradaDevolucao = Request("dataEntradaDevolucao")
    dataConclusao = Iif(Request("dataConclusao") <> "", "'" & Request("dataConclusao") & "'", "NULL")
    quantidadePaletes =  Iif(Request("quantidadePaletes") = "", 0, Request("quantidadePaletes"))
    nomeCliente = Request("nomeCliente")
    modoDefeito = Request("modoDefeito")
    area = Request("area")
    ordemProducao = Request("ordemProducao")

    SQL = "UPDATE FichaQM SET "
    SQL = SQL & "   CodPlanta = '" & codPlanta & "', "
    SQL = SQL & "   Status = '" & status & "', "
    SQL = SQL & "   IdFNC = '" & idFNC & "', "
    SQL = SQL & "   QuantidadeRecebida = '" & quantidadeRecebida & "', "
    SQL = SQL & "   QuantidadeRefugada = '" & quantidadeRefugada & "', "
    SQL = SQL & "   NFDevolucao = '" & NFDevolucao & "', "
    SQL = SQL & "   Valor = '" & Replace(valor, ",", ".") & "', "
    SQL = SQL & "   ValorUnitarioMaterial = '" & Replace(valorUnitarioMaterial, ",", ".") & "', "
    SQL = SQL & "   CodigoVA = '" & codigoVA & "', "
    SQL = SQL & "   CodigoQM = '" & codigoQM & "', "
    SQL = SQL & "   DataEntradaDevolucao = '" & dataEntradaDevolucao & "', "
    SQL = SQL & "   QuantidadePaletes = '" & quantidadePaletes & "', "
    SQL = SQL & "   NomeCliente = '" & nomeCliente & "', "
    SQL = SQL & "   ModoDefeito = '" & modoDefeito & "', "
    SQL = SQL & "   Area = '" & area & "', "
    SQL = SQL & "   OrdemProducao = '" & ordemProducao & "', "
    SQL = SQL & "   DataConclusao = " & dataConclusao & " "
    SQL = SQL & "WHERE Id = " & id 

    On Error Resume Next
    SET rsEditar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar ficha: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEditar.Close()
    SET rsEditar = NOTHING
    On Error Goto 0

    Call RegistrarLog("edicao", id)
    Editar = "{""sucesso"": true, ""mensagem"": ""Ficha editada com sucesso!""}"
End Function

Function Excluir()
    id = Request("id")
    
    If id = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    SQL = "UPDATE FichaQM SET FlagInativo = 1 WHERE Id = " & id
    
    On Error Resume Next
    SET rsExcluir = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir ficha: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsExcluir.Close()
    SET rsExcluir = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("exclusao", id)
    Excluir = "{""sucesso"": true, ""mensagem"": ""Ficha excluída com sucesso!""}"
End Function

Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AGUARDANDO_RETRABALHO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aguardando Retrabalho" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "AGUARDANDO_FATURAMENTO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aguardando Faturamento" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "REFATURADO_TOTAL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Refaturado Total" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","


    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "REFUGADO_TOTAL" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Refugado Total" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "RETRABALHO_FINALIZADO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Retrabalho Finalizado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "SIMBOLICO_REFATURADO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Simbólico Refaturado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "SIMBOLICO_REFUGADO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Simbólico Refugado" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterDadosUsuario()
    nivelUsuario = ""
    setorUsuario = ""
    plantas = ""

    ' Query para verificar se é Analista ou Responsável
    SQL = "SELECT TOP 1" 
    SQL = SQL & "    CASE WHEN r.Tipo = 1 THEN 'ANALISTA' ELSE 'RESPONSAVEL' END Nivel, "
    SQL = SQL & "    r.IdSetor, "
    SQL = SQL & "    LTRIM(RTRIM(s.Descricao)) as Setor, "
    SQL = SQL & "    STUFF(( "
    SQL = SQL & "        SELECT ', ' + r2.Planta "
    SQL = SQL & "        FROM Responsavel r2 "
    SQL = SQL & "        WHERE r2.Id = r.Id "
    SQL = SQL & "        AND r2.Tipo = r.Tipo "
    SQL = SQL & "        AND r2.Email = r.Email "
    SQL = SQL & "        AND (r2.Excluido = 0 OR r2.Excluido IS NULL) "
    SQL = SQL & "        FOR XML PATH('')), 1, 2, '' "
    SQL = SQL & "    ) AS Plantas "
    SQL = SQL & "FROM Responsavel r "
    SQL = SQL & "LEFT JOIN Setor s ON r.IdSetor = s.Id "
    SQL = SQL & "WHERE (r.Excluido = 0 OR r.Excluido IS NULL) "
    SQL = SQL & "AND r.Id = " & usuarioLogado("id")

    SET rsResponsavel = Conecta("BEN-CORPORATIVO", SQL)
    
    ' Se der erro finaliza fluxo
    If Err.Number <> 0 OR Not IsObject(rsResponsavel) Then
        ObterDadosUsuario = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsResponsavel.EOF Then
        nivelUsuario = rsResponsavel("Nivel")
        setorUsuario = rsResponsavel("Setor")
        plantas = rsResponsavel("Plantas")
    End If

    rsResponsavel.Close()
    SET rsResponsavel = NOTHING

    ' Se não tiver nenhum nível consideramos que é somente visualização
    If nivelUsuario = "" Then
        nivelUsuario = "VISUALIZACAO"
        plantas = usuarioLogado("planta")
        setorUsuario = ""
    End If
    
    dadosJSON = "{""setor"": """ & setorUsuario & """, ""nivel"": """ & nivelUsuario & """, ""plantas"": """ & plantas & """}"
    ObterDadosUsuario = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Sub RegistrarLog(acao, id)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Ficha de QM"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a ficha DE id " & id
        Case "edicao"
            acaoLog = "Edição de Ficha de QM"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou a ficha de id " & id
        Case "exclusao"
            acaoLog = "Exclusão de Ficha de QM"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu a ficha de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Ficha de QM', "
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