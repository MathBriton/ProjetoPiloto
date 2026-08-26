<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp" -->
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
usuarioLogado.Add "codigo", Session("CODUSER")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "email", Session("EMAIL")
usuarioLogado.Add "nivel", SESSION("NIVEL")

Dim acao
acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "obterDadosDashboard"
        resultado = ObterDadosDashboard()
    Case "obterDadosGrafico"
        resultado = ObterDadosGrafico()
    Case "listarResponsaveis"
        resultado = ListarResponsaveis()
    Case "listarSetores"
        resultado = ListarSetores()
    Case "listarDefeitos"
        resultado = ListarDefeitos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function ObterDadosDashboard()
    Dim totaisGerais, fncPorCliente, fncPorDefeito, fncPorSetor, fncPorFilial

    Set totaisGerais = ObterTotaisGerais()
    totalFNC = totaisGerais("totalFNC")
    totalAbertas = totaisGerais("totalAbertas")
    totalRespondidas = totaisGerais("totalRespondidas")
    totalEncerradas = totaisGerais("totalEncerradas")

    jsonfncPorFilial = ObterFncPorFilial(false)
    jsonfncPorDefeito = ObterFncPorDefeito(false)
    jsonfncPorSetor = ObterFncPorSetor(false)
    jsonfncPorCliente = ObterFncPorCliente(false)

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """totalFNC"": " & totalFNC & ", "
    jsonObjeto = jsonObjeto & """totalAbertas"": " & totalAbertas & ", "
    jsonObjeto = jsonObjeto & """totalRespondidas"": " & totalRespondidas & ", "
    jsonObjeto = jsonObjeto & """totalEncerradas"": " & totalEncerradas & ", "
    jsonObjeto = jsonObjeto & """fncPorFilial"": " & jsonfncPorFilial & ", "
    jsonObjeto = jsonObjeto & """fncPorDefeito"": " & jsonfncPorDefeito & ", "
    jsonObjeto = jsonObjeto & """fncPorCliente"": " & jsonfncPorCliente & ", "
    jsonObjeto = jsonObjeto & """fncPorSetor"": " & jsonfncPorSetor
    jsonObjeto = jsonObjeto & "}"
    
    ObterDadosDashboard = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function ObterDadosGrafico()
    Dim fncPorCliente, fncPorDefeito, fncPorSetor, fncPorFilial

    tipoGrafico = Request("tipoGrafico")

    Dim jsonDados
    Select Case tipoGrafico
        Case "fncPorFilial"
            jsonDados = ObterFncPorFilial(true)
        Case "fncPorDefeito"
            jsonDados = ObterFncPorDefeito(true)
        Case "fncPorSetor"
            jsonDados = ObterFncPorSetor(true)
        Case "fncPorCliente"
            jsonDados = ObterFncPorCliente(true)
        Case Else
            jsonDados = "[]"
    End Select
    
    ObterDadosGrafico = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonDados & "}"
End Function

Function ObterTotaisGerais()
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")

    Dim dados
    Set dados = Server.CreateObject("Scripting.Dictionary")
    
    dados("totalFNC") = 0
    dados("totalAbertas") = 0
    dados("totalRespondidas") = 0
    dados("totalEncerradas") = 0
    
    SQL = "SELECT COUNT(f.Id) as totalFNC, "
    SQL = SQL & "SUM(CASE WHEN f.Situacao = 1 THEN 1 ELSE 0 END) as totalAbertas, "
    SQL = SQL & "SUM(CASE WHEN f.Situacao = 2 THEN 1 ELSE 0 END) as totalRespondidas, "
    SQL = SQL & "SUM(CASE WHEN f.Situacao = 3 THEN 1 ELSE 0 END) as totalEncerradas "
    SQL = SQL & "FROM FNC f WITH(NOLOCK) "
    SQL = SQL & "WHERE ISNULL(f.Excluido, 0) = 0 " 

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND f.IdSetor IN (" & idSetor & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND f.Id = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao IN (" & situacao & ")"
    End If

    If validacao <> "" Then
        SQL = SQL & " AND f.Status IN (" & validacao & ")"
    End If

    If disposicao <> "" Then
        SQL = SQL & " AND f.IdDisposicao IN (" & disposicao & ")"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND f.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND f.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If dataEnvioClienteDe <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente >= '" & Replace(dataEnvioClienteDe, "-", "") & "'"
    End If

    If dataEnvioClienteAte <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente < DATEADD(DAY, 1, '" & Replace(dataEnvioClienteAte & "", "-", "") & "')"
    End If


    On Error Resume Next
    SET rsTotais = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTotais) Then
        ObterTotaisGerais = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar totais gerais: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsTotais.EOF Then
        dados("totalFNC") = Iif(isNull(rsTotais("totalFNC")), 0, rsTotais("totalFNC"))
        dados("totalAbertas") = Iif(isNull(rsTotais("totalAbertas")), 0, rsTotais("totalAbertas"))
        dados("totalRespondidas") = Iif(isNull(rsTotais("totalRespondidas")), 0, rsTotais("totalRespondidas"))
        dados("totalEncerradas") = Iif(isNull(rsTotais("totalEncerradas")), 0, rsTotais("totalEncerradas"))
    End If
    
    Set ObterTotaisGerais = dados
End Function

Function ObterFncPorFilial(carregarTodos)
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")

    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT f.Planta as label, COUNT(f.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM FNC f WITH(NOLOCK) "
    SQL = SQL & "   WHERE ISNULL(f.Excluido, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND f.IdSetor IN (" & idSetor & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND f.Id = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao IN (" & situacao & ")"
    End If

    If validacao <> "" Then
        SQL = SQL & " AND f.Status IN (" & validacao & ")"
    End If

    If disposicao <> "" Then
        SQL = SQL & " AND f.IdDisposicao IN (" & disposicao & ")"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND f.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND f.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If dataEnvioClienteDe <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente >= '" & Replace(dataEnvioClienteDe, "-", "") & "'"
    End If

    If dataEnvioClienteAte <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente < DATEADD(DAY, 1, '" & Replace(dataEnvioClienteAte & "", "-", "") & "')"
    End If
    
    SQL = SQL & "   GROUP BY f.Planta "
    SQL = SQL & ") "

    maximoItems = Iif(carregarTodos, 9999, 20)

    SQL = SQL & "SELECT "
    SQL = SQL & " CASE "
    SQL = SQL & "   WHEN ranking <=" & maximoItems & "THEN CAST(label AS VARCHAR(50))" 
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END as label,"
    SQL = SQL & " SUM(total) as total"
    SQL = SQL & " FROM Item"
    SQL = SQL & " GROUP BY "
    SQL = SQL & "   CASE WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END "
    SQL = SQL & " ORDER BY"
    SQL = SQL & "   CASE WHEN CASE  WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros' END = 'Outros' THEN 1 ELSE 0"
    SQL = SQL & " END, "
    SQL = SQL & " total DESC"

    On Error Resume Next
    SET rsPlantas = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsPlantas) Then
        ObterFncPorFilial = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsPlantas.EOF Then
        Do While Not rsPlantas.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(ObterNomeFilial(rsPlantas("label"))) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsPlantas("label")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsPlantas("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsPlantas.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsPlantas.Close()
    SET rsPlantas = NOTHING
    
    ObterFncPorFilial = jsonArray
End Function

Function ObterFncPorDefeito(carregarTodos)
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")
    
    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT def.Descricao as label, def.Id as valor, COUNT(f.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM FNC f WITH(NOLOCK) "
    SQL = SQL & "   INNER JOIN Defeito def WITH(NOLOCK) ON f.IdDefeito = def.Id "
    SQL = SQL & "   WHERE ISNULL(f.Excluido, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND f.IdSetor IN (" & idSetor & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND f.Id = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao IN (" & situacao & ")"
    End If

    If validacao <> "" Then
        SQL = SQL & " AND f.Status IN (" & validacao & ")"
    End If

    If disposicao <> "" Then
        SQL = SQL & " AND f.IdDisposicao IN (" & disposicao & ")"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND f.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND f.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If dataEnvioClienteDe <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente >= '" & Replace(dataEnvioClienteDe, "-", "") & "'"
    End If

    If dataEnvioClienteAte <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente < DATEADD(DAY, 1, '" & Replace(dataEnvioClienteAte & "", "-", "") & "')"
    End If

    SQL = SQL & "   GROUP BY def.Descricao, def.Id"
    SQL = SQL & ") "

    maximoItems = Iif(carregarTodos, 9999, 10)

    SQL = SQL & "SELECT "
    SQL = SQL & " CASE "
    SQL = SQL & "   WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))" 
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END as label,"
    SQL = SQL & " CASE "
    SQL = SQL & "   WHEN ranking <=" & maximoItems & " THEN valor" 
    SQL = SQL & "   ELSE NULL"
    SQL = SQL & " END as valor,"
    SQL = SQL & " SUM(total) as total"
    SQL = SQL & " FROM Item"
    SQL = SQL & " GROUP BY "
    SQL = SQL & "   CASE WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END, "
    SQL = SQL & "   CASE WHEN ranking <=" & maximoItems & " THEN valor"
    SQL = SQL & " END "
    SQL = SQL & " ORDER BY"
    SQL = SQL & "   CASE WHEN CASE  WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros' END = 'Outros' THEN 1 ELSE 0"
    SQL = SQL & " END, "
    SQL = SQL & " total DESC"
    

    On Error Resume Next
    SET rsDefeitos = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDefeitos) Then
        ObterFncPorDefeito = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDefeitos.EOF Then
        Do While Not rsDefeitos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsDefeitos("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsDefeitos("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsDefeitos("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDefeitos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDefeitos.Close()
    SET rsDefeitos = NOTHING
    
    ObterFncPorDefeito = jsonArray
End Function

Function ObterFncPorSetor(carregarTodos)
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")

    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT s.Descricao as label, f.IdSetor as valor, COUNT(f.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM FNC f WITH(NOLOCK) "
    SQL = SQL & "   LEFT JOIN Setor s ON f.IdSetor = s.Id "
    SQL = SQL & "   WHERE ISNULL(f.Excluido, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND f.IdSetor IN (" & idSetor & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND f.Id = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao IN (" & situacao & ")"
    End If

    If validacao <> "" Then
        SQL = SQL & " AND f.Status IN (" & validacao & ")"
    End If

    If disposicao <> "" Then
        SQL = SQL & " AND f.IdDisposicao IN (" & disposicao & ")"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND f.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND f.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If dataEnvioClienteDe <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente >= '" & Replace(dataEnvioClienteDe, "-", "") & "'"
    End If

    If dataEnvioClienteAte <> "" Then
        SQL = SQL & " AND f.DataEnvioCliente < DATEADD(DAY, 1, '" & Replace(dataEnvioClienteAte & "", "-", "") & "')"
    End If

    SQL = SQL & "   GROUP BY f.IdSetor, s.Descricao"
    SQL = SQL & ") "

    maximoItems = Iif(carregarTodos, 9999, 10)

    SQL = SQL & "SELECT "
    SQL = SQL & " CASE "
    SQL = SQL & "   WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))" 
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END as label,"
    SQL = SQL & " CASE "
    SQL = SQL & "   WHEN ranking <=" & maximoItems & " THEN valor" 
    SQL = SQL & "   ELSE NULL"
    SQL = SQL & " END as valor,"
    SQL = SQL & " SUM(total) as total"
    SQL = SQL & " FROM Item"
    SQL = SQL & " GROUP BY "
    SQL = SQL & "   CASE WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros'"
    SQL = SQL & " END, "
    SQL = SQL & "   CASE WHEN ranking <=" & maximoItems & " THEN valor"
    SQL = SQL & " END "
    SQL = SQL & " ORDER BY"
    SQL = SQL & "   CASE WHEN CASE  WHEN ranking <=" & maximoItems & " THEN CAST(label AS VARCHAR(50))"
    SQL = SQL & "   ELSE 'Outros' END = 'Outros' THEN 1 ELSE 0"
    SQL = SQL & " END, "
    SQL = SQL & " total DESC"
    
    On Error Resume Next
    SET rsSetores = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSetores) Then
        ObterFncPorSetor = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSetores.EOF Then
        Do While Not rsSetores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsSetores("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsSetores("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsSetores("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSetores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSetores.Close()
    SET rsSetores = NOTHING
    
    ObterFncPorSetor = jsonArray
End Function

Function ObterFncPorCliente(carregarTodos)
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")

    SQL = "WITH Fnc AS "
    SQL = SQL & "("
    SQL = SQL & "   SELECT"
    SQL = SQL & "    Planta,"
    SQL = SQL & "    Cliente,"
    SQL = SQL & "    Loja,"
    SQL = SQL & "    COUNT(*) AS Quantidade,"
    SQL = SQL & "    ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) AS MAIOR"
    SQL = SQL & "    FROM FNC"
    SQL = SQL & "    WHERE ISNULL(Excluido, 0) = 0"

    If codPlanta <> "" Then
        SQL = SQL & " AND Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND Loja = '" & codLoja & "'"
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND IdSetor IN (" & idSetor & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND Id = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND Situacao IN (" & situacao & ")"
    End If

    If validacao <> "" Then
        SQL = SQL & " AND Status IN (" & validacao & ")"
    End If

    If disposicao <> "" Then
        SQL = SQL & " AND IdDisposicao IN (" & disposicao & ")"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If dataEnvioClienteDe <> "" Then
        SQL = SQL & " AND DataEnvioCliente >= '" & Replace(dataEnvioClienteDe, "-", "") & "'"
    End If

    If dataEnvioClienteAte <> "" Then
        SQL = SQL & " AND DataEnvioCliente < DATEADD(DAY, 1, '" & Replace(dataEnvioClienteAte & "", "-", "") & "')"
    End If

    SQL = SQL & "    GROUP BY Cliente, Planta, Loja"
    SQL = SQL & " ),"
    SQL = SQL & " Cliente AS "
    SQL = SQL & " ("
    SQL = SQL & "    SELECT CIDADE, COD_CLIENTE, LOJA_CLIENTE, FANTASIA_CLIENTE, CODIGO_ERP "
    SQL = SQL & "    FROM BI_CLIENTES"
    SQL = SQL & " )"
    SQL = SQL & " SELECT FANTASIA_CLIENTE as label, Quantidade as total, MAIOR,"
    SQL = SQL & " CASE WHEN f.Planta IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END as valor"
    SQL = SQL & " FROM Fnc f"
    SQL = SQL & " LEFT JOIN Cliente c ON c.CIDADE = CASE f.Planta "
    SQL = SQL & "   WHEN '00' THEN 'PIRAPETINGA' "
    SQL = SQL & "   WHEN '04' THEN 'UBERABA' "
    SQL = SQL & "   WHEN '05' THEN 'BENTO' "
    SQL = SQL & "   WHEN '06' THEN 'MARANGUAPE' "
    SQL = SQL & "   WHEN '07' THEN 'SAQUAREMA' "
    SQL = SQL & "   WHEN '08' THEN 'ARACATUBA' "
    SQL = SQL & "   WHEN '09' THEN 'BLUMENAU' "
    SQL = SQL & "   WHEN '10' THEN 'PACAJUS' "
    SQL = SQL & "   WHEN '11' THEN 'PORTO FELIZ' "
    SQL = SQL & "   WHEN '12' THEN 'TRES BARRAS' "
    SQL = SQL & "   WHEN '13' THEN 'FLORESTAL' "
    SQL = SQL & " END "
    SQL = SQL & " AND CASE WHEN f.Planta IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END = f.Cliente COLLATE SQL_Latin1_General_CP1_CI_AS "
    SQL = SQL & " AND CASE WHEN f.Planta IN ('09') THEN '01' ELSE LOJA_CLIENTE END = f.Loja COLLATE SQL_Latin1_General_CP1_CI_AS "
    
    If NOT carregarTodos Then
        SQL = SQL & " WHERE MAIOR <= 10"
    End If
    
    SQL = SQL & " ORDER BY MAIOR"
     
    On Error Resume Next
    SET rsSetores = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSetores) Then
        ObterFncPorCliente = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSetores.EOF Then
        Do While Not rsSetores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsSetores("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsSetores("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsSetores("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSetores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSetores.Close()
    SET rsSetores = NOTHING
    
    ObterFncPorCliente = jsonArray
End Function

Function ListarDefeitos()
    Dim SQL, rsDefeitos, jsonArray, primeiroItem

    tipo = request("filtroTipoDefeito")

    SQL = "SELECT Id, Descricao FROM Defeito"

    If tipo <> "" Then
        SQL = SQL & " WHERE Tipo = '" & tipo & "'"
    End If

    On Error Resume Next
    SET rsDefeitos = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDefeitos) Then
        ListarDefeitos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de defeitos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDefeitos.EOF Then
        Do While Not rsDefeitos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsDefeitos("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsDefeitos("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDefeitos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDefeitos.Close()
    SET rsDefeitos = NOTHING
    
    ListarDefeitos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarSetores()
    Dim SQL, rsSetores, jsonArray, primeiroItem

    SQL = "SELECT Id, Descricao FROM Setor"

    On Error Resume Next
    SET rsSetores = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSetores) Then
        ListarSetores = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de setores: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSetores.EOF Then
        Do While Not rsSetores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsSetores("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsSetores("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSetores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSetores.Close()
    SET rsSetores = NOTHING
    
    ListarSetores = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterNomeFilial(codPlanta)
    Dim nomeFilial 
    nomeFilial = ""
    Select Case codPlanta
        Case "00"
            nomeFilial = "Pirapetinga"
        Case "01"
            nomeFilial = "Corporativo RJ"
        Case "02"
            nomeFilial = "Corporativo SP"
        Case "03"
            nomeFilial = "Corporativo"
        Case "04"
            nomeFilial = "Uberaba"
        Case "05"
            nomeFilial = "Bento"
        Case "06"
            nomeFilial = "Maranguape"
        Case "07"
            nomeFilial = "Saquarema"
        Case "08"
            nomeFilial = "Araçatuba"
        Case "09"
            nomeFilial = "Blumenau"
        Case "10"
            nomeFilial = "Pacajus"
        Case "11"
            nomeFilial = "Porto Feliz"
        Case "12"
            nomeFilial = "Três Barras"
        Case "13"
            nomeFilial = "Florestal"
        Case Else
            nomeFilial = "Outros"
    End Select
    
    ObterNomeFilial = nomeFilial
End Function

%>