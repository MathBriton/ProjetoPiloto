<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../includes/variaveis.asp" -->
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "codigo", Session("CODUSER")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "email", Session("EMAIL")
usuarioLogado.Add "nivel", SESSION("NIVEL")

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Select Case acao
    Case "obterDadosDashboard"
        resultado = ObterDadosDashboard()
    Case "obterDadosGrafico"
        resultado = ObterDadosGrafico()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function ObterDadosDashboard()
    Set totaisGerais = ObterTotaisGerais()
    totalAbertos = totaisGerais("totalAbertos")
    totalConcluidos = totaisGerais("totalConcluidos")
    totalAtrasados = totaisGerais("totalAtrasados")
    tempoMedioMinutos = totaisGerais("tempoMedioMinutos")

    jsonChamadoPorTipo = ObterChamadoPorTipo(false)
    jsonChamadoPorCategoria = ObterChamadoPorCategoria(false)
    jsonChamadoPorDepartamento = ObterChamadoPorDepartamento(false)
    jsonChamadoPorAtendente = ObterChamadoPorAtendente(false)
    jsonVolumeMensal = ObterVolumeMensal()
    jsonSLAPorUrgencia = ObterSLAPorUrgencia()

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """totalAbertos"": " & totalAbertos & ", "
    jsonObjeto = jsonObjeto & """totalConcluidos"": " & totalConcluidos & ", "
    jsonObjeto = jsonObjeto & """totalAtrasados"": " & totalAtrasados & ", "
    jsonObjeto = jsonObjeto & """tempoMedioMinutos"": " & tempoMedioMinutos & ", "
    jsonObjeto = jsonObjeto & """chamadoPorCategoria"": " & jsonChamadoPorCategoria & ", "
    jsonObjeto = jsonObjeto & """chamadoPorTipo"": " & jsonChamadoPorTipo & ", "
    jsonObjeto = jsonObjeto & """volumeMensal"": " & jsonVolumeMensal & ", "
    jsonObjeto = jsonObjeto & """chamadoPorDepartamento"": " & jsonChamadoPorDepartamento & ", "
    jsonObjeto = jsonObjeto & """chamadoPorAtendente"": " & jsonChamadoPorAtendente & ", "
    jsonObjeto = jsonObjeto & """SLAPorUrgencia"": " & jsonSLAPorUrgencia
    jsonObjeto = jsonObjeto & "}"
    
    ObterDadosDashboard = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function ObterDadosGrafico()
    tipoGrafico = Request("tipoGrafico")

    Select Case tipoGrafico
        Case "chamadoPorTipo"
            jsonDados = ObterChamadoPorTipo(true)
         Case "chamadoPorCategoria"
            jsonDados = ObterChamadoPorCategoria(true)
        Case "chamadoPorDepartamento"
            jsonDados = ObterChamadoPorDepartamento(true)
        Case Else
            jsonDados = "[]"
    End Select
    
    ObterDadosGrafico = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonDados & "}"
End Function

Function ObterTotaisGerais()
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")

    idAtendente = Request("filtroAtendente")
    codPlanta = Request("filtroFilial")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("totalAbertos") = 0
    dados("totalConcluidos") = 0
    dados("totalAtrasados") = 0
    dados("tempoMedioMinutos") = 0
    
    SQL = ""
    SQL = SQL & "DECLARE @DataInicio  DATETIME = '" & Replace(dataCriacaoDe, "-", "") & "'; "  
    SQL = SQL & "DECLARE @DataFim     DATETIME = " & "DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "'); " 

    SQL = SQL & "SELECT "
    SQL = SQL & "    SUM(CASE WHEN c.DataCriacao >= @DataInicio AND c.DataCriacao <  @DataFim THEN 1 ELSE 0 END) as totalAbertos,"
    SQL = SQL & "    SUM(CASE WHEN c.DataFinalizacao >= @DataInicio AND c.DataFinalizacao <  @DataFim THEN 1 ELSE 0 END) as totalConcluidos, "
    SQL = SQL & "    SUM("
    SQL = SQL & "        CASE WHEN "
    SQL = SQL & "            c.DataCriacao >= @DataInicio "
    SQL = SQL & "            AND c.DataCriacao < @DataFim "
    SQL = SQL & "            AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
    SQL = SQL & "        THEN 1 ELSE 0 END"
    SQL = SQL & "    ) as totalAtrasados, "
    SQL = SQL & "   CASE WHEN SUM("
	SQL = SQL & "   	CASE WHEN Status = 'FINALIZADO' AND DataFinalizacao >= @DataInicio AND DataFinalizacao <  @DataFim"
	SQL = SQL & "   	THEN 1"
	SQL = SQL & "   	ELSE 0 END"
	SQL = SQL & "   ) > 0 THEN"
	SQL = SQL & "   SUM("
	SQL = SQL & "   	CASE WHEN Status = 'FINALIZADO' AND DataFinalizacao >= @DataInicio AND DataFinalizacao <  @DataFim"
	SQL = SQL & "   	THEN CAST(DATEDIFF(MINUTE, c.DataCriacao, c.DataFinalizacao) AS BIGINT)"
	SQL = SQL & "   	ELSE 0 END"
	SQL = SQL & "   ) /"
	SQL = SQL & "   SUM("
	SQL = SQL & "   	CASE WHEN Status = 'FINALIZADO' AND DataFinalizacao >= @DataInicio AND DataFinalizacao <  @DataFim"
	SQL = SQL & "   	THEN 1"
	SQL = SQL & "   	ELSE 0 END"
	SQL = SQL & "   ) ELSE 0 END AS tempoMedioMinutos "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "WHERE ISNULL(c.FlagInativo, 0) = 0 "
    
    
    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If codPlanta <> "" Then
        SQL = SQL & "AND c.CodPlanta = '" & codPlanta & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    On Error Resume Next
    SET rsTotais = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTotais) Then
        ObterTotaisGerais = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar totais gerais: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsTotais.EOF Then
        dados("totalAbertos") = Iif(isNull(rsTotais("totalAbertos")), 0, rsTotais("totalAbertos"))
        dados("totalConcluidos") = Iif(isNull(rsTotais("totalConcluidos")), 0, rsTotais("totalConcluidos"))
        dados("totalAtrasados") = Iif(isNull(rsTotais("totalAtrasados")), 0, rsTotais("totalAtrasados"))
        dados("tempoMedioMinutos") = Iif(isNull(rsTotais("tempoMedioMinutos")), 0, rsTotais("tempoMedioMinutos"))
    End If
    
    Set ObterTotaisGerais = dados
End Function

Function ObterChamadoPorCategoria(carregarTodos)
    codPlanta = Request("filtroCodPlanta")
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    idAtendente = Request("filtroAtendente")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT cat.Nome as label, cat.Id as valor, COUNT(c.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM Chamados c WITH(NOLOCK) "
    SQL = SQL & "   INNER JOIN CategoriasChamado cat WITH(NOLOCK) ON c.IdCategoria = cat.Id "
    SQL = SQL & "   LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "   LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "   WHERE ISNULL(c.FlagInativo, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND c.Planta = '" & codPlanta & "'"
    End If
   
    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If dataCriacaoDe <> "" Then
        SQL = SQL & " AND c.DataCriacao >= '" & Replace(dataCriacaoDe, "-", "") & "'"
    End If

    If dataCriacaoAte <> "" Then
        SQL = SQL & " AND c.DataCriacao < DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "')"
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "   GROUP BY cat.Nome, cat.Id"
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
    SET rsCategorias = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCategorias) Then
        ObterChamadoPorCategoria = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCategorias.EOF Then
        Do While Not rsCategorias.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsCategorias("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsCategorias("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsCategorias("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCategorias.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCategorias.Close()
    SET rsCategorias = NOTHING
    
    ObterChamadoPorCategoria = jsonArray
End Function

Function ObterChamadoPorTipo(carregarTodos)
    codPlanta = Request("filtroCodPlanta")
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    idAtendente = Request("filtroAtendente")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT t.Nome as label, t.Id as valor, COUNT(c.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM Chamados c WITH(NOLOCK) "
    SQL = SQL & "   INNER JOIN TiposChamado t WITH(NOLOCK) ON c.IdTipo = t.Id "
    SQL = SQL & "   LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "   LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "   WHERE ISNULL(c.FlagInativo, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND c.Planta = '" & codPlanta & "'"
    End If
   
    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If dataCriacaoDe <> "" Then
        SQL = SQL & " AND c.DataCriacao >= '" & Replace(dataCriacaoDe, "-", "") & "'"
    End If

    If dataCriacaoAte <> "" Then
        SQL = SQL & " AND c.DataCriacao < DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "')"
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "   GROUP BY t.Nome, t.Id"
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
    SET rsTipos = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTipos) Then
        ObterChamadoPorTipo = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsTipos.EOF Then
        Do While Not rsTipos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsTipos("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsTipos("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsTipos("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsTipos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsTipos.Close()
    SET rsTipos = NOTHING
    
    ObterChamadoPorTipo = jsonArray
End Function

Function ObterChamadoPorDepartamento(carregarTodos)
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    idAtendente = Request("filtroAtendente")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "WITH Item AS ("
    SQL = SQL & "   SELECT d.Nome as label, c.IdDepartamento as valor, COUNT(c.Id) as total, "
    SQL = SQL & "   ROW_NUMBER() OVER (ORDER BY COUNT(*) DESC) as ranking" 
    SQL = SQL & "   FROM Chamados c WITH(NOLOCK) "
    SQL = SQL & "   LEFT JOIN DepartamentosChamado d ON c.IdDepartamento = d.Id "
    SQL = SQL & "   LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "   LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade"
    SQL = SQL & "   WHERE ISNULL(c.FlagInativo, 0) = 0 "

    If codPlanta <> "" Then
        SQL = SQL & " AND c.Planta = '" & codPlanta & "'"
    End If

    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If dataCriacaoDe <> "" Then
        SQL = SQL & " AND c.DataCriacao >= '" & Replace(dataCriacaoDe, "-", "") & "'"
    End If

    If dataCriacaoAte <> "" Then
        SQL = SQL & " AND c.DataCriacao < DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "')"
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "   GROUP BY c.IdDepartamento, d.Nome"
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
    SET rsDepartamentos = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDepartamentos) Then
        ObterChamadoPorDepartamento = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDepartamentos.EOF Then
        Do While Not rsDepartamentos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsDepartamentos("label")) & """, "
            jsonArray = jsonArray & """valor"": """ & EscapeJSON(rsDepartamentos("valor")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(rsDepartamentos("total")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDepartamentos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDepartamentos.Close()
    SET rsDepartamentos = NOTHING
    
    ObterChamadoPorDepartamento = jsonArray
End Function

Function ObterChamadoPorAtendente(carregarTodos)
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    idAtendente = Request("filtroAtendente")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "" 
    SQL = SQL & "DECLARE @DataInicio  DATETIME = '" & Replace(dataCriacaoDe, "-", "") & "'; "  
    SQL = SQL & "DECLARE @DataFim     DATETIME = " & "DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "'); " 

    SQL = SQL & "SELECT "
    SQL = SQL & "    a.Nome as Nome,"
    SQL = SQL & "    c.IdAtendente as IdAtendente,"
    SQL = SQL & "    COUNT(c.Id)as QuantidadeItens,"
    SQL = SQL & "    SUM(CASE WHEN c.DataFinalizacao >= @DataInicio AND c.DataFinalizacao < @DataFim THEN 1 ELSE 0 END ) as QuantidadeResolvida, "
    SQL = SQL & "    SUM(CASE WHEN c.DataFinalizacao >= @DataInicio AND c.DataFinalizacao < @DataFim AND c.DataFinalizacao < c.DataPrevista THEN 1 ELSE 0 END) as QuantidadeResolvidaNoPrazo "
    SQL = SQL & "FROM Chamados c WITH(NOLOCK)"
    SQL = SQL & "INNER JOIN AtendentesChamado a ON c.IdAtendente=a.Id "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "WHERE ISNULL(c.FlagInativo,0)=0 "
    SQL = SQL & "AND c.DataCriacao >= @DataInicio "
    SQL = SQL & "AND c.DataCriacao < @DataFim "

    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "GROUP BY c.IdAtendente,a.Nome "
    SQL = SQL & "ORDER BY SUM(CASE WHEN c.DataFinalizacao >= @DataInicio AND c.DataFinalizacao < @DataFim THEN 1 ELSE 0 END) "
    
    On Error Resume Next
    SET rsAtendentes = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAtendentes) Then
        ObterChamadoPorAtendente = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsAtendentes.EOF Then
        Do While Not rsAtendentes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsAtendentes("Nome")) & """, "
            jsonArray = jsonArray & """quantidadeResolvida"": """ & EscapeJSON(rsAtendentes("QuantidadeResolvida")) & """, "
            jsonArray = jsonArray & """quantidadeResolvidaNoPrazo"": """ & EscapeJSON(rsAtendentes("QuantidadeResolvidaNoPrazo")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsAtendentes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsAtendentes.Close()
    SET rsAtendentes = NOTHING
    
    ObterChamadoPorAtendente = jsonArray
End Function

Function ObterVolumeMensal()
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    idAtendente = Request("filtroAtendente")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "WITH Meses AS( "
    SQL = SQL & "    SELECT DATEADD(MONTH,-n,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1)) AS MesInicio "
    SQL = SQL & "    FROM(VALUES(0),(1),(2),(3),(4),(5),(6),(7),(8),(9),(10),(11))AS t(n) "
    SQL = SQL & "), "
    SQL = SQL & "CHAMADO AS ( "
    SQL = SQL & "    SELECT "
    SQL = SQL & "        m.MesInicio, "
    SQL = SQL & "        SUM(CASE WHEN c.DataCriacao >= m.MesInicio AND c.DataCriacao < DATEADD(MONTH,1,m.MesInicio) THEN 1 ELSE 0 END) as totalAbertos, "
    SQL = SQL & "        SUM(CASE WHEN c.DataFinalizacao >= m.MesInicio AND c.DataFinalizacao < DATEADD(MONTH,1,m.MesInicio) THEN 1 ELSE 0 END) as totalFinalizados "
    SQL = SQL & "    FROM Chamados c "
    SQL = SQL & "    CROSS JOIN Meses m "
    SQL = SQL & "    LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "    LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "    WHERE 1=1"

    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

    If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "GROUP BY m.MesInicio "
    SQL = SQL & ") "

    SQL = SQL & "SELECT "
    SQL = SQL & "    FORMAT(m.MesInicio,'yyyy-MM') AS Label, "
    SQL = SQL & "    ISNULL(c.totalAbertos,0) totalAbertos, "
    SQL = SQL & "    ISNULL(c.totalFinalizados,0) totalFinalizados "
    SQL = SQL & "FROM Meses m "
    SQL = SQL & "LEFT JOIN CHAMADO c ON c.MesInicio = m.MesInicio "
    SQL = SQL & "ORDER BY m.MesInicio "

    On Error Resume Next
    SET rsVolume = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsVolume) Then
        ObterVolumeMensal = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsVolume.EOF Then
        Do While Not rsVolume.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonValores = "["
            jsonValores = jsonValores & "{"
            jsonValores = jsonValores & """label"": """ & "Abertos" & """, "
            jsonValores = jsonValores & """total"": """ & rsVolume("totalAbertos") & """"
            jsonValores = jsonValores & "}"
            jsonValores = jsonValores & ",{"
            jsonValores = jsonValores & """label"": """ & "Finalizados" & """, "
            jsonValores = jsonValores & """total"": """ & rsVolume("totalFinalizados") & """"
            jsonValores = jsonValores & "}"
            jsonValores = jsonValores & "]" 

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(FormatarData(rsVolume("label"))) & """, "
            jsonArray = jsonArray & """total"": " & jsonValores & ""
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsVolume.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsVolume.Close()
    SET rsVolume = NOTHING
    
    ObterVolumeMensal = jsonArray
End Function

Function ObterSLAPorUrgencia()
    idDepartamento = Request("filtroDepartamento")
    idCategoria = Request("filtroCategoria")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    idAtendente = Request("filtroAtendente")
    codPlanta = Request("filtroFilial")
    status = Request("filtroStatus")
    prioridade = Request("filtroPrioridade")
    solicitante = Request("filtroSolicitante")

    SQL = "" 
    SQL = SQL & "DECLARE @DataInicio  DATETIME = '" & Replace(dataCriacaoDe, "-", "") & "'; "  
    SQL = SQL & "DECLARE @DataFim     DATETIME = " & "DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "'); " 
    
    SQL = SQL & "SELECT "
    SQL = SQL & "    p.Nivel, "
    SQL = SQL & "    COUNT(*) AS totalItens, "
    SQL = SQL & "    SUM(CASE "
    SQL = SQL & "        WHEN c.DataFinalizacao IS NOT NULL "
    SQL = SQL & "        AND c.DataFinalizacao <= c.DataPrevista" 
    SQL = SQL & "        THEN 1 ELSE 0 "
    SQL = SQL & "    END) AS Cumpridos, "
    SQL = SQL & "    SUM(CASE "
    SQL = SQL & "        WHEN (c.DataFinalizacao IS NOT NULL AND c.DataFinalizacao > c.DataPrevista) "
    SQL = SQL & "        OR (c.DataFinalizacao IS NULL     AND c.DataPrevista    < @DataFim) "
    SQL = SQL & "        THEN 1 ELSE 0 "
    SQL = SQL & "    END) AS Atrasados, "
    SQL = SQL & "    SUM(CASE "
    SQL = SQL & "        WHEN c.DataFinalizacao IS NULL AND c.DataPrevista >= @DataFim "
    SQL = SQL & "        THEN 1 ELSE 0 "
    SQL = SQL & "    END) AS EmAndamento, "
    SQL = SQL & "    SUM(CASE "
    SQL = SQL & "        WHEN c.DataFinalizacao IS NOT NULL "
    SQL = SQL & "        AND c.DataFinalizacao <= c.DataPrevista "
    SQL = SQL & "        THEN 1.0 ELSE 0 "
    SQL = SQL & "        END) * 100.0 "
    SQL = SQL & "    / "
    SQL = SQL & "    NULLIF( "
    SQL = SQL & "    SUM(CASE "
    SQL = SQL & "            WHEN c.DataFinalizacao IS NOT NULL "
    SQL = SQL & "            OR (c.DataFinalizacao IS NULL AND c.DataPrevista < @DataFim) "
    SQL = SQL & "            THEN 1 ELSE 0 "
    SQL = SQL & "        END), "
    SQL = SQL & "    0 "
    SQL = SQL & "    ) AS PercentualCumprido "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s on s.Id = c.IdSolicitante "
    SQL = SQL & "WHERE c.DataCriacao >= @DataInicio "

    If idCategoria <> "" Then
        SQL = SQL & "AND c.IdCategoria = " & idCategoria & " "
    End If

    If idDepartamento <> "" Then
        SQL = SQL & "AND c.IdDepartamento = " & idDepartamento & " "
    End If

     If idAtendente <> "" Then
        SQL = SQL & "AND c.IdAtendente = " & idAtendente & " "
    End If

    If solicitante <> "" Then
        SQL = SQL & "AND (s.Nome LIKE '%" & solicitante & "%' OR s.Email LIKE '%" & solicitante & "%') "
    End If

    If prioridade <> "" Then
        SQL = SQL & "AND p.Nivel = '" & prioridade & "' "
    End If

    If codPlanta <> "" Then
        SQL = SQL & "AND c.CodPlanta = '" & codPlanta & "' "
    End If

    If status <> "" Then
        If status = "PENDENTE" Then
            ' Status pendente é um filtro personalizado (tudo que não está finalizado ou cancelado)
            SQL = SQL & " AND c.Status NOT IN ('FINALIZADO', 'CANCELADO') "
        ElseIf status = "FINALIZADO_COM_ATRASO" Then
            'Também filtro personalizado (itens finalizados fora do prazo SLA)
            SQL = SQL & " AND c.Status = 'FINALIZADO' AND c.DataFinalizacao > c.DataPrevista "
        Else
            ' Outros status são os status padrão iguais do do registro
            SQL = SQL & " AND c.Status = '" & status & "' "
        End If
    End If

    SQL = SQL & "AND c.DataCriacao <  @DataFim "
    SQL = SQL & "GROUP BY p.Nivel "

    SQL = SQL & "ORDER BY "
    SQL = SQL & "CASE p.Nivel "
    SQL = SQL & "    WHEN 'URGENTE' THEN 1 "
    SQL = SQL & "    WHEN 'ALTA'    THEN 2 "
    SQL = SQL & "    WHEN 'MEDIA'   THEN 3 "
    SQL = SQL & "    WHEN 'BAIXA'   THEN 4 "
    SQL = SQL & "END;"

    On Error Resume Next
    SET rsSLA = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSLA) Then
        ObterSLAPorUrgencia = "[]"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSLA.EOF Then
        Do While Not rsSLA.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            percentualCumprido = Iif(IsNull(rsSLA("PercentualCumprido")), "0", rsSLA("PercentualCumprido"))
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """label"": """ & EscapeJSON(rsSLA("Nivel")) & """, "
            jsonArray = jsonArray & """total"": """ & EscapeJSON(Replace(percentualCumprido, ",", ".")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSLA.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSLA.Close()
    SET rsSLA = NOTHING
    
    ObterSLAPorUrgencia = jsonArray
End Function

Function FormatarData(data)
    If IsNull(data) Or data = "" Then
        FormatarData = ""
        Exit Function
    End If
    
    On Error Resume Next
    Dim dataFormatada
    dataFormatada = CDate(data)
    
    If Err.Number <> 0 Then
        FormatarData = ""
        Exit Function
    End If
    
    FormatarData = Right("0" & Month(dataFormatada), 2) & "/" & _
                             Year(dataFormatada)
    On Error Goto 0
End Function

Function ObterNomeFilial(codPlanta)
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