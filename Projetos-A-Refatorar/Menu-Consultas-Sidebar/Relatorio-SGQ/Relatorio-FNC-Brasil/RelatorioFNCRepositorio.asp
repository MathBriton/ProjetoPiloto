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
    Case "listarFNCs"
        resultado = ListarFNCs()
    Case "listarDefeitos"
        resultado = ListarDefeitos()
    Case "listarDisposicoes"
        resultado = ListarDisposicoes()
    Case "listarResponsaveis"
        resultado = ListarResponsaveis()
    Case "listarSetores"
        resultado = ListarSetores()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function ListarFNCs()
    Dim SQL, rsFNCs, jsonArray, primeiroItem

    idFNC = Request("filtroIdFNC")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idSetor = Request("filtroSetor")
    idResponsavel = Request("filtroResponsavel")
    idDefeito = Request("filtroDefeito")
    tipoDefeito = Request("filtroTipoDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    validacao = Request("filtroValidacao")
    disposicao = Request("filtroDisposicao")
    pedido = Request("filtroPedido")
    notaFiscal = Request("filtroNotaFiscal")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    dataEnvioClienteDe = Request("filtroDataEnvioClienteDe")
    dataEnvioClienteAte = Request("filtroDataEnvioClienteAte")

    'WITH PARA obter última data de resposta
    SQL = "WITH resp AS ("
    SQL = SQL & "   SELECT DISTINCT IdFNC, MIN(CASE WHEN r.DataFinalizacao IS NULL THEN r.DataInclusao ELSE r.DataFinalizacao END) AS DataResposta"
    SQL = SQL & "   FROM Resposta r "
    SQL = SQL & "   GROUP BY IdFNC"
    SQL = SQL & ") "

    SQL = SQL & "SELECT f.Id, f.Planta as CodFilial, cli.CIDADE as NomeFilial, f.Cliente as CodCliente, f.Loja as CodLoja, f.Situacao, "
    SQL = SQL & "f.Status as Validacao, f.IdDisposicao, dis.Descricao AS Disposicao, CONCAT(cli.MUNICIPIO_CLIENTE, ' - ', cli.ESTADO_CLIENTE) AS CidadeEstado,"
    SQL = SQL & "cli.NOME_CLIENTE as NomeCliente, f.Quantidade, ISNULL(f.Problema, def.Descricao) AS Problema, "
    SQL = SQL & "f.DataInclusao, f.DataOcorrencia, f.DataEnvioCliente, f.DataEncerramento, def.Descricao AS DescricaoDefeito, def.Tipo as TipoDefeito, "
    SQL = SQL & "res.Nome as NomeResponsavel, analista.Nome as NomeAnalista, f.DataConclusaoPrevista, f.DataProducao, "
    SQL = SQL & "("
    SQL = SQL & "   ISNULL(f.ValorTraslados, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorRefeicao, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorFrete, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorEstadias, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorDiaria, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorRefugo, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorRetrabalho, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorMaterial, 0) + " 
    SQL = SQL & "   ISNULL(f.ValorDesconto, 0) " 
    SQL = SQL & ") AS ValorAgregado, "
    SQL = SQL & "isnull(f.ValorReclamado,0) ValorReclamado, " 
    SQL = SQL & "f.Quantidade / (SELECT SUM(CASE WHEN p.Quantidade < 1 THEN 1 ELSE p.Quantidade END) FROM PedidoFnc p WHERE p.IdFNC = f.Id AND Excluido IS NULL) * 100 AS CoeficienteNQA,"
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Maquina), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Maquina, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Turno), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Turno, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS DescricaoDisposicao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), isnull(f.Eficacia,'')), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Eficacia, "
    SQL = SQL & "p.Pedidos, p.PrecoUnitario, p.Referencia, p.Quantidade as QuantidadePedidos, " 
    SQL = SQL & "DATEDIFF(DAY, CASE WHEN f.DataOcorrencia IS NULL THEN f.DataInclusao ELSE f.DataOcorrencia END, CASE WHEN f.DataEnvioCliente IS NULL THEN GETDATE() ELSE f.DataEnvioCliente END)"
    SQL = SQL & "- (DATEDIFF(WEEK, CASE WHEN f.DataOcorrencia IS NULL THEN f.DataInclusao ELSE f.DataOcorrencia END, CASE WHEN f.DataEnvioCliente IS NULL THEN GETDATE() ELSE f.DataEnvioCliente END) * 2) "
    SQL = SQL & "- CASE WHEN DATENAME(WEEKDAY, CASE WHEN f.DataOcorrencia IS NULL THEN f.DataInclusao ELSE f.DataOcorrencia END) = 'Sunday'  THEN 1 ELSE 0 END "
    SQL = SQL & "- CASE WHEN DATENAME(WEEKDAY, CASE WHEN f.DataEnvioCliente IS NULL THEN GETDATE() ELSE f.DataEnvioCliente END) = 'Saturday' THEN 1 ELSE 0 END "
    SQL = SQL & "AS DiasAteResposta, "
    SQL = SQL & "DATEDIFF(DAY, f.DataInclusao, CASE WHEN f.DataEncerramento IS NULL THEN GETDATE() ELSE f.DataEncerramento END) as DiasAteEncerramento, "
    SQL = SQL & "p.ValorPedido, p.CodProduto, p.GrupoProduto, nf.NotasFiscais, s.Descricao as Setor, "
    SQL = SQL & "nf.DataFaturamento, CASE WHEN f.Aprovada = 1 THEN '1' ELSE '0' END AS Status "
    SQL = SQL & "FROM FNC f WITH(NOLOCK) "
    SQL = SQL & "LEFT JOIN Disposicao dis ON f.IdDisposicao = dis.Id "
    SQL = SQL & "LEFT JOIN Responsavel analista WITH(NOLOCK) ON f.IdResponsavel = analista.Id AND analista.Planta = f.Planta AND analista.Tipo = 1 AND analista.Excluido IS NULL "
    SQL = SQL & "LEFT JOIN Responsavel res WITH(NOLOCK) ON f.IdAnalista = res.Id AND res.Planta = f.Planta AND res.Tipo = 2 AND res.Excluido IS NULL "
    SQL = SQL & "LEFT JOIN vw_PedidosFNC p ON f.Id = p.IdFNC " 
    SQL = SQL & "INNER JOIN Defeito def WITH(NOLOCK) ON f.IdDefeito = def.Id "
    SQL = SQL & "LEFT JOIN Setor s ON f.IdSetor = s.Id " 
    SQL = SQL & "LEFT JOIN vw_NotasFiscaisFNC nf ON f.Id = nf.IdFNC " 
    SQL = SQL & "LEFT JOIN resp r ON r.IdFNC = f.Id "
    SQL = SQL & "OUTER APPLY ( "
    SQL = SQL & "   SELECT TOP 1 NOME_CLIENTE, CIDADE, LTRIM(RTRIM(MUNICIPIO_CLIENTE)) AS MUNICIPIO_CLIENTE, ESTADO_CLIENTE, CNPJ_CLIENTE "
    SQL = SQL & "   FROM BI_CLIENTES "
    SQL = SQL & "   WHERE CIDADE = "
    SQL = SQL & "   CASE f.Planta "
    SQL = SQL & "       WHEN '00' THEN 'PIRAPETINGA' "
    SQL = SQL & "       WHEN '04' THEN 'UBERABA' "
    SQL = SQL & "       WHEN '05' THEN 'BENTO' "
    SQL = SQL & "       WHEN '06' THEN 'MARANGUAPE' "
    SQL = SQL & "       WHEN '07' THEN 'SAQUAREMA' "
    SQL = SQL & "       WHEN '08' THEN 'ARACATUBA' "
    SQL = SQL & "       WHEN '09' THEN 'BLUMENAU' "
    SQL = SQL & "       WHEN '10' THEN 'PACAJUS' "
    SQL = SQL & "       WHEN '11' THEN 'PORTO FELIZ' "
    SQL = SQL & "       WHEN '12' THEN 'TRES BARRAS' "
    SQL = SQL & "       WHEN '13' THEN 'FLORESTAL' "
    SQL = SQL & "   END "
    SQL = SQL & "   AND CASE WHEN f.Planta IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END = f.Cliente COLLATE SQL_Latin1_General_CP1_CI_AS "
    SQL = SQL & "   AND CASE WHEN f.Planta IN ('09') THEN '01' ELSE LOJA_CLIENTE END = f.Loja COLLATE SQL_Latin1_General_CP1_CI_AS "
    SQL = SQL &  ") AS cli "
    SQL = SQL & "WHERE ISNULL(f.Excluido, 0) = 0 "

    If usuarioLogado("nivel") = 1 Then
        SQL = SQL & "AND f.QuemIncluiu = '" & usuarioLogado("id") & "' OR f.QuemIncluiu = '" & usuarioLogado("codigo") & "' "
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

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If pedido <> "" Then
        SQL = SQL & " AND f.Id IN ( SELECT DISTINCT IdFNC FROM PedidoFnc WHERE Pedido = '" & pedido & "' ) "
    End If

    If notafiscal <> "" Then
        SQL = SQL & " AND f.Id IN ( SELECT DISTINCT IdFNC FROM NotaFiscal WHERE NotaFiscal = '" & notafiscal & "' ) "
    End If

    If idSetor <> "" Then
        SQL = SQL & " AND f.IdSetor IN (" & idSetor & ")"
    End If

    If idResponsavel <> "" Then
        SQL = SQL & " AND f.IdAnalista IN (" & idResponsavel & ")"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito IN (" & idDefeito & ")"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND f.Id IN (" & idFNC & ")"
    End If

    If tipoDefeito <> "" Then
        primeiroItem = True
        itens = Split(tipoDefeito, ",")

        SQL = SQL & " AND def.Tipo IN ("

        For Each item In itens
            If Not primeiroItem Then SQL = SQL & "," End If

            SQL = SQL & "'" & item & "'"
            primeiroItem = False
        Next

        SQL = SQL & ")"
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

    SQL = SQL & " ORDER BY f.Id DESC"

    On Error Resume Next
    SET rsFNCs = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsFNCs) Then
        ListarFNCs = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de respostas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsFNCs.EOF Then
        Do While Not rsFNCs.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            coeficienteNQA = Iif(ISNULL(rsFNCs("CoeficienteNQA")), 0, rsFNCs("CoeficienteNQA"))
 
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """idFNC"": """ & EscapeJSON(rsFNCs("Id")) & """, "
            jsonArray = jsonArray & """nomeFilial"": """ & EscapeJSON(rsFNCs("NomeFilial")) & """, "
            jsonArray = jsonArray & """codFilial"": """ & EscapeJSON(rsFNCs("CodFilial")) & """, "
            jsonArray = jsonArray & """codCliente"": """ & EscapeJSON(rsFNCs("CodCliente")) & """, "
            jsonArray = jsonArray & """codLoja"": """ & EscapeJSON(rsFNCs("CodLoja")) & """, "
            jsonArray = jsonArray & """situacao"": """ & EscapeJSON(rsFNCs("Situacao")) & """, "
            jsonArray = jsonArray & """validacao"": """ & EscapeJSON(rsFNCs("Validacao")) & """, "
            jsonArray = jsonArray & """maquina"": """ & EscapeJSON(rsFNCs("Maquina")) & """, "
            jsonArray = jsonArray & """turno"": """ & EscapeJSON(rsFNCs("Turno")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsFNCs("Status")) & """, "
            jsonArray = jsonArray & """idDisposicao"": """ & EscapeJSON(rsFNCs("IdDisposicao")) & """, "
            jsonArray = jsonArray & """disposicao"": """ & EscapeJSON(rsFNCs("Disposicao")) & """, "
            jsonArray = jsonArray & """coeficienteNQA"": """ & EscapeJSON(Replace(coeficienteNQA, ",", ".")) & """, "
            jsonArray = jsonArray & """descricaoDisposicao"": """ & EscapeJSON(rsFNCs("DescricaoDisposicao")) & """, "
            jsonArray = jsonArray & """eficacia"": """ & EscapeJSON(rsFNCs("Eficacia")) & """, "
            jsonArray = jsonArray & """cidadeEstado"": """ & EscapeJSON(rsFNCs("CidadeEstado")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsFNCs("NomeCliente")) & """, "
            jsonArray = jsonArray & """quantidade"": """ & EscapeJSON(rsFNCs("Quantidade")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsFNCs("DataInclusao")) & """, "
            jsonArray = jsonArray & """dataOcorrencia"": """ & EscapeJSON(rsFNCs("DataOcorrencia")) & """, "
            jsonArray = jsonArray & """dataEnvioCliente"": """ & EscapeJSON(rsFNCs("DataEnvioCliente")) & """, "
            jsonArray = jsonArray & """dataEncerramento"": """ & EscapeJSON(rsFNCs("DataEncerramento")) & """, "
            jsonArray = jsonArray & """dataConclusaoPrevista"": """ & EscapeJSON(rsFNCs("DataConclusaoPrevista")) & """, "
            jsonArray = jsonArray & """dataProducao"": """ & EscapeJSON(rsFNCs("DataProducao")) & """, "
            jsonArray = jsonArray & """descricaoDefeito"": """ & EscapeJSON(rsFNCs("DescricaoDefeito")) & """, "
            jsonArray = jsonArray & """nomeResponsavel"": """ & EscapeJSON(rsFNCs("NomeResponsavel")) & """, "
            jsonArray = jsonArray & """nomeAnalista"": """ & EscapeJSON(rsFNCs("NomeAnalista")) & """, "
            jsonArray = jsonArray & """valorAgregado"": """ & EscapeJSON(rsFNCs("ValorAgregado")) & """, "
            jsonArray = jsonArray & """valorReclamado"": """ & EscapeJSON(rsFNCs("ValorReclamado")) & """, "
            jsonArray = jsonArray & """pedidos"": """ & EscapeJSON(rsFNCs("Pedidos")) & """, "
            jsonArray = jsonArray & """precoUnitario"": """ & EscapeJSON(rsFNCs("PrecoUnitario")) & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsFNCs("Referencia")) & """, "
            jsonArray = jsonArray & """quantidadePedidos"": """ & EscapeJSON(rsFNCs("QuantidadePedidos")) & """, "
            jsonArray = jsonArray & """valorPedido"": """ & EscapeJSON(rsFNCs("ValorPedido")) & """, "
            jsonArray = jsonArray & """codProduto"": """ & EscapeJSON(rsFNCs("CodProduto")) & """, "
            jsonArray = jsonArray & """grupoProduto"": """ & EscapeJSON(rsFNCs("GrupoProduto")) & """, "
            jsonArray = jsonArray & """notasFiscais"": """ & EscapeJSON(rsFNCs("NotasFiscais")) & """, "
            jsonArray = jsonArray & """dataFaturamento"": """ & EscapeJSON(rsFNCs("DataFaturamento")) & """, "
            jsonArray = jsonArray & """diasAteResposta"": """ & EscapeJSON(rsFNCs("DiasAteResposta")) & """, "
            jsonArray = jsonArray & """diasAteEncerramento"": """ & EscapeJSON(rsFNCs("DiasAteEncerramento")) & """, "
            jsonArray = jsonArray & """tipoDefeito"": """ & EscapeJSON(rsFNCs("TipoDefeito")) & """, "
            jsonArray = jsonArray & """setor"": """ & EscapeJSON(rsFNCs("Setor")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsFNCs.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsFNCs.Close()
    SET rsFNCs = NOTHING
    
    ListarFNCs = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
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


Function ListarDisposicoes()
    Dim SQL, rsDisposicoes, jsonArray, primeiroItem

    SQL = "SELECT Id, Descricao FROM Disposicao"

    On Error Resume Next
    SET rsDisposicoes = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDisposicoes) Then
        ListarDisposicoes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de disposições: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDisposicoes.EOF Then
        Do While Not rsDisposicoes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsDisposicoes("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsDisposicoes("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDisposicoes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDisposicoes.Close()
    SET rsDisposicoes = NOTHING
    
    ListarDisposicoes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function


Function ListarResponsaveis()
    Dim SQL, rsResponsaveis, jsonArray, primeiroItem

    codplanta = Request("filtroFilial")

    SQL = "SELECT DISTINCT Id, Nome FROM Responsavel "
    SQL = SQL & "WHERE ISNULL(Excluido,'') = '' AND Tipo = 2"

    If codPlanta <> "" Then
        SQL = SQL & " AND Planta = " & codPlanta
    End If

    On Error Resume Next
    SET rsResponsaveis = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsResponsaveis) Then
        ListarResponsaveis = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de responsáveis: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsResponsaveis.EOF Then
        Do While Not rsResponsaveis.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsResponsaveis("Id")) & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsResponsaveis("Nome")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsResponsaveis.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsResponsaveis.Close()
    SET rsResponsaveis = NOTHING
    
    ListarResponsaveis = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarSetores()
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

%>