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
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "nivel", Session("NIVEL")
usuarioLogado.Add "cod", Session("CODUSER")
usuarioLogado.Add "planta", SESSION("FILIAL")

' Limite NQA
Set limitesNQA = CreateObject("Scripting.Dictionary")
limitesNQA.Add "CRITICO", 0.5
limitesNQA.Add "GRAVE", 0.8

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
    Case "editarSolicitante"
        resultado = EditarSolicitante()
    Case "avaliar"
        resultado = Avaliar()
    Case "responder"
        resultado = Responder()
    Case "cancelar"
        resultado = Cancelar()
    Case "enviarParaResponsavel"
        resultado = EnviarParaResponsavel()
    Case "enviarParaAprovacao"
        resultado = EnviarParaAprovacao()
    Case "excluir"
        resultado = Excluir()
    Case "obterPorId"
        resultado = ObterPorId()
    Case "obterLimitesNQA"
        resultado = ObterLimitesNQA()
    Case "obterNivelUsuario"
        resultado = ObterNivelUsuario()
    Case "listarTiposDefeito"
        resultado = ListarTiposDefeito()
    Case "listarSituacoes"
        resultado = ListarSituacoes()
    Case "listarStatus"
        resultado = ListarStatus()
    Case "listarResponsaveis"
        resultado = ListarResponsaveis()
    Case "listarSetores"
        resultado = ListarSetores()
    Case "listarDefeitos"
        resultado = ListarDefeitos()
    Case "listarDisposicoes"
        resultado = ListarDisposicoes()
    Case "listarNotasCliente"
        resultado = ListarNotasCliente()
    Case "listarAprovadores"
        resultado = ListarAprovadores()
    Case "listarHistoricoRespostas"
        resultado = ListarHistoricoRespostas()
    Case "listarHistoricoAprovacoes"
        resultado = ListarHistoricoAprovacoes()
    Case "listarHistoricoObservacoes"
        resultado = ListarHistoricoObservacoes()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select


' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    idFNC = Request("filtroIdFNC")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    codPlanta = Request("filtroFilial")
    status = Request("filtroStatus")
    situacao = Request("filtroSituacao")
    aprovada = Request("filtroAprovada")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    tipoReclamacao = Request("filtroTipoReclamacao")
    flagApenasPendentes = Request("flagApenasPendentes")

    nivelUsuario = ""
    plantasUsuario = ""

    respostaNivelUsuario = ObterNivelUsuario()
    Set respostaObj = JsonToDictionary(respostaNivelUsuario)
    
    If Not respostaObj("sucesso") Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário""}"
        Exit Function
    End If

    Set dadosUsuario = respostaObj("dados")
    nivelUsuario = dadosUsuario("nivel")
    plantasVisiveis = dadosUsuario("plantas")
    setoresVisiveis = dadosUsuario("setores")

    If nivelUsuario = "" OR plantasVisiveis = "" Then
        ValidarSolicitacao = "{""sucesso"": false, ""mensagem"": ""Nível de usuário inválido""}"
        Exit Function
    End If

    'MONTA clausula planta
    clausulaPlanta = ""
    primeiroItem = True
    arrayPlantas = Split(plantasVisiveis, ",")

    clausulaPlanta = clausulaPlanta & " AND f.Planta in ("
    For Each item In arrayPlantas
        If Not primeiroItem Then 
            clausulaPlanta = clausulaPlanta & "," 
        End If

        clausulaPlanta = clausulaPlanta & "'" & trim(item) & "'"
        primeiroItem = False
    Next

    clausulaPlanta = clausulaPlanta & ")"

    'MONTA clausula setor
    clausulaSetor = ""
    primeiroItem = True
    arraySetores = Split(setoresVisiveis, ",")

    clausulaSetor = clausulaSetor & " AND f.IdSetor in ("
    For Each item In arraySetores
        If Not primeiroItem Then 
            clausulaSetor = clausulaSetor & "," 
        End If

        clausulaSetor = clausulaSetor & "'" & trim(item) & "'"
        primeiroItem = False
    Next

    clausulaSetor = clausulaSetor & ")"

    SQL = "SELECT TOP 200 f.Id, f.Planta as CodFilial, cli.CIDADE as NomeFilial, f.ValorReclamado, d.Tipo as tipoDefeito,"
    SQL = SQL & "f.Cliente as CodCliente, f.Loja as CodLoja, f.Situacao, f.Status, f.IdAnalista as ResponsavelId,"
    SQL = SQL & "CONVERT(varchar,f.DataInclusao,103) as DataInclusao, f.Quantidade as QuantidadeReclamada, "
    SQL = SQL & "d.Descricao AS DescricaoTipoDefeito, dis.Descricao AS DescricaoDisposicao, f.TipoReclamacao, "
    SQL = SQL & "CASE WHEN f.ComSetorResponsavel = 1 THEN 1 ELSE 0 END AS ComSetorResponsavel, "
    SQL = SQL & "cli.NOME_CLIENTE as NomeCliente, CONCAT(cli.MUNICIPIO_CLIENTE, ' - ', cli.ESTADO_CLIENTE) AS CidadeEstado, "
    SQL = SQL & "t2.CRM_NOME as NomeResponsavel, f.NivelAprovacaoAtual, s.Descricao as NomeSetor, "
	SQL = SQL & "f.Quantidade / (SELECT SUM(CASE WHEN p.Quantidade < 1 THEN 1 ELSE p.Quantidade END) FROM PedidoFnc p WHERE p.IdFNC = f.Id AND Excluido IS NULL) * 100 AS CoeficienteNQA, "
    SQL = SQL & "CASE isnull(f.Aprovada,0) WHEN 1 THEN 1 ELSE 0 END AS Aprovada "
    SQL = SQL & "FROM FNC f "
    SQL = SQL & "LEFT JOIN Setor s ON f.IdSetor = s.Id " 
    SQL = SQL & "LEFT JOIN Disposicao dis ON f.IdDisposicao = dis.Id "
    SQL = SQL & "LEFT JOIN Defeito d ON d.Id = f.IdDefeito "
    SQL = SQL & "LEFT JOIN CRM_DTI t2 ON t2.CRM_R_E_C_N_O = f.IdAnalista "
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
    SQL = SQL & "WHERE 1=1"

    If nivelUsuario = "SOLICITANTE" Then
        'Solicitante visualiza apenas o que ele incluiu
        SQL = SQL & " AND ("
        SQL = SQL & "  f.QuemIncluiu = '" & usuarioLogado("id") & "'"
        SQL = SQL & "  OR f.QuemIncluiu = '" & usuarioLogado("codigo") & "'" 
        SQL = SQL & ") " 
    ElseIf nivelUsuario = "RESPONSAVEL" Then
        ' Responsável visualizam tudo das suas plantas
        SQL = SQL & clausulaPlanta & " "
    ElseIf nivelUsuario = "APROVADOR" Then
        ' Aprovador visualiza tudo das suas plantas e dos seus setores
        ' Só filtra setores se o aprovador for apenas nivel 1         
        If Not AprovadorNivelMaiorQueUm(usuarioLogado("id")) then
            SQL = SQL & clausulaSetor & " "
        End If

        SQL = SQL & clausulaPlanta & " "
    ElseIf nivelUsuario = "ANALISTA" Then
        ' Analista visualiza tudo das suas plantas
        SQL = SQL & clausulaPlanta & " "
    End If

    If flagApenasPendentes = "1" Then
        If nivelUsuario = "SOLICITANTE" Then
            ' Pendências de solicitante: tudo que ainda não está com situação encerrada
            SQL = SQL & " AND f.Situacao <> 3 "
        ElseIf nivelUsuario = "ANALISTA" Then
            ' Pendências de analista: Abertas, Respondidas e Aprovadas
            SQL = SQL & " AND ( "
            SQL = SQL & "    (f.Situacao = 1 and (f.Aprovada <> 1 or f.Aprovada is null) and (f.NivelAprovacaoAtual = 0 or f.NivelAprovacaoAtual is null) and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null)) "
            SQL = SQL & "    OR (f.Situacao = 2 and (f.Aprovada <> 1 or f.Aprovada is null) and (f.NivelAprovacaoAtual = 0 or f.NivelAprovacaoAtual is null) and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null)) "
            SQL = SQL & "    OR (f.Situacao = 2 and f.Aprovada = 1 and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null)) "
            SQL = SQL & ")"
            SQL = SQL & " AND f.TipoReclamacao = 'FNC' "
        ElseIf nivelUsuario = "RESPONSAVEL" Then
            ' Pendências de responsável: tudo que está pendente de resposta
            SQL = SQL & " AND f.IdAnalista = '" & usuarioLogado("id") & "' "
            SQL = SQL & " AND f.ComSetorResponsavel = 1"
            SQL = SQL & " AND EXISTS ("
            SQL = SQL & "   SELECT TOP 1 Id FROM Resposta r "
            SQL = SQL & "   WHERE r.DataFinalizacao IS NULL "
            SQL = SQL & "   AND f.IdAnalista = r.QuemFinalizou "
            SQL = SQL & "   AND r.IdFNC = f.Id"
            SQL = SQL & ")"
            SQL = SQL & " AND f.TipoReclamacao = 'FNC' "
        ElseIf nivelUsuario = "APROVADOR" Then
            ' Pendências de aprovador: tudo que está pendente de aprovação
            idsFNC = ObterIdsFNCAprovador(usuarioLogado("id"))
            SQL = SQL & " AND f.DataAvaliacao IS NULL"
            SQL = SQL & " AND f.Situacao = 2"
            SQL = SQL & " AND f.Id IN (" & Iif (idsFNC = "", 0, idsFNC) & ")"
            SQL = SQL & " AND f.TipoReclamacao = 'FNC' "
        End If
    End If

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
    End If
    
    If idFNC <> "" Then
        SQL = SQL & " AND f.Id = " & idFNC & ""
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND f.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND f.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao = " & situacao & ""
    End If

    If aprovada <> "" Then
        SQL = SQL & Iif (aprovada = "1", "AND f.Aprovada = 1", "AND ISNULL(f.Aprovada, 0) = 0")
    End If

    If status <> "" Then
        SQL = SQL & " AND f.Status = " & status & ""
    End If

    If tipoReclamacao <> "" Then
        SQL = SQL & " AND f.TipoReclamacao = '" & tipoReclamacao & "'"
    End If

    SQL = SQL & " ORDER BY f.Id DESC"

    On Error Resume Next
    SET rsListar = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de FNCs: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            coeficienteNQA = Iif (isNull(rsListar("CoeficienteNQA")), 0, rsListar("CoeficienteNQA"))

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id")) & """, "
            jsonArray = jsonArray & """comSetorResponsavel"": """ & EscapeJSON(rsListar("ComSetorResponsavel")) & """, "
            jsonArray = jsonArray & """nivelAprovacaoAtual"": """ & EscapeJSON(rsListar("NivelAprovacaoAtual")) & """, "
            jsonArray = jsonArray & """nomeSetor"": """ & EscapeJSON(rsListar("NomeSetor")) & """, "
            jsonArray = jsonArray & """tipoDefeito"": """ & EscapeJSON(rsListar("TipoDefeito")) & """, "
            jsonArray = jsonArray & """tipoReclamacao"": """ & EscapeJSON(rsListar("TipoReclamacao")) & """, "
            jsonArray = jsonArray & """responsavelId"": """ & EscapeJSON(rsListar("ResponsavelId")) & """, "
            jsonArray = jsonArray & """codFilial"": """ & EscapeJSON(rsListar("CodFilial")) & """, "
            jsonArray = jsonArray & """coeficienteNQA"": """ & EscapeJSON(Replace(coeficienteNQA, ",", ".")) & """, "
            jsonArray = jsonArray & """nomeFilial"": """ & EscapeJSON(rsListar("NomeFilial")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsListar("NomeCliente")) & """, "
            jsonArray = jsonArray & """nomeResponsavel"": """ & EscapeJSON(rsListar("NomeResponsavel")) & """, "
            jsonArray = jsonArray & """descricaoDisposicao"": """ & EscapeJSON(rsListar("DescricaoDisposicao")) & """, "
            jsonArray = jsonArray & """descricaoTipoDefeito"": """ & EscapeJSON(rsListar("DescricaoTipoDefeito")) & """, "
            jsonArray = jsonArray & """codCliente"": """ & EscapeJSON(rsListar("CodCliente")) & """, "
            jsonArray = jsonArray & """codLoja"": """ & EscapeJSON(rsListar("CodLoja")) & """, "
            jsonArray = jsonArray & """situacao"": """ & EscapeJSON(rsListar("Situacao")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao")) & """, "
            jsonArray = jsonArray & """aprovada"": """ & EscapeJSON(rsListar("Aprovada")) & """, "
            jsonArray = jsonArray & """cidadeEstado"": """ & EscapeJSON(rsListar("CidadeEstado")) & """, "
            jsonArray = jsonArray & """quantidadeReclamada"": """ & EscapeJSON(rsListar("QuantidadeReclamada")) & """"
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
    idFNC = Request("id")

    SQL = "SELECT f.Id, f.Planta as CodFilial, cli.CIDADE as NomeFilial, d.Tipo as tipoDefeito, f.DataEncerramento, "
    SQL = SQL & "f.Contato, f.Telefone, f.Email, f.IdAnalista as ResponsavelId, f.IdSetor as SetorId, "
    SQL = SQL & "f.Cliente as CodCliente, f.Loja as CodLoja, f.Situacao, f.Status, t3.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), f.Eficacia), 'json') AS Eficacia, f.TipoReclamacao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao), 'json') AS DescricaoDisposicao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), f.Descricao), 'json') AS DescricaoDefeito, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), dis.Descricao), 'json') AS Disposicao, "
    SQL = SQL & "FORMAT(f.DataInclusao,'yyyy-MM-dd') as DataInclusao, f.Quantidade as QuantidadeReclamada, "
    SQL = SQL & "f.ValorReclamado, cli.CNPJ_CLIENTE as CNPJ, cli.COD_VENDEDOR AS CodVendedor, cli.NOME_VENDEDOR AS NomeVendedor, "
    SQL = SQL & "cli.NOME_CLIENTE as NomeCliente, CONCAT(cli.MUNICIPIO_CLIENTE, ' - ', cli.ESTADO_CLIENTE) AS CidadeEstado, "
    SQL = SQL & "FORMAT(f.DataOcorrencia,'yyyy-MM-dd') as DataOcorrencia, f.NivelAprovacaoAtual, "
    SQL = SQL & "t2.CRM_NOME as NomeResponsavel, f.IdDefeito as DefeitoId, f.IdDisposicao as DisposicaoId, "
    SQL = SQL & "f.Turno, f.Maquina, FORMAT(f.DataConclusaoPrevista,'yyyy-MM-dd') as DataConclusaoPrevista,"
    SQL = SQL & "FORMAT(f.DataProducao,'yyyy-MM-dd') as DataProducao, d.Descricao AS DescricaoTipoDefeito, "
    SQL = SQL & "CASE WHEN f.ComSetorResponsavel = 1 THEN 1 ELSE 0 END AS ComSetorResponsavel, "
    SQL = SQL & "CASE f.DataEnvioCliente WHEN '1900-01-01' THEN '' ELSE FORMAT(f.DataEnvioCliente,'yyyy-MM-dd') END as DataEnvioCliente,"
    SQL = SQL & "CASE isnull(f.Aprovada,0) WHEN 1 THEN 1 ELSE 0 END AS Aprovada, f.DefeitoIshikawa, f.Assunto5W2H, "
    SQL = SQL & "f.Quantidade / (SELECT SUM(CASE WHEN p.Quantidade < 1 THEN 1 ELSE p.Quantidade END) FROM PedidoFnc p WHERE p.IdFNC = f.Id AND Excluido IS NULL) * 100 AS CoeficienteNQA "
    SQL = SQL & "FROM FNC f "
    SQL = SQL & "LEFT JOIN Disposicao dis ON f.IdDisposicao = dis.Id "
    SQL = SQL & "LEFT JOIN Defeito d ON d.Id = f.IdDefeito "
    SQL = SQL & "LEFT JOIN CRM_DTI t2 ON t2.CRM_R_E_C_N_O = f.IdAnalista "
    SQL = SQL & "LEFT JOIN CRM_DTI t3 ON t3.CRM_R_E_C_N_O = f.QuemIncluiu "
    SQL = SQL & "OUTER APPLY ( "
    SQL = SQL & "   SELECT TOP 1 NOME_CLIENTE, CIDADE, LTRIM(RTRIM(MUNICIPIO_CLIENTE)) AS MUNICIPIO_CLIENTE, ESTADO_CLIENTE, CNPJ_CLIENTE, "
    SQL = SQL & "   COD_VENDEDOR, NOME_VENDEDOR "
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
    SQL = SQL & "WHERE f.Id = '" & idFNC & "'"
    
    On Error Resume Next
    SET rsObter = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de FNCs: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""FNC nao encontrada.""}"
        Exit Function
    End If

    ' Obtem Notas
    SQL = "SELECT DISTINCT n.Pedido as NumeroPedido, n.IdFNC, n.Item as ItemPedido, p.PrecoUnitario, "
    SQL = SQL & "p.Quantidade, p.CodProduto, p.GrupoProduto, n.NotaFiscal, p.QuantidadeReclamada, "
    SQL = SQL & "n.ItemNotaFiscal, n.Serie as SerieNotaFiscal, n.DataEmissao as DataEmissaoNota, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), p.Referencia), 'json') AS Referencia "
    SQL = SQL & "FROM NotaFiscal n "
    SQL = SQL & "LEFT JOIN PedidoFnc p on p.IdFNC = n.IdFNC AND n.Pedido = p.Pedido AND (n.Quantidade = 0 OR n.Quantidade = p.Quantidade) "
    SQL = SQL & "WHERE n.IdFNC = '" & idFNC & "' "
    SQL = SQL & "AND n.Excluido is null AND p.Excluido is null"

    On Error Resume Next
    SET rsNotas = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsNotas) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    notasJSON = "["

    If Not rsNotas.EOF Then
        Do While Not rsNotas.EOF
            If Not primeiroItem Then notasJSON = notasJSON & "," End If
            referencia = Iif(IsNull(rsNotas("Referencia")), "", rsNotas("Referencia"))

            notasJSON = notasJSON & "{"
            notasJSON = notasJSON & """numeroPedido"": """ & rsNotas("NumeroPedido") & """, "
            notasJSON = notasJSON & """idFNC"": """ & rsNotas("idFNC") & """, "
            notasJSON = notasJSON & """itemPedido"": """ & rsNotas("ItemPedido") & """, "
            notasJSON = notasJSON & """precoUnitario"": """ & rsNotas("PrecoUnitario") & """, "
            notasJSON = notasJSON & """quantidade"": """ & rsNotas("Quantidade") & """, "
            notasJSON = notasJSON & """quantidadeReclamada"": """ & rsNotas("QuantidadeReclamada") & """, "
            notasJSON = notasJSON & """codProduto"": """ & rsNotas("CodProduto") & """, "
            notasJSON = notasJSON & """referencia"": """ & EscapeJSON(Replace(referencia, "'", "")) & """, "
            notasJSON = notasJSON & """grupoProduto"": """ & rsNotas("GrupoProduto") & """, "
            notasJSON = notasJSON & """notaFiscal"": """ & rsNotas("NotaFiscal") & """, "
            notasJSON = notasJSON & """itemNotaFiscal"": """ & rsNotas("ItemNotaFiscal") & """, "
            notasJSON = notasJSON & """serieNotaFiscal"": """ & rsNotas("SerieNotaFiscal") & """, "
            notasJSON = notasJSON & """dataEmissaoNota"": """ & rsNotas("DataEmissaoNota") & """"
            notasJSON = notasJSON & "}"
            
            primeiroItem = False
            rsNotas.MoveNext()
        Loop
    End If

    notasJSON = notasJSON & "]"
    rsNotas.Close()
    SET rsNotas = NOTHING

    ' Obtem lista de anexos
    SQL = "SELECT idFNCDocumento, idFNC, Filial, Nome, Usuario, Caminho, "
    SQL = SQL & "FORMAT(DataDaOperacao , 'MM/dd/yyyy HH:mm:ss') AS DataDaOperacao "
    SQL = SQL & "FROM FNCDocumento "
    SQL = SQL & "WHERE idFNC = " & idFNC
    SQL = SQL & "ORDER BY DataDaOperacao Desc"

    On Error Resume Next
    SET rsAnexos = Conecta("BEN-CORPORATIVO", SQL)
    
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
            anexosJSON = anexosJSON & """id"": """ & EscapeJSON(rsAnexos("idFNCDocumento")) & """, "
            anexosJSON = anexosJSON & """idFNC"": """ & EscapeJSON(rsAnexos("idFNC")) & """, "
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

    ' Obtem lista de aprovações
    SQL = "SELECT "
    SQL = SQL & "   Aprovacao.IdAprovador, "
    SQL = SQL & "   Aprovacao.IdUsuarioInclusao, "
    SQL = SQL & "   u.CRM_NOME as NomeUsuarioAvaliacao, "
    SQL = SQL & "   u2.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "   CASE WHEN DataCriacaoRegistro >= '20260504'"
    SQL = SQL & "   THEN "
    SQL = SQL & "       CASE Aprovacao.Nivel "
    SQL = SQL & "           WHEN 1 THEN 'Gerente Setor'"
    SQL = SQL & "           WHEN 2 THEN 'Gerente Qualidade'"
    SQL = SQL & "           WHEN 3 THEN 'Controller'"
    SQL = SQL & "           WHEN 4 THEN 'Gerente Geral'"
    SQL = SQL & "       END"
    SQL = SQL & "   ELSE "
    SQL = SQL & "       CASE Aprovacao.Nivel "
    SQL = SQL & "           WHEN 1 THEN 'Responsavel Qualidade' "
    SQL = SQL & "           WHEN 2 THEN 'Responsavel Producao' "
    SQL = SQL & "           WHEN 3 THEN 'Responsavel Comercial' "
    SQL = SQL & "           WHEN 4 THEN 'Controller' "
    SQL = SQL & "           WHEN 5 THEN 'Gerencia Geral' " 
    SQL = SQL & "       END "
    SQL = SQL & "   END AS Nivel, "
    SQL = SQL & "   CASE isnull(Avaliacao,0) WHEN 0 THEN 'Não' ELSE 'Sim' END Avaliacao, "
    SQL = SQL & "   DataCriacaoRegistro as DataInclusao, "
    SQL = SQL & "   DataAvaliacao, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), Motivo), 'json') AS Motivo, "
    SQL = SQL & "   CASE WHEN Invalido = 1 THEN '1' ELSE '0' END as Invalido "
    SQL = SQL & "FROM AprovacaoFNC Aprovacao "
    SQL = SQL & "LEFT JOIN CRM_DTI u ON Aprovacao.IdAprovador = u.CRM_R_E_C_N_O "
    SQL = SQL & "LEFT JOIN CRM_DTI u2 ON Aprovacao.IdUsuarioInclusao = u2.CRM_R_E_C_N_O "
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "
    SQL = SQL & "ORDER BY DataCriacaoRegistro, Aprovacao.Nivel"

    On Error Resume Next
    SET rsAprovacoes = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAprovacoes) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    aprovacoesJSON = "["

    If Not rsAprovacoes.EOF Then
        Do While Not rsAprovacoes.EOF
            If Not primeiroItem Then aprovacoesJSON = aprovacoesJSON & "," End If
            
            aprovacoesJSON = aprovacoesJSON & "{"
            aprovacoesJSON = aprovacoesJSON & """idAprovador"": """ & EscapeJSON(rsAprovacoes("IdAprovador")) & """, "
            aprovacoesJSON = aprovacoesJSON & """nivel"": """ & EscapeJSON(rsAprovacoes("Nivel")) & """, "
            aprovacoesJSON = aprovacoesJSON & """invalido"": """ & EscapeJSON(rsAprovacoes("Invalido")) & """, "
            aprovacoesJSON = aprovacoesJSON & """nomeUsuarioAvaliacao"": """ & EscapeJSON(rsAprovacoes("NomeUsuarioAvaliacao")) & """, "
            aprovacoesJSON = aprovacoesJSON & """nomeUsuarioInclusao"": """ & EscapeJSON(rsAprovacoes("NomeUsuarioInclusao")) & """, "
            aprovacoesJSON = aprovacoesJSON & """avaliacao"": """ & EscapeJSON(rsAprovacoes("Avaliacao")) & """, "
            aprovacoesJSON = aprovacoesJSON & """dataAvaliacao"": """ & EscapeJSON(rsAprovacoes("DataAvaliacao")) & """, "
            aprovacoesJSON = aprovacoesJSON & """dataInclusao"": """ & EscapeJSON(rsAprovacoes("DataInclusao")) & """, "
            aprovacoesJSON = aprovacoesJSON & """motivo"": """ & EscapeJSON(rsAprovacoes("Motivo")) & """ "
            aprovacoesJSON = aprovacoesJSON & "}"
            
            primeiroItem = False
            rsAprovacoes.MoveNext
        Loop
    End If

    aprovacoesJSON = aprovacoesJSON & "]"

    rsAprovacoes.Close()
    SET rsAprovacoes = NOTHING

    ' Obtem lista de respostas
    SQL = "SELECT" 
    SQL = SQL & "    r.IdFNC, "
    SQL = SQL & "    isnull(Finalizada,0) Finalizada,"
    SQL = SQL & "    STRING_ESCAPE(CONVERT(VARCHAR(MAX), r.Descricao), 'json') AS Descricao, "
    SQL = SQL & "    STRING_ESCAPE(CONVERT(VARCHAR(MAX), r.Justificativa), 'json') AS Justificativa, "
    SQL = SQL & "    c1.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "    c2.CRM_NOME as NomeUsuarioFinalizacao, "
    SQL = SQL & "    r.DataInclusao, "
    SQL = SQL & "    r.DataFinalizacao, "
    SQL = SQL & "    Avaliacao "
    SQL = SQL & "FROM Resposta r "
    SQL = SQL & "LEFT JOIN FNC f on f.Id=r.IdFNC "
    SQL = SQL & "LEFT JOIN CRM_DTI c1 on c1.CRM_R_E_C_N_O = r.QuemIncluiu "
    SQL = SQL & "LEFT JOIN CRM_DTI c2 on c2.CRM_R_E_C_N_O = r.QuemFinalizou " 
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "

    On Error Resume Next
    SET rsRespostas = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRespostas) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    respostasJSON = "["

    If Not rsRespostas.EOF Then
        Do While Not rsRespostas.EOF
            If Not primeiroItem Then respostasJSON = respostasJSON & "," End If

            avaliacao = ""
            If rsRespostas("Avaliacao") = True Then 
                avaliacao = 1
            ElseIf rsRespostas("Avaliacao") = False Then 
                avaliacao = 0 
            End If
            
            respostasJSON = respostasJSON & "{"
            respostasJSON = respostasJSON & """descricao"": """ & EscapeJSON(rsRespostas("Descricao")) & """, "
            respostasJSON = respostasJSON & """justificativa"": """ & EscapeJSON(rsRespostas("Justificativa")) & """, "
            respostasJSON = respostasJSON & """finalizada"": """ & EscapeJSON(rsRespostas("Finalizada")) & """, "
            respostasJSON = respostasJSON & """avaliacao"": """ & avaliacao & """, "
            respostasJSON = respostasJSON & """nomeUsuarioInclusao"": """ & EscapeJSON(rsRespostas("NomeUsuarioInclusao")) & """, "
            respostasJSON = respostasJSON & """nomeUsuarioFinalizacao"": """ & EscapeJSON(rsRespostas("NomeUsuarioFinalizacao")) & """, "
            respostasJSON = respostasJSON & """dataInclusao"": """ & EscapeJSON(rsRespostas("DataInclusao")) & """, "
            respostasJSON = respostasJSON & """dataFinalizacao"": """ & EscapeJSON(rsRespostas("DataFinalizacao")) & """"
            respostasJSON = respostasJSON & "}"
            
            primeiroItem = False
            rsRespostas.MoveNext
        Loop
    End If

    respostasJSON = respostasJSON & "]"

    rsRespostas.Close()
    SET rsRespostas = NOTHING

    ' Obtém observações
    SQL = "SELECT o.Id, o.IdFnc, o.IdUsuarioInclusao, o.DataCriacao, u.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), o.Descricao), 'json') AS Descricao "
    SQL = SQL & "FROM ObservacaoFnc o "
    SQL = SQL & "LEFT JOIN CRM_DTI u ON o.IdUsuarioInclusao = u.CRM_R_E_C_N_O "
    SQL = SQL & "WHERE o.IdFnc = " & idFNC & " "
    SQL = SQL & "AND o.FlagInativo = 0"

    On Error Resume Next
    SET rsObservacoes = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObservacoes) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    observacoesJSON = "["

    If Not rsObservacoes.EOF Then
        Do While Not rsObservacoes.EOF
            If Not primeiroItem Then observacoesJSON = observacoesJSON & "," End If
            
            observacoesJSON = observacoesJSON & "{"
            observacoesJSON = observacoesJSON & """id"": """ & rsObservacoes("Id") & """, "
            observacoesJSON = observacoesJSON & """idFNC"": """ & rsObservacoes("IdFnc") & """, "
            observacoesJSON = observacoesJSON & """descricao"": """ & EscapeJSON(rsObservacoes("Descricao")) & """, "
            observacoesJSON = observacoesJSON & """idUsuarioInclusao"": """ & rsObservacoes("IdUsuarioInclusao") & """, "
            observacoesJSON = observacoesJSON & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObservacoes("NomeUsuarioInclusao")) & """, "
            observacoesJSON = observacoesJSON & """dataCriacao"": """ & rsObservacoes("DataCriacao") & """"
            observacoesJSON = observacoesJSON & "}"
            
            primeiroItem = False
            rsObservacoes.MoveNext()
        Loop
    End If

    observacoesJSON = observacoesJSON & "]"
    rsObservacoes.Close()
    SET rsObservacoes = NOTHING

    ' Obtém lista de causas ishikawa
    SQL = "SELECT Id, IdFNC, Categoria, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Descricao), 'json') AS Descricao, "
    SQL = SQL & "CausaRaiz, DataCriacao "
    SQL = SQL & "FROM FNCIshikawaCausas "
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "

    On Error Resume Next
    SET rsIshikawa = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsIshikawa) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar causas ishikawa: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    causasIshikawaJSON = "["

    If Not rsIshikawa.EOF Then
        Do While Not rsIshikawa.EOF
            If Not primeiroItem Then causasIshikawaJSON = causasIshikawaJSON & "," End If
            causaRaiz = Iif(rsIshikawa("CausaRaiz") = "True", 1, 0)
            
            causasIshikawaJSON = causasIshikawaJSON & "{"
            causasIshikawaJSON = causasIshikawaJSON & """id"": """ & rsIshikawa("Id") & """, "
            causasIshikawaJSON = causasIshikawaJSON & """idFNC"": """ & rsIshikawa("IdFNC") & """, "
            causasIshikawaJSON = causasIshikawaJSON & """categoria"": """ & rsIshikawa("Categoria") & """, "
            causasIshikawaJSON = causasIshikawaJSON & """causaRaiz"": """ & causaRaiz & """, "
            causasIshikawaJSON = causasIshikawaJSON & """descricao"": """ & EscapeJSON(rsIshikawa("Descricao")) & """, "
            causasIshikawaJSON = causasIshikawaJSON & """dataCriacao"": """ & rsIshikawa("DataCriacao") & """"
            causasIshikawaJSON = causasIshikawaJSON & "}"
            
            primeiroItem = False
            rsIshikawa.MoveNext()
        Loop
    End If

    causasIshikawaJSON = causasIshikawaJSON & "]"
    rsIshikawa.Close()
    SET rsIshikawa = NOTHING

    ' Obtém lista de ações 5w2h
    SQL = "SELECT Id, IdFNC, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Acao), 'json') AS Acao, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Oque), 'json') AS Oque, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Porque), 'json') AS Porque, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Onde), 'json') AS Onde, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Quando), 'json') AS Quando, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Quem), 'json') AS Quem, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Como), 'json') AS Como, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), Quanto), 'json') AS Quanto, "
    SQL = SQL & "DataCriacao "
    SQL = SQL & "FROM FNC5W2HAcoes "
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "

    On Error Resume Next
    SET rs5w2h = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rs5w2h) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar 5w2h: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    acoes5W2HJSON = "["

    If Not rs5w2h.EOF Then
        Do While Not rs5w2h.EOF
            If Not primeiroItem Then acoes5W2HJSON = acoes5W2HJSON & "," End If
            
            acoes5W2HJSON = acoes5W2HJSON & "{"
            acoes5W2HJSON = acoes5W2HJSON & """id"": """ & rs5w2h("Id") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """idFNC"": """ & rs5w2h("IdFNC") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """acao"": """ & rs5w2h("Acao") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """oque"": """ & rs5w2h("Oque") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """porque"": """ & rs5w2h("Porque") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """onde"": """ & rs5w2h("Onde") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """quando"": """ & rs5w2h("Quando") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """quanto"": """ & rs5w2h("Quanto") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """quem"": """ & rs5w2h("Quem") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """como"": """ & rs5w2h("Como") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """quanto"": """ & rs5w2h("Quanto") & """, "
            acoes5W2HJSON = acoes5W2HJSON & """dataCriacao"": """ & rs5w2h("DataCriacao") & """"
            acoes5W2HJSON = acoes5W2HJSON & "}"
            
            primeiroItem = False
            rs5w2h.MoveNext()
        Loop
    End If

    acoes5W2HJSON = acoes5W2HJSON & "]"
    rs5w2h.Close()
    SET rs5w2h = NOTHING
    
    jsonObjeto = "{"
    If Not rsObter.EOF Then
        valorReclamado = Replace(FormatNumber(rsObter("ValorReclamado"), 2, -1, 0, 0), ",", ".")
        coeficienteNQA = Iif(ISNULL(rsObter("CoeficienteNQA")), "", rsObter("CoeficienteNQA"))

        jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id")) & """, "
        jsonObjeto = jsonObjeto & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObter("NomeUsuarioInclusao")) & """, "
        jsonObjeto = jsonObjeto & """comSetorResponsavel"": """ & EscapeJSON(rsObter("ComSetorResponsavel")) & """, "
        jsonObjeto = jsonObjeto & """nivelAprovacaoAtual"": """ & EscapeJSON(rsObter("NivelAprovacaoAtual")) & """, "
        jsonObjeto = jsonObjeto & """codVendedor"": """ & EscapeJSON(rsObter("CodVendedor")) & """, "
        jsonObjeto = jsonObjeto & """nomeVendedor"": """ & EscapeJSON(rsObter("NomeVendedor")) & """, "
        jsonObjeto = jsonObjeto & """cidadeEstado"": """ & EscapeJSON(rsObter("CidadeEstado")) & """, "
        jsonObjeto = jsonObjeto & """cnpj"": """ & EscapeJSON(rsObter("CNPJ")) & """, "
        jsonObjeto = jsonObjeto & """turno"": """ & EscapeJSON(rsObter("Turno")) & """, "
        jsonObjeto = jsonObjeto & """maquina"": """ & EscapeJSON(rsObter("Maquina")) & """, "
        jsonObjeto = jsonObjeto & """eficacia"": """ & EscapeJSON(rsObter("Eficacia")) & """, "
        jsonObjeto = jsonObjeto & """responsavelId"": """ & EscapeJSON(rsObter("ResponsavelId")) & """, "
        jsonObjeto = jsonObjeto & """tipoDefeito"": """ & EscapeJSON(rsObter("tipoDefeito")) & """, "
        jsonObjeto = jsonObjeto & """tipoReclamacao"": """ & EscapeJSON(rsObter("tipoReclamacao")) & """, "
        jsonObjeto = jsonObjeto & """coeficienteNQA"": """ & EscapeJSON(Replace(coeficienteNQA, ",", ".")) & """, "
        jsonObjeto = jsonObjeto & """descricaoDisposicao"": """ & EscapeJSON(rsObter("DescricaoDisposicao")) & """, "
        jsonObjeto = jsonObjeto & """descricaoDefeito"": """ & EscapeJSON(rsObter("DescricaoDefeito")) & """, "
        jsonObjeto = jsonObjeto & """descricaoTipoDefeito"": """ & EscapeJSON(rsObter("DescricaoTipoDefeito")) & """, "
        jsonObjeto = jsonObjeto & """defeitoIshikawa"": """ & EscapeJSON(rsObter("DefeitoIshikawa")) & """, "
        jsonObjeto = jsonObjeto & """assunto5W2H"": """ & EscapeJSON(rsObter("Assunto5W2H")) & """, "
        jsonObjeto = jsonObjeto & """defeitoId"": """ & EscapeJSON(rsObter("DefeitoId")) & """, "
        jsonObjeto = jsonObjeto & """setorId"": """ & EscapeJSON(rsObter("SetorId")) & """, "
        jsonObjeto = jsonObjeto & """disposicaoId"": """ & EscapeJSON(rsObter("DisposicaoId")) & """, "
        jsonObjeto = jsonObjeto & """disposicao"": """ & EscapeJSON(rsObter("Disposicao")) & """, "
        jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodFilial")) & """, "
        jsonObjeto = jsonObjeto & """nomeFilial"": """ & EscapeJSON(rsObter("NomeFilial")) & """, "
        jsonObjeto = jsonObjeto & """nomeCliente"": """ & EscapeJSON(rsObter("NomeCliente")) & """, "
        jsonObjeto = jsonObjeto & """nomeResponsavel"": """ & EscapeJSON(rsObter("NomeResponsavel")) & """, "
        jsonObjeto = jsonObjeto & """codCliente"": """ & EscapeJSON(rsObter("CodCliente")) & """, "
        jsonObjeto = jsonObjeto & """codLoja"": """ & EscapeJSON(rsObter("CodLoja")) & """, "
        jsonObjeto = jsonObjeto & """situacao"": """ & EscapeJSON(rsObter("Situacao")) & """, "
        jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
        jsonObjeto = jsonObjeto & """dataEncerramento"": """ & EscapeJSON(rsObter("DataEncerramento")) & """, "
        jsonObjeto = jsonObjeto & """dataOcorrencia"": """ & EscapeJSON(rsObter("DataOcorrencia")) & """, "
        jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao")) & """, "
        jsonObjeto = jsonObjeto & """dataConclusaoPrevista"": """ & EscapeJSON(rsObter("DataConclusaoPrevista")) & """, "
        jsonObjeto = jsonObjeto & """dataProducao"": """ & EscapeJSON(rsObter("DataProducao")) & """, "
        jsonObjeto = jsonObjeto & """dataEnvioCliente"": """ & EscapeJSON(rsObter("DataEnvioCliente")) & """, "
        jsonObjeto = jsonObjeto & """aprovada"": """ & EscapeJSON(rsObter("Aprovada")) & """, "
        jsonObjeto = jsonObjeto & """valorReclamado"": """ & valorReclamado & """, "
        jsonObjeto = jsonObjeto & """quantidadeReclamada"": """ & EscapeJSON(rsObter("QuantidadeReclamada")) & """, "
        jsonObjeto = jsonObjeto & """contato"": """ & EscapeJSON(rsObter("Contato")) & """, "
        jsonObjeto = jsonObjeto & """telefone"": """ & EscapeJSON(rsObter("Telefone")) & """, "
        jsonObjeto = jsonObjeto & """email"": """ & EscapeJSON(rsObter("Email")) & """, "
        jsonObjeto = jsonObjeto & """notas"": " & notasJSON & ", "
        jsonObjeto = jsonObjeto & """anexos"": " & anexosJSON & ", "
        jsonObjeto = jsonObjeto & """aprovacoes"": " & aprovacoesJSON & ", "
        jsonObjeto = jsonObjeto & """respostas"": " & respostasJSON & ", "
        jsonObjeto = jsonObjeto & """causasIshikawa"": " & causasIshikawaJSON & ", "
        jsonObjeto = jsonObjeto & """acoes5W2H"": " & acoes5W2HJSON & ", "
        jsonObjeto = jsonObjeto & """observacoes"": " & observacoesJSON
    End If

    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()    
    tipoReclamacao = Request("tipoReclamacao")
    codPlanta = Request("codPlanta")
    nomeCliente = Request("nomeCliente")
    codCliente = Request("codCliente")
    codLoja = Request("codLoja")
    dataOcorrencia = Request("dataOcorrencia")
    tipoDefeito = Request("tipoDefeito")
    defeitoId = Request("defeitoId")
    disposicaoId = Request("disposicaoId")
    quantidadeReclamada = Request("quantidadeReclamada")
    valorReclamado = Request("valorReclamado")
    descricaoDisposicao = Request("descricaoDisposicao")
    descricaoDefeito = Request("descricaoDefeito")
    contato = Request("contato")
    email = Request("email")
    telefone = Request("telefone")
    notasJSON = Request("notas")
    anexosJSON = Request("anexos")
    descricaoPlanta = ObterDescricaoPlanta(codPlanta)
    eficacia = ""

    ehPlantaWestrock = VerificarPlantaWestrock(codPlanta)
    If ehPlantaWestrock Then
        codCliente = Right(String(10, "0") & codCliente, 10)
    End If

    ' Se usuário criando FNC for solicitante, precisamos ver se cliente pertence a carteira dele
    If usuarioLogado("nivel") > 0 Then
        temPermissao = ClientePertenceCarteira(codPlanta, codCliente, codLoja, descricaoPlanta)
    End If

    'If Not temPermissao Then
        'Criar = "{""sucesso"": false, ""mensagem"": ""Nao tem permissao para criar FNC para este cliente""}"
        'Exit Function
    'End If

    setorId = ObterSetor(defeitoId)
    responsavelId = ObterResponsavel(codPlanta)

    duplicado = ExisteDuplicidade(usuarioLogado("id"), quantidadeReclamada, valorReclamado, codCliente, disposicaoId, defeitoId, setorId, codPlanta)
    If duplicado Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Já existe uma FNC com os dados informados""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "INSERT INTO FNC ( "
    SQL = SQL & "   DataOcorrencia, DataInclusao, QuemIncluiu, Quantidade, IdDisposicao, ClienteDisposicao, "
    SQL = SQL & "   Descricao, Cliente, Loja, Contato, Telefone, Eficacia, Email, IdDefeito, IdResponsavel, "
    SQL = SQL & "   ValorReclamado, IdSetor, Planta, Status, TipoReclamacao "
    SQL = SQL & ") OUTPUT INSERTED.Id VALUES ("
    SQL = SQL & "'" & dataOcorrencia & "', "
    SQL = SQL & "getdate(), "
    SQL = SQL & "" & usuarioLogado("id") & ", "
    SQL = SQL & "" & quantidadeReclamada & ", "
    SQL = SQL & "" & disposicaoId & ", "
    SQL = SQL & "'" & descricaoDisposicao & "', "
    SQL = SQL & "'" & descricaoDefeito & "', "
    SQL = SQL & "'" & codCliente & "', "
    SQL = SQL & "'" & codLoja & "', "
    SQL = SQL & "'" & contato & "', "
    SQL = SQL & "'" & Left(telefone, 12) & "', "
    SQL = SQL & "'" & eficacia & "', "
    SQL = SQL & "'" & email & "', "
    SQL = SQL & "" & defeitoId & ", "
    SQL = SQL & "" & responsavelId & ", "
    SQL = SQL & "" & valorReclamado & ", "
    SQL = SQL & "" & setorId & ", "
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "1, "
    SQL = SQL & "'" & tipoReclamacao & "' "
    SQL = SQL & "); "

    ' Monta o SQL com as notas e pedidos
    notasJSON = Replace(notasJSON, "[", "")
    notasJSON = Replace(notasJSON, "]", "")
    notasArray = Split(notasJSON, "},{")
    
    listaPedidos = "" ' Auxiliar que será utilizada no envio de email
    valoresPedidoFnc = ""
    valoresNotasFiscal = ""
    For i = 0 To UBound(notasArray)
        nota = notasArray(i)
        
        idFNC = ExtrairValorJSON(nota, "idFNC")
        numeroPedido = ExtrairValorJSON(nota, "numeroPedido")
        itemPedido = ExtrairValorJSON(nota, "itemPedido")
        itemNotaFiscal = ExtrairValorJSON(nota, "itemNotaFiscal")
        precoUnitarioPedido = ExtrairValorJSON(nota, "precoUnitario")
        quantidadePedidoFNC = ExtrairValorJSON(nota, "quantidade")
        quantidadeReclamadaPedidoFNC = ExtrairValorJSON(nota, "quantidadeReclamada")
        codProduto = ExtrairValorJSON(nota, "codProduto")
        grupoProduto = ExtrairValorJSON(nota, "grupoProduto")
        referenciaProduto = ExtrairValorJSON(nota, "referencia")
        referenciaProduto = Replace(referenciaProduto, "'", "")
        notaFiscal = ExtrairValorJSON(nota, "notaFiscal")
        serieNotaFiscal = ExtrairValorJSON(nota, "serieNotaFiscal")
        dataEmissaoNota = ExtrairValorJSON(nota, "dataEmissaoNota")

        If i > 0 Then 
            valoresPedidoFnc = valoresPedidoFnc & ","
            valoresNotasFiscal = valoresNotasFiscal & ","
        End If

        valoresPedidoFnc = valoresPedidoFnc & "("
        valoresPedidoFnc = valoresPedidoFnc & "(SELECT max(Id) FROM FNC WHERE QuemIncluiu = " & usuarioLogado("id") & "), "
        valoresPedidoFnc = valoresPedidoFnc & "'" & numeroPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & itemPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & Replace(precoUnitarioPedido, ",", ".") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "'" & referenciaProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadePedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadeReclamadaPedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & usuarioLogado("id") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "GETDATE(), "
        valoresPedidoFnc = valoresPedidoFnc & "'" & codProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & grupoProduto & "' "
        valoresPedidoFnc = valoresPedidoFnc & ")"

        valoresNotasFiscal = valoresNotasFiscal & "("
        valoresNotasFiscal = valoresNotasFiscal & "(SELECT max(Id) FROM FNC WHERE QuemIncluiu = " & usuarioLogado("id") & "), "
        valoresNotasFiscal = valoresNotasFiscal & "'" & numeroPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & notaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & serieNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & quantidadePedidoFNC & ", "
        valoresNotasFiscal = valoresNotasFiscal & "'" & dataEmissaoNota & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & usuarioLogado("id") & ", "
        valoresNotasFiscal = valoresNotasFiscal & "GETDATE() "
        valoresNotasFiscal = valoresNotasFiscal & ")"

        listaPedidos = listaPedidos & numeroPedido & "; "
    Next

    ' Monta o SQL com os anexos
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    valoresAnexos = ""
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")

        If i > 0 Then 
            valoresAnexos = valoresAnexos & ","
        End If
        
        valoresAnexos = valoresAnexos & "("
        valoresAnexos = valoresAnexos & "(SELECT max(Id) FROM FNC WHERE QuemIncluiu = " & usuarioLogado("id") & "), "
        valoresAnexos = valoresAnexos & "'" & codPlanta & "', "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    If valoresPedidoFnc <> "" Then
        SQL = SQL & "INSERT INTO PedidoFnc ( "
        SQL = SQL & "   IdFNC, Pedido, Item, PrecoUnitario, Referencia, Quantidade, "
        SQL = SQL & "   QuantidadeReclamada, QuemIncluiu, DataInclusao, CodProduto, GrupoProduto "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresPedidoFnc & "; "
    End If

    If valoresNotasFiscal <> "" Then
        SQL = SQL & "INSERT INTO NotaFiscal ( "
        SQL = SQL & "   IdFNC, Pedido, Item, NotaFiscal, Serie, "
        SQL = SQL & "   ItemNotaFiscal, Quantidade, DataEmissao, QuemIncluiu, DataInclusao"
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresNotasFiscal & "; "
    End If

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO FNCDocumento ( "
        SQL = SQL & "   idFNC, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAnexos & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsCriar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    idNovaFNC = rsCriar("Id")
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0

    Call EnviarEmailCriacao(idNovaFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, listaPedidos)
    Call RegistrarLog("criacao", idNovaFNC, "")
    Criar = "{""sucesso"": true, ""mensagem"": ""FNC criada com sucesso!""}"
End Function

Function EditarSolicitante()   
    idFNC = Request("id")
    dataOcorrencia = Request("dataOcorrencia")
    tipoDefeito = Request("tipoDefeito")
    defeitoId = Iif(Request("defeitoId") = "", 0, Request("defeitoId"))
    disposicaoId = Iif(Request("disposicaoId") = "", 0, Request("disposicaoId"))
    quantidadeReclamada = Request("quantidadeReclamada")
    valorReclamado = Request("valorReclamado")
    descricaoDisposicao = Request("descricaoDisposicao")
    descricaoDefeito = Request("descricaoDefeito")
    contato = Request("contato")
    email = Request("email")
    telefone = Request("telefone")
    flagConverterFNC = Request("flagConverterFNC")

    notasJSON = Request("notas")
    anexosJSON = Request("anexos")
    
    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE FNC SET " 
    
    If flagConverterFNC = "1" Then
        'Se marcado para converter atualizar status de pre-cadastro
        SQL = SQL & "   TipoReclamacao = 'FNC', " 
    End If

    SQL = SQL & "   DataOcorrencia = '" & dataOcorrencia & "', "
    SQL = SQL & "   Quantidade = " & quantidadeReclamada & ", "
    SQL = SQL & "   IdDisposicao = " & disposicaoId & ", "
    SQL = SQL & "   ClienteDisposicao = '" & descricaoDisposicao & "', "
    SQL = SQL & "   Descricao = '" & descricaoDefeito & "', "
    SQL = SQL & "   Contato = '" & contato & "', "
    SQL = SQL & "   Telefone = '" & telefone & "', "
    SQL = SQL & "   Email = '" & email & "', "
    SQL = SQL & "   IdDefeito = " & defeitoId & ", "
    SQL = SQL & "   ValorReclamado = " & valorReclamado & " "
    SQL = SQL & "WHERE Id = " & idFNC & "; "

    ' Monta o SQL com as notas e pedidos
    notasJSON = Replace(notasJSON, "[", "")
    notasJSON = Replace(notasJSON, "]", "")
    notasArray = Split(notasJSON, "},{")
    
    listaPedidos = "" ' Auxiliar que será utilizada no envio de email
    valoresPedidoFnc = ""
    valoresNotasFiscal = ""
    For i = 0 To UBound(notasArray)
        nota = notasArray(i)
        
        numeroPedido = ExtrairValorJSON(nota, "numeroPedido")
        itemPedido = ExtrairValorJSON(nota, "itemPedido")
        itemNotaFiscal = ExtrairValorJSON(nota, "itemNotaFiscal")
        precoUnitarioPedido = ExtrairValorJSON(nota, "precoUnitario")
        quantidadePedidoFNC = ExtrairValorJSON(nota, "quantidade")
        quantidadeReclamadaPedidoFNC = ExtrairValorJSON(nota, "quantidadeReclamada")
        codProduto = ExtrairValorJSON(nota, "codProduto")
        grupoProduto = ExtrairValorJSON(nota, "grupoProduto")
        referenciaProduto = ExtrairValorJSON(nota, "referencia")
        referenciaProduto = Replace(referenciaProduto, "'", "")
        notaFiscal = ExtrairValorJSON(nota, "notaFiscal")
        serieNotaFiscal = ExtrairValorJSON(nota, "serieNotaFiscal")
        dataEmissaoNota = ExtrairValorJSON(nota, "dataEmissaoNota")

        If i > 0 Then 
            valoresPedidoFnc = valoresPedidoFnc & ","
            valoresNotasFiscal = valoresNotasFiscal & ","
        End If

        valoresPedidoFnc = valoresPedidoFnc & "("
        valoresPedidoFnc = valoresPedidoFnc & "'" & idFNC & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & numeroPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & itemPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & Replace(precoUnitarioPedido, ",", ".") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "'" & referenciaProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadePedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadeReclamadaPedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & usuarioLogado("id") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "GETDATE(), "
        valoresPedidoFnc = valoresPedidoFnc & "'" & codProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & grupoProduto & "' "
        valoresPedidoFnc = valoresPedidoFnc & ")"

        valoresNotasFiscal = valoresNotasFiscal & "("
        valoresNotasFiscal = valoresNotasFiscal & "'" & idFNC & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & numeroPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & notaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & serieNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & quantidadePedidoFNC & ", "
        valoresNotasFiscal = valoresNotasFiscal & "'" & dataEmissaoNota & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & usuarioLogado("id") & ", "
        valoresNotasFiscal = valoresNotasFiscal & "GETDATE() "
        valoresNotasFiscal = valoresNotasFiscal & ")"

        listaPedidos = listaPedidos & numeroPedido & "; "
    Next

    ' Monta o SQL com os anexos
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    anexosIds = ""
    valoresAnexos = ""
    primeiroItem = True
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        id = ExtrairValorJSON(anexo, "id")
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
         
        If anexosIds = "" Then 
            anexosIds = "'" & id & "'"
        Else
            anexosIds = anexosIds & ",'" & id & "'"
        End If

        If id = "" OR isNull(id) Then
            If Not primeiroItem Then 
                valoresAnexos = valoresAnexos & "," 
            End If

            valoresAnexos = valoresAnexos & "("
            valoresAnexos = valoresAnexos & "'" & idFNC & "', "
            valoresAnexos = valoresAnexos & "'" & codPlanta & "', "
            valoresAnexos = valoresAnexos & "'" & nome & "', "
            valoresAnexos = valoresAnexos & "'" & caminho & "', "
            valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
            valoresAnexos = valoresAnexos & "GETDATE()"
            valoresAnexos = valoresAnexos & ")"

            primeiroItem = False
        End If
    Next

    SQL = SQL & "DELETE FROM NotaFiscal WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM PedidoFnc WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM FNCDocumento WHERE idFNC = '" & idFNC & "' AND idFNCDocumento NOT IN (" & Iif(anexosIds = "", "0", anexosIds) & "); "

    If valoresPedidoFnc <> "" Then
        SQL = SQL & "INSERT INTO PedidoFnc ( "
        SQL = SQL & "   IdFNC, Pedido, Item, PrecoUnitario, Referencia, Quantidade, "
        SQL = SQL & "   QuantidadeReclamada, QuemIncluiu, DataInclusao, CodProduto, GrupoProduto "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresPedidoFnc & "; "
    End If
    
    If valoresNotasFiscal <> "" Then
        SQL = SQL & "INSERT INTO NotaFiscal ( "
        SQL = SQL & "   IdFNC, Pedido, Item, NotaFiscal, Serie, "
        SQL = SQL & "   ItemNotaFiscal, Quantidade, DataEmissao, QuemIncluiu, DataInclusao"
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresNotasFiscal & "; "
    End If

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO FNCDocumento ( "
        SQL = SQL & "   idFNC, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAnexos & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsEditar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        EditarSolicitante = "{""sucesso"": false, ""mensagem"": ""Erro ao editar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEditar.Close()
    SET rsEditar = NOTHING
    On Error Goto 0

    Call RegistrarLog("edicao", idFNC, "")
    EditarSolicitante = "{""sucesso"": true, ""mensagem"": ""FNC atualizada com sucesso!""}"
End Function

Function Editar()   
    idFNC = Request("id")
    codPlanta = Request("codPlanta")
    nomeCliente = Request("nomeCliente")
    codCliente = Request("codCliente")
    codLoja = Request("codLoja")
    dataOcorrencia = Request("dataOcorrencia")
    tipoDefeito = Request("tipoDefeito")
    defeitoId = Iif(Request("defeitoId") = "", 0, Request("defeitoId"))
    disposicaoId = Iif(Request("disposicaoId") = "", 0, Request("disposicaoId"))
    quantidadeReclamada = Request("quantidadeReclamada")
    valorReclamado = Request("valorReclamado")
    descricaoDisposicao = Request("descricaoDisposicao")
    descricaoDefeito = Request("descricaoDefeito")
    contato = Request("contato")
    email = Request("email")
    telefone = Request("telefone")
    descricaoPlanta = ObterDescricaoPlanta(codPlanta)

    notasJSON = Request("notas")
    anexosJSON = Request("anexos")
    observacoesJSON = Request("observacoes")

    defeitoIshikawa = Request("defeitoIshikawa")
    assunto5W2H = Request("assunto5W2H")
    causasIshikawaJSON = Request("causasIshikawa")
    acoes5W2HJSON = Request("acoes5W2H")

    dataEnvioCliente = Request("dataEnvioCliente")
    dataConclusaoPrevista = Request("dataConclusaoPrevista")
    dataProducao = Request("dataProducao")
    status = Request("status")
    situacao = Request("situacao")
    eficacia = Request("eficacia")
    maquina = Request("maquina")
    turno = Request("turno")

    setorId = Iif(Request("setorId") = "", 0, Request("setorId"))
    responsavelId = Iif(Request("responsavelId") = "", 0, Request("responsavelId"))
    
    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE FNC SET "   
    SQL = SQL & "   DataOcorrencia = '" & dataOcorrencia & "', "
    SQL = SQL & "   Quantidade = " & quantidadeReclamada & ", "
    SQL = SQL & "   IdDisposicao = " & disposicaoId & ", "
    SQL = SQL & "   ClienteDisposicao = '" & descricaoDisposicao & "', "
    SQL = SQL & "   Descricao = '" & descricaoDefeito & "', "
    SQL = SQL & "   Contato = '" & contato & "', "
    SQL = SQL & "   Telefone = '" & telefone & "', "
    SQL = SQL & "   Email = '" & email & "', "
    SQL = SQL & "   IdDefeito = " & defeitoId & ", "
    SQL = SQL & "   IdSetor = " & setorId & ", "
    SQL = SQL & "   IdAnalista = '" & responsavelId & "', "
    SQL = SQL & "   ValorReclamado = " & valorReclamado & ", "
    SQL = SQL & "   DataEnvioCliente = " & Iif(dataEnvioCliente = "", "NULL", "'" & dataEnvioCliente & "'") & ", "
    SQL = SQL & "   DataConclusaoPrevista = " & Iif(dataConclusaoPrevista = "", "NULL", "'" & dataConclusaoPrevista & "'") & ", "
    SQL = SQL & "   DataProducao = " & Iif(dataProducao = "", "NULL", "'" & dataProducao & "'") & ", "
    SQL = SQL & "   Status = " & status & ", "
    SQL = SQL & "   Situacao = " & situacao & ", "
    SQL = SQL & "   Eficacia = '" & eficacia & "', "
    SQL = SQL & "   Maquina = '" & maquina & "', "
    SQL = SQL & "   Turno = '" & turno & "', "
    SQL = SQL & "   DefeitoIshikawa = '" & defeitoIshikawa & "', "
    SQL = SQL & "   Assunto5W2H = '" & assunto5W2H & "' "
    SQL = SQL & Iif(situacao = 3, ", DataEncerramento = GETDATE(), QuemEncerrou = " & usuarioLogado("id") & " ", "")
    SQL = SQL & "WHERE Id = " & idFNC & "; "

    ' Monta o SQL com as notas e pedidos
    notasJSON = Replace(notasJSON, "[", "")
    notasJSON = Replace(notasJSON, "]", "")
    notasArray = Split(notasJSON, "},{")
    
    listaPedidos = "" ' Auxiliar que será utilizada no envio de email
    valoresPedidoFnc = ""
    valoresNotasFiscal = ""
    For i = 0 To UBound(notasArray)
        nota = notasArray(i)
        
        numeroPedido = ExtrairValorJSON(nota, "numeroPedido")
        itemPedido = ExtrairValorJSON(nota, "itemPedido")
        itemNotaFiscal = ExtrairValorJSON(nota, "itemNotaFiscal")
        precoUnitarioPedido = ExtrairValorJSON(nota, "precoUnitario")
        quantidadePedidoFNC = ExtrairValorJSON(nota, "quantidade")
        quantidadeReclamadaPedidoFNC = ExtrairValorJSON(nota, "quantidadeReclamada")
        codProduto = ExtrairValorJSON(nota, "codProduto")
        grupoProduto = ExtrairValorJSON(nota, "grupoProduto")
        referenciaProduto = ExtrairValorJSON(nota, "referencia")
        referenciaProduto = Replace(referenciaProduto, "'", "")
        notaFiscal = ExtrairValorJSON(nota, "notaFiscal")
        serieNotaFiscal = ExtrairValorJSON(nota, "serieNotaFiscal")
        dataEmissaoNota = ExtrairValorJSON(nota, "dataEmissaoNota")

        If i > 0 Then 
            valoresPedidoFnc = valoresPedidoFnc & ","
            valoresNotasFiscal = valoresNotasFiscal & ","
        End If

        valoresPedidoFnc = valoresPedidoFnc & "("
        valoresPedidoFnc = valoresPedidoFnc & "'" & idFNC & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & numeroPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & itemPedido & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & Replace(precoUnitarioPedido, ",", ".") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "'" & referenciaProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadePedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & quantidadeReclamadaPedidoFNC & ", "
        valoresPedidoFnc = valoresPedidoFnc & "" & usuarioLogado("id") & ", "
        valoresPedidoFnc = valoresPedidoFnc & "GETDATE(), "
        valoresPedidoFnc = valoresPedidoFnc & "'" & codProduto & "', "
        valoresPedidoFnc = valoresPedidoFnc & "'" & grupoProduto & "' "
        valoresPedidoFnc = valoresPedidoFnc & ")"

        valoresNotasFiscal = valoresNotasFiscal & "("
        valoresNotasFiscal = valoresNotasFiscal & "'" & idFNC & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & numeroPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemPedido & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & notaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & serieNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "'" & itemNotaFiscal & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & quantidadePedidoFNC & ", "
        valoresNotasFiscal = valoresNotasFiscal & "'" & dataEmissaoNota & "', "
        valoresNotasFiscal = valoresNotasFiscal & "" & usuarioLogado("id") & ", "
        valoresNotasFiscal = valoresNotasFiscal & "GETDATE() "
        valoresNotasFiscal = valoresNotasFiscal & ")"

        listaPedidos = listaPedidos & numeroPedido & "; "
    Next

    ' Monta o SQL com os anexos
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    anexosIds = ""
    valoresAnexos = ""
    primeiroItem = True
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        id = ExtrairValorJSON(anexo, "id")
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
         
        If anexosIds = "" Then 
            anexosIds = "'" & id & "'"
        Else
            anexosIds = anexosIds & ",'" & id & "'"
        End If

        If id = "" OR isNull(id) Then
            If Not primeiroItem Then 
                valoresAnexos = valoresAnexos & "," 
            End If

            valoresAnexos = valoresAnexos & "("
            valoresAnexos = valoresAnexos & "'" & idFNC & "', "
            valoresAnexos = valoresAnexos & "'" & codPlanta & "', "
            valoresAnexos = valoresAnexos & "'" & nome & "', "
            valoresAnexos = valoresAnexos & "'" & caminho & "', "
            valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
            valoresAnexos = valoresAnexos & "GETDATE()"
            valoresAnexos = valoresAnexos & ")"

            primeiroItem = False
        End If
    Next

    ' Monta o SQL com as observações
    observacoesJSON = Replace(observacoesJSON, "[", "")
    observacoesJSON = Replace(observacoesJSON, "]", "")
    observacoesArray = Split(observacoesJSON, "},{")
    
    observacoesIds = ""
    valoresObservacao = ""
    primeiroItem = True
    For i = 0 To UBound(observacoesArray)
        observacao = observacoesArray(i)
        id = ExtrairValorJSON(observacao, "id")
        descricao = ExtrairValorJSON(observacao, "descricao")
        
        If observacoesIds = "" Then 
            observacoesIds = "'" & id & "'"
        Else
            observacoesIds = observacoesIds & ",'" & id & "'"
        End If

        If id = "" OR isNull(id) Then
            If Not primeiroItem Then 
                valoresObservacao = valoresObservacao & "," 
            End If

            valoresObservacao = valoresObservacao & "("
            valoresObservacao = valoresObservacao & "'" & idFNC & "', "
            valoresObservacao = valoresObservacao & "'" & descricao & "', "
            valoresObservacao = valoresObservacao & "" & usuarioLogado("id") & ", "
            valoresObservacao = valoresObservacao & "GETDATE(), "
            valoresObservacao = valoresObservacao & "0 "
            valoresObservacao = valoresObservacao & ")"

            primeiroItem = False
        End If
    Next

    ' Monta o SQL com as observações
    causasIshikawaJSON = Replace(causasIshikawaJSON, "[", "")
    causasIshikawaJSON = Replace(causasIshikawaJSON, "]", "")
    causasIshikawaArray = Split(causasIshikawaJSON, "},{")
    
    valoresCausasIshikawa = ""
    primeiroItem = True
    For i = 0 To UBound(causasIshikawaArray)
        causa = causasIshikawaArray(i)
        categoria = ExtrairValorJSON(causa, "categoria")
        descricao = ExtrairValorJSON(causa, "descricao")
        causaRaiz = ExtrairValorJSON(causa, "causaRaiz")

        If Not primeiroItem Then 
            valoresCausasIshikawa = valoresCausasIshikawa & "," 
        End If

        valoresCausasIshikawa = valoresCausasIshikawa & "("
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & idFNC & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & Replace(categoria, "'", "") & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & Replace(descricao, "'", "") & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & Replace(causaRaiz, "'", "") & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "GETDATE() "
        valoresCausasIshikawa = valoresCausasIshikawa & ")"

        primeiroItem = False
    Next

    ' Monta o SQL com as ações
    acoes5W2HJSON = Replace(acoes5W2HJSON, "[", "")
    acoes5W2HJSON = Replace(acoes5W2HJSON, "]", "")
    acoes5W2HArray = Split(acoes5W2HJSON, "},{")
    
    valoresAcoes5W2H = ""
    primeiroItem = True
    For i = 0 To UBound(acoes5W2HArray)
        acao = acoes5W2HArray(i)
        oque = ExtrairValorJSON(acao, "oque")
        porque = ExtrairValorJSON(acao, "porque")
        onde = ExtrairValorJSON(acao, "onde")
        quando = ExtrairValorJSON(acao, "quando")
        quem = ExtrairValorJSON(acao, "quem")
        como = ExtrairValorJSON(acao, "como")
        quanto = ExtrairValorJSON(acao, "quanto")


        If Not primeiroItem Then 
            valoresAcoes5W2H = valoresAcoes5W2H & "," 
        End If

        valoresAcoes5W2H = valoresAcoes5W2H & "("
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & idFNC & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(oque, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(porque, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(onde, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(quando, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(quem, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(como, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & Replace(quanto, "'", "") & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "GETDATE() "
        valoresAcoes5W2H = valoresAcoes5W2H & ")"

        primeiroItem = False
    Next

    SQL = SQL & "DELETE FROM FNC5W2HAcoes WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM FNCIshikawaCausas WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM NotaFiscal WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM PedidoFnc WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM ObservacaoFnc WHERE IdFnc = '" & idFNC & "' AND Id NOT IN (" & Iif(observacoesIds = "", "0", observacoesIds) & "); "
    SQL = SQL & "DELETE FROM FNCDocumento WHERE idFNC = '" & idFNC & "' AND idFNCDocumento NOT IN (" & Iif(anexosIds = "", "0", anexosIds) & "); "

    If valoresPedidoFnc <> "" Then
        SQL = SQL & "INSERT INTO PedidoFnc ( "
        SQL = SQL & "   IdFNC, Pedido, Item, PrecoUnitario, Referencia, Quantidade, "
        SQL = SQL & "   QuantidadeReclamada, QuemIncluiu, DataInclusao, CodProduto, GrupoProduto "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresPedidoFnc & "; "
    End If
    
    If valoresNotasFiscal <> "" Then
        SQL = SQL & "INSERT INTO NotaFiscal ( "
        SQL = SQL & "   IdFNC, Pedido, Item, NotaFiscal, Serie, "
        SQL = SQL & "   ItemNotaFiscal, Quantidade, DataEmissao, QuemIncluiu, DataInclusao"
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresNotasFiscal & "; "
    End If

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO FNCDocumento ( "
        SQL = SQL & "   idFNC, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAnexos & "; "
    End If

    If valoresObservacao <> "" Then
        SQL = SQL & "INSERT INTO ObservacaoFnc ( "
        SQL = SQL & "   IdFnc, Descricao, IdUsuarioInclusao, DataCriacao, FlagInativo "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresObservacao & "; "
    End If

    If valoresCausasIshikawa <> "" Then
        SQL = SQL & "INSERT INTO FNCIshikawaCausas ( "
        SQL = SQL & "   IdFNC, Categoria, Descricao, CausaRaiz, DataCriacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresCausasIshikawa & "; "
    End If

    If valoresAcoes5W2H <> "" Then
        SQL = SQL & "INSERT INTO FNC5W2HAcoes ( "
        SQL = SQL & "   IdFNC, Oque, Porque, Onde, Quando, Quem, Como, Quanto, DataCriacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAcoes5W2H & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsEditar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEditar.Close()
    SET rsEditar = NOTHING
    On Error Goto 0

    Call RegistrarLog("edicao", idFNC, "")
    Editar = "{""sucesso"": true, ""mensagem"": ""FNC atualizada com sucesso!""}"
End Function

Function EnviarParaResponsavel()   
    idFNC = Request("idFNC")
    responsavelId = Request("responsavelId")

    If responsavelId = "" Then
        EnviarParaResponsavel = "{""sucesso"": false, ""mensagem"": ""Selecione um responsável""}"
        Exit Function
    End If
    
    'respostaPendente = VerificarRespostaPendente(idFNC)
    'If respostaPendente Then
        'EnviarParaResponsavel = "{""sucesso"": false, ""mensagem"": ""Já existe uma resposta pendente para essa FNC""}"
        'Exit Function
    'End If

    ' Obtém dados FNC 
    SQL = "SELECT "
    SQL = SQL & "f.Id, f.Cliente as CodCliente, cli.NOME_CLIENTE as NomeCliente, "
    SQL = SQL & "f.Loja as CodLoja, f.Planta as CodPlanta, cli.CIDADE as NomeFilial, "
    SQL = SQL & "f.Quantidade as QuantidadeReclamada, f.ValorReclamado, f.Situacao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDisposicao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Descricao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDefeito, "
    SQL = SQL & "CONVERT(varchar, f.DataInclusao,103) as DataInclusao "
    SQL = SQL & "FROM FNC f "
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
    SQL = SQL & "WHERE f.Id = '" & idFNC & "'"

    On Error Resume Next
    SET rsDados = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        EnviarParaResponsavel = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar para responsável : " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    nomeCliente = ""
    codCliente = ""
    codLoja = ""
    codPlanta = ""
    descricaoPlanta = ""
    quantidadeReclamada = ""
    valorReclamado = ""
    descricaoDefeito = ""
    descricaoDisposicao = ""

    If Not rsDados.EOF Then
        ' Se falta apenas a aprovação atual, marca como finalizado
        situacao = rsDados("Situacao")
        nomeCliente = rsDados("NomeCliente")
        codCliente = rsDados("CodCliente")
        codLoja = rsDados("CodLoja")
        codPlanta = rsDados("CodPlanta")
        quantidadeReclamada = rsDados("QuantidadeReclamada")
        valorReclamado = rsDados("ValorReclamado")
        descricaoDefeito = rsDados("DescricaoDefeito")
        descricaoDisposicao = rsDados("DescricaoDisposicao")
        descricaoPlanta = ObterDescricaoPlanta(codPlanta)    
    End If

    rsDados.Close()
    SET rsDados = NOTHING
    On Error Goto 0

    If situacao = "3" Then
        EnviarParaResponsavel = "{""sucesso"": false, ""mensagem"": ""Essa FNC já foi encerrada.""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE FNC SET "   
    SQL = SQL & "   IdAnalista = '" & responsavelId & "', "
    SQL = SQL & "   Situacao = 1, "
    SQL = SQL & "   DataEnvioSetorResponsavel = GETDATE(), "
    SQL = SQL & "   ComSetorResponsavel = 1"
    SQL = SQL & "WHERE Id = " & idFNC & "; "

    SQL = SQL & "INSERT INTO Resposta (IdFNC, Descricao, Justificativa, Finalizada, QuemIncluiu, Excluido, DataInclusao, QuemFinalizou, DataFinalizacao) VALUES ("
    SQL = SQL & idFNC & ", "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "0,"
    SQL = SQL & usuarioLogado("id") & ","
    SQL = SQL & "0, "
    SQL = SQL & "GETDATE(),"
    SQL = SQL & "'" & responsavelId & "', "
    SQL = SQL & "NULL"
    SQL = SQL & "); " 
    
    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsEnviar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEnviar) Then
        EnviarParaResponsavel = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar para responsável: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEnviar.Close()
    SET rsEnviar = NOTHING
    On Error Goto 0
    
    Call EnviarEmailResponsavel(responsavelId, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, dataInclusao)
    Call RegistrarLog("envio_responsavel", idFNC, "")
    EnviarParaResponsavel = "{""sucesso"": true, ""mensagem"": ""FNC enviada com sucesso!""}"
End Function

Function EnviarParaAprovacao()   
    idFNC = Request("idFNC")
 
    ' Obtém dados FNC 
    SQL = "SELECT "
    SQL = SQL & "f.Id, f.Cliente as CodCliente, cli.NOME_CLIENTE as NomeCliente, f.IdSetor, "
    SQL = SQL & "f.Loja as CodLoja, f.Planta as CodPlanta, cli.CIDADE as NomeFilial, "
    SQL = SQL & "f.Quantidade as QuantidadeReclamada, f.ValorReclamado, f.Situacao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDisposicao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Descricao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDefeito, "
    SQL = SQL & "CONVERT(varchar, f.DataInclusao,103) as DataInclusao, NivelAprovacaoAtual "
    SQL = SQL & "FROM FNC f "
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
    SQL = SQL & "WHERE f.Id = '" & idFNC & "'"

    On Error Resume Next
    SET rsDados = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        EnviarParaAprovacao = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar para aprovacao: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsDados.EOF Then
        ' Se falta apenas a aprovação atual, marca como finalizado
        situacao = rsDados("Situacao")
        idSetor = rsDados("IdSetor")
        nomeCliente = rsDados("NomeCliente")
        codCliente = rsDados("CodCliente")
        codLoja = rsDados("CodLoja")
        codPlanta = rsDados("CodPlanta")
        quantidadeReclamada = rsDados("QuantidadeReclamada")
        valorReclamado = rsDados("ValorReclamado")
        descricaoDefeito = rsDados("DescricaoDefeito")
        descricaoDisposicao = rsDados("DescricaoDisposicao")
        descricaoPlanta = ObterDescricaoPlanta(codPlanta)    
        nivelAprovacaoAtual = rsDados("NivelAprovacaoAtual")
    End If

    rsDados.Close()
    SET rsDados = NOTHING
    On Error Goto 0

    'If situacao = "3" Then
        'EnviarParaAprovacao = "{""sucesso"": false, ""mensagem"": ""Essa FNC já foi encerrada.""}"
        'Exit Function
    'End If

    quantidadeAprovacoes = ObterQuantidadeAprovacoes(codPlanta, valorReclamado)

    SQL = "BEGIN TRANSACTION; "
    
    SQL = SQL & "UPDATE FNC SET "
    SQL = SQL & "   NivelAprovacaoAtual = '1', "
    SQL = SQL & "   Aprovada = NULL, "
    SQL = SQL & "   DataAvaliacao = NULL "   
    SQL = SQL & "WHERE Id = " & idFNC & "; "

    SQL = SQL & "UPDATE AprovacaoFNC SET Invalido = 1 WHERE IdFNC = " & idFNC & "; "

    SQL = SQL & "INSERT INTO AprovacaoFNC ("
    SQL = SQL & "    IdFNC, "
    SQL = SQL & "    IdAprovador, "
    SQL = SQL & "    IdUsuarioInclusao, "
    SQL = SQL & "    Avaliacao, "
    SQL = SQL & "    DataAvaliacao, "
    SQL = SQL & "    Motivo, "
    SQL = SQL & "    Nivel, "
    SQL = SQL & "    DataCriacaoRegistro, "
    SQL = SQL & "    Invalido "
    SQL = SQL & ") VALUES "

    For i = 1 To quantidadeAprovacoes        
        If i > 1 Then  SQL = SQL & "," End If
        
        SQL = SQL & "("
        SQL = SQL & "'" & idFNC & "', "
        SQL = SQL & "0, "
        SQL = SQL & "'" & usuarioLogado("id") & "', "
        SQL = SQL & "0, "
        SQL = SQL & "NULL, "
        SQL = SQL & "'', "
        SQL = SQL & "'" & i & "', "
        SQL = SQL & "GETDATE(), "
        SQL = SQL & "0"
        SQL = SQL & ")"
    Next
    
    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsEnviar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEnviar) Then
        EnviarParaAprovacao = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar para responsável: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEnviar.Close()
    SET rsEnviar = NOTHING
    On Error Goto 0

    Call EnviarEmailAprovadores(1, idSetor, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao)
    Call RegistrarLog("envio_aprovacao", idFNC, "")
    EnviarParaAprovacao = "{""sucesso"": true, ""mensagem"": ""FNC enviada com sucesso!""}"
End Function

Function Avaliar()   
    idFNC = Request("id")
    flagAprovar = Request("flagAprovar")
    observacao = Request("observacao")
    anexosJSON = Request("anexos")

    ' Obtém dados FNC 
    SQL = "SELECT "
    SQL = SQL & "("
	SQL = SQL & "   SELECT COUNT(IdFNC)as Total "
	SQL = SQL & "   FROM AprovacaoFNC WHERE IdFNC = f.Id "
    SQL = SQL & "   AND DataAvaliacao IS NULL "
    SQL = SQL & "   AND (Invalido IS NULL OR Invalido = 0) "
	SQL = SQL & ") as TotalAprovacoesPendentes, "
    SQL = SQL & "f.Id, f.IdSetor, f.Cliente as CodCliente, cli.NOME_CLIENTE as NomeCliente, "
    SQL = SQL & "f.Loja as CodLoja, f.Planta as CodPlanta, cli.CIDADE as NomeFilial, f.QuemIncluiu as IdUsuarioInclusao, "
    SQL = SQL & "f.Quantidade as QuantidadeReclamada, f.ValorReclamado, f.Situacao, f.NivelAprovacaoAtual, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDisposicao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Descricao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDefeito, "
    SQL = SQL & "CONVERT(varchar, f.DataInclusao,103) as DataInclusao "
    SQL = SQL & "FROM FNC f "
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
    SQL = SQL & "WHERE f.Id = '" & idFNC & "'"

    On Error Resume Next
    SET rsDados = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        Avaliar = "{""sucesso"": false, ""mensagem"": ""Erro ao avaliar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    Set dadosFNC = Server.CreateObject("Scripting.Dictionary")

    If Not rsDados.EOF Then
        ' Se falta apenas a aprovação atual (<=1), marca como finalizado
        dadosFNC("ehUltimaAprovacao") = Iif(rsDados("TotalAprovacoesPendentes") < 2, 1, 0)
        dadosFNC("idUsuarioInclusao") = rsDados("IdUsuarioInclusao")
        dadosFNC("idSetor") = rsDados("IdSetor")
        dadosFNC("situacao") = rsDados("Situacao")
        dadosFNC("nomeCliente") = rsDados("NomeCliente")
        dadosFNC("codCliente") = rsDados("CodCliente")
        dadosFNC("codLoja") = rsDados("CodLoja")
        dadosFNC("codPlanta") = rsDados("CodPlanta")
        dadosFNC("quantidadeReclamada") = rsDados("QuantidadeReclamada")
        dadosFNC("valorReclamado") = rsDados("ValorReclamado")
        dadosFNC("descricaoDefeito") = rsDados("DescricaoDefeito")
        dadosFNC("descricaoDisposicao") = rsDados("DescricaoDisposicao")
        dadosFNC("dataInclusao") = rsDados("DataInclusao")
        dadosFNC("descricaoPlanta") = ObterDescricaoPlanta(dadosFNC("codPlanta"))
        dadosFNC("nivelAprovacaoAtual") = rsDados("NivelAprovacaoAtual")
    End If

    rsDados.Close()
    SET rsDados = NOTHING
    On Error Goto 0


    If dadosFNC("situacao") <> "2" Then
        Avaliar = "{""sucesso"": false, ""mensagem"": ""Essa FNC não está aguardando aprovação""}"
        Exit Function
    End If

    ' Verificamos se usuário pode validar na etapa atual
    If not UsuarioPodeValidarEtapa(dadosFNC, usuarioLogado("id")) Then
        Avaliar = "{""sucesso"": false, ""mensagem"": ""Você não tem permisão para aprovar essa FNC nesta etapa""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    ' Atualiza registro de aprovação
    SQL = SQL & "UPDATE AprovacaoFNC SET "
    SQL = SQL & "   IdAprovador = '" & usuarioLogado("id") & "', "
    SQL = SQL & "   Avaliacao = " & flagAprovar & ", "
    SQL = SQL & "   DataAvaliacao = GETDATE(), "
    SQL = SQL & "   Motivo = '" & observacao & "' "
    SQL = SQL & "WHERE IdFNC = '" & idFNC & "' "
    SQL = SQL & "AND Nivel = '" & dadosFNC("nivelAprovacaoAtual") & "' "
    SQL = SQL & "AND (Invalido = 0 OR Invalido IS NULL); "

    If flagAprovar = 0 Then
        ' Se foi rejeitado precisamos marcar a FNC como rejeitada, resetar nível de aprovacao e invalidar registros de aprovação atuais
        SQL = SQL & "UPDATE FNC SET Aprovada = 0, NivelAprovacaoAtual = NULL, DataAvaliacao = NULL WHERE Id = '" & idFNC & "'; "
        SQL = SQL & "UPDATE AprovacaoFNC SET Invalido = 1 WHERE IdFNC = '" & idFNC & "'; "
    Else 
        SQL = SQL & "UPDATE FNC SET  "
        
        If dadosFNC("ehUltimaAprovacao") = 1 Then
            ' Se foi aprovado e finalizou todas aprovacoes marca como aprovado
            SQL = SQL & "Aprovada = 1, DataAvaliacao = GETDATE(), NivelAprovacaoAtual = NULL "
        Else
            ' Se não apenas vai para próximo nivel de aprovação
            SQL = SQL & "NivelAprovacaoAtual = NivelAprovacaoAtual + 1 "
        End If 

        SQL = SQL & "Where Id = " & idFNC & "; "
    End If

    ' Anexos (se houver)
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    valoresAnexos = ""
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")

        If i > 0 Then 
            valoresAnexos = valoresAnexos & ","
        End If
        
        valoresAnexos = valoresAnexos & "("
        valoresAnexos = valoresAnexos & "'" & idFNC & "', "
        valoresAnexos = valoresAnexos & "'" & codPlanta & "', "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO FNCDocumento ( "
        SQL = SQL & "   idFNC, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAnexos & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
      
    On Error Resume Next
    SET rsAvaliar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAvaliar) Then
        Avaliar = "{""sucesso"": false, ""mensagem"": ""Erro ao avaliar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsAvaliar.Close()
    SET rsAvaliar = NOTHING
    On Error Goto 0

    ' Envia email de avaliação para analista, independente de aprovar ou não
    Call EnviarEmailAvaliacao(flagAprovar, idFNC, dadosFNC("nomeCliente"), dadosFNC("codCliente"), dadosFNC("codLoja"), dadosFNC("codPlanta"), dadosFNC("descricaoPlanta"), dadosFNC("quantidadeReclamada"), dadosFNC("valorReclamado"), dadosFNC("descricaoDefeito"), dadosFNC("descricaoDisposicao"), dataInclusao)
    
    If flagAprovar = 1 AND dadosFNC("ehUltimaAprovacao") = 0 Then
        ' Se foi aprovado e não finalizou todas aprovacoes, envia email para o próximo aprovador
        Call EnviarEmailAprovadores(dadosFNC("nivelAprovacaoAtual") + 1, dadosFNC("idSetor"), idFNC, dadosFNC("nomeCliente"), dadosFNC("codCliente"), dadosFNC("codLoja"), dadosFNC("codPlanta"), dadosFNC("descricaoPlanta"), dadosFNC("quantidadeReclamada"), dadosFNC("valorReclamado"), dadosFNC("descricaoDefeito"), dadosFNC("descricaoDisposicao"))
    ElseIf flagAprovar = 1 AND dadosFNC("ehUltimaAprovacao") = 1 Then
        Call EnviarEmailAprovadoTotal(flagAprovar, idFNC, dadosFNC("nomeCliente"), dadosFNC("codCliente"), dadosFNC("codLoja"), dadosFNC("codPlanta"), dadosFNC("descricaoPlanta"), dadosFNC("quantidadeReclamada"), dadosFNC("valorReclamado"), dadosFNC("descricaoDefeito"), dadosFNC("descricaoDisposicao"), dataInclusao, dadosFNC("idUsuarioInclusao"))
    End If
    
    Call RegistrarLog(Iif(flagAprovar = 1, "aprovacao", "rejeicao"), idFNC, "")
    Avaliar = "{""sucesso"": true, ""mensagem"": ""FNC avaliada com sucesso!""}"
End Function

Function Responder()   
    idFNC = Request("id")
    flagAprovar = Request("flagAprovar")
    resposta = Request("resposta")
    justificativa = Request("justificativa")
    anexosJSON = Request("anexos")
    causasIshikawaJSON = Request("causasIshikawa")
    acoes5W2HJSON = Request("acoes5W2H")
    defeitoIshikawa = Request("defeitoIshikawa")
    assunto5W2H = Request("assunto5W2H")

    ' Verifica se usuário é o responsável
    temPermissao = ResponsavelExiste(idFNC, usuarioLogado("id"))
    If Not temPermissao Then
        Responder = "{""sucesso"": false, ""mensagem"": ""Você não tem permissão para responder essa FNC""}"
        Exit Function
    End If

    ' Obtém dados FNC 
    SQL = "SELECT "
    SQL = SQL & "f.Id, f.Cliente as CodCliente, cli.NOME_CLIENTE as NomeCliente, f.QuemIncluiu as IdUsuarioInclusao, "
    SQL = SQL & "f.Loja as CodLoja, f.Planta as CodPlanta, cli.CIDADE as NomeFilial, "
    SQL = SQL & "f.Quantidade as QuantidadeReclamada, f.ValorReclamado, f.Situacao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.ClienteDisposicao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDisposicao, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Descricao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS DescricaoDefeito, "
    SQL = SQL & "CONVERT(varchar, f.DataInclusao,103) as DataInclusao "
    SQL = SQL & "FROM FNC f "
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
    SQL = SQL & "WHERE f.Id = '" & idFNC & "'"

    On Error Resume Next
    SET rsDados = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        Responder = "{""sucesso"": false, ""mensagem"": ""Erro ao Responder FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsDados.EOF Then
        ' Se falta apenas a aprovação atual, marca como finalizado
        idUsuarioInclusao = rsDados("IdUsuarioInclusao")
        situacao = rsDados("Situacao")
        nomeCliente = rsDados("NomeCliente")
        codCliente = rsDados("CodCliente")
        codLoja = rsDados("CodLoja")
        codPlanta = rsDados("CodPlanta")
        quantidadeReclamada = rsDados("QuantidadeReclamada")
        valorReclamado = rsDados("ValorReclamado")
        descricaoDefeito = rsDados("DescricaoDefeito")
        descricaoDisposicao = rsDados("DescricaoDisposicao")
        descricaoPlanta = ObterDescricaoPlanta(codPlanta)    
    End If

    rsDados.Close()
    SET rsDados = NOTHING
    On Error Goto 0

    If situacao = "3" Then
        Responder = "{""sucesso"": false, ""mensagem"": ""Essa FNC já foi respondida ou está encerrada.""}"
        Exit Function
    End If

    ' Atualiza registro de aprovação
    SQL = "BEGIN TRANSACTION; "

    SQL = SQL & "UPDATE r SET "
    SQL = SQL & "   r.Descricao = '" & resposta & "', "
    SQL = SQL & "   r.Justificativa = '" & justificativa & "',"
    SQL = SQL & "   r.Avaliacao = '" & flagAprovar & "', "
    SQL = SQL & "   r.Finalizada = 1, "
    SQL = SQL & "   r.DataFinalizacao = GETDATE() "   
    SQL = SQL & "FROM Resposta r "
    SQL = SQL & "WHERE r.IdFNC = " & idFNC & " "
    SQL = SQL & "AND r.Id = (                                   "
    SQL = SQL & "    SELECT TOP 1 r2.Id "
    SQL = SQL & "    FROM Resposta r2 "
    SQL = SQL & "    WHERE r2.IdFNC = r.IdFNC "
    SQL = SQL & "        AND r2.QuemFinalizou = r.QuemFinalizou "
    SQL = SQL & "        AND r2.Finalizada = 0 "
    SQL = SQL & "    ORDER BY r2.Id DESC "
    SQL = SQL & ") "
    SQL = SQL & "AND r.QuemFinalizou = " & usuarioLogado("id") & " "
    SQL = SQL & "AND r.Finalizada = 0; "

    SQL = SQL & "UPDATE FNC SET "
    SQL = SQL & "   ComSetorResponsavel = 0, "
    SQL = SQL & "   Situacao = 2, "
    SQL = SQL & "   DefeitoIshikawa = '" & defeitoIshikawa & "', "
    SQL = SQL & "   Assunto5W2H = '" & assunto5W2H & "', "
    SQL = SQL & "   Status = '" & Iif(flagAprovar = 1, 3, 4) & "'"
    SQL = SQL & "Where Id = " & idFNC 

    ' Anexos (se houver)
    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    valoresAnexos = ""
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")

        If i > 0 Then 
            valoresAnexos = valoresAnexos & ","
        End If
        
        valoresAnexos = valoresAnexos & "("
        valoresAnexos = valoresAnexos & "'" & idFNC & "', "
        valoresAnexos = valoresAnexos & "'" & codPlanta & "', "
        valoresAnexos = valoresAnexos & "'" & nome & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "'" & usuarioLogado("login") & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    ' Monta o SQL com as observações
    causasIshikawaJSON = Replace(causasIshikawaJSON, "[", "")
    causasIshikawaJSON = Replace(causasIshikawaJSON, "]", "")
    causasIshikawaArray = Split(causasIshikawaJSON, "},{")
    
    valoresCausasIshikawa = ""
    primeiroItem = True
    For i = 0 To UBound(causasIshikawaArray)
        causa = causasIshikawaArray(i)
        categoria = ExtrairValorJSON(causa, "categoria")
        descricao = ExtrairValorJSON(causa, "descricao")
        causaRaiz = ExtrairValorJSON(causa, "causaRaiz")

        If Not primeiroItem Then 
            valoresCausasIshikawa = valoresCausasIshikawa & "," 
        End If

        valoresCausasIshikawa = valoresCausasIshikawa & "("
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & idFNC & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & categoria & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & descricao & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "'" & causaRaiz & "', "
        valoresCausasIshikawa = valoresCausasIshikawa & "GETDATE() "
        valoresCausasIshikawa = valoresCausasIshikawa & ")"

        primeiroItem = False
    Next

    ' Monta o SQL com as ações
    acoes5W2HJSON = Replace(acoes5W2HJSON, "[", "")
    acoes5W2HJSON = Replace(acoes5W2HJSON, "]", "")
    acoes5W2HArray = Split(acoes5W2HJSON, "},{")
    
    valoresAcoes5W2H = ""
    primeiroItem = True
    For i = 0 To UBound(acoes5W2HArray)
        acao = acoes5W2HArray(i)
        oque = ExtrairValorJSON(acao, "oque")
        porque = ExtrairValorJSON(acao, "porque")
        onde = ExtrairValorJSON(acao, "onde")
        quando = ExtrairValorJSON(acao, "quando")
        quem = ExtrairValorJSON(acao, "quem")
        como = ExtrairValorJSON(acao, "como")
        quanto = ExtrairValorJSON(acao, "quanto")

        If Not primeiroItem Then 
            valoresAcoes5W2H = valoresAcoes5W2H & "," 
        End If

        valoresAcoes5W2H = valoresAcoes5W2H & "("
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & idFNC & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & oque & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & porque & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & onde & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & quando & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & quem & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & como & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "'" & quanto & "', "
        valoresAcoes5W2H = valoresAcoes5W2H & "GETDATE() "
        valoresAcoes5W2H = valoresAcoes5W2H & ")"

        primeiroItem = False
    Next

    SQL = SQL & "DELETE FROM FNC5W2HAcoes WHERE IdFNC = '" & idFNC & "'; "
    SQL = SQL & "DELETE FROM FNCIshikawaCausas WHERE IdFNC = '" & idFNC & "'; "

    If valoresAnexos <> "" Then
        SQL = SQL & "INSERT INTO FNCDocumento ( "
        SQL = SQL & "   idFNC, Filial, Nome, Caminho, Usuario, DataDaOperacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAnexos & "; "
    End If

    If valoresCausasIshikawa <> "" Then
        SQL = SQL & "INSERT INTO FNCIshikawaCausas ( "
        SQL = SQL & "   IdFNC, Categoria, Descricao, CausaRaiz, DataCriacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresCausasIshikawa & "; "
    End If

    If valoresAcoes5W2H <> "" Then
        SQL = SQL & "INSERT INTO FNC5W2HAcoes ( "
        SQL = SQL & "   IdFNC, Oque, Porque, Onde, Quando, Quem, Como, Quanto, DataCriacao "
        SQL = SQL & ") VALUES " 
        SQL = SQL & valoresAcoes5W2H & "; "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
      
    On Error Resume Next
    SET rsResponder = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsResponder) Then
        Responder = "{""sucesso"": false, ""mensagem"": ""Erro ao Responder FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsResponder.Close()
    SET rsResponder = NOTHING
    On Error Goto 0

    Call EnviarEmailResposta(flagAprovar, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, idUsuarioInclusao)
    Call RegistrarLog(Iif(flagAprovar = 1, "resposta", "devolucao"), idFNC, "")
    Responder = "{""sucesso"": true, ""mensagem"": ""FNC respondida com sucesso!""}"
End Function

Function Cancelar()   
    idFNC = Request("idFNC")

    ' Atualiza registro de aprovação
    SQL = SQL & "UPDATE FNC SET "
    SQL = SQL & "   Status = '5' "
    SQL = SQL & "WHERE Id = '" & idFNC & "' "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
      
    On Error Resume Next
    SET rsCancelar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCancelar) Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Erro ao avaliar FNC: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsCancelar.Close()
    SET rsCancelar = NOTHING
    On Error Goto 0

    Call RegistrarLog("cancelamento", idFNC, "")
    Cancelar = "{""sucesso"": true, ""mensagem"": ""FNC cancelada com sucesso!""}"
End Function

Function ListarSituacoes()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "ABERTA" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "RESPONDIDA" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "ENCERRADA" & """"
    jsonArray = jsonArray & "}"


    jsonArray = jsonArray & "]"
    
    ListarSituacoes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposDefeito()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "ADM" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Administrativo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "PRO" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Industrial" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposDefeito = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function
                             
Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "REGISTRADA" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "EM ANALISE" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "PROCEDE" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "NÃO PROCEDE" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "5" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CANCELADA" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarResponsaveis()
    codplanta = Request("codPlanta")

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
    SQL = "SELECT Id, Descricao FROM Setor WHERE ISNULL(FlagInativo, 0) = 0"

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

Function ListarDefeitos()
    Dim SQL, rsDefeitos, jsonArray, primeiroItem

    tipo = request("tipoDefeito")

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

Function ListarAprovadores()
    codPlanta = Request("codPlanta")
    tipoDefeito = Request("tipoDefeito")

    SQL = SQL & "SELECT Id, Nivel, Planta, Nome, "  
    SQL = SQL & "CASE " 
    SQL = SQL & "   WHEN Nivel = 1 THEN 'Gerente Setor' "
    SQL = SQL & "   WHEN Nivel = 2 THEN 'Gerente Qualidade'  "
    SQL = SQL & "   WHEN Nivel = 3 THEN 'Controller' " 
    SQL = SQL & "   WHEN Nivel = 4 THEN 'Gerente Geral' "  
    SQL = SQL & "END as DescricaoNivel "
    SQL = SQL & "FROM AprovadorFNC "
    SQL = SQL & "WHERE Planta = '" & codPlanta & "' "
    SQL = SQL & "AND ( "
    SQL = SQL & "   AprovadorTipo = '" & tipoDefeito & "' "
    SQL = SQL & "   OR AprovadorTipo = 'AMB' "
    SQL = SQL & ") "
    SQL = SQL & "AND isnull(Excluido,'') = '' "
    SQL = SQL & "ORDER BY Nivel" 

    On Error Resume Next
    SET rsAprovadores = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAprovadores) Then
        ListarAprovadores = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de aprovadores: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsAprovadores.EOF Then
        Do While Not rsAprovadores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsAprovadores("Id")) & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsAprovadores("Nome")) & """, "
            jsonArray = jsonArray & """nivel"": """ & EscapeJSON(rsAprovadores("Nivel")) & """, "
            jsonArray = jsonArray & """planta"": """ & EscapeJSON(rsAprovadores("Planta")) & """, "
            jsonArray = jsonArray & """descricaoNivel"": """ & EscapeJSON(rsAprovadores("DescricaoNivel")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsAprovadores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsAprovadores.Close()
    SET rsAprovadores = NOTHING
    
    ListarAprovadores = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarNotasCliente()
    codCliente = Request("codCliente")
    codLoja = Request("codLoja")
    codPlanta = Request("codPlanta")
    pesquisa = Request("pesquisa")

    descricaoPlanta = ObterDescricaoPlanta(codPlanta)
    ehPlantaWestrock = VerificarPlantaWestrock(codPlanta)

    If ehPlantaWestrock Then
        SQL = "SELECT NUMERO_PEDIDO AS PEDIDO, ITEM_PEDIDO AS ITEM, PRODUTO, DESCRICAO_PRODUTO, "
        SQL = SQL & " CODIGO_CLIENTE_SAP AS COD_CLI, NOME_CLIENTE AS NOME_CLI, GRUPO_PRODUTO, "
        SQL = SQL & " QUANTIDADE AS QTDE, VALOR_UNITARIO_NET AS PRECO, NUMERO_NF, SERIE_NF, EMISSAO "
        SQL = SQL & " FROM VW_FNC_NOTAS_PEDIDOS_WRK " 
        SQL = SQL & " WHERE UPPER(CIDADE) = UPPER('" + descricaoPlanta + "') " 
        SQL = SQL & " AND CODIGO_CLIENTE_SAP = RIGHT(REPLICATE('0', 10) + '" & codCliente & "', 10) " 
        SQL = SQL & " AND ("
        SQL = SQL & "   NUMERO_PEDIDO LIKE '%" & pesquisa & "%' "
        SQL = SQL & "   OR NUMERO_NF LIKE '%" & pesquisa & "%' "
        SQL = SQL & ") "
        SQL = SQL & " ORDER BY NOME_CLIENTE "
    Else
        SQL = "SELECT DISTINCT PEDIDO, ITEM, PRODUTO, DESCRICAO_PRODUTO, NUMERO_NF,"
        SQL = SQL & " SERIE_NF, EMISSAO, COD_CLI, NOME_CLI,GRUPO_PRODUTO, "
        SQL = SQL & " QUANT_VENDIDA AS QTDE, VALOR_UNITARIO AS PRECO"
        SQL = SQL & " FROM VW_FNC_PEDIDOS WHERE 1=1" 
        
        If codPlanta = "00" OR codPlanta = "04" OR codPlanta = "07" THEN 
            SQL = SQL & " AND CIDADE = '" & descricaoPlanta & "' AND COD_CLI = '" & codCliente & "' "
        End If

        If codPlanta = "05" THEN 
            SQL = SQL & " AND CIDADE = '" & descricaoPlanta & "' AND CODIGO_ERP = '" & codCliente & "' "
        End If

        If codPlanta = "06" THEN
            SQL = SQL &  "AND CIDADE IN ('FORTALEZA','MARANGUAPE') AND CODIGO_ERP = '" & codCliente & "' "
        End If

        If codPlanta = "07" THEN 
            SQL = SQL & "AND LOJA_CLI = '" & codLoja & "' "
        End If
        
        SQL = SQL & " AND ("
        SQL = SQL & "   PEDIDO LIKE '%" & pesquisa & "%' "
        SQL = SQL & "   OR NUMERO_NF LIKE '%" & pesquisa & "%' "
        SQL = SQL & ") "
    End If

    On Error Resume Next
    SET rsNotas = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsNotas) Then
        ListarNotasCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsNotas.EOF Then
        Do While Not rsNotas.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If
            
            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """numeroPedido"": """ & EscapeJSON(rsNotas("PEDIDO")) & """, "
            dadosJSON = dadosJSON & """itemPedido"": """ & EscapeJSON(rsNotas("ITEM")) & """, "
            dadosJSON = dadosJSON & """precoUnitario"": """ & EscapeJSON(rsNotas("PRECO")) & """, "
            dadosJSON = dadosJSON & """quantidade"": """ & EscapeJSON(rsNotas("QTDE")) & """, "
            dadosJSON = dadosJSON & """codProduto"": """ & EscapeJSON(rsNotas("PRODUTO")) & """, "
            dadosJSON = dadosJSON & """referencia"": """ & EscapeJSON(Replace(rsNotas("DESCRICAO_PRODUTO"), "'", "")) & """, "
            dadosJSON = dadosJSON & """grupoProduto"": """ & EscapeJSON(rsNotas("GRUPO_PRODUTO")) & """, "
            dadosJSON = dadosJSON & """notaFiscal"": """ & EscapeJSON(rsNotas("NUMERO_NF")) & """, "
            dadosJSON = dadosJSON & """serieNotaFiscal"": """ & EscapeJSON(rsNotas("SERIE_NF")) & """, "
            dadosJSON = dadosJSON & """dataEmissaoNota"": """ & EscapeJSON(rsNotas("EMISSAO")) & """"
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsNotas.MoveNext()
        Loop
    End If

    dadosJSON = dadosJSON & "]"
    rsNotas.Close()
    SET rsNotas = NOTHING

    ListarNotasCliente = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ListarHistoricoRespostas()
    idFNC = Request("id")

    SQL = SQL & "SELECT" 
    SQL = SQL & "    r.IdFNC, "
    SQL = SQL & "    REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),r.Descricao),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS Descricao, "
    SQL = SQL & "    REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),r.Justificativa),CHAR(7),''),CHAR(34),''),CHAR(0),'') AS Justificativa, "
    SQL = SQL & "    c1.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "    c2.CRM_NOME as NomeUsuarioFinalizacao, "
    SQL = SQL & "    r.DataInclusao, "
    SQL = SQL & "    r.DataFinalizacao, "
    SQL = SQL & "    Avaliacao "
    SQL = SQL & "FROM Resposta r "
    SQL = SQL & "LEFT JOIN FNC f on f.Id=r.IdFNC "
    SQL = SQL & "LEFT JOIN CRM_DTI c1 on c1.CRM_R_E_C_N_O = r.QuemIncluiu "
    SQL = SQL & "LEFT JOIN CRM_DTI c2 on c2.CRM_R_E_C_N_O = r.QuemFinalizou " 
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "
    
    On Error Resume Next
    SET rsHistorico = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistoricoRespostas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            avaliacao = ""
            If rsHistorico("Avaliacao") = True Then 
                avaliacao = 1
            ElseIf rsHistorico("Avaliacao") = False Then 
                avaliacao = 0 
            End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsHistorico("Descricao")) & """, "
            jsonArray = jsonArray & """justificativa"": """ & EscapeJSON(rsHistorico("Justificativa")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsHistorico("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioFinalizacao"": """ & EscapeJSON(rsHistorico("NomeUsuarioFinalizacao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """, "
            jsonArray = jsonArray & """dataFinalizacao"": """ & EscapeJSON(rsHistorico("DataFinalizacao")) & """, "
            jsonArray = jsonArray & """avaliacao"": """ & avaliacao & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistoricoRespostas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistoricoAprovacoes()
    codPlanta = Request("codPlanta")
    idFNC = Request("id")

    SQL = "SELECT "
    SQL = SQL & "   Aprovacao.IdAprovador, "
    SQL = SQL & "   Aprovacao.IdUsuarioInclusao, "
    SQL = SQL & "   u.CRM_NOME as NomeUsuarioAvaliacao, "
    SQL = SQL & "   u2.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "   Cast(Aprovacao.Nivel as varchar(2)) + ' - ' + "
    SQL = SQL & "   CASE WHEN DataCriacaoRegistro >= '20260504'"
    SQL = SQL & "   THEN "
    SQL = SQL & "       CASE Aprovacao.Nivel "
    SQL = SQL & "           WHEN 1 THEN 'Gerente Setor'"
    SQL = SQL & "           WHEN 2 THEN 'Gerente Qualidade'"
    SQL = SQL & "           WHEN 3 THEN 'Controller'"
    SQL = SQL & "           WHEN 4 THEN 'Gerente Geral'"
    SQL = SQL & "       END"
    SQL = SQL & "   ELSE "
    SQL = SQL & "       CASE Aprovacao.Nivel "
    SQL = SQL & "           WHEN 1 THEN 'Responsavel Qualidade' "
    SQL = SQL & "           WHEN 2 THEN 'Responsavel Producao' "
    SQL = SQL & "           WHEN 3 THEN 'Responsavel Comercial' "
    SQL = SQL & "           WHEN 4 THEN 'Controller' "
    SQL = SQL & "           WHEN 5 THEN 'Gerencia Geral' " 
    SQL = SQL & "       END "
    SQL = SQL & "   END AS Nivel, "
    SQL = SQL & "   CASE isnull(Avaliacao,0) WHEN 0 THEN 'Não' ELSE 'Sim' END Avaliacao, "
    SQL = SQL & "   DataCriacaoRegistro as DataInclusao, "
    SQL = SQL & "   DataAvaliacao, "
    SQL = SQL & "   Motivo, "
    SQL = SQL & "   CASE WHEN Invalido = 1 THEN '1' ELSE '0' END Invalido "
    SQL = SQL & "FROM AprovacaoFNC Aprovacao "
    SQL = SQL & "LEFT JOIN CRM_DTI u ON Aprovacao.IdAprovador = u.CRM_R_E_C_N_O "
    SQL = SQL & "LEFT JOIN CRM_DTI u2 ON Aprovacao.IdUsuarioInclusao = u2.CRM_R_E_C_N_O "
    SQL = SQL & "WHERE IdFNC = " & idFNC & " "
    SQL = SQL & "ORDER BY DataCriacaoRegistro, Aprovacao.Nivel"

    On Error Resume Next
    SET rsHistorico = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistoricoAprovacoes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """idAprovador"": """ & EscapeJSON(rsHistorico("IdAprovador")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsHistorico("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioAvaliacao"": """ & EscapeJSON(rsHistorico("NomeUsuarioAvaliacao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsHistorico("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nivel"": """ & EscapeJSON(rsHistorico("Nivel")) & """, "
            jsonArray = jsonArray & """avaliacao"": """ & EscapeJSON(rsHistorico("Avaliacao")) & """, "
            jsonArray = jsonArray & """dataAvaliacao"": """ & EscapeJSON(rsHistorico("DataAvaliacao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """, "
            jsonArray = jsonArray & """motivo"": """ & EscapeJSON(rsHistorico("Motivo")) & """, "
            jsonArray = jsonArray & """invalido"": """ & EscapeJSON(rsHistorico("Invalido")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistoricoAprovacoes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistoricoObservacoes()
    idFNC = Request("id")

   ' Obtém observações
    SQL = "SELECT o.Id, o.IdFnc, o.Descricao, o.IdUsuarioInclusao, o.DataCriacao, u.CRM_LOGIN as LoginUsuarioInclusao, u.CRM_NOME AS NomeUsuarioInclusao "
    SQL = SQL & "FROM ObservacaoFnc o "
    SQL = SQL & "LEFT JOIN CRM_DTI u ON o.IdUsuarioInclusao = u.CRM_R_E_C_N_O "
    SQL = SQL & "WHERE o.IdFnc = " & idFNC & " "
    SQL = SQL & "AND o.FlagInativo = 0"

    On Error Resume Next
    SET rsObservacoes = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObservacoes) Then
        ListarHistoricoObservacoes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    jsonArray = "["

    If Not rsObservacoes.EOF Then
        Do While Not rsObservacoes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & rsObservacoes("Id") & """, "
            jsonArray = jsonArray & """idFNC"": """ & rsObservacoes("IdFnc") & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsObservacoes("Descricao")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & rsObservacoes("IdUsuarioInclusao") & """, "
            jsonArray = jsonArray & """loginUsuarioInclusao"": """ & EscapeJSON(rsObservacoes("LoginUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsObservacoes("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & rsObservacoes("DataCriacao") & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsObservacoes.MoveNext()
        Loop
    End If

    jsonArray = jsonArray & "]"
    rsObservacoes.Close()
    SET rsObservacoes = NOTHING

    ListarHistoricoObservacoes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterNivelUsuario()
    nivelUsuario = ""
    plantas = ""
    setores = ""

    ' Query para verificar se é Analista ou Responsável
    SQL = "SELECT TOP 1" 
    SQL = SQL & "    r.IdSetor as Setores, "
    SQL = SQL & "    CASE WHEN r.Tipo = 1 THEN 'ANALISTA' ELSE 'RESPONSAVEL' END Nivel, "
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
    SQL = SQL & "WHERE (r.Excluido = 0 OR r.Excluido IS NULL) "
    SQL = SQL & "AND r.Id = " & usuarioLogado("id")

    SET rsResponsavel = Conecta("BEN-CORPORATIVO", SQL)
    
    ' Se der erro finaliza fluxo
    If Err.Number <> 0 OR Not IsObject(rsResponsavel) Then
        ObterNivelUsuario = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsResponsavel.EOF Then
        nivelUsuario = rsResponsavel("Nivel")
        plantas = rsResponsavel("Plantas")
        setores = rsResponsavel("Setores")
    End If

    rsResponsavel.Close()
    SET rsResponsavel = NOTHING
    
    If nivelUsuario <> "" AND plantas <> "" Then
        dadosJSON = "{""nivel"": """ & nivelUsuario & """, " & """plantas"": """ & plantas & """, " & """setores"": """ & setores & """}"
        ObterNivelUsuario = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
        Exit Function
    End If

    SQL = "SELECT TOP 1 "
    SQL = SQL & "    a.Id, "
    SQL = SQL & "    STUFF(( "
    SQL = SQL & "        SELECT ', ' + a2.CodPlanta "
    SQL = SQL & "        FROM AprovadorFNCParametro a2 "
    SQL = SQL & "        WHERE a2.IdUsuario = a.Id"
    SQL = SQL & "        ORDER BY a2.CodPlanta"
    SQL = SQL & "        FOR XML PATH('')), 1, 2, '' "
    SQL = SQL & "    ) AS Plantas,"
    SQL = SQL & "    STUFF(( "
    SQL = SQL & "        SELECT ', ' + CONVERT(VARCHAR(8), a2.IdSetor) "
    SQL = SQL & "        FROM AprovadorFNCParametro a2 "
    SQL = SQL & "        WHERE a2.IdUsuario = a.Id"
    SQL = SQL & "        ORDER BY a2.IdSetor"
    SQL = SQL & "        FOR XML PATH('')), 1, 2, '' "
    SQL = SQL & "    ) AS Setores "
    SQL = SQL & "FROM AprovadorFNC a "
    SQL = SQL & "WHERE a.Id = " & usuarioLogado("id") & " "
    SQL = SQL & "AND (a.Excluido = 0 OR a.Excluido IS NULL)"

    SET rsAprovador = Conecta("BEN-CORPORATIVO", SQL)
    
    ' Se der erro finaliza fluxo
    If Err.Number <> 0 OR Not IsObject(rsAprovador) Then
        ObterNivelUsuario = "{""sucesso"": false, ""mensagem"": ""Erro ao obter nível de usuário: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    If Not rsAprovador.EOF Then
        nivelUsuario = "APROVADOR"
        plantas = rsAprovador("Plantas")
        setores = rsAprovador("Setores")
    End If

    rsAprovador.Close()
    SET rsAprovador = NOTHING

    ' Se não tiver nenhum nível consideramos que é solicitante
    If nivelUsuario = "" Then
        nivelUsuario = "SOLICITANTE"
        plantas = usuarioLogado("planta")
        setores = ""
    End If
    
    dadosJSON = "{""nivel"": """ & nivelUsuario & """, " & """plantas"": """ & plantas & """, " & """setores"": """ & setores & """}"
    ObterNivelUsuario = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ObterIdsFNCAprovador(usuarioId)
    retorno = ""

    SQL = SQL & "WITH APROVADOR AS ( "
    SQL = SQL & "   SELECT  "
    SQL = SQL & "       IdSetor, "
    SQL = SQL & "       CodPlanta, "
    SQL = SQL & "       Nivel "
    SQL = SQL & "   FROM AprovadorFNCParametro "
    SQL = SQL & "   WHERE IdUsuario = " & usuarioId
    SQL = SQL & ")"

    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND F.IdSetor = A.IdSetor AND A.Nivel = F.NivelAprovacaoAtual "
    
    SQL = SQL & "UNION "
   
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 2 "
    
    SQL = SQL & "UNION "
    
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 3 "
    
    SQL = SQL & "UNION "
    
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 4 "

    On Error Resume Next
    SET rsIds = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsIds) Then
        ObterIdsFNCAprovador = retorno
        Exit Function
    End If

    retorno = ""
    primeiroItem = True
    While Not rsIds.EOF
        If Not primeiroItem Then retorno = retorno & "," End If
        
        retorno = retorno & rsIds("Id")
        primeiroItem = False
        rsIds.MoveNext()
    Wend
    
    rsIds.Close()
    SET rsIds = NOTHING
    On Error Goto 0
    
    ObterIdsFNCAprovador = retorno
End Function

' Auxiliares
Function VerificarRespostaPendente(idFNC)
    flagExiste = False

    SQL = "SELECT 1 FROM Resposta WHERE IdFNC = " & idFNC & " AND Finalizada = 0 "

    On Error Resume Next
    SET rsExiste = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        VerificarRespostaPendente = flagExiste
        Exit Function
    End If

    If Not rsExiste.EOF Then
       flagExiste = True
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    On Error Goto 0
    
    VerificarRespostaPendente = flagExiste
End Function

Function ObterQuantidadeAprovacoes(codPlanta, valorReclamado)
    plantasMediaMenorQue8m2 = Array("05", "06", "07")
    plantasMediaMaiorQue8m2 = Array("00", "04", "08", "09", "10", "11", "12", "13")
    quantidadeAprovacoes = 0

    If valorReclamado = 0 Then
        quantidadeAprovacoes = 1
        ObterQuantidadeAprovacoes = quantidadeAprovacoes
        Exit Function
    End If

    If ContidoNoArray(plantasMediaMenorQue8m2, codPlanta) Then
        'Bento Gonçalves, Maranguape e Saquarema: 3 níveis
        If valorReclamado < 15001 Then
            quantidadeAprovacoes = 3
        Else
            quantidadeAprovacoes = 4
        End If
    ElseIf ContidoNoArray(plantasMediaMaiorQue8m2, codPlanta) Then
        If valorReclamado < 50001 Then
            quantidadeAprovacoes = 3
        Else
            quantidadeAprovacoes = 4
        End If
    End If

    ObterQuantidadeAprovacoes = quantidadeAprovacoes
End Function

Function ObterLimitesNQA() 
    dadosJSON = "{""critico"": """ & Replace(limitesNQA("CRITICO"), ",", ".") & """, " & """grave"": """ & Replace(limitesNQA("GRAVE"), ",", ".") & """}"
    ObterLimitesNQA = "{""sucesso"": true, ""mensagem"": ""Nível obtido com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function UsuarioPodeValidarEtapa(dadosFNC, idUsuario)
    temPermissao = False

    SQL = "SELECT 1 FROM AprovadorFNCParametro "
    SQL = SQL & "WHERE IdUsuario = '" & idUsuario & "' "
    SQL = SQL & "AND CodPlanta = '" & dadosFNC("codPlanta") & "' "
    SQL = SQL & "AND Nivel = '" & dadosFNC("nivelAprovacaoAtual") & "' "

    If nivelAprovacaoAtual = 1 Then 
        ' Só valida setor caso nível de aprovação seja 1
        SQL = SQL & "AND IdSetor = '" & dadosFNC("idSetor") & "' "
    End If

    On Error Resume Next
    SET rsPermissao = Conecta("BEN-CORPORATIVO", SQL)

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


Function ResponsavelExiste(idFNC, idResponsavel)
    flagExiste = False

    SQL = "SELECT 1 FROM FNC WHERE Id = " & idFNC & " and IdAnalista = " & idResponsavel 

    On Error Resume Next
    SET rsExiste = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        ResponsavelExiste = flagExiste
        Exit Function
    End If

    If Not rsExiste.EOF Then
       flagExiste = True
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    On Error Goto 0
    
    ResponsavelExiste = flagExiste
End Function

Function ClientePertenceCarteira(codPlanta, codCliente, codLoja, descricaoPlanta)
    flagPertence = False

    SQL = "SELECT COD_VENDEDOR " 
    SQL = SQL & "FROM VW_CLIENTES " 
    SQL = SQL & "WHERE CIDADE = '" & plantaDescricao & "'"   
    SQL = SQL & " AND CASE WHEN '" & codPlanta & "' IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END = '" & codCliente & "'  COLLATE SQL_Latin1_General_CP1_CI_AS " 
    SQL = SQL & " AND CASE WHEN '" & codPlanta & "' IN ('06','05') THEN '01' ELSE LOJA_CLIENTE END = '" &  codLoja & "' COLLATE SQL_Latin1_General_CP1_CI_AS " 
    SQL = SQL & " AND COD_VENDEDOR = " & usuarioLogado("cod")

    On Error Resume Next
    SET rsPertence = Conecta("SPO-BI", SQL)

    If Err.Number <> 0 OR Not IsObject(rsPertence) Then
        ClientePertenceCarteira = flagPertence
        Exit Function
    End If

    If Not rsPertence.EOF Then
       flagPertence = True
    End If
    
    rsPertence.Close()
    SET rsPertence = NOTHING
    On Error Goto 0
    
    ClientePertenceCarteira = flagPertence
End Function

Function ObterResponsavel(codPlanta)
    responsavelId = 0

    SQL = "SELECT TOP 1 Id FROM Responsavel WHERE isnull(Excluido,'') = '' and Tipo = 1 and Planta = '" & codPlanta & "'"

    On Error Resume Next
    SET rsResponsavel = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsResponsavel) Then
        ObterResponsavel = responsavelId
        Exit Function
    End If

    If Not rsResponsavel.EOF Then
       responsavelId = rsResponsavel("Id")
    End If
    
    rsResponsavel.Close()
    SET rsResponsavel = NOTHING
    On Error Goto 0
    
    ObterResponsavel = responsavelId
End Function

Function ObterSetor(defeitoId)
    setorId = 0

    SQL = "SELECT IdSetor, IdDefeito FROM SetorDefeito WHERE IdDefeito = " & defeitoId

    On Error Resume Next
    SET rsSetor = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsSetor) Then
        ObterSetor = setorId
        Exit Function
    End If

    If Not rsSetor.EOF Then
       setorId = rsSetor("IdSetor")
    End If
    
    rsSetor.Close()
    SET rsSetor = NOTHING
    On Error Goto 0
    
    ObterSetor = setorId
End Function

Function ExisteDuplicidade(pQuemIncluiu, pQuantidade, pValorReclamado, pCodCliente, pIdDisposicao, pIdDefeito, pIdSetor, pPlanta )
    flagExiste = False

    SQL = "SELECT COUNT(Id) AS CONTADOR " 
    SQL = SQL & "FROM FNC " 
    SQL = SQL & "WHERE QuemIncluiu = '" & pQuemIncluiu & "' " 
    SQL = SQL & "   AND Quantidade = " & pQuantidade & " " 
    SQL = SQL & "   AND ValorReclamado = " & pValorReclamado & " " 
    SQL = SQL & "   AND Cliente = '" & pCodCliente & "' " 
    SQL = SQL & "   AND IdDisposicao = " & pIdDisposicao & " " 
    SQL = SQL & "   AND IdDefeito = " & pIdDefeito & " " 
    SQL = SQL & "   AND IdSetor = " & pIdSetor & " " 
    SQL = SQL & "   AND Planta = '" & pPlanta & "'"
    SQL = SQL & "   AND Situacao <> 3 "
  
    On Error Resume Next
    SET rsExiste = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        ExisteDuplicidade = flagExiste
        Exit Function
    End If

    If Not rsExiste.EOF Then
       flagExiste = Iif(rsExiste("CONTADOR") > 0, True, False) 
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    On Error Goto 0
    
    ExisteDuplicidade = flagExiste
End Function

Function ObterDescricaoPlanta(codPlanta)
    Dim descricaoPlanta 
    descricaoPlanta = ""

    Select Case codPlanta
        Case "00"
            descricaoPlanta = "PIRAPETINGA"
        Case "04"
            descricaoPlanta = "UBERABA"
        Case "05"
            descricaoPlanta = "BENTO"
        Case "06"
            descricaoPlanta = "MARANGUAPE"
        Case "07"
            descricaoPlanta = "SAQUAREMA"
        Case "08"
            descricaoPlanta = "ARACATUBA"
        Case "09"
            descricaoPlanta = "BLUMENAU"
        Case "10"
            descricaoPlanta = "PACAJUS"
        Case "11"
            descricaoPlanta = "PORTO FELIZ"
        Case "12"
            descricaoPlanta = "TRES BARRAS"
        Case "13"
            descricaoPlanta = "FLORESTAL"
        Case Else
            descricaoPlanta = "Desconhecido"
    End Select
    
    ObterDescricaoPlanta = descricaoPlanta
End Function

Function VerificarPlantaWestrock(codPlanta)
    ehPlantaWestrock = codPlanta = "08" OR codPlanta = "09" OR codPlanta = "10" OR codPlanta = "11" OR codPlanta = "12" OR codPlanta = "13"
    VerificarPlantaWestrock = ehPlantaWestrock
End Function

Function AprovadorNivelMaiorQueUm(idUsuario)
    ehAprovadorNivelMaiorQueUm = False

    SQL = SQL & "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS nivelMaiorQueUm FROM AprovadorFNCParametro "
    SQL = SQL & "WHERE IdUsuario = '" & idUsuario & "' "
    SQL = SQL & "AND Nivel <> 1"

    On Error Resume Next
    SET rsVerificar = Conecta("BEN-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsVerificar) Then
        AprovadorNivelMaiorQueUm = ehAprovadorNivelMaiorQueUm
        Exit Function
    End If
    On Error Goto 0

    If rsVerificar("nivelMaiorQueUm") = 1 Then
        ehAprovadorNivelMaiorQueUm = True
    End If

    rsVerificar.Close()
    SET rsVerificar = NOTHING

    AprovadorNivelMaiorQueUm = ehAprovadorNivelMaiorQueUm
End Function

Sub EnviarEmailCriacao(idNovaFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, listaPedidos)
    emailAnalistaFNC = ""
    emailUsuarioCriacao = ""

    SQL = "SELECT Email FROM Responsavel WHERE Planta = '" & codPlanta & "' " 
    SQL = SQL & "AND Tipo = 1 " 
    SQL = SQL & "AND ISNULL(Excluido,'0') = '0' "

    On Error Resume Next
    SET rsResponsavel = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsResponsavel) Then
        emailAnalistaFNC = rsResponsavel("Email")

        While Not rsResponsavel.EOF
            emailAnalistaFNC = emailAnalistaFNC & ";" & rsResponsavel("Email")
            rsResponsavel.MoveNext()
        Wend
        
        rsResponsavel.Close()
        SET rsResponsavel = NOTHING
    End If
    On Error Goto 0

    SQL = "SELECT CRM_EMAILPRIMARIO as email FROM CRM_DTI WHERE CRM_LOGIN = '" & usuarioLogado("login") & "'"
    
    On Error Resume Next
    SET rsCriador = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsCriador) Then
        emailUsuarioCriacao = rsCriador("Email")
        rsCriador.Close()
        SET rsCriador = NOTHING
    End If 
    On Error Goto 0

    assunto = "FNC: " & idNovaFNC & " incluída por " & usuarioLogado("nome")

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
    mensagem = mensagem & "<p>Uma FNC foi inserida por " & usuarioLogado("nome") & ".</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idNovaFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Pedidos:</strong> " & listaPedidos & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Data de Inclusão:</strong> " & Now() & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailUsuarioCriacao <> "" or emailAnalistaFNC <> "" Then
        'emailAnalistaFNC = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(emailAnalistaFNC, emailUsuarioCriacao, "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailAvaliacao(flagAprovar, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, dataInclusao)
    emailAnalistaFNC = ""

    SQL = "SELECT Email FROM Responsavel WHERE Planta = '" & codPlanta & "' " 
    SQL = SQL & "AND Tipo = 1 " 
    SQL = SQL & "AND ISNULL(Excluido,'0') = '0' "

    On Error Resume Next
    SET rsAnalista = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsAnalista) Then
        emailAnalistaFNC = rsAnalista("Email")

        While Not rsAnalista.EOF
            emailAnalistaFNC = emailAnalistaFNC & ";" & rsAnalista("Email")
            rsAnalista.MoveNext()
        Wend
        
        rsAnalista.Close()
        SET rsAnalista = NOTHING
    End If

    If codPlanta = "10" Then
		'emailAnalistaFNC = emailAnalistaFNC & ";grp-qm-devolucoes@WestRockCo.onmicrosoft.com;joao.k.assuncao@westrock.com;CS-PACAJUS@westrock.com"		
	End If

    On Error Goto 0

    assunto = "FNC: " & idFNC & " foi " & Iif(flagAprovar = 1, "aprovada por ", "rejeitada por ") & usuarioLogado("nome")

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
    mensagem = mensagem & "<p>Uma FNC foi" & Iif(flagAprovar = 1, " aprovada", " rejeitada") &  " por " & usuarioLogado("nome") & ".</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Data de Inclusão:</strong> " & dataInclusao & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If emailAnalistaFNC <> "" Then
        'emailAnalistaFNC = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(emailAnalistaFNC, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailResposta(flagAprovar, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, idUsuarioInclusao)
    listaEmails = ""

    SQL = "SELECT Email FROM Responsavel WHERE Planta = '" & codPlanta & "' " 
    SQL = SQL & "AND Tipo = 1 " 
    SQL = SQL & "AND ISNULL(Excluido,'0') = '0' "

    On Error Resume Next
    SET rsAnalista = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsAnalista) Then
        listaEmails = rsAnalista("Email")

        While Not rsAnalista.EOF
            If rsAnalista("Email") <> "" Then
                listaEmails = listaEmails & ";" & rsAnalista("Email")
            End If
            rsAnalista.MoveNext()
        Wend
        
        rsAnalista.Close()
        SET rsAnalista = NOTHING
    End If
    On Error Goto 0

    SQL = "SELECT CRM_EMAILPRIMARIO as Email FROM CRM_DTI WHERE CRM_R_E_C_N_O = '" & idUsuarioInclusao & "'"
    
    On Error Resume Next
    SET rsCriador = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsCriador) Then
        If rsCriador("Email") <> "" Then
            listaEmails = listaEmails & ";" & rsCriador("Email")
        End If

        rsCriador.Close()
        SET rsCriador = NOTHING
    End If 
    On Error Goto 0

    assunto = "FNC: " & idFNC & " Devolvida/Respondida a você por " & usuarioLogado("nome")

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
    mensagem = mensagem & "<p>Uma FNC foi" & Iif(flagAprovar = 1, " respondida", " devolvida") &  " por " & usuarioLogado("nome") & ".</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Data de Inclusão:</strong> " & Now() & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If listaEmails <> "" Then
        'listaEmails = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(listaEmails, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailResponsavel(idResponsavel, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, dataInclusao)
    emailResponsavelFNC = ""

    SQL = "SELECT Email FROM Responsavel WHERE Planta = '" & codPlanta & "' " 
    SQL = SQL & "AND Tipo = 2 " 
    SQL = SQL & "AND ISNULL(Excluido,'0') = '0' "
    SQL = SQL & "AND Id = " & idResponsavel & " "

    On Error Resume Next
    SET rsAnalista = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsAnalista) Then
        emailResponsavelFNC = rsAnalista("Email")

        While Not rsAnalista.EOF
            emailResponsavelFNC = emailResponsavelFNC & ";" & rsAnalista("Email")
            rsAnalista.MoveNext()
        Wend
        
        rsAnalista.Close()
        SET rsAnalista = NOTHING
    End If
    On Error Goto 0

    assunto = "FNC: " & idFNC & " foi direcionada a você pelo setor de Qualidade"

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
    mensagem = mensagem & "<p>Uma FNC foi direcionada a voce pelo setor de qualidade e aguarda sua resposta.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If emailResponsavelFNC <> "" Then  
        'emailResponsavelFNC = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(emailResponsavelFNC, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailAprovadores(nivelAprovador, idSetor, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao)
    emailAprovadoresFNC = ""

    SQL = "SELECT a.Email " 
    SQL = SQL & " FROM AprovadorFNC a"
    SQL = SQL & " WHERE (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " AND EXISTS ( "
    SQL = SQL & "    SELECT TOP 1 a.Id "
    SQL = SQL & "    FROM AprovadorFNCParametro p "
    SQL = SQL & "    WHERE p.IdUsuario = a.Id "
    SQL = SQL & "    AND p.CodPlanta = " & codPlanta
    SQL = SQL & "    AND p.Nivel = " & nivelAprovador

    If nivelAprovador = 1 Then 
        ' Só valida setor caso nível de aprovador seja 1
        SQL = SQL & "    AND p.IdSetor = " & idSetor
    End If

    SQL = SQL & ") "

    On Error Resume Next
    SET rsAprovador = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsAprovador) Then
        emailAprovadoresFNC = rsAprovador("Email")

        While Not rsAprovador.EOF
            emailAprovadoresFNC = emailAprovadoresFNC & ";" & rsAprovador("Email")
            rsAprovador.MoveNext()
        Wend
        
        rsAprovador.Close()
        SET rsAprovador = NOTHING
    End If
    On Error Goto 0

    assunto = "FNC: " & idFNC & " está pendente de aprovação"

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
    mensagem = mensagem & "<p>Uma FNC foi direcionada a voce e aguarda sua aprovação.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If emailAprovadoresFNC <> "" Then  
        'emailAprovadoresFNC = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(emailAprovadoresFNC, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailAprovadoTotal(flagAprovar, idFNC, nomeCliente, codCliente, codLoja, codPlanta, descricaoPlanta, quantidadeReclamada, valorReclamado, descricaoDefeito, descricaoDisposicao, dataInclusao, idUsuarioInclusao)
    listaEmails = ""

    SQL = "SELECT Email FROM Responsavel WHERE Planta = '" & codPlanta & "' " 
    SQL = SQL & "AND Tipo = 1 " 
    SQL = SQL & "AND ISNULL(Excluido,'0') = '0' "

    On Error Resume Next
    SET rsAnalista = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsAnalista) Then
        While Not rsAnalista.EOF
            If rsAnalista("Email") <> "" Then
                listaEmails = listaEmails & ";" & rsAnalista("Email")
            End If

            rsAnalista.MoveNext()
        Wend
        
        rsAnalista.Close()
        SET rsAnalista = NOTHING
    End If
    On Error Goto 0

    SQL = "SELECT CRM_EMAILPRIMARIO as Email FROM CRM_DTI WHERE CRM_R_E_C_N_O = '" & idUsuarioInclusao & "'"
    
    On Error Resume Next
    SET rsCriador = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsCriador) Then
        If rsCriador("Email") <> "" Then
            listaEmails = listaEmails & ";" & rsCriador("Email")
        End If

        rsCriador.Close()
        SET rsCriador = NOTHING
    End If 
    On Error Goto 0

    assunto = "FNC: " & idFNC & " finalizou seu fluxo de aprovações"

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
    mensagem = mensagem & "<p>Uma FNC finalizou seu fluxo de aprovações</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Número:</strong> " & idFNC & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & descricaoPlanta & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Quantidade Reclamada:</strong> " & quantidadeReclamada & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Valor Reclamado:</strong> " & valorReclamado & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Problema Apontado:</strong> " & descricaoDefeito & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Descrição Cliente:</strong> " & descricaoDisposicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Data de Inclusão:</strong> " & dataInclusao & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If listaEmails <> "" Then
        'listaEmails = "thiago.pereira@smurfitwestrock.com.br"
        Call EnviaEmail(listaEmails, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub RegistrarLog(acao, id, idResponsavel)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a FNC de id " & id
        Case "edicao"
            acaoLog = "Edição de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou a FNC de id " & id
        Case "aprovacao"
            acaoLog = "Aprovação de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a FNC de id " & id
        Case "rejeicao"
            acaoLog = "Rejeição de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " rejeitou a FNC de id " & id
        Case "envio_responsavel"
            acaoLog = "Envio para Responsável FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou a FNC de id " & id & " para o responsável de id " & idResponsavel
        Case "envio_aprovacao"
            acaoLog = "Envio para Aprovação FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " enviou a FNC de id " & id & " para a aprovação"
        Case "resposta"
            acaoLog = "Resposta de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " respondeu a FNC de id " & id 
        Case "devolucao"
            acaoLog = "Devolução de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " devolveu a FNC de id " & id 
        Case "cancelamento"
            acaoLog = "Cancelamento de FNC"
            descricaoLog = "Usuário " & usuarioLogado("login") & " cancelou a FNC de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'FNC', "
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