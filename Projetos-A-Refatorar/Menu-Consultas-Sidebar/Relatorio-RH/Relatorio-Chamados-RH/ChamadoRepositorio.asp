<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
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

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Select Case acao
    Case "listar"
        resultado = Listar()
    Case "listarFila"
        resultado = ListarFila()
    Case "criar"
        resultado = Criar()
    Case "enviarMensagem"
        resultado = EnviarMensagem()
    Case "obterChamado"
        resultado = ObterChamado()
    Case "obterDadosKPI"
        resultado = ObterDadosKPI()
    Case "avaliarChamado"
        resultado = AvaliarChamado()
    Case "finalizarChamado"
        resultado = FinalizarChamado()
    Case "reabrirChamado"
        resultado = ReabrirChamado()
    Case "atribuirAtendente"
        resultado = AtribuirAtendente()
    Case "alterarStatus"
        resultado = AlterarStatus()
    Case "alterarPrioridade"
        resultado = AlterarPrioridade()
    Case "alterarCategoria"
        resultado = AlterarCategoria()
    Case "alterarDepartamento"
        resultado = AlterarDepartamento()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    codPlanta = Request("filtroFilial")
    dataCriacaoDe = Request("filtroDataCriacaoDe")
    dataCriacaoAte = Request("filtroDataCriacaoAte")
    dataFinalizacaoDe = Request("filtroDataFinalizacaoDe")
    dataFinalizacaoAte = Request("filtroDataFinalizacaoAte")
    status = Request("filtroStatus")
    solicitante = Request("filtroSolicitante")
    prioridade = Request("filtroPrioridade")
    idChamado = Request("filtroIdChamado")
    idAtendente = Request("filtroAtendente")
    idCategoria = Request("filtroCategoria")
    idDepartamento = Request("filtroDepartamento")
    flagMeusChamados = Request("flagMeusChamados")

    SQL = "SELECT "
    SQL = SQL & "   c.Id, c.CodPlanta, c.IdSolicitante, c.IdAtendente, c.IdPrioridade, "
    SQL = SQL & "   c.IdDepartamento, d.Nome as NomeDepartamento, c.IdCategoria,"
    SQL = SQL & "   c.TokenAcesso, c.Status, c.Titulo, s.Nome as NomeSolicitante,  "
    SQL = SQL & "   c.DataCriacao, c.DataFinalizacao, c.Descricao, a.Nome as NomeAtendente, "
    SQL = SQL & "   ca.Nome as NomeCategoria, c.DataPrevista, p.Nivel as NivelPrioridade "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN AtendentesChamado a ON a.Id = c.IdAtendente "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "LEFT JOIN CategoriasChamado ca ON ca.Id = c.IdCategoria "
    SQL = SQL & "LEFT JOIN DepartamentosChamado d on d.Id = c.IdDepartamento "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "

    If idChamado <> "" Then
        SQL = SQL & " AND c.Id = " & idChamado
    Else  
        If flagMeusChamados = "1" Then
            SQL = SQL & " AND s.IdUsuario = '" & usuarioLogado("id") & "'"
        End If

        If codPlanta <> "" Then
            SQL = SQL & " AND c.CodPlanta = '" & codPlanta & "'"
        End If

        If dataCriacaoDe <> "" Then
            SQL = SQL & " AND c.DataCriacao >= '" & Replace(dataCriacaoDe, "-", "") & "'"
        End If

        If dataCriacaoAte <> "" Then
            SQL = SQL & " AND c.DataCriacao < DATEADD(DAY, 1, '" & Replace(dataCriacaoAte & "", "-", "") & "')"
        End If

        If dataFinalizacaoDe <> "" Then
            SQL = SQL & " AND c.DataFinalizacao >= '" & Replace(dataFinalizacaoDe, "-", "") & "'"
        End If

        If dataFinalizacaoAte <> "" Then
            SQL = SQL & " AND c.DataFinalizacao < DATEADD(DAY, 1, '" & Replace(dataFinalizacaoAte & "", "-", "") & "')"
        End If

        If solicitante <> "" Then
            SQL = SQL & " AND ("
            SQL = SQL & "   s.Nome LIKE '%" & Replace(solicitante, "'", "''") & "%'"
            SQL = SQL & "   OR s.Email LIKE '%" & Replace(solicitante, "'", "''") & "%'"
            SQL = SQL & " )"
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

        If prioridade <> "" Then
            SQL = SQL & " AND p.Nivel = '" & prioridade & "'"
        End If

        If idAtendente <> "" Then
            SQL = SQL & " AND c.IdAtendente = '" & idAtendente & "'"
        End If

        If idCategoria <> "" Then
            SQL = SQL & " AND c.IdCategoria = '" & idCategoria & "'"
        End If

        If idDepartamento <> "" Then
            SQL = SQL & " AND c.IdDepartamento = '" & idDepartamento & "'"
        End If
    End If

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de chamados: " & EscapeJSON(Err.Description) & """}"
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
            jsonArray = jsonArray & """idSolicitante"": """ & EscapeJSON(rsListar("IdSolicitante") & "") & """, "
            jsonArray = jsonArray & """nomeSolicitante"": """ & EscapeJSON(rsListar("NomeSolicitante") & "") & """, "
            jsonArray = jsonArray & """idAtendente"": """ & EscapeJSON(rsListar("IdAtendente") & "") & """, "
            jsonArray = jsonArray & """nomeAtendente"": """ & EscapeJSON(rsListar("NomeAtendente") & "") & """, "
            jsonArray = jsonArray & """idPrioridade"": """ & EscapeJSON(rsListar("IdPrioridade") & "") & """, "
            jsonArray = jsonArray & """nivelPrioridade"": """ & EscapeJSON(rsListar("NivelPrioridade") & "") & """, "
            jsonArray = jsonArray & """idDepartamento"": """ & EscapeJSON(rsListar("IdDepartamento") & "") & """, "
            jsonArray = jsonArray & """nomeDepartamento"": """ & EscapeJSON(rsListar("NomeDepartamento") & "") & """, "
            jsonArray = jsonArray & """idCategoria"": """ & EscapeJSON(rsListar("IdCategoria") & "") & """, "
            jsonArray = jsonArray & """nomeCategoria"": """ & EscapeJSON(rsListar("NomeCategoria") & "") & """, "
            jsonArray = jsonArray & """tokenAcesso"": """ & EscapeJSON(rsListar("TokenAcesso") & "") & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status") & "") & """, "
            jsonArray = jsonArray & """titulo"": """ & EscapeJSON(rsListar("Titulo") & "") & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsListar("Descricao") & "") & """, "
            jsonArray = jsonArray & """dataPrevista"": """ & EscapeJSON(rsListar("DataPrevista") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """, "
            jsonArray = jsonArray & """dataFinalizacao"": """ & EscapeJSON(rsListar("DataFinalizacao") & "") & """"
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

Function ListarFila()
    idAtendente = ConsultarAtendente(usuarioLogado("id"))
    If idAtendente = "" Then
        ListarFila = "{""sucesso"": false, ""mensagem"": ""Você não possui acesso a essa página""}"
        Exit Function
    End If

    SQL = "SELECT c.Id, c.CodPlanta, c.IdSolicitante, c.IdAtendente, c.IdPrioridade, c.IdDepartamento, "
    SQL = SQL & "c.IdCategoria, c.TokenAcesso, c.Status, c.Titulo, s.Nome as NomeSolicitante, a.Nome as NomeAtendente, "
    SQL = SQL & "c.DataCriacao, c.DataFinalizacao, c.Descricao, c.DataPrevista, p.Nivel as NivelPrioridade "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN AtendentesChamado a ON a.Id = c.IdAtendente "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "
    SQL = SQL & "AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        ListarFila = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de chamados: " & EscapeJSON(Err.Description) & """}"
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
            jsonArray = jsonArray & """idSolicitante"": """ & EscapeJSON(rsListar("IdSolicitante") & "") & """, "
            jsonArray = jsonArray & """nomeSolicitante"": """ & EscapeJSON(rsListar("NomeSolicitante") & "") & """, "
            jsonArray = jsonArray & """idAtendente"": """ & EscapeJSON(rsListar("IdAtendente") & "") & """, "
            jsonArray = jsonArray & """nomeAtendente"": """ & EscapeJSON(rsListar("NomeAtendente") & "") & """, "
            jsonArray = jsonArray & """idPrioridade"": """ & EscapeJSON(rsListar("IdPrioridade") & "") & """, "
            jsonArray = jsonArray & """nivelPrioridade"": """ & EscapeJSON(rsListar("NivelPrioridade") & "") & """, "
            jsonArray = jsonArray & """idDepartamento"": """ & EscapeJSON(rsListar("IdDepartamento") & "") & """, "
            jsonArray = jsonArray & """idCategoria"": """ & EscapeJSON(rsListar("IdCategoria") & "") & """, "
            jsonArray = jsonArray & """tokenAcesso"": """ & EscapeJSON(rsListar("TokenAcesso") & "") & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status") & "") & """, "
            jsonArray = jsonArray & """titulo"": """ & EscapeJSON(rsListar("Titulo") & "") & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsListar("Descricao") & "") & """, "
            jsonArray = jsonArray & """dataPrevista"": """ & EscapeJSON(rsListar("DataPrevista") & "") & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsListar("DataCriacao") & "") & """, "
            jsonArray = jsonArray & """dataFinalizacao"": """ & EscapeJSON(rsListar("DataFinalizacao") & "") & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsListar.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsListar.Close()
    SET rsListar = NOTHING
    
    ListarFila = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterChamado()    
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    SQL = "SELECT "
    SQL = SQL & "    c.Id, "
    SQL = SQL & "    c.CodPlanta, "
    SQL = SQL & "    c.IdSolicitante, "
    SQL = SQL & "    c.IdAtendente, "
    SQL = SQL & "    c.IdPrioridade, "
    SQL = SQL & "    c.IdDepartamento, "
    SQL = SQL & "    c.IdCategoria, "
    SQL = SQL & "    ca.Nome as Categoria, "
    SQL = SQL & "    p.Nivel as Prioridade, "
    SQL = SQL & "    d.Nome as Departamento, "
    SQL = SQL & "    c.TokenAcesso, "
    SQL = SQL & "    c.Status, "
    SQL = SQL & "    c.Titulo, "
    SQL = SQL & "    c.Descricao, "
    SQL = SQL & "    a.Nome as NomeAtendente, "
    SQL = SQL & "    c.CentroCusto, "
    SQL = SQL & "    c.IdTipo, "
    SQL = SQL & "    t.Caminho as TipoCaminho, "
    SQL = SQL & "    STRING_ESCAPE(c.Detalhamento, 'json') AS Detalhamento,"
    SQL = SQL & "    CASE WHEN EXISTS ("
    SQL = SQL & "        SELECT TOP 1 a.Id"
    SQL = SQL & "        FROM AvaliacaoChamado a"
    SQL = SQL & "        WHERE a.IdChamado = c.Id"
    SQL = SQL & "    ) THEN 1 ELSE 0 END AS Avaliado,"
    SQL = SQL & "    FORMAT(c.DataCriacao, 'yyyy-MM-dd HH:mm:ss') AS DataCriacao, "
    SQL = SQL & "    FORMAT(c.DataFinalizacao, 'yyyy-MM-dd HH:mm:ss') AS DataFinalizacao, "
    SQL = SQL & "    FORMAT(c.DataPrevista, 'yyyy-MM-dd HH:mm:ss') AS DataPrevista, "
    SQL = SQL & "    FORMAT(c.DataUltimaAtualizacao, 'yyyy-MM-dd HH:mm:ss') AS DataUltimaAtualizacao "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN AtendentesChamado a on a.Id = c.IdAtendente "
    SQL = SQL & "LEFT JOIN CategoriasChamado ca ON ca.Id = c.IdCategoria "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "LEFT JOIN DepartamentosChamado d on d.Id = c.IdDepartamento "
    SQL = SQL & "LEFT JOIN TiposChamado t on t.Id = c.IdTipo "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.Id = '" & autor("idChamado") & "' "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "
    
    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados do chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Chamado não encontrado""}"
        Exit Function
    End If

    ' Obtem Mensagens
    SQL = "SELECT i.Id, i.IdChamado, i.IdAutor, i.TipoAutor, i.NomeAutor, i.NotaInterna, REPLACE(i.Descricao, '""', '''') as Descricao, i.FlagInativo, "
    SQL = SQL & "FORMAT(i.DataCriacao,'yyyy-MM-dd HH:mm') as DataCriacao, "
    SQL = SQL & "( "
    SQL = SQL & "    SELECT a.Id as id, a.Nome as nome, a.Caminho as caminho "
    SQL = SQL & "    FROM AnexosChamado a "
    SQL = SQL & "    WHERE a.FlagInativo = 0 "
    SQL = SQL & "    AND a.IdMensagem = i.Id "
    SQL = SQL & "    FOR JSON PATH "
    SQL = SQL & ") AS Anexos "
    SQL = SQL & "FROM MensagensChamado i "
    SQL = SQL & "WHERE i.IdChamado = '" & autor("idChamado") & "' "
    SQL = SQL & "AND i.FlagInativo = 0 "
    
    If autor("tipo") = "SOLICITANTE" Then
        SQL = SQL & "AND i.NotaInterna = 0"
    End If

    On Error Resume Next
    SET rsMensagens = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsMensagens) Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    mensagensJSON = "["

    If Not rsMensagens.EOF Then
        Do While Not rsMensagens.EOF
            If Not primeiroItem Then mensagensJSON = mensagensJSON & "," End If
            notaInterna = Iif(rsMensagens("NotaInterna") = "True", 1, 0)
            
            mensagensJSON = mensagensJSON & "{"
            mensagensJSON = mensagensJSON & """id"": """ & rsMensagens("Id") & """, "
            mensagensJSON = mensagensJSON & """idChamado"": """ & rsMensagens("IdChamado") & """, "
            mensagensJSON = mensagensJSON & """idAutor"": """ & rsMensagens("IdAutor") & """, "
            mensagensJSON = mensagensJSON & """tipoAutor"": """ & rsMensagens("TipoAutor") & """, "
            mensagensJSON = mensagensJSON & """nomeAutor"": """ & rsMensagens("NomeAutor") & """, "
            mensagensJSON = mensagensJSON & """notaInterna"": """ & notaInterna & """, "
            mensagensJSON = mensagensJSON & """descricao"": """ & EscapeJSON(rsMensagens("Descricao")) & """, "
            mensagensJSON = mensagensJSON & """anexos"": " & Iif(IsNull(rsMensagens("Anexos")), "[]", rsMensagens("Anexos")) & ", "
            mensagensJSON = mensagensJSON & """dataCriacao"": """ & rsMensagens("DataCriacao") & """"
            mensagensJSON = mensagensJSON & "}"
            
            primeiroItem = False
            rsMensagens.MoveNext()
        Loop
    End If

    mensagensJSON = mensagensJSON & "]"
    rsMensagens.Close()
    SET rsMensagens = NOTHING

    ' Obtem histórico
    SQL = "SELECT Id, IdChamado, IdAtendente, NomeAtendente, Acao, DeValor, ParaValor, Descricao, FlagInativo, "
    SQL = SQL & "FORMAT(DataCriacao,'yyyy-MM-dd HH:mm') as DataCriacao "
    SQL = SQL & "FROM HistoricoChamado "
    SQL = SQL & "WHERE IdChamado = " & autor("idChamado")
    SQL = SQL & "AND FlagInativo = 0 "

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    historicoJSON = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then historicoJSON = historicoJSON & "," End If
            
            historicoJSON = historicoJSON & "{"
            historicoJSON = historicoJSON & """id"": """ & rsHistorico("Id") & """, "
            historicoJSON = historicoJSON & """idChamado"": """ & rsHistorico("IdChamado") & """, "
            historicoJSON = historicoJSON & """idAtendente"": """ & rsHistorico("IdAtendente") & """, "
            historicoJSON = historicoJSON & """acao"": """ & rsHistorico("Acao") & """, "
            historicoJSON = historicoJSON & """deValor"": """ & rsHistorico("DeValor") & """, "
            historicoJSON = historicoJSON & """paraValor"": """ & rsHistorico("ParaValor") &  """, "
            historicoJSON = historicoJSON & """descricao"": """ & rsHistorico("Descricao") & """, "
            historicoJSON = historicoJSON & """dataCriacao"": """ & rsHistorico("DataCriacao") &  """"
            historicoJSON = historicoJSON & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext()
        Loop
    End If

    historicoJSON = historicoJSON & "]"
    rsHistorico.Close()
    SET rsHistorico = NOTHING

    ' Obtem Anexos
    SQL = "SELECT Id, IdChamado, Nome, Caminho, FlagInativo, "
    SQL = SQL & "FORMAT(DataCriacao,'yyyy-MM-dd HH:mm') as DataCriacao "
    SQL = SQL & "FROM AnexosChamado "
    SQL = SQL & "WHERE IdChamado = " & autor("idChamado")
    SQL = SQL & "AND FlagInativo = 0 "

    On Error Resume Next
    SET rsAnexo = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAnexo) Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar notas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    anexosJSON = "["

    If Not rsAnexo.EOF Then
        Do While Not rsAnexo.EOF
            If Not primeiroItem Then anexosJSON = anexosJSON & "," End If
            
            anexosJSON = anexosJSON & "{"
            anexosJSON = anexosJSON & """id"": """ & rsAnexo("Id") & """, "
            anexosJSON = anexosJSON & """idChamado"": """ & rsAnexo("IdChamado") & """, "
            anexosJSON = anexosJSON & """nome"": """ & rsAnexo("Nome") & """, "
            anexosJSON = anexosJSON & """caminho"": """ & rsAnexo("Caminho") &  """, "
            anexosJSON = anexosJSON & """dataCriacao"": """ & rsAnexo("DataCriacao") &  """"
            anexosJSON = anexosJSON & "}"
            
            primeiroItem = False
            rsAnexo.MoveNext()
        Loop
    End If

    anexosJSON = anexosJSON & "]"
    rsAnexo.Close()
    SET rsAnexo = NOTHING

    ' Obtem dados do Solicitante
    SQL = "SELECT TOP 1 s.Id, s.Nome, s.Email, s.Telefone, "
    SQL = SQL & "( "
    SQL = SQL & "    SELECT TOP 4 c.Id, c.Titulo, c.DataCriacao "
    SQL = SQL & "    FROM Chamados c "
    SQL = SQL & "    WHERE c.FlagInativo = 0 "
    SQL = SQL & "    AND c.IdSolicitante = s.Id "
    SQL = SQL & "    AND c.Id <> " & autor("idChamado")
    SQL = SQL & "    ORDER BY Id DESC"
    SQL = SQL & "    FOR JSON PATH "
    SQL = SQL & ") AS Historico "
    SQL = SQL & "FROM SolicitantesChamado s "
    SQL = SQL & "WHERE s.Id = '" & rsObter("IdSolicitante") & "' "
    SQL = SQL & "AND s.FlagInativo = 0 "

    On Error Resume Next
    SET rsSolicitante = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSolicitante) Then
        ObterChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados do solicitante: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    solicitanteJSON = "{"

    If Not rsSolicitante.EOF Then        
        solicitanteJSON = solicitanteJSON & """id"": """ & rsSolicitante("Id") & """, "
        solicitanteJSON = solicitanteJSON & """nome"": """ & rsSolicitante("Nome") & """, "
        solicitanteJSON = solicitanteJSON & """email"": """ & rsSolicitante("Email") &  """, " 
        solicitanteJSON = solicitanteJSON & """telefone"": """ & rsSolicitante("Telefone") &  """, " 
        solicitanteJSON = solicitanteJSON & """historico"": " & Iif(IsNull(rsSolicitante("Historico")), "[]", rsSolicitante("Historico")) & ""
    End If

    solicitanteJSON = solicitanteJSON & "}"
    rsSolicitante.Close()
    SET rsSolicitante = NOTHING

    chamadoJSON = chamadoJSON & "{"
    chamadoJSON = chamadoJSON & """id"": """ & EscapeJSON(rsObter("Id") & "") & """, "
    chamadoJSON = chamadoJSON & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta") & "") & """, "
    chamadoJSON = chamadoJSON & """idSolicitante"": """ & EscapeJSON(rsObter("IdSolicitante") & "") & """, "
    chamadoJSON = chamadoJSON & """idAtendente"": """ & EscapeJSON(rsObter("IdAtendente") & "") & """, "
    chamadoJSON = chamadoJSON & """nomeAtendente"": """ & EscapeJSON(rsObter("NomeAtendente") & "") & """, "
    chamadoJSON = chamadoJSON & """idPrioridade"": """ & EscapeJSON(rsObter("IdPrioridade") & "") & """, "
    chamadoJSON = chamadoJSON & """idDepartamento"": """ & EscapeJSON(rsObter("IdDepartamento") & "") & """, "
    chamadoJSON = chamadoJSON & """idCategoria"": """ & EscapeJSON(rsObter("IdCategoria") & "") & """, "
    chamadoJSON = chamadoJSON & """idTipo"": """ & EscapeJSON(rsObter("IdTipo") & "") & """, "
    chamadoJSON = chamadoJSON & """tipoCaminho"": """ & EscapeJSON(rsObter("TipoCaminho") & "") & """, "
    chamadoJSON = chamadoJSON & """categoria"": """ & EscapeJSON(rsObter("Categoria") & "") & """, "
    chamadoJSON = chamadoJSON & """avaliado"": """ & EscapeJSON(rsObter("Avaliado") & "") & """, "
    chamadoJSON = chamadoJSON & """prioridade"": """ & EscapeJSON(rsObter("Prioridade") & "") & """, "
    chamadoJSON = chamadoJSON & """departamento"": """ & EscapeJSON(rsObter("Departamento") & "") & """, "
    chamadoJSON = chamadoJSON & """centroCusto"": """ & EscapeJSON(rsObter("CentroCusto") & "") & """, "
    chamadoJSON = chamadoJSON & """tokenAcesso"": """ & EscapeJSON(rsObter("TokenAcesso") & "") & """, "
    chamadoJSON = chamadoJSON & """detalhamento"": """ & EscapeJSON(rsObter("Detalhamento") & "") & """, "
    chamadoJSON = chamadoJSON & """status"": """ & EscapeJSON(rsObter("Status") & "") & """, "
    chamadoJSON = chamadoJSON & """titulo"": """ & EscapeJSON(rsObter("Titulo") & "") & """, "
    chamadoJSON = chamadoJSON & """descricao"": """ & EscapeJSON(rsObter("Descricao") & "") & """, "
    chamadoJSON = chamadoJSON & """dataCriacao"": """ & EscapeJSON(rsObter("DataCriacao") & "") & """, "
    chamadoJSON = chamadoJSON & """dataUltimaAtualizacao"": """ & EscapeJSON(rsObter("DataUltimaAtualizacao") & "") & """, "
    chamadoJSON = chamadoJSON & """dataFinalizacao"": """ & EscapeJSON(rsObter("DataFinalizacao") & "") & """, "
    chamadoJSON = chamadoJSON & """dataPrevista"": """ & EscapeJSON(rsObter("DataPrevista") & "") & """, "
    chamadoJSON = chamadoJSON & """solicitante"": " & solicitanteJSON & ", "
    chamadoJSON = chamadoJSON & """mensagens"": " & mensagensJSON & ", "
    chamadoJSON = chamadoJSON & """anexos"": " & anexosJSON & ", "
    chamadoJSON = chamadoJSON & """historico"": " & historicoJSON & ""
    chamadoJSON = chamadoJSON & "}"

    usuarioJSON = usuarioJSON & "{"
    usuarioJSON = usuarioJSON & """nome"": """ & autor("nome") & """, "
    usuarioJSON = usuarioJSON & """tipo"": """ & autor("tipo") & """"
    usuarioJSON = usuarioJSON & "}"
    
    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """chamado"": " & chamadoJSON & ", "
    jsonObjeto = jsonObjeto & """usuario"": " & usuarioJSON & ""
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterChamado = "{""sucesso"": true, ""mensagem"": ""Chamado encontrado"", ""dados"": " & jsonObjeto & "}"
End Function

Function ObterDadosKPI()
    idAtendente = ConsultarAtendente(usuarioLogado("id"))

    SQL = SQL & "DECLARE @Agora         DATETIME2 = SYSDATETIME(); "
    SQL = SQL & "DECLARE @InicioHoje    DATETIME2 = CAST(SYSDATETIME() AS DATE); "    
    SQL = SQL & "DECLARE @QuatroHFuturo DATETIME2 = DATEADD(HOUR,  4, @Agora); "

    SQL = SQL & "SELECT"
    SQL = SQL & "    ( "
    SQL = SQL & "        SELECT COUNT(*)"
    SQL = SQL & "        FROM Chamados c"
    SQL = SQL & "        WHERE Status NOT IN ('FINALIZADO', 'CANCELADO') "
    SQL = SQL & "        AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "        AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"
    SQL = SQL & "    ) AS EmAndamento,"
    SQL = SQL & "    ("
    SQL = SQL & "        SELECT COUNT(*)"
    SQL = SQL & "        FROM Chamados c"
    SQL = SQL & "        WHERE Status = 'FINALIZADO'"
    SQL = SQL & "        AND c.DataFinalizacao >= @InicioHoje"
    SQL = SQL & "        AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "        AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"
    SQL = SQL & "    ) AS ResolvidosHoje, "
    SQL = SQL & "    ( "
    SQL = SQL & "        SELECT COUNT(*)"
    SQL = SQL & "        FROM Chamados c"
    SQL = SQL & "        WHERE Status NOT IN ('FINALIZADO', 'CANCELADO')"
    SQL = SQL & "        AND c.DataPrevista IS NOT NULL"
    SQL = SQL & "        AND (c.DataPrevista <= @QuatroHFuturo AND c.DataPrevista >= GETDATE())"
    SQL = SQL & "        AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "        AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"
    SQL = SQL & "    ) AS SlaVencendo,"
    SQL = SQL & "    ( "
    SQL = SQL & "        SELECT COUNT(*)"
    SQL = SQL & "        FROM Chamados c"
    SQL = SQL & "        WHERE Status NOT IN ('FINALIZADO', 'CANCELADO')"
    SQL = SQL & "        AND c.DataPrevista IS NOT NULL"
    SQL = SQL & "        AND c.DataPrevista < GETDATE()"
    SQL = SQL & "        AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "        AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"
    SQL = SQL & "    ) AS SlaVencido"

    On Error Resume Next
    SET rsDados = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        ObterDadosKPI = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de chamados: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    emAndamento = 0
    resolvidosHoje = 0
    slaVencendo = 0
    slaVencido = 0

    If Not rsDados.EOF Then
        emAndamento = rsDados("EmAndamento")
        resolvidosHoje = rsDados("ResolvidosHoje")
        slaVencendo = rsDados("SlaVencendo")
        slaVencido = rsDados("SlaVencido")
    End If

    jsonObjeto = jsonObjeto & "{"
    jsonObjeto = jsonObjeto & """emAndamento"": """ & EscapeJSON(emAndamento & "") & """, "
    jsonObjeto = jsonObjeto & """resolvidosHoje"": """ & EscapeJSON(resolvidosHoje & "") & """, "
    jsonObjeto = jsonObjeto & """slaVencendo"": """ & EscapeJSON(slaVencendo & "") & """, "
    jsonObjeto = jsonObjeto & """slaVencido"": """ & EscapeJSON(slaVencido & "") & """"
    jsonObjeto = jsonObjeto & "}" 

    rsDados.Close()
    SET rsDados = NOTHING
    
    ObterDadosKPI = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    codPlanta = Request("codPlanta")
    tipo = Request("tipo")
    nome = Request("nome")
    email = Request("email")
    telefone = Request("telefone")
    idDepartamento = Request("idDepartamento")
    idTipo = Request("idTipo")
    titulo = Request("titulo")
    descricao = Request("descricao")
    anexosJSON = Request("anexos")
    centroCusto = Request("centroCusto")
    detalhamento = Iif(Request("detalhamento") <> "", "'" & Replace(Request("detalhamento"), "'", "''") & "'", "NULL")
    idCategoria = Iif(Request("idCategoria") <> "", Request("idCategoria"), "NULL")

    Set dadosSolicitante = ObterOuCriarSolicitante(codPlanta, nome, email, telefone)
    If dadosSolicitante("idSolicitante") = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao obter solicitante: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    idPrioridade = ObterPrioridade(codPlanta, "MEDIA")
    If idPrioridade = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao obter prioridade: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "
    SQL = SQL & "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @DadosChamado TABLE (Id int, TokenAcesso VARCHAR(100)); "

    ' Gera registro de chamado            
    SQL = SQL & "INSERT INTO Chamados ("
    SQL = SQL & "   CodPlanta, IdSolicitante, IdPrioridade, IdDepartamento, IdTipo, IdCategoria, "
    SQL = SQL & "   Status, Titulo, Descricao, CentroCusto, Detalhamento, DataPrevista, "
    SQL = SQL & "   DataCriacao, DataUltimaAtualizacao, DataFinalizacao"
    SQL = SQL & ") OUTPUT INSERTED.Id, INSERTED.TokenAcesso INTO @DadosChamado VALUES ("
    SQL = SQL & "'" & codPlanta & "', "
    SQL = SQL & "'" & dadosSolicitante("idSolicitante") & "', "
    SQL = SQL & "'" & idPrioridade & "', "
    SQL = SQL & "'" & idDepartamento & "', "
    SQL = SQL & "'" & idTipo & "', "
    SQL = SQL & "" & idCategoria & ", "
    SQL = SQL & "'ABERTO', "
    SQL = SQL & "'" & titulo & "', "
    SQL = SQL & "'" & descricao & "', "
    SQL = SQL & "'" & centroCusto & "', "
    SQL = SQL & "" & detalhamento & ", "
    SQL = SQL & "'" & ObterPrazoSLA(idPrioridade) & "', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "NULL"
    SQL = SQL & "); "

    SQL = SQL & "SELECT Id, TokenAcesso FROM @DadosChamado; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
    
    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    idChamado = rsCriar("Id")
    tokenAcesso = rsCriar("TokenAcesso")
    
    rsCriar.Close()
    SET rsCriar = NOTHING

    If anexosJSON <> "" AND anexosJSON <> "[]" Then
        Call SalvarAnexos(anexosJSON, idChamado, "")
    End If

    Call EnviarEmailCriacaoParaSolicitante(titulo, nome, email, tokenAcesso)
    Call EnviarEmailCriacaoParaAtendentes(titulo, idDepartamento)
    Call RegistrarLog("criacao", idChamado, nome)
    Criar = "{""sucesso"": true, ""mensagem"": ""Chamado criado com sucesso!""}"
End Function

Function EnviarMensagem()  
    mensagem = Request("mensagem")
    ehNotaInterna = Request("ehNotaInterna")
    anexosJSON = Request("anexos")
    acaoAdicional = Request("acaoAdicional")

    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Esse chamado já foi finalizado""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "
    SQL = SQL & "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @DadosMensagem TABLE (Id int); "

    SQL = SQL & "INSERT INTO MensagensChamado ("
    SQL = SQL & "   IdChamado, IdAutor, TipoAutor, NomeAutor, NotaInterna, Descricao, DataCriacao"
    SQL = SQL & ") OUTPUT INSERTED.Id INTO @DadosMensagem VALUES ("
    SQL = SQL & "'" & autor("idChamado") & "', "
    SQL = SQL & "'" & autor("idAutor") & "', "
    SQL = SQL & "'" & autor("tipo") & "', "
    SQL = SQL & "'" & autor("nome") & "', "
    SQL = SQL & "'" & ehNotaInterna & "', "
    SQL = SQL & "'" & mensagem & "', "
    SQL = SQL & "GETDATE()"
    SQL = SQL & "); "
    
    If acaoAdicional = "finalizar" Then
        If dadosChamado("idAtendente") = "" Then
            EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Antes de finalizar um chamado é necessário atribuir um atendente.""}"
            Exit Function
        End If

        SQL = SQL & "INSERT INTO HistoricoChamado ("
        SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
        SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
        SQL = SQL & ") VALUES ("
        SQL = SQL & "   '" & autor("idChamado") & "', "
        SQL = SQL & "   '" & autor("idAutor") & "', "
        SQL = SQL & "   '" & autor("nome") & "', "
        SQL = SQL & "   'FINALIZADO', "
        SQL = SQL & "   '" & dadosChamado("status") & "', "
        SQL = SQL & "   'FINALIZADO', "
        SQL = SQL & "   'Chamado foi marcado como finalizado', "
        SQL = SQL & "   GETDATE(), "
        SQL = SQL & "   0"
        SQL = SQL & "); "

        SQL = SQL & "UPDATE Chamados SET "
        SQL = SQL & "   Status = 'FINALIZADO', "
        SQL = SQL & "   DataFinalizacao = GETDATE(), "
        SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
        SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "
    ElseIf acaoAdicional = "aguardandoSolicitante" Then
        SQL = SQL & "INSERT INTO HistoricoChamado ("
        SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
        SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
        SQL = SQL & ") VALUES ("
        SQL = SQL & "   '" & autor("idChamado") & "', "
        SQL = SQL & "   '" & autor("idAutor") & "', "
        SQL = SQL & "   '" & autor("nome") & "', "
        SQL = SQL & "   'STATUS_ALTERADO', "
        SQL = SQL & "   '" & dadosChamado("status") & "', "
        SQL = SQL & "   'AGUARDANDO_SOLICITANTE', "
        SQL = SQL & "   'Chamado foi marcado como aguardando solicitante', "
        SQL = SQL & "   GETDATE(), "
        SQL = SQL & "   0"
        SQL = SQL & "); "

        SQL = SQL & "UPDATE Chamados SET "
        SQL = SQL & "   Status = 'AGUARDANDO_SOLICITANTE', "
        SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
        SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "
    ElseIf acaoAdicional = "manterAberto" Then
        SQL = SQL & "UPDATE Chamados SET "
        SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
        SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "
    End If

    SQL = SQL & "SELECT Id FROM @DadosMensagem; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"
    
    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        EnviarMensagem = "{""sucesso"": false, ""mensagem"": ""Erro ao enviar mensagem: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    idMensagem = rsCriar("Id")

    rsCriar.Close()
    SET rsCriar = NOTHING

    If anexosJSON <> "" AND anexosJSON <> "[]" Then
        Call SalvarAnexos(anexosJSON, "", idMensagem)
    End If

    If autor("tipo") = "ATENDENTE" AND ehNotaInterna = "0" Then
        Call EnviarEmailMensagemAtendente(autor("idChamado"), autor("nome"), mensagem)
    ElseIf autor("tipo") = "SOLICITANTE" Then
        Call EnviarEmailMensagemSolicitante(autor("idChamado"), autor("nome"), mensagem)
    End If

    Call RegistrarLog("envio_mensagem", autor("idChamado"), autor("nome"))
    EnviarMensagem = "{""sucesso"": true, ""mensagem"": ""Mensagem enviada com sucesso!""}"
End Function

Function FinalizarChamado()   
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Esse chamado já foi finalizado""}"
        Exit Function
    End If

    If dadosChamado("idAtendente") = "" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Antes de fechar o chamado é necessário atribuir um atendente""}"
        Exit Function
    End If

    If dadosChamado("idDepartamento") = "" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Antes de fechar o chamado é necessário atribuir um departamento""}"
        Exit Function
    End If

    If dadosChamado("idCategoria") = "" Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Antes de fechar o chamado é necessário atribuir uma categoria""}"
        Exit Function
    End If

    statusAnterior = dadosChamado("status")
    novoStatus = "FINALIZADO"

    SQL = "BEGIN TRANSACTION; "
 
    SQL = SQL & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'FINALIZADO', "
    SQL = SQL & "   '" & statusAnterior & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   'Chamado foi marcado como finalizado', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   Status = '" & novoStatus & "', "
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE(), "
    SQL = SQL & "   DataFinalizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsFinalizar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsFinalizar) Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao finalizar chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call EnviarEmailFinalizacao(autor("idChamado"))
    Call RegistrarLog("finalizacao", autor("idChamado"), autor("nome"))
    FinalizarChamado = "{""sucesso"": true, ""mensagem"": ""Chamado finalizado com sucesso!""}"
End Function

Function ReabrirChamado()   
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        ReabrirChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        ReabrirChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        ReabrirChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") <> "CANCELADO" AND  dadosChamado("status") <> "FINALIZADO" Then
        ReabrirChamado = "{""sucesso"": false, ""mensagem"": ""Esse chamado já está aberto""}"
        Exit Function
    End If

    statusAnterior = dadosChamado("status")
    novoStatus = "EM_ANALISE"

    SQL = "BEGIN TRANSACTION; "
 
    SQL = sql & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'REABERTO', "
    SQL = SQL & "   '" & statusAnterior & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   'Chamado foi reaberto', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   Status = '" & novoStatus & "', "
    SQL = SQL & "   DataFinalizacao = NULL ,"
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsReabrir = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsReabrir) Then
        ReabrirChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao reabrir chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call EnviarEmailReabertura(autor("idChamado"))
    Call RegistrarLog("reabertura", autor("idChamado"), autor("nome"))
    ReabrirChamado = "{""sucesso"": true, ""mensagem"": ""Chamado reaberto com sucesso!""}"
End Function

Function AtribuirAtendente()   
    idAtendente = Request("idAtendente")

    If idAtendente = "" Then
        AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Atendente inválido""}"
        Exit Function
    End If

    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    'If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        'AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        'Exit Function
    'End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Esse chamado está finalizado""}"
        Exit Function
    End If

    SQL = "SELECT Nome FROM AtendentesChamado WHERE Id = '" & idAtendente & "'; "
    
    On Error Resume Next
    SET rsAtendente = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAtendente) Then
        FinalizarChamado = "{""sucesso"": false, ""mensagem"": ""Erro atribuir atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    nomeAtendenteAnterior = Iif(dadosChamado("nomeAtendente") <> "", dadosChamado("nomeAtendente"), "Nenhum Atendente")
    nomeNovoAtendente = rsAtendente("Nome")

    SQL = "BEGIN TRANSACTION; "
 
    SQL = SQL & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'ATRIBUIDO', "
    SQL = SQL & "   '" & nomeAtendenteAnterior & "', "
    SQL = SQL & "   '" & nomeNovoAtendente & "', "
    SQL = SQL & "   'Atendente alterado de " & nomeAtendenteAnterior & " para " & nomeNovoAtendente & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   IdAtendente = '" & idAtendente & "', " 
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsAtribuir = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAtribuir) Then
        AtribuirAtendente = "{""sucesso"": false, ""mensagem"": ""Erro ao atribuir atendente: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call RegistrarLog("atribuicao", autor("idChamado"), autor("nome"))
    AtribuirAtendente = "{""sucesso"": true, ""mensagem"": ""Atendente atribuído com sucesso!""}"
End Function

Function AlterarStatus()       
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Esse chamado está finalizado""}"
        Exit Function
    End If

    statusAnterior = dadosChamado("status")
    novoStatus = Request("status")

    SQL = "BEGIN TRANSACTION; "
 
    SQL = sql & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'STATUS_ALTERADO', "
    SQL = SQL & "   '" & statusAnterior & "', "
    SQL = SQL & "   '" & novoStatus & "', "
    SQL = SQL & "   'Status alterado de " & statusAnterior & " para " & novoStatus & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   Status = '" & novoStatus & "', "
    SQL = SQL & "   DataFinalizacao = NULL, "
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsAlterar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAlterar) Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar status: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("alteracao_status", autor("idChamado"), autor("nome"))
    AlterarStatus = "{""sucesso"": true, ""mensagem"": ""Status alterado com sucesso!""}"
End Function

Function AlterarPrioridade()       
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Esse chamado está finalizado""}"
        Exit Function
    End If

    nivelPrioridadeAnterior = dadosChamado("prioridade")

    ' Obtem dados da nova prioridade   
    nivelPrioridade = Request("prioridade")
    SQL = "SELECT p.Id, p.Nivel, p.PrazoSLA FROM Chamados c "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.CodPlanta = c.CodPlanta AND p.Nivel = '" & nivelPrioridade & "' "
    SQL = SQL & "WHERE c.Id = '" & autor("idChamado") & "' "

    On Error Resume Next
    SET rsNovaPrioridade = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsNovaPrioridade) Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Erro ao reabrir chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    idNovaPrioridade = rsNovaPrioridade("Id")
    nivelNovaPrioridade = rsNovaPrioridade("Nivel")
    prazoSLA = rsNovaPrioridade("PrazoSLA")

    rsNovaPrioridade.Close()
    SET rsNovaPrioridade = NOTHING

    If nivelPrioridadeAnterior = nivelNovaPrioridade Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Prioridade atual e nova prioridade são iguais""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "
 
    SQL = SQL & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'PRIORIDADE_ALTERADA', "
    SQL = SQL & "   '" & nivelPrioridadeAnterior & "', "
    SQL = SQL & "   '" & nivelNovaPrioridade & "', "
    SQL = SQL & "   'Prioridade alterada de " & nivelPrioridadeAnterior & " para " & nivelNovaPrioridade & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   IdPrioridade = '" & idNovaPrioridade & "', "
    SQL = SQL & "   DataPrevista = DATEADD(hour, " & prazoSLA & ", DataCriacao), "
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsAlterar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAlterar) Then
        AlterarPrioridade = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar prioridade: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call RegistrarLog("alteracao_prioridade", autor("idChamado"), autor("nome"))
    AlterarPrioridade = "{""sucesso"": true, ""mensagem"": ""Prioridade alterada com sucesso!""}"
End Function

Function AlterarCategoria()       
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AlterarCategoria = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AlterarCategoria = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        AlterarCategoria = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        AlterarCategoria = "{""sucesso"": false, ""mensagem"": ""Esse chamado está finalizado""}"
        Exit Function
    End If

    idCategoriaAnterior = dadosChamado("idCategoria")
    nomeCategoriaAnterior = dadosChamado("categoria")

    idCategoriaNova = Request("idCategoria")
    nomeCategoriaNova = Request("nomeCategoria")

    SQL = "BEGIN TRANSACTION; "
 
    SQL = sql & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'CATEGORIA_ALTERADA', "
    SQL = SQL & "   '" & idCategoriaAnterior & "', "
    SQL = SQL & "   '" & idCategoriaNova & "', "
    SQL = SQL & "   'Categoria alterada de " & nomeCategoriaAnterior & " para " & nomeCategoriaNova & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   IdCategoria = '" & idCategoriaNova & "', "
    SQL = SQL & "   DataFinalizacao = NULL, "
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsAlterar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAlterar) Then
        AlterarCategoria = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar status: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("alteracao_categoria", autor("idChamado"), autor("nome"))
    AlterarCategoria = "{""sucesso"": true, ""mensagem"": ""Categoria alterada com sucesso!""}"
End Function

Function AlterarDepartamento()     
    idNovoDepartamento = Request("idDepartamento")
    If idNovoDepartamento = "" Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Departamento inválido""}"
        Exit Function
    End If

    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") = "ATENDENTE" AND Not AtendentePodeAtender(autor("idAutor"), autor("idChamado")) Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") = "CANCELADO" OR dadosChamado("status") = "FINALIZADO" Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Esse chamado está finalizado""}"
        Exit Function
    End If

    idDepartamentoAnterior = dadosChamado("idDepartamento")
    idAtendenteAnterior = dadosChamado("idAtendente")
    nomeDepartamentoAnterior = dadosChamado("departamento")

    ' Obtem dados do novo departamento
    SQL = "SELECT Id, Nome From DepartamentosChamado WHERE Id = " & idNovoDepartamento

    On Error Resume Next
    SET rsNovoDepartamento = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsNovoDepartamento) Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    nomeNovoDepartamento = rsNovoDepartamento("Nome")

    rsNovoDepartamento.Close()
    SET rsNovoDepartamento = NOTHING

    If CStr(idDepartamentoAnterior) = CStr(idNovoDepartamento) Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Departamento atual e novo departamento são iguais""}"
        Exit Function
    End If

    ' Precisamos verificar se o atendente atual possui acesso ao novo departamento
    ' Se não possuir precisamos desassociar ele do chamado 
    deveDesassociarAtendente = False
    If Not isNull(idAtendenteAnterior) AND Not AtendentePossuiDepartamento(idNovoDepartamento, autor("idAutor")) Then
        deveDesassociarAtendente = True
    End If

    SQL = "BEGIN TRANSACTION; "
 
    SQL = sql & "INSERT INTO HistoricoChamado ("
    SQL = SQL & "   IdChamado, IdAtendente, NomeAtendente, Acao, "
    SQL = SQL & "   DeValor, ParaValor, Descricao, DataCriacao, FlagInativo "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & autor("idAutor") & "', "
    SQL = SQL & "   '" & autor("nome") & "', "
    SQL = SQL & "   'DEPARTAMENTO_ALTERADO', "
    SQL = SQL & "   '" & nomeDepartamentoAnterior & "', "
    SQL = SQL & "   '" & nomeNovoDepartamento & "', "
    SQL = SQL & "   'Departamento alterado de " & nomeDepartamentoAnterior & " para " & nomeNovoDepartamento & "', "
    SQL = SQL & "   GETDATE(), "
    SQL = SQL & "   0"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE Chamados SET "
    SQL = SQL & "   IdDepartamento = '" & idNovoDepartamento & "', "
    SQL = SQL & "   IdCategoria = NULL, "
    SQL = SQL & "   DataUltimaAtualizacao = GETDATE() "
    SQL = SQL &     Iif(deveDesassociarAtendente, ", IdAtendente = NULL ", " ")
    SQL = SQL & "WHERE Id = '" & autor("idChamado") & "'; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION"

    On Error Resume Next
    SET rsAlterar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAlterar) Then
        AlterarDepartamento = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar departamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call RegistrarLog("alteracao_departamento", autor("idChamado"), autor("nome"))
    AlterarDepartamento = "{""sucesso"": true, ""mensagem"": ""Departamento alterado com sucesso!""}"
End Function

Function AvaliarChamado()   
    Set autor = IdentificarAutor()

    If autor("tipo") = "anonimo" OR autor("idChamado") = "" Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    ' Token foi passado mas é inválido/expirado
    If autor("erro") <> "" Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If
    
    If autor("tipo") <> "SOLICITANTE" Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Não autorizado""}"
        Exit Function
    End If

    Set dadosChamado = ObterDadosChamado(autor("idChamado"))
    If dadosChamado("status") <> "FINALIZADO" Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Esse chamado ainda não foi finalizado""}"
        Exit Function
    End If

    jaExisteAvaliacao = False
    SQL = "SELECT Id FROM AvaliacaoChamado WHERE IdChamado = '" & autor("idChamado") & "'"

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao avaliar chamado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.EOF Then
        jaExisteAvaliacao = True
    End If

    If jaExisteAvaliacao Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Chamado ja avaliado""}"
        Exit Function
    End If

    notaAtendimento = Request("notaAtendimento")
    notaRecomendacao = Request("notaRecomendacao")
    motivos = Request("motivos")
    comentario = Request("comentario")
 
    SQL = SQL & "INSERT INTO AvaliacaoChamado ("
    SQL = SQL & "   IdChamado, NotaAtendimento, NotaRecomendacao, Motivos, Comentario, DataCriacao"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "   '" & autor("idChamado") & "', "
    SQL = SQL & "   '" & notaAtendimento & "', "
    SQL = SQL & "   '" & notaRecomendacao & "', "
    SQL = SQL & "   '" & motivos & "', "
    SQL = SQL & "   '" & comentario & "', "
    SQL = SQL & "   GETDATE()"
    SQL = SQL & "); "

    On Error Resume Next
    SET rsAvaliar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAvaliar) Then
        AvaliarChamado = "{""sucesso"": false, ""mensagem"": ""Erro ao finalizar avaliar: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    Call RegistrarLog("avaliacao", autor("idChamado"), autor("nome"))
    AvaliarChamado = "{""sucesso"": true, ""mensagem"": ""Chamado finalizado com sucesso!""}"
End Function


'----------------------------------------------
'                MÉTODOS AUXILIARES
'----------------------------------------------

Function ObterDadosChamado(idChamado)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados.Add "idChamado", idChamado
    dados.Add "idSolicitante", ""
    dados.Add "idAtendente", ""
    dados.Add "idPrioridade", ""
    dados.Add "idDepartamento", ""
    dados.Add "idCategoria", ""
    dados.Add "nomeSolicitante", ""
    dados.Add "emailSolicitante", ""
    dados.Add "nomeAtendente", ""
    dados.Add "categoria", ""
    dados.Add "prioridade", ""
    dados.Add "departamento", ""
    dados.Add "tokenAcesso", ""
    dados.Add "status", ""
    dados.Add "titulo", ""
    dados.Add "descricao", ""
    dados.Add "dataCriacao", ""
    dados.Add "dataFinalizacao", ""
    dados.Add "dataPrevista", ""
      
    SQL = "SELECT "
    SQL = SQL & "    c.Id, "
    SQL = SQL & "    c.IdSolicitante, "
    SQL = SQL & "    c.IdAtendente, "
    SQL = SQL & "    c.IdPrioridade, "
    SQL = SQL & "    c.IdDepartamento, "
    SQL = SQL & "    c.IdCategoria, "
    SQL = SQL & "    s.Nome as NomeSolicitante, "
    SQL = sql & "    s.Email as EmailSolicitante, "
    SQL = SQL & "    a.Nome as NomeAtendente, "
    SQL = SQL & "    ca.Nome as Categoria, "
    SQL = SQL & "    p.Nivel as Prioridade, "
    SQL = SQL & "    d.Nome as Departamento, "
    SQL = SQL & "    c.TokenAcesso, "
    SQL = SQL & "    c.Status, "
    SQL = SQL & "    c.Titulo, "
    SQL = SQL & "    c.Descricao, "
    SQL = SQL & "FORMAT(c.DataCriacao , 'MM/dd/yyyy HH:mm:ss') AS DataCriacao, "
    SQL = SQL & "FORMAT(c.DataFinalizacao , 'MM/dd/yyyy HH:mm:ss') AS DataFinalizacao, "
    SQL = SQL & "FORMAT(c.DataPrevista , 'MM/dd/yyyy HH:mm:ss') AS DataPrevista "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "LEFT JOIN AtendentesChamado a on a.Id = c.IdAtendente "
    SQL = SQL & "LEFT JOIN CategoriasChamado ca ON ca.Id = c.IdCategoria "
    SQL = SQL & "LEFT JOIN PrioridadesChamado p on p.Id = c.IdPrioridade "
    SQL = SQL & "LEFT JOIN DepartamentosChamado d on d.Id = c.IdDepartamento "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.Id = " & idChamado & " "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "
    
    On Error Resume Next
    SET rsDados = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        Set ObterDadosChamado = dados
        Exit Function
    End If
    On Error Goto 0

    If Not rsDados.EOF Then
        dados("idChamado") = rsDados("Id") & ""
        dados("idSolicitante") = rsDados("IdSolicitante") & ""
        dados("idAtendente") = rsDados("IdAtendente") & ""
        dados("idPrioridade") = rsDados("IdPrioridade") & ""
        dados("idDepartamento") = rsDados("IdDepartamento") & ""
        dados("idCategoria") = rsDados("IdCategoria") & ""
        dados("nomeSolicitante") = rsDados("NomeSolicitante") & ""
        dados("emailSolicitante") = rsDados("EmailSolicitante") & ""
        dados("nomeAtendente") = rsDados("NomeAtendente") & ""
        dados("categoria") = rsDados("Categoria") & ""
        dados("prioridade") = rsDados("Prioridade") & ""
        dados("departamento") = rsDados("Departamento") & ""
        dados("tokenAcesso") = rsDados("TokenAcesso") & ""
        dados("status") = rsDados("Status") & ""
        dados("titulo") = rsDados("Titulo") & ""
        dados("descricao") = rsDados("Descricao") & ""
        dados("dataCriacao") = rsDados("DataCriacao") & ""
        dados("dataFinalizacao") = rsDados("DataFinalizacao") & ""
        dados("dataPrevista") = rsDados("DataPrevista") & ""
    End If

    Set ObterDadosChamado = dados
End Function

' Determina quem está acessando o chamado:
'   - Atendente logado: variável global usuariologado(id) preenchida)
'   - Solicitante autenticado via Chamados.TokenAcesso na URL
'   - Anônimo (nenhum dos dois)
Function IdentificarAutor()
    Set ctx = Server.CreateObject("Scripting.Dictionary")
    ctx.Add "tipo",            "anonimo"
    ctx.Add "idAutor",         "0"
    ctx.Add "nome",            ""
    ctx.Add "idChamado",       "0"
    ctx.Add "podeNotaInterna", False
    ctx.Add "erro",            ""

    ' Solicitante via token 
    token = Trim("" & Request("token"))

    If token <> "" Then
        Set dadosToken = ValidarTokenChamado(token)

        If dadosToken("ok") Then
            ctx("tipo")            = "SOLICITANTE"
            ctx("idAutor")         = dadosToken("idSolicitante")
            ctx("nome")            = dadosToken("nomeSolicitante")
            ctx("idChamado")       = dadosToken("idChamado")
            ctx("podeNotaInterna") = False
        Else
            ctx("erro") = dadosToken("erro")
        End If

        Set IdentificarAutor = ctx
        Exit Function
    End If

    ' Atendente Logado
    If usuarioLogado("id") <> "" Then
        idAtendente = ConsultarAtendente(usuarioLogado("id"))

        If idAtendente <> "" Then
            ctx("tipo")            = "ATENDENTE"
            ctx("idAutor")         = idAtendente & ""
            ctx("nome")            = usuarioLogado("nome")
            ctx("idChamado")       = Request("idChamado")
            ctx("podeNotaInterna") = True

            Set IdentificarAutor = ctx
            Exit Function
        End If
    End If

    Set IdentificarAutor = ctx
End Function

Function AtendentePossuiDepartamento(idAtendente, idDepartamento)
    temDepartamento = False
    
    SQL = "SELECT TOP 1 ISNULL(a.Id, 0) as Id "
    SQL = SQL & "FROM AtendentesChamado a " 
    SQL = SQL & "WHERE Id = '" & idAtendente & "' "
    SQL = SQL & "AND a.FlagInativo = 0 "
    SQL = SQL & "AND EXISTS ( "
    SQL = SQL & "    SELECT TOP 1 a.Id "
    SQL = SQL & "    FROM AtendenteDepartamentoChamado ad "
    SQL = SQL & "    WHERE ad.IdAtendente = a.Id "
    SQL = SQL & "    AND ad.IdDepartamento = " & idDepartamento
    SQL = SQL & ") "

    On Error Resume Next
    SET rsVerificacao = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsVerificacao) Then
        AtendentePossuiDepartamento = False
        Exit Function
    End If
    On Error Goto 0

    If Not rsVerificacao.EOF Then
        If CStr(rsVerificacao("Id")) <> "0" Then
            temDepartamento = True
        End If
        
        rsVerificacao.Close()
        SET rsVerificacao = NOTHING
    End If

    AtendentePossuiDepartamento = temDepartamento
End Function

Function ConsultarAtendente(idUsuario)
    idAtendente = ""

    SQL = "SELECT TOP 1 Id FROM AtendentesChamado"
    SQL = SQL & " WHERE IdUsuario = '" & idUsuario & "'"
    SQL = SQL & " AND FlagInativo = 0"

    On Error Resume Next
    SET rsAtendente = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsAtendente) Then
        ConsultarAtendente = idAtendente
        Exit Function
    End If
    On Error Goto 0

    If Not rsAtendente.EOF Then
        idAtendente = rsAtendente("Id")
    End If

    rsAtendente.Close()
    SET rsAtendente = NOTHING

    ConsultarAtendente = idAtendente
End Function

Function ValidarTokenChamado(token)
    Set retorno = Server.CreateObject("Scripting.Dictionary")
    retorno.Add "ok",              False
    retorno.Add "idChamado",       0
    retorno.Add "idSolicitante",   0
    retorno.Add "nomeSolicitante", ""
    retorno.Add "erro",            ""

    SQL = "SELECT c.Id, c.IdSolicitante, c.Status, s.Nome as NomeSolicitante "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "WHERE c.TokenAcesso = '" & token & "'"

    On Error Resume Next
    SET rsValidar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsValidar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao validar token: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsValidar.EOF Then
        retorno("erro") = "Token não encontrado"
    Else
        retorno("ok") = True
        retorno("idChamado") = rsValidar("Id")
        retorno("idSolicitante") = rsValidar("IdSolicitante")
        retorno("nomeSolicitante") = rsValidar("NomeSolicitante")
    End If

    rsValidar.Close()
    SET rsValidar = NOTHING

    Set ValidarTokenChamado = retorno
End Function


Function AtendentePodeAtender(idAtendente, idChamado)
    podeAntender = False
    
    SQL = SQL & "SELECT TOP 1 ISNULL(c.IdAtendente, 0) IdAtendente FROM Chamados c "
    SQL = SQL & "LEFT JOIN AtendentesChamado a on a.Id = c.IdAtendente "
    SQL = SQL & "WHERE EXISTS ( "
    SQL = SQL & "    SELECT TOP 1 ad.IdDepartamento "
    SQL = SQL & "    FROM AtendenteDepartamentoChamado ad "
    SQL = SQL & "    WHERE ad.IdAtendente = '" & idAtendente & "' "
    SQL = SQL & "    AND ad.IdDepartamento = c.IdDepartamento "
    SQL = SQL & ") "
    SQL = SQL & "AND c.Id = '" & idChamado & "'" 

    On Error Resume Next
    SET rsVerificacao = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsVerificacao) Then
        AtendentePodeAtender = False
        Exit Function
    End If
    On Error Goto 0

    If Not rsVerificacao.EOF Then
        If CStr(rsVerificacao("IdAtendente")) = "0" OR CStr(rsVerificacao("IdAtendente")) = CStr(idAtendente) Then
            podeAntender = True
        End If
        
        rsVerificacao.Close()
        SET rsVerificacao = NOTHING
    End If

    AtendentePodeAtender = podeAntender
End Function

Sub SalvarAnexos(anexosJSON, idChamado, idMensagem)
    idChamado = Iif(idChamado = "", "NULL", idChamado)
    idMensagem = Iif(idMensagem = "", "NULL", idMensagem)

    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    anexosArray = Split(anexosJSON, "},{")
    
    valoresAnexos = ""
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)

        nomeAnexo = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")

        If i > 0 Then 
            valoresAnexos = valoresAnexos & ","
        End If
        
        valoresAnexos = valoresAnexos & "("
        valoresAnexos = valoresAnexos & "" & idChamado & ", "
        valoresAnexos = valoresAnexos & "" & idMensagem & ", "
        valoresAnexos = valoresAnexos & "'" & nomeAnexo & "', "
        valoresAnexos = valoresAnexos & "'" & caminho & "', "
        valoresAnexos = valoresAnexos & "GETDATE()"
        valoresAnexos = valoresAnexos & ")"
    Next

    SQLAnexos = "INSERT INTO AnexosChamado ( "
    SQLAnexos = SQLAnexos & "  IdChamado, IdMensagem, Nome, Caminho, DataCriacao "
    SQLAnexos = SQLAnexos & ") VALUES " 
    SQLAnexos = SQLAnexos & valoresAnexos & " "

    On Error Resume Next
    SET rsAnexos = Conecta("PIR-CORPORATIVO", SQLAnexos)
    
    If Err.Number <> 0 OR Not IsObject(rsAnexos) Then
        Exit Sub
    End If
    On Error Goto 0
End Sub

Function ObterPrazoSLA(idPrioridade)
    prazoHoras = 24

    SQL = "SELECT PrazoSLA FROM PrioridadesChamado WHERE Id = " & idPrioridade

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        ObterPrazoSLA = DateAdd("h", prazoHoras, Now())
        Exit Function
    End If

    If Not rsExiste.EOF Then
       prazoHoras = rsExiste("PrazoSLA")
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    On Error Goto 0

    data = DateAdd("h", prazoHoras, Now())
    dataFormatada = Year(data) & "-" & _
                Right("0" & Month(data),  2) & "-" & _
                Right("0" & Day(data),    2) & " " & _
                Right("0" & Hour(data),   2) & ":" & _
                Right("0" & Minute(data), 2) & ":" & _
                Right("0" & Second(data), 2)

    ObterPrazoSLA = dataFormatada
End Function

' Verificamos se solicitante já existe, se existir retorna id
' Se não existir cria solicitante e retorna id do registro criado
Function ObterOuCriarSolicitante(codPlanta, nome, email, telefone)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados.Add "idSolicitante", ""
    dados.Add "idUsuario", Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
    dados.Add "codPlanta", codPlanta
    dados.Add "nome", nome
    dados.Add "email", email
    dados.Add "telefone", telefone
   
    SQL = "SELECT Id, CodPlanta, IdUsuario, Nome, Email, Telefone FROM SolicitantesChamado WHERE Email = '" & email & "'"

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        Set ObterOuCriarSolicitante = dados
        Exit Function
    End If
    On Error Goto 0

    ' Se encontrou retorna id e finaliza fluxo
    If Not rsExiste.EOF Then
        dados("idSolicitante") = rsExiste("Id")
        dados("idUsuario") = rsExiste("IdUsuario")
        dados("codPlanta") = rsExiste("CodPlanta")
        dados("nome") = rsExiste("Nome")
        dados("email") = rsExiste("Email")
        dados("telefone") = rsExiste("Telefone")
        
        rsExiste.Close()
        SET rsExiste = NOTHING
        
        Set ObterOuCriarSolicitante = dados
        Exit Function
    End If

    ' Caso continue cria registro de solicitante
    SQL = "INSERT INTO SolicitantesChamado (IdUsuario, CodPlanta, Nome, Email, Telefone) "
    SQL = SQL & "OUTPUT INSERTED.Id, INSERTED.IdUsuario, INSERTED.CodPlanta, INSERTED.Nome, INSERTED.Email, INSERTED.Telefone "
    SQL = SQL & "VALUES ( "
    SQL = SQL & "" & dados("idUsuario") & ", "
    SQL = SQL & "'" & dados("codPlanta") & "', "
    SQL = SQL & "'" & dados("nome") & "', "
    SQL = SQL & "'" & dados("email") & "', "
    SQL = SQL & "'" & dados("telefone") & "' "
    SQL = SQL & ")"

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Set ObterOuCriarSolicitante = dados
        Exit Function
    End If
    On Error Goto 0

    dados("idSolicitante") = rsCriar("Id")
    dados("idUsuario") = rsCriar("IdUsuario")
    dados("codPlanta") = rsCriar("CodPlanta")
    dados("nome") = rsCriar("Nome")
    dados("email") = rsCriar("Email")
    dados("telefone") = rsCriar("Telefone")
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    
    Call RegistrarLog("criacao_solicitante", "", "")
    Set ObterOuCriarSolicitante = dados
End Function

Function ObterPrioridade(codPlanta, nivel)
    idPrioridade = ""
 
    ' Caso continue cria registro de solicitante
    SQL = "SELECT TOP 1 Id FROM PrioridadesChamado "
    SQL = SQL & "WHERE CodPlanta = '" & codPlanta & "' "
    SQL = SQL & "AND Nivel = '" & nivel & "' "
    SQL = SQL & "AND FlagInativo = 0"

    On Error Resume Next
    SET rsPrioridade = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsPrioridade) Then
        ObterPrioridade = idPrioridade
        Exit Function
    End If
    On Error Goto 0

    idPrioridade = rsPrioridade("Id")
    
    rsPrioridade.Close()
    SET rsPrioridade = NOTHING
    
    ObterPrioridade = idPrioridade
End Function

Function ObterTipo(codTipo)
    idTipo = ""
 
    ' Caso continue cria registro de solicitante
    SQL = "SELECT TOP 1 Id FROM TiposChamado "
    SQL = SQL & "WHERE Codigo = '" & codTipo & "' "
    SQL = SQL & "AND FlagInativo = 0"

    On Error Resume Next
    SET rsTipo = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTipo) Then
        ObterTipo = idTipo
        Exit Function
    End If
    On Error Goto 0

    idTipo = rsTipo("Id")
    
    rsTipo.Close()
    SET rsTipo = NOTHING
    
    ObterTipo = idTipo
End Function


Sub EnviarEmailCriacaoParaSolicitante(titulo, nomeSolicitante, emailSolicitante, tokenAcesso)
    url = "http://br-pir1-m012/corporativo/chamados/Chamado/ChamadoDetalhes.asp?token=" & tokenAcesso

    assunto = "RE: " & titulo

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
    mensagem = mensagem & "<p>Olá, " & nomeSolicitante & "!</p>"
    mensagem = mensagem & "<p>Recebemos seu chamado e ele já está na fila de atendimento.</p>"
    mensagem = mensagem & "<p><strong>Assunto:</strong> " & titulo & "</p>"
    mensagem = mensagem & "<div class=""info-item"">Para acompanhar o andamento do seu chamado, acesse:</div>"
    mensagem = mensagem & "<div class=""info-item""><a href=""" & url & """>" & url & "</a></div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailSolicitante <> "" Then
        Call EnviaEmail(emailSolicitante, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailCriacaoParaAtendentes(titulo, idDepartamento)
    SQL = "SELECT Id, Email "
    SQL = SQL & "FROM AtendentesChamado a " 
    SQL = SQL & "WHERE a.FlagInativo = 0 "
    SQL = SQL & "AND EXISTS ( "
    SQL = SQL & "    SELECT TOP 1 a.Id "
    SQL = SQL & "    FROM AtendenteDepartamentoChamado ad "
    SQL = SQL & "    WHERE ad.IdAtendente = a.Id "
    SQL = SQL & "    AND ad.IdDepartamento = '" & idDepartamento & "' "
    SQL = SQL & ") "

    On Error Resume Next
    SET rsAtendentes = Conecta("PIR-CORPORATIVO", SQL)
    On Error Goto 0

    listaEmails = ""
    If Not rsAtendentes.EOF Then
        Do While Not rsAtendentes.EOF
            If rsAtendentes("Email") <> "" Then
                listaEmails = listaEmails & rsAtendentes("Email") & "; "
            End If
            rsAtendentes.MoveNext
        Loop
    End If

    assunto = "RE: " & titulo

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
    mensagem = mensagem & "<p>Olá, </p>"
    mensagem = mensagem & "<p>Um novo chamado foi criado e está aguardando atendimento.</p>"
    mensagem = mensagem & "<p><strong>Assunto:</strong> " & titulo & "</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If listaEmails <> "" Then
        Call EnviaEmail(listaEmails, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

' Notifica solicitante da mensagem do atendente
Sub EnviarEmailMensagemAtendente(idChamado, nomeAtendente, mensagemUsuario)
    SQL = "SELECT c.Titulo, s.Nome as NomeSolicitante, s.Email as EmailSolicitante, c.TokenAcesso "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "WHERE c.Id = '" & idChamado & "' "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "

    On Error Resume Next
    SET rsChamado = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsChamado) Then
        Exit Sub
    End If
    On Error Goto 0

    titulo = rsChamado("Titulo")
    nomeSolicitante = rsChamado("NomeSolicitante")
    emailSolicitante = rsChamado("EmailSolicitante")
    tokenAcesso = rsChamado("TokenAcesso")

    rsChamado.Close()
    SET rsChamado = NOTHING

    url = "http://br-pir1-m012/corporativo/chamados/Chamado/ChamadoDetalhes.asp?token=" & tokenAcesso

    assunto = "RE: " & titulo

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
    mensagem = mensagem & "<p>Olá, " & nomeSolicitante & "!</p>"
    mensagem = mensagem & "<p>" & nomeAtendente & ", respondeu seu chamado.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Assunto:</strong> " & titulo & "</div>"
    mensagem = mensagem & "<div class=""info-item""> <i>" & mensagemUsuario & "</i></div>"
    mensagem = mensagem & "<div style=""margin-top: 20px; class=""info-item"">Para responder ou ver o histórico completo, acesse:</div>"
    mensagem = mensagem & "<div class=""info-item""><a href=""" & url & """>" & url & "</a></div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailSolicitante <> "" Then
        Call EnviaEmail(emailSolicitante, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

' Notifica atendente da mensagem do solicitante
Sub EnviarEmailMensagemSolicitante(idChamado, nomeSolicitante, mensagemUsuario)
    SQL = "SELECT c.Titulo, a.Nome as NomeAtendente, a.Email as EmailAtendente, c.TokenAcesso "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN AtendentesChamado a ON a.Id = c.IdAtendente "
    SQL = SQL & "WHERE c.Id = '" & idChamado & "' "
    SQL = SQL & "AND c.FlagInativo = 0 "

    On Error Resume Next
    SET rsChamado = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsChamado) Then
        Exit Sub
    End If
    On Error Goto 0

    titulo = rsChamado("Titulo")
    nomeAtendente = rsChamado("NomeAtendente")
    emailAtendente = rsChamado("EmailAtendente")
    tokenAcesso = rsChamado("TokenAcesso")

    rsChamado.Close()
    SET rsChamado = NOTHING

    url = "http://br-pir1-m012/corporativo/chamados/Chamado/ChamadoDetalhes.asp?token=" & tokenAcesso

    assunto = "RE: " & titulo

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
    mensagem = mensagem & "<p>Olá, " & nomeAtendente & "!</p>"
    mensagem = mensagem & "<p>" & nomeSolicitante & ", respondeu um chamado.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Assunto:</strong> " & titulo & "</div>"
    mensagem = mensagem & "<div class=""info-item""> <i>" & mensagemUsuario & "</i></div>"
    mensagem = mensagem & "<div style=""margin-top: 20px; class=""info-item"">Para responder ou ver o histórico completo, acesse:</div>"
    mensagem = mensagem & "<div class=""info-item""><a href=""" & url & """>" & url & "</a></div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailAtendente <> "" Then
        Call EnviaEmail(emailAtendente, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailFinalizacao(idChamado)
    SQL = "SELECT c.Titulo, s.Nome as NomeSolicitante, s.Email as EmailSolicitante, c.TokenAcesso "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "WHERE c.Id = '" & idChamado & "' "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "

    On Error Resume Next
    SET rsChamado = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsChamado) Then
        Exit Sub
    End If
    On Error Goto 0

    titulo = rsChamado("Titulo")
    nomeSolicitante = rsChamado("NomeSolicitante")
    emailSolicitante = rsChamado("EmailSolicitante")
    tokenAcesso = rsChamado("TokenAcesso")

    rsChamado.Close()
    SET rsChamado = NOTHING

    assunto = "RE: " & titulo

    url = "http://br-pir1-m012/corporativo/chamados/Chamado/ChamadoDetalhes.asp?token=" & tokenAcesso

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
    mensagem = mensagem & "<p>Olá, " & nomeSolicitante & "!</p>"
    mensagem = mensagem & "<p>Seu chamado foi finalizado</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Assunto:</strong> " & titulo & "</div>"
    mensagem = mensagem & "<div style=""margin-top: 20px; class=""info-item"">Para mais informações, acesse:</div>"
    mensagem = mensagem & "<div class=""info-item""><a href=""" & url & """>" & url & "</a></div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailSolicitante <> "" Then
        Call EnviaEmail(emailSolicitante, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub EnviarEmailReabertura(idChamado)
    SQL = "SELECT c.Titulo, s.Nome as NomeSolicitante, s.Email as EmailSolicitante, c.TokenAcesso "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "LEFT JOIN SolicitantesChamado s ON s.Id = c.IdSolicitante "
    SQL = SQL & "WHERE c.Id = '" & idChamado & "' "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND s.FlagInativo = 0 "

    On Error Resume Next
    SET rsChamado = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsChamado) Then
        Exit Sub
    End If
    On Error Goto 0

    titulo = rsChamado("Titulo")
    nomeSolicitante = rsChamado("NomeSolicitante")
    emailSolicitante = rsChamado("EmailSolicitante")
    tokenAcesso = rsChamado("TokenAcesso")

    rsChamado.Close()
    SET rsChamado = NOTHING

    assunto = "RE: " & titulo

    url = "http://br-pir1-m012/corporativo/chamados/Chamado/ChamadoDetalhes.asp?token=" & tokenAcesso

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
    mensagem = mensagem & "<p>Olá, " & nomeSolicitante & "!</p>"
    mensagem = mensagem & "<p>Seu chamado foi reaberto</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Assunto:</strong> " & titulo & "</div>"
    mensagem = mensagem & "<div style=""margin-top: 20px; class=""info-item"">Para mais informações, acesse:</div>"
    mensagem = mensagem & "<div class=""info-item""><a href=""" & url & """>" & url & "</a></div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"
    
    If emailSolicitante <> "" Then
        Call EnviaEmail(emailSolicitante, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub RegistrarLog(acao, idChamado, nomeUsuario)   
    Select Case LCase(acao)
        Case "criacao"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Criação de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " criou o chamado de id " & idChamado
        Case "criacao_solicitante"
            idUsuario = "NULL"
            acaoLog = "Criação de Solicitante"
            descricaoLog = "Sistema criou o solicitante de id " & idChamado & " após abertura de chamado"
        Case "envio_mensagem"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Envio de Mensagem"
            descricaoLog = "Usuário " & nomeUsuario & " enviou uma mensagem ao chamado de id " & idChamado
        Case "finalizacao"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Finalização de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " finalizou o chamado de id " & idChamado
        Case "reabertura"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Reabertura de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " reabriu o chamado de id " & idChamado
        Case "atribuicao"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Atribuição de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " atribuiu o chamado de id " & idChamado 
          Case "avaliacao"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Avaliação de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " avaliou o chamado de id " & idChamado 
        Case "alteracao_status"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Alteração de Status de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " alterou o status do chamado de id " & idChamado
        Case "alteracao_prioridade"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Alteração de Prioridade de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " alterou a prioridade do chamado de id " & idChamado
        Case "alteracao_departamento"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Alteração de Departamento de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " alterou o departamento do chamado de id " & idChamado
        Case "alteracao_categoria"
            idUsuario = Iif(usuarioLogado("id") <> "", usuarioLogado("id"), "NULL")
            acaoLog = "Alteração de Categoria de Chamado"
            descricaoLog = "Usuário " & nomeUsuario & " alterou a categoria do chamado de id " & idChamado
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & idUsuario & ", "
    SQL = SQL & "'Chamados', "
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