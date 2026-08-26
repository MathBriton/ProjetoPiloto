<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")

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
    Case "listarOrcamentos"
        resultado = ListarOrcamentos()
    Case "desativarOrcamentos"
        resultado = DesativarOrcamentos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


'--------------------------------------
'           FUNÇÕES PRINCIPAIS
'--------------------------------------

Function ListarOrcamentos()    
    primeiroItem = True
    termoBusca = Request("termo")
    paginaAtual = Request("pagina")
    itensPorPagina = Request("itensPorPagina")

    ' ****** Logica de paginacao
    If paginaAtual = "" Or Not IsNumeric(paginaAtual) Then paginaAtual = 1
    If itensPorPagina = "" Or Not IsNumeric(itensPorPagina) Then itensPorPagina = 10
    
    paginaAtual = CLng(paginaAtual)
    itensPorPagina = CLng(itensPorPagina)
    offset = (paginaAtual - 1) * itensPorPagina
    
    ' Query para obter total de registos
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM ORCAMENTO_BRASIL t1 "
    SQL = SQL & "WHERE 1 = 1 "
    SQL = SQL & "   AND PLANTA IN (SELECT PLANTA FROM ANALISTAS_ORCAMENTO WHERE EXCLUIDO IS NULL AND CODIGO_USUARIO = '" & usuarioLogado("id") & "') "
    SQL = SQL & "   AND t1.STATUS_ORCAMENTO = 'Aprovado Total' "
    SQL = SQL & "   AND t1.VALIDADE_ATE >= GETDATE() "

    If termoBusca <> "" Then
        SQL = SQL & " AND ("
        SQL = SQL & "    t1.NR_ORCAMENTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.ID_ORCAMENTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.COD_CLIENTE LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.NOME_CLIENTE LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.PRODUTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & ")"
    End If

    On Error Resume Next
    SET rsCount = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCount) Then
        ListarOrcamentos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar orçamentos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    totalRegistros = 0
    If Not rsCount.EOF Then
        totalRegistros = rsCount("Total")
    End If
    
    If totalRegistros > 0 Then
        totalPaginas = Int((totalRegistros - 1) / itensPorPagina) + 1
    Else
        totalPaginas = 0
    End If

    ' ****** Query principal
    SQL = "SELECT "
    SQL = SQL & "    t1.ID_ORCAMENTO as IdOrcamento, "
    SQL = SQL & "    t1.NR_ORCAMENTO as NumeroOrcamento, "
    SQL = SQL & "    CASE t1.PLANTA "
    SQL = SQL & "        WHEN '00' THEN 'PIRAPETINGA' "               
    SQL = SQL & "        WHEN '04' THEN 'UBERABA' "            
    SQL = SQL & "        WHEN '05' THEN 'BENTO' "              
    SQL = SQL & "        WHEN '06' THEN 'FORTALEZA' "                  
    SQL = SQL & "        WHEN '07' THEN 'SAQUAREMA' "
    SQL = SQL & "    END as Planta, "
    SQL = SQL & "    CONCAT(t1.COD_CLIENTE,' - ',t1.NOME_CLIENTE) as NomeCliente, "
    SQL = SQL & "    t1.CNPJ, "
    SQL = SQL & "    CONCAT(t1.CIDADE,' - ',t1.UF) as Cidade, "
    SQL = SQL & "    t1.NOME_REPRESENTANTE as Representante, "
    SQL = SQL & "    CONVERT(VARCHAR,t1.DATA_INCLUSAO, 103) as DataInclusao, "
    SQL = SQL & "    CONVERT(VARCHAR,t1.VALIDADE_DE, 103) as DataValidadeDe, "
    SQL = SQL & "    CONVERT(VARCHAR,t1.VALIDADE_ATE, 103) as DataValidadeAte "
    SQL = SQL & "FROM ORCAMENTO_BRASIL t1 "
    SQL = SQL & "WHERE 1 = 1 "
    SQL = SQL & "   AND PLANTA IN (SELECT PLANTA FROM ANALISTAS_ORCAMENTO WHERE EXCLUIDO IS NULL AND CODIGO_USUARIO = '" & usuarioLogado("id") & "') "
    SQL = SQL & "   AND t1.STATUS_ORCAMENTO = 'Aprovado Total' "
    SQL = SQL & "   AND t1.VALIDADE_ATE >= GETDATE() "

    If termoBusca <> "" Then
        SQL = SQL & " AND ("
        SQL = SQL & "    t1.NR_ORCAMENTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.ID_ORCAMENTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.COD_CLIENTE LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.NOME_CLIENTE LIKE '%" & termoBusca & "%'"
        SQL = SQL & "    OR t1.PRODUTO LIKE '%" & termoBusca & "%'"
        SQL = SQL & ")"
    End If

    SQL = SQL & "ORDER BY NR_ORCAMENTO DESC "
    SQL = SQL & "OFFSET " & offset & " ROWS "
    SQL = SQL & "FETCH NEXT " & itensPorPagina & " ROWS ONLY"

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        ListarOrcamentos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar orçamentos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
        
    dadosJson = "["
    
    If Err.Number = 0 AND IsObject(rsListar) Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then
                dadosJson = dadosJson & ","
            End If
            primeiroItem = False
            
            dadosJson = dadosJson & "{"
            dadosJson = dadosJson & """idOrcamento"":""" & rsListar("IdOrcamento") & """"
            dadosJson = dadosJson & ",""numeroOrcamento"":""" & rsListar("NumeroOrcamento") & """"
            dadosJson = dadosJson & ",""planta"":""" & rsListar("Planta") & """"
            dadosJson = dadosJson & ",""nomeCliente"":""" & rsListar("NomeCliente") & """"
            dadosJson = dadosJson & ",""cnpj"":""" & rsListar("CNPJ") & """"
            dadosJson = dadosJson & ",""cidade"":""" & rsListar("Cidade") & """"
            dadosJson = dadosJson & ",""representante"":""" & rsListar("Representante") & """"
            dadosJson = dadosJson & ",""dataInclusao"":""" & rsListar("DataInclusao") & """"
            dadosJson = dadosJson & ",""dataValidadeDe"":""" & rsListar("DataValidadeDe") & """"
            dadosJson = dadosJson & ",""dataValidadeAte"":""" & rsListar("DataValidadeAte") & """"
            dadosJson = dadosJson & "}"
            
            rsListar.MoveNext
        Loop
        
        rsListar.Close()
        SET rsListar = NOTHING
    End If
        
    dadosJson = dadosJson & "]"
    
    jsonResposta = "{"
    jsonResposta = jsonResposta & """sucesso"": true"
    jsonResposta = jsonResposta & ",""mensagem"": ""Dados obtidos com sucesso!"""
    jsonResposta = jsonResposta & ",""dados"": " & dadosJson
    jsonResposta = jsonResposta & ",""totalRegistros"": " & totalRegistros
    jsonResposta = jsonResposta & ",""paginaAtual"": " & paginaAtual
    jsonResposta = jsonResposta & ",""totalPaginas"": " & totalPaginas
    jsonResposta = jsonResposta & ",""itensPorPagina"": " & itensPorPagina
    jsonResposta = jsonResposta & "}"
    
    ListarOrcamentos = jsonResposta
End Function


Function DesativarOrcamentos()    
    idsOrcamento = Request("idsOrcamento")

    If idsOrcamento = "" Then
        DesativarOrcamentos = "{""sucesso"": false, ""mensagem"": ""Nenhum orçamento informado!""}"
    End If

    SQL = "UPDATE ORCAMENTO_BRASIL SET VALIDADE_ATE = GETDATE() - 1"
    SQL = SQL & "WHERE ID_ORCAMENTO IN (" & idsOrcamento & ") "

    On Error Resume Next
    SET rsDesativar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not isObject(rsDesativar) then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Erro ao atualizar orçamentos""}"
        Exit Function
    End If
    
    rsDesativar.Close()
    SET rsDesativar = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("desativar_orcamentos", idsOrcamento)
    DesativarOrcamentos = "{""sucesso"": true, ""mensagem"": ""Orçamento atualizado com sucesso!""}"
End Function

Sub RegistrarLog(acao, idsOrcamento)
    Select Case LCase(acao)
        Case "desativar_orcamentos"
            acaoLog = "Desativar Orçamentos"
            descricaoLog = "Usuário " & usuarioLogado("login") & " desativou os orcamentos " & idsOrcamento 
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Desativar Orçamentos', "
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
%>