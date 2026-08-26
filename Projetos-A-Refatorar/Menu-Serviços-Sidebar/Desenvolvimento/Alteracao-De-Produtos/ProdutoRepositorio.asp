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

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")

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
    Case "obterProdutos"
        resultado = ObterProdutos()
    Case "alterarStatus"
        resultado = AlterarStatus()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


'--------------------------------------
'           FUNÇÕES PRINCIPAIS
'--------------------------------------

Function ObterProdutos()    
    primeiroItem = True
    termoBusca = Request("termo")
    paginaAtual = Request("pagina")
    itensPorPagina = Request("itensPorPagina")

    If SESSION("FILIAL") = "00" Then
        filialUsuario = "I"
    ElseIf SESSION("FILIAL") = "04" Then
        filialUsuario = "U"
    End If
    
    ' ****** Logica de paginacao
    If paginaAtual = "" Or Not IsNumeric(paginaAtual) Then paginaAtual = 1
    If itensPorPagina = "" Or Not IsNumeric(itensPorPagina) Then itensPorPagina = 10
    
    paginaAtual = CLng(paginaAtual)
    itensPorPagina = CLng(itensPorPagina)
    
    offset = (paginaAtual - 1) * itensPorPagina
    
    clausulaWhere = "WHERE 1=1 "
    
    If termoBusca <> "" Then
        clausulaWhere = clausulaWhere & "AND ( "
        clausulaWhere = clausulaWhere & "   B1_COD LIKE '%" & Replace(termoBusca, "'", "''") & "%' "
        clausulaWhere = clausulaWhere & "   OR B1_DESC LIKE '%" & Replace(termoBusca, "'", "''") & "%'"
        clausulaWhere = clausulaWhere & ")"
    End If
    
    sqlCount = "SELECT COUNT(*) as Total FROM SB1010 " & clausulaWhere
    
    On Error Resume Next
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", sqlCount)
    
    totalRegistros = 0
    If Err.Number = 0 AND IsObject(PIR_DADOSADV) Then
        If Not PIR_DADOSADV.EOF Then
            totalRegistros = PIR_DADOSADV("Total")
        End If
        PIR_DADOSADV.Close()
        SET PIR_DADOSADV = NOTHING
    End If
    
    If totalRegistros > 0 Then
        totalPaginas = Int((totalRegistros - 1) / itensPorPagina) + 1
    Else
        totalPaginas = 0
    End If

    ' ****** Query principal

    SQL = SQL & "WITH SB1 AS ("
    SQL = SQL & "    SELECT  "
    SQL = SQL & "        B1_COD, "
    SQL = SQL & "        B1_STATUS, "
    SQL = SQL & "        B1_DESC, "
    SQL = SQL & "        CASE  "
    SQL = SQL & "            WHEN SUBSTRING(B1_COD,1,3) = '750' THEN CONCAT('710',RIGHT(RTRIM(B1_COD),6))  "
    SQL = SQL & "            WHEN SUBSTRING(B1_COD,1,3) = '760' THEN CONCAT('761',RIGHT(RTRIM(B1_COD),6)) "
    SQL = SQL & "            ELSE B1_COD  "
    SQL = SQL & "        END CODIGO "
    SQL = SQL & "    FROM SB1010 "
    SQL = SQL & "    WHERE B1_COD BETWEEN '710000000'AND '769999999'AND B1_MSBLQL <>'1' AND SB1010.D_E_L_E_T_='' "
    SQL = SQL & "), "
    SQL = SQL & "SB5 AS ( "
    SQL = SQL & "    SELECT  "
    SQL = SQL & "        B5_COD, "
    SQL = SQL & "        B5_CLI, "
    SQL = SQL & "        B5_LOJA, "
    SQL = SQL & "        B5_NOMECLI, "
    SQL = SQL & "        ROW_NUMBER() OVER ( "
    SQL = SQL & "            PARTITION BY B5_COD, B5_CLI, B5_LOJA "
    SQL = SQL & "            ORDER BY B5_LOJA "
    SQL = SQL & "        ) VALIDADOR "
    SQL = SQL & "    FROM SB5010 "
    SQL = SQL & "    WHERE D_E_L_E_T_ = '' "
    SQL = SQL & ") "
    SQL = SQL & "SELECT "
    SQL = SQL & "    B1_COD CodProduto, "
    SQL = SQL & "    B1_STATUS Status, "
    SQL = SQL & "    B1_DESC Descricao, "
    SQL = SQL & "    B5_NOMECLI NomeCliente, "
    SQL = SQL & "    B5_CLI CodCLiente, "
    SQL = SQL & "    B5_LOJA CodLoja "
    SQL = SQL & "FROM SB1  "
    SQL = SQL & "LEFT JOIN SB5 ON CODIGO = B5_COD AND VALIDADOR = '1' "
    SQL = SQL & clausulaWhere
    SQL = SQL & "ORDER BY B1_COD "
    SQL = SQL & "OFFSET " & offset & " ROWS "
    SQL = SQL & "FETCH NEXT " & itensPorPagina & " ROWS ONLY"
        
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
    
    dadosJson = "["
    
    If Err.Number = 0 AND IsObject(PIR_DADOSADV) Then
        Do While Not PIR_DADOSADV.EOF
            If Not primeiroItem Then
                dadosJson = dadosJson & ","
            End If
            primeiroItem = False
            
            dadosJson = dadosJson & "{"
            dadosJson = dadosJson & """codProduto"":""" & PIR_DADOSADV("CodProduto") & """"
            dadosJson = dadosJson & ",""status"":""" & PIR_DADOSADV("Status") & """"
            dadosJson = dadosJson & ",""descricao"":""" & PIR_DADOSADV("Descricao") & """"
            dadosJson = dadosJson & ",""nomeCliente"":""" & PIR_DADOSADV("NomeCliente") & """"
            dadosJson = dadosJson & ",""codCliente"":""" & PIR_DADOSADV("CodCLiente") & """"
            dadosJson = dadosJson & ",""codLoja"":""" & PIR_DADOSADV("CodLoja") & """"
            dadosJson = dadosJson & "}"
            
            PIR_DADOSADV.MoveNext
        Loop
        
        PIR_DADOSADV.Close()
        SET PIR_DADOSADV = NOTHING
    End If
    
    On Error Goto 0
    
    dadosJson = dadosJson & "]"
    
    jsonResposta = "{"
    jsonResposta = jsonResposta & """sucesso"": true"
    jsonResposta = jsonResposta & ",""mensagem"": ""Dados obtidos com sucesso!"""
    jsonResposta = jsonResposta & ",""produtos"": " & dadosJson
    jsonResposta = jsonResposta & ",""totalRegistros"": " & totalRegistros
    jsonResposta = jsonResposta & ",""paginaAtual"": " & paginaAtual
    jsonResposta = jsonResposta & ",""totalPaginas"": " & totalPaginas
    jsonResposta = jsonResposta & ",""itensPorPagina"": " & itensPorPagina
    jsonResposta = jsonResposta & "}"
    
    ObterProdutos = jsonResposta
End Function


Function AlterarStatus()    
    status = Request("status")
    produtosIds = Request("produtosIds")

    If status = "" Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Nenhum status informado!""}"
    End If

    If produtosIds = "" Then
        AlterarStatus = "{""sucesso"": false, ""mensagem"": ""Nenhum produto informado!""}"
    End If

    SQL = "UPDATE SB1010"
    SQL = SQL & " SET B1_STATUS = '" & status & "'"
    SQL = SQL & " WHERE B1_COD IN (" & produtosIds & ") " 

    On Error Resume Next
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
    
    If Err.Number <> 0 OR Not isObject(PIR_DADOSADV) then
        ObterDadosTransferencia = "{""sucesso"": false, ""mensagem"": ""Erro ao atualizar status""}"
        Exit Function
    End If
    
    PIR_DADOSADV.Close()
    SET PIR_DADOSADV = NOTHING
    On Error Goto 0
    
    ' Alterações para V (Verificado) é necessário enviar email para lista de usuários 
    If status = "V" Then
        Call EnviarEmailNotificacao(produtosIds)
    End If

    Call RegistrarLog("alterarstatus", status, produtosIds)
    AlterarStatus = "{""sucesso"": true, ""mensagem"": ""Status atualizado com sucesso!""}"
End Function

Sub EnviarEmailNotificacao(produtoIds)
    emailUsuarios = ""
    produtoIds = Replace(produtoIds, "'", "")

    SQL = "SELECT X6_VAR, X6_CONTEUD AS Email FROM SX6010 WHERE X6_VAR = 'SK_MAILFTU' AND D_E_L_E_T_ = ''"

    On Error Resume Next
    SET rsEmail = Conecta("PIR-DADOSADV", SQL)

    If Err.Number = 0 AND IsObject(rsEmail) Then
        While Not rsEmail.EOF
            emailUsuarios = emailUsuarios & ";" & rsEmail("Email")
            rsEmail.MoveNext()
        Wend

        rsEmail.Close()
        SET rsEmail = NOTHING
    End If
    On Error Goto 0

    detalhesProdutos = ""

    SQL = "SELECT "
    SQL = SQL & "   B1_COD AS CodProduto, "
    SQL = SQL & "   DadosCliente.B5_NOMECLI AS NomeCliente, "
    SQL = SQL & "   DadosCliente.B5_CLI AS CodigoCliente, "
    SQL = SQL & "   DadosCliente.B5_LOJA AS CodigoLoja "
    SQL = SQL & "FROM SB1010 "
    SQL = SQL & "OUTER APPLY ( "
    SQL = SQL & "   SELECT TOP 1 B5_NOMECLI, B5_CLI, B5_LOJA "
    SQL = SQL & "   FROM SB5010 "
    SQL = SQL & "   WHERE RIGHT(RTRIM(B5_COD), 6) = RIGHT(RTRIM(B1_COD), 6) "
    SQL = SQL & "   AND SB5010.D_E_L_E_T_ = '' "
    SQL = SQL & "   ORDER BY CASE WHEN B5_COD = B1_COD THEN 0 ELSE 1 END, B5_COD "
    SQL = SQL & ") AS DadosCliente "
    SQL = SQL & "WHERE B1_COD IN ('" & Replace(produtoIds, ",", "','") & "') "
    SQL = SQL & "AND SB1010.D_E_L_E_T_ = '' "
    SQL = SQL & "ORDER BY B1_COD"

    On Error Resume Next
    SET rsProdutos = Conecta("PIR-DADOSADV", SQL)

    If Err.Number = 0 AND IsObject(rsProdutos) Then
        While Not rsProdutos.EOF
            detalhesProdutos = detalhesProdutos & "<div class=""produto-item"">"
            detalhesProdutos = detalhesProdutos & "<div class=""info-item""><strong>Produto:</strong> " & rsProdutos("CodProduto") & "</div>"
            detalhesProdutos = detalhesProdutos & "<div class=""info-item""><strong>Nome Cliente:</strong> " & rsProdutos("NomeCliente") & "</div>"
            detalhesProdutos = detalhesProdutos & "<div class=""info-item""><strong>Código cliente:</strong> " & rsProdutos("CodigoCliente") & "</div>"
            detalhesProdutos = detalhesProdutos & "<div class=""info-item""><strong>Código loja:</strong> " & rsProdutos("CodigoLoja") & "</div>"
            detalhesProdutos = detalhesProdutos & "</div>"
            rsProdutos.MoveNext()
        Wend

        rsProdutos.Close()
        SET rsProdutos = NOTHING
    End If
    On Error Goto 0

    assunto = "Produtos verificados: " & produtoIds
    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".produto-item { border-bottom: 1px solid #ddd; padding: 10px 0; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>Os seguintes produtos tiveram seu status alterado para V - Verificado:</p>"
    mensagem = mensagem & detalhesProdutos
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    If emailUsuarios <> "" Then
        Call EnviaEmail(emailUsuarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    End If
End Sub

Sub RegistrarLog(acao, statusNovo, produtosIds)
    Dim SQL, LOG_CONN
    Dim descricaoLog, acaoLog
    
    Select Case LCase(acao)
        Case "alterarstatus"
            acaoLog = "Alteração de status produto"
            descricaoLog = "Usuário " & usuarioLogado("login") & " alterou dos produtos " & produtosIds & " para " & ObterStatusTraduzido(statusNovo) & ": " & statusNovo
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Alteração de Produtos', "
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

Function ObterStatusTraduzido(status) 
    Dim statusTraduzido 
    statusTraduzido = ""

    Select Case status
        Case "A"
            statusTraduzido = "Aprovado"
        Case "C"
            statusTraduzido = "Cancelado"
        Case "P"
            statusTraduzido = "Projeto"
        Case "V"
            statusTraduzido = "Verificado"
        Case "W"
            statusTraduzido = "Validado"
        Case "E"
            statusTraduzido = "Em Espera"
        Case Else
            statusTraduzido = ""
    End Select

    ObterStatusTraduzido = statusTraduzido
End Function

%>