<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../../cnn/cnn.asp" -->
<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
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
    Case "obterPedido"
        resultado = ObterPedido()
    Case "salvarSolicitacaoAlteracao"
        resultado = SalvarSolicitacaoAlteracao()
    Case "listarTiposAlteracao"
        resultado = ListarTiposAlteracao()
    Case "listarFiliais"
        resultado = ListarFiliais()
    Case "listarCondicoesPagamento"
        resultado = ListarCondicoesPagamento()
    Case "listarTiposFrete"
        resultado = ListarTiposFrete()
    Case "listarStatus"
        resultado = ListarStatus()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

Function ObterPedido()
    codPedido = Request("codPedido")
    
    ' Query para obter dados do pedido
    SQL = "SELECT c6.C6_PRCVEN as Preco, c6.C6_QTDVEN as Quantidade, C5_NUM as CodPedido, "
    SQL = SQL & "C6_DESCRI as Referencia, C6_PRODUTO as CodProduto, sa1.A1_NREDUZ as NomeCliente, "
    SQL = SQL & "c6.C6_PRONTO as QuantidadePronto, c5.C5_TPFRETE as Frete, "
    SQL = SQL & "c5.C5_ESTATUS as Status, c6.C6_PEDCLI as PedidoCliente, "
    SQL = SQL & "FORMAT(CONVERT(DATE, c6.C6_ENTREG), 'yyyy-MM-dd') as DataEntregaDe, "
    SQL = SQL & "FORMAT(CONVERT(DATE, c6.C6_ENTREG2), 'yyyy-MM-dd') as DataEntregaAte, "
    SQL = SQL & "CASE c5.C5_INPUBE WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END NomeFilial, "
    SQL = SQL & "c5.C5_INPUBE as CodFilial, "
    SQL = SQL & "c5.C5_CLIENTE as CodCliente, "
    SQL = SQL & "c5.C5_LOJACLI as CodLoja, "
    SQL = SQL & "C5_FRETUNI as PrecoFrete, e4.E4_DESCRI as DescricaoCondicaoPagamento, "
    SQL = SQL & "e4.E4_CODIGO AS CodCondicaoPagamento "
    SQL = SQL & "FROM SC5010 c5 "
    SQL = SQL & "LEFT JOIN SC6010 c6 on c6.C6_NUM = c5.C5_NUM AND c6.D_E_L_E_T_ = '' "
    SQL = SQL & "LEFT JOIN SE4010 e4 on e4.E4_CODIGO = c5.C5_CONDPAG AND e4.D_E_L_E_T_ = '' "
    SQL = SQL & " LEFT JOIN SA1010 sa1 ON sa1.A1_COD = c5.C5_CLIENTE AND sa1.A1_LOJA = c5.C5_LOJACLI "
    SQL = SQL & "WHERE c5.D_E_L_E_T_ = '' "
    SQL = SQL & "AND c5.C5_NUM = '" & codPedido & "' "

    On Error Resume Next
    SET rsPedido = Conecta("PIR-DADOSADV", SQL)

    If Err.Number <> 0 OR Not IsObject(rsPedido) OR rsPedido.EOF Then
        ObterPedido = "{""sucesso"": false, ""mensagem"": ""Não foi possível encontrar este pedido ""}"
        Exit Function
    End If

    ' Query dados do cliente
    SQL = "SELECT TOP 1 A1_DTEMAIO AS IntervaloDataEntrega, A1_TOLMAIS AS Tolerancia " 
    SQL = SQL & "FROM VW_BI_CLIENTES "
    SQL = SQL & "WHERE COD_CLIENTE = '" & rsPedido("CodCliente") & "' "
    SQL = SQL & "AND LOJA_CLIENTE = '" & rsPedido("CodLoja") & "' "

    On Error Resume Next
    SET rsCliente = Conecta("PIR-DADOSADV", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCliente) OR rsCliente.EOF Then
        ObterPedido = "{""sucesso"": false, ""mensagem"": ""Não foi possível encontrar este pedido ""}"
        Exit Function
    End If

    ' Query para verificar se cliente aceita entrega final de semana
    SQL = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS permiteEntregaFinalDeSemana FROM DEF_ENTREGA_FDS "
    SQL = SQL & "WHERE COD_CLIENTE = '" & rsPedido("CodCliente") & "' "
    SQL = SQL & "AND (QUEM_EXCLUIU = 0 OR QUEM_EXCLUIU = '') "
    SQL = SQL & "AND ATIVO = 1 "

    On Error Resume Next
    SET rsEntrega = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsEntrega) OR rsEntrega.EOF Then
        ObterPedido = "{""sucesso"": false, ""mensagem"": ""Não foi possível verificar disponibilidade de entrega ""}"
        Exit Function
    End If

    dadosJSON = "{"
    dadosJSON = dadosJSON & """codPedido"": """ & EscapeJSON(rsPedido("CodPedido")) & """, "
    dadosJSON = dadosJSON & """referencia"": """ & EscapeJSON(rsPedido("Referencia")) & """, "
    dadosJSON = dadosJSON & """codProduto"": """ & EscapeJSON(rsPedido("CodProduto")) & """, "
    dadosJSON = dadosJSON & """preco"": """ & EscapeJSON(rsPedido("Preco")) & """, "
    dadosJSON = dadosJSON & """quantidade"": """ & EscapeJSON(rsPedido("Quantidade")) & """, "
    dadosJSON = dadosJSON & """quantidadePronto"": """ & EscapeJSON(rsPedido("QuantidadePronto")) & """, "
    dadosJSON = dadosJSON & """frete"": """ & EscapeJSON(rsPedido("Frete")) & """, "
    dadosJSON = dadosJSON & """status"": """ & EscapeJSON(rsPedido("Status")) & """, "
    dadosJSON = dadosJSON & """nomeCliente"": """ & EscapeJSON(rsPedido("NomeCliente")) & """, "
    dadosJSON = dadosJSON & """codCliente"": """ & EscapeJSON(rsPedido("CodCliente")) & """, "
    dadosJSON = dadosJSON & """codLoja"": """ & EscapeJSON(rsPedido("CodLoja")) & """, "
    dadosJSON = dadosJSON & """dataEntregaDe"": """ & EscapeJSON(rsPedido("DataEntregaDe")) & """, "
    dadosJSON = dadosJSON & """dataEntregaAte"": """ & EscapeJSON(rsPedido("DataEntregaAte")) & """, "
    dadosJSON = dadosJSON & """pedidoCliente"": """ & EscapeJSON(rsPedido("PedidoCliente")) & """, "
    dadosJSON = dadosJSON & """nomeFilial"": """ & EscapeJSON(rsPedido("NomeFilial")) & """, "
    dadosJSON = dadosJSON & """codFilial"": """ & EscapeJSON(rsPedido("CodFilial")) & """, "
    dadosJSON = dadosJSON & """precoFrete"": """ & EscapeJSON(rsPedido("PrecoFrete")) & """, "
    dadosJSON = dadosJSON & """descricaoCondicaoPagamento"": """ & EscapeJSON(rsPedido("DescricaoCondicaoPagamento")) & """, "
    dadosJSON = dadosJSON & """codCondicaoPagamento"": """ & EscapeJSON(rsPedido("CodCondicaoPagamento")) & """, "
    dadosJSON = dadosJSON & """permiteEntregaFinalDeSemana"": """ & rsEntrega("permiteEntregaFinalDeSemana") & """, "
    dadosJSON = dadosJSON & """intervaloDataEntrega"": """ & EscapeJSON(rsCliente("IntervaloDataEntrega")) & """, "
    dadosJSON = dadosJSON & """tolerancia"": """ & EscapeJSON(rsCliente("Tolerancia")) & """ "
    dadosJSON = dadosJSON & "}"

    rsEntrega.Close()
    SET rsEntrega = NOTHING

    rsPedido.Close()
    SET rsPedido = NOTHING

    rsCliente.Close()
    SET rsCliente = NOTHING

    ObterPedido = "{""sucesso"": true, ""mensagem"": ""Dados obtidos com sucesso!"", ""dados"": " & dadosJSON & "}"
End Function


Function ListarTiposFrete()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CIF" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "F" & """, "
    jsonArray = jsonArray & """descricao"": """ & "FOB" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposFrete = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "A" & """, "
    jsonArray = jsonArray & """descricao"": """ & "ATIVO" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "PROGRAMADO" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CANCELADO" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "L" & """, "
    jsonArray = jsonArray & """descricao"": """ & "LIQUIDADO" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarFiliais()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "U" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Uberaba" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """id"": """ & "I" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Pirapetinga" & """ "
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarFiliais = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarCondicoesPagamento()
    SQL = "SELECT E4_CODIGO, E4_DESCRI FROM ZZ7010 "
    SQL = SQL & "INNER JOIN SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' "
    SQL = SQL & "WHERE ZZ7_GRUPO = 'SE4' "
    SQL = SQL & "AND ZZ7010.D_E_L_E_T_ = '' "
    SQL = SQL & "AND E4_LIBVEN <> 'N' "
    SQL = SQL & "AND E4_MSBLQL <> '1' "
    SQL = SQL & "AND E4_CODIGO <> '090' "
    SQL = SQL & "ORDER BY E4_DESCRI"

    On Error Resume Next
    SET rsCondicoesPagamento = Conecta("PIR-DADOSADV", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCondicoesPagamento) Then
        ListarCondicoesPagamento = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de tipos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCondicoesPagamento.EOF Then
        Do While Not rsCondicoesPagamento.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsCondicoesPagamento("E4_CODIGO")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsCondicoesPagamento("E4_DESCRI")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCondicoesPagamento.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCondicoesPagamento.Close()
    SET rsCondicoesPagamento = NOTHING
    
    ListarCondicoesPagamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function


Function ListarTiposAlteracao()
    SQL = "EXEC dbo.SP_INPANET_HISTORICO_LISTA_OPERACOES_ALTERACOES_PEDIDO 14"

    On Error Resume Next
    SET rsTiposAlteracao = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTiposAlteracao) Then
        ListarTiposAlteracao = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de tipos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsTiposAlteracao.EOF Then
        Do While Not rsTiposAlteracao.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsTiposAlteracao("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsTiposAlteracao("Operacao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsTiposAlteracao.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsTiposAlteracao.Close()
    SET rsTiposAlteracao = NOTHING
    
    ListarTiposAlteracao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function SalvarSolicitacaoAlteracao() 
	idSolicitacao = 0
    evidenciasJSON = request("evidencias")
    codFilial = request("codFilial")
    codPedido = request("codPedido")
    tipoAlteracao = request("tipoAlteracao")
    valorDe = request("valorDe")
    valorPara = request("valorPara")
    valorPara2 = request("valorPara2")
    observacao = request("observacao")

    ehObrigatorioEvidencia = VerificaObrigatoriedadeEvidencias(tipoAlteracao, codFilial, codPedido, valorDe, valorPara, valorPara2)

    If ehObrigatorioEvidencia AND (evidenciasJSON = "" OR evidenciasJSON = "[]") Then
        SalvarSolicitacaoAlteracao = "{""sucesso"": false, ""mensagem"": ""O anexo de evidência é obrigatório para esse tipo de alteração.""}"
        Exit Function
    End If
    
    SQL = "BEGIN TRANSACTION TRANS; " 

    SQL = SQL & "INSERT INTO InHistoricoAlteracoesPedidosSolicitacao( "
    SQL = SQL & "Pedido, "
    SQL = SQL & "Tipo, "
    SQL = SQL & "De, "
    SQL = SQL & "Para, "
    SQL = SQL & "Para2, "
    SQL = SQL & "Observacao, "
    SQL = SQL & "Realizado, "
    SQL = SQL & "Descartado, "
    SQL = SQL & "Data, "
    SQL = SQL & "Hora, "
    SQL = SQL & "Login "
    SQL = SQL & ") OUTPUT INSERTED.Id VALUES ( "
    SQL = SQL & " '" & codPedido & "', "
    SQL = SQL & " '" & tipoAlteracao & "', "
    SQL = SQL & " '" & Replace(valorDe, ".", ",") & "', "
    SQL = SQL & " '" & Replace(valorPara, ".", ",") & "', "
    SQL = SQL & " '" & Replace(valorPara2, ".", ",") & "', "
    SQL = SQL & " '" & observacao & "', "
    SQL = SQL & " 0, "
    SQL = SQL & " 0, "
    SQL = SQL & " FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & " FORMAT(GETDATE(), 'HH:mm:ss'), "
    SQL = SQL & " '" & usuarioLogado("login") & "' "
    SQL = SQL & "); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION TRANS "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION TRANS "

	SET rsSalvar = Conecta( "PIR-HELPDESK", SQL )
	If Err.Number = 0 AND IsObject(rsSalvar) Then		
		idSolicitacao = rsSalvar("Id")
		
		rsSalvar.Close()
		SET rsSalvar = NOTHING
	End If
  
	If evidenciasJSON <> "" Then
		Call SalvarEvidencias(idSolicitacao, evidenciasJSON)
	End If

    Call RegistrarLog("criacao", idSolicitacao)
    SalvarSolicitacaoAlteracao = "{""sucesso"": true, ""mensagem"": ""Solicitação salva com sucesso""}"
End Function

' Funções auxiliares
Function VerificaObrigatoriedadeEvidencias(tipoAlteracao,codFilial, codPedido, valorDe, valorPara, valorPara2)
    ehObrigatorio = False

    Select Case tipoAlteracao
        Case "1"
            ehObrigatorio = False
        Case "2"
            ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, True)
        Case "3"
            ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, True)
        Case "4"
            ehObrigatorio = False
        Case "5"
            ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, True)
        Case "6"
            ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, True)
        Case "7"
            ' Para tipo 7 é necessário verificar se o pedido possui programação
            stringConexao = Iif(codFilial = "U", "UBE-TRIMBOX", "PIR-TRIMBOX")

            SQL = "SELECT dbo.fq_PedidoProduzidoOuProgramado('" & codPedido & "', '" & "01" & "') as programado"
            SET rsVerificar = Conecta(stringConexao, SQL)

            If Not rsVerificar.EOF Then
                ehProgramado = Iif(rsVerificar("programado") = 0, False, True)
                ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, ehProgramado)
            End If

            rsVerificar.Close()
            SET rsVerificar = NOTHING
        Case "8"
            ehObrigatorio = Iif(usuarioLogado("nivel") = 1, False, True)
        Case "9"
            ' Só é obrigatório se mudar de L para A
            If valorDe = "L" And valorPara = "A" Then
                ehObrigatorio = True
            End If
        Case "10"
            ehObrigatorio = False
        Case "14"
            ehObrigatorio = False
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
    End Select

    VerificaObrigatoriedadeEvidencias = ehObrigatorio
End Function

Sub SalvarEvidencias(idSolicitacao, evidenciasJSON)
    evidenciasJSON = Replace(evidenciasJSON, "[", "")
    evidenciasJSON = Replace(evidenciasJSON, "]", "")
    
    Dim anexosArray, anexo, i
    anexosArray = Split(evidenciasJSON, "},{")
    
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
        
        SQL = "INSERT INTO SolicitacaoAlteracaoPedidoEvidencia "
        SQL = SQL & "(IdSolicitacaoAlteracaoPedido, Nome, Caminho, DataCriacao) VALUES ("
        SQL = SQL & idSolicitacao & ", "
        SQL = SQL & "'" & nome & "', "
        SQL = SQL & "'" & caminho & "', "
        SQL = SQL & "GETDATE()"
        SQL = SQL & ")"
        
        On Error Resume Next
        SET rsEvidencias = Conecta("PIR-HELPDESK", SQL)
        
        If Err.Number = 0 AND IsObject(rsEvidencias) Then
            rsEvidencias.Close()
            SET rsEvidencias = NOTHING
        End If
    Next
End Sub

Sub RegistrarLog(acao, id)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Solicitação de Alteração de Pedido"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a solicitação de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Alteração de Pedidos', "
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