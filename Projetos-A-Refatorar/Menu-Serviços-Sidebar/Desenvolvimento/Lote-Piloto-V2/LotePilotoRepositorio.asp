<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<!--#Include File="../includes/variaveis.asp"-->
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
    Case "excluir"
        resultado = Excluir()
    Case "obterPorId"
        resultado = ObterPorId()
    Case "listarClientes"
        resultado = ListarClientes()
    Case "obterDetalhesPedido"
        resultado = ObterDetalhesPedido()
    Case "listarTiposItem"
        resultado = ListarTiposItem()
    Case "listarMotivos"
        resultado = ListarMotivos()
    Case "listarStatus"
        resultado = ListarStatus()
    Case "listarTiposLotePiloto"
        resultado = ListarTiposLotePiloto()
    Case "listarHistoricoStatus"
        resultado = ListarHistoricoStatus()
    Case "listarHistoricoJustificativas"
        resultado = ListarHistoricoJustificativas()
    Case "listarProdutos"
        resultado = ListarProdutos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select


' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    id = Request("filtroId")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")
    codigoPlanta = Request("filtroFilial")
    status = Request("filtroStatus")
    codigoCliente = Request("filtroCodigoCliente")

    SQL = "SELECT "
    SQL = SQL & "l.ID AS Id, "
    SQL = SQL & "l.PLANTA AS CodigoPlanta, "
    SQL = SQL & "l.CLIENTE AS CodigoCliente, "
    SQL = SQL & "l.LOJA AS CodigoLoja, "
    SQL = SQL & "l.NOME_CLIENTE AS NomeCliente, "
    SQL = SQL & "l.PRODUTO AS CodigoProduto, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), l.DESCRICAO_PRODUTO), 'json') AS Referencia, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), l.MOTIVO_DESCRICAO), 'json') AS MotivoDescricao, "
    SQL = SQL & "l.QUANTIDADE_PALETES AS QuantidadePaletes, "
    SQL = SQL & "CASE WHEN l.FLAG_DDP = '1' THEN 1 ELSE 0 END AS AcompanhamentoDDP, "
    SQL = SQL & "CASE WHEN l.FLAG_CLICHERIA = '1' THEN 1 ELSE 0 END AS AcompanhamentoClicheria, "
    SQL = SQL & "CASE WHEN l.FLAG_ONDULADEIRA = '1' THEN 1 ELSE 0 END AS AcompanhamentoOnduladeira, "
    SQL = SQL & "CASE WHEN l.FLAG_IMPRESSORA = '1' THEN 1 ELSE 0 END AS AcompanhamentoImpressora, "
    SQL = SQL & "CASE WHEN l.FLAG_EXPEDICAO = '1' THEN 1 ELSE 0 END AS AcompanhamentoExpedicao, "
    SQL = SQL & "CASE WHEN l.FLAG_CLIENTE = '1' THEN 1 ELSE 0 END AS AcompanhamentoCliente, "
    SQL = SQL & "l.QUANTIDADE_PEDIDOS AS QuantidadePedidos, "
    SQL = SQL & "l.MOTIVO AS Motivo, "
    SQL = SQL & "l.QUEM_CADASTROU AS IdUsuarioInclusao, "
    SQL = SQL & "c.CRM_NOME as NomeUsuarioInclusao, "
    SQL = SQL & "CASE WHEN l.LIBERACAO = '1' THEN 1 ELSE 0 END AS PrecisaLiberacao, "
    SQL = SQL & "CASE WHEN l.APROVACAO = '1' THEN 1 ELSE 0 END AS LotePilotoAprovado, "
    SQL = SQL & "FORMAT(l.DATA_APROVACAO,'yyyy-MM-dd') as DataAprovacao, "
    SQL = SQL & "FORMAT(l.DATA_HORA,'yyyy-MM-dd') as DataInclusao, "
    SQL = SQL & "l.STATUS AS Status, "
    SQL = SQL & "l.TIPO_LP AS TipoLP, "
    SQL = SQL & "l.PEDIDO_WRK AS PedidoWRK, "
    SQL = SQL & "l.PRODUTO_ANTERIOR AS ProdutoAnterior, "
    SQL = SQL & "("
	SQL = SQL & "	SELECT "
	SQL = SQL & "	FORMAT(MAX(l2.DATA_INCLUSAO), 'yyyy-MM-dd')"
	SQL = SQL & "	FROM LOTE_PILOTO_HISTORICO_STATUS l2 "
	SQL = SQL & "	WHERE l.ID = l2.IDLOTEPILOTO"
	SQL = SQL & ") AS DataStatus "
    SQL = SQL & "FROM LOTE_PILOTO l "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c on c.CRM_R_E_C_N_O = l.QUEM_CADASTROU  "
    SQL = SQL & "WHERE 1=1 "
    
    If codigoPlanta <> "" Then
        SQL = SQL & "AND l.PLANTA = '" & codigoPlanta & "' "
    End If

    If codigoCliente <> "" Then
        SQL = SQL & "AND l.CLIENTE = '" & codigoCliente & "' "
    End If

    If status <> "" Then
        SQL = SQL & "AND l.STATUS = '" & status & "' "
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND l.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND l.DataInclusao < DATEADD(DAY, 1, '" & Replace(dataInclusaoAte & "", "-", "") & "')"
    End If

    SQL = SQL & "ORDER BY l.ID DESC"

    On Error Resume Next
    SET rsListar = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de Cadastros: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id")) & """, "
            jsonArray = jsonArray & """codigoPlanta"": """ & EscapeJSON(rsListar("CodigoPlanta")) & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & EscapeJSON(rsListar("CodigoCliente")) & """, "
            jsonArray = jsonArray & """codigoLoja"": """ & EscapeJSON(rsListar("CodigoLoja")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsListar("NomeCliente")) & """, "
            jsonArray = jsonArray & """codigoProduto"": """ & EscapeJSON(rsListar("CodigoProduto")) & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsListar("Referencia")) & """, "
            jsonArray = jsonArray & """quantidadePaletes"": """ & EscapeJSON(rsListar("QuantidadePaletes")) & """, "
            jsonArray = jsonArray & """quantidadePedidos"": """ & EscapeJSON(rsListar("QuantidadePedidos")) & """, "
            jsonArray = jsonArray & """acompanhamentoDDP"": """ & EscapeJSON(rsListar("AcompanhamentoDDP")) & """, "
            jsonArray = jsonArray & """acompanhamentoClicheria"": """ & EscapeJSON(rsListar("AcompanhamentoClicheria")) & """, "
            jsonArray = jsonArray & """acompanhamentoOnduladeira"": """ & EscapeJSON(rsListar("AcompanhamentoOnduladeira")) & """, "
            jsonArray = jsonArray & """acompanhamentoImpressora"": """ & EscapeJSON(rsListar("AcompanhamentoImpressora")) & """, "
            jsonArray = jsonArray & """acompanhamentoExpedicao"": """ & EscapeJSON(rsListar("AcompanhamentoExpedicao")) & """, "
            jsonArray = jsonArray & """acompanhamentoCliente"": """ & EscapeJSON(rsListar("AcompanhamentoCliente")) & """, "
            jsonArray = jsonArray & """motivo"": """ & EscapeJSON(rsListar("Motivo")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsListar("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsListar("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """precisaLiberacao"": """ & EscapeJSON(rsListar("PrecisaLiberacao")) & """, "
            jsonArray = jsonArray & """lotePilotoAprovado"": """ & EscapeJSON(rsListar("LotePilotoAprovado")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """motivoDescricao"": """ & EscapeJSON(rsListar("MotivoDescricao")) & """, "
            jsonArray = jsonArray & """tipoLP"": """ & EscapeJSON(rsListar("TipoLP")) & """, "
            jsonArray = jsonArray & """pedidoWRK"": """ & EscapeJSON(rsListar("PedidoWRK")) & """, "
            jsonArray = jsonArray & """produtoAnterior"": """ & EscapeJSON(rsListar("ProdutoAnterior")) & """, "
            jsonArray = jsonArray & """dataAprovacao"": """ & EscapeJSON(rsListar("DataAprovacao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsListar("DataInclusao")) & """, "
            jsonArray = jsonArray & """dataStatus"": """ & EscapeJSON(rsListar("DataStatus")) & """ "
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
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Nenhum código informado.""}"
        Exit Function
    End If
 
    SQL = "SELECT TOP 1 "
    SQL = SQL & "ID AS Id, "
    SQL = SQL & "PLANTA AS CodigoPlanta, "
    SQL = SQL & "CLIENTE AS CodigoCliente, "
    SQL = SQL & "LOJA AS CodigoLoja, "
    SQL = SQL & "NOME_CLIENTE AS NomeCliente, "
    SQL = SQL & "PRODUTO AS CodigoProduto, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), DESCRICAO_PRODUTO), 'json') AS Referencia, "
    SQL = SQL & "STRING_ESCAPE(CONVERT(VARCHAR(MAX), MOTIVO_DESCRICAO), 'json') AS MotivoDescricao, "
    SQL = SQL & "QUANTIDADE_PALETES AS QuantidadePaletes, "
    SQL = SQL & "CASE WHEN FLAG_DDP = '1' THEN 1 ELSE 0 END AS AcompanhamentoDDP, "
    SQL = SQL & "CASE WHEN FLAG_CLICHERIA = '1' THEN 1 ELSE 0 END AS AcompanhamentoClicheria, "
    SQL = SQL & "CASE WHEN FLAG_ONDULADEIRA = '1' THEN 1 ELSE 0 END AS AcompanhamentoOnduladeira, "
    SQL = SQL & "CASE WHEN FLAG_IMPRESSORA = '1' THEN 1 ELSE 0 END AS AcompanhamentoImpressora, "
    SQL = SQL & "CASE WHEN FLAG_EXPEDICAO = '1' THEN 1 ELSE 0 END AS AcompanhamentoExpedicao, "
    SQL = SQL & "CASE WHEN FLAG_CLIENTE = '1' THEN 1 ELSE 0 END AS AcompanhamentoCliente, "
    SQL = SQL & "CASE WHEN LIBERACAO = '1' THEN 1 ELSE 0 END AS PrecisaLiberacao, "
    SQL = SQL & "CASE WHEN APROVACAO = '1' THEN 1 ELSE 0 END AS LotePilotoAprovado, "
    SQL = SQL & "QUANTIDADE_PEDIDOS AS QuantidadePedidos, "
    SQL = SQL & "MOTIVO AS Motivo, "
    SQL = SQL & "QUEM_CADASTROU AS IdUsuarioInclusao, "
    SQL = SQL & "DATA_HORA AS DataInclusao, "
    SQL = SQL & "DATA_APROVACAO AS DataAprovacao, "
    SQL = SQL & "STATUS AS Status, "
    SQL = SQL & "TIPO_LP AS TipoLP, "
    SQL = SQL & "PEDIDO_WRK AS PedidoWRK, "
    SQL = SQL & "PRODUTO_ANTERIOR AS ProdutoAnterior "
    SQL = SQL & "FROM LOTE_PILOTO "
    SQL = SQL & "WHERE ID = '" & id & "' "

    On Error Resume Next
    SET rsObter = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de lotes piloto: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsObter.EOF Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Lote Piloto não encontrado.""}"
        Exit Function
    End If

    jsonObjeto = "{"
    jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id")) & """, "
    jsonObjeto = jsonObjeto & """codigoPlanta"": """ & EscapeJSON(rsObter("CodigoPlanta")) & """, "
    jsonObjeto = jsonObjeto & """codigoCliente"": """ & EscapeJSON(rsObter("CodigoCliente")) & """, "
    jsonObjeto = jsonObjeto & """codigoLoja"": """ & EscapeJSON(rsObter("CodigoLoja")) & """, "
    jsonObjeto = jsonObjeto & """nomeCliente"": """ & EscapeJSON(rsObter("NomeCliente")) & """, "
    jsonObjeto = jsonObjeto & """codigoProduto"": """ & EscapeJSON(rsObter("CodigoProduto")) & """, "
    jsonObjeto = jsonObjeto & """referencia"": """ & EscapeJSON(rsObter("Referencia")) & """, "
    jsonObjeto = jsonObjeto & """quantidadePaletes"": """ & EscapeJSON(rsObter("QuantidadePaletes")) & """, "
    jsonObjeto = jsonObjeto & """quantidadePedidos"": """ & EscapeJSON(rsObter("QuantidadePedidos")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoDDP"": """ & EscapeJSON(rsObter("AcompanhamentoDDP")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoClicheria"": """ & EscapeJSON(rsObter("AcompanhamentoClicheria")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoOnduladeira"": """ & EscapeJSON(rsObter("AcompanhamentoOnduladeira")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoImpressora"": """ & EscapeJSON(rsObter("AcompanhamentoImpressora")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoExpedicao"": """ & EscapeJSON(rsObter("AcompanhamentoExpedicao")) & """, "
    jsonObjeto = jsonObjeto & """acompanhamentoCliente"": """ & EscapeJSON(rsObter("AcompanhamentoCliente")) & """, "
    jsonObjeto = jsonObjeto & """motivo"": """ & EscapeJSON(rsObter("Motivo")) & """, "
    jsonObjeto = jsonObjeto & """idUsuarioInclusao"": """ & EscapeJSON(rsObter("IdUsuarioInclusao")) & """, "
    jsonObjeto = jsonObjeto & """dataInclusao"": """ & EscapeJSON(rsObter("DataInclusao")) & """, "
    jsonObjeto = jsonObjeto & """precisaLiberacao"": """ & EscapeJSON(rsObter("PrecisaLiberacao")) & """, "
    jsonObjeto = jsonObjeto & """lotePilotoAprovado"": """ & EscapeJSON(rsObter("LotePilotoAprovado")) & """, "
    jsonObjeto = jsonObjeto & """dataAprovacao"": """ & EscapeJSON(rsObter("DataAprovacao")) & """, "
    jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
    jsonObjeto = jsonObjeto & """motivoDescricao"": """ & EscapeJSON(rsObter("MotivoDescricao")) & """, "
    jsonObjeto = jsonObjeto & """tipoLP"": """ & EscapeJSON(rsObter("TipoLP")) & """, "
    jsonObjeto = jsonObjeto & """pedidoWRK"": """ & EscapeJSON(rsObter("PedidoWRK")) & """, "
    jsonObjeto = jsonObjeto & """produtoAnterior"": """ & EscapeJSON(rsObter("ProdutoAnterior")) & """ "
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()  
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("codigoCliente") = Request("codigoCliente")
    Dados("codigoLoja") = Request("codigoLoja")
    Dados("nomeCliente") = Request("nomeCliente")
    Dados("tipoItem") = Request("tipoItem")
    Dados("codigoProduto") = Request("codigoProduto")
    Dados("referencia") = Request("referencia")
    Dados("quantidadePaletes") = Request("quantidadePaletes")
    Dados("quantidadePedidos") = Request("quantidadePedidos")
    Dados("pedidoWRK") = Request("pedidoWRK")
    Dados("motivo") = Request("motivo")
    Dados("motivoDescricao") = Request("motivoDescricao")
    Dados("precisaLiberacao") = Request("precisaLiberacao")
    Dados("lotePilotoAprovado") = Request("lotePilotoAprovado")
    Dados("tipoLP") = Request("tipoLP")
    Dados("status") = Request("status")
    Dados("acompanhamentoOnduladeira") = Request("acompanhamentoOnduladeira")
    Dados("acompanhamentoImpressora") = Request("acompanhamentoImpressora")
    Dados("acompanhamentoExpedicao") = Request("acompanhamentoExpedicao")
    Dados("acompanhamentoCliente") = Request("acompanhamentoCliente")
    Dados("acompanhamentoDDP") = Request("acompanhamentoDDP")
    Dados("acompanhamentoClicheria") = Request("acompanhamentoClicheria")
    Dados("produtoAnterior") = Request("produtoAnterior")
    Dados("justificativasJSON") = Request("justificativas")

    jaExiste = VerificarLotePilotoJaExiste(Dados("codigoPlanta"), Dados("codigoProduto"))
    'If jaExiste Then
        'Criar = "{""sucesso"": false, ""mensagem"": ""Já existe um lote piloto para este produto""}"
        'Exit Function
    'End If

    SQL = "BEGIN TRANSACTION; "
    SQL = SQL & "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @DadosLotePiloto TABLE (Id int); "

    SQL = SQL & "INSERT INTO LOTE_PILOTO ("
    SQL = SQL & "   PLANTA, "
    SQL = SQL & "   CLIENTE, "
    SQL = SQL & "   LOJA, "
    SQL = SQL & "   NOME_CLIENTE, "
    SQL = SQL & "   PRODUTO, "
    SQL = SQL & "   DESCRICAO_PRODUTO, "
    SQL = SQL & "   QUANTIDADE_PALETES, "
    SQL = SQL & "   QUANTIDADE_PEDIDOS, "
    SQL = SQL & "   FLAG_ONDULADEIRA, "
    SQL = SQL & "   FLAG_IMPRESSORA, "
    SQL = SQL & "   FLAG_EXPEDICAO, "
    SQL = SQL & "   FLAG_CLIENTE, "
    SQL = SQL & "   FLAG_DDP, "
    SQL = SQL & "   FLAG_CLICHERIA, "
    SQL = SQL & "   MOTIVO, "
    SQL = SQL & "   MOTIVO_DESCRICAO, "
    SQL = SQL & "   QUEM_CADASTROU, "
    SQL = SQL & "   DATA_HORA, "
    SQL = SQL & "   LIBERACAO, "
    SQL = SQL & "   APROVACAO, "
    SQL = SQL & "   DATA_APROVACAO, "
    SQL = SQL & "   STATUS, "
    SQL = SQL & "   TIPO_LP, "
    SQL = SQL & "   PEDIDO_WRK, "
    SQL = SQL & "   PRODUTO_ANTERIOR"
    SQL = SQL & ") OUTPUT INSERTED.ID INTO @DadosLotePiloto VALUES ("
    SQL = SQL & "'" & Dados("codigoPlanta") & "', "
    SQL = SQL & "'" & Dados("codigoCliente") & "', "
    SQL = SQL & "'" & Dados("codigoLoja") & "', "
    SQL = SQL & "'" & Dados("nomeCliente") & "', "
    SQL = SQL & "'" & Dados("codigoProduto") & "', "
    SQL = SQL & "'" & Dados("referencia") & "', "
    SQL = SQL & "'" & Dados("quantidadePaletes") & "', "
    SQL = SQL & "'" & Dados("quantidadePedidos") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoOnduladeira") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoImpressora") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoExpedicao") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoCliente") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoDDP") & "', "
    SQL = SQL & "'" & Dados("acompanhamentoClicheria") & "', "
    SQL = SQL & "'" & Dados("motivo") & "', "
    SQL = SQL & "'" & Dados("motivoDescricao") & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "'" & Dados("precisaLiberacao") & "', "
    SQL = SQL & "'" & Dados("lotePilotoAprovado") & "', "
    SQL = SQL & "NULL, "
    SQL = SQL & "'1', "
    SQL = SQL & "'" & Dados("tipoLP") & "', "
    SQL = SQL & "'" & Dados("pedidoWRK") & "', "
    SQL = SQL & "'" & Dados("produtoAnterior") & "' "
    SQL = SQL & "); "

    ' Justificativas (se houver)
    Dados("justificativasJSON") = Replace(Dados("justificativasJSON"), "[", "")
    Dados("justificativasJSON") = Replace(Dados("justificativasJSON"), "]", "")
    justificativasArray = Split(Dados("justificativasJSON"), "},{")
    
    valoresJustificativas = ""
    For i = 0 To UBound(justificativasArray)
        justificativa = justificativasArray(i)

        tipoJustificativa = ExtrairValorJSON(justificativa, "tipoJustificativa")
        descricaoJustificativa = ExtrairValorJSON(justificativa, "descricaoJustificativa")

        If i > 0 Then valoresJustificativas = valoresJustificativas & "," End If

        valoresJustificativas = valoresJustificativas & "("
        valoresJustificativas = valoresJustificativas & "(SELECT Id FROM @DadosLotePiloto), "
        valoresJustificativas = valoresJustificativas & "'" & tipoJustificativa & "', "
        valoresJustificativas = valoresJustificativas & "'" & descricaoJustificativa & "', "
        valoresJustificativas = valoresJustificativas & "NULL, "
        valoresJustificativas = valoresJustificativas & "NULL, "
        valoresJustificativas = valoresJustificativas & "NULL,"
        valoresJustificativas = valoresJustificativas & "GETDATE(), "  
        valoresJustificativas = valoresJustificativas & "'" & usuarioLogado("id") & "'"
        valoresJustificativas = valoresJustificativas & "); "
    Next

    If valoresJustificativas <> "" Then
        SQL = SQL & "INSERT INTO LOTE_PILOTO_JUSTIFICATIVA ("
        SQL = SQL & "   IDLOTEPILOTO, TIPO_JUSTIFICATIVA, JUSTIFICATIVA, "
        SQL = SQL & "   EXCLUIDO, DATA_EXCLUSAO, QUEM_EXCLUIU, DATA_INCLUSAO, USUARIO "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresJustificativas & "; "
    End If

    SQL = SQL & "INSERT INTO LOTE_PILOTO_HISTORICO_STATUS (IDLOTEPILOTO, STATUS_ANTIGO, STATUS_NOVO, ALTERADO_POR, DATA_INCLUSAO, USUARIO "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "(SELECT Id FROM @DadosLotePiloto), "
    SQL = SQL & "NULL, "
    SQL = SQL & "'" & ObterDescricaoStatus(Dados("status")) & "', "
    SQL = SQL & "'LOTE PILOTO', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "'" & usuarioLogado("id") & "'"
    SQL = SQL & "); "

    SQL = SQL & "SELECT Id FROM @DadosLotePiloto; "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL &   "ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL &   "COMMIT TRANSACTION "

    On Error Resume Next
    SET rsCriar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar Lote Piloto: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    id = rsCriar("ID")
    
    rsCriar.Close()
    SET rsCriar = NOTHING

    Call RegistrarLog("criacao", id)
    Criar = "{""sucesso"": true, ""mensagem"": ""Lote Piloto criado com sucesso!""}"
End Function

Function Editar()
    Set Dados = Server.CreateObject("Scripting.Dictionary")

    Dados("id") = Request("ID")
    Dados("quantidadePaletes") = Request("quantidadePaletes")
    Dados("quantidadePedidos") = Request("quantidadePedidos")
    Dados("pedidoWRK") = Request("pedidoWRK")
    Dados("motivo") = Request("motivo")
    Dados("motivoDescricao") = Request("motivoDescricao")
    Dados("precisaLiberacao") = Request("precisaLiberacao")
    Dados("aprovandoLotePiloto") = Request("aprovandoLotePiloto")
    Dados("lotePilotoAprovado") = Request("lotePilotoAprovado")
    Dados("tipoLP") = Request("tipoLP")
    Dados("status") = Request("status")
    Dados("statusAnterior") = Request("statusAnterior")
    Dados("acompanhamentoOnduladeira") = Request("acompanhamentoOnduladeira")
    Dados("acompanhamentoImpressora") = Request("acompanhamentoImpressora")
    Dados("acompanhamentoExpedicao") = Request("acompanhamentoExpedicao")
    Dados("acompanhamentoCliente") = Request("acompanhamentoCliente")
    Dados("acompanhamentoDDP") = Request("acompanhamentoDDP")
    Dados("acompanhamentoClicheria") = Request("acompanhamentoClicheria")
    Dados("produtoAnterior") = Request("produtoAnterior")
    Dados("justificativasJSON") = Request("justificativas")

    ' Se estiver alterando para lote piloto aprovado e status atual for diferente de 4 (Em Teste)
    If Dados("aprovandoLotePiloto") = "1" AND Dados("status") <> "4" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Só é possível marcar lote piloto aprovado quando status é Em Teste Cliente.""}"
        Exit Function
    End If

    SQL = "BEGIN TRANSACTION; "

    ' Insere na tabela de log de campos as alterações
    SQL = SQL & "INSERT INTO LOG_LOTE_PILOTO ("
    SQL = SQL & "   IDLOTEPILOTO, "
    SQL = SQL & "   LOG_SOURCE, "
    SQL = SQL & "   LOG_FIELD_NAME, "
    SQL = SQL & "   LOG_FIELD_TYPE, "
    SQL = SQL & "   LOG_FIELD_VALUE_OLD, "
    SQL = SQL & "   LOG_FIELD_VALUE_NEW, "
    SQL = SQL & "   LOG_FIELD_RECNO, "
    SQL = SQL & "   LOG_USER_ID, "
    SQL = SQL & "   LOG_USER_NAME, "
    SQL = SQL & "   LOG_HTTP_NAME, "
    SQL = SQL & "   LOG_OPERATION, "
    SQL = SQL & "   LOG_DATETIME "
    SQL = SQL & ") "
    SQL = SQL & "SELECT "
    SQL = SQL & "   '" & Dados("id") & "', "
    SQL = SQL & "   'LOTEPILOTOREPOSITORIO', "
    SQL = SQL & "   CAMPOS.NOME_CAMPO, "
    SQL = SQL & "   CAMPOS.TIPO_CAMPO, "
    SQL = SQL & "   CAMPOS.VALOR_ANTIGO, "
    SQL = SQL & "   CAMPOS.VALOR_NOVO, "
    SQL = SQL & "   NULL, "
    SQL = SQL & "   '" & SanitizeSQL(usuarioLogado("id")) & "', "
    SQL = SQL & "   '" & SanitizeSQL(usuarioLogado("login")) & "', "
    SQL = SQL & "   '" & Iif(Session("Ambiente") = "HOMOLOGACAO", "BR-PIR1-M012", "BR-PIR1-M010") & "', "
    SQL = SQL & "   'U', "
    SQL = SQL & "   GETDATE() "
    SQL = SQL & "FROM LOTE_PILOTO LP WITH (UPDLOCK, HOLDLOCK) "
    SQL = SQL & "CROSS APPLY (VALUES "

    SQL = SQL & "   ('QUANTIDADE_PALETES', 'NUMERIC', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.QUANTIDADE_PALETES), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("quantidadePaletes")) & "')), "

    SQL = SQL & "   ('QUANTIDADE_PEDIDOS', 'NUMERIC', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.QUANTIDADE_PEDIDOS), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("quantidadePedidos")) & "')), "

    SQL = SQL & "   ('FLAG_ONDULADEIRA', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_ONDULADEIRA), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoOnduladeira")) & "')), "

    SQL = SQL & "   ('FLAG_IMPRESSORA', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_IMPRESSORA), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoImpressora")) & "')), "

    SQL = SQL & "   ('FLAG_EXPEDICAO', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_EXPEDICAO), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoExpedicao")) & "')), "

    SQL = SQL & "   ('FLAG_CLIENTE', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_CLIENTE), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoCliente")) & "')), "

    SQL = SQL & "   ('FLAG_DDP', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_DDP), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoDDP")) & "')), "

    SQL = SQL & "   ('FLAG_CLICHERIA', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.FLAG_CLICHERIA), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("acompanhamentoClicheria")) & "')), "

    SQL = SQL & "   ('MOTIVO', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.MOTIVO), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("motivo")) & "')), "

    SQL = SQL & "   ('MOTIVO_DESCRICAO', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.MOTIVO_DESCRICAO), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("motivoDescricao")) & "')), "

    SQL = SQL & "   ('LIBERACAO', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.LIBERACAO), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("precisaLiberacao")) & "')), "

    SQL = SQL & "   ('APROVACAO', 'BOOLEAN', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.APROVACAO), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("lotePilotoAprovado")) & "')), "

    SQL = SQL & "   ('STATUS', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.STATUS), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("status")) & "')), "

    SQL = SQL & "   ('TIPO_LP', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.TIPO_LP), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("tipoLP")) & "')), "

    SQL = SQL & "   ('PEDIDO_WRK', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.PEDIDO_WRK), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("pedidoWRK")) & "')), "

    SQL = SQL & "   ('PRODUTO_ANTERIOR', 'VARCHAR', "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), LP.PRODUTO_ANTERIOR), "
    SQL = SQL & "       CONVERT(NVARCHAR(MAX), N'" & SanitizeSQL(Dados("produtoAnterior")) & "')) "

    SQL = SQL & ") CAMPOS ("
    SQL = SQL & "   NOME_CAMPO, "
    SQL = SQL & "   TIPO_CAMPO, "
    SQL = SQL & "   VALOR_ANTIGO, "
    SQL = SQL & "   VALOR_NOVO "
    SQL = SQL & ") "
    SQL = SQL & "WHERE LP.ID = '" & SanitizeSQL(Dados("id")) & "' "
    SQL = SQL & "AND ISNULL(CAMPOS.VALOR_ANTIGO, N'{NULL}') "
    SQL = SQL & "    <> ISNULL(CAMPOS.VALOR_NOVO, N'{NULL}'); "

    ' Atualiza dados tabela principal
    SQL = SQL & "UPDATE LOTE_PILOTO SET "
    SQL = SQL & "   QUANTIDADE_PALETES = '" & SanitizeSQL(Dados("quantidadePaletes")) & "', "
    SQL = SQL & "   QUANTIDADE_PEDIDOS = '" & SanitizeSQL(Dados("quantidadePedidos")) & "', "
    SQL = SQL & "   FLAG_ONDULADEIRA = '" & SanitizeSQL(Dados("acompanhamentoOnduladeira")) & "', "
    SQL = SQL & "   FLAG_IMPRESSORA = '" & SanitizeSQL(Dados("acompanhamentoImpressora")) & "', "
    SQL = SQL & "   FLAG_EXPEDICAO = '" & SanitizeSQL(Dados("acompanhamentoExpedicao")) & "', "
    SQL = SQL & "   FLAG_CLIENTE = '" & SanitizeSQL(Dados("acompanhamentoCliente")) & "', "
    SQL = SQL & "   FLAG_DDP = '" & SanitizeSQL(Dados("acompanhamentoDDP")) & "', "
    SQL = SQL & "   FLAG_CLICHERIA = '" & SanitizeSQL(Dados("acompanhamentoClicheria")) & "', "
    SQL = SQL & "   MOTIVO = '" & SanitizeSQL(Dados("motivo")) & "', "
    SQL = SQL & "   MOTIVO_DESCRICAO = '" & SanitizeSQL(Dados("motivoDescricao")) & "', "
    SQL = SQL & "   LIBERACAO = '" & SanitizeSQL(Dados("precisaLiberacao")) & "', "
    SQL = SQL & "   APROVACAO = '" & SanitizeSQL(Dados("lotePilotoAprovado")) & "', "
    SQL = SQL & "   DATA_APROVACAO = " & SanitizeSQL(Iif(Dados("lotePilotoAprovado") = "1", "GETDATE()", "NULL")) & ", "
    SQL = SQL & "   STATUS = '" & SanitizeSQL(Dados("status")) & "', "
    SQL = SQL & "   TIPO_LP = '" & SanitizeSQL(Dados("tipoLP")) & "', "
    SQL = SQL & "   PEDIDO_WRK = '" & SanitizeSQL(Dados("pedidoWRK")) & "', "
    SQL = SQL & "   PRODUTO_ANTERIOR = '" & SanitizeSQL(Dados("produtoAnterior")) & "' "
    SQL = SQL & "WHERE ID = '" & SanitizeSQL(Dados("id")) & "'; "

    ' Insere Justificativas, se houver
    valoresJustificativas = ""

    If Trim(Dados("justificativasJSON")) <> "" Then
        Dados("justificativasJSON") = Replace(Dados("justificativasJSON"), "[", "")
        Dados("justificativasJSON") = Replace(Dados("justificativasJSON"), "]", "")
        justificativasArray = Split(Dados("justificativasJSON"), "},{")

        For i = 0 To UBound(justificativasArray)
            justificativa = justificativasArray(i)
            tipoJustificativa = ExtrairValorJSON(justificativa, "tipoJustificativa")
            descricaoJustificativa = ExtrairValorJSON(justificativa, "descricaoJustificativa")

            If i > 0 Then valoresJustificativas = valoresJustificativas & "," End If

            valoresJustificativas = valoresJustificativas & "("
            valoresJustificativas = valoresJustificativas & "'" & SanitizeSQL(Dados("id")) & "', "
            valoresJustificativas = valoresJustificativas & "'" & SanitizeSQL(tipoJustificativa) & "', "
            valoresJustificativas = valoresJustificativas & "'" & SanitizeSQL(descricaoJustificativa) & "', "
            valoresJustificativas = valoresJustificativas & "NULL, "
            valoresJustificativas = valoresJustificativas & "NULL, "
            valoresJustificativas = valoresJustificativas & "NULL, "
            valoresJustificativas = valoresJustificativas & "GETDATE(), "
            valoresJustificativas = valoresJustificativas & "'" & SanitizeSQL(usuarioLogado("id")) & "'"
            valoresJustificativas = valoresJustificativas & ")"
        Next
    End If

    If valoresJustificativas <> "" Then
        SQL = SQL & "INSERT INTO LOTE_PILOTO_JUSTIFICATIVA ("
        SQL = SQL & "   IDLOTEPILOTO, "
        SQL = SQL & "   TIPO_JUSTIFICATIVA, "
        SQL = SQL & "   JUSTIFICATIVA, "
        SQL = SQL & "   EXCLUIDO, "
        SQL = SQL & "   DATA_EXCLUSAO, "
        SQL = SQL & "   QUEM_EXCLUIU, "
        SQL = SQL & "   DATA_INCLUSAO, "
        SQL = SQL & "   USUARIO "
        SQL = SQL & ") VALUES "
        SQL = SQL & valoresJustificativas & "; "
    End If

    ' Histórico de alteração de status
    If Dados("statusAnterior") <> Dados("status") Then
        SQL = SQL & "INSERT INTO LOTE_PILOTO_HISTORICO_STATUS ("
        SQL = SQL & "   IDLOTEPILOTO, "
        SQL = SQL & "   STATUS_ANTIGO, "
        SQL = SQL & "   STATUS_NOVO, "
        SQL = SQL & "   ALTERADO_POR, "
        SQL = SQL & "   DATA_INCLUSAO, "
        SQL = SQL & "   USUARIO "
        SQL = SQL & ") VALUES ("
        SQL = SQL & "'" & SanitizeSQL(Dados("id")) & "', "
        SQL = SQL & "'" & SanitizeSQL(ObterDescricaoStatus(Dados("statusAnterior"))) & "', "
        SQL = SQL & "'" & SanitizeSQL(ObterDescricaoStatus(Dados("status"))) & "', "
        SQL = SQL & "'LOTE PILOTO', "
        SQL = SQL & "GETDATE(), "
        SQL = SQL & "'" & SanitizeSQL(usuarioLogado("id")) & "'"
        SQL = SQL & "); "
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION; "

    On Error Resume Next
    Set rsEditar = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsEditar) Then
        On Error GoTo 0
        Editar = "{""sucesso"": false, " & """mensagem"": ""Erro ao editar Lote Piloto: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Call RegistrarLog("edicao", Dados("id"))
    Editar = "{""sucesso"": true, " & """mensagem"": ""Lote Piloto atualizado com sucesso!""}"
End Function

Function Excluir()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")
    Dados("codigoPlanta") = Request("codigoPlanta")
    Dados("codigoProduto") = Request("codigoProduto")
    Dados("descricaoJustificativa") = Request("descricaoJustificativa")

    jaExiste = VerificarExistePedido(Dados("codigoPlanta"), Dados("codigoProduto"), Dados("id"))
    If jaExiste Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Já existe um pedido criado, não é possível excluir o lote piloto.""}"
        Exit Function
    End If

    SQL = "UPDATE LOTE_PILOTO SET "
    SQL = SQL & "   STATUS = '99', "
    SQL = SQL & "   QUEM_EXCLUIU = " & usuarioLogado("id") & ", "
    SQL = SQL & "   DATA_EXCLUSAO = GETDATE() "
    SQL = SQL & "WHERE ID = '" & Dados("id") & "'; "

    SQL = SQL & "INSERT INTO LOTE_PILOTO_JUSTIFICATIVA ("
    SQL = SQL & "   IDLOTEPILOTO, TIPO_JUSTIFICATIVA, JUSTIFICATIVA, "
    SQL = SQL & "   EXCLUIDO, DATA_EXCLUSAO, QUEM_EXCLUIU, DATA_INCLUSAO, USUARIO "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'EX', "
    SQL = SQL & "'" & Dados("descricaoJustificativa") & "', "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL,"
    SQL = SQL & "GETDATE(), "  
    SQL = SQL & "'" & usuarioLogado("id") & "'"
    SQL = SQL & "); "

    SQL = SQL & "INSERT INTO LOTE_PILOTO_HISTORICO_STATUS ("
    SQL = SQL & "   IDLOTEPILOTO, STATUS_ANTIGO, STATUS_NOVO, "
    SQL = SQL & "   ALTERADO_POR, DATA_INCLUSAO, USUARIO"
    SQL = SQL & ") "
    SQL = SQL & " VALUES ("
    SQL = SQL & "'" & Dados("id") & "', "
    SQL = SQL & "'" & Dados("statusAnterior") & "', "
    SQL = SQL & "'EXCLUIDO', "
    SQL = SQL & "'LOTE PILOTO', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "'" & usuarioLogado("id") & "')"
  
    On Error Resume Next
    SET rsExcluir = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao Excluir Lote Piloto: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Call RegistrarLog("exclusao", Dados("codigo"))
    Excluir = "{""sucesso"": true, ""mensagem"": ""Lote Piloto excluído com sucesso!""}"
End Function

Function ListarClientes()
    codigoPlanta = request("codigoPlanta")
    codigoCliente = request("codigoCliente")
    codigoLoja = request("codigoLoja")
    busca = Iif(request("busca") = "", "", trim(request("busca")))
    
    If codigoPlanta = "" Then
        ListarClientes = "{""sucesso"": false, ""mensagem"": ""Nenhuma planta informada.""}"
        Exit Function
    End If

    ehPlantaWestRock = VerificarEhPlantaWestrock(codigoPlanta)
    nomePlanta = ObterNomePlanta(codigoPlanta)
    
    If ehPlantaWestRock Then
        nomeTabela = "VW_CLIENTES_WRK"
    Else
        nomeTabela = IIf(atualizandoBD(), "VW_CLIENTES_TEMP", "VW_CLIENTES")
    End If

    SQL = "SELECT "
    SQL = SQL & "   CIDADE as NomePlanta, "
    SQL = SQL & "   CASE WHEN CIDADE IN ('MARANGUAPE','BENTO') THEN CODIGO_ERP ELSE COD_CLIENTE END as CodigoCliente, "
    SQL = SQL & "   LOJA_CLIENTE AS CodigoLoja, "
    SQL = SQL & "   NOME_CLIENTE as Nome, "
    SQL = SQL & "   FANTASIA_CLIENTE as Fantasia, "
    SQL = SQL & "   MUNICIPIO_CLIENTE as Municipio, "
    SQL = SQL & "   COD_MUNICIPIOS_IBGE as CodMunicipio, "
    SQL = SQL & "   ESTADO_CLIENTE as UF, "
    SQL = SQL & "   COD_VENDEDOR as CodVendedor, "
    SQL = SQL & "   NOME_VENDEDOR as NomeVendedor, "
    SQL = SQL & "   STATUS as Status, "
    SQL = SQL & "   CNPJ_CLIENTE as Cnpj, "
    SQL = SQL & "   BAIRRO AS Bairro "
    SQL = SQL & "FROM " & nomeTabela & " "
    SQL = SQL & "WHERE 1=1"
    
    If busca <> "" Then
        SQL = SQL & " AND ("
        SQL = SQL & "  NOME_CLIENTE like '%" & Ucase(busca) & "%' "
        SQL = SQL & "  OR CASE WHEN CIDADE IN ('PIRAPETINGA','UBERABA','SAQUAREMA') THEN CAST(COD_CLIENTE as varchar(20)) ELSE CODIGO_ERP END LIKE '%" & busca & "%' "
        SQL = SQL & ")"
    End If

    If nomePlanta <> "" Then
        SQL = SQL & "AND CIDADE ='" & nomePlanta & "' "
    End If

    If codigoCliente <> "" Then
        SQL = SQL & "AND (COD_CLIENTE ='" & codigoCliente & "' OR CODIGO_ERP ='" & codigoCliente & "') "
    End If

    If codigoLoja <> "" Then
        SQL = SQL & "AND LOJA_CLIENTE ='" & codigoLoja & "' "
    End If

    If usuarioLogado("nivel") <> 0 then
        SQL = SQL & " AND COD_VENDEDOR = '" & usuarioLogado("cod") & "' "
    End If

    SQL = SQL & " ORDER BY NOME_CLIENTE "

    On Error Resume Next
    SET rsClientes = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsClientes) Then
        ListarClientes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsClientes.EOF Then
        Do While Not rsClientes.EOF            
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """nomePlanta"": """ & EscapeJSON(rsClientes("NomePlanta")) & """, "
            dadosJSON = dadosJSON & """codigoCliente"": """ & EscapeJSON(rsClientes("CodigoCliente")) & """, "
            dadosJSON = dadosJSON & """codigoLoja"": """ & EscapeJSON(rsClientes("CodigoLoja")) & """, "
            dadosJSON = dadosJSON & """nome"": """ & EscapeJSON(rsClientes("Nome")) & """, "
            dadosJSON = dadosJSON & """fantasia"": """ & EscapeJSON(rsClientes("Fantasia")) & """, "            
            dadosJSON = dadosJSON & """municipio"": """ & EscapeJSON(rsClientes("Municipio")) & """, "
            dadosJSON = dadosJSON & """codMunicipio"": """ & EscapeJSON(rsClientes("CodMunicipio")) & """, "
            dadosJSON = dadosJSON & """UF"": """ & EscapeJSON(rsClientes("UF")) & """, "
            dadosJSON = dadosJSON & """codVendedor"": """ & EscapeJSON(rsClientes("CodVendedor")) & """, "
            dadosJSON = dadosJSON & """nomeVendedor"": """ & EscapeJSON(rsClientes("NomeVendedor")) & """, "
            dadosJSON = dadosJSON & """bairro"": """ & EscapeJSON(rsClientes("Bairro")) & """, "
            dadosJSON = dadosJSON & """cnpj"": """ & EscapeJSON(rsClientes("Cnpj")) & """ "
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsClientes.MoveNext()
        Loop
    End If

    dadosJSON = dadosJSON & "]"
    rsClientes.Close()
    SET rsClientes = NOTHING

    ListarClientes = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ListarStatus()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aguardando Pedido LP" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Convertido em Venda" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aprovado sem Pedido de Conversão" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Em Teste Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "5" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Reprovado pelo Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "6" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Reprovado Internamente/Retrabalho" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "7" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aguardando Produção" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "8" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Descontinuado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "9" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sem Retorno Vendas/Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "10" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Processo Segue em Novo Produto/Ficha Técnica" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "11" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Validação Máquina Nova/Homologação" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "99" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Excluído" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposLotePiloto()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","
    
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Incremento de Volume" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Manutenção de Volume" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposLotePiloto = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarMotivos()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Inovação" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","
    
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "2" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Melhoria Significativa  - Processo SW e Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Melhoria - Paletização e Transporte" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - SW + Cliente  - Acompanhamento detalhado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "5" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - SW - Acompanhamento detalhado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "6" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - Layout de Impressão" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "7" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - Dimensional" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "8" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - Padrão de Cor / Cromia / Reticula" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "9" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Novo SKU - Paletização e Transporte" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarMotivos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposItem()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CX" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CAIXA" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","
    
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CH" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CHAPA" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CJ" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CONJUNTO" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposItem = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistoricoStatus()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")

    SQL = "SELECT "
    SQL = SQL & "   l.IDHISTORICO_STATUS AS Id, "
    SQL = SQL & "   l.IDLOTEPILOTO AS IdLotePiloto, "	
    SQL = SQL & "   l.STATUS_ANTIGO AS StatusAnterior, "
    SQL = SQL & "   l.STATUS_NOVO AS StatusNovo, "	
    SQL = SQL & "   l.ALTERADO_POR AlteradoPor, "
    SQL = SQL & "   l.USUARIO AS IdUsuarioInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, "
    SQL = SQL & "   l.DATA_INCLUSAO as DataInclusao "
    SQL = SQL & "FROM LOTE_PILOTO_HISTORICO_STATUS l "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c on c.CRM_R_E_C_N_O = l.USUARIO "
    SQL = SQL & "WHERE l.IDLOTEPILOTO = '" & Dados("id") & "' "

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistoricoStatus = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsHistorico("Id")) & """, "
            jsonArray = jsonArray & """idLotePiloto"": """ & EscapeJSON(rsHistorico("IdLotePiloto")) & """, "
            jsonArray = jsonArray & """statusAnterior"": """ & EscapeJSON(rsHistorico("StatusAnterior")) & """, "
            jsonArray = jsonArray & """statusNovo"": """ & EscapeJSON(rsHistorico("StatusNovo")) & """, "
            jsonArray = jsonArray & """alteradoPor"": """ & EscapeJSON(rsHistorico("AlteradoPor")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsHistorico("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsHistorico("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistoricoStatus = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarHistoricoJustificativas()
    Set Dados = Server.CreateObject("Scripting.Dictionary")
    Dados("id") = Request("id")

    SQL = "SELECT "
    SQL = SQL & "   l.IDLOTEPILOTOJUSTIFICATIVA AS Id, "
    SQL = SQL & "   l.IDLOTEPILOTO AS IdLotePiloto, "	
    SQL = SQL & "   CASE l.TIPO_JUSTIFICATIVA "
    SQL = SQL & "       WHEN '1' THEN 'Pedido' "
    SQL = SQL & "       WHEN 'EX' THEN 'Exclusão' "
    SQL = SQL & "       ELSE 'Status' "
    SQL = SQL & "   END AS TipoJustificativa, "
    SQL = SQL & "   l.JUSTIFICATIVA AS DescricaoJustificativa, "	
    SQL = SQL & "   l.USUARIO AS IdUsuarioInclusao, "
    SQL = SQL & "   c.CRM_NOME AS NomeUsuarioInclusao, "
    SQL = SQL & "   l.DATA_INCLUSAO as DataInclusao "
    SQL = SQL & "FROM LOTE_PILOTO_JUSTIFICATIVA l "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI c on c.CRM_R_E_C_N_O = l.USUARIO "
    SQL = SQL & "WHERE l.IDLOTEPILOTO = '" & Dados("id") & "' "
    SQL = SQL & "AND EXCLUIDO IS NULL "

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsHistorico) Then
        ListarHistoricoJustificativas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar histórico de aprovações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsHistorico.EOF Then
        Do While Not rsHistorico.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsHistorico("Id")) & """, "
            jsonArray = jsonArray & """idLotePiloto"": """ & EscapeJSON(rsHistorico("IdLotePiloto")) & """, "
            jsonArray = jsonArray & """tipoJustificativa"": """ & EscapeJSON(rsHistorico("TipoJustificativa")) & """, "
            jsonArray = jsonArray & """descricaoJustificativa"": """ & EscapeJSON(rsHistorico("DescricaoJustificativa")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsHistorico("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsHistorico("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsHistorico("DataInclusao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsHistorico.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsHistorico.Close()
    SET rsHistorico = NOTHING
    
    ListarHistoricoJustificativas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterDetalhesPedido()
    codigoPlanta = Trim(Request("codigoPlanta"))
    codigoCliente = Trim(Request("codigoCliente"))
    codigoLoja = Trim(Request("codigoLoja"))
    codigoProduto = Trim(Request("codigoProduto"))
    quantidadePedidos = Trim(Request("quantidadePedidos"))
    nomePlanta = ObterNomePlanta(codigoPlanta)
    nomeCliente = ObterNomeCliente(codigoPlanta, codigoCliente, codigoLoja)

    If codigoPlanta = "" Then
        ObterDetalhesPedido = "{""sucesso"": false, ""mensagem"": ""Erro ao listar detalhes de pedido: planta não informada""}"
        Exit Function
    End If

    nomeConexao = ""

    Select Case codigoPlanta
        Case "00", "04"
            SQL = "SELECT "
            SQL = SQL & "   COUNT(CASE WHEN C5_LOTEPIL = 'S' THEN C5_NUM END) AS PedidosComLotePiloto,  "
            SQL = SQL & "   ISNULL(SUM(CASE WHEN C5_LOTEPIL = 'S' THEN C6_QTDVEN END), 0) AS QuantidadeComLotePiloto,  "
            SQL = SQL & "   COUNT(CASE WHEN C5_LOTEPIL = 'N' THEN C5_NUM END) AS PedidosSemLotePiloto,  "
            SQL = SQL & "   ISNULL(SUM(CASE WHEN C5_LOTEPIL = 'N' THEN C6_QTDVEN END), 0) AS QuantidadeSemLotePiloto,  "
            SQL = SQL & "   COUNT(C5_NUM) AS TotalPedidos,  "
            SQL = SQL & "   ISNULL(SUM(C6_QTDVEN), 0) AS TotalQuantidade " 
            SQL = SQL & "FROM SC5010 "
            SQL = SQL & "LEFT JOIN SC6010 ON C6_NUM = C5_NUM "
            SQL = SQL & "AND SC6010.D_E_L_E_T_ = '' "
            SQL = SQL & "WHERE C6_PRODUTO = '" & codigoProduto & "' "
            SQL = SQL & "   AND C5_CLIENTE = '" & codigoCliente & "'"
            SQL = SQL & "   AND C5_LOJACLI = '" & codigoLoja & "'" 
            SQL = SQL & "   AND SC5010.D_E_L_E_T_ = ''"
            SQL = SQL & "   AND C5_ESTATUS <> 'C'"
            nomeConexao = "PIR-DADOSADV"

        Case "05",  "06",  "07" 
            SQL = "SELECT "
            SQL = SQL & "     COUNT(CASE WHEN LotePiloto = '1' THEN Pedido END)  AS PedidosComLotePiloto, "
            SQL = SQL & "     ISNULL(SUM(CASE WHEN LotePiloto = '1' THEN Quantidade END),0) AS QuantidadeComLotePiloto, "
            SQL = SQL & "     COUNT(CASE WHEN LotePiloto = '0' THEN Pedido END) AS PedidosSemLotePiloto, "
            SQL = SQL & "     ISNULL(SUM(CASE WHEN LotePiloto = '0' THEN Quantidade END),0)  AS QuantidadeSemLotePiloto, "
            SQL = SQL & "     COUNT(Pedido) AS TotalPedidos, "
            SQL = SQL & "     ISNULL(SUM(Quantidade),0) AS TotalQuantidade "
            SQL = SQL & " FROM Pedido S "
            SQL = SQL & " OUTER APPLY(SELECT top 1 StatusPedido FROM TRIMBOX..Pedidos T WHERE S.Pedido = T.Pedido_ERP COLLATE Latin1_General_CI_AS) Ped "
            SQL = SQL & " WHERE Produto = '" & codigoProduto & "' "
            SQL = SQL & " AND Cliente = '" & codigoCliente & "' "
            SQL = SQL & " AND ISNULL(Status,0) = 1 "
            SQL = SQL & " AND ISNULL(Excluido,0) = 0 "
            SQL = SQL & " AND StatusPedido <> '4' "

            If codigoPlanta = "05" Then 
                nomeConexao = "BEN-CORPORATIVO"
            ElseIf codigoPlanta = "06" Then  
                nomeConexao = "FOR-CORPORATIVO"
            ElseIf codigoPlanta = "07" Then     
                nomeConexao = "SAQ-CORPORATIVO"
            End If 
        Case "08", "09","10","11","12","13"               
            SQL = "WITH PEDIDOS AS ( "
            SQL = SQL & "   SELECT "
            SQL = SQL & "       PEDIDO, "
            SQL = SQL & "       QUANT_VENDIDA, "
            SQL = SQL & "       ROW_NUMBER() OVER (ORDER BY EMISSAO ASC) MENOR "
            SQL = SQL & "   FROM BI_PEDIDOS_WRK "
            SQL = SQL & "   WHERE "
            SQL = SQL & "   CIDADE = '" & nomePlanta & "'"  
			SQL = SQL & "   AND PRODUTO = '" & codigoProduto & "'"  
			SQL = SQL & " ) "
            SQL = SQL & " SELECT "   
            SQL = SQL & "   COUNT(CASE WHEN MENOR <= " & quantidadePedidos & " THEN 1 END) AS PedidosComLotePiloto, "
            SQL = SQL & "   ISNULL(SUM(CASE WHEN MENOR <= " & quantidadePedidos & " THEN QUANT_VENDIDA END), 0) AS QuantidadeComLotePiloto, "
            SQL = SQL & "   COUNT(CASE WHEN MENOR > " & quantidadePedidos & " THEN 1 END) AS PedidosSemLotePiloto, "
            SQL = SQL & "   ISNULL(SUM(CASE WHEN MENOR > " & quantidadePedidos & " THEN QUANT_VENDIDA END), 0) AS QuantidadeSemLotePiloto, "
            SQL = SQL & "   ("
            SQL = SQL & "       COUNT(CASE WHEN MENOR <= " & quantidadePedidos & " THEN 1 END) "
            SQL = SQL & "       + COUNT(CASE WHEN MENOR > " & quantidadePedidos & " THEN 1 END)"
            SQL = SQL & "   ) AS TotalPedidos "
            SQL = SQL & "   ("
            SQL = SQL & "       ISNULL(SUM(CASE WHEN MENOR <= " & quantidadePedidos & " THEN QUANT_VENDIDA END), 0) "
            SQL = SQL & "       ISNULL(SUM(CASE WHEN MENOR > " & quantidadePedidos & " THEN QUANT_VENDIDA END), 0) "
            SQL = SQL & "   ) AS TotalQuantidade "
            SQL = SQL & "FROM PEDIDOS "
            nomeConexao = "SPO-BI"
    End Select

    On Error Resume Next
    Set rsDetalhes = Conecta(nomeConexao, SQL)

    If Err.Number <> 0 Or Not IsObject(rsDetalhes) Then
        ObterDetalhesPedido = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar produtos: " & EscapeJSON(Err.Description) & """}"
        Err.Clear

        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    dadosJSON = dadosJSON & "{"

    If Not rsDetalhes.EOF Then
        dadosJSON = dadosJSON & """codigoProduto"": """ & EscapeJSON(codigoProduto) & """, "
        dadosJSON = dadosJSON & """codigoCliente"": """ & EscapeJSON(codigoCliente) & """, "
        dadosJSON = dadosJSON & """nomeCliente"": """ & EscapeJSON(nomeCliente) & """, "
        dadosJSON = dadosJSON & """pedidosComLotePiloto"": """ & EscapeJSON(rsDetalhes("PedidosComLotePiloto")) & """, "
        dadosJSON = dadosJSON & """quantidadeComLotePiloto"": """ & EscapeJSON(rsDetalhes("QuantidadeComLotePiloto")) & """, "
        dadosJSON = dadosJSON & """pedidosSemLotePiloto"": """ & EscapeJSON(rsDetalhes("PedidosSemLotePiloto")) & """, "
        dadosJSON = dadosJSON & """quantidadeSemLotePiloto"": """ & EscapeJSON(rsDetalhes("QuantidadeSemLotePiloto")) & """, "
        dadosJSON = dadosJSON & """totalPedidos"": """ & EscapeJSON(rsDetalhes("TotalPedidos")) & """, "
        dadosJSON = dadosJSON & """totalQuantidade"": """ & EscapeJSON(rsDetalhes("TotalQuantidade")) & """ "
    End If

    dadosJSON = dadosJSON & "}"

    rsDetalhes.Close
    Set rsDetalhes = Nothing

    ObterDetalhesPedido = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ObterNomeCliente(codigoPlanta, codigoCliente, codigoLoja)
    SQL = "SELECT A1_NOME AS Nome FROM SA1010 "
    SQL = SQL & "WHERE A1_COD = '" & codigoCliente & "' "
    SQL = SQL & "AND A1_LOJA = '" & codigoLoja & "' "

    On Error Resume Next
    SET rsCliente = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCliente) Then
        ObterNomeCliente = ""
        Exit Function
    End If
    On Error Goto 0

    If Not rsCliente.EOF Then
        nomeCliente = rsCliente("Nome")
    End If

    rsCliente.Close()
    SET rsCliente = NOTHING
    
    ObterNomeCliente = nomeCliente
End Function

Function ListarProdutos()
    codigoPlanta = Trim(Request("codigoPlanta"))
    codigoCliente = Trim(Request("codigoCliente"))
    codigoLoja = Trim(Request("codigoLoja"))
    codigoProduto = Trim(Request("codigoProduto"))
    tipoItem = Trim(Request("tipoItem"))
    busca = Trim(Request("busca"))

    If codigoPlanta = "" Then
        ListarProdutos = "{""sucesso"": false, ""mensagem"": ""Erro ao listar produtos: planta não informada""}"
        Exit Function
    End If

    filtroProduto = codigoProduto

    If filtroProduto = "" Then
        filtroProduto = busca
    End If

    SQL = ""
    nomeConexao = ""

    Select Case codigoPlanta
        Case "00", "04"
            SQL = SentencaProdutosPirapetinga(codigoCliente, tipoItem, filtroProduto)
            nomeConexao = "PIR-DADOSADV"
        Case "05"
            SQL = SentencaProdutosBento(codigoCliente, tipoItem, filtroProduto)
            nomeConexao = "BEN-TRIMBOX"
        Case "06"
            SQL = SentencaProdutosFortaleza(codigoCliente, codigoLoja, tipoItem, filtroProduto)
            nomeConexao = "FOR-TRIMBOX"
        Case "07"
            SQL = SentencaProdutosSaquarema(codigoCliente, codigoLoja, tipoItem, filtroProduto)
            nomeConexao = "SAQ-TRIMBOX"
        Case Else
            SQL = SentencaProdutosPlantasWRK(codigoPlanta, codigoCliente, filtroProduto)
            nomeConexao = "SPO-BI"
    End Select

    If SQL = "" Then
        ListarProdutos = "{""sucesso"": false, ""mensagem"": ""Erro ao listar produtos: tipo de item inválido""}"
        Exit Function
    End If

    On Error Resume Next
    Set rsProdutos = Conecta(nomeConexao, SQL)

    If Err.Number <> 0 Or Not IsObject(rsProdutos) Then
        ListarProdutos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar produtos: " & EscapeJSON(Err.Description) & """}"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    primeiroItem = True
    dadosJSON = "["

    If Not rsProdutos.EOF Then
        Do While Not rsProdutos.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & ","  End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """codigo"": """ & EscapeJSON(rsProdutos("Codigo")) & """, "
            dadosJSON = dadosJSON & """referencia"": """ & EscapeJSON(rsProdutos("Referencia")) & """"
            dadosJSON = dadosJSON & "}"

            primeiroItem = False
            rsProdutos.MoveNext
        Loop
    End If

    dadosJSON = dadosJSON & "]"

    rsProdutos.Close
    Set rsProdutos = Nothing

    ListarProdutos = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function SentencaProdutosBento(codigoCliente, tipoItem, filtroProduto)
    SQL = ""

    Select Case tipoItem
        Case "CJ"
            SQL = "SELECT TOP 1 " & _
                  "       LTRIM(RTRIM(CODIGO_CONJUNTO)) AS Codigo, " & _
                  "       LTRIM(RTRIM(REFERENCIA_CONJUNTO)) AS Referencia " & _
                  "FROM VW_BI_PRODUTOS_PEDIDO " & _
                  "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
                  "  AND GRUPO_PRODUTO <> 'CHAPA' " & _
                  "  AND (REFERENCIA_CONJUNTO LIKE '%" & filtroProduto & "%' " & _
                  "       OR CODIGO_CONJUNTO = '" & filtroProduto & "') " & _
                  "ORDER BY CODIGO_ITEM"

        Case "CX"
            SQL = "SELECT " & _
                  "       LTRIM(RTRIM(CODIGO_ITEM)) AS Codigo, " & _
                  "       LTRIM(RTRIM(REFERENCIA_ITEM)) AS Referencia " & _
                  "FROM VW_BI_PRODUTOS_PEDIDO " & _
                  "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
                  "  AND GRUPO_PRODUTO <> 'CHAPA' " & _
                  "  AND (REFERENCIA_ITEM LIKE '%" & filtroProduto & "%' " & _
                  "       OR CODIGO_ITEM = '" & filtroProduto & "') " & _
                  "ORDER BY REFERENCIA_ITEM"

        Case "CH"
            SQL = "SELECT " & _
                  "       LTRIM(RTRIM(CODIGO_ITEM)) AS Codigo, " & _
                  "       LTRIM(RTRIM(REFERENCIA_ITEM)) AS Referencia " & _
                  "FROM VW_BI_PRODUTOS_PEDIDO " & _
                  "WHERE CODIGO_CLIENTE = '707' " & _
                  "  AND GRUPO_PRODUTO = 'CHAPA' " & _
                  "  AND (REFERENCIA_ITEM LIKE '%" & filtroProduto & "%' " & _
                  "       OR CODIGO_ITEM = '" & filtroProduto & "') " & _
                  "ORDER BY REFERENCIA_ITEM"
    End Select

    SentencaProdutosBento = SQL
End Function

Function SentencaProdutosPirapetinga(codigoCliente, tipoItem, filtroProduto)
    SQL = ""

    If tipoItem <> "CX" Then
        SQL = SQL & "SELECT DISTINCT "
        SQL = SQL & "       LTRIM(RTRIM(SB1.B1_COD)) AS Codigo, "
        SQL = SQL & "       LTRIM(RTRIM(SB1.B1_DESC)) AS Referencia, "
        SQL = SQL & "       SB1.B1_INPUBE "
        SQL = SQL & "FROM SB1010 SB1 WITH(NOLOCK) "
        SQL = SQL & "LEFT JOIN SB5010 SB5 WITH(NOLOCK) "
        SQL = SQL & "       ON SB5.B5_COD = SB1.B1_COD "
        SQL = SQL & "      AND ISNULL(SB5.D_E_L_E_T_, '') = '' "
        SQL = SQL & "WHERE ISNULL(SB1.D_E_L_E_T_, '') = '' "
        SQL = SQL & "  AND SB1.B1_STATUS IN ('V', 'A', 'W', 'P') "
        SQL = SQL & "  AND SUBSTRING(ISNULL(SB1.B1_COD, ''), 1, 2) <> '77' "

        If tipoItem = "CH" Then
            SQL = SQL & "  AND SB1.B1_GRUPO = '72' "
            SQL = SQL & "  AND (ISNULL(SB5.B5_CLI, '') = '" & codigoCliente & "' "
            SQL = SQL & "       OR ISNULL(SB5.B5_CLI, '') = '') "
            SQL = SQL & "  AND (SB1.B1_COD = '" & filtroProduto & "' "
            SQL = SQL & "       OR UPPER(SB1.B1_DESC) LIKE UPPER('%" & filtroProduto & "%') "
            SQL = SQL & "       OR UPPER(SB1.B1_ESTRUCH) LIKE UPPER('%" & filtroProduto & "%')) "
        End If

        If tipoItem = "CJ" Then
            SQL = SQL & "  AND SB1.B1_GRUPO = '75' "
            SQL = SQL & "  AND (ISNULL(SB5.B5_CLI, '') = '" & codigoCliente & "' "
            SQL = SQL & "       OR ISNULL(SB5.B5_CLI, '') = '') "
            SQL = SQL & "  AND (SB1.B1_COD = '" & filtroProduto & "' "
            SQL = SQL & "       OR UPPER(SB1.B1_DESC) LIKE UPPER('%" & filtroProduto & "%')) "
        End If

        SQL = SQL & "ORDER BY SB1.B1_INPUBE, Descricao"
    Else
        SQL = SQL & "SELECT DISTINCT "
        SQL = SQL & "       LTRIM(RTRIM(SB1.B1_COD)) AS Codigo, "
        SQL = SQL & "       LTRIM(RTRIM(SB5.B5_REF)) AS Referencia "
        SQL = SQL & "FROM SB5010 SB5 WITH(NOLOCK) "
        SQL = SQL & "INNER JOIN SB1010 SB1 WITH(NOLOCK) "
        SQL = SQL & "        ON SB5.B5_COD = SB1.B1_COD "
        SQL = SQL & "       AND ISNULL(SB5.D_E_L_E_T_, '') = '' "
        SQL = SQL & "WHERE (UPPER(SB1.B1_DESC) LIKE UPPER('%" & filtroProduto & "%') "
        SQL = SQL & "       OR SB1.B1_COD = '" & filtroProduto & "') "
        SQL = SQL & "  AND SB1.D_E_L_E_T_ <> '*' "
        SQL = SQL & "  AND SB5.B5_TIPO IN ('C', 'D', 'S', 'T') "
        SQL = SQL & "  AND SB1.B1_STATUS IN ('V', 'A', 'W', 'P') "
        SQL = SQL & "  AND SB1.B1_GRUPO IN ('71', '76') "
        SQL = SQL & "  AND SB5.B5_CLI = '" & codigoCliente & "' "
        SQL = SQL & "  AND SUBSTRING(ISNULL(SB1.B1_COD, ''), 1, 2) <> '77' "
        SQL = SQL & "  AND SB1.B1_SUBGRUP NOT IN ('750', '760') "

        SQL = SQL & "UNION "

        SQL = SQL & "SELECT DISTINCT "
        SQL = SQL & "       LTRIM(RTRIM(SB1.B1_COD)) AS Codigo, "
        SQL = SQL & "       LTRIM(RTRIM(("
        SQL = SQL & "           SELECT TOP 1 SB5A.B5_ESTRUTU "
        SQL = SQL & "           FROM SB5010 SB5A "
        SQL = SQL & "           WHERE SB5A.B5_COD = MIN(SG1.G1_COMP) "
        SQL = SQL & "             AND SB5A.D_E_L_E_T_ = ''"
        SQL = SQL & "       ))) AS Referencia "
        SQL = SQL & "FROM SB1010 SB1 "
        SQL = SQL & "INNER JOIN SG1010 SG1 "
        SQL = SQL & "        ON SG1.G1_COD = SB1.B1_COD "
        SQL = SQL & "       AND SG1.D_E_L_E_T_ = '' "
        SQL = SQL & "INNER JOIN SB5010 SB5 "
        SQL = SQL & "        ON SB5.B5_COD = SG1.G1_COMP "
        SQL = SQL & "       AND SB5.D_E_L_E_T_ = '' "
        SQL = SQL & "WHERE (UPPER(SB1.B1_DESC) LIKE UPPER('%" & filtroProduto & "%') "
        SQL = SQL & "       OR SB1.B1_COD = '" & filtroProduto & "') "
        SQL = SQL & "  AND SB1.D_E_L_E_T_ = '' "
        SQL = SQL & "  AND SB1.B1_SUBGRUP IN ('750', '760') "
        SQL = SQL & "  AND SB5.B5_CLI = '" & codigoCliente & "' "
        SQL = SQL & "  AND SUBSTRING(SB1.B1_COD, 1, 2) <> '77' "
        SQL = SQL & "  AND SB1.B1_STATUS IN ('V', 'A', 'W', 'P') "
        SQL = SQL & "GROUP BY SB1.B1_COD, SB1.B1_DESC, SB5.B5_CLI, SB1.B1_INPUBE"
    End If

    SentencaProdutosPirapetinga = SQL
End Function

Function SentencaProdutosFortaleza(codigoCliente, codigoLoja, tipoItem, filtroProduto)
    SQL = ""

    If tipoItem = "CJ" Then
        SQL = "SELECT TOP 1 " & _
              "       LTRIM(RTRIM(CODIGO_CONJUNTO)) AS Codigo, " & _
              "       LTRIM(RTRIM(REFERENCIA_CONJUNTO)) AS Referencia " & _
              "FROM VW_BI_PRODUTOS_PEDIDO " & _
              "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
              "  AND LOJA_CLIENTE = '" & codigoLoja & "' " & _
              "  AND (REFERENCIA_CONJUNTO LIKE '%" & filtroProduto & "%' " & _
              "       OR CODIGO_CONJUNTO = '" & filtroProduto & "') " & _
              "ORDER BY CODIGO_ITEM"
    Else
        SQL = "SELECT " & _
              "       LTRIM(RTRIM(CODIGO_ITEM)) AS Codigo, " & _
              "       LTRIM(RTRIM(REFERENCIA_ITEM)) AS Referencia " & _
              "FROM VW_BI_PRODUTOS_PEDIDO " & _
              "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
              "  AND (REFERENCIA_ITEM LIKE '%" & filtroProduto & "%' " & _
              "       OR CODIGO_ITEM = '" & filtroProduto & "')"
    End If

    SentencaProdutosFortaleza = SQL
End Function

Function SentencaProdutosSaquarema(codigoCliente, codigoLoja, tipoItem, filtroProduto)
    SQL = ""

    If tipoItem = "CJ" Then
        SQL = "SELECT TOP 1 " & _
              "       LTRIM(RTRIM(CODIGO_CONJUNTO)) AS Codigo, " & _
              "       LTRIM(RTRIM(REFERENCIA_CONJUNTO)) AS Referencia " & _
              "FROM VW_BI_PRODUTOS_PEDIDO " & _
              "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
              "  AND LOJA_CLIENTE = '" & codigoLoja & "' " & _
              "  AND (REFERENCIA_CONJUNTO LIKE '%" & filtroProduto & "%' " & _
              "       OR CODIGO_CONJUNTO = '" & filtroProduto & "') " & _
              "ORDER BY CODIGO_ITEM"
    Else
        SQL = "SELECT " & _
              "       LTRIM(RTRIM(CODIGO_ITEM)) AS Codigo, " & _
              "       LTRIM(RTRIM(REFERENCIA_ITEM)) AS Referencia " & _
              "FROM VW_BI_PRODUTOS_PEDIDO " & _
              "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' " & _
              "  AND LOJA_CLIENTE = '" & codigoLoja & "' " & _
              "  AND (REFERENCIA_ITEM LIKE '%" & filtroProduto & "%' " & _
              "       OR CODIGO_ITEM = '" & filtroProduto & "')"
    End If

    SentencaProdutosSaquarema = SQL
End Function

Function SentencaProdutosPlantasWRK(codigoPlanta, codigoCliente, filtroProduto)
    nomePlanta = fNomePlanta(codigoPlanta)

    SQL = "SELECT " & _
          "       LTRIM(RTRIM(CAST(PRODUTO AS VARCHAR(20)))) AS Codigo, " & _
          "       LTRIM(RTRIM(DESCRICAO_PRODUTO)) AS Referencia " & _
          "FROM VW_PRODUTOS_CLIENTES_WRK " & _
          "WHERE CIDADE = '" & nomePlanta & "' " & _
          "  AND CODIGO_CLIENTE = '" & codigoCliente & "' " & _
          "  AND (UPPER(DESCRICAO_PRODUTO) LIKE '%" & UCase(filtroProduto) & "%' " & _
          "       OR CAST(PRODUTO AS VARCHAR(20)) = '" & filtroProduto & "') " & _
          "GROUP BY PRODUTO, DESCRICAO_PRODUTO " & _
          "ORDER BY DESCRICAO_PRODUTO"

    SentencaProdutosPlantasWRK = SQL
End Function

Function AdicionarEmailLista(listaEmails, email)
    email = Trim(email & "")

    If email <> "" Then
        If InStr(1, ";" & LCase(listaEmails) & ";", ";" & LCase(email) & ";", vbTextCompare) = 0 Then
            If listaEmails <> "" Then listaEmails = listaEmails & ";"
            listaEmails = listaEmails & email
        End If
    End If

    AdicionarEmailLista = listaEmails
End Function

' Auxiliares
Function VerificarExistePedido(codigoPlanta, codigoProduto, idLotePiloto)
    flagExiste = False

    SQL = ""
    nomeConexao = ""

    Select Case codigoPlanta
        Case "00", "04"
            SQL = "SELECT Count(PRODUTO) AS Total FROM VW_BI_PEDIDOS WHERE PRODUTO = '" & codigoProduto & "' AND STATUS_PEDIDO <> 'CANCELADO'"
            nomeConexao = "PIR-DADOSADV"
        Case "05"
            SQL = "SELECT Count(PRODUTO) AS Total FROM VW_BI_PEDIDOS WHERE PRODUTO = '" & codigoProduto  & "' AND STATUS_PEDIDO <> 'CANCELADO'"
            nomeConexao = "BEN-DADOSADV"
        Case "06"
            SQL = "SELECT Count(PRODUTO) AS Total FROM VW_BI_PEDIDOS WHERE PRODUTO = '" & codigoProduto  & "' AND STATUS_PEDIDO <> 'CANCELADO'"  
            nomeConexao = "FOR-DADOSADV"
        Case "07"
            SQL = "SELECT Count(PRODUTO) AS Total FROM VW_BI_PEDIDOS WHERE PRODUTO = '" & codigoProduto  & "' AND STATUS_PEDIDO <> 'CANCELADO'"
            nomeConexao = "SAQ-DADOSADV"
        Case Else
            SQL = "SELECT CASE WHEN PEDIDO_WRK IS NOT NULL AND PEDIDO_WRK <> '' THEN 1 ELSE 0 END AS Total " & _
                " FROM LOTE_PILOTO " & _
                " WHERE ID = " & idLotePiloto & _
                " AND DATA_EXCLUSAO IS NULL "
            nomeConexao = "PIR-CORPORATIVO"
    End Select

    On Error Resume Next
    SET rsExiste = Conecta(nomeConexao, SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        VerificarExistePedido = flagExiste
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.EOF Then
        If rsExiste("Total") > 0 Then
            flagExiste = True
        End If
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    
    VerificarExistePedido = flagExiste                  
End Function

Function VerificarLotePilotoJaExiste(codigoPlanta, codigoProduto)
    flagExiste = False
    
    SQL = "SELECT PRODUTO FROM LOTE_PILOTO "
    SQL = SQL & "WHERE PLANTA = '" & codigoPlanta & "' "
    SQL = SQL & "AND PRODUTO = '" & codigoProduto & "' " 
    SQL = SQL & "AND QUEM_EXCLUIU IS NULL"

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsExiste) Then
        VerificarLotePilotoJaExiste = flagExiste
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.EOF Then
        flagExiste = True
    End If
    
    rsExiste.Close()
    SET rsExiste = NOTHING
    
    VerificarLotePilotoJaExiste = flagExiste
End Function

Function VerificarEhPlantaWestrock(codigoPlanta)
    ehPlantaWestRock = False

    If codigoPlanta = "08" OR codigoPlanta = "09" OR codigoPlanta = "10" OR codigoPlanta = "11" OR codigoPlanta = "12" OR codigoPlanta = "13" Then
        ehPlantaWestRock = True
    End If

    VerificarEhPlantaWestrock = ehPlantaWestRock
End Function

Function ObterConexaoProtheus(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "04": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "05": ObterConexaoProtheus = "BEN-DADOSADV"
        Case "06": ObterConexaoProtheus = "FOR-DADOSADV"
        Case "07": ObterConexaoProtheus = "SAQ-DADOSADV"
        Case Else: ObterConexaoProtheus = ""
    End Select
End Function

Function ObterNomePlanta(codPlanta)
    nomePlanta = ""
    Select Case codPlanta
        Case "00"
            nomePlanta = "PIRAPETINGA"
        Case "04"
            nomePlanta = "UBERABA"
        Case "05"
            nomePlanta = "BENTO"
        Case "06"
            nomePlanta = "MARANGUAPE"
        Case "07"
            nomePlanta = "SAQUAREMA"
        Case "08"
            nomePlanta = "ARACATUBA"
        Case "09"
            nomePlanta = "BLUMENAU"
        Case "10"
            nomePlanta = "PACAJUS"
        Case "11"
            nomePlanta = "PORTO FELIZ"
        Case "12"
            nomePlanta = "TRES BARRAS"
        Case "13"
            nomePlanta = "FLORESTAL"
        Case Else
            nomePlanta = ""
    End Select
    
    ObterNomePlanta = nomePlanta
End Function

Function ObterDescricaoStatus(codPlanta)
    descricao = ""
    Select Case codPlanta
        Case "1"
            descricao = "Aguardando Pedido LP"
        Case "2"
            descricao = "Convertido em Venda"
        Case "3"
            descricao = "Aprovado sem Pedido de Conversão"
        Case "4"
            descricao = "Em Teste Cliente"
        Case "5"
            descricao = "Reprovado pelo Cliente"
        Case "6"
            descricao = "Reprovado Internamente/Retrabalho"
        Case "7"
            descricao = "Aguardando Produção"
        Case "8"
            descricao = "Descontinuado"
        Case "9"
            descricao = "Sem Retorno Vendas/Cliente"
        Case "10"
            descricao = "Processo Segue em Novo Produto/Ficha Técnica"
        Case "11"
            descricao = "Validação Máquina Nova/Homologação"
        Case Else
            descricao = ""
    End Select
    
    ObterDescricaoStatus = descricao
End Function


Function ObterDadosEmailCadastro(codigo)
    Set dadosEmail = Server.CreateObject("Scripting.Dictionary")
    dadosEmail("codigo") = codigo
    dadosEmail("razaoSocial") = ""
    dadosEmail("nomeFantasia") = ""
    dadosEmail("documento") = ""
    dadosEmail("codigoPlanta") = ""
    dadosEmail("nomeSolicitante") = ""
    dadosEmail("emailSolicitante") = ""

    SQL = "SELECT TOP 1 "
    SQL = SQL & "   a.Codigo, "
    SQL = SQL & "   a.RazaoSocial, "
    SQL = SQL & "   a.NomeFantasia, "
    SQL = SQL & "   ISNULL(NULLIF(a.Cnpj, ''), a.Cpf) AS Documento, "
    SQL = SQL & "   a.Filial AS CodigoPlanta, "
    SQL = SQL & "   ISNULL(t.CRM_NOME, a.QuemIncluiu) AS NomeSolicitante, "
    SQL = SQL & "   ISNULL(t.CRM_EMAILPRIMARIO, '') AS EmailSolicitante "
    SQL = SQL & "FROM LOTE_PILOTO a "
    SQL = SQL & "LEFT JOIN HELPDESK..CRM_DTI t ON t.CRM_LOGIN = a.QuemIncluiu "
    SQL = SQL & "WHERE a.Codigo = '" & Replace(codigo, "'", "''") & "' "

    On Error Resume Next
    SET rsDadosEmail = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(rsDadosEmail) Then
        If Not rsDadosEmail.EOF Then
            dadosEmail("codigo") = rsDadosEmail("Codigo") & ""
            dadosEmail("razaoSocial") = rsDadosEmail("RazaoSocial") & ""
            dadosEmail("nomeFantasia") = rsDadosEmail("NomeFantasia") & ""
            dadosEmail("documento") = rsDadosEmail("Documento") & ""
            dadosEmail("codigoPlanta") = rsDadosEmail("CodigoPlanta") & ""
            dadosEmail("nomeSolicitante") = rsDadosEmail("NomeSolicitante") & ""
            dadosEmail("emailSolicitante") = rsDadosEmail("EmailSolicitante") & ""
            'dadosEmail("emailSolicitante") = "thiago.pereira@smurfitwestrock.com.br"
        End If

        rsDadosEmail.Close()
        SET rsDadosEmail = NOTHING
    End If

    On Error Goto 0

    Set ObterDadosEmailCadastro = dadosEmail
End Function

Function montarHTMLEmail(dadosEmail, titulo, texto, observacao)
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
    mensagem = mensagem & "<p>" & EscapeHTML(texto) & "</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Código:</strong> " & EscapeHTML(dadosEmail("codigo")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Razão Social:</strong> " & EscapeHTML(dadosEmail("razaoSocial")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Nome Fantasia:</strong> " & EscapeHTML(dadosEmail("nomeFantasia")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>CNPJ/CPF:</strong> " & EscapeHTML(dadosEmail("documento")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Planta:</strong> " & EscapeHTML(dadosEmail("codigoPlanta")) & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Solicitante:</strong> " & EscapeHTML(dadosEmail("nomeSolicitante")) & "</div>"

    If Trim(observacao & "") <> "" Then
        mensagem = mensagem & "<br>"
        mensagem = mensagem & "<div class=""info-item""><strong>Observação:</strong> " & EscapeHTML(observacao) & "</div>"
    End If

    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    montarHTMLEmail = mensagem
End Function

Sub EnviarEmailCadastroFinalizado(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi finalizado."
    titulo = "Cadastro finalizado"    

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailDevolucao(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi devolvido ao solicitante para ajustes."
    titulo = "Cadastro devolvido ao solicitante"

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Sub EnviarEmailCancelamento(Dados)
    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    destinatarios = dadosEmail("emailSolicitante")
    If destinatarios = "" Then Exit Sub

    texto = "O cadastro de cliente foi cancelado."
    titulo = "Cadastro cancelado"

    Set dadosEmail = ObterDadosEmailCadastro(Dados("codigo"))
    assunto = "[Cadastro Unificado] " & titulo & " - " & dadosEmail("codigo")
    mensagem = montarHTMLEmail(dadosEmail, titulo, texto, Dados("observacao"))

    On Error Resume Next
    Call EnviaEmail(destinatarios, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
    On Error Goto 0
End Sub

Function SanitizeSQL(Value)
    If IsNull(Value) Then
        SanitizeSQL = ""
    Else
        SanitizeSQL = Replace(CStr(Value), "'", "''")
    End If
End Function

Sub RegistrarLog(acao, id)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de Lote Piloto"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou o lote piloto de id " & id
        Case "edicao"
            acaoLog = "Edição de Lote Piloto"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou o lote piloto de id " & id
        Case "cancelamento"
            acaoLog = "Cancelamento de Lote Piloto"
            descricaoLog = "Usuário " & usuarioLogado("login") & " cancelou o lote piloto de id " & id 
        Case "exclusao"
            acaoLog = "Exclusão de Lote Piloto"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu o lote piloto de id " & id 
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Lote Piloto', "
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