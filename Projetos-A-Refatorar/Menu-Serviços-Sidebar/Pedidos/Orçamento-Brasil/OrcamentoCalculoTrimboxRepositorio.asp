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

Dim codPlanta, itemsIds, tipoCalculo, valorBase

tipoCalculo = Request("tipoCalculo")
valorBase = Request("valorBase")
codPlanta = Request("codPlanta")
itemsIds = Request("itemsIds")

If codPlanta = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Nenhuma planta informada""}")
    Response.End()
End If

If tipoCalculo = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Nenhum tipo de cálculo informado""}")
    Response.End()
End If

If itemsIds = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Nenhum item informado""}")
    Response.End()
End If

'--------------------------------------
'      VARIÁVEIS GLOBAIS
'--------------------------------------
Dim itensCalculo
Dim itensSucesso(), itensErro()
Dim countSucesso, countErro
Dim sucessoIds(), erroIds()

'--------------------------------------
'      PROCESSAMENTO PRINCIPAL
'--------------------------------------
itensCalculo = BuscarItensOrcamento(itemsIds)

Call GravarProdutosTrimbox(codPlanta)
Call AtualizarCustoTransporte(codPlanta, itensCalculo(0)("IdMunicipio"), itensCalculo(0)("TipoCaminhao"))
Call RealizarCalculos(itensCalculo, codPlanta, tipoCalculo, valorBase)
Call AtualizarClienteSeNecessario(itensCalculo(0), codPlanta)
Call InserirOrcamentos(itensCalculo, codPlanta, tipoCalculo)

Dim respostaJson, i
respostaJson = "{""sucesso"": true, ""mensagem"": ""Items calculados com sucesso!"", ""sucessoIds"": ["

If UBound(sucessoIds) >= 0 Then
    For i = 0 To UBound(sucessoIds)
        If i > 0 Then respostaJson = respostaJson & ","
        respostaJson = respostaJson & sucessoIds(i)
    Next
End If

respostaJson = respostaJson & "], ""erroIds"": ["

If UBound(erroIds) >= 0 Then
    For i = 0 To UBound(erroIds)
        If i > 0 Then respostaJson = respostaJson & ","
        respostaJson = respostaJson & erroIds(i)
    Next
End If

respostaJson = respostaJson & "]}"

Response.Write(respostaJson)


'--------------------------------------
'       FUNÇÕES PRINCIPAIS
'--------------------------------------

Function BuscarItensOrcamento(itemsIds)
    Dim SQL, PIR_CORPORATIVO, itensCalculo(), idx
    
    SQL = "SELECT PRODUTO as Produto, ID_ORCAMENTO as IdOrcamentoSknet, REFERENCIA as Referencia, COD_CLIENTE CodCliente, "
    SQL = SQL & "NOME_CLIENTE as NomeCliente, CASE WHEN FLAG_PRODUTO = 1 THEN 1 WHEN TIPO_PRODUTO = 'CH' AND PLANTA <> '06' THEN 3 ELSE 2 END as FlagProduto, "
    SQL = SQL & "CASE WHEN FLAG_CLIENTE = 1 THEN 1 ELSE 0 END as FlagCliente, QUANTIDADE as Quantidade, PLANTA AS Planta, "
    SQL = SQL & "CONDICAO_PAGAMENTO as CondicaoPagamento, TIPO_CAMINHAO as TipoCaminhao, "
    SQL = SQL & "PRECO_KG as PrecoKg, VALOR_UNITARIO * 1000 as ValorUnitario, "
    SQL = SQL & "NR_ORCAMENTO as NrOrcamento, IDMUNICIPIO as IdMunicipio, "
    SQL = SQL & "CASE WHEN TIPO_FRETE = 'FOB' THEN 0 ELSE 1 END as TipoFrete "
    SQL = SQL & "FROM ORCAMENTO_BRASIL WHERE ID_ORCAMENTO IN (" & itemsIds & ")"

    ReDim itensCalculo(0)
    idx = 0

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        While NOT PIR_CORPORATIVO.EOF
            ReDim Preserve itensCalculo(idx)
            Set itensCalculo(idx) = Server.CreateObject("Scripting.Dictionary")
            
            itensCalculo(idx)("IdOrcamentoSknet") = PIR_CORPORATIVO("IdOrcamentoSknet") & ""
            itensCalculo(idx)("Produto") = PIR_CORPORATIVO("Produto") & ""
            itensCalculo(idx)("Planta") = PIR_CORPORATIVO("Planta") & ""
            itensCalculo(idx)("Referencia") = PIR_CORPORATIVO("Referencia") & ""
            itensCalculo(idx)("CodCliente") = PIR_CORPORATIVO("CodCliente") & ""
            itensCalculo(idx)("NomeCliente") = PIR_CORPORATIVO("NomeCliente") & ""
            itensCalculo(idx)("FlagProduto") = PIR_CORPORATIVO("FlagProduto") & ""
            itensCalculo(idx)("FlagCliente") = PIR_CORPORATIVO("FlagCliente") & ""
            itensCalculo(idx)("Quantidade") = PIR_CORPORATIVO("Quantidade") & ""
            itensCalculo(idx)("CondicaoPagamento") = PIR_CORPORATIVO("CondicaoPagamento") & ""
            itensCalculo(idx)("TipoCaminhao") = PIR_CORPORATIVO("TipoCaminhao") & ""
            itensCalculo(idx)("PrecoKg") = PIR_CORPORATIVO("PrecoKg") & ""
            itensCalculo(idx)("ValorUnitario") = PIR_CORPORATIVO("ValorUnitario") & ""
            itensCalculo(idx)("NrOrcamento") = PIR_CORPORATIVO("NrOrcamento") & ""
            itensCalculo(idx)("IdMunicipio") = PIR_CORPORATIVO("IdMunicipio") & ""
            itensCalculo(idx)("TipoFrete") = PIR_CORPORATIVO("TipoFrete") & ""
                        
            idx = idx + 1
            PIR_CORPORATIVO.MoveNext()
        Wend
        
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
    
    BuscarItensOrcamento = itensCalculo
End Function


Sub GravarProdutosTrimbox(codPlanta)
    Dim SQL, CONEXAO_TRIMBOX, CONEXAO_PRODUTO
    Dim produtosExistentes, listaProdutos
    Dim countExistentes, countNaoExistentes, i
    Dim itensExistentes(), itensNaoExistentes()
    
    listaProdutos = ""
    For i = 0 To UBound(itensCalculo)
        If i > 0 Then listaProdutos = listaProdutos & ","
        listaProdutos = listaProdutos & "'" & itensCalculo(i)("Produto") & "'"
    Next
    
    SQL = "SELECT DISTINCT Conjunto as Produto FROM Itens "
    SQL = SQL & "WHERE Item IN (" & listaProdutos & ") OR Conjunto IN (" & listaProdutos & ")"
    
    Set produtosExistentes = Server.CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(CONEXAO_TRIMBOX) Then
        While NOT CONEXAO_TRIMBOX.EOF
            produtosExistentes(CONEXAO_TRIMBOX("Produto") & "") = True
            CONEXAO_TRIMBOX.MoveNext()
        Wend
        CONEXAO_TRIMBOX.Close()
        SET CONEXAO_TRIMBOX = NOTHING
    End If
    On Error Goto 0
    
    ' Separa itens existentes e não existentes
    ReDim itensExistentes(0)
    ReDim itensNaoExistentes(0)
    countExistentes = 0
    countNaoExistentes = 0
    
    For i = 0 To UBound(itensCalculo)
        If produtosExistentes.Exists(itensCalculo(i)("Produto")) Then
            If countExistentes > 0 Then ReDim Preserve itensExistentes(countExistentes)
            Set itensExistentes(countExistentes) = Server.CreateObject("Scripting.Dictionary")
            itensExistentes(countExistentes)("IdOrcamentoSknet") = itensCalculo(i)("IdOrcamentoSknet") & ""
            itensExistentes(countExistentes)("Produto") = itensCalculo(i)("Produto") & ""
            countExistentes = countExistentes + 1
        Else
            If countNaoExistentes > 0 Then ReDim Preserve itensNaoExistentes(countNaoExistentes)
            Set itensNaoExistentes(countNaoExistentes) = Server.CreateObject("Scripting.Dictionary")
            itensNaoExistentes(countNaoExistentes)("IdOrcamentoSknet") = itensCalculo(i)("IdOrcamentoSknet") & ""
            itensNaoExistentes(countNaoExistentes)("Produto") = itensCalculo(i)("Produto") & ""

            ' Se não existe devemos setar flagProduto = 1 (Pre-ft)
            itensCalculo(i)("FlagProduto") = 1
            countNaoExistentes = countNaoExistentes + 1
        End If
    Next

    ' Grava itens existentes (FT)
    If countExistentes > 0 Then
        For i = 0 To UBound(itensExistentes)
            SQL = "EXEC SP_GRAVA_TSK_FT " & itensExistentes(i)("IdOrcamentoSknet")
            
            On Error Resume Next
            SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
            If IsObject(CONEXAO_TRIMBOX) Then 
                CONEXAO_TRIMBOX.Close()
                SET CONEXAO_TRIMBOX = NOTHING
            End If
            On Error Goto 0
        Next
    End If
    
    ' Grava itens não existentes (PRÉ-FT)
    If countNaoExistentes > 0 Then
        For i = 0 To UBound(itensNaoExistentes)
            SQL = "EXEC SP_GRAVA_TSK_PREFT " & itensNaoExistentes(i)("IdOrcamentoSknet")
            
            On Error Resume Next
            SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
            If IsObject(CONEXAO_TRIMBOX) Then 
                CONEXAO_TRIMBOX.Close()
                SET CONEXAO_TRIMBOX = NOTHING
            End If
            On Error Goto 0
        Next
    End If
End Sub

Sub AtualizarCustoTransporte(codPlanta, idMunicipio, tipoCaminhao)
    Dim SQL, CONEXAO_TRANSPORTE
    
    SQL = "IF EXISTS("
    SQL = SQL & "   SELECT NULL FROM CstCstTrnspCidade WHERE Versao = dbo.fq_VersaoCustoAtivaNoMomento(GETDATE())"
    SQL = SQL & "   AND IDMunicipio = " & idMunicipio & " AND LTRIM(RTRIM(TipoCaminhao)) = LTRIM(RTRIM('" & tipoCaminhao & "'))"
    SQL = SQL & "   AND CustoViagem = 99999"
    SQL = SQL & ") "
    SQL = SQL & "BEGIN "
    SQL = SQL & "   UPDATE CstCstTrnspCidade "
    SQL = SQL & "   SET CustoViagem = 88888, CustoViagemEdicao = 88888, CustoPedagio = 0, "
    SQL = SQL & "       CustoPedagioEdicao = 0, CustoPorcent = 0, CustoPorcentEdicao = 0 "
    SQL = SQL & "   WHERE Versao = dbo.fq_VersaoCustoAtivaNoMomento(GETDATE())"
    SQL = SQL & "       AND IDMunicipio = " & idMunicipio & " AND LTRIM(RTRIM(TipoCaminhao)) = LTRIM(RTRIM('" & tipoCaminhao & "'))"
    SQL = SQL & "       AND CustoViagem = 99999 "
    SQL = SQL & "END"

    On Error Resume Next
    SET CONEXAO_TRANSPORTE = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
    If IsObject(CONEXAO_TRANSPORTE) Then 
        CONEXAO_TRANSPORTE.Close()
        SET CONEXAO_TRANSPORTE = NOTHING
    End If
    On Error Goto 0
End Sub

Sub RealizarCalculos(itensCalculo, codPlanta, tipoCalculo, valorBase)
    Dim SQL, CONEXAO_TRIMBOX, i
    Dim argumento, descricaoCalculo, tipoCalculoNum
    
    For i = 0 To UBound(itensCalculo)
        ' Define parâmetros do cálculo
        Select Case tipoCalculo
            Case "VU"
                argumento = Iif(valorBase <> 0, ConverterNumeroParaCalcular(valorBase) * 1000, itensCalculo(i)("ValorUnitario"))
                descricaoCalculo = "VALOR UNITARIO"
                tipoCalculoNum = 2
            Case "PK"
                argumento = Iif(valorBase <> 0, ConverterNumeroParaCalcular(valorBase), itensCalculo(i)("PrecoKg"))
                descricaoCalculo = "PRECO KG"
                tipoCalculoNum = 3
            Case "MC"
                argumento = ConverterNumeroParaCalcular(valorBase)
                descricaoCalculo = "MARGEM"
                tipoCalculoNum = 3
        End Select

        'Atualiza com o código gerado para o produto
        SQL = "SELECT PRODUTO as CodProduto FROM ORCAMENTO_BRASIL WHERE ID_ORCAMENTO = " & itensCalculo(i)("IdOrcamentoSknet")
        
        On Error Resume Next
        SET CONEXAO_PRODUTO = Conecta("PIR-CORPORATIVO", SQL)
        
        If NOT CONEXAO_PRODUTO.EOF Then
            itensCalculo(i)("Produto") = CONEXAO_PRODUTO("CodProduto")
        End If
        
        If IsObject(CONEXAO_PRODUTO) Then
            CONEXAO_PRODUTO.Close()
            SET CONEXAO_PRODUTO = NOTHING
        End If
        On Error Goto 0

        SQL = "Declare @orcamento INT, @RETURN_VALUES INT, @erro VARCHAR(8000) "
        SQL = SQL & "EXEC @RETURN_VALUES = ssp_ConsultaOrcamentoProdutosPeloInpanet_SK "
        SQL = SQL & "@conjunto = '" & itensCalculo(i)("Produto") & "', "
        SQL = SQL & "@tipoProduto = '" & itensCalculo(i)("FlagProduto") & "', "
        SQL = SQL & "@quantidade = " & itensCalculo(i)("Quantidade") & ", "
        SQL = SQL & "@condpag = '" & itensCalculo(i)("CondicaoPagamento") & "', "
        SQL = SQL & "@tipocaminhao = '" & itensCalculo(i)("TipoCaminhao") & "', "
        SQL = SQL & "@tipocalcfrete = " & itensCalculo(i)("TipoFrete") & ", "
        SQL = SQL & "@tipocalculo = '" & tipoCalculoNum & "', "
        SQL = SQL & "@argumento = '" & Replace(CStr(argumento), ",", ".") & "', "
        SQL = SQL & "@ORCAMENTO = @orcamento OUTPUT, @ERRO = @erro OUTPUT "
        SQL = SQL & "SELECT @orcamento AS ORCAMENTO, @erro as ERRO"

        On Error Resume Next
        SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
        
        If Err.Number = 0 AND IsObject(CONEXAO_TRIMBOX) Then
            If NOT CONEXAO_TRIMBOX.EOF Then
                itensCalculo(i)("Calculado") = Iif(CONEXAO_TRIMBOX("ERRO") <> "", 0, 1)
                itensCalculo(i)("IdOrcamentoTrim") = CONEXAO_TRIMBOX("ORCAMENTO")
                itensCalculo(i)("Erro") = CONEXAO_TRIMBOX("ERRO")
                itensCalculo(i)("TipoCalculoNum") = tipoCalculoNum
                itensCalculo(i)("DescricaoCalculo") = descricaoCalculo
                itensCalculo(i)("Argumento") = Iif(tipoCalculoNum = 2, argumento / 1000, argumento)
            End If
            CONEXAO_TRIMBOX.Close()
            SET CONEXAO_TRIMBOX = NOTHING
        End If
        On Error Goto 0
    Next
End Sub

Sub AtualizarClienteSeNecessario(primeiroItem, codPlanta)
    Dim SQL, CONEXAO_TRIMBOX
    Dim codCliente, criouCliente
    
    If primeiroItem("FlagCliente") <> "1" OR primeiroItem("CodCliente") <> "1" Then Exit Sub
    
    SQL = "SELECT TOP 1 CodCliente FROM TSK_IN_PreFTs "
    SQL = SQL & "WHERE LTRIM(RTRIM(Obs1)) = " & primeiroItem("IdOrcamentoSknet") & " "
    SQL = SQL & "ORDER BY id_SK_InPreFTs DESC"

    On Error Resume Next
    SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(CONEXAO_TRIMBOX) Then
        If NOT CONEXAO_TRIMBOX.EOF Then
            codCliente = CONEXAO_TRIMBOX("CodCliente")
            criouCliente = IIf(codCliente <> "1", 1, 0)
        End If
        CONEXAO_TRIMBOX.Close()
        SET CONEXAO_TRIMBOX = NOTHING
    End If
    On Error Goto 0

    If criouCliente = 1 Then
        SQL = "UPDATE ORCAMENTO_BRASIL SET COD_CLIENTE ='" & codCliente & "', FLAG_CLIENTE = 0 "
        SQL = SQL & "WHERE NR_ORCAMENTO =" & primeiroItem("NrOrcamento")

        On Error Resume Next
        SET CONEXAO_TRIMBOX = Conecta("PIR-CORPORATIVO", SQL)
        If IsObject(CONEXAO_TRIMBOX) Then
            CONEXAO_TRIMBOX.Close()
            SET CONEXAO_TRIMBOX = NOTHING
        End If
        On Error Goto 0
    End If
End Sub

Sub InserirOrcamentos(itensCalculo, codPlanta, tipoCalculo)
    Dim SQL, CONEXAO_TRIMBOX, PIR_CORPORATIVO, i
    Dim orcamentoTrimIds, orcamentoSkNetIds
        
    ' Inicializa arrays de IDs vazios
    ReDim sucessoIds(-1)
    ReDim erroIds(-1)
    
    countSucesso = 0
    countErro = 0

    ' Separa itens com sucesso e com erro
    For i = 0 To UBound(itensCalculo)
        If itensCalculo(i)("Calculado") = 1 Then
            ReDim Preserve itensSucesso(countSucesso)
            Set itensSucesso(countSucesso) = itensCalculo(i)
            countSucesso = countSucesso + 1
        Else
            ReDim Preserve itensErro(countErro)
            Set itensErro(countErro) = itensCalculo(i)
            countErro = countErro + 1
        End If
    Next

    ' Processa itens com sucesso
    If countSucesso > 0 Then
        orcamentoTrimIds = ""
        orcamentoSkNetIds = ""

        For i = 0 To UBound(itensSucesso)
            If i > 0 Then
                orcamentoTrimIds = orcamentoTrimIds & ","
                orcamentoSkNetIds = orcamentoSkNetIds & ","
            End If
            orcamentoTrimIds = orcamentoTrimIds & itensSucesso(i)("IdOrcamentoTrim")
            orcamentoSkNetIds = orcamentoSkNetIds & itensSucesso(i)("IdOrcamentoSknet")
        Next

        ' Validamos registros que foram criados no trimbox, apenas itens criados no trimbox serão criados na tabela de sucesso
        SQL = "SELECT Orcamento, "
        SQL = SQL & "MAX(CASE WHEN Descricao = 'ICMS' THEN ISNULL(PorcValorVenda, 0) ELSE 0 END) as ICMS "
        SQL = SQL & "FROM vw_ConsultaCustosItensOrcamentosPelaInpaNet "
        SQL = SQL & "WHERE Orcamento IN (" & orcamentoTrimIds & ") "
        SQL = SQL & "GROUP BY Orcamento"
        
        Dim orcamentosICMS
        Set orcamentosICMS = CreateObject("Scripting.Dictionary")
        
        On Error Resume Next
        SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
        
        If Err.Number = 0 AND IsObject(CONEXAO_TRIMBOX) Then
            Dim orcTemp, icmsTemp
            While NOT CONEXAO_TRIMBOX.EOF
                orcTemp = CStr(CONEXAO_TRIMBOX("Orcamento"))
                icmsTemp = CONEXAO_TRIMBOX("ICMS")
                orcamentosICMS.Add orcTemp, icmsTemp
                CONEXAO_TRIMBOX.MoveNext()
            Wend
            CONEXAO_TRIMBOX.Close()
            SET CONEXAO_TRIMBOX = NOTHING
        End If
        On Error Goto 0
        
        Dim itensSucessoFiltrados()
        Dim countSucessoFiltrado, idOrcTrim
        countSucessoFiltrado = 0
        
        For i = 0 To UBound(itensSucesso)
            idOrcTrim = CStr(itensSucesso(i)("IdOrcamentoTrim"))
            
            If orcamentosICMS.Exists(idOrcTrim) Then
                itensSucesso(i)("ICMS") = orcamentosICMS(idOrcTrim)
                
                ReDim Preserve itensSucessoFiltrados(countSucessoFiltrado)
                Set itensSucessoFiltrados(countSucessoFiltrado) = itensSucesso(i)
                countSucessoFiltrado = countSucessoFiltrado + 1
            End If
        Next
        
        Set orcamentosICMS = Nothing
        
        If countSucessoFiltrado > 0 Then
            orcamentoSkNetIds = ""
            For i = 0 To UBound(itensSucessoFiltrados)
                If i > 0 Then orcamentoSkNetIds = orcamentoSkNetIds & ","
                orcamentoSkNetIds = orcamentoSkNetIds & itensSucessoFiltrados(i)("IdOrcamentoSknet")
            Next

            ' Atualiza registros antigos
            SQL = "UPDATE ORCAMENTO_VALIDA_ERRO SET ESTADO = 0 "
            SQL = SQL & "WHERE ORCAMENTO_SKNET IN (" & orcamentoSkNetIds & ") AND ESTADO = 1"
            
            On Error Resume Next
            SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
            If IsObject(PIR_CORPORATIVO) Then
                PIR_CORPORATIVO.Close()
                SET PIR_CORPORATIVO = NOTHING
            End If
            On Error Goto 0

            SQL = "UPDATE ORCAMENTO_SKNET_TRIM SET VALIDADOR = 0 "
            SQL = SQL & "WHERE ORCAMENTO_SKNET IN (" & orcamentoSkNetIds & ") AND VALIDADOR = 1"
            
            On Error Resume Next
            SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
            If IsObject(PIR_CORPORATIVO) Then
                PIR_CORPORATIVO.Close()
                SET PIR_CORPORATIVO = NOTHING
            End If
            On Error Goto 0
                
            ' Insere novos orçamentos e captura IDs
            Dim clausulasInsert
            clausulasInsert = ""
            
            For i = 0 To UBound(itensSucessoFiltrados)
                If i > 0 Then clausulasInsert = clausulasInsert & ","
                
                clausulasInsert = clausulasInsert & "("
                clausulasInsert = clausulasInsert & "'" & ObterNomePlanta(codPlanta) & "', "
                clausulasInsert = clausulasInsert & "'" & itensSucessoFiltrados(i)("IdOrcamentoSknet") & "', "
                clausulasInsert = clausulasInsert & "'" & itensSucessoFiltrados(i)("IdOrcamentoTrim") & "', "
                clausulasInsert = clausulasInsert & "'" & itensSucessoFiltrados(i)("NomeCliente") & "', "
                clausulasInsert = clausulasInsert & "'" & itensSucessoFiltrados(i)("Produto") & " - " & itensSucessoFiltrados(i)("Referencia") & "', "
                clausulasInsert = clausulasInsert & "'" & itensSucessoFiltrados(i)("DescricaoCalculo") & "', "
                clausulasInsert = clausulasInsert & Replace(CStr(itensSucessoFiltrados(i)("Argumento")), ",", ".") & ", "
                clausulasInsert = clausulasInsert & Replace(CStr(itensSucessoFiltrados(i)("ICMS")), ",", ".") & ", "
                clausulasInsert = clausulasInsert & "GETDATE(), 1)"
            Next
            
            If clausulasInsert <> "" Then
                SQL = "INSERT INTO ORCAMENTO_SKNET_TRIM "
                SQL = SQL & "(PLANTA, ORCAMENTO_SKNET, ORCAMENTO_TRIM, CLIENTE, PRODUTO, "
                SQL = SQL & "TIPO_CALCULO, VALOR_ANALISADO, ALIQ_ICMS, DATA_INCLUSAO, VALIDADOR) "
                SQL = SQL & "OUTPUT INSERTED.ID "
                SQL = SQL & "VALUES " & clausulasInsert

                On Error Resume Next
                SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
                
                If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
                    Dim contadorSucesso
                    contadorSucesso = 0
                    
                    If Not PIR_CORPORATIVO.EOF Then
                        Do While Not PIR_CORPORATIVO.EOF
                            ReDim Preserve sucessoIds(contadorSucesso)
                            sucessoIds(contadorSucesso) = PIR_CORPORATIVO("ID")
                            contadorSucesso = contadorSucesso + 1
                            PIR_CORPORATIVO.MoveNext
                        Loop
                    End If
                    
                    PIR_CORPORATIVO.Close()
                    SET PIR_CORPORATIVO = NOTHING
                End If
                On Error Goto 0
            End If
        End If
    End If

    ' Processa itens com erro
    If countErro > 0 Then
        orcamentoSkNetIds = ""
        For i = 0 To UBound(itensErro)
            If i > 0 Then orcamentoSkNetIds = orcamentoSkNetIds & ","
            orcamentoSkNetIds = orcamentoSkNetIds & itensErro(i)("IdOrcamentoSknet")
        Next

        SQL = "UPDATE ORCAMENTO_VALIDA_ERRO SET ESTADO = 0 "
        SQL = SQL & "WHERE ORCAMENTO_SKNET IN (" & orcamentoSkNetIds & ") AND ESTADO = 1"
        
        On Error Resume Next
        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        If IsObject(PIR_CORPORATIVO) Then
            PIR_CORPORATIVO.Close()
            SET PIR_CORPORATIVO = NOTHING
        End If
        On Error Goto 0

        Dim clausulasInsertErro, tipoInsumo, chaveInsumo
        clausulasInsertErro = ""

        For i = 0 To UBound(itensErro)
            tipoInsumo = ""
            chaveInsumo = ""

            Select Case itensErro(i)("Erro")
                Case " Orçamento com custos não preenchidos"
                    SQL = "SELECT TipoDeCusto AS TipoInsumo, ChaveSemCusto AS ChaveInsumo "
                    SQL = SQL & "FROM fq_VerificaCustosPreenchidosOrcamento_PeloInpaNET('" & itensErro(i)("IdOrcamentoTrim") & "')"
                Case " Orçamento com insumos inativos" 
                    SQL = "SELECT TipoInsumo, ChaveInsumo "
                    SQL = SQL & "FROM fq_VerificaInsumosInativosOrcamento_PeloInpaNET('" & itensErro(i)("IdOrcamentoTrim") & "')"
                Case Else
                    SQL = ""
            End Select

            If SQL <> "" Then            
                On Error Resume Next
                SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
                If NOT CONEXAO_TRIMBOX.EOF Then
                    tipoInsumo = CONEXAO_TRIMBOX("TipoInsumo")
                    chaveInsumo = CONEXAO_TRIMBOX("ChaveInsumo")
                End If
                If IsObject(CONEXAO_TRIMBOX) Then
                    CONEXAO_TRIMBOX.Close()
                    SET CONEXAO_TRIMBOX = NOTHING
                End If
                On Error Goto 0
            End If

            If i > 0 Then clausulasInsertErro = clausulasInsertErro & ","
            
            clausulasInsertErro = clausulasInsertErro & "("
            clausulasInsertErro = clausulasInsertErro & "'" & ObterNomePlanta(codPlanta) & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & itensErro(i)("IdOrcamentoSknet") & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & itensErro(i)("IdOrcamentoTrim") & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & itensErro(i)("NomeCliente") & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & itensErro(i)("Produto") & " - " & itensErro(i)("Referencia") & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & itensErro(i)("Erro") & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & tipoInsumo & "', "
            clausulasInsertErro = clausulasInsertErro & "'" & chaveInsumo & "', "
            clausulasInsertErro = clausulasInsertErro & "1, GETDATE())"
        Next

        If clausulasInsertErro <> "" Then
            SQL = "INSERT INTO ORCAMENTO_VALIDA_ERRO "
            SQL = SQL & "(PLANTA, ORCAMENTO_SKNET, ORCAMENTO_TRIM, CLIENTE, PRODUTO, "
            SQL = SQL & "ERRO, TIPO_CUSTO, CHAVE_CUSTO, ESTADO, DATA_INCLUSAO) "
            SQL = SQL & "OUTPUT INSERTED.ID "
            SQL = SQL & "VALUES " & clausulasInsertErro

            On Error Resume Next
            SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
            
            If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
                Dim contadorErro
                contadorErro = 0
                
                If Not PIR_CORPORATIVO.EOF Then
                    Do While Not PIR_CORPORATIVO.EOF
                        ReDim Preserve erroIds(contadorErro)
                        erroIds(contadorErro) = PIR_CORPORATIVO("ID")
                        contadorErro = contadorErro + 1
                        PIR_CORPORATIVO.MoveNext
                    Loop
                End If
                
                PIR_CORPORATIVO.Close()
                SET PIR_CORPORATIVO = NOTHING
            End If
            On Error Goto 0
        End If
    End If  
    
    ' Atualiza PreFTs
    Dim listaConjuntos
    listaConjuntos = ""
    For i = 0 To UBound(itensCalculo)
        If i > 0 Then listaConjuntos = listaConjuntos & ","
        listaConjuntos = listaConjuntos & "'" & itensCalculo(i)("Produto") & "'"
    Next

    orcamentoSkNetIds = ""
    For i = 0 To UBound(itensCalculo)
        If i > 0 Then
            orcamentoSkNetIds = orcamentoSkNetIds & ","
        End If
        orcamentoSkNetIds = orcamentoSkNetIds & itensCalculo(i)("IdOrcamentoSknet")
    Next
    
    If listaConjuntos <> "" Then
        SQL = "UPDATE TSK_IN_PreFTs SET Estado = 0, Obs2 = " & itensCalculo(0)("TipoCalculoNum") & " "
        SQL = SQL & "WHERE Conjunto IN (" & listaConjuntos & ") AND Obs1 IN (" & orcamentoSkNetIds & ")"

        On Error Resume Next
        SET CONEXAO_TRIMBOX = Conecta(ObterConexaoTrimbox(codPlanta), SQL)
        If IsObject(CONEXAO_TRIMBOX) Then
            CONEXAO_TRIMBOX.Close()
            SET CONEXAO_TRIMBOX = NOTHING
        End If
        On Error Goto 0
    End If

    ' Executa cálculo básico
    If countSucesso > 0 Then
        For i = 0 To UBound(itensSucesso)
            SQL = "SP_CALCULO_BASICO_ORCAMENTO " & itensSucesso(i)("IdOrcamentoSknet") & ""

            On Error Resume Next
            SET rsCalculoBasico = Conecta("PIR-CORPORATIVO", SQL)

            If IsObject(rsCalculoBasico) Then
                rsCalculoBasico.Close()
                SET rsCalculoBasico = NOTHING
            End If
            
            On Error Goto 0
        Next
    End If
End Sub


'--------------------------------------
'       FUNÇÕES AUXILIARES
'--------------------------------------

Function ObterConexaoTrimbox(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoTrimbox = "PIR-TRIMBOX"
        Case "04": ObterConexaoTrimbox = "UBE-TRIMBOX"
        Case "05": ObterConexaoTrimbox = "BEN-TRIMBOX"
        Case "06": ObterConexaoTrimbox = "FOR-TRIMBOX"
        Case "07": ObterConexaoTrimbox = "SAQ-TRIMBOX"
        Case Else: ObterConexaoTrimbox = ""
    End Select
End Function

Function ObterNomePlanta(codPlanta)
    Select Case codPlanta
        Case "00": ObterNomePlanta = "PIRAPETINGA"
        Case "04": ObterNomePlanta = "UBERABA"
        Case "05": ObterNomePlanta = "BENTO"
        Case "06": ObterNomePlanta = "FORTALEZA"
        Case "07": ObterNomePlanta = "SAQUAREMA"
        Case Else: ObterNomePlanta = ""
    End Select
End Function
%>