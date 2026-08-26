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
    Case "validarItens"
        resultado = ValidarItens()
    Case "listarTiposItem"
        resultado = ListarTiposItem()
    Case "listarClientesRemessa"
        resultado = ListarClientesRemessa()
    Case "listarClientes"
        resultado = ListarClientes()
    Case "listarProdutos"
        resultado = ListarProdutos()
    Case "listarTes"
        resultado = ListarTes()
    Case "listarCondicoesPagamento"
        resultado = ListarCondicoesPagamento()
    Case "listarGramaturas"
        resultado = ListarGramaturas()
    Case "listarProdutoRoteiro"
        resultado = ListarProdutoRoteiro()
    Case "listarTubetes"
        resultado = ListarTubetes()
    Case "obterValorUnitarioProdutoOrcamento"
        resultado = ObterValorUnitarioProdutoOrcamento()
    Case "obterSaldoCliente"
        resultado = ObterSaldoClienteRequest()
    Case "verificarExisteOrcamento"
        resultado = VerificarExisteOrcamento()
    Case "verificarExisteLotePiloto"
        resultado = VerificarExisteLotePiloto()
    Case "verificarExisteMaquinaInativa"
        resultado = VerificarExisteMaquinaInativa()
    Case "verificarPedidoCompraDuplicado"
        resultado = VerificarPedidoCompraDuplicado()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

Function ValidarItens()
    ' Campos Gerais
    codigoPlanta = Trim(Request("codigoPlanta"))
    codigoCliente = Trim(Request("codigoCliente"))
    codigoClienteRemessa = Trim(Request("codigoClienteRemessa"))
    codigoLoja = Trim(Request("codigoLoja"))
    toleranciaMaior = Request("toleranciaMaior")
    tipoFrete = Request("tipoFrete")

    tes = Trim(Request("tes"))
    poder3 = Request("poder3")
    tesRemessa = Request("tesRemessa")
    validaLimiteCredito = Request("validaLimiteCredito")

    ' Itens
    pedidosJSON = Request("itens")
    pedidosJSON = Replace(pedidosJSON, "[", "")
    pedidosJSON = Replace(pedidosJSON, "]", "")
    pedidosArray = Split(pedidosJSON, "},{")

    quantidadeItens = UBound(pedidosArray)

    ' Arrays que armazenam as informações dos itens
    ReDim tiposItem(quantidadeItens)
    ReDim codigosProduto(quantidadeItens)
    ReDim referenciasItem(quantidadeItens)
    ReDim quantidadesItem(quantidadeItens)
    ReDim valoresUnitarios(quantidadeItens)
    ReDim largurasItem(quantidadeItens)
    ReDim comprimentosItem(quantidadeItens)
    ReDim pesosItem(quantidadeItens)
    ReDim areasItem(quantidadeItens)
    ReDim datasEntrega(quantidadeItens)
    ReDim produtosEstoque(quantidadeItens)
    ReDim itensLotePiloto(quantidadeItens)

    pesoTotalGeralComTolerancia = 0
    valorTotalGeralComTolerancia = 0

    ' LEITURA DOS ITENS E CÁLCULO DOS TOTAIS GERAIS
    For i = 0 To quantidadeItens
        itemJSON = Trim(pedidosArray(i))

        tiposItem(i) = UCase(Trim(ExtrairValorJSON(itemJSON, "itemTipo")))
        codigosProduto(i) = Trim(ExtrairValorJSON(itemJSON, "codigoProduto"))
        referenciasItem(i) = Trim(ExtrairValorJSON(itemJSON, "referencia"))
        quantidadesItem(i) = ExtrairValorJSON(itemJSON, "quantidade")
        valoresUnitarios(i) = ExtrairValorJSON(itemJSON, "valorUnitario")
        largurasItem(i) = ExtrairValorJSON(itemJSON, "largura")
        comprimentosItem(i) = ExtrairValorJSON(itemJSON, "comprimento")
        pesosItem(i) = ExtrairValorJSON(itemJSON, "peso")
        areasItem(i) = ExtrairValorJSON(itemJSON, "areaRefile")
        datasEntrega(i) = Trim(ExtrairValorJSON(itemJSON, "dataEntrega"))
        produtosEstoque(i) = Trim(ExtrairValorJSON(itemJSON, "produtoEstoque"))
        itensLotePiloto(i) = Iif(ExtrairValorJSON(itemJSON, "lotePiloto") = "1", True, False)
        quantidadesItem(i) = CDbl(quantidadesItem(i))

        ' Precisamos substituir . por , por conta da configuração de decimais do IIS
        valoresUnitarios(i) = CDbl(Replace(valoresUnitarios(i), ".", ","))
        pesosItem(i) = CDbl(Replace(pesosItem(i), ".", ","))
        areasItem(i) = CDbl(Replace(areasItem(i), ".", ","))

        ' Calcula Totais
        quantidadeComTolerancia = quantidadesItem(i) * (1 + (toleranciaMaior / 100))

        ' Papel e lote piloto não entram nas validações de saldo
        If tiposItem(i) <> "PA" AND NOT itensLotePiloto(i) Then
            pesoTotalGeralComTolerancia = pesoTotalGeralComTolerancia + (pesosItem(i) * quantidadeComTolerancia)
            valorTotalGeralComTolerancia = valorTotalGeralComTolerancia + (valoresUnitarios(i) * quantidadeComTolerancia)
        End If

        ' Se item for lote piloto precisamos validar
        If itensLotePiloto(i) Then
            Set respostaLotePiloto = ValidarLotePiloto( _
                codigoCliente, _
                codigoLoja, _
                codigosProduto(i), _
                codigoPlanta, _
                quantidadesItem(i) _
            )

            If Not respostaLotePiloto("sucesso") Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciasItem(i) & ": " & respostaLotePiloto("mensagem")) & """}"
                Exit Function
            End If

            If Not respostaLotePiloto("valido") Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciasItem(i) & ": " & respostaLotePiloto("mensagem")) & """}"
                Exit Function
            End If

            Set respostaLotePiloto = Nothing
        End If
    Next

    ' Validação Geral para Saldo de Aparas (Todos os itens que não são nem lote piloto nem papel)
    If poder3 = "S" Then
        Set respostaSaldoAparas = ObterSaldoAparasCliente(codigoCliente, codigoPlanta)

        If Not respostaSaldoAparas("sucesso") Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON(respostaSaldoAparas("mensagem")) & """}"
            Exit Function
        End If

        If respostaSaldoAparas("saldo") = 0 Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Saldo de Aparas está zerado.") & """}"
            Exit Function
        End If

        If respostaSaldoAparas("saldo") < pesoTotalGeralComTolerancia Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Quantidade emitida no Pedido ultrapassa o saldo de Aparas de " & respostaSaldoAparas("saldo") & ". Favor ajustar os itens do pedido.") & """}"
            Exit Function
        End If

        Set respostaSaldoAparas = Nothing
    End If

    ' Validação Geral para Saldo Comum (Todos os itens que não são nem lote piloto nem papel)
    If tesRemessa = "N" AND CStr(validaLimiteCredito) = "S" Then
        Set respostaSaldo = ObterSaldoCliente(codigoCliente, codigoLoja, codigoPlanta)

        If Not respostaSaldo("sucesso") Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON(respostaSaldo("mensagem")) & """}"
            Exit Function
        End If

        If respostaSaldo("saldo") < valorTotalGeralComTolerancia Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Quantidade emitida no Pedido ultrapassa o saldo de " & respostaSaldo("saldo") & ". Favor ajustar os itens do pedido.") & """}"
            Exit Function
        End If

        Set respostaSaldo = Nothing
    End If

    ' Parâmetros da planta
    Set respostaParametros = ObterParametrosPlanta(codigoPlanta)
    If Not respostaParametros("sucesso") Then
        ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON(respostaParametros("mensagem")) & """}"
        Exit Function
    End If

    Set parametrosPlanta = respostaParametros("dados")
    larguraMinima = parametrosPlanta("larguraMin")
    larguraMaxima = parametrosPlanta("larguraMax")
    comprimentoMinimo = parametrosPlanta("comprimentoMin")
    comprimentoMaximo = parametrosPlanta("comprimentoMax")
    diasViagem = parametrosPlanta("diasViagem")
    validaCargaMaquina = parametrosPlanta("cargaMaquina")

    ' VALIDAÇÕES ESPECÍFICAS DOS ITENS
    For i = 0 To quantidadeItens
        tipoItem = tiposItem(i)
        codigoProduto = codigosProduto(i)
        referenciaItem = referenciasItem(i)
        quantidadeItem = quantidadesItem(i)
        valorUnitario = valoresUnitarios(i)
        larguraItem = largurasItem(i)
        comprimentoItem = comprimentosItem(i)
        pesoItem = pesosItem(i)
        areaItem = areasItem(i)
        dataEntrega = datasEntrega(i)
        produtoEstoque = produtosEstoque(i)
        ehLotePiloto = itensLotePiloto(i)

        ' Reinicia os parâmetros específicos do item
        dataBloqueioMaquinaInicio = ""
        dataBloqueioMaquinaFim = ""
        bloqueiaMaquinaPeriodo = 0
        diasAcrescimo = 0
        pesoMinimo = 0
        areaMinima = 0
        precoMinimo = 0

        ' Parâmetros por tipo de item

        If tipoItem = "PA" Then
            ' Papel
            diasAcrescimo = parametrosPlanta("diasAcrescimoPA")
            pesoMinimo = parametrosPlanta("pesoCH")

        ElseIf tipoItem = "CH" Then
            ' Chapa
            bloqueiaMaquinaPeriodo = parametrosPlanta("bloqueiaMaquinaChapa")
            dataBloqueioMaquinaInicio = parametrosPlanta("deChapa")
            dataBloqueioMaquinaFim = parametrosPlanta("ateChapa")
            diasAcrescimo = parametrosPlanta("diasAcrescimoCH")
            pesoMinimo = parametrosPlanta("pesoCH")
            precoMinimo = parametrosPlanta("precoKgChapa")
            areaMinima = parametrosPlanta("mtsCHmin")

        ElseIf tipoItem = "CX" OR tipoItem = "CJ" Then
            diasAcrescimo = parametrosPlanta("diasAcrescimoCX")

            ' Corte Vinco
            If tipoProduto = "CAIXA VINCO" Then
                bloqueiaMaquinaPeriodo = parametrosPlanta("bloqueiaMaquinaCorteVinco")
                dataBloqueioMaquinaInicio = parametrosPlanta("deCorteVinco")
                dataBloqueioMaquinaFim = parametrosPlanta("ateCorteVinco")
                pesoMinimo = parametrosPlanta("pesoCV")
                areaMinima = parametrosPlanta("mtsCVmin")
                precoMinimo = parametrosPlanta("precoKgCorteVinco")

            Else
            ' Caixa
                bloqueiaMaquinaPeriodo = parametrosPlanta("bloqueiaMaquinaCaixa")
                dataBloqueioMaquinaInicio = parametrosPlanta("deMaleta")
                dataBloqueioMaquinaFim = parametrosPlanta("ateMaleta")
                pesoMinimo = parametrosPlanta("pesoMN")
                areaMinima = parametrosPlanta("mtsCXmin")
                precoMinimo = parametrosPlanta("precoKgMaleta")
            End If
        End If

        ' Totais do Item
        pesoTotal = pesoItem * quantidadeItem
        areaTotal = areaItem * quantidadeItem
        valorTotal = valorUnitario * quantidadeItem

        quantidadeComTolerancia = quantidadeItem * (1 + (toleranciaMaior / 100))

        pesoTotalComTolerancia = pesoItem * quantidadeComTolerancia
        areaTotalComTolerancia = areaItem * quantidadeComTolerancia
        valorTotalComTolerancia = valorUnitario * quantidadeComTolerancia

        ' Validamos area minima e peso minimo
        ' Não validamos área e peso para itens que forem lote piloto ou papel
        If tipoItem <> "PA" AND NOT ehLotePiloto Then
            If CLng(pesoTotal) < CLng(pesoMinimo) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": peso total menor que o mínimo de " & pesoMinimo & " kg.") & """}"
                Exit Function
            End If

            If CLng(areaTotal) < CLng(areaMinima) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": área total menor que o mínimo de " & areaMinima & " m2.") & """}"
                Exit Function
            End If
        End If

        ' Validações para chapas
        If tipoItem = "CH" Then
            If CLng(larguraItem) < CLng(larguraMinima) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": largura mínima permitida: " & larguraMinima & " mm.") & """}"
                Exit Function
            End If

            If CLng(larguraItem) > CLng(larguraMaxima) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": largura máxima permitida: " & larguraMaxima & " mm.") & """}"
                Exit Function
            End If

            If CLng(comprimentoItem) < CLng(comprimentoMinimo) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": comprimento mínimo permitido: " & comprimentoMinimo & " mm.") & """}"
                Exit Function
            End If

            If CLng(comprimentoItem) > CLng(comprimentoMaximo) Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": comprimento máximo permitido: " & comprimentoMaximo & " mm.") & """}"
                Exit Function
            End If

            If CLng(larguraItem) > 250 AND CLng(larguraItem) < 300 AND CLng(comprimentoItem) < 1000 Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": para largura entre 250 e 300 mm, o comprimento mínimo é 1000 mm.") & """}"
                Exit Function
            End If

            If CLng(larguraItem) >= 300 AND CLng(comprimentoItem) < 700 AND codigoPlanta <> "06" Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": para largura a partir de 300 mm, o comprimento mínimo é 700 mm.") & """}"
                Exit Function
            End If
        End If

        ' Validações Gerais de Data
        partesData = Split(dataEntrega, "-")
        dataEntregaVB = DateSerial(CInt(partesData(0)), CInt(partesData(1)), CInt(partesData(2)))

        If Weekday(dataEntregaVB, vbSunday) = 1 Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": a data de entrega não pode ser domingo.") & """}"
            Exit Function
        End If

        If codigoPlanta = "07" AND Weekday(dataEntregaVB, vbSunday) = 7 Then
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": para a planta 07, a data de entrega não pode ser sábado.") & """}"
            Exit Function
        End If

        dataMinima = Date()

        If CInt(diasAcrescimo) > 0 And produtoEstoque <> "1" Then
            dataMinima = DateAdd("d", CInt(diasAcrescimo), dataMinima)
        End If

        If DateValue(dataEntregaVB) < DateValue(dataMinima) Then
            dataMinima = DateAdd("d", 1, dataMinima)
            ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": data de Entrega Indisponível. Data mínima do Cliente será: " & dataMinima & ".") & """}"
            Exit Function
        End If

        ' Validação de Carga Máquina
        If validaCargaMaquina = "1" Then
            ' Parâmetro que bloqueia máquina em um período específico
            If bloqueiaMaquinaPeriodo = "1" Then
                If IsDate(dataEntregaVB) And IsDate(dataBloqueioMaquinaInicio) And IsDate(dataBloqueioMaquinaFim) Then
                    If DateValue(dataEntregaVB) >= DateValue(dataBloqueioMaquinaInicio) And DateValue(dataEntregaVB) <= DateValue(dataBloqueioMaquinaFim) Then
                        ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": carga Máquina para Caixas fechada no período de: " & FormatarData(dataBloqueioMaquinaInicio) & " até " & FormatarData(dataBloqueioMaquinaFim) & ".") & """}"
                        Exit Function
                    End If
                End If
            End If

            Set dadosCargaMaquina = ValidarCargaMaquina( _
                codigoPlanta, _
                codigoProduto, _
                tipoFrete, _
                codigoCliente, _
                codigoClienteRemessa, _
                quantidadeItem, _
                dataEntrega, _
                produtoEstoque _
            )

            If Not dadosCargaMaquina("sucesso") Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": " & dadosCargaMaquina("mensagem")) & """}"
                Exit Function
            End If

            If dadosCargaMaquina("cargaMaquinaFechada") Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": carga maquina fechada.") & """}"
                Exit Function
            End If

            If dadosCargaMaquina("disponivel") <> 1 Then
                ValidarItens = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON("Item " & referenciaItem & ": data de Entrega Indisponível. Próxima Data Disponível: " & dadosCargaMaquina("dataRetorno") & ".") & """}"
                Exit Function
            End If

            Set dadosCargaMaquina = Nothing
        End If
    Next

    Set parametrosPlanta = Nothing
    Set respostaParametros = Nothing
    ValidarItens = "{""sucesso"": true, ""mensagem"": ""Itens validados com sucesso.""}"
End Function

Function Listar()
    pedido = Request("filtroPedido")
    codigoPlanta = Request("filtroCodigoPlanta")
    dataEmissaoDe = Request("filtroDataEmissaoDe")
    dataEmissaoAte = Request("filtroDataEmissaoAte")
    codigoCliente = Request("filtroCodigoCliente")

    SQL = "SELECT " 
    SQL = SQL & "   Pedido, " 
    SQL = SQL & "   Planta as CodigoPlanta, " 
    SQL = SQL & "   '' as NomeCliente, "
    SQL = SQL & "   Cliente as CodigoCliente, " 
    SQL = SQL & "   LojaCli as CodigoLoja, " 
    SQL = SQL & "   Status, " 
    SQL = SQL & "   LotePiloto, " 
    SQL = SQL & "   ClienteRemessa, " 
    SQL = SQL & "   Produto as CodProduto, " 
    SQL = SQL & "   Referencia as Referencia, " 
    SQL = SQL & "   Estrutura, " 
    SQL = SQL & "   ISNULL(PrecoUnitario, 0) ValorUnitario, " 
    SQL = SQL & "   Quantidade, " 
    SQL = SQL & "   Observacao, " 
    SQL = SQL & "   Largura, " 
    SQL = SQL & "   Comprimento, " 
    SQL = SQL & "   DataEmissao, " 
    SQL = SQL & "   PrimeiraDataEntrega as DataEntrega, " 
    SQL = SQL & "   DescricaoBloqueio, " 
    SQL = SQL & "   NumeroOrdemCompra, " 
    SQL = SQL & "   ItemOrdemCompra " 
    SQL = SQL & "FROM Pedido " 
    SQL = SQL & "WHERE ISNULL(Status, 0) = 0 " 
    SQL = SQL & "  AND ISNULL(Excluido, 0) = 0 " 
    SQL = SQL & "  AND QuemIncluiu = '" & usuarioLogado("id") & "' " 

    If codigoCliente <> "" Then
        SQL = SQL & "AND Cliente = " & codigoCliente & " "
    End If

    If pedido <> "" Then 
        SQL = SQL & "AND Pedido = " & pedido & " "
    End If

    If dataEmissaoDe <> "" Then
        SQL = SQL & " AND DataEmissao >= '" & Replace(dataEmissaoDe, "-", "") & "'"
    End If

    If dataEmissaoAte <> "" Then
        SQL = SQL & " AND DataEmissao < DATEADD(DAY, 1, '" & Replace(dataEmissaoAte & "", "-", "") & "')"
    End If

    SQL = SQL & "ORDER BY Pedido DESC "

    On Error Resume Next
    SET rsListar = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de Pedidos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """pedido"": """ & EscapeJSON(rsListar("Pedido")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsListar("NomeCliente")) & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & EscapeJSON(rsListar("CodigoCliente")) & """, "
            jsonArray = jsonArray & """codigoLoja"": """ & EscapeJSON(rsListar("CodigoLoja")) & """, "
            jsonArray = jsonArray & """codigoPlanta"": """ & EscapeJSON(rsListar("CodigoPlanta")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """lotePiloto"": """ & EscapeJSON(rsListar("LotePiloto")) & """, "
            jsonArray = jsonArray & """clienteRemessa"": """ & EscapeJSON(rsListar("ClienteRemessa")) & """, "
            jsonArray = jsonArray & """codProduto"": """ & EscapeJSON(rsListar("CodProduto")) & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsListar("Referencia")) & """, "
            jsonArray = jsonArray & """estrutura"": """ & EscapeJSON(rsListar("Estrutura")) & """, "
            jsonArray = jsonArray & """valorUnitario"": """ & Replace(EscapeJSON(rsListar("ValorUnitario")), ",", ".") & """, "
            jsonArray = jsonArray & """quantidade"": """ & EscapeJSON(rsListar("Quantidade")) & """, "
            jsonArray = jsonArray & """observacao"": """ & EscapeJSON(rsListar("Observacao")) & """, "
            jsonArray = jsonArray & """largura"": """ & EscapeJSON(rsListar("Largura")) & """, "
            jsonArray = jsonArray & """comprimento"": """ & EscapeJSON(rsListar("Comprimento")) & """, "
            jsonArray = jsonArray & """dataEmissao"": """ & EscapeJSON(rsListar("DataEmissao")) & """, "
            jsonArray = jsonArray & """dataEntrega"": """ & EscapeJSON(rsListar("DataEntrega")) & """, "
            jsonArray = jsonArray & """descricaoBloqueio"": """ & EscapeJSON(rsListar("DescricaoBloqueio")) & """, "
            jsonArray = jsonArray & """numeroOrdemCompra"": """ & EscapeJSON(rsListar("NumeroOrdemCompra")) & """, "
            jsonArray = jsonArray & """itemOrdemCompra"": """ & EscapeJSON(rsListar("ItemOrdemCompra")) & """ "
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

Function Criar()
    codigoCliente = Request("codigoCliente")
    codigoLoja = Request("codigoLoja")
    codigoPlanta = Request("codigoPlanta")
    tesRemessa = Request("tesRemessa")
    TES = Request("tes")
    validaLimiteCredito = Request("validaLimiteCredito")
    toleranciaMaior = Request("toleranciaMaior")
    toleranciaDataMaior = Request("toleranciaDataMaior")
    codigoClienteRemessa = Request("codigoClienteRemessa")
    codigoLojaRemessa = Request("codigoLojaRemessa")
    tipoFrete = Request("tipoFrete")
    condicaoPagamento = Request("condicaoPagamento")
    moeda = Request("moeda")

    quantidadeInserida = 0
    quantidadeErros = 0
    resultadosJSON = ""
    pedidosEmail = ""

    Select Case codigoPlanta
        Case "05": nomePlanta = "BENTO"
        Case "06": nomePlanta = "MARANGUAPE"
        Case "07": nomePlanta = "SAQUAREMA"
        Case Else: nomePlanta = ""
    End Select

    pedidosJSON = Request("itens")
    pedidosJSON = Replace(pedidosJSON, "[", "")
    pedidosJSON = Replace(pedidosJSON, "]", "")

    If Trim(pedidosJSON) = "" Then
        Criar = "{""sucesso"":false,""sucessoParcial"":false,""mensagem"":""Nenhum item foi enviado."",""quantidadeInserida"":0,""quantidadeErros"":0,""resultados"":[]}"
        Exit Function
    End If

    pedidosArray = Split(pedidosJSON, "},{")

    SQL = "SELECT COD_VENDEDOR, ESTADO_CLIENTE FROM VW_CLIENTES_PEDIDO WHERE COD_CLIENTE = '" & codigoCliente & "'"
    On Error Resume Next
    Set rsRepresentante = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    If Err.Number <> 0 Or Not IsObject(rsRepresentante) Then
        Criar = "{""sucesso"":false,""sucessoParcial"":false,""mensagem"":""Erro ao consultar o representante: " & EscapeJSON(Err.Description) & """,""quantidadeInserida"":0,""quantidadeErros"":0,""resultados"":[]}"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    codigoRepresentante = ValorOuVazio(rsRepresentante("COD_VENDEDOR"))
    estado = UCase(ValorOuVazio(rsRepresentante("ESTADO_CLIENTE")))
    rsRepresentante.Close()
    Set rsRepresentante = Nothing

    SQL = "SELECT CALC_IPI, CALC_ICM, PIS_COF, ISENCAO_FECP, "
    SQL = SQL & "DIFERE_ICMS, RED_BASE_ICMS, PERC_DIF "
    SQL = SQL & "FROM VW_BI_TES WHERE CODIGO = '" & TES & "'"

    On Error Resume Next
    Set rsTES = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    If Err.Number <> 0 Or Not IsObject(rsTES) Then
        Criar = "{""sucesso"":false,""sucessoParcial"":false,""mensagem"":""Erro ao consultar a TES: " & EscapeJSON(Err.Description) & """,""quantidadeInserida"":0,""quantidadeErros"":0,""resultados"":[]}"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    calculaIPI = UCase(ValorOuVazio(rsTES("CALC_IPI")))
    calculaICMS = UCase(ValorOuVazio(rsTES("CALC_ICM")))
    isentoFECP = UCase(ValorOuVazio(rsTES("ISENCAO_FECP")))
    difereICMS = UCase(ValorOuVazio(rsTES("DIFERE_ICMS")))
    percentualBaseICMS = NumeroVB(rsTES("RED_BASE_ICMS"))
    percentualDIF = NumeroVB(rsTES("PERC_DIF"))

    Select Case UCase(ValorOuVazio(rsTES("PIS_COF")))
        Case "PIS": calculaPISCofins = "P"
        Case "COFINS": calculaPISCofins = "C"
        Case "AMBOS": calculaPISCofins = "A"
        Case Else: calculaPISCofins = "N"
    End Select

    rsTES.Close()
    Set rsTES = Nothing

    For i = 0 To UBound(pedidosArray)
        itemJSON = Trim(pedidosArray(i))
        idTemporario = ExtrairValorJSON(itemJSON, "idTemporario")
        numeroPedidoAtual = ""
        mensagemItem = ""
        pedidoGravado = False

        pedidoValido = ProcessarPedido( _
            itemJSON, codigoCliente, codigoLoja, codigoPlanta, _
            codigoClienteRemessa, codigoLojaRemessa, nomePlanta, TES, _
            tesRemessa, tipoFrete, condicaoPagamento, moeda, _
            toleranciaMaior, toleranciaDataMaior, validaLimiteCredito, _
            codigoRepresentante, estado, calculaIPI, calculaICMS, _
            calculaPISCofins, isentoFECP, difereICMS, percentualBaseICMS, _
            percentualDIF, numeroPedidoAtual, mensagemItem, pedidoGravado _
        )

        If resultadosJSON <> "" Then resultadosJSON = resultadosJSON & ","

        If pedidoGravado Then
            quantidadeInserida = quantidadeInserida + 1

            If pedidoValido Then
                statusItem = "GRAVADO"
                If mensagemItem = "" Then
                    mensagemItem = "Pedido inserido com sucesso."
                End If

                If pedidosEmail <> "" Then
                    pedidosEmail = pedidosEmail & ","
                End If
                pedidosEmail = pedidosEmail & "'" & Replace(CStr(numeroPedidoAtual), "'", "''") & "'"
            Else
                statusItem = "GRAVADO_COM_PENDENCIA"
                If mensagemItem = "" Then mensagemItem = "Pedido gravado, mas houve uma pendencia no pos-processamento."
            End If

            resultadosJSON = resultadosJSON & "{" & _
                """idTemporario"":""" & EscapeJSON(idTemporario) & """," & _
                """gravado"":true," & _
                """status"":""" & statusItem & """," & _
                """pedido"":""" & EscapeJSON(CStr(numeroPedidoAtual)) & """," & _
                """mensagem"":""" & EscapeJSON(mensagemItem) & """}"
        Else
            quantidadeErros = quantidadeErros + 1
            If mensagemItem = "" Then mensagemItem = "Nao foi possivel inserir o item."

            resultadosJSON = resultadosJSON & "{" & _
                """idTemporario"":""" & EscapeJSON(idTemporario) & """," & _
                """gravado"":false," & _
                """status"":""NAO_GRAVADO""," & _
                """pedido"":""""," & _
                """mensagem"":""" & EscapeJSON(mensagemItem) & """}"
        End If
    Next

    If pedidosEmail <> "" Then
        Call ExtratoPedidos(codigoPlanta, pedidosEmail)
    End If

    If quantidadeErros = 0 Then
        Criar = "{" & _
            """sucesso"":true," & _
            """sucessoParcial"":false," & _
            """mensagem"":""Todos os pedidos foram inseridos com sucesso.""," & _
            """quantidadeInserida"":" & quantidadeInserida & "," & _
            """quantidadeErros"":0," & _
            """resultados"":" & "[" & resultadosJSON & "]}"
    ElseIf quantidadeInserida > 0 Then
        Criar = "{" & _
            """sucesso"":false," & _
            """sucessoParcial"":true," & _
            """mensagem"":""Alguns pedidos foram inseridos e removidos da lista. Corrija os itens restantes.""," & _
            """quantidadeInserida"":" & quantidadeInserida & "," & _
            """quantidadeErros"":" & quantidadeErros & "," & _
            """resultados"":" & "[" & resultadosJSON & "]}"
    Else
        Criar = "{" & _
            """sucesso"":false," & _
            """sucessoParcial"":false," & _
            """mensagem"":""Nenhum pedido foi inserido. Corrija os itens indicados e tente novamente.""," & _
            """quantidadeInserida"":0," & _
            """quantidadeErros"":" & quantidadeErros & "," & _
            """resultados"":" & "[" & resultadosJSON & "]}"
    End If
End Function

Function ProcessarPedido( _
    ByVal itemJSON, _
    ByVal codigoCliente, _
    ByVal codigoLoja, _
    ByVal codigoPlanta, _
    ByVal codigoClienteRemessa, _
    ByVal codigoLojaRemessa, _
    ByVal nomePlanta, _
    ByVal TES, _
    ByVal tesRemessa, _
    ByVal tipoFrete, _
    ByVal condicaoPagamento, _
    ByVal moeda, _
    ByVal toleranciaMaior, _
    ByVal toleranciaDataMaior, _
    ByVal validaLimiteCredito, _
    ByVal codigoRepresentante, _
    ByVal estado, _
    ByVal calculaIPI, _
    ByVal calculaICMS, _
    ByVal calculaPISCofins, _
    ByVal isentoFECP, _
    ByVal difereICMS, _
    ByVal percentualBaseICMS, _
    ByVal percentualDIF, _
    ByRef numeroPedidoAtual, _
    ByRef mensagemItem, _
    ByRef pedidoGravado _
)
    pedidoValido = True
    pedidoGravado = False
    mensagemItem = ""

    itemTipo = UCase(ExtrairValorJSON(itemJSON, "itemTipo"))
    codigoProduto = ExtrairValorJSON(itemJSON, "codigoProduto")
    referencia = ExtrairValorJSON(itemJSON, "referencia")
    largura = NumeroVB(ExtrairValorJSON(itemJSON, "largura"))
    comprimento = NumeroVB(ExtrairValorJSON(itemJSON, "comprimento"))
    vinco1 = NumeroVB(ExtrairValorJSON(itemJSON, "vinco1"))
    vinco2 = NumeroVB(ExtrairValorJSON(itemJSON, "vinco2"))
    vinco3 = NumeroVB(ExtrairValorJSON(itemJSON, "vinco3"))
    vinco4 = NumeroVB(ExtrairValorJSON(itemJSON, "vinco4"))
    vinco5 = NumeroVB(ExtrairValorJSON(itemJSON, "vinco5"))
    gramatura = ExtrairValorJSON(itemJSON, "gramatura")
    tubeteBobina = ExtrairValorJSON(itemJSON, "tubeteBobina")
    diametroBobina = ExtrairValorJSON(itemJSON, "diametroBobina")
    larguraBobina = ExtrairValorJSON(itemJSON, "larguraBobina")
    quantidade = NumeroVB(ExtrairValorJSON(itemJSON, "quantidade"))
    valorUnitario = NumeroVB(ExtrairValorJSON(itemJSON, "valorUnitario"))
    dataEntrega = ExtrairValorJSON(itemJSON, "dataEntrega")
    dataEntrega2 = AdicionarDias(dataEntrega, toleranciaDataMaior)
    dataPrevistaEntrega = ExtrairValorJSON(itemJSON, "dataPrevistaEntrega")
    pedidoCompra = ExtrairValorJSON(itemJSON, "pedidoCompra")
    itemPedido = ExtrairValorJSON(itemJSON, "itemPedido")
    observacao = ExtrairValorJSON(itemJSON, "observacao")
    modalFrete = ExtrairValorJSON(itemJSON, "modalFrete")
    comercioExterior = ExtrairValorJSON(itemJSON, "comercioExterior")
    ehLotePiloto =  Iif(ExtrairValorJSON(itemJSON, "lotePiloto") = "1", True, False)

    If dataPrevistaEntrega = "" Then
        dataPrevistaEntrega = dataEntrega
    End If

    numeroPedidoAtual = ObterUltimoPedido(codigoPlanta)

    If numeroPedidoAtual = "" Then
        pedidoValido = False
        mensagemItem = "Nao foi possivel gerar o numero do pedido."
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    ' BUSCA PRODUTO, ESTRUTURA, PESO E IPI
    If pedidoValido Then
        percentualIPI = 0
        estrutura = ""
        peso = 0

        stringConexao = ""
        Select Case itemTipo
            Case "CJ"
                stringConexao = ObterConexaoTrimbox(codigoPlanta)

                SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO " 
                SQL = SQL & "FROM VW_BI_PRODUTOS_PEDIDO " 
                SQL = SQL & "WHERE CODIGO_CONJUNTO = '" & codigoProduto & "' " 
                SQL = SQL & "ORDER BY CODIGO_ITEM"
            Case "PA"
                If codigoPlanta = "05" then
                    stringConexao = "BEN-DADOSADV"
                Else
                    stringConexao = ObterConexaoTrimbox(codigoPlanta)  
                End If

                SQL = "SELECT ALIQUOTA_IPI, ESTRUTURA, PESO "
                SQL = SQL & "FROM VW_PAPEL_PEDIDO "
                SQL = SQL & "WHERE CODIGO_ITEM = '" & codigoProduto & "'"
            Case Else
                stringConexao = ObterConexaoTrimbox(codigoPlanta)

                SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO " 
                SQL = SQL & "FROM VW_BI_PRODUTOS_PEDIDO " 
                SQL = SQL & "WHERE CODIGO_CONJUNTO = '" & codigoProduto & "'"
        End Select

        On Error Resume Next
        Set Consulta = Conecta(stringConexao, SQL)

        If Err.Number <> 0 Or Not IsObject(Consulta) Then
            pedidoValido = False
            mensagemItem = "Erro ao consultar o produto: " & Err.Description
            Err.Clear
        ElseIf Consulta.EOF Then
            pedidoValido = False
            mensagemItem = "Produto " & codigoProduto & " nao encontrado."
        Else
            percentualIPI = NumeroVB(Consulta("ALIQUOTA_IPI"))
            estrutura = ValorOuVazio(Consulta("ESTRUTURA"))
            peso = NumeroVB(Consulta("PESO"))
        End If

        Consulta.Close()
        Set Consulta = Nothing
        On Error GoTo 0
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    ' RECALCULA PESO PARA CHAPA
    If pedidoValido AND itemTipo = "CH" Then
        If largura > 0 AND comprimento > 0 Then
            SQL = "SELECT PESO FROM dbo.FUNC_CALCULA_AREA_PESO(" & SQLNumero(largura) & "," & SQLNumero(comprimento) & ", '" & estrutura & "')"

            On Error Resume Next
            Set Consulta = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)

            If Err.Number <> 0 Or Not IsObject(Consulta) Then
                pedidoValido = False
                mensagemItem = "Erro ao recalcular o peso: " & Err.Description
                Err.Clear
            ElseIf Not Consulta.EOF Then
                peso = NumeroVB(Consulta("PESO"))
            End If

            Consulta.Close()
            Set Consulta = Nothing
            On Error GoTo 0
        End If
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    ' ----------------------------------------------------------------
    ' BUSCA ALIQUOTA DE ICMS
    ' ----------------------------------------------------------------
    If pedidoValido Then
        percentualICMS = 0

        SQL = "SELECT dbo.FUNC_PERCENTUAL_ICMS('" & nomePlanta & "', '" & estado & "') AS AliquotaICMS"

        On Error Resume Next
        Set Consulta = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)

        If Err.Number <> 0 Or Not IsObject(Consulta) Then
            pedidoValido = False
            mensagemItem = "Erro ao consultar a aliquota de ICMS: " & Err.Description
            Err.Clear
        ElseIf Not Consulta.EOF Then
            percentualICMS = NumeroVB(Consulta("AliquotaICMS"))

            If estado = "RJ" AND isentoFECP = "N" Then
                percentualICMS = percentualICMS + 2
            End If
        End If

        Consulta.Close()
        Set Consulta = Nothing
        On Error GoTo 0
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If


    ' ----------------------------------------------------------------
    ' CALCULO DOS IMPOSTOS DO ITEM
    ' ----------------------------------------------------------------
    If pedidoValido Then
        valorBaseIPIUnitario = 0
        valorIPIUnitario = 0
        valorBaseICMSUnitario = 0
        valorICMSUnitario = 0
        percentualBasePIS = 0
        percentualPIS = 0
        valorPISUnitario = 0
        percentualBaseCOFINS = 0
        percentualCOFINS = 0
        valorCOFINSUnitario = 0

        If calculaIPI = "S" Then
            valorBaseIPIUnitario = valorUnitario
            valorIPIUnitario = valorUnitario * percentualIPI / 100
        End If

        If calculaICMS = "S" AND difereICMS = "N" Then
            valorBaseICMSUnitario = valorUnitario * percentualBaseICMS / 100
            valorICMSUnitario = valorBaseICMSUnitario * percentualICMS / 100
        End If

        If calculaICMS = "S" AND difereICMS = "S" Then
            valorBaseICMSUnitario = valorUnitario
            icmsIntegral = valorBaseICMSUnitario * percentualICMS / 100

            If percentualDIF < 100 Then
                valorICMSDIF = icmsIntegral * percentualDIF / 100
                valorICMSUnitario = icmsIntegral - valorICMSDIF
            Else
                valorICMSUnitario = 0
            End If
        End If

        If calculaPISCofins <> "N" Then
            basePISCofins = valorUnitario - valorICMSUnitario

            If calculaPISCofins = "P" Or calculaPISCofins = "A" Then
                percentualPIS = 1.65
                valorPISUnitario = basePISCofins * percentualPIS / 100
            End If

            If calculaPISCofins = "C" Or calculaPISCofins = "A" Then
                percentualCOFINS = 7.6
                valorCOFINSUnitario = basePISCofins * percentualCOFINS / 100
            End If
        End If
    End If
      
    ' ----------------------------------------------------------------
    ' VALIDACAO DO LIMITE DE CREDITO
    ' ----------------------------------------------------------------
    If pedidoValido AND validaLimiteCredito = "S" Then
        Set respostaSaldo = ObterLimiteCredito(codigoCliente, codigoLoja, codigoPlanta)
        If Not respostaSaldo("sucesso") Then
            pedidoValido = False
            mensagemItem = "Erro obter limite de crédito: " & Err.Description
        End If

        saldoDisponivel = respostaSaldo("saldo")
        quantidadeComTolerancia = quantidade

        If tolerancia > 0 Then
            quantidadeComTolerancia = quantidadeComTolerancia + (quantidadeComTolerancia * tolerancia / 100)
        End If

        valorComprometido = _
            (quantidadeComTolerancia * valorUnitario) + _
            (quantidadeComTolerancia * valorIPIUnitario)

        saldoDisponivel = NumeroVB(saldoDisponivel) - valorComprometido

        If saldoDisponivel < 0 Then
            pedidoValido = False
            mensagemItem = "Cliente nao possui limite suficiente para o produto " & codigoProduto & "."
        End If
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    ' ----------------------------------------------------------------
    ' VALIDACAO DE LOTE PILOTO
    ' ----------------------------------------------------------------
    atualizaLotePiloto = False
    If pedidoValido AND ehLotePiloto Then
        Set respostaLotePiloto = ValidarLotePiloto( _
            codigoCliente, _
            codigoLoja, _
            codigoProduto, _
            codigoPlanta, _
            quantidade _
        )

        If Not respostaLotePiloto("sucesso") OR Not respostaLotePiloto("valido") Then
            pedidoValido = False
            mensagemItem = "Erro de validação de lote piloto para o item " & codigoProduto & ": " & respostaLotePiloto("mensagem")
        End If

        If respostaLotePiloto("observacoes") <> "" Then
            observacao = observacao & " - " & respostaLotePiloto("observacoes")
        End If

        If respostaLotePiloto("atualizaStatus") = 1 Then
            atualizaLotePiloto = True
        End If
    End If

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    centroCusto = "0"
    contaContabil = "0"
    depositoOriginal = "0"
    codigoTabelaPreco = "1"

    If TES = "245" OR TES = "272" OR TES = "288" Then
        codigoTabelaPreco = "35"
    End If

    If codigoPlanta <> "07" Then
        Select Case itemTipo
            Case "CX"
                centroCusto = "508"
                depositoOriginal = "508"
                contaContabil = "900014"
            Case "CH"
                centroCusto = "509"
                depositoOriginal = "509"
                contaContabil = "900015"
            Case "PA"
                centroCusto = "503"
                depositoOriginal = "503"
                contaContabil = "900016"
        End Select
    End If

    ' A partir daqui começam os inserts e updates
    SQL = "INSERT INTO Pedido ("
    SQL = SQL & "    Pedido,"
    SQL = SQL & "    Planta, "
    SQL = SQL & "    TESFatura,"
    SQL = SQL & "    DataEmissao,"
    SQL = SQL & "    PrimeiraDataEntrega,"
    SQL = SQL & "    SegundaDataEntrega,"
    SQL = SQL & "    Cliente,"
    SQL = SQL & "    FaturaRemessa,"
    SQL = SQL & "    TipoFrete,"
    SQL = SQL & "    Representante,"
    SQL = SQL & "    CondicaoPagamento,"
    SQL = SQL & "    TabelaComissao,"
    SQL = SQL & "    Moeda,"
    SQL = SQL & "    ZonaVenda,"
    SQL = SQL & "    PesoLiquido,"
    SQL = SQL & "    Observacao,"
    SQL = SQL & "    GeradoRepresentante,"
    SQL = SQL & "    TESRemessa,"
    SQL = SQL & "    Status,"
    SQL = SQL & "    ClienteRemessa,"
    SQL = SQL & "    Produto,"
    SQL = SQL & "    Referencia,"
    SQL = SQL & "    Largura,"
    SQL = SQL & "    Comprimento,"
    SQL = SQL & "    Vinco1,"
    SQL = SQL & "    Vinco2,"
    SQL = SQL & "    Vinco3,"
    SQL = SQL & "    Vinco4,"
    SQL = SQL & "    Vinco5,"
    SQL = SQL & "    Tubete,"
    SQL = SQL & "    Diametro,"
    SQL = SQL & "    Gramatura,"
    SQL = SQL & "    Quantidade,"
    SQL = SQL & "    PrecoUnitario,"
    SQL = SQL & "    CentroCusto,"
    SQL = SQL & "    CodigoOperacaoContaContabil,"
    SQL = SQL & "    NumeroOrdemCompra,"
    SQL = SQL & "    NumeroIPI,"
    SQL = SQL & "    TabelaPreco,"
    SQL = SQL & "    ItemOrdemCompra,"
    SQL = SQL & "    CST_IPI,"
    SQL = SQL & "    AplicouSuspencaoIPI,"
    SQL = SQL & "    CodigoEnquadramentoIPI,"
    SQL = SQL & "    ValorIPIUnitario,"
    SQL = SQL & "    CST_ICMS,"
    SQL = SQL & "    ValorBaseICMSUnitario,"
    SQL = SQL & "    PercentualBaseICMS,"
    SQL = SQL & "    ValorICMSUnitario,"
    SQL = SQL & "    ValorICMS6142,"
    SQL = SQL & "    PercentualICMS,"
    SQL = SQL & "    PercentualBasePIS,"
    SQL = SQL & "    ValorPISUnitario,"
    SQL = SQL & "    PercentualPIS,"
    SQL = SQL & "    ValorCOFINSUnitario,"
    SQL = SQL & "    PercentualCOFINS,"
    SQL = SQL & "    PercentualBaseCOFINS,"
    SQL = SQL & "    DepositoOriginal,"
    SQL = SQL & "    QuemIncluiu,"
    SQL = SQL & "    DataInclusao,"
    SQL = SQL & "    QuemValidou,"
    SQL = SQL & "    StatusBloqueio,"
    SQL = SQL & "    DescricaoBloqueio,"
    SQL = SQL & "    QuemBloqueou,"
    SQL = SQL & "    QuemExcluiu,"
    SQL = SQL & "    Excluido,"
    SQL = SQL & "    LotePiloto,"
    SQL = SQL & "    Estrutura,"
    SQL = SQL & "    LarguraBobina,"
    SQL = SQL & "    DescricaoComplementar,"
    SQL = SQL & "    LojaCli,"
    SQL = SQL & "    LojaCliRemessa,"
    SQL = SQL & "    Item,"
    SQL = SQL & "    TipoItem,"

    If codigoPlanta <> "07" Then
        SQL = SQL & "    StatusImportacao,"
    End If

    If codigoPlanta = "05" Then
        SQL = SQL & "    ModalFrete, "
        SQL = SQL & "    ComercioExterior, "
    End If

    SQL = SQL & "    DataPrevistaEntrega "
    SQL = SQL & ") VALUES("
    SQL = SQL & "'" & numeroPedidoAtual & "', " 
    SQL = SQL & "'" & codigoPlanta & "', " 
    SQL = SQL & "'" & TES & "', " 
    SQL = SQL & "GETDATE(), " 
    SQL = SQL & "'" & dataEntrega & "', " 
    SQL = SQL & "'" & dataEntrega2 & "', " 
    SQL = SQL & "'" & codigoCliente & "', " 
    SQL = SQL & "'" & Iif(tesRemessa = "S", "1", "0") & "', " 
    SQL = SQL & "'" & tipoFrete & "', " 
    SQL = SQL & "'" & codigoRepresentante & "', " 
    SQL = SQL & "'" & condicaoPagamento & "', " 
    SQL = SQL & "''," 
    SQL = SQL & "'" & moeda & "', " 
    SQL = SQL & "''," 
    SQL = SQL & "'" & SQLNumero(peso) & "', " 
    SQL = SQL & "'" & observacao & "', " 
    SQL = SQL & "'', "
    SQL = SQL & "'', " ' TODO: TESRemessa retestar
    SQL = SQL & "0,"
    SQL = SQL & "'" & codigoClienteRemessa & "', "
    SQL = SQL & "'" & codigoProduto & "', "
    SQL = SQL & "'" & referencia & "', "
    SQL = SQL & "'" & SQLNumero(largura) & "', "
    SQL = SQL & "'" & SQLNumero(comprimento) & "', "
    SQL = SQL & "'" & SQLNumero(vinco1) & "', "
    SQL = SQL & "'" & SQLNumero(vinco2) & "', "
    SQL = SQL & "'" & SQLNumero(vinco3) & "', "
    SQL = SQL & "'" & SQLNumero(vinco4) & "', "
    SQL = SQL & "'" & SQLNumero(vinco5) & "', "
    SQL = SQL & "'" & tubeteBobina & "', "
    SQL = SQL & "'" & SQLNumero(diametroBobina) & "', "
    SQL = SQL & "'" & SQLNumero(gramatura) & "', "
    SQL = SQL & "'" & SQLNumero(quantidade) & "', "
    SQL = SQL & "'" & SQLNumero(valorUnitario) & "', "
    SQL = SQL & "'" & centroCusto & "', "
    SQL = SQL & "'" & contaContabil & "', "
    SQL = SQL & "'" & pedidoCompra & "', "
    SQL = SQL & "1,"
    SQL = SQL & "'" & codigoTabelaPreco & "', "
    SQL = SQL & "'" & itemPedido & "', "
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "'" & SQLNumero(valorIPIUnitario) & "', "
    SQL = SQL & "0,"
    SQL = SQL & "'" & SQLNumero(valorBaseICMSUnitario) & "', "
    SQL = SQL & "'" & SQLNumero(percentualBaseICMS) & "', "
    SQL = SQL & "'" & SQLNumero(valorICMSUnitario) & "', "
    SQL = SQL & "0,"
    SQL = SQL & "'" & SQLNumero(percentualICMS) & "', "
    SQL = SQL & "'" & SQLNumero(percentualBasePIS) & "', "
    SQL = SQL & "'" & SQLNumero(valorPISUnitario) & "', "
    SQL = SQL & "'" & SQLNumero(percentualPIS) & "', "
    SQL = SQL & "'" & SQLNumero(valorCOFINSUnitario) & "', "
    SQL = SQL & "'" & SQLNumero(percentualCOFINS) & "', "
    SQL = SQL & "'" & SQLNumero(percentualBaseCOFINS) & "', "
    SQL = SQL & "'" & depositoOriginal & "', "
    SQL = SQL & "'" & usuarioLogado("id") & "', "
    SQL = SQL & "GETDATE(),"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "'',"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "'" & Iif(ehLotePiloto, "1", "0") & "', "
    SQL = SQL & "'" & estrutura & "', "
    SQL = SQL & "'" & SQLNumero(larguraBobina) & "', "
    SQL = SQL & "'',"
    SQL = SQL & "'" & codigoLoja & "', "
    SQL = SQL & "'" & codigoLojaRemessa & "', "
    SQL = SQL & "'1',"
    SQL = SQL & "'" & itemTipo & "', " ' A partir daqui campos são diferentes por planta

    If codigoPlanta <> "07" Then
        SQL = SQL & "0, "
    End If
    
    If codigoPlanta = "05" Then
        SQL = SQL & "'" & modalFrete & "', " 
        SQL = SQL & "'" & comercioExterior & "', " 
    End If

    SQL = SQL & "'" & dataPrevistaEntrega & "' " 
    SQL = SQL & ")"

    On Error Resume Next
    SET rsInsert = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)
    
    If Err.Number <> 0 Then
        pedidoValido = False
        mensagemItem = "Erro ao inserir o pedido: " & Err.Description
        Err.Clear
    Else
        pedidoGravado = True
    End If

    If IsObject(rsInsert) Then
        If Not rsInsert Is Nothing Then
            On Error Resume Next
            If rsInsert.State = 1 Then rsInsert.Close()
            On Error GoTo 0
        End If
    End If
    Set rsInsert = Nothing
    On Error GoTo 0

    If Not pedidoValido Then
        ProcessarPedido = False
        Exit Function
    End If

    If pedidoValido AND itemTipo <> "PA" Then
        SQL = "EXEC SP_VALIDA_PEDIDO_COM_ORCAMENTO '" & codigoPlanta & "', '" & numeroPedidoAtual & "' "
        Set rsValida = Conecta("PIR-CORPORATIVO", SQL)

        SQL = "SELECT QuemValidou FROM Pedido " 
        SQL = SQL & "WHERE Pedido = '" & numeroPedidoAtual & "' " 
        SQL = SQL & "AND QuemValidou = '999999' AND DataValidacao IS NOT NULL"

        On Error Resume Next
        Set rsConsulta = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)

        If IsObject(rsConsulta) Then
            If Not rsConsulta Is Nothing Then
                If Not rsConsulta.EOF Then
                    SQL = "EXEC SP_INCLUI_PEDIDO_TRIMBOX_" & nomePlanta & " '" & numeroPedidoAtual & "', '1'"
                    Set rsConsulta = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)
                End If
            End If
        End If

        rsConsulta.Close()
        Set rsConsulta = Nothing
        On Error GoTo 0
    End If

    If atualizaLotePiloto Then
        Call AtualizarLotePiloto(codigoCliente, codigoPlanta, codigoProduto)
    End If

    ProcessarPedido = pedidoValido
End Function


Function AtualizarLotePiloto(codigoCliente, codigoPlanta, codigoProduto)
    'Busca o status atual do lote piloto
    SQL = "SELECT TOP 1 ID, "
    SQL = SQL & " CASE STATUS "
    SQL = SQL & "   WHEN '1' THEN 'Aguardando Pedido LP' "
    SQL = SQL & "   WHEN '2' THEN 'Convertido em Venda'  "
    SQL = SQL & "	WHEN '3' THEN 'Aprovado sem Pedido de Conversao'  "
    SQL = SQL & "	WHEN '4' THEN 'Em Teste Cliente'  "
    SQL = SQL & "	WHEN '5' THEN 'Reprovado Pelo Cliente'  "
    SQL = SQL & "	WHEN '6' THEN 'Reprovado Internamente/Retrabalho'  " 
    SQL = SQL & "	WHEN '7' THEN 'Aguardando Producao'  "
    SQL = SQL & "	WHEN '8' THEN 'Descontinuado'  "
    SQL = SQL & "	WHEN '9' THEN 'Sem Retorno Vendas/Cliente'  "
    SQL = SQL & "	WHEN '10' THEN 'Processo segue em Novo Produto/Ficha Técnica' "
    SQL = SQL & "	WHEN '11' THEN 'Validacao Maquina Nova/Homologaca' END STATUS  "
    SQL = SQL & " FROM LOTE_PILOTO  "
    SQL = SQL & " WHERE CLIENTE = '" & codigoCliente & "' "
    SQL = SQL & "   AND PLANTA =  '" & codigoPlanta & "' "
    SQL = SQL & "   AND APROVACAO = 0  "
    SQL = SQL & "   AND QUEM_EXCLUIU IS NULL " 
    SQL = SQL & "   AND PRODUTO = '" & codigoProduto  & "' " 
    SQL = SQL & " ORDER BY DATEDIFF(DAY,DATA_HORA,GETDATE()) " 

    On Error Resume Next
    SET rsLotePiloto = Conecta("PIR-CORPORATIVO", SQL )
    
    If Not IsObject(rsLotePiloto) OR (rsLotePiloto IS NOTHING) Then
        Exit Function
    End If
    On Error GoTo 0

    SQL = "UPDATE LOTE_PILOTO SET STATUS = '7' " 
    SQL = SQL & "WHERE CLIENTE = '" & codigoCliente & "' "
    SQL = SQL & "AND PLANTA = '" & codigoPlanta & "' "
    SQL = SQL & "AND APROVACAO = 0 "
    SQL = SQL & "AND QUEM_EXCLUIU IS NULL AND STATUS = '1' " 
    SQL = SQL & "AND PRODUTO = '" & codigoProduto & "'"

    On Error Resume Next
    SET rsUpdate = Conecta("PIR-CORPORATIVO", SQL)

    If Not IsObject(rsUpdate) OR (rsUpdate IS NOTHING) Then
        Exit Function
    End If
    On Error GoTo 0
    
    ' Log o status na tabela de historico
    SQL = "INSERT INTO LOTE_PILOTO_HISTORICO_STATUS (IDLOTEPILOTO, STATUS_ANTIGO, STATUS_NOVO, ALTERADO_POR, DATA_INCLUSAO, USUARIO) "
    SQL = SQL & "VALUES ("
    SQL = SQL & "'" & rsLotePiloto("ID") & "', "
    SQL = SQL & "'" & rsLotePiloto("STATUS") & "', "
    SQL = SQL & "'Aguardando Producao', "
    SQL = SQL & "'PEDIDO', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "'" & usuarioLogado("id") & "'"
    SQL = SQL & ")"

    On Error Resume Next
    SET rsHistorico = Conecta("PIR-CORPORATIVO", SQL )
    If Not IsObject(rsHistorico) OR (rsHistorico IS NOTHING) Then
        Exit Function
    End If
    On Error GoTo 0

    rsLotePiloto.Close()
    SET rsLotePiloto = NOTHING
End Function 

Function ListarClientes()
    codigoPlanta = request("codigoPlanta")
    codigoCliente = request("codigoCliente")
    codigoLoja = request("codigoLoja")
    cpfCnpj = request("cpfCnpj")
    busca = Iif(request("busca") = "", "", trim(request("busca")))
    
    If codigoPlanta = "" Then
        ListarClientes = "{""sucesso"": false, ""mensagem"": ""Nenhuma planta informada.""}"
        Exit Function
    End If

    SQL = "SELECT COD_CLIENTE AS CodigoCliente, " 
    SQL = SQL & "LOJA_CLIENTE AS CodigoLoja, " 
    SQL = SQL & "NOME_CLIENTE AS RazaoSocial, " 
    SQL = SQL & "MUNICIPIO_CLIENTE AS Municipio, " 
    SQL = SQL & "ESTADO_CLIENTE AS UF, "
    SQL = SQL & "COD_VENDEDOR AS CodigoVendedor, " 
    SQL = SQL & "CNPJ_CLIENTE AS CpfCnpj, " 
    SQL = SQL & "DATA_MAIOR AS ToleranciaDataMaior, "
    SQL = SQL & "TOL_MAIOR AS ToleranciaMaior, "
    SQL = SQL & "NOME_VENDEDOR AS NomeVendedor, "
    SQL = SQL & "CONDICAO_PAGAMENTO AS CondicaoPagamento, " 
    SQL = SQL & "'J' AS TipoPessoa, " 
    SQL = SQL & "CASE WHEN TIPO_FRETE = 'C' THEN 0 ELSE 1 END AS TipoFrete, " 

    If codigoPlanta = "05" then 
        SQL = SQL & "CASE WHEN PODER3 = 'S' THEN (SELECT dbo.FUNC_SALDO_APARAS(COD_CLIENTE)) ELSE 0 END as SaldoAparas "  
    Else 
        SQL = SQL & "0 as SaldoAparas "  
    End If    

    SQL = SQL & "FROM VW_CLIENTES_PEDIDO  " 
    SQL = SQL & "WHERE STATUS = 'ATIVO'"
    
    If busca <> "" Then
        SQL = SQL & " AND ("
        SQL = SQL & "   FANTASIA_CLIENTE LIKE '%" & Ucase(busca) & "%' "
        SQL = SQL & "   OR NOME_CLIENTE LIKE '%" & Ucase(busca) & "%' "
        SQL = SQL & "   OR COD_CLIENTE LIKE '%" & busca & "%' "
        SQL = SQL & ")"
    End If

    If usuarioLogado("nivel") <> 0 then
        SQL = SQL & " AND s.A1_VEND = '" & usuarioLogado("cod") & "' "
    End If

    SQL = SQL & "ORDER BY FANTASIA_CLIENTE "

    On Error Resume Next
    SET rsClientes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
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
            dadosJSON = dadosJSON & """codigoCliente"": """ & EscapeJSON(rsClientes("CodigoCliente")) & """, "
            dadosJSON = dadosJSON & """codigoLoja"": """ & EscapeJSON(rsClientes("CodigoLoja")) & """, "
            dadosJSON = dadosJSON & """razaoSocial"": """ & EscapeJSON(rsClientes("RazaoSocial")) & """, "
            dadosJSON = dadosJSON & """tipoPessoa"": """ & EscapeJSON(rsClientes("TipoPessoa")) & """, "
            dadosJSON = dadosJSON & """tipoFrete"": """ & EscapeJSON(rsClientes("TipoFrete")) & """, "
            dadosJSON = dadosJSON & """municipio"": """ & EscapeJSON(rsClientes("Municipio")) & """, "
            dadosJSON = dadosJSON & """uf"": """ & EscapeJSON(rsClientes("UF")) & """, "
            dadosJSON = dadosJSON & """toleranciaMaior"": """ & EscapeJSON(rsClientes("ToleranciaMaior")) & """, "
            dadosJSON = dadosJSON & """toleranciaDataMaior"": """ & EscapeJSON(rsClientes("ToleranciaDataMaior")) & """, "
            dadosJSON = dadosJSON & """cpfCnpj"": """ & EscapeJSON(rsClientes("CpfCnpj")) & """, "
            dadosJSON = dadosJSON & """codigoVendedor"": """ & EscapeJSON(rsClientes("CodigoVendedor")) & """, "
            dadosJSON = dadosJSON & """nomeVendedor"": """ & EscapeJSON(rsClientes("NomeVendedor")) & """, "
            dadosJSON = dadosJSON & """condicaoPagamento"": """ & EscapeJSON(rsClientes("CondicaoPagamento")) & """, "
            dadosJSON = dadosJSON & """saldoAparas"": """ & EscapeJSON(rsClientes("SaldoAparas")) & """ "
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

Function ListarClientesRemessa()
    codigoPlanta = request("codigoPlanta")
    codigoCliente = request("codigoCliente")
    codigoLoja = request("codigoLoja")

    SQL = "SELECT A1_NFFAT "
    SQL = SQL & "FROM VW_CLIENTES_PEDIDO "
    SQL = SQL & "WHERE COD_CLIENTE = '" & codigoCliente & "' " 
    SQL = SQL & "AND LOJA_CLIENTE = '" & codigoLoja & "' "
    
    If usuarioLogado("nivel") <> 0 then
        SQL = SQL & " AND COD_VENDEDOR = '" & usuarioLogado("cod") & "' "
    End If

    On Error Resume Next
    SET rsClientes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsClientes) Then
        ListarClientesRemessa = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsClientes.EOF Then
        ListarClientesRemessa = "{""sucesso"": false, ""mensagem"": ""Nenhum Cliente Encontrado""}"
        Exit Function
    End If

    'stringCliente = "73293"
    stringCliente = Trim(rsClientes("A1_NFFAT")) 
    If Right(stringCliente, 1) = "," Then
        stringCliente = Left(stringCliente, Len(stringCliente) - 1)
    End If
    
    SQL = "SELECT  "
    SQL = SQL & "   COD_CLIENTE AS CodigoCliente, "
    SQL = SQL & "   NOME_CLIENTE AS RazaoSocial, "
    SQL = SQL & "   LOJA_CLIENTE AS CodigoLoja, "
    SQL = SQL & "   MUNICIPIO_CLIENTE AS Municipio, " 
    SQL = SQL & "   ESTADO_CLIENTE AS UF, "
    SQL = SQL & "   CNPJ_CLIENTE AS CpfCnpj " 
    SQL = SQL & "  FROM VW_CLIENTES_PEDIDO "

    If codigoPlanta = "07" Then 
        SQL = SQL & "WHERE CONCAT(LTRIM(RTRIM(COD_CLIENTE)), LTRIM(RTRIM(LOJA_CLIENTE))) = ('" & stringCliente & "') "
    Else
        SQL = SQL & "WHERE COD_CLIENTE IN (" & stringCliente & ") "
    End If     

    If usuarioLogado("nivel") <> 0 Then
        SQL = SQL & " AND COD_VENDEDOR = '" & usuarioLogado("cod") & "' "
    End If

    On Error Resume Next
    SET rsClientes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsClientes) Then
        ListarClientesRemessa = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsClientes.EOF Then
        ListarClientesRemessa = "{""sucesso"": false, ""mensagem"": ""Nenhuma TES Encontrada""}"
        Exit Function
    End If
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsClientes.EOF Then
        Do While Not rsClientes.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """codigoCliente"": """ & EscapeJSON(rsClientes("CodigoCliente")) & """, "
            dadosJSON = dadosJSON & """razaoSocial"": """ & EscapeJSON(rsClientes("RazaoSocial")) & """, "
            dadosJSON = dadosJSON & """codigoLoja"": """ & EscapeJSON(rsClientes("CodigoLoja")) & """, "
            dadosJSON = dadosJSON & """municipio"": """ & EscapeJSON(rsClientes("Municipio")) & """, "
            dadosJSON = dadosJSON & """uf"": """ & EscapeJSON(rsClientes("UF")) & """, "
            dadosJSON = dadosJSON & """cpfCnpj"": """ & EscapeJSON(rsClientes("CpfCnpj")) & """ "
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsClientes.MoveNext()
        Loop
    End If

    dadosJSON = dadosJSON & "]"
    rsClientes.Close()
    SET rsClientes = NOTHING

    ListarClientesRemessa = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ListarTes()
    codigoPlanta = request("codigoPlanta")
    codigoCliente = request("codigoCliente")
    codigoLoja = request("codigoLoja")

    nomeTabela = Iif(codigoPlanta = "07", "VW_BI_CLIENTES", "VW_CLIENTES_PEDIDO")

    SQL = "SELECT LTRIM(RTRIM(TES)) AS TES "
    SQL = SQL & "FROM " & nomeTabela & " "
    SQL = SQL & "WHERE COD_CLIENTE = '"  & codigoCliente & "' "
    SQL = SQL & "AND LOJA_CLIENTE = '" & codigoLoja & "' "

    On Error Resume Next
    SET rsTes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTes) Then
        ListarTes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsTes.EOF Then
        ListarTes = "{""sucesso"": false, ""mensagem"": ""Nenhuma TES Encontrada""}"
        Exit Function
    End If

    stringTes = rsTes("TES")
    If Right(stringTes, 1) = "," Then
        stringTes = Left(stringTes, Len(stringTes) - 1)
    End If
    
    SQL = "SELECT "
    SQL = SQL & "   RTRIM(LTRIM(CODIGO)) AS Codigo, "
    SQL = SQL & "   RTRIM(LTRIM(FINALIDADE)) AS Descricao, "
    SQL = SQL & "   ISNULL(TES_REMESSA,'') AS TesRemessa, "
    SQL = SQL & "   PODER3 AS Poder3, "
    SQL = SQL & "   DUPLICATA AS ValidaLimiteCredito "
    SQL = SQL & "FROM VW_BI_TES "
    SQL = SQL & "WHERE CODIGO IN (" & stringTes & ")"

    On Error Resume Next
    SET rsTes = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTes) Then
        ListarTes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If rsTes.EOF Then
        ListarTes = "{""sucesso"": false, ""mensagem"": ""Nenhuma TES Encontrada""}"
        Exit Function
    End If
    
    primeiroItem = True
    dadosJSON = "["

    If Not rsTes.EOF Then
        Do While Not rsTes.EOF
            If Not primeiroItem Then dadosJSON = dadosJSON & "," End If

            dadosJSON = dadosJSON & "{"
            dadosJSON = dadosJSON & """codigo"": """ & EscapeJSON(rsTes("Codigo")) & """, "
            dadosJSON = dadosJSON & """descricao"": """ & EscapeJSON(rsTes("Descricao")) & """, "
            dadosJSON = dadosJSON & """tesRemessa"": """ & EscapeJSON(rsTes("TesRemessa")) & """, "
            dadosJSON = dadosJSON & """poder3"": """ & EscapeJSON(rsTes("Poder3")) & """, "
            dadosJSON = dadosJSON & """validaLimiteCredito"": """ & EscapeJSON(rsTes("ValidaLimiteCredito")) & """ "
            dadosJSON = dadosJSON & "}"
            
            primeiroItem = False
            rsTes.MoveNext()
        Loop
    End If

    dadosJSON = dadosJSON & "]"
    rsTes.Close()
    SET rsTes = NOTHING

    ListarTes = "{""sucesso"": true, ""mensagem"": ""Lista obtida com sucesso"", ""dados"": " & dadosJSON & "}"
End Function

Function ListarTiposItem()
    codigoPlanta = request("codigoPlanta")

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

    ' Bento também tem tipo papel
    If codigoPlanta = "05" Then
        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigo"": """ & "PA" & """, "
        jsonArray = jsonArray & """descricao"": """ & "PAPEL" & """"
        jsonArray = jsonArray & "}"
        jsonArray = jsonArray & ","
    End If

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "CJ" & """, "
    jsonArray = jsonArray & """descricao"": """ & "CONJUNTO" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTiposItem = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarCondicoesPagamento()
    codigoPlanta = Request("codigoPlanta")
    condicaoPagamento = Request("condicaoPagamento")

    SQL = "SELECT "
    SQL = SQL & "   LTRIM(RTRIM(CODIGO)) AS Codigo, "
    SQL = SQL & "   LTRIM(RTRIM(DESCRICAO)) AS Descricao "
    SQL = SQL & "FROM VW_BI_CONDICAO_PAGAMENTO "
    SQL = SQL & "WHERE STATUS = 'ATIVO' "
    
    If condicaoPagamento <> "" Then 
        SQL = SQL & "AND CODIGO = '" & condicaoPagamento & "' "
    End If

    SQL = SQL & "ORDER BY DESCRICAO"

    On Error Resume Next
    SET rsCondicao = Conecta((ObterConexaoProtheus(codigoPlanta)), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsCondicao) Then
        ListarCondicoesPagamento = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de responsáveis: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsCondicao.EOF Then
        Do While Not rsCondicao.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsCondicao("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsCondicao("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsCondicao.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsCondicao.Close()
    SET rsCondicao = NOTHING
    ListarCondicoesPagamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarProdutos()
    codigoPlanta = Request("codigoPlanta")
    codigoCliente = Request("codigoCliente")
    codigoLoja = Request("codigoLoja")
    tipoItem = Request("itemTipo")
    busca = trim(request("busca"))

    If tipoItem = "CJ" then 
        SQL = "SELECT TOP 1 "  
        SQL = SQL & "   CODIGO_CONJUNTO as Codigo, "  
        SQL = SQL & "   REFERENCIA_CONJUNTO as Descricao, "  
        SQL = SQL & "   REFERENCIA_CONJUNTO AS Referencia, "  
        SQL = SQL & "   ESTRUTURA as Estrutura, "  
        SQL = SQL & "   LARGURA AS Largura, "  
        SQL = SQL & "   ALTURA AS Altura, "  
        SQL = SQL & "   COMPRIMENTO AS Comprimento, "  
        SQL = SQL & "   LARGURA_CHAPA as LarguraChapa, "  
        SQL = SQL & "   COMPRIMENTO_CHAPA as ComprimentoChapa, "  
        SQL = SQL & "   AREA as Area, "  
        SQL = SQL & "   AREA_REFILE as AreaRefile, "  
        SQL = SQL & "   CIDADE as Cidade, "  
        SQL = SQL & "   CODIGO_CLIENTE as CodigoCliente, "  
        SQL = SQL & "   0 as Cidade, "  
        SQL = SQL & "   VINCO1 as Vinco1, " 
        SQL = SQL & "   VINCO2 as Vinco2, " 
        SQL = SQL & "   VINCO3 as Vinco3, " 
        SQL = SQL & "   VINCO4 as Vinco4, " 
        SQL = SQL & "   VINCO5 as Vinco5, "  
        SQL = SQL & "   0 as B5_TPPALE as Palete, "  
        SQL = SQL & "   PRECO as ValorUnitario, "  
        SQL = SQL & "   0 as GramaturaBobina, "  
        SQL = SQL & "   0 as LarguraBobina, "  
        SQL = SQL & "   0 as DiametroBobina, "  
        SQL = SQL & "   0 as TubeteBobina, "  
        SQL = SQL & "   GRUPO_PALETIZACAO as GrupoPaletizacao, "   
        SQL = SQL & "   ESTOQUE as Estoque, "   
        SQL = SQL & "   PESO as Peso, "   
        SQL = SQL & "   TIPO_PRODUTO as TipoProduto "   
        SQL = SQL & "FROM VW_BI_PRODUTOS_PEDIDO "  
        SQL = SQL & "WHERE CODIGO_CLIENTE = '" & codigoCliente & "' "  
        SQL = SQL & "   AND STATUS = 'ATIVO' "  
        SQL = SQL & "   AND LOJA_CLIENTE = '" & codigoLoja & "' "  
        SQL = SQL & "   AND (REFERENCIA_CONJUNTO like UPPER ('%" & busca & "%') OR CODIGO_CONJUNTO = '" & busca & "') " 
        SQL = SQL & "   AND TIPO = 'CJ' "  
        SQL = SQL & "ORDER BY CODIGO_ITEM"
    Else
        SQL = "SELECT "
        SQL = SQL & "   CASE WHEN TIPO  = 'FT' THEN CODIGO_CONJUNTO ELSE CODIGO_ITEM END AS Codigo, "
        SQL = SQL & "   REFERENCIA_ITEM as Descricao, "
        SQL = SQL & "   REFERENCIA_ITEM AS Referencia, "
        SQL = SQL & "   ESTRUTURA AS Estrutura, "
        SQL = SQL & "   LARGURA AS Largura, "  
        SQL = SQL & "   ALTURA AS Altura, "  
        SQL = SQL & "   COMPRIMENTO AS Comprimento, " 
        SQL = SQL & "   LARGURA_CHAPA as LarguraChapa, "
        SQL = SQL & "   COMPRIMENTO_CHAPA as ComprimentoChapa, " 
        SQL = SQL & "   AREA AS Area, "
        SQL = SQL & "   CIDADE as Cidade, "
        SQL = SQL & "   CODIGO_CLIENTE as CodigoCliente, "
        SQL = SQL & "   VINCO1 AS Vinco1, "
        SQL = SQL & "   VINCO2 AS Vinco2, "
        SQL = SQL & "   VINCO3 AS Vinco3, "
        SQL = SQL & "   VINCO4 AS Vinco4, "
        SQL = SQL & "   VINCO5 AS Vinco5, "
        SQL = SQL & "   0 AS Palete, "
        SQL = SQL & "   PRECO AS ValorUnitario, "

        If tipoItem = "PA" Then
            SQL = SQL & "0 as AreaRefile, "
            SQL = SQL & "GramaturaBobina, "
            SQL = SQL & "LarguraBobina, "
            SQL = SQL & "DiametroBobina, "
            SQL = SQL & "TubeteBobina, "
        Else
            SQL = SQL & "AREA_REFILE as AreaRefile, "  
            SQL = SQL & "0 as GramaturaBobina, "
            SQL = SQL & "0 as LarguraBobina, "
            SQL = SQL & "0 as DiametroBobina, "
            SQL = SQL & "0 as TubeteBobina, "
        End If

        SQL = SQL & "   GRUPO_PALETIZACAO as GrupoPaletizacao, "
        SQL = SQL & "   ESTOQUE AS Estoque, "
        SQL = SQL & "   PESO AS Peso, "
        SQL = SQL & "   TIPO_PRODUTO AS TipoProduto "
        SQL = SQL & "FROM " & IIf(tipoItem = "PA", "VW_PAPEL_PEDIDO", "VW_BI_PRODUTOS_PEDIDO") & " "
        SQL = SQL & "WHERE 1 = 1 " 

        If tipoItem = "CH" AND (codigoPlanta = "05" OR codigoPlanta = "07") Then
            SQL = SQL & "AND GRUPO_PRODUTO = 'CHAPA' "
            
            If codigoPlanta = "07" Then
                SQL = SQL & "AND TIPO_CHAPA = 'GENERICA'"
            End If
        Else
            If tipoItem <> "PA" Then
                SQL = SQL & "AND CODIGO_CLIENTE = '" & codigoCliente & "' "
                SQL = SQL & "AND LOJA_CLIENTE = '" & codigoLoja & "' "
            End If
        End If     

        SQL = SQL & "AND STATUS = 'ATIVO' " 
        SQL = SQL & "AND (REFERENCIA_ITEM LIKE UPPER ('%" & busca & "%') OR CASE WHEN TIPO = 'FT' THEN CODIGO_CONJUNTO ELSE CODIGO_ITEM END LIKE '%" & busca & "%') " 
    End If 

    stringConexao = ""
    Select Case codigoPlanta
        Case "05" 
            If tipoItem = "PA" then
                stringConexao = "BEN-DADOSADV"
            Else
                stringConexao = "BEN-TRIMBOX"
            End If
        Case "06" 
            stringConexao = "FOR-TRIMBOX" 
        Case "07" 
            stringConexao = "SAQ-TRIMBOX"
    End Select

    On Error Resume Next
    SET rsProdutos = Conecta(stringConexao, SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsProdutos) Then
        ListarProdutos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de responsáveis: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsProdutos.EOF Then
        Do While Not rsProdutos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsProdutos("Codigo")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsProdutos("Descricao")) & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsProdutos("Referencia")) & """, "
            jsonArray = jsonArray & """estrutura"": """ & EscapeJSON(rsProdutos("Estrutura")) & """, "
            jsonArray = jsonArray & """largura"": """ & Replace(EscapeJSON(rsProdutos("Largura")), ",", ".") & """, "
            jsonArray = jsonArray & """altura"": """ & Replace(EscapeJSON(rsProdutos("Altura")), ",", ".") & """, "
            jsonArray = jsonArray & """comprimento"": """ & Replace(EscapeJSON(rsProdutos("Comprimento")), ",", ".") & """, "
            jsonArray = jsonArray & """larguraChapa"": """ & Replace(EscapeJSON(rsProdutos("LarguraChapa")), ",", ".") & """, "
            jsonArray = jsonArray & """comprimentoChapa"": """ & Replace(EscapeJSON(rsProdutos("ComprimentoChapa")), ",", ".") & """, "
            jsonArray = jsonArray & """area"": """ & Replace(EscapeJSON(rsProdutos("Area")), ",", ".") & """, "
            jsonArray = jsonArray & """areaRefile"": """ & Replace(EscapeJSON(rsProdutos("AreaRefile")), ",", ".") & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & Replace(EscapeJSON(rsProdutos("CodigoCliente")), ",", ".") & """, "
            jsonArray = jsonArray & """vinco1"": """ & Replace(EscapeJSON(rsProdutos("Vinco1")), ",", ".") & """, "
            jsonArray = jsonArray & """vinco2"": """ & Replace(EscapeJSON(rsProdutos("Vinco2")), ",", ".") & """, "
            jsonArray = jsonArray & """vinco3"": """ & Replace(EscapeJSON(rsProdutos("Vinco3")), ",", ".") & """, "
            jsonArray = jsonArray & """vinco4"": """ & Replace(EscapeJSON(rsProdutos("Vinco4")), ",", ".") & """, "
            jsonArray = jsonArray & """vinco5"": """ & Replace(EscapeJSON(rsProdutos("Vinco5")), ",", ".") & """, "
            jsonArray = jsonArray & """palete"": """ & Replace(EscapeJSON(rsProdutos("Palete")), ",", ".") & """, "
            jsonArray = jsonArray & """valorUnitario"": """ & Replace(EscapeJSON(rsProdutos("ValorUnitario")), ",", ".") & """, "
            jsonArray = jsonArray & """gramaturaBobina"": """ & Replace(EscapeJSON(rsProdutos("GramaturaBobina")), ",", ".") & """, "
            jsonArray = jsonArray & """tubeteBobina"": """ & Replace(EscapeJSON(rsProdutos("TubeteBobina")), ",", ".") & """, "
            jsonArray = jsonArray & """diametroBobina"": """ & Replace(EscapeJSON(rsProdutos("DiametroBobina")), ",", ".") & """, "
            jsonArray = jsonArray & """larguraBobina"": """ & Replace(EscapeJSON(rsProdutos("LarguraBobina")), ",", ".") & """, "
            jsonArray = jsonArray & """grupoPaletizacao"": """ & Replace(EscapeJSON(rsProdutos("GrupoPaletizacao")), ",", ".") & """, "
            jsonArray = jsonArray & """estoque"": """ & Replace(EscapeJSON(rsProdutos("Estoque")), ",", ".") & """, "
            jsonArray = jsonArray & """peso"": """ & Replace(EscapeJSON(rsProdutos("Peso")), ",", ".") & """, "
            jsonArray = jsonArray & """tipoProduto"": """ & EscapeJSON(rsProdutos("TipoProduto")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsProdutos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsProdutos.Close()
    SET rsProdutos = NOTHING
    
    ListarProdutos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarGramaturas()
    codigoPlanta = Request("codigoPlanta")

    SQL = "Select DISTINCT Gramatura from dbo.E001_GN060E82 Where Gramatura > 0 order by Gramatura"

    On Error Resume Next
    SET rsGramatura = Conecta("BEN-SOLPEN", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsGramatura) Then
        ListarGramaturas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de responsáveis: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsGramatura.EOF Then
        Do While Not rsGramatura.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsGramatura("Gramatura")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsGramatura("Gramatura")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsGramatura.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsGramatura.Close()
    SET rsGramatura = NOTHING
    
    ListarGramaturas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTubetes()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "3" & """, "
    jsonArray = jsonArray & """descricao"": """ & "3" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "4" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarTubetes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarProdutoRoteiro()
    codigoProduto = Trim(Request("codigoProduto"))
    codigoCliente = Trim(Request("codigoCliente"))
    codigoLoja = Trim(Request("codigoLoja"))
    tipoProduto = Trim(Request("tipoProduto"))
    codigoPlanta = Trim(Request("codigoPlanta"))

    If tipoProduto = "CJ" Then
        clausulaWhere = "WHERE P.CODIGO_CONJUNTO = '" & codigoProduto & "'"
    Else
        clausulaWhere = "WHERE P.CODIGO_ITEM = '" & codigoProduto & "'"
    End If

    ' Adiciona condições específicas por codigoPlanta
    If codigoPlanta = "07" Then
        ' codigoPlanta 07: adiciona filtro por código cliente e loja
        If codigoCliente <> "" Then
            clausulaWhere = clausulaWhere & " AND P.CODIGO_CLIENTE = '" & codigoCliente & "'"
        End If

        If codigoLoja <> "" Then
            clausulaWhere = clausulaWhere & " AND P.LOJA_CLIENTE = '" & codigoLoja & "'"
        End If
    ElseIf codigoPlanta = "05" Or codigoPlanta = "06" Then
        ' Plantas 05 e 06: adiciona filtro apenas por código cliente
        If codigoCliente <> "" Then
            clausulaWhere = clausulaWhere & " AND P.CODIGO_CLIENTE = '" & codigoCliente & "'"
        End If
    End If

    SQL = "SELECT " 
    SQL = SQL & "   P.TIPO AS Tipo, " 
    SQL = SQL & "   P.CIDADE AS Cidade, " 
    SQL = SQL & "   P.CODIGO_CONJUNTO AS CodigoConjunto, " 
    SQL = SQL & "   P.CODIGO_ITEM AS CodigoItem, " 
    SQL = SQL & "   P.REFERENCIA_CONJUNTO AS ReferenciaConjunto, " 
    SQL = SQL & "   P.REFERENCIA_ITEM AS ReferenciaItem, " 
    SQL = SQL & "   R.MAQUINA_1 AS Maquina1, " 
    SQL = SQL & "   R.NOME_MAQUINA_1 AS NomeMaquina1, " 
    SQL = SQL & "   R.STATUS_MAQUINA1 AS StatusMaquina1, " 
    SQL = SQL & "   R.MAQUINA_2 AS Maquina2,  " 
    SQL = SQL & "   R.NOME_MAQUINA_2 AS NomeMaquina2, " 
    SQL = SQL & "   R.STATUS_MAQUINA2 AS StatusMaquina2, " 
    SQL = SQL & "   R.MAQUINA_3 AS Maquina3, " 
    SQL = SQL & "   R.NOME_MAQUINA_3 AS NomeMaquina3, " 
    SQL = SQL & "   R.STATUS_MAQUINA3 AS StatusMaquina3, " 
    SQL = SQL & "   R.MAQUINA_4 AS Maquina4, " 
    SQL = SQL & "   R.NOME_MAQUINA_4 AS NomeMaquina4, " 
    SQL = SQL & "   R.STATUS_MAQUINA4 AS StatusMaquina4, " 
    SQL = SQL & "   R.MAQUINA_5 AS Maquina5, " 
    SQL = SQL & "   R.NOME_MAQUINA_5 AS NomeMaquina5, " 
    SQL = SQL & "   R.STATUS_MAQUINA5 AS StatusMaquina5, " 
    SQL = SQL & "   R.MAQUINA_6 AS Maquina6, " 
    SQL = SQL & "   R.NOME_MAQUINA_6 AS NomeMaquina6, " 
    SQL = SQL & "   R.STATUS_MAQUINA6 AS StatusMaquina6 " 
    SQL = SQL & "FROM VW_BI_PRODUTOS_PEDIDO P " 
    SQL = SQL & "LEFT JOIN VW_BI_ROTEIRO R ON R.PRODUTO = P.CODIGO_ITEM " 
    SQL = SQL & clausulaWhere

    On Error Resume Next
    SET rsRoteiro = Conecta(ObterConexaoTrimbox(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRoteiro) Then
        ListarGramaturas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de responsáveis: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroRegistro = True
    jsonArray = "["

    If Not rsRoteiro.EOF Then
        Do While Not rsRoteiro.EOF
            If Not primeiroRegistro Then
                jsonArray = jsonArray & ","
            End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """tipo"": """ & EscapeJSON(Trim(rsRoteiro("Tipo") & "")) & ""","
            jsonArray = jsonArray & """cidade"": """ & EscapeJSON(Trim(rsRoteiro("Cidade") & "")) & ""","
            jsonArray = jsonArray & """codigoConjunto"": """ & EscapeJSON(Trim(rsRoteiro("CodigoConjunto") & "")) & ""","
            jsonArray = jsonArray & """codigoItem"": """ & EscapeJSON(Trim(rsRoteiro("CodigoItem") & "")) & ""","
            jsonArray = jsonArray & """referenciaConjunto"": """ & EscapeJSON(Trim(rsRoteiro("ReferenciaConjunto") & "")) & ""","
            jsonArray = jsonArray & """referenciaItem"": """ & EscapeJSON(Trim(rsRoteiro("ReferenciaItem") & "")) & ""","
            jsonArray = jsonArray & """maquinas"": ["

            flagPrimeiraMaquina = True

            For i = 1 To 6
                codigoMaquina =  Trim(rsRoteiro("maquina" & i) & "")
                nomeMaquina = Trim(rsRoteiro("nomeMaquina" & i) & "")
                statusMaquina = Trim(rsRoteiro("statusMaquina" & i) & "")

                If codigoMaquina <> "" And Not IsNull(rsRoteiro("maquina" & i)) Then
                    If Not flagPrimeiraMaquina Then
                        jsonArray = jsonArray & ","
                    End If

                    jsonArray = jsonArray & "{"
                    jsonArray = jsonArray & """codigo"": """ & EscapeJSON(codigoMaquina) & ""","
                    jsonArray = jsonArray & """nome"": """ & EscapeJSON(nomeMaquina) & """, "
                    jsonArray = jsonArray & """status"": """ & EscapeJSON(statusMaquina) & """ "
                    jsonArray = jsonArray & "}"

                    flagPrimeiraMaquina = False
                End If
            Next

            jsonArray = jsonArray & "]"
            jsonArray = jsonArray & "}"

            primeiroRegistro = False
            rsRoteiro.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsRoteiro.Close
    Set rsRoteiro = Nothing

    ListarProdutoRoteiro = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterValorUnitarioProdutoOrcamento()
    codigoPlanta = Request("codigoPlanta")
    codigoCliente = Request("codigoCliente")
    codigoLoja = Request("codigoLoja")
    codigoProduto = Request("codigoProduto")
    tipoFrete = Request("tipoFrete")
    condicaoPagamento = Request("condicaoPagamento")

    valor = "0"

    SQL = "SELECT dbo.FUNC_VALIDA_ORCAMENTO_RETORNA_PRECO(" 
    SQL = SQL & "'" & codigoPlanta & "', "
    SQL = SQL & "'" & codigoCliente & "', "
    SQL = SQL & "'" & codigoLoja & "', "
    SQL = SQL & "'" & codigoProduto & "', "
    SQL = SQL & "'" & tipoFrete & "', "
    SQL = SQL & "'" & condicaoPagamento & "', "
    SQL = SQL & "CONVERT(VARCHAR(8), GETDATE(), 112)"
    SQL = SQL & ") AS PrecoOrcamento "

    On Error Resume Next
    SET rsValor = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsValor) Then
        ObterValorUnitarioProdutoOrcamento = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar valor unitario orcamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If not rsValor.eof then
        valor = rsValor("PrecoOrcamento") 
    End If 
  
    ObterValorUnitarioProdutoOrcamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""valor"": " & Replace(valor, ",", ".")  & "}"
End Function

Function VerificarExisteOrcamento()
    codigoPlanta = Request("codigoPlanta")
    codigoCliente = Request("codigoCliente")
    codigoLoja = Request("codigoLoja")
    codigoProduto = Request("codigoProduto")
    tipoFrete = Request("tipoFrete")
    condicaoPagamento = Request("condicaoPagamento")
    valorUnitario = Request("valorUnitario")

    flagExiste = "0"

    SQL = "SELECT dbo.FUNC_VALIDA_ORCAMENTO(" 
    SQL = SQL & "'" & codigoPlanta & "', "
    SQL = SQL & "'" & codigoCliente & "', "
    SQL = SQL & "'" & codigoLoja & "', "
    SQL = SQL & "'" & codigoProduto & "', "
    SQL = SQL & "'" & tipoFrete & "', "
    SQL = SQL & "'" & condicaoPagamento & "', "
    SQL = SQL & "CONVERT(VARCHAR(8), GETDATE(), 112), "
    SQL = SQL & "'" & Replace(valorUnitario, ",", ".") & "'"
    SQL = SQL & ") AS Existe "

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsExiste) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar se orçamento para item: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If not rsExiste.eof then
        flagExiste = rsExiste("Existe") 
    End If 
  
    VerificarExisteOrcamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""existe"": " & flagExiste & "}"
End Function

' Verifica que já existe pedido com mesma ordem de compra
' Aviso para evitar emissão de pedido duplicado
Function VerificarPedidoCompraDuplicado()
    codigoPlanta = Request("codigoPlanta")
    codigoCliente = Request("codigoCliente")
    codigoProduto = Request("codigoProduto")
    pedidoCompra = Request("pedidoCompra")
    itemPedido = Request("itemPedido")

    flagExiste = "0"

    If pedidoCompra = "" OR itemPedido = "" Then
        VerificarPedidoCompraDuplicado = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""existe"": " & flagExiste & "}"
        Exit Function
    End If

    SQL = "SELECT COUNT(Id) as Existe FROM Pedido "
    SQL = SQL & "WHERE Cliente = '" & codigoCliente & "' "
    SQL = SQL & "AND NumeroOrdemCompra = '" & pedidoCompra & "' "
    SQL = SQL & "AND ItemOrdemCompra = '" & itemPedido & "' " 
    SQL = SQL & "AND Excluido = 0"

    On Error Resume Next
    SET rsExisteCorporativo = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)

    If Err.Number <> 0 Or Not IsObject(rsExisteCorporativo) Then
        VerificarPedidoCompraDuplicado = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar se existe pedido duplicado para item: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsExisteCorporativo.EOF then
        If rsExisteCorporativo("Existe") > 0 Then
            flagExiste = "1"
        End If
    End If 

    ' Se não encontrou também busca no trimbox
    If flagExiste = "0" Then
        SQL = "SELECT COUNT(IdPedidos) as Existe "
        SQL = SQL & "FROM Pedidos WHERE IDCliente = '" & codigoCliente & "' "
        SQL = SQL & "AND PedCliente = '" & pedidoCompra & "' "
        SQL = SQL & "AND ItemPedCliente = '" & itemPedido & "' "
        SQL = SQL & "AND StatusPedido <> '4'"

        On Error Resume Next
        SET rsExisteTrimbox = Conecta(ObterConexaoTrimbox(codigoPlanta), SQL)

        If Err.Number <> 0 Or Not IsObject(rsExisteTrimbox) Then
            VerificarPedidoCompraDuplicado = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar se existe pedido duplicado para item: " & EscapeJSON(Err.Description) & """}"
            Exit Function
        End If
        On Error Goto 0

        If Not rsExisteTrimbox.EOF then
            If rsExisteTrimbox("Existe") > 0 Then
                flagExiste = "1"
            End If
        End If 
    End If
  
    VerificarPedidoCompraDuplicado = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""existe"": " & flagExiste & "}"
End Function

Function VerificarExisteMaquinaInativa()
    codigoProduto = Trim(Request("codigoProduto"))
    codigoCliente = Trim(Request("codigoCliente"))
    codigoLoja = Trim(Request("codigoLoja"))
    tipoProduto = Trim(Request("tipoProduto"))
    codigoPlanta = Trim(Request("codigoPlanta"))

    flagExiste = "0"

    If tipoProduto = "CJ" Then
        clausulaWhere = "WHERE P.CODIGO_CONJUNTO = '" & codigoProduto & "'"
    Else
        clausulaWhere = "WHERE P.CODIGO_ITEM = '" & codigoProduto & "'"
    End If

    ' Adiciona condições específicas por codigoPlanta
    If codigoPlanta = "07" Then
        ' codigoPlanta 07: adiciona filtro por código cliente e loja
        If codigoCliente <> "" Then
            clausulaWhere = clausulaWhere & " AND P.CODIGO_CLIENTE = '" & codigoCliente & "'"
        End If

        If codigoLoja <> "" Then
            clausulaWhere = clausulaWhere & " AND P.LOJA_CLIENTE = '" & codigoLoja & "'"
        End If
    ElseIf codigoPlanta = "05" Or codigoPlanta = "06" Then
        ' Plantas 05 e 06: adiciona filtro apenas por código cliente
        If codigoCliente <> "" Then
            clausulaWhere = clausulaWhere & " AND P.CODIGO_CLIENTE = '" & codigoCliente & "'"
        End If
    End If
    
    SQL = SQL & "SELECT"
    SQL = SQL & "    CASE "
    SQL = SQL & "        WHEN ( "
    SQL = SQL & "            (R.NOME_MAQUINA_1 IS NOT NULL AND R.STATUS_MAQUINA1 = 'INATIVA') "
    SQL = SQL & "            OR (R.NOME_MAQUINA_2 IS NOT NULL AND R.STATUS_MAQUINA2 = 'INATIVA') "
    SQL = SQL & "            OR (R.NOME_MAQUINA_3 IS NOT NULL AND R.STATUS_MAQUINA3 = 'INATIVA') "
    SQL = SQL & "            OR (R.NOME_MAQUINA_4 IS NOT NULL AND R.STATUS_MAQUINA4 = 'INATIVA') "
    SQL = SQL & "            OR (R.NOME_MAQUINA_5 IS NOT NULL AND R.STATUS_MAQUINA5 = 'INATIVA') "
    SQL = SQL & "            OR (R.NOME_MAQUINA_6 IS NOT NULL AND R.STATUS_MAQUINA6 = 'INATIVA') "
    SQL = SQL & "        ) "
    SQL = SQL & "        THEN 'SIM' "
    SQL = SQL & "        ELSE 'NAO' "
    SQL = SQL & "    END AS PossuiMaquinaInativa "
    SQL = SQL & "FROM VW_BI_PRODUTOS_PEDIDO AS P "
    SQL = SQL & "LEFT JOIN VW_BI_ROTEIRO AS R ON R.PRODUTO = P.CODIGO_ITEM "

    On Error Resume Next
    SET rsExiste = Conecta(ObterConexaoTrimbox(codigoPlanta), SQL)

    If Err.Number <> 0 Or Not IsObject(rsExiste) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar se orçamento para item: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.eof then
        flagExiste = Iif(rsExiste("PossuiMaquinaInativa") = "SIM", "1", "0")
    End If 
  
    VerificarExisteMaquinaInativa = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""existe"": " & flagExiste & "}"
End Function

Function VerificarExisteLotePiloto()
    codigoPlanta = Request("codigoPlanta")
    codigoCliente = Request("codigoCliente")
    codigoLoja = Request("codigoLoja")
    codigoProduto = Request("codigoProduto")

    flagExiste = "0"

    SQL = SQL & "SELECT COUNT(*) as Total FROM LOTE_PILOTO "
    SQL = SQL & "WHERE PLANTA = '" & codigoPlanta & "' "
    SQL = SQL & "AND CLIENTE = '" & codigoCliente & "' "
    SQL = SQL & "AND LOJA = '" & codigoLoja & "' "
    SQL = SQL & "AND PRODUTO = '" & codigoProduto & "' "
    SQL = SQL & "AND APROVACAO = 0"

    On Error Resume Next
    SET rsExiste = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Or Not IsObject(rsExiste) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao verificar se orçamento para item: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    If Not rsExiste.eof then
        flagExiste = Iif(rsExiste("Total") > 0, "1", "0")
    End If 
  
    VerificarExisteLotePiloto = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""existe"": " & flagExiste & "}"
End Function

' ========= Auxiliares ==========

Function ObterLimiteCredito(codigoCliente, codigoLoja, codigoPlanta)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrão de retorno
    resultado.Add "sucesso", False
    resultado.Add "mensagem", ""
    resultado.Add "saldo", 0
    resultado.Add "planta", codigoPlanta
    resultado.Add "cliente", codigoCliente
    resultado.Add "loja", codigoLoja

    ' Clientes que não devem ter restrição de saldo
    If (codigoCliente = "76728" Or codigoCliente = "50894") Then
        saldoEspecial = 999999999999

        resultado("saldo") = saldoEspecial
        resultado("sucesso") = True
        resultado("mensagem") = "Cliente com limite de crédito especial."

        Set ObterLimiteCredito = resultado
        Exit Function
    End If

    If codigoPlanta = "07" Then
        SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO(" & _
              "'" & codigoCliente & "', " & _
              "'" & codigoLoja & "') AS Saldo"
    Else
        SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO(" & _
              "'" & codigoCliente & "') AS Saldo"
    End If

    On Error Resume Next

    Err.Clear
    Set rsSaldo = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If Err.Number <> 0 Or Not IsObject(rsSaldo) Then
        resultado("mensagem") = _
            "Erro ao consultar o limite de crédito: " & Err.Description

        Err.Clear
        On Error GoTo 0

        Set ObterLimiteCredito = resultado
        Exit Function
    End If

    On Error GoTo 0

    If Not rsSaldo.EOF Then
        resultado("saldo") = rsSaldo("Saldo")
        resultado("sucesso") = True
        resultado("mensagem") = "Limite de crédito obtido com sucesso."
    Else
        resultado("mensagem") = _
            "Nenhuma informação de limite de crédito foi encontrada."
    End If

    rsSaldo.Close
    Set rsSaldo = Nothing

    Set ObterLimiteCredito = resultado
End Function

Function ObterSaldoCliente(codigoCliente, codigoLoja, codigoPlanta)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrão de retorno
    resultado.Add "sucesso", False
    resultado.Add "mensagem", ""
    resultado.Add "aberto", 0
    resultado.Add "limite", 0
    resultado.Add "pedidos", 0
    resultado.Add "pedidosLiberar", 0
    resultado.Add "saldo", 0
    resultado.Add "planta", codigoPlanta
    resultado.Add "cliente", codigoCliente
    resultado.Add "loja", codigoLoja

    ' Clientes que não devem ter restrição de saldo
    If (codigoCliente = "76728" Or codigoCliente = "50894") Then
        saldoEspecial = 999999999999.99

        resultado("limite") = saldoEspecial
        resultado("saldo") = saldoEspecial
        resultado("sucesso") = True
        resultado("mensagem") = "Cliente com limite de crédito especial."

        Set ObterSaldoCliente = resultado
        Exit Function
    End If

    If codigoPlanta = "07" Then
        SQL = "SELECT "
        SQL = SQL & "   Aberto, "
        SQL = SQL & "   Limite, "
        SQL = SQL & "   Pedidos, "
        SQL = SQL & "   PedidoLiberar, "
        SQL = SQL & "   Saldo "
        SQL = SQL & "FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA(" & _
                    "'" & codigoCliente & "', " & _
                    "'" & codigoLoja & "')"
    Else
        SQL = "SELECT "
        SQL = SQL & "   Aberto, "
        SQL = SQL & "   Limite, "
        SQL = SQL & "   Pedidos, "
        SQL = SQL & "   PedidoLiberar, "
        SQL = SQL & "   Saldo "
        SQL = SQL & "FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA(" & _
                    "'" & codigoCliente & "')"
    End If

    On Error Resume Next
    Err.Clear
    Set rsSaldo = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If Err.Number <> 0 Or Not IsObject(rsSaldo) Then
        resultado("mensagem") = _
            "Erro ao consultar o limite de crédito: " & Err.Description

        Err.Clear
        On Error GoTo 0

        Set ObterSaldoCliente = resultado
        Exit Function
    End If

    On Error GoTo 0

    If Not rsSaldo.EOF Then
        resultado("aberto") = rsSaldo("Aberto")
        resultado("limite") = rsSaldo("Limite")
        resultado("pedidos") = rsSaldo("Pedidos")
        resultado("pedidosLiberar") = rsSaldo("PedidoLiberar")
        resultado("saldo") = rsSaldo("Saldo")
        resultado("sucesso") = True
        resultado("mensagem") = "Limite de crédito obtido com sucesso."
    Else
        resultado("mensagem") = _
            "Nenhuma informação de limite de crédito foi encontrada."
    End If

    rsSaldo.Close
    Set rsSaldo = Nothing

    Set ObterSaldoCliente = resultado
End Function

Function ObterSaldoClienteRequest()
    codigoCliente = Request("codigoCliente") 
    codigoLoja = Request("codigoLoja")
    codigoPlanta = Request("codigoPlanta")

    If codigoPlanta = "07" Then
        SQL = "SELECT "
        SQL = SQL & "   Aberto, "
        SQL = SQL & "   Limite, "
        SQL = SQL & "   Pedidos, "
        SQL = SQL & "   PedidoLiberar, "
        SQL = SQL & "   Saldo "
        SQL = SQL & "FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA(" & _
                    "'" & codigoCliente & "', " & _
                    "'" & codigoLoja & "')"
    Else
        SQL = "SELECT "
        SQL = SQL & "   Aberto, "
        SQL = SQL & "   Limite, "
        SQL = SQL & "   Pedidos, "
        SQL = SQL & "   PedidoLiberar, "
        SQL = SQL & "   Saldo "
        SQL = SQL & "FROM dbo.FUNC_CALCULA_LIMITE_CREDITO_TABELA(" & _
                    "'" & codigoCliente & "')"
    End If

    On Error Resume Next
    Set rsSaldo = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If Err.Number <> 0 Or Not IsObject(rsSaldo) Then
        ObterSaldoClienteRequest = "{""sucesso"": false, ""mensagem"": ""Erro ao obter limite de crédito""}"
        Set ObterSaldoClienteRequest = resultado
        Exit Function
    End If
    On Error GoTo 0

    jsonObjeto = "{"

    If Not rsSaldo.EOF Then
        jsonObjeto = jsonObjeto & """aberto"": """ & EscapeJSON(Replace(rsSaldo("Aberto"), ",", ".")) & """, "
        jsonObjeto = jsonObjeto & """limite"": """ & EscapeJSON(Replace(rsSaldo("Limite"), ",", ".")) & """, "
        jsonObjeto = jsonObjeto & """pedidos"": """ & EscapeJSON(Replace(rsSaldo("Pedidos"), ",", ".")) & """, "
        jsonObjeto = jsonObjeto & """pedidosLiberar"": """ & EscapeJSON(Replace(rsSaldo("PedidoLiberar"), ",", ".")) & """, "
        jsonObjeto = jsonObjeto & """saldo"": """ & EscapeJSON(Replace(rsSaldo("Saldo"), ",", ".")) & """ "
    End If

    jsonObjeto = jsonObjeto & "}"

    rsSaldo.Close
    Set rsSaldo = Nothing

    ObterSaldoClienteRequest = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function ObterLimiteCredito(codigoCliente, codigoLoja, codigoPlanta)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrão de retorno
    resultado.Add "sucesso", False
    resultado.Add "mensagem", ""
    resultado.Add "saldo", 0
    resultado.Add "planta", codigoPlanta
    resultado.Add "cliente", codigoCliente
    resultado.Add "loja", codigoLoja

    If codigoPlanta = "07" Then
        SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO('" & codigoCliente & "', '" & codigoLoja & "') AS Saldo"
    Else
        SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO('" & codigoCliente & "') AS Saldo"
    End If

    On Error Resume Next
    Set rsSaldo = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If Err.Number <> 0 OR Not IsObject(rsSaldo) Then
        resultado("mensagem") = "Erro ao consultar o limite de crédito: " & Err.Description
        Set ObterLimiteCredito = resultado
        Exit Function
    End If
    On Error GoTo 0

    If NOT rsSaldo.EOF Then
        resultado("saldo") = rsSaldo("Saldo")
        resultado("sucesso") = True
        resultado("mensagem") = "Limite de crédito obtido com sucesso."      
    End If

    rsSaldo.Close()
    SET rsSaldo = NOTHING

    Set ObterLimiteCredito = resultado
End Function

Function ObterSaldoAparasCliente(codigoCliente, codigoPlanta)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrão de retorno
    resultado.Add "sucesso", False
    resultado.Add "mensagem", ""
    resultado.Add "saldo", 0
    resultado.Add "planta", codigoPlanta
    resultado.Add "cliente", codigoCliente

    SQL = "SELECT dbo.FUNC_SALDO_APARAS('" & codigoCliente & "') AS Saldo" 
    
    On Error Resume Next
    SET rsSaldo = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSaldo) Then
        resultado("mensagem") = "Erro ao obter saldo de aparas: " & Err.Description
        Set ObterSaldoAparasCliente = resultado
        Exit Function
    End If
    On Error Goto 0

    If Not rsSaldo.EOF Then
        resultado("saldo") = rsSaldo("Saldo")
        resultado("sucesso") = True
        resultado("mensagem") = "Saldo de Aparas obtido com sucesso."      
    End If

    rsSaldo.Close()
    SET rsSaldo = NOTHING

    Set ObterSaldoAparasCliente = resultado
End Function

Function ValidarLotePiloto(codigoCliente, codigoLoja, codigoProduto, codigoPlanta, quantidade)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrao de retorno
    resultado.Add "sucesso", False
    resultado.Add "valido", False
    resultado.Add "mensagem", ""
    resultado.Add "lotePilotoEncontrado", False
    resultado.Add "atualizaStatus", 0
    resultado.Add "totalPalete", 0
    resultado.Add "observacoes", ""

    quantidadeNumerica = CDbl(quantidade)

    SQL = "SELECT "
    SQL = SQL & "    LOTE_PILOTO_ENCONTRADO, "
    SQL = SQL & "    VALIDA_PEDIDO, "
    SQL = SQL & "    VALIDA_PALETE, "
    SQL = SQL & "    VALIDA_ONDULADEIRA, "
    SQL = SQL & "    VALIDA_IMPRESSORA, "
    SQL = SQL & "    VALIDA_EXPEDICAO, "
    SQL = SQL & "    VALIDA_CLIENTE, "
    SQL = SQL & "    VALIDA_DDP, "
    SQL = SQL & "    VALIDA_CLICHERIA, "
    SQL = SQL & "    TOTAL_PALETE, "
    SQL = SQL & "    MENSAGEM, "
    SQL = SQL & "    MENSAGEM_PEDIDO, "
    SQL = SQL & "    MENSAGEM_PALETE, "
    SQL = SQL & "    ATUALIZA_STATUS "
    SQL = SQL & "FROM dbo.FUNC_VALIDA_LOTE_PILOTO("
    SQL = SQL & "'" & codigoCliente & "', "
    SQL = SQL & "'" & codigoLoja & "', "
    SQL = SQL & "'" & codigoProduto & "', "
    SQL = SQL & "'" & codigoPlanta & "', "
    SQL = SQL & Replace(CStr(quantidadeNumerica), ",", ".")
    SQL = SQL & ")"
    
    On Error Resume Next
    Set rsLotePiloto = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Then
        resultado("mensagem") = "Erro ao consultar o lote piloto: " & Err.Description

        Err.Clear
        On Error GoTo 0

        Set ValidarLotePiloto = resultado
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsLotePiloto) Then
        resultado("mensagem") = "Não foi possível consultar as informações do lote piloto."
        Set ValidarLotePiloto = resultado
        Exit Function
    End If

    If rsLotePiloto.EOF Then
        resultado("sucesso") = True
        resultado("valido") = False
        resultado("mensagem") = "Nenhuma configuração de lote piloto foi encontrada."

        rsLotePiloto.Close
        Set rsLotePiloto = Nothing

        Set ValidarLotePiloto = resultado
        Exit Function
    End If

    ' Inicializa as variaveis booleanas
    lotePilotoEncontrado = False
    validaPedido = False
    validaPalete = False
    validaOnduladeira = False
    validaImpressora = False
    validaExpedicao = False
    validaCliente = False
    validaDDP = False
    validaClicheria = False

    ' Converte os campos retornados pelo banco
    If Not IsNull(rsLotePiloto("LOTE_PILOTO_ENCONTRADO")) Then
        lotePilotoEncontrado = CBool(rsLotePiloto("LOTE_PILOTO_ENCONTRADO"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_PEDIDO")) Then
        validaPedido = CBool(rsLotePiloto("VALIDA_PEDIDO"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_PALETE")) Then
        validaPalete = CBool(rsLotePiloto("VALIDA_PALETE"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_ONDULADEIRA")) Then
        validaOnduladeira = CBool(rsLotePiloto("VALIDA_ONDULADEIRA"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_IMPRESSORA")) Then
        validaImpressora = CBool(rsLotePiloto("VALIDA_IMPRESSORA"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_EXPEDICAO")) Then
        validaExpedicao = CBool(rsLotePiloto("VALIDA_EXPEDICAO"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_CLIENTE")) Then
        validaCliente = CBool(rsLotePiloto("VALIDA_CLIENTE"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_DDP")) Then
        validaDDP = CBool(rsLotePiloto("VALIDA_DDP"))
    End If

    If Not IsNull(rsLotePiloto("VALIDA_CLICHERIA")) Then
        validaClicheria = CBool(rsLotePiloto("VALIDA_CLICHERIA"))
    End If

    ' A consulta foi executada corretamente
    resultado("sucesso") = True
    resultado("lotePilotoEncontrado") = lotePilotoEncontrado

    ' Total de paletes
    If Not IsNull(rsLotePiloto("TOTAL_PALETE")) Then
        If IsNumeric(rsLotePiloto("TOTAL_PALETE")) Then
            resultado("totalPalete") = CDbl(rsLotePiloto("TOTAL_PALETE"))
        End If
    End If

    ' Status que devera ser atualizado posteriormente
    If Not IsNull(rsLotePiloto("ATUALIZA_STATUS")) Then
        If IsNumeric(rsLotePiloto("ATUALIZA_STATUS")) Then
            resultado("atualizaStatus") = CLng(rsLotePiloto("ATUALIZA_STATUS"))
        End If
    End If

    ' Verifica se foi encontrado um lote piloto
    If Not lotePilotoEncontrado Then
        If Not IsNull(rsLotePiloto("MENSAGEM")) Then
            resultado("mensagem") = Trim(CStr(rsLotePiloto("MENSAGEM")))
        End If

        If resultado("mensagem") = "" Then
            resultado("mensagem") = "O item informado não corresponde a um lote piloto."
        End If

        resultado("valido") = False

        rsLotePiloto.Close
        Set rsLotePiloto = Nothing

        Set ValidarLotePiloto = resultado
        Exit Function
    End If

    ' Verifica a validacao do pedido
    If Not validaPedido Then
        mensagemValidacao = ""

        If Not IsNull(rsLotePiloto("MENSAGEM_PEDIDO")) Then
            mensagemValidacao = Trim(CStr(rsLotePiloto("MENSAGEM_PEDIDO")))
        End If

        If mensagemValidacao = "" Then
            mensagemValidacao =  "Pedido bloqueado pelas regras do lote piloto."
        End If

        resultado("valido") = False
        resultado("mensagem") = mensagemValidacao

        rsLotePiloto.Close
        Set rsLotePiloto = Nothing

        Set ValidarLotePiloto = resultado
        Exit Function
    End If

    ' Verifica a validacao da quantidade de paletes
    If Not validaPalete Then
        mensagemValidacao = ""

        If Not IsNull(rsLotePiloto("MENSAGEM_PALETE")) Then
            mensagemValidacao = Trim(CStr(rsLotePiloto("MENSAGEM_PALETE")))
        End If

        If mensagemValidacao = "" Then
            mensagemValidacao = "Pedido Bloqueado: Quantidade de paletes solicitada maior que a permitida para o lote piloto"
        End If

        resultado("valido") = False
        resultado("mensagem") = mensagemValidacao

        rsLotePiloto.Close
        Set rsLotePiloto = Nothing

        Set ValidarLotePiloto = resultado
        Exit Function
    End If

    ' Monta as observacoes de acompanhamento
    observacoes = ""

    If validaOnduladeira Then
        observacoes = "ACOMPANHAR NA ONDULADEIRA"
    End If

    If validaImpressora Then
        If observacoes <> "" Then
            observacoes = observacoes & " - "
        End If

        observacoes = observacoes & "ACOMPANHAR NA IMPRESSORA"
    End If

    If validaExpedicao Then
        If observacoes <> "" Then
            observacoes = observacoes & " - "
        End If

        observacoes = observacoes & "ACOMPANHAR NA EXPEDIÇÃO"
    End If

    If validaCliente Then
        If observacoes <> "" Then
            observacoes = observacoes & " - "
        End If

        observacoes = observacoes & "ACOMPANHAR O CLIENTE"
    End If

    If validaDDP Then
        If observacoes <> "" Then
            observacoes = observacoes & " - "
        End If

        observacoes = observacoes & "ACOMPANHAR DDP"
    End If

    If validaClicheria Then
        If observacoes <> "" Then
            observacoes = observacoes & " - "
        End If

        observacoes = observacoes & "ACOMPANHAR CLICHERIA"
    End If

    resultado("observacoes") = observacoes
    resultado("valido") = True

    ' Mensagem final
    mensagemValidacao = ""

    If Not IsNull(rsLotePiloto("MENSAGEM").Value) Then
        mensagemValidacao = Trim(CStr(rsLotePiloto("MENSAGEM").Value))
    End If

    If mensagemValidacao = "" Then
        mensagemValidacao = "Lote piloto validado com sucesso."
    End If

    resultado("mensagem") = mensagemValidacao

    rsLotePiloto.Close()
    Set rsLotePiloto = Nothing
    Set ValidarLotePiloto = resultado
End Function

Function ObterParametrosPlanta(codigoPlanta)
    Set parametrosPlanta = CreateObject("Scripting.Dictionary")
    Set dados = CreateObject("Scripting.Dictionary")

    SQL = "SELECT TOP 1 "
    SQL = SQL & "CargaMaquina, "
    SQL = SQL & "LimiteDeCredito, "
    SQL = SQL & "DiasViagem, "
    SQL = SQL & "DiaUtil, "
    SQL = SQL & "LarguraMin, "
    SQL = SQL & "LarguraMax, "
    SQL = SQL & "ComprimentoMin, "
    SQL = SQL & "ComprimentoMax, "
    SQL = SQL & "DiasAcrescimoPA, "
    SQL = SQL & "DiasAcrescimoCX, "
    SQL = SQL & "DiasAcrescimoCH, "
    SQL = SQL & "BloqueiaMaquinaCaixa, "
    SQL = SQL & "BloqueiaMaquinaChapa, "
    SQL = SQL & "BloqueiaMaquinaCorteVinco, "
    SQL = SQL & "DeMaleta, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(DeChapa, ''), 112), 23), '') AS DeChapa, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(AteChapa, ''), 112), 23), '') AS AteChapa, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(DeMaleta, ''), 112), 23), '') AS DeMaleta, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(AteMaleta, ''), 112), 23), '') AS AteMaleta, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(DeCorteVinco, ''), 112), 23), '') AS DeCorteVinco, "
    SQL = SQL & "ISNULL(CONVERT(char(10), CONVERT(datetime, NULLIF(AteCorteVinco, ''), 112), 23), '') AS AteCorteVinco, "
    SQL = SQL & "MtsCVmin, "
    SQL = SQL & "MtsCVmax, "
    SQL = SQL & "MtsCXmin, "
    SQL = SQL & "MtsCXmax, "
    SQL = SQL & "MtsCHmin, "
    SQL = SQL & "MtsCHmax, "
    SQL = SQL & "PrecoKgCorteVinco, "
    SQL = SQL & "PrecoKgMaleta, "
    SQL = SQL & "PrecoKgChapa, "
    SQL = SQL & "PesoMN, "
    SQL = SQL & "PesoCV, "
    SQL = SQL & "PesoCH "
    SQL = SQL & "FROM PARAMETROS_PEDIDO "
    SQL = SQL & "WHERE Planta = '" & codigoPlanta & "'"

    On Error Resume Next
    Set rsParametros = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 Then
        mensagemErro = Err.Description
        Err.Clear
        On Error GoTo 0

        parametrosPlanta.Add "sucesso", False
        parametrosPlanta.Add "mensagem", "Erro ao buscar parâmetros: " & mensagemErro

        Set ObterParametrosPlanta = parametrosPlanta
        Exit Function
    End If
    On Error GoTo 0

    If Not IsObject(rsParametros) Then
        parametrosPlanta.Add "sucesso", False
        parametrosPlanta.Add "mensagem", "Não foi possível obter os parâmetros da planta"

        Set ObterParametrosPlanta = parametrosPlanta
        Exit Function
    End If

    If rsParametros.EOF Then
        rsParametros.Close
        Set rsParametros = Nothing

        parametrosPlanta.Add "sucesso", False
        parametrosPlanta.Add "mensagem", "Nenhum parâmetro encontrado"

        Set ObterParametrosPlanta = parametrosPlanta
        Exit Function
    End If

    ' Campos booleanos convertidos para 0 ou 1.
    dados.Add "cargaMaquina", 0

    If Not IsNull(rsParametros("CargaMaquina")) Then
        If CBool(rsParametros("CargaMaquina")) Then
            dados("cargaMaquina") = 1
        End If
    End If

    dados.Add "limiteDeCredito", 0

    If Not IsNull(rsParametros("LimiteDeCredito")) Then
        If CBool(rsParametros("LimiteDeCredito")) Then
            dados("limiteDeCredito") = 1
        End If
    End If

    dados.Add "diasUtil", 0

    If Not IsNull(rsParametros("DiaUtil")) Then
        If CBool(rsParametros("DiaUtil")) Then
            dados("diasUtil") = 1
        End If
    End If

    dados.Add "diasViagem", 0

    If Not IsNull(rsParametros("DiasViagem")) Then
        If CBool(rsParametros("DiasViagem")) Then
            dados("diasViagem") = 1
        End If
    End If

    dados.Add "bloqueiaMaquinaCaixa", 0

    If Not IsNull(rsParametros("BloqueiaMaquinaCaixa")) Then
        If CBool(rsParametros("BloqueiaMaquinaCaixa")) Then
            dados("bloqueiaMaquinaCaixa") = 1
        End If
    End If

    dados.Add "bloqueiaMaquinaChapa", 0

    If Not IsNull(rsParametros("BloqueiaMaquinaChapa")) Then
        If CBool(rsParametros("BloqueiaMaquinaChapa")) Then
            dados("bloqueiaMaquinaChapa") = 1
        End If
    End If

    dados.Add "bloqueiaMaquinaCorteVinco", 0

    If Not IsNull(rsParametros("BloqueiaMaquinaCorteVinco")) Then
        If CBool(rsParametros("BloqueiaMaquinaCorteVinco")) Then
            dados("bloqueiaMaquinaCorteVinco") = 1
        End If
    End If

    ' Campos que podem retornar Null.
    dados.Add "larguraMin", ""
    If Not IsNull(rsParametros("LarguraMin")) Then
        dados("larguraMin") = rsParametros("LarguraMin")
    End If

    dados.Add "larguraMax", ""
    If Not IsNull(rsParametros("LarguraMax")) Then
        dados("larguraMax") = rsParametros("LarguraMax")
    End If

    dados.Add "comprimentoMin", ""
    If Not IsNull(rsParametros("ComprimentoMin")) Then
        dados("comprimentoMin") = rsParametros("ComprimentoMin")
    End If

    dados.Add "comprimentoMax", ""
    If Not IsNull(rsParametros("ComprimentoMax")) Then
        dados("comprimentoMax") = rsParametros("ComprimentoMax")
    End If

    dados.Add "diasAcrescimoPA", ""
    If Not IsNull(rsParametros("DiasAcrescimoPA")) Then
        dados("diasAcrescimoPA") = rsParametros("DiasAcrescimoPA")
    End If

    dados.Add "diasAcrescimoCX", ""
    If Not IsNull(rsParametros("DiasAcrescimoCX")) Then
        dados("diasAcrescimoCX") = rsParametros("DiasAcrescimoCX")
    End If

    dados.Add "diasAcrescimoCH", ""
    If Not IsNull(rsParametros("DiasAcrescimoCH")) Then
        dados("diasAcrescimoCH") = rsParametros("DiasAcrescimoCH")
    End If

    dados.Add "deMaleta", ""
    If Not IsNull(rsParametros("DeMaleta")) Then
        dados("deMaleta") = rsParametros("DeMaleta")
    End If

    dados.Add "ateMaleta", ""
    If Not IsNull(rsParametros("AteMaleta")) Then
        dados("ateMaleta") = rsParametros("AteMaleta")
    End If

    dados.Add "deCorteVinco", ""
    If Not IsNull(rsParametros("DeCorteVinco")) Then
        dados("deCorteVinco") = rsParametros("DeCorteVinco")
    End If

    dados.Add "ateCorteVinco", ""
    If Not IsNull(rsParametros("AteCorteVinco")) Then
        dados("ateCorteVinco") = rsParametros("AteCorteVinco")
    End If

    dados.Add "deChapa", ""
    If Not IsNull(rsParametros("DeChapa")) Then
        dados("deChapa") = rsParametros("DeChapa")
    End If

    dados.Add "ateChapa", ""
    If Not IsNull(rsParametros("AteChapa")) Then
        dados("ateChapa") = rsParametros("AteChapa")
    End If

    dados.Add "mtsCVmin", ""
    If Not IsNull(rsParametros("MtsCVmin")) Then
        dados("mtsCVmin") = rsParametros("MtsCVmin")
    End If

    dados.Add "mtsCVmax", ""
    If Not IsNull(rsParametros("MtsCVmax")) Then
        dados("mtsCVmax") = rsParametros("MtsCVmax")
    End If

    dados.Add "mtsCXmin", ""
    If Not IsNull(rsParametros("MtsCXmin")) Then
        dados("mtsCXmin") = rsParametros("MtsCXmin")
    End If

    dados.Add "mtsCXmax", ""
    If Not IsNull(rsParametros("MtsCXmax")) Then
        dados("mtsCXmax") = rsParametros("MtsCXmax")
    End If

    dados.Add "mtsCHmin", ""
    If Not IsNull(rsParametros("MtsCHmin")) Then
        dados("mtsCHmin") = rsParametros("MtsCHmin")
    End If

    dados.Add "mtsCHmax", ""
    If Not IsNull(rsParametros("MtsCHmax")) Then
        dados("mtsCHmax") = rsParametros("MtsCHmax")
    End If

    dados.Add "precoKgCorteVinco", ""
    If Not IsNull(rsParametros("PrecoKgCorteVinco")) Then
        dados("precoKgCorteVinco") = rsParametros("PrecoKgCorteVinco")
    End If

    dados.Add "precoKgMaleta", ""
    If Not IsNull(rsParametros("PrecoKgMaleta")) Then
        dados("precoKgMaleta") = rsParametros("PrecoKgMaleta")
    End If

    dados.Add "precoKgChapa", ""
    If Not IsNull(rsParametros("PrecoKgChapa")) Then
        dados("precoKgChapa") = rsParametros("PrecoKgChapa")
    End If

    dados.Add "pesoMN", ""
    If Not IsNull(rsParametros("PesoMN")) Then
        dados("pesoMN") = rsParametros("PesoMN")
    End If

    dados.Add "pesoCV", ""
    If Not IsNull(rsParametros("PesoCV")) Then
        dados("pesoCV") = rsParametros("PesoCV")
    End If

    dados.Add "pesoCH", ""
    If Not IsNull(rsParametros("PesoCH")) Then
        dados("pesoCH") = rsParametros("PesoCH")
    End If

    rsParametros.Close
    Set rsParametros = Nothing

    parametrosPlanta.Add "sucesso", True
    parametrosPlanta.Add "mensagem", "Lista obtida com sucesso"

    Set parametrosPlanta("dados") = dados
    Set ObterParametrosPlanta = parametrosPlanta
End Function

Function ValidarCargaMaquina( _
    codigoPlanta, _
    codigoProduto, _
    tipoFrete, _
    codigoCliente, _
    clienteRemessa, _
    quantidade, _
    dataEntrega, _
    estoque _
)
    Set resultado = CreateObject("Scripting.Dictionary")

    ' Estrutura padrão do retorno.
    resultado.Add "sucesso", False
    resultado.Add "mensagem", ""
    resultado.Add "dataDesejadaDisponivel", False
    resultado.Add "disponivel", 0
    resultado.Add "dataPossivel", Null
    resultado.Add "dataRetorno", Null
    resultado.Add "cargaMaquinaFechada", False
    resultado.Add "planta", codigoPlanta

    If Trim(dataEntrega) = "" Then
        resultado("mensagem") = "A data de entrega não foi informada."
        Set ValidarCargaMaquina = resultado
        Exit Function
    End If

    Select Case codigoPlanta
        Case "05"
            SQL = "SELECT "
            SQL = SQL & "DataDesejadaDisponivel, "
            SQL = SQL & "DataPossivel "
            SQL = SQL & "FROM dbo.fq_PreviaPedidos_InfoDataPossivel_Paema("
            SQL = SQL & "'" & codigoProduto & "', "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & tipoFrete & ", "
            SQL = SQL & "'" & codigoCliente & "', "
            SQL = SQL & "'" & clienteRemessa & "', "
            SQL = SQL & quantidade & ", "
            SQL = SQL & "'" & dataEntrega & "'" 
            SQL = SQL & ")"

        Case "06"
            SQL = "SELECT " 
            SQL = SQL & "DataDesejadaDisponivel, " 
            SQL = SQL & "DataPossivel " 
            SQL = SQL & "FROM dbo.fq_PreviaPedidos_InfoDataPossivel_SKMaranguape(" 
            SQL = SQL & "'" & codigoProduto & "', "
            SQL = SQL & tipoFrete & ", " 
            SQL = SQL & "'" & codigoCliente & "', "
            SQL = SQL & "'" & clienteRemessa & "', "
            SQL = SQL & quantidade & ", " 
            SQL = SQL & "'" & dataEntrega & "'" 
            SQL = SQL & ")"

        Case "07"
            SQL = "SELECT "
            SQL = SQL & "DataDesejadaDisponivel, "
            SQL = SQL & "DataPossivel "
            SQL = SQL & "FROM dbo.fq_PreviaPedidos_InfoDataPossivel_Paperbox("
            SQL = SQL & "'" & codigoProduto & "', "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & "NULL, "
            SQL = SQL & tipoFrete & ", "
            SQL = SQL & "'" & codigoCliente & "', "
            SQL = SQL & "'" & clienteRemessa & "', "
            SQL = SQL & quantidade & ", "
            SQL = SQL & "'" & dataEntrega & "'" 
            SQL = SQL & ")"

        Case Else
            resultado("mensagem") = "A planta '" & codigoPlanta & "' não está configurada para consultar a data possível de entrega."
            Set ValidarCargaMaquina = resultado
            Exit Function
    End Select

    On Error Resume Next
    Set rsConsulta = Conecta(ObterConexaoTrimbox(codigoPlanta), SQL)

    If Err.Number <> 0 Then
        descricaoErro = Err.Description
        Err.Clear

        resultado("mensagem") = "Erro ao consultar a data possível de entrega: " & descricaoErro
        Set ValidarCargaMaquina = resultado
        Exit Function
    End If

    If Not IsObject(rsConsulta) Then
        resultado("mensagem") = "Não foi possível obter o resultado da consulta."
        Set ValidarCargaMaquina = resultado
        Exit Function
    End If

    If rsConsulta.EOF Then
        resultado("mensagem") = "Carga Máquina não encontrada."
        rsConsulta.Close()
        Set rsConsulta = Nothing
        Set ValidarCargaMaquina = resultado
        Exit Function
    End If

    If Not IsNull(rsConsulta("DataDesejadaDisponivel")) Then
        disponivel = CBool(rsConsulta("DataDesejadaDisponivel"))
    Else
        disponivel = False
    End If

    resultado("dataDesejadaDisponivel") = disponivel

    If disponivel Then
        resultado("disponivel") = 1
    Else
        resultado("disponivel") = 0
    End If

    If Not IsNull(rsConsulta("DataPossivel")) Then
        dataPossivel = CDate(rsConsulta("DataPossivel"))
        resultado("dataPossivel") = dataPossivel

        If dataPossivel > DateSerial(3000, 1, 1) Then
            cargaMaquinaFechada = True

            resultado("cargaMaquinaFechada") = True
            resultado("mensagem") = "Carga máquina fechada."
            resultado("dataRetorno") = Null
        Else
            cargaMaquinaFechada = False
            dataRetorno = dataPossivel

            ' Regra específica para maranguape.
            If codigoPlanta = "06" And CStr(estoque) = "2" Then
                dataMinima = DateAdd("d", 7, Date())

                If dataPossivel < dataMinima Then
                    dataRetorno = dataPossivel
                Else
                    dataRetorno = dataMinima
                End If
            End If

            resultado("cargaMaquinaFechada") = False
            resultado("dataRetorno") = dataRetorno
            resultado("mensagem") = "Data possível encontrada."
        End If
    Else
        resultado("mensagem") = "A consulta não retornou uma data possível."
    End If

    rsConsulta.Close
    Set rsConsulta = Nothing

    If resultado("mensagem") = "Data possível encontrada." OR resultado("cargaMaquinaFechada") Then
        resultado("sucesso") = True
    End If

    If Err.Number <> 0 Then
        resultado("sucesso") = False
        resultado("mensagem") = "Erro ao processar o retorno da consulta: " & Err.Description
    End If
    On Error GoTo 0

    Set ValidarCargaMaquina = resultado
End Function

Function ObterConexaoCorporativo(codigoPlanta)
    Select Case codigoPlanta
        Case "00": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "04": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "05": ObterConexaoCorporativo = "BEN-CORPORATIVO"
        Case "06": ObterConexaoCorporativo = "FOR-CORPORATIVO"
        Case "07": ObterConexaoCorporativo = "SAQ-CORPORATIVO"
        Case Else: ObterConexaoCorporativo = ""
    End Select
End Function

Function ObterConexaoProtheus(codigoPlanta)
    Select Case codigoPlanta
        Case "00": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "04": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "05": ObterConexaoProtheus = "BEN-DADOSADV"
        Case "06": ObterConexaoProtheus = "FOR-DADOSADV"
        Case "07": ObterConexaoProtheus = "SAQ-DADOSADV"
        Case Else: ObterConexaoProtheus = ""
    End Select
End Function

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

Function ObterUltimoPedido(codigoPlanta)
    numeroPedido = 0

    SQL = "SELECT dbo.FUNC_NOVO_NUM_PEDIDO_LETRAS() as NumeroPedido "
  
    On Error Resume Next
    SET rsObterFT = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsObterFT) Then
        If Not rsObterFT.EOF Then
            numeroPedido = rsObterFT("NumeroPedido")
        End If

        rsObterFT.Close()
        SET rsObterFT = NOTHING
    End If
    On Error Goto 0

    ObterUltimoPedido = numeroPedido
End Function

Function ExtratoPedidos(codigoPlanta, pedidosSQL)
    If Trim(pedidosSQL) = "" Then
        Exit Function
    End If

    SQL = "SELECT Planta, Pedido, DataEmissao, Cliente, LojaCli, Quantidade, "
    SQL = SQL & "Largura, Comprimento, PrecoUnitario, Produto, Referencia, "
    SQL = SQL & "NumeroOrdemCompra, ItemOrdemCompra, TipoItem, "
    SQL = SQL & "PrecoUnitario * Quantidade AS TotalPedido, "
    SQL = SQL & "PrimeiraDataEntrega, DataPrevistaEntrega, "
    SQL = SQL & "ISNULL(CondicaoPagamento, 0) AS CondicaoPagamento, "
    SQL = SQL & "ClienteRemessa, FaturaRemessa, TESFatura, TESRemessa, TipoFrete, "
    SQL = SQL & "Gramatura, Diametro, Tubete, LarguraBobina, Estrutura, LotePiloto, "
    SQL = SQL & "Vinco1, Vinco2, Vinco3, Vinco4, Vinco5, Observacao "
    SQL = SQL & "FROM Pedido "
    SQL = SQL & "WHERE Pedido IN (" & pedidosSQL & ") "
    SQL = SQL & "AND ISNULL(Excluido, 0) = 0 "
    SQL = SQL & "ORDER BY Pedido"

    descricaoPlantaEmail = ""

    Select Case codigoPlanta
        Case "05"
            Set rsPedidosEmail = Conecta("BEN-CORPORATIVO", SQL)
            descricaoPlantaEmail = "Bento Golçalves"
        Case "06"
            Set rsPedidosEmail = Conecta("FOR-CORPORATIVO", SQL)
            descricaoPlantaEmail = "Maranguape"
        Case "07"
            Set rsPedidosEmail = Conecta("SAQ-CORPORATIVO", SQL)
            descricaoPlantaEmail = "Saquarema"
        Case Else
            Exit Function
    End Select

    If Not IsObject(rsPedidosEmail) Then
        Exit Function
    End If

    If rsPedidosEmail.EOF Then
        rsPedidosEmail.Close()
        Set rsPedidosEmail = Nothing
        Exit Function
    End If

    dataEmissaoEmail = rsPedidosEmail("DataEmissao")
    codigoClienteEmail = ValorOuVazio(rsPedidosEmail("Cliente"))
    codigoLojaEmail = ValorOuVazio(rsPedidosEmail("LojaCli"))
    codigoClienteRemessaEmail = ValorOuVazio(rsPedidosEmail("ClienteRemessa"))
    condicaoPagamentoEmail = ValorOuVazio(rsPedidosEmail("CondicaoPagamento"))
    tesFaturaEmail = ValorOuVazio(rsPedidosEmail("TESFatura"))
    tesRemessaEmail = ValorOuVazio(rsPedidosEmail("TESRemessa"))
    tipoFreteEmail = ValorOuVazio(rsPedidosEmail("TipoFrete"))
    faturaRemessaEmail = NumeroVB(rsPedidosEmail("FaturaRemessa"))
    totalPedidosEmail = UBound(Split(pedidosSQL, ",")) + 1

    razaoClienteEmail = ""
    municipioClienteEmail = ""

    SQL = "SELECT NOME_CLIENTE AS Nome, MUNICIPIO_CLIENTE AS Municipio "
    SQL = SQL & "FROM VW_CLIENTES_PEDIDO "
    SQL = SQL & "WHERE COD_CLIENTE = '" & Replace(codigoClienteEmail, "'", "''") & "' "
    SQL = SQL & "AND LOJA_CLIENTE = '" & Replace(codigoLojaEmail, "'", "''") & "'"

    Set rsAuxEmail = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If IsObject(rsAuxEmail) Then
        If Not rsAuxEmail.EOF Then
            razaoClienteEmail = Trim(ValorOuVazio(rsAuxEmail("Nome")))
            municipioClienteEmail = Trim(ValorOuVazio(rsAuxEmail("Municipio")))
        End If

        rsAuxEmail.Close()
        Set rsAuxEmail = Nothing
    End If

    descricaoCondicaoEmail = ""

    SQL = "SELECT DESCRICAO FROM VW_BI_CONDICAO_PAGAMENTO "
    SQL = SQL & "WHERE CODIGO = '" & Replace(condicaoPagamentoEmail, "'", "''") & "'"

    Set rsAuxEmail = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If IsObject(rsAuxEmail) Then
        If Not rsAuxEmail.EOF Then
            descricaoCondicaoEmail = Trim(ValorOuVazio(rsAuxEmail("DESCRICAO")))
        End If

        rsAuxEmail.Close()
        Set rsAuxEmail = Nothing
    End If

    descricaoTesFaturaEmail = ""

    SQL = "SELECT DESCRICAO FROM VW_BI_TES "
    SQL = SQL & "WHERE CODIGO = '" & Replace(tesFaturaEmail, "'", "''") & "'"

    Set rsAuxEmail = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

    If IsObject(rsAuxEmail) Then
        If Not rsAuxEmail.EOF Then
            descricaoTesFaturaEmail = Trim(ValorOuVazio(rsAuxEmail("DESCRICAO")))
        End If
        rsAuxEmail.Close()
        Set rsAuxEmail = Nothing
    End If

    tipoNotaFiscalEmail = "Normal"
    localEntregaEmail = codigoClienteRemessaEmail
    descricaoTesRemessaEmail = ""

    If faturaRemessaEmail > 0 Then
        tipoNotaFiscalEmail = "Opera&ccedil;&atilde;o triangular"

        SQL = "SELECT NOME_CLIENTE AS Nome, MUNICIPIO_CLIENTE AS Municipio "
        SQL = SQL & "FROM VW_CLIENTES_PEDIDO "
        SQL = SQL & "WHERE COD_CLIENTE = '" & Replace(codigoClienteRemessaEmail, "'", "''") & "'"

        Set rsAuxEmail = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

        If IsObject(rsAuxEmail) Then
            If Not rsAuxEmail.EOF Then
                localEntregaEmail = codigoClienteRemessaEmail
                localEntregaEmail = localEntregaEmail & " - " & Trim(ValorOuVazio(rsAuxEmail("Nome")))
                localEntregaEmail = localEntregaEmail & " - " & Trim(ValorOuVazio(rsAuxEmail("Municipio")))
            End If

            rsAuxEmail.Close()
            Set rsAuxEmail = Nothing
        End If

        SQL = "SELECT DESCRICAO FROM VW_BI_TES "
        SQL = SQL & "WHERE CODIGO = '" & Replace(tesRemessaEmail, "'", "''") & "'"

        Set rsAuxEmail = Conecta(ObterConexaoProtheus(codigoPlanta), SQL)

        If IsObject(rsAuxEmail) Then
            If Not rsAuxEmail.EOF Then
                descricaoTesRemessaEmail = Trim(ValorOuVazio(rsAuxEmail("DESCRICAO")))
            End If
            
            rsAuxEmail.Close()
            Set rsAuxEmail = Nothing
        End If
    End If

    If IsDate(dataEmissaoEmail) Then
        dataEmissaoFormatada = Right("0" & Day(CDate(dataEmissaoEmail)), 2)
        dataEmissaoFormatada = dataEmissaoFormatada & "/" & Right("0" & Month(CDate(dataEmissaoEmail)), 2)
        dataEmissaoFormatada = dataEmissaoFormatada & "/" & Year(CDate(dataEmissaoEmail))
    Else
        dataEmissaoFormatada = ValorOuVazio(dataEmissaoEmail)
    End If

    varEmail = "<html>"
    varEmail = varEmail & "<head><meta charset=""utf-8"" /></head>"
    varEmail = varEmail & "<body style=""margin:0;padding:24px;background:#ffffff;"
    varEmail = varEmail & "font-family:Arial,Helvetica,sans-serif;color:#1f2937;"">"

    varEmail = varEmail & "<div style=""padding:18px 20px;background:#263849;color:#ffffff;"">"
    varEmail = varEmail & "<div style=""font-size:21px;font-weight:bold;line-height:26px;"">"
    varEmail = varEmail & "Extrato de Pedidos Emitidos</div>"
    varEmail = varEmail & "<div style=""margin-top:5px;font-size:12px;color:#dbe2e8;"">"
    varEmail = varEmail & totalPedidosEmail & " pedido(s) emitido(s) </div>"
    varEmail = varEmail & "</div>"

    varEmail = varEmail & "<div style=""margin-top:24px;padding:7px 10px;background:#3f5265;"
    varEmail = varEmail & "color:#ffffff;font-size:12px;font-weight:bold;"">Dados Gerais</div>"
    varEmail = varEmail & "<table width=""100%"" cellspacing=""0"" cellpadding=""0"" "
    varEmail = varEmail & "style=""border-collapse:collapse;border:1px solid #d1d5db;font-size:12px;"">"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Data de Emiss&atilde;o</td>"
    varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(dataEmissaoFormatada) & "</td>"
    varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Usu&aacute;rio Emitente</td>"
    varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(usuarioLogado("nome")) & "</td>"
    varEmail = varEmail & "</tr>"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Cliente</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(codigoClienteEmail & " - " & razaoClienteEmail) & "</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Unidade</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(codigoLojaEmail & " - " & municipioClienteEmail) & "</td>"
    varEmail = varEmail & "</tr>"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Condi&ccedil;&atilde;o de Pagamento</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(condicaoPagamentoEmail & " - " & descricaoCondicaoEmail) & "</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Local de Entrega</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(localEntregaEmail) & "</td>"
    varEmail = varEmail & "</tr>"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Tipo de NF</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(tipoNotaFiscalEmail) & "</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Frete</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">"
    
    If tipoFreteEmail = "0" Then
        varEmail = varEmail & "CIF"
    Else
        varEmail = varEmail & "FOB"
    End If

    varEmail = varEmail & "</td></tr>"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">TES</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(tesFaturaEmail) & "</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Tipo OP Fatura</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(tesFaturaEmail & " - " & descricaoTesFaturaEmail) & "</td>"
    varEmail = varEmail & "</tr>"

    varEmail = varEmail & "<tr>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Tipo OP Remessa</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(tesRemessaEmail & " - " & descricaoTesRemessaEmail) & "</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Local de Fabrica&ccedil;&atilde;o</td>"
    varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(descricaoPlantaEmail) & "</td>"
    varEmail = varEmail & "</tr></table>"

    listaNumerosEmail = ""
    primeiroPedidoEmail = True

    Do While Not rsPedidosEmail.EOF
        numeroPedidoEmail = ValorOuVazio(rsPedidosEmail("Pedido"))
        tipoItem = ValorOuVazio(rsPedidosEmail("TipoItem"))

        If listaNumerosEmail <> "" Then
            listaNumerosEmail = listaNumerosEmail & ", "
        End If
        listaNumerosEmail = listaNumerosEmail & numeroPedidoEmail

        If primeiroPedidoEmail Then
            espacoPedidoEmail = "38px"
            primeiroPedidoEmail = False
        Else
            espacoPedidoEmail = "64px"
        End If

        If CStr(ValorOuVazio(rsPedidosEmail("LotePiloto"))) = "1" Then
            lotePilotoEmail = "Sim"
        Else
            lotePilotoEmail = "Nao"
        End If

        prazoEntregaEmail = ValorOuVazio(rsPedidosEmail("PrimeiraDataEntrega"))
        If IsDate(prazoEntregaEmail) Then
            prazoEntregaFormatado = Right("0" & Day(CDate(prazoEntregaEmail)), 2)
            prazoEntregaFormatado = prazoEntregaFormatado & "/" & Right("0" & Month(CDate(prazoEntregaEmail)), 2)
            prazoEntregaFormatado = prazoEntregaFormatado & "/" & Year(CDate(prazoEntregaEmail))
        Else
            prazoEntregaFormatado = prazoEntregaEmail
        End If

        dataPrevistaEmail = ValorOuVazio(rsPedidosEmail("DataPrevistaEntrega"))
        If IsDate(dataPrevistaEmail) Then
            dataPrevistaFormatada = Right("0" & Day(CDate(dataPrevistaEmail)), 2)
            dataPrevistaFormatada = dataPrevistaFormatada & "/" & Right("0" & Month(CDate(dataPrevistaEmail)), 2)
            dataPrevistaFormatada = dataPrevistaFormatada & "/" & Year(CDate(dataPrevistaEmail))
        Else
            dataPrevistaFormatada = dataPrevistaEmail
        End If

        varEmail = varEmail & "<div style=""margin-top:" & espacoPedidoEmail & ";padding-bottom:9px;"
        varEmail = varEmail & "border-bottom:2px solid #3f5265;font-size:16px;font-weight:bold;color:#263849;"">"
        varEmail = varEmail & "Pedido " & HtmlEmail(numeroPedidoEmail) & "</div>"

        varEmail = varEmail & "<div style=""margin-top:18px;padding:7px 10px;background:#3f5265;"
        varEmail = varEmail & "color:#ffffff;font-size:12px;font-weight:bold;"">Dados do Pedido</div>"
        varEmail = varEmail & "<table width=""100%"" cellspacing=""0"" cellpadding=""0"" "
        varEmail = varEmail & "style=""border-collapse:collapse;border:1px solid #d1d5db;font-size:12px;"">"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Produto</td>"
        varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Produto"))) & "</td>"
        varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Total</td>"
        varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">R$ " & FormatNumber(NumeroVB(rsPedidosEmail("TotalPedido")), 2) & "</td>"
        varEmail = varEmail & "</tr>"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Pedido Cliente</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("NumeroOrdemCompra"))) & "</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Item Pedido Cliente</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("ItemOrdemCompra"))) & "</td>"
        varEmail = varEmail & "</tr>"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Prazo de Entrega</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(prazoEntregaFormatado) & "</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Data Prevista de Entrega</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(dataPrevistaFormatada) & "</td>"
        varEmail = varEmail & "</tr>"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Local de Fabrica&ccedil;&atilde;o</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(descricaoPlantaEmail) & "</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Lote Piloto</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & lotePilotoEmail & "</td>"
        varEmail = varEmail & "</tr>"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Quantidade</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Quantidade")), 0) & "</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Pre&ccedil;o Unit&aacute;rio</td>"
        varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">R$ " & FormatNumber(NumeroVB(rsPedidosEmail("PrecoUnitario")), 5) & "</td>"
        varEmail = varEmail & "</tr>"

        varEmail = varEmail & "<tr>"
        varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Refer&ecirc;ncia</td>"
        varEmail = varEmail & "<td colspan=""3"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Referencia"))) & "</td>"
        varEmail = varEmail & "</tr></table>"

        If tipoItem = "PA" Then
            varEmail = varEmail & "<div style=""margin-top:18px;padding:7px 10px;background:#3f5265;color:#ffffff;font-size:12px;font-weight:bold;"">Dados do Papel</div>"
            varEmail = varEmail & "<table width=""100%"" cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;border:1px solid #d1d5db;font-size:12px;"">"
            varEmail = varEmail & "<tr>"
            varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Gramatura</td>"
            varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Gramatura"))) & "</td>"
            varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Di&acirc;metro</td>"
            varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Diametro"))) & "</td>"
            varEmail = varEmail & "</tr><tr>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Formato</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("LarguraBobina"))) & "</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Tubete</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Tubete"))) & "</td>"
            varEmail = varEmail & "</tr></table>"
        End If

        If tipoItem <> "PA" Then
            varEmail = varEmail & "<div style=""margin-top:18px;padding:7px 10px;background:#3f5265;color:#ffffff;font-size:12px;font-weight:bold;"">Dados Chapa</div>"
            varEmail = varEmail & "<table width=""100%"" cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;border:1px solid #d1d5db;font-size:12px;"">"
            varEmail = varEmail & "<tr>"
            varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Estrutura</td>"
            varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & HtmlEmail(ValorOuVazio(rsPedidosEmail("Estrutura"))) & "</td>"
            varEmail = varEmail & "<td width=""19%"" style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Largura</td>"
            varEmail = varEmail & "<td width=""31%"" style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Largura")), 0) & "</td>"
            varEmail = varEmail & "</tr><tr>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Comprimento</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Comprimento")), 0) & "</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Vinco 1</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Vinco1")), 0) & "</td>"
            varEmail = varEmail & "</tr><tr>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Vinco 2</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Vinco2")), 0) & "</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Vinco 3</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Vinco3")), 0) & "</td>"
            varEmail = varEmail & "</tr><tr>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Vinco 4</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Vinco4")), 0) & "</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;background:#f3f4f6;border:1px solid #d1d5db;font-weight:bold;"">Vinco 5</td>"
            varEmail = varEmail & "<td style=""padding:7px 8px;border:1px solid #d1d5db;"">" & FormatNumber(NumeroVB(rsPedidosEmail("Vinco5")), 0) & "</td>"
            varEmail = varEmail & "</tr></table>"
        End If

        varEmail = varEmail & "<div style=""margin-top:18px;padding:7px 10px;background:#3f5265;color:#ffffff;font-size:12px;font-weight:bold;"">Observa&ccedil;&otilde;es do Pedido</div>"
        varEmail = varEmail & "<div style=""padding:10px 12px;border:1px solid #d1d5db;background:#fafafa;font-size:12px;line-height:18px;"">"
        varEmail = varEmail & HtmlEmail(ValorOuVazio(rsPedidosEmail("Observacao"))) & "</div>"

        rsPedidosEmail.MoveNext()
    Loop

    varEmail = varEmail & "</body></html>"

    rsPedidosEmail.Close()
    Set rsPedidosEmail = Nothing

    emailPrincipal = ""
    SQL = "SELECT CRM_EMAILPRIMARIO FROM CRM_DTI "
    SQL = SQL & "WHERE CRM_R_E_C_N_O = " & Session("CODIGO")

    Set rsAuxEmail = Conecta("PIR-HELPDESK", SQL)

    If IsObject(rsAuxEmail) Then
        If Not rsAuxEmail.EOF Then
            emailPrincipal = ValorOuVazio(rsAuxEmail("CRM_EMAILPRIMARIO"))
        End If
        rsAuxEmail.Close()
        Set rsAuxEmail = Nothing
    End If

    If Trim(emailPrincipal) <> "" Then
        assuntoEmail = "Extrato de " & totalPedidosEmail & " pedido(s): " & listaNumerosEmail
        Call EnviaEmail( _
            emailPrincipal, _
            "", _
            "", _
            "", _
            "<inpanet.inpa@smurfitkappa.com.br>Sistema de Automacao Smurfit Kappa", _
            varEmail, _
            assuntoEmail _
        )
    End If
End Function

Function HtmlEmail(valor)
    If IsNull(valor) Then
        textoEmail = ""
    Else
        textoEmail = CStr(valor)
    End If

    textoEmail = Replace(textoEmail, "&", "&amp;")
    textoEmail = Replace(textoEmail, "<", "&lt;")
    textoEmail = Replace(textoEmail, ">", "&gt;")
    textoEmail = Replace(textoEmail, Chr(34), "&quot;")
    textoEmail = Replace(textoEmail, vbCrLf, "<br />")
    textoEmail = Replace(textoEmail, vbLf, "<br />")

    If Trim(Replace(textoEmail, "<br />", "")) = "" Then
        textoEmail = "Nao informado"
    End If

    HtmlEmail = textoEmail
End Function

Sub RegistrarLog(acao, id)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Emissão de Pedido"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a Pedido de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'Pedidos', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET rsLog = Conecta(ObterConexaoCorporativo(codigoPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsLog) Then
        rsLog.Close()
        SET rsLog = NOTHING
    End If
    On Error Goto 0
End Sub
%>