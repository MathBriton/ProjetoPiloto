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
    Case "listarVendasClientes"
        resultado = ListarVendasClientes()
    Case "listarVendasLojas"
        resultado = ListarVendasLojas()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function ListarVendasClientes()
    codPlanta = Request("filtroFilial")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    dataInicial = Request("filtroDataInicial")
    dataFinal = Request("filtroDataFinal")
    nomePlanta = descricaoDaPlanta(codPlanta)
    representante = Request("filtroRepresentante")
    valores = iif(Request("filtroTipoProduto") = "1", 1, request("filtroValores")) 
    tipoProduto = Request("filtroTipoProduto")
    tipoPapelao = Request("filtroTipoPapelao")
    clienteGrupo = Request("filtroClienteGrupo")
    impostos =  Request("filtroImpostos")

    If codPlanta = "" Then
        ListarVendasClientes = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    If dataInicial = "" OR dataFinal = "" Then
        ListarVendasClientes = "{""sucesso"": false, ""mensagem"": ""Data não informada""}"
        Exit Function
    End If

    If nomePlanta = "MARANGUAPE" Then
        clausulaPlanta = "'MARANGUAPE', 'FORTALEZA'"
    Else
        clausulaPlanta = "'" & nomePlanta & "'"
    End If

    nomeTabela = Iif(atualizandoBD(), "VW_RESUMO_VENDAS_TEMP", "VW_RESUMO_VENDAS") 
    

    SQL = SQL & "SELECT                                                                         "
    SQL = SQL & "    LTRIM(RTRIM(NOME_VENDEDOR)) AS NomeVendedor, "
    SQL = SQL & "    LTRIM(RTRIM(GRUPO_DO_CLIENTE)) AS GrupoCliente, "
    SQL = SQL & "    LTRIM(RTRIM(FANTASIA_CLIENTE)) AS NomeCliente, "
    SQL = SQL & "    CONCAT(ANO,MES) AS Ano,"
    SQL = SQL & "    SUM(PESO) AS Peso,"
    SQL = SQL & "    SUM(AREA) AS Area,"
    SQL = SQL & "    CASE " & impostos & " "
    SQL = SQL & "        WHEN 0 THEN SUM(VALOR_TOTAL)"
    SQL = SQL & "        WHEN 1 THEN SUM(VALOR_SEM_IPI)"
    SQL = SQL & "        WHEN 2 THEN SUM(VALOR_NET)"
    SQL = SQL & "    END as VALOR "
    SQL = SQL & "FROM " & nomeTabela &  " "
    SQL = SQL & "WHERE CIDADE IN (" & clausulaPlanta & ") "
    SQL = SQL & "AND DATA_EMISSAO >= '" & Replace(dataInicial, "-", "") & "' "
    SQL = SQL & "AND DATA_EMISSAO <= '" & Replace(dataFinal, "-", "") & "' "
    
    If tipoProduto <> "0" Then
        grupoProduto = iif(tipoProduto = "1", "PAPEL","PAPELAO")
        SQL = SQL & "AND GRUPO_PRODUTO = '" & grupoProduto & "' "
    End If

    If tipoProduto <> "1" Then
        If tipoPapelao = "1" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CAIXA' "
        End If

        If tipoPapelao = "2" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CHAPA' "
        End If

        If tipoPapelao = "3" Then
            SQL = SQL & "AND TIPO_CAIXA = 'CORTE VINCO' "
        End If

        If tipoPapelao = "4" Then
            SQL = SQL & "AND TIPO_CAIXA = 'MALETA' "
        End If
    End If

    If clienteGrupo <> "T" Then
        SQL = SQL & "AND CLIENTE_GRUPO = '" & clienteGrupo & "' "
    End If

    If representante <> "" Then
        SQL = SQL & "AND COD_VENDEDOR = '" & representante & "' "
    End If

    If codCliente <> "" Then
        If codPlanta = "06" then
            SQL = SQL & "AND (CODIGO_CLIENTE = '" & codCliente & "' OR CODIGO_ERP = '" & codCliente & "') "
        Else
            SQL = SQL & "AND CODIGO_CLIENTE = '" & codCliente & "' "
        End If
    End If

    If tipoProduto <> "0" Then
        grupoProduto = iif(tipoProduto = "1", "PAPEL","PAPELAO")
        SQL = SQL & "AND GRUPO_PRODUTO = '" & grupoProduto & "' "
    End If

    If tipoProduto <> "1" Then
        If tipoPapelao = "1" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CAIXA' "
        End If

        If tipoPapelao = "2" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CHAPA' "
        End If

        If tipoPapelao = "3" Then
            SQL = SQL & "AND TIPO_CAIXA = 'CORTE VINCO' "
        End If

        If tipoPapelao = "4" Then
            SQL = SQL & "AND TIPO_CAIXA = 'MALETA' "
        End If
    End If
    
    SQL = SQL & "GROUP BY "
    SQL = SQL & "    NOME_VENDEDOR, "
    SQL = SQL & "    GRUPO_DO_CLIENTE, "
    SQL = SQL & "    FANTASIA_CLIENTE,"
    SQL = SQL & "    CONCAT(ANO,MES)"
    SQL = SQL & "ORDER BY "
    SQL = SQL & "    NOME_VENDEDOR, "
    SQL = SQL & "    FANTASIA_CLIENTE"


    On Error Resume Next
    SET rsDados = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        ListarVendasClientes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jsonItens = "["
    primeiroItem = True

    If Not rsDados.EOF Then
        Do While Not rsDados.EOF
            If Not primeiroItem Then jsonItens = jsonItens & "," End If
            
            jsonItens = jsonItens & "{"
            jsonItens = jsonItens & """nomeVendedor"": """ & EscapeJSON(rsDados("NomeVendedor")) & """, "
            jsonItens = jsonItens & """grupoCliente"": """ & EscapeJSON(rsDados("GrupoCliente")) & """, "
            jsonItens = jsonItens & """nomeCliente"": """ & EscapeJSON(rsDados("NomeCliente")) & """, "
            jsonItens = jsonItens & """ano"": """ & EscapeJSON(rsDados("Ano")) & """, "
            jsonItens = jsonItens & """peso"": """ & Replace(EscapeJSON(rsDados("Peso")), ",", ".") & """, "
            jsonItens = jsonItens & """area"": """ &  Replace(EscapeJSON(rsDados("Area")), ",", ".") & """, "
            jsonItens = jsonItens & """valor"": """ &  Replace(EscapeJSON(rsDados("Valor")), ",", ".") & """ "
            jsonItens = jsonItens & "}"

            primeiroItem = False
            rsDados.MoveNext
        Loop
    End If

    jsonItens = jsonItens & "]"

    rsDados.Close()
    SET rsDados = NOTHING 

    ListarVendasClientes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonItens & "}"
End Function

Function ListarVendasLojas()
    codPlanta = Request("filtroFilial")
    codGrupo = Request("filtroCodGrupo")
    codLoja = Request("filtroCodLoja")
    dataInicial = Request("filtroDataInicial")
    dataFinal = Request("filtroDataFinal")
    nomePlanta  = descricaoDaPlanta(codPlanta)
    representante = Request("filtroRepresentante")
    valores = iif(Request("filtroTipoProduto") = "1", 1, request("filtroValores")) 
    tipoProduto = Request("filtroTipoProduto")
    tipoPapelao = Request("filtroTipoPapelao")
    clienteGrupo = Request("filtroClienteGrupo")
    impostos =  Request("filtroImpostos")

    If codPlanta = "" Then
        ListarVendasLojas = "{""sucesso"": false, ""mensagem"": ""Planta não encontrada""}"
        Exit Function
    End If

    If dataInicial = "" OR dataFinal = "" Then
        ListarVendasLojas = "{""sucesso"": false, ""mensagem"": ""Data não informada""}"
        Exit Function
    End If

    nomeTabela = Iif(atualizandoBD(), "VW_RESUMO_VENDAS_LOJA_TEMP", "VW_RESUMO_VENDAS_LOJA") 

    If nomePlanta = "MARANGUAPE" Then
        clausulaPlanta = "'MARANGUAPE', 'FORTALEZA'"
    Else
        clausulaPlanta = "'" & nomePlanta & "'"
    End If
    
    SQL = SQL & "SELECT "
    SQL = SQL & "LTRIM(RTRIM(NOME_VENDEDOR)) AS NomeVendedor, "
    SQL = SQL & "CONCAT(LTRIM(RTRIM(CODIGO_CLIENTE)), ' | ', LTRIM(RTRIM(LOJA_CLIENTE)) ) AS CodigoCliente, "
    SQL = SQL & "CONCAT(LTRIM(RTRIM(FANTASIA_CLIENTE)), ' | ', LTRIM(RTRIM(LOJA_CLIENTE)), ' | ' , LTRIM(RTRIM(CIDADE_CLIENTE)), '/', LTRIM(RTRIM(ESTADO_CLIENTE))) AS Loja, "
    SQL = SQL & "    CONCAT(ANO,MES) AS Ano,"
    SQL = SQL & "    SUM(PESO) AS Peso,"
    SQL = SQL & "    SUM(AREA) AS Area,"
    SQL = SQL & "    CASE " & impostos & " "
    SQL = SQL & "        WHEN 0 THEN SUM(VALOR_TOTAL)"
    SQL = SQL & "        WHEN 1 THEN SUM(VALOR_SEM_IPI)"
    SQL = SQL & "        WHEN 2 THEN SUM(VALOR_NET)"
    SQL = SQL & "    END as VALOR "
    SQL = SQL & "FROM " & nomeTabela &  " "
    SQL = SQL & "WHERE CIDADE IN (" & clausulaPlanta & ") "
    SQL = SQL & "AND DATA_EMISSAO >= '" & Replace(dataInicial, "-", "") & "' "
    SQL = SQL & "AND DATA_EMISSAO <= '" & Replace(dataFinal, "-", "") & "' "

    If tipoProduto <> "0" Then
        grupoProduto = iif(tipoProduto = "1", "PAPEL","PAPELAO")
        SQL = SQL & "AND GRUPO_PRODUTO = '" & grupoProduto & "' "
    End If

    If tipoProduto <> "1" Then
        If tipoPapelao = "1" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CAIXA' "
        End If

        If tipoPapelao = "2" Then
            SQL = SQL & "AND TIPO_PRODUTO = 'CHAPA' "
        End If

        If tipoPapelao = "3" Then
            SQL = SQL & "AND TIPO_CAIXA = 'CORTE VINCO' "
        End If

        If tipoPapelao = "4" Then
            SQL = SQL & "AND TIPO_CAIXA = 'MALETA' "
        End If
    End If

    If representante <> "" Then
        SQL = SQL & "AND COD_VENDEDOR = '" & representante & "' "
    End If

    If codGrupo <> "" Then
        SQL = SQL & "AND GRUPO_DO_CLIENTE = '" & codGrupo & "' "
    End If
    

    SQL = SQL & "GROUP BY NOME_VENDEDOR, CONCAT(ANO,MES), GRUPO_DO_CLIENTE, CODIGO_CLIENTE, FANTASIA_CLIENTE, CIDADE_CLIENTE, LOJA_CLIENTE, ESTADO_CLIENTE "
    SQL = SQL & "ORDER BY NOME_VENDEDOR, FANTASIA_CLIENTE"

    On Error Resume Next
    SET rsDados = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        ListarVendasLojas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jsonItens = "["
    primeiroItem = True

    If Not rsDados.EOF Then
        Do While Not rsDados.EOF
            If Not primeiroItem Then jsonItens = jsonItens & "," End If
            
            jsonItens = jsonItens & "{"
            jsonItens = jsonItens & """nomeVendedor"": """ & EscapeJSON(rsDados("NomeVendedor")) & """, "
            jsonItens = jsonItens & """codigoCliente"": """ & EscapeJSON(rsDados("CodigoCliente")) & """, "
            jsonItens = jsonItens & """loja"": """ & EscapeJSON(rsDados("Loja")) & """, "
            jsonItens = jsonItens & """ano"": """ & EscapeJSON(rsDados("Ano")) & """, "
            jsonItens = jsonItens & """peso"": """ & Replace(EscapeJSON(rsDados("Peso")), ",", ".") & """, "
            jsonItens = jsonItens & """area"": """ &  Replace(EscapeJSON(rsDados("Area")), ",", ".") & """, "
            jsonItens = jsonItens & """valor"": """ &  Replace(EscapeJSON(rsDados("Valor")), ",", ".") & """ "            
            jsonItens = jsonItens & "}"

            primeiroItem = False
            rsDados.MoveNext
        Loop
    End If

    jsonItens = jsonItens & "]"

    rsDados.Close()
    SET rsDados = NOTHING 

    ListarVendasLojas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonItens & "}"
End Function

' Auxiliares
Function descricaoDaPlanta(codigo)
    Select Case codigo
        Case "00" descricaoDaPlanta = "PIRAPETINGA"
        Case "04" descricaoDaPlanta = "UBERABA"
        Case "05" descricaoDaPlanta = "BENTO"
        Case "06" descricaoDaPlanta = "MARANGUAPE"
        Case "07" descricaoDaPlanta = "SAQUAREMA"
        Case "10" descricaoDaPlanta = "ESHOP"
    End Select
End Function
%>