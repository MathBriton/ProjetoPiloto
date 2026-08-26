<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Buffer       = True
Response.Charset      = "UTF-8"
Response.CodePage     = 65001
Response.ContentType  = "application/json"
Response.Expires      = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

'--------------------------------------------------------------------------------------------------
' HELPERS DE SERIALIZAÇÃO JSON
'--------------------------------------------------------------------------------------------------
Function jsonEscape(ByVal valor)
  Dim s
  If IsNull(valor) Then
    jsonEscape = ""
    Exit Function
  End If
  s = CStr(valor)
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, Chr(13) & Chr(10), "\n")
  s = Replace(s, Chr(13), "\n")
  s = Replace(s, Chr(10), "\n")
  s = Replace(s, Chr(9), "\t")
  jsonEscape = s
End Function

Function jsonStr(ByVal valor)
  jsonStr = """" & jsonEscape(valor) & """"
End Function

Function jsonBool(ByVal valor)
  If CBool(valor) Then jsonBool = "true" Else jsonBool = "false"
End Function

' Converte número para JSON com ponto como separador decimal
Function jsonNum(ByVal valor)
  Dim s
  If IsNull(valor) Or Trim(valor & "") = "" Then
    jsonNum = "0"
    Exit Function
  End If
  s = CStr(valor)
  s = Replace(s, ",", ".")
  jsonNum = s
End Function

' Converte data (yyyy-mm-dd vinda do <input type="date">) para AAAAMMDD
Function dataParaYYYYMMDD(ByVal data)
  dataParaYYYYMMDD = Replace(Trim(data & ""), "-", "")
End Function

' Formata data US (yyyy-mm-dd) -> dd/mm/yyyy
Function formataDataUStoBR(ByVal data)
  Dim dia, mes, ano, ret
  dia = "" : mes = "" : ano = "" : ret = ""
  If Trim(data & "") <> "" Then
    If Len(data) = 10 Then
      If Mid(data, 5, 1) = "-" Then
        ano = Mid(data, 1, 4)
        mes = Mid(data, 6, 2)
        dia = Mid(data, 9, 2)
        ret = dia & "/" & mes & "/" & ano
      Else
        ret = data
      End If
    End If
  End If
  formataDataUStoBR = ret
End Function

' Sanitização básica para evitar SQL Injection (escapa aspas simples)
Function sqlSafe(ByVal valor)
  sqlSafe = Replace(Trim(valor & ""), "'", "''")
End Function

Sub responderJson(ByVal json)
  Response.Write json
  Response.End
End Sub

Sub responderErro(ByVal mensagem)
  Response.Write "{""sucesso"":false,""mensagem"":" & jsonStr(mensagem) & "}"
  Response.End
End Sub

'--------------------------------------------------------------------------------------------------
' ROTEAMENTO DE AÇÕES
'--------------------------------------------------------------------------------------------------

Dim acao
acao = LCase(Trim(Request("acao")))

Select Case acao
  Case "obterpedidos"
    Call obterPedidos()
  Case "obterdetalhes"
    Call obterDetalhes()
  Case "validarcliente"
    Call validarCliente()
  Case Else
    Call responderErro("Ação inválida ou não informada.")
End Select

'==================================================================================================
' AÇÃO: obterPedidos
' DESCRIÇÃO: Retorna a lista de pedidos em carteira aplicando os filtros recebidos.
'==================================================================================================
Sub obterPedidos()

  Dim pCidade, pCodCli, pLojaCli, pPedido, pRepre
  Dim pGrupoProduto, pTipoProduto, pTipoFrete, pPeriodoPor
  Dim pDataIni, pDataFim
  Dim flagAtualizando, SQL, RST_VALIDACAO, CONSULTA
  Dim primeiro, json
  Dim somaValorTotal, somaValorNet, somaAreaVendida, somaPesoVendido
  Dim somaAreaSaldo, somaPesoSaldo, somaSaldoSemIPI, somaSaldoComIPI, somaVlrNetSaldo
  Dim totalRegistros

  pCidade        = sqlSafe(Request("cboCidade"))
  pCodCli        = sqlSafe(Request("codcli"))
  pLojaCli       = sqlSafe(Request("lojacli"))
  pPedido        = UCase(sqlSafe(Request("pedido")))
  pRepre         = sqlSafe(Request("representante"))
  pGrupoProduto  = sqlSafe(Request("cboGrupoProduto"))
  pTipoProduto   = sqlSafe(Request("cboTipoProduto"))
  pTipoFrete     = sqlSafe(Request("cboTipoFrete"))
  pPeriodoPor    = sqlSafe(Request("cboPeriodoPor"))
  pDataIni       = dataParaYYYYMMDD(Request("dataini"))
  pDataFim       = dataParaYYYYMMDD(Request("datafin"))

  ' Validações mínimas
  If pCidade = "" Then
    Call responderErro("Selecione a filial!")
  End If
  If pDataIni = "" Then
    Call responderErro("Informe a data inicial!")
  End If
  If pDataFim = "" Then
    Call responderErro("Informe a data final!")
  End If

  ' Verifica se a base está em atualização
  SQL = "SELECT TOP 1 VALIDACAO FROM BI_ATUALIZACAO"
  Set RST_VALIDACAO = Conecta("SPO-BI", SQL)
  If Not IsObject(RST_VALIDACAO) Then
    Call responderErro("Falha ao validar status da base.")
  End If
  flagAtualizando = False
  If Not RST_VALIDACAO.EOF Then flagAtualizando = RST_VALIDACAO("VALIDACAO")
  RST_VALIDACAO.Close()
  Set RST_VALIDACAO = Nothing

  ' Monta a query
  SQL = "" & _
    "SELECT " & _
      "PEDIDO " & _
      ",ITEM " & _
      ",EMISSAO " & _
      ",DT_ENTREGA1 " & _
      ",DATA_EMBARQUE " & _
      ",CASE WHEN CIDADE IN ('MARANGUAPE','BENTO') THEN CODIGO_ERP ELSE COD_CLI END AS COD_CLI " & _
      ",LOJA_CLI " & _
      ",NOME_REDUZIDO_CLI " & _
      ",PEDIDO_CLIENTE " & _
      ",ITEM_PEDIDO_CLIENTE " & _
      ",TIPO_FRETE " & _
      ",CONDICAO_PAGAMENTO " & _
      ",PRODUTO " & _
      ",DESCRICAO_PRODUTO " & _
      ",ESTRUTURA " & _
      ",GRUPO_PRODUTO " & _
      ",TIPO_PRODUTO " & _
      ",MAQUINA " & _
      ",COD_REPRE " & _
      ",NOME_REDUZIDO_REP " & _
      ",QUANT_VENDIDA " & _
      ",QUANT_SALDO_ENTREGAR " & _
      ",VALOR_UNITARIO " & _
      ",AREA_UNITARIA " & _
      ",PESO_UNITARIO " & _
      ",VALOR_TOTAL " & _
      ",VALOR_NET " & _
      ",AREA_VENDIDA " & _
      ",PESO_VENDIDO " & _
      ",AREA_SALDO " & _
      ",PESO_SALDO " & _
      ",VALOR_TOTAL_SALDO_SEM_IPI " & _
      ",VALOR_TOTAL_SALDO_COM_IPI " & _
      ",VALOR_NET_SALDO " & _
      ",PRECO_KG_NET " & _
      ",PRECO_M2_NET " & _
      ",CIDADE_CLIENTE " & _
      ",ESTADO_CLIENTE " & _
      iif(flagAtualizando, ",CAST(NULL AS VARCHAR(10)) AS TES ", ",TES ") & _
    iif(flagAtualizando, "FROM VW_PEDIDO_CARTEIRA_TEMP ", "FROM VW_PEDIDO_CARTEIRA ") & _
    "WHERE 1=1 " & _
      iif(Nivel > 0                       , "AND COD_REPRE = '" & CODUSER & "' "                          , "") & _
      iif(pCidade <> ""                   , "AND CIDADE = '" & pCidade & "' "                             , "") & _
      iif(pCodCli <> ""                   , "AND CASE WHEN CIDADE IN ('MARANGUAPE','BENTO') THEN CODIGO_ERP ELSE COD_CLI END = '" & pCodCli & "' ", "") & _
      iif(pLojaCli <> ""                  , "AND LOJA_CLI = '" & pLojaCli & "' "                          , "") & _
      iif(pPedido <> ""                   , "AND PEDIDO = '" & pPedido & "' "                             , "") & _
      iif(pRepre <> ""                    , "AND COD_REPRE = '" & pRepre & "' "                           , "") & _
      iif(pGrupoProduto = "PAPEL"         , "AND GRUPO_PRODUTO IN ('PAPEL','BOBINA') "                    , "") & _
      iif(pGrupoProduto = "PAPELÃO"       , "AND GRUPO_PRODUTO IN ('CAIXA','CHAPA','ACESSORIO','CONJUNTO') ", "") & _
      iif(pTipoProduto = "CAIXA"          , "AND TIPO_PRODUTO IN ('CORTE VINCO','MALETA') "               , "") & _
      iif(pTipoProduto = "CHAPA"          , "AND TIPO_PRODUTO IN ('CHAPA') "                              , "") & _
      iif(pTipoProduto = "CORTE VINCO"    , "AND TIPO_PRODUTO IN ('CORTE VINCO') "                        , "") & _
      iif(pTipoProduto = "MALETA"         , "AND TIPO_PRODUTO IN ('MALETA') "                             , "") & _
      iif(pTipoFrete = "CIF"              , "AND TIPO_FRETE = 'CIF' "                                     , "") & _
      iif(pTipoFrete = "FOB"              , "AND TIPO_FRETE = 'FOB' "                                     , "") & _
      iif(pPeriodoPor = "DTEMI" And pDataIni <> "" And pDataFim <> "", "AND CONVERT(varchar(8),EMISSAO,112) BETWEEN '" & pDataIni & "' AND '" & pDataFim & "' ", "") & _
      iif(pPeriodoPor = "DTENT" And pDataIni <> "" And pDataFim <> "", "AND CONVERT(varchar(8),DT_ENTREGA1,112) BETWEEN '" & pDataIni & "' AND '" & pDataFim & "' ", "") & _
      iif(pPeriodoPor = "DTEMB" And pDataIni <> "" And pDataFim <> "", "AND CONVERT(varchar(8),DATA_EMBARQUE,112) BETWEEN '" & pDataIni & "' AND '" & pDataFim & "' ", "") & _
    "ORDER BY DATA_EMBARQUE, COD_CLI"

  Set CONSULTA = Conecta("SPO-BI", SQL)

  somaValorTotal  = 0
  somaValorNet    = 0
  somaAreaVendida = 0
  somaPesoVendido = 0
  somaAreaSaldo   = 0
  somaPesoSaldo   = 0
  somaSaldoSemIPI = 0
  somaSaldoComIPI = 0
  somaVlrNetSaldo = 0
  totalRegistros  = 0

  Response.Write "{""sucesso"":true,""dados"":["

  primeiro = True

  If Not CONSULTA.EOF Then
    Do While Not CONSULTA.EOF

      somaValorTotal  = somaValorTotal  + CONSULTA("VALOR_TOTAL")
      somaValorNet    = somaValorNet    + CONSULTA("VALOR_NET")
      somaAreaVendida = somaAreaVendida + CONSULTA("AREA_VENDIDA")
      somaPesoVendido = somaPesoVendido + CONSULTA("PESO_VENDIDO")
      somaAreaSaldo   = somaAreaSaldo   + CONSULTA("AREA_SALDO")
      somaPesoSaldo   = somaPesoSaldo   + CONSULTA("PESO_SALDO")
      somaSaldoSemIPI = somaSaldoSemIPI + CONSULTA("VALOR_TOTAL_SALDO_SEM_IPI")
      somaSaldoComIPI = somaSaldoComIPI + CONSULTA("VALOR_TOTAL_SALDO_COM_IPI")
      somaVlrNetSaldo = somaVlrNetSaldo + CONSULTA("VALOR_NET_SALDO")
      totalRegistros  = totalRegistros + 1

      If Not primeiro Then Response.Write ","
      primeiro = False

      json = "{" & _
        """pedido"":"               & jsonStr(CONSULTA("PEDIDO")) & "," & _
        """tes"":"                  & jsonStr(CONSULTA("TES")) & "," & _
        """emissao"":"              & jsonStr(formataDataUStoBR(CONSULTA("EMISSAO"))) & "," & _
        """entrega"":"              & jsonStr(formataDataUStoBR(CONSULTA("DT_ENTREGA1"))) & "," & _
        """embarque"":"             & jsonStr(formataDataUStoBR(CONSULTA("DATA_EMBARQUE"))) & "," & _
        """codCli"":"               & jsonStr(Trim(CONSULTA("COD_CLI") & "")) & "," & _
        """lojaCli"":"              & jsonStr(CONSULTA("LOJA_CLI")) & "," & _
        """nomeCli"":"              & jsonStr(CONSULTA("NOME_REDUZIDO_CLI")) & "," & _
        """cidade"":"               & jsonStr(CONSULTA("CIDADE_CLIENTE") & " - " & CONSULTA("ESTADO_CLIENTE")) & "," & _
        """pedidoCliente"":"        & jsonStr(CONSULTA("PEDIDO_CLIENTE") & "-" & CONSULTA("ITEM_PEDIDO_CLIENTE")) & "," & _
        """tipoFrete"":"            & jsonStr(CONSULTA("TIPO_FRETE")) & "," & _
        """condPagamento"":"        & jsonStr(CONSULTA("CONDICAO_PAGAMENTO")) & "," & _
        """produto"":"              & jsonStr(CONSULTA("PRODUTO")) & "," & _
        """descricaoProduto"":"     & jsonStr(CONSULTA("DESCRICAO_PRODUTO")) & "," & _
        """estrutura"":"            & jsonStr(CONSULTA("ESTRUTURA")) & "," & _
        """grupoProduto"":"         & jsonStr(CONSULTA("GRUPO_PRODUTO")) & "," & _
        """tipoProduto"":"          & jsonStr(CONSULTA("TIPO_PRODUTO")) & "," & _
        """maquina"":"              & jsonStr(CONSULTA("MAQUINA")) & "," & _
        """codRepre"":"             & jsonStr(CONSULTA("COD_REPRE")) & "," & _
        """nomeRepre"":"            & jsonStr(CONSULTA("NOME_REDUZIDO_REP")) & "," & _
        """quantVendida"":"         & jsonNum(CONSULTA("QUANT_VENDIDA")) & "," & _
        """quantSaldoEntregar"":"   & jsonNum(CONSULTA("QUANT_SALDO_ENTREGAR")) & "," & _
        """valorUnitario"":"        & jsonNum(CONSULTA("VALOR_UNITARIO")) & "," & _
        """areaUnitaria"":"         & jsonNum(CONSULTA("AREA_UNITARIA")) & "," & _
        """pesoUnitario"":"         & jsonNum(CONSULTA("PESO_UNITARIO")) & "," & _
        """valorTotal"":"           & jsonNum(CONSULTA("VALOR_TOTAL")) & "," & _
        """valorNet"":"             & jsonNum(CONSULTA("VALOR_NET")) & "," & _
        """areaVendida"":"          & jsonNum(CONSULTA("AREA_VENDIDA")) & "," & _
        """pesoVendido"":"          & jsonNum(CONSULTA("PESO_VENDIDO")) & "," & _
        """areaSaldo"":"            & jsonNum(CONSULTA("AREA_SALDO")) & "," & _
        """pesoSaldo"":"            & jsonNum(CONSULTA("PESO_SALDO")) & "," & _
        """saldoSemIPI"":"          & jsonNum(CONSULTA("VALOR_TOTAL_SALDO_SEM_IPI")) & "," & _
        """saldoComIPI"":"          & jsonNum(CONSULTA("VALOR_TOTAL_SALDO_COM_IPI")) & "," & _
        """valorNetSaldo"":"        & jsonNum(CONSULTA("VALOR_NET_SALDO")) & "," & _
        """precoKgNet"":"           & jsonNum(CONSULTA("PRECO_KG_NET")) & "," & _
        """precoM2Net"":"           & jsonNum(CONSULTA("PRECO_M2_NET")) & _
      "}"

      Response.Write json
      CONSULTA.MoveNext()
    Loop
  End If

  CONSULTA.Close()
  Set CONSULTA = Nothing

  Response.Write "],""totais"":{" & _
    """valorTotal"":"     & jsonNum(somaValorTotal) & "," & _
    """valorNet"":"       & jsonNum(somaValorNet) & "," & _
    """areaVendida"":"    & jsonNum(somaAreaVendida) & "," & _
    """pesoVendido"":"    & jsonNum(somaPesoVendido) & "," & _
    """areaSaldo"":"      & jsonNum(somaAreaSaldo) & "," & _
    """pesoSaldo"":"      & jsonNum(somaPesoSaldo) & "," & _
    """saldoSemIPI"":"    & jsonNum(somaSaldoSemIPI) & "," & _
    """saldoComIPI"":"    & jsonNum(somaSaldoComIPI) & "," & _
    """valorNetSaldo"":"  & jsonNum(somaVlrNetSaldo) & _
  "},""totalRegistros"":" & totalRegistros & "}"

  Response.End
End Sub

Sub obterDetalhes()
  	pedido = sqlSafe(Request("pedido"))
  	cidade = sqlSafe(Request("cidade"))

	If pedido = "" Or cidade = "" Then
		Response.Write "{""sucesso"":false,""mensagem"":""Pedido ou cidade não informada.""}"
		Response.End
	End If

    SQL = SQL & "SELECT "
    SQL = SQL & "	PLANTA as Planta, "
    SQL = SQL & "	CIDADE as Cidade, "
    SQL = SQL & "	CONCAT(PEDIDO, '/', ITEM, '-', REPROGRAMACAO) OP, "
    SQL = SQL & "	PEDIDO AS Pedido, "
    SQL = SQL & "	ITEM AS Item, "
    SQL = SQL & "	REPROGRAMACAO AS Reprogramacao, "
    SQL = SQL & "	QUANT_PEDIDA AS QuantidadePedida, "
    SQL = SQL & "	QUANT_PROGRAMADA AS QuantidadeProgramada, "
    SQL = SQL & "	QUANT_ONDULADA AS QuantidadeOndulada, "
    SQL = SQL & "	QUANT_PRODUZIDA AS QuantidadeProduzida, "
    SQL = SQL & "	QUANT_APROVEITADA_ESTOQUE as QuantidadeAproveitadaEstoque,  "
    SQL = SQL & "	QUANT_EXPEDIDA as QuantidadeExpedida "
    SQL = SQL & "FROM BI_OFS "
    SQL = SQL & "WHERE PEDIDO = '" & pedido & "' "
    SQL = SQL & "AND Cidade ='" & cidade & "'"

    On Error Resume Next
    SET rsDetalhes = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDetalhes) Then
        Response.Write "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de chamados: " & EscapeJSON(Err.Description) & """}"
		Response.End
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDetalhes.EOF Then
        Do While Not rsDetalhes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """planta"": """ & EscapeJSON(rsDetalhes("Planta") & "") & """, "
			      jsonArray = jsonArray & """cidade"": """ & EscapeJSON(rsDetalhes("Cidade") & "") & """, "
            jsonArray = jsonArray & """OP"": """ & EscapeJSON(rsDetalhes("OP") & "") & """, "
            jsonArray = jsonArray & """pedido"": """ & EscapeJSON(rsDetalhes("Pedido") & "") & """, "
            jsonArray = jsonArray & """item"": """ & EscapeJSON(rsDetalhes("Item") & "") & """, "
            jsonArray = jsonArray & """reprogramacao"": """ & EscapeJSON(rsDetalhes("Reprogramacao") & "") & """, "
            jsonArray = jsonArray & """quantidadePedida"": """ & EscapeJSON(rsDetalhes("QuantidadePedida") & "") & """, "
            jsonArray = jsonArray & """quantidadeProgramada"": """ & EscapeJSON(rsDetalhes("QuantidadeProgramada") & "") & """, "
            jsonArray = jsonArray & """quantidadeOndulada"": """ & EscapeJSON(rsDetalhes("QuantidadeOndulada") & "") & """, "
            jsonArray = jsonArray & """quantidadeProduzida"": """ & EscapeJSON(rsDetalhes("QuantidadeProduzida") & "") & """, "
            jsonArray = jsonArray & """quantidadeAproveitadaEstoque"": """ & EscapeJSON(rsDetalhes("QuantidadeAproveitadaEstoque") & "") & """, "
			      jsonArray = jsonArray & """quantidadeExpedida"": """ & EscapeJSON(rsDetalhes("QuantidadeExpedida") & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDetalhes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDetalhes.Close()
    SET rsDetalhes = NOTHING

	Response.Write "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Sub

Sub validarCliente()

  Dim pCidade, pCodCli, pLojaCli
  Dim mapaUnidade, unidade, baseConexao
  Dim SQL, RST, nomeCliente

  pCidade  = sqlSafe(Request("planta"))
  pCodCli  = sqlSafe(Request("cliente"))
  pLojaCli = sqlSafe(Request("loja"))

  If pCodCli = "" Or pLojaCli = "" Then
    Response.Write "{""sucesso"":false,""mensagem"":""Informe os campos Cliente e Loja.""}"
    Response.End
  End If

  ' Mapa filial -> string de conexão Protheus
  ' (Mantém a mesma lógica usada no projeto SWnet — ajuste conforme padrão de cada planta)
  Select Case UCase(pCidade)
    Case "PIRAPETINGA" : baseConexao = "PIR-DADOSADV" : unidade = "00"
    Case "UBERABA"     : baseConexao = "UBA-DADOSADV" : unidade = "04"
    Case "BENTO"       : baseConexao = "BEN-DADOSADV" : unidade = "05"
    Case "MARANGUAPE"  : baseConexao = "MAR-DADOSADV" : unidade = "06"
    Case "SAQUAREMA"   : baseConexao = "SAQ-DADOSADV" : unidade = "07"
    Case Else
      Response.Write "{""sucesso"":false,""mensagem"":""Filial inválida para validação de cliente.""}"
      Response.End
  End Select

  SQL = "SELECT TOP 1 A1_NREDUZ FROM SA1010 " & _
        "WHERE A1_FILIAL = '" & unidade & "' " & _
          "AND A1_COD = '"    & pCodCli & "' " & _
          "AND A1_LOJA = '"   & pLojaCli & "' " & _
          "AND D_E_L_E_T_ <> '*'"

  Set RST = Conecta(baseConexao, SQL)
  If Not IsObject(RST) Then
    Response.Write "{""sucesso"":false,""mensagem"":""Falha ao consultar cliente.""}"
    Response.End
  End If

  If RST.EOF Then
    RST.Close()
    Set RST = Nothing
    Response.Write "{""sucesso"":false,""mensagem"":""Cliente não encontrado.""}"
    Response.End
  End If

  nomeCliente = Trim(RST("A1_NREDUZ") & "")
  RST.Close()
  Set RST = Nothing

  Response.Write "{""sucesso"":true,""nomeCliente"":" & jsonStr(nomeCliente) & "}"
  Response.End
End Sub
%>
