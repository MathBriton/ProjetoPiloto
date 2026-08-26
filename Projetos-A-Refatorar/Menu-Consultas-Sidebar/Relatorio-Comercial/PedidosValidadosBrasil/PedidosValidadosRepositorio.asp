<!--#include file="../../../cnn2/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp"-->
<%
Response.Buffer      = True
Response.ContentType = "application/json"
Response.Charset     = "UTF-8"
Response.Clear

Session.LCID = 1046

' ── Helper: escapa string para JSON ──────────────────────────────────────────
Function EscapeJson(v)
  Dim s, i
  s = CStr(v)
  s = Replace(s, "\",    "\\")
  s = Replace(s, """",   "\""")
  s = Replace(s, Chr(8),  "\b")
  s = Replace(s, Chr(9),  "\t")
  s = Replace(s, Chr(10), "\n")
  s = Replace(s, Chr(12), "\f")
  s = Replace(s, Chr(13), "\r")
  ' Remove qualquer outro caractere de controle (Chr 0-7, 11, 14-31)
  For i = 0 To 31
    If i <> 8 And i <> 9 And i <> 10 And i <> 12 And i <> 13 Then
      s = Replace(s, Chr(i), "")
    End If
  Next
  EscapeJson = s
End Function

' ── Helper: serializa string nullable ────────────────────────────────────────
Function strJson(v)
  If IsNull(v) Or IsEmpty(v) Then
    strJson = ""
  Else
    strJson = EscapeJson(CStr(v))
  End If
End Function

' ── Helper: serializa numero para JSON (remove formatacao pt-BR) ──────────────
Function numJson(v)
  Dim s
  If IsNull(v) Or IsEmpty(v) Then
    numJson = "0"
  Else
    s = CStr(v)
    s = Replace(s, ".", "")  ' remove separador de milhar pt-BR (ponto)
    s = Replace(s, ",", ".") ' troca virgula decimal por ponto JS
    numJson = s
  End If
End Function

' ── Helper: serializa data ISO (YYYY-MM-DD) ou null ──────────────────────────
Function dateJson(v)
  Dim d
  If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v)) = "" Then
    dateJson = "null"
  Else
    On Error Resume Next
    d = CDate(v)
    If Err.Number <> 0 Then
      Err.Clear
      dateJson = "null"
    Else
      dateJson = """" & Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & """"
    End If
    On Error GoTo 0
  End If
End Function

' ── Helper: envelope de erro ──────────────────────────────────────────────────
Function jsonErro(mensagem)
  jsonErro = "{""sucesso"":false,""dados"":[],""mensagem"":""" & EscapeJson(mensagem) & """}"
End Function

' ── Roteamento por action ────────────────────────────────────────────────────────────────────────────────
Dim vAction
vAction = Trim(Request("action"))

If vAction = "buscaProduto" Then

  Dim bpFilial, bpParametro, bpTpPedido, bpCodCli, bpLoja, bpEshop, bpParamSafe
  Dim bpStringConexao, bpSQL

  bpFilial    = Trim(Request("filial"))
  bpParametro = Trim(Request("parametro"))
  bpTpPedido  = UCase(Trim(Request("tppedido")))
  bpCodCli    = Replace(Trim("" & Request("codcli")), "'", "''")
  bpLoja      = Replace(Trim("" & Request("loja")),   "'", "''")
  bpEshop     = IIf(Trim(Request("eshop")) = "1", "1", "0")
  bpParamSafe = Replace(bpParametro, "'", "''")

  If bpFilial = "" Then
    Response.Write jsonErro("Informe a filial.")
    Response.End
  End If
  If bpParametro = "" Then
    Response.Write jsonErro("Informe o parametro de busca.")
    Response.End
  End If

  Select Case bpFilial

    Case "00", "04"
      bpStringConexao = "PIR-DADOSADV"

      If bpTpPedido <> "CX" Then
        Dim bpWhereCH, bpWherePA

        If bpTpPedido = "CH" Then
          Dim bpWhereCHCliente
          If bpCodCli <> "" Then
            bpWhereCHCliente = IIf(bpEshop = "0", "AND (ISNULL(B5_CLI,'') = '" & bpCodCli & "' OR ISNULL(B5_CLI,'') = '') ", _
                                                   "AND ISNULL(B5_CLI,'" & bpCodCli & "') IN ('9990','9099') ")
          Else
            bpWhereCHCliente = ""
          End If
          bpWhereCH = "AND B1_GRUPO = '72' " & _
            bpWhereCHCliente & _
            "AND (B1_COD = '" & bpParamSafe & "' OR UPPER(B1_DESC) LIKE UPPER('%" & bpParamSafe & "%') OR UPPER(B1_ESTRUCH) LIKE UPPER('%" & bpParamSafe & "%') OR UPPER(ISNULL(B1_REFERENCIA,'')) LIKE UPPER('%" & bpParamSafe & "%')) "
        Else
          bpWhereCH = ""
        End If

        If bpTpPedido = "PA" Or bpTpPedido = "M2" Then
          bpWherePA = "AND (B1_GRUPO = '70') " & _
            "AND (B1_COD = '" & bpParamSafe & "' OR UPPER(B1_DESC) LIKE UPPER('%" & bpParamSafe & "%') OR UPPER(ISNULL(B1_REFERENCIA,'')) LIKE UPPER('%" & bpParamSafe & "%')) "
        Else
          bpWherePA = ""
        End If

        Dim bpEstruturaCol
        If bpTpPedido = "CH" Or bpTpPedido = "PA" Or bpTpPedido = "M2" Then
          bpEstruturaCol = "CASE WHEN RTRIM(ISNULL(B5_ESTRUTU,'')) = '' THEN B1_ESTRUCH ELSE B5_ESTRUTU END AS ESTRUTURA, "
        Else
          bpEstruturaCol = "'' AS ESTRUTURA, "
        End If

        bpSQL = "SELECT DISTINCT " & _
          "0 GramaturaBobina, 0 LarguraBobina, 0 DiametroBobina, 0 TubeteBobina, " & _
          "B1_COD, ISNULL(B5_CLI,'') B5_CLI, " & _
          "CASE WHEN B1_INPUBE = 'U' THEN 'UBERABA' ELSE 'PIRAPETINGA' END AS B1_INPUBE, " & _
          "ISNULL(B5_LARGCH,'') B5_LARGCH, ISNULL(B5_COMPCH,'') B5_COMPCH, " & _
          "ISNULL(B5_VINCO1,'') B5_VINCO1, ISNULL(B5_VINCO2,'') B5_VINCO2, " & _
          "ISNULL(B5_VINCO3,'') B5_VINCO3, ISNULL(B5_VINCO4,'') B5_VINCO4, " & _
          "ISNULL(B5_VINCO5,'') B5_VINCO5, " & _
          "B1_DESC AS DESCRICAO, B1_DESC AS Referencia, " & _
          bpEstruturaCol & _
          "0 AS PRECO, 0 AS ESTOQUE, 0 AS PESO, '' AS TIPO_PRODUTO " & _
          "FROM SB1010 SB1 WITH(NOLOCK) " & _
          "LEFT JOIN SB5010 SB5 WITH(NOLOCK) ON B5_COD = B1_COD AND ISNULL(SB5.D_E_L_E_T_,'') = '' " & _
          "WHERE ISNULL(SB1.D_E_L_E_T_,'') = '' " & _
          "AND B1_STATUS IN ('V','A','W') AND B1_MSBLQL <> '1' " & _
          "AND SUBSTRING(ISNULL(B1_COD,''),1,2) <> '77' " & _
          bpWhereCH & bpWherePA & _
          "ORDER BY B1_INPUBE, B1_DESC"

      Else
        Dim bpWhereInpube
        bpWhereInpube = IIf(bpFilial = "00", "AND B1_INPUBE = 'I' ", "AND B1_INPUBE = 'U' ")

        ' ── Busca livre: sem cliente, remove filtro B5_CLI; com cliente, filtra normalmente
        Dim bpCXPreco, bpCXWhereCli
        If bpCodCli <> "" Then
          bpCXPreco    = "ISNULL((SELECT TOP 1 C6_PRCVEN FROM SC6010 WITH(NOLOCK) " & _
                         "WHERE C6_PRODUTO = B1_COD AND C6_CLI = '" & bpCodCli & "' AND D_E_L_E_T_ <> '*' ORDER BY C6_DTVALID DESC),0) AS PRECO, "
          bpCXWhereCli = IIf(bpEshop = "0", "AND SB5010.B5_CLI = '" & bpCodCli & "' ", "AND SB5010.B5_CLI IN ('9099','9990') ")
        Else
          bpCXPreco    = "0 AS PRECO, "
          bpCXWhereCli = ""
        End If

        bpSQL = "SELECT DISTINCT " & _
          "0 GramaturaBobina, 0 LarguraBobina, 0 DiametroBobina, 0 TubeteBobina, " & _
          "SB1010.B1_DESC AS DESCRICAO, SB1010.B1_DESC AS Referencia, " & _
          "ISNULL(B5_CLI,'') B5_CLI, SB1010.B1_COD, " & _
          "B5_ESTRUTU AS ESTRUTURA, " & _
          bpCXPreco & _
          "CASE WHEN B1_INPUBE = 'U' THEN 'UBERABA' ELSE 'PIRAPETINGA' END AS B1_INPUBE, " & _
          "'' B5_LARGCH, '' B5_COMPCH, '' B5_VINCO1, '' B5_VINCO2, " & _
          "'' B5_VINCO3, '' B5_VINCO4, '' B5_VINCO5, " & _
          "0 AS ESTOQUE, 0 AS PESO, '' AS TIPO_PRODUTO " & _
          "FROM SB5010 WITH(NOLOCK) " & _
          "INNER JOIN SB1010 WITH(NOLOCK) ON SB5010.B5_COD = SB1010.B1_COD AND ISNULL(SB5010.D_E_L_E_T_,'') = '' " & _
          "WHERE (UPPER(SB1010.B1_DESC) LIKE UPPER('%" & bpParamSafe & "%') OR B1_COD = '" & bpParamSafe & "' OR UPPER(ISNULL(SB1010.B1_REFERENCIA,'')) LIKE UPPER('%" & bpParamSafe & "%')) " & _
          bpWhereInpube & _
          "AND SB1010.D_E_L_E_T_ <> '*' " & _
          "AND SB5010.B5_TIPO IN ('C','D','S','T') " & _
          "AND SB1010.B1_STATUS IN ('V','A','W') " & _
          "AND SB1010.B1_GRUPO IN ('71','76') AND B1_SUBGRUP NOT IN ('750','760') " & _
          bpCXWhereCli & _
          "AND SB1010.B1_MSBLQL <> '1' " & _
          "AND SUBSTRING(ISNULL(SB1010.B1_COD,''),1,2) <> '77' " & _
          "ORDER BY SB1010.B1_DESC"
      End If

    Case "05", "06", "07"
      If bpFilial = "05" Then
        bpStringConexao = IIf(bpTpPedido = "PA", "BEN-DADOSADV", "BEN-TRIMBOX")
      ElseIf bpFilial = "06" Then
        bpStringConexao = "FOR-TRIMBOX"
      Else
        bpStringConexao = "SAQ-TRIMBOX"
      End If

      ' ── Busca livre: sem cliente no Else, abre para todos
      Dim bpFrom, bpWhere
      If bpTpPedido = "PA" Then
        bpFrom  = "FROM VW_PAPEL_PEDIDO "
        bpWhere = "WHERE 1 = 1 "
      ElseIf (bpFilial = "05" Or bpFilial = "07") And bpTpPedido = "CH" Then
        bpFrom  = "FROM VW_BI_PRODUTOS_PEDIDO "
        bpWhere = "WHERE GRUPO_PRODUTO = 'CHAPA' " & IIf(bpFilial = "07", "AND TIPO_CHAPA = 'GENERICA' ", "")
      ElseIf bpCodCli <> "" Then
        bpFrom  = "FROM VW_BI_PRODUTOS_PEDIDO "
        bpWhere = "WHERE CODIGO_CLIENTE = '" & bpCodCli & "' AND LOJA_CLIENTE = '" & bpLoja & "' "
      Else
        bpFrom  = "FROM VW_BI_PRODUTOS_PEDIDO "
        bpWhere = "WHERE 1 = 1 "
      End If

      Dim bpPaCol
      bpPaCol = IIf(bpTpPedido = "PA", "GramaturaBobina, LarguraBobina, DiametroBobina, TubeteBobina, ", "0 AS GramaturaBobina, 0 AS LarguraBobina, 0 AS DiametroBobina, 0 AS TubeteBobina, ")

      bpSQL = "SELECT " & _
          "CASE WHEN TIPO = 'FT' THEN CODIGO_CONJUNTO ELSE CODIGO_ITEM END AS B1_COD, " & _
          "REFERENCIA_ITEM AS DESCRICAO, REFERENCIA_ITEM AS Referencia, " & _
          "ESTRUTURA, " & _
          "LARGURA_CHAPA     AS B5_LARGCH, COMPRIMENTO_CHAPA AS B5_COMPCH, " & _
          "VINCO1 AS B5_VINCO1, VINCO2 AS B5_VINCO2, VINCO3 AS B5_VINCO3, " & _
          "VINCO4 AS B5_VINCO4, VINCO5 AS B5_VINCO5, 0 AS B5_TPPALE, PRECO, " & _
          bpPaCol & _
          "ESTOQUE, PESO, TIPO_PRODUTO, " & _
          "CODIGO_CLIENTE AS B5_CLI, CIDADE AS B1_INPUBE " & _
          bpFrom & bpWhere & _
          "AND STATUS = 'ATIVO' " & _
          "AND (REFERENCIA_ITEM LIKE UPPER('%" & bpParamSafe & "%') " & _
          "OR CASE WHEN TIPO = 'FT' THEN CODIGO_CONJUNTO ELSE CODIGO_ITEM END LIKE '%" & bpParamSafe & "%') "

    Case Else
      Response.Write jsonErro("Filial invalida: " & bpFilial)
      Response.End
  End Select

  ' ── Executa e serializa ──────────────────────────────────────────────────────────────────────────────
  Dim bpConnStr
  bpConnStr = ""
  On Error Resume Next
  bpConnStr = BuscaStringConexao(UCase(bpStringConexao))
  Err.Clear : On Error GoTo 0

  If Trim(bpConnStr) = "" Then
    Response.Write jsonErro("String de conexao nao encontrada: " & bpStringConexao)
    Response.End
  End If

  Dim bpCnn, bpCmd, bpRS
  On Error Resume Next
  Set bpCnn = Server.CreateObject("ADODB.Connection")
  bpCnn.CommandTimeout = 120
  bpCnn.Open bpConnStr
  If Err.Number <> 0 Then
    Dim bpErrConn : bpErrConn = Err.Description
    Err.Clear : On Error GoTo 0
    Response.Write jsonErro("Erro ao conectar: " & EscapeJson(bpErrConn))
    Response.End
  End If

  Set bpCmd = Server.CreateObject("ADODB.Command")
  bpCmd.ActiveConnection = bpCnn
  bpCmd.CommandType      = 1
  bpCmd.CommandTimeout   = 120
  bpCmd.CommandText      = bpSQL
  Set bpRS = bpCmd.Execute
  If Err.Number <> 0 Then
    Dim bpErrQ : bpErrQ = Err.Description
    Err.Clear : On Error GoTo 0
    bpCnn.Close : Set bpCnn = Nothing
    Response.Write jsonErro("Erro na consulta: " & EscapeJson(bpErrQ))
    Response.End
  End If
  On Error GoTo 0

  Dim bpJson, bpPrimeiro
  bpJson     = "["
  bpPrimeiro = True

  Do While Not bpRS.EOF
    Dim bpTemEst, bpTemRef
    bpTemEst = (strJson(bpRS("ESTRUTURA")) <> "" Or bpTpPedido = "PA")
    bpTemRef = (Trim(strJson(bpRS("Referencia"))) <> "")

    If Not bpPrimeiro Then bpJson = bpJson & ","
    bpPrimeiro = False

    bpJson = bpJson & "{" & _
      """codigo"":"     & """" & strJson(bpRS("B1_COD"))      & """," & _
      """descricao"":"  & """" & strJson(bpRS("DESCRICAO"))   & """," & _
      """referencia"":" & """" & strJson(bpRS("Referencia"))  & """," & _
      """estrutura"":"  & """" & strJson(bpRS("ESTRUTURA"))   & """," & _
      """filial"":"     & """" & strJson(bpRS("B1_INPUBE"))   & """," & _
      """cliente"":"    & """" & strJson(bpRS("B5_CLI"))      & """," & _
      """largCh"":"     & """" & strJson(bpRS("B5_LARGCH"))   & """," & _
      """compCh"":"     & """" & strJson(bpRS("B5_COMPCH"))   & """," & _
      """vinco1"":"     & """" & strJson(bpRS("B5_VINCO1"))   & """," & _
      """vinco2"":"     & """" & strJson(bpRS("B5_VINCO2"))   & """," & _
      """vinco3"":"     & """" & strJson(bpRS("B5_VINCO3"))   & """," & _
      """vinco4"":"     & """" & strJson(bpRS("B5_VINCO4"))   & """," & _
      """vinco5"":"     & """" & strJson(bpRS("B5_VINCO5"))   & """," & _
      """gramatura"":"  & numJson(bpRS("GramaturaBobina"))    & "," & _
      """largBobina"":" & numJson(bpRS("LarguraBobina"))      & "," & _
      """diamBobina"":" & numJson(bpRS("DiametroBobina"))     & "," & _
      """tubete"":"     & numJson(bpRS("TubeteBobina"))       & "," & _
      """preco"":"      & numJson(bpRS("PRECO"))              & "," & _
      """estoque"":"    & numJson(bpRS("ESTOQUE"))            & "," & _
      """peso"":"       & numJson(bpRS("PESO"))               & "," & _
      """tipoProduto"":" & """" & strJson(bpRS("TIPO_PRODUTO")) & """," & _
      """disponivel"":"  & IIf(bpTemEst And bpTemRef, "true", "false") & _
      "}"
    bpRS.MoveNext()
  Loop

  bpJson = bpJson & "]"
  bpRS.Close  : Set bpRS  = Nothing
  bpCnn.Close : Set bpCnn = Nothing

  Response.Write "{""sucesso"":true,""dados"":" & bpJson & ",""mensagem"":""""}"
  Response.End

End If
' ── /action buscaProduto ───────────────────────────────────────────────────────────────────────────────

' ── Leitura e sanitizacao dos parametros ──────────────────────────────────────
Dim vCboFilial, vCboProduto, vCodProduto, vNomeProduto, vCboTipo
Dim vDataIni1, vDataFin1, vDataIni2, vDataFin2
Dim vCodCli, vPedidoFiltro, cbxLotePiloto
Dim NomePlanta, vStringConexao

vCboFilial    = Trim(Request("cboFilial"))
vCboProduto   = Trim(Request("cboProduto"))
vCodProduto   = UCase(Replace(Trim("" & Request("codProduto")),   "'", "''"))
vNomeProduto  = UCase(Replace(Trim("" & Request("nomeProduto")), "'", "''"))

' Remove sufixo de dimensoes " (L=...)" concatenado pelo frontend em CHAPAs
Dim posL
posL = InStr(1, vNomeProduto, " (L=", 1)
If posL > 0 Then vNomeProduto = Left(vNomeProduto, posL - 1)
vCboTipo      = Trim(Request("cboTipo"))
vDataIni1     = Replace(Trim("" & Request("dataInicial1")), "-", "")
vDataFin1     = Replace(Trim("" & Request("dataFinal1")),   "-", "")
vDataIni2     = Replace(Trim("" & Request("dataInicial2")), "-", "")
vDataFin2     = Replace(Trim("" & Request("dataFinal2")),   "-", "")
vCodCli       = Replace(Trim("" & Request("codcli")),       "'", "''")
vPedidoFiltro = UCase(Replace(Trim("" & Request("pedido")), "'", "''"))
cbxLotePiloto = IIf(Request("lotepiloto") = "1", "1", "0")

' ── Validacoes obrigatorias ────────────────────────────────────────────────────
If vCboFilial = "" Then
  Response.Write jsonErro("Informe a filial.")
  Response.End
End If

If vCboTipo = "" Then
  Response.Write jsonErro("Informe o tipo de filtro.")
  Response.End
End If

If vDataIni1 = "" Or vDataFin1 = "" Then
  Response.Write jsonErro("Informe o periodo de datas.")
  Response.End
End If

If vCboTipo = "EE" And (vDataIni2 = "" Or vDataFin2 = "") Then
  Response.Write jsonErro("Informe o periodo de entrega.")
  Response.End
End If

' ── Resolve nome da planta e string de conexao ────────────────────────────────
Select Case vCboFilial
  Case "05"
    NomePlanta     = "BENTO"
    vStringConexao = "BEN-CORPORATIVO"
  Case "06"
    NomePlanta     = "MARANGUAPE"
    vStringConexao = "FOR-CORPORATIVO"
  Case "07"
    NomePlanta     = "SAQUAREMA"
    vStringConexao = "SAQ-CORPORATIVO"
  Case Else
    Response.Write jsonErro("Filial invalida.")
    Response.End
End Select

' ── Monta a query (logica preservada integralmente do original) ───────────────
Dim SQL
SQL = "" & _
    "SELECT " & _
      "Pedido, " & _
      "LotePiloto, " & _
      "Cliente, " & _
      "Razao AS Razao_social_cliente, " & _
      "NumeroOrdemCompra, " & _
      "ItemOrdemCompra, " & _
      "Status AS Situacao_Pedido, " & _
      "StatusPedido, " & _
      "Produto, " & _
      "Referencia, " & _
      "Largura, " & _
      "Comprimento, " & _
      "TESFatura, " & _
      "Quantidade, " & _
      "isnull(Qtde_Faturada, 0) Qtde_Faturada, " & _
      "Area_Unitaria AS Area, " & _
      "(Quantidade * Area_Unitaria) AS area_total, " & _
      "PrecoUnitario, " & _
      "PrecoAtual, " & _
      "Preco_Net, " & _
      "Preco_KG_Net, " & _
      "PrimeiraDataEntrega, " & _
      "DataEmissao, " & _
      "QuemIncluiu AS QuemInc, " & _
      "DataValidacao, " & _
      "QuemValidou AS QuemVal, " & _
      "GRUPO_PRODUTO, " & _
      "(Quantidade - Qtde_Faturada) AS Saldo_Pedido " & _
    "FROM VW_CONSULTA_PEDIDOS_VALIDADOS " & _
    "LEFT JOIN OPENQUERY (BRSPA1M006, 'SELECT " & _
                                        "NUMERO_PEDIDO, " & _
                                        "CIDADE, " & _
                                        "SUM(QUANTIDADE) Qtde_Faturada " & _
                                      "FROM BI.dbo.BI_NOTAS " & _
                                      "WHERE CIDADE = ''" & NomePlanta & "'' " & _
                                        "AND EMISSAO >= ''" & vDataIni1 & "'' " & _
                                        "GROUP BY NUMERO_PEDIDO, CIDADE, EMISSAO') AS NF ON " & _
          "NF.NUMERO_PEDIDO = Pedido COLLATE SQL_Latin1_General_CP1_CI_AS " & _
    "WHERE 1=1 " & _
      IIf(vCboProduto = "CX", "AND GRUPO_PRODUTO LIKE '%CAIXA' ", "") & _
      IIf(vCboProduto = "CH", "AND GRUPO_PRODUTO LIKE '%CHAPA' ", "") & _
      IIf(vCboProduto = "PA", "AND GRUPO_PRODUTO LIKE '%PAPEL' ", "") & _
      IIf(vCodProduto <> "",  "AND Produto = '" & vCodProduto & "' ", "") & _
      IIf(vNomeProduto <> "" And vCodProduto = "", "AND (UPPER(Referencia) LIKE UPPER('%" & vNomeProduto & "%') OR UPPER(Produto) LIKE UPPER('%" & vNomeProduto & "%')) ", "") & _
      IIf(vCboTipo = "EM",  "AND CONVERT(varchar(8),DataEmissao,112)         BETWEEN '" & vDataIni1 & "' AND '" & vDataFin1 & "' ", "") & _
      IIf(vCboTipo = "VL",  "AND CONVERT(varchar(8),DataValidacao,112)       BETWEEN '" & vDataIni1 & "' AND '" & vDataFin1 & "' ", "") & _
      IIf(vCboTipo = "EN",  "AND CONVERT(varchar(8),PrimeiraDataEntrega,112) BETWEEN '" & vDataIni1 & "' AND '" & vDataFin1 & "' ", "") & _
      IIf(vCboTipo = "EE",  "AND ( CONVERT(varchar(8),DataEmissao,112) BETWEEN '" & vDataIni1 & "' AND '" & vDataFin1 & "' " & _
                                "AND CONVERT(varchar(8),PrimeiraDataEntrega,112) BETWEEN '" & vDataIni2 & "' AND '" & vDataFin2 & "' ) ", "") & _
      IIf(Nivel > 0,          "AND Representante = '"  & CODUSER        & "' ", "") & _
      IIf(vCodCli       <> "", "AND Cliente = '"        & vCodCli        & "' ", "") & _
      IIf(vPedidoFiltro <> "", "AND Pedido = '"         & vPedidoFiltro  & "' ", "") & _
      IIf(cbxLotePiloto = "1", "AND LotePiloto = '1' ", "") & _
      IIf(vCboTipo = "EM",  "ORDER BY DataEmissao, Cliente ",                       "") & _
      IIf(vCboTipo = "VL",  "ORDER BY DataValidacao, Cliente ",                     "") & _
      IIf(vCboTipo = "EN",  "ORDER BY PrimeiraDataEntrega, Cliente ",               "") & _
      IIf(vCboTipo = "EE",  "ORDER BY DataEmissao, PrimeiraDataEntrega, Cliente ",  "")

' ── Executa a consulta via ADODB direto (sem Conecta do cnn2) ────────────────
' Motivo: Conecta() chama Response.Write + Response.End ao falhar,
' contaminando a resposta JSON mesmo com Response.Buffer = True.
' Aqui tratamos o erro nós mesmos e emitimos o envelope JSON de erro.
Dim oCnn, oCmd, CONSULTA
Dim sStringConexao

sStringConexao = ""
On Error Resume Next
sStringConexao = BuscaStringConexao(UCase(vStringConexao))
Err.Clear
On Error GoTo 0

If Trim(sStringConexao) = "" Then
  Response.Clear
  Response.ContentType = "application/json"
  Response.Charset     = "UTF-8"
  Response.Write jsonErro("String de conexao nao encontrada: " & vStringConexao)
  Response.End
End If

On Error Resume Next

Set oCnn = Server.CreateObject("ADODB.Connection")
oCnn.CommandTimeout = 300
oCnn.Open sStringConexao

If Err.Number <> 0 Then
  Dim eCnnDesc : eCnnDesc = Err.Description
  Err.Clear
  On Error GoTo 0
  Response.Clear
  Response.ContentType = "application/json"
  Response.Charset     = "UTF-8"
  Response.Write jsonErro("Erro ao conectar ao banco de dados: " & EscapeJson(eCnnDesc))
  Response.End
End If

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = oCnn
oCmd.CommandType      = 1  ' adCmdText
oCmd.CommandTimeout   = 300
oCmd.CommandText      = SQL

Set CONSULTA = oCmd.Execute

If Err.Number <> 0 Then
  Dim eQryDesc : eQryDesc = Err.Description
  Err.Clear
  On Error GoTo 0
  oCnn.Close
  Set oCnn = Nothing
  Response.Clear
  Response.ContentType = "application/json"
  Response.Charset     = "UTF-8"
  Response.Write jsonErro("Erro ao executar a consulta: " & EscapeJson(eQryDesc))
  Response.End
End If

On Error GoTo 0

If CONSULTA Is Nothing Then
  oCnn.Close
  Set oCnn = Nothing
  Response.Clear
  Response.ContentType = "application/json"
  Response.Charset     = "UTF-8"
  Response.Write jsonErro("A consulta nao retornou um Recordset valido.")
  Response.End
End If

' ── Serializa o JSON ──────────────────────────────────────────────────────────
Dim jsonDados, primeiro
Dim fPedido, fCliente, fRazao, fProduto, fReferencia, fGrupoProd
Dim fLargura, fComprimento, fTES, fQtde, fQtdeFat, fSaldo
Dim fPrecoInc, fPrecoAtual, fPrecoKgNet
Dim fQuemInc, fQuemVal, fLotePiloto
Dim fOC, fItemOC, fSitPedido, fStatusPedido, fAreaUnit, fAreaTotal
Dim status, vReferencia, vPedCli, vLinkPedido

jsonDados = "["
primeiro  = True

Do While Not CONSULTA.EOF

  fPedido      = "" & CONSULTA("Pedido")
  fCliente     = "" & CONSULTA("Cliente")
  fRazao       = "" & CONSULTA("Razao_social_cliente")
  fProduto     = "" & CONSULTA("Produto")
  fReferencia  = "" & CONSULTA("Referencia")
  fGrupoProd   = "" & CONSULTA("GRUPO_PRODUTO")
  fLargura     = "" & CONSULTA("Largura")
  fComprimento = "" & CONSULTA("Comprimento")
  fTES         = "" & CONSULTA("TESFatura")
  fQtde        = CONSULTA("Quantidade")
  fQtdeFat     = CONSULTA("Qtde_Faturada")
  fSaldo       = CONSULTA("Saldo_Pedido")
  fPrecoInc    = CONSULTA("PrecoUnitario")
  fPrecoAtual  = CONSULTA("PrecoAtual")
  fPrecoKgNet  = CONSULTA("Preco_KG_Net")
  fQuemInc     = "" & CONSULTA("QuemInc")
  fQuemVal     = "" & CONSULTA("QuemVal")
  fLotePiloto  = "" & CONSULTA("LotePiloto")
  fOC          = "" & CONSULTA("NumeroOrdemCompra")
  fItemOC      = "" & CONSULTA("ItemOrdemCompra")
  fSitPedido   = "" & CONSULTA("Situacao_Pedido")
  fStatusPedido = "" & CONSULTA("StatusPedido")
  fAreaUnit    = CONSULTA("Area")
  fAreaTotal   = CONSULTA("area_total")

  ' Traducao do status (regra de negocio preservada)
  Select Case fSitPedido
    Case "A" : status = "Atendido"
    Case "C" : status = "Cancelado"
    Case "N" : status = "Normal"
    Case Else: status = fSitPedido
  End Select

  ' Referencia com dimensoes para CHAPA (regra de negocio preservada)
  vReferencia = fReferencia
  If (InStr(1, fReferencia, "CHAPA", 1) > 0) Or (InStr(1, fGrupoProd, "CHAPA", 1) > 0) Then
    vReferencia = fReferencia & " (L=" & fLargura & ", C=" & fComprimento & ")"
  End If

  ' Pedido do cliente (OC + Item)
  vPedCli = IIf(fOC = "", "", fOC & "-" & fItemOC)

  ' Link de detalhamento
  vLinkPedido = "PedidosValidadosDetalhamento.asp?Planta=" & vCboFilial & _
                "&Pedido=" & fPedido & _
                "&PrecoKGnet=" & Replace(FormatNumber(CONSULTA("Preco_KG_Net"), 2), ".", "")

  If Not primeiro Then jsonDados = jsonDados & ","
  primeiro = False

  jsonDados = jsonDados & "{" & _
    """pedido"":"       & """" & strJson(fPedido)      & """," & _
    """linkPedido"":"   & """" & strJson(vLinkPedido)  & """," & _
    """cliente"":"      & """" & strJson(fCliente)     & """," & _
    """razao"":"        & """" & strJson(fRazao)       & """," & _
    """pedCli"":"       & """" & strJson(vPedCli)      & """," & _
    """produto"":"      & """" & strJson(fProduto)     & """," & _
    """referencia"":"   & """" & strJson(vReferencia)  & """," & _
    """grupoProduto"":""" & strJson(fGrupoProd)        & """," & _
    """tes"":"          & """" & strJson(fTES)         & """," & _
    """qtde"":"         & numJson(fQtde)               & ","   & _
    """qtdeFat"":"      & numJson(fQtdeFat)            & ","   & _
    """saldo"":"        & numJson(fSaldo)              & ","   & _
    """areaUnit"":"     & numJson(fAreaUnit)           & ","   & _
    """areaTotal"":"    & numJson(fAreaTotal)          & ","   & _
    """precoInc"":"     & numJson(fPrecoInc)           & ","   & _
    """precoAtual"":"   & numJson(fPrecoAtual)         & ","   & _
    """precoKgNet"":"   & numJson(fPrecoKgNet)         & ","   & _
    """entrega"":"      & dateJson(CONSULTA("PrimeiraDataEntrega")) & "," & _
    """emissao"":"      & dateJson(CONSULTA("DataEmissao"))        & "," & _
    """validacao"":"    & dateJson(CONSULTA("DataValidacao"))      & "," & _
    """status"":"       & """" & strJson(status)       & """," & _
    """statusPedido"":" & """" & strJson(fStatusPedido) & """," & _
    """emitiu"":"       & """" & strJson(fQuemInc)     & """," & _
    """validou"":"      & """" & strJson(fQuemVal)     & """," & _
    """lotePiloto"":"   & """" & strJson(fLotePiloto)  & """" & _
    "}"

  CONSULTA.MoveNext()
Loop

jsonDados = jsonDados & "]"

CONSULTA.Close()
Set CONSULTA = Nothing
oCnn.Close()
Set oCnn = Nothing

' ── Emite envelope final ──────────────────────────────────────────────────────
Response.Write "{""sucesso"":true,""dados"":" & jsonDados & ",""mensagem"":""""}"
Response.End
%>
