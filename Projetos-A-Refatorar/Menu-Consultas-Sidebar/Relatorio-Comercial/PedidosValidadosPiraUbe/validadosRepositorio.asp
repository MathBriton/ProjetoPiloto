<%@ Language="VBScript" %>
<!--#Include File="../../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../../cnn2/conexoes.asp"-->
<!--#Include File="../../../../includes/variaveis.asp"-->
<!--#Include File="../../../../funcoes/funcoes.asp"-->
<%
Response.ContentType = "application/json"
Response.Charset     = "UTF-8"

Session.LCID = 1046

' ── Sanitização básica ────────────────────────────────────────────────────────
Function S(val)
  S = Replace(Trim(val), "'", "''")
End Function

' ── EscapeJson: escapa caracteres especiais para JSON ─────────────────────────
Function EscapeJson(s)
  s = Replace(s, "\",  "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, Chr(8),  "\b")
  s = Replace(s, Chr(9),  "\t")
  s = Replace(s, Chr(10), "\n")
  s = Replace(s, Chr(12), "\f")
  s = Replace(s, Chr(13), "\r")
  EscapeJson = s
End Function

' ── FN: serializa número como JSON (troca vírgula pt-BR por ponto) ───────────
Function FN(val)
  If IsNull(val) Or IsEmpty(val) Then
    FN = "0"
  Else
    FN = Replace(CStr(CDbl(val)), ",", ".")
  End If
End Function

' ── Leitura dos parâmetros ────────────────────────────────────────────────────
Dim pTipo       : pTipo        = Trim(Request.QueryString("tipo"))
Dim pParam      : pParam       = Trim(Request.QueryString("param"))
Dim pPlanta     : pPlanta      = Trim(Request.QueryString("planta"))
Dim pCodcli     : pCodcli      = S(Request.QueryString("codcli"))
Dim pLojacli    : pLojacli     = S(Request.QueryString("lojacli"))
Dim pDataIni    : pDataIni     = Trim(Request.QueryString("datainicial"))
Dim pDataFin    : pDataFin     = Trim(Request.QueryString("datafinal"))
Dim pPedIni     : pPedIni      = S(Request.QueryString("pedidoInicial"))
Dim pPedFin     : pPedFin      = S(Request.QueryString("pedidoFinal"))
Dim pTipoDat    : pTipoDat     = Trim(Request.QueryString("tipoDat"))  ' "E" = Entrega | "M" = Emissão

' ── Conecta ao banco ──────────────────────────────────────────────────────────
If NOT ConectaDB("PIR-DADOSADV") Then
  Response.Write "{""sucesso"":false,""mensagem"":""Falha na conexão com o banco de dados.""}"
  Response.End
End If

' ── param <> "1": retorna envelope vazio ─────────────────────────────────────
If pParam <> "1" Then
  Response.Write "{""sucesso"":true,""dados"":[],""total"":0}"
  Response.End
End If

' ── Monta cláusulas de filtro ─────────────────────────────────────────────────
Dim filtroPlanta  : filtroPlanta  = ""
Dim filtroCodcli  : filtroCodcli  = ""
Dim filtroLojacli : filtroLojacli = ""
Dim filtroData    : filtroData    = ""
Dim filtroPedido  : filtroPedido  = ""
Dim filtroNivel   : filtroNivel   = ""

If pPlanta <> "" Then
  filtroPlanta = "AND SC5010.C5_INPUBE = '" & IIf(pPlanta = "00", "I", "U") & "' "
End If

If pCodcli <> "" Then
  filtroCodcli = "AND SC6010.C6_CLI = '" & pCodcli & "' "
End If

If pLojacli <> "" Then
  filtroLojacli = "AND SC6010.C6_LOJA = '" & pLojacli & "' "
End If

If pTipoDat = "M" Then
  ' Filtro por Emissão (ZF_EMISSAO na SZF010)
  If pDataIni <> "" And pDataFin = "" Then
    filtroData = "AND SZF010.ZF_EMISSAO = '" & formataDataYYYYMMDD(pDataIni) & "' "
  ElseIf pDataIni <> "" And pDataFin <> "" Then
    filtroData = "AND SZF010.ZF_EMISSAO BETWEEN '" & formataDataYYYYMMDD(pDataIni) & "' AND '" & formataDataYYYYMMDD(pDataFin) & "' "
  End If
Else
  ' Filtro por Entrega (padrão: C6_ENTREG)
  If pDataIni <> "" And pDataFin = "" Then
    filtroData = "AND SC6010.C6_ENTREG = '" & formataDataYYYYMMDD(pDataIni) & "' "
  ElseIf pDataIni <> "" And pDataFin <> "" Then
    filtroData = "AND SC6010.C6_ENTREG BETWEEN '" & formataDataYYYYMMDD(pDataIni) & "' AND '" & formataDataYYYYMMDD(pDataFin) & "' "
  End If
End If

If pPedIni <> "" And pPedFin = "" Then
  filtroPedido = "AND SC6010.C6_NUM = '" & pPedIni & "' "
ElseIf pPedIni <> "" And pPedFin <> "" Then
  filtroPedido = "AND SC6010.C6_NUM BETWEEN '" & pPedIni & "' AND '" & pPedFin & "' "
End If

If nivel = 1 Then
  filtroNivel = "AND SC5010.C5_VEND1 = '" & CODUSER & "' "
End If

' ── Monta SQL conforme tipo (D = Data / P = Pedido) ───────────────────────────
Dim SQL : SQL = ""

If pTipo = "D" Then

  SQL = "" &_
    "SELECT " &_
      "C6_ENTREG, " &_
      "A1_MUN, " &_
      "A1_EST, " &_
      "C6_NUM, " &_
      "C6_LOJA, " &_
      "A1_LOJA, " &_
      "C5_EMISSAO, " &_
      "C6_PEDCLI, " &_
      "C6_ITEMPC, " &_
      "C5_ESTATUS, " &_
      "C6_PRODUTO, " &_
      "C6_DESCRI, " &_
      "B1_DESC, " &_
      "DATAENTREG1, " &_
      "DATAENTREG2, " &_
      "C6_QTDVEN, " &_
      "C6_PRCVEN, " &_
      "C6_QTDENT, " &_
      "C6_LARG, " &_
      "C6_COMP, " &_
      "C6_VINC01, " &_
      "C6_VINC02, " &_
      "C6_VINC03, " &_
      "C6_VINC04, " &_
      "C6_VINC05, " &_
      "D2, " &_
      "C6_QTDORIG, " &_
      "C5_INPUBE, " &_
      "SB5010.B5_COD, " &_
      "ISNULL( dbo.funcCalculaArea( C6_PRODUTO, C6_NUM ), 0 ) * C6_QTDVEN AS AREA, " &_
      "ISNULL( dbo.funcCalculaPeso( C6_PRODUTO, C6_NUM, 'C' ), 0 ) AS PESO, " &_
      "C6_TES, " &_
      "CASE " &_
        "WHEN SUBSTRING(C6_PRODUTO,1,3) = '720' AND ISNULL(SB5010.B5_ESTRUTU,'') = '' THEN " &_
          "B1_ESTRUCH " &_
        "ELSE " &_
          "SB5010.B5_ESTRUTU " &_
      "END AS ESTRUTURA, " &_
      "A1_COD, " &_
      "A1_NREDUZ, " &_
      "INCLUIU, " &_
      "VALIDOU, " &_
      "CONDICAO_PAGAMENTO, " &_
      "VALOR_UNITARIO_NET, " &_
      "C5_LOTEPIL " &_
    "FROM (" &_
      "SELECT " &_
        "SC6010.C6_ENTREG, " &_
        "B1_ESTRUCH, " &_
        "SA1010.A1_MUN, " &_
        "SA1010.A1_EST, " &_
        "SA1010.A1_COD, " &_
        "SA1010.A1_LOJA, " &_
        "SA1010.A1_NREDUZ, " &_
        "SC6010.C6_NUM, " &_
        "SC6010.C6_LOJA, " &_
        "SC5010.C5_EMISSAO, " &_
        "SC6010.C6_PEDCLI, " &_
        "SC6010.C6_ITEMPC, " &_
        "SC5010.C5_ESTATUS, " &_
        "SC6010.C6_PRODUTO, " &_
        "SC6010.C6_DESCRI, " &_
        "SB1010.B1_DESC, " &_
        "SC6010.C6_TES, " &_
        "SUBSTRING(SC6010.C6_ENTREG, 7, 2) + '/' + SUBSTRING(SC6010.C6_ENTREG, 5, 2)+ '/' + SUBSTRING(SC6010.C6_ENTREG, 1, 4) AS DATAENTREG1, " &_
        "SUBSTRING(SC6010.C6_ENTREG2, 7, 2) + '/' + SUBSTRING(SC6010.C6_ENTREG2, 5, 2)+ '/' + SUBSTRING(SC6010.C6_ENTREG2, 1, 4) AS DATAENTREG2, " &_
        "CASE " &_
          "WHEN C6_UM <> 'KG' THEN " &_
            "SC6010.C6_QTDVEN " &_
          "ELSE " &_
            "SC6010.C6_QTDVEN / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) " &_
        "END C6_QTDVEN, " &_
        "CASE " &_
          "WHEN C6_UM <> 'KG' THEN " &_
            "SC6010.C6_QTDORIG " &_
          "ELSE " &_
            "SC6010.C6_QTDORIG / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) " &_
        "END C6_QTDORIG, " &_
        "CASE " &_
          "WHEN C6_UM <> 'KG' THEN " &_
            "SC6010.C6_QTDENT " &_
          "ELSE " &_
            "SC6010.C6_QTDENT / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) " &_
        "END C6_QTDENT, " &_
        "CASE " &_
          "WHEN C6_UM <> 'KG' THEN " &_
            "SC6010.C6_PRCVEN " &_
          "ELSE " &_
            "SC6010.C6_PRCVEN * (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) " &_
        "END C6_PRCVEN, " &_
        "SC6010.C6_LARG, " &_
        "SC6010.C6_COMP, " &_
        "C6_VINC01, " &_
        "C6_VINC02, " &_
        "C6_VINC03, " &_
        "C6_VINC04, " &_
        "C6_VINC05, " &_
        "SUBSTRING(SC6010.C6_DATFAT, 7, 2)+ '/' + SUBSTRING(SC6010.C6_DATFAT, 5, 2)+ '/' + SUBSTRING(SC6010.C6_DATFAT, 1, 4) AS D2, " &_
        "SZF010.ZF_DIGITAD AS INCLUIU, " &_
        "SC5010.C5_QUEMLIB AS VALIDOU, " &_
        "SC5010.C5_INPUBE, " &_
        "CASE " &_
          "WHEN SUBSTRING(C6_PRODUTO,1,3) = '750' THEN " &_
            "(SELECT MIN(G1_COMP) FROM SG1010 (NOLOCK) WHERE G1_COD = C6_PRODUTO AND SG1010.D_E_L_E_T_='') " &_
          "WHEN SUBSTRING(C6_PRODUTO,1,3) = '760' THEN " &_
            "(SELECT MIN(G1_COMP) FROM SG1010 (NOLOCK) WHERE G1_COD = C6_PRODUTO AND SG1010.D_E_L_E_T_='') " &_
          "WHEN SUBSTRING(C6_PRODUTO,1,3) = '720' THEN " &_
            "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
          "WHEN SUBSTRING(C6_PRODUTO,1,3) = '710' THEN " &_
            "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
          "WHEN SUBSTRING(C6_PRODUTO,1,3) BETWEEN '763' AND '769' THEN " &_
            "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
        "END AS B5_COD, " &_
        "E4_DESCRI AS CONDICAO_PAGAMENTO, " &_
        "C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100) - " &_
        "(C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100)) * (CASE F4_PISCOF WHEN '3' THEN 0.0165 ELSE 0 END) - " &_
        "(C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100)) * (CASE F4_PISCOF WHEN '3' THEN 0.0760 ELSE 0 END) AS VALOR_UNITARIO_NET, " &_
        "C5_LOTEPIL " &_
      "FROM SC6010 (NOLOCK) " &_
        "INNER JOIN SC5010 (NOLOCK) ON " &_
              "C5_NUM = C6_NUM " &_
          "AND SC5010.D_E_L_E_T_='' " &_
        "INNER JOIN SA1010 (NOLOCK) ON " &_
              "A1_COD = C6_CLI " &_
          "AND A1_LOJA = C6_LOJA " &_
          "AND SA1010.D_E_L_E_T_='' " &_
        "INNER JOIN SB1010 (NOLOCK) ON " &_
              "B1_COD = C6_PRODUTO " &_
          "AND SB1010.D_E_L_E_T_ = '' " &_
        "LEFT JOIN SE4010  (NOLOCK) ON " &_
              "C5_CONDPAG = E4_CODIGO " &_
          "AND SE4010.D_E_L_E_T_ = '' " &_
        "LEFT JOIN SF4010 (NOLOCK) ON " &_
              "C6_TES = F4_CODIGO " &_
          "AND SF4010.D_E_L_E_T_ = '' " &_
        "LEFT JOIN SZF010 (NOLOCK) ON " &_
              "C6_NUM = ZF_NUMERO " &_
          "AND SZF010.D_E_L_E_T_ = '' " &_
      "WHERE " &_
        "(SUBSTRING(SC6010.C6_PRODUTO, 1,1) = '7') AND " &_
        "(SC6010.D_E_L_E_T_ ='') AND " &_
        "(C5_NUM BETWEEN '000000' AND 'ZZ9999') " &_
        filtroPlanta &_
        filtroCodcli &_
        filtroLojacli &_
        filtroData &_
        filtroNivel &_
    ") AS CONSULTA " &_
    "LEFT JOIN SB5010 (NOLOCK) ON CONSULTA.B5_COD = SB5010.B5_COD " &_
    "ORDER BY " &_
      "A1_COD, " &_
      "A1_LOJA, " &_
      "C6_ENTREG, " &_
      "A1_MUN, " &_
      "A1_EST, " &_
      "C6_NUM, " &_
      "C6_PRODUTO, " &_
      "SB5010.B5_COD "

ElseIf pTipo = "P" Then

  SQL = "" &_
    "SELECT " &_
      "SC6010.C6_ENTREG, " &_
      "SB5010.B5_COD, " &_
      "SA1010.A1_MUN, " &_
      "SA1010.A1_EST, " &_
      "SA1010.A1_COD, " &_
      "SA1010.A1_NREDUZ, " &_
      "SC6010.C6_NUM, " &_
      "SC6010.C6_LOJA, " &_
      "SC5010.C5_EMISSAO, " &_
      "SC6010.C6_PEDCLI, " &_
      "SC6010.C6_ITEMPC, " &_
      "SC5010.C5_ESTATUS, " &_
      "SC6010.C6_PRODUTO, " &_
      "SC6010.C6_DESCRI, " &_
      "SB1010.B1_DESC, " &_
      "SC6010.C6_TES, " &_
      "SUBSTRING(SC6010.C6_ENTREG, 7, 2) + '/' + SUBSTRING(SC6010.C6_ENTREG, 5, 2)+ '/' + SUBSTRING(SC6010.C6_ENTREG, 1, 4) AS DATAENTREG1, " &_
      "SUBSTRING(SC6010.C6_ENTREG2, 7, 2) + '/' + SUBSTRING(SC6010.C6_ENTREG2, 5, 2)+ '/' + SUBSTRING(SC6010.C6_ENTREG2, 1, 4) AS DATAENTREG2, " &_
      "CASE WHEN C6_UM <> 'KG' THEN SC6010.C6_QTDVEN else SC6010.C6_QTDVEN / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) END C6_QTDVEN, " &_
      "CASE WHEN C6_UM <> 'KG' THEN SC6010.C6_PRCVEN else SC6010.C6_PRCVEN * (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) END C6_PRCVEN, " &_
      "CASE WHEN SUBSTRING(C6_PRODUTO,1,3) = '720' AND ISNULL(SB5010.B5_ESTRUTU,'') = '' THEN B1_ESTRUCH else SB5010.B5_ESTRUTU END AS ESTRUTURA, " &_
      "CASE WHEN C6_UM <> 'KG' THEN SC6010.C6_QTDENT else SC6010.C6_QTDENT / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) END C6_QTDENT, " &_
      "SC6010.C6_LARG, " &_
      "SC6010.C6_COMP, " &_
      "C6_VINC01, " &_
      "C6_VINC02, " &_
      "C6_VINC03, " &_
      "C6_VINC04, " &_
      "C6_VINC05, " &_
      "SUBSTRING(SC6010.C6_DATFAT, 7, 2)+ '/' + SUBSTRING(SC6010.C6_DATFAT, 5, 2)+ '/' + SUBSTRING(SC6010.C6_DATFAT, 1, 4) AS D2, " &_
      "SZF010.ZF_DIGITAD AS INCLUIU, " &_
      "SC5010.C5_QUEMLIB AS VALIDOU, " &_
      "CASE WHEN C6_UM <> 'KG' THEN SC6010.C6_QTDORIG else SC6010.C6_QTDORIG / (CASE C6_PESOUNI WHEN 0 THEN 1 else C6_PESOUNI END) END C6_QTDORIG, " &_
      "SC5010.C5_INPUBE, " &_
      "SB5010.B5_COD, " &_
      "ISNULL( dbo.funcCalculaArea( SC6010.C6_PRODUTO, SC6010.C6_NUM ), 0 ) * SC6010.C6_QTDVEN AS AREA, " &_
      "ISNULL( dbo.funcCalculaPeso( SC6010.C6_PRODUTO, SC6010.C6_NUM, 'C' ), 0 ) AS PESO, " &_
      "E4_DESCRI AS CONDICAO_PAGAMENTO, " &_
      "C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100) - " &_
      "(C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100)) * (CASE F4_PISCOF WHEN '3' THEN 0.0165 ELSE 0 END) - " &_
      "(C6_PRCVEN - (C6_PRCVEN * CONVERT(FLOAT,(CASE F4_ICM WHEN 'S' THEN dbo.FUNC_PERC_ICM_ESTADO(A1_EST) ELSE 0 END ))/100)) * (CASE F4_PISCOF WHEN '3' THEN 0.0760 ELSE 0 END) AS VALOR_UNITARIO_NET, " &_
      "C5_LOTEPIL " &_
    "FROM SC6010 (NOLOCK) " &_
      "INNER JOIN SC5010 (NOLOCK) ON " &_
              "C5_NUM = C6_NUM " &_
          "AND SC5010.D_E_L_E_T_='' " &_
          "AND SC5010.C5_NUM BETWEEN '000000' AND 'ZZ9999' " &_
      "INNER JOIN SA1010 (NOLOCK) ON " &_
              "A1_COD = C6_CLI " &_
          "AND A1_LOJA = C6_LOJA " &_
          "AND SA1010.D_E_L_E_T_='' " &_
          "AND C5_NUM BETWEEN '000000' AND 'ZZ9999' " &_
      "INNER JOIN SB1010 (NOLOCK) ON " &_
              "B1_COD = C6_PRODUTO " &_
          "AND SB1010.D_E_L_E_T_ = '' " &_
      "LEFT JOIN SB5010 (NOLOCK) ON " &_
              "B5_COD = ( " &_
                  "CASE " &_
                  "WHEN SUBSTRING(C6_PRODUTO,1,3) = '750' THEN " &_
                      "(SELECT MIN(G1_COMP) FROM SG1010 (NOLOCK) WHERE G1_COD = C6_PRODUTO AND SG1010.D_E_L_E_T_='') " &_
                  "WHEN SUBSTRING(C6_PRODUTO,1,3) = '760' THEN " &_
                      "(SELECT MIN(G1_COMP) FROM SG1010 (NOLOCK) WHERE G1_COD = C6_PRODUTO AND SG1010.D_E_L_E_T_='') " &_
                  "WHEN SUBSTRING(C6_PRODUTO,1,3) = '720' THEN " &_
                      "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
                  "WHEN SUBSTRING(C6_PRODUTO,1,3) = '710' THEN " &_
                      "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
                  "WHEN SUBSTRING(C6_PRODUTO,1,3) BETWEEN '763' AND '769' THEN " &_
                      "(SELECT B5C.B5_COD FROM SB5010 (NOLOCK) B5C WHERE B5C.B5_COD = C6_PRODUTO AND B5C.D_E_L_E_T_ = '') " &_
                  "END ) " &_
          "AND SB5010.D_E_L_E_T_='' " &_
      "LEFT JOIN SZF010 (NOLOCK) ON " &_
              "C6_NUM = ZF_NUMERO " &_
          "AND SZF010.D_E_L_E_T_ = '' " &_
      "LEFT JOIN SE4010 (NOLOCK) ON " &_
              "C5_CONDPAG = E4_CODIGO " &_
          "AND SE4010.D_E_L_E_T_ = '' " &_
      "LEFT JOIN SF4010 (NOLOCK) ON " &_
              "C6_TES = F4_CODIGO " &_
          "AND SF4010.D_E_L_E_T_ = '' " &_
    "WHERE (SC6010.D_E_L_E_T_ ='') " &_
      filtroCodcli &_
      filtroLojacli &_
      filtroPedido &_
      filtroNivel &_
    "ORDER BY " &_
      "SC6010.C6_NUM, " &_
      "SA1010.A1_MUN, " &_
      "SA1010.A1_EST, " &_
      "SC6010.C6_ENTREG, " &_
      "SC6010.C6_PRODUTO, " &_
      "SB5010.B5_COD "

Else
  ' tipo inválido ou ausente
  Response.Write "{""sucesso"":false,""mensagem"":""Parâmetro tipo inválido. Use D (data) ou P (pedido).""}"
  Response.End
End If

' ── Executa a query ───────────────────────────────────────────────────────────
Dim rs
On Error Resume Next
Set rs = Conecta("PIR-DADOSADV", SQL)
On Error GoTo 0

If Not IsObject(rs) Or (rs Is Nothing) Then
  Response.Write "{""sucesso"":false,""mensagem"":""Erro ao executar a consulta.""}"
  Response.End
End If

' ── Serializa o resultado em JSON ─────────────────────────────────────────────
Dim json    : json    = ""
Dim primeiro : primeiro = True
Dim total    : total    = 0

Dim CodigoAtual, CodigoOLD
CodigoOLD = ""

json = "{""sucesso"":true,""dados"":["

Do While Not rs.EOF

  CodigoAtual = rs("C6_NUM")

  ' Lógica de negócio original: agrupa por C6_NUM (ignora duplicatas de B5_COD)
  If CodigoAtual <> CodigoOLD Then

    ' ── Campos derivados (replicam lógica do validados.asp) ──────────────────
    Dim preco, peso, PKG, vKILO_NET

    preco = rs("C6_PRCVEN")
    peso  = rs("PESO")

    If CDbl(preco) > 0 And CDbl(peso) > 0 Then
      PKG = CDbl(preco) / CDbl(peso)
    Else
      PKG = 0
    End If

    If rs("VALOR_UNITARIO_NET") > 0 And rs("PESO") > 0 Then
      vKILO_NET = CDbl(rs("VALOR_UNITARIO_NET")) / CDbl(rs("PESO"))
    Else
      vKILO_NET = 0
    End If

    ' Descrição completa (igual ao validados.asp)
    Dim vLarg, vComp, vinco1, vinco2, vinco3, vinco4, vinco5, descricaoCompleta
    vLarg   = "L=" & rs("C6_LARG")   & ", "
    vComp   = "C=" & rs("C6_COMP")   & ", "
    vinco1  = "V1=" & rs("C6_VINC01") & ", "
    vinco2  = "V2=" & rs("C6_VINC02") & ", "
    vinco3  = "V3=" & rs("C6_VINC03") & ", "
    vinco4  = "V4=" & rs("C6_VINC04") & ", "
    vinco5  = "V5=" & rs("C6_VINC05")
    descricaoCompleta = Trim(rs("B1_DESC")) & " " & "(" & vLarg & vComp & vinco1 & vinco2 & vinco3 & vinco4 & vinco5 & ")"

    ' Fábrica
    Dim fabrica
    fabrica = IIf(Trim(rs("C5_INPUBE")) = "U", "UBERABA", "MATRIZ")

    ' Pedido do cliente
    Dim pedCli
    If Mid(rs("C6_PEDCLI"), 1, 6) <> Space(6) Then
      pedCli = Trim(rs("C6_PEDCLI")) & " - " & Trim(rs("C6_ITEMPC"))
    Else
      pedCli = ""
    End If

    ' Status
    Dim status
    If Mid(rs("C5_ESTATUS"), 1, 6) > Space(6) Then
      status = Trim(rs("C5_ESTATUS"))
    Else
      status = ""
    End If

    ' Estrutura
    Dim estrutura
    If IsNull(rs("ESTRUTURA")) Then
      estrutura = ""
    Else
      estrutura = Trim(rs("ESTRUTURA"))
    End If

    ' Lote piloto
    Dim lotePiloto
    lotePiloto = IIf(Trim(rs("C5_LOTEPIL")) = "S", "SIM", "NÃO")

    ' ── Monta objeto JSON ─────────────────────────────────────────────────────
    If Not primeiro Then json = json & ","
    json = json & "{"

    json = json & """C6_NUM"":"""           & EscapeJson(CStr(Trim(rs("C6_NUM"))))             & ""","
    json = json & """C5_EMISSAO"":"""       & EscapeJson(CStr(IIf(IsNull(rs("C5_EMISSAO")), "", ConverterData(rs("C5_EMISSAO"), "DD/MM/AAAA")))) & ""","
    json = json & """PED_CLI_ITEM"":"""     & EscapeJson(pedCli)                                & ""","
    json = json & """C5_ESTATUS"":"""       & EscapeJson(status)                                & ""","
    json = json & """CLIENTE"":"""          & EscapeJson(CStr(Trim(rs("A1_COD")) & " - " & Trim(rs("C6_LOJA")) & " - " & Trim(rs("A1_NREDUZ")))) & ""","
    json = json & """LOJA_MUN_EST"":"""     & EscapeJson(CStr(Trim(rs("C6_LOJA")) & "-" & Trim(rs("A1_MUN")) & "-" & Trim(rs("A1_EST")))) & ""","
    json = json & """C6_PRODUTO"":"""       & EscapeJson(CStr(Trim(rs("C6_PRODUTO"))))          & ""","
    json = json & """DESCRICAO"":"""        & EscapeJson(descricaoCompleta)                     & ""","
    json = json & """C6_TES"":"""           & EscapeJson(CStr(Trim(rs("C6_TES"))))              & ""","
    json = json & """CONDICAO_PAGAMENTO"":"
    If IsNull(rs("CONDICAO_PAGAMENTO")) Then
      json = json & "null,"
    Else
      json = json & """" & EscapeJson(CStr(Trim(rs("CONDICAO_PAGAMENTO")))) & ""","
    End If
    json = json & """DATAENTREG1"":"""      & EscapeJson(CStr(IIf(IsNull(rs("DATAENTREG1")), "", Trim(rs("DATAENTREG1"))))) & ""","
    json = json & """DATAENTREG2"":"""      & EscapeJson(CStr(IIf(IsNull(rs("DATAENTREG2")), "", Trim(rs("DATAENTREG2"))))) & ""","
    json = json & """C6_QTDVEN"":"          & FN(rs("C6_QTDVEN"))         & ","
    json = json & """C6_QTDORIG"":"         & FN(rs("C6_QTDORIG"))        & ","
    json = json & """C6_QTDENT"":"          & FN(rs("C6_QTDENT"))         & ","
    json = json & """C6_PRCVEN"":"          & FN(rs("C6_PRCVEN"))         & ","
    json = json & """PRECO_KG"":"           & FN(PKG)                     & ","
    json = json & """VALOR_UNITARIO_NET"":"  & FN(rs("VALOR_UNITARIO_NET")) & ","
    json = json & """KG_NET"":"             & FN(vKILO_NET)               & ","
    json = json & """AREA"":"               & FN(rs("AREA"))              & ","
    json = json & """PESO"":"               & FN(rs("PESO"))              & ","
    json = json & """SALDO"":"              & FN(CDbl(IIf(IsNull(rs("C6_QTDORIG")), 0, rs("C6_QTDORIG"))) - CDbl(IIf(IsNull(rs("C6_QTDENT")), 0, rs("C6_QTDENT")))) & ","
    json = json & """ESTRUTURA"":"""        & EscapeJson(estrutura)                                  & ""","
    json = json & """D2"":"""               & EscapeJson(CStr(IIf(IsNull(rs("D2")), "", Trim(rs("D2"))))) & ""","
    json = json & """FABRICA"":"""          & EscapeJson(fabrica)                                    & ""","
    json = json & """INCLUIU"":"""          & EscapeJson(CStr(IIf(IsNull(rs("INCLUIU")), "", Trim(rs("INCLUIU"))))) & ""","
    json = json & """VALIDOU"":"""          & EscapeJson(CStr(IIf(IsNull(rs("VALIDOU")), "", Trim(rs("VALIDOU"))))) & ""","
    json = json & """C5_LOTEPIL"":"""       & EscapeJson(lotePiloto)                                & ""","
    json = json & """C5_INPUBE"":"""        & EscapeJson(CStr(IIf(IsNull(rs("C5_INPUBE")), "", Trim(rs("C5_INPUBE"))))) & """"

    json = json & "}"

    primeiro    = False
    total       = total + 1
    CodigoOLD   = CodigoAtual

  End If

  rs.MoveNext
Loop

json = json & "],""total"":" & total & "}"

rs.Close
Set rs = Nothing

Response.Write json
Response.End
%>
