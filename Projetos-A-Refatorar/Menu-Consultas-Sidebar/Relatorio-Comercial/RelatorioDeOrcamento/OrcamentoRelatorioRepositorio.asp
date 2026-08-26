<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Buffer      = True
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID         = 1046
Response.ContentType = "application/json"

' ── Contexto do usuário logado ──────────────────────────────────────────────
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id",    Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome",  Session("NOME")

' ── Dispatcher ──────────────────────────────────────────────────────────────
Dim acao : acao = Trim(Request("acao"))

If acao = "" Then
    Response.Write("{""sucesso"":false,""mensagem"":""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listar"       : resultado = ListarOrcamentos()
    Case "filtros"      : resultado = ObterFiltros()
    Case "logAprovacao" : resultado = ObterLogAprovacao()
    Case "buscaCliente" : resultado = BuscarCliente()
    Case "buscaProduto" : resultado = BuscarProduto()
    Case Else           : resultado = "{""sucesso"":false,""mensagem"":""Ação inválida: " & EscapeJSON(acao) & """}"
End Select

Response.Clear
Response.Write resultado

' ═══════════════════════════════════════════════════════════════════════════
' LISTAR ORÇAMENTOS
' ═══════════════════════════════════════════════════════════════════════════
Function ListarOrcamentos()
    On Error Resume Next

    Dim filtroPlanta, filtroStatus, filtroCliente
    Dim filtroDataDe, filtroDataAte, filtroProduto, filtroNrOrcamento
    Dim cmd, rs, SQL

    filtroPlanta      = Trim(Request("planta"))
    filtroNrOrcamento = Trim(Request("nrOrcamento"))
    filtroStatus      = Trim(Request("status"))
    filtroCliente     = Trim(Request("cliente"))  ' legado — mantido para compatibilidade
    filtroCodCli      = Trim(Request("codcli"))   ' código exato (modal ou digitação)
    filtroCodLoja     = Trim(Request("codloja"))  ' loja exata (opcional — vazio = todas as lojas)
    filtroDataDe      = Trim(Request("dataDe"))
    filtroDataAte     = Trim(Request("dataAte"))
    filtroProduto     = Trim(Request("produto"))

    ' ── Conexão via ADODB.Command (permite parâmetros tipados) ──────────────
    CriaConexao(1)
    Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
    conn.CommandTimeout = 180
    conn.Open Session("conexaoString")

    If Err.Number <> 0 Or conn Is Nothing Then
        ListarOrcamentos = "{""sucesso"":false,""mensagem"":""Erro na conexão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    SQL = "SET NOCOUNT ON; " & _
          ";WITH CondPag AS ( " & _
          "    SELECT CP.CODIGO, CP.CIDADE, CP.DESCRICAO " & _
          "    FROM BRSPA1M006.BI.dbo.BI_COND_PAG_TRIMBOX CP " & _
          "    WHERE CP.CIDADE IN ('PIRAPETINGA','UBERABA','BENTO','MARANGUAPE','SAQUAREMA') " & _
          ") " & _
          "SELECT " & _
          "    CASE O.PLANTA " & _
          "        WHEN '00' THEN 'PIRAPETINGA' " & _
          "        WHEN '04' THEN 'UBERABA' " & _
          "        WHEN '05' THEN 'BENTO' " & _
          "        WHEN '06' THEN 'MARANGUAPE' " & _
          "        WHEN '07' THEN 'SAQUAREMA' " & _
          "    END AS PLANTA, " & _
          "    O.NR_ORCAMENTO, " & _
          "    CONCAT(O.COD_CLIENTE, ' - ', O.LOJA_CLIENTE, '   ', O.NOME_CLIENTE) AS CLIENTE, " & _
          "    O.CNPJ, " & _
          "    O.CIDADE, " & _
          "    O.UF, " & _
          "    CONCAT(O.COD_REPRESENTANTE, ' - ', O.NOME_REPRESENTANTE) AS REPRESENTANTE, " & _
          "    O.TIPO_FRETE, " & _
          "    CP.DESCRICAO, " & _
          "    CONVERT(VARCHAR(10), O.VALIDADE_DE,    103) AS VALIDADE_DE, " & _
          "    CONVERT(VARCHAR(10), O.VALIDADE_ATE,   103) AS VALIDADE_ATE, " & _
          "    O.PRODUTO, " & _
          "    O.REFERENCIA, " & _
          "    O.VALOR_UNITARIO, " & _
          "    O.STATUS_ORCAMENTO, " & _
          "    O.MARGEM, " & _
          "    CONVERT(VARCHAR(10), O.DATA_INCLUSAO, 103) + ' ' + CONVERT(VARCHAR(5), O.DATA_INCLUSAO, 108) AS DATA_INCLUSAO " & _
          "FROM ORCAMENTO_BRASIL O " & _
          "LEFT JOIN CondPag CP " & _
          "    ON  CP.CODIGO = O.CONDICAO_PAGAMENTO COLLATE Latin1_General_CI_AS " & _
          "    AND CP.CIDADE = CASE O.PLANTA " & _
          "                        WHEN '00' THEN 'PIRAPETINGA' " & _
          "                        WHEN '04' THEN 'UBERABA' " & _
          "                        WHEN '05' THEN 'BENTO' " & _
          "                        WHEN '06' THEN 'MARANGUAPE' " & _
          "                        WHEN '07' THEN 'SAQUAREMA' " & _
          "                    END " & _
          "WHERE 1=1 "

    If filtroNrOrcamento <> "" Then SQL = SQL & " AND O.NR_ORCAMENTO LIKE ?"
    If filtroPlanta      <> "" Then SQL = SQL & " AND O.PLANTA = ?"
    If filtroStatus      <> "" Then SQL = SQL & " AND O.STATUS_ORCAMENTO = ?"

    ' Filtro de cliente — prioridade: codcli (exato) > cliente (LIKE)
    If filtroCodCli <> "" Then
        SQL = SQL & " AND LTRIM(RTRIM(O.COD_CLIENTE)) = ?"
        If filtroCodLoja <> "" Then SQL = SQL & " AND LTRIM(RTRIM(O.LOJA_CLIENTE)) = ?"
    ElseIf filtroCliente <> "" Then
        SQL = SQL & " AND (O.NOME_CLIENTE COLLATE Latin1_General_CI_AS LIKE ? OR O.COD_CLIENTE COLLATE Latin1_General_CI_AS LIKE ?)"
        If filtroCodLoja <> "" Then SQL = SQL & " AND LTRIM(RTRIM(O.LOJA_CLIENTE)) = ?"
    End If

    If filtroDataDe <> "" And filtroDataAte <> "" Then
        SQL = SQL & " AND O.DATA_INCLUSAO >= ? AND O.DATA_INCLUSAO < DATEADD(DAY, 1, ?)"
    ElseIf filtroDataDe <> "" Then
        SQL = SQL & " AND O.DATA_INCLUSAO >= ?"
    ElseIf filtroDataAte <> "" Then
        SQL = SQL & " AND O.DATA_INCLUSAO < DATEADD(DAY, 1, ?)"
    End If

    If filtroProduto <> "" Then SQL = SQL & " AND (O.PRODUTO LIKE ? OR O.REFERENCIA LIKE ?)"

    SQL = SQL & " ORDER BY O.NR_ORCAMENTO DESC"

    Const adVarChar    = 200
    Const adDate       = 7
    Const adParamInput = 1
    Const adCmdText    = 1

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText    = SQL
    cmd.CommandType    = adCmdText
    cmd.CommandTimeout = 180

    If filtroNrOrcamento <> "" Then cmd.Parameters.Append cmd.CreateParameter("pNrOrc",  adVarChar, adParamInput, 100, "%" & filtroNrOrcamento & "%")
    If filtroPlanta      <> "" Then cmd.Parameters.Append cmd.CreateParameter("pPlanta", adVarChar, adParamInput, 10,  filtroPlanta)
    If filtroStatus      <> "" Then cmd.Parameters.Append cmd.CreateParameter("pStatus", adVarChar, adParamInput, 100, filtroStatus)

    ' Bind do cliente — espelha a lógica do WHERE acima
    If filtroCodCli <> "" Then
        ' Normaliza: remove espaços; loja com zero à esquerda se tiver 1 dígito
        Dim pCodCliNorm : pCodCliNorm = Trim(filtroCodCli)
        cmd.Parameters.Append cmd.CreateParameter("pCodCli", adVarChar, adParamInput, 20, pCodCliNorm)
        If filtroCodLoja <> "" Then
            Dim pCodLojaNorm : pCodLojaNorm = Trim(filtroCodLoja)
            If Len(pCodLojaNorm) = 1 Then pCodLojaNorm = "0" & pCodLojaNorm
            cmd.Parameters.Append cmd.CreateParameter("pCodLoja", adVarChar, adParamInput, 10, pCodLojaNorm)
        End If
    ElseIf filtroCliente <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("pCliNome", adVarChar, adParamInput, 200, "%" & filtroCliente & "%")
        cmd.Parameters.Append cmd.CreateParameter("pCliCod",  adVarChar, adParamInput, 200, "%" & filtroCliente & "%")
        If filtroCodLoja <> "" Then
            Dim pLojaLivre : pLojaLivre = Trim(filtroCodLoja)
            If Len(pLojaLivre) = 1 Then pLojaLivre = "0" & pLojaLivre
            cmd.Parameters.Append cmd.CreateParameter("pCliLoja", adVarChar, adParamInput, 10, pLojaLivre)
        End If
    End If

    If filtroDataDe <> "" And filtroDataAte <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("pDataDe",  adDate, adParamInput, , CDate(filtroDataDe))
        cmd.Parameters.Append cmd.CreateParameter("pDataAte", adDate, adParamInput, , CDate(filtroDataAte))
    ElseIf filtroDataDe <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("pDataDe",  adDate, adParamInput, , CDate(filtroDataDe))
    ElseIf filtroDataAte <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("pDataAte", adDate, adParamInput, , CDate(filtroDataAte))
    End If

    If filtroProduto <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("pProd", adVarChar, adParamInput, 200, "%" & filtroProduto & "%")
        cmd.Parameters.Append cmd.CreateParameter("pRef",  adVarChar, adParamInput, 200, "%" & filtroProduto & "%")
    End If

    On Error Resume Next
    Set rs = cmd.Execute

    If Err.Number <> 0 Or Not IsObject(rs) Then
        Dim errListar : errListar = Err.Description
        conn.Close : Set cmd = Nothing : Set conn = Nothing
        ListarOrcamentos = "{""sucesso"":false,""mensagem"":""Erro na consulta: " & EscapeJSON(errListar) & """}"
        Exit Function
    End If
    On Error GoTo 0

    ' ── Serialização: flag primeiroItem conforme padrão do skill ────────────
    Dim jsonArray, primeiroItem
    jsonArray    = "["
    primeiroItem = True

    Do While Not rs.EOF
        If Not primeiroItem Then jsonArray = jsonArray & ","
        primeiroItem = False

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """planta"":"""         & EscapeJSON(rs("PLANTA")           & "") & ""","
        jsonArray = jsonArray & """nrOrcamento"":"""    & EscapeJSON(rs("NR_ORCAMENTO")      & "") & ""","
        jsonArray = jsonArray & """cliente"":"""        & EscapeJSON(rs("CLIENTE")           & "") & ""","
        jsonArray = jsonArray & """cnpj"":"""           & EscapeJSON(rs("CNPJ")              & "") & ""","
        jsonArray = jsonArray & """cidade"":"""         & EscapeJSON(rs("CIDADE")            & "") & ""","
        jsonArray = jsonArray & """uf"":"""             & EscapeJSON(rs("UF")                & "") & ""","
        jsonArray = jsonArray & """representante"":"""  & EscapeJSON(rs("REPRESENTANTE")     & "") & ""","
        jsonArray = jsonArray & """tipoFrete"":"""      & EscapeJSON(rs("TIPO_FRETE")        & "") & ""","
        jsonArray = jsonArray & """descricao"":"""      & EscapeJSON(rs("DESCRICAO")         & "") & ""","
        jsonArray = jsonArray & """validadeDe"":"""     & EscapeJSON(rs("VALIDADE_DE")       & "") & ""","
        jsonArray = jsonArray & """validadeAte"":"""    & EscapeJSON(rs("VALIDADE_ATE")      & "") & ""","
        jsonArray = jsonArray & """produto"":"""        & EscapeJSON(rs("PRODUTO")           & "") & ""","
        jsonArray = jsonArray & """referencia"":"""     & EscapeJSON(rs("REFERENCIA")        & "") & ""","
        jsonArray = jsonArray & """valorUnitario"":"    & FormatarNumeroJSON(rs("VALOR_UNITARIO"))  & ","
        jsonArray = jsonArray & """statusOrcamento"":"""& EscapeJSON(rs("STATUS_ORCAMENTO")  & "") & ""","
        jsonArray = jsonArray & """margem"":"           & FormatarNumeroJSON(rs("MARGEM"))           & ","
        jsonArray = jsonArray & """dataInclusao"":"""   & EscapeJSON(rs("DATA_INCLUSAO")     & "") & """"
        jsonArray = jsonArray & "}"

        rs.MoveNext
    Loop

    jsonArray = jsonArray & "]"

    rs.Close  : Set rs   = Nothing
    conn.Close: Set cmd  = Nothing : Set conn = Nothing

    ListarOrcamentos = "{""sucesso"":true,""mensagem"":"""",""dados"":" & jsonArray & "}"
End Function

' ═══════════════════════════════════════════════════════════════════════════
' OBTER FILTROS (plantas + status)
' ═══════════════════════════════════════════════════════════════════════════
Function ObterFiltros()
    On Error Resume Next

    ' ── Conexão pela sessão (mesma base de ORCAMENTO_BRASIL) ────────────────
    CriaConexao(1)
    Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
    conn.CommandTimeout = 180
    conn.Open Session("conexaoString")

    If Err.Number <> 0 Or conn Is Nothing Then
        ObterFiltros = "{""sucesso"":false,""mensagem"":""Erro na conexão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Const adCmdText_F    = 1
    Const adParamInput_F = 1

    ' ── Plantas ─────────────────────────────────────────────────────────────
    Dim cmdF, rsPlantas, jsonPlantas, primPlantas
    jsonPlantas = "["
    primPlantas = True

    Set cmdF = Server.CreateObject("ADODB.Command")
    Set cmdF.ActiveConnection = conn
    cmdF.CommandType    = adCmdText_F
    cmdF.CommandTimeout = 180
    cmdF.CommandText    = _
        "SELECT DISTINCT PLANTA, " & _
        "CASE PLANTA " & _
        "    WHEN '00' THEN 'PIRAPETINGA' " & _
        "    WHEN '04' THEN 'UBERABA' " & _
        "    WHEN '05' THEN 'BENTO' " & _
        "    WHEN '06' THEN 'MARANGUAPE' " & _
        "    WHEN '07' THEN 'SAQUAREMA' " & _
        "    ELSE PLANTA " & _
        "END AS PLANTA_NOME " & _
        "FROM ORCAMENTO_BRASIL " & _
        "WHERE PLANTA IS NOT NULL " & _
        "ORDER BY PLANTA_NOME"

    On Error Resume Next
    Set rsPlantas = cmdF.Execute

    If Err.Number <> 0 Or Not IsObject(rsPlantas) Then
        Dim errPlantas : errPlantas = Err.Description
        conn.Close : Set cmdF = Nothing : Set conn = Nothing
        ObterFiltros = "{""sucesso"":false,""mensagem"":""Erro ao buscar plantas: " & EscapeJSON(errPlantas) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Do While Not rsPlantas.EOF
        If Not primPlantas Then jsonPlantas = jsonPlantas & ","
        primPlantas = False
        jsonPlantas = jsonPlantas & "{"
        jsonPlantas = jsonPlantas & """codigo"":""" & EscapeJSON(rsPlantas("PLANTA")      & "") & ""","
        jsonPlantas = jsonPlantas & """nome"":"""   & EscapeJSON(rsPlantas("PLANTA_NOME") & "") & """"
        jsonPlantas = jsonPlantas & "}"
        rsPlantas.MoveNext
    Loop
    jsonPlantas = jsonPlantas & "]"
    rsPlantas.Close : Set rsPlantas = Nothing : Set cmdF = Nothing

    ' ── Status ──────────────────────────────────────────────────────────────
    Dim cmdS, rsStatus, jsonStatus, primStatus
    jsonStatus = "["
    primStatus = True

    Set cmdS = Server.CreateObject("ADODB.Command")
    Set cmdS.ActiveConnection = conn
    cmdS.CommandType    = adCmdText_F
    cmdS.CommandTimeout = 180
    cmdS.CommandText    = _
        "SELECT DISTINCT STATUS_ORCAMENTO " & _
        "FROM ORCAMENTO_BRASIL " & _
        "WHERE STATUS_ORCAMENTO IS NOT NULL " & _
        "  AND STATUS_ORCAMENTO <> '' " & _
        "  AND UPPER(STATUS_ORCAMENTO) NOT IN ('APROVADO PARCIAL') " & _
        "ORDER BY STATUS_ORCAMENTO"

    On Error Resume Next
    Set rsStatus = cmdS.Execute

    If Err.Number <> 0 Or Not IsObject(rsStatus) Then
        Dim errStatus : errStatus = Err.Description
        conn.Close : Set cmdS = Nothing : Set conn = Nothing
        ObterFiltros = "{""sucesso"":false,""mensagem"":""Erro ao buscar status: " & EscapeJSON(errStatus) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Do While Not rsStatus.EOF
        If Not primStatus Then jsonStatus = jsonStatus & ","
        primStatus = False
        jsonStatus = jsonStatus & """" & EscapeJSON(rsStatus("STATUS_ORCAMENTO") & "") & """"
        rsStatus.MoveNext
    Loop
    jsonStatus = jsonStatus & "]"
    rsStatus.Close : Set rsStatus = Nothing : Set cmdS = Nothing

    conn.Close : Set conn = Nothing

    ObterFiltros = "{""sucesso"":true,""mensagem"":"""",""dados"":{""plantas"":" & jsonPlantas & ",""status"":" & jsonStatus & "}}"
End Function

' ═══════════════════════════════════════════════════════════════════════════
' OBTER LOG DE APROVAÇÃO
' ═══════════════════════════════════════════════════════════════════════════
Function ObterLogAprovacao()
    On Error Resume Next

    Dim nrOrcamento, idsOrcamento, flagMigracao
    Dim SQL, rs, cmd

    nrOrcamento = Trim(Request("nrOrcamento"))

    If nrOrcamento = "" Then
        ObterLogAprovacao = "{""sucesso"":false,""mensagem"":""Nº do Orçamento não informado""}"
        Exit Function
    End If

    ' ── Validação: nrOrcamento deve ser numérico (evita SQL Injection) ──────
    If Not IsNumeric(nrOrcamento) Then
        ObterLogAprovacao = "{""sucesso"":false,""mensagem"":""Nº do Orçamento inválido""}"
        Exit Function
    End If

    CriaConexao(1)
    Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
    conn.CommandTimeout = 180
    conn.Open Session("conexaoString")

    If Err.Number <> 0 Or conn Is Nothing Then
        ObterLogAprovacao = "{""sucesso"":false,""mensagem"":""Erro na conexão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    ' ── Buscar IDs e flag de migração ────────────────────────────────────────
    Const adVarChar_L    = 200
    Const adParamInput_L = 1
    Const adCmdText_L    = 1
    Const adInteger_L    = 3

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType    = adCmdText_L
    cmd.CommandTimeout = 180
    cmd.CommandText    = "SELECT ID_ORCAMENTO, FLAG_MIGRACAO_APROVACAO FROM ORCAMENTO_BRASIL WHERE NR_ORCAMENTO = ?"
    cmd.Parameters.Append cmd.CreateParameter("pNrOrc", adInteger_L, adParamInput_L, , CLng(nrOrcamento))

    On Error Resume Next
    Set rs = cmd.Execute

    If Err.Number <> 0 Or Not IsObject(rs) Then
        Dim errLog1 : errLog1 = Err.Description
        conn.Close : Set cmd = Nothing : Set conn = Nothing
        ObterLogAprovacao = "{""sucesso"":false,""mensagem"":""Erro ao buscar orçamento: " & EscapeJSON(errLog1) & """}"
        Exit Function
    End If
    On Error GoTo 0

    idsOrcamento = ""
    flagMigracao = False

    Do While Not rs.EOF
        idsOrcamento = idsOrcamento & CStr(rs("ID_ORCAMENTO") & "") & ","
        flagMigracao = rs("FLAG_MIGRACAO_APROVACAO")
        rs.MoveNext
    Loop
    rs.Close : Set rs = Nothing : Set cmd = Nothing

    If Len(idsOrcamento) > 0 Then idsOrcamento = Left(idsOrcamento, Len(idsOrcamento) - 1)

    ' ── Buscar log de aprovação ──────────────────────────────────────────────
    SQL = "SELECT DISTINCT " & _
          "    CASE AO.PLANTA " & _
          "        WHEN '00' THEN 'PIRAPETINGA' " & _
          "        WHEN '04' THEN 'UBERABA' " & _
          "        WHEN '05' THEN 'BENTO' " & _
          "        WHEN '06' THEN 'MARANGUAPE' " & _
          "        WHEN '07' THEN 'SAQUAREMA' " & _
          "        ELSE ISNULL(AO.PLANTA, '-') " & _
          "    END AS PLANTA_NOME, " & _
          "    CASE WHEN ISNULL(AO.ID_APROVADOR, 0) = 0 THEN 'FALTA APROVAR' ELSE ISNULL(U.CRM_NOME, 'FALTA APROVAR') END AS NOME_APROVADOR, " & _
          "    ISNULL(AO.NIVEL, 0) AS NIVEL, " & _
          "    ISNULL(CONVERT(VARCHAR(10), AO.DATA_APROVACAO, 103) + ' ' + CONVERT(VARCHAR(5), AO.DATA_APROVACAO, 108), '-') AS DATA_APROVACAO, " & _
          "    ISNULL(CONVERT(VARCHAR(8000), AO.JUSTIFICATIVA), '') AS JUSTIFICATIVA " & _
          "FROM APROVACAO_ORCAMENTO AO " & _
          "LEFT JOIN HELPDESK..CRM_DTI U ON U.CRM_R_E_C_N_O = AO.ID_APROVADOR "

    ' NR_ORCAMENTO já foi validado como numérico acima — concatenação segura
    If flagMigracao Then
        SQL = SQL & "WHERE AO.NR_ORCAMENTO = " & CLng(nrOrcamento) & " "
    Else
        If idsOrcamento = "" Then
            conn.Close : Set conn = Nothing
            ObterLogAprovacao = "{""sucesso"":true,""mensagem"":"""",""dados"":{""log"":[]}}"
            Exit Function
        End If
        SQL = SQL & "WHERE AO.ID_ORCAMENTO IN (" & idsOrcamento & ") "
    End If

    SQL = SQL & "AND AO.EXCLUIDO IS NULL ORDER BY NIVEL"

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType    = adCmdText_L
    cmd.CommandTimeout = 180
    cmd.CommandText    = SQL

    On Error Resume Next
    Set rs = cmd.Execute

    If Err.Number <> 0 Or Not IsObject(rs) Then
        Dim errLog2 : errLog2 = Err.Description
        conn.Close : Set cmd = Nothing : Set conn = Nothing
        ObterLogAprovacao = "{""sucesso"":false,""mensagem"":""Erro ao buscar log: " & EscapeJSON(errLog2) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonLog, primLog
    jsonLog = "["
    primLog = True

    Do While Not rs.EOF
        If Not primLog Then jsonLog = jsonLog & ","
        primLog = False

        jsonLog = jsonLog & "{"
        jsonLog = jsonLog & """planta"":"""       & EscapeJSON(rs("PLANTA_NOME")    & "") & ""","
        jsonLog = jsonLog & """aprovador"":"""    & EscapeJSON(rs("NOME_APROVADOR") & "") & ""","
        jsonLog = jsonLog & """nivel"":"""        & EscapeJSON(rs("NIVEL")          & "") & ""","
        jsonLog = jsonLog & """data"":"""         & EscapeJSON(rs("DATA_APROVACAO") & "") & ""","
        jsonLog = jsonLog & """justificativa"":"""& EscapeJSON(rs("JUSTIFICATIVA")  & "") & """"
        jsonLog = jsonLog & "}"

        rs.MoveNext
    Loop
    jsonLog = jsonLog & "]"

    rs.Close  : Set rs  = Nothing
    conn.Close: Set cmd = Nothing : Set conn = Nothing

    ObterLogAprovacao = "{""sucesso"":true,""mensagem"":"""",""dados"":{""log"":" & jsonLog & "}}"
End Function

' ═══════════════════════════════════════════════════════════════════════════
' BUSCA CLIENTE — SPA Modal
' Parâmetros: parametro (nome/código), planta (código 00|04|05|06|07)
' ═══════════════════════════════════════════════════════════════════════════
Function BuscarCliente()
    On Error Resume Next

    Dim parametro, planta, SQL, rs

    parametro = Trim(Request("parametro"))
    planta    = Trim(Request("planta"))

    If parametro = "" Then
        BuscarCliente = "{""sucesso"":false,""mensagem"":""Parâmetro de busca não informado""}"
        Exit Function
    End If

    ' ── SQL por planta ───────────────────────────────────────────────────────
    Select Case planta

        Case "00", "04"   ' PIRAPETINGA / UBERABA
            Dim pUpper00 : pUpper00 = Replace(UCase(parametro), "'", "''")
            Dim pExato00 : pExato00 = Replace(parametro, "'", "''")
            SQL = "SELECT TOP 100 A1_COD, A1_LOJA, A1_NOME, A1_NREDUZ, A1_MUN, A1_EST, A1_BAIRRO, A1_CGC AS CNPJ, A1_CODSAP, A1_MSBLQL " & _
                  "FROM SA1010 WITH(NOLOCK) " & _
                  "WHERE D_E_L_E_T_ = '' " & _
                  "  AND A1_MSBLQL <> '1' " & _
                  "  AND (UPPER(A1_NREDUZ) LIKE '%" & pUpper00 & "%' " & _
                  "    OR UPPER(A1_NOME)   LIKE '%" & pUpper00 & "%' " & _
                  "    OR A1_COD           =    '"  & pExato00 & "') " & _
                  "ORDER BY A1_NREDUZ"

        Case "05"   ' BENTO
            Dim sqlBentoCod05 : sqlBentoCod05 = ""
            Dim pUpper05 : pUpper05 = Replace(UCase(parametro), "'", "''")
            If IsNumeric(parametro) Then sqlBentoCod05 = " OR (T1.CODIGO_ERP = " & CLng(parametro) & ")"
            SQL = "SELECT TOP 100 T1.CODIGO_ERP AS A1_COD, T1.LOJA_CLIENTE AS A1_LOJA, T1.NOME_CLIENTE AS A1_NOME, " & _
                  "       T1.FANTASIA_CLIENTE AS A1_NREDUZ, T1.MUNICIPIO_CLIENTE AS A1_MUN, T1.ESTADO_CLIENTE AS A1_EST, " & _
                  "       T1.A1_BAIRRO, T1.CNPJ_CLIENTE AS CNPJ, '' AS A1_CODSAP, '2' AS A1_MSBLQL " & _
                  "FROM VW_BI_CLIENTES T1 WITH(NOLOCK) " & _
                  "WHERE ((T1.FANTASIA_CLIENTE LIKE '%" & pUpper05 & "%') " & _
                  "    OR (T1.NOME_CLIENTE     LIKE '%" & pUpper05 & "%') " & _
                  sqlBentoCod05 & _
                  ") AND T1.STATUS = 'ATIVO' " & _
                  "ORDER BY T1.FANTASIA_CLIENTE"

        Case "06"   ' MARANGUAPE — VW_CLIENTES_PEDIDO retorna COD_SAP_CLIENTE real
            Dim sqlMarCod06 : sqlMarCod06 = ""
            Dim pUpper06 : pUpper06 = Replace(UCase(parametro), "'", "''")
            If IsNumeric(parametro) Then sqlMarCod06 = " OR (COD_CLIENTE = " & CLng(parametro) & ")"
            SQL = "SELECT TOP 100 COD_CLIENTE AS A1_COD, " & _
                  "       LOJA_CLIENTE AS A1_LOJA, " & _
                  "       NOME_CLIENTE AS A1_NOME, " & _
                  "       FANTASIA_CLIENTE AS A1_NREDUZ, " & _
                  "       MUNICIPIO_CLIENTE AS A1_MUN, " & _
                  "       ESTADO_CLIENTE AS A1_EST, " & _
                  "       BAIRRO AS A1_BAIRRO, " & _
                  "       CNPJ_CLIENTE AS CNPJ, " & _
                  "       COD_SAP_CLIENTE AS A1_CODSAP, " & _
                  "       '2' AS A1_MSBLQL " & _
                  "FROM VW_CLIENTES_PEDIDO " & _
                  "WHERE ((FANTASIA_CLIENTE LIKE '%" & pUpper06 & "%') " & _
                  "    OR (NOME_CLIENTE     LIKE '%" & pUpper06 & "%') " & _
                  sqlMarCod06 & _
                  ") AND STATUS = 'ATIVO' " & _
                  "ORDER BY FANTASIA_CLIENTE"

        Case "07"          ' SAQUAREMA
            Dim pUpper07 : pUpper07 = Replace(UCase(parametro), "'", "''")
            Dim pExato07 : pExato07 = Replace(parametro, "'", "''")
            SQL = "SELECT TOP 100 COD_CLIENTE AS A1_COD, LOJA_CLIENTE AS A1_LOJA, NOME_CLIENTE AS A1_NOME, " & _
                  "       FANTASIA_CLIENTE AS A1_NREDUZ, MUNICIPIO_CLIENTE AS A1_MUN, ESTADO_CLIENTE AS A1_EST, " & _
                  "       '' AS A1_BAIRRO, CNPJ_CLIENTE AS CNPJ, '' AS A1_CODSAP, '2' AS A1_MSBLQL " & _
                  "FROM VW_BI_CLIENTES " & _
                  "WHERE ((FANTASIA_CLIENTE LIKE '%" & pUpper07 & "%') " & _
                  "    OR (NOME_CLIENTE     LIKE '%" & pUpper07 & "%') " & _
                  "    OR (COD_CLIENTE = '"          & pExato07  & "')) " & _
                  "  AND STATUS = 'ATIVO' " & _
                  "ORDER BY FANTASIA_CLIENTE"

        Case Else          ' Demais (Westrock/SPO-BI)
            Dim pUpperElse : pUpperElse = Replace(UCase(parametro), "'", "''")
            Dim pExatoElse : pExatoElse = Replace(parametro, "'", "''")
            SQL = "SELECT TOP 100 CODIGO_SAP AS A1_COD, '01' AS A1_LOJA, NOME_CLIENTE AS A1_NOME, " & _
                  "       NOME_REDUZIDO_CLI AS A1_NREDUZ, CIDADE_CLIENTE AS A1_MUN, ESTADO_CLIENTE AS A1_EST, " & _
                  "       '' AS A1_BAIRRO, CNPJ AS CNPJ, CODIGO_SAP AS A1_CODSAP, '2' AS A1_MSBLQL " & _
                  "FROM BI_CLIENTES_WRK " & _
                  "WHERE ((NOME_CLIENTE LIKE '%" & pUpperElse & "%') " & _
                  "    OR (CODIGO_SAP = '"        & pExatoElse & "')) " & _
                  "  AND STATUS_CLI = 'Active' " & _
                  "ORDER BY NOME_CLIENTE"

    End Select

    ' ── Conexão por planta ───────────────────────────────────────────────────
    Select Case planta
        Case "00", "04" : Set rs = Conecta("PIR-DADOSADV", SQL)
        Case "05"       : Set rs = Conecta("BEN-DADOSADV", SQL)
        Case "06"       : Set rs = Conecta("FOR-DADOSADV", SQL)
        Case "07"       : Set rs = Conecta("SAQ-DADOSADV", SQL)
        Case Else       : Set rs = Conecta("SPO-BI",       SQL)
    End Select

    If Err.Number <> 0 Or Not IsObject(rs) Then
        BuscarCliente = "{""sucesso"":false,""mensagem"":""Erro na consulta de clientes: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonCli, primCli
    jsonCli = "["
    primCli = True

    Do While Not rs.EOF
        If Not primCli Then jsonCli = jsonCli & ","
        primCli = False

        Dim nomeExib
        If Trim(rs("A1_NREDUZ") & "") <> "" Then
            nomeExib = Trim(rs("A1_NREDUZ") & "")
        Else
            nomeExib = Trim(rs("A1_NOME") & "")
        End If

        jsonCli = jsonCli & "{"
        jsonCli = jsonCli & """codigo"":"""   & EscapeJSON(Trim(rs("A1_COD")    & "")) & ""","
        jsonCli = jsonCli & """loja"":"""     & EscapeJSON(Trim(rs("A1_LOJA")   & "")) & ""","
        jsonCli = jsonCli & """nome"":"""     & EscapeJSON(nomeExib)                    & ""","
        jsonCli = jsonCli & """cnpj"":"""     & EscapeJSON(Trim(rs("CNPJ")      & "")) & ""","
        jsonCli = jsonCli & """bairro"":"""   & EscapeJSON(Trim(rs("A1_BAIRRO") & "")) & ""","
        jsonCli = jsonCli & """municipio"":"""& EscapeJSON(Trim(rs("A1_MUN")    & "")) & ""","
        jsonCli = jsonCli & """uf"":"""       & EscapeJSON(Trim(rs("A1_EST")    & "")) & ""","
        jsonCli = jsonCli & """codSap"":"""   & EscapeJSON(Trim(rs("A1_CODSAP") & "")) & ""","
        jsonCli = jsonCli & """bloqueado"":"""& EscapeJSON(Trim(rs("A1_MSBLQL") & "")) & """"
        jsonCli = jsonCli & "}"

        rs.MoveNext
    Loop
    jsonCli = jsonCli & "]"

    rs.Close : Set rs = Nothing

    BuscarCliente = "{""sucesso"":true,""mensagem"":"""",""dados"":" & jsonCli & "}"
End Function

' ═══════════════════════════════════════════════════════════════════════════
' BUSCA PRODUTO — SPA Modal
' Parâmetros: parametro (descrição/código), planta (código)
' ═══════════════════════════════════════════════════════════════════════════
Function BuscarProduto()
    On Error Resume Next

    Dim parametro, planta, SQL, rs

    parametro = Trim(Request("parametro"))
    planta    = Trim(Request("planta"))

    If parametro = "" Then
        BuscarProduto = "{""sucesso"":false,""mensagem"":""Parâmetro de busca não informado""}"
        Exit Function
    End If

    ' ── SQL por planta ───────────────────────────────────────────────────────
    If planta = "00" Or planta = "04" Then   ' PIRAPETINGA / UBERABA

        SQL = "SELECT DISTINCT B1_COD, B1_DESC AS DESCRICAO, " & _
              "    CASE WHEN RTRIM(ISNULL(B5_ESTRUTU,'')) = '' THEN B1_ESTRUCH ELSE B5_ESTRUTU END AS ESTRUTURA, " & _
              "    CASE WHEN B1_INPUBE = 'U' THEN 'UBERABA' ELSE 'PIRAPETINGA' END AS FILIAL, " & _
              "    ISNULL(B5_CLI,'') AS B5_CLI " & _
              "FROM SB1010 SB1 WITH(NOLOCK) " & _
              "    LEFT JOIN SB5010 SB5 WITH(NOLOCK) ON B5_COD = B1_COD AND ISNULL(SB5.D_E_L_E_T_,'') = '' " & _
              "WHERE ISNULL(SB1.D_E_L_E_T_,'') = '' " & _
              "  AND B1_STATUS IN ('V','A','W') " & _
              "  AND B1_MSBLQL <> '1' " & _
              "  AND SUBSTRING(ISNULL(B1_COD,''),1,2) <> '77' " & _
              "  AND (B1_COD = '"          & Replace(parametro,       "'", "''") & "' " & _
              "    OR UPPER(B1_DESC)    LIKE UPPER('%" & Replace(parametro, "'", "''") & "%') " & _
              "    OR UPPER(B1_ESTRUCH) LIKE UPPER('%" & Replace(parametro, "'", "''") & "%')) " & _
              "ORDER BY FILIAL, B1_DESC"

    ElseIf planta = "05" Or planta = "06" Or planta = "07" Then  ' BENTO / MARANGUAPE / SAQUAREMA

        SQL = "SELECT CODIGO_ITEM AS B1_COD, REFERENCIA_ITEM AS DESCRICAO, ESTRUTURA, " & _
              "       CIDADE AS FILIAL, CODIGO_CLIENTE AS B5_CLI " & _
              "FROM VW_BI_PRODUTOS_PEDIDO " & _
              "WHERE STATUS = 'ATIVO' " & _
              "  AND (REFERENCIA_ITEM LIKE UPPER('%" & Replace(parametro, "'", "''") & "%') " & _
              "    OR CODIGO_ITEM     LIKE '%"        & Replace(parametro, "'", "''") & "%') " & _
              "ORDER BY REFERENCIA_ITEM"

    Else  ' Fallback

        SQL = "SELECT DISTINCT B1_COD, B1_DESC AS DESCRICAO, B1_ESTRUCH AS ESTRUTURA, " & _
              "    'PIRAPETINGA' AS FILIAL, '' AS B5_CLI " & _
              "FROM SB1010 WITH(NOLOCK) " & _
              "WHERE ISNULL(D_E_L_E_T_,'') = '' " & _
              "  AND B1_STATUS IN ('V','A','W') " & _
              "  AND B1_MSBLQL <> '1' " & _
              "  AND SUBSTRING(ISNULL(B1_COD,''),1,2) <> '77' " & _
              "  AND (B1_COD = '"       & Replace(parametro, "'", "''") & "' " & _
              "    OR UPPER(B1_DESC) LIKE UPPER('%" & Replace(parametro, "'", "''") & "%')) " & _
              "ORDER BY B1_DESC"

    End If

    ' ── Conexão por planta ───────────────────────────────────────────────────
    Select Case planta
        Case "00", "04" : Set rs = Conecta("PIR-DADOSADV", SQL)
        Case "05"       : Set rs = Conecta("BEN-TRIMBOX",  SQL)
        Case "06"       : Set rs = Conecta("FOR-TRIMBOX",  SQL)
        Case "07"       : Set rs = Conecta("SAQ-TRIMBOX",  SQL)
        Case Else       : Set rs = Conecta("PIR-DADOSADV", SQL)
    End Select

    If Err.Number <> 0 Or Not IsObject(rs) Then
        BuscarProduto = "{""sucesso"":false,""mensagem"":""Erro na consulta de produtos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonProd, primProd
    jsonProd = "["
    primProd = True

    Do While Not rs.EOF
        If Not primProd Then jsonProd = jsonProd & ","
        primProd = False

        Dim estrutura, disponivel
        estrutura  = Trim(rs("ESTRUTURA") & "")
        disponivel = IIf(estrutura <> "", "1", "0")

        jsonProd = jsonProd & "{"
        jsonProd = jsonProd & """codigo"":"""    & EscapeJSON(Trim(rs("B1_COD")    & "")) & ""","
        jsonProd = jsonProd & """descricao"":""" & EscapeJSON(Trim(rs("DESCRICAO") & "")) & ""","
        jsonProd = jsonProd & """estrutura"":""" & EscapeJSON(estrutura)                   & ""","
        jsonProd = jsonProd & """filial"":"""    & EscapeJSON(Trim(rs("FILIAL")    & "")) & ""","
        jsonProd = jsonProd & """cliente"":"""   & EscapeJSON(Trim(rs("B5_CLI")    & "")) & ""","
        jsonProd = jsonProd & """disponivel"":"""& disponivel                              & """"
        jsonProd = jsonProd & "}"

        rs.MoveNext
    Loop
    jsonProd = jsonProd & "]"

    rs.Close : Set rs = Nothing

    BuscarProduto = "{""sucesso"":true,""mensagem"":"""",""dados"":" & jsonProd & "}"
End Function

' ═══════════════════════════════════════════════════════════════════════════
' HELPERS
' ═══════════════════════════════════════════════════════════════════════════
Function EscapeJSON(ByVal str)
    If IsNull(str) Then EscapeJSON = "" : Exit Function
    str = Replace(str, "\",  "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, Chr(13), "")
    str = Replace(str, Chr(10), " ")
    str = Replace(str, Chr(9),  " ")
    EscapeJSON = Trim(str)
End Function

Function FormatarNumeroJSON(valor)
    If IsNull(valor) Or IsEmpty(valor) Or valor = "" Then
        FormatarNumeroJSON = "0"
    Else
        FormatarNumeroJSON = Replace(CStr(valor), ",", ".")
    End If
End Function
%>
