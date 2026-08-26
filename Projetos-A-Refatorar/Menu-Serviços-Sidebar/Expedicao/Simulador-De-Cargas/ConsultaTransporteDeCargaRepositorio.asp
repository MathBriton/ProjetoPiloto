<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../funcoes/funcoes.asp"-->
<!--#Include File="../../../includes/variaveis.asp"-->
<%
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Buffer      = True
Response.Expires     = 0
Session.LCID         = 1046
Response.ContentType = "application/json"

' ── Contexto de sessão ──────────────────────────────────────────────────
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id",    Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome",  Session("NOME")

' ── Dispatcher ──────────────────────────────────────────────────────────
Dim acao, resultado
acao = Trim(Request("acao"))

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada.""}")
    Response.End()
End If

Select Case acao
    Case "validaCliente"  : resultado = ValidaCliente()
    Case "validaProduto"  : resultado = ValidaProduto()
    Case "buscaCliente"   : resultado = BuscaCliente()
    Case "buscaProduto"   : resultado = BuscaProduto()
    Case "consultar"      : resultado = Consultar()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}"
End Select

Response.Write(resultado)

' ════════════════════════════════════════════════════════════════════════
'  FUNÇÕES
' ════════════════════════════════════════════════════════════════════════

' ── validaCliente ────────────────────────────────────────────────────────
' Valida se um cliente existe na BI_CLIENTES e retorna nome/razão social.
' Parâmetros: planta, cliente, loja, codsiga, nivel, tipo
' ────────────────────────────────────────────────────────────────────────
Function ValidaCliente()
    Dim vc_planta, vc_cliente, vc_loja, vc_nivel, vc_codsiga, vc_tipo
    Dim vc_SQL, vc_RS

    vc_planta  = Trim(Request("planta"))
    vc_cliente = Trim(Request("cliente"))
    vc_loja    = Trim(Request("loja"))
    vc_nivel   = Trim(Request("nivel"))
    vc_codsiga = Trim(Request("codsiga"))
    vc_tipo    = Trim(Request("tipo"))

    ' Guarda obrigatório: cliente ou tipo = Grupo
    If vc_cliente = "" And vc_tipo <> "Grupo" Then
        ValidaCliente = "{""sucesso"": false, ""mensagem"": ""Informe o cliente."", ""erro"": ""ErroCliente""}"
        Exit Function
    End If

    ' Guarda obrigatório: codsiga quando nivel <> 0
    If vc_nivel <> "" And vc_nivel <> "0" And vc_codsiga = "" Then
        ValidaCliente = "{""sucesso"": false, ""mensagem"": ""Sessão expirada."", ""erro"": ""ErroLogof""}"
        Exit Function
    End If

    If vc_nivel = "" Then
        ValidaCliente = "{""sucesso"": false, ""mensagem"": ""Nível não informado."", ""erro"": ""ErroLogof""}"
        Exit Function
    End If

    ' Campo de pesquisa varia por planta (BENTO/MARANGUAPE usam CODIGO_ERP)
    Dim vc_campoCod
    If InArray(vc_planta, Array("BENTO", "MARANGUAPE")) Then
        vc_campoCod = "CODIGO_ERP"
    Else
        vc_campoCod = "COD_CLIENTE"
    End If

    vc_SQL = "SELECT NOME_CLIENTE " &_
             "FROM BI_CLIENTES " &_
             "WHERE " & vc_campoCod & " = '" & Replace(vc_cliente, "'", "''") & "' " &_
             "AND LOJA_CLIENTE = '" & Replace(vc_loja,    "'", "''") & "' " &_
             "AND CIDADE       = '" & Replace(vc_planta,  "'", "''") & "' " &_
             IIf(vc_nivel <> "0", "AND COD_VENDEDOR = '" & Replace(vc_codsiga, "'", "''") & "' ", "") &_
             "ORDER BY NOME_CLIENTE, LOJA_CLIENTE"

    On Error Resume Next
    Set vc_RS = Conecta("SPO-BI", vc_SQL)
    If Err.Number <> 0 Or Not IsObject(vc_RS) Then
        ValidaCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar base de clientes."", ""erro"": ""ErroDB""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not vc_RS.EOF Then
        Dim vc_nome
        vc_nome = EscapeJSON(vc_RS("NOME_CLIENTE") & "")
        vc_RS.Close()
        Set vc_RS = Nothing
        ValidaCliente = "{""sucesso"": true, ""mensagem"": """", ""dados"": {""nomeCliente"": """ & vc_nome & """, ""tipo"": """ & EscapeJSON(vc_tipo) & """}}"
    Else
        vc_RS.Close()
        Set vc_RS = Nothing
        ValidaCliente = "{""sucesso"": false, ""mensagem"": ""Cliente não encontrado."", ""erro"": ""ErroNEncontrado"", ""dados"": {""tipo"": """ & EscapeJSON(vc_tipo) & """}}"
    End If
End Function


' ── validaProduto ────────────────────────────────────────────────────────
' Valida se um produto existe na VW_QUANTIDADE_CAMINHAO para o cliente.
' Parâmetros: cliente, loja, produto, nivel, tipo
' ────────────────────────────────────────────────────────────────────────
Function ValidaProduto()
    Dim vp_cliente, vp_loja, vp_produto, vp_nivel, vp_tipo
    Dim vp_SQL, vp_RS

    vp_cliente = Trim(Request("cliente"))
    vp_loja    = Trim(Request("loja"))
    vp_produto = Trim(Request("produto"))
    vp_nivel   = Trim(Request("nivel"))
    vp_tipo    = Trim(Request("tipo"))

    If vp_cliente = "" And vp_tipo <> "Produto" Then
        ValidaProduto = "{""sucesso"": false, ""mensagem"": ""Informe o produto."", ""erro"": ""ErroPRODUTO""}"
        Exit Function
    End If

    If vp_nivel = "" Then
        ValidaProduto = "{""sucesso"": false, ""mensagem"": ""Sessão expirada."", ""erro"": ""ErroLogof""}"
        Exit Function
    End If

    vp_SQL = "SELECT DISTINCT CODIGO, REFERENCIA " &_
             "FROM VW_QUANTIDADE_CAMINHAO " &_
             "WHERE CODIGO_CLIENTE = '" & Replace(vp_cliente, "'", "''") & "' " &_
             "AND LOJA_CLIENTE     = '" & Replace(vp_loja,    "'", "''") & "' " &_
             "AND CODIGO           = '" & Replace(vp_produto, "'", "''") & "'"

    On Error Resume Next
    Set vp_RS = Conecta("PIR-DADOSADV", vp_SQL)
    If Err.Number <> 0 Or Not IsObject(vp_RS) Then
        ValidaProduto = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar base de produtos."", ""erro"": ""ErroDB""}"
        Exit Function
    End If
    On Error GoTo 0

    If Not vp_RS.EOF Then
        Dim vp_ref
        vp_ref = EscapeJSON(Replace(Replace(vp_RS("REFERENCIA") & "", ".", ""), ",", "."))
        vp_RS.Close()
        Set vp_RS = Nothing
        ValidaProduto = "{""sucesso"": true, ""mensagem"": """", ""dados"": {""referencia"": """ & vp_ref & """, ""tipo"": """ & EscapeJSON(vp_tipo) & """}}"
    Else
        vp_RS.Close()
        Set vp_RS = Nothing
        ValidaProduto = "{""sucesso"": false, ""mensagem"": ""Produto não encontrado."", ""erro"": ""ErroPRODUTONAOENCONTRADO"", ""dados"": {""tipo"": """ & EscapeJSON(vp_tipo) & """}}"
    End If
End Function


' ── buscaCliente ─────────────────────────────────────────────────────────
' Pesquisa clientes na BI_CLIENTES por nome, fantasia ou código.
' Parâmetros: parametro, planta (id6), status (id9), nivel, codsiga (id4=1 → nivel 0)
' ────────────────────────────────────────────────────────────────────────
Function BuscaCliente()
    Dim bc_parametro, bc_planta, bc_status, bc_nivel, bc_codsiga
    Dim bc_campoCod, bc_SQL, bc_RS
    Dim bc_primeiroItem, bc_jsonArray, bc_jsonItem

    bc_parametro = Trim(Request("parametro"))
    bc_planta    = Trim(Request("planta"))
    bc_status    = Trim(Request("status"))
    bc_codsiga   = Trim(Request("codsiga"))
    bc_nivel     = Trim(Request("nivel"))

    ' id4 = "1" sinaliza que o usuário é admin (nivel 0)
    If Trim(Request("id4")) = "1" Then bc_nivel = "0"
    If bc_codsiga = "" Then bc_codsiga = " "

    If bc_parametro = "" Then
        BuscaCliente = "{""sucesso"": false, ""mensagem"": ""Informe o parâmetro de pesquisa.""}"
        Exit Function
    End If

    If InArray(bc_planta, Array("BENTO", "MARANGUAPE")) Then
        bc_campoCod = "CODIGO_ERP"
    Else
        bc_campoCod = "COD_CLIENTE"
    End If

    Dim bc_paramEsc, bc_plantaEsc, bc_codsigaEsc
    bc_paramEsc  = Replace(UCase(Trim(bc_parametro)), "'", "''")
    bc_plantaEsc = Replace(bc_planta, "'", "''")
    bc_codsigaEsc = Replace(bc_codsiga, "'", "''")

    bc_SQL = "SELECT " &_
               "PLANTA, CIDADE, COD_CLIENTE, LOJA_CLIENTE, NOME_CLIENTE, FANTASIA_CLIENTE, " &_
               "CNPJ_CLIENTE, MUNICIPIO_CLIENTE, ESTADO_CLIENTE, COD_VENDEDOR, NOME_VENDEDOR, " &_
               "CNPJ_VENDEDOR, COD_SAP_CLIENTE, CHAVE_CLIENTE, END_CLIENTE, CEP_CLIENTE, STATUS, " &_
               "CLIENTE_GRUPO, GRUPO_DO_CLIENTE, LIMITE_CREDITO, INSCRICAO_ESTADUAL, " &_
               "COD_MUNICIPIOS_IBGE, TES, EMAIL, BAIRRO, CONDICAO_PAGAMENTO, DIAS_TRANSPORTE, " &_
               "CODIGO_SEGMENTO, DESCRICAO_SEGMENTO, CONTATO, TELEFONE, CODIGO_ERP " &_
             "FROM BI_CLIENTES " &_
             "WHERE ((FANTASIA_CLIENTE LIKE '%" & bc_paramEsc & "%') " &_
                 "OR (NOME_CLIENTE     LIKE '%" & bc_paramEsc & "%') " &_
                 "OR (" & bc_campoCod & " LIKE '%" & Replace(Trim(bc_parametro), "'", "''") & "%')) " &_
               "AND CIDADE LIKE '" & bc_plantaEsc & "' " &_
               IIf(bc_status <> "", "AND STATUS = '" & Replace(bc_status, "'", "''") & "' ", "") &_
               IIf(bc_nivel <> "0", "AND COD_VENDEDOR = '" & bc_codsigaEsc & "' ", "") &_
             "ORDER BY NOME_CLIENTE, LOJA_CLIENTE"

    On Error Resume Next
    Set bc_RS = Conecta("SPO-BI", bc_SQL)
    If Err.Number <> 0 Or Not IsObject(bc_RS) Then
        BuscaCliente = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar base de clientes.""}"
        Exit Function
    End If
    On Error GoTo 0

    bc_primeiroItem = True
    bc_jsonArray = "["

    Do While Not bc_RS.EOF
        If Not bc_primeiroItem Then bc_jsonArray = bc_jsonArray & ","

        ' Código de exibição varia por planta
        Dim bc_codExibicao, bc_sap, bc_stRs
        If InArray(bc_RS("CIDADE") & "", Array("BENTO", "MARANGUAPE")) Then
            bc_codExibicao = EscapeJSON(bc_RS("CODIGO_ERP") & "")
        Else
            bc_codExibicao = EscapeJSON(bc_RS("COD_CLIENTE") & "")
        End If

        bc_sap  = Trim(bc_RS("COD_SAP_CLIENTE") & "")
        bc_stRs = Trim(bc_RS("STATUS") & "")

        ' Flag de selecionável: precisa de SAP e não pode ser INATIVO
        Dim bc_selecionavel
        bc_selecionavel = "true"
        If bc_sap = "" Or bc_stRs = "INATIVO" Then bc_selecionavel = "false"

        bc_jsonItem = "{"
        bc_jsonItem = bc_jsonItem & """codExibicao"": """      & bc_codExibicao                                      & """, "
        bc_jsonItem = bc_jsonItem & """codCliente"": """       & EscapeJSON(bc_RS("COD_CLIENTE")    & "")            & """, "
        bc_jsonItem = bc_jsonItem & """lojaCliente"": """      & EscapeJSON(bc_RS("LOJA_CLIENTE")   & "")            & """, "
        bc_jsonItem = bc_jsonItem & """nomeCliente"": """      & EscapeJSON(bc_RS("NOME_CLIENTE")   & "")            & """, "
        bc_jsonItem = bc_jsonItem & """fantasiaCliente"": """  & EscapeJSON(bc_RS("FANTASIA_CLIENTE") & "")          & """, "
        bc_jsonItem = bc_jsonItem & """cnpj"": """             & EscapeJSON(bc_RS("CNPJ_CLIENTE")   & "")            & """, "
        bc_jsonItem = bc_jsonItem & """municipio"": """        & EscapeJSON(bc_RS("MUNICIPIO_CLIENTE") & "")         & """, "
        bc_jsonItem = bc_jsonItem & """estado"": """           & EscapeJSON(bc_RS("ESTADO_CLIENTE") & "")            & """, "
        bc_jsonItem = bc_jsonItem & """bairro"": """           & EscapeJSON(bc_RS("BAIRRO")         & "")            & """, "
        bc_jsonItem = bc_jsonItem & """codVendedor"": """      & EscapeJSON(bc_RS("COD_VENDEDOR")   & "")            & """, "
        bc_jsonItem = bc_jsonItem & """nomeVendedor"": """     & EscapeJSON(bc_RS("NOME_VENDEDOR")  & "")            & """, "
        bc_jsonItem = bc_jsonItem & """codSap"": """           & EscapeJSON(bc_sap)                                  & """, "
        bc_jsonItem = bc_jsonItem & """status"": """           & EscapeJSON(bc_stRs)                                 & """, "
        bc_jsonItem = bc_jsonItem & """condPagamento"": """    & EscapeJSON(bc_RS("CONDICAO_PAGAMENTO") & "")        & """, "
        bc_jsonItem = bc_jsonItem & """diasTransporte"": """   & EscapeJSON(bc_RS("DIAS_TRANSPORTE") & "")           & """, "
        bc_jsonItem = bc_jsonItem & """codigoERP"": """        & EscapeJSON(bc_RS("CODIGO_ERP")     & "")            & """, "
        bc_jsonItem = bc_jsonItem & """cidade"": """           & EscapeJSON(bc_RS("CIDADE")         & "")            & """, "
        bc_jsonItem = bc_jsonItem & """selecionavel"": "       & bc_selecionavel
        bc_jsonItem = bc_jsonItem & "}"

        bc_jsonArray = bc_jsonArray & bc_jsonItem
        bc_primeiroItem = False
        bc_RS.MoveNext()
    Loop

    bc_jsonArray = bc_jsonArray & "]"
    bc_RS.Close()
    Set bc_RS = Nothing

    BuscaCliente = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & bc_jsonArray & "}"
End Function


' ── buscaProduto ─────────────────────────────────────────────────────────
' Pesquisa produtos na VW_QUANTIDADE_CAMINHAO por código ou referência.
' Parâmetros: parametro, codcli (id3), lojacli (id4)
' ────────────────────────────────────────────────────────────────────────
Function BuscaProduto()
    Dim bp_codcli, bp_lojcli, bp_parametro
    Dim bp_SQL, bp_RS
    Dim bp_primeiroItem, bp_jsonArray, bp_jsonItem

    bp_codcli    = Trim(Request("codcli"))
    bp_lojcli    = Trim(Request("lojacli"))
    bp_parametro = Trim(Request("parametro"))

    If bp_parametro = "" Then
        BuscaProduto = "{""sucesso"": false, ""mensagem"": ""Informe o parâmetro de pesquisa.""}"
        Exit Function
    End If

    bp_SQL = "SELECT DISTINCT CODIGO, REFERENCIA " &_
             "FROM VW_QUANTIDADE_CAMINHAO " &_
             "WHERE CODIGO_CLIENTE = '" & Replace(bp_codcli,    "'", "''") & "' " &_
             "AND LOJA_CLIENTE     = '" & Replace(bp_lojcli,    "'", "''") & "' " &_
             "AND (CODIGO = '"          & Replace(bp_parametro, "'", "''") & "' " &_
                  "OR UPPER(REFERENCIA) LIKE UPPER('%" & Replace(bp_parametro, "'", "''") & "%'))"

    On Error Resume Next
    Set bp_RS = Conecta("PIR-DADOSADV", bp_SQL)
    If Err.Number <> 0 Or Not IsObject(bp_RS) Then
        BuscaProduto = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar base de produtos.""}"
        Exit Function
    End If
    On Error GoTo 0

    bp_primeiroItem = True
    bp_jsonArray    = "["

    Do While Not bp_RS.EOF
        If Not bp_primeiroItem Then bp_jsonArray = bp_jsonArray & ","

        Dim bp_ref
        bp_ref = EscapeJSON(Replace(Replace(bp_RS("REFERENCIA") & "", ".", ""), ",", "."))

        bp_jsonItem = "{"
        bp_jsonItem = bp_jsonItem & """codigo"": """    & EscapeJSON(bp_RS("CODIGO") & "") & """, "
        bp_jsonItem = bp_jsonItem & """referencia"": """ & bp_ref                          & """"
        bp_jsonItem = bp_jsonItem & "}"

        bp_jsonArray    = bp_jsonArray & bp_jsonItem
        bp_primeiroItem = False
        bp_RS.MoveNext()
    Loop

    bp_jsonArray = bp_jsonArray & "]"
    bp_RS.Close()
    Set bp_RS = Nothing

    BuscaProduto = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & bp_jsonArray & "}"
End Function


' ── consultar ────────────────────────────────────────────────────────────
' Retorna os dados de quantidade por caminhão para o produto/cliente.
' Parâmetros: planta, codcli, lojacli, produto
' ────────────────────────────────────────────────────────────────────────
Function Consultar()
    Dim ct_planta, ct_codcli, ct_lojacli, ct_produto
    Dim ct_SQL, ct_RS, ct_dsn
    Dim ct_primeiroItem, ct_jsonArray, ct_jsonItem

    ct_planta  = Trim(Request("planta"))
    ct_codcli  = Trim(Request("codcli"))
    ct_lojacli = Trim(Request("lojacli"))
    ct_produto = Trim(Request("produto"))

    If ct_planta = "" Then
        Consultar = "{""sucesso"": false, ""mensagem"": ""Selecione uma filial.""}"
        Exit Function
    End If
    If ct_codcli = "" Or ct_lojacli = "" Then
        Consultar = "{""sucesso"": false, ""mensagem"": ""Informe o cliente e a loja.""}"
        Exit Function
    End If
    If ct_produto = "" Then
        Consultar = "{""sucesso"": false, ""mensagem"": ""Informe o produto.""}"
        Exit Function
    End If

    ' Roteamento de DSN por planta
    Select Case ct_planta
        Case "PIRAPETINGA" : ct_dsn = "PIR-DADOSADV"
        Case "UBERABA"     : ct_dsn = "UBE-DADOSADV"
        Case "BENTO"       : ct_dsn = "BEN-DADOSADV"
        Case "MARANGUAPE"  : ct_dsn = "FOR-DADOSADV"
        Case "SAQUAREMA"   : ct_dsn = "SAQ-DADOSADV"
        Case Else
            Consultar = "{""sucesso"": false, ""mensagem"": ""Filial inválida.""}"
            Exit Function
    End Select

    ct_SQL = "SELECT TIPO, TOTAL_CAIXAS, TOTAL_PALETE, M2_TOTAL, " &_
                    "TOTAL_PALETE_2800, TOTAL_CAIXAS_2800, M2_TOTAL_2800 " &_
             "FROM VW_QUANTIDADE_CAMINHAO " &_
             "WHERE RTRIM(CODIGO_CLIENTE) = '" & Replace(ct_codcli,  "'", "''") & "' " &_
               "AND RTRIM(LOJA_CLIENTE)   = '" & Replace(ct_lojacli, "'", "''") & "' " &_
               "AND RTRIM(CODIGO)         = '" & Replace(ct_produto, "'", "''") & "' " &_
             "ORDER BY TOTAL_CAIXAS DESC"

    On Error Resume Next
    Set ct_RS = Conecta(ct_dsn, ct_SQL)
    If Err.Number <> 0 Or Not IsObject(ct_RS) Then
        Consultar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar a base de dados.""}"
        Exit Function
    End If
    On Error GoTo 0

    ct_primeiroItem = True
    ct_jsonArray    = "["

    Do While Not ct_RS.EOF
        If Not ct_primeiroItem Then ct_jsonArray = ct_jsonArray & ","

        ct_jsonItem = "{"
        ct_jsonItem = ct_jsonItem & """tipo"": """           & EscapeJSON(ct_RS("TIPO")             & "") & """, "
        ct_jsonItem = ct_jsonItem & """totalCaixas"": """    & EscapeJSON(ct_RS("TOTAL_CAIXAS")      & "") & """, "
        ct_jsonItem = ct_jsonItem & """totalPalete"": """    & EscapeJSON(ct_RS("TOTAL_PALETE")      & "") & """, "
        ct_jsonItem = ct_jsonItem & """m2Total"": """        & EscapeJSON(ct_RS("M2_TOTAL")          & "") & """, "
        ct_jsonItem = ct_jsonItem & """totalCaixas2800"": """ & EscapeJSON(ct_RS("TOTAL_CAIXAS_2800") & "") & """, "
        ct_jsonItem = ct_jsonItem & """totalPalete2800"": """ & EscapeJSON(ct_RS("TOTAL_PALETE_2800") & "") & """, "
        ct_jsonItem = ct_jsonItem & """m2Total2800"": """    & EscapeJSON(ct_RS("M2_TOTAL_2800")     & "") & """"
        ct_jsonItem = ct_jsonItem & "}"

        ct_jsonArray    = ct_jsonArray & ct_jsonItem
        ct_primeiroItem = False
        ct_RS.MoveNext()
    Loop

    ct_jsonArray = ct_jsonArray & "]"
    ct_RS.Close()
    Set ct_RS = Nothing

    Consultar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & ct_jsonArray & "}"
End Function
%>
