<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Session.LCID = 1046

Dim idUsuarioLogado, nomeUsuario
idUsuarioLogado = Session("CODIGO")
nomeUsuario = Session("NOME")
nivelAlteracaoPedido = Session("NIVELSOLICITACAOALTERACAOPEDIDO")

' Permissões do usuário
temPermissaoChamados = False
temPermissaoAnaliseCredito = False
temPermissaoAprovacaoTransferencia = False
temPermissaoAprovacaoFNC = False
temPermissaoAnaliseFNC = False
temPermissaoRespostaFNC = False
temPermissaoAprovacaoOrcamento = False
temPermissaoAnaliseOrcamento = False
temPermissaoAprovacaoCredito = False
temPermissaoCadastroUnificado = False
temPermissaoAlteracaoPedido = False

' Auxiliares
Dim nivelAprovadorOrcamento, nivelValidadorSolicitacao, nivelCadastroUnificado
Dim plantasAprovadorOrcamento, plantasValidadorSolicitacao, plantasAnalistaOrcamento
Dim minimoSolicitacao, maximoSolicitacao

nivelCadastroUnificado = ""
nivelAprovadorOrcamento = ""
nivelValidadorSolicitacao = ""

plantasAprovadorOrcamento = ""
plantasValidadorSolicitacao = ""

minimoSolicitacao = 0
maximoSolicitacao = 0

idAtendente = ""

' Contadores
totalChamados = 0
totalAnaliseCredito = 0
totalAprovacaoCredito = 0
totalTransferenciaCredito = 0
totalAnaliseFNC = 0
totalAnaliseFNCAbertas = 0
totalAnaliseFNCRespondidas = 0
totalAnaliseFNCAprovadas = 0
totalAprovacaoFNC = 0
totalRespostaFNC = 0
totalAprovacaoOrcamento = 0
totalAnaliseOrcamento = 0
totalCadastroUnificado = 0
totalAlteracaoPedidoPiraUbe = 0
totalAlteracaoPedidoOutrasPlantas = 0
totalGeral = 0

Dim qtdCardsVisiveis
qtdCardsVisiveis = 0

Call VerificarPermissoes()

If temPermissaoAnaliseCredito Then
    Call ContarAnalisesCredito()
End If

If temPermissaoChamados Then
    Call ContarChamados()
End If

If temPermissaoAprovacaoTransferencia Then
    Call ContarTransferenciasCredito()
End If

If temPermissaoAprovacaoCredito Then
    Call ContarAprovacoesCredito()
End If

If temPermissaoAnaliseFNC Then
    Call ContarAnalisesFNC()
End If

If temPermissaoAprovacaoFNC Then
    Call ContarAprovacoesFNC()
End If

If temPermissaoRespostaFNC Then
    Call ContarRespostasFNC()
End If

If temPermissaoAprovacaoOrcamento Then
    Call ContarAprovacoesOrcamento()
End If

If temPermissaoAnaliseOrcamento Then
    Call ContarAnalisesOrcamento()
End If

If temPermissaoCadastroUnificado Then
    Call ContarCadastroUnificado()
End If

If temPermissaoAlteracaoPedido Then
    Call ContarAlteracoesPedido()
End If

totalGeral = totalChamados + totalAnaliseCredito + totalAprovacaoCredito + totalTransferenciaCredito + totalAprovacaoFNC + totalAnaliseFNC + totalRespostaFNC + totalAprovacaoOrcamento + totalAnaliseOrcamento + totalCadastroUnificado + totalAlteracaoPedidoPiraUbe + totalAlteracaoPedidoOutrasPlantas

'--------------------
'    Sub-rotinas
'--------------------

Sub VerificarPermissoes()
    Dim SQL
    
    ' Permissão para Solicitações de Crédito
    SQL = "SELECT TOP 1 "
    SQL = SQL & " MIN(vc.Minimo) AS Minimo,"
    SQL = SQL & " MAX(vc.Maximo) AS Maximo,"
    SQL = SQL & " CAST(vc.Nivel AS NVARCHAR(10)) as CodNivel,"
    SQL = SQL & " ISNULL('' + STUFF(("
    SQL = SQL & "     SELECT ',' + QUOTENAME(vc2.Planta, '''')  "
    SQL = SQL & "     FROM ValidadoresCredito vc2"
    SQL = SQL & "     INNER JOIN HELPDESK..Empresas e2 ON vc2.Planta = e2.Empresa"
    SQL = SQL & "     WHERE vc2.IdValidador = vc.IdValidador"
    SQL = SQL & "     AND vc2.DataExclusao IS NULL"
    SQL = SQL & "     ORDER BY vc2.Planta"
    SQL = SQL & "     FOR XML PATH(''), TYPE"
    SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + '', '[]') AS Plantas"
    SQL = SQL & " FROM ValidadoresCredito vc"
    SQL = SQL & " WHERE vc.IdValidador = " & idUsuarioLogado
    SQL = SQL & " AND vc.DataExclusao is NULL"
    SQL = SQL & " GROUP BY vc.IdValidador, vc.Nivel"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            plantasValidadorSolicitacao = PIR_CORPORATIVO("Plantas")
            nivelValidadorSolicitacao = PIR_CORPORATIVO("codNivel")
            minimoSolicitacao = PIR_CORPORATIVO("Minimo")
            maximoSolicitacao = PIR_CORPORATIVO("Maximo")
    
            If nivelValidadorSolicitacao = "10" Then
                temPermissaoAnaliseCredito = True 
                temPermissaoAprovacaoTransferencia = True
            End If

            temPermissaoAprovacaoCredito = True
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    ' Permissão para Analise de FNC
    SQL = "SELECT TOP 1 IdResponsavel"
    SQL = SQL & " FROM Responsavel a"
    SQL = SQL & " WHERE a.Id = " & idUsuarioLogado
    SQL = SQL & " AND a.Tipo = '1'" 
    SQL = SQL & " AND (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " GROUP BY a.IdResponsavel, a.Nome, a.Email"
    SQL = SQL & " ORDER BY a.Nome" 

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        If NOT BEN_CORPORATIVO.EOF Then
            temPermissaoAnaliseFNC = True
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
    
    ' Permissão para Aprovacao de FNC
    SQL = "SELECT TOP 1 Id"
    SQL = SQL & " FROM AprovadorFNC a"
    SQL = SQL & " WHERE a.Id = " & idUsuarioLogado
    SQL = SQL & " AND (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " GROUP BY a.Id, a.Nome, a.Nivel, a.Email, a.AprovadorTipo"
    SQL = SQL & " ORDER BY a.Nome" 

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        If NOT BEN_CORPORATIVO.EOF Then
            temPermissaoAprovacaoFNC = True
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    ' Permissão para Resposta de FNC
    SQL = "SELECT TOP 1 IdResponsavel"
    SQL = SQL & " FROM Responsavel a"
    SQL = SQL & " WHERE a.Id = " & idUsuarioLogado
    SQL = SQL & " AND a.Tipo = '2'" 
    SQL = SQL & " AND (a.Excluido = 0 OR a.Excluido IS NULL)"
    SQL = SQL & " GROUP BY a.IdResponsavel, a.Nome, a.Email"
    SQL = SQL & " ORDER BY a.Nome" 

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        If NOT BEN_CORPORATIVO.EOF Then
            temPermissaoRespostaFNC = True
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
    
    ' Permissão para Aprovação de Orçamento
    SQL = "SELECT TOP 1 a.NIVEL as Nivel, "
    SQL = SQL & " ISNULL('' + STUFF(("
    SQL = SQL & "    SELECT ',' + QUOTENAME(a2.PLANTA, '''')  "
    SQL = SQL & "    FROM APROVADORES_ORCAMENTO a2"
    SQL = SQL & "    WHERE a2.ID_APROVADOR = a.ID_APROVADOR"
    SQL = SQL & "    AND (a2.EXCLUIDO = 0 OR a2.EXCLUIDO IS NULL)"
    SQL = SQL & "    ORDER BY a2.PLANTA"
    SQL = SQL & "    FOR XML PATH(''), TYPE"
    SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + '', '[]') AS Plantas"
    SQL = SQL & " FROM APROVADORES_ORCAMENTO a"
    SQL = SQL & " WHERE a.ID_APROVADOR = " & idUsuarioLogado
    SQL = SQL & " AND (a.EXCLUIDO = 0 OR a.EXCLUIDO IS NULL)"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            temPermissaoAprovacaoOrcamento = True
            nivelAprovadorOrcamento = PIR_CORPORATIVO("Nivel")
            plantasAprovadorOrcamento = PIR_CORPORATIVO("Plantas")
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    ' Permissão para Analises de Orçamento
    SQL = "SELECT TOP 1 "
    SQL = SQL & " ISNULL('' + STUFF(("
    SQL = SQL & "    SELECT ',' + QUOTENAME(a2.PLANTA, '''')  "
    SQL = SQL & "    FROM ANALISTAS_ORCAMENTO a2"
    SQL = SQL & "    WHERE a2.CODIGO_USUARIO = a.CODIGO_USUARIO"
    SQL = SQL & "    AND (a2.EXCLUIDO = 0 OR a2.EXCLUIDO IS NULL)"
    SQL = SQL & "    ORDER BY a2.PLANTA"
    SQL = SQL & "    FOR XML PATH(''), TYPE"
    SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + '', '[]') AS Plantas"
    SQL = SQL & " FROM ANALISTAS_ORCAMENTO a"
    SQL = SQL & " WHERE a.CODIGO_USUARIO = " & idUsuarioLogado
    SQL = SQL & " AND (a.EXCLUIDO = 0 OR a.EXCLUIDO IS NULL)"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            temPermissaoAnaliseOrcamento = True
            plantasAnalistaOrcamento = PIR_CORPORATIVO("Plantas")
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    ' Permissão para Cadastro Unificado
    SQL = "SELECT CASE WHEN ("
    SQL = SQL & "   c.NivelE2OPEN = 1 "
    SQL = SQL & "   OR c.NivelControlesInternos = 1 "
    SQL = SQL & "   OR c.NivelControladoria = 1 "
    SQL = SQL & "   OR c.NivelFiscal = 1 "
    SQL = SQL & "   OR c.NivelE2OPEN = 1 "
    SQL = SQL & ") THEN 1 ELSE 0 END as temPermissao, "
    SQL = SQL & "c.NivelE2OPEN, c.NivelControlesInternos, c.NivelControladoria, c.NivelFiscal "
    SQL = SQL & "FROM CRM_DTI c WHERE CRM_D_E_L_E_T <> '*' "
    SQL = SQL & "AND c.CRM_R_E_C_N_O = " & idUsuarioLogado & " "
    SQL = SQL & "ORDER BY CRM_NOME"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If NOT PIR_HELPDESK.EOF Then
            temPermissaoCadastroUnificado = PIR_HELPDESK("temPermissao")
            If PIR_HELPDESK("NivelE2OPEN") <> "" Then nivelCadastroUnificado = "E"
            If PIR_HELPDESK("NivelControlesInternos") <> "" Then nivelCadastroUnificado = "I"
            If PIR_HELPDESK("NivelControladoria") <> "" Then nivelCadastroUnificado = "C"
            If PIR_HELPDESK("NivelFiscal") <> "" Then nivelCadastroUnificado = "S"

            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    ' Permissão para Validação de Alteração de pedidos
    SQL = "SELECT CASE WHEN c.NivelSolicitacaoAlteracaoPedido IS NOT NULL THEN 1 ELSE 0 END as temPermissao, c.NivelSolicitacaoAlteracaoPedido "
    SQL = SQL & "FROM CRM_DTI c WHERE CRM_D_E_L_E_T <> '*' "
    SQL = SQL & "AND c.CRM_R_E_C_N_O = " & idUsuarioLogado & " "
    SQL = SQL & "ORDER BY CRM_NOME"

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If NOT PIR_HELPDESK.EOF Then
            temPermissaoAlteracaoPedido = PIR_HELPDESK("temPermissao")
            NivelSolicitacaoAlteracaoPedido = PIR_HELPDESK("NivelSolicitacaoAlteracaoPedido")
    
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    ' Permissão Visualização de chamados
    SQL = "SELECT TOP 1 Id FROM AtendentesChamado"
    SQL = SQL & " WHERE IdUsuario = '" & idUsuarioLogado & "'"
    SQL = SQL & " AND FlagInativo = 0"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            temPermissaoChamados = True
            idAtendente = PIR_CORPORATIVO("Id")
    
            qtdCardsVisiveis = qtdCardsVisiveis + 1
        End If

        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAnalisesCredito()
    Dim SQL, statusValidacao
    
    statusValidacao = ObterStatusPermitidoSolicitacao(nivelValidadorSolicitacao, "ANALISE")

    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM SOLCREDITOCLIENTE "
    SQL = SQL & "INNER JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = LOGININPANET COLLATE SQL_Latin1_General_CP437_BIN AND ISNULL(CRM_D_E_L_E_T,'') = ''"
    SQL = SQL & "WHERE STATUS = '" & statusValidacao & "'"
    SQL = SQL & "AND Planta IN (" & plantasValidadorSolicitacao & ") "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) >= " & minimoSolicitacao & " "
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalAnaliseCredito = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarChamados()   
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM Chamados c "
    SQL = SQL & "WHERE 1=1 "
    SQL = SQL & "AND c.FlagInativo = 0 "
    SQL = SQL & "AND (c.IdAtendente = " & idAtendente & " OR c.IdAtendente IS NULL)"
    SQL = SQL & "AND c.IdDepartamento IN (SELECT IdDepartamento FROM AtendenteDepartamentoChamado WHERE IdAtendente = " & idAtendente & " )"
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalChamados = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAprovacoesCredito()
    Dim SQL, statusValidacao
    
    statusValidacao = ObterStatusPermitidoSolicitacao(nivelValidadorSolicitacao, "APROVACAO")

    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM SOLCREDITOCLIENTE "
    SQL = SQL & "INNER JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = LOGININPANET COLLATE SQL_Latin1_General_CP437_BIN AND ISNULL(CRM_D_E_L_E_T,'') = ''"
    SQL = SQL & "WHERE STATUS = '" & statusValidacao & "'"
    SQL = SQL & "AND Planta IN (" & plantasValidadorSolicitacao & ") "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) >= " & minimoSolicitacao & " "
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalAprovacaoCredito = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarTransferenciasCredito()
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM TransferenciaCredito "
    SQL = SQL & "WHERE Status = 'PENDENTE' "
    SQL = SQL & "AND PlantaDestino IN (" & plantasValidadorSolicitacao & ") "
    
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalTransferenciaCredito = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAnalisesFNC()    
    SQL = SQL & "SELECT "
    SQL = SQL & "    SUM(CASE WHEN f.Situacao = 1 and (f.Aprovada <> 1 or f.Aprovada is null) and (f.NivelAprovacaoAtual = 0 or f.NivelAprovacaoAtual is null) and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null) THEN 1 ELSE 0 END) AS totalAbertas,"
    SQL = SQL & "    SUM(CASE WHEN f.Situacao = 2 and (f.Aprovada <> 1 or f.Aprovada is null) and (f.NivelAprovacaoAtual = 0 or f.NivelAprovacaoAtual is null) and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null) THEN 1 ELSE 0 END) AS totalRespondidas, "
    SQL = SQL & "    SUM(CASE WHEN f.Situacao = 2 and f.Aprovada = 1 and (f.ComSetorResponsavel = 0 or f.ComSetorResponsavel is null) THEN 1 ELSE 0 END) AS totalAprovadas  "
    SQL = SQL & "FROM FNC f "
    SQL = SQL & "WHERE ISNULL(f.Excluido, 0) = 0 "
    SQL = SQL & "AND f.DataExclusao IS NULL "
    SQL = SQL & "AND f.Planta IN ( "
    SQL = SQL & "    SELECT Planta FROM Responsavel r "
    SQL = SQL & "    WHERE r.Id = " & idUsuarioLogado & " "
    SQL = SQL & "    AND r.Tipo = 1 "
    SQL = SQL & "    AND ISNULL(r.Excluido, 0) = 0 "
    SQL = SQL & ")"

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        totalAnaliseFNC = 0
        totalAnaliseFNCAbertas = 0
        totalAnaliseFNCRespondidas = 0
        totalAnaliseFNCAprovadas = 0

        If NOT BEN_CORPORATIVO.EOF Then
            totalAnaliseFNC = BEN_CORPORATIVO("totalAbertas") + BEN_CORPORATIVO("totalRespondidas") + BEN_CORPORATIVO("totalAprovadas")
            totalAnaliseFNCAbertas = BEN_CORPORATIVO("totalAbertas")
            totalAnaliseFNCRespondidas = BEN_CORPORATIVO("totalRespondidas")
            totalAnaliseFNCAprovadas = BEN_CORPORATIVO("totalAprovadas")
            
        End If

        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAprovacoesFNC()    
    SQL = SQL & "WITH APROVADOR AS ( "
    SQL = SQL & "   SELECT  "
    SQL = SQL & "       IdSetor, "
    SQL = SQL & "       CodPlanta, "
    SQL = SQL & "       Nivel "
    SQL = SQL & "   FROM AprovadorFNCParametro "
    SQL = SQL & "   WHERE IdUsuario = " & idUsuarioLogado
    SQL = SQL & ")"

    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND F.IdSetor = A.IdSetor AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE F.DataAvaliacao IS NULL "
    SQL = SQL & "AND F.Situacao = 2 "
    
    SQL = SQL & "UNION "
   
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 2 "
    SQL = SQL & "AND F.DataAvaliacao IS NULL "
    SQL = SQL & "AND F.Situacao = 2 "
    
    SQL = SQL & "UNION "
    
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 3 "
    SQL = SQL & "AND F.DataAvaliacao IS NULL "
    SQL = SQL & "AND F.Situacao = 2 "
    
    SQL = SQL & "UNION "
    
    SQL = SQL & "SELECT distinct F.Id "
    SQL = SQL & "FROM FNC F "
    SQL = SQL & "INNER JOIN APROVADOR A ON A.CodPlanta = F.Planta AND A.Nivel = F.NivelAprovacaoAtual "
    SQL = SQL & "WHERE NivelAprovacaoAtual = 4 "
    SQL = SQL & "AND F.DataAvaliacao IS NULL "
    SQL = SQL & "AND F.Situacao = 2 "

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        totalAprovacaoFNC = 0
        While Not BEN_CORPORATIVO.EOF        
            BEN_CORPORATIVO.MoveNext()
            totalAprovacaoFNC = totalAprovacaoFNC + 1
        Wend

        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarRespostasFNC()
    Dim SQL
    
    SQL = "SELECT COUNT(*) as Total FROM FNC "
    SQL = SQL & "WHERE IdAnalista = " & idUsuarioLogado & " "
    SQL = SQL & "AND Situacao <> 3"
    SQL = SQL & "AND ComSetorResponsavel = 1"

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        If NOT BEN_CORPORATIVO.EOF Then
            totalRespostaFNC = BEN_CORPORATIVO("Total")
        End If
        
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAprovacoesOrcamento()
    Dim SQL
    
    SQL = "SELECT COUNT(DISTINCT o.NR_ORCAMENTO) as Total FROM APROVACAO_ORCAMENTO a "
    SQL = SQL & "LEFT JOIN ORCAMENTO_BRASIL o on (o.ID_ORCAMENTO = a.ID_ORCAMENTO OR o.NR_ORCAMENTO = a.NR_ORCAMENTO) "
    SQL = SQL & "WHERE a.EXCLUIDO IS NULL "
    SQL = SQL & "AND a.APROVADO = '0' "
    SQL = SQL & "AND a.PLANTA IN (" & plantasAprovadorOrcamento & ") "
    SQL = SQL & "AND a.NIVEL = " & nivelAprovadorOrcamento & " "
    SQL = SQL & "AND o.STATUS_ORCAMENTO = 'Em Aprovacao'"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalAprovacaoOrcamento = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAnalisesOrcamento()
    Dim SQL
    
    SQL = "SELECT COUNT(DISTINCT o.NR_ORCAMENTO) as Total FROM ORCAMENTO_BRASIL o "
    SQL = SQL & "WHERE o.PLANTA IN (" & plantasAnalistaOrcamento & ") "
    SQL = SQL & "AND o.STATUS_ORCAMENTO = 'Em Analise'"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalAnaliseOrcamento = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarCadastroUnificado()
    Dim SQL
    
    SQL = "SELECT COUNT(*) as Total FROM UnificadoCliente o "
    SQL = SQL & "WHERE StatusCli = '" & nivelCadastroUnificado & "' "

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If NOT PIR_CORPORATIVO.EOF Then
            totalCadastroUnificado = PIR_CORPORATIVO("Total")
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Sub ContarAlteracoesPedido()
    Dim SQL
    tiposPermitidos = ""
    filialUsuario = ""
    
    If nivelAlteracaoPedido = "V" Then
        tiposPermitidos = "1,2,6"
    Else
        tiposPermitidos = "3,4,5,7,8,9,10,14"
    End If

    If SESSION("FILIAL") = "00" Then
        filialUsuario = "I"
    ElseIf SESSION("FILIAL") = "04" Then
        filialUsuario = "U"
    Else
        filialUsuario = "X"
    End If

    nomeTabelaProtheus = ""
    If Session("Ambiente") = "HOMOLOGACAO" Then
        nomeTabelaProtheus = "DADOSADV2410"
    Else
        nomeTabelaProtheus = "DADOSADV"
    End if

    ' Alteracoes Pirapetinga/Uberaba
    SQL = "SELECT COUNT (*) as Total FROM InHistoricoAlteracoesPedidosSolicitacao a  "
    SQL = SQL & "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id  "
    SQL = SQL & "INNER JOIN " & nomeTabelaProtheus & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN "
    SQL = SQL & "INNER JOIN " & nomeTabelaProtheus & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN "
    SQL = SQL & "INNER JOIN " & nomeTabelaProtheus & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI "
    SQL = SQL & "WHERE a.Descartado <> 1 "
    SQL = SQL & "AND a.Realizado = 0 "
    SQL = SQL & "AND Data >= '20150101' "  
    SQL = sql & "AND Para <> 'L' "
    SQL = SQL & "AND a.Tipo IN (" & tiposPermitidos & ") " 
    SQL = SQL & iif(filialUsuario <> "" and nivelAlteracaoPedido = "P", "AND (C5_INPUBE = '" & filialUsuario & "' OR C5_INPUBE = '') ", "" )

    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
        
    If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
        If NOT PIR_HELPDESK.EOF Then
            totalAlteracaoPedidoPiraUbe = PIR_HELPDESK("Total")
        End If
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
    On Error Goto 0

    ' Alteracoes Outras planta
    SQL = "SELECT COUNT(*) AS Total "
    SQL = SQL & "FROM HistoricoSolicitacaoAlteracaoPedido p "
    SQL = SQL & "LEFT JOIN HistoricoAlteracaoPedido h on h.Id = p.Id "
    SQL = SQL & "WHERE 1=1"
    SQL = SQL & "AND ISNULL(DataExclusao,'') = '' "
    SQL = SQL & "AND p.IdOperacao IN (" & tiposPermitidos & ") "
    SQL = SQL & "AND h.Status IS NULL "
    SQL = SQL & iif(SESSION("ALTERAPO") <> "S", "AND Solicitante = '" & SESSION("CODIGO") & "' ", "") 
    SQL = SQL & iif(nivelAlteracaoPedido = "P", "AND p.Planta = '" & SESSION("FILIAL") & "'", "" )

    On Error Resume Next
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
        
    If Err.Number = 0 AND IsObject(BEN_CORPORATIVO) Then
        If NOT BEN_CORPORATIVO.EOF Then
            totalAlteracaoPedidoOutrasPlantas = BEN_CORPORATIVO("Total")
        End If
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
    On Error Goto 0
End Sub

Function ObterPrimeiroNome(nomeCompleto)
    Dim partes
    If nomeCompleto = "" Or IsNull(nomeCompleto) Then
        ObterPrimeiroNome = "Usuário"
        Exit Function
    End If
    partes = Split(nomeCompleto, " ")
    ObterPrimeiroNome = partes(0)
End Function

Function ObterStatusPermitidoSolicitacao(nivelUsuario, tipoAcao)
    Dim status 
    status = ""
    Select Case nivelUsuario
        Case "10"
            If tipoAcao = "ANALISE" Then
                status = "A"
            Else
                status = "D"
            End If
        Case "20"
            status = "C"
        Case "30"
            status = "E"
        Case "40"
            status = "F"
        Case "50"
            status = "G"
        Case "60"
            status = "H"
        Case Else
            status = "Desconhecido"
    End Select
    ObterStatusPermitidoSolicitacao = status
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Pendências - Painel Administrativo</title>
    <meta charset="utf-8"/>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="refresh" content="300">
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="expires" content="Mon, 06 Jan 1990 00:00:01 GMT">
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

    <script type="text/javascript">
        $(document).ready(function() {
            setInterval(() => { window.location.reload(); }, 60000);
        });

        function navegarPara(url, kind, label) {
            if (window.parent && typeof window.parent.navegarDoMid === 'function') {
                window.parent.navegarDoMid(url, kind, label);
            } else {
                window.location.href = url;
            }
        }
    </script>

    <style>
        body {
            background-color: #f8f9fa;
        }
        .notification-card {
            transition: all 0.3s ease;
            border: none;
            border-radius: 1rem;
            cursor: pointer;
            height: 100%;
            color: white;
        }
        .notification-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 24px rgba(0, 0, 0, 0.15);
        }
        .card-icon {
            width: 48px;
            height: 48px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 0.625rem;
            font-size: 1.35rem;
        }
        .count-display {
            font-size: 2.75rem;
            font-weight: 700;
            line-height: 1;
        }
        .card-description {
            opacity: 0.9;
            font-size: 0.813rem;
        }
        .card-link {
            font-size: 0.813rem;
            opacity: 0.9;
        }
        .bg-gradient-blue {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        }
        .bg-gradient-purple {
            background: linear-gradient(135deg, #a855f7 0%, #9333ea 100%);
        }
        .bg-gradient-orange {
            background: linear-gradient(135deg, #f58435 0%, #ea580c 100%);
        }
        .bg-gradient-yellow {
            background: linear-gradient(135deg, #e1b12a 0%, #bd8721 100%);
        }
        .bg-gradient-green {
            background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%);
        }
        .bg-gradient-teal {
            background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%);
        }
        .bg-gradient-red {
            background: linear-gradient(135deg, #cd5c5c 0%, #b93030 100%);
        }
        .bg-gradient-pink {
            background: linear-gradient(135deg, #df4d82 0%, #cb2260 100%);
        }
        .bg-gradient-dark-blue {
            background: linear-gradient(135deg, #685ad9 0%, #402cdb 100%);
        }
        .bg-gradient-light-blue {
            background: linear-gradient(135deg, #38abcd 0%, #3388a1 100%)
        }
        .bg-gradient-orange-pink {
            background: linear-gradient(135deg, #e97c48 0%, #f34c5a 100%);
        }
        .summary-banner {
            background: white;
            border-radius: 1rem;
            border: 1px solid #e5e7eb;
        }
        .no-permissions {
            text-align: center;
            padding: 4rem 2rem;
            background: white;
            border-radius: 1rem;
            border: 1px solid #e5e7eb;
        }
        .no-permissions i {
            font-size: 3rem;
            color: #94a3b8;
            margin-bottom: 1rem;
        }
        .no-permissions h3 {
            font-size: 1.25rem;
            margin-bottom: 0.5rem;
        }
    </style>
</head>
<body>
    <div class="container py-4">
        <div class="mb-4">
            <h1 class="h3 fw-bold text-dark mb-1">Olá, <%=ObterPrimeiroNome(nomeUsuario)%></h1>
            <p class="text-muted mb-0" style="font-size: 0.875rem;">Gerencie suas pendências e mantenha tudo em dia</p>
        </div>
 
        <% If qtdCardsVisiveis > 0 Then %>
        <div class="summary-banner p-3 mb-4 shadow-sm">
            <div class="d-flex align-items-center justify-content-between flex-wrap gap-3">
                <div>
                    <p class="text-muted small mb-1" style="font-size: 0.75rem;">Total de Pendências</p>
                    <p class="h2 fw-bold text-dark mb-0"><%=totalGeral%></p>
                </div>
            </div>
        </div>

        <div class="row g-3 g-md-4">
            <% If temPermissaoAnaliseCredito Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-orange shadow" onclick="navegarPara('/corporativo/SolicitacaoCredito/SolicitacaoCreditoGrade.asp?aba=analises', 'Serviço', 'Análises de Crédito')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-credit-card"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Análises de Crédito</h5>
                        <p class="card-description mb-3">Solicitações aguardando análise</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAnaliseCredito%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAprovacaoTransferencia Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-red shadow" onclick="navegarPara('/corporativo/SolicitacaoCredito/SolicitacaoCreditoGrade.asp?aba=transferencias', 'Serviço', 'Transferências de Crédito')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-exchange-alt"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Transferências de Crédito</h5>
                        <p class="card-description mb-3">Transferências aguardando aprovação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalTransferenciaCredito%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAprovacaoCredito Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-pink shadow" onclick="navegarPara('/corporativo/SolicitacaoCredito/SolicitacaoCreditoGrade.asp?aba=aprovacoes', 'Serviço', 'Aprovações de Crédito')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-wallet"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Aprovações de Crédito</h5>
                        <p class="card-description mb-3">Solicitações aguardando aprovação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAprovacaoCredito%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAnaliseFNC Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-yellow shadow" onclick="navegarPara('/corporativo/sgq/fnc/Principal/FNCGrade.asp', 'Serviço', 'Análise de FNC')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-clipboard-check"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Análise de FNC</h5>

                        <div class="card-description mb-3">
                            <div class="d-flex gap-2">
                                <span><%=totalAnaliseFNCAbertas%> Abertas</span>
                                <span><%=totalAnaliseFNCRespondidas%> Respondidas</span>
                                <span><%=totalAnaliseFNCAprovadas%> Aprovadas</span>
                            </div>
                        </div>
    

                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAnaliseFNC%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAprovacaoFNC Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-green shadow" onclick="navegarPara('/corporativo/sgq/fnc/Principal/FNCGrade.asp', 'Serviço', 'Aprovação de FNC')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-clipboard-check"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Aprovação de FNC</h5>
                        <p class="card-description mb-3">FNCs aguardando aprovação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAprovacaoFNC%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoRespostaFNC Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-blue shadow" onclick="navegarPara('/corporativo/sgq/fnc/Principal/FNCGrade.asp', 'Serviço', 'Respostas de FNC')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-reply"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Respostas de FNC</h5>
                        <p class="card-description mb-3">FNCs aguardando sua resposta</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalRespostaFNC%></span>
                            <span class="card-link">Ver todas <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAprovacaoOrcamento Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-purple shadow" onclick="navegarPara('/corporativo/Orcamento/OrcamentoGrade.asp', 'Serviço', 'Aprovação de Orçamento')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-circle-check"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Aprovação de Orçamento</h5>
                        <p class="card-description mb-3">Orçamentos aguardando aprovação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAprovacaoOrcamento%></span>
                            <span class="card-link">Ver todos <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoAnaliseOrcamento Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-teal shadow" onclick="navegarPara('/corporativo/Orcamento/OrcamentoGrade.asp', 'Serviço', 'Análises de Orçamento')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-file-invoice-dollar"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Análises de Orçamento</h5>
                        <p class="card-description mb-3">Orçamentos aguardando análise</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAnaliseOrcamento%></span>
                            <span class="card-link">Ver todos <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoCadastroUnificado Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-dark-blue shadow" onclick="navegarPara('/corporativo/Unificado/index.asp', 'Serviço', 'Cadastro Unificado')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-file-invoice-dollar"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Cadastro Unificado</h5>
                        <p class="card-description mb-3">Cadastros de cliente aguardando liberação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalCadastroUnificado%></span>
                            <span class="card-link">Ver todos <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% 
                urlSolicitacaoAlteracaoPedido = ""
                If nivelAlteracaoPedido = "P" Then
                    ' PCP Visualiza apenas solicitações da planta dele
                    If SESSION("FILIAL") = "00" OR SESSION("FILIAL") = "04" Then
                        urlSolicitacaoAlteracaoPedido = "/corporativo/frameset/sfa/pedidos/alteracoes/SolicitaAlteracao.asp"
                    Else
                        urlSolicitacaoAlteracaoPedido = "/corporativo/comercial/Brasil/alteracaopedido/SelecionaPlantaAlteracaoPedido.asp"
                    End If
                Else
                    ' Validador visualiza solicitações de todas as plantas
                    If totalAlteracaoPedidoPiraUbe > 0 Then
                        urlSolicitacaoAlteracaoPedido = "/corporativo/frameset/sfa/pedidos/alteracoes/SolicitaAlteracao.asp"
                    Else
                        urlSolicitacaoAlteracaoPedido = "/corporativo/comercial/Brasil/alteracaopedido/SelecionaPlantaAlteracaoPedido.asp"
                    End If
                End If
            %>
            
            <% If temPermissaoAlteracaoPedido Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-light-blue shadow" onclick="navegarPara('<%=urlSolicitacaoAlteracaoPedido%>', 'Serviço', 'Alteração de Pedidos')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-file-invoice-dollar"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Alteração de Pedidos</h5>
                        <p class="card-description mb-3">Solicitações de alteração aguardando aprovação</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalAlteracaoPedidoOutrasPlantas + totalAlteracaoPedidoPiraUbe%></span>
                            <span class="card-link">Ver todos <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>

            <% If temPermissaoChamados Then %>
            <div class="col-12 col-md-6 col-lg-4">
                <div class="card notification-card bg-gradient-orange-pink shadow" onclick="navegarPara('/corporativo/chamados/Chamado/ChamadoFilaAtendimento.asp', 'Serviço', 'Chamados')">
                    <div class="card-body p-3">
                        <div class="d-flex align-items-start justify-content-between mb-2">
                            <div class="card-icon">
                                <i class="fas fa-credit-card"></i>
                            </div>
                        </div>
                        <h5 class="card-title fw-semibold mb-1" style="font-size: 0.95rem;">Chamados</h5>
                        <p class="card-description mb-3">Chamados Aguardando Resposta</p>
                        <div class="d-flex align-items-end justify-content-between">
                            <span class="count-display"><%=totalChamados%></span>
                            <span class="card-link">Ver todos <i class="fas fa-arrow-right ms-1"></i></span>
                        </div>
                    </div>
                </div>
            </div>
            <% End If %>
        </div>
       
        <% Else %>
        <div class="no-permissions">
            <h6 class="mb-0">Nenhuma pendência</h6>
        </div>
        <% End If %>
    </div>
</body>
</html>
