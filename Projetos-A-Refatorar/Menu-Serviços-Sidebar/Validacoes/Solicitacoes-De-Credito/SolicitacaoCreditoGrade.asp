<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<!--#include file="./util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim solicitacoes()
Dim transferencias()

'Variaveis auxiliares'
Dim idx
idx = 0
totalSolicitacoes = 0
totalTransferencias = 0

Dim idUsuarioLogado
idUsuarioLogado = Session("CODIGO")

'----------------------------------------------
'         Carrega Grade principal
'----------------------------------------------

ReDim solicitacoes(0)
ReDim transferencias(0)
idx = 0

' Verifica se usuário é validador
Dim validador, codNivel, minimo, maximo, plantas
validadorEncontrado = False

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")


' Contadores
Dim totalAnaliseCredito, totalAprovacaoCredito, totalTransferenciaCredito
totalAnaliseCredito = 0
totalAprovacaoCredito = 0
totalTransferenciaCredito = 0

SQL = "SELECT DISTINCT "
SQL = SQL & " CAST(vc.IdValidador AS INT) AS IdValidador,"
SQL = SQL & " u.CRM_NOME AS Nome,"
SQL = SQL & " MIN(vc.Minimo) AS Minimo,"
SQL = SQL & " MAX(vc.Maximo) AS Maximo,"
SQL = SQL & " CAST(vc.Nivel AS NVARCHAR(10)) as CodNivel,"
SQL = SQL & " STUFF(("
SQL = SQL & "     SELECT ',' + vc2.Planta"
SQL = SQL & "     FROM ValidadoresCredito vc2"
SQL = SQL & "     INNER JOIN HELPDESK..Empresas e2 ON vc2.Planta = e2.Empresa"
SQL = SQL & "     WHERE vc2.IdValidador = vc.IdValidador"
SQL = SQL & "     AND vc2.DataExclusao IS NULL"
SQL = SQL & "     ORDER BY vc2.Planta"
SQL = SQL & "     FOR XML PATH(''), TYPE"
SQL = SQL & " ).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS Plantas"
SQL = SQL & " FROM ValidadoresCredito vc"
SQL = SQL & " LEFT JOIN HELPDESK..CRM_DTI u ON vc.IdValidador = u.CRM_R_E_C_N_O"
SQL = SQL & " WHERE vc.IdValidador = " & idUsuarioLogado
SQL = SQL & " AND vc.DataExclusao is NULL"
SQL = SQL & " GROUP BY vc.IdValidador, u.CRM_NOME, vc.Nivel"
SQL = SQL & " ORDER BY Nome"

On Error Resume Next
SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)

If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
    If Not PIR_CORPORATIVO.EOF Then
        plantas = PIR_CORPORATIVO("Plantas")
        codNivel = PIR_CORPORATIVO("codNivel")
        minimo = PIR_CORPORATIVO("Minimo")
        maximo = PIR_CORPORATIVO("Maximo")
        validadorEncontrado = True
    End If
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If
On Error Goto 0

' Aba selecionada (apenas para analista de crédito)
Dim abaSelecionada
abaSelecionada = Request.QueryString("aba")
If abaSelecionada = "" Then
    abaSelecionada = "analises"
End If

Dim filtroPlanta
filtroPlanta = Request.Form("filtroPlanta")
If filtroPlanta <> "" And filtroPlanta <> "TODOS" Then
    plantas = filtroPlanta
End If

Dim statusSelecionado
If codNivel = "10" Then
    Select Case abaSelecionada
        Case "analises"
            statusSelecionado = "A"
        Case "aprovacoes"
            statusSelecionado = "D"
        Case "transferencias"
            statusSelecionado = "TRANSFERENCIA"
        Case Else
            statusSelecionado = "A"
    End Select
Else
    ' Outros níveis sempre usam o status definido pelo nível
    statusSelecionado = ObterStatusPermitidoParaUsuario(codNivel, "")
End If

If codNivel = "10" Then
    Call ContarAnalisesCredito()
    Call ContarAprovacoesCredito()
    Call ContarTransferenciasCredito()
End If

If validadorEncontrado Then
    ' Carrega Transferências de Crédito (apenas para analista de crédito)
    If abaSelecionada = "transferencias" And codNivel = "10" Then
        SQL = "SELECT "
        SQL = SQL & "Id, NomeCliente, NomeSolicitante, NomeAprovador, "
        SQL = SQL & "Status, LimiteAtualOrigem, LimiteAtualDestino, ValorTransferencia, PlantaOrigem, "
        SQL = SQL & "CodClienteOrigem, CodLojaOrigem, PlantaDestino, CodClienteDestino, "
        SQL = SQL & "CodLojaDestino, Observacao, FlagInativo, CNPJCPF as CNPJ, "
        SQL = SQL & "CASE "
        SQL = SQL & "    WHEN ISDATE(DataCriacao) = 1 "
        SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, DataCriacao, 120), 126) "
        SQL = SQL & "    ELSE NULL "
        SQL = SQL & "END AS DataCriacao "
        SQL = SQL & "FROM TransferenciaCredito "
        SQL = SQL & "WHERE FlagInativo = 0 "
        SQL = SQL & "AND PlantaDestino IN (" & plantas & ") "
        SQL = SQL & "AND Status = 'PENDENTE' "
        SQL = SQL & "ORDER BY Id DESC"
        
        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        idx = 0
        While NOT PIR_CORPORATIVO.EOF
            ReDim Preserve transferencias(idx)
            Set transferencias(idx) = Server.CreateObject("Scripting.Dictionary")
            transferencias(idx)("Id") = PIR_CORPORATIVO("Id")
            transferencias(idx)("NomeCliente") = PIR_CORPORATIVO("NomeCliente")
            transferencias(idx)("CNPJ") = PIR_CORPORATIVO("CNPJ")
            transferencias(idx)("NomeSolicitante") = PIR_CORPORATIVO("NomeSolicitante")
            transferencias(idx)("LimiteAtualOrigem") = PIR_CORPORATIVO("LimiteAtualOrigem")
            transferencias(idx)("LimiteAtualDestino") = PIR_CORPORATIVO("LimiteAtualDestino")
            transferencias(idx)("ValorTransferencia") = PIR_CORPORATIVO("ValorTransferencia")
            transferencias(idx)("PlantaOrigem") = PIR_CORPORATIVO("PlantaOrigem")
            transferencias(idx)("PlantaDestino") = PIR_CORPORATIVO("PlantaDestino")
            transferencias(idx)("DataCriacao") = PIR_CORPORATIVO("DataCriacao")
            transferencias(idx)("Status") = PIR_CORPORATIVO("Status")
            idx = idx + 1
            PIR_CORPORATIVO.MoveNext()
        Wend
        If isObject(PIR_CORPORATIVO) Then
            PIR_CORPORATIVO.Close()
            SET PIR_CORPORATIVO = NOTHING
        End If
        totalTransferencias = idx
    Else
        ' Carrega Solicitações de Crédito normais
        Dim statusValidacao
        statusValidacao = ObterStatusPermitidoParaUsuario(codNivel, statusSelecionado)
        
        SQL = "SELECT DISTINCT " 
        SQL = SQL & "s.NUM as IdSolicitacao, "
        SQL = SQL & "s.LOGININPANET as LoginSolicitante, "
        SQL = SQL & "s.SOLICITANTE as NomeSolicitante, "
        SQL = SQL & "s.NomeCliente, "
        SQL = SQL & "s.COD as CodCliente, "
        SQL = SQL & "s.CODLOJA as CodLoja, "
        SQL = SQL & "s.Planta as CodPlanta, "
        SQL = SQL & "s.LCATUAL as LimiteAtual, "
        SQL = SQL & "s.LCSOLICIT as LimiteSolicitado, "
        SQL = SQL & "s.LCNOVO as LimiteNovo, "
        SQL = SQL & "s.LCVALID as LimiteValidado, "
        SQL = SQL & "s.STATUS as Status, "
        SQL = SQL & "CASE " 
        SQL = SQL & "    WHEN ISDATE(s.DTVALID) = 1 "
        SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTVALID, 112), 126) "
        SQL = SQL & "    ELSE NULL "
        SQL = SQL & "END AS DataValidacao, "
        SQL = SQL & "CASE "
        SQL = SQL & "    WHEN ISDATE(s.DTSOLICIT) = 1 "
        SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTSOLICIT, 112), 126) "
        SQL = SQL & "    ELSE NULL "
        SQL = SQL & "END AS DataCadastro "
        SQL = SQL & "FROM SOLCREDITOCLIENTE s " 
        SQL = SQL & "INNER JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = LOGININPANET COLLATE SQL_Latin1_General_CP437_BIN AND ISNULL(CRM_D_E_L_E_T,'') = ''"
        SQL = SQL & "WHERE s.STATUS = '" & statusValidacao & "' "
        SQL = SQL & "AND s.Planta IN (" & plantas & ") "
        SQL = SQL & "AND (CASE WHEN s.LCVALID = 0 THEN s.LCSOLICIT ELSE s.LCVALID END) >= " & minimo & " "
        SQL = SQL & "ORDER BY NUM DESC"

        SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
        idx = 0
        While NOT PIR_CORPORATIVO.EOF
            ReDim Preserve solicitacoes(idx)
            Set solicitacoes(idx) = Server.CreateObject("Scripting.Dictionary")
            solicitacoes(idx)("IdSolicitacao") = PIR_CORPORATIVO("IdSolicitacao")
            solicitacoes(idx)("CodPlanta") = PIR_CORPORATIVO("CodPlanta")
            solicitacoes(idx)("LoginSolicitante") = PIR_CORPORATIVO("LoginSolicitante")
            solicitacoes(idx)("NomeSolicitante") = PIR_CORPORATIVO("NomeSolicitante")
            solicitacoes(idx)("NomeCliente") = PIR_CORPORATIVO("NomeCliente")
            solicitacoes(idx)("LimiteAtual") = PIR_CORPORATIVO("LimiteAtual")
            solicitacoes(idx)("LimiteNovo") = PIR_CORPORATIVO("LimiteNovo")
            solicitacoes(idx)("LimiteValidado") = PIR_CORPORATIVO("LimiteValidado")
            solicitacoes(idx)("LimiteSolicitado") = PIR_CORPORATIVO("LimiteSolicitado")
            solicitacoes(idx)("Status") = PIR_CORPORATIVO("Status")
            solicitacoes(idx)("DataCadastro") = PIR_CORPORATIVO("DataCadastro")
            solicitacoes(idx)("CodCliente") = PIR_CORPORATIVO("CodCliente")
            solicitacoes(idx)("CodLoja") = PIR_CORPORATIVO("CodLoja")
            idx = idx + 1
            PIR_CORPORATIVO.MoveNext()
        Wend
        If isObject(PIR_CORPORATIVO) Then
            PIR_CORPORATIVO.Close()
            SET PIR_CORPORATIVO = NOTHING
        End If
        totalSolicitacoes = idx
    End If
End If

Function ObterStatusPermitidoParaUsuario(nivelUsuario, statusSelecionado)
    Dim status 
    status = ""
    Select Case nivelUsuario
        Case "10"
            ' Para nível 10 (Analista), retorna o status selecionado pelo usuário
            status = statusSelecionado
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
    
    ObterStatusPermitidoParaUsuario = status
End Function

Sub ContarAnalisesCredito()
    Dim SQL
    
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM SOLCREDITOCLIENTE "
    SQL = SQL & "INNER JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = LOGININPANET COLLATE SQL_Latin1_General_CP437_BIN AND ISNULL(CRM_D_E_L_E_T,'') = '' "
    SQL = SQL & "WHERE STATUS = 'A' "
    SQL = SQL & "AND Planta IN (" & plantas & ") "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) >= " & minimo & " "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) <= " & maximo & " "
    
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

Sub ContarAprovacoesCredito()
    Dim SQL
    
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM SOLCREDITOCLIENTE "
    SQL = SQL & "INNER JOIN HELPDESK..CRM_DTI ON CRM_LOGIN = LOGININPANET COLLATE SQL_Latin1_General_CP437_BIN AND ISNULL(CRM_D_E_L_E_T,'') = '' "
    SQL = SQL & "WHERE STATUS = 'D' "
    SQL = SQL & "AND Planta IN (" & plantas & ") "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) >= " & minimo & " "
    SQL = SQL & "AND (CASE WHEN LCVALID = 0 THEN LCSOLICIT ELSE LCVALID END) <= " & maximo & " "
    
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
    Dim SQL
    
    SQL = "SELECT COUNT(*) as Total "
    SQL = SQL & "FROM TransferenciaCredito "
    SQL = SQL & "WHERE Status = 'PENDENTE' "
    SQL = SQL & "AND PlantaDestino IN (" & plantas & ") "
    SQL = SQL & "AND FlagInativo = 0 "
    
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
%>


<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="./util/Util.js"></script>
        <script type="text/javascript">
            function plantaAlterada() {
                $('#Pesquisar').submit();
            }
            function validar(idSolicitacao) {
                window.location.href = 'SolicitacaoCreditoValidar.asp?flagRejeitar=0&idSolicitacao=' + idSolicitacao;
            }

            function alterarAba(aba) {
                window.location.href = window.location.pathname + '?aba=' + aba;
            }

            function validarTransferencia(idTransferencia) {
                window.location.href = 'Transferencia/TransferenciaCreditoValidar.asp?flagRejeitar=0&idTransferencia=' + idTransferencia;
            }

            function rejeitarTransferencia(idTransferencia) {
                window.location.href = 'Transferencia/TransferenciaCreditoValidar.asp?flagRejeitar=1&idTransferencia=' + idTransferencia;
            }

            function abrirHistoricoCliente(codCliente) {
                var url = `/corporativo/SolicitacaoCredito/SolicitacaoCreditoRelatorio.asp?aba=solicitacoes&filtroCodCliente=${codCliente}`;
                window.open(url, '', "width=1200,height=800,scrollbars=yes,resizable=yes");
            }
        </script>
        <style type="text/css">
            body{
                background-color: #f5f5f5;
                font-size: .9rem;
            }
            .card {
                padding: 0rem;
            }
            .form-label {
                font-weight: 600;
            }
            .filters-card {
                background-color: white;
                border-radius: 8px;
                padding: 15px;
                margin-bottom: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            .table-wrapper {
                overflow: auto;
                max-height: 80vh;
            }
            thead {
                vertical-align: middle !important;
                position: sticky; 
                inset-block-start: 0; 
                background: #f5f5f5;
            }
            td, th {
                padding: 8px 12px !important;
            }
            th {
                font-weight: 600;
                color: #495057;
            }
            tr:hover {
                background: #f5f5f5;
            }
            .no-records {
                padding: 10px;
                text-align: center;
                background: white;
            }
            .status-badge {
                display: inline-block;
                text-align: center;
                padding: 6px 12px;
                border-radius: 6px;
                font-size: 0.75rem;
                font-weight: 600;
            }
            .badge-novo {
                background-color: #dbeafe;
                color: #1e40af;
            }
            .badge-rejeitado {
                background-color: #fee2e2;
                color: #991b1b;
            }
            .badge-aprovado {
                background-color: #d1fae5;
                color: #065f46;
            }        
            .badge-validado-analista {
                background-color: #e9d5ff;
                color: #6b21a8;
            }
            .badge-analise-finalizada {
                background-color: #fed7aa;
                color: #c2410c;
            }
            .badge-validado-coordenador {
                background-color: #cffafe;
                color: #155e75;
            }
            .badge-validado-gerente-tesouraria {
                background-color: #e0e7ff;
                color: #3730a3;
            }
            .badge-validado-controller {
                background-color: #fce7f3;
                color: #9f1239;
            }
            .badge-validado-gerente-planta {
                background-color: #fef3c7;
                color: #92400e;
            }
            .badge-default {
                background-color: #f3f4f6;
                color: #374151;
            }
            .acesso-negado {
                background-color: #fff3cd;
                border: 1px solid #ffc107;
                color: #856404;
                padding: 20px;
                border-radius: 8px;
                text-align: center;
                margin: 40px auto;
                max-width: 600px;
            }
            .acesso-negado i {
                font-size: 3rem;
                margin-bottom: 15px;
            }
            .acesso-negado h4 {
                margin-bottom: 10px;
            }
            .tabs-container {
                display: flex;
                padding: 0;
                gap: 0;
                min-width: max-content;
                background-color: #f8f9fa;
                border-radius: 8px 8px 0 0;
            }
            .tab-button {
                flex: 0 0 auto; 
                padding: 15px 20px;
                background: transparent;
                border: none;
                border-bottom: 3px solid transparent;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: baseline;
                gap: 8px;
                color: #6c757d;
                font-weight: 500;
            }
            .tab-button:hover {
                color: #212529;
            }
            .tab-button.active {
                color: #0d6efd;
                border-bottom-color: #0d6efd;
                background-color: #fff;
            }
            .tab-total {
                display: flex;
                justify-content: center;
                align-items: center;
                background-color: #8d8d8dff;
                color: #fff;
                border-radius: 50%;
                height: 18px;
                width: 18px;
                font-size: 0.75rem;
                margin-left: 10px;
            }
            .tab-total.active {
                background-color: #0d6efd;
            }
        </style>
    </head>
    <body>
        <% If Not validadorEncontrado Then %>
            <div class="container">
                <div class="acesso-negado">
                    <i class="fas fa-exclamation-triangle"></i>
                    <h4>Acesso Negado</h4>
                    <p>Você não possui permissão para acessar esta página.</p>
                </div>
            </div>
        <% Else %>
            <div class="container">
                <h3 class="mt-4 mb-4">Solicitações de Crédito</h3>

                <div class="filters-card">
                    <form method="post" action="" id="Pesquisar">
                        <div class="row">
                            <div class="col-md-3 mb-3">
                                <label for="filtroPlanta" class="form-label">Planta</label>
                                <select onchange="plantaAlterada()" class="form-select form-select-sm" id="filtroPlanta" name="filtroPlanta">
                                    <option value="TODOS" <%=IIf(filtroPlanta = "TODOS" Or filtroPlanta = "", "selected", "")%>>Todas</option>
                                    <option value="00" <%=IIf(filtroPlanta = "00", "selected", "")%>>Pirapetinga</option>
                                    <option value="04" <%=IIf(filtroPlanta = "04", "selected", "")%>>Uberaba</option>
                                    <option value="05" <%=IIf(filtroPlanta = "05", "selected", "")%>>Bento</option>
                                    <option value="06" <%=IIf(filtroPlanta = "06", "selected", "")%>>Maranguape</option>
                                    <option value="07" <%=IIf(filtroPlanta = "07", "selected", "")%>>Saquarema</option>
                                    <option value="08" <%=IIf(filtroPlanta = "08", "selected", "")%>>Aracatuba</option>
                                    <option value="09" <%=IIf(filtroPlanta = "09", "selected", "")%>>Blumenau</option>
                                    <option value="10" <%=IIf(filtroPlanta = "10", "selected", "")%>>Pacajus</option>
                                    <option value="11" <%=IIf(filtroPlanta = "11", "selected", "")%>>Porto Feliz</option>
                                    <option value="12" <%=IIf(filtroPlanta = "12", "selected", "")%>>Tres Barras</option>
                                    <option value="13" <%=IIf(filtroPlanta = "13", "selected", "")%>>Florestal</option>
                                </select>
                            </div>
                        </div>
                    </form>
                </div>
                
                
                <!-- Analista pode navegar entre três abas, outros níveis tem acesso apenas uma aba sem navegação -->
                <% If codNivel = "10" Then %>
                    <div class="tabs-container">
                        <button type="button" class="tab-button <%= IIf(abaSelecionada = "analises", "active", "") %>" onclick="alterarAba('analises')">
                            <i class="fa fa-list me-2"></i>
                            <span class="tab-title">Análises Pendentes</span>
                            <div class="tab-total <%= IIf(abaSelecionada = "analises", "active", "") %>"> 
                                <%= totalAnaliseCredito %> 
                            </div>
                        </button>
                        <button type="button" class="tab-button <%= IIf(abaSelecionada = "aprovacoes", "active", "") %>" onclick="alterarAba('aprovacoes')">
                            <i class="far fa-clock me-2"></i>
                            <span class="tab-title">Aprovações Pendentes</span>
                            <div class="tab-total <%= IIf(abaSelecionada = "aprovacoes", "active", "") %>"> 
                                <%= totalAprovacaoCredito %> 
                            </div>
                        </button>
                        <button type="button" class="tab-button <%= IIf(abaSelecionada = "transferencias", "active", "") %>" onclick="alterarAba('transferencias')">
                            <i class="fas fa-exchange-alt me-2"></i>
                            <span class="tab-title">Transferências Pendentes</span>
                            <div class="tab-total <%= IIf(abaSelecionada = "transferencias", "active", "") %>"> 
                                <%= totalTransferenciaCredito %> 
                            </div>
                        </button>
                    </div>
                <% End If %>

                <div class="card">
                    <div class="table-wrapper">
                        <% If abaSelecionada = "transferencias" And codNivel = "10" Then %>
                            <table class="table" name="GradeTransferencias" id="GradeTransferencias">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Solicitante</th>
                                        <th>Cliente</th>
                                        <th>CNPJ</th>
                                        <th>Planta Origem</th>
                                        <th>Planta Destino</th>
                                        <th>Limite Anterior Origem</th>
                                        <th>Limite Anterior Destino</th>
                                        <th>Valor Transferência</th>
                                        <th>Status</th>
                                        <th>Data de Solicitação</th>
                                        <th style="width: 150px;">Ações</th>
                                    </tr>              
                                </thead>

                                <% IF totalTransferencias > 0 THEN %>
                                <tbody>
                                    <% 
                                    Dim t
                                    For t = 0 To UBound(transferencias)
                                    %>
                                        <tr>   
                                            <td class="text-left"><%=transferencias(t)("Id")%></td>
                                            <td class="text-left"><%=transferencias(t)("NomeSolicitante")%></td>
                                            <td class="text-left"><%=transferencias(t)("CodCliente") & " - " & transferencias(t)("NomeCliente")%></td>
                                            <td class="text-left"><%=FormataCnpj(transferencias(t)("CNPJ"))%></td>
                                            <td class="text-left"><%=ObterNomePlanta(transferencias(t)("PlantaOrigem"))%></td>
                                            <td class="text-left"><%=ObterNomePlanta(transferencias(t)("PlantaDestino"))%></td>
                                            <td class="text-right"><%=FormatarMoeda(transferencias(t)("LimiteAtualOrigem"))%></td>
                                            <td class="text-right"><%=FormatarMoeda(transferencias(t)("LimiteAtualDestino"))%></td>
                                            <td class="text-right"><%=FormatarMoeda(transferencias(t)("ValorTransferencia"))%></td>
                                            <td class="text-left">
                                                <span class="status-badge badge-novo">
                                                    Novo
                                                </span>
                                            </td>
                                            <td class="text-left"><%=FormatarData(transferencias(t)("DataCriacao"))%></td>
                                            <td class="text-left">
                                                <div class="d-flex gap-1">
                                                    <button type="button" 
                                                        class="btn btn-sm btn-success" 
                                                        title="Validar Transferência" 
                                                        onclick="validarTransferencia(<%=transferencias(t)("Id")%>)">
                                                        <i class="fa fa-check"></i>
                                                    </button>
                                                    <button type="button" 
                                                        class="btn btn-sm btn-danger" 
                                                        title="Rejeitar Transferência" 
                                                        onclick="rejeitarTransferencia(<%=transferencias(t)("Id")%>)">
                                                        <i class="fa fa-ban"></i>
                                                    </button>
                                                </div>
                                            </td>
                                        </tr>
                                    <% Next %>
                                </tbody>
                                <% END IF %>
                            </table>
                            <% IF totalTransferencias = 0 THEN %>
                                <div class="no-records">
                                    Nenhum registro encontrado
                                </div>
                            <% END IF %>
                        <% Else %>
                            <table class="table" name="GradeSolicitacoesCredito" id="GradeSolicitacoesCredito">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Planta</th>
                                        <th>Solicitante</th>
                                        <th>Cliente</th>
                                        <th>Limite Atual</th>
                                        <th>Limite Solicitado</th>
                                        <th>Limite Validado</th>
                                        <th>Status</th>
                                        <th>Data de Solicitação</th>
                                        <th style="width: 100px;">Ações</th>
                                    </tr>              
                                </thead>

                                <% IF totalSolicitacoes > 0 THEN %>
                                <tbody>
                                    <% 
                                    Dim i
                                    For i = 0 To UBound(solicitacoes)
                                    %>
                                        <tr>   
                                            <td class="text-left"><%=solicitacoes(i)("IdSolicitacao")%></td>
                                            <td class="text-left"><%=ObterNomePlanta(solicitacoes(i)("CodPlanta"))%></td>
                                            <td class="text-left"><%=solicitacoes(i)("NomeSolicitante")%></td>
                                            <td class="text-left"><%=solicitacoes(i)("CodCliente") & " - " & solicitacoes(i)("NomeCliente")%></td>
                                            <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteAtual"))%></td>
                                            <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteSolicitado"))%></td>
                                            <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteValidado"))%></td>
                                            <td class="text-left">
                                                <span class="status-badge <%=ObterClasseStatus(solicitacoes(i)("Status"))%>">
                                                    <%=ObterNomeStatus(solicitacoes(i)("Status"))%>
                                                </span>
                                            </td>
                                            <td class="text-left"><%=FormatarData(solicitacoes(i)("DataCadastro"))%></td>
                                            <td class="text-left">
                                                <div class="d-flex gap-1">
                                                    <button type="button" 
                                                        class="btn btn-sm btn-primary" 
                                                        title="Validar Solicitacao" 
                                                        onclick="validar(<%=solicitacoes(i)("IdSolicitacao")%>)">
                                                        <i class="fa fa-search-plus"></i>
                                                    </button>
                                                    <button type="button" class="btn btn-sm btn-secondary btn-outlined" onclick="abrirHistoricoCliente(<%= "'" & solicitacoes(i)("CodCliente") & "'" %>)" title="Ver Histórico do cliente">
                                                        <i class="fas fa-history"></i>
                                                    </button>
                                                </div>
                                            </td>
                                        </tr>
                                    <% Next %>
                                </tbody>
                                <% END IF %>
                            </table>
                            <% IF totalSolicitacoes = 0 THEN %>
                                <div class="no-records">
                                    Nenhum registro encontrado
                                </div>
                            <% END IF %>
                        <% End If %>
                    </div>
                </div>
            </div>
        <% End If %>
    </body>
</html>
