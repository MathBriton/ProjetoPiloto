<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../cnn2/conexoes.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file = "Classes/clsAspUtil.asp"-->

<%
response.Expires = -1
response.ExpiresAbsolute = Now() -1
response.AddHeader "pragma", "no-store"
response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
session.LCID = 1046

Dim objUtil
Set objUtil = New AspUtil

acao   = IIf( Trim(Request("acao"))   = "", "BuscaSolicitacoes", Trim(Request("acao")) )
filtro = Trim(Request("filtro"))
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <title>SWnet - Solicitação de Alteração de Pedidos</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />

    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
    /* ── Tokens de Design ─────────────────────────────────────────────── */
    :root {
        --color-bg:            hsl(240, 9%, 98%);
        --color-surface:       hsl(0, 0%, 100%);
        --color-surface-muted: hsl(240, 9%, 97%);
        --color-border:        hsl(240, 6%, 90%);
        --color-border-strong: hsl(240, 6%, 82%);

        --color-text:          hsl(240, 6%, 10%);
        --color-text-muted:    hsl(240, 4%, 46%);
        --color-text-faint:    hsl(240, 4%, 64%);

        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        --color-success:       hsl(145, 63%, 32%);
        --color-success-soft:  hsl(145, 60%, 95%);
        --color-success-row:   hsl(145, 55%, 88%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);
        --color-danger-row:    hsl(3, 80%, 88%);
        --color-warning:       hsl(35, 100%, 48%);
        --color-warning-soft:  hsl(40, 100%, 95%);
        --color-warning-row:   hsl(40, 90%, 82%);
        --color-accent-row:    hsl(210, 75%, 85%);
        --color-neutral-row:   hsl(240, 6%, 86%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
        --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

        --radius-sm:  6px;
        --radius-md:  10px;
        --radius-lg:  14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);
        --shadow-lg: 0 12px 32px hsla(240, 6%, 10%, .12);

        --transition: 160ms ease;

        --space-1: 4px;  --space-2: 8px;  --space-3: 12px;
        --space-4: 16px; --space-5: 24px; --space-6: 32px; --space-7: 48px;

        --z-dropdown: 1000; --z-sticky: 1020; --z-overlay: 1030;
        --z-modal: 1050;    --z-toast: 1100;
    }

    /* ── Base ─────────────────────────────────────────────────────────── */
    body {
        background-color: var(--color-bg);
        min-height: 100vh;
        font-family: var(--font-sans);
        color: var(--color-text);
    }

    .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
    }

    /* ── Card de Filtros ──────────────────────────────────────────────── */
    .filter-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
    }
    .filter-card-header {
        background-color: var(--color-surface);
        color: var(--color-text);
        padding: .7rem 1.1rem;
        border-radius: var(--radius-md) var(--radius-md) 0 0;
        border-bottom: 1px solid var(--color-border);
        font-weight: 600;
        font-size: .85rem;
    }
    .filter-card-body {
        padding: 1.1rem 1.25rem 1.35rem;
    }

    .filter-card-body label {
        font-size: .78rem;
        font-weight: 500;
        color: var(--color-text-muted);
        letter-spacing: .01em;
        margin-bottom: .3rem;
    }

    .form-control:focus,
    .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px var(--color-accent-soft);
    }

    /* ── Botões ───────────────────────────────────────────────────────── */
    .btn-primary {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
        font-weight: 500;
        font-size: .85rem;
        transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-primary:hover {
        background-color: var(--color-accent-hover);
        border-color: var(--color-accent-hover);
    }
    .btn-primary:active { transform: scale(.98); }

    .btn-light {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        color: var(--color-text);
        font-weight: 500;
        font-size: .85rem;
        transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-light:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
    }

    .btn-danger {
        font-weight: 500;
        font-size: .82rem;
        transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-danger:active { transform: scale(.98); }

    /* ── Botões de navegação com aparência de badge ───────────────────── */
    .btn-aba {
        font-size: .78rem;
        font-weight: 600;
        padding: .3rem .75rem;
        border-radius: var(--radius-sm);
        border: none;
        letter-spacing: .01em;
        cursor: pointer;
        transition: filter var(--transition), transform var(--transition);
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: .3rem;
    }
    .btn-aba:active { transform: scale(.97); }

    .btn-aba-pendente {
        background-color: var(--color-warning-soft);
        color: var(--color-warning);
    }
    .btn-aba-pendente:hover  { filter: brightness(.94); color: var(--color-warning); }

    .btn-aba-realizada {
        background-color: var(--color-success-soft);
        color: var(--color-success);
    }
    .btn-aba-realizada:hover { filter: brightness(.94); color: var(--color-success); }

    .btn-aba-descartada {
        background-color: var(--color-danger-soft);
        color: var(--color-danger);
    }
    .btn-aba-descartada:hover { filter: brightness(.94); color: var(--color-danger); }

    .btn-aba-registrar {
        background-color: var(--color-accent-soft);
        color: var(--color-accent);
    }
    .btn-aba-registrar:hover { filter: brightness(.94); color: var(--color-accent); }

    /* estado ativo (aba selecionada) — borda evidente sem alterar cor */
    .btn-aba.aba-ativa {
        box-shadow: 0 0 0 2px currentColor;
    }

    /* ── Card de Tabela ───────────────────────────────────────────────── */
    .table-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        position: relative;
    }

    /* ── Toolbar da Tabela ────────────────────────────────────────────── */
    .table-toolbar {
        background-color: var(--color-surface);
        border-radius: var(--radius-md) var(--radius-md) 0 0;
    }

    .badge-status {
        font-size: .75rem;
        font-weight: 600;
        padding: .3rem .65rem;
        border-radius: var(--radius-sm);
        letter-spacing: .01em;
    }
    .badge-status-pendente {
        background-color: var(--color-warning-soft);
        color: var(--color-warning);
    }
    .badge-status-realizada {
        background-color: var(--color-success-soft);
        color: var(--color-success);
    }
    .badge-status-descartada {
        background-color: var(--color-danger-soft);
        color: var(--color-danger);
    }

    /* ── Wrapper + Tabela ─────────────────────────────────────────────── */
    .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
    }
    .table-wrapper table { min-width: max-content; }

    .table thead th {
        position: sticky;
        top: 0;
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        font-size: .76rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .03em;
        text-align: center;
        white-space: nowrap;
        z-index: 2;
        border-color: var(--color-border);
        padding: .6rem .75rem;
    }
    .table tbody td {
        font-size: .85rem;
        text-align: center;
        white-space: nowrap;
        color: var(--color-text);
        vertical-align: middle;
        padding: .55rem .75rem;
    }
    .table tbody tr:not([class*="row-"]):hover td {
        background-color: var(--color-surface-muted);
    }

    /* Classes de estado de linha */
    .table > tbody > tr.row-success > td { background-color: var(--color-success-row); }
    .table > tbody > tr.row-danger  > td { background-color: var(--color-danger-row); }
    .table > tbody > tr.row-warning > td { background-color: var(--color-warning-row); }
    .table > tbody > tr.row-info    > td { background-color: var(--color-accent-row); }
    .table > tbody > tr.row-muted   > td { background-color: var(--color-neutral-row); }

    .table > tbody > tr.row-success:hover > td,
    .table > tbody > tr.row-danger:hover  > td,
    .table > tbody > tr.row-warning:hover > td,
    .table > tbody > tr.row-info:hover    > td,
    .table > tbody > tr.row-muted:hover   > td { filter: brightness(.95); }

    /* ── Placeholder (estado vazio) ───────────────────────────────────── */
    .table-placeholder {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 4rem 1rem;
        color: var(--color-text-muted);
        gap: .85rem;
    }
    .table-placeholder .bi { font-size: 2.4rem; color: var(--color-text-faint); }
    .table-placeholder p   { margin: 0; font-size: .88rem; text-align: center; }

    /* ── Overlay de carregamento ──────────────────────────────────────── */
    .overlay-loading {
        position: absolute;
        inset: 0;
        background: hsla(0, 0%, 100%, .82);
        backdrop-filter: blur(2px);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 10;
        gap: .85rem;
        border-radius: var(--radius-md);
    }
    .overlay-loading .spinner-border {
        width: 1.75rem;
        height: 1.75rem;
        border-width: 2px;
        color: var(--color-text-faint);
    }
    .overlay-loading span {
        font-size: .85rem;
        font-weight: 500;
        color: var(--color-text-muted);
    }

    /* ── Toast ────────────────────────────────────────────────────────── */
    .toast-container-custom {
        position: fixed; top: 1rem; right: 1rem;
        z-index: var(--z-toast);
        display: flex; flex-direction: column; gap: .5rem;
    }
    .toast-custom {
        min-width: 280px; max-width: 360px;
        border-radius: var(--radius-md); overflow: hidden;
        box-shadow: var(--shadow-md);
        background: var(--color-surface); border: 1px solid var(--color-border);
        animation: toastSlideIn .22s ease;
    }
    @keyframes toastSlideIn {
        from { opacity: 0; transform: translateX(24px); }
        to   { opacity: 1; transform: translateX(0); }
    }
    .toast-custom .toast-progress-bar {
        animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
        from { width: 100%; } to { width: 0%; }
    }

    /* ── Modal de Anexos ──────────────────────────────────────────────── */
    .anexo-item {
        padding: var(--space-3);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        margin-bottom: var(--space-2);
        transition: background-color var(--transition);
    }
    .anexo-item:hover { background-color: var(--color-surface-muted); }
    .anexo-icon {
        font-size: 1.5rem;
        color: var(--color-accent);
        margin-right: var(--space-3);
    }
    .anexo-nome {
        font-weight: 600;
        color: var(--color-text);
        margin-bottom: 4px;
        font-size: .88rem;
    }
    .anexo-data {
        font-size: .82rem;
        color: var(--color-text-muted);
    }

    /* ── Acessibilidade ───────────────────────────────────────────────── */
    @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
            animation-duration: 0.01ms !important;
            transition-duration: 0.01ms !important;
        }
    }
    </style>
</head>

<body>

<!-- ── Toast Container ────────────────────────────────────────────────── -->
<div class="toast-container-custom" id="toastContainer" aria-live="polite"></div>

<div class="container-fluid py-3 px-3 px-lg-4">

    <!-- ── Cabeçalho da Página ──────────────────────────────────────── -->
    <div class="d-flex align-items-center gap-3 mb-4">
        <span class="page-title">Solicitação de Alteração de Pedido</span>
    </div>

    <!-- ── Card de Filtros ──────────────────────────────────────────── -->
    <div class="filter-card mb-4">
        <div class="filter-card-header">
            <i class="bi bi-funnel me-2" style="color: var(--color-text-muted);"></i>Filtros
        </div>
        <div class="filter-card-body">
            <form id="form" name="form" method="post" action="">
                <input id="acao"   name="acao"   type="hidden" value="<%=acao%>" />
                <input id="filtro" name="filtro" type="hidden" value="<%=filtro%>" />

                <div class="d-flex flex-wrap align-items-end gap-3">

                    <!-- Filtro de liquidação -->
                    <div style="min-width:140px;">
                        <label for="cmbFiltro">Exibição</label>
                        <select name="cmbFiltro" id="cmbFiltro" class="form-select form-select-sm" onchange="filtrar();">
                            <option value="SL" <%if Request("cmbFiltro") = "SL" then%> selected <%end if%>>Sem liquidações</option>
                            <option value="CL" <%if Request("cmbFiltro") = "CL" then%> selected <%end if%>>Só liquidações</option>
                        </select>
                    </div>

                    <!-- Ações -->
                    <div class="ms-auto">
                        <label>Ações</label>
                        <div class="d-flex flex-wrap gap-2 justify-content-end">
                            <% If SESSION("PERMISSAO") <> "C" Then %>
                            <a href="SolicitacaoAlteracaoPedidoFormulario.asp" class="btn-aba btn-aba-registrar">
                                <i class="bi bi-plus-circle"></i>Registrar solicitação
                            </a>
                            <% End If %>
                            <button type="button" id="btnAtuList"
                                    class="btn-aba btn-aba-pendente <%=IIf(acao="BuscaSolicitacoes","aba-ativa","")%>"
                                    onclick="trocarAcao('BuscaSolicitacoes')">
                                <i class="bi bi-arrow-clockwise"></i>A realizar
                            </button>
                            <button type="button" id="btnSolReal"
                                    class="btn-aba btn-aba-realizada <%=IIf(acao="BuscaSolicitacoesRealizadas","aba-ativa","")%>"
                                    onclick="trocarAcao('BuscaSolicitacoesRealizadas')">
                                <i class="bi bi-check2-circle"></i>Realizadas
                            </button>
                            <button type="button" id="btnSolDesc"
                                    class="btn-aba btn-aba-descartada <%=IIf(acao="BuscaSolicitacoesDescartadas","aba-ativa","")%>"
                                    onclick="trocarAcao('BuscaSolicitacoesDescartadas')">
                                <i class="bi bi-x-circle"></i>Descartadas
                            </button>
                        </div>
                    </div>

                </div><!-- /d-flex -->
            </form>
        </div>
    </div>

    <!-- ── Card de Tabela ───────────────────────────────────────────── -->
    <div class="table-card">

        <!-- Overlay de carregamento -->
        <div id="overlayLoading" class="overlay-loading d-none">
            <div class="spinner-border" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
            <span>Carregando...</span>
        </div>

        <!-- Toolbar -->
        <div class="table-toolbar d-flex align-items-center justify-content-center px-3 py-2 border-bottom flex-wrap gap-2">
            <!-- Centro: subtítulo / badge de contexto -->
            <div class="d-flex align-items-center gap-2">
                <%
                Dim badgeClass, badgeLabel
                Select Case acao
                    Case "BuscaSolicitacoes"
                        badgeClass = "badge-status-pendente"
                        badgeLabel = "A Realizar"
                    Case "BuscaSolicitacoesRealizadas"
                        badgeClass = "badge-status-realizada"
                        badgeLabel = "Realizadas — últimos 15 dias"
                    Case "BuscaSolicitacoesDescartadas"
                        badgeClass = "badge-status-descartada"
                        badgeLabel = "Descartadas — últimos 15 dias"
                End Select

                Dim filtroLabel
                If filtro = "SL" Then
                    filtroLabel = " · sem liquidação"
                ElseIf filtro = "CL" Then
                    filtroLabel = " · só liquidação"
                Else
                    filtroLabel = ""
                End If
                %>
                <span class="badge-status <%=badgeClass%>"><%=badgeLabel%><%=filtroLabel%></span>
            </div>
        </div>

        <!-- Conteúdo da tabela -->
        <div id="conteudo" class="table-wrapper">
            <%
            Select Case acao
                Case "BuscaSolicitacoes"            : Response.Write BuscaSolicitacoes()
                Case "BuscaSolicitacoesRealizadas"  : Response.Write BuscaSolicitacoesRealizadas()
                Case "BuscaSolicitacoesDescartadas" : Response.Write BuscaSolicitacoesDescartadas()
            End Select
            %>
        </div>

    </div>
    <!-- /.table-card -->

</div>
<!-- /.container-fluid -->

<!-- ── Modal de Anexos ──────────────────────────────────────────────── -->
<div class="modal fade" id="modalAnexos" tabindex="-1" aria-labelledby="modalAnexosLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content" style="border-radius: var(--radius-lg); border: 1px solid var(--color-border);">
            <div class="modal-header" style="border-bottom: 1px solid var(--color-border);">
                <h5 class="modal-title" id="modalAnexosLabel" style="font-size: .95rem; font-weight: 600; color: var(--color-text);">
                    Anexos da Solicitação
                </h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
            </div>
            <div class="modal-body" id="modalAnexosBody" style="padding: var(--space-4);">
            </div>
            <div class="modal-footer" style="border-top: 1px solid var(--color-border); padding: .7rem 1.25rem;">
                <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Fechar</button>
            </div>
        </div>
    </div>
</div>

<!-- ── VBScript: funções de consulta ──────────────────────────────────── -->
<%
' ── BuscaSolicitacoes ─────────────────────────────────────────────────────
Function BuscaSolicitacoes()

    Dim tiposPermitidos, filialUsuario, SQL, DADOS, Table

    If Session("Nivel") = 0 Then ' NIVEL USUÁRIO

        tiposPermitidos = ""
        filialUsuario   = ""

        If SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "V" Then
            tiposPermitidos = "1,2,6"
        ElseIf SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "P" Then
            tiposPermitidos = "3,4,5,7,8,9,10,14"
        Else
            tiposPermitidos = "1,2,3,4,5,6,7,8,9,10,14"
        End If

        If SESSION("FILIAL") = "00" Then
            filialUsuario = "I"
        ElseIf SESSION("FILIAL") = "04" Then
            filialUsuario = "U"
        End If

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "Para, " &_
                "Para2, " &_
                "CASE C5_INPUBE WHEN 'U' THEN 'Uberaba' Else 'Pirapetinga' END FILIALDESCRICAO, " &_
                "CASE b.Id " &_
                "WHEN 3 THEN Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg'  " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN (   SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                                "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                    "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 SE4 ON ZZ7_COD = SE4.E4_CODIGO AND SE4.D_E_L_E_T_ = '' " &_
                                "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "SC6.C6_QTDENT, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado <> 1 " &_
                "AND Realizado = 0 " &_
                "AND Data >= '20150101' " &_
                "AND b.Id IN (" & tiposPermitidos & ") " &_
                IIf( filialUsuario <> "" And SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "P", "AND (C5_INPUBE = '" & filialUsuario & "' OR C5_INPUBE = '') ", "" ) &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND Para <> 'L' " ) &_
            " ORDER BY Realizado, Data, Hora "

    Else ' NIVEL REPRESENTANTE

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
                "Para2, " &_
                "CASE C5_INPUBE WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END FILIALDESCRICAO, " &_
                "CASE b.Id " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN (   SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                                "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                    "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 SE4 ON ZZ7_COD = SE4.E4_CODIGO AND SE4.D_E_L_E_T_ = '' " &_
                                "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "SC6.C6_QTDENT, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado <> 1 " &_
                "AND Realizado = 0 " &_
                "AND Data >= '20230101' " &_
                "AND lower(Login) = lower('" & Session("Usr") & "') " &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND Para <> 'L' " ) &_
            " ORDER BY Realizado, Data, Hora "

    End If

    Set DADOS = Conecta("PIR-HELPDESK", SQL)

    If Not IsObject(DADOS) Or (DADOS Is Nothing) Then
        Response.End
    End If

    If Not DADOS.EOF Then

        Table = "" &_
            "<table class='table table-bordered table-hover table-sm mb-0'>" &_
                "<thead>" &_
                    "<tr>" &_
                        "<th>Ação</th>" &_
                        "<th>Pedido</th>" &_
                        "<th>Planta</th>" &_
                        "<th>Cliente</th>" &_
                        "<th>#</th>" &_
                        "<th>Tipo</th>" &_
                        "<th>De</th>" &_
                        "<th>Para</th>" &_
                        "<th>Realizado</th>" &_
                        "<th>Obs</th>" &_
                        "<th>Solicitante</th>" &_
                        "<th>Data</th>" &_
                        "<th>Hora</th>" &_
                        "<th>Qtde. Entr.</th>" &_
                        "<th>Anexos</th>" &_
                    "</tr>" &_
                "</thead>" &_
                "<tbody>"

        While Not DADOS.EOF

            If (DateDiff("d", objUtil.AcertaData(DADOS("Data")), Date()) < 4 And DADOS("Realizado") = True) Or DADOS("Realizado") = False Then

                Table = Table & "<tr>"

                If Session("Nivel") = 0 And SESSION("PERMISSAO") = "C" Then

                    Dim lnk
                    lnk = "AlteraPedido.asp" &_
                            "?Pedido="     & Trim(DADOS("Pedido"))      &_
                            "&Tipo="       & Trim(DADOS("OperacaoId"))  &_
                            "&De="         & Trim(DADOS("De"))          &_
                            "&Para="       & Trim(DADOS("Para"))        &_
                            "&Para2="      & Trim(DADOS("Para2"))       &_
                            "&Descricao="  & Trim(DADOS("Descricao"))   &_
                            "&Id="         & Trim(DADOS("Id"))          &_
                            "&filtro="     & Request("filtro")

                    Table = Table &_
                        "<td>" &_
                            "<button type='button' class='btn btn-danger btn-sm'" &_
                                    " title='Excluir esta solicitação'" &_
                                    " onclick=""excluirSolicitacao('" & Trim(DADOS("Id")) & "')"">" &_
                                "<i class='bi bi-dash-circle'></i>" &_
                            "</button>" &_
                        "</td>" &_
                        "<td><a href='" & lnk & "'>" & Trim(DADOS("Pedido")) & "</a></td>"

                Else

                    Table = Table &_
                        "<td></td>" &_
                        "<td>" & Trim(DADOS("Pedido")) & "</td>"

                End If

                Table = Table &_
                    "<td>" & Trim(DADOS("FILIALDESCRICAO")) & "</td>" &_
                    "<td>" & Trim(DADOS("Cliente"))         & "</td>" &_
                    "<td>" & Trim(DADOS("OperacaoId"))      & "</td>" &_
                    "<td>" & Trim(DADOS("Operacao"))        & "</td>" &_
                    "<td>" & Trim(DADOS("De"))              & "</td>"

                If Trim(DADOS("OperacaoId")) = "5" And Trim(DADOS("Para2")) <> "" Then
                    Table = Table & "<td>" & FormatarDataPTBR(Trim(DADOS("Para"))) & " a " & FormatarDataPTBR(Trim(DADOS("Para2"))) & "</td>"
                Else
                    Table = Table & "<td>" & Trim(DADOS("Descricao")) & "</td>"
                End If

                If Trim(DADOS("Realizado")) Then
                    Table = Table & "<td><span class='badge-status badge-status-realizada'>SIM</span></td>"
                Else
                    Table = Table & "<td><span class='badge-status badge-status-pendente'>NÃO</span></td>"
                End If

                Table = Table &_
                    "<td>" & Trim(DADOS("Observacao")) & "</td>" &_
                    "<td>" & Trim(DADOS("Login"))      & "</td>" &_
                    "<td>" & objUtil.AcertaData(Trim(DADOS("Data"))) & " <small style='color:var(--color-text-muted)'>(" & DateDiff("d", objUtil.AcertaData(DADOS("Data")), Date()) & "d)</small></td>" &_
                    "<td>" & Trim(DADOS("Hora"))       & "</td>" &_
                    "<td>" & Trim(DADOS("C6_QTDENT"))  & "</td>" &_
                    "<td>" &_
                        "<button type='button' class='btn btn-primary btn-sm'" &_
                                " onclick=""abrirModalAnexos(" & Trim(DADOS("Id")) & ")""" &_
                                " title='Ver Anexos'>" &_
                            "<i class='bi bi-paperclip'></i>" &_
                        "</button>" &_
                    "</td>" &_
                    "</tr>"

            End If

            DADOS.MoveNext

        Wend

        Table = Table & "</tbody></table>"

    Else

        Table = "" &_
            "<div class='table-placeholder'>" &_
                "<i class='bi bi-clipboard-data'></i>" &_
                "<p>Nenhuma solicitação pendente encontrada.</p>" &_
            "</div>"

    End If

    DADOS.Close()
    Set DADOS = Nothing

    BuscaSolicitacoes = Table

End Function

' ── BuscaSolicitacoesRealizadas ────────────────────────────────────────────
Function BuscaSolicitacoesRealizadas()

    Dim SQL, DADOS, Table

    If Session("Nivel") = 0 Then ' NIVEL USUARIO

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "ISNULL(QuemRealizou,'') QuemRealizou, " &_
                "ISNULL(DtRealizou,'') DtRealizou, " &_
                "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
                "Para2, " &_
                "CASE b.Id " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                            "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                            "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "SC6.C6_QTDENT, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado <> 1 " &_
                "AND Data >= convert(varchar(8),dateadd(Day,-15,getdate()),112) " &_
                "AND Realizado = 1 " &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
            "ORDER BY Realizado, Data, Hora"

    Else ' NIVEL REPRESENTANTE

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "ISNULL(QuemRealizou,'') QuemRealizou, " &_
                "ISNULL(DtRealizou,'') DtRealizou, " &_
                "CASE WHEN b.Id = 3 then De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "CASE WHEN b.Id = 3 then Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
                "Para2, " &_
                "CASE b.Id " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                                "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                                "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "SC6.C6_QTDENT, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado <> 1 " &_
                "AND Data >= CONVERT(VARCHAR(8),DATEADD(DAY,-15,GETDATE()),112) " &_
                "AND Realizado = 1 " &_
                "AND LOWER(Login) = LOWER('" & Session("Usr") & "') " &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
            "ORDER BY Realizado, Data, Hora"

    End If

    Set DADOS = Conecta("PIR-HELPDESK", SQL)

    If Not IsObject(DADOS) Or (DADOS Is Nothing) Then
        Response.End
    End If

    If Not DADOS.EOF Then

        Table = "" &_
            "<table class='table table-bordered table-hover table-sm mb-0'>" &_
                "<thead>" &_
                    "<tr>" &_
                        "<th>Ação</th>" &_
                        "<th>Pedido</th>" &_
                        "<th>Cliente</th>" &_
                        "<th>#</th>" &_
                        "<th>Tipo</th>" &_
                        "<th>De</th>" &_
                        "<th>Para</th>" &_
                        "<th>Realizado</th>" &_
                        "<th>Obs</th>" &_
                        "<th>Solicitante</th>" &_
                        "<th>Data</th>" &_
                        "<th>Hora</th>" &_
                        "<th>Realizou</th>" &_
                        "<th>Dt. Realização</th>" &_
                        "<th>Qtde. Entr.</th>" &_
                        "<th>Anexos</th>" &_
                    "</tr>" &_
                "</thead>" &_
                "<tbody>"

        While Not DADOS.EOF

            Table = Table &_
                "<tr>" &_
                "<td></td>" &_
                "<td>" & Trim(DADOS("Pedido"))     & "</td>" &_
                "<td>" & Trim(DADOS("Cliente"))    & "</td>" &_
                "<td>" & Trim(DADOS("OperacaoId")) & "</td>" &_
                "<td>" & Trim(DADOS("Operacao"))   & "</td>" &_
                "<td>" & Trim(DADOS("De"))         & "</td>"

            If Trim(DADOS("OperacaoId")) = "5" And Trim(DADOS("Para2")) <> "" Then
                Table = Table & "<td>" & FormatarDataPTBR(Trim(DADOS("Para"))) & " a " & FormatarDataPTBR(Trim(DADOS("Para2"))) & "</td>"
            Else
                Table = Table & "<td>" & Trim(DADOS("Descricao")) & "</td>"
            End If

            If Trim(DADOS("Realizado")) = "1" Then
                Table = Table & "<td><span class='badge-status badge-status-realizada'>SIM</span></td>"
            Else
                Table = Table & "<td><span class='badge-status badge-status-pendente'>NÃO</span></td>"
            End If

            Table = Table &_
                "<td>" & Trim(DADOS("Observacao")) & "</td>" &_
                "<td>" & Trim(DADOS("Login"))      & "</td>"

            If Trim(DADOS("Data")) <> "" Then
                Table = Table & "<td>" & objUtil.AcertaData(Trim(DADOS("Data"))) & " <small style='color:var(--color-text-muted)'>(" & DateDiff("d", objUtil.AcertaData(DADOS("Data")), Date()) & "d)</small></td>"
            Else
                Table = Table & "<td></td>"
            End If

            Table = Table & "<td>" & Trim(DADOS("Hora")) & "</td>" &_
                            "<td>" & Trim(DADOS("QuemRealizou")) & "</td>"

            If Trim(DADOS("DtRealizou")) <> "" Then
                Table = Table & "<td>" & objUtil.AcertaData(Trim(DADOS("DtRealizou"))) & " <small style='color:var(--color-text-muted)'>(" & DateDiff("d", objUtil.AcertaData(DADOS("DtRealizou")), Date()) & "d)</small></td>"
            Else
                Table = Table & "<td></td>"
            End If

            Table = Table &_
                "<td>" & Trim(DADOS("C6_QTDENT")) & "</td>" &_
                "<td>" &_
                    "<button type='button' class='btn btn-primary btn-sm'" &_
                            " onclick=""abrirModalAnexos(" & Trim(DADOS("Id")) & ")""" &_
                            " title='Ver Anexos'>" &_
                        "<i class='bi bi-paperclip'></i>" &_
                    "</button>" &_
                "</td>" &_
                "</tr>"

            DADOS.MoveNext

        Wend

        Table = Table & "</tbody></table>"

    Else

        Table = "" &_
            "<div class='table-placeholder'>" &_
                "<i class='bi bi-clipboard-data'></i>" &_
                "<p>Nenhuma solicitação realizada nos últimos 15 dias.</p>" &_
            "</div>"

    End If

    DADOS.Close()
    Set DADOS = Nothing

    BuscaSolicitacoesRealizadas = Table

End Function

' ── BuscaSolicitacoesDescartadas ──────────────────────────────────────────
Function BuscaSolicitacoesDescartadas()

    Dim SQL, DADOS, Table

    If Session("Nivel") = 0 Then ' NIVEL USUARIO

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
                "Para2, " &_
                "CASE b.Id " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                            "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                            "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "Descartado, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "SC6.C6_QTDENT, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado = 1 " &_
                "AND Data >= CONVERT(VARCHAR(8),DATEADD(DAY,-15,GETDATE()),112) " &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
            "ORDER BY Realizado, Data, Hora"

    Else ' NIVEL REPRESENTANTE

        SQL = "" &_
            "SELECT DISTINCT " &_
                "a.Id AS Id, " &_
                "Pedido AS Pedido, " &_
                "A1_NREDUZ AS Cliente, " &_
                "b.Id AS OperacaoId, " &_
                "Operacao AS Operacao, " &_
                "CASE WHEN b.Id = 3 then De + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
                "CASE WHEN b.Id = 3 then Para + ' - ' + REPLACE(cast(cast((" & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') AS float)) AS int) AS varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
                "Para2, " &_
                "CASE b.Id " &_
                "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
                "WHEN 6 THEN ( SELECT TOP 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                                "FROM " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.ZZ7010 " &_
                                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                                "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
                "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
                "ELSE Para END AS Descricao, " &_
                "Observacao AS Observacao, " &_
                "Realizado AS Realizado, " &_
                "Data AS Data, " &_
                "Hora AS Hora, " &_
                "SC6.C6_QTDENT, " &_
                "Login AS Login " &_
            "FROM InHistoricoAlteracoesPedidosSolicitacao AS a " &_
                "INNER JOIN InHistoricoTipoOperacao AS b on a.Tipo = b.Id " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC5010 SC5 ON SC5.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SC6010 SC6 ON SC6.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
                "INNER JOIN " & cnx(seekCnx("PIR-DADOSADV")).getDataBase & ".dbo.SA1010 SA1 ON SA1.A1_COD = SC5.C5_CLIENTE AND SA1.A1_LOJA = SC5.C5_LOJACLI " &_
            "WHERE Descartado = 1 " &_
                "AND Data >= CONVERT(VARCHAR(8),DATEADD(Day,-15,GETDATE()),112) " &_
                "AND LOWER(Login) = LOWER('" & Session("Usr") & "') " &_
                IIf( Request("filtro") = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
            "ORDER BY Realizado, Data, Hora"

    End If

    Set DADOS = Conecta("PIR-HELPDESK", SQL)

    If Not IsObject(DADOS) Or (DADOS Is Nothing) Then
        Response.End
    End If

    If Not DADOS.EOF Then

        Table = "" &_
            "<table class='table table-bordered table-hover table-sm mb-0'>" &_
                "<thead>" &_
                    "<tr>" &_
                        "<th>Ação</th>" &_
                        "<th>Pedido</th>" &_
                        "<th>Cliente</th>" &_
                        "<th>#</th>" &_
                        "<th>Tipo</th>" &_
                        "<th>De</th>" &_
                        "<th>Para</th>" &_
                        "<th>Realizado</th>" &_
                        "<th>Obs</th>" &_
                        "<th>Solicitante</th>" &_
                        "<th>Data</th>" &_
                        "<th>Hora</th>" &_
                        "<th>Qtde. Entr.</th>" &_
                        "<th>Anexos</th>" &_
                    "</tr>" &_
                "</thead>" &_
                "<tbody>"

        While Not DADOS.EOF

            Table = Table &_
                "<tr class='row-danger'>" &_
                "<td></td>" &_
                "<td>" & Trim(DADOS("Pedido"))     & "</td>" &_
                "<td>" & Trim(DADOS("Cliente"))    & "</td>" &_
                "<td>" & Trim(DADOS("OperacaoId")) & "</td>" &_
                "<td>" & Trim(DADOS("Operacao"))   & "</td>" &_
                "<td>" & Trim(DADOS("De"))         & "</td>"

            If Trim(DADOS("OperacaoId")) = "5" And Trim(DADOS("Para2")) <> "" Then
                Table = Table & "<td>" & FormatarDataPTBR(Trim(DADOS("Para"))) & " a " & FormatarDataPTBR(Trim(DADOS("Para2"))) & "</td>"
            Else
                Table = Table & "<td>" & Trim(DADOS("Descricao")) & "</td>"
            End If

            If Trim(DADOS("Realizado")) = "1" Then
                Table = Table & "<td><span class='badge-status badge-status-realizada'>SIM</span></td>"
            Else
                Table = Table & "<td><span class='badge-status badge-status-descartada'>DESCARTADA</span></td>"
            End If

            Table = Table &_
                "<td>" & Trim(DADOS("Observacao")) & "</td>" &_
                "<td>" & Trim(DADOS("Login"))      & "</td>" &_
                "<td>" & objUtil.AcertaData(Trim(DADOS("Data"))) & " <small style='color:var(--color-text-muted)'>(" & DateDiff("d", objUtil.AcertaData(DADOS("Data")), Date()) & "d)</small></td>" &_
                "<td>" & Trim(DADOS("Hora"))       & "</td>" &_
                "<td>" & Trim(DADOS("C6_QTDENT"))  & "</td>" &_
                "<td>" &_
                    "<button type='button' class='btn btn-primary btn-sm'" &_
                            " onclick=""abrirModalAnexos(" & Trim(DADOS("Id")) & ")""" &_
                            " title='Ver Anexos'>" &_
                        "<i class='bi bi-paperclip'></i>" &_
                    "</button>" &_
                "</td>" &_
                "</tr>"

            DADOS.MoveNext

        Wend

        Table = Table & "</tbody></table>"

    Else

        Table = "" &_
            "<div class='table-placeholder'>" &_
                "<i class='bi bi-clipboard-data'></i>" &_
                "<p>Nenhuma solicitação descartada nos últimos 15 dias.</p>" &_
            "</div>"

    End If

    DADOS.Close()
    Set DADOS = Nothing

    BuscaSolicitacoesDescartadas = Table

End Function

' ── FormatarDataPTBR ──────────────────────────────────────────────────────
Function FormatarDataPTBR(aData)

    Dim ano, mes, dia

    If InStr(1, aData, "/") > 0 Then

        FormatarDataPTBR = aData

    Else

        If InStr(1, aData, "-") > 0 Then
            ano = Left(aData, 4)
            mes = Mid(aData, 6, 2)
            dia = Right(aData, 2)
            FormatarDataPTBR = dia & "/" & mes & "/" & ano
        Else
            ano = Left(aData, 4)
            mes = Mid(aData, 5, 2)
            dia = Right(aData, 2)
            FormatarDataPTBR = dia & "/" & mes & "/" & ano
        End If

    End If

End Function
%>

<!-- ── Bootstrap JS ──────────────────────────────────────────────────────── -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>

<!-- jQuery 4.0.0 -->
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
/* ── Namespace principal ────────────────────────────────────────────────── */
const SolicitaAlteracao = (() => {

    /* ── Toast ────────────────────────────────────────────────────────── */
    const TOAST_DELAY = 4500;
    const ICONE_TOAST = {
        sucesso : 'bi-check-circle-fill',
        erro    : 'bi-x-circle-fill',
        aviso   : 'bi-exclamation-triangle-fill',
        info    : 'bi-info-circle-fill',
    };
    const COR_TOAST = {
        sucesso : 'var(--color-success)',
        erro    : 'var(--color-danger)',
        aviso   : 'var(--color-warning)',
        info    : 'var(--color-text-muted)',
    };

    function exibirToast(tipo, mensagem) {
        const icone = ICONE_TOAST[tipo] ?? ICONE_TOAST.info;
        const cor   = COR_TOAST[tipo]   ?? COR_TOAST.info;
        const el    = document.createElement('div');
        el.className = 'toast-custom p-3';
        el.innerHTML = `
            <div class="d-flex align-items-start gap-2">
                <i class="bi ${icone}" style="color:${cor};font-size:1.1rem;flex-shrink:0;margin-top:1px;"></i>
                <div style="flex:1;font-size:.85rem;color:var(--color-text);">${mensagem}</div>
                <button style="background:none;border:none;padding:0;cursor:pointer;color:var(--color-text-muted);"
                        onclick="this.closest('.toast-custom').remove()">&times;</button>
            </div>
            <div class="toast-progress-bar mt-2" style="height:3px;background:${cor};border-radius:2px;"></div>`;
        document.getElementById('toastContainer').appendChild(el);
        setTimeout(() => el.remove(), TOAST_DELAY + 300);
    }

    /* ── Overlay ─────────────────────────────────────────────────────── */
    function mostrarOverlay() { $('#overlayLoading').removeClass('d-none'); }
    function ocultarOverlay() { $('#overlayLoading').addClass('d-none'); }

    /* ── Trocar ação (submete o form) ────────────────────────────────── */
    function trocarAcao(novaAcao) {
        $('#conteudo').empty();
        $('#acao').val(novaAcao);
        mostrarOverlay();
        document.getElementById('form').submit();
    }

    /* ── Filtrar (change no select de liquidação) ─────────────────────── */
    function filtrar() {
        $('#filtro').val($('#cmbFiltro').val());
        mostrarOverlay();
        document.getElementById('form').submit();
    }

    /* ── Excluir solicitação ─────────────────────────────────────────── */
    function excluirSolicitacao(id) {
        if (!confirm('Deseja realmente EXCLUIR esta solicitação?')) return;
        mostrarOverlay();
        $.ajax({
            url      : 'ObjetosAsp/objAjax.asp',
            method   : 'GET',
            dataType : 'text',
            data     : { acao: 'Exclui', tipo: id },
            success  : function() {
                exibirToast('sucesso', 'Solicitação excluída com sucesso.');
                trocarAcao('BuscaSolicitacoes');
            },
            error    : function(xhr) {
                ocultarOverlay();
                exibirToast('erro', 'Erro ao excluir solicitação.');
                console.error(xhr.responseText);
            },
        });
    }

    /* ── Modal de Anexos ─────────────────────────────────────────────── */
    function abrirModalAnexos(idSolicitacao) {
        const modal = new bootstrap.Modal(document.getElementById('modalAnexos'));
        $('#modalAnexosBody').html(
            '<div class="d-flex align-items-center gap-2 py-3 text-muted">' +
                '<div class="spinner-border spinner-border-sm" role="status"></div>' +
                '<span style="font-size:.85rem;">Carregando anexos...</span>' +
            '</div>'
        );
        modal.show();
        carregarAnexos(idSolicitacao);
    }

    function carregarAnexos(idSolicitacao) {
        $.ajax({
            url      : 'SolicitacaoAlteracaoPedidoAnexos.asp?v=' + Math.random(),
            method   : 'GET',
            dataType : 'html',
            data     : { idSolicitacao },
            success  : function(response) {
                $('#modalAnexosBody').html(response);
            },
            error    : function(xhr) {
                $('#modalAnexosBody').html(
                    '<div class="d-flex align-items-center gap-2" style="color:var(--color-danger);font-size:.85rem;">' +
                        '<i class="bi bi-exclamation-octagon"></i>' +
                        '<span>Erro ao carregar anexos.</span>' +
                    '</div>'
                );
                console.error(xhr.responseText);
            },
        });
    }

    /* ── Init ────────────────────────────────────────────────────────── */
    function init() {
        /* expõe funções usadas inline pelo ASP gerado */
        window.trocarAcao       = trocarAcao;
        window.filtrar          = filtrar;
        window.excluirSolicitacao = excluirSolicitacao;
        window.abrirModalAnexos = abrirModalAnexos;
    }

    return Object.freeze({ init });
})();

$(function() { SolicitaAlteracao.init(); });
</script>

</body>
</html>
