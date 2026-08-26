<%SESSION.LCID = 1046 %>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
    ' Formata data de YYYY-MM-DD para DD/MM/YYYY
    Function FormatarDataBR(d)
        Dim s
        s = Trim(CStr(d))
        If Len(s) >= 10 And Mid(s,5,1) = "-" Then
            FormatarDataBR = Mid(s,9,2) & "/" & Mid(s,6,2) & "/" & Left(s,4)
        ElseIf Len(s) = 8 Then
            FormatarDataBR = Mid(s,7,2) & "/" & Mid(s,5,2) & "/" & Left(s,4)
        Else
            FormatarDataBR = s
        End If
    End Function

    ' [Data] na tabela FechamentoRefeitorio é varchar(8) no formato YYYYMMDD.
    ' Converte qualquer entrada para YYYYMMDD para comparação correta como string.
    Function FormatarDataSQL(d)
        Dim s
        s = Trim(CStr(d))
        If Len(s) = 10 And Mid(s,3,1) = "/" And Mid(s,6,1) = "/" Then
            ' Entrada DD/MM/YYYY → YYYYMMDD
            FormatarDataSQL = Right(s,4) & Mid(s,4,2) & Left(s,2)
        ElseIf Len(s) = 10 And Mid(s,5,1) = "-" Then
            ' Entrada YYYY-MM-DD → YYYYMMDD (remove hífens)
            FormatarDataSQL = Left(s,4) & Mid(s,6,2) & Right(s,2)
        ElseIf Len(s) = 8 And IsNumeric(s) Then
            ' Já está em YYYYMMDD
            FormatarDataSQL = s
        Else
            FormatarDataSQL = s
        End If
    End Function

    Dim ConexaoFechamento
    Dim StringConexaoFechamento
    Dim Dados

    StringConexaoFechamento = BuscaStringConexao("PIR-RBACESSO")

    If Trim(StringConexaoFechamento) = "" Then
        Response.Write "A string de conexão PIR-RBACESSO não foi encontrada no arquivo conexao.txt."
        Response.End
    End If

    On Error Resume Next

    Set ConexaoFechamento = Server.CreateObject("ADODB.Connection")
    ConexaoFechamento.CommandTimeout = 300
    ConexaoFechamento.Open StringConexaoFechamento

    If Err.Number <> 0 Then
        Response.Write "Não foi possível conectar ao banco de dados PIR-RBACESSO."
        Response.End
    End If

    On Error GoTo 0

    Set Dados = Server.CreateObject("ADODB.Recordset")
    Dados.ActiveConnection = ConexaoFechamento

    if True then  ' sempre executa o fechamento
        ' Usa YYYY-MM-DD para garantir conversão correta independente do SET DATEFORMAT do servidor
        SQL = "EXEC dbo.sp_ExecutaFechamentoRefeitorio " & _
              "'" & FormatarDataSQL(Request("dtini")) & "'," & _
              "'" & FormatarDataSQL(Request("dtfim")) & "'"

        On Error Resume Next
        ConexaoFechamento.Execute SQL

        If Err.Number <> 0 Then
            Dim NumeroErroFechamento
            Dim DescricaoErroFechamento

            NumeroErroFechamento = Err.Number
            DescricaoErroFechamento = Err.Description
            Err.Clear

            On Error GoTo 0

            Response.Write "Erro ao executar o fechamento do refeitório.<br>"
            Response.Write "Número: " & Server.HTMLEncode(CStr(NumeroErroFechamento)) & "<br>"
            Response.Write "Descrição: " & Server.HTMLEncode(DescricaoErroFechamento)
            Response.End
        End If

        On Error GoTo 0
    end if

    Dim strDtIni, strDtFim
    strDtIni = FormatarDataSQL(Request("dtini"))
    strDtFim = FormatarDataSQL(Request("dtfim"))

    if trim(request("tipo")) = "1" then
        SQL = "  SELECT  "
        SQL = SQL &"Credencial"
        SQL = SQL &",[Data]  "
        SQL = SQL &",Matricula "
        SQL = SQL &",Nome"
        SQL = SQL &",Ano "
        SQL = SQL &",Mes "
        SQL = SQL &",Dia "
        SQL = SQL &",Hora "
        SQL = SQL &",Minuto "
        SQL = SQL &",Isnull(Valor,0) Valor "
        SQL = SQL &"FROM FechamentoRefeitorio "
        SQL = SQL &"Where [Data] >= '"&strDtIni&"' AND [Data] <= '"&strDtFim&"' "
        SQL = SQL &"And Planta = '"&Trim(Request("Planta"))&"' "
        SQL = SQL &"ORDER BY [Data],Nome "
    Else
        SQL = "  SELECT  "
        SQL = SQL &"Credencial"
        SQL = SQL &",Matricula "
        SQL = SQL &",Nome"
        SQL = SQL &",Sum(Valor) Valor "
        SQL = SQL &"FROM FechamentoRefeitorio "
        SQL = SQL &"Where [Data] >= '"&strDtIni&"' AND [Data] <= '"&strDtFim&"' "
        SQL = SQL &"And Planta = '"&Trim(Request("Planta"))&"' "
        SQL = SQL &"Group By Matricula,Credencial,Nome "
        SQL = SQL &"ORDER BY Nome "
    End if


    Dados.Source = SQL
    Dados.open()
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Fechamento Refeitório</title>

    <!-- Bootstrap 5.3 -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" />
    <!-- Bootstrap Icons -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />
    <!-- jsPDF + AutoTable (geração nativa de PDF) -->
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js"></script>

    <style>
        /* ─── Tokens de Design ─────────────────────────────────────── */
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
            --color-danger:        hsl(3, 90%, 58%);
            --color-danger-soft:   hsl(3, 100%, 96%);
            --color-warning:       hsl(35, 100%, 48%);
            --color-warning-soft:  hsl(40, 100%, 95%);

            --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

            --radius-sm:  6px;
            --radius-md:  10px;
            --radius-lg:  14px;

            --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
            --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

            --transition: 160ms ease;

            --space-1: 4px;
            --space-2: 8px;
            --space-3: 12px;
            --space-4: 16px;
            --space-5: 24px;
            --space-6: 32px;
            --space-7: 48px;
        }

        /* ─── Base ────────────────────────────────────────────────── */
        *, *::before, *::after { box-sizing: border-box; }

        body {
            background-color: var(--color-bg);
            color: var(--color-text);
            font-family: var(--font-sans);
            font-size: .875rem;
            line-height: 1.5;
            -webkit-font-smoothing: antialiased;
        }

        /* ─── Layout de página ────────────────────────────────────── */
        .page-wrapper {
            max-width: 1200px;
            margin: 0 auto;
            padding: var(--space-6) var(--space-5);
        }

        /* ─── Cabeçalho da página ─────────────────────────────────── */
        .page-header {
            display: flex;
            align-items: stretch;
            justify-content: space-between;
            gap: 0;
            margin-bottom: var(--space-6);
            flex-wrap: wrap;
        }

        .page-header__title-group {
            flex: 1;
            min-width: 0;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .page-title {
            font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
            font-weight: 600;
            color: var(--color-text);
            margin: 0 0 var(--space-1) 0;
            letter-spacing: -.01em;
        }

        .page-subtitle {
            font-size: .8125rem;
            color: var(--color-text-muted);
            margin: 0;
        }

        /* ─── Aba "Exportar PDF" no cabeçalho ────────────────────── */
        .header-tab-pdf {
            display: inline-flex;
            align-items: center;
            gap: var(--space-2);
            padding: 0 var(--space-5);
            background: var(--color-danger);
            border: 1px solid var(--color-danger);
            border-bottom: none;
            border-radius: var(--radius-sm) var(--radius-sm) 0 0;
            color: #fff;
            font-size: .8125rem;
            font-weight: 500;
            cursor: pointer;
            text-decoration: none;
            align-self: flex-end;
            height: 38px;
            transition: background-color var(--transition), border-color var(--transition);
            white-space: nowrap;
        }

        .header-tab-pdf:hover {
            background: hsl(3, 90%, 50%);
            border-color: hsl(3, 90%, 50%);
            color: #fff;
        }

        .header-tab-pdf:focus-visible {
            outline: none;
            box-shadow: 0 0 0 3px hsla(3, 90%, 58%, .3);
        }

        /* ─── Aba "Exportar Excel" no cabeçalho ───────────────────── */
        .header-tab-excel {
            display: inline-flex;
            align-items: center;
            gap: var(--space-2);
            padding: 0 var(--space-5);
            background: var(--color-success);
            border: 1px solid var(--color-success);
            border-bottom: none;
            border-radius: var(--radius-sm) var(--radius-sm) 0 0;
            color: #fff;
            font-size: .8125rem;
            font-weight: 500;
            cursor: pointer;
            text-decoration: none;
            align-self: flex-end;
            height: 38px;
            transition: background-color var(--transition), border-color var(--transition);
            white-space: nowrap;
        }

        .header-tab-excel:hover {
            background: hsl(145, 63%, 27%);
            border-color: hsl(145, 63%, 27%);
            color: #fff;
        }

        .header-tab-excel:focus-visible {
            outline: none;
            box-shadow: 0 0 0 3px hsla(145, 63%, 32%, .3);
        }

        /* ─── Card principal ──────────────────────────────────────── */
        .data-card {
            background: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-sm);
            overflow: hidden;
        }

        /* ─── Toolbar interna (filtro + rows per page) ────────────── */
        .table-toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: var(--space-3);
            padding: var(--space-3) var(--space-5);
            border-bottom: 1px solid var(--color-border);
            background: var(--color-surface-muted);
            flex-wrap: wrap;
        }

        .filter-wrap {
            position: relative;
            flex: 1;
            min-width: 180px;
            max-width: 340px;
        }

        .filter-icon {
            position: absolute;
            left: 10px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--color-text-faint);
            font-size: .875rem;
            pointer-events: none;
        }

        .filter-input {
            width: 100%;
            height: 34px;
            padding: 0 10px 0 32px;
            border: 1px solid var(--color-border);
            border-radius: var(--radius-sm);
            background: var(--color-surface);
            color: var(--color-text);
            font-size: .8125rem;
            outline: none;
            transition: border-color var(--transition), box-shadow var(--transition);
        }

        .filter-input:focus {
            border-color: var(--color-accent);
            box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .15);
        }

        .filter-input::placeholder { color: var(--color-text-faint); }

        .rows-wrap {
            display: flex;
            align-items: center;
            gap: var(--space-2);
            white-space: nowrap;
            font-size: .8125rem;
            color: var(--color-text-muted);
        }

        .rows-select {
            height: 34px;
            padding: 0 28px 0 10px;
            border: 1px solid var(--color-border);
            border-radius: var(--radius-sm);
            background: var(--color-surface);
            color: var(--color-text);
            font-size: .8125rem;
            outline: none;
            appearance: none;
            background-image:
                linear-gradient(45deg, transparent 50%, var(--color-text-muted) 50%),
                linear-gradient(135deg, var(--color-text-muted) 50%, transparent 50%);
            background-position:
                calc(100% - 12px) 50%,
                calc(100% - 7px) 50%;
            background-repeat: no-repeat;
            background-size: 4px 4px, 4px 4px;
            cursor: pointer;
            transition: border-color var(--transition);
        }

        .rows-select:focus { border-color: var(--color-accent); }

        /* ─── Tabela de dados ─────────────────────────────────────── */

        /* Wrapper com scroll vertical: cabeçalho fixo, corpo rolável > 12 linhas */
        .table-scroll {
            overflow-x: auto;
        }

        /* Container interno que aplica o scroll vertical no tbody */
        .table-scroll-inner {
            /* Altura de 20 linhas aprox. (linha ~41px + border) */
            max-height: calc(20 * 41px);
            overflow-y: auto;
            overflow-x: visible;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: .8125rem;
        }

        /* thead sticky para que o cabeçalho permaneça visível durante o scroll */
        .data-table thead {
            position: sticky;
            top: 0;
            z-index: 2;
        }

        .data-table thead tr {
            background-color: var(--color-surface-muted);
            border-bottom: 2px solid var(--color-border-strong);
        }

        .data-table thead th {
            padding: var(--space-3) var(--space-4);
            font-size: .6875rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: .04em;
            color: var(--color-text-muted);
            text-align: center;
            white-space: nowrap;
            /* Linha vertical de separação nas colunas do cabeçalho */
            border-right: 1px solid var(--color-border-strong);
        }
        .data-table thead th:last-child { border-right: none; }

        .data-table tbody tr {
            border-bottom: 1px solid var(--color-border);
            transition: background-color var(--transition);
        }

        .data-table tbody tr.row-hidden { display: none; }

        .data-table tbody tr:last-child { border-bottom: none; }

        .data-table tbody tr:hover {
            background-color: var(--color-surface-muted);
        }

        .data-table tbody td {
            padding: var(--space-3) var(--space-4);
            color: var(--color-text);
            text-align: center;
            /* Linha vertical de separação nas células do corpo */
            border-right: 1px solid var(--color-border);
        }
        .data-table tbody td:last-child { border-right: none; }
            vertical-align: middle;
        }

        .data-table td.col-nome {
            text-align: left;
            font-weight: 500;
            white-space: nowrap;
        }

        .badge-credencial {
            display: inline-block;
            padding: 2px var(--space-2);
            border-radius: var(--radius-sm);
            background: var(--color-surface-muted);
            border: 1px solid var(--color-border);
            color: var(--color-text-muted);
            font-size: .6875rem;
            font-weight: 500;
            font-family: "SF Mono", "Cascadia Code", Consolas, monospace;
            letter-spacing: .02em;
        }

        .col-valor {
            font-variant-numeric: tabular-nums;
            font-weight: 500;
        }

        /* ─── Sem resultados ──────────────────────────────────────── */
        .no-results-row td {
            padding: var(--space-6) var(--space-5);
            text-align: center;
            color: var(--color-text-faint);
            font-size: .8125rem;
        }

        /* ─── Paginação ───────────────────────────────────────────── */
        .table-footer {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: var(--space-4);
            padding: var(--space-3) var(--space-5);
            border-top: 1px solid var(--color-border);
            background: var(--color-surface-muted);
            flex-wrap: wrap;
        }

        .table-footer__summary {
            font-size: .8125rem;
            color: var(--color-text-muted);
        }

        .table-footer__summary strong {
            color: var(--color-text);
            font-weight: 600;
        }

        .pagination-controls {
            display: flex;
            align-items: center;
            gap: var(--space-1);
        }

        .page-btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 32px;
            height: 32px;
            padding: 0 var(--space-2);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-sm);
            background: var(--color-surface);
            color: var(--color-text);
            font-size: .75rem;
            font-weight: 500;
            cursor: pointer;
            transition: background-color var(--transition), border-color var(--transition);
            line-height: 1;
        }

        .page-btn:hover:not(:disabled) {
            background: var(--color-surface-muted);
            border-color: var(--color-border-strong);
        }

        .page-btn.active {
            background: var(--color-accent);
            border-color: var(--color-accent);
            color: #fff;
        }

        .page-btn:disabled {
            opacity: .4;
            cursor: not-allowed;
        }

        .page-info {
            font-size: .75rem;
            color: var(--color-text-muted);
            padding: 0 var(--space-2);
            white-space: nowrap;
        }

        /* ─── Botões de ação (fora do card) ──────────────────────── */
        .page-actions {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: var(--space-3);
            margin-top: var(--space-5);
        }

        .btn-quiet {
            display: inline-flex;
            align-items: center;
            gap: var(--space-2);
            padding: var(--space-2) var(--space-4);
            font-size: .8125rem;
            font-weight: 500;
            line-height: 1.5;
            border-radius: var(--radius-sm);
            border: 1px solid var(--color-border);
            background: var(--color-surface);
            color: var(--color-text);
            cursor: pointer;
            text-decoration: none;
            transition: background-color var(--transition), border-color var(--transition), color var(--transition), box-shadow var(--transition);
            white-space: nowrap;
        }

        .btn-quiet:hover {
            background-color: var(--color-surface-muted);
            border-color: var(--color-border-strong);
            color: var(--color-text);
        }

        .btn-quiet:focus-visible {
            outline: none;
            box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .25);
        }

        .btn-quiet--primary {
            background: var(--color-accent);
            border-color: var(--color-accent);
            color: #fff;
        }

        .btn-quiet--primary:hover {
            background: var(--color-accent-hover);
            border-color: var(--color-accent-hover);
            color: #fff;
        }

        /* ─── Responsividade ──────────────────────────────────────── */
        @media (max-width: 768px) {
            .page-wrapper { padding: var(--space-4); }

            .page-header {
                flex-direction: column;
                align-items: flex-start;
                gap: var(--space-3);
            }

            .header-tab-pdf,
            .header-tab-excel {
                border-radius: var(--radius-sm);
                border-bottom: 1px solid currentColor;
                align-self: auto;
                height: 38px;
            }

            .table-footer {
                flex-direction: column;
                align-items: stretch;
                text-align: center;
            }

            .pagination-controls { justify-content: center; }

            .page-actions {
                flex-wrap: wrap;
            }

            .btn-quiet { flex: 1 1 auto; justify-content: center; }
        }

        /* ─── Acessibilidade ──────────────────────────────────────── */
        @media (prefers-reduced-motion: reduce) {
            *, *::before, *::after { transition: none !important; }
        }
    </style>
</head>
<body>
    <div class="page-wrapper">

        <!-- Cabeçalho com aba Excel -->
        <header class="page-header">
            <div class="page-header__title-group">
                <h1 class="page-title">Fechamento Refeitório</h1>
                <p class="page-subtitle">
                    <%if trim(request("tipo")) = "1" then%>
                    Relatório detalhado por acesso
                    <%else%>
                    Relatório consolidado por colaborador
                    <%end if%>
                    &mdash; período de <strong><%=FormatarDataBR(trim(request("dtini")))%></strong> a <strong><%=FormatarDataBR(trim(request("dtfim")))%></strong>
                </p>
            </div>
            <!-- Botões de exportação como abas, alinhados à direita no header -->
            <div style="display:inline-flex;align-items:flex-end;gap:4px;">
                <button type="button" class="header-tab-pdf" onclick="exportarPDF()" title="Exportar tabela para PDF">
                    <i class="bi bi-file-earmark-pdf me-1" aria-hidden="true"></i>Exportar PDF
                </button>
                <button type="button" class="header-tab-excel" onclick="exportarCSV()" title="Exportar tabela para Excel">
                    <i class="bi bi-file-earmark-spreadsheet me-1" aria-hidden="true"></i>Exportar Excel
                </button>
            </div>
        </header>

        <!-- Card com tabela -->
        <div class="data-card">

            <!-- Toolbar: filtro + linhas por página -->
            <div class="table-toolbar">
                <div class="filter-wrap">
                    <i class="bi bi-search filter-icon" aria-hidden="true"></i>
                    <input
                        type="text"
                        id="filtroTabela"
                        class="filter-input"
                        placeholder="Filtrar por qualquer coluna..."
                        oninput="filtrarTabela()"
                        autocomplete="off"
                    />
                </div>
                <div class="rows-wrap">
                    <span>Linhas por página:</span>
                    <select id="rowsPerPage" class="rows-select" onchange="mudarRowsPerPage()">
                        <option value="100">100</option>
                        <option value="500">500</option>
                        <option value="1000">1000</option>
                        <option value="5000">5000</option>
                        <option value="10000">10000</option>
                    </select>
                </div>
            </div>

            <div class="table-scroll">
            <div class="table-scroll-inner" id="tabelaScrollInner">

                <%if trim(request("tipo")) = "1" then%>
                <!-- Tabela detalhada -->
                <table class="data-table" id="tabelaDados" aria-label="Fechamento refeitório — detalhado">
                    <thead>
                        <tr>
                            <th scope="col">Credencial</th>
                            <th scope="col">Matrícula</th>
                            <th scope="col" class="col-nome">Nome</th>
                            <th scope="col">Ano</th>
                            <th scope="col">Mês</th>
                            <th scope="col">Dia</th>
                            <th scope="col">Hora</th>
                            <th scope="col">Minuto</th>
                            <th scope="col">Valor</th>
                        </tr>
                    </thead>
                    <tbody id="tabelaBody">
<%
                    cont = 0
                    while not Dados.Eof
%>
                        <tr>
                            <td><span class="badge-credencial"><%= Dados("Credencial")%></span></td>
                            <td><%= Dados("Matricula")%></td>
                            <td class="col-nome"><%= Dados("Nome")%></td>
                            <td><%= Dados("Ano")%></td>
                            <td><%= Dados("Mes")%></td>
                            <td><%= Dados("Dia")%></td>
                            <td><%= Dados("Hora")%></td>
                            <td><%= Dados("Minuto")%></td>
                            <td class="col-valor"><%= Dados("Valor")%></td>
                        </tr>
<%
                        cont = cont + 1
                        response.Flush()
                        Dados.Movenext()
                    Wend
%>
                    </tbody>
                </table>
                <%else%>
                <!-- Tabela consolidada -->
                <table class="data-table" id="tabelaDados" aria-label="Fechamento refeitório — consolidado">
                    <thead>
                        <tr>
                            <th scope="col">Credencial</th>
                            <th scope="col">Matrícula</th>
                            <th scope="col" class="col-nome">Nome</th>
                            <th scope="col">Valor Total</th>
                        </tr>
                    </thead>
                    <tbody id="tabelaBody">
<%
                    cont = 0
                    while not Dados.Eof
%>
                        <tr>
                            <td><span class="badge-credencial"><%= Dados("Credencial")%></span></td>
                            <td><%= Dados("Matricula")%></td>
                            <td class="col-nome"><%= Dados("Nome")%></td>
                            <td class="col-valor"><%= FormatNumber(Dados("Valor"),2)%></td>
                        </tr>
<%
                        cont = cont + 1
                        response.Flush()
                        Dados.Movenext()
                    Wend
%>
                    </tbody>
                </table>
                <%end if%>

            </div><!-- /.table-scroll-inner -->
            </div><!-- /.table-scroll -->

            <!-- Rodapé com paginação — paginação à esquerda, resumo à direita -->
            <div class="table-footer">
                <div class="pagination-controls" id="paginacaoContainer" aria-label="Navegação de páginas">
                    <!-- Preenchido via JS -->
                </div>
                <p class="table-footer__summary mb-0" id="resumoRegistros">
                    <strong id="contadorVisiveis"><%=cont%></strong>
                    de <strong><%=cont%></strong>
                    <%if trim(request("tipo")) = "1" then%>
                    registros
                    <%else%>
                    colaboradores
                    <%end if%>
                </p>
            </div>

        </div><!-- /.data-card -->

        <!-- ── Botões fora do card ── -->
        <div class="page-actions">
            <button type="button" class="btn-quiet btn-quiet--primary" onclick="location.href='ParametrosFechamento.asp'">
                <i class="bi bi-arrow-left me-1" aria-hidden="true"></i>Voltar
            </button>

        </div>

    </div><!-- /.page-wrapper -->

    <%
        If Dados.State = 1 Then Dados.Close
        Set Dados = Nothing

        If ConexaoFechamento.State = 1 Then ConexaoFechamento.Close
        Set ConexaoFechamento = Nothing
    %>

    <script>
        /* ──────────────────────────────────────────────────────────
           Estado global de paginação + filtro
        ────────────────────────────────────────────────────────── */
        var _todasAsLinhas   = [];   // Array com todas as <tr> originais
        var _linhasFiltradas = [];   // Subconjunto após filtro
        var _paginaAtual     = 1;
        var _rowsPerPage     = 100;
        var _totalOriginal   = 0;

        /* Inicialização: captura todas as linhas do tbody */
        (function init() {
            var tbody = document.getElementById('tabelaBody');
            if (!tbody) return;

            var trs = tbody.querySelectorAll('tr');
            _totalOriginal = trs.length;

            for (var i = 0; i < trs.length; i++) {
                _todasAsLinhas.push(trs[i]);
            }

            _linhasFiltradas = _todasAsLinhas.slice();
            renderizarPagina();
        })();

        /* ── Filtro dinâmico ─────────────────────────────────────── */
        function filtrarTabela() {
            var termo = document.getElementById('filtroTabela').value.toLowerCase().trim();

            if (termo === '') {
                _linhasFiltradas = _todasAsLinhas.slice();
            } else {
                _linhasFiltradas = [];
                for (var i = 0; i < _todasAsLinhas.length; i++) {
                    var texto = _todasAsLinhas[i].textContent.toLowerCase();
                    if (texto.indexOf(termo) !== -1) {
                        _linhasFiltradas.push(_todasAsLinhas[i]);
                    }
                }
            }

            _paginaAtual = 1;
            renderizarPagina();
        }

        /* ── Mudança de rows per page ────────────────────────────── */
        function mudarRowsPerPage() {
            _rowsPerPage = parseInt(document.getElementById('rowsPerPage').value, 10);
            _paginaAtual = 1;
            renderizarPagina();
        }

        /* ── Renderização da página atual ────────────────────────── */
        function renderizarPagina() {
            var totalFiltrado = _linhasFiltradas.length;
            var totalPaginas  = Math.max(1, Math.ceil(totalFiltrado / _rowsPerPage));

            if (_paginaAtual > totalPaginas) _paginaAtual = totalPaginas;

            var inicio = (_paginaAtual - 1) * _rowsPerPage;
            var fim    = Math.min(inicio + _rowsPerPage, totalFiltrado);

            /* Esconde todas */
            for (var i = 0; i < _todasAsLinhas.length; i++) {
                _todasAsLinhas[i].style.display = 'none';
            }

            /* Mostra as da página atual */
            for (var j = inicio; j < fim; j++) {
                _linhasFiltradas[j].style.display = '';
            }

            /* Linha "sem resultados" */
            var noResults = document.getElementById('noResultsRow');
            if (totalFiltrado === 0) {
                if (!noResults) {
                    var tbody = document.getElementById('tabelaBody');
                    var tr    = document.createElement('tr');
                    tr.id     = 'noResultsRow';
                    tr.className = 'no-results-row';
                    var cols = document.querySelector('#tabelaDados thead tr').querySelectorAll('th').length;
                    tr.innerHTML = '<td colspan="' + cols + '">Nenhum registro encontrado para o filtro aplicado.</td>';
                    tbody.appendChild(tr);
                }
            } else {
                if (noResults) noResults.parentNode.removeChild(noResults);
            }

            /* Atualiza contador */
            var exibidos = fim - inicio;
            document.getElementById('contadorVisiveis').textContent = totalFiltrado > 0
                ? (inicio + 1) + '–' + fim
                : '0';

            /* Reconstrói paginação */
            construirPaginacao(totalPaginas);
        }

        /* ── Construção dos botões de página ─────────────────────── */
        function construirPaginacao(totalPaginas) {
            var container = document.getElementById('paginacaoContainer');
            container.innerHTML = '';

            if (totalPaginas <= 1) return;

            /* Botão anterior */
            var btnPrev = criarBotaoPagina('‹', _paginaAtual === 1, function() {
                irParaPagina(_paginaAtual - 1);
            });
            container.appendChild(btnPrev);

            /* Janela de páginas: mostra até 7 botões centrados na página atual */
            var janela = 7;
            var inicio = Math.max(1, _paginaAtual - Math.floor(janela / 2));
            var fim    = Math.min(totalPaginas, inicio + janela - 1);
            inicio     = Math.max(1, fim - janela + 1);

            if (inicio > 1) {
                container.appendChild(criarBotaoPagina('1', false, function() { irParaPagina(1); }));
                if (inicio > 2) {
                    var ell = document.createElement('span');
                    ell.className = 'page-info';
                    ell.textContent = '…';
                    container.appendChild(ell);
                }
            }

            for (var p = inicio; p <= fim; p++) {
                (function(pg) {
                    var btn = criarBotaoPagina(String(pg), false, function() { irParaPagina(pg); });
                    if (pg === _paginaAtual) btn.classList.add('active');
                    container.appendChild(btn);
                })(p);
            }

            if (fim < totalPaginas) {
                if (fim < totalPaginas - 1) {
                    var ell2 = document.createElement('span');
                    ell2.className = 'page-info';
                    ell2.textContent = '…';
                    container.appendChild(ell2);
                }
                container.appendChild(criarBotaoPagina(String(totalPaginas), false, function() { irParaPagina(totalPaginas); }));
            }

            /* Botão próximo */
            var btnNext = criarBotaoPagina('›', _paginaAtual === totalPaginas, function() {
                irParaPagina(_paginaAtual + 1);
            });
            container.appendChild(btnNext);
        }

        function criarBotaoPagina(texto, desabilitado, onClick) {
            var btn = document.createElement('button');
            btn.type      = 'button';
            btn.className = 'page-btn';
            btn.textContent = texto;
            btn.disabled  = desabilitado;
            btn.addEventListener('click', onClick);
            return btn;
        }

        function irParaPagina(p) {
            _paginaAtual = p;
            renderizarPagina();
        }

        /* ── Exportar PDF (jsPDF + autoTable) ───────────────────── */
        function exportarPDF() {
            var tabela = document.getElementById('tabelaDados');
            if (!tabela) return;

            var { jsPDF } = window.jspdf;
            /* compress:true reduz significativamente o tamanho do PDF gerado */
            var doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });

            var pageW = doc.internal.pageSize.getWidth();
            var margin = 14;

            /* ── Logo ─────────────────────────────────────────────── */
            var logoPath = '\\corporativo\\sesi\\Logo\\logo-smurfit1.png';
            // Carrega a logo como imagem via elemento <img> temporário
            // Em ambiente ASP a URL relativa é mapeada pelo servidor;
            // aqui construímos o src relativo para o fetch do browser.
            var logoSrc = '/corporativo/sesi/Logo/logo-smurfit1.png';
            var img = new Image();
            img.crossOrigin = 'anonymous';

            function gerarComLogo(dataUrl) {
                var curY = margin;

                if (dataUrl) {
                    try {
                        // Logo no canto superior esquerdo (h=10mm, proporcional)
                        var tmpImg = new Image();
                        tmpImg.src = dataUrl;
                        var ratio = tmpImg.naturalWidth / tmpImg.naturalHeight || 3;
                        var logoH = 10;
                        var logoW = logoH * ratio;
                        doc.addImage(dataUrl, 'JPEG', margin, curY, logoW, logoH);
                    } catch(e) { /* ignora se falhar */ }
                }

                /* ── Títulos ─────────────────────────────────────── */
                doc.setFontSize(13);
                doc.setFont('helvetica', 'bold');
                doc.setTextColor(30, 30, 30);
                doc.text('Fechamento Refeitório', pageW / 2, curY + 5, { align: 'center' });

                /* Subtítulo */
                var tipo = '<%=trim(request("tipo"))%>';
                var dtIni = '<%=FormatarDataBR(trim(request("dtini")))%>';
                var dtFim = '<%=FormatarDataBR(trim(request("dtfim")))%>';
                var subtitulo = (tipo === '1' ? 'Relatório Detalhado' : 'Relatório Consolidado') +
                                ' — Período: ' + dtIni + ' a ' + dtFim;
                doc.setFontSize(8);
                doc.setFont('helvetica', 'normal');
                doc.setTextColor(100, 100, 100);
                doc.text(subtitulo, pageW / 2, curY + 11, { align: 'center' });

                curY += 18;

                /* ── Cabeçalhos da tabela ────────────────────────── */
                var ths = tabela.querySelectorAll('thead th');
                var head = [[]];
                for (var h = 0; h < ths.length; h++) {
                    head[0].push(ths[h].textContent.trim());
                }

                /* ── Dados (todas as linhas, sem paginação) ──────── */
                var body = [];
                for (var i = 0; i < _todasAsLinhas.length; i++) {
                    var tds = _todasAsLinhas[i].querySelectorAll('td');
                    var row = [];
                    for (var c = 0; c < tds.length; c++) {
                        row.push(tds[c].textContent.trim());
                    }
                    body.push(row);
                }

                /* ── Gera tabela ─────────────────────────────────── */
                /* Larguras fixas por tipo de relatório para alinhar cabeçalho e dados */
                var tipoRel = tipo; /* reutiliza a variável já declarada acima */
                var colStyles = {};
                if (tipoRel === '1') {
                    /* Detalhado: Credencial | Matrícula | Nome | Ano | Mês | Dia | Hora | Minuto | Valor */
                    colStyles = {
                        0: { halign: 'center', cellWidth: 22 },
                        1: { halign: 'center', cellWidth: 20 },
                        2: { halign: 'center', cellWidth: 'auto' },
                        3: { halign: 'center', cellWidth: 14 },
                        4: { halign: 'center', cellWidth: 12 },
                        5: { halign: 'center', cellWidth: 12 },
                        6: { halign: 'center', cellWidth: 14 },
                        7: { halign: 'center', cellWidth: 16 },
                        8: { halign: 'right',  cellWidth: 20 }
                    };
                } else {
                    /* Consolidado: Credencial | Matrícula | Nome | Valor Total */
                    colStyles = {
                        0: { halign: 'center', cellWidth: 26 },
                        1: { halign: 'center', cellWidth: 22 },
                        2: { halign: 'center', cellWidth: 'auto' },
                        3: { halign: 'center', cellWidth: 28 }
                    };
                }

                doc.autoTable({
                    head: head,
                    body: body,
                    startY: curY,
                    margin: { left: margin, right: margin },
                    tableLineColor: [180, 180, 180],
                    tableLineWidth: 0.2,
                    styles: {
                        fontSize: 7,
                        cellPadding: { top: 2, right: 3, bottom: 2, left: 3 },
                        valign: 'middle',
                        overflow: 'linebreak',
                        lineColor: [200, 200, 200],
                        lineWidth: 0.2
                    },
                    headStyles: {
                        fillColor: [41, 98, 163],
                        textColor: 255,
                        fontStyle: 'bold',
                        halign: 'center',
                        lineColor: [255, 255, 255],
                        lineWidth: 0.3
                    },
                    alternateRowStyles: { fillColor: [245, 246, 250] },
                    columnStyles: colStyles,
                    didDrawCell: function(data) {
                        /* Linha vertical direita em cada célula (exceto última coluna) */
                        var colCount = data.table.columns.length;
                        if (data.column.index < colCount - 1) {
                            var x = data.cell.x + data.cell.width;
                            var y = data.cell.y;
                            var h = data.cell.height;
                            doc.setDrawColor(180, 180, 180);
                            doc.setLineWidth(0.2);
                            doc.line(x, y, x, y + h);
                        }
                    },
                    didDrawPage: function(data) {
                        /* Rodapé com nº de página */
                        doc.setFontSize(7);
                        doc.setTextColor(150);
                        var pgStr = 'Página ' + data.pageNumber;
                        doc.text(pgStr, pageW - margin, doc.internal.pageSize.getHeight() - 6, { align: 'right' });

                        /* Data de geração */
                        var agora = new Date();
                        var pad = function(n){ return n < 10 ? '0'+n : n; };
                        var dataStr = 'Gerado em: ' + pad(agora.getDate()) + '/' + pad(agora.getMonth()+1) + '/' + agora.getFullYear() +
                                      ' ' + pad(agora.getHours()) + ':' + pad(agora.getMinutes());
                        doc.text(dataStr, margin, doc.internal.pageSize.getHeight() - 6);
                    }
                });

                doc.save('FechamentoRefeitorio.pdf');
            }

            /* Tenta carregar a logo; gera o PDF independente do resultado */
            img.onload = function() {
                var canvas = document.createElement('canvas');
                /* Limita a resolução da logo a 300px de largura para reduzir o peso do PDF.
                   Logos maiores são redimensionadas proporcionalmente. */
                var maxW = 300;
                var scale = img.naturalWidth > maxW ? maxW / img.naturalWidth : 1;
                canvas.width  = Math.round(img.naturalWidth  * scale);
                canvas.height = Math.round(img.naturalHeight * scale);
                canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
                /* JPEG com qualidade 0.65 é ~5–10× menor que PNG sem perda visual perceptível */
                gerarComLogo(canvas.toDataURL('image/jpeg', 0.65));
            };
            img.onerror = function() { gerarComLogo(null); };
            img.src = logoSrc;
        }

        /* ── Exportar Excel (CSV UTF-8 com BOM) ──────────────────── */
        function exportarCSV() {
            var tabela = document.getElementById('tabelaDados');
            if (!tabela) return;

            var esc = function(v) {
                var s = v == null ? '' : String(v).trim();
                return (s.indexOf(';') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
                    ? '"' + s.replace(/"/g, '""') + '"' : s;
            };

            var linhas = [];

            /* Cabeçalhos */
            var ths = tabela.querySelectorAll('thead th');
            var hdrs = [];
            for (var h = 0; h < ths.length; h++) hdrs.push(esc(ths[h].textContent));
            linhas.push(hdrs.join(';'));

            /* Todas as linhas (sem filtro, sem paginação) */
            for (var i = 0; i < _todasAsLinhas.length; i++) {
                var tds = _todasAsLinhas[i].querySelectorAll('td');
                var cols = [];
                for (var c = 0; c < tds.length; c++) cols.push(esc(tds[c].textContent));
                linhas.push(cols.join(';'));
            }

            var blob = new Blob(['\uFEFF' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
            var url  = URL.createObjectURL(blob);
            var a    = document.createElement('a');
            a.href     = url;
            a.download = 'FechamentoRefeitorio.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    </script>

    <!-- Bootstrap 5.3 JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
