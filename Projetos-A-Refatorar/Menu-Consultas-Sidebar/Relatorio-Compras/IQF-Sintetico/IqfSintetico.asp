<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer  = True
Response.Charset = "utf-8"
Response.CodePage = 65001
Session.LCID = 1046
%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp"-->
<%
' ============================================================
' BLOCO DE API JSON
' Ativado quando o parâmetro "acao" está presente na requisição
' ============================================================
If Trim(Request("acao")) <> "" Then

    Response.ContentType = "application/json"
    Response.Expires     = 0

    Dim acao
    acao = Trim(Request("acao"))

    Select Case acao
        Case "listar"
            Response.Write(Listar())
        Case Else
            Response.Write("{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}")
    End Select

    Response.End()

End If

' ============================================================
' HELPERS JSON LOCAIS
' EscapeJSON de funcoes.asp tem bug conhecido — usar local
' ============================================================
Function EscapeJSONLocal(s)
    Dim v
    v = CStr(s)
    v = Replace(v, "\",  "\\")
    v = Replace(v, """", "\""")
    v = Replace(v, Chr(13) & Chr(10), "\n")
    v = Replace(v, Chr(10), "\n")
    v = Replace(v, Chr(13), "\n")
    v = Replace(v, Chr(9),  "\t")
    EscapeJSONLocal = v
End Function

Function JsonNum(v)
    If IsNull(v) Or IsEmpty(v) Then
        JsonNum = "0"
    Else
        JsonNum = Replace(CStr(CDbl(v)), ",", ".")
    End If
End Function

' ============================================================
' CONEXÃO ÚNICA
' vw_MSNotasAnalisadas_Fornecedores, FornecedoresInsumos e
' vw_MSIQF existem apenas em PIR-CORPORATIVO.
' O filtro por filial é feito via CO_FILIAL na query.
' ============================================================
Const CONEXAO_IQF = "PIR-CORPORATIVO"

' ============================================================
' ENDPOINT: listar
' Retorna array JSON com os registros de IQF Sintético
' ============================================================
Function Listar()
    Dim dataIni, dataFim, filialParam, fornecedor
    Dim filialCodigo, SQL
    Dim rs, jsonArray, primeiroItem

    dataIni    = Trim(Request("DataIni"))
    dataFim    = Trim(Request("DataFim"))
    filialParam = Trim(Request("Filial"))
    fornecedor = Trim(Request("Fornecedor"))

    If dataIni = "" Or dataFim = "" Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Informe a Data Inicial e a Data Final.""}"
        Exit Function
    End If

    filialCodigo = ""
    Select Case UCase(filialParam)
        Case "B" : filialCodigo = "'05'"
        Case "F" : filialCodigo = "'06'"
        Case "P" : filialCodigo = "'00'"
        Case "U" : filialCodigo = "'04'"
        Case "S" : filialCodigo = "'07'"
    End Select


    SQL  = "SELECT vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE, vw_MSNotasAnalisadas_Fornecedores.CO_LOJA, "
    SQL = SQL & "SUBSTRING(Razao,1,20) + '...' A2_NREDUZ, Razao A2_NOME, "
    SQL = SQL & "SUM(vw_MSNotasAnalisadas_Fornecedores.PONTOS)               AS SOMA_PONTOS, "
    SQL = SQL & "SUM(vw_MSNotasAnalisadas_Fornecedores.PESO)                 AS SOMA_ENCONTRADA, "
    SQL = SQL & "AVG(ISNULL(NOTA,100))                                       AS IQF, "
    SQL = SQL & "SUM(ISNULL(vw_MSNotasAnalisadas_Fornecedores.CO_TOTAL,0))   AS TOTAL, "
    SQL = SQL & "COUNT(DISTINCT vw_MSNotasAnalisadas_Fornecedores.CO_DOC)    AS NFS "
    SQL = SQL & "FROM vw_MSNotasAnalisadas_Fornecedores "
    SQL = SQL & "INNER JOIN FornecedoresInsumos "
    SQL = SQL & "    ON vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE COLLATE SQL_Latin1_General_CP437_BIN = Codigo "
    SQL = SQL & "LEFT JOIN vw_MSIQF "
    SQL = SQL & "    ON FORNECEDOR = CO_FORNECE COLLATE SQL_Latin1_General_CP437_BIN "
    SQL = SQL & "    AND MES = SUBSTRING(CO_DATCONF,1,6) "
    SQL = SQL & "WHERE CONVERT(DATE, vw_MSNotasAnalisadas_Fornecedores.CO_DATCONF, 103) "
    SQL = SQL & "    BETWEEN '" & Replace(dataIni, "'", "''") & "' AND '" & Replace(dataFim, "'", "''") & "' "

    If filialCodigo <> "" Then
        SQL = SQL & "AND vw_MSNotasAnalisadas_Fornecedores.CO_FILIAL = " & filialCodigo & " "
    End If

    SQL = SQL & "GROUP BY vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE, "
    SQL = SQL & "         vw_MSNotasAnalisadas_Fornecedores.CO_LOJA, "
    SQL = SQL & "         SUBSTRING(Razao,1,20), Razao "

    If fornecedor <> "" Then
        SQL = SQL & "HAVING vw_MSNotasAnalisadas_Fornecedores.CO_FORNECE = '" & Replace(fornecedor, "'", "''") & "' "
    End If

    SQL = SQL & "ORDER BY IQF "

    On Error Resume Next
    Response.Clear()
    Set rs = Conecta(CONEXAO_IQF, SQL)

    If Err.Number <> 0 Or Not IsObject(rs) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar banco de dados: " & EscapeJSONLocal(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    primeiroItem = True
    jsonArray    = "["

    Do While Not rs.EOF
        If Not primeiroItem Then jsonArray = jsonArray & ","  End If

        Dim codFornece, codLoja, nomeReduzido, nomeCompleto
        Dim somaTotal, nfs, iqf

        codFornece   = EscapeJSONLocal(rs("CO_FORNECE") & "")
        codLoja      = EscapeJSONLocal(rs("CO_LOJA") & "")
        nomeCompleto = EscapeJSONLocal(rs("A2_NOME") & "")
        somaTotal    = JsonNum(rs("TOTAL"))
        nfs          = rs("NFS") & ""
        iqf          = JsonNum(rs("IQF"))

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """codigoCompleto"": """  & codFornece & """, "
        jsonArray = jsonArray & """fornecedor"": """       & nomeCompleto & """, "
        jsonArray = jsonArray & """total"": "              & somaTotal & ", "
        jsonArray = jsonArray & """nfs"": "                & nfs & ", "
        jsonArray = jsonArray & """iqf"": "                & iqf
        jsonArray = jsonArray & "}"

        primeiroItem = False
        rs.MoveNext
    Loop

    jsonArray = jsonArray & "]"

    rs.Close()
    Set rs = Nothing

    Listar = "{""sucesso"": true, ""mensagem"": """", ""dados"": " & jsonArray & "}"
End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>IQF - Sintético</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

    <style>
        body { background-color: #f0f0f0; min-height: 100vh; }

        .page-title {
            font-size: 1.15rem;
            font-weight: 700;
            color: #212529;
        }

        .select-filial { width: 145px; }

        /* Card de filtros */
        .filter-card {
            background-color: #fff;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
        }
        .filter-card-header {
            background-color: #fff;
            color: #212529;
            padding: .6rem 1rem;
            border-radius: 6px 6px 0 0;
            border-bottom: 1px solid #dee2e6;
            font-weight: 600;
            font-size: .9rem;
        }
        .filter-card-body { padding: 1rem 1.25rem 1.25rem; }
        .filter-card label {
            font-size: .82rem;
            font-weight: 600;
            color: #6c757d;
            margin-bottom: .25rem;
        }
        .input-data { width: 155px; }

        /* Card de tabela */
        .table-card {
            background-color: #fff;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            position: relative;
        }

        /* Toolbar */
        .table-toolbar { background-color: #fff; border-radius: 6px 6px 0 0; }

        /* Search box */
        .search-box { position: relative; max-width: 200px; }
        .search-box .form-control { padding-right: 76px; font-size: .82rem; }
        .search-box .search-hotkey {
            position: absolute; right: 8px; top: 50%;
            transform: translateY(-50%);
            pointer-events: none; user-select: none;
            display: flex; gap: 2px;
        }
        .search-box .search-hotkey kbd {
            font-size: .65rem; padding: .1rem .3rem;
            background-color: #f8f9fa; border: 1px solid #dee2e6;
            border-radius: 3px; color: #6c757d;
        }

        /* Placeholder estado vazio */
        .table-placeholder {
            display: flex; flex-direction: column;
            align-items: center; justify-content: center;
            padding: 3.5rem 1rem; color: #6c757d; gap: .75rem;
        }
        .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
        .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

        /* Wrapper de scroll */
        .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
        .table-wrapper table { min-width: max-content; }

        /* Cabeçalho sticky */
        .table thead th {
            position: sticky; top: 0;
            background-color: #fff; color: #212529;
            font-size: .82rem; font-weight: 700;
            text-align: center; white-space: nowrap;
            z-index: 2; border-color: #dee2e6;
            padding: .55rem .75rem;
        }
        .table tbody td {
            font-size: .82rem; text-align: center;
            white-space: nowrap; color: #212529;
            vertical-align: middle; padding: .45rem .75rem;
        }

        /* Rodapé da tabela */
        .table-footer {
            display: flex; align-items: center;
            justify-content: space-between;
            padding: .6rem 1rem;
            border-top: 1px solid #dee2e6;
            flex-wrap: wrap; gap: .5rem;
        }
        .table-footer-left  { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
        .table-footer-center { flex: 1; display: flex; justify-content: center; }
        .table-footer-right { min-width: 140px; text-align: right; }
        .info-registros { font-size: .8rem; color: #6c757d; }
        .select-limite  { width: auto; }

        /* Paginação */
        #paginacao .page-link                    { color: #6c757d; border-color: #dee2e6; }
        #paginacao .page-item.active .page-link  { background-color: #6c757d; border-color: #6c757d; color: #fff; }
        #paginacao .page-link:hover              { background-color: #e9ecef; color: #495057; }
        #paginacao .page-item.disabled .page-link{ color: #adb5bd; }

        /* Overlay */
        .overlay-loading {
            position: absolute; inset: 0;
            background: rgba(255,255,255,.80);
            display: flex; flex-direction: column;
            align-items: center; justify-content: center;
            z-index: 10; gap: 1rem; border-radius: 6px;
        }
        .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
        .overlay-loading span { font-size: .9rem; font-weight: 600; color: #495057; }

        /* Toast */
        .toast-container-custom {
            position: fixed; top: 1rem; right: 1rem;
            z-index: 1100; display: flex; flex-direction: column; gap: .5rem;
        }
        .toast-custom {
            min-width: 280px; max-width: 360px;
            border-radius: 10px; overflow: hidden;
            box-shadow: 0 4px 18px rgba(0,0,0,.13);
            background: #fff; border: 1px solid #dee2e6;
            animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
        }
        @keyframes toastSlideIn {
            from { opacity: 0; transform: translateX(40px); }
            to   { opacity: 1; transform: translateX(0); }
        }
        .toast-custom .toast-progress-bar {
            animation: toastProgressShrink 4.5s linear forwards;
        }
        @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }
    </style>
</head>
<body>
    <div class="container-fluid py-3 px-3 px-lg-4">

        <!-- Cabeçalho da página -->
        <div class="d-flex align-items-center mb-3 gap-2">
            <span class="page-title">Índice de Qualificação de Fornecimento - Sintético</span>
        </div>

        <!-- Card de Filtros -->
        <div class="filter-card mb-3">
            <div class="filter-card-header">
                <i class="bi bi-funnel me-2"></i>Filtros
            </div>
            <div class="filter-card-body">
                <div class="row g-3 align-items-end">

                    <div class="col-auto">
                        <label for="Filial">Filial</label>
                        <select id="Filial" class="form-select form-select-sm select-filial">
                            <option value="">Todas</option>
                            <option value="B">Bento Gonçalves</option>
                            <option value="F">Fortaleza</option>
                            <option value="P">Pirapetinga</option>
                            <option value="U">Uberaba</option>
                            <option value="S">Saquarema</option>
                        </select>
                    </div>

                    <div class="col-sm-auto">
                        <label for="DataIni">Data Inicial</label>
                        <input type="date" id="DataIni"
                               class="form-control form-control-sm input-data" />
                    </div>

                    <div class="col-sm-auto">
                        <label for="DataFim">Data Final</label>
                        <input type="date" id="DataFim"
                               class="form-control form-control-sm input-data" />
                    </div>

                    <div class="col-sm-2">
                        <label for="Fornecedor">Fornecedor</label>
                        <input type="text" id="Fornecedor"
                               class="form-control form-control-sm"
                               placeholder="Código do fornecedor" />
                    </div>

                    <div class="col-sm-auto ms-auto">
                        <div class="d-flex gap-2 justify-content-end">
                            <button type="button" id="btnLimpar" class="btn btn-light btn-sm">
                                <i class="bi bi-eraser me-2"></i>Limpar
                            </button>
                            <button type="button" id="btnFiltrar" class="btn btn-primary btn-sm">
                                <i class="bi bi-funnel-fill me-2"></i>Filtrar
                            </button>
                        </div>
                    </div>

                </div>
            </div>
        </div>

        <!-- Card de Tabela -->
        <div class="table-card">

            <!-- Overlay de carregamento -->
            <div id="overlayLoading" class="overlay-loading d-none">
                <div class="spinner-border" role="status">
                    <span class="visually-hidden">Carregando...</span>
                </div>
                <span>Carregando...</span>
            </div>

            <!-- Toolbar -->
            <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
                <div class="d-flex gap-2">
                    <button id="btnExportCsv" type="button" class="btn btn-success btn-sm" disabled>
                        <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
                    </button>
                    <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm" disabled>
                        <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
                    </button>
                </div>
                <div class="search-box">
                    <input type="text" id="searchBox" class="form-control form-control-sm"
                           placeholder="Buscar..." autocomplete="off" />
                    <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
                </div>
            </div>

            <!-- Área de conteúdo -->
            <div id="tablePlaceholder" class="table-placeholder">
                <i class="bi bi-clipboard-data"></i>
                <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
            </div>

            <div id="tableWrapper" class="table-wrapper d-none">
                <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaDados">
                    <thead>
                        <tr>
                            <th>Código</th>
                            <th>Fornecedor</th>
                            <th>Valor Total</th>
                            <th>Notas Conferidas</th>
                            <th>IQF Médio</th>
                        </tr>
                    </thead>
                    <tbody id="corpoTabela"></tbody>
                </table>
            </div>

            <!-- Rodapé da tabela -->
            <div id="tableFooter" class="table-footer d-none">
                <div class="table-footer-left">
                    <label for="selectLimite" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">
                        Linhas por página:
                    </label>
                    <select id="selectLimite" class="form-select form-select-sm select-limite">
                        <option value="10">10</option>
                        <option value="25">25</option>
                        <option value="50" selected>50</option>
                        <option value="100">100</option>
                        <option value="250">250</option>
                        <option value="500">500</option>
                    </select>
                </div>
                <div class="table-footer-center">
                    <nav aria-label="Navegação de páginas">
                        <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
                    </nav>
                </div>
                <div class="table-footer-right">
                    <span id="infoRegistros" class="info-registros"></span>
                </div>
            </div>

        </div>
    </div>

    <!-- Container de toasts -->
    <div class="toast-container-custom" id="toastContainer"></div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <script>
    "use strict";

    const SWNET_LOGO     = 'http://<%=Request.ServerVariables("SERVER_NAME")%>/corporativo/suprimentos/Logo/logo-smurfit1.png';
    const SWNET_OPERADOR = '<%=EscapeJSONLocal(USR & " - " & NOME)%>';
    const REPOSITORIO    = 'IqfSintetico.asp';

    const IqfSintetico = (function () {

        /* ── constantes ── */
        const TOAST_DELAY = 4500;
        const TOAST_ICON  = Object.freeze({
            sucesso : 'bi-check-circle-fill',
            erro    : 'bi-x-circle-fill',
            aviso   : 'bi-exclamation-triangle-fill',
            info    : 'bi-info-circle-fill'
        });
        const TOAST_COR = Object.freeze({
            sucesso : '#198754',
            erro    : '#dc3545',
            aviso   : '#856404',
            info    : '#6c757d'
        });
        const TOAST_BARRA = Object.freeze({
            sucesso : '#198754',
            erro    : '#dc3545',
            aviso   : '#ffc107',
            info    : '#6c757d'
        });

        /* ── estado ── */
        let dadosCompletos  = [];
        let dadosFiltrados  = [];
        let paginaAtual     = 1;
        let itensPorPagina  = 50;

        /* ── helpers ── */
        const $el = id => document.getElementById(id);

        function exibirToast(tipo, mensagem) {
            const icon  = TOAST_ICON[tipo]  || TOAST_ICON.info;
            const cor   = TOAST_COR[tipo]   || TOAST_COR.info;
            const barra = TOAST_BARRA[tipo] || TOAST_BARRA.info;
            const el    = document.createElement('div');
            el.className = 'toast-custom p-0';
            el.innerHTML = `
                <div class="d-flex align-items-start px-3 pt-3 pb-2 gap-2">
                    <i class="bi ${icon} mt-1" style="color:${cor};font-size:1rem;"></i>
                    <div class="flex-grow-1" style="font-size:.88rem;">${mensagem}</div>
                    <button type="button" class="btn-close btn-sm"
                            onclick="this.closest('.toast-custom').remove()"></button>
                </div>
                <div style="height:3px;background:#e9ecef;border-radius:0 0 10px 10px;">
                    <div class="toast-progress-bar"
                         style="height:100%;background:${barra};border-radius:inherit;"></div>
                </div>`;
            $el('toastContainer').appendChild(el);
            setTimeout(() => el.remove(), TOAST_DELAY + 300);
        }

        /* ── formata número pt-BR ── */
        function fmtNum(v, casas) {
            const n = parseFloat(v);
            if (isNaN(n)) return '0,00';
            return n.toLocaleString('pt-BR', {
                minimumFractionDigits: casas,
                maximumFractionDigits: casas
            });
        }

        /* ── constrói linhas TR a partir dos dados ── */
        function _criarLinha(item) {
            const tr = document.createElement('tr');
            tr.dataset.busca = (item.codigoCompleto + ' ' + item.fornecedor).toLowerCase();

            const tds = [
                item.codigoCompleto,
                item.fornecedor,
                fmtNum(item.total, 2),
                item.nfs,
                '<strong>' + fmtNum(item.iqf, 2) + '%</strong>'
            ];

            tds.forEach(function(conteudo) {
                const td = document.createElement('td');
                td.innerHTML = conteudo;
                tr.appendChild(td);
            });

            return tr;
        }

        /* ── renderiza tabela com paginação ── */
        function _renderTabela() {
            const inicio  = (paginaAtual - 1) * itensPorPagina;
            const fim     = inicio + itensPorPagina;
            const $tbody  = $('#corpoTabela').empty();
            const frag    = document.createDocumentFragment();

            dadosFiltrados.slice(inicio, fim).forEach(function(item) {
                frag.appendChild(_criarLinha(item));
            });

            $tbody[0].appendChild(frag);

            const totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / itensPorPagina));
            _renderPaginacao(paginaAtual, totalPaginas);
            $('#infoRegistros').text('Total: ' + dadosFiltrados.length + ' registro(s)');
        }

        function _renderPaginacao(atual, total) {
            const $ul = $('#paginacao').empty();
            if (total <= 1) return;

            const mkItem = function(label, page, disabled, active) {
                const $li = $('<li>').addClass('page-item');
                if (disabled) $li.addClass('disabled');
                if (active)   $li.addClass('active');
                const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
                if (!disabled && !active) {
                    $a.on('click', function(e) {
                        e.preventDefault();
                        paginaAtual = page;
                        _renderTabela();
                    });
                }
                return $li.append($a);
            };

            $ul.append(mkItem('&laquo;', 1,         atual === 1,     false));
            $ul.append(mkItem('&lsaquo;', atual - 1, atual === 1,    false));

            const inicio = Math.max(1, atual - 2);
            const fim    = Math.min(total, atual + 2);
            for (let p = inicio; p <= fim; p++) {
                $ul.append(mkItem(p, p, false, p === atual));
            }

            $ul.append(mkItem('&rsaquo;', atual + 1, atual === total, false));
            $ul.append(mkItem('&raquo;', total,      atual === total, false));
        }

        /* ── filtro dinâmico (search box) ── */
        function _aplicarFiltro(termo) {
            const t = termo.toLowerCase().trim();
            dadosFiltrados = t === ''
                ? dadosCompletos.slice()
                : dadosCompletos.filter(function(item) {
                    return (item.codigoCompleto + ' ' + item.fornecedor)
                               .toLowerCase().includes(t);
                });
            paginaAtual = 1;
            _renderTabela();
        }

        /* ── mostrar / ocultar áreas ── */
        function _mostrarTabela() {
            $('#tablePlaceholder').addClass('d-none');
            $('#tableWrapper').removeClass('d-none');
            $('#tableFooter').removeClass('d-none');
            $('#btnExportCsv').prop('disabled', false);
            $('#btnExportPdf').prop('disabled', false);
        }

        function _mostrarPlaceholder(mensagem) {
            $('#tablePlaceholder')
                .removeClass('d-none')
                .find('p')
                .html(mensagem || 'Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.');
            $('#tableWrapper').addClass('d-none');
            $('#tableFooter').addClass('d-none');
            $('#btnExportCsv').prop('disabled', true);
            $('#btnExportPdf').prop('disabled', true);
        }

        /* ── buscar dados via AJAX ── */
        function _filtrar() {
            const dataIni  = $('#DataIni').val().trim();
            const dataFim  = $('#DataFim').val().trim();

            if (!dataIni || !dataFim) {
                exibirToast('aviso', 'Informe a Data Inicial e a Data Final.');
                return;
            }

            $('#overlayLoading').removeClass('d-none');

            $.ajax({
                url    : REPOSITORIO,
                method : 'POST',
                data   : {
                    acao       : 'listar',
                    DataIni    : dataIni,
                    DataFim    : dataFim,
                    Filial     : $('#Filial').val(),
                    Fornecedor : $('#Fornecedor').val().trim()
                },
                dataType : 'json',
                success  : function(resp) {
                    $('#overlayLoading').addClass('d-none');

                    if (!resp.sucesso) {
                        exibirToast('erro', resp.mensagem || 'Erro ao carregar dados.');
                        _mostrarPlaceholder();
                        return;
                    }

                    dadosCompletos = resp.dados || [];
                    dadosFiltrados = dadosCompletos.slice();
                    paginaAtual    = 1;
                    $('#searchBox').val('');

                    if (dadosCompletos.length === 0) {
                        _mostrarPlaceholder('Nenhum registro encontrado para os filtros informados.');
                        return;
                    }

                    _mostrarTabela();
                    _renderTabela();
                },
                error : function() {
                    $('#overlayLoading').addClass('d-none');
                    exibirToast('erro', 'Erro de comunicação com o servidor.');
                    _mostrarPlaceholder();
                }
            });
        }

        /* ── exportação CSV ── */
        function _exportarCSV() {
            const cabecalhos = ['Código', 'Fornecedor', 'Valor Total', 'Notas Conferidas', 'IQF Médio'];

            const esc = function(v) {
                const s = v == null ? '' : String(v);
                return (s.includes(';') || s.includes('"') || s.includes('\n'))
                    ? '"' + s.replace(/"/g, '""') + '"' : s;
            };

            const linhas = dadosFiltrados.map(function(item) {
                return [
                    item.codigoCompleto,
                    item.fornecedor,
                    fmtNum(item.total, 2),
                    item.nfs,
                    fmtNum(item.iqf, 2) + '%'
                ].map(esc).join(';');
            });

            const rows = [cabecalhos.map(esc).join(';'), ...linhas];
            const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href     = url;
            a.download = 'iqf-sintetico.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        /* ── exportação PDF ── */
        function _exportarPDF() {
            const { jsPDF } = window.jspdf;
            const doc    = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
            const pageW  = doc.internal.pageSize.getWidth();
            const img    = new Image();
            img.src      = SWNET_LOGO;
            const canvas = document.createElement('canvas');
            canvas.width = 900;

            img.onload = function() {
                const ratio   = img.height / img.width;
                canvas.height = Math.round(900 * ratio);
                canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
                doc.addImage(canvas.toDataURL('image/png'), 'PNG', 10, 4, 38, Math.round(38 * ratio));
                _gerarAutoTable(doc, pageW);
            };
            img.onerror = function() {
                _gerarAutoTable(doc, pageW);
            };
        }

        function _gerarAutoTable(doc, pageW) {
            const agora    = new Date();
            const pad      = function(n) { return String(n).padStart(2, '0'); };
            const dataHora = pad(agora.getDate()) + '/' + pad(agora.getMonth() + 1) + '/' + agora.getFullYear()
                           + ' ' + pad(agora.getHours()) + ':' + pad(agora.getMinutes()) + ':' + pad(agora.getSeconds());

            doc.setFontSize(13);
            doc.text('Índice de Qualificação de Fornecimento - Sintético', pageW / 2, 14, { align: 'center' });
            doc.setFontSize(8);
            doc.text('Relatório gerado pelo SWnet', pageW / 2, 20, { align: 'center' });
            doc.setFontSize(7);
            doc.text('Emitido por: ' + SWNET_OPERADOR + '   |   ' + dataHora, pageW - 10, 8, { align: 'right' });

            const cabecalhos = [['Código', 'Fornecedor', 'Valor Total', 'Notas Conferidas', 'IQF Médio']];
            const corpo = dadosFiltrados.map(function(item) {
                return [
                    item.codigoCompleto,
                    item.fornecedor,
                    fmtNum(item.total, 2),
                    String(item.nfs),
                    fmtNum(item.iqf, 2) + '%'
                ];
            });

            doc.autoTable({
                startY            : 26,
                theme             : 'grid',
                head              : cabecalhos,
                body              : corpo,
                headStyles        : { fillColor: [74, 74, 74], textColor: [255, 255, 255], fontSize: 4.5, cellPadding: 1.5 },
                alternateRowStyles: { fillColor: [245, 245, 245] },
                bodyStyles        : { fontSize: 4.5, cellPadding: 1.5, textColor: [33, 37, 41] },
                tableLineColor    : [180, 180, 180],
                tableLineWidth    : 0.2
            });

            doc.save('iqf-sintetico.pdf');
        }

        /* ── inicialização ── */
        function init() {

            $('#btnFiltrar').on('click', function() {
                _filtrar();
            });

            $('#btnLimpar').on('click', function() {
                $('#Filial').val('');
                $('#DataIni').val('');
                $('#DataFim').val('');
                $('#Fornecedor').val('');
                $('#searchBox').val('');
                dadosCompletos = [];
                dadosFiltrados = [];
                _mostrarPlaceholder();
            });

            let debounce;
            $('#searchBox').on('input', function() {
                clearTimeout(debounce);
                const termo = $(this).val();
                debounce = setTimeout(function() { _aplicarFiltro(termo); }, 250);
            });

            document.addEventListener('keydown', function(e) {
                if (e.ctrlKey && e.key === 'k') {
                    e.preventDefault();
                    $el('searchBox') && $el('searchBox').focus();
                }
            });

            $('#selectLimite').on('change', function() {
                itensPorPagina = parseInt($(this).val(), 10);
                paginaAtual    = 1;
                _renderTabela();
            });

            $('#btnExportCsv').on('click', function() { _exportarCSV(); });
            $('#btnExportPdf').on('click', function() { _exportarPDF(); });
        }

        return Object.freeze({ init });
    })();

    $(document).ready(function() {
        IqfSintetico.init();
    });
    </script>
</body>
</html>
