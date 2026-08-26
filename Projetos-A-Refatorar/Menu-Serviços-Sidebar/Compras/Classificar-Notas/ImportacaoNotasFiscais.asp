<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Session.LCID     = 1046

' ============================================================
' DISPATCHER — responde JSON quando "acao" está presente
' ============================================================
Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.Expires     = 0
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "importar" : resultado = Importar()
        Case Else
            resultado = "{""sucesso"":false,""mensagem"":""Ação inválida: " & EscapeJSON(acao) & """,""dados"":null}"
    End Select

    Response.Write resultado
    Response.End()

End If

' ============================================================
' FUNÇÕES DE NEGÓCIO
' ============================================================

' ------------------------------------------------------------
' Importar — lê o Excel, grava no SQL Server, devolve JSON
' ------------------------------------------------------------
Function Importar()

    Const CAMINHO_EXCEL    = "C:\inetpub\wwwroot\corporativo\suprimentos\Arquivos\Notas.xlsx"
    Const CONN_CORPORATIVO = "PIR-CORPORATIVO"
    Const USR_IMPORTACAO   = "SWNET"

    Dim dataHora, dataInvertida
    dataInvertida = InverteData(Date())
    dataHora      = dataInvertida & " " & Time()

    ' ----------------------------------------------------------
    ' 1. Abre o arquivo Excel via ADODB direto
    ' ----------------------------------------------------------
    Dim strConnExcel
    strConnExcel = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & CAMINHO_EXCEL & ";" & _
                   "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

    Dim excelConn, rsExcel
    On Error Resume Next

    Set excelConn = Server.CreateObject("ADODB.Connection")
    excelConn.Open strConnExcel

    If Err.Number <> 0 Or (excelConn Is Nothing) Then
        Importar = "{""sucesso"":false,""mensagem"":""Não foi possível abrir o arquivo Excel. Verifique se o arquivo existe em: " & EscapeJSON(CAMINHO_EXCEL) & """,""dados"":null}"
        Err.Clear
        Exit Function
    End If

    ' ----------------------------------------------------------
    ' Descobre o nome real da primeira aba via OpenSchema
    ' (evita falha por nomes como Sheet1, Planilha1, etc.)
    ' ----------------------------------------------------------
    Dim rsSchema, nomeAba, queryExcel
    On Error Resume Next

    Set rsSchema = excelConn.OpenSchema(20) ' adSchemaTables = 20

    If Err.Number <> 0 Or (rsSchema Is Nothing) Then
        Importar = "{""sucesso"":false,""mensagem"":""Não foi possível listar as abas do arquivo Excel: " & EscapeJSON(Err.Description) & """,""dados"":null}"
        Err.Clear
        excelConn.Close
        Set excelConn = Nothing
        Exit Function
    End If

    On Error GoTo 0

    nomeAba = ""
    Do While Not rsSchema.EOF And nomeAba = ""
        Dim tNome
        tNome = rsSchema("TABLE_NAME") & ""
        ' O driver ACE lista as abas com $ no final (ex.: "Plan1$")
        ' Ignora entradas de named ranges que não terminam em $
        If Right(tNome, 1) = "$" Then
            nomeAba = tNome
        End If
        rsSchema.MoveNext
    Loop

    rsSchema.Close
    Set rsSchema = Nothing

    If nomeAba = "" Then
        Importar = "{""sucesso"":false,""mensagem"":""Nenhuma aba encontrada no arquivo Excel."",""dados"":null}"
        excelConn.Close
        Set excelConn = Nothing
        Exit Function
    End If

    ' Monta a query usando o nome real da aba (ex.: "Plan1$" → [Plan1$A1:O100000])
    Dim nomeAbaBase
    nomeAbaBase  = Left(nomeAba, Len(nomeAba) - 1) ' remove o $ final
    queryExcel   = "SELECT * FROM [" & nomeAbaBase & "$A1:O100000]"

    On Error Resume Next

    Set rsExcel = Server.CreateObject("ADODB.Recordset")
    rsExcel.Open queryExcel, excelConn

    If Err.Number <> 0 Or (rsExcel Is Nothing) Then
        Importar = "{""sucesso"":false,""mensagem"":""Erro ao ler a planilha Excel: " & EscapeJSON(Err.Description) & """,""dados"":null}"
        Err.Clear
        excelConn.Close
        Set excelConn = Nothing
        Exit Function
    End If

    On Error GoTo 0

    ' ----------------------------------------------------------
    ' 2. Itera linhas do Excel e monta o bloco SQL
    ' ----------------------------------------------------------
    Dim SQL1
    SQL1 = ""

    Dim material, centro, deposito, tipoMovimento, docMaterial
    Dim item, dataLancamento, quantidade, um, descricao
    Dim notaFiscal, serie, codigoFornecedor, notaSerie, campoNota

    Do Until rsExcel.EOF

        material      = rsExcel.Fields(0).Value & ""
        centro        = rsExcel.Fields(1).Value & ""
        deposito      = rsExcel.Fields(2).Value & ""
        tipoMovimento = rsExcel.Fields(3).Value & ""
        docMaterial   = rsExcel.Fields(4).Value & ""
        item          = rsExcel.Fields(5).Value & ""

        dataLancamento = rsExcel.Fields(6).Value & ""
        If dataLancamento <> "" Then
            dataLancamento = InverteData(dataLancamento)
        End If

        quantidade = Replace(Replace(rsExcel.Fields(7).Value & "", ".", ""), ",", ".")
        um         = rsExcel.Fields(8).Value & ""
        descricao  = rsExcel.Fields(9).Value & ""

        campoNota = rsExcel.Fields(10).Value & ""
        If InStr(campoNota, "-") > 0 Then
            notaSerie  = Split(campoNota, "-")
            notaFiscal = notaSerie(0)
            serie      = notaSerie(1)
        Else
            notaFiscal = campoNota
            serie      = ""
        End If

        codigoFornecedor = rsExcel.Fields(11).Value & ""

        SQL1 = SQL1 & "IF NOT EXISTS(SELECT NotaFiscal FROM NotasFiscaisInsumos WHERE CodigoFornecedor = '" & codigoFornecedor & "' AND NotaFiscal = '" & notaFiscal & "' AND Serie = '" & serie & "' AND Item = '" & item & "') "
        SQL1 = SQL1 & "INSERT INTO [dbo].[NotasFiscaisInsumos]([Material],[Centro],[Deposito],[TipoMovimnto],[DocMaterial],[Item],[DataLancamento],[Quantidade],[UM],[Descricao],[NotaFiscal],[Serie],[CodigoFornecedor],[DataHoraIntegracao],[Usuario]) VALUES"
        SQL1 = SQL1 & "('" & material & "','" & centro & "','" & deposito & "','" & tipoMovimento & "','" & docMaterial & "','" & item & "','" & dataLancamento & "'," & quantidade & ",'" & um & "','" & descricao & "','" & notaFiscal & "','" & serie & "','" & codigoFornecedor & "','" & dataHora & "','" & USR_IMPORTACAO & "') "

        rsExcel.MoveNext
    Loop

    rsExcel.Close
    Set rsExcel = Nothing
    excelConn.Close
    Set excelConn = Nothing

    ' ----------------------------------------------------------
    ' 3. Executa a transação no SQL Server
    ' Usa ADODB.Connection.Execute diretamente — Conecta() não
    ' serve para DML pois tenta abrir Recordset sobre o resultado,
    ' causando erro 800a0e78 quando a query não retorna linhas.
    ' ----------------------------------------------------------
    Dim SQL, connExec, strConnExec
    SQL = "BEGIN TRANSACTION "
    SQL = SQL & SQL1
    SQL = SQL & " IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

    strConnExec = BuscaStringConexao(UCase(CONN_CORPORATIVO))

    If Trim(strConnExec) = "" Then
        Importar = "{""sucesso"":false,""mensagem"":""String de conexão não encontrada para: " & CONN_CORPORATIVO & """,""dados"":null}"
        Exit Function
    End If

    On Error Resume Next

    Set connExec = Server.CreateObject("ADODB.Connection")
    connExec.CommandTimeout = 300
    connExec.Open strConnExec

    If Err.Number <> 0 Or (connExec Is Nothing) Then
        Importar = "{""sucesso"":false,""mensagem"":""Erro ao conectar ao banco de dados: " & EscapeJSON(Err.Description) & """,""dados"":null}"
        Err.Clear
        Exit Function
    End If

    connExec.Execute SQL, , 128 ' adExecuteNoRecords = 128

    If Err.Number <> 0 Then
        Importar = "{""sucesso"":false,""mensagem"":""Erro ao executar a transação: " & EscapeJSON(Err.Description) & """,""dados"":null}"
        Err.Clear
        connExec.Close
        Set connExec = Nothing
        Exit Function
    End If

    On Error GoTo 0
    connExec.Close
    Set connExec = Nothing

    ' ----------------------------------------------------------
    ' 4. Consulta os registros recém inseridos
    ' ----------------------------------------------------------
    Dim SQL2, rsResultado
    SQL2 = "SELECT [Material],[Centro],[Deposito],[TipoMovimnto],[DocMaterial],[Item],[DataLancamento],[Quantidade],[UM],[Descricao],[NotaFiscal],[Serie],[CodigoFornecedor],[DataHoraIntegracao],[Usuario] "
    SQL2 = SQL2 & "FROM NotasFiscaisInsumos "
    SQL2 = SQL2 & "WHERE DataHoraIntegracao = '" & dataHora & "' AND Usuario = '" & USR_IMPORTACAO & "' "
    SQL2 = SQL2 & "ORDER BY IdNota"

    On Error Resume Next
    Set rsResultado = Conecta(CONN_CORPORATIVO, SQL2)

    If Err.Number <> 0 Or Not IsObject(rsResultado) Then
        Importar = "{""sucesso"":false,""mensagem"":""Erro ao consultar registros importados: " & EscapeJSON(Err.Description) & """,""dados"":null}"
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0

    ' ----------------------------------------------------------
    ' 5. Serializa o resultado em JSON
    ' ----------------------------------------------------------
    If rsResultado.EOF Then
        rsResultado.Close
        Set rsResultado = Nothing
        Importar = "{""sucesso"":false,""mensagem"":""Nenhuma nota fiscal foi importada. Verifique se estes registros já foram importados anteriormente."",""dados"":[]}"
        Exit Function
    End If

    Dim jsonArray, primeiroItem
    jsonArray    = "["
    primeiroItem = True

    Dim fMaterial, fCentro, fDeposito, fTipoMov, fDocMat
    Dim fItem, fDataLanc, fQtd, fUM, fDesc, fNF, fSerie, fCodForn, fDataHoraI, fUsr

    Do Until rsResultado.EOF

        If Not primeiroItem Then jsonArray = jsonArray & "," End If

        fMaterial  = EscapeJSON(rsResultado("Material")           & "")
        fCentro    = EscapeJSON(rsResultado("Centro")             & "")
        fDeposito  = EscapeJSON(rsResultado("Deposito")           & "")
        fTipoMov   = EscapeJSON(rsResultado("TipoMovimnto")       & "")
        fDocMat    = EscapeJSON(rsResultado("DocMaterial")        & "")
        fItem      = EscapeJSON(rsResultado("Item")               & "")
        fDataLanc  = EscapeJSON(rsResultado("DataLancamento")     & "")
        fQtd       = EscapeJSON(rsResultado("Quantidade")         & "")
        fUM        = EscapeJSON(rsResultado("UM")                 & "")
        fDesc      = EscapeJSON(rsResultado("Descricao")          & "")
        fNF        = EscapeJSON(rsResultado("NotaFiscal")         & "")
        fSerie     = EscapeJSON(rsResultado("Serie")              & "")
        fCodForn   = EscapeJSON(rsResultado("CodigoFornecedor")   & "")
        fDataHoraI = EscapeJSON(rsResultado("DataHoraIntegracao") & "")
        fUsr       = EscapeJSON(rsResultado("Usuario")            & "")

        jsonArray = jsonArray & "{"
        jsonArray = jsonArray & """material"":"""           & fMaterial  & ""","
        jsonArray = jsonArray & """centro"":"""             & fCentro    & ""","
        jsonArray = jsonArray & """deposito"":"""           & fDeposito  & ""","
        jsonArray = jsonArray & """tipoMovimento"":"""      & fTipoMov   & ""","
        jsonArray = jsonArray & """docMaterial"":"""        & fDocMat    & ""","
        jsonArray = jsonArray & """item"":"""               & fItem      & ""","
        jsonArray = jsonArray & """dataLancamento"":"""     & fDataLanc  & ""","
        jsonArray = jsonArray & """quantidade"":"""         & fQtd       & ""","
        jsonArray = jsonArray & """um"":"""                 & fUM        & ""","
        jsonArray = jsonArray & """descricao"":"""          & fDesc      & ""","
        jsonArray = jsonArray & """notaFiscal"":"""         & fNF        & ""","
        jsonArray = jsonArray & """serie"":"""              & fSerie     & ""","
        jsonArray = jsonArray & """codigoFornecedor"":"""   & fCodForn   & ""","
        jsonArray = jsonArray & """dataHoraIntegracao"":""" & fDataHoraI & ""","
        jsonArray = jsonArray & """usuario"":"""            & fUsr       & """"
        jsonArray = jsonArray & "}"

        primeiroItem = False
        rsResultado.MoveNext
    Loop

    jsonArray = jsonArray & "]"

    rsResultado.Close
    Set rsResultado = Nothing

    Importar = "{""sucesso"":true,""mensagem"":""Notas fiscais cadastradas com sucesso!"",""dados"":" & jsonArray & "}"

End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Importação de Notas Fiscais SAP</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
        body { background-color: #f0f0f0; min-height: 100vh; }

        /* ── Filter card ── */
        .filter-card         { background:#fff; border:1px solid #dee2e6; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,.08); }
        .filter-card-header  { background:#fff; color:#212529; padding:.6rem 1rem; border-radius:6px 6px 0 0; border-bottom:1px solid #dee2e6; font-weight:600; font-size:.9rem; }
        .filter-card-body    { padding:1rem 1.25rem 1.25rem; }

        /* ── Table card ── */
        .table-card     { background:#fff; border:1px solid #dee2e6; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,.08); position:relative; }
        .table-toolbar  { background:#fff; border-radius:6px 6px 0 0; }
        .table-wrapper  { overflow-x:auto; overflow-y:auto; max-height:68vh; }
        .table-wrapper table { min-width:max-content; }

        .table thead th {
            position:sticky; top:0; background-color:#fff; color:#212529;
            font-size:.82rem; font-weight:700; text-align:center;
            white-space:nowrap; z-index:2; border-color:#dee2e6; padding:.55rem .75rem;
        }
        .table tbody td {
            font-size:.82rem; text-align:center; white-space:nowrap;
            color:#212529; vertical-align:middle; padding:.45rem .75rem;
        }

        /* ── Placeholder ── */
        .table-placeholder      { display:flex; flex-direction:column; align-items:center; justify-content:center; padding:3.5rem 1rem; color:#6c757d; gap:.75rem; }
        .table-placeholder .bi  { font-size:2.8rem; opacity:.35; }
        .table-placeholder p    { margin:0; font-size:.9rem; text-align:center; }

        /* ── Overlay ── */
        .overlay-loading { position:absolute; inset:0; background:rgba(255,255,255,.80); display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:10; gap:1rem; border-radius:6px; }
        .overlay-loading .spinner-border { width:2.5rem; height:2.5rem; color:#6c757d; }
        .overlay-loading span            { font-size:.9rem; font-weight:600; color:#495057; }

        /* ── Toast ── */
        .toast-container-custom { position:fixed; top:1rem; right:1rem; z-index:1100; display:flex; flex-direction:column; gap:.5rem; }
        .toast-custom { min-width:280px; max-width:360px; border-radius:10px; overflow:hidden; box-shadow:0 4px 18px rgba(0,0,0,.13); background:#fff; border:1px solid #dee2e6; animation:toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
        @keyframes toastSlideIn { from { opacity:0; transform:translateX(40px); } to { opacity:1; transform:translateX(0); } }
        .toast-custom .toast-progress-bar { animation:toastProgressShrink 4.5s linear forwards; }
        @keyframes toastProgressShrink { from { width:100%; } to { width:0%; } }

        .page-title { font-size:1.15rem; font-weight:700; color:#212529; }
    </style>
</head>
<body>
<div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center gap-3 mb-3">
        <button type="button" class="btn btn-light btn-sm" onclick="history.back()">
            <i class="bi bi-arrow-left me-1"></i>Voltar
        </button>
            <span class="page-title">Importação de Notas Fiscais SAP</span>
    </div>

    <!-- Card de ação -->
    <div class="filter-card mb-3">
        <div class="filter-card-header">
            <i class="bi bi-file-earmark-arrow-up me-2"></i>Importação
        </div>
        <div class="filter-card-body">
            <p class="mb-3 text-muted" style="font-size:.875rem;">
                Clique em <strong>Importar</strong> para processar o arquivo
                <code>Notas.xlsx</code> e inserir os registros novos na base de dados.
            </p>
            <div class="d-flex gap-2">
                <button type="button" id="btnImportar" class="btn btn-primary btn-sm">
                    <i class="bi bi-cloud-upload me-2"></i>Importar
                </button>
            </div>
        </div>
    </div>

    <!-- Card de resultados -->
    <div class="table-card">

        <div id="overlayLoading" class="overlay-loading d-none">
            <div class="spinner-border" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
            <span>Processando importação...</span>
        </div>

        <!-- Toolbar -->
        <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
            <span style="font-size:.85rem;font-weight:600;color:#495057;">
                <i class="bi bi-table me-2"></i>Registros Importados
            </span>
            <div class="d-flex gap-2">
                <button type="button" id="btnExportCsv" class="btn btn-success btn-sm d-none">
                    <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
                </button>
            </div>
        </div>

        <!-- Área de conteúdo -->
        <div id="areaConteudo">
            <div class="table-placeholder">
                <i class="bi bi-clipboard-data"></i>
                <p>Clique em <strong>Importar</strong> para processar o arquivo e visualizar os registros inseridos.</p>
            </div>
        </div>

    </div><!-- /.table-card -->

</div><!-- /.container-fluid -->

<div class="toast-container-custom" id="toastContainer"></div>

<!-- Template da tabela de resultados -->
<template id="tplTabela">
    <div class="table-wrapper">
        <table class="table table-bordered table-striped table-hover table-sm mb-0">
            <thead>
                <tr>
                    <th>#</th>
                    <th>Material</th>
                    <th>Centro</th>
                    <th>Depósito</th>
                    <th>Tipo Movimento</th>
                    <th>Doc. Material</th>
                    <th>Item</th>
                    <th>Data Lançamento</th>
                    <th>Quantidade</th>
                    <th>UM</th>
                    <th>Descrição</th>
                    <th>Nota Fiscal</th>
                    <th>Série</th>
                    <th>Cód. Fornecedor</th>
                    <th>Data/Hora Integração</th>
                    <th>Usuário</th>
                </tr>
            </thead>
            <tbody id="corpoTabela"></tbody>
        </table>
    </div>
</template>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
"use strict";

const ImportacaoNotasFiscais = (function () {

    const REPOSITORIO = 'ImportacaoNotasFiscais.asp';
    const TOAST_DELAY = 4500;

    const TOAST_ICON = Object.freeze({
        sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
        erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
        aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
        info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' }
    });

    /* ── Toast ── */
    function exibirToast(tipo, titulo, corpo) {
        const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
        const el  = document.createElement('div');
        el.className = 'toast-custom';
        el.innerHTML = `
            <div style="height:4px;background:${cfg.barra};" class="toast-progress-bar"></div>
            <div class="p-3">
                <div class="d-flex align-items-center justify-content-between mb-1">
                    <span style="font-weight:700;font-size:.88rem;color:${cfg.cor};">
                        <i class="bi ${cfg.icon} me-1"></i>${titulo}
                    </span>
                    <button style="background:none;border:none;cursor:pointer;font-size:1rem;color:#6c757d;padding:0 0 0 .5rem;"
                            onclick="this.closest('.toast-custom').remove()">&times;</button>
                </div>
                ${corpo ? `<div style="font-size:.82rem;color:#495057;">${corpo}</div>` : ''}
            </div>`;
        document.getElementById('toastContainer').appendChild(el);
        setTimeout(() => el.remove(), TOAST_DELAY);
    }

    /* ── Placeholder ── */
    function renderPlaceholder(icone, html) {
        document.getElementById('areaConteudo').innerHTML =
            `<div class="table-placeholder">
                <i class="bi ${icone}"></i>
                <p>${html}</p>
            </div>`;
    }

    /* ── Tabela ── */
    function renderTabela(dados) {
        const node  = document.getElementById('tplTabela').content.cloneNode(true);
        const tbody = node.getElementById('corpoTabela');
        const frag  = document.createDocumentFragment();

        dados.forEach((r, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td>${r.material            || ''}</td>
                <td>${r.centro              || ''}</td>
                <td>${r.deposito            || ''}</td>
                <td>${r.tipoMovimento       || ''}</td>
                <td>${r.docMaterial         || ''}</td>
                <td>${r.item               || ''}</td>
                <td>${r.dataLancamento      || ''}</td>
                <td>${r.quantidade          || ''}</td>
                <td>${r.um                  || ''}</td>
                <td>${r.descricao           || ''}</td>
                <td>${r.notaFiscal          || ''}</td>
                <td>${r.serie              || ''}</td>
                <td>${r.codigoFornecedor    || ''}</td>
                <td>${r.dataHoraIntegracao  || ''}</td>
                <td>${r.usuario             || ''}</td>`;
            frag.appendChild(tr);
        });

        tbody.appendChild(frag);

        const area = document.getElementById('areaConteudo');
        area.innerHTML = '';
        area.appendChild(node);
    }

    /* ── Exportação CSV ── */
    function exportarCSV(nomeArquivo, cabecalhos, linhas) {
        const esc = v => {
            const s = v == null ? '' : String(v);
            return (s.includes(';') || s.includes('"') || s.includes('\n'))
                ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const rows = [
            cabecalhos.map(esc).join(';'),
            ...linhas.map(row => row.map(esc).join(';'))
        ];
        const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href     = url;
        a.download = nomeArquivo.endsWith('.csv') ? nomeArquivo : nomeArquivo + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /* ── Estado em memória ── */
    let dadosUltimaImportacao = [];

    /* ── AJAX ── */
    function executarImportacao() {
        $.ajax({
            url      : REPOSITORIO,
            method   : 'POST',
            dataType : 'json',
            data     : { acao: 'importar' },
            beforeSend: function () {
                $('#overlayLoading').removeClass('d-none');
                $('#btnImportar').prop('disabled', true);
                $('#btnExportCsv').addClass('d-none');
                dadosUltimaImportacao = [];
            },
            success: function (resp) {
                if (!resp.sucesso) {
                    const ehVazio = Array.isArray(resp.dados) && resp.dados.length === 0;
                    if (ehVazio) {
                        renderPlaceholder('bi-inbox', 'Nenhuma nota fiscal foi importada. Os registros do arquivo podem já ter sido inseridos anteriormente.');
                        exibirToast('aviso', 'Sem novos registros', resp.mensagem);
                    } else {
                        renderPlaceholder('bi-exclamation-triangle text-danger" style="opacity:1', '<span class="text-danger"><strong>Erro:</strong> ' + resp.mensagem + '</span>');
                        exibirToast('erro', 'Erro na importação', resp.mensagem);
                    }
                    return;
                }

                dadosUltimaImportacao = resp.dados || [];
                renderTabela(dadosUltimaImportacao);
                $('#btnExportCsv').removeClass('d-none');
                exibirToast('sucesso', 'Importação concluída', resp.mensagem);
            },
            error: function (xhr) {
                renderPlaceholder('bi-wifi-off', 'Erro de comunicação com o servidor. Tente novamente.');
                exibirToast('erro', 'Erro de comunicação', 'Não foi possível concluir a importação.');
                console.error(xhr.responseText);
            },
            complete: function () {
                $('#overlayLoading').addClass('d-none');
                $('#btnImportar').prop('disabled', false);
            }
        });
    }

    /* ── Init ── */
    function init() {
        $('#btnImportar').on('click', function () {
            executarImportacao();
        });

        $('#btnExportCsv').on('click', function () {
            if (!dadosUltimaImportacao.length) return;

            const cabecalhos = [
                'Material', 'Centro', 'Depósito', 'Tipo Movimento', 'Doc. Material',
                'Item', 'Data Lançamento', 'Quantidade', 'UM', 'Descrição',
                'Nota Fiscal', 'Série', 'Cód. Fornecedor', 'Data/Hora Integração', 'Usuário'
            ];
            const linhas = dadosUltimaImportacao.map(r => [
                r.material, r.centro, r.deposito, r.tipoMovimento, r.docMaterial,
                r.item, r.dataLancamento, r.quantidade, r.um, r.descricao,
                r.notaFiscal, r.serie, r.codigoFornecedor, r.dataHoraIntegracao, r.usuario
            ]);

            exportarCSV('NotasFiscaisImportadas', cabecalhos, linhas);
        });
    }

    return Object.freeze({ init });

}());

$(function () {
    ImportacaoNotasFiscais.init();
});
</script>
</body>
</html>
