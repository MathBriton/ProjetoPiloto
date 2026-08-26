<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../corporativo/cnn2/cnn2.asp"-->
<!--#include file="../../corporativo/funcoes/funcoes.asp"-->
<!--#include file="../../corporativo/includes/variaveis.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Session.LCID         = 1046

' ════════════════════════════════════════════════════════════════════════════
'  DISPATCHER JSON  —  executado apenas quando ?acao= está presente
' ════════════════════════════════════════════════════════════════════════════
Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.Expires     = 0
    Response.ContentType = "application/json"

    Set usuarioLogado = CreateObject("Scripting.Dictionary")
    usuarioLogado.Add "id",    Session("CODIGO")
    usuarioLogado.Add "login", Session("USR")
    usuarioLogado.Add "nome",  Session("NOME")

    Dim resultado
    Select Case acao
        Case "listar"  : resultado = Listar()
        Case Else
            resultado = "{""sucesso"":false,""mensagem"":""Ação inválida: " & EscapeJSON(acao) & """}"
    End Select

    Response.Write resultado
    Response.End()

End If

' ════════════════════════════════════════════════════════════════════════════
'  FUNÇÕES DE DOMÍNIO
' ════════════════════════════════════════════════════════════════════════════

' ---------------------------------------------------------------------------
'  Listar — processa a query IQA agrupada por fornecedor e devolve JSON
' ---------------------------------------------------------------------------
Function Listar()
    Dim sFilial, sDataIni, sDataFim, sFornecedor
    sFilial     = Trim(Request("Filial"))
    sDataIni    = Trim(Request("DataIni"))
    sDataFim    = Trim(Request("DataFim"))
    sFornecedor = Trim(Request("lbFornecedor"))

    ' Validações de entrada
    If sFilial = "" Then
        Listar = "{""sucesso"":false,""mensagem"":""Selecione uma filial para realizar a pesquisa.""}"
        Exit Function
    End If

    If sDataIni = "" And sDataFim = "" And sFornecedor = "" Then
        Listar = "{""sucesso"":false,""mensagem"":""Informe ao menos um filtro de data ou fornecedor.""}"
        Exit Function
    End If

    ' Mapeamento filial → conexão
    Dim sConn
    sConn = ""
    Select Case sFilial
        Case "B" : sConn = "BEN-CORPORATIVO"
        Case "F" : sConn = "FOR-CORPORATIVO"
        Case "P" : sConn = "PIR-CORPORATIVO"
        Case "U" : sConn = "PIR-CORPORATIVO"
        Case "S" : sConn = "SAQ-CORPORATIVO"
    End Select

    If sConn = "" Then
        Listar = "{""sucesso"":false,""mensagem"":""Filial inválida.""}"
        Exit Function
    End If

    ' Montagem dos filtros SQL
    Dim sFiltroData, sFiltroForn, SQL
    sFiltroData = ""
    sFiltroForn = ""

    If sDataIni <> "" And sDataFim <> "" Then
        sFiltroData = " AND CO_DATCONF BETWEEN '" & formataDataYYYYMMDD(sDataIni) & "' AND '" & formataDataYYYYMMDD(sDataFim) & "'"
    End If
    If sFornecedor <> "" Then
        sFiltroForn = " AND CO_FORNECE = '" & Replace(sFornecedor, "'", "''") & "'"
    End If

    SQL = "SELECT CO_FORNECE, CO_DESCRI, CO_DATCONF, A2_NOME, A2_IQS, 100 AS PONTOS, ENCONTRADO "
    SQL = SQL & "FROM vw_MSNotasAnalisadas_IQA "
    SQL = SQL & "WHERE 1=1" & sFiltroData & sFiltroForn
    SQL = SQL & " ORDER BY CO_FORNECE"

    Dim rsMain
    On Error Resume Next
    Set rsMain = Conecta(sConn, SQL)
    If Err.Number <> 0 Or Not IsObject(rsMain) Then
        Listar = "{""sucesso"":false,""mensagem"":""Erro ao conectar: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    If rsMain.EOF Then
        rsMain.Close()
        Set rsMain = Nothing
        Listar = "{""sucesso"":false,""mensagem"":""Nenhum registro encontrado para os filtros selecionados.""}"
        Exit Function
    End If

    ' ── Processamento agrupado por fornecedor ─────────────────────────────
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")

    Dim iqa_fornecedor, iqa_nfornecedor, iqa_descricao, iqa_descri
    Dim iqa_contaFornece, iqa_somaIQS, iqa_cont
    Dim iqa_iqs, iqa_iqf, iqa_iqa
    Dim iqa_Tiqs, iqa_Tiqf, iqa_Tiqa
    Dim iqa_rsSub, iqa_SQLsub, iqa_DataIniM, iqa_DataFimM
    Dim iqa_num, iqa_totIQF

    iqa_fornecedor   = rsMain("CO_FORNECE") & ""
    iqa_nfornecedor  = rsMain("A2_NOME") & ""
    iqa_contaFornece = 0
    iqa_somaIQS      = 0
    iqa_cont         = 0
    iqa_descri       = ""
    iqa_Tiqs = 0 : iqa_Tiqf = 0 : iqa_Tiqa = 0

    Dim jsonArray, primeiroItem
    jsonArray   = "["
    primeiroItem = True

    ' Sub-função local: calcula IQF para um fornecedor e devolve o valor
    ' (VBScript não tem lambdas; extraímos inline com GoSub simulado via Sub)

    Do While Not rsMain.EOF

        If iqa_fornecedor <> (rsMain("CO_FORNECE") & "") Then

            ' ── Fechar grupo do fornecedor anterior ──────────────────────
            iqa_DataIniM = Replace(Left(sDataIni, 7), "-", "")
            iqa_DataFimM = Replace(Left(sDataFim, 7), "-", "")

            iqa_num    = 0
            iqa_totIQF = 0

            iqa_SQLsub = "SELECT MES, MIN(NOTA) IQF, MAX(QTD) QTD FROM ("
            iqa_SQLsub = iqa_SQLsub & "(SELECT DISTINCT SUBSTRING(CO_DATCONF,1,6) MES, 100 NOTA, COUNT(CO_DOC) QTD "
            iqa_SQLsub = iqa_SQLsub & " FROM vw_MSNotasAnalisadas_Fornecedores WHERE CO_FORNECE = SUBSTRING('" & Replace(iqa_fornecedor,"'","''") & "',1,6) GROUP BY SUBSTRING(CO_DATCONF,1,6)) "
            iqa_SQLsub = iqa_SQLsub & "UNION "
            iqa_SQLsub = iqa_SQLsub & "(SELECT MES, NOTA, 0 QTD FROM vw_MSIQF WHERE FORNECEDOR = SUBSTRING('" & Replace(iqa_fornecedor,"'","''") & "',1,6))"
            iqa_SQLsub = iqa_SQLsub & ") X"
            If sDataIni <> "" And sDataFim <> "" Then
                iqa_SQLsub = iqa_SQLsub & " WHERE MES BETWEEN SUBSTRING('" & iqa_DataIniM & "',1,6) AND SUBSTRING('" & iqa_DataFimM & "',1,6)"
            End If
            iqa_SQLsub = iqa_SQLsub & " GROUP BY MES ORDER BY MES"

            On Error Resume Next
            Set iqa_rsSub = Conecta(sConn, iqa_SQLsub)
            On Error GoTo 0

            If Not (iqa_rsSub Is Nothing) Then
                Do While Not iqa_rsSub.EOF
                    iqa_num    = iqa_num + iqa_rsSub("QTD")
                    iqa_totIQF = iqa_totIQF + (iqa_rsSub("IQF") * iqa_rsSub("QTD"))
                    iqa_rsSub.MoveNext()
                Loop
                iqa_rsSub.Close()
                Set iqa_rsSub = Nothing
            End If

            If iqa_num > 0 Then iqa_totIQF = iqa_totIQF / iqa_num

            iqa_iqf = iqa_totIQF
            If iqa_cont > 0 Then
                iqa_iqs = CDbl(FormatNumber(iqa_somaIQS / iqa_cont, 0))
            Else
                iqa_iqs = 0
            End If
            iqa_iqa = (iqa_iqs * 0.3 * (100.0/26.0)) + (iqa_iqf * 0.7)

            iqa_Tiqs = iqa_Tiqs + iqa_iqs
            iqa_Tiqf = iqa_Tiqf + iqa_iqf
            iqa_Tiqa = iqa_Tiqa + iqa_iqa

            ' Serializar item do array JSON
            If Not primeiroItem Then jsonArray = jsonArray & ","
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """fornecedor"":"""   & EscapeJSON(iqa_fornecedor)  & ""","
            jsonArray = jsonArray & """nome"":"""         & EscapeJSON(iqa_nfornecedor) & ""","
            jsonArray = jsonArray & """materiais"":"""    & EscapeJSON(Mid(iqa_descri, 3, Len(iqa_descri)-3)) & ""","
            jsonArray = jsonArray & """iqs"":"            & Replace(CStr(CDbl(FormatNumber(iqa_iqs, 2))), ",", ".") & ","
            jsonArray = jsonArray & """iqf"":"            & Replace(CStr(CDbl(FormatNumber(iqa_iqf, 2))), ",", ".") & ","
            jsonArray = jsonArray & """iqa"":"            & Replace(CStr(CDbl(FormatNumber(iqa_iqa, 2))), ",", ".")
            jsonArray = jsonArray & "}"
            primeiroItem = False

            iqa_contaFornece = iqa_contaFornece + 1
            iqa_somaIQS = 0
            iqa_cont    = 0
            d.RemoveAll
            iqa_descri  = ""

        End If ' troca de fornecedor

        iqa_somaIQS     = iqa_somaIQS + CDbl("" & rsMain("A2_IQS"))
        iqa_fornecedor  = rsMain("CO_FORNECE") & ""
        iqa_nfornecedor = rsMain("A2_NOME") & ""
        iqa_descricao   = rsMain("CO_DESCRI") & ""

        If d.Exists(iqa_descricao) = False Then
            d.Add iqa_descricao, "s"
            iqa_descri = iqa_descri & " | " & iqa_descricao
        End If

        iqa_cont = iqa_cont + 1
        rsMain.MoveNext()
    Loop

    ' ── Último fornecedor ─────────────────────────────────────────────────
    If iqa_cont > 0 Then
        iqa_iqs = CDbl(FormatNumber(iqa_somaIQS / iqa_cont, 0))
    Else
        iqa_iqs = 0
    End If
    ' Último fornecedor: sem sub-query IQF, mantém cálculo original (iqa_iqs nas duas parcelas)
    iqa_iqa = (iqa_iqs * 0.3 * (100.0/26.0)) + (iqa_iqs * 0.7)

    iqa_Tiqs = iqa_Tiqs + iqa_iqs
    iqa_Tiqa = iqa_Tiqa + iqa_iqa

    If Not primeiroItem Then jsonArray = jsonArray & ","
    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """fornecedor"":"""   & EscapeJSON(iqa_fornecedor)  & ""","
    jsonArray = jsonArray & """nome"":"""         & EscapeJSON(iqa_nfornecedor) & ""","
    jsonArray = jsonArray & """materiais"":"""    & EscapeJSON(Mid(iqa_descri, 3, Len(iqa_descri)-3)) & ""","
    jsonArray = jsonArray & """iqs"":"            & Replace(CStr(CDbl(FormatNumber(iqa_iqs, 2))), ",", ".") & ","
    jsonArray = jsonArray & """iqf"":0,"
    jsonArray = jsonArray & """iqa"":"            & Replace(CStr(CDbl(FormatNumber(iqa_iqa, 2))), ",", ".")
    jsonArray = jsonArray & "}"

    iqa_contaFornece = iqa_contaFornece + 1

    rsMain.Close()
    Set rsMain = Nothing
    Set d      = Nothing

    jsonArray = jsonArray & "]"

    ' Totais gerais
    Dim divisor
    divisor = IIf(iqa_contaFornece > 0, iqa_contaFornece, 1)

    Dim totalIqs, totalIqf, totalIqa
    totalIqs = Replace(CStr(CDbl(FormatNumber(iqa_Tiqs / divisor, 2))), ",", ".")
    totalIqf = Replace(CStr(CDbl(FormatNumber(iqa_Tiqf / divisor, 2))), ",", ".")
    totalIqa = Replace(CStr(CDbl(FormatNumber(iqa_Tiqa / divisor, 2))), ",", ".")

    Listar = "{"
    Listar = Listar & """sucesso"":true,"
    Listar = Listar & """mensagem"":"""","
    Listar = Listar & """dados"":{"
    Listar = Listar & """itens"":"   & jsonArray & ","
    Listar = Listar & """totalFornecedores"":" & iqa_contaFornece & ","
    Listar = Listar & """totalIqs"":"  & totalIqs & ","
    Listar = Listar & """totalIqf"":"  & totalIqf & ","
    Listar = Listar & """totalIqa"":"  & totalIqa
    Listar = Listar & "}}"

End Function

' ════════════════════════════════════════════════════════════════════════════
'  HTML  —  renderizado quando não há ?acao= (acesso direto ao formulário)
' ════════════════════════════════════════════════════════════════════════════
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet — IQA de Fornecedores</title>

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

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    .filter-card        { background: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .filter-card-header { background: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
    .filter-card-body   { padding: 1rem 1.25rem 1.25rem; }
    .filter-card .form-label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    .table-card { background: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); position: relative; }

    .table-toolbar { background: #fff; border-radius: 6px 6px 0 0; }

    .search-box .form-control { padding-right: 80px; }
    .search-hotkey { position: absolute; right: .5rem; top: 50%; transform: translateY(-50%); display: flex; gap: 2px; pointer-events: none; }
    .search-hotkey kbd { font-size: .65rem; background: #e9ecef; color: #495057; border: 1px solid #ced4da; border-radius: 3px; padding: 1px 4px; }

    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
    .table-wrapper table { min-width: max-content; }
    .table thead th {
      position: sticky; top: 0;
      background: #fff; color: #212529;
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
    .table tbody td:first-child,
    .table thead th:first-child,
    .table tfoot th:first-child {
      white-space: normal; word-break: break-word;
      max-width: 420px; min-width: 200px; text-align: left;
    }

    .col-sticky-iqs, .col-sticky-iqf, .col-sticky-iqa {
      position: sticky; background: #fff; z-index: 3; text-align: center;
    }
    .col-sticky-iqa { right: 0;    border-left: 2px solid #dee2e6; }
    .col-sticky-iqf { right: 80px; }
    .col-sticky-iqs { right: 180px; }
    .table thead th.col-sticky-iqs,
    .table thead th.col-sticky-iqf,
    .table thead th.col-sticky-iqa { z-index: 4; }
    .table-striped tbody tr:nth-of-type(odd) td.col-sticky-iqs,
    .table-striped tbody tr:nth-of-type(odd) td.col-sticky-iqf,
    .table-striped tbody tr:nth-of-type(odd) td.col-sticky-iqa { background: #f2f2f2; }
    .table-hover tbody tr:hover td.col-sticky-iqs,
    .table-hover tbody tr:hover td.col-sticky-iqf,
    .table-hover tbody tr:hover td.col-sticky-iqa { background: #e9ecef; }
    .table tfoot th.col-sticky-iqs,
    .table tfoot th.col-sticky-iqf,
    .table tfoot th.col-sticky-iqa { background: #212529; }

    .table-footer { display: flex; align-items: center; justify-content: space-between; padding: .6rem 1rem; border-top: 1px solid #dee2e6; flex-wrap: wrap; gap: .5rem; }
    .table-footer-right { min-width: 140px; text-align: right; }
    .info-registros { font-size: .8rem; color: #6c757d; }

    .table-placeholder { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 3.5rem 1rem; color: #6c757d; gap: .75rem; }
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

    .overlay-loading { position: absolute; inset: 0; background: rgba(255,255,255,.80); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 10; gap: 1rem; border-radius: 6px; }
    .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
    .overlay-loading span { font-size: .9rem; font-weight: 600; color: #495057; }

    .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
    .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
    @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
    .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

    .legenda-card { background: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .legenda-badge-a { background-color: #198754; color: #fff; }
    .legenda-badge-b { background-color: #0d6efd; color: #fff; }
    .legenda-badge-c { background-color: #ffc107; color: #212529; }
  </style>
</head>
<body>

<div class="container-fluid py-3 px-3 px-lg-4">

  <div class="toast-container-custom" id="toastContainer"></div>

  <!-- Cabeçalho -->
  <div class="d-flex align-items-center justify-content-between mb-3">
    <span class="page-title">Índice de Qualificação e Avaliação de Desempenho de Fornecedores</span>
    <div class="d-flex gap-2">
      <button type="button" class="btn btn-outline-primary btn-sm"
              onclick="location.href='../../corporativo/suprimentos/ImportacaoFornecedores.asp'">
        <i class="bi bi-people me-2"></i>Fornecedores
      </button>
      <button type="button" class="btn btn-outline-primary btn-sm"
              onclick="location.href='../../corporativo/suprimentos/ImportacaoNotasFiscais.asp'">
        <i class="bi bi-upc-scan me-2"></i>Notas
      </button>
    </div>
  </div>

  <!-- Card de Filtros -->
  <div class="filter-card mb-3">
    <div class="filter-card-header">
      <i class="bi bi-funnel me-2"></i>Filtros
    </div>
    <div class="filter-card-body">
      <div class="row g-3 align-items-end">

        <div class="col-12 col-sm-6 col-md-2">
          <label for="Filial" class="form-label">Filial</label>
          <select id="Filial" class="form-select form-select-sm">
            <option value="">Selecione</option>
            <option value="B">Bento Gonçalves</option>
            <option value="F">Fortaleza</option>
            <option value="P">Pirapetinga</option>
            <option value="U">Uberaba</option>
            <option value="S">Saquarema</option>
          </select>
        </div>

        <div class="col-12 col-sm-6 col-md-2">
          <label for="DataIni" class="form-label">Data Inicial</label>
          <input type="date" id="DataIni" class="form-control form-control-sm" />
        </div>

        <div class="col-12 col-sm-6 col-md-2">
          <label for="DataFim" class="form-label">Data Final</label>
          <input type="date" id="DataFim" class="form-control form-control-sm" />
        </div>

        <div class="col-12 col-sm-6 col-md-3">
          <label for="lbFornecedor" class="form-label">Fornecedor (código)</label>
          <input type="text" id="lbFornecedor" class="form-control form-control-sm" autocomplete="off" />
        </div>

        <div class="col-12 col-md-3">
          <div class="d-flex gap-2 justify-content-end">
            <button type="button" id="btnLimpar" class="btn btn-light btn-sm">
              <i class="bi bi-eraser me-2"></i>Limpar
            </button>
            <button type="button" id="btnConsultar" class="btn btn-primary btn-sm">
              <i class="bi bi-funnel-fill me-2"></i>Filtrar
            </button>
          </div>
        </div>

      </div>
    </div>
  </div>

  <!-- Legenda IQA -->
  <div class="legenda-card p-3 mb-3">
    <p class="fw-semibold mb-2" style="font-size:.85rem;">
      <i class="bi bi-info-circle me-2 text-secondary"></i>Legenda de Classificação IQA
    </p>
    <div class="row g-2 mb-2" style="font-size:.82rem;">
      <div class="col">
        <div class="d-flex align-items-center gap-2 h-100 p-2 rounded justify-content-center" style="background:#f8f9fa;">
          <span class="badge legenda-badge-a" style="min-width:90px;">85% — 100%</span>
          <span>Qualificado — Nível <strong>A</strong> (qualidade assegurada)</span>
        </div>
      </div>
      <div class="col">
        <div class="d-flex align-items-center gap-2 h-100 p-2 rounded justify-content-center" style="background:#f8f9fa;">
          <span class="badge legenda-badge-b" style="min-width:90px;">71% — 84%</span>
          <span>Qualificado — Nível <strong>B</strong> (intermediário)</span>
        </div>
      </div>
      <div class="col">
        <div class="d-flex align-items-center gap-2 h-100 p-2 rounded justify-content-center" style="background:#f8f9fa;">
          <span class="badge legenda-badge-c" style="min-width:90px;">50% — 70%</span>
          <span>Qualificado — Nível <strong>C</strong> (restritivo)</span>
        </div>
      </div>
    </div>
    <div class="row g-0" style="font-size:.75rem; color:#6c757d;">
      <div class="col text-center">IQA = Índice de Qualidade Assegurada</div>
      <div class="col text-center">IQS = Índice de Qualidade do Sistema</div>
      <div class="col text-center">IQF = Índice de Qualidade de Fornecimento</div>
    </div>
  </div>

  <!-- Card de Tabela -->
  <div class="table-card mb-3">

    <div id="overlayLoading" class="overlay-loading d-none">
      <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
      <span>Carregando...</span>
    </div>

    <!-- Toolbar -->
    <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
      <div class="d-flex gap-2">
        <button id="btnExportCsv" type="button" class="btn btn-success btn-sm">
          <i class="bi bi-file-earmark-excel me-2"></i>Exportar Excel
        </button>
        <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm">
          <i class="bi bi-file-earmark-pdf me-2"></i>Exportar PDF
        </button>
      </div>
      <div class="search-box position-relative">
        <input type="text" id="searchBox" class="form-control form-control-sm"
               placeholder="Buscar..." autocomplete="off" />
        <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
      </div>
    </div>

    <!-- Placeholder inicial -->
    <div id="tablePlaceholder" class="table-placeholder">
      <i class="bi bi-clipboard-data"></i>
      <p>Selecione os filtros e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
    </div>

    <!-- Tabela (oculta até carregar) -->
    <div id="tableWrapper" class="table-wrapper d-none">
      <table id="tblIQA" class="table table-bordered table-striped table-hover table-sm mb-0">
        <thead>
          <tr>
            <th style="text-align:left;">Fornecedor / Materiais</th>
            <th class="col-sticky-iqs" style="min-width:80px;width:80px;">IQS</th>
            <th class="col-sticky-iqf" style="min-width:100px;width:100px;">IQF (Médio)</th>
            <th class="col-sticky-iqa" style="min-width:80px;width:80px;">IQA</th>
          </tr>
        </thead>
        <tbody id="tbodyIQA"></tbody>
        <tfoot>
          <tr class="table-dark">
            <th class="text-start">Índice Geral do Período</th>
            <th id="tfIqs" class="col-sticky-iqs text-end"></th>
            <th id="tfIqf" class="col-sticky-iqf text-end"></th>
            <th id="tfIqa" class="col-sticky-iqa text-end"></th>
          </tr>
        </tfoot>
      </table>
    </div>

    <!-- Footer -->
    <div id="tableFooter" class="table-footer d-none">
      <div class="table-footer-left"></div>
      <div class="table-footer-center"></div>
      <div class="table-footer-right">
        <span class="info-registros">
          <i class="bi bi-building me-2"></i>
          <strong id="totalFornecedores">0</strong> fornecedor(es)
        </span>
      </div>
    </div>

  </div>

</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
"use strict";

(function IQA() {

  // ── Constantes ─────────────────────────────────────────────────────────
  const REPOSITORIO = 'iqa.asp';
  const LOGO_PATH   = '../../corporativo/suprimentos/Logo/logo-smurfit1.png';
  const TOAST_DELAY = 4500;

  const TOAST_ICON = Object.freeze({
    sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754' },
    erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545' },
    aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404' },
    info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d' },
  });

  // ── Estado ─────────────────────────────────────────────────────────────
  let dadosCompletos = [];

  // ── Helpers ────────────────────────────────────────────────────────────
  function $el(id) { return document.getElementById(id); }

  function fmt(num) {
    return Number(num).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function exibirToast(tipo, mensagem) {
    const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
    const el  = document.createElement('div');
    el.className = 'toast-custom';
    el.innerHTML = `
      <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
        <i class="bi ${cfg.icon}" style="color:${cfg.cor};font-size:1.1rem;"></i>
        <span class="fw-semibold flex-grow-1" style="color:${cfg.cor};font-size:.88rem;">${mensagem}</span>
        <button class="btn-close btn-close-sm" onclick="this.closest('.toast-custom').remove()"></button>
      </div>
      <div class="px-3 pb-2" style="font-size:.82rem;color:#495057;">${mensagem}</div>
      <div style="height:3px;background:#e9ecef;border-radius:0 0 10px 10px;">
        <div class="toast-progress-bar" style="height:100%;background:${cfg.cor};border-radius:inherit;"></div>
      </div>`;
    $el('toastContainer').appendChild(el);
    setTimeout(() => el.remove(), TOAST_DELAY + 300);
  }

  // ── Renderização da tabela ──────────────────────────────────────────────
  function renderTabela(itens) {
    const tbody = $el('tbodyIQA');
    const frag  = document.createDocumentFragment();

    itens.forEach(function (item) {
      // Linha de cabeçalho do fornecedor
      const trGrupo = document.createElement('tr');
      trGrupo.className = 'table-active';
      const tdGrupo = document.createElement('td');
      tdGrupo.colSpan = 4;
      tdGrupo.className = 'fw-bold text-start';
      tdGrupo.innerHTML = `<span class="text-muted me-2" style="font-size:.75rem;">${item.fornecedor}</span>${item.nome}`;
      trGrupo.appendChild(tdGrupo);
      frag.appendChild(trGrupo);

      // Linha de dados
      const trDados = document.createElement('tr');
      trDados.innerHTML = `
        <td class="text-start text-muted" style="font-size:.75rem;">${item.materiais}</td>
        <td class="col-sticky-iqs text-end fw-semibold">${fmt(item.iqs)}%</td>
        <td class="col-sticky-iqf text-end fw-semibold">${fmt(item.iqf)}%</td>
        <td class="col-sticky-iqa text-end fw-bold text-primary">${fmt(item.iqa)}%</td>`;
      frag.appendChild(trDados);
    });

    tbody.innerHTML = '';
    tbody.appendChild(frag);
  }

  function mostrarTabela(dados) {
    renderTabela(dados.itens);

    $el('tfIqs').textContent = fmt(dados.totalIqs) + '%';
    $el('tfIqf').textContent = fmt(dados.totalIqf) + '%';
    $el('tfIqa').textContent = fmt(dados.totalIqa) + '%';
    $el('totalFornecedores').textContent = dados.totalFornecedores;

    $el('tablePlaceholder').classList.add('d-none');
    $el('tableWrapper').classList.remove('d-none');
    $el('tableFooter').classList.remove('d-none');
  }

  function ocultarTabela() {
    $el('tableWrapper').classList.add('d-none');
    $el('tableFooter').classList.add('d-none');
    $el('tablePlaceholder').classList.remove('d-none');
    dadosCompletos = [];
  }

  // ── Consulta AJAX ───────────────────────────────────────────────────────
  function consultar() {
    const filial     = $el('Filial').value;
    const dataIni    = $el('DataIni').value;
    const dataFim    = $el('DataFim').value;
    const fornecedor = $el('lbFornecedor').value.trim();

    if (!filial) {
      exibirToast('aviso', 'Selecione uma filial.');
      return;
    }
    if (!dataIni && !dataFim && !fornecedor) {
      exibirToast('aviso', 'Informe ao menos um filtro de data ou fornecedor.');
      return;
    }

    $el('overlayLoading').classList.remove('d-none');
    ocultarTabela();

    $.ajax({
      url     : REPOSITORIO,
      method  : 'POST',
      data    : { acao: 'listar', Filial: filial, DataIni: dataIni, DataFim: dataFim, lbFornecedor: fornecedor },
      dataType: 'json',
      success : function (resp) {
        $el('overlayLoading').classList.add('d-none');
        if (!resp.sucesso) {
          exibirToast('aviso', resp.mensagem);
          return;
        }
        dadosCompletos = resp.dados.itens;
        mostrarTabela(resp.dados);
      },
      error: function () {
        $el('overlayLoading').classList.add('d-none');
        exibirToast('erro', 'Erro de comunicação com o servidor. Tente novamente.');
      }
    });
  }

  // ── Busca em memória ────────────────────────────────────────────────────
  function initSearch() {
    const box = $el('searchBox');
    if (!box) return;

    box.addEventListener('input', function () {
      const q = this.value.toLowerCase();
      $el('tbodyIQA').querySelectorAll('tr').forEach(function (tr) {
        tr.style.display = (!q || tr.textContent.toLowerCase().includes(q)) ? '' : 'none';
      });
    });

    document.addEventListener('keydown', function (e) {
      if (e.ctrlKey && e.key === 'k') { e.preventDefault(); box.focus(); }
    });
  }

  // ── Exportar CSV ────────────────────────────────────────────────────────
  function onExportCsv() {
    if (!dadosCompletos.length) { exibirToast('aviso', 'Nenhum dado disponível para exportar.'); return; }

    const esc = function (v) {
      const s = v == null ? '' : String(v);
      return (s.includes(';') || s.includes('"') || s.includes('\n')) ? `"${s.replace(/"/g, '""')}"` : s;
    };

    const cab  = ['Fornecedor', 'Nome', 'Materiais', 'IQS', 'IQF (Médio)', 'IQA'];
    const rows = [cab.map(esc).join(';')];

    dadosCompletos.forEach(function (item) {
      rows.push([item.fornecedor, item.nome, item.materiais,
                 fmt(item.iqs) + '%', fmt(item.iqf) + '%', fmt(item.iqa) + '%'].map(esc).join(';'));
    });

    const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'IQA_Fornecedores.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    exibirToast('sucesso', 'Arquivo CSV gerado com sucesso.');
  }

  // ── Exportar PDF ────────────────────────────────────────────────────────
  function gerarPdfComLogo(logoDataUrl) {
    const { jsPDF } = window.jspdf;
    const doc    = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW  = doc.internal.pageSize.getWidth();

    if (logoDataUrl) doc.addImage(logoDataUrl, 'PNG', 14, 5, 38, 13);

    const now = new Date();
    const pad = function (n) { return String(n).padStart(2, '0'); };
    const dataHora = `${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;

    doc.setFontSize(8); doc.setTextColor(100);
    doc.text('Emitido em: ' + dataHora, pageW - 10, 8, { align: 'right' });

    doc.setFontSize(13); doc.setTextColor(33);
    doc.text('Índice de Qualificação e Avaliação de Desempenho de Fornecedores', pageW / 2, 20, { align: 'center' });

    const head = [['Fornecedor', 'Nome', 'Materiais', 'IQS', 'IQF (Médio)', 'IQA']];
    const body = dadosCompletos.map(function (item) {
      return [item.fornecedor, item.nome, item.materiais,
              fmt(item.iqs) + '%', fmt(item.iqf) + '%', fmt(item.iqa) + '%'];
    });

    doc.autoTable({
      head,
      body,
      startY    : 26,
      theme     : 'striped',
      styles    : { fontSize: 4, cellPadding: 1.5 },
      headStyles: { fillColor: [74,74,74], textColor: [255,255,255] },
      alternateRowStyles: { fillColor: [248,248,248] },
    });

    doc.save('IQA_Fornecedores.pdf');
    exibirToast('sucesso', 'PDF gerado com sucesso.');
  }

  function onExportPdf() {
    if (!dadosCompletos.length) { exibirToast('aviso', 'Nenhum dado disponível para exportar.'); return; }

    const img = new Image();
    img.onload = function () {
      const canvas  = document.createElement('canvas');
      canvas.width  = 900; canvas.height = 315;
      canvas.getContext('2d').drawImage(img, 0, 0, 900, 315);
      gerarPdfComLogo(canvas.toDataURL('image/png'));
    };
    img.onerror = function () {
      exibirToast('info', 'Logo não encontrada — PDF gerado sem logotipo.');
      gerarPdfComLogo(null);
    };
    img.src = LOGO_PATH;
  }

  // ── Limpar ──────────────────────────────────────────────────────────────
  function onLimpar() {
    $el('Filial').value       = '';
    $el('DataIni').value      = '';
    $el('DataFim').value      = '';
    $el('lbFornecedor').value = '';
    ocultarTabela();
  }

  // ── Init ────────────────────────────────────────────────────────────────
  function init() {
    initSearch();
    $el('btnConsultar').addEventListener('click', consultar);
    $el('btnLimpar').addEventListener('click', onLimpar);
    $el('btnExportCsv').addEventListener('click', onExportCsv);
    $el('btnExportPdf').addEventListener('click', onExportPdf);
  }

  $(document).ready(init);

})();
</script>

</body>
</html>
