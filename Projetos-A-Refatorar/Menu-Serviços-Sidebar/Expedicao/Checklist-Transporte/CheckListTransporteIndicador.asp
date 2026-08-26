<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<%
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046

' ================================================================
' BLOCO DE API — só executa quando vier o parâmetro "acao"
' ================================================================
Dim acao
acao = Trim("" & Request("acao"))

If acao <> "" Then

    Response.Charset     = "UTF-8"
    Response.CodePage    = 65001
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "obterIndicadores" : resultado = ObterIndicadores()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}"
    End Select

    Response.Write resultado
    Response.End()

End If

' ================================================================
' FUNÇÃO: ObterIndicadores
' ================================================================
Function ObterIndicadores()

    Dim dataInicial, dataFinal
    dataInicial = Trim("" & Request("parDataInicial"))
    dataFinal   = Trim("" & Request("parDataFinal"))

    If dataInicial = "" Or dataFinal = "" Then
        ObterIndicadores = "{""sucesso"": false, ""mensagem"": ""Informe o período (parDataInicial e parDataFinal).""}"
        Exit Function
    End If

    ' Sanitização básica
    dataInicial = Replace(Replace(dataInicial, "'", ""), ";", "")
    dataFinal   = Replace(Replace(dataFinal,   "'", ""), ";", "")

    ' ── 1. Proprietários (top 20) ────────────────────────────────
    Dim sqlProprietarios
    sqlProprietarios = "SELECT TOP 20 " & _
                       "  NOMEPROPRIETARIO, " & _
                       "  COUNT(NOMEPROPRIETARIO) AS TOTALVEZES " & _
                       "FROM CHECKLIST_TRANSPORTES " & _
                       "WHERE CAST(DATACADASTRO AS DATE) >= '" & dataInicial & "' " & _
                       "  AND CAST(DATACADASTRO AS DATE) <= '" & dataFinal   & "' " & _
                       "GROUP BY NOMEPROPRIETARIO " & _
                       "ORDER BY TOTALVEZES DESC"

    On Error Resume Next
    Dim rsProprietarios
    Set rsProprietarios = Conecta("PIR-CORPORATIVO", sqlProprietarios)
    If Err.Number <> 0 Or Not IsObject(rsProprietarios) Then
        ObterIndicadores = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar proprietários: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonProprietarios, primeiroP
    primeiroP         = True
    jsonProprietarios = "["
    Do While Not rsProprietarios.EOF
        If Not primeiroP Then jsonProprietarios = jsonProprietarios & ","
        jsonProprietarios = jsonProprietarios & "{"
        jsonProprietarios = jsonProprietarios & """label"": """ & EscapeJSON(rsProprietarios("NOMEPROPRIETARIO") & "") & """, "
        jsonProprietarios = jsonProprietarios & """value"": "  & (rsProprietarios("TOTALVEZES") & "")
        jsonProprietarios = jsonProprietarios & "}"
        primeiroP = False
        rsProprietarios.MoveNext()
    Loop
    jsonProprietarios = jsonProprietarios & "]"
    rsProprietarios.Close()
    Set rsProprietarios = Nothing

    ' ── 2. Transportadoras ───────────────────────────────────────
    Dim sqlTransportadoras
    sqlTransportadoras = "SELECT " & _
                         "  CASE ID_TIPOTRANSPORTADORA " & _
                         "    WHEN '1' THEN 'SMURFIT WESTROCK' " & _
                         "    WHEN '2' THEN 'TRANSPORTADORA' " & _
                         "    WHEN '3' THEN 'FRETEIRO' " & _
                         "    ELSE 'NÃO INFORMADO' " & _
                         "  END AS TRANSPORTADORA, " & _
                         "  COUNT(ID_TIPOTRANSPORTADORA) AS QTDE " & _
                         "FROM CHECKLIST_TRANSPORTES " & _
                         "WHERE CAST(DATACADASTRO AS DATE) >= '" & dataInicial & "' " & _
                         "  AND CAST(DATACADASTRO AS DATE) <= '" & dataFinal   & "' " & _
                         "GROUP BY ID_TIPOTRANSPORTADORA " & _
                         "ORDER BY QTDE DESC"

    On Error Resume Next
    Dim rsTransportadoras
    Set rsTransportadoras = Conecta("PIR-CORPORATIVO", sqlTransportadoras)
    If Err.Number <> 0 Or Not IsObject(rsTransportadoras) Then
        ObterIndicadores = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar transportadoras: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonTransportadoras, primeiroT
    primeiroT           = True
    jsonTransportadoras = "["
    Do While Not rsTransportadoras.EOF
        If Not primeiroT Then jsonTransportadoras = jsonTransportadoras & ","
        jsonTransportadoras = jsonTransportadoras & "{"
        jsonTransportadoras = jsonTransportadoras & """label"": """ & EscapeJSON(rsTransportadoras("TRANSPORTADORA") & "") & """, "
        jsonTransportadoras = jsonTransportadoras & """value"": "  & (rsTransportadoras("QTDE") & "")
        jsonTransportadoras = jsonTransportadoras & "}"
        primeiroT = False
        rsTransportadoras.MoveNext()
    Loop
    jsonTransportadoras = jsonTransportadoras & "]"
    rsTransportadoras.Close()
    Set rsTransportadoras = Nothing

    ' ── 3. Tipos de Veículo ───────────────────────────────────────
    Dim sqlTipos
    sqlTipos = "SELECT " & _
               "  ISNULL(T2.DESCRICAO, 'NÃO INFORMADO') AS TIPOVEICULO, " & _
               "  COUNT(T1.ID) AS QTDE " & _
               "FROM CHECKLIST_TRANSPORTES T1 " & _
               "  LEFT JOIN CHECKLIST_TIPOVEICULO T2 ON T2.ID = T1.ID_TIPOVEICULO " & _
               "WHERE CAST(T1.DATACADASTRO AS DATE) >= '" & dataInicial & "' " & _
               "  AND CAST(T1.DATACADASTRO AS DATE) <= '" & dataFinal   & "' " & _
               "GROUP BY T2.DESCRICAO " & _
               "ORDER BY QTDE DESC"

    On Error Resume Next
    Dim rsTipos
    Set rsTipos = Conecta("PIR-CORPORATIVO", sqlTipos)
    If Err.Number <> 0 Or Not IsObject(rsTipos) Then
        ObterIndicadores = "{""sucesso"": false, ""mensagem"": ""Erro ao consultar tipos de veículo: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    Dim jsonTipos, primeiroV
    primeiroV = True
    jsonTipos = "["
    Do While Not rsTipos.EOF
        If Not primeiroV Then jsonTipos = jsonTipos & ","
        jsonTipos = jsonTipos & "{"
        jsonTipos = jsonTipos & """label"": """ & EscapeJSON(rsTipos("TIPOVEICULO") & "") & """, "
        jsonTipos = jsonTipos & """value"": "  & (rsTipos("QTDE") & "")
        jsonTipos = jsonTipos & "}"
        primeiroV = False
        rsTipos.MoveNext()
    Loop
    jsonTipos = jsonTipos & "]"
    rsTipos.Close()
    Set rsTipos = Nothing

    ' ── Resposta final ────────────────────────────────────────────
    Dim json
    json = "{"
    json = json & """sucesso"": true, "
    json = json & """mensagem"": """", "
    json = json & """dados"": {"
    json = json & """proprietarios"": "  & jsonProprietarios  & ", "
    json = json & """transportadoras"": " & jsonTransportadoras & ", "
    json = json & """tiposVeiculo"": "   & jsonTipos
    json = json & "}}"

    ObterIndicadores = json

End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Indicadores — CheckList Transporte</title>

  <link href="css/bootstrap.min.css" rel="stylesheet" />
  <link href="css/bootstrap-icons.min.css" rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    /* Filtros */
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
    .filter-card-body  { padding: 1rem 1.25rem 1.25rem; }
    .filter-label      { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }
    .input-data        { width: 155px; }

    /* Cards de gráfico */
    .chart-card {
      background-color: #fff;
      color: #212529;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      position: relative;
    }
    .chart-card-header {
      background-color: #cfcfcf;
      color: #050000;
      padding: .55rem 1rem;
      border-radius: 6px 6px 0 0;
      font-weight: 600;
      font-size: .88rem;
      display: flex;
      align-items: center;
      gap: .5rem;
    }
    .chart-card-body { padding: 1.25rem; }
    .chart-card-body-rosca { max-width: 420px; margin: 0 auto; }

    /* Overlay */
    .overlay-loading {
      position: absolute; inset: 0;
      background: rgba(255,255,255,.80);
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      z-index: 10; gap: 1rem; border-radius: 6px;
    }
    .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
    .overlay-loading span            { font-size: .9rem; font-weight: 600; color: #495057; }

    /* Placeholder */
    .chart-placeholder {
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      padding: 3.5rem 1rem; color: #6c757d; gap: .75rem;
    }
    .chart-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .chart-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

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

    #badgePeriodo { font-size: .8rem; }
  </style>
</head>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center justify-content-between mb-3 flex-wrap gap-2">
      <div class="d-flex align-items-center gap-2">
        <button type="button" class="btn btn-light btn-sm" id="btnVoltar">
          <i class="bi bi-arrow-left me-2"></i>Voltar
        </button>
        <span class="page-title">Indicadores — CheckList de Transporte</span>
        <span id="badgePeriodo" class="badge bg-secondary d-none"></span>
      </div>
    </div>

    <!-- Filtros -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <div class="row g-3 align-items-end">
          <div class="col-auto">
            <label for="parDataInicial" class="form-label filter-label mb-1">Data Inicial</label>
            <input type="date" id="parDataInicial" class="form-control form-control-sm input-data" />
          </div>
          <div class="col-auto">
            <label for="parDataFinal" class="form-label filter-label mb-1">Data Final</label>
            <input type="date" id="parDataFinal" class="form-control form-control-sm input-data" />
          </div>
          <div class="col-auto">
            <div class="d-flex gap-2">
              <button id="btnFiltrar" type="button" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-1"></i>Filtrar
              </button>
              <button id="btnLimpar" type="button" class="btn btn-light btn-sm">
                <i class="bi bi-eraser me-1"></i>Limpar
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Gráficos: Transportadoras + Tipos de Veículo -->
    <div class="row g-3 mb-3 justify-content-center">

      <div class="col-12 col-md-6 col-xl-4">
        <div class="chart-card h-100">
          <div class="chart-card-header">
            <i class="bi bi-truck me-2"></i>Por Tipo de Transporte
          </div>
          <div class="chart-card-body chart-card-body-rosca">
            <div id="placeholderTransporte" class="chart-placeholder">
              <i class="bi bi-bar-chart me-2"></i>
              <p>Selecione um período e clique em <strong>Filtrar</strong>.</p>
            </div>
            <div id="wrapperTransporte" class="d-none">
              <canvas id="chartTransporte" height="160"></canvas>
            </div>
          </div>
          <div id="overlayTransporte" class="overlay-loading d-none">
            <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
            <span>Carregando...</span>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6 col-xl-4">
        <div class="chart-card h-100">
          <div class="chart-card-header">
            <i class="bi bi-ev-front me-2"></i>Por Tipo de Veículo
          </div>
          <div class="chart-card-body chart-card-body-rosca">
            <div id="placeholderVeiculo" class="chart-placeholder">
              <i class="bi bi-bar-chart"></i>
              <p>Selecione um período e clique em <strong>Filtrar</strong>.</p>
            </div>
            <div id="wrapperVeiculo" class="d-none">
              <canvas id="chartVeiculo" height="160"></canvas>
            </div>
          </div>
          <div id="overlayVeiculo" class="overlay-loading d-none">
            <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
            <span>Carregando...</span>
          </div>
        </div>
      </div>

    </div>

    <!-- Gráfico: Top 20 Proprietários -->
    <div class="chart-card mb-3">
      <div class="chart-card-header">
        <i class="bi bi-person-lines-fill me-2"></i>Top 20 — Por Proprietário
      </div>
      <div class="chart-card-body">
        <div id="placeholderProprietario" class="chart-placeholder">
          <i class="bi bi-bar-chart"></i>
          <p>Selecione um período e clique em <strong>Filtrar</strong>.</p>
        </div>
        <div id="wrapperProprietario" class="d-none">
          <canvas id="chartProprietario" height="85"></canvas>
        </div>
      </div>
      <div id="overlayProprietario" class="overlay-loading d-none">
        <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
        <span>Carregando...</span>
      </div>
    </div>

  </div>

  <!-- Toast container -->
  <div class="toast-container-custom" id="toastContainer"></div>

  <script src="js/jquery.min.js"></script>
  <script src="js/bootstrap.bundle.min.js"></script>

  <!-- Chart.js -->
  <script src="js/chart.umd.min.js"></script>

  <script>
    const CheckListIndicador = (function () {

      'use strict';

      const ENDPOINT   = 'CheckListTransporteIndicador.asp';
      const TOAST_DELAY = 4500;

      const CORES_ROSCA = [
        '#1a6fb5', '#2a9d3f', '#e07b25',
        '#8b2fc9', '#c0392b', '#16a085', '#d4ac0d'
      ];

      let chartProprietario = null;
      let chartTransporte   = null;
      let chartVeiculo      = null;

      /* ── Toast ─────────────────────────────────────────────────── */
      const TOAST_ICON = {
        sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754' },
        erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545' },
        aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404' },
        info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d' }
      };

      function exibirToast(tipo, mensagem) {
        const t      = TOAST_ICON[tipo] || TOAST_ICON.info;
        const barCor = tipo === 'aviso' ? '#ffc107' : t.cor;
        const el = $(
          '<div class="toast-custom p-3">' +
            '<div class="d-flex align-items-start gap-2">' +
              '<i class="bi ' + t.icon + ' mt-1" style="color:' + t.cor + ';font-size:1.1rem;"></i>' +
              '<div class="flex-grow-1">' +
                '<div style="font-size:.85rem;font-weight:600;color:' + t.cor + ';">' + mensagem + '</div>' +
              '</div>' +
              '<button class="btn-close btn-close-sm ms-auto" style="font-size:.7rem;" ' +
                      'onclick="this.closest(\'.toast-custom\').remove()"></button>' +
            '</div>' +
            '<div class="mt-2" style="height:3px;background:#e9ecef;border-radius:2px;">' +
              '<div class="toast-progress-bar" ' +
                   'style="height:100%;background:' + barCor + ';border-radius:2px;"></div>' +
            '</div>' +
          '</div>'
        );
        $('#toastContainer').append(el);
        setTimeout(function () { el.remove(); }, TOAST_DELAY);
      }

      /* ── Overlay por seção ─────────────────────────────────────── */
      function mostrarOverlay(secao) { $('#overlay' + secao).removeClass('d-none'); }
      function ocultarOverlay(secao) { $('#overlay' + secao).addClass('d-none');    }

      /* ── Formatar data ISO → DD/MM/AAAA ───────────────────────── */
      function formatarDataBR(iso) {
        if (!iso) return '';
        const p = iso.split('-');
        return p[2] + '/' + p[1] + '/' + p[0];
      }

      /* ── Destruir instância anterior do Chart.js ───────────────── */
      function destruirChart(ref) {
        if (ref) ref.destroy();
        return null;
      }

      /* ── Barras horizontais (Proprietários) ────────────────────── */
      function renderizarBarras(canvasId, dados) {
        return new Chart(document.getElementById(canvasId), {
          type: 'bar',
          data: {
            labels: dados.map(function (d) { return d.label; }),
            datasets: [{
              label: 'Visitas',
              data: dados.map(function (d) { return d.value; }),
              backgroundColor: '#1a6fb5',
              borderRadius: 4,
              borderSkipped: false
            }]
          },
          options: {
            indexAxis: 'y',
            responsive: true,
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: function (ctx) { return ' ' + ctx.raw + ' visita(s)'; }
                }
              }
            },
            scales: {
              x: { beginAtZero: true, ticks: { font: { size: 11 } }, grid: { color: '#e9ecef' } },
              y: { ticks: { font: { size: 11 }, autoSkip: false }, grid: { display: false } }
            }
          }
        });
      }

      /* ── Rosca (Transportadoras / Tipos de Veículo) ────────────── */
      function renderizarRosca(canvasId, dados) {
        const cores = dados.map(function (_, i) { return CORES_ROSCA[i % CORES_ROSCA.length]; });
        return new Chart(document.getElementById(canvasId), {
          type: 'doughnut',
          data: {
            labels: dados.map(function (d) { return d.label; }),
            datasets: [{
              data: dados.map(function (d) { return d.value; }),
              backgroundColor: cores,
              borderWidth: 2,
              borderColor: '#fff'
            }]
          },
          options: {
            responsive: true,
            plugins: {
              legend: { position: 'bottom', labels: { font: { size: 12 }, padding: 16 } },
              tooltip: {
                callbacks: {
                  label: function (ctx) {
                    const total = ctx.dataset.data.reduce(function (a, b) { return a + b; }, 0);
                    const pct   = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : '0.0';
                    return ' ' + ctx.raw + ' (' + pct + '%)';
                  }
                }
              }
            }
          }
        });
      }

      /* ── Carregar indicadores via AJAX ─────────────────────────── */
      function carregarIndicadores(dataIni, dataFim) {
        ['Proprietario', 'Transporte', 'Veiculo'].forEach(mostrarOverlay);

        $.ajax({
          url      : ENDPOINT,
          method   : 'GET',
          dataType : 'json',
          data     : { acao: 'obterIndicadores', parDataInicial: dataIni, parDataFinal: dataFim },
          success: function (resp) {
            ['Proprietario', 'Transporte', 'Veiculo'].forEach(ocultarOverlay);

            if (!resp.sucesso) {
              exibirToast('erro', resp.mensagem || 'Erro ao carregar indicadores.');
              return;
            }

            const d = resp.dados;

            $('#badgePeriodo')
              .text(formatarDataBR(dataIni) + ' a ' + formatarDataBR(dataFim))
              .removeClass('d-none');

            /* Proprietários */
            if (d.proprietarios && d.proprietarios.length > 0) {
              $('#placeholderProprietario').addClass('d-none');
              $('#wrapperProprietario').removeClass('d-none');
              chartProprietario = destruirChart(chartProprietario);
              chartProprietario = renderizarBarras('chartProprietario', d.proprietarios);
            } else {
              $('#wrapperProprietario').addClass('d-none');
              $('#placeholderProprietario').removeClass('d-none');
            }

            /* Transportadoras */
            if (d.transportadoras && d.transportadoras.length > 0) {
              $('#placeholderTransporte').addClass('d-none');
              $('#wrapperTransporte').removeClass('d-none');
              chartTransporte = destruirChart(chartTransporte);
              chartTransporte = renderizarRosca('chartTransporte', d.transportadoras);
            } else {
              $('#wrapperTransporte').addClass('d-none');
              $('#placeholderTransporte').removeClass('d-none');
            }

            /* Tipos de Veículo */
            if (d.tiposVeiculo && d.tiposVeiculo.length > 0) {
              $('#placeholderVeiculo').addClass('d-none');
              $('#wrapperVeiculo').removeClass('d-none');
              chartVeiculo = destruirChart(chartVeiculo);
              chartVeiculo = renderizarRosca('chartVeiculo', d.tiposVeiculo);
            } else {
              $('#wrapperVeiculo').addClass('d-none');
              $('#placeholderVeiculo').removeClass('d-none');
            }
          },
          error: function () {
            ['Proprietario', 'Transporte', 'Veiculo'].forEach(ocultarOverlay);
            exibirToast('erro', 'Erro de comunicação com o servidor.');
          }
        });
      }

      /* ── Filtrar ────────────────────────────────────────────────── */
      function filtrar() {
        const dataIni = $('#parDataInicial').val();
        const dataFim = $('#parDataFinal').val();
        if (!dataIni) { exibirToast('aviso', 'Informe a Data Inicial.'); return; }
        if (!dataFim) { exibirToast('aviso', 'Informe a Data Final.');   return; }
        if (dataIni > dataFim) {
          exibirToast('aviso', 'A Data Inicial não pode ser posterior à Data Final.');
          return;
        }
        carregarIndicadores(dataIni, dataFim);
      }

      /* ── Limpar ─────────────────────────────────────────────────── */
      function limpar() {
        $('#parDataInicial, #parDataFinal').val('');
        $('#badgePeriodo').addClass('d-none');
        ['Proprietario', 'Transporte', 'Veiculo'].forEach(function (s) {
          $('#wrapper'     + s).addClass('d-none');
          $('#placeholder' + s).removeClass('d-none');
        });
        chartProprietario = destruirChart(chartProprietario);
        chartTransporte   = destruirChart(chartTransporte);
        chartVeiculo      = destruirChart(chartVeiculo);
      }

      /* ── Init ───────────────────────────────────────────────────── */
      function init() {
        $('#btnFiltrar').on('click', function () { filtrar(); });
        $('#btnLimpar').on('click',  function () { limpar();  });
        $('#btnVoltar').on('click',  function () { history.back(); });
        $('#parDataInicial, #parDataFinal').on('keydown', function (e) {
          if (e.key === 'Enter') filtrar();
        });

        /* Pré-preencher datas se vier da querystring (redirecionamento da Grade) */
        const params  = new URLSearchParams(window.location.search);
        const dataIni = params.get('parDataInicial') || '';
        const dataFim = params.get('parDataFinal')   || '';
        if (dataIni) $('#parDataInicial').val(dataIni);
        if (dataFim) $('#parDataFinal').val(dataFim);
        if (dataIni && dataFim) carregarIndicadores(dataIni, dataFim);
      }

      return Object.freeze({ init: init });

    }());

    $(function () { CheckListIndicador.init(); });
  </script>
</body>
</html>
