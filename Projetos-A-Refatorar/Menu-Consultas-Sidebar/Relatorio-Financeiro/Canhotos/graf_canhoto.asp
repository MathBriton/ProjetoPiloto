<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID      = 1046

Dim dtiniPre, dtfinPre, tipoPre, filialPre, respostaPre
dtiniPre    = Trim(Request("dtini"))
dtfinPre    = Trim(Request("dtfin"))
tipoPre     = Trim(Request("tipo"))
filialPre   = Trim(Request("filial"))
respostaPre = Trim(Request("resposta"))
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet — Gráfico Consulta Canhoto</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    .filter-card, .chart-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }

    .filter-card-header, .chart-card-header {
      background-color: #fff;
      color: #212529;
      padding: .6rem 1rem;
      border-radius: 6px 6px 0 0;
      border-bottom: 1px solid #dee2e6;
      font-weight: 600;
      font-size: .9rem;
    }

    .filter-card-body { padding: 1rem 1.25rem 1.25rem; }

    .filter-card-body label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }

    .input-data { width: 155px; }

    /* ── Área do gráfico: altura aumentada para dar espaço vertical ── */
    .chart-wrapper {
      padding: 1.25rem;
      height: 480px;
      position: relative;
    }

    .chart-wrapper canvas {
      width: 100% !important;
      height: 100% !important;
    }

    .chart-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3.5rem 1rem;
      color: #6c757d;
      gap: .75rem;
    }

    .chart-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .chart-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

    .overlay-loading {
      position: fixed;
      inset: 0;
      background: rgba(255,255,255,.72);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 1055;
      gap: 1rem;
    }

    .overlay-loading .spinner-border { width: 3rem; height: 3rem; }
    .overlay-loading span { font-size: .95rem; font-weight: 600; }

    .toast-container-custom {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 1100;
      display: flex;
      flex-direction: column;
      gap: .5rem;
    }

    .toast-custom {
      min-width: 280px;
      max-width: 360px;
      border-radius: 10px;
      overflow: hidden;
      box-shadow: 0 4px 18px rgba(0,0,0,.13);
      background: #fff;
      border: 1px solid #dee2e6;
      animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
    }

    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(40px); }
      to   { opacity: 1; transform: translateX(0); }
    }

    .toast-inner {
      display: flex;
      align-items: flex-start;
      padding: .75rem 1rem;
      gap: .6rem;
    }

    .toast-icon { font-size: 1.15rem; flex-shrink: 0; margin-top: .1rem; }
    .toast-body-text { flex: 1; }
    .toast-title { font-weight: 700; font-size: .82rem; margin-bottom: .1rem; }
    .toast-msg   { font-size: .82rem; color: #495057; }
    .toast-close { background: none; border: none; font-size: 1.1rem; cursor: pointer; color: #6c757d; padding: 0; }

    .toast-sucesso .toast-icon, .toast-sucesso .toast-title { color: #198754; }
    .toast-erro    .toast-icon, .toast-erro    .toast-title { color: #dc3545; }
    .toast-aviso   .toast-icon, .toast-aviso   .toast-title { color: #856404; }
    .toast-info    .toast-icon, .toast-info    .toast-title { color: #6c757d; }

    .toast-progress { height: 3px; }
    .toast-progress-bar {
      height: 100%;
      animation: toastProgressShrink 4.5s linear forwards;
    }

    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    .toast-sucesso .toast-progress-bar { background-color: #198754; }
    .toast-erro    .toast-progress-bar { background-color: #dc3545; }
    .toast-aviso   .toast-progress-bar { background-color: #ffc107; }
    .toast-info    .toast-progress-bar { background-color: #6c757d; }

    .dd-multi-toggle {
      width: 100%;
      text-align: left;
      font-size: .875rem;
      position: relative;
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: .25rem;
      padding: .25rem 2rem .25rem .5rem;
      height: calc(1.5em + .5rem + 2px);
      line-height: 1.5;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      color: #212529;
      cursor: pointer;
    }

    .dd-multi-toggle .chevron {
      position: absolute;
      right: .5rem;
      top: 50%;
      transform: translateY(-50%);
      pointer-events: none;
      color: #6c757d;
      font-size: .75rem;
      transition: transform .2s;
    }

    .dd-multi-toggle.show .chevron {
      transform: translateY(-50%) rotate(180deg);
    }

    .dd-multi-menu {
      min-width: 100%;
      padding: .25rem 0;
    }

    .dd-multi-item {
      display: flex;
      align-items: center;
      gap: .5rem;
      padding: .35rem .75rem;
      font-size: .875rem;
      cursor: pointer;
      user-select: none;
    }

    .dd-multi-item:hover { background-color: #f8f9fa; }
    .dd-multi-item input[type="checkbox"] { cursor: pointer; flex-shrink: 0; }

    .dd-multi-placeholder { color: #6c757d; }
  </style>
</head>
<body>

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border text-secondary" role="status"></div>
    <span>Carregando...</span>
  </div>

  <div id="toastMsg" class="toast-container-custom"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="d-flex align-items-center mb-3">
      <a href="rlto_canhoto.asp" class="btn btn-light btn-sm" title="Voltar">
        <i class="bi bi-arrow-left"></i>
      </a>
      <span class="page-title ms-2">Gráfico Consulta Canhoto</span>
    </div>

    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <form id="form" onsubmit="return false;">
          <div class="row g-3 align-items-end">

            <div class="col-12 col-md-3 col-xl-2">
              <label for="ddFilialToggle">Planta</label>
              <div class="dropdown" id="ddFilial">
                <button id="ddFilialToggle" type="button"
                        class="dd-multi-toggle"
                        data-bs-toggle="dropdown"
                        data-bs-auto-close="outside"
                        aria-expanded="false">
                  <span id="ddFilialLabel" class="dd-multi-placeholder">Todas</span>
                  <i class="bi bi-chevron-down chevron"></i>
                </button>
                <ul class="dropdown-menu dd-multi-menu">
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input dd-todos" value="" /><span>Todos</span></label></li>
                  <li><hr class="dropdown-divider my-1" /></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="0" /><span>Pirapetinga</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="4" /><span>Uberaba</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="5" /><span>Bento</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="6" /><span>Fortaleza</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="7" /><span>Saquarema</span></label></li>
                </ul>
                <input type="hidden" id="filial" name="filial" />
              </div>
            </div>

            <div class="col-12 col-md-3 col-xl-2">
              <label for="ddRespostaToggle">Resposta</label>
              <div class="dropdown" id="ddResposta">
                <button id="ddRespostaToggle" type="button"
                        class="dd-multi-toggle"
                        data-bs-toggle="dropdown"
                        data-bs-auto-close="outside"
                        aria-expanded="false">
                  <span id="ddRespostaLabel" class="dd-multi-placeholder">Todas</span>
                  <i class="bi bi-chevron-down chevron"></i>
                </button>
                <ul class="dropdown-menu dd-multi-menu">
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input dd-todos" value="" /><span>Todos</span></label></li>
                  <li><hr class="dropdown-divider my-1" /></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="0" /><span>Não Respondeu</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="5" /><span>Muito Satisfeito</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="4" /><span>Satisfeito</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="3" /><span>Indiferente</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="2" /><span>Insatisfeito</span></label></li>
                  <li><label class="dd-multi-item"><input type="checkbox" class="form-check-input" value="1" /><span>Muito Insatisfeito</span></label></li>
                </ul>
                <input type="hidden" id="resposta" name="resposta" />
              </div>
            </div>

            <div class="col-6 col-md-2 col-xl-1">
              <label for="tipo">Tipo</label>
              <select id="tipo" name="tipo" class="form-select form-select-sm">
                <option value=""  <% If tipoPre = ""  Then Response.Write "selected" %>>Todos</option>
                <option value="P" <% If tipoPre = "P" Then Response.Write "selected" %>>Produto</option>
                <option value="S" <% If tipoPre = "S" Then Response.Write "selected" %>>Serviço</option>
              </select>
            </div>

            <div class="col-6 col-md-auto">
              <label for="dtini">Data Inicial</label>
              <input id="dtini" name="dtini" type="date" class="form-control form-control-sm input-data" value="<%= dtiniPre %>" />
            </div>

            <div class="col-6 col-md-auto">
              <label for="dtfin">Data Final</label>
              <input id="dtfin" name="dtfin" type="date" class="form-control form-control-sm input-data" value="<%= dtfinPre %>" />
            </div>

            <div class="col-12 col-md-auto">
              <div class="d-flex gap-2">
                <button type="button" id="btnFiltrar" class="btn btn-primary btn-sm">
                  <i class="bi bi-funnel-fill me-2"></i>Filtrar
                </button>
                <button type="button" id="btnLimpar" class="btn btn-light btn-sm">
                  <i class="bi bi-eraser"></i>
                </button>
                <button type="button" id="btnCsatModal" class="btn btn-secondary btn-sm"
                        data-bs-toggle="modal" data-bs-target="#modalCsat" disabled>
                  Tabela CSAT
                </button>
              </div>
            </div>

          </div>
        </form>
      </div>
    </div>

    <div class="chart-card">
      <div class="chart-wrapper">
        <div class="chart-placeholder" id="chartPlaceholder">
          <i class="bi bi-clipboard-data"></i>
          <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir o gráfico.</p>
        </div>
        <canvas id="graficoCanhoto" class="d-none"></canvas>
      </div>
    </div>

    <div id="legendaTipo" class="filter-card mt-2 d-none">
      <div class="filter-card-header">
        <i class="bi bi-circle-half me-2"></i>Legenda
      </div>
      <div id="legendaTipoItens" class="d-flex flex-wrap gap-3 px-3 py-2" style="font-size:.82rem;color:#495057;"></div>
    </div>

  </div>

  <div class="modal fade" id="modalCsat" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title fs-6 fw-bold">Tabela CSAT por Planta</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body p-0">
          <table class="table table-bordered table-hover table-sm mb-0">
            <thead>
              <tr>
                <th class="text-center">Planta</th>
                <th class="text-center">Tipo</th>
                <th class="text-center">CSAT</th>
                <th class="text-center">Média</th>
              </tr>
            </thead>
            <tbody id="tbodyCsat"></tbody>
          </table>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous">
  </script>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>

  <script>
  (() => {
    "use strict";

    const REPOSITORIO = "graf_canhotoRepositorio.asp";
    const TOAST_DELAY = 4500;

    const TOAST_CFG = Object.freeze({
      sucesso: { icon: "bi-check-circle-fill",         title: "Sucesso"    },
      erro:    { icon: "bi-x-circle-fill",             title: "Erro"       },
      aviso:   { icon: "bi-exclamation-triangle-fill", title: "Atenção"    },
      info:    { icon: "bi-info-circle-fill",          title: "Informação" }
    });

    const $el = (id) => document.getElementById(id);

    let chartInstance = null;
    let csatAtual     = [];

    const pluginSeparadorPlantas = {
      id: "separadorPlantas",
      afterDraw(chart) {
        const { ctx, scales } = chart;
        const xScale = scales.x;
        if (!xScale) return;

        // Índices dos separadores no array de labels (posições ímpares)
        const sepIndices = [1, 3, 5, 7];

        ctx.save();
        ctx.strokeStyle = "rgba(150,150,150,0.55)";
        ctx.lineWidth   = 1.5;
        ctx.setLineDash([4, 3]);

        sepIndices.forEach((idx) => {
          const meta = chart.getDatasetMeta(0);
          if (!meta || !meta.data || !meta.data[idx]) return;

          // Posição X central do label separador
          const x = xScale.getPixelForValue(idx);
          const topY    = chart.chartArea.top;
          const bottomY = chart.chartArea.bottom;

          ctx.beginPath();
          ctx.moveTo(x, topY - 4);
          ctx.lineTo(x, bottomY + 4);
          ctx.stroke();
        });

        ctx.restore();
      }
    };

    const pluginRotuloPlanta = {
      id: "rotuloPlanta",
      afterDraw(chart) {
        const { ctx, scales, chartArea } = chart;
        const xScale = scales.x;
        if (!xScale) return;

        // Mapa: índice real da planta → nome
        const plantas = [
          { idx: 0, nome: "Pirapetinga" },
          { idx: 2, nome: "Uberaba"     },
          { idx: 4, nome: "Bento"       },
          { idx: 6, nome: "Fortaleza"   },
          { idx: 8, nome: "Saquarema"   }
        ];

        ctx.save();
        ctx.font         = "bold 11px sans-serif";
        ctx.fillStyle    = "#495057";
        ctx.textAlign    = "center";
        ctx.textBaseline = "bottom";

        plantas.forEach(({ idx, nome }) => {
          const x = xScale.getPixelForValue(idx);
          ctx.fillText(nome, x, chartArea.top - 6);
        });

        ctx.restore();
      }
    };

    function toast(msg, tipo) {
      const t  = TOAST_CFG[tipo] ? tipo : "info";
      const c  = TOAST_CFG[t];
      const el = document.createElement("div");
      el.className = `toast-custom toast-${t}`;
      el.innerHTML = `
        <div class="toast-inner">
          <i class="bi ${c.icon} toast-icon"></i>
          <div class="toast-body-text">
            <div class="toast-title">${c.title}</div>
            <div class="toast-msg">${msg}</div>
          </div>
          <button class="toast-close" aria-label="Fechar">&times;</button>
        </div>
        <div class="toast-progress"><div class="toast-progress-bar"></div></div>`;
      el.querySelector(".toast-close").addEventListener("click", () => el.remove());
      $el("toastMsg").appendChild(el);
      setTimeout(() => { if (el.parentNode) el.remove(); }, TOAST_DELAY);
    }

    function overlay(on) {
      $el("overlayLoading").classList.toggle("d-none", !on);
    }

    function validar() {
      const di = $el("dtini").value;
      const df = $el("dtfin").value;
      if (!di) { toast("Informe a data inicial.", "aviso"); return false; }
      if (!df) { toast("Informe a data final.",   "aviso"); return false; }
      if (new Date(di) > new Date(df)) {
        toast("A data inicial deve ser menor ou igual à data final.", "aviso");
        return false;
      }
      return true;
    }

    function coletarFiltros() {
      return {
        dtini:    $el("dtini").value,
        dtfin:    $el("dtfin").value,
        tipo:     $el("tipo").value,
        filial:   $el("filial").value,
        resposta: $el("resposta").value
      };
    }

    function formatarNumero(v, dec) {
      return Number(v).toLocaleString("pt-BR", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec
      });
    }

    function tickCallbackX(value, index, ticks) {
      const label = this.getLabelForValue(value);
      if (label === "|") return "";   // separador — label invisível
      return "";                      // nome já desenhado pelo plugin
    }

    function renderizarGrafico(labels, datasets) {
      const canvas      = $el("graficoCanhoto");
      const placeholder = $el("chartPlaceholder");
      const legendaTipo = $el("legendaTipo");

      if (!datasets || datasets.length === 0) {
        canvas.classList.add("d-none");
        legendaTipo.classList.add("d-none");
        placeholder.classList.remove("d-none");
        placeholder.querySelector("p").innerHTML =
          "Nenhum dado encontrado para os filtros informados.";
        if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
        return;
      }

      placeholder.classList.add("d-none");
      canvas.classList.remove("d-none");
      legendaTipo.classList.remove("d-none");

      // Popula a legenda DOM com um item por dataset (ignora separadores com null)
      const legendaItens = $el("legendaTipoItens");
      legendaItens.innerHTML = "";
      datasets.forEach((ds) => {
        const item = document.createElement("span");
        item.style.cssText = "display:inline-flex;align-items:center;gap:.35rem;";
        const dot = document.createElement("span");
        dot.style.cssText = `display:inline-block;width:12px;height:12px;border-radius:2px;` +
                            `background-color:${ds.backgroundColor};` +
                            `border:1.5px solid ${ds.borderColor};flex-shrink:0;`;
        const txt = document.createTextNode(ds.label);
        item.appendChild(dot);
        item.appendChild(txt);
        legendaItens.appendChild(item);
      });

      if (chartInstance) chartInstance.destroy();

      chartInstance = new Chart(canvas, {
        type: "bar",
        plugins: [pluginSeparadorPlantas, pluginRotuloPlanta],
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,

          layout: {
            padding: { top: 24 }   // espaço para os rótulos de planta
          },

          plugins: {
            legend: { display: false },
            title: { display: false },
            tooltip: {
              // Não exibe tooltip para labels separadores
              filter: (item) => item.label !== "|",
              callbacks: {
                label: (ctx) => {
                  if (ctx.parsed.y === null) return null;
                  return ` ${ctx.dataset.label}: ${formatarNumero(ctx.parsed.y, 0)} respostas`;
                }
              }
            }
          },

          scales: {
            x: {
              // Largura do separador menor que a das barras reais
              grid:  { display: false },
              ticks: {
                font: { size: 11 },
                callback: tickCallbackX,
                maxRotation: 0
              }
            },
            y: {
              beginAtZero: true,
              ticks: { stepSize: 1, font: { size: 12 } },
              title: {
                display: true,
                text: "Total de Respostas",
                font: { size: 12 }
              }
            }
          }
        }
      });
    }

    function renderizarCsat(csat) {
      csatAtual = csat || [];
      const tbody = $el("tbodyCsat");
      tbody.innerHTML = "";

      $el("btnCsatModal").disabled = csatAtual.length === 0;

      if (csatAtual.length === 0) return;

      const grupos = new Map();
      for (const row of csatAtual) {
        if (!grupos.has(row.planta)) grupos.set(row.planta, []);
        grupos.get(row.planta).push(row);
      }

      for (const [planta, linhas] of grupos) {
        const validas  = linhas.filter(l => l.divisor > 0);
        let mediaTxt   = "";
        if (validas.length > 0) {
          const media = validas.reduce((acc, l) => acc + (l.dividendo / l.divisor), 0) / validas.length * 100;
          mediaTxt = `${formatarNumero(media, 2)}%`;
        }

        linhas.forEach((l, idx) => {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td class="text-center">${l.planta}</td>
            <td class="text-center">${l.tipo}</td>
            <td class="text-center">
              (${formatarNumero(l.dividendo, 0)}/${formatarNumero(l.divisor, 0)})*100 =
              <strong>${formatarNumero(l.percentual, 2)}</strong>%
            </td>`;
          if (idx === 0) {
            const td = document.createElement("td");
            td.className = "text-center align-middle";
            td.rowSpan   = linhas.length;
            td.textContent = mediaTxt;
            tr.appendChild(td);
          }
          tbody.appendChild(tr);
        });
      }
    }

    async function carregar() {
      if (!validar()) return;

      const filtros = coletarFiltros();
      const params  = new URLSearchParams(filtros);

      const url = new URL(window.location.href);
      for (const [k, v] of Object.entries(filtros)) {
        if (v) url.searchParams.set(k, v); else url.searchParams.delete(k);
      }
      window.history.replaceState({}, "", url);

      overlay(true);
      try {
        const resp = await fetch(`${REPOSITORIO}?${params.toString()}`, {
          headers: { "X-Requested-With": "XMLHttpRequest" }
        });
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const data = await resp.json();

        if (data.erro) {
          toast(data.erro, "erro");
          return;
        }

        renderizarGrafico(data.labels, data.datasets);
        renderizarCsat(data.csat);
        toast("Dados carregados.", "sucesso");
      } catch (e) {
        toast(`Falha ao carregar: ${e.message}`, "erro");
      } finally {
        overlay(false);
      }
    }

    function limpar() {
      limparDropdown("ddFilial",   "filial");
      limparDropdown("ddResposta", "resposta");

      $el("tipo").value  = "";
      $el("dtini").value = "";
      $el("dtfin").value = "";

      const url = new URL(window.location.href);
      ["dtini", "dtfin", "tipo", "filial", "resposta"].forEach(k => url.searchParams.delete(k));
      window.history.replaceState({}, "", url);

      if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
      $el("graficoCanhoto").classList.add("d-none");
      $el("legendaTipo").classList.add("d-none");
      $el("chartPlaceholder").classList.remove("d-none");
      $el("chartPlaceholder").querySelector("p").innerHTML =
        "Selecione um período e clique em <strong>Filtrar</strong> para exibir o gráfico.";

      renderizarCsat([]);
    }

    function limparDropdown(ddId, hiddenId) {
      const dd      = $el(ddId);
      const cbTodos = dd.querySelector("input.dd-todos");
      dd.querySelectorAll("input[type='checkbox']").forEach(cb => { cb.checked = false; });
      if (cbTodos) cbTodos.checked = true;
      $el(hiddenId).value = "";
      atualizarLabel(ddId, hiddenId);
    }

    function atualizarLabel(ddId, hiddenId) {
      const dd       = $el(ddId);
      const label    = dd.querySelector("[id$='Label']");
      const cbTodos  = dd.querySelector("input.dd-todos");
      const cbItens  = [...dd.querySelectorAll("input[type='checkbox']")]
                         .filter(cb => !cb.classList.contains("dd-todos"));
      const marcados = cbItens.filter(cb => cb.checked);

      if (cbTodos && cbTodos.checked) {
        cbItens.forEach(cb => { cb.checked = true; });
        $el(hiddenId).value = "";
        label.textContent = "Todos";
        label.classList.remove("dd-multi-placeholder");
        return;
      }

      $el(hiddenId).value = marcados.map(cb => cb.value).join(",");

      if (marcados.length === 0) {
        label.textContent = "Todos";
        label.classList.add("dd-multi-placeholder");
      } else {
        const primeiro = marcados[0].closest("label").querySelector("span").textContent;
        label.textContent = marcados.length === 1
          ? primeiro
          : `${primeiro} +${marcados.length - 1}`;
        label.classList.remove("dd-multi-placeholder");
      }
    }

    function iniciarDropdownMulti(ddId, hiddenId, preVal) {
      const dd      = $el(ddId);
      const toggle  = dd.querySelector("button[data-bs-toggle='dropdown']");
      const cbTodos = dd.querySelector("input.dd-todos");
      const cbItens = [...dd.querySelectorAll("input[type='checkbox']")]
                        .filter(cb => !cb.classList.contains("dd-todos"));

      if (cbTodos) {
        cbTodos.addEventListener("change", () => {
          cbItens.forEach(cb => { cb.checked = cbTodos.checked; });
          atualizarLabel(ddId, hiddenId);
        });
      }

      cbItens.forEach(cb => {
        cb.addEventListener("change", () => {
          if (cbTodos) cbTodos.checked = cbItens.every(c => c.checked);
          atualizarLabel(ddId, hiddenId);
        });
      });

      toggle.addEventListener("show.bs.dropdown", () => toggle.classList.add("show"));
      toggle.addEventListener("hide.bs.dropdown", () => toggle.classList.remove("show"));

      if (preVal) {
        // Pré-seleção vinda da querystring: marca apenas os itens indicados
        preVal.split(",").forEach((v) => {
          const cb = dd.querySelector(`input[value="${v.trim()}"]:not(.dd-todos)`);
          if (cb) cb.checked = true;
        });
        // Verifica se todos ficaram marcados para sincronizar o Todos
        if (cbTodos) cbTodos.checked = cbItens.every(c => c.checked);
        atualizarLabel(ddId, hiddenId);
      } else {
        // Sem pré-seleção → Todos marcado + todos os filhos marcados
        if (cbTodos) cbTodos.checked = true;
        cbItens.forEach(cb => { cb.checked = true; });
        atualizarLabel(ddId, hiddenId);
      }
    }

    function init() {
      iniciarDropdownMulti("ddFilial",   "filial",   "<%= filialPre %>");
      iniciarDropdownMulti("ddResposta", "resposta", "<%= respostaPre %>");
      $el("btnFiltrar").addEventListener("click", carregar);
      $el("btnLimpar").addEventListener("click", limpar);

      if ($el("dtini").value && $el("dtfin").value) carregar();
    }

    document.addEventListener("DOMContentLoaded", init);
  })();
  </script>

</body>
</html>
