<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp"-->
<!--#include file="../../../funcoes/funcoes.asp"-->
<!--#include file="../../../includes/variaveis.asp"-->
<%
Response.Buffer   = True
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Response.Expires  = 0
Response.CacheControl = "no-cache"
Session.LCID = 1046

Session("Modulo")   = Modulo
Session("Login")    = USR
Session("IdUsuario") = IdUsuario
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <title>SWnet – Relatório Matriz por Período</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA"        content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    /* ── Base ───────────────────────────────────────────────────────────────── */
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }

    /* ── Filtros ────────────────────────────────────────────────────────────── */
    .filter-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }
    .filter-card-header {
      background-color: #fff;
      color: #212529;
      padding: 0.6rem 1rem;
      border-radius: 6px 6px 0 0;
      border-bottom: 1px solid #dee2e6;
      font-weight: 600;
      font-size: 0.9rem;
    }
    .filter-card-body {
      padding: 1rem 1.25rem 1.25rem;
    }
    .filter-label {
      font-size: 0.82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: 0.25rem;
    }
    .input-data {
      width: 155px;
    }

    /* ── Tabela ─────────────────────────────────────────────────────────────── */
    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }

    .table-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3.5rem 1rem;
      color: #6c757d;
      gap: 0.75rem;
    }
    .table-placeholder .bi {
      font-size: 2.8rem;
      opacity: 0.35;
    }
    .table-placeholder p {
      margin: 0;
      font-size: 0.9rem;
      text-align: center;
    }

    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 68vh;
    }
    .table-wrapper table {
      min-width: max-content;
    }
    .table thead th {
      position: sticky;
      top: 0;
      background-color: #fff;
      color: #212529;
      font-size: 0.82rem;
      font-weight: 700;
      text-align: center;
      white-space: nowrap;
      z-index: 2;
      border-color: #dee2e6;
      padding: 0.55rem 0.75rem;
    }
    .table tbody td {
      font-size: 0.82rem;
      text-align: center;
      white-space: nowrap;
      color: #212529;
      vertical-align: middle;
      padding: 0.45rem 0.75rem;
    }
    .table tfoot td {
      font-size: 0.82rem;
      font-weight: 700;
      white-space: nowrap;
      color: #212529;
      padding: 0.45rem 0.75rem;
    }

    /* ── Rodapé da tabela (paginação + info) ────────────────────────────────── */
    .table-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0.6rem 1rem;
      border-top: 1px solid #dee2e6;
      flex-wrap: wrap;
      gap: 0.5rem;
    }
    .table-footer-left {
      min-width: 120px;
    }
    .table-footer-center {
      flex: 1;
      display: flex;
      justify-content: center;
    }
    .table-footer-right {
      min-width: 120px;
      text-align: right;
    }
    .info-registros {
      font-size: 0.8rem;
    }
    .select-limite {
      width: auto;
    }

    /* ── Paginação ──────────────────────────────────────────────────────────── */
    #paginacao .page-link {
      color: #6c757d;
      border-color: #dee2e6;
    }
    #paginacao .page-item.active .page-link {
      background-color: #6c757d;
      border-color: #6c757d;
      color: #fff;
    }
    #paginacao .page-link:hover {
      background-color: #e9ecef;
      color: #495057;
    }
    #paginacao .page-item.disabled .page-link {
      color: #adb5bd;
    }

    /* ── Busca ──────────────────────────────────────────────────────────────── */
    .search-box {
      position: relative;
      max-width: 180px;
    }
    .search-box .form-control {
      padding-right: 72px;
      font-size: 0.82rem;
    }
    .search-box .search-hotkey {
      position: absolute;
      right: 8px;
      top: 50%;
      transform: translateY(-50%);
      pointer-events: none;
      user-select: none;
    }
    .search-box .search-hotkey kbd {
      font-size: 0.65rem;
      padding: 0.1rem 0.3rem;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      color: #6c757d;
    }

    /* ── Overlay de carregamento ────────────────────────────────────────────── */
    .overlay-loading {
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, 0.72);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 1055;
      gap: 1rem;
    }
    .overlay-loading .spinner-border {
      width: 3rem;
      height: 3rem;
    }
    .overlay-loading span {
      font-size: 0.95rem;
      font-weight: 600;
    }

    /* ── Toast ──────────────────────────────────────────────────────────────── */
    .toast-container-custom {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 1100;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }
    .toast-custom {
      min-width: 280px;
      max-width: 360px;
      border-radius: 10px;
      overflow: hidden;
      box-shadow: 0 4px 18px rgba(0, 0, 0, 0.13);
      background: #fff;
      border: 1px solid #dee2e6;
      animation: toastSlideIn 0.28s cubic-bezier(0.22, 1, 0.36, 1);
    }
    @keyframes toastSlideIn {
      from {
        opacity: 0;
        transform: translateX(40px);
      }
      to {
        opacity: 1;
        transform: translateX(0);
      }
    }
    .toast-custom .toast-progress-bar {
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from {
        width: 100%;
      }
      to {
        width: 0%;
      }
    }

  </style>
</head>
<body>

<!-- ── Overlay de carregamento ─────────────────────────────────────────────── -->
<div id="overlayLoading" class="overlay-loading d-none">
  <div class="spinner-border text-secondary" role="status"></div>
  <span>Carregando...</span>
</div>

<!-- ── Contêiner de toasts ─────────────────────────────────────────────────── -->
<div class="toast-container-custom" id="toastContainer"></div>

<div class="container-fluid py-3 px-3 px-lg-4">

  <!-- ── Cabeçalho da página ───────────────────────────────────────────────── -->
  <div class="d-flex align-items-center justify-content-between mb-3">
    <span class="page-title">Relatório Matriz por Período</span>
  </div>

  <!-- ── Card de filtros ───────────────────────────────────────────────────── -->
  <div class="filter-card mb-3">
    <div class="filter-card-header">
      <i class="bi bi-funnel me-2"></i>Filtros
    </div>
    <div class="filter-card-body">
      <div class="row g-3 align-items-end">

        <div class="col-auto">
          <label class="filter-label d-block" for="cboPlanta">Filial</label>
          <select id="cboPlanta" class="form-select form-select-sm" style="min-width:160px;">
            <option value="">Selecione...</option>
            <option value="PIRAPETINGA">Pirapetinga</option>
            <option value="UBERABA">Uberaba</option>
            <option value="BENTO">Bento Gonçalves</option>
            <option value="MARANGUAPE">Maranguape</option>
            <option value="SAQUAREMA">Saquarema</option>
          </select>
        </div>

        <div class="col-auto">
          <label class="filter-label d-block" for="txtFaca">Faca / Matriz</label>
          <input id="txtFaca" type="text" maxlength="6" placeholder="Código"
                 class="form-control form-control-sm" style="width:120px;" />
        </div>

        <div class="col-auto">
          <label class="filter-label d-block" for="dtInicial">Data Inicial</label>
          <input id="dtInicial" type="date" class="form-control form-control-sm input-data" />
        </div>

        <div class="col-auto">
          <label class="filter-label d-block" for="dtFinal">Data Final</label>
          <input id="dtFinal" type="date" class="form-control form-control-sm input-data" />
        </div>

        <div class="col-auto">
          <div class="d-flex gap-2">
            <button id="btnFiltrar" class="btn btn-primary btn-sm" type="button">
              <i class="bi bi-funnel-fill me-2"></i>Filtrar
            </button>
            <button id="btnLimpar" class="btn btn-light btn-sm" type="button" title="Limpar filtro">
              <i class="bi bi-eraser"></i>
            </button>
            <button id="btnExcel" class="btn btn-success btn-sm" type="button" title="Exportar Excel" disabled>
              <i class="bi bi-file-earmark-excel"></i>
            </button>
            <button id="btnPdf" class="btn btn-danger btn-sm" type="button" title="Exportar PDF" disabled>
              <i class="bi bi-file-earmark-pdf"></i>
            </button>
          </div>
        </div>

      </div>
    </div>
  </div>

  <!-- ── Card da tabela ────────────────────────────────────────────────────── -->
  <div class="table-card">

    <!-- Placeholder inicial (sem dados) -->
    <div id="areaPlaceholder" class="table-placeholder">
      <i class="bi bi-clipboard-data"></i>
      <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
    </div>

    <!-- Área de dados (visível após consulta com resultados) -->
    <div id="areaTabela" class="d-none">

      <div class="d-flex justify-content-end px-3 pt-3">
        <div class="search-box">
          <input id="txtBusca" type="text" class="form-control form-control-sm" placeholder="Buscar...">
          <span class="search-hotkey"><kbd>CTRL+K</kbd></span>
        </div>
      </div>

      <div class="table-wrapper px-3 pt-2">
        <table id="tblDados" class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead>
            <tr>
              <th>Item</th>
              <th>Data</th>
              <th>Máquina</th>
              <th>Clichê</th>
              <th>Turma</th>
              <th>Qtde. Prog.</th>
            </tr>
          </thead>
          <tbody id="tblBody"></tbody>
          <tfoot>
            <tr>
              <td colspan="5" class="text-end">Total:</td>
              <td id="tdTotal" class="text-center"></td>
            </tr>
          </tfoot>
        </table>
      </div>

      <!-- Rodapé: itens por página / paginação / contador -->
      <div class="table-footer">
        <div class="table-footer-left">
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
          <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
        </div>
        <div class="table-footer-right">
          <span id="infoRegistros" class="info-registros text-muted"></span>
        </div>
      </div>

    </div>
  </div>

</div>

<!-- Campo oculto para passar o login do usuário ao JS -->
<input type="hidden" id="swnetLogin" value="<%= USR %>" />

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

<script>
(function () {
  "use strict";

  /* ── Constants ─────────────────────────────────────────────────── */
  const REPOSITORIO = "RelatorioMatrizPorPeriodoRepositorio.asp";
  const LOGO_PATH = "Logo/logo-smurfit1.png";
  const TOAST_DELAY = 4500;

  const TOAST_ICON = {
    sucesso: { icon: "bi-check-circle-fill", cor: "#198754", barra: "#198754" },
    erro: { icon: "bi-x-circle-fill", cor: "#dc3545", barra: "#dc3545" },
    aviso: {
      icon: "bi-exclamation-triangle-fill",
      cor: "#856404",
      barra: "#ffc107",
    },
    info: { icon: "bi-info-circle-fill", cor: "#6c757d", barra: "#6c757d" },
  };

  /* ── State ──────────────────────────────────────────────────────── */
  let dadosCompletos = [];
  let dadosFiltrados = [];
  let paginaAtual = 1;
  let totalPaginas = 1;

  /* ── Helpers ────────────────────────────────────────────────────── */
  const $el = (id) => document.getElementById(id);

  function fmtNum(v) {
    if (v === null || v === undefined || v === "") return "";
    return Number(v).toLocaleString("pt-BR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    });
  }

  function somaTotal(arr) {
    return arr.reduce((acc, r) => acc + (Number(r.qtdeProg) || 0), 0);
  }

  /* ── Toast ──────────────────────────────────────────────────────── */
  function exibirToast(tipo, titulo, mensagem) {
    const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
    const el = document.createElement("div");
    el.className = "toast-custom";
    el.innerHTML = `
      <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
        <i class="bi ${cfg.icon}" style="color:${cfg.cor};font-size:1.1rem;"></i>
        <strong class="flex-grow-1" style="color:${cfg.cor};font-size:.9rem;">${titulo}</strong>
        <button style="background:none;border:none;font-size:1rem;cursor:pointer;color:#6c757d;"
                onclick="this.closest('.toast-custom').remove()">&times;</button>
      </div>
      <div class="px-3 pb-2" style="font-size:.82rem;color:#495057;">${mensagem}</div>
      <div style="height:3px;background:#e9ecef;">
        <div class="toast-progress-bar" style="height:100%;background:${cfg.barra};"></div>
      </div>`;
    $el("toastContainer").appendChild(el);
    setTimeout(() => el.remove(), TOAST_DELAY);
  }

  /* ── Overlay ────────────────────────────────────────────────────── */
  function showOverlay() {
    $el("overlayLoading").classList.remove("d-none");
  }
  function hideOverlay() {
    $el("overlayLoading").classList.add("d-none");
  }

  /* ── Busca em memória ───────────────────────────────────────────── */
  function aplicarBusca(termo) {
    const t = termo.trim().toLowerCase();
    dadosFiltrados =
      t === ""
        ? dadosCompletos.slice()
        : dadosCompletos.filter(
            (r) =>
              (r.item || "").toLowerCase().includes(t) ||
              (r.inicio || "").toLowerCase().includes(t) ||
              (r.maquina || "").toLowerCase().includes(t) ||
              (r.nomeMaq || "").toLowerCase().includes(t) ||
              (r.codFI || "").toLowerCase().includes(t) ||
              (r.turma || "").toLowerCase().includes(t),
          );
    paginaAtual = 1;
    renderTabela();
  }

  /* ── Render ─────────────────────────────────────────────────────── */
  function renderTabela() {
    const limite = parseInt($el("selectLimite").value, 10);
    totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / limite));
    paginaAtual = Math.min(paginaAtual, totalPaginas);

    const inicio = (paginaAtual - 1) * limite;
    const pagina = dadosFiltrados.slice(inicio, inicio + limite);

    const tbody = $el("tblBody");
    if (pagina.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">Nenhum registro encontrado.</td></tr>`;
    } else {
      tbody.innerHTML = pagina
        .map(
          (r) => `
        <tr>
          <td>${r.item || ""}</td>
          <td>${r.inicio || ""}</td>
          <td class="text-center">${r.maquina ? r.maquina + " - " + (r.nomeMaq || "") : r.nomeMaq || ""}</td>
          <td>${r.codFI || ""}</td>
          <td>${r.turma || ""}</td>
          <td class="text-center">${fmtNum(r.qtdeProg)}</td>
        </tr>`,
        )
        .join("");
    }

    const totalGeral = somaTotal(dadosFiltrados);
    $el("tdTotal").textContent = fmtNum(totalGeral);

    const de = dadosFiltrados.length === 0 ? 0 : inicio + 1;
    const ate = Math.min(inicio + limite, dadosFiltrados.length);
    $el("infoRegistros").textContent =
      `${de}–${ate} de ${dadosFiltrados.length}`;

    renderPaginacao(limite);
  }

  /* ── Paginação ──────────────────────────────────────────────────── */
  function renderPaginacao(limite) {
    const ul = $el("paginacao");
    const half = 2;
    const start = Math.max(1, paginaAtual - half);
    const end = Math.min(totalPaginas, paginaAtual + half);
    let html = "";

    const pgItem = (lbl, pg, disabled, active) =>
      `<li class="page-item${disabled ? " disabled" : ""}${active ? " active" : ""}">
             <a class="page-link" href="#" data-pg="${pg}">${lbl}</a>
           </li>`;

    html += pgItem("&laquo;", paginaAtual - 1, paginaAtual === 1);
    if (start > 1) {
      html += pgItem("1", 1);
      if (start > 2) html += pgItem("…", paginaAtual, true);
    }
    for (let i = start; i <= end; i++)
      html += pgItem(i, i, false, i === paginaAtual);
    if (end < totalPaginas) {
      if (end < totalPaginas - 1) html += pgItem("…", paginaAtual, true);
      html += pgItem(totalPaginas, totalPaginas);
    }
    html += pgItem("&raquo;", paginaAtual + 1, paginaAtual === totalPaginas);

    ul.innerHTML = html;
    ul.querySelectorAll("a.page-link").forEach((a) => {
      a.addEventListener("click", (e) => {
        e.preventDefault();
        const pg = parseInt(a.dataset.pg, 10);
        if (!isNaN(pg) && pg >= 1 && pg <= totalPaginas) {
          paginaAtual = pg;
          renderTabela();
        }
      });
    });
  }

  /* ── Filtrar (AJAX) ─────────────────────────────────────────────── */
  function filtrar() {
    const planta = $el("cboPlanta").value.trim();
    const faca = $el("txtFaca").value.trim();
    const di = $el("dtInicial").value.trim();
    const df = $el("dtFinal").value.trim();

    if (!planta) {
      exibirToast("aviso", "Atenção", "Selecione uma filial.");
      return;
    }
    if (!di) {
      exibirToast("aviso", "Atenção", "Informe a data inicial.");
      return;
    }
    if (!df) {
      exibirToast("aviso", "Atenção", "Informe a data final.");
      return;
    }
    if (di > df) {
      exibirToast(
        "aviso",
        "Atenção",
        "A data inicial deve ser menor ou igual à data final.",
      );
      return;
    }

    showOverlay();

    const params =
      "acao=FILTRAR" +
      "&planta=" +
      encodeURIComponent(planta) +
      "&faca=" +
      encodeURIComponent(faca) +
      "&datainicial=" +
      encodeURIComponent(di) +
      "&datafinal=" +
      encodeURIComponent(df);

    const xhr = new XMLHttpRequest();
    xhr.open("POST", REPOSITORIO, true);
    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      hideOverlay();
      if (xhr.status !== 200) {
        exibirToast(
          "erro",
          "Erro",
          "Falha na comunicação com o servidor (" + xhr.status + ").",
        );
        return;
      }
      let dados;
      try {
        dados = JSON.parse(xhr.responseText);
      } catch (e) {
        exibirToast("erro", "Erro", "Resposta inválida do servidor.");
        return;
      }
      if (dados.erro) {
        exibirToast("erro", "Erro", dados.erro);
        return;
      }
      dadosCompletos = Array.isArray(dados) ? dados : [];
      dadosFiltrados = dadosCompletos.slice();
      paginaAtual = 1;
      $el("txtBusca").value = "";

      const temDados = dadosCompletos.length > 0;
      $el("areaPlaceholder").classList.toggle("d-none", temDados);
      $el("areaTabela").classList.toggle("d-none", !temDados);
      $el("btnExcel").disabled = !temDados;
      $el("btnPdf").disabled = !temDados;

      if (temDados) {
        renderTabela();
        exibirToast(
          "sucesso",
          "Sucesso",
          dadosCompletos.length + " registro(s) encontrado(s).",
        );
      } else {
        exibirToast(
          "info",
          "Sem resultados",
          "Nenhum registro para esse filtro.",
        );
      }
    };
    xhr.send(params);
  }

  /* ── Export Excel ───────────────────────────────────────────────── */
  function exportarExcel() {
    const planta =
      $el("cboPlanta").options[$el("cboPlanta").selectedIndex].text;
    const faca = $el("txtFaca").value.trim() || "Todos";
    const di = $el("dtInicial").value;
    const df = $el("dtFinal").value;

    const headers = [
      "Item",
      "Data",
      "Máquina",
      "Clichê",
      "Turma",
      "Qtde. Prog.",
    ];
    const rows = dadosFiltrados.map((r) => [
      r.item || "",
      r.inicio || "",
      r.maquina ? r.maquina + " - " + (r.nomeMaq || "") : r.nomeMaq || "",
      r.codFI || "",
      r.turma || "",
      Number(r.qtdeProg) || 0,
    ]);
    rows.push(["", "", "", "", "Total", somaTotal(dadosFiltrados)]);

    const ws = XLSX.utils.aoa_to_sheet([
      ["Relatório Matriz por Período"],
      [`Filial: ${planta} | Faca/Matriz: ${faca} | Período: ${di} a ${df}`],
      [],
      headers,
      ...rows,
    ]);
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Matriz por Período");
    XLSX.writeFile(wb, "RelatorioMatrizPorPeriodo.xlsx");
  }

  /* ── Export PDF ─────────────────────────────────────────────────── */
  function exportarPdf() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });
    const pageW = doc.internal.pageSize.getWidth();

    const planta =
      $el("cboPlanta").options[$el("cboPlanta").selectedIndex].text;
    const faca = $el("txtFaca").value.trim() || "Todos";
    const di = $el("dtInicial").value;
    const df = $el("dtFinal").value;
    const now = new Date().toLocaleString("pt-BR");
    const login = ($el("swnetLogin") ? $el("swnetLogin").value : "") || "N/D";

    function fmtData(iso) {
      if (!iso) return "";
      const p = iso.split("-");
      return p.length === 3 ? p[2] + "/" + p[1] + "/" + p[0] : iso;
    }
    const diFormatado = fmtData(di);
    const dfFormatado = fmtData(df);

    const rows = dadosFiltrados.map((r) => [
      r.item || "",
      r.inicio || "",
      r.maquina ? r.maquina + " - " + (r.nomeMaq || "") : r.nomeMaq || "",
      r.codFI || "",
      r.turma || "",
      fmtNum(r.qtdeProg),
    ]);
    rows.push(["", "", "", "", "Total", fmtNum(somaTotal(dadosFiltrados))]);

    function gerarCorpo(logoImg) {
      if (logoImg) {
        const ratio = logoImg.naturalHeight / logoImg.naturalWidth;
        const canvas = document.createElement("canvas");
        canvas.width = 900;
        canvas.height = Math.round(900 * ratio);
        canvas
          .getContext("2d")
          .drawImage(logoImg, 0, 0, canvas.width, canvas.height);
        const logoH = Math.round(38 * ratio);
        doc.addImage(canvas.toDataURL("image/png"), "PNG", 10, 5, 38, logoH);
      }

      doc.setFontSize(7);
      doc.setTextColor(100);
      doc.text(
        "Emitido em: " + now + "   |   Emitido por: " + login,
        pageW - 10,
        8,
        { align: "right" },
      );

      doc.setFontSize(14);
      doc.setTextColor(30);
      doc.text("Relatório Matriz por Período", pageW / 2, 14, {
        align: "center",
      });

      doc.setFontSize(8);
      doc.setTextColor(80);
      doc.text(
        "Filial: " +
          planta +
          "   |   Faca/Matriz: " +
          faca +
          "   |   Período: " +
          diFormatado +
          " a " +
          dfFormatado,
        pageW / 2,
        20,
        { align: "center" },
      );
      doc.text(dadosFiltrados.length + " registro(s)", pageW / 2, 25, {
        align: "center",
      });

      doc.autoTable({
        startY: 28,
        head: [["Item", "Data", "Máquina", "Clichê", "Turma", "Qtde. Prog."]],
        body: rows,
        theme: "striped",
        headStyles: {
          fillColor: [74, 74, 74],
          textColor: [255, 255, 255],
          fontSize: 4.5,
          cellPadding: 1.5,
          halign: "center",
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        styles: {
          fontSize: 4.5,
          cellPadding: 1.5,
          lineColor: [200, 200, 200],
          lineWidth: 0.15,
          valign: "middle",
          halign: "center",
        },
        columnStyles: {
          2: { halign: "center" },
          5: { halign: "center" },
        },
        didDrawPage: function (data) {
          const pg = doc.internal.getCurrentPageInfo().pageNumber;
          const total = doc.internal.getNumberOfPages();
          doc.setFontSize(6);
          doc.setTextColor(150);
          doc.text(
            "Página " + pg + " / " + total,
            pageW / 2,
            doc.internal.pageSize.getHeight() - 4,
            { align: "center" },
          );
        },
      });

      doc.save("RelatorioMatrizPorPeriodo.pdf");
    }

    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = function () {
      gerarCorpo(img);
    };
    img.onerror = function () {
      gerarCorpo(null);
    };
    img.src = LOGO_PATH;
  }

  /* ── Event listeners ────────────────────────────────────────────── */
  $el("btnFiltrar").addEventListener("click", filtrar);
  $el("btnExcel").addEventListener("click", exportarExcel);
  $el("btnPdf").addEventListener("click", exportarPdf);

  $el("btnLimpar").addEventListener("click", function () {
    $el("cboPlanta").value = "";
    $el("txtFaca").value = "";
    $el("dtInicial").value = "";
    $el("dtFinal").value = "";
    $el("txtBusca").value = "";
    dadosCompletos = [];
    dadosFiltrados = [];
    $el("areaTabela").classList.add("d-none");
    $el("areaPlaceholder").classList.remove("d-none");
    $el("btnExcel").disabled = true;
    $el("btnPdf").disabled = true;
  });

  $el("selectLimite").addEventListener("change", function () {
    paginaAtual = 1;
    renderTabela();
  });

  $el("txtBusca").addEventListener("input", function () {
    aplicarBusca(this.value);
  });

  document.addEventListener("keydown", function (e) {
    if (e.ctrlKey && e.key === "k") {
      e.preventDefault();
      $el("txtBusca").focus();
    }
  });
})();

</script>
</body>
</html>
