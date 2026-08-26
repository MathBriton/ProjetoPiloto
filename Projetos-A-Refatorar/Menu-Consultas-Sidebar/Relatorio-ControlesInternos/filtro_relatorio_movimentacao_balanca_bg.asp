<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Buffer = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

Dim IdUsuarioLogado, LoginUsuarioLogado
IdUsuarioLogado    = Session("CODIGO")
LoginUsuarioLogado = Session("USR")
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet — Relatório de Movimentação de Balança — Bento Gonçalves</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }

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

    .filter-card-body {
      padding: 1rem 1.25rem 1.25rem;
    }

    .form-label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }

    .input-data {
      width: 155px;
    }

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }

    .table-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3.5rem 1rem;
      color: #6c757d;
      gap: .75rem;
    }

    .table-placeholder .bi {
      font-size: 2.8rem;
      opacity: .35;
    }

    .table-placeholder p {
      margin: 0;
      font-size: .9rem;
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
      font-size: .82rem;
      font-weight: 700;
      text-align: center;
      white-space: nowrap;
      z-index: 2;
      border-color: #dee2e6;
      padding: .55rem .75rem;
    }

    .table tbody td {
      font-size: .82rem;
      text-align: center;
      white-space: nowrap;
      color: #212529;
      vertical-align: middle;
      padding: .45rem .75rem;
    }

    .table-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: .6rem 1rem;
      border-top: 1px solid #dee2e6;
      flex-wrap: wrap;
      gap: .5rem;
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
      font-size: .8rem;
    }

    .select-limite {
      width: auto;
    }

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

    .overlay-loading .spinner-border {
      width: 3rem;
      height: 3rem;
    }

    .overlay-loading span {
      font-size: .95rem;
      font-weight: 600;
    }

    .search-box {
      position: relative;
      max-width: 180px;
    }

    .search-box .form-control {
      padding-right: 72px;
      font-size: .82rem;
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
      font-size: .65rem;
      padding: .1rem .3rem;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      color: #6c757d;
    }

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

    .toast-custom .toast-inner {
      display: flex;
      align-items: flex-start;
      padding: .65rem .85rem .5rem;
      gap: .65rem;
    }

    .toast-custom .toast-icon {
      font-size: 1.2rem;
      flex-shrink: 0;
      margin-top: .05rem;
    }

    .toast-custom .toast-body-text {
      flex: 1;
    }

    .toast-custom .toast-title {
      font-size: .78rem;
      font-weight: 700;
      margin-bottom: .1rem;
    }

    .toast-custom .toast-msg {
      font-size: .82rem;
      color: #495057;
      line-height: 1.35;
    }

    .toast-custom .toast-close {
      background: none;
      border: none;
      font-size: 1rem;
      line-height: 1;
      color: #adb5bd;
      cursor: pointer;
      padding: 0;
      flex-shrink: 0;
    }

    .toast-custom .toast-close:hover { color: #495057; }

    .toast-custom .toast-progress {
      height: 3px;
      width: 100%;
      border-radius: 0 0 10px 10px;
      overflow: hidden;
      background: #e9ecef;
    }

    .toast-custom .toast-progress-bar {
      height: 100%;
      width: 100%;
      animation: toastProgressShrink 4.5s linear forwards;
      border-radius: 0 0 10px 10px;
    }

    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    .toast-custom.toast-sucesso .toast-icon  { color: #198754; }
    .toast-custom.toast-sucesso .toast-title { color: #198754; }
    .toast-custom.toast-sucesso .toast-progress-bar { background: #198754; }

    .toast-custom.toast-aviso .toast-icon   { color: #856404; }
    .toast-custom.toast-aviso .toast-title  { color: #856404; }
    .toast-custom.toast-aviso .toast-progress-bar { background: #ffc107; }

    .toast-custom.toast-erro .toast-icon   { color: #dc3545; }
    .toast-custom.toast-erro .toast-title  { color: #dc3545; }
    .toast-custom.toast-erro .toast-progress-bar { background: #dc3545; }

    .toast-custom.toast-info .toast-icon   { color: #6c757d; }
    .toast-custom.toast-info .toast-title  { color: #6c757d; }
    .toast-custom.toast-info .toast-progress-bar { background: #6c757d; }

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
  </style>
</head>
<body>

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border" role="status" aria-hidden="true"></div>
    <span>Carregando...</span>
  </div>

  <div id="toastMsg" class="toast-container-custom"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <h5 class="page-title mb-3">
      Relatório de Movimentação da Balança — Bento Gonçalves
    </h5>

    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtro
      </div>
      <div class="filter-card-body">
        <div class="row g-3 align-items-end">
          <div class="col-auto">
            <label class="form-label" for="datainicial">Data Inicial</label>
            <input type="date" id="datainicial" name="datainicial" class="form-control input-data" autocomplete="off" />
          </div>
          <div class="col-auto">
            <label class="form-label" for="datafinal">Data Final</label>
            <input type="date" id="datafinal" name="datafinal" class="form-control input-data" autocomplete="off" />
          </div>
          <div class="col">
            <div class="d-flex gap-2 justify-content-end">
              <button type="button" id="btnFiltrar" class="btn btn-primary" title="Filtrar registros">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" id="btnLimpar" class="btn btn-light" title="Limpar filtros">
                <i class="bi bi-eraser"></i>
              </button>
              <button type="button" id="btnExcel" class="btn btn-success" title="Exportar para Excel">
                <i class="bi bi-file-earmark-excel"></i>
              </button>
              <button type="button" id="btnPdf" class="btn btn-danger" title="Exportar para PDF">
                <i class="bi bi-file-earmark-pdf"></i>
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="table-card">
      <div class="p-0">

        <div id="placeholderTabela" class="table-placeholder">
          <i class="bi bi-clipboard-data"></i>
          <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
        </div>

        <div id="areaTabela" class="d-none">

          <div class="d-flex align-items-center justify-content-end px-3 pt-3 pb-2 border-bottom">
            <div class="search-box">
              <input type="text" id="inputBusca" class="form-control form-control-sm" placeholder="Buscar..." autocomplete="off" />
              <span class="search-hotkey"><kbd>CTRL+K</kbd></span>
            </div>
          </div>

          <div class="table-wrapper px-0">
            <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaDados">
              <thead>
                <tr>
                  <th>Ticket</th>
                  <th>Data da Criação</th>
                  <th>Placa do Veículo</th>
                  <th>Placa da Carreta</th>
                  <th>Transportadora</th>
                  <th>Cód.Emissor</th>
                  <th>Nome do Emissor</th>
                  <th>Cód.Item</th>
                  <th>Descrição do Item</th>
                  <th>Campo Adicional</th>
                  <th>Nota Fiscal</th>
                  <th>Peso Inicial</th>
                  <th>Data Peso Inicial</th>
                  <th>Hora Peso Inicial</th>
                  <th>Peso Final</th>
                  <th>Data Peso Final</th>
                  <th>Hora Peso Final</th>
                  <th>Peso Líquido</th>
                </tr>
              </thead>
              <tbody id="tbodyDados"></tbody>
            </table>
          </div>

          <div class="table-footer">
            <div class="table-footer-left">
              <div class="d-flex align-items-center gap-2">
                <label class="form-label mb-0 me-1" for="selectLimite">Exibir</label>
                <select id="selectLimite" class="form-select form-select-sm select-limite">
                  <option value="10">10</option>
                  <option value="25">25</option>
                  <option value="50" selected>50</option>
                  <option value="100">100</option>
                  <option value="250">250</option>
                  <option value="500">500</option>
                </select>
                <span class="form-label mb-0">por página</span>
              </div>
            </div>
            <div class="table-footer-center">
              <nav aria-label="Paginação">
                <ul class="pagination pagination-sm mb-0" id="paginacao"></ul>
              </nav>
            </div>
            <div class="table-footer-right">
              <span id="infoRegistros" class="text-muted info-registros"></span>
            </div>
          </div>

        </div>
      </div>
    </div>

  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

  <script>
    (() => {
  "use strict";

  const REPOSITORIO = "RelatorioMovimentacaoBalancaBGRepositorio.asp";
  const CAMINHO_LOGO = "/corporativo/balanca/bg_v2/Logo/logo-smurfit1.png";
  const ID_USUARIO = "<%=IdUsuarioLogado%>";
  const LOGIN_USUARIO = "<%=LoginUsuarioLogado%>";
  const TOAST_DELAY = 4500;

  const COLS = Object.freeze([
    "Ticket",
    "Data da Criação",
    "Placa do veículo",
    "Placa da Carreta",
    "Transportadora",
    "Cód.Emissor",
    "Nome do Emissor",
    "Cód.Item",
    "Descrição do Item",
    "Campo Adicional",
    "Nota Fiscal",
    "Peso Inicial",
    "Data Peso Inicial",
    "Hora Peso Inicial",
    "Peso Final",
    "Data Peso Final",
    "Hora Peso Final",
    "Peso Líquido",
  ]);

  const KEYS = Object.freeze([
    "ticket",
    "dataCriacao",
    "placaVei",
    "placaCar",
    "transp",
    "codEmi",
    "nomeEmi",
    "codItem",
    "descrItem",
    "campoAdic",
    "nf",
    "pesoIni",
    "dataPesoIni",
    "horaPesoIni",
    "pesoFin",
    "dataPesoFin",
    "horaPesoFin",
    "pesoLiq",
  ]);

  const TOAST_CFG = Object.freeze({
    sucesso: { icon: "bi-check-circle-fill", title: "Sucesso" },
    erro: { icon: "bi-x-circle-fill", title: "Erro" },
    aviso: { icon: "bi-exclamation-triangle-fill", title: "Atenção" },
    info: { icon: "bi-info-circle-fill", title: "Informação" },
  });

  let _todos = [];
  let _filtrados = [];
  let _paginaAtual = 1;
  let _limite = 50;
  let _periodoDE = "";
  let _periodoATE = "";

  const $id = (id) => document.getElementById(id);

  function exibirToast(msg, tipo) {
    const t = Object.hasOwn(TOAST_CFG, tipo) ? tipo : "info"; // ES2022: Object.hasOwn()
    const c = TOAST_CFG[t];
    const el = document.createElement("div");
    el.className = `toast-custom toast-${t}`;
    el.innerHTML = `
      <div class="toast-inner">
        <i class="bi ${c.icon} toast-icon"></i>
        <div class="toast-body-text">
          <div class="toast-title">${c.title}</div>
          <div class="toast-msg">${msg}</div>
        </div>
        <button class="toast-close" onclick="this.closest('.toast-custom').remove()" aria-label="Fechar">&times;</button>
      </div>
      <div class="toast-progress"><div class="toast-progress-bar"></div></div>`;
    $id("toastMsg").appendChild(el);
    setTimeout(() => el.isConnected && el.remove(), TOAST_DELAY); // ES2022: Element.isConnected
  }

  function exibirOverlay() {
    $id("overlayLoading").classList.remove("d-none");
  }
  function ocultarOverlay() {
    $id("overlayLoading").classList.add("d-none");
  }

  function validar() {
    const di = $id("datainicial").value;
    const df = $id("datafinal").value;
    if (!di) {
      exibirToast("Informe a data inicial.", "aviso");
      return false;
    }
    if (!df) {
      exibirToast("Informe a data final.", "aviso");
      return false;
    }
    if (new Date(di) > new Date(df)) {
      exibirToast("A data inicial deve ser menor que a data final.", "aviso");
      return false;
    }
    const diff = (new Date(df) - new Date(di)) / (1000 * 60 * 60 * 24);
    if (diff + 1 > 15) {
      exibirToast("O período deve ser de no máximo 15 dias.", "aviso");
      return false;
    }
    return true;
  }

  async function filtrar() {
    // ES2022: async/await com Error.cause
    if (!validar()) return;

    const di = $id("datainicial").value;
    const df = $id("datafinal").value;

    exibirOverlay();

    try {
      const res = await fetch(REPOSITORIO, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body:
          "acao=FILTRAR&datainicial=" +
          encodeURIComponent(di) +
          "&datafinal=" +
          encodeURIComponent(df),
      });

      if (!res.ok) {
        throw new Error(`Status HTTP: ${res.status}`, { cause: res }); // ES2022: Error.cause
      }

      const dados = await res.json();

      if (Object.hasOwn(dados, "erro")) {
        // ES2022: Object.hasOwn()
        exibirToast(dados.erro, "erro");
        return;
      }

      if (dados.length === 0) {
        exibirToast(
          "Nenhum registro encontrado para o período informado.",
          "aviso",
        );
        $id("placeholderTabela").classList.remove("d-none");
        $id("areaTabela").classList.add("d-none");
        return;
      }

      _periodoDE = di;
      _periodoATE = df;
      _todos = dados;
      _filtrados = structuredClone(dados);
      $id("inputBusca").value = "";
      exibirTabela();
    } catch (e) {
      const isComunicacao = !(e.cause instanceof Response);
      exibirToast(
        isComunicacao
          ? "Falha de comunicação com o servidor."
          : `Erro ao consultar os dados. ${e.message}`,
        "erro",
      );
    } finally {
      ocultarOverlay();
    }
  }

  function limpar() {
    $id("datainicial").value = "";
    $id("datafinal").value = "";
    $id("inputBusca").value = "";
    _todos = [];
    _filtrados = [];
    _periodoDE = "";
    _periodoATE = "";
    $id("placeholderTabela").classList.remove("d-none");
    $id("areaTabela").classList.add("d-none");
  }

  function construirLinha(r) {
    const tr = document.createElement("tr");
    for (const k of KEYS) {
      // ES2022: for...of em vez de forEach para melhor perf em loops grandes
      const td = document.createElement("td");
      td.textContent = r[k] ?? "";
      tr.appendChild(td);
    }
    return tr;
  }

  function renderizarPagina() {
    const tbody = $id("tbodyDados");
    tbody.innerHTML = "";
    const inicio = (_paginaAtual - 1) * _limite;
    const fim = Math.min(inicio + _limite, _filtrados.length);
    const frag = document.createDocumentFragment();
    for (let i = inicio; i < fim; i++) {
      frag.appendChild(construirLinha(_filtrados.at(i))); // ES2022: Array.at()
    }
    tbody.appendChild(frag);
    atualizarInfo(inicio, fim);
    renderizarPaginacao();
  }

  function atualizarInfo(inicio, fim) {
    const total = _filtrados.length;
    $id("infoRegistros").textContent =
      total === 0
        ? "Nenhum registro encontrado."
        : `Exibindo ${inicio + 1}–${fim} de ${total} registros`;
  }

  function renderizarPaginacao() {
    const totalPags = Math.ceil(_filtrados.length / _limite);
    const ul = $id("paginacao");
    ul.innerHTML = "";

    const criarItem = (label, pg, desabilitado, ativo) => {
      const li = document.createElement("li");
      li.className =
        "page-item" +
        (desabilitado ? " disabled" : "") +
        (ativo ? " active" : "");
      const a = document.createElement("a");
      a.className = "page-link";
      a.href = "#";
      a.innerHTML = label;
      a.addEventListener("click", (e) => {
        e.preventDefault();
        if (!desabilitado && !ativo) {
          _paginaAtual = pg;
          renderizarPagina();
        }
      });
      li.appendChild(a);
      return li;
    };

    ul.appendChild(criarItem("&laquo;", 1, _paginaAtual === 1, false));
    ul.appendChild(
      criarItem("&lsaquo;", _paginaAtual - 1, _paginaAtual === 1, false),
    );

    let inicio = Math.max(1, _paginaAtual - 2);
    let fim = Math.min(totalPags, inicio + 4);
    if (fim - inicio < 4) inicio = Math.max(1, fim - 4);

    for (let p = inicio; p <= fim; p++) {
      ul.appendChild(criarItem(p, p, false, p === _paginaAtual));
    }

    ul.appendChild(
      criarItem(
        "&rsaquo;",
        _paginaAtual + 1,
        _paginaAtual === totalPags || totalPags === 0,
        false,
      ),
    );
    ul.appendChild(
      criarItem(
        "&raquo;",
        totalPags,
        _paginaAtual === totalPags || totalPags === 0,
        false,
      ),
    );
  }

  function exibirTabela() {
    _paginaAtual = 1;
    $id("placeholderTabela").classList.add("d-none");
    $id("areaTabela").classList.remove("d-none");
    renderizarPagina();
  }

  function aplicarBusca() {
    const termo = $id("inputBusca").value.trim().toLowerCase();
    _filtrados =
      termo === ""
        ? structuredClone(_todos)
        : _todos.filter((r) =>
            KEYS.some((k) =>
              String(r[k] ?? "")
                .toLowerCase()
                .includes(termo),
            ),
          );
    _paginaAtual = 1;
    renderizarPagina();
  }

  function exportarExcel() {
    if (_filtrados.length === 0) {
      exibirToast("Nenhum dado para exportar.", "aviso");
      return;
    }
    const linhas = _filtrados.map((r) => KEYS.map((k) => r[k] ?? ""));
    const ws = XLSX.utils.aoa_to_sheet([COLS, ...linhas]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Movimentação Balança");
    XLSX.writeFile(wb, "relatorio-balanca-bento-goncalves.xlsx");
  }

  function _gerarAutoTable(doc, pageW) {
    const titulo = "Relatório de Movimentação da Balança — Bento Gonçalves";
    const subtitulo = `Período: ${_periodoDE} a ${_periodoATE}   |   Total: ${_filtrados.length} registros`;
    const emitido = `Emitido por: ${LOGIN_USUARIO} (ID: ${ID_USUARIO})`;

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(33, 37, 41);
    doc.text(titulo, pageW / 2, 14, { align: "center" });

    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(100, 100, 100);
    doc.text(subtitulo, pageW / 2, 20, { align: "center" });

    doc.setFontSize(7);
    doc.setTextColor(130, 130, 130);
    doc.text(emitido, pageW - 10, 8, { align: "right" });

    doc.autoTable({
      head: [COLS],
      body: _filtrados.map((r) => KEYS.map((k) => r[k] ?? "")),
      startY: 26,
      theme: "grid",
      styles: {
        fontSize: 4.5,
        cellPadding: 1.5,
        halign: "center",
        lineColor: [180, 180, 180],
        lineWidth: 0.2,
      },
      headStyles: {
        fillColor: [74, 74, 74],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [74, 74, 74],
        lineWidth: 0.2,
      },
      alternateRowStyles: { fillColor: [245, 245, 245] },
    });

    doc.save("relatorio-balanca-bento-goncalves.pdf");
    ocultarOverlay();
  }

  async function exportarPdf() {
    // ES2022: async/await com Error.cause
    if (_filtrados.length === 0) {
      exibirToast("Nenhum dado para exportar.", "aviso");
      return;
    }
    exibirOverlay();

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });
    const pageW = doc.internal.pageSize.getWidth();

    try {
      const dataUrl = await _carregarLogoDataUrl(CAMINHO_LOGO); // ES2022: await direto no topo do try
      const img = await _imagemFromDataUrl(dataUrl);
      const logoDispW = 38;
      const logoDispH = logoDispW * (img.naturalHeight / img.naturalWidth);
      const canvas = Object.assign(document.createElement("canvas"), {
        // ES2022: Object.assign idiom
        width: 900,
        height: Math.round((img.naturalHeight / img.naturalWidth) * 900),
      });
      canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
      doc.addImage(
        canvas.toDataURL("image/png"),
        "PNG",
        10,
        4,
        logoDispW,
        logoDispH,
      );
    } catch {
      // ES2022: catch sem binding (omissão do parâmetro `e`)
    } finally {
      _gerarAutoTable(doc, pageW);
    }
  }

  function _carregarLogoDataUrl(src) {
    // helper para exportarPdf
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => resolve(src);
      img.onerror = () =>
        reject(new Error("Logo não carregou", { cause: src })); // ES2022: Error.cause
      img.src = src;
    });
  }

  function _imagemFromDataUrl(src) {
    // helper para exportarPdf
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () =>
        reject(new Error("Erro ao criar imagem", { cause: src })); // ES2022: Error.cause
      img.src = src;
    });
  }

  function init() {
    $id("btnFiltrar").addEventListener("click", filtrar);
    $id("btnLimpar").addEventListener("click", limpar);
    $id("btnExcel").addEventListener("click", exportarExcel);
    $id("btnPdf").addEventListener("click", exportarPdf);

    $id("selectLimite").addEventListener("change", () => {
      _limite = parseInt($id("selectLimite").value, 10);
      _paginaAtual = 1;
      renderizarPagina();
    });

    $id("inputBusca").addEventListener("input", aplicarBusca);

    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "k") {
        e.preventDefault();
        $id("inputBusca").focus();
        $id("inputBusca").select();
      }
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();

  </script>

</body>
</html>
