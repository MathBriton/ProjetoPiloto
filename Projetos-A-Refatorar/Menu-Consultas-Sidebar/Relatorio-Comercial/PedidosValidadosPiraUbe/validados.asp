<%@ Language="VBScript" %>
<!--#Include File="../../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../../cnn2/conexoes.asp"-->
<!--#Include File="../../../../includes/variaveis.asp"-->
<!--#Include File="../../../../funcoes/funcoes.asp"-->
<%
  Session.LCID = 1046

  Dim swnet_nivel   : swnet_nivel   = CInt(nivel)
  Dim swnet_coduser : swnet_coduser = CStr(CODUSER)
  Dim swnet_usr     : swnet_usr     = CStr(USR)
  Dim swnet_helpdesk : swnet_helpdesk = CStr(CODIGOHELPDESK)
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>Consulta Pedidos Validados - SWnet</title>

  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    /* ── Filtros ───────────────────────────────────────────────────────────── */
    .filter-card         { background:#fff; border:1px solid #dee2e6; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,.08); }
    .filter-card-header  { background:#fff; color:#212529; padding:.6rem 1rem; border-radius:6px 6px 0 0; border-bottom:1px solid #dee2e6; font-weight:600; font-size:.9rem; }
    .filter-card-body    { padding:1rem 1.25rem 1.25rem; }
    .filter-label        { font-size:.82rem; font-weight:600; color:#6c757d; margin-bottom:.25rem; display:block; }
    .input-data          { width:155px; }
    .filter-divider      { width:1px; background-color:#dee2e6; align-self:stretch; margin:0 .25rem; }

    /* ── Tabela ────────────────────────────────────────────────────────────── */
    .table-card     { background:#fff; border:1px solid #dee2e6; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,.08); position:relative; }
    .table-toolbar  { background:#fff; border-radius:6px 6px 0 0; }
    .search-box     { position:relative; max-width:220px; }
    .search-box .form-control { padding-right:76px; font-size:.82rem; }
    .search-box .search-hotkey {
      position:absolute; right:8px; top:50%; transform:translateY(-50%);
      pointer-events:none; user-select:none; display:flex; gap:2px;
    }
    .search-box .search-hotkey kbd { font-size:.65rem; padding:.1rem .3rem; background:#f8f9fa; border:1px solid #dee2e6; border-radius:3px; color:#6c757d; }

    .table-wrapper            { overflow-x:auto; overflow-y:auto; max-height:68vh; }
    .table-wrapper table      { min-width:max-content; }
    .table thead th           { position:sticky; top:0; background:#fff; color:#212529; font-size:.82rem; font-weight:700; text-align:center; white-space:nowrap; z-index:2; border-color:#dee2e6; padding:.55rem .75rem; }
    .table tbody td           { font-size:.82rem; text-align:center; white-space:nowrap; color:#212529; vertical-align:middle; padding:.45rem .75rem; }

    /* ── Placeholder ───────────────────────────────────────────────────────── */
    .table-placeholder        { display:flex; flex-direction:column; align-items:center; justify-content:center; padding:3.5rem 1rem; color:#6c757d; gap:.75rem; }
    .table-placeholder .bi    { font-size:2.8rem; opacity:.35; }
    .table-placeholder p      { margin:0; font-size:.9rem; text-align:center; }

    /* ── Overlay ───────────────────────────────────────────────────────────── */
    .overlay-loading          { position:absolute; inset:0; background:rgba(255,255,255,.82); display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:10; gap:1rem; border-radius:6px; }
    .overlay-loading .spinner-border { width:2.5rem; height:2.5rem; color:#6c757d; }
    .overlay-loading span     { font-size:.9rem; font-weight:600; color:#495057; }

    /* ── Footer / Paginação ────────────────────────────────────────────────── */
    .table-footer             { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:.5rem; padding:.6rem .75rem; border-top:1px solid #dee2e6; background:#fafafa; border-radius:0 0 6px 6px; min-height:44px; }
    .table-footer-left,
    .table-footer-center,
    .table-footer-right       { display:flex; align-items:center; gap:.4rem; }
    .table-footer-center      { flex:1; justify-content:center; }
    .table-footer-right       { justify-content:flex-end; font-size:.82rem; color:#6c757d; white-space:nowrap; }
    .table-footer-left label  { font-size:.82rem; color:#6c757d; white-space:nowrap; }
    .select-limite            { width:auto; }
    .info-registros           { font-size:.8rem; color:#6c757d; }

    #paginacao .page-link                    { color:#6c757d; border-color:#dee2e6; }
    #paginacao .page-item.active .page-link  { background-color:#6c757d; border-color:#6c757d; color:#fff; }
    #paginacao .page-link:hover              { background-color:#e9ecef; color:#495057; }
    #paginacao .page-item.disabled .page-link { color:#adb5bd; }

    /* ── Toast ─────────────────────────────────────────────────────────────── */
    .toast-custom { min-width:280px; max-width:360px; border-radius:10px; overflow:hidden; box-shadow:0 4px 18px rgba(0,0,0,.13); background:#fff; border:1px solid #dee2e6; animation:toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
    @keyframes toastSlideIn { from{opacity:0;transform:translateX(40px)} to{opacity:1;transform:translateX(0)} }
    .toast-progress-bar { animation:toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from{width:100%} to{width:0%} }

    /* ── page-title ────────────────────────────────────────────────────────── */
    .page-title { font-size:1.15rem; font-weight:700; color:#212529; }

    @media print { .table-footer { display:none; } }
  </style>
</head>
<body>

  <!-- Toast container -->
  <div id="toastContainer"
       style="position:fixed;top:1rem;right:1rem;z-index:1100;display:flex;flex-direction:column;gap:.5rem;"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center mb-3">
      <span class="page-title">Consulta Pedidos Validados</span>
    </div>

    <!-- ===== CARD DE FILTROS ===== -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <div class="d-flex flex-wrap align-items-end gap-3">

          <div>
            <label class="filter-label" for="tipoDat">Tipo de data</label>
            <select id="tipoDat" class="form-select form-select-sm" style="width:120px;">
              <option value="E" selected>Entrega</option>
              <option value="M">Emissão</option>
            </select>
          </div>
          <!-- Datas -->
          <div>
            <label class="filter-label" for="datainicial">Data de</label>
            <input id="datainicial" type="date" class="form-control form-control-sm input-data" />
          </div>
          <div>
            <label class="filter-label" for="datafinal">Data até</label>
            <input id="datafinal" type="date" class="form-control form-control-sm input-data" />
          </div>

          <!-- Alternância Entrega / Emissão -->

          <div class="filter-divider d-none d-md-block"></div>

          <!-- Pedidos -->
          <div>
            <label class="filter-label" for="pedidoInicial">Pedido inicial</label>
            <input id="pedidoInicial" type="text" class="form-control form-control-sm"
                   style="width:140px;" placeholder="n° do pedido..."
                   oninput="this.value = this.value.toUpperCase();" />
          </div>
          <div>
            <label class="filter-label" for="pedidoFinal">Pedido final</label>
            <input id="pedidoFinal" type="text" class="form-control form-control-sm"
                   style="width:140px;" placeholder="n° do pedido..."
                   oninput="this.value = this.value.toUpperCase();" />
          </div>

          <div class="filter-divider d-none d-md-block"></div>

          <!-- Cliente -->
          <div>
            <label class="filter-label" for="codcli">Cliente</label>
            <div class="input-group input-group-sm">
              <button type="button" class="btn btn-secondary" id="btnBuscarCliente" title="Buscar no modal">
                Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
              </button>

              <input id="codcli" type="text" class="form-control form-control-sm" style="min-width:260px;"
                     maxlength="50" autocomplete="off" placeholder="Código ou nome do cliente..."
                     oninput="this.value = this.value.toUpperCase();" />
              <input id="razao" type="hidden" value="" />

              <input id="lojacli" type="text" class="form-control form-control-sm" style="width:70px;"
                     maxlength="2" placeholder="Loja" />

              <button type="button" class="btn btn-outline-secondary" id="btnLimparCliente" title="Limpar cliente">
                <i class="bi bi-x-lg"></i>
              </button>
            </div>
          </div>

          <!-- Ações -->
          <div class="ms-auto">
            <label class="filter-label">&nbsp;</label>
            <div class="d-flex gap-2">
              <button type="button" class="btn btn-primary btn-sm" id="btnFiltrar">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" class="btn btn-light btn-sm" id="btnLimpar">
                <i class="bi bi-eraser me-2"></i>Limpar
              </button>
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- ===== CARD DE TABELA ===== -->
    <div class="table-card" id="tableCard">

      <!-- Overlay de carregamento -->
      <div id="overlayLoading" class="overlay-loading d-none">
        <div class="spinner-border text-secondary" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <!-- Toolbar (oculta até haver dados) -->
      <div id="tableToolbar" class="table-toolbar d-none d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          <button type="button" class="btn btn-success btn-sm" id="btnExcel">
            <i class="bi bi-file-excel me-1"></i>Exportar Excel
          </button>
          <button type="button" class="btn btn-danger btn-sm" id="btnPdf">
            <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
          </button>
        </div>
        <div class="search-box">
          <input id="filtroTabela" type="text" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <!-- Placeholder inicial -->
      <div id="tablePlaceholder" class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Preencha o período de datas ou o intervalo de pedidos e clique em <strong>Filtrar</strong>.</p>
      </div>

      <!-- Wrapper da tabela (oculto até haver dados) -->
      <div id="tableWrapper" class="table-wrapper d-none">
        <table class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead id="tableThead"></thead>
          <tbody id="tableTbody"></tbody>
        </table>
      </div>

      <!-- Footer / paginação -->
      <div class="table-footer d-none" id="tableFooter">
        <div class="table-footer-left">
          <label for="selectLimite">Exibir</label>
          <select id="selectLimite" class="form-select form-select-sm select-limite">
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="100">100</option>
            <option value="200">200</option>
            <option value="9999" selected>Todos</option>
          </select>
          <label>por página</label>
        </div>
        <div class="table-footer-center">
          <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
        </div>
        <div class="table-footer-right">
          <span id="infoRegistros" class="info-registros"></span>
        </div>
      </div>

    </div><!-- /.table-card -->
  </div><!-- /.container-fluid -->

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

  <script>
  // ── Variáveis de sessão injetadas pelo servidor ───────────────────────────
  window.SWNET_NIVEL    = <%= swnet_nivel %>;
  window.SWNET_CODUSER  = '<%= swnet_coduser %>';
  window.SWNET_USR      = '<%= swnet_usr %>';
  window.SWNET_HELPDESK = '<%= swnet_helpdesk %>';
  </script>

  <script>
  (function () {
  "use strict";

  // ── Constantes ────────────────────────────────────────────────────────
  const REPOSITORIO = "validadosRepositorio.asp";
  const CACHE_KEY_DADOS = "val_dados";
  const CACHE_KEY_FILTROS = "val_filtros";
  const CACHE_KEY_PAGINA = "val_pagina";
  const CACHE_KEY_LIMITE = "val_limite";
  const CACHE_FLAG_RETORNO = "val_retornando";

  // ── Estado ────────────────────────────────────────────────────────────
  let _dadosCompletos = [];
  let _filtrados = [];
  let _paginaAtual = 1;
  let _limite = 9999;
  let _tipoAtual = ""; // "D" ou "P"

  // ── Atalho ────────────────────────────────────────────────────────────
  const $id = (id) => document.getElementById(id);

  // ── Toast ─────────────────────────────────────────────────────────────
  const TOAST_DELAY = 4500;
  const TOAST_CFG = {
    sucesso: { icone: "bi-check-circle-fill", cor: "#198754" },
    erro: { icone: "bi-x-circle-fill", cor: "#dc3545" },
    aviso: {
      icone: "bi-exclamation-triangle-fill",
      cor: "#856404",
      barCor: "#ffc107",
    },
    info: { icone: "bi-info-circle-fill", cor: "#6c757d" },
  };

  function mostrarToast(mensagem, tipo) {
    tipo = tipo || "info";
    const cfg = TOAST_CFG[tipo] || TOAST_CFG.info;
    const barCor = cfg.barCor || cfg.cor;
    const el = document.createElement("div");
    el.className = "toast-custom";
    el.innerHTML =
      '<div style="display:flex;align-items:flex-start;padding:.75rem 1rem;gap:.6rem;">' +
      '<i class="bi ' +
      cfg.icone +
      '" style="color:' +
      cfg.cor +
      ';font-size:1rem;margin-top:.1rem;flex-shrink:0;"></i>' +
      '<span style="font-size:.87rem;color:#212529;flex:1;">' +
      mensagem +
      "</span>" +
      '<button type="button" style="background:none;border:none;font-size:1.1rem;line-height:1;color:#6c757d;cursor:pointer;padding:0;margin-left:.25rem;" ' +
      "onclick=\"this.closest('.toast-custom').remove()\">&times;</button>" +
      "</div>" +
      '<div style="height:3px;background:#f0f0f0;">' +
      '<div class="toast-progress-bar" style="height:100%;background:' +
      barCor +
      ';"></div>' +
      "</div>";
    $id("toastContainer").appendChild(el);
    setTimeout(function () {
      if (el.parentNode) el.remove();
    }, TOAST_DELAY);
  }

  // ── Overlay ───────────────────────────────────────────────────────────
  function exibirOverlay() {
    $id("overlayLoading").classList.remove("d-none");
  }
  function ocultarOverlay() {
    $id("overlayLoading").classList.add("d-none");
  }

  // ── Formatadores ──────────────────────────────────────────────────────
  function fmtNum(valor, casas) {
    if (valor === null || valor === undefined || valor === "") return "-";
    const n = parseFloat(valor);
    if (isNaN(n)) return "-";
    return n.toLocaleString("pt-BR", {
      minimumFractionDigits: casas,
      maximumFractionDigits: casas,
    });
  }

  // ── Validação dos filtros ─────────────────────────────────────────────
  function validar() {
    const datini = $id("datainicial").value.trim();
    const datfin = $id("datafinal").value.trim();
    const pedini = $id("pedidoInicial").value.trim();
    const pedfin = $id("pedidoFinal").value.trim();

    const temData = datini.length > 0 || datfin.length > 0;
    const temPedido = pedini.length > 0 || pedfin.length > 0;

    if (temData && temPedido) {
      mostrarToast(
        "Preencha apenas um critério: OU período de datas OU intervalo de pedidos.",
        "aviso",
      );
      return null;
    }

    if (!temData && !temPedido) {
      mostrarToast(
        "Preencha o período de datas ou o intervalo de pedidos.",
        "aviso",
      );
      return null;
    }

    if (temData) {
      if (!datini) {
        mostrarToast("Digite a data inicial.", "aviso");
        return null;
      }
      if (!datfin) {
        mostrarToast("Digite a data final.", "aviso");
        return null;
      }
      return "D";
    }

    if (temPedido) {
      if (!pedini) {
        mostrarToast("Digite o pedido inicial.", "aviso");
        return null;
      }
      if (!pedfin) {
        mostrarToast("Digite o pedido final.", "aviso");
        return null;
      }
      return "P";
    }

    return null;
  }

  // ── Coleta parâmetros para o AJAX ─────────────────────────────────────
  function _coletarFiltros(tipo) {
    const params = {
      param: "1",
      tipo: tipo,
      codcli: $id("codcli").value.trim().split(" - ")[0].trim(),
      lojacli: $id("lojacli").value.trim(),
    };
    if (tipo === "D") {
      params.datainicial = $id("datainicial").value.trim();
      params.datafinal   = $id("datafinal").value.trim();
      params.tipoDat     = $id("tipoDat").value; // "E" = Entrega | "M" = Emissão
    } else {
      params.pedidoInicial = $id("pedidoInicial").value.trim();
      params.pedidoFinal   = $id("pedidoFinal").value.trim();
    }
    return params;
  }

  // ── Pesquisar ─────────────────────────────────────────────────────────
  function filtrar() {
    const tipo = validar();
    if (!tipo) return;

    _tipoAtual = tipo;
    _paginaAtual = 1;

    exibirOverlay();
    $("#btnFiltrar").prop("disabled", true);

    $.ajax({
      url: REPOSITORIO,
      method: "GET",
      dataType: "json",
      cache: false,
      data: _coletarFiltros(tipo),
      success: function (resp) {
        if (!resp.sucesso) {
          mostrarToast(resp.mensagem || "Erro ao carregar dados.", "erro");
          _mostrarPlaceholder();
          return;
        }
        if (!resp.dados || resp.dados.length === 0) {
          mostrarToast(
            "Nenhum registro encontrado para os filtros informados.",
            "aviso",
          );
          _mostrarPlaceholder();
          return;
        }
        _dadosCompletos = resp.dados;
        _filtrados = _dadosCompletos.slice();
        _renderThead();
        _aplicarBusca(); // respeita eventual termo já digitado
        _mostrarTabela();
      },
      error: function (xhr) {
        mostrarToast("Erro na comunicação com o servidor.", "erro");
        console.error(xhr.responseText);
        _mostrarPlaceholder();
      },
      complete: function () {
        ocultarOverlay();
        $("#btnFiltrar").prop("disabled", false);
      },
    });
  }

  // ── Limpar ────────────────────────────────────────────────────────────
  function limpar() {
    $id("datainicial").value = "";
    $id("datafinal").value = "";
    $id("tipoDat").value   = "E";
    $id("pedidoInicial").value = "";
    $id("pedidoFinal").value = "";
    $id("codcli").value  = "";
    $id("lojacli").value = "";
    $id("razao").value   = "";
    if ($id("filtroTabela")) $id("filtroTabela").value = "";

    _dadosCompletos = [];
    _filtrados = [];
    _tipoAtual = "";
    _paginaAtual = 1;

    sessionStorage.removeItem(CACHE_KEY_DADOS);
    sessionStorage.removeItem(CACHE_KEY_FILTROS);
    sessionStorage.removeItem(CACHE_KEY_PAGINA);
    sessionStorage.removeItem(CACHE_KEY_LIMITE);

    _mostrarPlaceholder();
  }

  // ── Cabeçalho da tabela (nivel controla colunas de preço) ─────────────
  function _renderThead() {
    const mostrarPreco = window.SWNET_NIVEL !== 2;
    let h = "";

    // Linha 1 — grupos
    h += "<tr>";
    h +=
      "<th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th>";
    h += '<th class="text-center">Condição de</th>';
    h += '<th class="text-center" colspan="2">Data Entrega</th>';
    h += "<th></th>";
    h += '<th class="text-center" colspan="2">Quantidade</th>';
    if (mostrarPreco) {
      h += "<th></th><th></th><th></th><th></th><th></th><th></th>";
    }
    h += '<th class="text-center">Último</th>';
    h += '<th class="text-center">Unidade</th>';
    h += '<th class="text-center">Quem</th>';
    h += '<th class="text-center">Quem</th>';
    h += "<th></th>";
    h += "</tr>";

    // Linha 2 — rótulos
    h += "<tr>";
    h += '<th class="text-center">Número</th>';
    h += '<th class="text-center">Emissão</th>';
    h += '<th class="text-center" style="width:150px">Ped.Cli-Item</th>';
    h += "<th>Status</th>";
    h += '<th style="width:150px">Cliente</th>';
    h += '<th style="width:150px">Loja</th>';
    h += "<th>Produto</th>";
    h += '<th style="width:500px">Referência</th>';
    h += "<th>TES</th>";
    h += '<th class="text-center">Pagto.</th>';
    h += '<th class="text-center">1ª</th>';
    h += '<th class="text-center">2ª</th>';
    h += '<th class="text-center">Qtde.</th>';
    h += '<th class="text-center">Original</th>';
    h += '<th class="text-center">Entregue</th>';
    if (mostrarPreco) {
      h += '<th class="text-center">Preço</th>';
      h += '<th class="text-center">Preço(Kg)</th>';
      h += '<th class="text-center">KG NET</th>';
      h += '<th class="text-center">Área(m²)</th>';
      h += '<th class="text-center">Saldo</th>';
      h += '<th class="text-center">Estrutura</th>';
    }
    h += '<th class="text-center">Faturamento</th>';
    h += '<th class="text-center">Fabril</th>';
    h += '<th class="text-center">Incluiu</th>';
    h += '<th class="text-center">Validou</th>';
    h += '<th class="text-center">Piloto</th>';
    h += "</tr>";

    $id("tableThead").innerHTML = h;
  }

  // ── Renderização de uma página de dados ───────────────────────────────
  function _renderTbody() {
    const mostrarPreco = window.SWNET_NIVEL !== 2;
    const inicio = (_paginaAtual - 1) * _limite;
    const fim = Math.min(inicio + _limite, _filtrados.length);
    const pagina = _filtrados.slice(inicio, fim);

    if (pagina.length === 0) {
      $id("tableTbody").innerHTML =
        '<tr><td colspan="99" class="text-center text-muted py-3">Nenhum registro encontrado.</td></tr>';
      _atualizarInfo(0, 0);
      _renderPaginacao();
      return;
    }

    let h = "";
    $.each(pagina, function (i, item) {
      // link detalhe
      const lnk =
        "detalha_nf.asp" +
        "?Planta=00" +
        "&Pedido=" +
        encodeURIComponent(item.C6_NUM || "") +
        "&tipo=" +
        encodeURIComponent(_tipoAtual) +
        "&codcli=" +
        encodeURIComponent($id("codcli").value.split(" - ")[0].trim()) +
        "&data1=" +
        encodeURIComponent($id("datainicial").value) +
        "&data2=" +
        encodeURIComponent($id("datafinal").value) +
        "&pedini=" +
        encodeURIComponent($id("pedidoInicial").value) +
        "&pedfim=" +
        encodeURIComponent($id("pedidoFinal").value) +
        "&kgnet=" +
        encodeURIComponent(item.KG_NET || 0);

      h += "<tr>";
      h +=
        '<td><a href="' +
        lnk +
        '" class="link-detalhe">' +
        (item.C6_NUM || "-") +
        "</a></td>";
      h += "<td>" + (item.C5_EMISSAO || "-") + "</td>";
      h += "<td>" + (item.PED_CLI_ITEM || "-") + "</td>";
      h += "<td>" + (item.C5_ESTATUS || "-") + "</td>";
      h += "<td>" + (item.CLIENTE || "-") + "</td>";
      h += "<td>" + (item.LOJA_MUN_EST || "-") + "</td>";
      h += "<td>" + (item.C6_PRODUTO || "-") + "</td>";
      h += "<td>" + (item.DESCRICAO || "-") + "</td>";
      h += "<td>" + (item.C6_TES || "-") + "</td>";
      h += "<td>" + (item.CONDICAO_PAGAMENTO || "-") + "</td>";
      h += "<td>" + (item.DATAENTREG1 || "-") + "</td>";
      h += "<td>" + (item.DATAENTREG2 || "-") + "</td>";
      h += '<td class="text-end">' + fmtNum(item.C6_QTDVEN, 0) + "</td>";
      h += '<td class="text-end">' + fmtNum(item.C6_QTDORIG, 0) + "</td>";
      h += '<td class="text-end">' + fmtNum(item.C6_QTDENT, 0) + "</td>";
      if (mostrarPreco) {
        h += '<td class="text-end">' + fmtNum(item.C6_PRCVEN, 5) + "</td>";
        h += '<td class="text-end">' + fmtNum(item.PRECO_KG, 5) + "</td>";
        h += '<td class="text-end">' + fmtNum(item.KG_NET, 5) + "</td>";
        h += '<td class="text-end">' + fmtNum(item.AREA, 5) + "</td>";
        h += '<td class="text-end">' + fmtNum(item.SALDO, 0) + "</td>";
        h += "<td>" + (item.ESTRUTURA || "-") + "</td>";
      }
      h += "<td>" + (item.D2 || "-") + "</td>";
      h += "<td>" + (item.FABRICA || "-") + "</td>";
      h += "<td>" + (item.INCLUIU || "-") + "</td>";
      h += "<td>" + (item.VALIDOU || "-") + "</td>";
      h += "<td>" + (item.C5_LOTEPIL || "-") + "</td>";
      h += "</tr>";
    });

    $id("tableTbody").innerHTML = h;
    _atualizarInfo(inicio, fim);
    _renderPaginacao();
  }

  // ── Busca rápida na tabela ────────────────────────────────────────────
  function _aplicarBusca() {
    const termo = ($id("filtroTabela") ? $id("filtroTabela").value : "")
      .toLowerCase()
      .trim();
    _filtrados =
      termo === ""
        ? _dadosCompletos.slice()
        : _dadosCompletos.filter(function (item) {
            return JSON.stringify(item).toLowerCase().includes(termo);
          });
    _paginaAtual = 1;
    _renderTbody();
  }

  // ── Paginação ─────────────────────────────────────────────────────────
  function _renderPaginacao() {
    const container = $id("paginacao");
    if (!container) return;
    container.innerHTML = "";

    const total = Math.ceil(_filtrados.length / _limite);
    if (total <= 1) return;

    const criarItem = function (label, pagina, ativo, desabilitado) {
      const li = document.createElement("li");
      li.className =
        "page-item" +
        (ativo ? " active" : "") +
        (desabilitado ? " disabled" : "");
      const a = document.createElement("a");
      a.className = "page-link";
      a.href = "#";
      a.innerHTML = label;
      if (!ativo && !desabilitado) {
        a.addEventListener("click", function (e) {
          e.preventDefault();
          _paginaAtual = pagina;
          _renderTbody();
        });
      } else {
        a.addEventListener("click", function (e) {
          e.preventDefault();
        });
      }
      li.appendChild(a);
      return li;
    };

    container.appendChild(criarItem("&laquo;", 1, false, _paginaAtual === 1));
    container.appendChild(
      criarItem("&lsaquo;", _paginaAtual - 1, false, _paginaAtual === 1),
    );

    const ini = Math.max(1, _paginaAtual - 2);
    const fim = Math.min(total, _paginaAtual + 2);
    for (let p = ini; p <= fim; p++) {
      container.appendChild(criarItem(p, p, p === _paginaAtual, false));
    }

    container.appendChild(
      criarItem("&rsaquo;", _paginaAtual + 1, false, _paginaAtual === total),
    );
    container.appendChild(
      criarItem("&raquo;", total, false, _paginaAtual === total),
    );
  }

  function _atualizarInfo(inicio, fim) {
    const el = $id("infoRegistros");
    if (!el) return;
    if (_filtrados.length === 0) {
      el.textContent = "Nenhum registro";
      return;
    }
    el.textContent =
      inicio + 1 + "–" + fim + " de " + _filtrados.length + " registros";
  }

  // ── Controle de visibilidade das seções ───────────────────────────────
  function _mostrarTabela() {
    $id("tablePlaceholder").classList.add("d-none");
    $id("tableToolbar").classList.remove("d-none");
    $id("tableWrapper").classList.remove("d-none");
    $id("tableFooter").classList.remove("d-none");
    _renderTbody();
  }

  function _mostrarPlaceholder() {
    $id("tablePlaceholder").classList.remove("d-none");
    $id("tableToolbar").classList.add("d-none");
    $id("tableWrapper").classList.add("d-none");
    $id("tableFooter").classList.add("d-none");
    $id("tableThead").innerHTML = "";
    $id("tableTbody").innerHTML = "";
    _dadosCompletos = [];
    _filtrados = [];
    _renderPaginacao();
    _atualizarInfo(0, 0);
  }

  // ── Export CSV ────────────────────────────────────────────────────────
  function exportarExcel() {
    if (_filtrados.length === 0) {
      mostrarToast("Nenhum dado disponível para exportar.", "aviso");
      return;
    }

    const mostrarPreco = window.SWNET_NIVEL !== 2;

    function escaparCSV(valor) {
      if (valor === null || valor === undefined) return '""';
      const str = String(valor).trim().replace(/"/g, '""');
      return '"' + str + '"';
    }

    // Colunas — mantém alinhamento com _renderTbody
    const cabecalho = [
      "Número",
      "Emissão",
      "Ped.Cli-Item",
      "Status",
      "Cliente",
      "Loja",
      "Produto",
      "Referência",
      "TES",
      "Pagto.",
      "1ª Entrega",
      "2ª Entrega",
      "Qtde.",
      "Original",
      "Entregue",
    ];
    const colunasData = [1, 10, 11]; // Emissão, 1ª, 2ª

    if (mostrarPreco) {
      cabecalho.push(
        "Preço",
        "Preço(Kg)",
        "KG NET",
        "Área(m²)",
        "Saldo",
        "Estrutura",
      );
    }
    cabecalho.push("Faturamento", "Fabril", "Incluiu", "Validou", "Piloto");
    if (mostrarPreco) colunasData.push(cabecalho.indexOf("Faturamento"));

    let linhas = [cabecalho.map(escaparCSV).join(";")];

    $.each(_filtrados, function (i, item) {
      const cols = [
        item.C6_NUM,
        item.C5_EMISSAO,
        item.PED_CLI_ITEM,
        item.C5_ESTATUS,
        item.CLIENTE,
        item.LOJA_MUN_EST,
        item.C6_PRODUTO,
        item.DESCRICAO,
        item.C6_TES,
        item.CONDICAO_PAGAMENTO,
        item.DATAENTREG1,
        item.DATAENTREG2,
        fmtNum(item.C6_QTDVEN, 0),
        fmtNum(item.C6_QTDORIG, 0),
        fmtNum(item.C6_QTDENT, 0),
      ];
      if (mostrarPreco) {
        cols.push(
          fmtNum(item.C6_PRCVEN, 5),
          fmtNum(item.PRECO_KG, 5),
          fmtNum(item.KG_NET, 5),
          fmtNum(item.AREA, 5),
          fmtNum(item.SALDO, 0),
          item.ESTRUTURA || "",
        );
      }
      cols.push(
        item.D2,
        item.FABRICA,
        item.INCLUIU,
        item.VALIDOU,
        item.C5_LOTEPIL,
      );

      const row = cols.map(function (val, idx) {
        if (
          colunasData.indexOf(idx) !== -1 &&
          /^\d{2}\/\d{2}\/\d{4}$/.test(String(val || ""))
        ) {
          return '="' + val + '"'; // força texto no Excel — elimina hashtag
        }
        return escaparCSV(val);
      });
      linhas.push(row.join(";"));
    });

    const csv = "\uFEFF" + linhas.join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pedidos_validados.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ── Export PDF ────────────────────────────────────────────────────────
  function exportarPDF() {
    if (_filtrados.length === 0) {
      mostrarToast("Nenhum dado disponível para exportar.", "aviso");
      return;
    }
    const img = new Image();
    img.onload = function () {
      const canvas = document.createElement("canvas");
      canvas.width = 900;
      canvas.height = 315;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, 900, 315);
      ctx.drawImage(img, 0, 0, 900, 315);
      _finalizarPDF(canvas.toDataURL("image/jpeg", 0.92));
    };
    img.onerror = function () {
      _finalizarPDF(null);
    };
    img.src = "Logo/logo-smurfit1.png";
  }

  function _finalizarPDF(imgData) {
    const mostrarPreco = window.SWNET_NIVEL !== 2;
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });
    const pgW = doc.internal.pageSize.getWidth();

    if (imgData) doc.addImage(imgData, "JPEG", 14, 5, 38, 13);

    doc.setFontSize(7);
    doc.setTextColor(120);
    doc.text(window.SWNET_HELPDESK + " - " + window.SWNET_USR, pgW - 10, 8, {
      align: "right",
    });

    doc.setFontSize(13);
    doc.setTextColor(40);
    doc.text("Consulta Pedidos Validados", pgW / 2, 14, { align: "center" });

    doc.setFontSize(8);
    doc.setTextColor(100);
    doc.text("Gerado em: " + new Date().toLocaleString("pt-BR"), pgW / 2, 20, {
      align: "center",
    });

    const headers = [
      "Número",
      "Emissão",
      "Ped.Cli-Item",
      "Status",
      "Cliente",
      "Loja",
      "Produto",
      "Referência",
      "TES",
      "Pagto.",
      "1ª",
      "2ª",
      "Qtde.",
      "Original",
      "Entregue",
    ];
    if (mostrarPreco)
      headers.push(
        "Preço",
        "Preço(Kg)",
        "KG NET",
        "Área(m²)",
        "Saldo",
        "Estrutura",
      );
    headers.push("Faturamento", "Fabril", "Incluiu", "Validou", "Piloto");

    const body = _filtrados.map(function (item) {
      const row = [
        item.C6_NUM,
        item.C5_EMISSAO,
        item.PED_CLI_ITEM,
        item.C5_ESTATUS,
        item.CLIENTE,
        item.LOJA_MUN_EST,
        item.C6_PRODUTO,
        item.DESCRICAO,
        item.C6_TES,
        item.CONDICAO_PAGAMENTO,
        item.DATAENTREG1,
        item.DATAENTREG2,
        fmtNum(item.C6_QTDVEN, 0),
        fmtNum(item.C6_QTDORIG, 0),
        fmtNum(item.C6_QTDENT, 0),
      ];
      if (mostrarPreco) {
        row.push(
          fmtNum(item.C6_PRCVEN, 5),
          fmtNum(item.PRECO_KG, 5),
          fmtNum(item.KG_NET, 5),
          fmtNum(item.AREA, 5),
          fmtNum(item.SALDO, 0),
          item.ESTRUTURA || "",
        );
      }
      row.push(
        item.D2 || "",
        item.FABRICA || "",
        item.INCLUIU || "",
        item.VALIDOU || "",
        item.C5_LOTEPIL || "",
      );
      return row;
    });

    doc.autoTable({
      head: [headers],
      body: body,
      startY: 26,
      margin: { left: 10 },
      theme: "grid",
      styles: { fontSize: 4.5, cellPadding: 1.5 },
      headStyles: {
        fillColor: [74, 74, 74],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [245, 245, 245] },
      tableLineColor: [180, 180, 180],
      tableLineWidth: 0.2,
    });

    doc.save("pedidos_validados.pdf");
  }

  // ── Callback global para popup buscaCLIENTE.asp ───────────────────────
  function objReturn(obj) {
    const limpar = function () {
      $id("codcli").value = "";
      $id("lojacli").value = "";
      $id("razao").value = "";
    };
    if (obj.erro === "ErroLogof") {
      mostrarToast("Seu logon EXPIROU. Favor se conectar novamente.", "erro");
      if (obj.tipo === "Grupo") limpar();
      return false;
    }
    if (obj.erro === "ErroNEncontrado") {
      if (obj.tipo === "Grupo") {
        limpar();
        mostrarToast("Cliente não encontrado.", "aviso");
      }
      return false;
    }
    if (obj.erro === "ErroCliente") {
      mostrarToast("Favor preencher os campos Cliente e Loja.", "aviso");
      if (obj.tipo === "Grupo") limpar();
      return false;
    }
    if (obj.tipo === "Grupo") {
      // razao já foi preenchido pelo modal via opener; monta exibição unificada
      var cod  = $id("codcli").value.trim();
      var nome = $id("razao").value.trim();
      if (cod && nome) $id("codcli").value = cod + " - " + nome;
    }
  }

  // ── Init ──────────────────────────────────────────────────────────────
  function init() {
    $id("btnFiltrar").addEventListener("click", filtrar);
    $id("btnLimpar").addEventListener("click", limpar);
    $id("btnExcel").addEventListener("click", exportarExcel);
    $id("btnPdf").addEventListener("click", exportarPDF);

    // Salva estado antes de navegar para o detalhamento
    document.addEventListener("click", function (e) {
      const a = e.target.closest(".link-detalhe");
      if (!a) return;
      try {
        sessionStorage.setItem(
          CACHE_KEY_DADOS,
          JSON.stringify(_dadosCompletos),
        );
        sessionStorage.setItem(
          CACHE_KEY_FILTROS,
          JSON.stringify({
            datainicial: $id("datainicial").value,
            datafinal: $id("datafinal").value,
            pedidoInicial: $id("pedidoInicial").value,
            pedidoFinal: $id("pedidoFinal").value,
            codcli: $id("codcli").value,
            lojacli: $id("lojacli").value,
            razao: $id("razao").value,
            tipo: _tipoAtual,
          tipoDat: $id("tipoDat").value,
          }),
        );
        sessionStorage.setItem(CACHE_KEY_PAGINA, String(_paginaAtual));
        sessionStorage.setItem(CACHE_KEY_LIMITE, String(_limite));
        sessionStorage.setItem(CACHE_FLAG_RETORNO, "1");
      } catch (e) {
        // sessionStorage cheio ou indisponível — ignora
      }
    });

    $id("btnLimparCliente").addEventListener("click", function () {
      $id("codcli").value  = "";
      $id("lojacli").value = "";
      $id("razao").value   = "";
    });

    $id("btnBuscarCliente").addEventListener("click", function () {
      window.open(
        "../../../../buscaCLIENTE.asp?id1=codcli&id2=lojacli&id3=razao&id6=00",
        "Consulta Cliente",
        "width=650,height=400,scrollbars=1",
      );
    });

    // ── Enter no campo unificado: busca e exibe "COD - NOME" ─────────────
    $id("codcli").addEventListener("keydown", function (e) {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const cod = $id("codcli").value.trim().split(" - ")[0].trim();
      if (cod.length === 0) return;
      $.ajax({
        url: REPOSITORIO,
        method: "GET",
        dataType: "json",
        cache: false,
        data: { action: "buscarClientes", termo: cod },
        success: function (resp) {
          if (!resp.sucesso || resp.clientes.length === 0) {
            mostrarToast("Cliente \"" + cod + "\" não encontrado.", "aviso");
            return;
          }
          const partes = resp.clientes[0].label.split(" - ");
          const nome = partes.length >= 3 ? partes.slice(2).join(" - ") : resp.clientes[0].label;
          $id("codcli").value = cod + " - " + nome;
          $id("razao").value  = nome;
        }
      });
    });

    $id("filtroTabela").addEventListener("input", _aplicarBusca);

    document.addEventListener("keydown", function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === "k") {
        e.preventDefault();
        const inp = $id("filtroTabela");
        if (inp) {
          inp.focus();
          inp.select();
        }
      }
    });

    $id("selectLimite").addEventListener("change", function () {
      _limite = parseInt($id("selectLimite").value, 10);
      _paginaAtual = 1;
      _renderTbody();
    });

    // ── Restaurar estado ao retornar do detalhamento ────────────────────
    (function _restaurarCache() {
      const retornando = sessionStorage.getItem(CACHE_FLAG_RETORNO) === "1";
      sessionStorage.removeItem(CACHE_FLAG_RETORNO);

      if (!retornando) {
        sessionStorage.removeItem(CACHE_KEY_DADOS);
        sessionStorage.removeItem(CACHE_KEY_FILTROS);
        sessionStorage.removeItem(CACHE_KEY_PAGINA);
        sessionStorage.removeItem(CACHE_KEY_LIMITE);
        return;
      }

      try {
        const dadosRaw = sessionStorage.getItem(CACHE_KEY_DADOS);
        const filtrosRaw = sessionStorage.getItem(CACHE_KEY_FILTROS);
        if (!dadosRaw || !filtrosRaw) return;

        const dados = JSON.parse(dadosRaw);
        const filtros = JSON.parse(filtrosRaw);
        const pagina = parseInt(
          sessionStorage.getItem(CACHE_KEY_PAGINA) || "1",
          10,
        );
        const limite = parseInt(
          sessionStorage.getItem(CACHE_KEY_LIMITE) || "9999",
          10,
        );

        // Restaura filtros no formulário
        $id("datainicial").value = filtros.datainicial || "";
        $id("datafinal").value = filtros.datafinal || "";
        $id("pedidoInicial").value = filtros.pedidoInicial || "";
        $id("pedidoFinal").value = filtros.pedidoFinal || "";
        $id("codcli").value   = filtros.codcli  || "";
        $id("lojacli").value  = filtros.lojacli || "";
        $id("tipoDat").value  = filtros.tipoDat || "E";

        // Restaura estado interno
        _tipoAtual = filtros.tipo || "";
        _limite = limite;
        _paginaAtual = pagina;
        _dadosCompletos = dados;
        _filtrados = dados.slice();

        $id("selectLimite").value = String(limite);

        // Reconstrói thead (depende do tipo) e exibe tabela
        _renderThead();
        _mostrarTabela();
      } catch (e) {
        // Cache corrompido — ignora e aguarda nova filtragem
      }
    })();
  }

  document.addEventListener("DOMContentLoaded", init);

  // APIs globais necessárias para popups externos
  window.objReturn = objReturn;
  window.janela = function (url, titulo, opcoes) {
    window.open(url, titulo, opcoes);
  };
})();

  </script>

</body>
</html>
