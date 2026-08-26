<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#include file="../includes/variaveis.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Relatório de Menu</title>

    <link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

    <style>
      /* ─── BASE ─────────────────────────────────────────────── */
      body {
        background: #f0f0f0;
        color: #212529;
        font-size: .8125rem;
        min-height: 100vh;
      }

      .page-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #212529;
      }

      /* ─── FILTER CARD ───────────────────────────────────────── */
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

      .form-control,
      .form-select {
        font-size: .85rem;
      }

      .form-control:focus,
      .form-select:focus {
        border-color: #4a4a4a;
        box-shadow: 0 0 0 .2rem rgba(74,74,74,.15);
      }

      /* ─── TABLE CARD ────────────────────────────────────────── */
      .table-card {
        background-color: #fff;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        position: relative;
      }

      /* ─── TOOLBAR ───────────────────────────────────────────── */
      .table-toolbar {
        background-color: #fff;
        border-radius: 6px 6px 0 0;
      }

      /* ─── TABLE WRAPPER ─────────────────────────────────────── */
      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }

      .table-wrapper::-webkit-scrollbar       { width: 5px; height: 5px; }
      .table-wrapper::-webkit-scrollbar-track  { background: #f0f0f0; border-radius: 3px; }
      .table-wrapper::-webkit-scrollbar-thumb  { background: #c0c0c0; border-radius: 3px; }
      .table-wrapper::-webkit-scrollbar-thumb:hover { background: #a0a0a0; }

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

      /* ─── TABLE FOOTER ──────────────────────────────────────── */
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
        min-width: 160px;
        display: flex;
        align-items: center;
        gap: .5rem;
      }

      .table-footer-center {
        flex: 1;
        display: flex;
        justify-content: center;
      }

      .table-footer-right {
        min-width: 140px;
        text-align: right;
      }

      .info-registros { font-size: .8rem; color: #6c757d; }
      .select-limite  { width: auto; }

      .paginacao .page-link {
        font-size: .8rem;
        padding: .25rem .55rem;
        color: #4a4a4a;
        border-color: #dee2e6;
      }

      .paginacao .page-item.active .page-link {
        background: #4a4a4a;
        border-color: #4a4a4a;
        color: #fff;
      }

      .paginacao .page-item.disabled .page-link {
        color: #adb5bd;
        pointer-events: none;
      }

      /* ─── PLACEHOLDER ───────────────────────────────────────── */
      .table-placeholder {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 3.5rem 1rem;
        color: #6c757d;
        gap: .75rem;
      }

      .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
      .table-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

      /* ─── OVERLAY LOADING ───────────────────────────────────── */
      .overlay-loading {
        position: absolute;
        inset: 0;
        background: rgba(255,255,255,.80);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 10;
        gap: 1rem;
        border-radius: 6px;
      }

      .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
      .overlay-loading span            { font-size: .9rem; font-weight: 600; color: #495057; }

      /* ─── TOAST ─────────────────────────────────────────────── */
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

      .toast-custom .toast-progress-bar {
        height: 3px;
        animation: toastProgressShrink 4.5s linear forwards;
      }

      @keyframes toastProgressShrink {
        from { width: 100%; }
        to   { width: 0%; }
      }

      /* ─── FILIAL BADGES ─────────────────────────────────────── */
      .filial-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 3px;
        font-size: .72rem;
        font-weight: 600;
        letter-spacing: .2px;
      }

      .filial-00 { background: #dbeafe; color: #1e40af; }
      .filial-01 { background: #fce7f3; color: #9d174d; }
      .filial-02 { background: #d1fae5; color: #065f46; }
      .filial-03 { background: #ede9fe; color: #4c1d95; }
      .filial-04 { background: #fef3c7; color: #78350f; }
      .filial-05 { background: #cffafe; color: #155e75; }
      .filial-06 { background: #fee2e2; color: #7f1d1d; }
      .filial-07 { background: #d1fae5; color: #064e3b; }
      .filial-08 { background: #ffedd5; color: #7c2d12; }
      .filial-09 { background: #e0e7ff; color: #1e1b4b; }
      .filial-10 { background: #fef9c3; color: #713f12; }
      .filial-11 { background: #f0fdf4; color: #14532d; }
      .filial-12 { background: #fdf4ff; color: #581c87; }
      .filial-13 { background: #ecfdf5; color: #065f46; }

      /* ─── STATUS BADGES ─────────────────────────────────────── */
      .status-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: .72rem; font-weight: 600; }
      .status-ativo   { background: #dcfce7; color: #166534; }
      .status-inativo { background: #fee2e2; color: #991b1b; }

      /* ─── MODAL ─────────────────────────────────────────────── */
      .modal-content { border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 8px 30px rgba(0,0,0,.2); }

      .modal-header-sk {
        background: #fff;
        color: #212529;
        padding: .65rem 1rem;
        border-bottom: 1px solid #dee2e6;
        border-radius: 6px 6px 0 0;
      }

      .modal-header-sk .modal-title { font-size: .8125rem; font-weight: 600; }

      .modal-footer-sk {
        background: #f8f9fa;
        border-top: 1px solid #dee2e6;
        padding: .6rem 1rem;
      }

      .modal-menu-list {
        max-height: 400px;
        overflow-y: auto;
        padding-bottom: 1rem;
      }

      .modal-menu-list::-webkit-scrollbar       { width: 5px; }
      .modal-menu-list::-webkit-scrollbar-track  { background: #f0f0f0; }
      .modal-menu-list::-webkit-scrollbar-thumb  { background: #c0c0c0; border-radius: 3px; }

      .modal-categoria-header {
        background: #1a1a1a;
        color: #fff;
        padding: .4rem .75rem;
        font-size: .72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .3px;
        position: sticky;
        top: 0;
        z-index: 1;
        border-bottom: 1px solid #000;
      }

      .modal-menu-item {
        padding: .45rem .75rem .45rem 1.25rem;
        border-bottom: 1px solid #e8e8e8;
        cursor: pointer;
        font-size: .75rem;
        transition: all .15s ease;
        background: #fff;
      }

      .modal-menu-item:hover { background: rgba(0,0,0,.03); transform: translateX(2px); }
      .modal-menu-item:last-child { border-bottom: none; }

      .input-menu-trigger { cursor: pointer; background: #fff; }
      .input-menu-trigger:hover { border-color: #4a4a4a; }

      /* ─── SWITCH ────────────────────────────────────────────── */
      .form-check-input[role="switch"] {
        width: 2.75rem; height: 1.5rem; border-radius: 9999px;
        border: 1px solid #d1d5db; background-color: #e5e7eb;
        background-image: none; appearance: none; cursor: pointer;
        position: relative;
        transition: background-color .2s ease, border-color .2s ease;
        flex-shrink: 0;
      }

      .form-check-input[role="switch"]::before {
        content: "✕"; position: absolute; top: 50%; left: 3px;
        transform: translateY(-50%); width: 1.1rem; height: 1.1rem;
        border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.2);
        transition: transform .2s ease, left .2s ease;
        display: flex; align-items: center; justify-content: center;
        font-size: .5rem; color: #9ca3af; line-height: 1.1rem; text-align: center;
      }

      .form-check-input[role="switch"]:checked { background-color: #4f46e5; border-color: #4f46e5; }
      .form-check-input[role="switch"]:checked::before { content: "✓"; left: calc(100% - 1.1rem - 3px); color: #4f46e5; transform: translateY(-50%); }
      .form-check-input[role="switch"]:focus   { box-shadow: 0 0 0 .2rem rgba(79,70,229,.2); border-color: #818cf8; outline: none; }

      @media (max-width: 768px)  { .table-wrapper { max-height: calc(100vh - 300px); } }
      @media (min-width: 1920px) { .table-wrapper { max-height: calc(90vh - 220px); } }
    </style>
  </head>

  <body>
    <!-- Toast container -->
    <div class="toast-container-custom" id="toastWrapper"></div>

    <!-- Hidden: usuário logado -->
    <input type="hidden" id="loginUsuarioLogado" value="<%=Trim(Session("USR"))%>" />

    <div class="container-fluid py-3 px-3 px-lg-4">

      <!-- Título -->
      <div class="mb-3">
        <span class="page-title">Relatório de Menu</span>
      </div>

      <!-- ── FILTROS ───────────────────────────────────────────── -->
      <div class="filter-card mb-3">
        <div class="filter-card-header">
          <i class="bi bi-funnel me-2"></i>Filtros
        </div>
        <div class="filter-card-body">
          <div class="row g-2 align-items-end">

            <div class="col-12 col-md-auto" style="min-width:130px;max-width:160px;">
              <label class="form-label" for="cboTipo">Tipo</label>
              <select class="form-select form-select-sm" id="cboTipo">
                <option value="">Selecione...</option>
                <option value="servico">Serviços</option>
                <option value="consulta">Consultas</option>
              </select>
            </div>

            <div class="col-12 col-md-4">
              <label class="form-label">Menu</label>
              <div class="input-group input-group-sm">
                <input type="text" class="form-control input-menu-trigger" id="inputMenuNome"
                       placeholder="Clique para selecionar..." readonly />
                <button class="btn btn-outline-secondary" type="button" id="btnAbrirModal" title="Selecionar menu">
                  <i class="bi bi-box-arrow-in-down"></i>
                </button>
              </div>
              <input type="hidden" id="inputMenuCodigo" />
            </div>

            <div class="col-auto d-flex align-items-end pb-1 ms-2">
              <div class="form-check form-switch mb-0">
                <input class="form-check-input me-2" type="checkbox" role="switch" id="chkApenasAtivos" checked />
                <label class="form-check-label fw-semibold" for="chkApenasAtivos" style="font-size:.82rem;">Apenas Ativos</label>
              </div>
            </div>

            <div class="col-auto d-flex align-items-end gap-2 ms-auto">
              <button class="btn btn-light btn-sm" id="btnLimpar" type="button">
                <i class="bi bi-eraser me-2"></i>Limpar
              </button>
              <button class="btn btn-primary btn-sm" id="btnFiltrar" type="button">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              
            </div>

          </div>
        </div>
      </div>

      <!-- ── TABLE CARD ────────────────────────────────────────── -->
      <div class="table-card">

        <!-- Overlay de carregamento (posicionado dentro do card) -->
        <div id="overlayLoading" class="overlay-loading d-none">
          <div class="spinner-border" role="status">
            <span class="visually-hidden">Carregando...</span>
          </div>
          <span>Carregando...</span>
        </div>

        <!-- Toolbar: exportações -->
        <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
          <div class="d-flex gap-2">
            <button id="btnExportCsv" type="button" class="btn btn-success btn-sm" disabled>
              <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
            </button>
            <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm" disabled>
              <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
            </button>
          </div>
        </div>

        <!-- Wrapper da tabela -->
        <div class="table-wrapper" id="tabelaWrapper">
          <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaResultado">
            <thead>
              <tr>
                <th style="width:3rem;">Nº</th>
                <th style="width:6rem;">Matrícula</th>
                <th id="thStatus" class="d-none" style="width:5rem;">Status</th>
                <th>Nome</th>
                <th style="width:10rem;">Login</th>
                <th style="width:8rem;">Filial</th>
              </tr>
            </thead>
            <tbody id="tbodyResultado"></tbody>
          </table>

          <!-- Placeholder -->
          <div class="table-placeholder" id="placeholderTabela">
            <i class="bi bi-clipboard-data"></i>
            <p>Selecione um <strong>Tipo</strong> e um <strong>Menu</strong>, depois clique em <strong>Filtrar</strong> para exibir os resultados.</p>
          </div>
        </div>

        <!-- Rodapé da tabela -->
        <div class="table-footer d-none" id="rodapeTabela">
          <div class="table-footer-left">
            <label for="selPorPagina" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">
              Linhas por página:
            </label>
            <select id="selPorPagina" class="form-select form-select-sm select-limite">
              <option value="10">10</option>
              <option value="25" selected>25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
          </div>

          <div class="table-footer-center">
            <nav>
              <ul class="pagination pagination-sm mb-0 paginacao" id="paginacao"></ul>
            </nav>
          </div>

          <div class="table-footer-right">
            <span class="info-registros" id="badgeTotal"></span>
          </div>
        </div>

      </div>
    </div>

    <!-- ── MODAL SELECIONAR MENU ─────────────────────────────── -->
    <div class="modal fade" id="modalSelecionarMenu" tabindex="-1" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-lg">
        <div class="modal-content">
          <div class="modal-header-sk d-flex align-items-center justify-content-between">
            <h5 class="modal-title" id="modalMenuTitulo">Selecionar Menu</h5>
            <button type="button" class="btn-close btn-sm" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="px-3 pt-3 pb-2">
            <input type="text" class="form-control form-control-sm" id="inputBuscaModal" placeholder="Buscar menu..." />
          </div>
          <div class="modal-menu-list" id="listaMenusModal"></div>
          <div class="modal-footer-sk d-flex justify-content-end">
            <button type="button" class="btn btn-outline-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
          </div>
        </div>
      </div>
    </div>

    <!-- ── CDNs ──────────────────────────────────────────────── -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

    <!-- Variáveis globais injetadas pelo ASP -->
    <script>
      const idUsuarioLogado    = "<%=IdUsuarioLogado%>";
      const loginUsuarioLogado = "<%=LoginUsuarioLogado%>";
      const LOGO_PATH          = "/corporativo/CriarPermissao/Logo/logo-smurfit1.png";
    </script>

    <script>
    "use strict";

      /* ── CONSTANTES ─────────────────────────────────────────── */
      const REPOSITORIO  = "RelatorioDeMenusRepositorio.asp";
      const TOAST_DELAY  = 4500;
      let   POR_PAGINA   = 25;

      const TOAST_ICON = Object.freeze({
        sucesso: { cls: "bi-check-circle-fill",          cor: "#198754" },
        erro:    { cls: "bi-x-circle-fill",              cor: "#dc3545" },
        aviso:   { cls: "bi-exclamation-triangle-fill",  cor: "#856404" },
        info:    { cls: "bi-info-circle-fill",            cor: "#6c757d" },
      });

      const FILIAIS = Object.freeze({
        "00": { nome: "Pirapetinga",     cls: "filial-00" },
        "01": { nome: "Rio de Janeiro",  cls: "filial-01" },
        "02": { nome: "São Paulo",       cls: "filial-02" },
        "03": { nome: "Corporativo",     cls: "filial-03" },
        "04": { nome: "Uberaba",         cls: "filial-04" },
        "05": { nome: "Bento Gonçalves", cls: "filial-05" },
        "06": { nome: "Fortaleza",       cls: "filial-06" },
        "07": { nome: "Saquarema",       cls: "filial-07" },
        "08": { nome: "Araçatuba",       cls: "filial-08" },
        "09": { nome: "Blumenau",        cls: "filial-09" },
        "10": { nome: "Pacajus",         cls: "filial-10" },
        "11": { nome: "Porto Feliz",     cls: "filial-11" },
        "12": { nome: "Três Barras",     cls: "filial-12" },
        "13": { nome: "Florestal",       cls: "filial-13" },
      });

      /* ── ESTADO ─────────────────────────────────────────────── */
      const $el = (id) => document.getElementById(id);

      let dadosCompletos  = [];
      let paginaAtual     = 1;
      let totalPaginas    = 0;
      let totalRegistros  = 0;
      let nomeMenuAtual   = "";
      let tipoMenuAtual   = "";
      let menusCarregados = [];
      let _modalInstancia = null;

      const Estado = Object.freeze({
        getDados()   { return dadosCompletos; },
        setDados(d)  {
          dadosCompletos = d;
          totalRegistros = d.length;
          totalPaginas   = Math.ceil(d.length / POR_PAGINA) || 1;
        },
        getPagina()  { return paginaAtual; },
        setPagina(p) { paginaAtual = p; },
        getNomeMenu(){ return nomeMenuAtual; },
        setNomeMenu(n){ nomeMenuAtual = n; },
        getTipoMenu(){ return tipoMenuAtual; },
        setTipoMenu(t){ tipoMenuAtual = t; },
        limpar() {
          dadosCompletos = []; paginaAtual = 1; totalPaginas = 0;
          totalRegistros = 0; nomeMenuAtual = ""; tipoMenuAtual = "";
        },
      });

      /* ── HELPERS ────────────────────────────────────────────── */
      function obterNomeFilial(cod) {
        const f = FILIAIS[String(cod).padStart(2, "0")];
        return f ? f.nome : (cod || "");
      }

      function renderFilialBadge(cod) {
        const key = String(cod).padStart(2, "0");
        const f   = FILIAIS[key];
        if (!f) return cod || "";
        return `<span class="filial-badge ${f.cls}">${f.nome}</span>`;
      }

      function mostrarLoading(show) {
        $('#overlayLoading').toggleClass('d-none', !show);
      }

      /* ── TOAST ──────────────────────────────────────────────── */
      function exibirToast(tipo, msg) {
        const cfg       = TOAST_ICON[tipo] || TOAST_ICON.info;
        const barColor  = { sucesso: '#198754', erro: '#dc3545', aviso: '#ffc107', info: '#6c757d' }[tipo] || '#6c757d';
        const el        = document.createElement('div');
        el.className    = 'toast-custom';
        el.innerHTML    =
          `<div class="d-flex align-items-center gap-2 px-3 py-2">` +
          `<i class="bi ${cfg.cls} flex-shrink-0" style="color:${cfg.cor};font-size:1rem;"></i>` +
          `<span class="flex-grow-1" style="font-size:.85rem;color:#333;">${msg}</span>` +
          `<button style="background:none;border:none;font-size:1rem;color:#999;cursor:pointer;padding:0 0 0 .5rem;" ` +
          `onclick="this.closest('.toast-custom').remove()">&times;</button>` +
          `</div>` +
          `<div class="toast-progress-bar" style="background:${barColor};animation-duration:${TOAST_DELAY}ms;"></div>`;
        $el('toastWrapper').appendChild(el);
        setTimeout(() => {
          el.style.transition = 'opacity .3s';
          el.style.opacity    = '0';
          setTimeout(() => el.remove(), 300);
        }, TOAST_DELAY);
      }

      /* ── API ($.ajax encapsulado em Promise) ────────────────── */
      const Api = Object.freeze({
        get(params) {
          return new Promise((resolve, reject) => {
            $.ajax({
              url:      REPOSITORIO,
              method:   'GET',
              dataType: 'json',
              data:     params,
              success:  resolve,
              error:    (xhr) => reject(new Error(xhr.responseText || 'Erro HTTP')),
            });
          });
        },
      });

      /* ── MODAL MENU ─────────────────────────────────────────── */
      function obterModalInstancia() {
        if (!_modalInstancia) {
          _modalInstancia = new bootstrap.Modal($el('modalSelecionarMenu'), {
            backdrop: 'static', keyboard: true,
          });
        }
        return _modalInstancia;
      }

      const ModalMenu = Object.freeze({
        async abrir() {
          const tipo = $el('cboTipo').value;
          if (!tipo) {
            exibirToast('aviso', 'Selecione o tipo (Serviços ou Consultas) antes de escolher o menu.');
            return;
          }

          const tituloTipo = tipo === 'servico' ? 'Serviços' : 'Consultas';
          $el('modalMenuTitulo').textContent = `Selecionar Menu — ${tituloTipo}`;
          $el('inputBuscaModal').value = '';
          $el('listaMenusModal').innerHTML =
            `<div class="text-center py-4"><div class="spinner-border spinner-border-sm text-secondary" role="status"></div></div>`;

          obterModalInstancia().show();

          try {
            const resp = await Api.get({ acao: 'listarMenus', tipo });
            if (!resp.sucesso) {
              $el('listaMenusModal').innerHTML =
                `<div class="text-center py-3 text-danger" style="font-size:.78rem;">${resp.mensagem}</div>`;
              return;
            }
            menusCarregados = resp.dados || [];
            this.renderizar(menusCarregados);
          } catch {
            $el('listaMenusModal').innerHTML =
              `<div class="text-center py-3 text-danger" style="font-size:.78rem;">Erro ao carregar menus.</div>`;
          }
        },

        renderizar(lista) {
          const container = $el('listaMenusModal');
          container.innerHTML = '';

          if (!lista.length) {
            container.innerHTML = `<div class="text-center py-3 text-muted" style="font-size:.78rem;">Nenhum menu encontrado.</div>`;
            return;
          }

          const agrupado = {};
          lista.forEach((m) => {
            const cat = m.categoria || 'Sem Categoria';
            if (!agrupado[cat]) agrupado[cat] = [];
            agrupado[cat].push(m);
          });

          const frag = document.createDocumentFragment();
          Object.keys(agrupado).sort().forEach((cat) => {
            const header = document.createElement('div');
            header.className   = 'modal-categoria-header';
            header.textContent = cat;
            frag.appendChild(header);

            agrupado[cat].forEach((m) => {
              const item = document.createElement('div');
              item.className      = 'modal-menu-item';
              item.textContent    = m.nome;
              item.dataset.codigo = m.codigo;
              item.dataset.nome   = m.nome;
              frag.appendChild(item);
            });
          });
          container.appendChild(frag);
        },

        filtrar(termo) {
          if (!termo) { this.renderizar(menusCarregados); return; }
          const t         = termo.toLowerCase();
          const filtrados = menusCarregados.filter((m) =>
            m.nome.toLowerCase().includes(t) || (m.categoria || '').toLowerCase().includes(t)
          );
          this.renderizar(filtrados);
        },

        selecionar(codigo, nome) {
          $el('inputMenuCodigo').value = codigo;
          $el('inputMenuNome').value   = nome;
          obterModalInstancia().hide();
        },
      });

      /* ── TABELA ─────────────────────────────────────────────── */
      const Tabela = Object.freeze({
        renderizar() {
          const dados  = Estado.getDados();
          const pagina = Estado.getPagina();
          const inicio = (pagina - 1) * POR_PAGINA;
          const fim    = Math.min(inicio + POR_PAGINA, dados.length);
          const tbody  = $el('tbodyResultado');
          tbody.innerHTML = '';

          if (!dados.length) {
            $el('placeholderTabela').classList.remove('d-none');
            $el('placeholderTabela').querySelector('p').innerHTML =
              'Nenhum usuário encontrado com esta permissão.';
            $el('rodapeTabela').classList.add('d-none');
            $el('btnExportPdf').disabled = true;
            $el('btnExportCsv').disabled = true;
            return;
          }

          $el('placeholderTabela').classList.add('d-none');

          const mostrarStatus = !$el('chkApenasAtivos').checked;
          $el('thStatus').classList.toggle('d-none', !mostrarStatus);

          const frag = document.createDocumentFragment();
          for (let i = inicio; i < fim; i++) {
            const u  = dados[i];
            const tr = document.createElement('tr');
            const statusCell = mostrarStatus
              ? `<td><span class="status-badge ${u.ativo == '1' ? 'status-ativo' : 'status-inativo'}">${u.ativo == '1' ? 'Ativo' : 'Inativo'}</span></td>`
              : '';
            tr.innerHTML =
              `<td>${i + 1}</td>` +
              `<td>${u.matricula}</td>` +
              statusCell +
              `<td class="text-start">${u.nome}</td>` +
              `<td>${u.login}</td>` +
              `<td>${renderFilialBadge(u.filial)}</td>`;
            frag.appendChild(tr);
          }
          tbody.appendChild(frag);

          $el('rodapeTabela').classList.remove('d-none');
          $el('badgeTotal').textContent =
            `${dados.length} usuário${dados.length !== 1 ? 's' : ''} com esta permissão`;
          $el('btnExportPdf').disabled = false;
          $el('btnExportCsv').disabled = false;

          this.renderizarPaginacao();
        },

        renderizarPaginacao() {
          const pagina = Estado.getPagina();
          const total  = Math.ceil(Estado.getDados().length / POR_PAGINA) || 1;
          const ul     = $el('paginacao');
          ul.innerHTML = '';

          const criar = (label, pg, disabled, active) => {
            const li = document.createElement('li');
            li.className = `page-item${disabled ? ' disabled' : ''}${active ? ' active' : ''}`;
            const a = document.createElement('a');
            a.className = 'page-link';
            a.href      = '#';
            a.innerHTML = label;
            if (!disabled && !active) {
              a.addEventListener('click', (e) => {
                e.preventDefault();
                Estado.setPagina(pg);
                this.renderizar();
              });
            }
            li.appendChild(a);
            return li;
          };

          ul.appendChild(criar('«', 1, pagina === 1, false));
          ul.appendChild(criar('‹', pagina - 1, pagina === 1, false));

          let ini = Math.max(1, pagina - 2);
          let fim = Math.min(total, pagina + 2);
          if (fim - ini < 4) { ini = Math.max(1, fim - 4); fim = Math.min(total, ini + 4); }

          for (let p = ini; p <= fim; p++) ul.appendChild(criar(p, p, false, p === pagina));

          ul.appendChild(criar('›', pagina + 1, pagina === total, false));
          ul.appendChild(criar('»', total, pagina === total, false));
        },
      });

      /* ── FILTRAR ────────────────────────────────────────────── */
      async function filtrar() {
        const tipo       = $el('cboTipo').value;
        const codigoMenu = $el('inputMenuCodigo').value;

        if (!tipo)       { exibirToast('aviso', 'Selecione o tipo (Serviços ou Consultas).'); return; }
        if (!codigoMenu) { exibirToast('aviso', 'Selecione um menu.'); return; }

        const apenasAtivos = $el('chkApenasAtivos').checked ? '1' : '0';

        mostrarLoading(true);
        try {
          const resp = await Api.get({ acao: 'buscarUsuarios', tipo, codigoMenu, apenasAtivos });

          if (!resp.sucesso) {
            exibirToast('erro', resp.mensagem || 'Erro ao buscar usuários.');
            Estado.limpar();
            Tabela.renderizar();
            return;
          }

          const d = resp.dados || {};
          Estado.setDados(d.usuarios || []);
          Estado.setNomeMenu(d.nomeMenu || $el('inputMenuNome').value);
          Estado.setTipoMenu(tipo === 'servico' ? 'Serviço' : 'Consulta');
          Estado.setPagina(1);
          Tabela.renderizar();

          if (!(d.usuarios || []).length) {
            exibirToast('info', 'Nenhum usuário encontrado com esta permissão.');
          }
        } catch (err) {
          exibirToast('erro', 'Erro na requisição: ' + err.message);
        } finally {
          mostrarLoading(false);
        }
      }

      /* ── LIMPAR ─────────────────────────────────────────────── */
      function limpar() {
        $el('cboTipo').value           = '';
        $el('inputMenuCodigo').value   = '';
        $el('inputMenuNome').value     = '';
        $el('chkApenasAtivos').checked = true;
        menusCarregados = [];
        Estado.limpar();
        $el('tbodyResultado').innerHTML = '';
        const ph = $el('placeholderTabela');
        ph.classList.remove('d-none');
        ph.querySelector('p').innerHTML =
          'Selecione um <strong>Tipo</strong> e um <strong>Menu</strong>, depois clique em <strong>Filtrar</strong> para exibir os resultados.';
        $el('rodapeTabela').classList.add('d-none');
        $el('btnExportPdf').disabled = true;
        $el('btnExportCsv').disabled = true;
      }

      /* ── METADADOS (PDF / CSV) ──────────────────────────────── */
      function obterMetadados() {
        return {
          tipoMenu:       Estado.getTipoMenu(),
          nomeMenu:       Estado.getNomeMenu(),
          apenasAtivos:   $el('chkApenasAtivos').checked ? 'Sim' : 'Não',
          usuarioLogado:  loginUsuarioLogado || $el('loginUsuarioLogado').value || 'Sistema',
          dataHora:       new Date().toLocaleString('pt-BR', {
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit',
          }),
          totalRegistros: Estado.getDados().length,
        };
      }

      /* ── EXPORTAR CSV ───────────────────────────────────────── */
      function exportarCSV(nomeArquivo, cabecalhos, linhas) {
        const esc = (v) => {
          const s = v == null ? '' : String(v);
          return (s.includes(';') || s.includes('"') || s.includes('\n'))
            ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const rows = [
          cabecalhos.map(esc).join(';'),
          ...linhas.map((row) => row.map(esc).join(';')),
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

      function gerarExcel() {
        const dados = Estado.getDados();
        if (!dados.length) { exibirToast('aviso', 'Nenhum dado para exportar.'); return; }

        const meta       = obterMetadados();
        const cabecalhos = ['Nº', 'Matrícula', 'Nome', 'Login', 'Filial'];
        const linhas     = dados.map((u, i) => [
          String(i + 1),
          u.matricula,
          u.nome,
          u.login,
          obterNomeFilial(u.filial),
        ]);

        const data  = new Date();
        const stamp = `${String(data.getDate()).padStart(2,'0')}${String(data.getMonth()+1).padStart(2,'0')}${data.getFullYear()}`;
        exportarCSV(`Relatorio_Menus_${meta.tipoMenu}_${stamp}`, cabecalhos, linhas);
        exibirToast('sucesso', 'CSV gerado com sucesso!');
      }

      /* ── EXPORTAR PDF ───────────────────────────────────────── */
      async function carregarLogoPDF() {
        return new Promise((resolve) => {
          const img       = new Image();
          img.crossOrigin = 'Anonymous';
          img.onload = function () {
            const scale  = img.width > 900 ? 900 / img.width : 1;
            const canvas = document.createElement('canvas');
            canvas.width  = Math.round(img.width  * scale);
            canvas.height = Math.round(img.height * scale);
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            resolve(canvas.toDataURL('image/png'));
          };
          img.onerror = () => resolve(null);
          img.src     = LOGO_PATH;
        });
      }

      async function gerarPDF() {
        const dados = Estado.getDados();
        if (!dados.length) { exibirToast('aviso', 'Nenhum dado para exportar.'); return; }

        mostrarLoading(true);
        try {
          const { jsPDF } = window.jspdf;
          const pdf  = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
          const meta = obterMetadados();
          const pw   = pdf.internal.pageSize.getWidth();

          const logo = await carregarLogoPDF();
          if (logo) pdf.addImage(logo, 'PNG', 10, 4, 38, 13);

          pdf.setFontSize(7);
          pdf.setFont('helvetica', 'normal');
          pdf.setTextColor(120, 120, 120);
          pdf.text(`Emitido por: ${meta.usuarioLogado}`, pw - 10, 8, { align: 'right' });

          pdf.setFontSize(13);
          pdf.setFont('helvetica', 'bold');
          pdf.setTextColor(74, 74, 74);
          pdf.text('Relatório de Menu', pw / 2, 14, { align: 'center' });

          pdf.setFontSize(8);
          pdf.setFont('helvetica', 'normal');
          pdf.setTextColor(100, 100, 100);
          pdf.text(
            `Tipo: ${meta.tipoMenu}  |  Menu: ${meta.nomeMenu}  |  Apenas Ativos: ${meta.apenasAtivos}  |  Total: ${meta.totalRegistros}  |  ${meta.dataHora}`,
            pw / 2, 20, { align: 'center' }
          );

          const mostrarStatus = !$el('chkApenasAtivos').checked;
          const cabHead  = ['Nº', 'Matrícula'];
          if (mostrarStatus) cabHead.push('Status');
          cabHead.push('Nome', 'Login', 'Filial');

          const linhas = dados.map((u, i) => {
            const row = [i + 1, u.matricula];
            if (mostrarStatus) row.push(u.ativo == '1' ? 'Ativo' : 'Inativo');
            row.push(u.nome, u.login, obterNomeFilial(u.filial));
            return row;
          });

          pdf.autoTable({
            startY: 26,
            head:   [cabHead],
            body:   linhas,
            theme:  'grid',
            styles:     { fontSize: 4.5, cellPadding: 1.5, overflow: 'linebreak', lineColor: [180,180,180], lineWidth: .2 },
            headStyles: { fillColor: [74,74,74], textColor: [255,255,255], fontStyle: 'bold', fontSize: 4.5 },
            alternateRowStyles: { fillColor: [245,245,245] },
            columnStyles: {
              0: { cellWidth: 8,    halign: 'center' },
              1: { cellWidth: 18,   halign: 'center' },
              2: { cellWidth: 'auto' },
            },
            margin: { left: 10, right: 10 },
          });

          const ph = pdf.internal.pageSize.getHeight();
          const tp = pdf.internal.getNumberOfPages();
          for (let i = 1; i <= tp; i++) {
            pdf.setPage(i);
            pdf.setDrawColor(180, 180, 180);
            pdf.setLineWidth(.3);
            pdf.line(10, ph - 12, pw - 10, ph - 12);
            pdf.setTextColor(120, 120, 120);
            pdf.setFontSize(7);
            pdf.text(`Página ${i} de ${tp}`, pw / 2, ph - 6, { align: 'center' });
          }

          const data  = new Date();
          const stamp = `${String(data.getDate()).padStart(2,'0')}${String(data.getMonth()+1).padStart(2,'0')}${data.getFullYear()}`;
          pdf.save(`Relatorio_Menus_${meta.tipoMenu}_${stamp}.pdf`);
          exibirToast('sucesso', 'PDF gerado com sucesso!');
        } catch (err) {
          exibirToast('erro', 'Erro ao gerar PDF: ' + err.message);
        } finally {
          mostrarLoading(false);
        }
      }

      /* ── EVENTOS ────────────────────────────────────────────── */
      function anexarEventos() {
        $el('cboTipo').addEventListener('change', function () {
          $el('inputMenuCodigo').value = '';
          $el('inputMenuNome').value   = '';
          menusCarregados = [];
        });

        $el('inputMenuNome').addEventListener('click', () => ModalMenu.abrir());
        $el('btnAbrirModal').addEventListener('click', () => ModalMenu.abrir());

        $el('listaMenusModal').addEventListener('click', function (e) {
          const item = e.target.closest('.modal-menu-item');
          if (item) ModalMenu.selecionar(item.dataset.codigo, item.dataset.nome);
        });

        let timerBuscaModal;
        $el('inputBuscaModal').addEventListener('input', function () {
          clearTimeout(timerBuscaModal);
          const termo = this.value.trim();
          timerBuscaModal = setTimeout(() => ModalMenu.filtrar(termo), 200);
        });

        $el('btnFiltrar').addEventListener('click', () => filtrar());
        $el('btnLimpar').addEventListener('click',  () => limpar());
        $el('btnExportPdf').addEventListener('click', () => gerarPDF());
        $el('btnExportCsv').addEventListener('click', () => gerarExcel());

        $el('selPorPagina').addEventListener('change', function () {
          POR_PAGINA  = parseInt(this.value, 10);
          totalPaginas = Math.ceil(dadosCompletos.length / POR_PAGINA) || 1;
          Estado.setPagina(1);
          Tabela.renderizar();
        });
      }

      document.addEventListener('DOMContentLoaded', () => {
        anexarEventos();
      });
    </script>
  </body>
</html>
