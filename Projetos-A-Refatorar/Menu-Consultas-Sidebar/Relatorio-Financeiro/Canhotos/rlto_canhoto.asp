<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Buffer   = True
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Response.Expires  = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

Dim IdUsuarioLogado, LoginUsuarioLogado
IdUsuarioLogado    = Session("CODIGO")
LoginUsuarioLogado = Session("USR")

Dim dtini, dtfin
If Trim(Request("dtini")) <> "" And Trim(Request("dtfin")) <> "" Then
  dtini = Trim(Request("dtini"))
  dtfin = Trim(Request("dtfin"))
Else
  dtini = formataDataUS(Date())
  dtfin = formataDataUS(DateAdd("m", 1, DateAdd("d", -1, Date())))
End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet — Consulta Canhoto</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

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
    .filter-label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }
    .input-data { width: 155px; }

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
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p { margin: 0; font-size: .9rem; text-align: center; }

    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 68vh;
    }
    .table-wrapper table { min-width: max-content; }

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
    .table-footer-left  { min-width: 120px; }
    .table-footer-center { flex: 1; display: flex; justify-content: center; }
    .table-footer-right { min-width: 120px; text-align: right; }
    .info-registros { font-size: .8rem; }
    .select-limite  { width: auto; }

    #paginacao .page-link         { color: #6c757d; border-color: #dee2e6; }
    #paginacao .page-item.active .page-link {
      background-color: #6c757d;
      border-color: #6c757d;
      color: #fff;
    }
    #paginacao .page-link:hover   { background-color: #e9ecef; color: #495057; }
    #paginacao .page-item.disabled .page-link { color: #adb5bd; }

    .search-box { position: relative; max-width: 200px; }
    .search-box .form-control { padding-right: 72px; font-size: .82rem; }
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
      gap: .6rem;
      padding: .75rem 1rem;
    }
    .toast-icon      { font-size: 1.15rem; flex-shrink: 0; margin-top: .05rem; }
    .toast-body-text { flex: 1; }
    .toast-title     { font-weight: 700; font-size: .82rem; }
    .toast-msg       { font-size: .82rem; color: #495057; }
    .toast-close     {
      background: none;
      border: none;
      font-size: 1.1rem;
      line-height: 1;
      cursor: pointer;
      color: #6c757d;
      flex-shrink: 0;
    }
    .toast-progress     { height: 3px; background: #f0f0f0; }
    .toast-progress-bar { height: 3px; animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    .toast-sucesso .toast-icon,
    .toast-sucesso .toast-title  { color: #198754; }
    .toast-sucesso .toast-progress-bar { background-color: #198754; }

    .toast-erro .toast-icon,
    .toast-erro .toast-title     { color: #dc3545; }
    .toast-erro .toast-progress-bar { background-color: #dc3545; }

    .toast-aviso .toast-icon,
    .toast-aviso .toast-title    { color: #856404; }
    .toast-aviso .toast-progress-bar { background-color: #ffc107; }

    .toast-info .toast-icon,
    .toast-info .toast-title     { color: #6c757d; }
    .toast-info .toast-progress-bar { background-color: #6c757d; }

    #modalBuscaCliente .modal-cliente-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3rem 1rem;
      color: #6c757d;
      gap: .6rem;
      background-color: #f8f9fa;
      border-radius: 6px;
      border: 1px dashed #dee2e6;
    }
    #modalBuscaCliente .modal-cliente-placeholder .bi { font-size: 2.2rem; opacity: .3; }
    #modalBuscaCliente .modal-cliente-placeholder p   { margin: 0; font-size: .83rem; }
    #modalBuscaCliente .modal-header  { border-bottom: 1px solid #dee2e6; }
    #modalBuscaCliente .modal-body    { padding: 1.25rem; }
    #modalBuscaCliente .modal-footer  {
      background-color: #fff;
      border-top: 1px solid #dee2e6;
      padding: .6rem 1rem;
    }
    #modalBuscaCliente .input-busca-wrapper          { position: relative; }
    #modalBuscaCliente .input-busca-wrapper .bi-search {
      position: absolute;
      left: .75rem;
      top: 50%;
      transform: translateY(-50%);
      color: #adb5bd;
      font-size: .9rem;
      pointer-events: none;
    }
    #modalBuscaCliente #inputBuscaCliente { padding-left: 2.1rem; }
    #modalBuscaCliente .modal-resultado-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: .5rem;
    }
    #modalBuscaCliente .modal-resultado-header .resultado-count {
      font-size: .78rem;
      color: #6c757d;
    }
    #modalClienteTabela thead th {
      background-color: #f8f9fa;
      font-size: .78rem;
      font-weight: 700;
      color: #495057;
      text-align: left;
      padding: .5rem .75rem;
      border-color: #dee2e6;
      white-space: nowrap;
    }
    #modalClienteTabela tbody td {
      font-size: .82rem;
      color: #212529;
      text-align: left;
      vertical-align: middle;
      padding: .45rem .75rem;
      border-color: #dee2e6;
    }
    #modalClienteTabela tbody tr {
      cursor: pointer;
      transition: background-color .12s ease;
    }
    #modalClienteTabela tbody tr:hover td { background-color: #e8f0fe; }
    #modalClienteTabela .btn-importar {
      padding: .2rem .45rem;
      font-size: .78rem;
      line-height: 1;
    }

    /* ── Dropdown multiselect (Produto / Serviço) ──────────────────────── */
    .dd-toggle { width: 100%; text-align: left; }
    .dd-toggle::after { display: none; }
    .dd-menu { min-width: 100%; }
  </style>
</head>

<body>

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border text-secondary" role="status"></div>
    <span>Carregando...</span>
  </div>

  <div id="toastMsg" class="toast-container-custom"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="mb-3">
      <span class="page-title">Consulta Canhoto</span>
    </div>

    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <div class="row g-2 align-items-end">

          <%
          Dim reqFilial : reqFilial = Trim(Request("filial"))
          %>
          <div class="col-12 col-md-auto" style="min-width:140px;max-width:180px;">
            <label for="filial" class="filter-label d-block">Planta</label>
            <select id="filial" name="filial" class="form-select form-select-sm">
              <option value="">Todas</option>
              <option value="0" <%if reqFilial="0" then%>selected<%end if%>>Pirapetinga</option>
              <option value="4" <%if reqFilial="4" then%>selected<%end if%>>Uberaba</option>
              <option value="5" <%if reqFilial="5" then%>selected<%end if%>>Bento Gonçalves</option>
              <option value="6" <%if reqFilial="6" then%>selected<%end if%>>Fortaleza</option>
              <option value="7" <%if reqFilial="7" then%>selected<%end if%>>Saquarema</option>
            </select>
          </div>

          <div class="col-12 col-md-auto" style="min-width:155px;max-width:200px;">
            <label class="filter-label d-block">Produto</label>
            <div class="dropdown">
              <button id="btnCarga" type="button"
                      class="form-select form-select-sm dd-toggle dropdown-toggle"
                      data-bs-toggle="dropdown" data-bs-auto-close="outside"
                      aria-expanded="false">Selecione...</button>
              <ul class="dropdown-menu dd-menu py-1">
                <li><button class="dropdown-item" data-val="0" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Não Respondido</span></button></li>
                <li><button class="dropdown-item" data-val="5" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Muito Satisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="4" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Satisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="3" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Indiferente</span></button></li>
                <li><button class="dropdown-item" data-val="2" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Insatisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="1" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Muito Insatisfeito</span></button></li>
              </ul>
              <input type="hidden" id="carga" name="carga" />
            </div>
          </div>

          <div class="col-12 col-md-auto" style="min-width:155px;max-width:200px;">
            <label class="filter-label d-block">Serviço</label>
            <div class="dropdown">
              <button id="btnVeiculo" type="button"
                      class="form-select form-select-sm dd-toggle dropdown-toggle"
                      data-bs-toggle="dropdown" data-bs-auto-close="outside"
                      aria-expanded="false">Selecione...</button>
              <ul class="dropdown-menu dd-menu py-1">
                <li><button class="dropdown-item" data-val="0" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Não Respondido</span></button></li>
                <li><button class="dropdown-item" data-val="5" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Muito Satisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="4" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Satisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="3" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Indiferente</span></button></li>
                <li><button class="dropdown-item" data-val="2" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Insatisfeito</span></button></li>
                <li><button class="dropdown-item" data-val="1" type="button"><input type="checkbox" class="form-check-input me-1" /><span>Muito Insatisfeito</span></button></li>
              </ul>
              <input type="hidden" id="veiculo" name="veiculo" />
            </div>
          </div>

          <div class="col-12 col-md-auto" style="min-width:140px;max-width:300px;">
            <label for="clienteDisplay" class="filter-label d-block">Cliente</label>
            <div class="input-group input-group-sm">
              <input id="clienteDisplay" type="text" class="form-control form-control-sm"
                     readonly placeholder="Selecione..." />
              <input type="hidden" id="cliente" name="cliente" value="<%=Trim(Request("cliente"))%>" />
              <input type="hidden" id="loja" name="loja" />
              <button type="button" class="btn btn-outline-primary btn-sm"
                      data-bs-toggle="modal" data-bs-target="#modalBuscaCliente"
                      title="Buscar cliente">
                <i class="bi bi-search"></i>
              </button>
            </div>
          </div>

          <div class="col-6 col-md-auto" style="min-width:90px;max-width:110px;">
            <label for="nf" class="filter-label d-block">Nota Fiscal</label>
            <input type="text" id="nf" name="nf" class="form-control form-control-sm"
                   value="<%=Trim(Request("nf"))%>" />
          </div>

          <div class="col-6 col-md-auto"> 
            <label for="dtini" class="filter-label d-block">De</label>
            <input id="dtini" name="dtini" type="date" class="form-control form-control-sm input-data" value="<%=dtini%>" />
          </div>

          <div class="col-6 col-md-auto">
            <label for="dtfin" class="filter-label d-block">Até</label>
            <input id="dtfin" name="dtfin" type="date" class="form-control form-control-sm input-data" value="<%=dtfin%>" />
          </div>

          <div class="col-12 col-md">
            <div class="d-flex gap-2 justify-content-end flex-nowrap">
              <button id="btnFiltrar" type="button" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button id="btnLimpar" type="button" class="btn btn-light btn-sm" title="Limpar">
                <i class="bi bi-eraser"></i>
              </button>
              <button id="btnExcel" type="button" class="btn btn-success btn-sm" title="Exportar Excel">
                <i class="bi bi-file-earmark-excel"></i>
              </button>
              <button id="btnPdf" type="button" class="btn btn-danger btn-sm" title="Exportar PDF">
                <i class="bi bi-file-earmark-pdf"></i>
              </button>
              <a id="btnGrafico" href="graf_canhoto.asp" class="btn btn-warning btn-sm" title="Gráfico">
                <i class="bi bi-graph-up-arrow text-dark"></i>
              </a>
            </div>
          </div>

        </div>
      </div>
    </div>

    <div class="table-card">

      <div id="placeholderTabela" class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>

      <div id="areaTabela" class="d-none">

        <div class="d-flex align-items-center justify-content-between px-3 pt-3 pb-2 flex-wrap gap-2">
          <div class="search-box">
            <input id="inputBusca" type="text" class="form-control form-control-sm" placeholder="Buscar..." />
            <span class="search-hotkey"><kbd>CTRL+K</kbd></span>
          </div>
        </div>

        <div class="table-wrapper">
          <table id="tabelaDados" class="table table-bordered table-striped table-hover table-sm mb-0">
            <thead>
              <tr>
                <th>Nota Fiscal</th>
                <th>Série</th>
                <th>Cliente</th>
                <th>Filial</th>
                <th>Data Canhoto</th>
                <th>Produto</th>
                <th>Serviço</th>
                <th>Usuário</th>
                <th>Observação</th>
              </tr>
            </thead>
            <tbody id="tbodyDados"></tbody>
          </table>
        </div>

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

  <div class="modal fade" id="modalBuscaCliente" tabindex="-1" aria-labelledby="modalBuscaClienteLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="modalBuscaClienteLabel">Busca de Clientes</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <div class="input-busca-wrapper mb-3">
            <i class="bi bi-search"></i>
            <input id="inputBuscaCliente" type="text" class="form-control form-control-sm"
                   placeholder="Filtrar por nome ou código..." />
          </div>
          <div id="placeholderModalClienteSemDados" class="modal-cliente-placeholder">
            <i class="bi bi-funnel"></i>
            <p>Digite ao menos 2 caracteres para buscar.</p>
          </div>
          <div id="areaModalTabelaCliente" class="d-none">
            <div class="modal-resultado-header">
              <span id="modalClienteCount" class="resultado-count"></span>
            </div>
            <div class="table-wrapper" style="max-height:340px;">
              <table class="table table-bordered table-hover table-sm mb-0" id="modalClienteTabela">
                <thead>
                  <tr>
                    <th style="width:42px;"></th>
                    <th>Cliente</th>
                  </tr>
                </thead>
                <tbody id="tbodyModalClientes"></tbody>
              </table>
            </div>
          </div>
          <div id="placeholderModalClienteVazio" class="modal-cliente-placeholder d-none">
            <i class="bi bi-clipboard-x"></i>
            <p>Nenhum cliente encontrado.</p>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

  <script>
  (() => {
    "use strict";

    const REPOSITORIO   = "rlto_canhotoRepositorio.asp";
    const CAMINHO_LOGO  = "/corporativo/financeiro/canhoto/Logo/logo-smurfit1.png";
    const ID_USUARIO    = "<%=IdUsuarioLogado%>";
    const LOGIN_USUARIO = "<%=LoginUsuarioLogado%>";
    const TOAST_DELAY   = 4500;

    const COLS = Object.freeze([
      "Nota Fiscal", "Série", "Cliente", "Filial",
      "Data Canhoto", "Produto", "Serviço", "Usuário", "Observação"
    ]);

    const KEYS = Object.freeze([
      "nf", "serie", "cliente", "filial",
      "data", "carga", "veiculo", "usuario", "obs"
    ]);

    const TOAST_CFG = Object.freeze({
      sucesso: { icon: "bi-check-circle-fill",         title: "Sucesso"    },
      erro:    { icon: "bi-x-circle-fill",             title: "Erro"       },
      aviso:   { icon: "bi-exclamation-triangle-fill", title: "Atenção"    },
      info:    { icon: "bi-info-circle-fill",          title: "Informação" }
    });

    let _todos              = [];
    let _filtrados          = [];
    let _paginaAtual        = 1;
    let _limite             = 50;
    let _clientesConhecidos = new Set();
    let _debounceTimer      = null;
    let _xhrCliente         = null;


    const $id = (id) => document.getElementById(id);

    const exibirOverlay  = () => $id("overlayLoading").classList.remove("d-none");
    const ocultarOverlay = () => $id("overlayLoading").classList.add("d-none");

    function exibirToast(msg, tipo) {
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
          <button class="toast-close" onclick="this.closest('.toast-custom').remove()" aria-label="Fechar">&times;</button>
        </div>
        <div class="toast-progress"><div class="toast-progress-bar"></div></div>`;
      $id("toastMsg").appendChild(el);
      setTimeout(() => { if (el.parentNode) el.remove(); }, TOAST_DELAY);
    }

    function validar() {
      const di = $id("dtini").value;
      const df = $id("dtfin").value;
      if (!di) { exibirToast("Informe a data inicial.", "aviso"); return false; }
      if (!df) { exibirToast("Informe a data final.", "aviso"); return false; }
      if (new Date(di) > new Date(df)) {
        exibirToast("A data inicial deve ser menor que a data final.", "aviso");
        return false;
      }
      return true;
    }

    function filtrar() {
      if (!validar()) return;
      exibirOverlay();

      const di      = $id("dtini").value;
      const df      = $id("dtfin").value;
      const filial  = $id("filial").value;
      const carga   = $id("carga").value;
      const veiculo = $id("veiculo").value;
      const nf      = $id("nf").value.trim();
      const cliente = $id("cliente").value.trim();

      const params = new URLSearchParams({ acao: "FILTRAR", dtini: di, dtfin: df, filial, carga, veiculo, nf, cliente });

      fetch(REPOSITORIO, {
        method:  "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body:    params.toString()
      })
        .then((res) => {
          if (!res.ok) {
            ocultarOverlay();
            exibirToast(`Erro ${res.status} na comunicação.`, "erro");
            return null;
          }
          return res.json();
        })
        .then((dados) => {
          if (!dados) return;
          ocultarOverlay();
          if (dados.erro) { exibirToast(dados.erro, "erro"); return; }
          if (dados.length === 0) {
            exibirToast("Nenhum registro encontrado.", "aviso");
            mostrarPlaceholder();
            return;
          }
          _todos       = dados;
          _filtrados   = structuredClone(dados);
          _paginaAtual = 1;
          _limite      = parseInt($id("selectLimite").value, 10);

          for (const r of dados) {
            const label = (r.cliente ?? "").trim();
            if (label !== "") _clientesConhecidos.add(label);
          }

          exibirTabela();
          exibirToast(`${_todos.length} registro(s) encontrado(s).`, "sucesso");

          const lnk = $id("btnGrafico");
          if (lnk) {
            lnk.href = `graf_canhoto.asp?dtini=${encodeURIComponent(di)}&dtfin=${encodeURIComponent(df)}&filial=${encodeURIComponent(filial)}`;
          }
        })
        .catch(() => {
          ocultarOverlay();
          exibirToast("Falha de comunicação com o servidor.", "erro");
        });
    }

    function limpar() {
      $id("dtini").value          = "";
      $id("dtfin").value          = "";
      $id("nf").value             = "";
      $id("cliente").value        = "";
      $id("clienteDisplay").value = "";
      $id("loja").value           = "";
      $id("inputBusca").value     = "";
      $id("filial").value         = "";
      $id("carga").value          = "";
      $id("veiculo").value        = "";
      _limparDropdown("btnCarga");
      _limparDropdown("btnVeiculo");

      _todos     = [];
      _filtrados = [];
      mostrarPlaceholder();
    }

    const mostrarPlaceholder = () => {
      $id("placeholderTabela").classList.remove("d-none");
      $id("areaTabela").classList.add("d-none");
    };

    const exibirTabela = () => {
      $id("placeholderTabela").classList.add("d-none");
      $id("areaTabela").classList.remove("d-none");
      renderizarPagina();
    };

    function construirLinha(r) {
      const tr = document.createElement("tr");
      for (const k of KEYS) {
        const td = document.createElement("td");
        td.textContent = r[k] ?? "";
        tr.appendChild(td);
      }
      return tr;
    }

    function renderizarPagina() {
      const tbody  = $id("tbodyDados");
      tbody.innerHTML = "";
      const inicio = (_paginaAtual - 1) * _limite;
      const fim    = Math.min(inicio + _limite, _filtrados.length);
      const frag   = document.createDocumentFragment();
      for (let i = inicio; i < fim; i++) frag.appendChild(construirLinha(_filtrados[i]));
      tbody.appendChild(frag);
      atualizarInfo(inicio + 1, fim);
      renderizarPaginacao();
    }

    function atualizarInfo(de, ate) {
      const total = _filtrados.length;
      $id("infoRegistros").textContent = total > 0
        ? `Exibindo ${de}–${ate} de ${total} registros`
        : "Nenhum registro";
    }

    function renderizarPaginacao() {
      const ul    = $id("paginacao");
      ul.innerHTML = "";
      const total = Math.ceil(_filtrados.length / _limite);
      if (total <= 1) return;

      const criarItem = (label, pagina, desabilitado, ativo) => {
        const li = document.createElement("li");
        li.className = `page-item${desabilitado ? " disabled" : ""}${ativo ? " active" : ""}`;
        const a = document.createElement("a");
        a.className = "page-link";
        a.href      = "#";
        a.innerHTML = label;
        a.addEventListener("click", (e) => {
          e.preventDefault();
          if (!desabilitado) { _paginaAtual = pagina; renderizarPagina(); }
        });
        li.appendChild(a);
        return li;
      };

      ul.appendChild(criarItem("&laquo;",  1,                   _paginaAtual === 1,     false));
      ul.appendChild(criarItem("&lsaquo;", _paginaAtual - 1,    _paginaAtual === 1,     false));

      const inicio = Math.max(1, _paginaAtual - 2);
      const fim    = Math.min(total, _paginaAtual + 2);
      for (let p = inicio; p <= fim; p++) ul.appendChild(criarItem(p, p, false, p === _paginaAtual));

      ul.appendChild(criarItem("&rsaquo;", _paginaAtual + 1, _paginaAtual === total, false));
      ul.appendChild(criarItem("&raquo;",  total,            _paginaAtual === total, false));
    }

    function aplicarBusca() {
      const termo = ($id("inputBusca").value ?? "").toLowerCase().trim();
      _filtrados = termo === ""
        ? structuredClone(_todos)
        : _todos.filter((r) => KEYS.some((k) => (r[k] ?? "").toString().toLowerCase().includes(termo)));
      _paginaAtual = 1;
      renderizarPagina();
    }

    function exportarExcel() {
      if (_filtrados.length === 0) { exibirToast("Nenhum dado para exportar.", "aviso"); return; }
      const linhas = _filtrados.map((r) => KEYS.map((k) => r[k] ?? ""));
      const ws     = XLSX.utils.aoa_to_sheet([Array.from(COLS), ...linhas]);
      const wb     = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Canhotos");
      XLSX.writeFile(wb, "relatorio-canhoto.xlsx");
    }

    function exportarPdf() {
      if (_filtrados.length === 0) { exibirToast("Nenhum dado para exportar.", "aviso"); return; }
      exibirOverlay();

      const { jsPDF } = window.jspdf;
      const doc   = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      const pageW = doc.internal.pageSize.getWidth();

      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => {
        const canvas  = document.createElement("canvas");
        canvas.width  = 900;
        canvas.height = Math.round(900 * (img.height / img.width));
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        doc.addImage(canvas.toDataURL("image/png"), "PNG", 10, 4, 38, Math.round(38 * (img.height / img.width)));
        _gerarAutoTable(doc, pageW);
      };
      img.onerror = () => _gerarAutoTable(doc, pageW);
      img.src     = CAMINHO_LOGO;
    }

    function _gerarAutoTable(doc, pageW) {
      const di = $id("dtini").value;
      const df = $id("dtfin").value;

      doc.setFontSize(13);
      doc.setFont(undefined, "bold");
      doc.text("Relatório de Canhotos", pageW / 2, 14, { align: "center" });

      doc.setFontSize(8);
      doc.setFont(undefined, "normal");
      doc.text(`Período: ${di} a ${df} — ${_filtrados.length} registro(s)`, pageW / 2, 20, { align: "center" });

      doc.setFontSize(7);
      doc.text(`Emitido por: ${LOGIN_USUARIO} (${ID_USUARIO})`, pageW - 10, 8, { align: "right" });

      doc.autoTable({
        head:               [Array.from(COLS)],
        body:               _filtrados.map((r) => KEYS.map((k) => r[k] ?? "")),
        startY:             26,
        theme:              "grid",
        headStyles:         { fillColor: [74, 74, 74], textColor: [255, 255, 255] },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        styles:             { fontSize: 4.5, cellPadding: 1.5 },
        tableLineColor:     [180, 180, 180],
        tableLineWidth:     0.2
      });

      doc.save("relatorio-canhoto.pdf");
      ocultarOverlay();
    }

    function _abrirModalCliente() {
      const inp = $id("inputBuscaCliente");
      inp.value = "";

      // Estado inicial: aguardando digitação
      $id("placeholderModalClienteSemDados").classList.remove("d-none");
      $id("placeholderModalClienteVazio").classList.add("d-none");
      $id("areaModalTabelaCliente").classList.add("d-none");
      $id("tbodyModalClientes").innerHTML = "";
      $id("modalClienteCount").textContent = "";

      // Foco automático para o usuário já começar a digitar
      setTimeout(() => inp.focus(), 150);
    }

    function _buscarClientesRemoto(termo) {
      // Aborta XHR anterior se ainda estiver em voo
      if (_xhrCliente) { _xhrCliente.abort(); _xhrCliente = null; }

      const elSem  = $id("placeholderModalClienteSemDados");
      const elVaz  = $id("placeholderModalClienteVazio");
      const elArea = $id("areaModalTabelaCliente");
      const tbody  = $id("tbodyModalClientes");
      const count  = $id("modalClienteCount");

      // Feedback visual de carregamento
      elSem.classList.remove("d-none");
      elSem.querySelector("p").textContent = "Buscando...";
      elArea.classList.add("d-none");
      elVaz.classList.add("d-none");
      tbody.innerHTML = "";
      count.textContent = "";

      const xhr = new XMLHttpRequest();
      _xhrCliente = xhr;
      xhr.open("POST", REPOSITORIO, true);
      xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        _xhrCliente = null;

        // Restaura texto padrão do placeholder
        elSem.querySelector("p").textContent = "Digite para buscar clientes.";

        if (xhr.status === 0) return; // abortado
        elSem.classList.add("d-none");

        if (xhr.status !== 200) { elVaz.classList.remove("d-none"); return; }
        let dados;
        try { dados = JSON.parse(xhr.responseText); } catch (e) { elVaz.classList.remove("d-none"); return; }
        if (!Array.isArray(dados) || dados.length === 0) { elVaz.classList.remove("d-none"); return; }

        // Acumula no cache para reuso local
        for (const label of dados) {
          if (label && label.trim() !== "") _clientesConhecidos.add(label.trim());
        }

        _renderizarModalClientes(dados);
      };
      xhr.send("acao=CLIENTES&q=" + encodeURIComponent(termo));
    }

    function _renderizarModalClientes(lista) {
      const areaTabela = $id("areaModalTabelaCliente");
      const vazio      = $id("placeholderModalClienteVazio");
      const tbody      = $id("tbodyModalClientes");
      const count      = $id("modalClienteCount");

      tbody.innerHTML = "";

      if (!lista || lista.length === 0) {
        areaTabela.classList.add("d-none");
        vazio.classList.remove("d-none");
        count.textContent = "";
        return;
      }

      vazio.classList.add("d-none");
      areaTabela.classList.remove("d-none");
      count.textContent = `${lista.length} cliente(s) encontrado(s)${lista.length === 50 ? " (máx. 50)" : ""}`;

      const frag = document.createDocumentFragment();
      for (const label of lista) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td class="text-center">
            <button class="btn btn-outline-primary btn-importar" title="Importar cliente">
              <i class="bi bi-box-arrow-in-down"></i>
            </button>
          </td>
          <td>${label}</td>`;
        tr.querySelector(".btn-importar").addEventListener("click", (e) => {
          e.stopPropagation();
          _selecionarCliente(label);
        });
        tr.addEventListener("click", () => _selecionarCliente(label));
        frag.appendChild(tr);
      }
      tbody.appendChild(frag);
    }

    function _onInputBuscaCliente() {
      const termo = $id("inputBuscaCliente").value.trim();

      if (termo.length < 2) {
        // Abaixo do mínimo — cancela requisição pendente e volta ao estado inicial
        if (_xhrCliente) { _xhrCliente.abort(); _xhrCliente = null; }
        if (_debounceTimer) { clearTimeout(_debounceTimer); _debounceTimer = null; }
        $id("placeholderModalClienteSemDados").classList.remove("d-none");
        $id("placeholderModalClienteSemDados").querySelector("p").textContent = "Digite ao menos 2 caracteres para buscar.";
        $id("areaModalTabelaCliente").classList.add("d-none");
        $id("placeholderModalClienteVazio").classList.add("d-none");
        $id("tbodyModalClientes").innerHTML = "";
        $id("modalClienteCount").textContent = "";
        return;
      }

      // Debounce de 350 ms para não disparar a cada tecla
      if (_debounceTimer) clearTimeout(_debounceTimer);
      _debounceTimer = setTimeout(() => {
        _debounceTimer = null;
        _buscarClientesRemoto(termo);
      }, 350);
    }

    function _selecionarCliente(label) {
      const match  = label.match(/^\[([^\]]+)\]/);
      const codigo = match ? match[1].trim() : label.trim();
      $id("clienteDisplay").value = label;
      $id("cliente").value        = codigo;
      bootstrap.Modal.getInstance($id("modalBuscaCliente"))?.hide();
      exibirToast(`Cliente "${label}" selecionado.`, "info");
    }

    /* ── Dropdown multiselect helpers ────────────────────────────────────── */
    function _iniciarDropdown(btnId, hiddenId) {
      const btn    = $id(btnId);
      const hidden = $id(hiddenId);
      const items  = btn.closest(".dropdown").querySelectorAll(".dropdown-item");

      items.forEach((item) => {
        item.addEventListener("click", () => {
          const cb = item.querySelector("input[type='checkbox']");
          cb.checked = !cb.checked;

          const marcados = [...items].filter((i) => i.querySelector("input").checked);
          hidden.value   = marcados.map((i) => i.dataset.val).join(",");

          if (marcados.length === 0) {
            btn.textContent = "Selecione...";
          } else if (marcados.length === 1) {
            btn.textContent = marcados[0].querySelector("span").textContent;
          } else {
            btn.textContent = `${marcados[0].querySelector("span").textContent} +${marcados.length - 1}`;
          }
        });
      });
    }

    function _limparDropdown(btnId) {
      const btn   = $id(btnId);
      const items = btn.closest(".dropdown").querySelectorAll(".dropdown-item");
      items.forEach((item) => {
        item.querySelector("input[type='checkbox']").checked = false;
      });
      btn.textContent = "Selecione...";
    }

    function init() {
      _iniciarDropdown("btnCarga",   "carga");
      _iniciarDropdown("btnVeiculo", "veiculo");

      $id("btnFiltrar").addEventListener("click", filtrar);
      $id("btnLimpar").addEventListener("click", limpar);
      $id("btnExcel").addEventListener("click", exportarExcel);
      $id("btnPdf").addEventListener("click", exportarPdf);

      $id("selectLimite").addEventListener("change", () => {
        _limite      = parseInt($id("selectLimite").value, 10);
        _paginaAtual = 1;
        renderizarPagina();
      });

      $id("inputBusca").addEventListener("input", aplicarBusca);

      document.addEventListener("keydown", (e) => {
        if (e.ctrlKey && e.key === "k") {
          e.preventDefault();
          const inp = $id("inputBusca");
          inp.focus();
          inp.select();
        }
      });

      $id("inputBuscaCliente").addEventListener("input", _onInputBuscaCliente);

      $id("modalBuscaCliente").addEventListener("show.bs.modal", _abrirModalCliente);

      $id("modalBuscaCliente").addEventListener("hidden.bs.modal", () => {
        $id("inputBuscaCliente").value = "";
        if (_debounceTimer) { clearTimeout(_debounceTimer); _debounceTimer = null; }
        if (_xhrCliente)    { _xhrCliente.abort(); _xhrCliente = null; }
      });
    }

    document.addEventListener("DOMContentLoaded", init);

  })();
  </script>

</body>
</html>
