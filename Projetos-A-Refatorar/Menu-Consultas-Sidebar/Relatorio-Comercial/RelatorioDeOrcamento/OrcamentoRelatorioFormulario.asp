<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet - Relatório de Orçamentos</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />
  <style>
    body {
      background: #f0f0f0;
      color: #212529;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }

    /* ── Filter Card ── */
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
      font-size: 0.82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: 0.25rem;
    }

    .form-control:focus,
    .form-select:focus {
      border-color: #4a4a4a;
      box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
    }

    /* ── Table Card ── */
    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      position: relative;
    }

    /* ── Toolbar ── */
    .table-toolbar {
      background-color: #fff;
      border-radius: 6px 6px 0 0;
    }

    /* ── Search Box ── */
    .search-box { position: relative; max-width: 200px; }
    .search-box .form-control { padding-right: 76px; font-size: .82rem; }
    .search-box .search-hotkey {
      position: absolute;
      right: 8px;
      top: 50%;
      transform: translateY(-50%);
      pointer-events: none;
      user-select: none;
      display: flex;
      gap: 2px;
    }
    .search-box .search-hotkey kbd {
      font-size: .65rem;
      padding: .1rem .3rem;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      color: #6c757d;
    }

    /* ── Table Wrapper ── */
    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 68vh;
    }
    .table-wrapper table { min-width: max-content; }

    .table-wrapper thead th {
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

    .table-wrapper tbody td {
      font-size: .82rem;
      text-align: center;
      white-space: nowrap;
      color: #212529;
      vertical-align: middle;
      padding: .45rem .75rem;
    }

    /* ── Placeholder ── */
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
    .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

    /* ── Table Footer ── */
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

    /* ── Paginação (secondary) ── */
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

    /* ── Overlay de carregamento (escopo do card) ── */
    .overlay-loading {
      position: absolute;
      inset: 0;
      background: rgba(255, 255, 255, .80);
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

    /* ── Toast ── */
    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: 1100;
      display: flex; flex-direction: column; gap: .5rem;
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
    .toast-custom .toast-body-inner {
      display: flex;
      align-items: center;
      gap: .6rem;
      padding: .65rem .85rem;
      font-size: .85rem;
      color: #333;
    }
    .toast-custom .toast-progress-bar {
      height: 3px;
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; } to { width: 0%; }
    }
    .toast-fechar {
      margin-left: auto;
      background: none;
      border: none;
      font-size: 1rem;
      color: #999;
      cursor: pointer;
      padding: 0 0 0 .5rem;
    }
    .toast-fechar:hover { color: #333; }

    /* ── Badges ── */
    .badge-status,
    .planta-badge {
      display: inline-block;
      padding: 2px 7px;
      border-radius: 3px;
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.2px;
    }

    .badge-aprovado     { background: rgba(25, 135, 84, 0.12);   color: #146c43; }
    .badge-pendente     { background: rgba(255, 193, 7, 0.2);    color: #856404; }
    .badge-reprovado    { background: rgba(220, 53, 69, 0.12);   color: #b02a37; }
    .badge-digitacao    { background: rgba(13, 202, 240, 0.15);  color: #087990; }
    .badge-analise      { background: rgba(102, 16, 242, 0.1);   color: #5a12c0; }
    .badge-aberto       { background: rgba(13, 110, 253, 0.1);   color: #0a58ca; }
    .badge-cancelado    { background: rgba(108, 117, 125, 0.12); color: #495057; }
    .badge-em-aprovacao { background: rgba(255, 193, 7, 0.2);    color: #856404; }

    .valor-positivo { color: #198754; font-weight: 600; }
    .valor-negativo { color: #dc3545; font-weight: 600; }

    .planta-pirapetinga { background: #fff3cd; color: #856404; }
    .planta-uberaba     { background: #d1f0f5; color: #0c5460; }
    .planta-bento       { background: #e8e8e8; color: #2d3436; }
    .planta-maranguape  { background: #fcd9d0; color: #721c24; }
    .planta-saquarema   { background: #e2deff; color: #2d1b69; }

    .dropdown-acoes-menu { min-width: 140px; font-size: 0.82rem; }

    .btn-acoes {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 28px;
      height: 24px;
      padding: 0;
      border: none;
      border-radius: 4px;
      background: var(--bs-primary);
      color: #fff;
      font-size: 0.85rem;
      transition: background 0.15s;
    }
    .btn-acoes:hover { background: var(--bs-primary-hover, #0b5ed7); }

    .log-secao-header {
      background: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 4px 4px 0 0;
      padding: 0.5rem 0.75rem;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      font-size: 0.82rem;
      font-weight: 600;
      color: #495057;
    }
    .log-secao-body {
      border: 1px solid #dee2e6;
      border-top: none;
      border-radius: 0 0 4px 4px;
      overflow: auto;
      max-height: 260px;
    }
    .log-secao-body table { margin-bottom: 0; }
    .log-secao-body thead th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #f8f9fa;
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.2px;
      white-space: nowrap;
      text-align: center;
    }
    .log-secao-body tbody td { font-size: 0.78rem; vertical-align: middle; }
    .log-vazio   { text-align: center; padding: 1.5rem; color: #888; font-size: 0.85rem; }
    .log-loading { text-align: center; padding: 1.5rem; color: #6c757d; font-size: 0.85rem; }

    .nivel-badge {
      display: inline-block;
      padding: 2px 7px;
      border-radius: 3px;
      font-size: 0.72rem;
      font-weight: 600;
      background: rgba(13, 110, 253, 0.1);
      color: #0a58ca;
    }

    /* ── Modais de busca ── */
    .modal-busca .modal-header {
      background: #4a4a4a;
      color: #fff;
      padding: 0.6rem 1rem;
    }
    .modal-busca .modal-header .btn-close { filter: invert(1) grayscale(1); }

    .modal-busca .tbl-busca-wrapper {
      overflow: auto;
      max-height: 55vh;
      border: 1px solid #dee2e6;
      border-radius: 4px;
    }

    .modal-busca .tbl-busca-wrapper thead th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #4a4a4a;
      color: #fff;
      font-size: 0.73rem;
      font-weight: 600;
      white-space: nowrap;
      vertical-align: middle;
      text-align: center;
      padding: 0.5rem 0.5rem;
      border-color: #5a5a5a;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    .modal-busca .tbl-busca-wrapper tbody td {
      font-size: 0.8rem;
      vertical-align: middle;
      padding: 0.4rem 0.5rem;
      border-color: #e8e8e8;
      white-space: nowrap;
    }

    .modal-busca .tbl-busca-wrapper tbody tr:hover td {
      background-color: #f0f4ff !important;
      cursor: pointer;
    }

    .modal-busca .tbl-busca-wrapper tbody tr.tr-alerta td {
      background-color: #fff8e1 !important;
      color: #7d5a00 !important;
    }
    .modal-busca .tbl-busca-wrapper tbody tr.tr-alerta:hover td {
      background-color: #fff3cd !important;
    }

    .btn-selecionar-busca {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 26px;
      height: 26px;
      padding: 0;
      border-radius: 4px;
      background: #0d6efd;
      color: #fff;
      border: none;
      cursor: pointer;
      font-size: 0.82rem;
      transition: background 0.15s, transform 0.1s;
    }
    .btn-selecionar-busca:hover  { background: #0b5ed7; transform: translateY(-1px); }
    .btn-selecionar-busca:active { transform: scale(0.95); }

    .busca-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3.5rem 1rem;
      color: #adb5bd;
      gap: 0.6rem;
    }
    .busca-placeholder .bi  { font-size: 2.8rem; opacity: 0.4; }
    .busca-placeholder p    { margin: 0; font-size: 0.85rem; color: #6c757d; text-align: center; }
  </style>
</head>
<body>

  <div class="toast-container-custom" id="toastWrapper"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <h4 class="fw-bold mb-3 text-start">
      <span class="page-title">Relatório de Orçamentos</span>
    </h4>

    <!-- Card de Filtros -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">

        <!-- Linha 1: Planta · Nº Orçamento · Status -->
        <div class="row g-2 align-items-end mb-2">

          <div class="col-auto" style="min-width:140px;">
            <label class="form-label" for="filtroPlanta">Planta <span class="text-danger">*</span></label>
            <select id="filtroPlanta" class="form-select form-select-sm"></select>
          </div>

          <div class="col-auto" style="min-width:120px;">
            <label class="form-label" for="filtroNrOrcamento">Nº Orçamento</label>
            <input id="filtroNrOrcamento" type="text" class="form-control form-control-sm" placeholder="Código..." />
          </div>

          <div class="col-auto" style="min-width:160px;">
            <label class="form-label" for="filtroStatus">Status</label>
            <select id="filtroStatus" class="form-select form-select-sm">
              <option value="">Todos</option>
            </select>
          </div>

        </div>

        <!-- Linha 2: Cliente · Produto · Data De · Data Até · Botões -->
        <div class="row g-2 align-items-end">

          <!-- Campo Cliente -->
          <div class="col-auto" style="min-width:300px;">
            <label class="form-label">Cliente</label>
            <div class="input-group input-group-sm">
              <button type="button" class="btn btn-secondary btn-sm" id="btnAbrirBuscaCliente" title="Buscar cliente">
                Buscar <i class="bi bi-box-arrow-up-right ms-1"></i>
              </button>
              <input id="filtroCliente" type="text" class="form-control form-control-sm"
                     placeholder="Código ou nome" autocomplete="off" style="max-width:130px;" />
              <input id="filtroClienteLojaVisivel" type="text" class="form-control form-control-sm"
                     placeholder="Loja" autocomplete="off" style="max-width:55px;" title="Loja (opcional)" />
              <button type="button" class="btn btn-outline-secondary btn-sm" id="btnLimparCliente" title="Limpar cliente">
                <i class="bi bi-x-lg"></i>
              </button>
            </div>
            <input type="hidden" id="filtroClienteCod" />
            <input type="hidden" id="filtroClienteLoja" />
          </div>

          <!-- Campo Produto -->
          <div class="col-auto" style="min-width:220px;">
            <label class="form-label">Produto / Referência</label>
            <div class="input-group input-group-sm">
              <button type="button" class="btn btn-secondary btn-sm" id="btnAbrirBuscaProduto" title="Buscar produto">
                Buscar <i class="bi bi-box-arrow-up-right ms-1"></i>
              </button>
              <input id="filtroProduto" type="text" class="form-control form-control-sm"
                     placeholder="Código ou descrição" autocomplete="off" />
              <button type="button" class="btn btn-outline-secondary btn-sm" id="btnLimparProduto" title="Limpar produto">
                <i class="bi bi-x-lg"></i>
              </button>
            </div>
            <input type="hidden" id="filtroProdutoCod" />
          </div>

          <div class="col-lg-1 col-md-4 col-sm-6" style="min-width:150px;">
            <label class="form-label" for="filtroDataDe">Data Inclusão De <span class="text-danger">*</span></label>
            <input id="filtroDataDe" type="date" class="form-control form-control-sm" />
          </div>

          <div class="col-lg-1 col-md-4 col-sm-6" style="min-width:150px;">
            <label class="form-label" for="filtroDataAte">Data Inclusão Até <span class="text-danger">*</span></label>
            <input id="filtroDataAte" type="date" class="form-control form-control-sm" />
          </div>

          <!-- Botões alinhados à direita -->
          <div class="col-lg-auto ms-auto d-flex align-items-end gap-2">
            <button type="button" class="btn btn-light btn-sm" id="btnLimpar">
              <i class="bi bi-eraser"></i>
            </button>
            <button type="button" class="btn btn-primary btn-sm" id="btnFiltrar">
              <i class="bi bi-funnel-fill me-1"></i>Filtrar
            </button>
          </div>

        </div>

      </div>
    </div>

    <!-- Card de Tabela -->
    <div class="table-card">

      <div id="overlayLoading" class="overlay-loading d-none">
        <div class="spinner-border" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <!-- Toolbar: exportações (esq) + search box (dir) -->
      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          <button id="btnExportCsv" type="button" class="btn btn-success btn-sm">
            <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
          </button>
          <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm">
            <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
          </button>
        </div>
        <div class="search-box">
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <div class="table-wrapper" id="tblWrapper">
        <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaOrcamentos">
          <thead>
            <tr>
              <th>Nº</th>
              <th>Planta</th>
              <th>Nº Orçamento</th>
              <th>Cliente</th>
              <th>CNPJ</th>
              <th>Cidade/UF</th>
              <th>Representante</th>
              <th>Frete</th>
              <th>Cond. Pagto</th>
              <th>Data<br />Inclusão</th>
              <th>Validade<br />De</th>
              <th>Validade<br />Até</th>
              <th>Produto</th>
              <th>Referência</th>
              <th>Vlr. Unit.</th>
              <th>Status</th>
              <th>Margem</th>
              <th>Ações</th>
            </tr>
          </thead>
          <tbody id="tblCorpo">
            <tr>
              <td colspan="18" class="p-0">
                <div class="table-placeholder">
                  <i class="bi bi-clipboard-data"></i>
                  <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="table-footer d-none" id="rodapeTabela">

        <div class="table-footer-left">
          <label for="pgTamanho" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">
            Linhas por página:
          </label>
          <select id="pgTamanho" class="form-select form-select-sm select-limite">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50" selected>50</option>
            <option value="100">100</option>
            <option value="200">200</option>
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

  <!-- MODAL: Detalhes do Orçamento -->
  <div class="modal fade" id="modalDetalhes" tabindex="-1" aria-labelledby="modalDetalhesLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header bg-white border-bottom">
          <h6 class="modal-title text-dark" id="modalDetalhesLabel">
            <i class="bi bi-info-circle me-2 text-secondary"></i>Detalhes do Orçamento
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body" id="modalDetalhesBody"></div>
        <div class="modal-footer bg-white border-top">
          <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
        </div>
      </div>
    </div>
  </div>

  <!-- MODAL: Busca de Cliente (SPA) -->
  <div class="modal fade modal-busca" id="modalBuscaCliente" tabindex="-1"
       aria-labelledby="modalBuscaClienteLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header">
          <h6 class="modal-title" id="modalBuscaClienteLabel">
            <i class="bi bi-person-lines-fill me-2"></i>Consulta de Cliente
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <div class="d-flex gap-2 mb-3">
            <input type="text" id="inputBuscaCliente"
                   class="form-control form-control-sm flex-grow-1"
                   placeholder="Nome ou código do cliente..."
                   autocomplete="off" />
            <button type="button" class="btn btn-primary btn-sm flex-shrink-0" id="btnExecutarBuscaCliente">
              <i class="bi bi-funnel-fill me-1"></i>Filtrar
            </button>
          </div>
          <div class="tbl-busca-wrapper" id="tblBuscaClienteWrapper">
            <div class="busca-placeholder">
              <i class="bi bi-clipboard-data"></i>
              <p>Digite o nome ou código e clique em <strong>Filtrar</strong>.</p>
            </div>
          </div>
        </div>
        <div class="modal-footer bg-white border-top">
          <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
        </div>
      </div>
    </div>
  </div>

  <!-- MODAL: Busca de Produto (SPA) -->
  <div class="modal fade modal-busca" id="modalBuscaProduto" tabindex="-1"
       aria-labelledby="modalBuscaProdutoLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header">
          <h6 class="modal-title" id="modalBuscaProdutoLabel">
            <i class="bi bi-box-seam me-2"></i>Consulta de Produto
          </h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <div class="d-flex gap-2 mb-3">
            <input type="text" id="inputBuscaProduto"
                   class="form-control form-control-sm flex-grow-1"
                   placeholder="Descrição ou código do produto..."
                   autocomplete="off" />
            <button type="button" class="btn btn-primary btn-sm flex-shrink-0" id="btnExecutarBuscaProduto">
              <i class="bi bi-funnel-fill me-1"></i>Filtrar
            </button>
          </div>
          <div class="tbl-busca-wrapper" id="tblBuscaProdutoWrapper">
            <div class="busca-placeholder">
              <i class="bi bi-clipboard-data"></i>
              <p>Digite a descrição ou código e clique em <strong>Filtrar</strong>.</p>
            </div>
          </div>
        </div>
        <div class="modal-footer bg-white border-top">
          <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>
  <script>
    var IdUsuarioLogado    = "<%= Session("CODIGO") %>";
    var LoginUsuarioLogado = "<%= Session("USR") %>";

    (($) => {
      "use strict";

      const REPO     = "OrcamentoRelatorioRepositorio.asp";
      const LOGO_URL = "Logo/logo-smurfit1.png";
      const TOAST_DELAY = 4500;

      let _pgTamanho    = 50;
      let _pgAtual      = 1;
      let _dados        = [];
      let _dadosFiltrados = [];
      let _login        = "";
      let _logoB64      = "";

      const plantaAtual = () => $("#filtroPlanta").val() || "";

      const STATUS_MAP = Object.freeze({
        "APROVADO":        "badge-aprovado",
        "APROVADO TOTAL":  "badge-aprovado",
        "CANCELADO":       "badge-cancelado",
        "CANCELADA":       "badge-cancelado",
        "EM ANALISE":      "badge-analise",
        "EM ANÁLISE":      "badge-analise",
        "EM APROVACAO":    "badge-em-aprovacao",
        "EM APROVAÇÃO":    "badge-em-aprovacao",
        "EM DIGITACAO":    "badge-digitacao",
        "EM DIGITAÇÃO":    "badge-digitacao",
        "REPROVADO":       "badge-reprovado",
      });

      const TOAST_ICON = Object.freeze({
        sucesso: { cls: "bi-check-circle-fill",         cor: "#198754" },
        erro:    { cls: "bi-x-circle-fill",             cor: "#dc3545" },
        aviso:   { cls: "bi-exclamation-triangle-fill", cor: "#856404" },
        info:    { cls: "bi-info-circle-fill",          cor: "#6c757d" },
        /* aliases legados */
        success: { cls: "bi-check-circle-fill",         cor: "#198754" },
        danger:  { cls: "bi-x-circle-fill",             cor: "#dc3545" },
        warning: { cls: "bi-exclamation-triangle-fill", cor: "#856404" },
      });

      const PLANTA_MAP = Object.freeze({
        "pirapetinga":     "planta-pirapetinga",
        "uberaba":         "planta-uberaba",
        "bento gonçalves": "planta-bento",
        "bento":           "planta-bento",
        "maranguape":      "planta-maranguape",
        "saquarema":       "planta-saquarema",
      });

      const COLUNAS_EXPORT = Object.freeze([
        { key: "planta",          label: "Planta"        },
        { key: "nrOrcamento",     label: "Nº Orçamento"  },
        { key: "cliente",         label: "Cliente"       },
        { key: "cnpj",            label: "CNPJ"          },
        { key: "cidade",          label: "Cidade"        },
        { key: "uf",              label: "UF"            },
        { key: "representante",   label: "Representante" },
        { key: "tipoFrete",       label: "Frete"         },
        { key: "descricao",       label: "Cond. Pagto"   },
        { key: "dataInclusao",    label: "Data Inclusão" },
        { key: "validadeDe",      label: "Validade De"   },
        { key: "validadeAte",     label: "Validade Até"  },
        { key: "produto",         label: "Produto"       },
        { key: "referencia",      label: "Referência"    },
        { key: "valorUnitario",   label: "Vlr. Unit."    },
        { key: "statusOrcamento", label: "Status"        },
        { key: "margem",          label: "Margem"        },
      ]);

      const Html = Object.freeze({
        esc(txt) {
          if (!txt) return "";
          return String(txt)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
        },
        truncar(texto, tamanho) {
          if (!texto) return "-";
          return texto.length <= tamanho ? texto : `${texto.substring(0, tamanho)}…`;
        },
      });

      const Formato = Object.freeze({
        cnpj(v) {
          if (!v) return "-";
          const s = String(v).replace(/\D/g, "");
          return s.length === 14
            ? s.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, "$1.$2.$3/$4-$5")
            : s;
        },
        moeda(v) {
          if (v == null) return "-";
          const n = parseFloat(v);
          if (isNaN(n)) return "-";
          return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2, maximumFractionDigits: 5 });
        },
        porcentagem(v) {
          if (v == null || v === "") return "-";
          const n = parseFloat(String(v).replace(",", "."));
          if (isNaN(n)) return "-";
          return `${n.toFixed(2).replace(".", ",")}%`;
        },
        nomeArquivo(ext) {
          const d    = new Date();
          const dd   = String(d.getDate()).padStart(2, "0");
          const mm   = String(d.getMonth() + 1).padStart(2, "0");
          const aaaa = d.getFullYear();
          return `OrcamentoRelatorio_${aaaa}${mm}${dd}.${ext}`;
        },
        statusClasse(s) {
          if (!s) return "";
          return STATUS_MAP[s.toUpperCase().trim()] ?? "";
        },
        plantaClasse(p) {
          if (!p) return "";
          return PLANTA_MAP[p.toLowerCase()] ?? "";
        },
        /* CSV helpers */
        dataCSV(v) {
          if (!v) return "";
          /* Formato vindo do repositório: "dd/mm/aaaa HH:MM" */
          return String(v);
        },
        numCSV(v, casas = 2) {
          if (v == null || v === "") return "";
          return parseFloat(v).toFixed(casas).replace(".", ",");
        },
      });

      /* ════════════════════════════════════════════════════════
         UI helpers
      ════════════════════════════════════════════════════════ */
      const UI = Object.freeze({
        exibirLoader()  { $("#overlayLoading").removeClass("d-none"); },
        ocultarLoader() { $("#overlayLoading").addClass("d-none"); },
        notificar(mensagem, tipo = "info") {
          const cfg  = TOAST_ICON[tipo] ?? TOAST_ICON.info;
          const item = $(`
            <div class="toast-custom">
              <div class="toast-body-inner">
                <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
                <span class="flex-grow-1">${mensagem}</span>
                <button class="toast-fechar" aria-label="Fechar">&times;</button>
              </div>
              <div class="toast-progress-bar" style="background:${cfg.cor};animation-duration:${TOAST_DELAY}ms;"></div>
            </div>
          `);
          item.find(".toast-fechar").on("click", () => item.remove());
          $("#toastWrapper").append(item);
          setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_DELAY);
        },
      });

      /* ════════════════════════════════════════════════════════
         Filtro
      ════════════════════════════════════════════════════════ */
      const Filtro = Object.freeze({
        /*
         * Estratégias de filtro de cliente:
         *
         * A) Modal selecionou (filtroClienteCod preenchido)
         *    → codcli exato + codloja exato (se tiver)
         *
         * B) Digitação livre — campo Código parece código numérico (só dígitos)
         *    → codcli exato + codloja exato (se campo loja preenchido)
         *
         * C) Digitação livre — campo Código contém texto (nome)
         *    → cliente LIKE (busca em NOME_CLIENTE e COD_CLIENTE)
         *    → codloja ignorada (LIKE não filtra loja)
         */
        _clienteParams() {
          const codHidden  = $("#filtroClienteCod").val().trim();
          const lojaHidden = $("#filtroClienteLoja").val().trim();
          const textoLivre = $("#filtroCliente").val().trim();
          const lojaLivre  = $("#filtroClienteLojaVisivel").val().trim();

          /* A) Veio do modal */
          if (codHidden) {
            return { codcli: codHidden, codloja: lojaHidden, cliente: "" };
          }

          if (!textoLivre) {
            return { codcli: "", codloja: "", cliente: "" };
          }

          /* B) Parece código numérico → busca exata, respeita loja */
          if (/^\d+$/.test(textoLivre)) {
            const lojaNorm = lojaLivre.length === 1 ? "0" + lojaLivre : lojaLivre;
            return { codcli: textoLivre, codloja: lojaNorm, cliente: "" };
          }

          /* C) Texto livre (nome) → LIKE no nome/código + loja exata se preenchida */
          const lojaNomeNorm = lojaLivre.length === 1 ? "0" + lojaLivre : lojaLivre;
          return { codcli: "", codloja: lojaNomeNorm, cliente: textoLivre };
        },

        valores() {
          const cli = this._clienteParams();
          const obj = {
            planta:      $("#filtroPlanta").val(),
            nrOrcamento: $("#filtroNrOrcamento").val().trim(),
            status:      $("#filtroStatus").val(),
            cliente:     cli.cliente,
            dataDe:      $("#filtroDataDe").val(),
            dataAte:     $("#filtroDataAte").val(),
            produto:     $("#filtroProdutoCod").val() || $("#filtroProduto").val().trim(),
          };
          /* Só envia codcli/codloja quando preenchidos — evita que ASP receba string vazia
             e ignore o filtro de loja por engano */
          if (cli.codcli)  obj.codcli  = cli.codcli;
          if (cli.codloja) obj.codloja = cli.codloja;
          return obj;
        },
        validar() {
          const f = this.valores();
          if (!f.nrOrcamento) {
            if (!f.planta) {
              UI.notificar("Selecione uma planta.", "warning");
              $("#filtroPlanta").trigger("focus");
              return false;
            }
            if (!f.dataDe) {
              UI.notificar("Informe a data inicial.", "warning");
              $("#filtroDataDe").trigger("focus");
              return false;
            }
          }
          return true;
        },
        limpar() {
          $("#filtroPlanta, #filtroStatus").val("");
          $("#filtroNrOrcamento, #filtroCliente, #filtroClienteLojaVisivel, #filtroClienteCod, #filtroClienteLoja").val("");
          $("#filtroProduto, #filtroProdutoCod").val("");
          $("#filtroDataDe, #filtroDataAte").val("");
          $("#searchBox").val("");
        },
      });

      /* ════════════════════════════════════════════════════════
         Paginação
      ════════════════════════════════════════════════════════ */
      const Paginacao = Object.freeze({
        totalPaginas() { return Math.max(1, Math.ceil(_dadosFiltrados.length / _pgTamanho)); },
        dadosPagina() {
          const inicio = (_pgAtual - 1) * _pgTamanho;
          return _dadosFiltrados.slice(inicio, inicio + _pgTamanho);
        },
        ir(pg) {
          _pgAtual = Math.max(1, Math.min(pg, this.totalPaginas()));
          Tabela.renderizarPagina();
          this.renderizarControles();
        },
        renderizarControles() {
          const total = this.totalPaginas();
          const atual = _pgAtual;
          const $ul   = $("#paginacao").empty();

          if (!_dadosFiltrados.length) { $("#rodapeTabela").addClass("d-none"); return; }

          $("#rodapeTabela").removeClass("d-none");
          $("#infoRegistros").text(`Total: ${_dadosFiltrados.length} registro(s)`);

          const mkItem = (label, page, disabled, active) => {
            const $li = $("<li>").addClass("page-item");
            if (disabled) $li.addClass("disabled");
            if (active)   $li.addClass("active");
            const $a = $("<a>").addClass("page-link").attr("href", "#").html(label);
            if (!disabled && !active) $a.on("click", e => { e.preventDefault(); Paginacao.ir(page); });
            return $li.append($a);
          };

          $ul.append(mkItem("&laquo;", 1,           atual === 1,     false));
          $ul.append(mkItem("&lsaquo;", atual - 1,  atual === 1,     false));

          const inicio = Math.max(1, atual - 2);
          const fim    = Math.min(total, atual + 2);
          for (let p = inicio; p <= fim; p++) {
            $ul.append(mkItem(p, p, false, p === atual));
          }

          $ul.append(mkItem("&rsaquo;", atual + 1, atual === total, false));
          $ul.append(mkItem("&raquo;",  total,     atual === total, false));
        },
      });

      /* ════════════════════════════════════════════════════════
         Tabela
      ════════════════════════════════════════════════════════ */
      const Tabela = Object.freeze({
        renderizar(dados) {
          _dados          = dados;
          _dadosFiltrados = dados;
          _pgAtual        = 1;
          $("#searchBox").val("");
          if (!_dados.length) {
            this.renderizarVazia();
            $("#rodapeTabela").addClass("d-none");
            return;
          }
          this.renderizarPagina();
          Paginacao.renderizarControles();
        },
        renderizarVazia() {
          $("#tblCorpo").html(`
            <tr>
              <td colspan="18" class="p-0">
                <div class="table-placeholder">
                  <i class="bi bi-clipboard-data"></i>
                  <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
                </div>
              </td>
            </tr>
          `);
        },
        renderizarPagina() {
          const linhas = Paginacao.dadosPagina().map((item, i) => {
            const seq       = (_pgAtual - 1) * _pgTamanho + i + 1;
            const idxAbs    = (_pgAtual - 1) * _pgTamanho + i;
            const statusCls = Formato.statusClasse(item.statusOrcamento);
            const plantaCls = Formato.plantaClasse(item.planta ?? "");
            const margem    = parseFloat(item.margem) || 0;
            const margemCls = margem >= 0 ? "valor-positivo" : "valor-negativo";
            const nrOrc     = Html.esc(item.nrOrcamento ?? "");
            const temCarta  = (item.statusOrcamento ?? "").toUpperCase().trim() === "APROVADO TOTAL";

            return `<tr>
              <td>${seq}</td>
              <td><span class="planta-badge ${plantaCls}">${Html.esc(item.planta) || "-"}</span></td>
              <td class="fw-semibold">${Html.esc(item.nrOrcamento) || "-"}</td>
              <td class="text-start" title="${Html.esc(item.cliente ?? "")}">${Html.truncar(item.cliente, 35)}</td>
              <td>${Formato.cnpj(item.cnpj)}</td>
              <td>${Html.esc(item.cidade) || "-"} / ${Html.esc(item.uf) || "-"}</td>
              <td class="text-start" title="${Html.esc(item.representante ?? "")}">${Html.truncar(item.representante, 25)}</td>
              <td>${Html.esc(item.tipoFrete) || "-"}</td>
              <td class="text-start" title="${Html.esc(item.descricao ?? "")}">${Html.truncar(item.descricao, 18)}</td>
              <td>${Html.esc(item.dataInclusao) || "-"}</td>
              <td>${Html.esc(item.validadeDe) || "-"}</td>
              <td>${Html.esc(item.validadeAte) || "-"}</td>
              <td class="text-start" title="${Html.esc(item.produto ?? "")}">${Html.truncar(item.produto, 18)}</td>
              <td>${Html.esc(item.referencia) || "-"}</td>
              <td class="text-end">${Formato.moeda(item.valorUnitario)}</td>
              <td><span class="badge-status ${statusCls}">${Html.esc(item.statusOrcamento) || "-"}</span></td>
              <td class="text-end ${margemCls}">${Formato.porcentagem(item.margem)}</td>
              <td>
                <div class="dropdown">
                  <button class="btn-acoes" type="button"
                    data-bs-toggle="dropdown" aria-expanded="false" title="Ações">
                    <i class="bi bi-eye"></i>
                  </button>
                  <ul class="dropdown-menu dropdown-menu-end dropdown-acoes-menu">
                    <li>
                      <button class="dropdown-item" data-nr="${nrOrc}" data-idx="${idxAbs}">
                        <i class="bi bi-info-circle me-2 text-secondary"></i>Detalhes
                      </button>
                    </li>
                    ${temCarta ? `<li>
                      <button class="dropdown-item btn-carta" data-nr="${nrOrc}">
                        <i class="bi bi-envelope-open me-2 text-secondary"></i>Carta
                      </button>
                    </li>` : ""}
                  </ul>
                </div>
              </td>
            </tr>`;
          });
          $("#tblCorpo").html(linhas.join(""));
        },
      });

      /* ════════════════════════════════════════════════════════
         Search Box (filtro dinâmico em memória)
      ════════════════════════════════════════════════════════ */
      let _searchTimer = null;
      const aplicarSearch = (termo) => {
        const t = termo.trim().toLowerCase();
        _dadosFiltrados = t
          ? _dados.filter(item =>
              COLUNAS_EXPORT.some(c => String(item[c.key] ?? "").toLowerCase().includes(t))
            )
          : _dados;
        _pgAtual = 1;
        if (!_dadosFiltrados.length) {
          Tabela.renderizarVazia();
          $("#rodapeTabela").addClass("d-none");
        } else {
          Tabela.renderizarPagina();
          Paginacao.renderizarControles();
        }
      };

      /* ════════════════════════════════════════════════════════
         Modal de Detalhes
      ════════════════════════════════════════════════════════ */
      const Modal = Object.freeze({
        abrirDetalhes(nrOrcamento, idxAbs) {
          let item = (idxAbs !== undefined && idxAbs !== null) ? _dadosFiltrados[idxAbs] : null;
          if (!item) item = _dados.find((d) => String(d.nrOrcamento) === String(nrOrcamento));
          if (!item) { UI.notificar("Orçamento não encontrado.", "warning"); return; }

          const statusCls  = Formato.statusClasse(item.statusOrcamento);
          const plantaCls  = Formato.plantaClasse(item.planta ?? "");
          const margem     = parseFloat(item.margem) || 0;
          const margemCls  = margem >= 0 ? "valor-positivo" : "valor-negativo";
          const statusNorm = (item.statusOrcamento ?? "").toUpperCase().trim();
          const temLog     = statusNorm === "EM APROVACAO" || statusNorm === "EM APROVAÇÃO" || statusNorm === "APROVADO TOTAL";

          $("#modalDetalhesBody").html(`
            <div class="row g-4">
              <div class="col-md-6">
                <h6 class="fw-semibold text-secondary mb-2"><i class="bi bi-building me-1"></i>Informações Gerais</h6>
                <table class="table table-sm">
                  <tr><th width="45%">Nº Orçamento:</th><td class="fw-semibold">${Html.esc(item.nrOrcamento) || "-"}</td></tr>
                  <tr><th>Planta:</th><td><span class="planta-badge ${plantaCls}">${Html.esc(item.planta) || "-"}</span></td></tr>
                  <tr><th>Status:</th><td><span class="badge-status ${statusCls}">${Html.esc(item.statusOrcamento) || "-"}</span></td></tr>
                  <tr><th>Data Inclusão:</th><td>${Html.esc(item.dataInclusao) || "-"}</td></tr>
                  <tr><th>Validade De:</th><td>${Html.esc(item.validadeDe) || "-"}</td></tr>
                  <tr><th>Validade Até:</th><td>${Html.esc(item.validadeAte) || "-"}</td></tr>
                  <tr><th>Tipo Frete:</th><td>${Html.esc(item.tipoFrete) || "-"}</td></tr>
                  <tr><th>Cond. Pagamento:</th><td>${Html.esc(item.descricao) || "-"}</td></tr>
                </table>
              </div>
              <div class="col-md-6">
                <h6 class="fw-semibold text-secondary mb-2"><i class="bi bi-person me-1"></i>Cliente e Representante</h6>
                <table class="table table-sm">
                  <tr><th width="40%">Cliente:</th><td>${Html.esc(item.cliente) || "-"}</td></tr>
                  <tr><th>CNPJ:</th><td>${Formato.cnpj(item.cnpj)}</td></tr>
                  <tr><th>Cidade/UF:</th><td>${Html.esc(item.cidade) || "-"} / ${Html.esc(item.uf) || "-"}</td></tr>
                  <tr><th>Representante:</th><td>${Html.esc(item.representante) || "-"}</td></tr>
                </table>
              </div>
            </div>
            <hr class="my-2" />
            <div class="row g-4">
              <div class="col-md-6">
                <h6 class="fw-semibold text-secondary mb-2"><i class="bi bi-box me-1"></i>Produto</h6>
                <table class="table table-sm">
                  <tr><th width="40%">Produto:</th><td>${Html.esc(item.produto) || "-"}</td></tr>
                  <tr><th>Referência:</th><td>${Html.esc(item.referencia) || "-"}</td></tr>
                </table>
              </div>
              <div class="col-md-6">
                <h6 class="fw-semibold text-secondary mb-2"><i class="bi bi-calculator me-1"></i>Valores</h6>
                <table class="table table-sm">
                  <tr><th width="40%">Valor Unitário:</th><td>${Formato.moeda(item.valorUnitario)}</td></tr>
                  <tr><th>Margem:</th><td class="${margemCls}">${Formato.porcentagem(item.margem)}</td></tr>
                </table>
              </div>
            </div>
            ${temLog ? `
            <hr class="my-2" />
            <div class="log-secao-header">
              <i class="bi bi-clock-history"></i>Log de Aprovações
            </div>
            <div class="log-secao-body" id="logAprovacaoBody">
              <div class="log-loading">
                <div class="spinner-border spinner-border-sm text-secondary me-2" role="status"></div>
                Carregando log de aprovações...
              </div>
            </div>
            ` : ""}
          `);

          bootstrap.Modal.getOrCreateInstance("#modalDetalhes").show();

          if (temLog) {
            $.ajax({
              url:      REPO,
              type:     "GET",
              dataType: "json",
              data:     { acao: "logAprovacao", nrOrcamento: item.nrOrcamento },
              success(resp) {
                if (!resp.sucesso) {
                  $("#logAprovacaoBody").html(`<div class="log-vazio text-danger"><i class="bi bi-exclamation-circle me-1"></i>${Html.esc(resp.mensagem)}</div>`);
                  return;
                }
                const registros = (resp.dados && resp.dados.log) ? resp.dados.log : [];
                if (!registros.length) {
                  $("#logAprovacaoBody").html(`<div class="log-vazio"><i class="bi bi-info-circle me-1 text-secondary"></i>Nenhuma aprovação registrada até o momento.</div>`);
                  return;
                }
                const linhas = registros.map((r) => `
                  <tr>
                    <td>${Html.esc(r.planta) || "-"}</td>
                    <td>${Html.esc(r.aprovador) || "-"}</td>
                    <td><span class="nivel-badge">${Html.esc(r.nivel) || "-"}</span></td>
                    <td>${Html.esc(r.data) || "-"}</td>
                    <td>${Html.esc(r.justificativa) || "-"}</td>
                  </tr>
                `).join("");
                $("#logAprovacaoBody").html(`
                  <table class="table table-sm table-bordered mb-0">
                    <thead>
                      <tr>
                        <th>Planta</th>
                        <th>Aprovador</th>
                        <th>Nível</th>
                        <th>Data da Aprovação</th>
                        <th>Justificativa</th>
                      </tr>
                    </thead>
                    <tbody>${linhas}</tbody>
                  </table>
                `);
              },
              error(_x, _s, err) {
                $("#logAprovacaoBody").html(`<div class="log-vazio text-danger"><i class="bi bi-exclamation-circle me-1"></i>Erro ao carregar log: ${Html.esc(err)}</div>`);
              },
            });
          }
        },
      });

      /* ════════════════════════════════════════════════════════
         BUSCA CLIENTE — SPA Modal
      ════════════════════════════════════════════════════════ */
      let _modalCliente = null;

      const BuscaCliente = Object.freeze({
        init() {
          _modalCliente = bootstrap.Modal.getOrCreateInstance("#modalBuscaCliente");
        },
        abrir() {
          const planta = plantaAtual();
          if (!planta) {
            UI.notificar("Selecione uma planta antes de buscar o cliente.", "warning");
            $("#filtroPlanta").trigger("focus");
            return;
          }
          const textoAtual = $("#filtroCliente").val().trim();
          $("#inputBuscaCliente").val(textoAtual);
          $("#tblBuscaClienteWrapper").html(`
            <div class="busca-placeholder">
              <i class="bi bi-clipboard-data"></i>
              <p>Digite o nome ou código e clique em <strong>Filtrar</strong>.</p>
            </div>
          `);
          _modalCliente.show();
          if (textoAtual) this.executar();
        },
        executar() {
          const planta    = plantaAtual();
          const parametro = $("#inputBuscaCliente").val().trim();
          if (!parametro) {
            UI.notificar("Informe um nome ou código para buscar.", "warning");
            return;
          }
          $("#tblBuscaClienteWrapper").html(`
            <div class="busca-placeholder">
              <div class="spinner-border spinner-border-sm text-secondary me-2" role="status"></div>
              Buscando...
            </div>
          `);
          $.ajax({
            url:      REPO,
            type:     "GET",
            dataType: "json",
            data:     { acao: "buscaCliente", parametro, planta },
            success(resp) {
              if (!resp.sucesso) {
                UI.notificar(`Erro: ${resp.mensagem ?? "Erro inesperado."}`, "danger");
                $("#tblBuscaClienteWrapper").html(`
                  <div class="busca-placeholder">
                    <i class="bi bi-exclamation-circle text-danger"></i>
                    <p>${Html.esc(resp.mensagem)}</p>
                  </div>
                `);
                return;
              }
              BuscaCliente.renderizar(resp.dados ?? []);
            },
            error(_x, _s, err) {
              UI.notificar(`Erro na busca: ${err}`, "danger");
              $("#tblBuscaClienteWrapper").html(`
                <div class="busca-placeholder">
                  <i class="bi bi-exclamation-circle text-danger"></i>
                  <p>Erro ao conectar ao servidor.</p>
                </div>
              `);
            },
          });
        },
        renderizar(lista) {
          if (!lista.length) {
            $("#tblBuscaClienteWrapper").html(`
              <div class="busca-placeholder">
                <i class="bi bi-person-x"></i>
                <p>Cliente não encontrado. Verifique o nome/código ou a planta selecionada.</p>
              </div>
            `);
            return;
          }
          const avisoLimite = lista.length >= 100
            ? `<div class="alert alert-warning alert-sm py-1 px-2 mb-2" style="font-size:.8rem;">
                 <i class="bi bi-exclamation-triangle-fill me-1"></i>
                 Exibindo os primeiros <strong>100</strong> resultados. Refine o termo de busca para encontrar mais facilmente.
               </div>`
            : "";
          const linhas = lista.map((c) => {
            const bloqueado = c.bloqueado == "1";
            const cls       = bloqueado ? " tr-alerta" : "";
            let icone = "";
            if (bloqueado) {
              icone = `<i class="bi bi-slash-circle text-danger" title="Cliente bloqueado. Solicite o desbloqueio com o comercial."></i>`;
            } else {
              icone = `<button type="button" class="btn-selecionar-busca btn-sel-cli"
                         data-cod="${Html.esc(c.codigo)}"
                         data-loja="${Html.esc(c.loja)}"
                         data-nome="${Html.esc(c.nome)}"
                         title="Selecionar cliente">
                         <i class="bi bi-download"></i>
                       </button>`;
            }
            return `<tr class="${cls}">
              <td>${icone}</td>
              <td>${Html.esc(c.codigo)}</td>
              <td>${Html.esc(c.loja)}</td>
              <td class="text-start">${Html.esc(c.nome)}</td>
              <td>${Html.esc(c.cnpj) || "-"}</td>
              <td class="text-start">${Html.esc(c.bairro) || "-"}</td>
              <td class="text-start">${Html.esc(c.municipio) || "-"}</td>
              <td>${Html.esc(c.uf) || "-"}</td>
            </tr>`;
          }).join("");
          $("#tblBuscaClienteWrapper").html(`
            ${avisoLimite}
            <table class="table table-bordered table-sm mb-0">
              <thead>
                <tr>
                  <th></th>
                  <th>Código</th>
                  <th>Loja</th>
                  <th>Nome</th>
                  <th>CNPJ</th>
                  <th>Bairro</th>
                  <th>Município</th>
                  <th>UF</th>
                </tr>
              </thead>
              <tbody>${linhas}</tbody>
            </table>
          `);
        },
        selecionar(codigo, loja, nome) {
          /* Preenche código e loja nos campos visíveis e nos hiddens */
          $("#filtroCliente").val(codigo);
          $("#filtroClienteLojaVisivel").val(loja);
          $("#filtroClienteCod").val(codigo);
          $("#filtroClienteLoja").val(loja);
          _modalCliente.hide();
          UI.notificar(`Cliente selecionado: ${codigo} - ${loja}   ${nome}`, "success");
        },
        limpar() {
          $("#filtroCliente, #filtroClienteLojaVisivel, #filtroClienteCod, #filtroClienteLoja").val("");
        },
      });

      /* ════════════════════════════════════════════════════════
         BUSCA PRODUTO — SPA Modal
      ════════════════════════════════════════════════════════ */
      let _modalProduto = null;

      const BuscaProduto = Object.freeze({
        init() {
          _modalProduto = bootstrap.Modal.getOrCreateInstance("#modalBuscaProduto");
        },
        abrir() {
          const planta = plantaAtual();
          if (!planta) {
            UI.notificar("Selecione uma planta antes de buscar o produto.", "warning");
            $("#filtroPlanta").trigger("focus");
            return;
          }
          const textoAtual = $("#filtroProduto").val().trim();
          $("#inputBuscaProduto").val(textoAtual);
          $("#tblBuscaProdutoWrapper").html(`
            <div class="busca-placeholder">
              <i class="bi bi-clipboard-data"></i>
              <p>Digite a descrição ou código e clique em <strong>Filtrar</strong>.</p>
            </div>
          `);
          _modalProduto.show();
          if (textoAtual) this.executar();
        },
        executar() {
          const planta    = plantaAtual();
          const parametro = $("#inputBuscaProduto").val().trim();
          if (!parametro) {
            UI.notificar("Informe uma descrição ou código para buscar.", "warning");
            return;
          }
          $("#tblBuscaProdutoWrapper").html(`
            <div class="busca-placeholder">
              <div class="spinner-border spinner-border-sm text-secondary me-2" role="status"></div>
              Buscando...
            </div>
          `);
          $.ajax({
            url:      REPO,
            type:     "GET",
            dataType: "json",
            data:     { acao: "buscaProduto", parametro, planta },
            success(resp) {
              if (!resp.sucesso) {
                UI.notificar(`Erro: ${resp.mensagem ?? "Erro inesperado."}`, "danger");
                $("#tblBuscaProdutoWrapper").html(`
                  <div class="busca-placeholder">
                    <i class="bi bi-exclamation-circle text-danger"></i>
                    <p>${Html.esc(resp.mensagem)}</p>
                  </div>
                `);
                return;
              }
              BuscaProduto.renderizar(resp.dados ?? []);
            },
            error(_x, _s, err) {
              UI.notificar(`Erro na busca: ${err}`, "danger");
              $("#tblBuscaProdutoWrapper").html(`
                <div class="busca-placeholder">
                  <i class="bi bi-exclamation-circle text-danger"></i>
                  <p>Erro ao conectar ao servidor.</p>
                </div>
              `);
            },
          });
        },
        renderizar(lista) {
          if (!lista.length) {
            $("#tblBuscaProdutoWrapper").html(`
              <div class="busca-placeholder">
                <i class="bi bi-box-x"></i>
                <p>Produto não encontrado. Verifique a descrição/código ou a planta selecionada.</p>
              </div>
            `);
            return;
          }
          const linhas = lista.map((p) => {
            const disponivel = p.disponivel != "0";
            let icone = "";
            if (disponivel) {
              icone = `<button type="button" class="btn-selecionar-busca btn-sel-prod"
                         data-cod="${Html.esc(p.codigo)}"
                         data-desc="${Html.esc(p.descricao)}"
                         title="Selecionar produto">
                         <i class="bi bi-download"></i>
                       </button>`;
            } else {
              icone = `<i class="bi bi-slash-circle text-danger" title="Produto indisponível para seleção."></i>`;
            }
            return `<tr>
              <td>${icone}</td>
              <td>${Html.esc(p.codigo)}</td>
              <td class="text-start">${Html.esc(p.descricao)}</td>
              <td class="text-start">${Html.esc(p.estrutura) || "-"}</td>
              <td>${Html.esc(p.filial) || "-"}</td>
              <td>${Html.esc(p.cliente) || "-"}</td>
            </tr>`;
          }).join("");
          $("#tblBuscaProdutoWrapper").html(`
            <table class="table table-bordered table-sm mb-0">
              <thead>
                <tr>
                  <th></th>
                  <th>Código</th>
                  <th>Descrição</th>
                  <th>Estrutura</th>
                  <th>Filial</th>
                  <th>Cliente</th>
                </tr>
              </thead>
              <tbody>${linhas}</tbody>
            </table>
          `);
        },
        selecionar(codigo, descricao) {
          $("#filtroProduto").val(descricao || codigo);
          $("#filtroProdutoCod").val(codigo);
          _modalProduto.hide();
          UI.notificar(`Produto selecionado: ${descricao || codigo}`, "success");
        },
        limpar() {
          $("#filtroProduto, #filtroProdutoCod").val("");
        },
      });

      /* ════════════════════════════════════════════════════════
         EXPORTAR
      ════════════════════════════════════════════════════════ */
      const Exportar = Object.freeze({
        csv() {
          if (!_dados.length) { UI.notificar("Nenhum dado para exportar.", "warning"); return; }

          const esc = v => {
            const s = v == null ? "" : String(v);
            return (s.includes(";") || s.includes('"') || s.includes("\n"))
              ? `"${s.replace(/"/g, '""')}"` : s;
          };

          const cabecalhos = ["#", ...COLUNAS_EXPORT.map(c => c.label)];
          const linhas = _dados.map((item, i) => {
            const row = [i + 1];
            COLUNAS_EXPORT.forEach(c => {
              let v = item[c.key] ?? "";
              if (c.key === "dataInclusao" || c.key === "validadeDe" || c.key === "validadeAte") {
                v = Formato.dataCSV(v);
              } else if (c.key === "valorUnitario") {
                v = Formato.numCSV(v, 5);
              } else if (c.key === "margem") {
                v = Formato.numCSV(v, 2);
              }
              row.push(v);
            });
            return row;
          });

          const rows = [
            cabecalhos.map(esc).join(";"),
            ...linhas.map(row => row.map(esc).join(";"))
          ];

          const blob = new Blob(["\uFEFF" + rows.join("\r\n")], { type: "text/csv;charset=utf-8;" });
          const url  = URL.createObjectURL(blob);
          const a    = document.createElement("a");
          a.href     = url;
          a.download = Formato.nomeArquivo("csv");
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          UI.notificar(`CSV gerado com ${_dados.length} registro(s).`, "success");
        },

        pdf() {
          if (!_dados.length) { UI.notificar("Nenhum dado para exportar.", "warning"); return; }

          const COLUNAS_PDF = COLUNAS_EXPORT.map((c) =>
            c.key === "nrOrcamento" ? { ...c, label: "Nº Orç." } : c
          );

          const head = [["Nº", ...COLUNAS_PDF.map((c) => c.label)]];
          const body = _dados.map((item, i) => [
            i + 1,
            ...COLUNAS_PDF.map((c) => {
              if (c.key === "valorUnitario")   return Formato.moeda(item[c.key]);
              if (c.key === "margem")          return Formato.porcentagem(item[c.key]);
              if (c.key === "cnpj")            return Formato.cnpj(item[c.key]);
              if (c.key === "statusOrcamento") return item[c.key] ?? "-";
              return item[c.key] ?? "";
            }),
          ]);

          const { jsPDF } = window.jspdf;
          const doc       = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
          const PAG_L     = doc.internal.pageSize.getWidth();
          const PAG_H     = doc.internal.pageSize.getHeight();
          const HDR_H     = 26;
          const MARGEM_L  = 10;
          const f         = Filtro.valores();
          const agora     = new Date();
          const dtStr     = agora.toLocaleDateString("pt-BR");
          const hrStr     = agora.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
          const loginExib = _login || "—";

          const _cabecalho = () => {
            doc.setFillColor(255, 255, 255);
            doc.rect(0, 0, PAG_L, HDR_H, "F");
            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.4);
            doc.line(0, HDR_H, PAG_L, HDR_H);

            if (_logoB64) {
              try { doc.addImage(_logoB64, "PNG", MARGEM_L, 5, 38, 13, undefined, "FAST"); } catch (_) {}
            }

            doc.setFontSize(13);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(74, 74, 74);
            doc.text("Relatório de Orçamentos", PAG_L / 2, 11, { align: "center" });

            doc.setFontSize(7.5);
            doc.setFont("helvetica", "normal");
            doc.setTextColor(100, 100, 100);
            const fmtData = (v) => v ? v.split("-").reverse().join("/") : "—";
            doc.text(
              `Período: ${fmtData(f.dataDe)} a ${fmtData(f.dataAte)}   |   Total: ${_dados.length} registro(s)`,
              PAG_L / 2, 18, { align: "center" }
            );

            doc.setFontSize(7);
            doc.setTextColor(130, 130, 130);
            const prefixo  = "Emitido por: ";
            const sufixo   = `   |   ${dtStr} ${hrStr}`;
            const largPref = doc.getTextWidth(prefixo);
            const largNome = doc.getTextWidth(loginExib);
            const largSuf  = doc.getTextWidth(sufixo);
            const xInic    = PAG_L - MARGEM_L - largPref - largNome - largSuf;

            doc.setFont("helvetica", "normal");
            doc.text(prefixo, xInic, 23);
            doc.setFont("helvetica", "bold");
            doc.text(loginExib, xInic + largPref, 23);
            doc.setFont("helvetica", "normal");
            doc.text(sufixo, xInic + largPref + largNome, 23);

            doc.setTextColor(0, 0, 0);
          };

          _cabecalho();

          const MARGEM_TB  = 4;
          const larguraTbl = PAG_L - MARGEM_TB * 2;
          const colWidths  = [5, 14, 14, 30, 20, 14, 6, 22, 12, 20, 16, 16, 16, 22, 16, 18, 20, 14];
          const fator      = larguraTbl / colWidths.reduce((a, b) => a + b, 0);

          const columnStyles = {};
          colWidths.forEach((w, idx) => {
            columnStyles[idx] = { cellWidth: parseFloat((w * fator).toFixed(2)), halign: "center" };
          });
          columnStyles[7].halign  = "left";
          columnStyles[14].halign = "left";

          const IDX_MARGEM = 17;

          doc.autoTable({
            head,
            body,
            startY:             29,
            tableWidth:         larguraTbl,
            styles:             { fontSize: 4.5, cellPadding: 1.5, valign: "middle", overflow: "linebreak", textColor: [33, 37, 41], lineColor: [180, 180, 180], lineWidth: 0.2 },
            headStyles:         { fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: "bold", halign: "center", fontSize: 4.5 },
            alternateRowStyles: { fillColor: [245, 245, 245] },
            columnStyles,
            margin:             { left: MARGEM_TB, right: MARGEM_TB },
            didDrawPage:        _cabecalho,
            didParseCell(data) {
              if (data.section !== "body" || data.column.index !== IDX_MARGEM) return;
              const raw = _dados[data.row.index];
              if (!raw) return;
              const n = parseFloat(String(raw.margem ?? "").replace(",", "."));
              if (isNaN(n)) return;
              data.cell.styles.textColor = n >= 0 ? [25, 135, 84] : [220, 53, 69];
              data.cell.styles.fontStyle = "bold";
            },
          });

          const totalPags = doc.getNumberOfPages();
          for (let p = 1; p <= totalPags; p++) {
            doc.setPage(p);
            doc.setFontSize(7);
            doc.setTextColor(150);
            doc.text(`Página ${p} de ${totalPags}`, PAG_L - MARGEM_L, PAG_H - 5, { align: "right" });
          }

          doc.save(Formato.nomeArquivo("pdf"));
          UI.notificar(`PDF gerado com ${_dados.length} registro(s).`, "success");
        },
      });

      const _carregarLogo = () =>
        new Promise((resolve) => {
          const img = new Image();
          img.crossOrigin = "anonymous";
          img.onload = () => {
            const SCALE  = 3;
            const canvas = document.createElement("canvas");
            canvas.width  = 152 * SCALE;
            canvas.height = 52  * SCALE;
            const ctx = canvas.getContext("2d");
            ctx.fillStyle = "#ffffff";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            _logoB64 = canvas.toDataURL("image/png");
            resolve();
          };
          img.onerror = () => resolve();
          img.src = `${LOGO_URL}?v=${Date.now()}`;
        });

      const Consulta = Object.freeze({
        executar() {
          if (!Filtro.validar()) return;
          UI.exibirLoader();
          $.ajax({
            url:      REPO,
            type:     "GET",
            dataType: "json",
            data:     { acao: "listar", ...Filtro.valores() },
            success(resp) {
              UI.ocultarLoader();
              if (!resp.sucesso) {
                UI.notificar(`Erro: ${resp.mensagem ?? "Erro inesperado."}`, "danger");
                Tabela.renderizar([]);
                return;
              }
              const dados = resp.dados ?? [];
              Tabela.renderizar(dados);
              if (!dados.length) UI.notificar("Nenhum registro encontrado para o filtro aplicado.", "warning");
              else UI.notificar(`${dados.length} registro(s) encontrado(s).`, "success");
            },
            error(_x, _s, err) {
              UI.ocultarLoader();
              UI.notificar(`Erro na consulta: ${err}`, "danger");
            },
          });
        },
      });

      const carregarFiltros = () => {
        $.ajax({
          url:      REPO,
          type:     "GET",
          dataType: "json",
          data:     { acao: "filtros" },
          success(resp) {
            if (!resp.sucesso) return;
            const $planta = $("#filtroPlanta").empty().append('<option value="">Selecione...</option>');
            ((resp.dados && resp.dados.plantas) ? resp.dados.plantas : []).forEach((p) => {
              $planta.append(`<option value="${Html.esc(p.codigo)}">${Html.esc(p.nome)}</option>`);
            });
            const $status = $("#filtroStatus").empty().append('<option value="">Todos</option>');
            ((resp.dados && resp.dados.status) ? resp.dados.status : []).forEach((s) => {
              $status.append(`<option value="${Html.esc(s)}">${Html.esc(s)}</option>`);
            });
          },
        });
      };

      const Sessao = Object.freeze({
        carregar() {
          if (typeof LoginUsuarioLogado !== "undefined") _login = LoginUsuarioLogado;
          _carregarLogo();
        },
      });

      const anexarEventos = () => {
        /* ── Filtro principal ── */
        $("#btnFiltrar").on("click", () => Consulta.executar());

        $("#btnLimpar").on("click", () => {
          Filtro.limpar();
          Tabela.renderizar([]);
        });

        $("#btnExportCsv").on("click", () => Exportar.csv());
        $("#btnExportPdf").on("click", () => Exportar.pdf());

        /* Enter nos inputs de filtro ── */
        $("#filtroNrOrcamento, #filtroCliente, #filtroProduto").on("keypress", function (e) {
          if (e.which === 13) Consulta.executar();
        });

        /* Limpar campos de cliente e produto ── */
        $("#btnLimparCliente").on("click", () => BuscaCliente.limpar());
        $("#btnLimparProduto").on("click", () => BuscaProduto.limpar());

        /* Digitar manualmente limpa os hiddens — o parse ocorre em Filtro._clienteParams() no submit ── */
        $("#filtroCliente").on("input", function () {
          $("#filtroClienteCod, #filtroClienteLoja").val("");
        });
        $("#filtroClienteLojaVisivel").on("input", function () {
          $("#filtroClienteLoja").val("");
        });
        $("#filtroProduto").on("input", function () {
          if (!$(this).val()) { $("#filtroProdutoCod").val(""); }
        });

        /* Page size ── */
        $("#pgTamanho").on("change", function () {
          _pgTamanho = parseInt($(this).val(), 10);
          _pgAtual   = 1;
          Tabela.renderizarPagina();
          Paginacao.renderizarControles();
        });

        /* Ações da tabela ── */
        $("#tblCorpo").on("click", "[data-nr]", function () {
          const nr = $(this).data("nr");
          if ($(this).hasClass("btn-carta")) {
            window.open(`../OrcamentoImprimir.asp?paramOrcamento=${nr}`, "_blank");
          } else {
            const idx = $(this).data("idx");
            Modal.abrirDetalhes(nr, idx);
          }
        });

        /* ── Abrir modais de busca ── */
        $("#btnAbrirBuscaCliente").on("click", () => BuscaCliente.abrir());
        $("#btnAbrirBuscaProduto").on("click", () => BuscaProduto.abrir());

        /* ── Enter dentro dos modais ── */
        $("#inputBuscaCliente").on("keypress", function (e) {
          if (e.which === 13) BuscaCliente.executar();
        });
        $("#inputBuscaProduto").on("keypress", function (e) {
          if (e.which === 13) BuscaProduto.executar();
        });

        /* ── Botões Filtrar dentro dos modais ── */
        $("#btnExecutarBuscaCliente").on("click", () => BuscaCliente.executar());
        $("#btnExecutarBuscaProduto").on("click", () => BuscaProduto.executar());

        /* ── Seleção nos modais (delegação) ── */
        $("#tblBuscaClienteWrapper").on("click", ".btn-sel-cli", function () {
          const $btn = $(this);
          BuscaCliente.selecionar($btn.data("cod"), $btn.data("loja"), $btn.data("nome"));
        });
        $("#tblBuscaProdutoWrapper").on("click", ".btn-sel-prod", function () {
          const $btn = $(this);
          BuscaProduto.selecionar($btn.data("cod"), $btn.data("desc"));
        });

        /* ── Search Box ── */
        $("#searchBox").on("input", function () {
          clearTimeout(_searchTimer);
          const val = $(this).val();
          _searchTimer = setTimeout(() => aplicarSearch(val), 250);
        });

        /* ── Atalho Ctrl+K ── */
        $(document).on("keydown", function (e) {
          if (e.ctrlKey && e.key === "k") {
            e.preventDefault();
            $("#searchBox").trigger("focus");
          }
        });
      };

      $(function () {
        UI.ocultarLoader();
        carregarFiltros();
        Sessao.carregar();
        BuscaCliente.init();
        BuscaProduto.init();
        $("#filtroDataAte").val(new Date().toISOString().split("T")[0]);
        anexarEventos();
      });

    })(jQuery);
  </script>

</body>
</html>
