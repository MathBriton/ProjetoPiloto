<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SWnet - Controle de Recebimento</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
  <style>
    body {
      background: #f0f0f0;
      color: #212529;
      min-height: 100vh;
    }

    .page-title {
      color: #4a4a4a;
      font-size: 1.25rem;
    }

    .form-label {
      font-size: 0.82rem;
      font-weight: 600;
      margin-bottom: 0.25rem;
      color: #212529;
    }
    .form-control,
    .form-select {
      font-size: 0.85rem;
    }
    .form-control:focus,
    .form-select:focus {
      border-color: #4a4a4a;
      box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
    }
    .form-control[readonly] {
      background-color: #f8f9fa;
    }
    .text-required {
      color: #dc3545;
    }

    .sw-container {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
    }
    .sw-container-header {
      background: #fff;
      color: #212529;
      border-bottom: 1px solid #dee2e6;
      padding: 0.6rem 1rem;
      border-radius: 6px 6px 0 0;
      font-size: 0.8rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.3px;
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }
    .sw-container-body {
      padding: 0.75rem 1rem;
    }

    .tbl-wrapper {
      overflow: auto;
      max-height: 78vh;
      border-radius: 4px;
    }
    .tbl-wrapper thead th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #f8f9fa;
      color: #212529;
      font-size: 0.71rem;
      font-weight: 600;
      white-space: nowrap;
      vertical-align: middle;
      text-align: center;
      padding: 0.45rem 0.75rem;
      border-color: #dee2e6;
      text-transform: uppercase;
      letter-spacing: 0.2px;
    }
    .tbl-wrapper tbody td {
      font-size: 0.78rem;
      vertical-align: middle;
      padding: 0.32rem 0.75rem;
      border-color: #e8e8e8;
      text-align: center;
    }
    .col-razao-social { min-width: 200px; }
    .col-produto      { min-width: 160px; }
    .col-local        { min-width: 120px; }
    .col-motorista    { min-width: 180px; }
    .tbl-wrapper tbody tr:hover td {
      background-color: rgba(0, 0, 0, 0.03);
    }
    .tbl-vazia {
      padding: 3rem 1rem;
      text-align: center;
      color: #888;
      font-size: 0.88rem;
    }

    .btn-acao {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 26px;
      height: 22px;
      padding: 0;
      border-radius: 4px;
      font-size: 0.78rem;
      cursor: pointer;
      transition: opacity 0.15s;
    }
    .btn-acao:hover {
      opacity: 0.82;
    }
    .btn-acao-outline {
      background: transparent;
      border: 1px solid #dee2e6;
      color: #6c757d;
    }
    .btn-acao-outline:hover {
      background: #f8f9fa;
      opacity: 1;
    }

    .toast-wrapper {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 11000;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
      min-width: 300px;
      max-width: 380px;
      pointer-events: none;
    }
    .toast-item {
      background: #fff;
      border-radius: 6px;
      overflow: hidden;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
      border: 1px solid #e8e8e8;
      pointer-events: all;
      animation: entrar-toast 0.2s ease;
    }
    .toast-item-success { border-left: 4px solid #198754; }
    .toast-item-warning { border-left: 4px solid #ffc107; }
    .toast-item-danger  { border-left: 4px solid #dc3545; }
    .toast-item-info    { border-left: 4px solid #6c757d; }
    @keyframes entrar-toast {
      from { opacity: 0; transform: translateX(20px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-corpo {
      display: flex;
      align-items: center;
      gap: 0.6rem;
      padding: 0.65rem 0.85rem;
      font-size: 0.85rem;
      color: #333;
    }
    .toast-fechar {
      margin-left: auto;
      background: none;
      border: none;
      font-size: 1rem;
      color: #999;
      cursor: pointer;
      padding: 0 0 0 0.5rem;
    }
    .toast-fechar:hover { color: #333; }
    .toast-barra { height: 3px; animation: encolher linear forwards; }
    .toast-barra-success { background: #198754; }
    .toast-barra-warning { background: #ffc107; }
    .toast-barra-danger  { background: #dc3545; }
    .toast-barra-info    { background: #6c757d; }
    @keyframes encolher {
      from { width: 100%; }
      to   { width: 0%; }
    }

    .overlay-loader {
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, 0.65);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .spinner-circular {
      width: 3rem;
      height: 3rem;
      border: 4px solid #e8e8e8;
      border-top-color: #4a4a4a;
      border-radius: 50%;
      animation: girar 0.8s linear infinite;
    }
    @keyframes girar { to { transform: rotate(360deg); } }

    .rodape-paginacao {
      display: grid;
      grid-template-columns: auto 1fr auto;
      align-items: center;
      padding: 0.5rem 0.75rem;
      background: #f8f9fa;
      border: 1px solid #dee2e6;
      border-top: none;
      border-radius: 0 0 6px 6px;
      font-size: 0.8rem;
      color: #555;
      gap: 0.5rem;
    }
    .rodape-esquerda { display: flex; align-items: center; gap: 0.4rem; }
    .rodape-centro   { display: flex; justify-content: center; }
    .rodape-direita  {
      display: flex;
      justify-content: flex-end;
      font-size: 0.78rem;
      white-space: nowrap;
      color: #212529;
      font-weight: 500;
    }
    .paginacao .page-item .page-link    { color: #6c757d; font-size: 0.78rem; }
    .paginacao .page-item.active .page-link {
      background: #4a4a4a;
      border-color: #4a4a4a;
      color: #fff;
    }
    .paginacao .page-item.disabled .page-link { color: #aaa; pointer-events: none; }

    .detalhe-campo-label {
      font-size: 0.72rem;
      font-weight: 600;
      color: #6c757d;
      text-transform: uppercase;
      letter-spacing: 0.2px;
      margin-bottom: 0.1rem;
    }
    .detalhe-campo-valor {
      font-size: 0.88rem;
      color: #212529;
      word-break: break-word;
      padding: 0.3rem 0.5rem;
      background: #fff;
      border-radius: 4px;
      border: 1px solid #e9ecef;
      min-height: 2rem;
      display: flex;
      align-items: center;
    }

    .form-mode-badge {
      font-size: 0.75rem;
      font-weight: 600;
      padding: 0.2rem 0.55rem;
      border-radius: 4px;
    }
    .badge-ins { background: rgba(25, 135, 84, 0.12);  color: #146c43; }
    .badge-alt { background: rgba(13, 110, 253, 0.12); color: #0a58ca; }
    .badge-apb { background: rgba(255, 193, 7, 0.18);  color: #856404; }
    .badge-del { background: rgba(220, 53, 69, 0.12);  color: #b02a37; }

    #feedbackCnpj.valido   { color: #198754; font-size: 0.75rem; margin-top: 2px; }
    #feedbackCnpj.invalido { color: #dc3545; font-size: 0.75rem; margin-top: 2px; }

    .modal-panel {
      background: #fff;
      border: 1px solid #e9ecef;
      border-radius: 6px;
      margin-bottom: 1rem;
    }
    .modal-panel-header {
      background: #f8f9fa;
      border-bottom: 1px solid #e9ecef;
      padding: 0.4rem 0.75rem;
      border-radius: 6px 6px 0 0;
      font-size: 0.75rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.3px;
      color: #495057;
      display: flex;
      align-items: center;
      gap: 0.35rem;
    }
    .modal-panel-body { padding: 0.85rem; }

    .modal-body {
      background: #f0f0f0;
      color: #212529;
    }

    .col-obs { min-width: 260px; }

    .sw-tbl-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0.5rem 0.75rem;
      border-bottom: 1px solid #dee2e6;
    }

    .sw-date-input  { width: 145px; }
    .sw-input-nf    { min-width: 150px; }

    .sw-label-nf-serie { font-size: 0.72rem; letter-spacing: .2px; }

    .btn-acao-ver {
      width: auto;
      padding: 0 6px;
      border: none;
    }

    .tbl-busca-forn {
      overflow-y: auto;
      max-height: 480px;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      background: #fff;
    }
    .tbl-busca-forn-table { font-size: 0.82rem; }
    .tbl-busca-th {
      text-align: center;
      font-size: 0.72rem;
      text-transform: uppercase;
      letter-spacing: .2px;
      background: #fff;
      position: sticky;
      top: 0;
      z-index: 1;
    }
    .tbl-busca-icone { font-size: 1.4rem; }
  </style>
</head>
<body>

<div class="toast-wrapper" id="toastWrapper"></div>
<div id="loader" class="overlay-loader d-none"><div class="spinner-circular"></div></div>

<div class="container-fluid py-3 pb-3">

  <div class="d-flex align-items-center justify-content-between mb-3">
    <h4 class="fw-bold mb-0 page-title">Controle de Recebimento</h4>
    <div class="d-flex gap-2">
      <a href="RelatorioControleRecebimento.asp" class="btn btn-outline-danger btn-sm" title="Relatório">
        <i class="bi bi-file-earmark-bar-graph me-2"></i>Relatório
      </a>
    </div>
  </div>

  <div class="sw-container mb-3">
    <div class="sw-container-header">
      <i class="bi bi-funnel"></i>Filtros
    </div>
    <div class="sw-container-body">
      <div class="d-flex flex-wrap align-items-end gap-2">
        <div class="d-flex flex-column">
          <label class="form-label" for="filtroNF">Nota Fiscal</label>
          <input type="text" id="filtroNF" class="form-control form-control-sm sw-input-nf" placeholder="Número N.F" inputmode="numeric">
        </div>
        <div class="d-flex flex-column flex-grow-1">
          <label class="form-label" for="filtroFornecedor">Fornecedor</label>
          <input type="text" id="filtroFornecedor" class="form-control form-control-sm" placeholder="Nome do fornecedor...">
        </div>
        <div class="d-flex flex-column flex-grow-1">
          <label class="form-label" for="filtroProduto">Produto</label>
          <input type="text" id="filtroProduto" class="form-control form-control-sm" placeholder="Nome do produto...">
        </div>
        <div class="d-flex flex-column">
          <label class="form-label" for="filtroDataIni">Inclusão De</label>
          <input type="date" id="filtroDataIni" class="form-control form-control-sm sw-date-input">
        </div>
        <div class="d-flex flex-column">
          <label class="form-label" for="filtroDataFim">Até</label>
          <input type="date" id="filtroDataFim" class="form-control form-control-sm sw-date-input">
        </div>
        <div class="d-flex align-items-end gap-2">
          <button type="button" class="btn btn-primary btn-sm" id="btnFiltrar">
            <i class="bi bi-funnel-fill me-1"></i>Filtrar
          </button>
          <button type="button" class="btn btn-light border btn-sm" id="btnLimparFiltros" title="Limpar Filtros">
            <i class="bi bi-eraser me-1"></i>
          </button>
        </div>
      </div>
    </div>
  </div>

  <div class="sw-container">

    <div class="sw-tbl-header">
      <button type="button" class="btn btn-success btn-sm" id="btnCriar">
        <i class="bi bi-plus-lg me-2"></i>Criar
      </button>
    </div>

    <div class="tbl-wrapper">
        <table class="table table-bordered table-striped mb-0" id="tblDados">
          <thead>
            <tr>
              <th>Ações</th>
              <th>N.F / Série</th>
              <th>Data N.F</th>
              <th>Data/Hora Teste</th>
              <th>CNPJ</th>
              <th class="col-razao-social">Razão Social</th>
              <th class="col-produto">Produto</th>
              <th>Peso N.F</th>
              <th>Peso Balança</th>
              <th>% Umidade</th>
              <th>% Impureza</th>
              <th>% Proibitivo</th>
              <th>% Outros</th>
              <th class="col-local">Local</th>
              <th>Qtd.</th>
              <th>Fardos An.</th>
              <th>Placa</th>
              <th class="col-motorista">Motorista</th>
              <th class="col-obs">Obs</th>
            </tr>
          </thead>
          <tbody id="tbodyDados">
            <tr>
              <td colspan="19">
                <div class="tbl-vazia">
                  <i class="bi bi-search fs-5"></i>
                  <p class="mt-2 mb-0">Carregando registros...</p>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="rodape-paginacao d-none" id="rodapeGrid">
      <div class="rodape-esquerda">
        <label class="form-label mb-0 small" for="cmbLinhas">Exibir</label>
        <select id="cmbLinhas" class="form-select form-select-sm w-auto">
          <option value="10">10</option>
          <option value="20">20</option>
          <option value="50" selected>50</option>
          <option value="100">100</option>
        </select>
      </div>
      <div class="rodape-centro">
        <nav aria-label="Paginação">
          <ul class="pagination pagination-sm mb-0 paginacao" id="paginacaoGrid"></ul>
        </nav>
      </div>
      <div class="rodape-direita" id="contadorGrid"></div>
    </div>

  </div>

</div>

<div class="modal fade" id="modalDetalhe" tabindex="-1" aria-labelledby="labelModalDetalhe" aria-hidden="true">
  <div class="modal-dialog modal-lg modal-dialog-scrollable modal-dialog-centered">
    <div class="modal-content border-0 shadow">
      <div class="modal-header border-bottom px-3 py-2">
        <span class="modal-title fw-semibold text-uppercase small ls-sm" id="labelModalDetalhe">
          Detalhe do Registro
        </span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body" id="modalDetalheBody"></div>
    </div>
  </div>
</div>

<div class="modal fade" id="modalForm" tabindex="-1" aria-labelledby="labelModalForm" aria-hidden="true" data-bs-backdrop="static">
  <div class="modal-dialog modal-xl modal-dialog-scrollable modal-dialog-centered">
    <div class="modal-content border-0 shadow">

      <div class="modal-header border-bottom px-3 py-2">
        <div class="d-flex align-items-center gap-2">
          <span class="modal-title fw-semibold text-uppercase small" id="labelModalForm">
            Registro
          </span>
          <span id="formModeBadge" class="form-mode-badge"></span>
        </div>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>

      <div class="modal-body">

        <div class="modal-panel">
          <div class="modal-panel-header">
            <i class="bi bi-file-earmark-text"></i> Dados da Nota Fiscal
          </div>
          <div class="modal-panel-body">
            <div class="row g-3">
              <div class="col-md-3">
                <label class="form-label" for="edtNF">Nota Fiscal <span class="text-required">*</span></label>
                <input type="text" id="edtNF" class="form-control form-control-sm" maxlength="9">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtSERIE">Série <span class="text-required">*</span></label>
                <input type="text" id="edtSERIE" class="form-control form-control-sm" maxlength="3">
              </div>
              <div class="col-md-3">
                <label class="form-label" for="edtDATA_NF">Data da N.F <span class="text-required">*</span></label>
                <input type="date" id="edtDATA_NF" class="form-control form-control-sm">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtDATA_RECEBE">Data do Teste <span class="text-required">*</span></label>
                <input type="date" id="edtDATA_RECEBE" class="form-control form-control-sm">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtHORA_RECEBE">Hora do Teste <span class="text-required">*</span></label>
                <input type="time" id="edtHORA_RECEBE" class="form-control form-control-sm">
              </div>
            </div>
          </div>
        </div>

        <div class="modal-panel">
          <div class="modal-panel-header d-flex align-items-center justify-content-between">
            <span class="d-flex align-items-center gap-1">
              <i class="bi bi-building"></i> Fornecedor
            </span>
            <span id="btnsFornecedor" class="d-flex gap-1">
              <button type="button" id="btnPesquisarFornecedor" class="btn btn-primary btn-sm py-0 px-2"
                      onclick="pesquisarFornecedor()" title="Pesquisar">
                <i class="bi bi-search"></i>
              </button>
              <button type="button" id="btnLimparFornecedor" class="btn btn-outline-secondary btn-sm py-0 px-2"
                      onclick="limparFornecedor()" title="Limpar">
                <i class="bi bi-x-lg"></i>
              </button>
            </span>
          </div>
          <div class="modal-panel-body">
            <div class="row g-3">
              <div class="col-md-4">
                <label class="form-label" for="edtCNPJ">CNPJ</label>
                <input type="text" id="edtCNPJ" class="form-control form-control-sm" readonly>
              </div>
              <div class="col-md-8">
                <label class="form-label" for="edtNOME_FORN">Razão Social</label>
                <input type="text" id="edtNOME_FORN" class="form-control form-control-sm" readonly>
              </div>
            </div>
          </div>
        </div>

        <div class="modal-panel">
          <div class="modal-panel-header">
            <i class="bi bi-box-seam"></i> Produto e Pesos
          </div>
          <div class="modal-panel-body">
            <div class="row g-3">
              <div class="col-md-4">
                <label class="form-label" for="edtPRODUTO">Produto <span class="text-required">*</span></label>
                <input type="text" id="edtPRODUTO" class="form-control form-control-sm" maxlength="100">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtPESO_NF">Peso N.F/KG <span class="text-required">*</span></label>
                <input type="text" id="edtPESO_NF" class="form-control form-control-sm text-end" placeholder="0,000" maxlength="10">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtPESO_BAL">Peso Balança/KG</label>
                <input type="text" id="edtPESO_BAL" class="form-control form-control-sm text-end" placeholder="0,000" maxlength="10">
              </div>
            </div>
          </div>
        </div>

        <div class="modal-panel">
          <div class="modal-panel-header">
            <i class="bi bi-percent"></i> Percentuais de Qualidade
          </div>
          <div class="modal-panel-body">
            <div class="row g-3">
              <div class="col-md-3">
                <label class="form-label" for="edtUMIDADE">% Umidade <span class="text-required">*</span></label>
                <input type="text" id="edtUMIDADE" class="form-control form-control-sm text-end" placeholder="0,00">
              </div>
              <div class="col-md-3">
                <label class="form-label" for="edtIMPUREZA">% Impureza</label>
                <input type="text" id="edtIMPUREZA" class="form-control form-control-sm text-end" placeholder="0,00">
              </div>
              <div class="col-md-3">
                <label class="form-label" for="edtPROIBITIVO">% Proibitivo</label>
                <input type="text" id="edtPROIBITIVO" class="form-control form-control-sm text-end" placeholder="0,00">
              </div>
              <div class="col-md-3">
                <label class="form-label" for="edtOUTROS">% Outros</label>
                <input type="text" id="edtOUTROS" class="form-control form-control-sm text-end" placeholder="0,00">
              </div>
            </div>
          </div>
        </div>

        <div class="modal-panel mb-0">
          <div class="modal-panel-header">
            <i class="bi bi-truck"></i> Armazenagem e Transporte
          </div>
          <div class="modal-panel-body">
            <div class="row g-3">
              <div class="col-md-3">
                <label class="form-label" for="edtLOCAL">Local Armazenagem</label>
                <input type="text" id="edtLOCAL" class="form-control form-control-sm" maxlength="10">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtNUM_FARDOS">Quantidade</label>
                <input type="text" id="edtNUM_FARDOS" class="form-control form-control-sm text-end" placeholder="0">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtFARDOS">Fardos Analisados</label>
                <input type="text" id="edtFARDOS" class="form-control form-control-sm text-end" placeholder="0">
              </div>
              <div class="col-md-3">
                <label class="form-label" for="edtMOTORISTA">Nome do Motorista <span class="text-required">*</span></label>
                <input type="text" id="edtMOTORISTA" class="form-control form-control-sm" maxlength="300">
              </div>
              <div class="col-md-2">
                <label class="form-label" for="edtPLACA">Placa <span class="text-required">*</span></label>
                <input type="text" id="edtPLACA" class="form-control form-control-sm" maxlength="10">
              </div>
              <div class="col-12">
                <label class="form-label" for="edtOBS">Observação</label>
                <input type="text" id="edtOBS" class="form-control form-control-sm" maxlength="300">
              </div>
            </div>
          </div>
        </div>

      </div>

      <div class="modal-footer border-top px-3 py-2">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Fechar</button>
        <button type="button" class="btn btn-sm btn-success d-none" id="btnSalvar">
          <span class="spinner-border spinner-border-sm me-1 d-none" id="spinnerSalvar"></span>
          Salvar
        </button>
        <button type="button" class="btn btn-sm btn-danger d-none" id="btnApagar">
          <span class="spinner-border spinner-border-sm me-1 d-none" id="spinnerApagar"></span>
          Apagar
        </button>
      </div>

    </div>
  </div>
</div>

<div class="modal fade" id="modalConfirmarApagar" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" style="max-width:420px">
    <div class="modal-content border-0 shadow">
      <div class="modal-header border-bottom pb-2">
        <span class="modal-title fw-bold text-danger fs-6">Atenção</span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body pt-2">
        <br>
        <p class="mb-1 text-center text-secondary small">
          Esta ação é <strong class="text-danger">CRÍTICA</strong>. O registro abaixo será excluído:
        </p>
        <br>
        <div class="p-2 mt-2 rounded border text-center bg-danger-subtle border-danger-subtle">
          <div class="text-muted fw-semibold text-uppercase sw-label-nf-serie">
            N.F / Série
          </div>
          <div id="apagarNfSerie" class="fw-semibold mt-1"></div>
        </div>
      </div>
      <div class="modal-footer border-top pt-1">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Fechar</button>
        <button type="button" class="btn btn-sm btn-danger" id="btnConfirmarApagar">Excluir</button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="modalCadFornece" tabindex="-1" aria-labelledby="labelModalFornece" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content border-0 shadow">
      <div class="modal-header border-bottom px-3 py-2">
        <span class="modal-title fw-semibold text-uppercase small" id="labelModalFornece">
          Cadastro de Fornecedor
        </span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body">
        <div class="row g-3">
          <div class="col-md-5">
            <label class="form-label" for="codifornece">Código SAP <span class="text-required">*</span></label>
            <input id="codifornece" type="text" class="form-control form-control-sm" maxlength="20" placeholder="Código SAP...">
          </div>
          <div class="col-md-7">
            <label class="form-label" for="cnpjfornece">CNPJ <span class="text-required">*</span></label>
            <input id="cnpjfornece" type="text" class="form-control form-control-sm" maxlength="14" placeholder="Somente números...">
            <div id="feedbackCnpj"></div>
          </div>
          <div class="col-12">
            <label class="form-label" for="nomefornece">Razão Social <span class="text-required">*</span></label>
            <input id="nomefornece" type="text" class="form-control form-control-sm" placeholder="Razão Social...">
          </div>
        </div>
      </div>
      <div class="modal-footer border-top">
        <button type="button" class="btn btn-light border btn-sm" data-bs-dismiss="modal">Fechar</button>
        <button type="button" class="btn btn-success btn-sm" id="btnSalvarFornecedor">Salvar</button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="modalBuscaFornecedor" tabindex="-1" aria-hidden="true" data-bs-backdrop="static">
  <div class="modal-dialog modal-xl modal-dialog-scrollable modal-dialog-centered" style="max-width:860px;">
    <div class="modal-content border-0 shadow">
      <div class="modal-header border-bottom px-3 py-2">
        <span class="modal-title fw-semibold text-uppercase small">
          Busca de Fornecedor
        </span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body bg-white p-3">
        <div class="d-flex gap-2 mb-3">
          <input
            type="text"
            id="inputBuscaForn"
            class="form-control form-control-sm"
            placeholder="Buscar pelo Nome ou CNPJ do Fornecedor..."
            autocomplete="off"
          />
          <button type="button" id="btnExecutarBuscaForn" class="btn btn-primary btn-sm px-3 flex-shrink-0">
            <i class="bi bi-funnel"></i>
          </button>
        </div>
        <div class="tbl-busca-forn">
          <table class="table table-bordered table-striped mb-0 tbl-busca-forn-table" id="tblBuscaForn">
            <thead>
              <tr>
                <th class="tbl-busca-th" style="width:30px">AÇÃO</th>
                <th class="tbl-busca-th" style="width:200px">NOME</th>
                <th class="tbl-busca-th" style="width:50px">CÓDIGO SAP</th>
                <th class="tbl-busca-th" style="width:80px">CNPJ</th>
              </tr>
            </thead>
            <tbody id="tbodyBuscaForn">
              <tr>
                <td colspan="4" id="tdBuscaFornEstado">
                  <div class="text-center py-4 text-secondary">
                    <i class="bi bi-funnel d-block mb-2 tbl-busca-icone"></i>
                    <span class="small">Utilize o campo acima para buscar pelo <span class="text-danger">NOME</span> ou <span class="text-danger">CNPJ.</span></span>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
      <div class="modal-footer border-top px-3 py-2">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Fechar</button>
      </div>
    </div>
  </div>
</div>

<template id="tmplLinhaDados">
  <tr>
    <td class="text-center text-nowrap"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
    <td class="text-center"></td>
  </tr>
</template>

<template id="tmplLinhaVazia">
  <tr>
    <td colspan="19">
      <div class="tbl-vazia">
        <i class="bi bi-inbox fs-5"></i>
        <p class="mt-2 mb-0">Nenhum registro encontrado.</p>
      </div>
    </td>
  </tr>
</template>

<template id="tmplToast">
  <div class="toast-item">
    <div class="toast-corpo">
      <i class="toast-icone bi fs-6 flex-shrink-0"></i>
      <span class="toast-msg flex-grow-1"></span>
      <button class="toast-fechar" aria-label="Fechar">&times;</button>
    </div>
    <div class="toast-barra"></div>
  </div>
</template>

<template id="tmplDetalheCampo">
  <div class="mb-3 campo-detalhe">
    <div class="detalhe-campo-label"></div>
    <div class="detalhe-campo-valor"></div>
  </div>
</template>

<template id="tmplDetalheSecao">
  <div class="modal-panel mb-3">
    <div class="modal-panel-header">
      <i class="bi secao-icone"></i>
      <span class="secao-texto"></span>
    </div>
    <div class="modal-panel-body">
      <div class="row secao-campos"></div>
    </div>
  </div>
</template>

<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery-mask-plugin@1.14.16/dist/jquery.mask.min.js"></script>
<script src="../../funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
<script>
(($) => {
  "use strict";

  const REPO = "CrudControleRecebimentoRepositorio.asp";
  let _pgAtual     = 1;
  let _pgTam       = 50;
  let _dadosTodos  = [];
  let _dadosFiltrados = [];
  let _modo        = "";
  let _nivelCtrl   = false;
  let _nfAtual     = "";
  let _serieAtual  = "";

  let _modalFormInst    = null;
  let _modalDetalheInst = null;
  let _modalApagarInst  = null;

  const TOAST_CFG  = Object.freeze({ delay: 4500 });
  const TOAST_ICON = Object.freeze({
    success: { cls: "bi-check-circle-fill",       cor: "#198754" },
    warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
    danger:  { cls: "bi-x-circle-fill",            cor: "#dc3545" },
    info:    { cls: "bi-info-circle-fill",          cor: "#6c757d" },
  });

  const MODO_LABELS = Object.freeze({
    INS: { txt: "Novo Registro",           cls: "badge-ins" },
    ALT: { txt: "Alterar Registro",        cls: "badge-alt" },
    APB: { txt: "Alterar Peso Balança",    cls: "badge-apb" },
    DEL: { txt: "Apagar Registro",         cls: "badge-del" },
  });

  const UI = Object.freeze({
    exibirLoader() { $("#loader").removeClass("d-none"); },
    ocultarLoader() { $("#loader").addClass("d-none"); },
    notificar(mensagem, tipo) {
      tipo = tipo || "info";
      const cfg  = TOAST_ICON[tipo] || TOAST_ICON.info;
      const node = document.getElementById("tmplToast").content.cloneNode(true);
      const wrap = node.querySelector(".toast-item");
      wrap.classList.add("toast-item-" + tipo);
      const ic = wrap.querySelector(".toast-icone");
      ic.classList.add(cfg.cls);
      ic.style.color = cfg.cor;
      wrap.querySelector(".toast-msg").textContent = mensagem;
      const barra = wrap.querySelector(".toast-barra");
      barra.classList.add("toast-barra-" + tipo);
      barra.style.animationDuration = TOAST_CFG.delay + "ms";
      const item = $(wrap);
      item.find(".toast-fechar").on("click", () => item.remove());
      $("#toastWrapper").append(item);
      setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_CFG.delay);
    },
  });

  const Html = Object.freeze({
    esc(txt) {
      if (!txt) return "";
      return String(txt)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    },
  });

  const _fmtCache = {};
  function _getFormatter(dec) {
    if (!_fmtCache[dec]) {
      _fmtCache[dec] = new Intl.NumberFormat("pt-BR", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      });
    }
    return _fmtCache[dec];
  }

  const Formato = Object.freeze({
    num(v, dec) {
      const n = parseFloat(String(v ?? "0").replace(",", "."));
      return isNaN(n) ? "0" : _getFormatter(dec).format(n);
    },
  });

  function hoje() {
    const d = new Date();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${d.getFullYear()}-${mm}-${dd}`;
  }

  const Paginacao = Object.freeze({
    totalPaginas() {
      return Math.max(1, Math.ceil(_dadosFiltrados.length / _pgTam));
    },
    dadosPagina() {
      const s = (_pgAtual - 1) * _pgTam;
      return _dadosFiltrados.slice(s, s + _pgTam);
    },
    ir(pg) {
      _pgAtual = Math.max(1, Math.min(pg, Paginacao.totalPaginas()));
      Grid.renderizarPagina();
      Paginacao.renderizarControles();
    },
    renderizarControles() {
      const total = _dadosFiltrados.length;
      if (total === 0) {
        $("#rodapeGrid").addClass("d-none");
        return;
      }
      const totalPgs = Paginacao.totalPaginas();
      $("#rodapeGrid").removeClass("d-none");
      $("#contadorGrid").html(
        `<b>${total}</b>&nbsp;registro(s) - pág.&nbsp;<b>${_pgAtual}</b>&nbsp;de&nbsp;<b>${totalPgs}</b>`
      );
      let html = `<li class="page-item${_pgAtual === 1 ? " disabled" : ""}">` +
                 `<a class="page-link" data-pg="${_pgAtual - 1}">«</a></li>`;
      const delta = 2;
      for (let p = 1; p <= totalPgs; p++) {
        if (p === 1 || p === totalPgs || (p >= _pgAtual - delta && p <= _pgAtual + delta)) {
          html += `<li class="page-item${p === _pgAtual ? " active" : ""}">` +
                  `<a class="page-link" data-pg="${p}">${p}</a></li>`;
        } else if (p === _pgAtual - delta - 1 || p === _pgAtual + delta + 1) {
          html += `<li class="page-item disabled"><a class="page-link">…</a></li>`;
        }
      }
      html += `<li class="page-item${_pgAtual === totalPgs ? " disabled" : ""}">` +
              `<a class="page-link" data-pg="${_pgAtual + 1}">»</a></li>`;
      $("#paginacaoGrid").html(html);
      $("#cmbLinhas").val(_pgTam);
    },
  });

  function dataNfParaIso(dataNf) {
    if (!dataNf) return "";
    const s = String(dataNf).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const p = s.split("/");
    if (p.length === 3) return `${p[2]}-${p[1].padStart(2,"0")}-${p[0].padStart(2,"0")}`;
    return s;
  }

  function aplicarFiltroTexto() {
    const termoNF   = $("#filtroNF").val().trim().toLowerCase();
    const termoForn = $("#filtroFornecedor").val().trim().toLowerCase();
    const termoProd = $("#filtroProduto").val().trim().toLowerCase();
    const dataIni   = $("#filtroDataIni").val();
    const dataFim   = $("#filtroDataFim").val();

    _dadosFiltrados = _dadosTodos.filter((r) => {

      if (termoNF && !String(r.nf ?? "").toLowerCase().includes(termoNF)) return false;
      if (termoForn && !String(r.nomeForn ?? "").toLowerCase().includes(termoForn)) return false;
      if (termoProd && !String(r.produto ?? "").toLowerCase().includes(termoProd)) return false;

      const dataIso = dataNfParaIso(r.dataInclusao);

      if (dataIni && dataIso && dataIso < dataIni) return false;
      if (dataFim && dataIso && dataIso > dataFim) return false;

      return true;
    });

    _pgAtual = 1;
    Grid.renderizarPagina();
    Paginacao.renderizarControles();
  }

  function construirAcoes(r) {
    const nf    = Html.esc(r.nf);
    const serie = Html.esc(r.serie);
    const btnVer = `<button class="btn-acao bg-primary text-white btn-detalhe btn-acao-ver" ` +
                   `data-nf="${nf}" data-serie="${serie}" title="Detalhe"><i class="bi bi-eye"></i></button>`;
    if (_nivelCtrl) {
      return `${btnVer}
        <button class="btn-acao btn btn-outline-warning btn-sm ms-1 btn-alt" ` +
               `data-nf="${nf}" data-serie="${serie}" title="Editar"><i class="bi bi-pencil"></i></button>
        <button class="btn-acao btn btn-outline-danger btn-sm ms-1 btn-del" ` +
               `data-nf="${nf}" data-serie="${serie}" title="Apagar"><i class="bi bi-trash"></i></button>`;
    }
    return `${btnVer}
      <button class="btn-acao btn btn-outline-warning btn-sm ms-1 btn-apb" ` +
             `data-nf="${nf}" data-serie="${serie}" title="Alterar Peso Balança"><i class="bi bi-pencil"></i></button>`;
  }

  function construirLinha(r) {
    const node = document.getElementById("tmplLinhaDados").content.cloneNode(true);
    const tds  = node.querySelectorAll("td");
    const dataHora = r.dataRecebe + (r.horaRecebe ? "\u00a0" + r.horaRecebe : "");
    tds[0].innerHTML = construirAcoes(r);
    tds[1].textContent  = r.nf + "\u00a0" + r.serie;
    tds[2].textContent  = r.dataNf;
    tds[3].textContent  = dataHora;
    tds[4].textContent  = r.cnpj;
    tds[5].textContent  = r.nomeForn;
    tds[6].textContent  = r.produto;
    tds[7].textContent  = Formato.num(r.pesoNf, 3);
    tds[8].textContent  = Formato.num(r.pesoBal, 3);
    tds[9].textContent  = Formato.num(r.umidade, 2);
    tds[10].textContent = Formato.num(r.impureza, 2);
    tds[11].textContent = Formato.num(r.proibitivo, 2);
    tds[12].textContent = Formato.num(r.outros, 2);
    tds[13].textContent = r.local;
    tds[14].textContent = r.numFardos;
    tds[15].textContent = r.fardosAnal;
    tds[16].textContent = r.placa;
    tds[17].textContent = r.motorista;
    tds[18].textContent = r.obs;
    return node;
  }

  const Grid = Object.freeze({
    renderizarLinhas(dados) {
      const tbody = document.getElementById("tbodyDados");
      if (!dados || dados.length === 0) {
        const node = document.getElementById("tmplLinhaVazia").content.cloneNode(true);
        $(tbody).empty().append(node);
        return;
      }
      const CHUNK = 50;
      let idx = 0;
      const frag = document.createDocumentFragment();

      function renderChunk() {
        const fim = Math.min(idx + CHUNK, dados.length);
        for (; idx < fim; idx++) {
          frag.appendChild(construirLinha(dados[idx]));
        }
        if (idx < dados.length) {
          requestAnimationFrame(renderChunk);
        } else {
          tbody.innerHTML = "";
          tbody.appendChild(frag);
        }
      }
      requestAnimationFrame(renderChunk);
    },
    renderizarPagina() {
      Grid.renderizarLinhas(Paginacao.dadosPagina());
    },
    carregar() {
      UI.exibirLoader();
      $.ajax({
        url: REPO,
        type: "GET",
        data: { acao: "listar", qtd: 999 },
        dataType: "json",
        success(resp) {
          UI.ocultarLoader();
          if (!resp.sucesso) {
            UI.notificar(resp.mensagem, "danger");
            return;
          }
          _dadosTodos     = resp.dados || [];
          _dadosFiltrados = structuredClone(_dadosTodos);
          _pgAtual        = 1;
          aplicarFiltroTexto();
        },
        error() {
          UI.ocultarLoader();
          UI.notificar("Erro ao carregar registros.", "danger");
        },
      });
    },
  });

  function construirCampo(label, valor, colClass) {
    const node = document.getElementById("tmplDetalheCampo").content.cloneNode(true);
    const div  = node.querySelector(".campo-detalhe");
    if (colClass) colClass.trim().split(/\s+/).forEach(c => div.classList.add(c));
    node.querySelector(".detalhe-campo-label").textContent = label;
    node.querySelector(".detalhe-campo-valor").textContent = valor || "-";
    return node;
  }

  function construirSecao(icone, titulo, campos, colClass) {
    const node = document.getElementById("tmplDetalheSecao").content.cloneNode(true);
    node.querySelector(".secao-icone").classList.add(icone);
    node.querySelector(".secao-texto").textContent = titulo;
    const row = node.querySelector(".secao-campos");
    campos.forEach(([l, v]) => row.appendChild(construirCampo(l, v, colClass)));
    return node;
  }

  function construirDetalhe(r) {
    const frag = document.createDocumentFragment();
    frag.appendChild(construirSecao("bi-file-earmark-text", "Nota Fiscal", [
      ["N.F / Série",    r.nf + " / " + r.serie],
      ["Data da N.F",    r.dataNf],
      ["Data do Teste",  r.dataRecebe],
      ["Hora do Teste",  r.horaRecebe],
    ], "col-md-3 col-6"));
    frag.appendChild(construirSecao("bi-building", "Fornecedor", [
      ["CNPJ",         r.cnpj],
      ["Razão Social", r.nomeForn],
    ], "col-md-4 col-12"));
    frag.appendChild(construirSecao("bi-box-seam", "Produto e Pesos", [
      ["Produto",      r.produto],
      ["Peso N.F",     Formato.num(r.pesoNf, 2)],
      ["Peso Balança", Formato.num(r.pesoBal, 2)],
    ], "col-md-4 col-6"));
    frag.appendChild(construirSecao("bi-percent", "Percentuais de Qualidade", [
      ["% Umidade",    Formato.num(r.umidade, 2)],
      ["% Impureza",   Formato.num(r.impureza, 2)],
      ["% Proibitivo", Formato.num(r.proibitivo, 2)],
      ["% Outros",     Formato.num(r.outros, 2)],
    ], "col-md-3 col-6"));
    frag.appendChild(construirSecao("bi-truck", "Armazenagem e Transporte", [
      ["Local",       r.local],
      ["Quantidade",  r.numFardos],
      ["Fardos An.",  r.fardosAnal],
      ["Placa",       r.placa],
      ["Motorista",   r.motorista],
      ["Observação",  r.obs],
    ], "col-md-2 col-6"));
    return frag;
  }

  function abrirDetalhe(nf, serie) {
    UI.exibirLoader();
    $.ajax({
      url: REPO,
      type: "GET",
      data: { acao: "buscar_registro", nf, serie },
      dataType: "json",
      success(resp) {
        UI.ocultarLoader();
        if (!resp.sucesso) {
          UI.notificar(resp.mensagem, "danger");
          return;
        }
        const body = document.getElementById("modalDetalheBody");
        body.innerHTML = "";
        body.appendChild(construirDetalhe(resp.dado));
        if (!_modalDetalheInst) {
          _modalDetalheInst = bootstrap.Modal.getOrCreateInstance(
            document.getElementById("modalDetalhe")
          );
        }
        _modalDetalheInst.show();
      },
      error() {
        UI.ocultarLoader();
        UI.notificar("Erro ao carregar detalhe.", "danger");
      },
    });
  }

  function decParaMascara(v, casas) {
    const n = parseFloat(String(v ?? "0").replace(",", "."));
    if (isNaN(n)) return "";
    return n.toFixed(casas).replace(".", ",");
  }

  function percParaEnvio(id) {
    const raw = $("#" + id).val().replace(",", ".");
    const n   = parseFloat(raw);
    return isNaN(n) ? "0" : String(n);
  }

  function getModalForm() {
    if (!_modalFormInst) {
      _modalFormInst = bootstrap.Modal.getOrCreateInstance(
        document.getElementById("modalForm")
      );
    }
    return _modalFormInst;
  }

  const Formulario = Object.freeze({
    limpar() {
      $("#edtNF,#edtSERIE,#edtDATA_NF,#edtDATA_RECEBE,#edtHORA_RECEBE").val("");
      $("#edtCNPJ,#edtNOME_FORN,#edtPRODUTO").val("");
      $("#edtPESO_NF,#edtPESO_BAL,#edtUMIDADE,#edtIMPUREZA,#edtPROIBITIVO,#edtOUTROS").val("");
      $("#edtLOCAL,#edtNUM_FARDOS,#edtFARDOS,#edtMOTORISTA,#edtPLACA,#edtOBS").val("");
    },
    preencher(d) {
      $("#edtNF").val(d.nf);
      $("#edtSERIE").val(d.serie);
      $("#edtDATA_NF").val(d.dataNf);
      $("#edtDATA_RECEBE").val(d.dataRecebe);
      $("#edtHORA_RECEBE").val((d.horaRecebe || "").substring(0, 5));
      $("#edtCNPJ").val(d.cnpj);
      $("#edtNOME_FORN").val(d.nomeForn);
      $("#edtPRODUTO").val(d.produto);
      $("#edtPESO_NF").val(decParaMascara(d.pesoNf, 3));
      $("#edtPESO_BAL").val(decParaMascara(d.pesoBal, 3));
      $("#edtUMIDADE").val(decParaMascara(d.umidade, 2));
      $("#edtIMPUREZA").val(decParaMascara(d.impureza, 2));
      $("#edtPROIBITIVO").val(decParaMascara(d.proibitivo, 2));
      $("#edtOUTROS").val(decParaMascara(d.outros, 2));
      $("#edtLOCAL").val(d.local);
      $("#edtNUM_FARDOS").val(d.numFardos);
      $("#edtFARDOS").val(d.fardosAnal);
      $("#edtMOTORISTA").val(d.motorista);
      $("#edtPLACA").val(d.placa);
      $("#edtOBS").val(d.obs);
    },
    configurarCampos(modo) {
      const soLeitura = modo === "DEL";
      const camposTodos =
        "#edtNF,#edtSERIE,#edtDATA_NF,#edtDATA_RECEBE,#edtHORA_RECEBE," +
        "#edtPRODUTO,#edtPESO_NF,#edtUMIDADE,#edtIMPUREZA,#edtPROIBITIVO," +
        "#edtOUTROS,#edtLOCAL,#edtNUM_FARDOS,#edtFARDOS,#edtMOTORISTA,#edtPLACA,#edtOBS";

      $(camposTodos).prop("readonly", soLeitura);
      $("#edtPESO_BAL").prop("readonly", soLeitura);
      $("#edtNF,#edtSERIE").prop("readonly", modo !== "INS");

      const mostraBtnsForn = (modo === "INS" || modo === "ALT");
      $("#btnsFornecedor").toggle(mostraBtnsForn);
      $("#btnPesquisarFornecedor").toggle(modo === "INS");
      $("#btnLimparFornecedor").toggle(modo === "INS");

      $("#btnSalvar").toggleClass("d-none", modo === "DEL");
      $("#btnApagar").toggleClass("d-none", modo !== "DEL");

      if (modo === "APB") {
        $(camposTodos).prop("readonly", true);
        $("#edtPESO_BAL").prop("readonly", false);
      }

      $("#spinnerSalvar,#spinnerApagar").addClass("d-none");
      $("#btnSalvar,#btnApagar").prop("disabled", false);
    },
    abrir(modo, nf, serie) {
      _modo      = modo;
      _nfAtual   = nf   || "";
      _serieAtual = serie || "";

      const lbl = MODO_LABELS[modo] || { txt: modo, cls: "" };
      $("#labelModalForm").text("Registro");
      $("#formModeBadge").text(lbl.txt).attr("class", "form-mode-badge " + lbl.cls);

      Formulario.limpar();
      Formulario.configurarCampos(modo);

      if (modo === "INS") {
        getModalForm().show();
        return;
      }

      UI.exibirLoader();
      $.ajax({
        url: REPO,
        type: "GET",
        data: { acao: "buscar_registro", nf, serie },
        dataType: "json",
        success(resp) {
          UI.ocultarLoader();
          if (!resp.sucesso) {
            UI.notificar(resp.mensagem, "danger");
            return;
          }

          if (modo === "DEL") {
            $("#apagarNfSerie").text(resp.dado.nf + " / " + resp.dado.serie);
            if (!_modalApagarInst) {
              _modalApagarInst = bootstrap.Modal.getOrCreateInstance(
                document.getElementById("modalConfirmarApagar")
              );
            }
            _modalApagarInst.show();
            return;
          }
          Formulario.preencher(resp.dado);
          getModalForm().show();
        },
        error() {
          UI.ocultarLoader();
          UI.notificar("Erro ao carregar registro.", "danger");
        },
      });
    },
    salvar() {
      const campos = [
        { id: "edtNF",          msg: "Número da Nota Fiscal não informado." },
        { id: "edtSERIE",       msg: "Série não informada." },
        { id: "edtDATA_NF",     msg: "Data da N.F. não informada." },
        { id: "edtDATA_RECEBE", msg: "Data do Teste não informada." },
        { id: "edtHORA_RECEBE", msg: "Hora do Teste não informada." },
        { id: "edtCNPJ",        msg: "Fornecedor não selecionado." },
        { id: "edtPRODUTO",     msg: "Produto não informado." },
        { id: "edtPESO_NF",     msg: "Peso da N.F. não informado." },
        { id: "edtUMIDADE",     msg: "% Umidade não informado." },
        { id: "edtMOTORISTA",   msg: "Nome do Motorista não informado." },
        { id: "edtPLACA",       msg: "Placa não informada." },
      ];

      const camposVal = _modo === "APB"
        ? [{ id: "edtPESO_BAL", msg: "Peso Balança não informado." }]
        : campos;

      for (const c of camposVal) {
        if ($("#" + c.id).val().trim() === "") {
          UI.notificar(c.msg, "warning");
          $("#" + c.id).focus();
          return;
        }
      }

      if (_modo === "INS") {
        const nfVal    = $("#edtNF").val().trim();
        const serieVal = $("#edtSERIE").val().trim();

        $("#spinnerSalvar").removeClass("d-none");
        $("#btnSalvar").prop("disabled", true);

        $.ajax({
          url: REPO,
          type: "POST",
          data: { acao: "verificar_nf", nf: nfVal, serie: serieVal, cnpj: $("#edtCNPJ").val().trim() },
          dataType: "json",
          success(resp) {
            if (!resp.sucesso) {
              $("#spinnerSalvar").addClass("d-none");
              $("#btnSalvar").prop("disabled", false);
              UI.notificar("Erro ao verificar Nota Fiscal.", "danger");
              return;
            }
            if (resp.existe) {
              $("#spinnerSalvar").addClass("d-none");
              $("#btnSalvar").prop("disabled", false);
              UI.notificar(
                "NF " + nfVal + " / Série " + serieVal + " já está cadastrada para este fornecedor nesta filial.",
                "warning"
              );
              $("#edtNF").focus();
              return;
            }
            Formulario._enviarSalvar();
          },
          error() {
            $("#spinnerSalvar").addClass("d-none");
            $("#btnSalvar").prop("disabled", false);
            UI.notificar("Erro ao verificar Nota Fiscal.", "danger");
          },
        });
        return;
      }

      Formulario._enviarSalvar();
    },
    _enviarSalvar() {
      $("#spinnerSalvar").removeClass("d-none");
      $("#btnSalvar").prop("disabled", true);
      UI.exibirLoader();

      $.ajax({
        url: REPO,
        type: "POST",
        data: {
          acao:        "salvar",
          tipoAcao:    _modo,
          nf:          $("#edtNF").val().trim(),
          serie:       $("#edtSERIE").val().trim(),
          dataNf:      $("#edtDATA_NF").val(),
          dataRecebe:  $("#edtDATA_RECEBE").val(),
          horaRecebe:  $("#edtHORA_RECEBE").val().substring(0, 5),
          cnpj:        $("#edtCNPJ").val().trim(),
          nomeForn:    $("#edtNOME_FORN").val().trim(),
          produto:     $("#edtPRODUTO").val().trim(),
          pesoNf:      $("#edtPESO_NF").val().trim(),
          pesoBal:     $("#edtPESO_BAL").val().trim(),
          umidade:     percParaEnvio("edtUMIDADE"),
          impureza:    percParaEnvio("edtIMPUREZA"),
          proibitivo:  percParaEnvio("edtPROIBITIVO"),
          outros:      percParaEnvio("edtOUTROS"),
          local:       $("#edtLOCAL").val().trim(),
          numFardos:   $("#edtNUM_FARDOS").val().trim(),
          fardosAnal:  $("#edtFARDOS").val().trim(),
          motorista:   $("#edtMOTORISTA").val().trim(),
          placa:       $("#edtPLACA").val().trim(),
          obs:         $("#edtOBS").val().trim(),
        },
        dataType: "json",
        success(resp) {
          UI.ocultarLoader();
          $("#spinnerSalvar").addClass("d-none");
          $("#btnSalvar").prop("disabled", false);
          if (!resp.sucesso) {
            UI.notificar(resp.mensagem, "danger");
            return;
          }
          UI.notificar(resp.mensagem, "success");
          getModalForm().hide();
          Grid.carregar();
        },
        error() {
          UI.ocultarLoader();
          $("#spinnerSalvar").addClass("d-none");
          $("#btnSalvar").prop("disabled", false);
          UI.notificar("Erro ao salvar registro.", "danger");
        },
      });
    },
    apagar() {
      $("#spinnerApagar").removeClass("d-none");
      $("#btnConfirmarApagar").prop("disabled", true);
      UI.exibirLoader();

      $.ajax({
        url: REPO,
        type: "POST",
        data: { acao: "apagar", nf: _nfAtual, serie: _serieAtual },
        dataType: "json",
        success(resp) {
          UI.ocultarLoader();
          $("#spinnerApagar").addClass("d-none");
          $("#btnConfirmarApagar").prop("disabled", false);
          if (_modalApagarInst) _modalApagarInst.hide();
          if (!resp.sucesso) {
            UI.notificar(resp.mensagem, "danger");
            return;
          }
          UI.notificar(resp.mensagem, "success");
          Grid.carregar();
        },
        error() {
          UI.ocultarLoader();
          $("#spinnerApagar").addClass("d-none");
          $("#btnConfirmarApagar").prop("disabled", false);
          UI.notificar("Erro ao apagar registro.", "danger");
        },
      });
    },
  });

  // Validação de CNPJ — suporta o formato alfanumérico (Resolução DREI nº 59/2024,
  // vigente a partir de 01/07/2026) e o formato numérico legado.
  // O algoritmo de dígitos verificadores é aplicado apenas a CNPJs 100% numéricos
  // (legado). Para CNPJs alfanuméricos, valida-se exclusivamente o formato, pois
  // o novo algoritmo de DV ainda não foi publicado oficialmente pela Receita Federal.
  function validarCNPJ(cnpj) {
    // Remove máscara: pontos, barras e hífens
    var raw = String(cnpj).replace(/[.\-\/]/g, "").toUpperCase();
    if (raw.length !== 14) return false;

    // Formato: 12 caracteres alfanuméricos (A-Z0-9) + 2 dígitos verificadores numéricos
    if (!/^[A-Z0-9]{12}\d{2}$/.test(raw)) return false;

    // Se os 12 primeiros são todos dígitos, aplica validação aritmética clássica (legado)
    if (/^\d{14}$/.test(raw)) {
      if (/^(\d)\1{13}$/.test(raw)) return false;
      var t = 12, n = raw.slice(0, t), d = raw.slice(t), s = 0, p = t - 7, i;
      for (i = t; i >= 1; i--) { s += n[t - i] * p--; if (p < 2) p = 9; }
      if ((s % 11 < 2 ? 0 : 11 - s % 11) != d[0]) return false;
      t++; n = raw.slice(0, t); s = 0; p = t - 7;
      for (i = t; i >= 1; i--) { s += n[t - i] * p--; if (p < 2) p = 9; }
      return (s % 11 < 2 ? 0 : 11 - s % 11) == d[1];
    }

    // CNPJ alfanumérico: formato válido, DV não verificado (algoritmo pendente de publicação)
    return true;
  }

  // Retorna true se o valor informado parece um CNPJ alfanumérico (novo formato)
  function isCNPJAlfanumerico(cnpj) {
    var raw = String(cnpj).replace(/[.\-\/]/g, "").toUpperCase();
    return raw.length === 14 && /[A-Z]/.test(raw.slice(0, 12));
  }

  // Aplica máscara XX.XXX.XXX/XXXX-DD a uma string de 14 chars sem pontuação
  function mascararCNPJ(raw) {
    raw = raw.toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (raw.length > 14) raw = raw.slice(0, 14);
    var out = raw;
    if (raw.length > 12) out = raw.slice(0,2)+"."+raw.slice(2,5)+"."+raw.slice(5,8)+"/"+raw.slice(8,12)+"-"+raw.slice(12);
    else if (raw.length > 8)  out = raw.slice(0,2)+"."+raw.slice(2,5)+"."+raw.slice(5,8)+"/"+raw.slice(8);
    else if (raw.length > 5)  out = raw.slice(0,2)+"."+raw.slice(2,5)+"."+raw.slice(5);
    else if (raw.length > 2)  out = raw.slice(0,2)+"."+raw.slice(2);
    return out;
  }

  window.pesquisarFornecedor = function () {
    BuscaFornecedor.abrir();
  };

  window.limparFornecedor = function () {
    $("#edtCNPJ,#edtNOME_FORN").val("");
  };

  function inicializarMascaras() {
    $("#edtPESO_NF,#edtPESO_BAL").mask("####0,000", { reverse: true });

    const camposPerc = ["edtUMIDADE", "edtIMPUREZA", "edtPROIBITIVO", "edtOUTROS"];
    camposPerc.forEach(function (id) {
      const $el = $("#" + id);

      $el.on("keypress", function (e) {
        if (!/\d/.test(String.fromCharCode(e.which))) e.preventDefault();
      });

      $el.on("input", function () {
        const digits = this.value.replace(/\D/g, "");
        if (digits === "") { this.value = ""; return; }

        let formatted;
        if (digits.length < 3) {
          formatted = String(parseInt(digits, 10));
        } else {
          const intPart = String(parseInt(digits.slice(0, digits.length - 2), 10));
          const decPart = digits.slice(digits.length - 2);
          const intVal  = parseInt(intPart, 10);
          formatted = intVal >= 100 ? "100,00" : intPart + "," + decPart;
        }
        this.value = formatted;

        const len = this.value.length;
        this.setSelectionRange(len, len);
      });

      $el.on("blur", function () {
        if (this.value === "" || this.value === ",") { this.value = "0,00"; return; }
        const n = parseFloat(this.value.replace(",", "."));
        if (isNaN(n)) { this.value = "0,00"; return; }
        this.value = Math.min(n, 100).toFixed(2).replace(".", ",");
      });

      $el.on("focus", function () {
        if (this.value === "0,00") this.value = "";
      });
    });
  }

  const BuscaFornecedor = (() => {
    let _inst = null;

    function _getInst() {
      if (!_inst) _inst = bootstrap.Modal.getOrCreateInstance(document.getElementById("modalBuscaFornecedor"));
      return _inst;
    }

    function _estadoInicial() {
      $("#inputBuscaForn").val("");
      $("#tbodyBuscaForn").html(
        `<tr><td colspan="4"><div class="text-center py-4 text-secondary small">` +
        `<i class="bi bi-funnel tbl-busca-icone d-block mb-2"></i>` +
        `Utilize o campo acima para buscar pelo <span class="text-danger">NOME</span> ou <span class="text-danger">CNPJ.</span>` +
        `</div></td></tr>`
      );
    }

    function _renderLinhas(dados) {
      if (!dados || dados.length === 0) {
        $("#tbodyBuscaForn").html(
          `<tr><td colspan="4"><div class="text-center py-4 text-secondary small">` +
          `<i class="bi bi-inbox tbl-busca-icone d-block mb-2"></i>` +
          `Nenhum fornecedor encontrado.</div></td></tr>`
        );
        return;
      }
      const linhas = dados.map((f) =>
        `<tr>` +
        `<td class="text-center">` +
        `<button type="button" class="btn btn-primary btn-sm py-0 px-2 btn-importar-forn" ` +
        `data-cnpj="${Html.esc(f.cnpj)}" data-nome="${Html.esc(f.nome)}" title="Importar">` +
        `<i class="bi bi-box-arrow-in-down"></i></button>` +
        `</td>` +
        `<td class="text-center">${Html.esc(f.nome)}</td>` +
        `<td class="text-center">${Html.esc(f.codigo)}</td>` +
        `<td class="text-center">${Html.esc(f.cnpj)}</td>` +
        `</tr>`
      );
      $("#tbodyBuscaForn").html(linhas.join(""));
    }

    function _pesquisar() {
      const parametro = $("#inputBuscaForn").val().trim();
      if (!parametro) return;
      $("#tbodyBuscaForn").html(
        `<tr><td colspan="4"><div class="text-center py-3 text-secondary small">` +
        `<span class="spinner-border spinner-border-sm me-2"></span>Buscando...</div></td></tr>`
      );
      $.ajax({
        url: REPO,
        type: "GET",
        data: { acao: "buscar_fornecedores", parametro },
        dataType: "json",
        success(resp) {
          if (!resp.sucesso) { UI.notificar(resp.mensagem, "danger"); return; }
          _renderLinhas(resp.dados);
        },
        error() {
          UI.notificar("Erro ao buscar fornecedores.", "danger");
        },
      });
    }

    function _importar(cnpj, nome) {
      $("#edtCNPJ").val(cnpj);
      $("#edtNOME_FORN").val(nome);
      _getInst().hide();
    }

    function _inicializar() {
      const elModal = document.getElementById("modalBuscaFornecedor");
      elModal.addEventListener("show.bs.modal", _estadoInicial);
      elModal.addEventListener("shown.bs.modal", () => $("#inputBuscaForn").focus());

      elModal.addEventListener("show.bs.modal", () => {
        const ov = document.createElement("div");
        ov.id = "overlayBuscaForn";
        ov.style.cssText = "position:absolute;inset:0;background:rgba(0,0,0,0.45);border-radius:inherit;z-index:10;transition:opacity .15s ease;";
        const mc = document.querySelector("#modalForm .modal-content");
        if (mc) { mc.style.position = "relative"; mc.appendChild(ov); }
      });

      elModal.addEventListener("hide.bs.modal", () => {
        const ov = document.getElementById("overlayBuscaForn");
        if (ov) ov.remove();
      });

      $("#btnExecutarBuscaForn").on("click", _pesquisar);

      $("#inputBuscaForn").on("keydown", (e) => {
        if (e.key === "Enter") _pesquisar();
      });

      $("#tbodyBuscaForn").on("click", ".btn-importar-forn", function () {
        _importar($(this).data("cnpj"), $(this).data("nome"));
      });
    }

    return Object.freeze({
      inicializar: _inicializar,
      abrir() { _getInst().show(); },
    });
  })();

  function inicializarModalFornecedor() {
    document.getElementById("modalCadFornece").addEventListener("show.bs.modal", () => {
      $("#codifornece,#cnpjfornece,#nomefornece").val("");
      $("#feedbackCnpj").text("").attr("class", "");
    });

    $("#cnpjfornece").on("input", function () {
      // Permite letras (A-Z) e digitos; aplica mascara progressiva XX.XXX.XXX/XXXX-DD
      var raw = this.value.toUpperCase().replace(/[^A-Z0-9]/g, "");
      this.value = mascararCNPJ(raw);

      var fb  = $("#feedbackCnpj");
      var sem = this.value.replace(/[.\-\/]/g, "");
      if (sem.length === 14) {
        var ok      = validarCNPJ(this.value);
        var alfanum = isCNPJAlfanumerico(this.value);
        var msg     = ok ? (alfanum ? "CNPJ alfanumerico - formato valido" : "CNPJ valido") : "CNPJ invalido";
        fb.text(msg).attr("class", ok ? "valido" : "invalido");
      } else {
        fb.text("").attr("class", "");
      }
    });

    $("#btnSalvarFornecedor").on("click", function () {
      const codigo = $("#codifornece").val().trim();
      const cnpj   = $("#cnpjfornece").val().trim();
      const nome   = $("#nomefornece").val().trim();
      if (!codigo) { UI.notificar("Informe o código SAP.", "warning");    return; }
      if (!cnpj)   { UI.notificar("Informe o CNPJ", "warning");      return; }
      if (!nome)   { UI.notificar("Informe a razão social.", "warning");  return; }
      if (!validarCNPJ(cnpj)) { UI.notificar("CNPJ inválido.", "danger"); return; }

      UI.exibirLoader();
      $.ajax({
        url: REPO,
        type: "POST",
        data: { acao: "cadastrar_fornecedor", codigo, cnpj, nome },
        dataType: "json",
        success(resp) {
          UI.ocultarLoader();
          if (!resp.sucesso) { UI.notificar(resp.mensagem, "danger"); return; }
          UI.notificar(resp.mensagem, "success");
          $("#edtCNPJ").val(cnpj);
          $("#edtNOME_FORN").val(nome);
          bootstrap.Modal.getInstance(document.getElementById("modalCadFornece")).hide();
        },
        error() {
          UI.ocultarLoader();
          UI.notificar("Erro ao cadastrar fornecedor.", "danger");
        },
      });
    });
  }

  function carregarNivel() {
    $.ajax({
      url: REPO,
      type: "GET",
      data: { acao: "nivel" },
      dataType: "json",
      success(resp) {
        if (resp.sucesso) _nivelCtrl = resp.nivel === 1;
      },
    });
  }

  function debounce(fn, delay) {
    let t;
    return function (...args) {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(this, args), delay);
    };
  }

  function anexarEventos() {

    $("#btnCriar").on("click", () => Formulario.abrir("INS"));

    $("#btnSalvar").on("click", () => Formulario.salvar());

    $("#btnConfirmarApagar").on("click", () => Formulario.apagar());

    $("#filtroNF").on("input", function () {
      this.value = this.value.replace(/\D/g, "");
    });

    $("#btnFiltrar").on("click", aplicarFiltroTexto);

    $("#filtroNF,#filtroFornecedor,#filtroProduto,#filtroDataIni,#filtroDataFim").on("keydown", function (e) {
      if (e.key === "Enter") aplicarFiltroTexto();
    });

    $("#btnLimparFiltros").on("click", function () {
      $("#filtroNF").val("");
      $("#filtroFornecedor").val("");
      $("#filtroProduto").val("");
      $("#filtroDataIni").val("");
      $("#filtroDataFim").val(hoje());
      aplicarFiltroTexto();
    });

    $("#cmbLinhas").on("change", function () {
      _pgTam   = parseInt(this.value) || 50;
      _pgAtual = 1;
      Grid.renderizarPagina();
      Paginacao.renderizarControles();
    });

    $("#paginacaoGrid").on("click", ".page-link", function (e) {
      e.preventDefault();
      const pg = parseInt($(this).data("pg"));
      if (!isNaN(pg)) Paginacao.ir(pg);
    });

    $("#tbodyDados").on("click", ".btn-detalhe", function () {
      abrirDetalhe($(this).data("nf"), $(this).data("serie"));
    });
    $("#tbodyDados").on("click", ".btn-alt", function () {
      Formulario.abrir("ALT", $(this).data("nf"), $(this).data("serie"));
    });
    $("#tbodyDados").on("click", ".btn-del", function () {
      Formulario.abrir("DEL", $(this).data("nf"), $(this).data("serie"));
    });
    $("#tbodyDados").on("click", ".btn-apb", function () {
      Formulario.abrir("APB", $(this).data("nf"), $(this).data("serie"));
    });

    document.addEventListener("keydown", function (e) {
      if (e.ctrlKey && e.key === "k") {
        e.preventDefault();
        document.getElementById("filtroNF").focus();
      }
    });
  }

  $(function () {

    $("#filtroDataFim").val(hoje());

    UI.ocultarLoader();
    carregarNivel();
    inicializarMascaras();
    inicializarModalFornecedor();
    BuscaFornecedor.inicializar();
    anexarEventos();
    Grid.carregar();
  });

})(jQuery);
</script>
</body>
</html>
