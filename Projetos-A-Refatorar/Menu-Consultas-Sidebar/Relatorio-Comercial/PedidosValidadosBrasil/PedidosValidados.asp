<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp"-->
<%

Session.LCID = 1046
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Pedidos Validados</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
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

  /* ── Filter card ─────────────────────────────── */
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

  /* ── Table card ──────────────────────────────── */
  .table-card {
    background-color: #fff;
    border: 1px solid #dee2e6;
    border-radius: 6px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    position: relative;
  }

  /* ── Toolbar ─────────────────────────────────── */
  .table-toolbar {
    background-color: #fff;
    border-radius: 6px 6px 0 0;
  }

  /* ── Search box ──────────────────────────────── */
  .search-box {
    position: relative;
  }
  .search-box .search-hotkey {
    position: absolute;
    right: 0.5rem;
    top: 50%;
    transform: translateY(-50%);
    display: flex;
    gap: 0.15rem;
    pointer-events: none;
  }
  .search-box .form-control {
    padding-right: 5rem;
  }
  .search-hotkey kbd {
    font-size: 0.65rem;
    padding: 0.1rem 0.3rem;
    border-radius: 3px;
    background: #f1f3f5;
    border: 1px solid #ced4da;
    color: #6c757d;
  }

  /* ── Table wrapper ───────────────────────────── */
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

  /* ── Footer ──────────────────────────────────── */
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
    min-width: 160px;
    display: flex;
    align-items: center;
    gap: 0.5rem;
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
  .info-registros {
    font-size: 0.8rem;
    color: #6c757d;
  }
  .select-limite {
    width: auto;
  }

  /* ── Placeholder ─────────────────────────────── */
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

  /* ── Overlay loading ─────────────────────────── */
  .overlay-loading {
    position: absolute;
    inset: 0;
    background: rgba(255, 255, 255, 0.8);
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    z-index: 10;
    gap: 1rem;
    border-radius: 6px;
  }
  .overlay-loading .spinner-border {
    width: 2.5rem;
    height: 2.5rem;
    color: #6c757d;
  }
  .overlay-loading span {
    font-size: 0.9rem;
    font-weight: 600;
    color: #495057;
  }

  /* ── Toast ───────────────────────────────────── */
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
    from { opacity: 0; transform: translateX(40px); }
    to   { opacity: 1; transform: translateX(0);    }
  }
  .toast-custom .toast-progress-bar {
    animation: toastProgressShrink 4.5s linear forwards;
  }
  @keyframes toastProgressShrink {
    from { width: 100%; }
    to   { width: 0%;   }
  }

  /* ── Paginacao — secondary ───────────────────── */
  .pagination .page-link {
    color: #6c757d;
    border-color: #dee2e6;
  }
  .pagination .page-link:hover {
    color: #495057;
    background-color: #e9ecef;
    border-color: #dee2e6;
  }
  .pagination .page-item.active .page-link {
    background-color: #6c757d;
    border-color: #6c757d;
    color: #fff;
  }
  .pagination .page-item.disabled .page-link {
    color: #adb5bd;
    border-color: #dee2e6;
  }

  /* ── Linha vermelha — StatusPedido = 4 ──────── */
  /* Neutraliza o striped do Bootstrap redefinindo as variáveis na própria <tr> */
  #tbodyDados tr.tr-status-4 {
    --bs-table-striped-bg: #fff0f0;
    --bs-table-hover-bg:   #ffd6d6;
    --bs-table-bg:         #fff0f0;
    --bs-table-color:      #842029;
    --bs-table-striped-color: #842029;
  }
  #tbodyDados tr.tr-status-4 td {
    background-color: #fff0f0 !important;
    color: #842029 !important;
  }
  #tbodyDados tr.tr-status-4:hover td {
    background-color: #ffd6d6 !important;
  }

  /* ── Datas ocultas no modo EE ────────────────── */
  .div-dt-entrega {
    display: none;
  }

</style>

</head>
<body>

<!-- Toast container -->
<div class="toast-container-custom" id="toastContainer"></div>

<div class="container-fluid py-3 px-3 px-lg-4">

  <!-- Cabecalho da pagina -->
  <div class="mb-3">
    <span class="page-title">Pedidos Validados Brasil</span>
  </div>

  <!-- ── Card de Filtros ──────────────────────────────────────────────────── -->
  <div class="filter-card mb-3">
    <div class="filter-card-header">
      <i class="bi bi-funnel me-2"></i>Filtros
    </div>
    <div class="filter-card-body">

      <!-- Linha 1: Filial · Produto · Buscar Produto · Pedido -->
      <div class="row g-2 mb-2">

        <div class="col-3">
          <label for="cboFilial" class="filter-label d-block">Filial</label>
          <select id="cboFilial" class="form-select form-select-sm">
            <option value="">Selecione...</option>
            <option value="05">Bento Gonçalves</option>
            <option value="06">Maranguape</option>
            <option value="07">Saquarema</option>
          </select>
        </div>

        <div class="col-3">
          <label for="cboProduto" class="filter-label d-block">Tipo Produto</label>
          <select id="cboProduto" class="form-select form-select-sm">
            <option value="">Todos</option>
            <option value="CX">Caixa</option>
            <option value="CH">Chapa</option>
            <option value="PA">Papel</option>
          </select>
        </div>

        <div class="col-3">
          <label class="filter-label d-block">Buscar Produto</label>
          <div class="input-group input-group-sm">
            <button type="button" class="btn btn-secondary btn-sm" onclick="pesquisarProduto()" title="Buscar produto">
              Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
            </button>
            <input type="text" class="form-control form-control-sm" id="nomeProduto"
                   placeholder="Código ou descrição"
                   autocomplete="off" />
            <button type="button" class="btn btn-outline-secondary btn-sm"
                    onclick="limparCamposProduto()" title="Limpar produto">
              <i class="bi bi-x-lg"></i>
            </button>
          </div>
          <input type="hidden" id="codProduto" value="" />
        </div>

        <div class="col-3">
          <label for="pedido" class="filter-label d-block">Pedido</label>
          <input type="text" id="pedido"
                 class="form-control form-control-sm"
                 autocomplete="off" />
        </div>

      </div>

      <!-- Linha 2: Cliente · Tipo · Data Inicial · Data Final (+ datas EE ocultas) -->
      <div class="row g-2 align-items-end mb-2">

        <!-- Cliente -->
        <div class="col-3" id="colCliente">
          <label class="filter-label d-block">Cliente</label>
          <div class="input-group input-group-sm">
            <button type="button" class="btn btn-secondary btn-sm" onclick="pesquisarCliente()" title="Buscar cliente">
              Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
            </button>
            <input type="text" class="form-control form-control-sm" id="codcliVisivel"
                   placeholder="Código do cliente"
                   autocomplete="off" maxlength="10" />
            <button type="button" class="btn btn-outline-secondary btn-sm" id="btnLimparCliente"
                    onclick="limparCamposCliente()" title="Limpar cliente">
              <i class="bi bi-x-lg"></i>
            </button>
          </div>
          <input type="hidden" id="nomeCliente" />
          <input type="hidden" id="codloja" />
        </div>

        <!-- Tipo -->
        <div class="col-3">
          <label for="cboTipo" class="filter-label d-block">Tipo</label>
          <select id="cboTipo" class="form-select form-select-sm">
            <option value="">Selecione...</option>
            <option value="EM">Emissão</option>
            <option value="VL">Validação</option>
            <option value="EN">Entrega</option>
            <option value="EE">Emissão e Entrega</option>
          </select>
        </div>

        <!-- Data inicial principal -->
        <div class="col-3">
          <label for="dataInicial1" class="filter-label d-block" id="lblDataIni1Vis">Data Inicial</label>
          <input type="date" id="dataInicial1" class="form-control form-control-sm" />
        </div>

        <!-- Data final principal -->
        <div class="col-3">
          <label for="dataFinal1" class="filter-label d-block" id="lblDataFin1Vis">Data Final</label>
          <input type="date" id="dataFinal1" class="form-control form-control-sm" />
        </div>

        <!-- Dt. Entrega Inicial (EE) — oculta por padrao -->
        <div class="col-3 div-dt-entrega" id="divDtIniEnt">
          <label for="dataInicial2" class="filter-label d-block">Dt. Entrega Inicial</label>
          <input type="date" id="dataInicial2" class="form-control form-control-sm" />
        </div>

        <!-- Dt. Entrega Final (EE) — oculta por padrao -->
        <div class="col-3 div-dt-entrega" id="divDtFinEnt">
          <label for="dataFinal2" class="filter-label d-block">Dt. Entrega Final</label>
          <input type="date" id="dataFinal2" class="form-control form-control-sm" />
        </div>

      </div>

      <!-- Linha 3: Switch Lote Piloto + Botoes -->
      <div class="row g-2 align-items-center">
        <div class="col-12 d-flex align-items-center justify-content-between">

          <!-- Switch -->
          <div class="form-check form-switch mb-0">
            <input class="form-check-input" type="checkbox" id="cbxLotePiloto" />
            <label class="form-check-label filter-label mb-0" for="cbxLotePiloto">
              Somente Lote Piloto
            </label>
          </div>

          <!-- Botoes -->
          <div class="d-flex gap-2">
            <button type="button" class="btn btn-light btn-sm" onclick="limparFiltro()" title="Limpar filtros">
              <i class="bi bi-eraser me-2"></i>Limpar
            </button>
            <button type="button" class="btn btn-primary btn-sm" id="btnFiltrar">
              <i class="bi bi-funnel-fill me-2"></i>Filtrar
            </button>
          </div>

        </div>
      </div>

    </div>
  </div>

  <!-- ── Legenda ─────────────────────────────────────────────────────────── -->
  <div class="filter-card mb-3">
    <div class="filter-card-header">
      <i class="bi bi-info-circle me-2"></i>Legenda
    </div>
    <div class="filter-card-body">
      <span style="font-size:0.78rem;font-weight:600;padding:0.1rem 0.6rem;border-radius:4px;background:red;color:white">Cancelado</span>
    </div>
  </div>

  <!-- ── Card de Tabela ───────────────────────────────────────────────────── -->
  <div class="table-card">

    <!-- Toolbar -->
    <div id="tableToolbar" class="table-toolbar d-none d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
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

    <!-- Tabela / Placeholder -->
    <div id="areaTabela" style="position:relative">

      <!-- Overlay de carregamento -->
      <div id="overlayLoading" class="overlay-loading d-none">
        <div class="spinner-border" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <div class="table-placeholder" id="placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>

      <div class="table-wrapper d-none" id="wrapperTabela">
        <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tabelaDados">
          <thead>
            <tr>
              <th>Pedido</th>
              <th>Cliente</th>
              <th>Razão</th>
              <th>Pedido Cliente</th>
              <th>Produto</th>
              <th>Referência</th>
              <th>TES</th>
              <th>Qtde.</th>
              <th>Qtde. Faturada</th>
              <th>Saldo Qtde.</th>
              <th>Área Total</th>
              <th>Preço Inclusão</th>
              <th>Preço Atual</th>
              <th>Preço KG NET</th>
              <th>Entrega</th>
              <th>Status</th>
              <th>Emissão</th>
              <th>Emitiu</th>
              <th>Validação</th>
              <th>Validou</th>
              <th>Lote Piloto</th>
            </tr>
          </thead>
          <tbody id="tbodyDados"></tbody>
          <tfoot id="tfootTotais"></tfoot>
        </table>
      </div>
    </div>

    <!-- Rodape -->
    <div class="table-footer" id="tableFooter" style="display:none">
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
        <nav><ul class="pagination pagination-sm mb-0" id="paginacao"></ul></nav>
      </div>
      <div class="table-footer-right">
        <span class="info-registros" id="infoRegistros"></span>
      </div>
    </div>

  </div>
</div>

<!-- ── Modal Busca Produto ──────────────────────────────────────────────── -->
<div class="modal fade" id="modalBuscaProduto" tabindex="-1"
     aria-labelledby="modalBuscaProdutoLabel" aria-hidden="true">
  <div class="modal-dialog modal-xl modal-dialog-scrollable">
    <div class="modal-content">

      <div class="modal-header py-2 px-3" style="background:#f5f5f5">
        <span class="modal-title fw-bold text-black" id="modalBuscaProdutoLabel"
              style="font-size:.95rem">
          Buscar Produto
        </span>
        <button type="button" class="btn-close btn-close-white btn-sm"
                data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>

      <div class="modal-body pb-0 px-3 pt-3">
        <div class="input-group input-group-sm mb-3">
          <input type="text" id="mpInputBusca" class="form-control"
                 placeholder="Código, descrição ou referência do produto…"
                 autocomplete="off" maxlength="100" />
          <button type="button" id="mpBtnBuscar" class="btn btn-primary btn-sm">
            <i class="bi bi-search me-1"></i>Buscar
          </button>
          <button type="button" class="btn btn-outline-secondary btn-sm"
                  onclick="$('#mpInputBusca').val('').focus()">
            <i class="bi bi-x-lg"></i>
          </button>
        </div>

        <div id="mpAreaResultados" style="min-height:220px;position:relative">

          <div id="mpOverlay"
               class="d-none position-absolute top-0 start-0 w-100 h-100
                      d-flex flex-column align-items-center justify-content-center"
               style="background:rgba(255,255,255,.85);z-index:5;gap:.75rem">
            <div class="spinner-border text-secondary" role="status" style="width:2rem;height:2rem"></div>
            <span style="font-size:.85rem;color:#495057;font-weight:600">Buscando…</span>
          </div>

          <div id="mpPlaceholder"
               class="d-flex flex-column align-items-center justify-content-center py-5 text-muted">
            <i class="bi bi-box-seam" style="font-size:2.5rem;opacity:.3"></i>
            <p class="mt-2 mb-0" style="font-size:.88rem">
              Digite o código ou descrição e clique em <strong>Buscar</strong>.
            </p>
          </div>

          <div id="mpWrapperTabela" class="d-none" style="max-height:50vh;overflow-y:auto">
            <table class="table table-bordered table-hover table-sm mb-0"
                   style="font-size:.8rem">
              <thead style="position:sticky;top:0;z-index:2;background:#4a4a4a;color:#fff">
                <tr>
                  <th class="text-center" style="width:36px"></th>
                  <th class="text-center">Código</th>
                  <th>Descrição</th>
                  <th>Referência</th>
                  <th>Estrutura</th>
                  <th class="text-center">Filial</th>
                  <th class="text-center">Cliente</th>
                </tr>
              </thead>
              <tbody id="mpTbody"></tbody>
            </table>
          </div>

          <div id="mpSemResultados" class="d-none py-4 text-center text-muted"
               style="font-size:.88rem">
            <i class="bi bi-inbox" style="font-size:2rem;opacity:.3;display:block;margin-bottom:.5rem"></i>
            Nenhum produto encontrado para o termo informado.
          </div>

        </div>
      </div>

      <div class="modal-footer py-2 px-3 border-top"
           style="font-size:.78rem;color:#6c757d">
        <span id="mpInfoTotal"></span>
        <button type="button" class="btn btn-sm btn-secondary ms-auto"
                data-bs-dismiss="modal">Fechar</button>
      </div>

    </div>
  </div>
</div>

<!-- Bootstrap JS Bundle -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>

<!-- jQuery -->
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<!-- jsPDF + autoTable -->
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

<script>
"use strict";

const SWNET_NIVEL   = parseInt("<%=Nivel%>", 10) || 0;
const SWNET_CODUSER = "<%=CODUSER%>";
const SWNET_USR     = "<%=USR%>";
const LOGO_PATH     = "/corporativo/comercial/Brasil/Relatorios/Logo/logo-smurfit1.png";
const REPOSITORIO   = "PedidosValidadosRepositorio.asp";
const REPO_PRODUTO  = 'PedidosValidadosRepositorio.asp';
const TOAST_DELAY   = 4500;

const CACHE_KEY_DADOS    = 'pvbr_dados';
const CACHE_KEY_FILTROS  = 'pvbr_filtros';
const CACHE_KEY_PAGINA   = 'pvbr_pagina';
const CACHE_KEY_LIMITE   = 'pvbr_limite';
const CACHE_FLAG_RETORNO = 'pvbr_retornando';

// ── Estado ────────────────────────────────────────────────────────────────────
let _dadosServidor  = [];
let _dadosFiltrados = [];
let _paginaAtual    = 1;
let _itensPorPagina = 50;

// ── Toast ─────────────────────────────────────────────────────────────────────
const TOAST_ICON = {
  sucesso: { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
  erro:    { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
  aviso:   { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
  info:    { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' },
};

function exibirToast(tipo, mensagem) {
  const t  = TOAST_ICON[tipo] || TOAST_ICON.info;
  const el = document.createElement('div');
  el.className = 'toast-custom';
  el.innerHTML = `
    <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
      <i class="bi ${t.icon}" style="color:${t.cor};font-size:1.1rem"></i>
      <span style="flex:1;font-size:.88rem;color:#212529">${mensagem}</span>
      <button onclick="this.closest('.toast-custom').remove()" style="background:none;border:none;font-size:1rem;color:#6c757d;cursor:pointer">&times;</button>
    </div>
    <div style="height:3px;background:#f1f3f5">
      <div class="toast-progress-bar" style="height:100%;background:${t.barra}"></div>
    </div>`;
  document.getElementById('toastContainer').appendChild(el);
  setTimeout(() => el.remove(), TOAST_DELAY + 300);
}

// ── Formatadores ──────────────────────────────────────────────────────────────
function formatarDataCSV(valor) {
  if (!valor) return '';
  const m = String(valor).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `\t${m[3]}/${m[2]}/${m[1]}` : String(valor);
}

function formatarNumCSV(valor, casas = 2) {
  if (valor == null || valor === '') return '';
  return parseFloat(valor).toFixed(casas).replace('.', ',');
}

function formatarDataExib(valor) {
  if (!valor) return '';
  const m = String(valor).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : String(valor);
}

function formatarNum(valor, casas = 2) {
  if (valor == null || valor === '') return '';
  return parseFloat(valor).toLocaleString('pt-BR', {
    minimumFractionDigits: casas,
    maximumFractionDigits: casas
  });
}

// ── Exportação CSV ────────────────────────────────────────────────────────────
function exportarCSV(nomeArquivo, cabecalhos, linhas) {
  const esc = v => {
    const s = v == null ? '' : String(v);
    return (s.includes(';') || s.includes('"') || s.includes('\n') || s.includes('/'))
      ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const rows = [
    cabecalhos.map(esc).join(';'),
    ...linhas.map(row => row.map(esc).join(';'))
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

function onExportCsv() {
  if (!_dadosFiltrados.length) { exibirToast('aviso', 'Nenhum dado para exportar.'); return; }

  const cab = ['Pedido','Cliente','Razão','Pedido Cliente','Produto','Referência','TES',
               'Qtde.','Qtde. Faturada','Saldo Qtde.','Área Total','Preço Inclusão',
               'Preço Atual','Preço KG NET','Entrega','Status','Emissão','Emitiu',
               'Validação','Validou','Lote Piloto'];

  const lin = _dadosFiltrados.map(d => [
    d.pedido, d.cliente, d.razao, d.pedCli, d.produto, d.referencia, d.tes,
    formatarNumCSV(d.qtde, 0), formatarNumCSV(d.qtdeFat, 0), formatarNumCSV(d.saldo, 0),
    formatarNumCSV(d.areaTotal, 0),
    formatarNumCSV(d.precoInc, 5), formatarNumCSV(d.precoAtual, 5), formatarNumCSV(d.precoKgNet, 2),
    formatarDataCSV(d.entrega), d.status,
    formatarDataCSV(d.emissao), d.emitiu,
    formatarDataCSV(d.validacao), d.validou,
    d.lotePiloto === '1' ? 'Sim' : 'Não'
  ]);

  exportarCSV('pedidos_validados', cab, lin);
}

// ── Exportação PDF ────────────────────────────────────────────────────────────
function onExportPdf() {
  if (!_dadosFiltrados.length) { exibirToast('aviso', 'Nenhum dado para exportar.'); return; }

  const { jsPDF } = window.jspdf;
  const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();

  const img = new Image();
  img.src = LOGO_PATH;
  img.onload = function () {
    const canvas = document.createElement('canvas');
    canvas.width  = 900;
    canvas.height = Math.round(900 * (img.naturalHeight / img.naturalWidth));
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    const ratio = canvas.height / canvas.width;
    doc.addImage(canvas.toDataURL('image/png'), 'PNG', 10, 4, 38, 38 * ratio);

    doc.setFontSize(13); doc.setFont(undefined, 'bold');
    doc.text('Pedidos Validados', pageW / 2, 14, { align: 'center' });

    doc.setFontSize(8); doc.setFont(undefined, 'normal');
    doc.text('Relatório gerado em ' + new Date().toLocaleString('pt-BR'), pageW / 2, 20, { align: 'center' });

    doc.setFontSize(7);
    doc.text('Emitido por: ' + SWNET_CODUSER + ' - ' + SWNET_USR, pageW - 10, 8, { align: 'right' });

    const cab = ['Pedido','Cliente','Razão','Ped.Cli','Produto','Referência','TES',
                 'Qtde','Fat.','Saldo','Área','Pr.Inc','Pr.At','KG NET',
                 'Entrega','Status','Emissão','Emitiu','Validação','Validou','LP'];

    const lin = _dadosFiltrados.map(d => [
      d.pedido, d.cliente, d.razao, d.pedCli, d.produto, d.referencia, d.tes,
      formatarNumCSV(d.qtde, 0), formatarNumCSV(d.qtdeFat, 0), formatarNumCSV(d.saldo, 0),
      formatarNumCSV(d.areaTotal, 0),
      formatarNumCSV(d.precoInc, 5), formatarNumCSV(d.precoAtual, 5), formatarNumCSV(d.precoKgNet, 2),
      formatarDataExib(d.entrega), d.status, formatarDataExib(d.emissao), d.emitiu,
      formatarDataExib(d.validacao), d.validou,
      d.lotePiloto === '1' ? 'S' : 'N'
    ]);

    doc.autoTable({
      head: [cab], body: lin, startY: 26, theme: 'grid',
      headStyles:         { fillColor: [74,74,74], textColor: [255,255,255], fontSize: 4.5, cellPadding: 1.5 },
      alternateRowStyles: { fillColor: [245,245,245] },
      bodyStyles:         { fontSize: 4.5, cellPadding: 1.5, textColor: [33,37,41] },
      tableLineColor: [180,180,180], tableLineWidth: 0.2,
    });

    doc.save('pedidos_validados.pdf');
  };
  img.onerror = function () {
    exibirToast('erro', 'Não foi possível carregar o logotipo para o PDF.');
  };
}

// ── Renderização da tabela ────────────────────────────────────────────────────
function renderizarTabela() {
  const inicio = (_paginaAtual - 1) * _itensPorPagina;
  const pagina = _dadosFiltrados.slice(inicio, inicio + _itensPorPagina);
  const $tbody = $('#tbodyDados').empty();
  const $tfoot = $('#tfootTotais').empty();

  let totalArea = 0;

  pagina.forEach(d => {
    totalArea += parseFloat(d.areaTotal) || 0;
    const lp = d.lotePiloto === '1'
      ? '<span class="text-success fw-semibold">Sim</span>'
      : '<span class="text-danger">Não</span>';

    const trClass = String(d.statusPedido) === '4' ? ' class="tr-status-4"' : '';
    $tbody.append(`<tr${trClass}>
      <td><a href="${d.linkPedido}" class="link-detalhe">${d.pedido}</a></td>
      <td>${d.cliente}</td>
      <td class="text-start">${d.razao}</td>
      <td>${d.pedCli}</td>
      <td>${d.produto}</td>
      <td class="text-start">${d.referencia}</td>
      <td>${d.tes}</td>
      <td class="text-end">${formatarNum(d.qtde, 0)}</td>
      <td class="text-end">${formatarNum(d.qtdeFat, 0)}</td>
      <td class="text-end">${formatarNum(d.saldo, 0)}</td>
      <td class="text-end">${formatarNum(d.areaTotal, 0)}</td>
      <td class="text-end">${formatarNum(d.precoInc, 5)}</td>
      <td class="text-end">${formatarNum(d.precoAtual, 5)}</td>
      <td class="text-end">${formatarNum(d.precoKgNet, 2)}</td>
      <td>${formatarDataExib(d.entrega)}</td>
      <td>${d.status}</td>
      <td>${formatarDataExib(d.emissao)}</td>
      <td>${d.emitiu}</td>
      <td>${formatarDataExib(d.validacao)}</td>
      <td>${d.validou}</td>
      <td>${lp}</td>
    </tr>`);
  });

  // Linha de totais
  $tfoot.append(`<tr class="fw-bold">
    <td colspan="10"></td>
    <td class="text-end">${formatarNum(totalArea, 0)}</td>
    <td colspan="10"></td>
  </tr>`);

  renderizarPaginacao();
}

function renderizarPaginacao() {
  const total   = _dadosFiltrados.length;
  const paginas = Math.ceil(total / _itensPorPagina);
  const $ul     = $('#paginacao').empty();

  const btn = (label, pg, disabled, active) => {
    const li = $(`<li class="page-item${disabled ? ' disabled' : ''}${active ? ' active' : ''}"></li>`);
    const a  = $(`<a class="page-link" href="#">${label}</a>`);
    if (!disabled && !active) a.on('click', e => { e.preventDefault(); _paginaAtual = pg; renderizarTabela(); });
    return li.append(a);
  };

  $ul.append(btn('&laquo;',  1,              _paginaAtual <= 1,       false));
  $ul.append(btn('&lsaquo;', _paginaAtual-1, _paginaAtual <= 1,       false));

  let start = Math.max(1, _paginaAtual - 2);
  let end   = Math.min(paginas, start + 4);
  start     = Math.max(1, end - 4);

  for (let i = start; i <= end; i++) {
    $ul.append(btn(i, i, false, i === _paginaAtual));
  }

  $ul.append(btn('&rsaquo;', _paginaAtual+1, _paginaAtual >= paginas, false));
  $ul.append(btn('&raquo;',  paginas,         _paginaAtual >= paginas, false));

  const ini = (_paginaAtual - 1) * _itensPorPagina + 1;
  const fim = Math.min(_paginaAtual * _itensPorPagina, total);
  $('#infoRegistros').text(total > 0 ? `${ini}–${fim} de ${total} registros` : '0 registros');
}

// ── Exibe resultados recebidos do servidor ────────────────────────────────────
function exibirResultados(dados) {
  _dadosServidor  = dados;
  _dadosFiltrados = [...dados];
  _paginaAtual    = 1;

  if (!dados.length) {
    $('#placeholder').show();
    $('#wrapperTabela').addClass('d-none');
    $('#tableFooter').hide();
    $('#tableToolbar').addClass('d-none');
    exibirToast('info', 'Nenhum registro encontrado para o filtro informado.');
    return;
  }

  $('#placeholder').hide();
  $('#wrapperTabela').removeClass('d-none');
  $('#tableFooter').show();
  $('#tableToolbar').removeClass('d-none');
  renderizarTabela();
}

// ── Filtro em memória (search box) ────────────────────────────────────────────
function aplicarBusca(termo) {
  const t = termo.toLowerCase();
  _dadosFiltrados = _dadosServidor.filter(d =>
    Object.values(d).some(v => String(v).toLowerCase().includes(t))
  );
  _paginaAtual = 1;
  renderizarTabela();
}

// ── Coleta os filtros do formulário ───────────────────────────────────────────
function _coletarFiltros() {
  return {
    cboFilial:    $('#cboFilial').val(),
    cboProduto:   $('#cboProduto').val(),
    codProduto:   $('#codProduto').val()  || '',
    nomeProduto:  $('#nomeProduto').val() || '',
    cboTipo:      $('#cboTipo').val(),
    dataInicial1: $('#dataInicial1').val(),
    dataFinal1:   $('#dataFinal1').val(),
    dataInicial2: $('#dataInicial2').val(),
    dataFinal2:   $('#dataFinal2').val(),
    codcli:       $('#codcliVisivel').val().trim() || $('#codcli').val() || '',
    pedido:       $('#pedido').val(),
    lotepiloto:   $('#cbxLotePiloto').is(':checked') ? '1' : '0',
  };
}

// ── Filtrar — validação e chamada AJAX ────────────────────────────────────────
function filtrar() {
  const erros = [];
  if (!$('#cboFilial').val())    erros.push('Selecione a filial.');
  if (!$('#cboTipo').val())      erros.push('Selecione o tipo.');
  if (!$('#dataInicial1').val()) erros.push('Informe a data inicial.');
  if (!$('#dataFinal1').val())   erros.push('Informe a data final.');
  if ($('#dataInicial1').val() && $('#dataFinal1').val() &&
      $('#dataInicial1').val() > $('#dataFinal1').val()) {
    erros.push('A data inicial não pode ser maior que a data final.');
  }
  if ($('#cboTipo').val() === 'EE') {
    if (!$('#dataInicial2').val()) erros.push('Informe a data de entrega inicial.');
    if (!$('#dataFinal2').val())   erros.push('Informe a data de entrega final.');
    if ($('#dataInicial2').val() && $('#dataFinal2').val() &&
        $('#dataInicial2').val() > $('#dataFinal2').val()) {
      erros.push('A data de entrega inicial não pode ser maior que a data de entrega final.');
    }
  }
  if (erros.length) { erros.forEach(e => exibirToast('aviso', e)); return; }

  $.ajax({
    url:      REPOSITORIO,
    method:   'GET',
    dataType: 'text',
    cache:    false,
    data:     _coletarFiltros(),
    beforeSend: function () {
      $('#overlayLoading').removeClass('d-none');
      $('#btnFiltrar').prop('disabled', true);
    },
    success: function (respText) {
      var resp;
      try {
        resp = JSON.parse(respText);
      } catch (e) {
        // Localiza a posição do erro no JSON
        var pos = parseInt((e.message || '').replace(/\D/g, '')) || -1;
        var ctx = pos > 0
          ? '...pos ' + pos + ': [' + respText.substring(Math.max(0, pos - 40), pos) + '>>>AQUI<<<' + respText.substring(pos, pos + 40) + ']'
          : respText.substring(0, 400);
        exibirToast('erro', 'Resposta inválida do servidor.');
        console.error('JSON.parse falhou:', e.message);
        console.error('Contexto do erro:', ctx);
        return;
      }
      if (!resp.sucesso) {
        exibirToast('erro', resp.mensagem || 'Erro ao carregar os dados.');
        return;
      }
      exibirResultados(resp.dados);
    },
    error: function (xhr) {
      exibirToast('erro', 'Erro na requisição. Tente novamente.');
      console.error('HTTP ' + xhr.status + ':', xhr.responseText ? xhr.responseText.substring(0, 300) : '(sem body)');
    },
    complete: function () {
      $('#overlayLoading').addClass('d-none');
      $('#btnFiltrar').prop('disabled', false);
    }
  });
}

function limparFiltro() {
  $('#cboFilial, #cboProduto, #cboTipo').val('');
  $('#dataInicial1, #dataFinal1, #dataInicial2, #dataFinal2').val('');
  $('#pedido').val('');
  $('#cbxLotePiloto').prop('checked', false);
  limparCamposProduto();
  limparCamposCliente();
  onChangeCboTipo('');

  _dadosServidor  = [];
  _dadosFiltrados = [];

  sessionStorage.removeItem(CACHE_KEY_DADOS);
  sessionStorage.removeItem(CACHE_KEY_FILTROS);
  sessionStorage.removeItem(CACHE_KEY_PAGINA);
  sessionStorage.removeItem(CACHE_KEY_LIMITE);

  $('#placeholder').show();
  $('#wrapperTabela').addClass('d-none');
  $('#tableFooter').hide();
  $('#tableToolbar').addClass('d-none');
  $('#searchBox').val('');
}

function limparCamposCliente() {
  $('#codcliVisivel').val('');
  $('#codcli').val('');
  $('#nomeCliente').val('');
  $('#codloja').val('');
}

// ── Muda labels e visibilidade de datas conforme tipo ─────────────────────────
function onChangeCboTipo(val) {
  const labels = {
    EM: ['Dt. Emissão Inicial',   'Dt. Emissão Final'],
    VL: ['Dt. Validação Inicial', 'Dt. Validação Final'],
    EN: ['Dt. Entrega Inicial',   'Dt. Entrega Final'],
    EE: ['Dt. Emissão Inicial',   'Dt. Emissão Final'],
    '': ['Data Inicial',           'Data Final'],
  };
  const [ini, fin] = labels[val] || labels[''];
  $('#lblDataIni1Vis').text(ini);
  $('#lblDataFin1Vis').text(fin);

  if (val === 'EE') {
    $('#divDtIniEnt, #divDtFinEnt').show();
  } else {
    $('#divDtIniEnt, #divDtFinEnt').hide();
    $('#dataInicial2, #dataFinal2').val('');
  }
}

// ── Pesquisa de produto (modal Bootstrap) ────────────────────────────────────
let _mpModal = null;
function _getModal() {
  if (!_mpModal) _mpModal = new bootstrap.Modal(document.getElementById('modalBuscaProduto'));
  return _mpModal;
}

function pesquisarProduto() {
  const filial = $('#cboFilial').val();
  if (!filial) {
    exibirToast('aviso', 'Selecione a filial antes de buscar o produto.');
    return;
  }
  $('#mpInputBusca').val('');
  $('#mpTbody').empty();
  $('#mpWrapperTabela').addClass('d-none');
  $('#mpSemResultados').addClass('d-none');
  $('#mpPlaceholder').removeClass('d-none');
  $('#mpOverlay').addClass('d-none');
  $('#mpInfoTotal').text('');

  _getModal().show();
  $('#modalBuscaProduto').one('shown.bs.modal', function () {
    $('#mpInputBusca').focus();
  });
}

function limparCamposProduto() {
  $('#codProduto').val('');
  $('#nomeProduto').val('');
}

function _mpBuscar() {
  const termo = ($('#mpInputBusca').val() || '').trim();
  if (!termo) { exibirToast('aviso', 'Informe ao menos um caractere para buscar.'); return; }

  const filial   = $('#cboFilial').val();
  const tppedido = $('#cboProduto').val() || 'CX';
  const codcli   = $('#codcli').val()  || '';
  const loja     = $('#codloja').val() || '';

  $('#mpPlaceholder').addClass('d-none');
  $('#mpSemResultados').addClass('d-none');
  $('#mpWrapperTabela').addClass('d-none');
  $('#mpOverlay').removeClass('d-none');
  $('#mpBtnBuscar').prop('disabled', true);

  $.ajax({
    url:      REPO_PRODUTO,
    method:   'GET',
    dataType: 'text',
    cache:    false,
    data: { action: 'buscaProduto', filial, parametro: termo, tppedido, codcli, loja, eshop: '0' },
    success: function (raw) {
      var resp;
      try { resp = JSON.parse(raw); }
      catch (e) {
        exibirToast('erro', 'Resposta inválida do servidor.');
        console.error('buscaPRODUTO JSON error:', e, raw.substring(0, 300));
        return;
      }
      if (!resp.sucesso) { exibirToast('erro', resp.mensagem || 'Erro ao buscar produtos.'); return; }
      _mpRenderizar(resp.dados);
    },
    error: function (xhr) {
      exibirToast('erro', 'Falha na requisição. Tente novamente.');
      console.error('buscaPRODUTO HTTP ' + xhr.status, xhr.responseText.substring(0, 200));
    },
    complete: function () {
      $('#mpOverlay').addClass('d-none');
      $('#mpBtnBuscar').prop('disabled', false);
    }
  });
}

function _mpRenderizar(dados) {
  const $tbody = $('#mpTbody').empty();
  if (!dados.length) {
    $('#mpSemResultados').removeClass('d-none');
    $('#mpWrapperTabela').addClass('d-none');
    $('#mpInfoTotal').text('');
    return;
  }
  dados.forEach(function (d) {
    const icone = d.disponivel
      ? '<i class="bi bi-check-circle-fill text-success" style="font-size:1rem;cursor:pointer" title="Selecionar produto"></i>'
      : '<i class="bi bi-exclamation-circle-fill text-warning" style="font-size:1rem" title="Produto sem estrutura"></i>';
    const $tr = $('<tr></tr>').css('cursor', d.disponivel ? 'pointer' : 'default');
    $tr.html(
      '<td class="text-center align-middle">' + icone + '</td>' +
      '<td class="text-center align-middle">' + _esc(d.codigo) + '</td>' +
      '<td class="align-middle">' + _esc(d.descricao) + '</td>' +
      '<td class="align-middle">' + _esc(d.referencia) + '</td>' +
      '<td class="align-middle">' + _esc(d.estrutura) + '</td>' +
      '<td class="text-center align-middle">' + _esc(d.filial) + '</td>' +
      '<td class="text-center align-middle">' + _esc(d.cliente) + '</td>'
    );
    if (d.disponivel) { $tr.on('click', function () { _mpSelecionar(d); }); }
    $tbody.append($tr);
  });
  $('#mpWrapperTabela').removeClass('d-none');
  $('#mpInfoTotal').text(dados.length + ' produto(s) encontrado(s)');
}

function _mpSelecionar(d) {
  $('#codProduto').val(d.codigo);
  $('#nomeProduto').val(d.descricao);
  _getModal().hide();
}

function _esc(v) {
  if (!v) return '';
  return String(v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}


// ── Pesquisa de cliente (popup) ───────────────────────────────────────────────
const _codPlantaMap = { BENTO: '05', FORTALEZA: '06', MARANGUAPE: '06', SAQUAREMA: '07' };

function pesquisarCliente() {
  const nomePlanta = $('#cboFilial option:selected').text().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const cod = Object.entries(_codPlantaMap).find(([k]) => nomePlanta.includes(k))?.[1] || '';
  const url = `${location.protocol}//${location.host}/corporativo/buscaCLIENTE.asp?id1=codcli&id2=codloja&id3=nomeCliente&id6=${cod}`;
  window.open(url, '', 'width=650,height=400');
}

// ── Inicialização ─────────────────────────────────────────────────────────────
$(function () {

  // Busca de produto — eventos do modal
  $('#mpBtnBuscar').on('click', _mpBuscar);
  $('#mpInputBusca').on('keydown', function (e) {
    if (e.key === 'Enter') { e.preventDefault(); _mpBuscar(); }
  });

  // Digitação manual no campo nomeProduto: limpa codProduto se o campo for esvaziado
  $('#nomeProduto').on('input', function () {
    if (!($(this).val() || '').trim()) {
      $('#codProduto').val('');
    }
  });

  // Input hidden de codcli (preenchido pelo popup de cliente)
  if (!$('#codcli').length) {
    $('body').append('<input type="hidden" id="codcli" value="" />');
  }

  // Sincroniza o campo visível de código com o hidden (digitação manual)
  $('#codcliVisivel').on('input', function () {
    var v = $(this).val().trim().toUpperCase();
    $(this).val(v);
    $('#codcli').val(v);
    // Limpa o nome quando o usuário edita o código manualmente
    if (!v) $('#nomeCliente').val('');
  });

  // Quando o popup de cliente preencher o hidden codcli, reflete no campo visível
  (function _watchCodCli() {
    var last = '';
    setInterval(function () {
      var cur = $('#codcli').val() || '';
      if (cur !== last) {
        last = cur;
        $('#codcliVisivel').val(cur);
      }
    }, 300);
  })();

  // Salva estado antes de navegar para o detalhamento
  $(document).on('click', '.link-detalhe', function () {
    try {
      sessionStorage.setItem(CACHE_KEY_DADOS,    JSON.stringify(_dadosServidor));
      sessionStorage.setItem(CACHE_KEY_FILTROS,  JSON.stringify(_coletarFiltros()));
      sessionStorage.setItem(CACHE_KEY_PAGINA,   String(_paginaAtual));
      sessionStorage.setItem(CACHE_KEY_LIMITE,   String(_itensPorPagina));
      sessionStorage.setItem(CACHE_FLAG_RETORNO, '1');
    } catch (e) {
      // sessionStorage cheio ou indisponível — ignora
    }
  });

  // Uppercasear pedido ao digitar
  $('#pedido').on('input', function () { $(this).val($(this).val().toUpperCase()); });

  // Tipo de filtro
  $('#cboTipo').on('change', function () { onChangeCboTipo($(this).val()); });

  // Botao filtrar
  $('#btnFiltrar').on('click', filtrar);

  // Search box com debounce
  let _searchTimer = null;
  $('#searchBox').on('input', function () {
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(() => aplicarBusca($(this).val()), 300);
  });

  // Atalho CTRL+K
  document.addEventListener('keydown', function (e) {
    if (e.ctrlKey && e.key === 'k') {
      e.preventDefault();
      e.stopPropagation();
      document.getElementById('searchBox').focus();
    }
  }, true);

  // Page size
  $('#selectLimite').on('change', function () {
    _itensPorPagina = parseInt($(this).val());
    _paginaAtual    = 1;
    renderizarTabela();
  });

  // Botoes de exportacao
  $('#btnExportCsv').on('click', onExportCsv);
  $('#btnExportPdf').on('click', onExportPdf);

  // ── Restaurar estado ao retornar do detalhamento ──────────────────────────
  (function _restaurarCache() {
    const retornando = sessionStorage.getItem(CACHE_FLAG_RETORNO) === '1';
    sessionStorage.removeItem(CACHE_FLAG_RETORNO);

    if (!retornando) {
      sessionStorage.removeItem(CACHE_KEY_DADOS);
      sessionStorage.removeItem(CACHE_KEY_FILTROS);
      sessionStorage.removeItem(CACHE_KEY_PAGINA);
      sessionStorage.removeItem(CACHE_KEY_LIMITE);
      return;
    }

    try {
      const dadosRaw   = sessionStorage.getItem(CACHE_KEY_DADOS);
      const filtrosRaw = sessionStorage.getItem(CACHE_KEY_FILTROS);
      if (!dadosRaw || !filtrosRaw) return;

      const dados   = JSON.parse(dadosRaw);
      const filtros = JSON.parse(filtrosRaw);
      const pagina  = parseInt(sessionStorage.getItem(CACHE_KEY_PAGINA) || '1', 10);
      const limite  = parseInt(sessionStorage.getItem(CACHE_KEY_LIMITE) || '50', 10);

      $('#cboFilial').val(filtros.cboFilial       || '');
      $('#cboProduto').val(filtros.cboProduto     || '');
      $('#codProduto').val(filtros.codProduto     || '');
      if (filtros.codProduto) $('#nomeProduto').val(filtros.nomeProduto || filtros.codProduto);
      $('#cboTipo').val(filtros.cboTipo           || '');
      onChangeCboTipo(filtros.cboTipo             || '');
      $('#dataInicial1').val(filtros.dataInicial1 || '');
      $('#dataFinal1').val(filtros.dataFinal1     || '');
      $('#dataInicial2').val(filtros.dataInicial2 || '');
      $('#dataFinal2').val(filtros.dataFinal2     || '');
      $('#pedido').val(filtros.pedido             || '');
      $('#cbxLotePiloto').prop('checked', filtros.lotepiloto === '1');
      if ($('#codcli').length) $('#codcli').val(filtros.codcli || '');
      if (filtros.codcli) $('#codcliVisivel').val(filtros.codcli);

      _itensPorPagina = limite;
      $('#selectLimite').val(String(limite));
      _paginaAtual = pagina;

      exibirResultados(dados);
    } catch (e) {
      // Cache corrompido — ignora e aguarda nova filtragem
    }
  })();

});
</script>

</body>
</html>
