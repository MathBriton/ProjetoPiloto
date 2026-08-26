<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet - Relatório de Controle de Recebimento</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
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

  .panel {
    background: #fff;
    border: 1px solid #dee2e6;
    border-radius: 6px;
  }
  .panel-header {
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
  .panel-body {
    padding: 1rem;
  }

  .tbl-wrapper {
    overflow: auto;
    min-height: 62vh;
    max-height: 62vh;
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
  .tbl-wrapper tbody tr:hover td {
    background-color: rgba(0, 0, 0, 0.03);
  }
  .tbl-vazia {
    padding: 1rem;
    text-align: center;
    color: #888;
    font-size: 0.88rem;
    height: 55vh;
    vertical-align: middle;
  }
  .tbl-empty-icon {
    font-size: 2rem;
    margin-bottom: 0.5rem;
    color: #adb5bd;
  }
  .tbl-empty-title {
    font-size: 0.95rem;
    font-weight: 600;
    color: #495057;
    margin-bottom: 0.25rem;
  }
  .tbl-empty-sub {
    font-size: 0.82rem;
    color: #6c757d;
  }

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
  .rodape-esquerda {
    display: flex;
    align-items: center;
    gap: 0.4rem;
  }
  .rodape-centro {
    display: flex;
    justify-content: center;
  }
  .rodape-direita {
    display: flex;
    justify-content: flex-end;
    font-size: 0.78rem;
    white-space: nowrap;
  }

  .paginacao .page-item .page-link {
    color: #6c757d;
  }
  .paginacao .page-item.active .page-link {
    background: #4a4a4a;
    border-color: #4a4a4a;
    color: #fff;
  }
  .paginacao .page-item.disabled .page-link {
    color: #aaa;
    pointer-events: none;
  }

  .planta-badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 3px;
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.2px;
    white-space: nowrap;
  }
  .planta-pirapetinga { background: #fff3cd; color: #856404; }
  .planta-uberaba     { background: #d1f0f5; color: #0c5460; }
  .planta-bento       { background: #e8e8e8; color: #2d3436; }
  .planta-fortaleza   { background: #fcd9d0; color: #721c24; }
  .planta-saquarema   { background: #e2deff; color: #2d1b69; }

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
  @keyframes girar {
    to { transform: rotate(360deg); }
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
    to   { opacity: 1; transform: translateX(0);    }
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
  .toast-barra {
    height: 3px;
    animation: encolher linear forwards;
  }
  .toast-barra-success { background: #198754; }
  .toast-barra-warning { background: #ffc107; }
  .toast-barra-danger  { background: #dc3545; }
  .toast-barra-info    { background: #6c757d; }
  @keyframes encolher {
    from { width: 100%; }
    to   { width: 0%;   }
  }
  </style>
</head>
<body>

<div class="toast-wrapper" id="toastWrapper"></div>
<div id="loader" class="overlay-loader d-none">
  <div class="spinner-circular"></div>
</div>

<div class="container-fluid py-3 pb-5">

  <div class="d-flex align-items-center justify-content-between mb-3">
    <h4 class="fw-bold mb-0 page-title">Relatório de Controle de Recebimento</h4>
    <a href="CrudControleRecebimento.asp" class="btn btn-outline-secondary btn-sm">
      <i class="bi bi-pencil-square me-1"></i>Cadastro
    </a>
  </div>

  <div class="panel mb-3">
    <div class="panel-header">
      <i class="bi bi-funnel"></i> Filtros
    </div>
    <div class="panel-body">
      <div class="row g-2 align-items-end">
        <div class="col-lg col-md-4 col-sm-6">
          <label class="form-label" for="filtroPlanta">Planta</label>
          <select id="filtroPlanta" class="form-select form-select-sm">
            <option value="">Todos</option>
            <option value="00">Pirapetinga</option>
            <option value="04">Uberaba</option>
            <option value="05">Bento Gonçalves</option>
            <option value="06">Fortaleza</option>
            <option value="07">Saquarema</option>
          </select>
        </div>
        <div class="col-lg col-md-4 col-sm-6">
          <label class="form-label" for="filtroFornecedor">Fornecedor</label>
          <select id="filtroFornecedor" class="form-select form-select-sm">
            <option value="">Selecione...</option>
          </select>
        </div>
        <div class="col-lg-auto col-md-3 col-sm-6">
          <label class="form-label" for="filtroDataIni">De</label>
          <input id="filtroDataIni" type="date" class="form-control form-control-sm" />
        </div>
        <div class="col-lg-auto col-md-3 col-sm-6">
          <label class="form-label" for="filtroDataFim">Até</label>
          <input id="filtroDataFim" type="date" class="form-control form-control-sm" />
        </div>
        <div class="col-lg-auto d-flex align-items-end gap-1">
          <button type="button" class="btn btn-sm btn-primary" id="btnFiltrar">
            <i class="bi bi-funnel-fill me-1"></i>Filtrar
          </button>
          <button type="button" class="btn btn-sm btn-light border flex-shrink-0" id="btnLimpar" title="Limpar filtros">
            <i class="bi bi-eraser"></i>
          </button>
          <button type="button" class="btn btn-sm btn-success" id="btnExcel" title="Exportar para Excel">
            <i class="bi bi-file-earmark-excel"></i>
          </button>
          <button type="button" class="btn btn-sm btn-danger" id="btnPdf" title="Exportar para PDF">
            <i class="bi bi-file-earmark-pdf"></i>
          </button>
        </div>
      </div>
    </div>
  </div>

  <div class="panel">
    <div class="panel-header">
      <i class="bi bi-table"></i> Registros
    </div>
    <div class="panel-body p-0">
      <div class="tbl-wrapper">
        <table class="table table-bordered table-striped mb-0" id="tblDados">
          <thead>
            <tr>
              <th style="min-width:110px">Planta</th>
              <th style="min-width:80px">N.F</th>
              <th>Série</th>
              <th style="min-width:95px">Data N.F</th>
              <th style="min-width:95px">Data Teste</th>
              <th style="min-width:65px">Hora</th>
              <th style="min-width:130px">CNPJ</th>
              <th style="min-width:170px">Fornecedor</th>
              <th style="min-width:150px">Produto</th>
              <th>Peso N.F</th>
              <th>Peso Bal.</th>
              <th>% Umid.</th>
              <th>% Impurez.</th>
              <th>% Proib.</th>
              <th>% Outros</th>
              <th>Local</th>
              <th>Qtd.</th>
              <th>Fardos An.</th>
              <th>Placa</th>
              <th style="min-width:130px">Motorista</th>
              <th style="min-width:160px">Obs</th>
              <th style="min-width:110px">Incluído por</th>
            </tr>
          </thead>
          <tbody id="tblBody">
            <tr id="trAguardando">
              <td colspan="22" class="tbl-vazia">
                <div class="tbl-empty-icon"><i class="bi bi-funnel"></i></div>
                <div class="tbl-empty-title">Nenhum resultado exibido</div>
                <div class="tbl-empty-sub">Aplique os filtros acima para obter os resultados.</div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="rodape-paginacao d-none" id="rodapeTabela">
      <div class="rodape-esquerda">
        <label class="form-label mb-0" for="pgTamanho" style="font-size:0.78rem">Exibir</label>
        <select id="pgTamanho" class="form-select form-select-sm" style="width:auto">
          <option value="25">25</option>
          <option value="50" selected>50</option>
          <option value="100">100</option>
          <option value="200">200</option>
        </select>
      </div>
      <div class="rodape-centro">
        <nav aria-label="Paginação">
          <ul class="pagination pagination-sm mb-0 paginacao" id="paginacao"></ul>
        </nav>
      </div>
      <div class="rodape-direita" id="contador">
        <i class="bi bi-table me-1"></i>
        <b id="cntTotal">0</b>&nbsp;registro(s) &mdash; pág.&nbsp;<b id="pgAtual">1</b>&nbsp;de&nbsp;<b id="pgTotal">1</b>
      </div>
    </div>
  </div>

</div>

<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>

<script>
  var LoginUsuarioLogado = "<%= Session("USR") %>";

  (($) => {
    "use strict";

    const REPO     = "RelatorioControleRecebimentoRepositorio.asp";
    const LOGO_URL = "Logo/logo-smurfit1.png";

    const TOAST_CFG  = Object.freeze({ delay: 4500, borda: true });
    const TOAST_ICON = Object.freeze({
      success: { cls: "bi-check-circle-fill",         cor: "#198754" },
      warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
      danger:  { cls: "bi-x-circle-fill",             cor: "#dc3545" },
      info:    { cls: "bi-info-circle-fill",           cor: "#6c757d" },
    });

    const PLANTA_MAP = Object.freeze({
      "00": { nome: "PIRAPETINGA", cls: "planta-pirapetinga" },
      "04": { nome: "UBERABA",     cls: "planta-uberaba"     },
      "05": { nome: "BENTO",       cls: "planta-bento"       },
      "06": { nome: "FORTALEZA",   cls: "planta-fortaleza"   },
      "07": { nome: "SAQUAREMA",   cls: "planta-saquarema"   },
    });

    const COLUNAS_EXPORT = [
      "Planta","N.F.","Série","Data N.F.","Data Teste","Hora","CNPJ","Fornecedor",
      "Produto","Peso N.F.","Peso Bal.","% Umid.","% Impurez.","% Proib.",
      "% Outros","Local","Qtd.","Fardos An.","Placa","Motorista","Obs","Incluído por"
    ];

    let _pgTamanho = 50;
    let _pgAtual   = 1;
    let _dados     = [];
    let _logoB64   = null;

    const Html = Object.freeze({
      esc(txt) {
        if (!txt) return "";
        return String(txt)
          .replace(/&/g, "&amp;").replace(/</g, "&lt;")
          .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
      },
    });

    const Formato = Object.freeze({
      numero(v, dec) {
        const n = parseFloat(String(v ?? "").replace(",", "."));
        if (isNaN(n)) return "-";
        return n.toLocaleString("pt-BR", { minimumFractionDigits: dec, maximumFractionDigits: dec });
      },
      nomeArquivo(ext) {
        const d  = new Date();
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `ControleRecebimento_${d.getFullYear()}${mm}${dd}.${ext}`;
      },
    });

    const UI = Object.freeze({
      exibirLoader()  { $("#loader").removeClass("d-none"); },
      ocultarLoader() { $("#loader").addClass("d-none"); },
      notificar(mensagem, tipo = "info") {
        const cfg  = TOAST_ICON[tipo] ?? TOAST_ICON.info;
        const item = $(`
          <div class="toast-item toast-item-${tipo}">
            <div class="toast-corpo">
              <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
              <span class="flex-grow-1">${mensagem}</span>
              <button class="toast-fechar" aria-label="Fechar">&times;</button>
            </div>
            <div class="toast-barra toast-barra-${tipo}" style="animation-duration:${TOAST_CFG.delay}ms;"></div>
          </div>
        `);
        item.find(".toast-fechar").on("click", () => item.remove());
        $("#toastWrapper").append(item);
        setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_CFG.delay);
      },
    });

    const Filtro = Object.freeze({
      valores() {
        return {
          planta:      $("#filtroPlanta").val(),
          fornecedor:  $("#filtroFornecedor").val(),
          datainicial: $("#filtroDataIni").val(),
          datafinal:   $("#filtroDataFim").val(),
        };
      },
      validar() {
        const v = Filtro.valores();
        if (!v.datainicial) { UI.notificar("Informe a data inicial.", "warning"); return false; }
        if (!v.datafinal)   { UI.notificar("Informe a data final.",   "warning"); return false; }
        if (new Date(v.datainicial) > new Date(v.datafinal)) {
          UI.notificar("A data inicial deve ser menor que a data final.", "warning");
          return false;
        }
        return true;
      },
      limpar() {
        $("#filtroPlanta, #filtroFornecedor").val("");
        $("#filtroDataIni, #filtroDataFim").val("");
        _dados = [];
        Tabela.renderizarAguardando();
        $("#rodapeTabela").addClass("d-none");
      },
    });

    const Paginacao = Object.freeze({
      totalPaginas() { return Math.max(1, Math.ceil(_dados.length / _pgTamanho)); },
      ir(pg) {
        _pgAtual = Math.max(1, Math.min(pg, Paginacao.totalPaginas()));
        Tabela.renderizarPagina();
        Paginacao.renderizarControles();
      },
      renderizarControles() {
        const total = Paginacao.totalPaginas();
        const $pg   = $("#paginacao").empty();

        const criarItem = (label, pg, desabilitado, ativo) => {
          const $li = $(`<li class="page-item${desabilitado ? " disabled" : ""}${ativo ? " active" : ""}"></li>`);
          const $a  = $(`<a class="page-link" href="#">${label}</a>`);
          if (!desabilitado && !ativo) $a.on("click", (e) => { e.preventDefault(); Paginacao.ir(pg); });
          $li.append($a);
          return $li;
        };

        $pg.append(criarItem("&laquo;", _pgAtual - 1, _pgAtual === 1, false));

        const delta = 2;
        let ini = Math.max(1, _pgAtual - delta);
        let fim = Math.min(total, _pgAtual + delta);
        if (ini > 1) {
          $pg.append(criarItem(1, 1, false, false));
          if (ini > 2) $pg.append($('<li class="page-item disabled"><span class="page-link">…</span></li>'));
        }
        for (let p = ini; p <= fim; p++) $pg.append(criarItem(p, p, false, p === _pgAtual));
        if (fim < total) {
          if (fim < total - 1) $pg.append($('<li class="page-item disabled"><span class="page-link">…</span></li>'));
          $pg.append(criarItem(total, total, false, false));
        }
        $pg.append(criarItem("&raquo;", _pgAtual + 1, _pgAtual === total, false));

        $("#cntTotal").text(_dados.length);
        $("#pgAtual").text(_pgAtual);
        $("#pgTotal").text(total);
        $("#rodapeTabela").removeClass("d-none");
      },
    });

    const Tabela = Object.freeze({
      renderizarAguardando() {
        $("#tblBody").html(`
          <tr id="trAguardando">
            <td colspan="22" class="tbl-vazia">
              <div class="tbl-empty-icon"><i class="bi bi-funnel"></i></div>
              <div class="tbl-empty-title">Nenhum resultado exibido</div>
              <div class="tbl-empty-sub">Aplique os filtros acima para obter os resultados.</div>
            </td>
          </tr>
        `);
      },
      renderizarVazia() {
        $("#tblBody").html(`
          <tr>
            <td colspan="22" class="tbl-vazia">
              <div class="tbl-empty-icon"><i class="bi bi-search"></i></div>
              <div class="tbl-empty-title">Nenhum registro encontrado</div>
              <div class="tbl-empty-sub">Tente ajustar os filtros e pesquisar novamente.</div>
            </td>
          </tr>
        `);
      },
      renderizarLinha(d) {
        const p   = PLANTA_MAP[d.filial] ?? { nome: d.filial, cls: "" };
        const num = (v, dec) => Formato.numero(v, dec);
        return `<tr>
          <td><span class="planta-badge ${p.cls}">${p.nome}</span></td>
          <td>${Html.esc(d.nf)}</td>
          <td>${Html.esc(d.serie)}</td>
          <td>${Html.esc(d.dataNf)}</td>
          <td>${Html.esc(d.dataRecebe)}</td>
          <td>${Html.esc(d.horaRecebe)}</td>
          <td>${Html.esc(d.cnpj)}</td>
          <td style="text-align:left">${Html.esc(d.fornecedor)}</td>
          <td style="text-align:left">${Html.esc(d.produto)}</td>
          <td style="text-align:right">${num(d.pesoNf,2)}</td>
          <td style="text-align:right">${num(d.pesoBalanca,2)}</td>
          <td style="text-align:right">${num(d.percUmidade,2)}</td>
          <td style="text-align:right">${num(d.percImpureza,2)}</td>
          <td style="text-align:right">${num(d.percProibitivo,2)}</td>
          <td style="text-align:right">${num(d.percOutros,2)}</td>
          <td>${Html.esc(d.local)}</td>
          <td style="text-align:right">${Html.esc(d.numFardos)}</td>
          <td style="text-align:right">${Html.esc(d.fardosAnalizados)}</td>
          <td>${Html.esc(d.placa)}</td>
          <td style="text-align:left">${Html.esc(d.motorista)}</td>
          <td style="text-align:left">${Html.esc(d.obs)}</td>
          <td>${Html.esc(d.quemIncluiu)}</td>
        </tr>`;
      },
      renderizarPagina() {
        const inicio = (_pgAtual - 1) * _pgTamanho;
        const fatia  = _dados.slice(inicio, inicio + _pgTamanho);
        $("#tblBody").html(fatia.map(Tabela.renderizarLinha).join(""));
      },
    });

    const Exportar = Object.freeze({
      excel() {
          if (!_dados.length) { UI.notificar("Não há dados para exportar.", "warning"); return; }
          const toNum = (v) => { const n = parseFloat(String(v ?? "").replace(",", ".")); return isNaN(n) ? v : n; };
          const linhas = _dados.map(d => {
            const p = PLANTA_MAP[d.filial] ?? { nome: d.filial };
            return [
              p.nome, d.nf, d.serie, d.dataNf, d.dataRecebe, d.horaRecebe,
              d.cnpj, d.fornecedor, d.produto,
              toNum(d.pesoNf), toNum(d.pesoBalanca), toNum(d.percUmidade), toNum(d.percImpureza),
              toNum(d.percProibitivo), toNum(d.percOutros), d.local, toNum(d.numFardos),
              toNum(d.fardosAnalizados), d.placa, d.motorista, d.obs, d.quemIncluiu,
            ];
          });
        const ws  = XLSX.utils.aoa_to_sheet([COLUNAS_EXPORT, ...linhas]);
        const wb  = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Dados");
        XLSX.writeFile(wb, Formato.nomeArquivo("xlsx"));
      },
      pdf() {
        if (!_dados.length) { UI.notificar("Não há dados para exportar.", "warning"); return; }
        UI.exibirLoader();
        const gerar = () => {
          const { jsPDF } = window.jspdf;
          const doc       = new jsPDF({ orientation: "landscape", format: "a3" });
          const PAG_L     = doc.internal.pageSize.getWidth();
          const PAG_H     = doc.internal.pageSize.getHeight();
          const MARGEM_L  = 10;
          const MARGEM_TB = 4;
          const HDR_H     = 26;
          const v         = Filtro.valores();
          const fmtDt     = (s) => s ? s.split("-").reverse().join("/") : "";
          const subtitulo = `Período: ${fmtDt(v.datainicial)} a ${fmtDt(v.datafinal)} — ${_dados.length} registro(s)`;

          const _cabecalho = (data) => {
            doc.setFillColor(255, 255, 255);
            doc.rect(0, 0, PAG_L, HDR_H, "F");
            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.4);
            doc.line(0, HDR_H, PAG_L, HDR_H);
            if (_logoB64) {
              try { doc.addImage(_logoB64, "PNG", MARGEM_L, 5, 38, 13); } catch (_) {}
            }
            doc.setFontSize(13);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(74, 74, 74);
            doc.text("Relatório de Controle de Recebimento", PAG_L / 2, 11, { align: "center" });
            doc.setFontSize(7.5);
            doc.setFont("helvetica", "normal");
            doc.text(subtitulo, PAG_L / 2, 18, { align: "center" });
            doc.setFontSize(7);
            doc.setTextColor(120);
            const emitido = `Emitido por: ${LoginUsuarioLogado || "—"}  |  ${new Date().toLocaleString("pt-BR")}`;
            doc.text(emitido, PAG_L - MARGEM_L, 23, { align: "right" });
            doc.setTextColor(33, 37, 41);
          };

          const head = [COLUNAS_EXPORT];
          const body = _dados.map(d => {
            const p = PLANTA_MAP[d.filial] ?? { nome: d.filial };
            return [
              p.nome, d.nf, d.serie, d.dataNf, d.dataRecebe, d.horaRecebe,
              d.cnpj, d.fornecedor, d.produto,
              Formato.numero(d.pesoNf, 2), Formato.numero(d.pesoBalanca, 2),
              Formato.numero(d.percUmidade, 2), Formato.numero(d.percImpureza, 2),
              Formato.numero(d.percProibitivo, 2), Formato.numero(d.percOutros, 2),
              d.local, d.numFardos, d.fardosAnalizados,
              d.placa, d.motorista, d.obs, d.quemIncluiu,
            ];
          });

          const pesos      = [12,8,6,10,10,7,13,18,16,9,9,8,9,8,8,8,6,8,8,12,16,11];
          const larguraTbl = PAG_L - MARGEM_TB * 2;
          const soma       = pesos.reduce((a, b) => a + b, 0);
          const fator      = larguraTbl / soma;
          const columnStyles = {};
          pesos.forEach((w, i) => {
            columnStyles[i] = { cellWidth: parseFloat((w * fator).toFixed(2)), halign: "center" };
          });

          doc.autoTable({
            head, body,
            startY: 29,
            tableWidth: larguraTbl,
            styles: {
              fontSize: 5.8, cellPadding: 1.4, valign: "middle", overflow: "linebreak",
              textColor: [33, 37, 41], lineColor: [220, 220, 220],
              lineWidth: { top: 0, bottom: 0, left: 0.1, right: 0.1 },
            },
            headStyles: {
              fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: "bold",
              halign: "center", fontSize: 5.8,
              lineColor: [100, 100, 100], lineWidth: { top: 0, bottom: 0, left: 0.2, right: 0.2 },
            },
            alternateRowStyles: { fillColor: [248, 248, 248] },
            columnStyles,
            margin: { left: MARGEM_TB, right: MARGEM_TB },
            didDrawPage: _cabecalho,
          });

          const totalPags = doc.getNumberOfPages();
          for (let p = 1; p <= totalPags; p++) {
            doc.setPage(p);
            doc.setFontSize(7);
            doc.setTextColor(150);
            doc.text(`Página ${p} de ${totalPags}`, PAG_L - MARGEM_L, PAG_H - 5, { align: "right" });
          }

          doc.save(Formato.nomeArquivo("pdf"));
          UI.ocultarLoader();
        };

        if (_logoB64) { gerar(); return; }
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
          gerar();
        };
        img.onerror = () => gerar();
        img.src = `${LOGO_URL}?v=${Date.now()}`;
      },
    });

    const Consulta = Object.freeze({
      carregarFornecedores() {
        $.ajax({
          url: REPO,
          type: "GET",
          data: { acao: "fornecedores" },
          dataType: "json",
          success(resp) {
            if (!resp.sucesso) return;
            const $sel = $("#filtroFornecedor").empty().append('<option value="">Selecione...</option>');
            resp.dados.forEach(f => {
              $sel.append(`<option value="${Html.esc(f.cnpj)}">${Html.esc(f.nome)}</option>`);
            });
          },
        });
      },
      executar() {
        if (!Filtro.validar()) return;
        UI.exibirLoader();
        $.ajax({
          url: REPO,
          type: "GET",
          data: Object.assign({ acao: "listar" }, Filtro.valores()),
          dataType: "json",
          success(resp) {
            UI.ocultarLoader();
            if (!resp.sucesso) { UI.notificar(resp.mensagem || "Erro ao consultar.", "danger"); return; }
            _dados   = resp.dados ?? [];
            _pgAtual = 1;
            if (!_dados.length) {
              Tabela.renderizarVazia();
              $("#rodapeTabela").addClass("d-none");
              UI.notificar("Nenhum registro encontrado para os filtros informados.", "info");
              return;
            }
            Tabela.renderizarPagina();
            Paginacao.renderizarControles();
          },
          error() {
            UI.ocultarLoader();
            UI.notificar("Erro ao processar requisição.", "danger");
          },
        });
      },
    });

    function anexarEventos() {
      $("#btnFiltrar").on("click", () => Consulta.executar());
      $("#btnLimpar").on("click",  () => Filtro.limpar());
      $("#btnExcel").on("click",   () => Exportar.excel());
      $("#btnPdf").on("click",     () => Exportar.pdf());
      $("#pgTamanho").on("change", function () {
        _pgTamanho = parseInt($(this).val(), 10);
        _pgAtual   = 1;
        Tabela.renderizarPagina();
        Paginacao.renderizarControles();
      });
      $("#filtroDataIni, #filtroDataFim").on("keydown", function (e) {
        if (e.which === 13) Consulta.executar();
      });
    }

    $(function () {
      UI.ocultarLoader();
      const hoje    = new Date();
      const hojeFmt = `${hoje.getFullYear()}-${String(hoje.getMonth()+1).padStart(2,"0")}-${String(hoje.getDate()).padStart(2,"0")}`;
      $("#filtroDataFim").val(hojeFmt);
      Consulta.carregarFornecedores();
      anexarEventos();
    });

  })(jQuery);
</script>

</body>
</html>
