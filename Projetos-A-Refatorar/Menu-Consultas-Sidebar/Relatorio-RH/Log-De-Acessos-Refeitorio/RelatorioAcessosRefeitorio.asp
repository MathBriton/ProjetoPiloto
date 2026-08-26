<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
Response.Buffer      = True
Response.Charset     = "UTF-8"
Response.CodePage    = 65001
Response.Expires     = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma",        "no-cache"
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
  <meta http-equiv="PRAGMA"        content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet — Acesso Refeitório</title>

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
    
    /* ── Page title ─────────────────────────────────────────────── */
    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }
    
    /* ── Filter card ─────────────────────────────────────────────── */
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
    .filter-card-body label {
      font-size: 0.82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: 0.25rem;
    }
    .input-data {
      width: 155px;
    }
    
    /* ── Table card ──────────────────────────────────────────────── */
    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
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
    .table tbody td.td-left {
      text-align: left;
    }
    .table thead th.th-left {
      text-align: left;
    }
    
    /* ── Footer paginação ────────────────────────────────────────── */
    .table-footer {
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      padding: 0.6rem 1rem;
      border-top: 1px solid #dee2e6;
      gap: 0.5rem;
    }
    .table-footer-left {
      justify-self: start;
    }
    .table-footer-center {
      justify-self: center;
    }
    .table-footer-right {
      justify-self: end;
      text-align: right;
    }
    .info-registros {
      font-size: 0.8rem;
    }
    .select-limite {
      width: auto;
    }
    
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
    
    /* ── Badges refeição ─────────────────────────────────────────── */
    .badge-refeicao {
      display: inline-block;
      font-size: 0.72rem;
      font-weight: 700;
      padding: 0.2rem 0.55rem;
      border-radius: 4px;
      white-space: nowrap;
    }
    .badge-cafe    { background-color: #fff3cd; color: #7a4f01; }
    .badge-almoco  { background-color: #d1e7dd; color: #0a4b2e; }
    .badge-jantar  { background-color: #cfe2ff; color: #08306b; }
    .badge-ceia    { background-color: #e8d5f5; color: #4a1072; }

    /* ── Placeholder ─────────────────────────────────────────────── */
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
    
    /* ── Search box ──────────────────────────────────────────────── */
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
    
    /* ── Overlay loading ─────────────────────────────────────────── */
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
    
    /* ── Toast ───────────────────────────────────────────────────── */
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

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border text-secondary"></div>
    <span>Carregando...</span>
  </div>

  <div class="toast-container-custom" id="toastMsg"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="d-flex align-items-center mb-3">
      <span class="page-title">Relatório de Acesso Refeitório</span>
    </div>

    <!-- Filtro -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <div class="row g-3 align-items-end">
          <div class="col-auto">
            <label for="inputFiltroIdentidade">Busca</label>
            <input id="inputFiltroIdentidade" type="text" class="form-control form-control-sm" style="width:260px"
                   placeholder="Credencial, Matrícula ou Nome..." />
          </div>
            <div class="col-auto">
              <label for="cboRefeicao">Refeição</label>
              <select id="cboRefeicao" class="form-select form-select-sm" style="width:170px">
                <option value="">Todas</option>
                <option value="cafe">Café da manhã</option>
                <option value="almoco">Almoço</option>
                <option value="jantar">Jantar</option>
                <option value="ceia">Ceia</option>
              </select>
            </div>
          <div class="col-auto">
            <label for="datainicial">Data Inicial</label>
            <input id="datainicial" type="date" class="form-control form-control-sm input-data" />
          </div>
          <div class="col-auto">
            <label for="datafinal">Data Final</label>
            <input id="datafinal" type="date" class="form-control form-control-sm input-data" />
          </div>
          <div class="col-auto ms-auto">
            <div class="d-flex gap-2 justify-content-end">
              <button id="btnFiltrar" type="button" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button id="btnLimpar" type="button" class="btn btn-light btn-sm" title="Limpar filtro">
                <i class="bi bi-eraser"></i>
              </button>
              <button id="btnExcel" type="button" class="btn btn-success btn-sm" title="Exportar Excel">
                <i class="bi bi-file-earmark-excel"></i>
              </button>
              <button id="btnPdf" type="button" class="btn btn-danger btn-sm" title="Exportar PDF">
                <i class="bi bi-file-earmark-pdf"></i>
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Tabela -->
    <div class="table-card">

      <div id="placeholderTabela" class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Selecione um período e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>

      <div id="areaTabela" class="d-none">
        <div class="p-2 border-bottom d-flex justify-content-end">
          <div class="search-box">
            <input id="inputBusca" type="text" class="form-control form-control-sm" placeholder="Buscar..." />
            <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
          </div>
        </div>

        <div class="table-wrapper">
          <table id="tabelaDados" class="table table-bordered table-striped table-hover table-sm mb-0">
            <thead>
              <tr>
                <th>Credencial</th>
                <th>Matrícula</th>
                <th>Nome</th>
                <th>Equipamento</th>
                <th>Descrição</th>
                <th>Data/Hora</th>
                <th>Refeição</th>
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
              <option value="1000">1000</option>
              <option value="2500">2500</option>
              <option value="5000">5000</option>
              <option value="10000">10000</option>
            </select>
          </div>
          <div class="table-footer-center">
            <nav><ul id="paginacao" class="pagination pagination-sm mb-0"></ul></nav>
          </div>
          <div class="table-footer-right">
            <span id="infoRegistros" class="info-registros text-muted"></span>
          </div>
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
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
   (() => {
  "use strict";

  // ── Constantes ────────────────────────────────────────────────────────────────

  const REPOSITORIO = "RelatorioAcessosRefeitorioRepositorio.asp";
  const CAMINHO_LOGO = "/corporativo/sesi/Logo/logo-smurfit1.png";
  const ID_USUARIO = "<%=IdUsuarioLogado%>";
  const LOGIN_USUARIO = "<%=LoginUsuarioLogado%>";
  const TOAST_DELAY = 4500;
  const AJAX_TIMEOUT = 120000;

  const COLS = Object.freeze([
    "Credencial",
    "Matrícula",
    "Nome",
    "Equipamento",
    "Descrição",
    "Data/Hora",
    "Refeição",
  ]);

  const KEYS = Object.freeze([
    "credencial",
    "matricula",
    "nome",
    "equipamento",
    "descricao",
    "datahora",
    "refeicao",
  ]);

  const KEYS_LEFT = Object.freeze(new Set([]));

  const KEYS_EXPORTAR = Object.freeze([
    "credencial",
    "matricula",
    "nome",
    "equipamento",
    "descricao",
    "datahora",
  ]);

  const REFEICAO_CFG = Object.freeze({
    cafe: { label: "Café da manhã", css: "badge-cafe" },
    almoco: { label: "Almoço", css: "badge-almoco" },
    jantar: { label: "Jantar", css: "badge-jantar" },
    ceia: { label: "Ceia", css: "badge-ceia" },
  });

  const TOAST_CFG = Object.freeze({
    sucesso: { icon: "bi-check-circle-fill", cor: "#198754", barra: "#198754" },
    erro: { icon: "bi-x-circle-fill", cor: "#dc3545", barra: "#dc3545" },
    aviso: {
      icon: "bi-exclamation-triangle-fill",
      cor: "#856404",
      barra: "#ffc107",
    },
    info: { icon: "bi-info-circle-fill", cor: "#6c757d", barra: "#6c757d" },
  });

  // ── Tipagem via JSDoc ─────────────────────────────────────────────────────────

  /**
   * @typedef {{
   *   credencial : string,
   *   matricula  : string,
   *   nome       : string,
   *   equipamento: string,
   *   descricao  : string,
   *   datahora   : string,
   *   refeicao   : string | null
   * }} RegistroAcesso
   */

  // ── Estado ────────────────────────────────────────────────────────────────────

  let _todos = /** @type {RegistroAcesso[]} */ ([]);
  let _filtrados = /** @type {RegistroAcesso[]} */ ([]);
  let _paginaAtual = 1;
  let _limite = 50;

  // ── Utilitários ───────────────────────────────────────────────────────────────

  const $el = (id) => document.getElementById(id);

  /**
   * Converte uma string de data/hora para a chave de refeição correspondente.
   * @param {string} datahora
   * @returns {"cafe" | "almoco" | "jantar" | "ceia" | null}
   */
  function classificarRefeicao(datahora) {
    const match = datahora.match(/(\d{2}):(\d{2})/);
    if (!match) return null;
    const total = parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
    const hm = (h, m) => h * 60 + m;
    if (total >= hm(4, 30) && total <= hm(9, 0)) return "cafe";
    if (total >= hm(9, 30) && total <= hm(14, 0)) return "almoco";
    if (total >= hm(17, 0) && total <= hm(21, 0)) return "jantar";
    if (total >= hm(23, 30) || total <= hm(3, 30)) return "ceia";
    return null;
  }

  /**
   * Formata uma string ISO (yyyy-mm-dd) para o formato brasileiro (dd/mm/yyyy).
   * Função pura — sem efeitos colaterais.
   * @param {string} iso
   * @returns {string}
   */
  function fmtDataBr(iso) {
    if (!iso) return "";
    const [a, m, d] = iso.split("-");
    return `${d}/${m}/${a}`;
  }

  /**
   * Enriquece um registro bruto da API adicionando a propriedade `refeicao`.
   * Retorna um novo objeto sem mutar o original.
   * @param {object} r
   * @returns {RegistroAcesso}
   */
  function enriquecerRegistro(r) {
    return {
      ...r,
      refeicao: classificarRefeicao(r.datahora ?? ""),
    };
  }

  /**
   * Constrói uma linha <tr> para a tabela a partir de um RegistroAcesso.
   * Função pura: mesmo input -> mesmo output, sem efeitos colaterais.
   * Usa textContent (não innerHTML) para prevenir XSS.
   * @param {RegistroAcesso} r
   * @returns {HTMLTableRowElement}
   */
  function criarLinha(r) {
    const KEYS_LINHA = [
      "credencial",
      "matricula",
      "nome",
      "equipamento",
      "descricao",
      "datahora",
    ];
    const tr = document.createElement("tr");

    KEYS_LINHA.forEach((k) => {
      const td = document.createElement("td");
      if (KEYS_LEFT.has(k)) td.className = "td-left";
      td.textContent = r[k] ?? "";
      tr.appendChild(td);
    });

    const tdRef = document.createElement("td");
    const cfg = r.refeicao ? REFEICAO_CFG[r.refeicao] : null;
    if (cfg) {
      const span = document.createElement("span");
      span.className = "badge-refeicao " + cfg.css;
      span.textContent = cfg.label;
      tdRef.appendChild(span);
    }
    tr.appendChild(tdRef);

    return tr;
  }

  /**
   * Monta o array de linhas para exportação (Excel ou PDF).
   * Função pura.
   * @param {RegistroAcesso[]} registros
   * @returns {string[][]}
   */
  function montarLinhasExport(registros) {
    return registros.map((r) => [
      ...KEYS_EXPORTAR.map((k) => r[k] ?? ""),
      r.refeicao && REFEICAO_CFG[r.refeicao]
        ? REFEICAO_CFG[r.refeicao].label
        : "",
    ]);
  }

  // ── ToastService ──────────────────────────────────────────────────────────────

  const ToastService = Object.freeze({
    /**
     * Exibe uma notificação toast no container fixo.
     * @param {string} mensagem
     * @param {"sucesso"|"erro"|"aviso"|"info"} tipo
     */
    exibir(mensagem, tipo) {
      const cfg = TOAST_CFG[tipo] ?? TOAST_CFG.info;
      const el = document.createElement("div");
      el.className = "toast-custom";
      el.innerHTML = `
        <div style="height:3px;background:${cfg.barra};border-radius:10px 10px 0 0">
          <div class="toast-progress-bar" style="height:100%;background:${cfg.barra}"></div>
        </div>
        <div class="d-flex align-items-start p-3 gap-2">
          <i class="bi ${cfg.icon} mt-1" style="color:${cfg.cor};font-size:1.1rem;flex-shrink:0"></i>
          <div class="flex-grow-1">
            <div style="font-size:.85rem;font-weight:700;color:${cfg.cor}">${mensagem}</div>
          </div>
          <button type="button"
                  style="background:none;border:none;font-size:1.1rem;color:#adb5bd;cursor:pointer;line-height:1"
                  onclick="this.closest('.toast-custom').remove()">&times;</button>
        </div>`;
      $el("toastMsg").appendChild(el);
      setTimeout(() => {
        if (el.parentNode) el.remove();
      }, TOAST_DELAY);
    },
  });

  // ── ExportService ─────────────────────────────────────────────────────────────

  const ExportService = Object.freeze({
    /**
     * Exporta os registros filtrados para Excel via SheetJS.
     * Inclui ao final uma seção de totalizadores por tipo de refeição.
     * @param {RegistroAcesso[]} registros
     */
    excel(registros) {
      const contagem = { cafe: 0, almoco: 0, jantar: 0, ceia: 0, sem: 0 };
      registros.forEach((r) => {
        if (r.refeicao && contagem[r.refeicao] !== undefined) {
          contagem[r.refeicao]++;
        } else {
          contagem.sem++;
        }
      });

      const totalGeral = registros.length;
      const linhaVazia = Array(COLS.length).fill("");
      const linhaTotal = Array(COLS.length).fill("");
      linhaTotal[COLS.length - 2] = "TOTAL GERAL";
      linhaTotal[COLS.length - 1] = totalGeral;

      const linhasCfg = [
        { chave: "cafe",   rotulo: "Café da manhã" },
        { chave: "almoco", rotulo: "Almoço"         },
        { chave: "jantar", rotulo: "Jantar"          },
        { chave: "ceia",   rotulo: "Ceia"            },
      ];

      const linhasRefeicao = linhasCfg.map(({ chave, rotulo }) => {
        const l = Array(COLS.length).fill("");
        l[COLS.length - 2] = rotulo;
        l[COLS.length - 1] = contagem[chave];
        return l;
      });

      const ws = XLSX.utils.aoa_to_sheet([
        COLS,
        ...montarLinhasExport(registros),
        linhaVazia,
        linhaTotal,
        ...linhasRefeicao,
      ]);

      const totalRows = 1 + registros.length;
      const sepRow   = totalRows + 1;
      const totRow   = totalRows + 2;

      const colIdx  = COLS.length - 2;
      const valIdx  = COLS.length - 1;
      const colLetra = (n) => {
        let s = "";
        n++;
        while (n > 0) {
          const r = (n - 1) % 26;
          s = String.fromCharCode(65 + r) + s;
          n = Math.floor((n - 1) / 26);
        }
        return s;
      };
      const cCol = colLetra(colIdx);
      const vCol = colLetra(valIdx);

      const estiloTotal = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4A4A4A" } },
        alignment: { horizontal: "center" },
      };
      const estilosRefeicao = {
        cafe:   { font: { bold: true, color: { rgb: "7A4F01" } }, fill: { fgColor: { rgb: "FFF3CD" } }, alignment: { horizontal: "center" } },
        almoco: { font: { bold: true, color: { rgb: "0A4B2E" } }, fill: { fgColor: { rgb: "D1E7DD" } }, alignment: { horizontal: "center" } },
        jantar: { font: { bold: true, color: { rgb: "08306B" } }, fill: { fgColor: { rgb: "CFE2FF" } }, alignment: { horizontal: "center" } },
        ceia:   { font: { bold: true, color: { rgb: "4A1072" } }, fill: { fgColor: { rgb: "E8D5F5" } }, alignment: { horizontal: "center" } },
      };

      const aplicarEstilo = (addr, estilo) => {
        if (!ws[addr]) ws[addr] = { t: "s", v: ws[addr]?.v ?? "" };
        ws[addr].s = estilo;
      };

      aplicarEstilo(`${cCol}${totRow}`, estiloTotal);
      aplicarEstilo(`${vCol}${totRow}`, estiloTotal);

      linhasCfg.forEach(({ chave }, i) => {
        const rowNum = totRow + 1 + i;
        const est = estilosRefeicao[chave];
        aplicarEstilo(`${cCol}${rowNum}`, est);
        aplicarEstilo(`${vCol}${rowNum}`, est);
      });

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Acessos Refeitório");
      XLSX.writeFile(wb, "RelatorioAcessosRefeitorio.xlsx");
    },

    /**
     * Exporta os registros filtrados para PDF via jsPDF + autoTable.
     * Carrega o logo de forma assíncrona; gera o PDF com ou sem logo.
     * @param {RegistroAcesso[]} registros
     * @param {string} dataInicial  Valor ISO do campo datainicial
     * @param {string} dataFinal    Valor ISO do campo datafinal
     * @param {() => void} onInicio Callback chamado antes de iniciar (ex: exibir overlay)
     * @param {() => void} onFim    Callback chamado ao concluir (ex: ocultar overlay)
     */
    pdf(registros, dataInicial, dataFinal, onInicio, onFim) {
      onInicio();

      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
      });
      const pageW = doc.internal.pageSize.getWidth();

      const gerarTabela = () => {
        const sub = `Período: ${fmtDataBr(dataInicial)} a ${fmtDataBr(dataFinal)}  |  ${registros.length} registros`;

        doc.setFontSize(13);
        doc.text("Relatório de Acesso Refeitório", pageW / 2, 14, {
          align: "center",
        });
        doc.setFontSize(8);
        doc.text(sub, pageW / 2, 20, { align: "center" });
        doc.setFontSize(7);
        doc.text(
          `Emitido por: ${LOGIN_USUARIO} (${ID_USUARIO})`,
          pageW - 10,
          8,
          { align: "right" },
        );

        doc.autoTable({
          head: [COLS],
          body: montarLinhasExport(registros),
          startY: 26,
          theme: "grid",
          headStyles: { fillColor: [74, 74, 74], textColor: [255, 255, 255] },
          alternateRowStyles: { fillColor: [245, 245, 245] },
          styles: { fontSize: 4.5, cellPadding: 1.5 },
          tableLineColor: [180, 180, 180],
          tableLineWidth: 0.2,
        });

        doc.save("RelatorioAcessosRefeitorio.pdf");
        onFim();
      };

      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = 900;
        canvas.height = 315;
        canvas.getContext("2d").drawImage(img, 0, 0, 900, 315);
        doc.addImage(canvas.toDataURL("image/png"), "PNG", 10, 4, 38, 13);
        gerarTabela();
      };
      img.onerror = () => gerarTabela();
      img.src = CAMINHO_LOGO;
    },
  });

  // ── TabelaService ─────────────────────────────────────────────────────────────

  const TabelaService = Object.freeze({
    /**
     * Renderiza as linhas da página atual no tbody e atualiza o contador.
     * Usa DocumentFragment para minimizar reflows.
     * @param {RegistroAcesso[]} registros
     * @param {number} pagina
     * @param {number} limite
     */
    renderizarPagina(registros, pagina, limite) {
      const inicio = (pagina - 1) * limite;
      const fim = Math.min(inicio + limite, registros.length);
      const frag = document.createDocumentFragment();

      for (let i = inicio; i < fim; i++) {
        frag.appendChild(criarLinha(registros[i]));
      }

      const tbody = $el("tbodyDados");
      tbody.innerHTML = "";
      tbody.appendChild(frag);

      $el("infoRegistros").textContent =
        `Exibindo ${inicio + 1}–${fim} de ${registros.length} registros`;

      this.renderizarPaginacao(Math.ceil(registros.length / limite), pagina);
    },

    /**
     * Renderiza os controles de paginação com janela deslizante.
     * @param {number} totalPags
     * @param {number} paginaAtual
     */
    renderizarPaginacao(totalPags, paginaAtual) {
      const JANELA = 2;
      let html = "";

      html += `<li class="page-item${paginaAtual <= 1 ? " disabled" : ""}">
                 <a class="page-link" href="#" data-pg="${paginaAtual - 1}">&laquo;</a></li>`;

      for (let p = 1; p <= totalPags; p++) {
        const visivel =
          p === 1 ||
          p === totalPags ||
          (p >= paginaAtual - JANELA && p <= paginaAtual + JANELA);
        const elipsis =
          p === paginaAtual - JANELA - 1 || p === paginaAtual + JANELA + 1;

        if (visivel) {
          html += `<li class="page-item${p === paginaAtual ? " active" : ""}">
                     <a class="page-link" href="#" data-pg="${p}">${p}</a></li>`;
        } else if (elipsis) {
          html += `<li class="page-item disabled"><span class="page-link">…</span></li>`;
        }
      }

      html += `<li class="page-item${paginaAtual >= totalPags ? " disabled" : ""}">
                 <a class="page-link" href="#" data-pg="${paginaAtual + 1}">&raquo;</a></li>`;

      $el("paginacao").innerHTML = html;
    },
  });

  // ── Módulo principal ──────────────────────────────────────────────────────────

  const RelatorioAcessosRefeitorio = Object.freeze({
    // ── Overlay ───────────────────────────────────────────────────────────────

    exibirOverlay() {
      $el("overlayLoading").classList.remove("d-none");
    },
    ocultarOverlay() {
      $el("overlayLoading").classList.add("d-none");
    },

    // ── Validação ─────────────────────────────────────────────────────────────

    /**
     * Valida os campos obrigatórios do formulário.
     * @returns {boolean}
     */
    validar() {
      const di = $el("datainicial").value;
      const df = $el("datafinal").value;
      if (!di) {
        ToastService.exibir("Informe a data inicial.", "aviso");
        return false;
      }
      if (!df) {
        ToastService.exibir("Informe a data final.", "aviso");
        return false;
      }
      if (new Date(di) > new Date(df)) {
        ToastService.exibir(
          "A data inicial deve ser menor ou igual à data final.",
          "aviso",
        );
        return false;
      }
      return true;
    },

    // ── Filtrar (fetch + async/await) ─────────────────────────────────────────

    /**
     * Busca registros no servidor e popula o estado local.
     * Usa fetch + async/await em vez de $.ajax + callbacks.
     */
    async filtrar() {
      if (!this.validar()) return;
      this.exibirOverlay();

      const di = $el("datainicial").value;
      const df = $el("datafinal").value;
      const ident = ($el("inputFiltroIdentidade").value ?? "").trim();

      try {
        const res = await fetch(REPOSITORIO, {
          method: "POST",
          body: new URLSearchParams({
            acao: "FILTRAR",
            datainicial: di,
            datafinal: df,
            identidade: ident,
          }),
          signal: AbortSignal.timeout(AJAX_TIMEOUT),
        });

        if (!res.ok) {
          ToastService.exibir(`Erro ${res.status} ao buscar dados.`, "erro");
          return;
        }

        const dados = await res.json();

        if (!Array.isArray(dados)) {
          ToastService.exibir("Resposta inválida do servidor.", "erro");
          return;
        }
        if (dados.erro) {
          ToastService.exibir(dados.erro, "erro");
          return;
        }
        if (dados.length === 0) {
          ToastService.exibir("Nenhum registro encontrado.", "aviso");
          return;
        }

        _todos = dados.map(enriquecerRegistro);
        _filtrados = _todos.slice();
        _paginaAtual = 1;
        this.exibirTabela();
      } catch (err) {
        const isTimeout =
          err.name === "TimeoutError" || err.name === "AbortError";
        ToastService.exibir(
          isTimeout
            ? "Tempo limite excedido. Reduza o período consultado."
            : "Erro ao buscar dados.",
          "erro",
        );
      } finally {
        this.ocultarOverlay();
      }
    },

    // ── Limpar ────────────────────────────────────────────────────────────────

    limpar() {
      [
        "datainicial",
        "datafinal",
        "inputFiltroIdentidade",
        "inputBusca",
        "cboRefeicao",
      ].forEach((id) => {
        $el(id).value = "";
      });
      _todos = [];
      _filtrados = [];
      $el("placeholderTabela").classList.remove("d-none");
      $el("areaTabela").classList.add("d-none");
    },

    // ── Exibir tabela ─────────────────────────────────────────────────────────

    exibirTabela() {
      $el("placeholderTabela").classList.add("d-none");
      $el("areaTabela").classList.remove("d-none");
      TabelaService.renderizarPagina(_filtrados, _paginaAtual, _limite);
    },

    // ── Busca / filtro local ──────────────────────────────────────────────────

    aplicarBusca() {
      const termo = ($el("inputBusca").value ?? "").toLowerCase().trim();
      const refFiltro = $el("cboRefeicao").value;
      const KEYS_BUSCA = [
        "credencial",
        "matricula",
        "nome",
        "equipamento",
        "descricao",
        "datahora",
      ];

      _filtrados = _todos.filter((r) => {
        const passaRefeicao = refFiltro === "" || r.refeicao === refFiltro;
        const passaBusca =
          termo === "" ||
          KEYS_BUSCA.some((k) =>
            (r[k] ?? "").toString().toLowerCase().includes(termo),
          );
        return passaRefeicao && passaBusca;
      });

      _paginaAtual = 1;
      TabelaService.renderizarPagina(_filtrados, _paginaAtual, _limite);
    },

    // ── Export ────────────────────────────────────────────────────────────────

    exportarExcel() {
      if (_filtrados.length === 0) {
        ToastService.exibir("Nenhum dado para exportar.", "aviso");
        return;
      }
      ExportService.excel(_filtrados);
    },

    exportarPdf() {
      if (_filtrados.length === 0) {
        ToastService.exibir("Nenhum dado para exportar.", "aviso");
        return;
      }
      ExportService.pdf(
        _filtrados,
        $el("datainicial").value,
        $el("datafinal").value,
        () => this.exibirOverlay(),
        () => this.ocultarOverlay(),
      );
    },

    // ── Bind de eventos ───────────────────────────────────────────────────────

    init() {
      $el("btnFiltrar").addEventListener("click", () => this.filtrar());
      $el("btnLimpar").addEventListener("click", () => this.limpar());
      $el("btnExcel").addEventListener("click", () => this.exportarExcel());
      $el("btnPdf").addEventListener("click", () => this.exportarPdf());

      $el("selectLimite").addEventListener("change", function () {
        _limite = parseInt(this.value, 10);
        _paginaAtual = 1;
        TabelaService.renderizarPagina(_filtrados, _paginaAtual, _limite);
      });

      $el("inputBusca").addEventListener("input", () => this.aplicarBusca());
      $el("cboRefeicao").addEventListener("change", () => this.aplicarBusca());

      $el("paginacao").addEventListener("click", (e) => {
        e.preventDefault();
        const a = e.target.closest("a[data-pg]");
        if (!a) return;
        const pg = parseInt(a.dataset.pg, 10);
        const totalPags = Math.ceil(_filtrados.length / _limite);
        if (pg < 1 || pg > totalPags) return;
        _paginaAtual = pg;
        TabelaService.renderizarPagina(_filtrados, _paginaAtual, _limite);
      });

      document.addEventListener("keydown", (e) => {
        if (e.ctrlKey && e.key === "k") {
          e.preventDefault();
          const busca = $el("inputBusca");
          busca.focus();
          busca.select();
        }
      });
    },
  });

  document.addEventListener("DOMContentLoaded", () =>
    RelatorioAcessosRefeitorio.init(),
  );
})();


  </script>

</body>
</html>
