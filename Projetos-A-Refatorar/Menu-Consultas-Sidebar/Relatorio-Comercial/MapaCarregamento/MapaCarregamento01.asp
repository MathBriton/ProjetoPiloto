<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Mapa de Carregamento</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />

    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <!-- jsPDF + autoTable (geração de PDF nativa) -->
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

    <script src="../../../funcoes/funcoesJAVASCRIPT.js"></script>

    <style>
      /* ─── Tokens de Design ─── */
      :root {
        --color-bg:            hsl(240, 9%, 98%);
        --color-surface:       hsl(0, 0%, 100%);
        --color-surface-muted: hsl(240, 9%, 97%);
        --color-border:        hsl(240, 6%, 90%);
        --color-border-strong: hsl(240, 6%, 82%);

        --color-text:          hsl(240, 6%, 10%);
        --color-text-muted:    hsl(240, 4%, 46%);
        --color-text-faint:    hsl(240, 4%, 64%);

        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        --color-success:       hsl(145, 63%, 32%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);
        --color-warning:       hsl(35, 100%, 48%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
        --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

        --radius-sm: 6px;
        --radius-md: 10px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-3: 12px;
        --space-4: 16px;

        --z-toast: 1100;
      }

      /* ─── Base ─── */
      body {
        background-color: var(--color-bg);
        font-family: var(--font-sans);
        color: var(--color-text);
        min-height: 100vh;
      }

      /* ─── Cabeçalho da página ─── */
      .page-header {
        padding: var(--space-4) 0 var(--space-3);
      }

      .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
      }

      /* ─── Card de Filtros ─── */
      .filter-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
      }

      .filter-card-header {
        background-color: var(--color-surface);
        color: var(--color-text);
        padding: .7rem 1.1rem;
        border-radius: var(--radius-md) var(--radius-md) 0 0;
        border-bottom: 1px solid var(--color-border);
        font-weight: 600;
        font-size: .85rem;
      }

      .filter-card-body {
        padding: 1.1rem 1.25rem 1.35rem;
      }

      .filter-card-body .form-label {
        font-size: .78rem;
        font-weight: 500;
        color: var(--color-text-muted);
        letter-spacing: .01em;
        margin-bottom: .3rem;
      }

      .form-control:focus,
      .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px var(--color-accent-soft);
        outline: none;
      }

      /* ─── Botões ─── */
      .btn-primary {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
        color: hsl(0, 0%, 100%);
        font-weight: 500;
        font-size: .85rem;
        border-radius: var(--radius-sm);
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-primary:hover,
      .btn-primary:focus {
        background-color: var(--color-accent-hover);
        border-color: var(--color-accent-hover);
        color: hsl(0, 0%, 100%);
      }
      .btn-primary:active  { transform: scale(.98); }
      .btn-primary:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }
      .btn-primary:disabled,
      .btn-primary.disabled { opacity: .5; cursor: not-allowed; pointer-events: none; }

      .btn-light {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        color: var(--color-text);
        font-weight: 500;
        font-size: .85rem;
        border-radius: var(--radius-sm);
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-light:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
        color: var(--color-text);
      }
      .btn-light:active { transform: scale(.98); }
      .btn-light:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }
      .btn-light:disabled,
      .btn-light.disabled { opacity: .5; cursor: not-allowed; pointer-events: none; }

      /* ─── Card de Tabela ─── */
      .table-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        position: relative;
      }

      /* ─── Toolbar do topo da tabela ─── */
      .table-toolbar {
        background-color: var(--color-surface);
        border-radius: var(--radius-md) var(--radius-md) 0 0;
      }

      /* ─── Wrapper e tabela ─── */
      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }
      .table-wrapper table { min-width: max-content; }

      .table thead th {
        position: sticky;
        top: 0;
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        font-size: .76rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .03em;
        text-align: center;
        white-space: nowrap;
        z-index: 2;
        border-color: var(--color-border);
        padding: .6rem .75rem;
      }

      .table tbody td {
        font-size: .85rem;
        text-align: center;
        white-space: nowrap;
        color: var(--color-text);
        vertical-align: middle;
        padding: .55rem .75rem;
        border-color: var(--color-border);
      }

      .table tbody tr:not([class*="row-"]):hover td {
        background-color: var(--color-surface-muted);
      }

      /* ─── Linha de placa (cabeçalho de grupo) ─── */
      .row-placa td {
        background-color: var(--color-surface-muted);
        font-weight: 600;
        font-size: .82rem;
        color: var(--color-text);
        border-top: 2px solid var(--color-border-strong) !important;
        text-align: left !important;
        padding-left: .75rem !important;
      }

      /* ─── Linha de subtotal ─── */
      .row-subtotal td {
        font-weight: 600;
        font-size: .82rem;
        background-color: hsl(240, 6%, 94%);
        color: var(--color-text);
        border-top: 1px solid var(--color-border-strong) !important;
      }

      /* ─── Espaçador entre grupos ─── */
      .row-spacer td {
        height: 6px;
        padding: 0 !important;
        border: none !important;
        background-color: var(--color-bg);
      }

      /* ─── Tabela de resumo ─── */
      .summary-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        overflow: hidden;
      }

      .summary-card .summary-header {
        background-color: var(--color-surface);
        color: var(--color-text);
        padding: .7rem 1.1rem;
        border-bottom: 1px solid var(--color-border);
        font-weight: 600;
        font-size: .85rem;
      }

      .summary-table thead th {
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        font-size: .72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .03em;
        text-align: center;
        border-color: var(--color-border);
        padding: .5rem .6rem;
        white-space: nowrap;
      }

      .summary-table .th-group-peso {
        background-color: hsl(240, 6%, 93%);
        color: var(--color-text);
      }

      .summary-table .th-group-area {
        background-color: hsl(210, 40%, 92%);
        color: var(--color-text);
      }

      .summary-table tbody td {
        font-size: .85rem;
        text-align: center;
        color: var(--color-text);
        border-color: var(--color-border);
        padding: .55rem .6rem;
      }

      /* ─── Overlay de carregamento ─── */
      .overlay-loading {
        position: absolute;
        inset: 0;
        background: hsla(0, 0%, 100%, .82);
        backdrop-filter: blur(2px);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 10;
        gap: .85rem;
        border-radius: var(--radius-md);
      }
      .overlay-loading .spinner-border {
        width: 1.75rem;
        height: 1.75rem;
        border-width: 2px;
        color: var(--color-text-faint);
      }
      .overlay-loading span {
        font-size: .85rem;
        font-weight: 500;
        color: var(--color-text-muted);
      }

      /* ─── Toast personalizado ─── */
      .toast-container-custom {
        position: fixed;
        top: 1rem;
        right: 1rem;
        z-index: var(--z-toast);
        display: flex;
        flex-direction: column;
        gap: .5rem;
      }

      .toast-custom {
        min-width: 280px;
        max-width: 360px;
        border-radius: var(--radius-md);
        overflow: hidden;
        box-shadow: var(--shadow-md);
        background: var(--color-surface);
        border: 1px solid var(--color-border);
        animation: toastSlideIn .22s ease;
      }

      @keyframes toastSlideIn {
        from { opacity: 0; transform: translateX(24px); }
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

      /* ─── Botões de exportação (toolbar) ─── */
      .btn-export-pdf {
        background-color: var(--color-danger);
        border-color: var(--color-danger);
        color: hsl(0, 0%, 100%);
        font-weight: 500;
        font-size: .8rem;
        border-radius: var(--radius-sm);
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-export-pdf:hover { background-color: hsl(3, 90%, 50%); border-color: hsl(3, 90%, 50%); color: hsl(0, 0%, 100%); }
      .btn-export-pdf:focus-visible { outline: none; box-shadow: 0 0 0 3px var(--color-danger-soft); }
      .btn-export-pdf:active { transform: scale(.98); }

      /* ─── Search box ─── */
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
        padding: .1rem .35rem;
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        color: var(--color-text-muted);
        font-family: var(--font-mono);
      }

      /* ─── Acessibilidade ─── */
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
          animation-duration: 0.01ms !important;
          transition-duration: 0.01ms !important;
        }
      }

      /* ─── Impressão ─── */
      @media print {
        @page { size: landscape; }
        body  { margin: 0; background: white; }
        .no-print { display: none !important; }
        .table-card,
        .filter-card,
        .summary-card { box-shadow: none; border: 1px solid hsl(0,0%,70%); }
      }

      /* ─── Responsividade ─── */
      @media (max-width: 768px) {
        .table-toolbar {
          flex-direction: column;
          align-items: stretch !important;
          gap: .5rem;
        }
        .table-toolbar > div { justify-content: center; }
        .filter-card-body .row { row-gap: .75rem; }
      }
    </style>
  </head>

  <body>

    <div class="container-fluid py-3 px-3 px-lg-4">

      <!-- ── Cabeçalho da página ── -->
      <div class="page-header mb-3">
        <span class="page-title">Mapa de Carregamento</span>
      </div>

      <!-- ── Card de Filtros ── -->
      <div class="filter-card mb-3">
        <div class="filter-card-header">
          <i class="bi bi-funnel me-2" style="color:var(--color-text-muted);" aria-hidden="true"></i>Filtros
        </div>
        <div class="filter-card-body">
          <div class="row g-3 align-items-end">
            <div class="col-12 col-md-3">
              <label for="DataIn" class="form-label">Data de referência</label>
              <input id="DataIn" type="date" class="form-control form-control-sm" required />
            </div>

            <div class="col-12 col-md-3">
              <label for="cboFilial" class="form-label">Filial</label>
              <select id="cboFilial" class="form-select form-select-sm">
                <option value="00">PIRAPETINGA</option>
                <option value="04">UBERABA</option>
              </select>
            </div>

            <div class="col-12 col-md ms-auto">
              <div class="d-flex gap-2 justify-content-end flex-wrap">
                <button type="button" id="btnClear" class="btn btn-light btn-sm">
                  <i class="bi bi-eraser me-1" aria-hidden="true"></i>Limpar
                </button>
                <button type="button" id="btnFiltrar" class="btn btn-primary btn-sm">
                  <i class="bi bi-funnel-fill me-1" aria-hidden="true"></i>Filtrar
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- ── Card da Tabela ── -->
      <div class="table-card mb-3">

        <!-- Overlay de carregamento -->
        <div id="overlayLoading" class="overlay-loading d-none" role="status" aria-live="polite">
          <div class="spinner-border" aria-hidden="true"></div>
          <span>Carregando...</span>
        </div>

        <!-- Toolbar: exportação (esquerda) + search box (direita) -->
        <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
          <div class="d-flex gap-2">
            <button id="btnExpPDF" type="button" class="btn btn-export-pdf btn-sm"
                    title="Exportar para PDF">
              <i class="bi bi-file-earmark-pdf me-1" aria-hidden="true"></i>Exportar PDF
            </button>
          </div>
          <div class="search-box">
            <input type="text" id="searchBox" class="form-control form-control-sm"
                   placeholder="Buscar..." autocomplete="off" aria-label="Buscar na tabela" />
            <span class="search-hotkey" aria-hidden="true">
              <kbd>CTRL</kbd><kbd>K</kbd>
            </span>
          </div>
        </div>

        <!-- Tabela de dados -->
        <div class="table-wrapper">
          <table class="table table-bordered table-hover table-sm mb-0">
            <thead>
              <tr>
                <th scope="col">Pedido</th>
                <th scope="col">Cliente</th>
                <th scope="col">Ped. Compra</th>
                <th scope="col">Referência</th>
                <th scope="col">NF / Série</th>
                <th scope="col">Data</th>
                <th scope="col">Hora</th>
                <th scope="col">Qtde</th>
                <th scope="col">Peso</th>
                <th scope="col">Área</th>
                <th scope="col">Encarregado</th>
                <th scope="col">Dt. Entrega</th>
                <th scope="col">Local Entrega</th>
              </tr>
            </thead>
            <tbody id="tabelaBody">
              <!-- preenchido via AJAX -->
            </tbody>
          </table>
        </div>

      </div><!-- /table-card -->

      <!-- ── Resumo ── -->
      <div class="summary-card mb-3">
        <div class="summary-header">
          <i class="bi bi-bar-chart-line me-2" style="color:var(--color-text-muted);" aria-hidden="true"></i>Resumo do carregamento
        </div>
        <div class="table-responsive">
          <table class="table table-bordered table-sm mb-0 summary-table">
            <thead>
              <tr>
                <th rowspan="2">Registros</th>
                <th colspan="5" class="th-group-peso">Peso (kg)</th>
                <th colspan="5" class="th-group-area">Área (m²)</th>
                <th colspan="3">Frota</th>
              </tr>
              <tr>
                <th>Papel</th><th>Chapa</th><th>Maleta</th><th>Corte Vinco</th><th>Total</th>
                <th class="th-group-area">Papel</th>
                <th class="th-group-area">Chapa</th>
                <th class="th-group-area">Maleta</th>
                <th class="th-group-area">Corte Vinco</th>
                <th class="th-group-area">Total</th>
                <th>Total</th><th>Própria</th><th>Terceirizada</th>
              </tr>
            </thead>
            <tbody id="resumoBody">
              <!-- preenchido via AJAX -->
            </tbody>
          </table>
        </div>
      </div>

    </div><!-- /container-fluid -->

    <!-- ── Toasts ── -->
    <div class="toast-container-custom" aria-live="polite" aria-atomic="true"></div>

    <!-- ═══════════════════════════════════════════════════════════════
         BLOCO 1 — Regras de Negócio (puras, sem DOM/jQuery/jsPDF)
         Testável isoladamente: cada função recebe e devolve primitivos.
         ═══════════════════════════════════════════════════════════════ -->
    <script>
      'use strict';

      // ─── § 1. Formatação de data ───────────────────────────────────
      // Converte AAAA-MM-DD (input[type=date]) para DD/MM/AAAA (exibição)
      function formatarDataIso(iso) {
        if (!iso || typeof iso !== 'string') return '';
        const partes = iso.split('-');
        if (partes.length !== 3) return '';
        const [ano, mes, dia] = partes;
        if (!ano || !mes || !dia) return '';
        return `${dia}/${mes}/${ano}`;
      }

      // ─── § 2. Formatação de placa ──────────────────────────────────
      // 7 chars (Mercosul): últimos 3 + " " + primeiros 4
      // 6 chars (antigo):   últimos 2 + " " + primeiros 4
      // vazio: "(SEM PLACA)"
      function formatarPlaca(placa) {
        const p = (placa || '').trim();
        if (p === '' || p === '(SEM PLACA)') return '(SEM PLACA)';
        if (p.length === 7) return p.slice(-3) + ' ' + p.slice(0, 4);
        if (p.length === 6) return p.slice(-2) + ' ' + p.slice(0, 4);
        return p;
      }

      // ─── § 3. Classificação de frota ──────────────────────────────
      // Própria: nome começa com "INPA" (case-insensitive, após trim)
      function isFrotaPropria(nome) {
        return (nome || '').trim().toUpperCase().startsWith('INPA');
      }

      // ─── § 4. Cálculo de peso por item ────────────────────────────
      // Prioridade: vPesoSP > 0  →  B1_PESO × Qtd  →  0
      function calcularPesoItem(vPesoSP, vB1Peso, vQtd) {
        if (vPesoSP > 0) return vPesoSP;
        if (vB1Peso > 0 && vQtd > 0) return vB1Peso * vQtd;
        return 0;
      }

      // ─── § 5. Cálculo de área por item ────────────────────────────
      // Prioridade: vAreaSP > 0  →  (Larg mm / 1000) × (Comp mm / 1000) × Qtd  →  0
      function calcularAreaItem(vAreaSP, vLarg, vComp, vQtd) {
        if (vAreaSP > 0) return vAreaSP;
        if (vLarg > 0 && vComp > 0 && vQtd > 0) {
          return (vLarg / 1000) * (vComp / 1000) * vQtd;
        }
        return 0;
      }

      // ─── § 6. Classificação de produto (B1_DESCNF) ────────────────
      // "900"→PA  "901"→ML  "902"→CV  "903"→CH  outros→null
      function classificarTipoProduto(descnf) {
        const mapa = { '900': 'PA', '901': 'ML', '902': 'CV', '903': 'CH' };
        return mapa[(descnf || '').trim()] ?? null;
      }

      // ─── § 7. Classificação de linha DOM ──────────────────────────
      // Precedência: spacer > placa > subtotal > dado
      function classificarLinhaDom(classes) {
        if (classes.includes('row-spacer'))   return 'spacer';
        if (classes.includes('row-placa'))    return 'placa';
        if (classes.includes('row-subtotal')) return 'subtotal';
        return 'dado';
      }

      // ─── § 8. Linha de subtotal para o autoTable ──────────────────
      // Layout físico do DOM .row-subtotal:
      //   [0] <td colspan="7">  [1] "Subtotal:"  [2] Peso  [3] Área  [4] <td colspan="3">
      // Posições lógicas no cabeçalho da tabela: Subtotal=7  Peso=8  Área=9
      function montarLinhaSubtotalPdf(cellsDOM, totalColunas, estilo) {
        const linha = new Array(totalColunas).fill('');
        linha[7] = 'Subtotal:';
        linha[8] = cellsDOM[2] || '';
        linha[9] = cellsDOM[3] || '';
        return linha.map(v => ({ content: v, styles: estilo }));
      }

      // ─── § 9. Body completo para o autoTable ──────────────────────
      function montarBodyPdf(linhas, totalColunas, estilos) {
        const body = [];
        let grupoAnterior = '';

        for (const item of linhas) {
          if (item.tipo === 'dado') {
            if (item.grupo !== grupoAnterior) {
              body.push([{ content: item.grupo, colSpan: totalColunas, styles: estilos.grupo }]);
              grupoAnterior = item.grupo;
            }
            body.push(item.cells);
          } else if (item.tipo === 'subtotal') {
            body.push(montarLinhaSubtotalPdf(item.cells, totalColunas, estilos.subtotal));
          }
        }

        return body;
      }
    </script>

    <!-- ═══════════════════════════════════════════════════════════════
         BLOCO 2 — Aplicação (UI, DOM, jQuery, jsPDF)
         Consome as funções do Bloco 1. Não contém lógica de negócio.
         Testes: rodar mapaCarregamento.test.js localmente antes do deploy.
         ═══════════════════════════════════════════════════════════════ -->
    <script>
      'use strict';

      const TOAST_DELAY = 4500;
      const DEBOUNCE_MS = 280;

      // ─── Utilidades de formatação de números (exibição) ────────────
      // Converte número para string com separador de milhar brasileiro (ponto)
      // ex.: 1234567 → "1.234.567"
      function formatarInteiro(n) {
        return Math.round(n).toLocaleString('pt-BR');
      }

      // ─── Toast ────────────────────────────────────────────────────
      function exibirToast(tipo, mensagem) {
        const icones = { sucesso:'bi-check-circle-fill', erro:'bi-x-circle-fill', aviso:'bi-exclamation-triangle-fill', info:'bi-info-circle-fill' };
        const cores  = { sucesso:'var(--color-success)', erro:'var(--color-danger)', aviso:'var(--color-warning)', info:'var(--color-text-muted)' };
        const icone  = icones[tipo] || icones.info;
        const cor    = cores[tipo]  || cores.info;
        const $toast = $(`
          <div class="toast-custom">
            <div style="height:3px;background:${cor};animation:toastProgressShrink ${TOAST_DELAY}ms linear forwards;"></div>
            <div style="display:flex;align-items:flex-start;gap:.6rem;padding:.75rem 1rem;">
              <i class="bi ${icone}" style="font-size:1rem;color:${cor};flex-shrink:0;margin-top:.1rem;" aria-hidden="true"></i>
              <span style="font-size:.85rem;color:var(--color-text);flex:1;">${mensagem}</span>
              <button type="button" style="background:none;border:none;padding:0;cursor:pointer;color:var(--color-text-faint);font-size:.9rem;line-height:1;" aria-label="Fechar" onclick="$(this).closest('.toast-custom').remove()">&times;</button>
            </div>
          </div>`);
        $('.toast-container-custom').append($toast);
        setTimeout(() => $toast.remove(), TOAST_DELAY);
      }

      // ─── Renderização da tabela ────────────────────────────────────

      // Renderiza as linhas da tabela a partir do array de registros JSON.
      // O agrupamento por placa (cabeçalho + subtotal) é feito aqui,
      // usando as regras de negócio do Bloco 1.
      function renderizarTabela(registros) {
        const $tbody = $('#tabelaBody').empty();

        if (registros.length === 0) {
          $tbody.append(`
            <tr>
              <td colspan="13" class="text-center" style="color:var(--color-text-muted);font-weight:600;padding:2rem;">
                Nenhum registro encontrado para o filtro informado.
              </td>
            </tr>`);
          return;
        }

        let placaAtual    = null;
        let subTotalPeso  = 0;
        let subTotalArea  = 0;

        registros.forEach((reg, idx) => {
          const placaMudou = reg.placa !== placaAtual;

          // Fecha subtotal do grupo anterior
          if (placaMudou && placaAtual !== null) {
            $tbody.append(`
              <tr class="row-subtotal">
                <td colspan="7"></td>
                <td>Subtotal:</td>
                <td>${formatarInteiro(subTotalPeso)}</td>
                <td>${formatarInteiro(subTotalArea)}</td>
                <td colspan="3"></td>
              </tr>
              <tr class="row-spacer"><td colspan="13"></td></tr>`);
            subTotalPeso = 0;
            subTotalArea = 0;
          }

          // Cabeçalho de placa
          if (placaMudou) {
            const placaFmt = formatarPlaca(reg.placa);
            $tbody.append(`
              <tr class="row-placa">
                <td colspan="13">
                  <i class="bi bi-truck me-2" style="color:var(--color-text-muted);" aria-hidden="true"></i>
                  <strong>${placaFmt}</strong>
                  <span style="color:var(--color-text-muted);margin-left:.5rem;">${reg.transportadora} | ${reg.tel}</span>
                </td>
              </tr>`);
            placaAtual = reg.placa;
          }

          // Acumula subtotal
          subTotalPeso += reg.peso;
          subTotalArea += reg.area;

          // Linha de dado
          $tbody.append(`
            <tr>
              <td>${reg.pedido}</td>
              <td style="text-align:left !important;">${reg.cliente}</td>
              <td>${reg.pedCompra}</td>
              <td style="text-align:left !important;">${reg.referencia}</td>
              <td>${reg.nfSerie}</td>
              <td>${reg.dataEmissao}</td>
              <td>${reg.hora}</td>
              <td>${formatarInteiro(reg.qtd)}</td>
              <td>${formatarInteiro(reg.peso)}</td>
              <td>${formatarInteiro(reg.area)}</td>
              <td>${reg.encarregado}</td>
              <td>${reg.dtEntrega}</td>
              <td>${reg.localEntrega}</td>
            </tr>`);
        });

        // Fecha subtotal do último grupo
        if (placaAtual !== null) {
          $tbody.append(`
            <tr class="row-subtotal">
              <td colspan="7"></td>
              <td>Subtotal:</td>
              <td>${formatarInteiro(subTotalPeso)}</td>
              <td>${formatarInteiro(subTotalArea)}</td>
              <td colspan="3"></td>
            </tr>`);
        }
      }

      // ─── Renderização do resumo ────────────────────────────────────
      function renderizarResumo(resumo) {
        const p = resumo.peso;
        const a = resumo.area;
        const f = resumo.frota;

        $('#resumoBody').html(`
          <tr>
            <td>${resumo.totalRegistros}</td>
            <td>${formatarInteiro(p.PA)}</td>
            <td>${formatarInteiro(p.CH)}</td>
            <td>${formatarInteiro(p.ML)}</td>
            <td>${formatarInteiro(p.CV)}</td>
            <td style="font-weight:600;">${formatarInteiro(p.total)}</td>
            <td>${formatarInteiro(a.PA)}</td>
            <td>${formatarInteiro(a.CH)}</td>
            <td>${formatarInteiro(a.ML)}</td>
            <td>${formatarInteiro(a.CV)}</td>
            <td style="font-weight:600;">${formatarInteiro(a.total)}</td>
            <td>${f.total}</td>
            <td>${f.propria}</td>
            <td>${f.terceirizada}</td>
          </tr>`);
      }

      // ─── Módulo principal ──────────────────────────────────────────
      const MapaCarregamento = (() => {

        // Guarda o último payload JSON para o PDF evitar nova requisição
        let _dadosAtuais = null;

        const validar = () => {
          const $dataIn = $('#DataIn');
          if (!$dataIn.val()) {
            $dataIn.trigger('focus');
            exibirToast('aviso', 'Informe a data de referência.');
            return false;
          }
          return true;
        };

        const filtrar = () => {
          if (!validar()) return;

          const dataIn  = $('#DataIn').val();
          const filial  = $('#cboFilial').val();

          _dadosAtuais = null;
          $('#tabelaBody').empty();
          $('#resumoBody').empty();
          $('#overlayLoading').removeClass('d-none');

          $.ajax({
            url:      'MapaCarregamento01Repositorio.asp',
            method:   'GET',
            dataType: 'json',
            data:     { DataIn: dataIn, cboFilial: filial },
            success: function (data) {
              if (!data.ok) {
                exibirToast('erro', data.erro || 'Erro ao carregar os dados.');
                return;
              }
              _dadosAtuais = data;
              renderizarTabela(data.registros);
              renderizarResumo(data.resumo);
            },
            error: function (xhr) {
              console.error('AJAX error:', xhr.responseText);
              exibirToast('erro', 'Falha na comunicação com o servidor.');
            },
            complete: function () {
              $('#overlayLoading').addClass('d-none');
            }
          });
        };

        const exportarPdf = () => {
          const dataIn = $('#DataIn').val();
          const filial = $('#cboFilial option:selected').text().trim();

          if (!dataIn) {
            $('#DataIn').trigger('focus');
            exibirToast('aviso', 'Informe a data antes de exportar o PDF.');
            return;
          }

          if (!_dadosAtuais || _dadosAtuais.registros.length === 0) {
            exibirToast('aviso', 'Filtre os dados antes de exportar o PDF.');
            return;
          }

          const dataFormatada = formatarDataIso(dataIn);
          const { jsPDF } = window.jspdf;
          const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
          const pageW = doc.internal.pageSize.getWidth();
          const pageH = doc.internal.pageSize.getHeight();

          // ── Cabeçalho ──
          doc.setFontSize(13);
          doc.setFont('helvetica', 'bold');
          doc.text('Mapa de Carregamento', pageW / 2, 14, { align: 'center' });
          doc.setFontSize(8);
          doc.setFont('helvetica', 'normal');
          doc.text(`Data: ${dataFormatada}  |  Filial: ${filial}`, pageW / 2, 20, { align: 'center' });

          // ── Monta linhas para o autoTable diretamente do JSON ──
          // Sem precisar ler o DOM: os dados já estão em _dadosAtuais.
          const cabecalhos = ['Pedido','Cliente','Ped. Compra','Referência','NF / Série','Data','Hora','Qtde','Peso','Área','Encarregado','Dt. Entrega','Local Entrega'];

          const ESTILO_GRUPO    = { fontStyle:'bold', fillColor:[230,230,232], textColor:[26,26,28] };
          const ESTILO_SUBTOTAL = { fontStyle:'bold', fillColor:[218,218,220], textColor:[26,26,28] };

          // Constrói a estrutura de linhas intermediária (compatível com montarBodyPdf do Bloco 1)
          const linhas = [];
          let placaAnterior = null;

          _dadosAtuais.registros.forEach(reg => {
            if (reg.placa !== placaAnterior) {
              const placaFmt = formatarPlaca(reg.placa);
              linhas.push({
                tipo:  'placa',
                grupo: `${placaFmt}  ${reg.transportadora}  ${reg.tel}`,
              });
              if (placaAnterior !== null) {
                linhas.push({ tipo: 'subtotal-marker' });
              }
              placaAnterior = reg.placa;
            }
            linhas.push({
              tipo:  'dado',
              grupo: `${formatarPlaca(reg.placa)}  ${reg.transportadora}  ${reg.tel}`,
              cells: [
                reg.pedido, reg.cliente, reg.pedCompra, reg.referencia,
                reg.nfSerie, reg.dataEmissao, reg.hora,
                formatarInteiro(reg.qtd),
                formatarInteiro(reg.peso),
                formatarInteiro(reg.area),
                reg.encarregado, reg.dtEntrega, reg.localEntrega,
              ],
              _peso: reg.peso,
              _area: reg.area,
            });
          });

          // Monta body com subtotais calculados diretamente do JSON
          const body = [];
          let grupoAnterior = '';
          let subPeso = 0, subArea = 0;

          linhas.forEach((item, idx) => {
            if (item.tipo === 'dado') {
              if (item.grupo !== grupoAnterior) {
                if (grupoAnterior !== '') {
                  // Subtotal do grupo que terminou
                  const sub = new Array(13).fill('');
                  sub[7] = 'Subtotal:';
                  sub[8] = formatarInteiro(subPeso);
                  sub[9] = formatarInteiro(subArea);
                  body.push(sub.map(v => ({ content: v, styles: ESTILO_SUBTOTAL })));
                  body.push([{ content: '', colSpan: 13, styles: { fillColor: [247,247,248] } }]);
                  subPeso = 0; subArea = 0;
                }
                body.push([{ content: item.grupo, colSpan: 13, styles: ESTILO_GRUPO }]);
                grupoAnterior = item.grupo;
              }
              body.push(item.cells);
              subPeso += item._peso;
              subArea += item._area;
            }
          });

          // Subtotal do último grupo
          if (grupoAnterior !== '') {
            const sub = new Array(13).fill('');
            sub[7] = 'Subtotal:';
            sub[8] = formatarInteiro(subPeso);
            sub[9] = formatarInteiro(subArea);
            body.push(sub.map(v => ({ content: v, styles: ESTILO_SUBTOTAL })));
          }

          doc.autoTable({
            head: [cabecalhos],
            body,
            startY: 26,
            theme:  'grid',
            styles: {
              fontSize: 4.5, cellPadding: 1.5,
              textColor: [26,26,28], lineColor: [229,229,231], lineWidth: 0.2,
              halign: 'center', valign: 'middle',
            },
            headStyles:          { fillColor:[26,26,28], textColor:[255,255,255], fontStyle:'bold' },
            alternateRowStyles:  { fillColor:[247,247,248] },
            columnStyles:        { 1: { halign:'left' }, 3: { halign:'left' } },
          });

          // Rodapé com número de página
          const totalPaginas = doc.internal.getNumberOfPages();
          for (let i = 1; i <= totalPaginas; i++) {
            doc.setPage(i);
            doc.setFontSize(7);
            doc.setFont('helvetica', 'normal');
            doc.text(`Página ${i} de ${totalPaginas}`, pageW - 10, pageH - 5, { align: 'right' });
          }

          doc.save(`MapaCarregamento_${dataIn}.pdf`);
          exibirToast('sucesso', 'PDF gerado com sucesso.');
        };

        const limpar = () => {
          $('#DataIn').val('');
          $('#cboFilial').val('00');
          $('#tabelaBody').empty();
          $('#resumoBody').empty();
          _dadosAtuais = null;
          exibirToast('sucesso', 'Filtro limpo.');
        };

        const bind = () => {
          $('#btnFiltrar').on('click', (e) => { e.preventDefault(); filtrar(); });
          $('#btnExpPDF').on('click',  (e) => { e.preventDefault(); exportarPdf(); });
          $('#btnClear').on('click',   (e) => { e.preventDefault(); limpar(); });

          $('#DataIn').on('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); filtrar(); }
          });

          // Search box — filtro client-side com debounce
          let searchTimer;
          $('#searchBox').on('input', function () {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => {
              const termo = $(this).val().toLowerCase().trim();
              $('#tabelaBody tr').each(function () {
                const $tr = $(this);
                if ($tr.hasClass('row-placa') || $tr.hasClass('row-subtotal') || $tr.hasClass('row-spacer')) return;
                $tr.toggle(termo === '' || $tr.text().toLowerCase().includes(termo));
              });
            }, DEBOUNCE_MS);
          });

          // CTRL+K → foco no search box
          $(document).on('keydown', (e) => {
            if (e.ctrlKey && e.key === 'k') { e.preventDefault(); $('#searchBox').trigger('focus'); }
          });
        };

        return { bind };
      })();

      $(document).ready(() => { MapaCarregamento.bind(); });
    </script>

  </body>
</html>