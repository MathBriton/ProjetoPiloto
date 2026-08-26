<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->
<%
' ============================================================
' PROCESSAMENTO — remover pedido liberado da lista
' ============================================================
if trim(Request("Pedido")) <> "" then

  SQL = "" &_
    "SELECT " &_
      "ISNULL(ZF_EXTRATO,'') EXTRATO, " &_
      "ISNULL(ZF_BLOQ,'') BLOQUEIO " &_
    "FROM SZF010 " &_
      "INNER JOIN SC5010 ON C5_NUM = ZF_NUMERO " &_
                       "AND C5_CLIENTE = ZF_CLIENTE " &_
                       "AND C5_LOJACLI = ZF_LOJACLI " &_
                       "AND SC5010.D_E_L_E_T_ = '' " &_
    "WHERE ZF_NUMERO = '" & trim(Request("Pedido")) & "' " &_
      "AND ZF_ESTATUS = 'LI' " &_
      "AND SZF010.D_E_L_E_T_ = '' "

  SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

  if NOT PIR_DADOSADV.EOF then

    if trim(PIR_DADOSADV("EXTRATO")) = "" and trim(PIR_DADOSADV("BLOQUEIO")) <> "04" then
      msgFeedback = "Este pedido não pode ser excluído da lista, pois ainda não foi integrado."
      tipoFeedback = "aviso"
    else

      SQL = "" &_
        "UPDATE SZF010 SET " &_
          "ZF_ESTATUS = 'OK' " &_
        "WHERE ZF_NUMERO = '" & trim(Request("Pedido")) & "' " &_
          "AND D_E_L_E_T_ = '' "

      SET PIR_DADOSADV1 = Conecta( "PIR-DADOSADV", SQL )
      SET PIR_DADOSADV1 = NOTHING
      msgFeedback  = "Operação efetuada com sucesso."
      tipoFeedback = "sucesso"

    end if

  else
    msgFeedback  = "Este pedido não pode ser excluído da lista."
    tipoFeedback = "aviso"
  end if

  PIR_DADOSADV.Close()
  SET PIR_DADOSADV = NOTHING

end if

' ============================================================
' QUERY PRINCIPAL
' ============================================================
SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", "EXECUTE SP_LIBERAPEDIDOSZF" )
%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <title>Validação de Pedidos</title>
    <meta charset="utf-8" />
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

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>

    <style>
      /* ── Tokens de design ─────────────────────────────── */
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
        --color-success-soft:  hsl(145, 60%, 95%);
        --color-success-row:   hsl(145, 55%, 88%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);
        --color-danger-row:    hsl(3, 80%, 88%);
        --color-warning:       hsl(35, 100%, 48%);
        --color-warning-soft:  hsl(40, 100%, 95%);
        --color-warning-row:   hsl(40, 90%, 82%);
        --color-accent-row:    hsl(210, 75%, 85%);
        --color-neutral-row:   hsl(240, 6%, 86%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
        --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

        --radius-sm: 6px;
        --radius-md: 10px;
        --radius-lg: 14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-1: 4px;
        --space-2: 8px;
        --space-3: 12px;
        --space-4: 16px;
        --space-5: 24px;
      }

      /* ── Base ─────────────────────────────────────────── */
      body {
        background-color: var(--color-bg);
        color: var(--color-text);
        font-family: var(--font-sans);
        min-height: 100vh;
      }

      /* ── Título da página ─────────────────────────────── */
      .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
      }

      /* ── Card de tabela ───────────────────────────────── */
      .table-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        position: relative;
      }

      /* ── Legenda ──────────────────────────────────────── */
      .legenda-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        padding: .75rem 1.25rem;
        display: flex;
        flex-wrap: wrap;
        gap: .75rem 2rem;
      }
      .legenda-grupo {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: .35rem;
      }
      .legenda-label {
        font-size: .75rem;
        font-weight: 600;
        color: var(--color-text-muted);
        text-transform: uppercase;
        letter-spacing: .04em;
        margin-right: .25rem;
      }
      .badge-legenda {
        font-size: .72rem;
        font-weight: 500;
        padding: .25em .55em;
        border-radius: var(--radius-sm);
      }
      .badge-lg-status  { background-color: var(--color-surface-muted); color: var(--color-text-muted); border: 1px solid var(--color-border); }
      .badge-lg-success { background-color: var(--color-success-soft);  color: var(--color-success); border: 1px solid hsl(145, 55%, 78%); }
      .badge-lg-danger  { background-color: var(--color-danger-soft);   color: var(--color-danger);  border: 1px solid hsl(3, 80%, 78%);  }
      .badge-lg-warning { background-color: var(--color-warning-soft);  color: hsl(35, 100%, 34%);   border: 1px solid hsl(40, 90%, 72%);  }
      .badge-lg-info    { background-color: var(--color-accent-soft);   color: var(--color-accent);  border: 1px solid hsl(210, 75%, 75%); }
      .badge-lg-muted   { background-color: var(--color-neutral-row);   color: var(--color-text-muted); border: 1px solid var(--color-border-strong); }

      /* ── Toolbar (topo da tabela) ─────────────────────── */
      .table-toolbar {
        background-color: var(--color-surface);
        border-radius: var(--radius-md) var(--radius-md) 0 0;
        padding: .6rem 1.25rem;
        border-bottom: 1px solid var(--color-border);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: .5rem;
      }

      /* ── Search box ───────────────────────────────────── */
      .search-box { position: relative; max-width: 200px; }
      .search-box .form-control { padding-right: 76px; font-size: .82rem; }
      .search-box .search-hotkey {
        position: absolute; right: 8px; top: 50%;
        transform: translateY(-50%);
        pointer-events: none; user-select: none;
        display: flex; gap: 2px;
      }
      .search-box .search-hotkey kbd {
        font-size: .65rem; padding: .1rem .35rem;
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        color: var(--color-text-muted);
        font-family: var(--font-mono);
      }

      /* ── Tabela ───────────────────────────────────────── */
      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }
      .table-wrapper table { min-width: max-content; }

      .table thead th {
        position: sticky; top: 0;
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        font-size: .76rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: .03em;
        text-align: center; white-space: nowrap;
        z-index: 2; border-color: var(--color-border);
        padding: .6rem .75rem;
      }

      .table tbody td {
        font-size: .85rem; text-align: center;
        white-space: nowrap; color: var(--color-text);
        vertical-align: middle; padding: .55rem .75rem;
        border-color: var(--color-border);
      }

      /* Hover genérico — nunca sobrescreve linhas com estado */
      .table tbody tr:not([class*="row-"]):hover td {
        background-color: var(--color-surface-muted);
      }

      /* Classes de estado de linha */
      .table > tbody > tr.row-success > td { background-color: var(--color-success-row); }
      .table > tbody > tr.row-danger  > td { background-color: var(--color-danger-row);  }
      .table > tbody > tr.row-warning > td { background-color: var(--color-warning-row); }
      .table > tbody > tr.row-info    > td { background-color: var(--color-accent-row);  }
      .table > tbody > tr.row-muted   > td { background-color: var(--color-neutral-row); }

      .table > tbody > tr.row-success:hover > td,
      .table > tbody > tr.row-danger:hover  > td,
      .table > tbody > tr.row-warning:hover > td,
      .table > tbody > tr.row-info:hover    > td,
      .table > tbody > tr.row-muted:hover   > td { filter: brightness(.95); }

      /* Cor do texto nunca é alterada pelo estado da linha */
      .table > tbody > tr td { color: var(--color-text); }

      /* ── Link de número do pedido ─────────────────────── */
      .link-pedido {
        color: var(--color-accent);
        text-decoration: none;
        font-weight: 500;
        transition: color var(--transition);
      }
      .link-pedido:hover { color: var(--color-accent-hover); text-decoration: underline; }
      .link-pedido-bloq  { color: var(--color-text-muted); cursor: default; text-decoration: none; }

      /* ── Botão de remover linha ───────────────────────── */
      .btn-remover {
        padding: .25rem .45rem;
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        color: var(--color-danger);
        font-size: .8rem; line-height: 1;
        transition: background-color var(--transition), border-color var(--transition);
        cursor: pointer;
      }
      .btn-remover:hover {
        background-color: var(--color-danger-soft);
        border-color: hsl(3, 80%, 78%);
      }

      /* ── Rodapé da tabela ─────────────────────────────── */
      .table-footer {
        display: flex; align-items: center;
        justify-content: space-between;
        padding: .7rem 1.25rem;
        border-top: 1px solid var(--color-border);
        flex-wrap: wrap; gap: .5rem;
      }
      .table-footer-right { min-width: 140px; text-align: right; padding-right: .25rem; }
      .info-registros { font-size: .8rem; color: var(--color-text-muted); }

      /* ── Toast ────────────────────────────────────────── */
      .toast-container-custom {
        position: fixed; top: 1rem; right: 1rem;
        z-index: 1100;
        display: flex; flex-direction: column; gap: .5rem;
      }
      .toast-custom {
        min-width: 280px; max-width: 360px;
        border-radius: var(--radius-md); overflow: hidden;
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
      @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

      /* ── Placeholder (tabela vazia) ───────────────────── */
      .table-placeholder {
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        padding: 4rem 1rem;
        color: var(--color-text-muted); gap: .85rem;
      }
      .table-placeholder .bi { font-size: 2.4rem; color: var(--color-text-faint); }
      .table-placeholder p   { margin: 0; font-size: .88rem; text-align: center; }

      /* ── Loading overlay ──────────────────────────────── */
      #loading-overlay {
        display: none;
        position: fixed;
        inset: 0;
        z-index: 9999;
        background-color: hsla(0, 0%, 0%, .45);
        align-items: center;
        justify-content: center;
        flex-direction: column;
        gap: 1rem;
      }

      #loading-overlay.active {
        display: flex;
      }

      #loading-overlay .loading-box {
        background-color: var(--color-surface);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-md);
        padding: 2rem 2.5rem;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: .85rem;
        min-width: 220px;
      }

      #loading-overlay .loading-box .spinner-border {
        width: 2.25rem;
        height: 2.25rem;
        border-width: 3px;
        color: var(--color-accent);
      }

      #loading-overlay .loading-box .loading-label {
        font-size: .88rem;
        font-weight: 500;
        color: var(--color-text-muted);
      }

      /* ── Acessibilidade ───────────────────────────────── */
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
          animation-duration: 0.01ms !important;
          transition-duration: 0.01ms !important;
        }
      }

      /* ── Responsividade abaixo de 768 px ─────────────── */
      @media (max-width: 768px) {
        .legenda-card { flex-direction: column; gap: .65rem; }
        .table-toolbar { flex-direction: column; align-items: stretch; }
        .search-box { max-width: 100%; }
        .table-footer { flex-direction: column; align-items: stretch; }
        .table-footer-right { text-align: left; }
      }
    </style>
  </head>

  <body>

    <!-- ── Loading overlay ──────────────────────────────────────────── -->
    <div id="loading-overlay" role="status" aria-live="polite" aria-label="Carregando pedido">
      <div class="loading-box">
        <div class="spinner-border" aria-hidden="true"></div>
        <span class="loading-label">Aguarde, carregando pedido…</span>
      </div>
    </div>

    <div class="container-fluid py-3 px-3 px-lg-4">

      <!-- Cabeçalho da página -->
      <div class="d-flex align-items-center mb-3">
        <span class="page-title">Validação de Pedidos de Caixa e Chapa</span>
      </div>

      <!-- Legenda -->
      <div class="legenda-card mb-3" role="note" aria-label="Legendas da tabela">

        <div class="legenda-grupo">
          <span class="legenda-label">Status</span>
          <span class="badge-legenda badge-lg-status">LI — Liberado</span>
          <span class="badge-legenda badge-lg-status">PR — Bloq. por Preço</span>
          <span class="badge-legenda badge-lg-status">DE — Data de Entrega</span>
          <span class="badge-legenda badge-lg-status">CP — Bloq. Cond. Pagamento</span>
        </div>

        <div class="legenda-grupo">
          <span class="legenda-label">Cor da linha</span>
          <span class="badge-legenda badge-lg-danger">Bloqueados</span>
          <span class="badge-legenda badge-lg-warning">Livres de débito</span>
          <span class="badge-legenda badge-lg-success">Liberados</span>
          <span class="badge-legenda badge-lg-info">Pendentes de liberação</span>
          <span class="badge-legenda badge-lg-muted">Lote Piloto</span>
        </div>

      </div>

      <!-- Card de tabela -->
      <form name="ValidaPedidos" id="ValidaPedidos" action="ListaPedidosAValidar.asp" method="post">
        <input type="hidden" id="Pedido" name="Pedido" value="" />

        <div class="table-card">

          <!-- Toolbar -->
          <div class="table-toolbar">
            <div class="d-flex gap-2">
              <button type="button" id="btnExportCsv" class="btn btn-sm" style="background-color:var(--color-success);border-color:var(--color-success);color:hsl(0,0%,100%);font-weight:500;font-size:.8rem;">
                <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
              </button>
            </div>
            <div class="search-box">
              <input type="text" id="searchBox" class="form-control form-control-sm"
                     placeholder="Buscar..." autocomplete="off"
                     aria-label="Buscar na tabela" />
              <span class="search-hotkey" aria-hidden="true">
                <kbd>CTRL</kbd><kbd>K</kbd>
              </span>
            </div>
          </div>

          <!-- Wrapper da tabela -->
          <div class="table-wrapper">
            <table class="table table-bordered table-hover table-sm mb-0" id="tabelaPedidos">
              <thead>
                <tr>
                  <th>Pedido</th>
                  <th>Cliente</th>
                  <th>Loja</th>
                  <th>Nome</th>
                  <th>FT</th>
                  <th>Produto</th>
                  <th>Tipo</th>
                  <th>Fab.</th>
                  <th>Quant.</th>
                  <th>Preço</th>
                  <th>Preço/KG</th>
                  <th>Preço/Kg Net</th>
                  <th>M²</th>
                  <th>Emissão</th>
                  <th>Entrega</th>
                  <th>Obs.</th>
                  <th>Status</th>
                  <th>Bloqueio</th>
                  <th>Ação</th>
                  <th>Quem Emitiu</th>
                </tr>
              </thead>
              <tbody>
<%
                dim totalRegistros
                totalRegistros = 0

                while not PIR_DADOSADV.EOF

                  totalRegistros = totalRegistros + 1

                  Fabricacao = iif( (PIR_DADOSADV("B1_INPUBE") = "U"), "UBERABA", "MATRIZ" )

                  select case trim(PIR_DADOSADV("ZF_ESTATUS"))
                  case "LI"                 classeRow = "row-success"
                  case "PR","DE","CP","LC"  classeRow = "row-danger"
                  case else                 classeRow = "row-info"
                  end select

                  if trim(PIR_DADOSADV("ZF_LOTEPIL")) = "S" then
                    classeRow = "row-muted"
                  end if

                  if inArray( trim(PIR_DADOSADV("ZF_TES")), array("580", "581", "715", "716", "626", "809") ) then
                    classeRow  = "row-warning"
                    tesExcecao = 1
                  else
                    tesExcecao = 0
                  end if

                  lotepiloto = iif( (trim(PIR_DADOSADV("ZF_LOTEPIL")) = "S" and tesExcecao <> 1), 1, 0 )
                  remessa    = iif( (trim(PIR_DADOSADV("ZF_REMESSA")) = "F"), 1, 0 )

                  SQL = "EXEC SP_CALCULAPESO '" & trim(PIR_DADOSADV("ZF_PRODUTO")) & "'"
                  SET PIR_DADOSADV3 = Conecta( "PIR-DADOSADV", SQL )
                  if NOT isObject( PIR_DADOSADV3 ) then
                    response.end
                  end if

                  tipoProduto = iif( (NOT PIR_DADOSADV3.EOF), PIR_DADOSADV3("TIPO"), "" )
                  estrutura   = iif( (NOT PIR_DADOSADV3.EOF), iif( (mid(trim(PIR_DADOSADV3("ESTRUTURA")),7,1) = ""), "SIMPLES", "DUPLA" ), "" )
                  peso        = iif( (NOT PIR_DADOSADV3.EOF), PIR_DADOSADV3("NPESO"), 0 )

                  PIR_DADOSADV3.Close()
                  SET PIR_DADOSADV3 = NOTHING

                  if left(trim(PIR_DADOSADV("ZF_PRODUTO")),3) = "720" then

                    SQL = "SELECT dbo.funcCalculaArea('" & trim(PIR_DADOSADV("ZF_PRODUTO")) & "', '" & PIR_DADOSADV("ZF_NUMERO") & "') AS AREA"
                    SET PIR_DADOSADV3 = Conecta( "PIR-DADOSADV", SQL)
                    if NOT isObject( PIR_DADOSADV3 ) then
                      response.end
                    end if

                    if NOT PIR_DADOSADV3.EOF then
                      valorM2        = (PIR_DADOSADV3("AREA") * PIR_DADOSADV("ZF_QUANT"))
                      areaChapa      = PIR_DADOSADV3("AREA")
                      precoAreaChapa = iif( (cdbl(areaChapa) > 0) and (cdbl(PIR_DADOSADV("ZF_PRECO")) > 0), cdbl(PIR_DADOSADV("ZF_PRECO")) / cdbl(areaChapa), 0 )
                      areaTotalChapa = iif( (cdbl(areaChapa) > 0), areaChapa * PIR_DADOSADV("ZF_QUANT"), 0 )
                    end if

                    PIR_DADOSADV3.Close()
                    SET PIR_DADOSADV3 = NOTHING

                    if (peso > 0) and (PIR_DADOSADV("ZF_PRECO") > 0) and (PIR_DADOSADV("ZF_LARG") > 0) and (PIR_DADOSADV("ZF_COMP") > 0) then
                      PKG       = iif( (peso > 0), PIR_DADOSADV("ZF_PRECO") / ((PIR_DADOSADV("ZF_LARG") * PIR_DADOSADV("ZF_COMP") * peso) / 1000000), 0 )
                      pesoTotal = iif( (peso > 0), PIR_DADOSADV("ZF_QUANT") * ((PIR_DADOSADV("ZF_LARG") * PIR_DADOSADV("ZF_COMP") * peso) / 1000000), 0 )
                    else
                      PKG       = 0
                      pesoTotal = 0
                    end if

                  else

                    PKG       = iif( (peso > 0), PIR_DADOSADV("ZF_PRECO") / peso, 0 )
                    pesoTotal = iif( (peso > 0), PIR_DADOSADV("ZF_QUANT") * peso, 0 )

                    SQL = "SELECT dbo.funcCalculaArea('" & trim(PIR_DADOSADV("ZF_PRODUTO")) & "', '" & PIR_DADOSADV("ZF_NUMERO") & "') AS AREA"
                    SET PIR_DADOSADV4 = Conecta("PIR-DADOSADV", SQL)
                    if NOT isObject( PIR_DADOSADV4 ) then
                      response.end
                    end if

                    if NOT PIR_DADOSADV4.EOF then
                      valorM2 = (PIR_DADOSADV4("AREA") * PIR_DADOSADV("ZF_QUANT"))
                    end if

                    PIR_DADOSADV4.Close()
                    SET PIR_DADOSADV4 = NOTHING

                  end if

                  precoKG    = cstr(PKG)
                  precoKG    = Left( precoKG, InStr( 1, precoKG, ",") + 5 )
                  PrecoNet   = ( PIR_DADOSADV("ZF_PRECO") - PIR_DADOSADV("ZF_VLICM") - PIR_DADOSADV("ZF_VLPIS") - PIR_DADOSADV("ZF_VLCOFINS") )
                  PrecoKgNet = iif( (PIR_DADOSADV("ZF_PESOUNI") > 0), PrecoNet / PIR_DADOSADV("ZF_PESOUNI"), 0 )
%>
                <tr class="<%=classeRow%>">

                  <td>
<%
                  if tesExcecao = 1 AND USR <> "CARLOSS" and USR <> "sandro" AND USR <> "Adriana.Bazzarello" then
%>
                    <a href="#" class="link-pedido-bloq" onclick="alertaTesExcecao(); return false;" title="Requer Gerente Comercial">
                      <i class="bi bi-lock-fill me-1" style="font-size:.75rem;color:var(--color-text-faint);"></i><%=PIR_DADOSADV("ZF_NUMERO")%>
                    </a>
<%
                  elseif lotepiloto = 1 AND USR <> "CARLOSS" and USR <> "rocharo" AND USR <> "Adriana.Bazzarello" then
%>
                    <a href="#" class="link-pedido-bloq" onclick="alertaLotePiloto(); return false;" title="Requer usuário autorizado">
                      <i class="bi bi-lock-fill me-1" style="font-size:.75rem;color:var(--color-text-faint);"></i><%=PIR_DADOSADV("ZF_NUMERO")%>
                    </a>
<%
                  else
%>
                    <a href="ValidaPedido.asp?Pedido=<%=trim(PIR_DADOSADV("ZF_NUMERO"))%>" class="link-pedido" onclick="document.getElementById('loading-overlay').classList.add('active')"><%=PIR_DADOSADV("ZF_NUMERO")%></a>
<%
                  end if
%>
                  </td>

                  <td><%=PIR_DADOSADV("ZF_CLIENTE")%></td>
                  <td><%=PIR_DADOSADV("ZF_LOJACLI")%></td>
                  <td><%=trim(PIR_DADOSADV("CLIENTE"))%></td>
                  <td><%=trim(PIR_DADOSADV("ZF_PRODUTO"))%></td>
                  <td><%=trim(PIR_DADOSADV("REF"))%></td>
                  <td><%=PIR_DADOSADV("DESCNF")%></td>
                  <td><%=Fabricacao%></td>
                  <td><%=PIR_DADOSADV("ZF_QUANT")%></td>
                  <td><%=FormatNumber(PIR_DADOSADV("ZF_PRECO"),6)%></td>
                  <td><%=FormatNumber(precoKG,6)%></td>
                  <td><%=FormatNumber(PrecoKgNet,6)%></td>
                  <td><%=FormatNumber(valorM2,6)%></td>
                  <td><%=AcertaData(PIR_DADOSADV("ZF_EMISSAO"))%></td>
                  <td><%=AcertaData(PIR_DADOSADV("ZF_ENT1"))%></td>
                  <td><%=trim(PIR_DADOSADV("OBS"))%></td>
                  <td><%=PIR_DADOSADV("ZF_ESTATUS")%></td>
                  <td><%=PIR_DADOSADV("ZF_DESCBL")%></td>

                  <td>
<%
                  if trim(PIR_DADOSADV("ZF_ESTATUS")) = "LI" and (trim(PIR_DADOSADV("ZF_EXTRATO")) <> "" or trim(PIR_DADOSADV("ZF_BLOQ")) = "04") then
%>
                    <button type="button" class="btn-remover"
                            onclick="confirmarRemocao('<%=trim(PIR_DADOSADV("ZF_NUMERO"))%>')"
                            title="Remover este pedido da lista"
                            aria-label="Remover pedido <%=trim(PIR_DADOSADV("ZF_NUMERO"))%>">
                      <i class="bi bi-trash3"></i>
                    </button>
<%
                  end if
%>
                  </td>

                  <td><%=PIR_DADOSADV("ZF_DIGITAD")%></td>
                </tr>
<%
                  PIR_DADOSADV.MoveNext()
                wend
                PIR_DADOSADV.Close()
                SET PIR_DADOSADV = NOTHING
%>
              </tbody>
            </table>
          </div>

          <!-- Rodapé -->
          <div class="table-footer">
            <div></div>
            <div class="table-footer-right">
              <span class="info-registros" id="infoRegistros" aria-live="polite">
                Total: <%=totalRegistros%> registro(s)
              </span>
            </div>
          </div>

        </div>
      </form>

    </div>

    <!-- Toast container -->
    <div class="toast-container-custom" aria-live="polite" aria-atomic="true"></div>

    <!-- Modal de confirmação de remoção -->
    <div class="modal fade" id="modalConfirmarRemocao" tabindex="-1"
         aria-labelledby="modalConfirmarRemocaoLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content" style="border-radius:var(--radius-lg);border:1px solid var(--color-border);">
          <div class="modal-header" style="border-bottom:1px solid var(--color-border);padding:1rem 1.25rem;">
            <h5 class="modal-title fw-semibold" id="modalConfirmarRemocaoLabel" style="font-size:.95rem;color:var(--color-text);">
              <i class="bi bi-exclamation-triangle-fill me-2" style="color:var(--color-warning);"></i>Confirmar remoção
            </h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body" style="padding:1.25rem;color:var(--color-text-muted);font-size:.9rem;">
            Tem certeza que deseja remover o pedido <strong id="numeroPedidoModal" style="color:var(--color-text);"></strong> da lista?
          </div>
          <div class="modal-footer" style="border-top:1px solid var(--color-border);padding:.85rem 1.25rem;gap:.5rem;">
            <button type="button" class="btn btn-sm btn-light" data-bs-dismiss="modal"
                    style="background:var(--color-surface);border:1px solid var(--color-border);color:var(--color-text);font-weight:500;font-size:.85rem;">
              Cancelar
            </button>
            <button type="button" class="btn btn-sm" id="btnConfirmarRemocao"
                    style="background-color:var(--color-danger);border-color:var(--color-danger);color:hsl(0,0%,100%);font-weight:500;font-size:.85rem;">
              <i class="bi bi-trash3 me-1"></i>Sim, remover pedido
            </button>
          </div>
        </div>
      </div>
    </div>

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <script>
      'use strict'

      /* ── Constantes ──────────────────────────────── */
      const DEBOUNCE_MS   = 280
      const TOAST_DELAY   = 4500

      /* ── Feedback do servidor (se houver) ────────── */
      const MSG_SERVIDOR  = <%
        if msgFeedback <> "" then
          Response.Write """" & msgFeedback & """"
        else
          Response.Write "null"
        end if
      %>
      const TIPO_FEEDBACK = <%
        if tipoFeedback <> "" then
          Response.Write """" & tipoFeedback & """"
        else
          Response.Write "null"
        end if
      %>

      /* ── Alertas de restrição ─────────────────────── */
      function alertaTesExcecao() {
        exibirToast('aviso', 'Este pedido possui uma TES destinada a PEDIDOS LIVRES DE DÉBITO.\nSomente o Gerente Comercial pode proceder com a validação.')
      }
      function alertaLotePiloto() {
        exibirToast('aviso', 'Este pedido é um LOTE PILOTO.\nSomente usuários autorizados podem proceder com a validação.')
      }
      function alertaRemessa() {
        exibirToast('aviso', 'Este pedido é uma operação triangular.\nSomente o Gerente Corporativo pode proceder com a validação.')
      }

      /* ── Modal de confirmação de remoção ─────────── */
      let pedidoParaRemover = ''
      const modalConfirmar  = new bootstrap.Modal(document.getElementById('modalConfirmarRemocao'))

      function confirmarRemocao(pedido) {
        pedidoParaRemover = pedido
        document.getElementById('numeroPedidoModal').textContent = pedido
        modalConfirmar.show()
      }

      document.getElementById('btnConfirmarRemocao').addEventListener('click', function () {
        document.getElementById('Pedido').value = pedidoParaRemover
        modalConfirmar.hide()
        document.getElementById('ValidaPedidos').submit()
      })

      /* ── Toast ────────────────────────────────────── */
      function exibirToast(tipo, mensagem) {
        const cfg = {
          sucesso: { icone: 'bi-check-circle-fill',        cor: 'var(--color-success)', titulo: 'Sucesso' },
          erro:    { icone: 'bi-x-circle-fill',            cor: 'var(--color-danger)',  titulo: 'Erro'    },
          aviso:   { icone: 'bi-exclamation-triangle-fill', cor: 'var(--color-warning)', titulo: 'Atenção' },
          info:    { icone: 'bi-info-circle-fill',          cor: 'var(--color-text-muted)', titulo: 'Info' }
        }
        const c = cfg[tipo] || cfg.info
        const el = document.createElement('div')
        el.className = 'toast-custom'
        el.innerHTML = `
          <div class="d-flex align-items-start p-3 gap-2">
            <i class="bi ${c.icone} mt-1" style="color:${c.cor};font-size:1rem;flex-shrink:0;"></i>
            <div class="flex-grow-1">
              <div style="font-size:.82rem;font-weight:600;color:${c.cor};margin-bottom:.2rem;">${c.titulo}</div>
              <div style="font-size:.83rem;color:var(--color-text-muted);white-space:pre-line;">${mensagem}</div>
            </div>
            <button onclick="this.closest('.toast-custom').remove()" style="background:none;border:none;cursor:pointer;color:var(--color-text-faint);font-size:1rem;line-height:1;padding:0;" aria-label="Fechar">&times;</button>
          </div>
          <div class="toast-progress-bar" style="background-color:${c.cor};"></div>
        `
        document.querySelector('.toast-container-custom').appendChild(el)
        setTimeout(() => el.remove(), TOAST_DELAY + 300)
      }

      /* ── Exibir feedback do servidor (se houver) ── */
      if (MSG_SERVIDOR) {
        exibirToast(TIPO_FEEDBACK || 'info', MSG_SERVIDOR)
      }

      /* ── Busca dinâmica (client-side) ─────────────── */
      let debounceTimer
      document.getElementById('searchBox').addEventListener('input', function () {
        clearTimeout(debounceTimer)
        const termo = this.value
        debounceTimer = setTimeout(() => filtrarTabela(termo), DEBOUNCE_MS)
      })

      function filtrarTabela(termo) {
        const t = termo.toLowerCase().trim()
        let visiveis = 0
        document.querySelectorAll('#tabelaPedidos tbody tr').forEach(function (tr) {
          const texto = tr.textContent.toLowerCase()
          const visivel = !t || texto.includes(t)
          tr.style.display = visivel ? '' : 'none'
          if (visivel) visiveis++
        })
        document.getElementById('infoRegistros').textContent = t
          ? `Exibindo ${visiveis} de <%=totalRegistros%> registro(s)`
          : `Total: <%=totalRegistros%> registro(s)`
      }

      /* ── Atalho Ctrl+K → foco na busca ───────────── */
      document.addEventListener('keydown', function (e) {
        if (e.ctrlKey && e.key === 'k') {
          e.preventDefault()
          document.getElementById('searchBox').focus()
        }
      })

      /* ── Exportar CSV ─────────────────────────────── */
      document.getElementById('btnExportCsv').addEventListener('click', function () {
        const cabecalhos = ['Pedido','Cliente','Loja','Nome','FT','Produto','Tipo','Fab.','Quant.','Preço','Preço/KG','Preço/Kg Net','M²','Emissão','Entrega','Obs.','Status','Bloqueio','Quem Emitiu']

        const linhas = []
        document.querySelectorAll('#tabelaPedidos tbody tr').forEach(function (tr) {
          if (tr.style.display === 'none') return
          const cells = tr.querySelectorAll('td')
          const linha = []
          cells.forEach(function (td, i) {
            if (i === 18) return  /* coluna Ação — não exportar */
            linha.push(td.textContent.trim())
          })
          linhas.push(linha)
        })

        exportarCSV('PedidosAValidar', cabecalhos, linhas)
      })

      function exportarCSV(nomeArquivo, cabecalhos, linhas) {
        const esc = function (v) {
          const s = v == null ? '' : String(v)
          return (s.includes(';') || s.includes('"') || s.includes('\n'))
            ? '"' + s.replace(/"/g, '""') + '"' : s
        }
        const rows = [
          cabecalhos.map(esc).join(';'),
          ...linhas.map(function (row) { return row.map(esc).join(';') })
        ]
        const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' })
        const url  = URL.createObjectURL(blob)
        const a    = document.createElement('a')
        a.href     = url
        a.download = nomeArquivo.endsWith('.csv') ? nomeArquivo : nomeArquivo + '.csv'
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        URL.revokeObjectURL(url)
      }
    </script>
  </body>
</html>
