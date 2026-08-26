<%
' ============================================================
'  grid_validadoresdecredito.asp  — arquivo único
'
'  Modo JSON  → ?api=1  (chamado pelo $.ajax interno)
'  Modo HTML  → acesso normal (sem ?api=1)
' ============================================================
Response.Buffer = True
Session.LCID    = 1046
%>
<!-- #include file="../../cnn/cnn.asp" -->
<%
' ── Bloco JSON — responde e encerra antes de qualquer HTML ──────────────
If LCase(Trim(Request.QueryString("api"))) = "1" Then

  Response.Clear
  Response.ContentType = "application/json"

  On Error Resume Next

  SQL = "SELECT DISTINCT " & _
          "vc.IdValidador AS VLD_COD, " & _
          "u.CRM_NOME    AS VLD_NOM, " & _
          "vp.Plantas    AS VLD_PLANTAS, " & _
          "vc.Nivel      AS VLD_NIVCOD, " & _
          "CASE " & _
            "WHEN vc.Nivel = 10 THEN 'Analista de Crédito' " & _
            "WHEN vc.Nivel = 20 THEN 'Coordenador de Crédito' " & _
            "WHEN vc.Nivel = 30 THEN 'Gerente de Tesouraria' " & _
            "WHEN vc.Nivel = 40 THEN 'Divisional Controller' " & _
            "WHEN vc.Nivel = 50 THEN 'Gerente de Planta' " & _
            "WHEN vc.Nivel = 60 THEN 'CFO' " & _
            "WHEN vc.Nivel = 70 THEN 'CEO' " & _
          "END AS VLD_NIVDSC, " & _
          "vc.Minimo AS VLD_VALMIN, " & _
          "vc.Maximo AS VLD_VALMAX " & _
        "FROM ValidadoresCredito vc " & _
          "LEFT JOIN HELPDESK..CRM_DTI u   ON vc.IdValidador = u.CRM_R_E_C_N_O " & _
          "LEFT JOIN HELPDESK..Empresas e   ON vc.Planta     = e.Empresa " & _
          "LEFT JOIN vw_ValidadorPlantas vp  ON vc.IdValidador = vp.IdValidador " & _
        "WHERE ISNULL(QuemExcluiu, 0) = 0 " & _
          "AND u.CRM_D_E_L_E_T = '' " & _
        "ORDER BY u.CRM_NOME"

  ' Dados11 é um Recordset pré-configurado pelo cnn.asp — usar direto, como o legado fazia
  Dados11.Open SQL

  If Err.Number <> 0 Then
    Response.Write "{""ok"":false,""erro"":""" & Replace(Err.Description, """", "'") & """}"
    Response.End
  End If

  ' ── Helpers ──────────────────────────────────────────────────────────
  Function J(val)
    J = Replace(CStr(val), "\",  "\\")
    J = Replace(J, """", "\""")
    J = Replace(J, Chr(13), "\r")
    J = Replace(J, Chr(10), "\n")
  End Function

  Function ValNum(v)
    If IsNull(v) Or v = "" Then ValNum = 0 Else ValNum = CDbl(v)
  End Function

  ' ── Serialização ─────────────────────────────────────────────────────
  Dim json : json    = "{""ok"":true,""registros"":["
  Dim primeiro       : primeiro = True

  Do While Not Dados11.EOF
    If Not primeiro Then json = json & ","
    primeiro = False

    Dim plantasRaw, plantasArr, plantasJson, p
    plantasRaw  = Trim(Dados11("VLD_PLANTAS") & "")
    plantasJson = "["
    If plantasRaw <> "" Then
      plantasArr = Split(plantasRaw, "|")
      Dim pfirst : pfirst = True
      For Each p In plantasArr
        p = Trim(p)
        If p <> "" Then
          If Not pfirst Then plantasJson = plantasJson & ","
          plantasJson = plantasJson & """" & J(p) & """"
          pfirst = False
        End If
      Next
    End If
    plantasJson = plantasJson & "]"

    json = json & "{" & _
      """cod"":"      & Dados11("VLD_COD")           & "," & _
      """nom"":"""    & J(Dados11("VLD_NOM"))         & """," & _
      """plantas"":"  & plantasJson                   & "," & _
      """nivCod"":"   & Dados11("VLD_NIVCOD")         & "," & _
      """nivDsc"":""" & J(Dados11("VLD_NIVDSC"))      & """," & _
      """valMin"":"   & ValNum(Dados11("VLD_VALMIN"))  & "," & _
      """valMax"":"   & ValNum(Dados11("VLD_VALMAX"))  & _
      "}"

    Dados11.MoveNext
  Loop

  json = json & "]}"

  Dados11.Close

  Response.Write json
  Response.End

End If

' ── Modo HTML — a partir daqui só executa quando NÃO for ?api=1 ─────────
strMsg = ""
Select Case Trim(UCase(Request("strStatus")))
  Case "INC" strMsg = "Cadastro realizado com sucesso!"
  Case "ALT" strMsg = "Cadastro atualizado com sucesso!"
  Case "DEL" strMsg = "Cadastro apagado com sucesso!"
End Select
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — Validadores de Crédito</title>

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    /* ── Tokens de Design ─────────────────────────────────────────────── */
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
      --color-success-hover: hsl(145, 63%, 27%);
      --color-success-soft:  hsl(145, 60%, 95%);

      --color-danger:        hsl(3, 90%, 48%);
      --color-danger-hover:  hsl(3, 90%, 42%);
      --color-danger-soft:   hsl(3, 100%, 96%);

      --color-warning:       hsl(55, 100%, 64%);
      --color-warning-hover: hsl(56, 82%, 57%);
      --color-warning-soft:  hsl(40, 100%, 95%);

      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
      --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

      --radius-sm: 6px;
      --radius-md: 10px;
      --radius-lg: 14px;

      --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
      --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);
      --shadow-lg: 0 12px 32px hsla(240, 6%, 10%, .12);

      --transition: 160ms ease;
    }

    /* ── Base ─────────────────────────────────────────────────────────── */
    body {
      background-color: var(--color-bg);
      min-height: 100vh;
      font-family: var(--font-sans);
      color: var(--color-text);
    }

    /* ── Título da página ─────────────────────────────────────────────── */
    .page-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
      font-weight: 600;
      letter-spacing: -0.01em;
      color: var(--color-text);
    }

    /* ── Card de tabela ───────────────────────────────────────────────── */
    .table-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-sm);
      position: relative;
    }

    /* ── Toolbar do topo ──────────────────────────────────────────────── */
    .table-toolbar {
      background-color: var(--color-surface);
      border-radius: var(--radius-md) var(--radius-md) 0 0;
    }

    /* ── Botão Novo ───────────────────────────────────────────────────── */
    .btn-novo {
      background-color: var(--color-success);
      border-color: var(--color-success);
      color: hsl(0, 0%, 100%);
      font-weight: 500;
      font-size: .85rem;
      transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-novo:hover        { background-color: var(--color-success-hover); border-color: var(--color-success-hover); color: hsl(0, 0%, 100%); }
    .btn-novo:active       { transform: scale(.98); }
    .btn-novo:focus-visible { box-shadow: 0 0 0 3px var(--color-success-soft); outline: none; }

    /* ── Botões de ação da linha ──────────────────────────────────────── */
    .btn-row-view {
      background-color: var(--color-accent);
      border-color: var(--color-accent);
      color: hsl(0, 0%, 100%);
      font-size: .78rem; padding: .25rem .5rem;
      border-radius: var(--radius-sm); line-height: 1.4;
      transition: background-color var(--transition);
    }
    .btn-row-view:hover { background-color: var(--color-accent-hover); border-color: var(--color-accent-hover); color: hsl(0, 0%, 100%); }

    .btn-row-edit {
      background-color: var(--color-warning);
      border-color: var(--color-warning);
      color: hsl(0, 0%, 0%);
      font-size: .78rem; padding: .25rem .5rem;
      border-radius: var(--radius-sm); line-height: 1.4;
      transition: background-color var(--transition);
    }
    .btn-row-edit:hover { background-color: var(--color-warning-hover); border-color: var(--color-warning-hover); color: hsl(0, 0%, 0%); }

    .btn-row-del {
      background-color: var(--color-danger);
      border-color: var(--color-danger);
      color: hsl(0, 0%, 100%);
      font-size: .78rem; padding: .25rem .5rem;
      border-radius: var(--radius-sm); line-height: 1.4;
      transition: background-color var(--transition);
    }
    .btn-row-del:hover { background-color: var(--color-danger-hover); border-color: var(--color-danger-hover); color: hsl(0, 0%, 100%); }

    /* ── Badges ───────────────────────────────────────────────────────── */
    .badge-planta {
      display: inline-block;
      padding: .18rem .48rem;
      font-size: .72rem; font-weight: 600;
      border-radius: var(--radius-sm);
      margin: 1px 2px; white-space: nowrap;
      cursor: default; letter-spacing: .01em;
    }
    .badge-nivel {
      display: inline-block;
      padding: .22rem .6rem;
      font-size: .75rem; font-weight: 500;
      border-radius: var(--radius-sm);
      white-space: nowrap; letter-spacing: .01em;
    }

    /* ── Search box ───────────────────────────────────────────────────── */
    .search-box { position: relative; max-width: 200px; }
    .search-box .form-control { padding-right: 76px; font-size: .82rem; }
    .search-box .form-control:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
    }
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

    /* ── Tabela ───────────────────────────────────────────────────────── */
    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
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
    .table tbody tr:hover td { background-color: var(--color-surface-muted); }

    /* ── Skeleton loader ──────────────────────────────────────────────── */
    .skeleton-cell {
      display: inline-block; height: .75rem;
      border-radius: var(--radius-sm);
      background: linear-gradient(90deg,
        var(--color-border) 25%,
        var(--color-surface-muted) 50%,
        var(--color-border) 75%);
      background-size: 200% 100%;
      animation: shimmer 1.4s infinite;
    }
    @keyframes shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }

    /* ── Rodapé ───────────────────────────────────────────────────────── */
    .table-footer {
      display: flex; align-items: center; justify-content: space-between;
      padding: .7rem 1.25rem;
      border-top: 1px solid var(--color-border);
      flex-wrap: wrap; gap: .5rem;
    }
    .table-footer-left   { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
    .table-footer-center { flex: 1; display: flex; justify-content: center; }
    .table-footer-right  { min-width: 140px; text-align: right; padding-right: .25rem; }
    .info-registros { font-size: .8rem; color: var(--color-text-muted); }
    .select-limite  { width: auto; }

    /* ── Paginação ────────────────────────────────────────────────────── */
    #paginacao .page-link {
      color: var(--color-text-muted);
      border-color: var(--color-border);
      font-size: .82rem;
      transition: background-color var(--transition), border-color var(--transition), color var(--transition);
    }
    #paginacao .page-item.active .page-link    { background-color: var(--color-accent); border-color: var(--color-accent); color: var(--color-surface); }
    #paginacao .page-link:hover                { background-color: var(--color-surface-muted); border-color: var(--color-border-strong); color: var(--color-text); }
    #paginacao .page-item.disabled .page-link  { color: var(--color-text-faint); }

    /* ── Overlay de carregamento ──────────────────────────────────────── */
    .overlay-loading {
      position: absolute; inset: 0;
      background: hsla(0, 0%, 100%, .82);
      backdrop-filter: blur(2px);
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      z-index: 10; gap: .85rem;
      border-radius: var(--radius-md);
    }
    .overlay-loading .spinner-border { width: 1.75rem; height: 1.75rem; border-width: 2px; color: var(--color-text-faint); }
    .overlay-loading span            { font-size: .85rem; font-weight: 500; color: var(--color-text-muted); }

    /* ── Toast ────────────────────────────────────────────────────────── */
    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: 1100;
      display: flex; flex-direction: column; gap: .5rem;
    }
    .toast-custom {
      min-width: 280px; max-width: 360px;
      border-radius: var(--radius-md); overflow: hidden;
      box-shadow: var(--shadow-md);
      background: var(--color-surface); border: 1px solid var(--color-border);
      animation: toastSlideIn .22s ease;
    }
    @keyframes toastSlideIn { from { opacity: 0; transform: translateX(24px); } to { opacity: 1; transform: translateX(0); } }
    .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

    /* ── Modal SPA ────────────────────────────────────────────────────── */
    #modalSPA .modal-dialog { max-width: min(860px, 94vw); }
    #modalSPA .modal-content {
      border: 1px solid var(--color-border);
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-lg); overflow: hidden;
    }
    #modalSPA .modal-header {
      background-color: var(--color-surface-muted);
      border-bottom: 1px solid var(--color-border);
      padding: .85rem 1.25rem;
    }
    #modalSPA .modal-title { font-size: .95rem; font-weight: 600; color: var(--color-text); }
    #modalSPA .modal-body  { padding: 0; position: relative; }

    #iframeModal {
      width: 100%; height: 75vh;
      border: none; display: block;
      background-color: var(--color-surface);
    }
    #modalLoadingOverlay {
      position: absolute; inset: 0;
      background: hsla(0, 0%, 100%, .92);
      display: flex; align-items: center; justify-content: center;
      z-index: 5;
    }

    /* ── Acessibilidade ───────────────────────────────────────────────── */
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }
    }
    @media (max-width: 768px) {
      .table-footer { flex-direction: column; align-items: stretch; }
      .table-footer-right { text-align: center; padding-right: 0; }
      .table-footer-center { justify-content: center; }
      #modalSPA .modal-dialog { max-width: 98vw; margin: .5rem auto; }
      #iframeModal { height: 85vh; }
    }
  </style>
</head>
<body>

  <%
  ' Toast vindo de redirecionamento do form (ex.: ?strStatus=ALT)
  If Trim(strMsg) <> "" Then
    Response.Write "<script>document.addEventListener('DOMContentLoaded',function(){ValidadoresGrid.exibirToast('sucesso','" & strMsg & "');});<" & "/script>"
  End If
  %>

  <!-- Toast container -->
  <div class="toast-container-custom" aria-live="polite"></div>

  <!-- ══════════════════════════════════════════════════════════════════ -->
  <!--  Modal SPA — form_validadoresdecredito.asp carregado via iframe   -->
  <!-- ══════════════════════════════════════════════════════════════════ -->
  <div class="modal fade" id="modalSPA" tabindex="-1" aria-labelledby="modalSPALabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable modal-xl">
      <div class="modal-content">
        <div class="modal-header">
          <span class="modal-title" id="modalSPALabel">Validador de Crédito</span>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <div id="modalLoadingOverlay">
            <div class="spinner-border"
                 style="width:1.75rem;height:1.75rem;border-width:2px;color:var(--color-text-faint);"
                 role="status">
              <span class="visually-hidden">Carregando...</span>
            </div>
          </div>
          <iframe id="iframeModal" src="about:blank" title="Formulário de Validador"></iframe>
        </div>
      </div>
    </div>
  </div>

  <!-- ── Container principal ─────────────────────────────────────────── -->
  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="d-flex align-items-center mb-3">
      <span class="page-title">Validadores de Crédito</span>
    </div>

    <div class="table-card">

      <!-- Overlay de loading do grid -->
      <div id="overlayLoading" class="overlay-loading d-none" aria-live="polite">
        <div class="spinner-border" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <!-- Toolbar -->
      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom flex-wrap gap-2">
        <div class="d-flex gap-2 align-items-center">
          <button type="button" class="btn btn-novo btn-sm"
                  title="Inserir novo validador"
                  onclick="ValidadoresGrid.abrirModal('form_validadoresdecredito.asp?acao=INS&vldcod=', 'Novo Validador')">
            <i class="bi bi-plus-lg me-1"></i>Novo
          </button>

        </div>
        <div class="search-box">
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <!-- Tabela -->
      <div class="table-wrapper">
        <table id="tabelaDados" class="table table-bordered table-hover table-sm mb-0">
          <thead>
            <tr>
              <th style="width:96px;">Ações</th>
              <th>Código</th>
              <th>Nome</th>
              <th>Plantas</th>
              <th>Nível</th>
              <th>Mínimo</th>
              <th>Máximo</th>
            </tr>
          </thead>
          <tbody id="corpoTabela">
            <!-- populado via AJAX → ?api=1 -->
          </tbody>
        </table>
      </div>

      <!-- Rodapé -->
      <div class="table-footer">
        <div class="table-footer-left">
          <label for="selectLimite" class="form-label mb-0"
                 style="font-size:.8rem;white-space:nowrap;color:var(--color-text-muted);">
            Linhas por página:
          </label>
          <select id="selectLimite" class="form-select form-select-sm select-limite">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50" selected>50</option>
            <option value="100">100</option>
            <option value="250">250</option>
          </select>
        </div>
        <div class="table-footer-center">
          <nav aria-label="Navegação de páginas">
            <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
          </nav>
        </div>
        <div class="table-footer-right">
          <span id="infoRegistros" class="info-registros" aria-live="polite"></span>
        </div>
      </div>

    </div><!-- /table-card -->
  </div><!-- /container-fluid -->

  <!-- jQuery -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <script>
  /* ================================================================
     ValidadoresGrid — IIFE principal
     Fluxo: carregarDados() → ?api=1 (JSON) → renderizar() → paginar()
     Cache em sessionStorage com TTL de 5 min.
     O modal não dispara window.location.reload — apenas recarrega o JSON.
  ================================================================ */
  var ValidadoresGrid = (function () {

    /* ── Constantes ─────────────────────────────────────────────── */
    var CACHE_KEY    = 'validadoresCredito_cache';
    var CACHE_TTL_MS = 5 * 60 * 1000;   // 5 minutos
    var DEBOUNCE_MS  = 280;
    var TOAST_DELAY  = 4500;
    // O próprio arquivo serve o JSON quando chamado com ?api=1
    var API_URL      = 'grid_validadoresdecredito.asp?api=1';

    /* ── Mapas de dados (JS) ────────────────────────────────────── */
    var PLANTAS_MAP = {
      '00': { sigla: 'PIR',  desc: 'Pirapetinga',     bg: 'hsl(338,70%,92%)', fg: 'hsl(338,60%,38%)' },
      '01': { sigla: 'RIO',  desc: 'Rio de Janeiro',  bg: 'hsl(210,80%,91%)', fg: 'hsl(210,70%,35%)' },
      '02': { sigla: 'SPO',  desc: 'São Paulo',       bg: 'hsl(262,60%,91%)', fg: 'hsl(262,50%,38%)' },
      '03': { sigla: 'COR',  desc: 'Corporativo',     bg: 'hsl(180,55%,88%)', fg: 'hsl(180,50%,28%)' },
      '04': { sigla: 'UBE',  desc: 'Uberaba',         bg: 'hsl(28,80%,90%)',  fg: 'hsl(28,70%,35%)'  },
      '05': { sigla: 'BEN',  desc: 'Bento Gonçalves', bg: 'hsl(145,55%,88%)', fg: 'hsl(145,55%,28%)' },
      '06': { sigla: 'FOR',  desc: 'Fortaleza',       bg: 'hsl(50,80%,88%)',  fg: 'hsl(45,70%,30%)'  },
      '07': { sigla: 'SQA',  desc: 'Saquarema',       bg: 'hsl(190,65%,88%)', fg: 'hsl(190,60%,30%)' },
      '08': { sigla: 'ARA',  desc: 'Araçatuba',       bg: 'hsl(0,70%,91%)',   fg: 'hsl(0,60%,38%)'   },
      '09': { sigla: 'BLU',  desc: 'Blumenau',        bg: 'hsl(300,45%,91%)', fg: 'hsl(300,40%,38%)' },
      '10': { sigla: 'PAC',  desc: 'Pacajus',         bg: 'hsl(160,55%,88%)', fg: 'hsl(160,55%,28%)' },
      '11': { sigla: 'POR',  desc: 'Porto Feliz',     bg: 'hsl(35,65%,90%)',  fg: 'hsl(35,60%,32%)'  },
      '12': { sigla: 'TRES', desc: 'Três Barras',     bg: 'hsl(220,65%,91%)', fg: 'hsl(220,55%,36%)' },
      '13': { sigla: 'FLO',  desc: 'Florestal',       bg: 'hsl(90,55%,88%)',  fg: 'hsl(90,50%,28%)'  }
    };

    var NIVEL_MAP = {
      10: { bg: 'hsl(210,80%,91%)', fg: 'hsl(210,70%,35%)' },
      20: { bg: 'hsl(262,60%,91%)', fg: 'hsl(262,50%,38%)' },
      30: { bg: 'hsl(145,55%,88%)', fg: 'hsl(145,55%,28%)' },
      40: { bg: 'hsl(28,80%,90%)',  fg: 'hsl(28,70%,35%)'  },
      50: { bg: 'hsl(338,70%,92%)', fg: 'hsl(338,60%,38%)' },
      60: { bg: 'hsl(0,70%,91%)',   fg: 'hsl(0,60%,38%)'   },
      70: { bg: 'hsl(50,80%,88%)',  fg: 'hsl(45,70%,30%)'  }
    };
    var _defaultCor = { bg: 'hsl(240,6%,90%)', fg: 'hsl(240,6%,36%)' };

    /* ── Estado interno ─────────────────────────────────────────── */
    var _todos       = [];
    var _filtrados   = [];
    var _pagAtual    = 1;
    var _itensPorPag = 50;
    var _debounce;

    /* ── Referências DOM ────────────────────────────────────────── */
    var $corpo         = $('#corpoTabela');
    var $paginacao     = $('#paginacao');
    var $infoRegistros = $('#infoRegistros');
    var $overlay       = $('#overlayLoading');
    var $searchBox     = $('#searchBox');
    var $selectLimite  = $('#selectLimite');

    /* ── Modal SPA ──────────────────────────────────────────────── */
    var _modalEl    = document.getElementById('modalSPA');
    var _bsModal    = new bootstrap.Modal(_modalEl);
    var _iframe     = document.getElementById('iframeModal');
    var _modalTitle = document.getElementById('modalSPALabel');
    var _mOverlay   = document.getElementById('modalLoadingOverlay');

    /* ── Helpers de badge ───────────────────────────────────────── */
    function _badgePlantas(arr) {
      if (!arr || !arr.length) return '—';
      return arr.map(function (cod) {
        var p = PLANTAS_MAP[String(cod).trim()];
        if (!p) return '';
        return '<span class="badge-planta" title="' + p.desc + '" ' +
               'style="background-color:' + p.bg + ';color:' + p.fg + ';">' +
               p.sigla + '</span>';
      }).join('');
    }

    function _badgeNivel(nivCod, nivDsc) {
      var n = NIVEL_MAP[nivCod] || _defaultCor;
      return '<span class="badge-nivel" style="background-color:' + n.bg + ';color:' + n.fg + ';">' +
             _esc(nivDsc) + '</span>';
    }

    function _moeda(v) {
      return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    }

    function _esc(s) {
      return String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /* ── Skeleton loader ────────────────────────────────────────── */
    function _mostrarSkeleton() {
      var html = '';
      for (var i = 0; i < 8; i++) {
        html += '<tr>' +
          '<td><span class="skeleton-cell" style="width:80px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:40px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:140px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:80px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:110px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:70px;"></span></td>' +
          '<td><span class="skeleton-cell" style="width:70px;"></span></td>' +
          '</tr>';
      }
      $corpo.html(html);
    }

    function _mostrarVazio(msg) {
      $corpo.html(
        '<tr><td colspan="7">' +
        '<div class="d-flex flex-column align-items-center justify-content-center py-5 gap-2">' +
        '<i class="bi bi-clipboard-data" style="font-size:2.4rem;color:var(--color-text-faint);"></i>' +
        '<p style="color:var(--color-text-muted);font-size:.88rem;margin:0;">' +
        (msg || 'Nenhum validador de crédito encontrado.') + '</p>' +
        '</div></td></tr>'
      );
      $paginacao.empty();
      $infoRegistros.text('');
    }

    /* ── Cache sessionStorage ───────────────────────────────────── */
    function _lerCache() {
      try {
        var raw = sessionStorage.getItem(CACHE_KEY);
        if (!raw) return null;
        var obj = JSON.parse(raw);
        if (Date.now() - obj.ts > CACHE_TTL_MS) { sessionStorage.removeItem(CACHE_KEY); return null; }
        return obj.data;
      } catch (e) { return null; }
    }

    function _gravarCache(data) {
      try { sessionStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), data: data })); } catch (e) {}
    }

    function _invalidarCache() {
      try { sessionStorage.removeItem(CACHE_KEY); } catch (e) {}
    }

    /* ── Carregar dados via AJAX ─────────────────────────────────── */
    function carregarDados(forcar) {
      if (!forcar) {
        var cached = _lerCache();
        if (cached) { _todos = cached; _filtrados = _todos.slice(); _renderPagina(1); return; }
      }

      _mostrarSkeleton();
      $overlay.removeClass('d-none');

      $.ajax({
        url:      API_URL,
        type:     'GET',
        dataType: 'json',
        cache:    false,
        success: function (resp) {
          if (!resp.ok) {
            exibirToast('erro', 'Erro ao carregar dados: ' + (resp.erro || 'desconhecido'));
            _mostrarVazio('Erro ao carregar dados.');
            return;
          }
          _todos     = resp.registros || [];
          _filtrados = _todos.slice();
          _gravarCache(_todos);

          var termo = $searchBox.val().toLowerCase().trim();
          if (termo) _aplicarFiltro(termo);

          _renderPagina(1);
        },
        error: function (xhr) {
          exibirToast('erro', 'Falha na requisição (' + xhr.status + ').');
          _mostrarVazio('Não foi possível carregar os dados.');
        },
        complete: function () {
          $overlay.addClass('d-none');
        }
      });
    }

    /* ── Renderizar página atual ─────────────────────────────────── */
    function _renderPagina(pagina) {
      _pagAtual = pagina;
      var total  = _filtrados.length;
      var totPag = Math.ceil(total / _itensPorPag) || 1;

      if (!total) { _mostrarVazio(); return; }

      var inicio = (_pagAtual - 1) * _itensPorPag;
      var fatia  = _filtrados.slice(inicio, inicio + _itensPorPag);

      var html = '';
      for (var i = 0; i < fatia.length; i++) {
        var r = fatia[i];
        html += '<tr>' +
          '<td><div class="d-flex gap-1 justify-content-center">' +
            '<button type="button" class="btn btn-row-view btn-sm" title="Visualizar" ' +
                    'onclick="ValidadoresGrid.abrirModal(\'form_validadoresdecredito.asp?acao=CON&vldcod=' + r.cod + '\',\'Visualizar Validador\')">' +
              '<i class="bi bi-eye"></i></button>' +
            '<button type="button" class="btn btn-row-edit btn-sm" title="Editar" ' +
                    'onclick="ValidadoresGrid.abrirModal(\'form_validadoresdecredito.asp?acao=ALT&vldcod=' + r.cod + '\',\'Editar Validador\')">' +
              '<i class="bi bi-pencil"></i></button>' +
            '<button type="button" class="btn btn-row-del btn-sm" title="Apagar" ' +
                    'onclick="ValidadoresGrid.abrirModal(\'form_validadoresdecredito.asp?acao=DEL&vldcod=' + r.cod + '\',\'Apagar Validador\')">' +
              '<i class="bi bi-trash3"></i></button>' +
          '</div></td>' +
          '<td>' + _esc(r.cod)                    + '</td>' +
          '<td>' + _esc(r.nom)                    + '</td>' +
          '<td>' + _badgePlantas(r.plantas)        + '</td>' +
          '<td>' + _badgeNivel(r.nivCod, r.nivDsc) + '</td>' +
          '<td>' + _moeda(r.valMin)               + '</td>' +
          '<td>' + _moeda(r.valMax)               + '</td>' +
          '</tr>';
      }
      $corpo.html(html);

      _renderPaginacao(_pagAtual, totPag);
      $infoRegistros.text('Total: ' + total + ' registro(s)');
    }

    /* ── Paginação ───────────────────────────────────────────────── */
    function _renderPaginacao(pagAtual, totalPaginas) {
      $paginacao.empty();
      if (totalPaginas <= 1) return;

      function mkItem(label, page, disabled, active) {
        var $li = $('<li>').addClass('page-item');
        if (disabled) $li.addClass('disabled');
        if (active)   $li.addClass('active');
        var $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
        if (!disabled && !active) {
          $a.on('click', function (e) { e.preventDefault(); _renderPagina(page); });
        }
        return $li.append($a);
      }

      $paginacao.append(mkItem('&laquo;', 1, pagAtual === 1, false));
      $paginacao.append(mkItem('&lsaquo;', pagAtual - 1, pagAtual === 1, false));

      var ini = Math.max(1, pagAtual - 2);
      var fim = Math.min(totalPaginas, pagAtual + 2);
      for (var p = ini; p <= fim; p++) {
        $paginacao.append(mkItem(p, p, false, p === pagAtual));
      }

      $paginacao.append(mkItem('&rsaquo;', pagAtual + 1, pagAtual === totalPaginas, false));
      $paginacao.append(mkItem('&raquo;', totalPaginas, pagAtual === totalPaginas, false));
    }

    /* ── Filtro de busca ─────────────────────────────────────────── */
    function _aplicarFiltro(termo) {
      if (!termo) { _filtrados = _todos.slice(); return; }
      _filtrados = _todos.filter(function (r) {
        return [r.cod, r.nom, r.nivDsc, r.valMin, r.valMax]
          .concat(r.plantas || []).join(' ').toLowerCase().indexOf(termo) !== -1;
      });
    }

    /* ── Modal SPA ───────────────────────────────────────────────── */
    function abrirModal(url, titulo) {
      _modalTitle.textContent = titulo || 'Validador de Crédito';
      _mOverlay.style.display = 'flex';
      _iframe.src             = 'about:blank';
      _bsModal.show();

      _modalEl.addEventListener('shown.bs.modal', function carregar() {
        _iframe.src = url;
        _modalEl.removeEventListener('shown.bs.modal', carregar);
      });
    }

    /* ── Toast ───────────────────────────────────────────────────── */
    function exibirToast(tipo, mensagem) {
      var iconMap = {
        sucesso: { icon: 'bi-check-circle-fill',          cor: 'var(--color-success)'     },
        erro:    { icon: 'bi-x-circle-fill',              cor: 'var(--color-danger)'       },
        aviso:   { icon: 'bi-exclamation-triangle-fill',  cor: 'var(--color-warning)'      },
        info:    { icon: 'bi-info-circle-fill',           cor: 'var(--color-text-muted)'   }
      };
      var cfg = iconMap[tipo] || iconMap.info;
      var $el = $(
        '<div class="toast-custom">' +
          '<div style="height:3px;background:' + cfg.cor + ';" class="toast-progress-bar"></div>' +
          '<div class="d-flex align-items-start gap-2 p-3">' +
            '<i class="bi ' + cfg.icon + ' mt-1" style="color:' + cfg.cor + ';font-size:1rem;flex-shrink:0;"></i>' +
            '<span style="font-size:.85rem;color:var(--color-text);flex:1;">' + mensagem + '</span>' +
            '<button type="button" style="background:none;border:none;cursor:pointer;color:var(--color-text-faint);font-size:1rem;line-height:1;padding:0;" ' +
                    'onclick="this.closest(\'.toast-custom\').remove()">&times;</button>' +
          '</div>' +
        '</div>'
      );
      $('.toast-container-custom').append($el);
      setTimeout(function () { $el.remove(); }, TOAST_DELAY);
    }

    /* ── Eventos ─────────────────────────────────────────────────── */
    function _bindEventos() {

      $searchBox.on('input', function () {
        clearTimeout(_debounce);
        var termo = $(this).val().toLowerCase().trim();
        _debounce = setTimeout(function () {
          _aplicarFiltro(termo);
          _renderPagina(1);
        }, DEBOUNCE_MS);
      });

      document.addEventListener('keydown', function (e) {
        if (e.ctrlKey && e.key === 'k') { e.preventDefault(); $searchBox.trigger('focus'); }
      });

      $selectLimite.on('change', function () {
        _itensPorPag = parseInt($(this).val(), 10);
        _renderPagina(1);
      });

      // Overlay some quando o iframe termina de carregar (qualquer carga)
      _iframe.addEventListener('load', function () {
        _mOverlay.style.display = 'none';
      });

      // Escuta postMessage enviado pelo form após AJAX com sucesso
      window.addEventListener('message', function (e) {
        if (e.origin !== window.location.origin) return;            // segurança: só mesma origem
        var d = e.data;
        if (!d || d.tipo !== 'vldcredito:ok') return;

        _bsModal.hide();
        _invalidarCache();
        carregarDados(true);
        setTimeout(function () {
          exibirToast('sucesso', d.mensagem || 'Operação realizada com sucesso!');
        }, 300);
      });

      _modalEl.addEventListener('hidden.bs.modal', function () {
        _iframe.src             = 'about:blank';
        _mOverlay.style.display = 'flex';
      });
    }

    /* ── Init ────────────────────────────────────────────────────── */
    function init() {
      _itensPorPag = parseInt($selectLimite.val(), 10);
      _bindEventos();
      carregarDados(false);
    }

    /* ── API pública ─────────────────────────────────────────────── */
    return Object.freeze({
      init:        init,
      abrirModal:  abrirModal,
      exibirToast: exibirToast,
      recarregar:  function () { _invalidarCache(); carregarDados(true); }
    });

  }());

  $(document).ready(function () {
    ValidadoresGrid.init();
  });
  </script>

</body>
</html>
