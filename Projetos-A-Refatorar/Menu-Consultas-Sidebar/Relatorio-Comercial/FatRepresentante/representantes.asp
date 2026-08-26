<%RESPONSE.BUFFER=TRUE%>
<!--#include virtual="/corporativo/cnn2/cnn2.asp" -->
<!--#include virtual="/corporativo/includes/variaveis.asp" -->
<%
'*******************************
'***    Validando Acesso     ***
'*******************************

If CInt(NIVEL) <> 1 Then
%>
<script>
  alert("Acesso Negado.\nVoce nao tem permissao para esse modulo.");
  window.close();
</script>
<%
  Response.End
End If

'Calculando mes e ano atual
Dim MES, ANO
MES = MONTH(NOW)
if trim(request("select_mes")) <> "" then
  ANO = trim(request("select_ano"))
else
  ANO = YEAR(NOW)
end if

if MES < 10 then
  MES = "0" & MES
end if

'Montando lista de anos dinamicamente (2010 ate ano atual + 5 anos futuros)
Dim anoInicio, anoFim
anoInicio = 2010
anoFim    = YEAR(NOW) + 5
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Faturamento de Representantes</title>

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    /* ============================================================
       TOKENS DE DESIGN — SWnet Quiet UI
       ============================================================ */
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
      --color-danger:        hsl(3, 90%, 58%);
      --color-danger-soft:   hsl(3, 100%, 96%);
      --color-warning:       hsl(35, 100%, 48%);
      --color-warning-soft:  hsl(40, 100%, 95%);

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
      --space-6: 32px;
      --space-7: 48px;

      --z-toast: 1100;
    }

    /* ============================================================
       BASE
       ============================================================ */
    body {
      background-color: var(--color-bg);
      min-height: 100vh;
      font-family: var(--font-sans);
      color: var(--color-text);
    }

    /* ============================================================
       TÍTULO DA PÁGINA
       ============================================================ */
    .page-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
      font-weight: 600;
      letter-spacing: -0.01em;
      color: var(--color-text);
    }

    /* ============================================================
       CARD DE FILTROS
       ============================================================ */
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

    .filter-label {
      font-size: .78rem;
      font-weight: 500;
      color: var(--color-text-muted);
      letter-spacing: .01em;
      margin-bottom: .3rem;
    }

    /* ============================================================
       INPUTS — foco do design system
       ============================================================ */
    .form-control:focus,
    .form-select:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
    }

    /* ============================================================
       BOTÕES
       ============================================================ */
    .btn-primary {
      background-color: var(--color-accent);
      border-color: var(--color-accent);
      font-weight: 500;
      font-size: .85rem;
      transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-primary:hover {
      background-color: var(--color-accent-hover);
      border-color: var(--color-accent-hover);
    }
    .btn-primary:active { transform: scale(.98); }
    .btn-primary:focus-visible {
      box-shadow: 0 0 0 3px var(--color-accent-soft);
    }

    .btn-light {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      color: var(--color-text);
      font-weight: 500;
      font-size: .85rem;
      transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-light:hover {
      background-color: var(--color-surface-muted);
      border-color: var(--color-border-strong);
    }
    .btn-light:active { transform: scale(.98); }

    /* ============================================================
       CARD DE RESULTADO (área futura da tabela)
       ============================================================ */
    .result-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-sm);
    }

    .result-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 4rem 1rem;
      color: var(--color-text-muted);
      gap: .85rem;
    }
    .result-placeholder .bi {
      font-size: 2.4rem;
      color: var(--color-text-faint);
    }
    .result-placeholder p {
      margin: 0;
      font-size: .88rem;
      text-align: center;
    }

    /* ============================================================
       TOAST
       ============================================================ */
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
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    /* ============================================================
       ACESSIBILIDADE
       ============================================================ */
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.01ms !important;
        transition-duration: 0.01ms !important;
      }
    }

    /* ============================================================
       RESPONSIVIDADE
       ============================================================ */
    @media (max-width: 768px) {
      .filter-actions {
        justify-content: stretch !important;
        flex-direction: column;
      }
      .filter-actions .btn {
        width: 100%;
      }
    }
  </style>
</head>

<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho da página -->
    <div class="d-flex align-items-center mb-4 gap-3">
      <span class="page-title">Faturamento de Representantes</span>
    </div>

    <!-- Card de Filtros -->
    <div class="filter-card mb-4">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2" style="color: var(--color-text-muted);"></i>Filtros
      </div>
      <div class="filter-card-body">
        <div class="row g-3 align-items-end">

          <!-- Mês -->
          <div class="col-12 col-sm-auto">
            <label for="selectMes" class="form-label filter-label">Mês</label>
            <select id="selectMes" name="select_mes" class="form-select form-select-sm" style="width: 100px;">
              <option value="01" <%if MES = "01" then%>selected<%end if%>>01 — Jan</option>
              <option value="02" <%if MES = "02" then%>selected<%end if%>>02 — Fev</option>
              <option value="03" <%if MES = "03" then%>selected<%end if%>>03 — Mar</option>
              <option value="04" <%if MES = "04" then%>selected<%end if%>>04 — Abr</option>
              <option value="05" <%if MES = "05" then%>selected<%end if%>>05 — Mai</option>
              <option value="06" <%if MES = "06" then%>selected<%end if%>>06 — Jun</option>
              <option value="07" <%if MES = "07" then%>selected<%end if%>>07 — Jul</option>
              <option value="08" <%if MES = "08" then%>selected<%end if%>>08 — Ago</option>
              <option value="09" <%if MES = "09" then%>selected<%end if%>>09 — Set</option>
              <option value="10" <%if MES = "10" then%>selected<%end if%>>10 — Out</option>
              <option value="11" <%if MES = "11" then%>selected<%end if%>>11 — Nov</option>
              <option value="12" <%if MES = "12" then%>selected<%end if%>>12 — Dez</option>
            </select>
          </div>

          <!-- Ano -->
          <div class="col-12 col-sm-auto">
            <label for="selectAno" class="form-label filter-label">Ano</label>
            <select id="selectAno" name="select_ano" class="form-select form-select-sm" style="width: 100px;">
              <%
              Dim a
              For a = anoFim To anoInicio Step -1
              %>
              <option value="<%=a%>" <%if CStr(ANO) = CStr(a) then%>selected<%end if%>><%=a%></option>
              <%
              Next
              %>
            </select>
          </div>

          <!-- Ações -->
          <div class="col-12 col-sm-auto">
            <div class="d-flex gap-2 filter-actions">
              <button type="button" id="btnFiltrar" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" id="btnLimpar" class="btn btn-light btn-sm">
                <i class="bi bi-eraser me-1"></i>Limpar
              </button>
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- Card de Resultado -->
    <div class="result-card" id="resultCard">
      <div class="result-placeholder" id="resultPlaceholder">
        <i class="bi bi-people"></i>
        <p>Selecione o mês e o ano e clique em <strong>Filtrar</strong> para exibir o faturamento.</p>
      </div>
    </div>

  </div>

  <!-- Toast container -->
  <div class="toast-container-custom" aria-live="polite" id="toastContainer"></div>

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <!-- jQuery -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
    'use strict';

    const RepresentantesConsulta = (function () {

      /* ── Constantes ─────────────────────────────────────────── */
      const REPOSITORIO  = 'fat_rep_cli_acumulado.asp';
      const TOAST_DELAY  = 4500;
      const DEBOUNCE_MS  = 280;

      const TOAST_ICON = {
        sucesso: { icon: 'bi-check-circle-fill',          cor: 'var(--color-success)' },
        erro:    { icon: 'bi-x-circle-fill',              cor: 'var(--color-danger)'  },
        aviso:   { icon: 'bi-exclamation-triangle-fill',  cor: 'var(--color-warning)' },
        info:    { icon: 'bi-info-circle-fill',           cor: 'var(--color-text-muted)' }
      };

      /* ── Helpers ─────────────────────────────────────────────── */
      function $el(id) { return document.getElementById(id); }

      function exibirToast(tipo, mensagem) {
        const cfg  = TOAST_ICON[tipo] || TOAST_ICON.info;
        const id   = 'toast_' + Date.now();
        const html = `
          <div id="${id}" class="toast-custom">
            <div class="d-flex align-items-center px-3 py-2 border-bottom gap-2">
              <i class="bi ${cfg.icon}" style="color:${cfg.cor};font-size:1rem;"></i>
              <span style="font-size:.85rem;font-weight:600;color:${cfg.cor};flex:1;">${tipo.charAt(0).toUpperCase() + tipo.slice(1)}</span>
              <button type="button" style="background:none;border:none;font-size:1.1rem;cursor:pointer;color:var(--color-text-muted);"
                onclick="this.closest('.toast-custom').remove()">&#215;</button>
            </div>
            <div class="px-3 py-2" style="font-size:.84rem;color:var(--color-text);">${mensagem}</div>
            <div style="height:3px;background:var(--color-border);">
              <div class="toast-progress-bar" style="height:100%;background:${cfg.cor};"></div>
            </div>
          </div>`;
        $('#toastContainer').append(html);
        setTimeout(() => { const el = $el(id); if (el) el.remove(); }, TOAST_DELAY + 250);
      }

      /* ── Lógica principal ────────────────────────────────────── */
      /*
       * fat_rep_cli_acumulado.asp é legado e retorna HTML puro, não JSON.
       * O filtro navega diretamente via GET, preservando o comportamento original.
       * Quando esse destino for refatorado para JSON, substituir por $.ajax
       * com dataType:'json' e renderizar a tabela inline no padrão SWnet.
       */
      function _filtrar() {
        const mes = $('#selectMes').val();
        const ano = $('#selectAno').val();

        if (!mes || !ano) {
          exibirToast('aviso', 'Selecione o mês e o ano antes de filtrar.');
          return;
        }

        window.location.href = REPOSITORIO
          + '?select_mes=' + encodeURIComponent(mes)
          + '&select_ano=' + encodeURIComponent(ano);
      }

      function _limpar() {
        const hoje = new Date();
        const mesAtual = String(hoje.getMonth() + 1).padStart(2, '0');
        const anoAtual = String(hoje.getFullYear());
        $('#selectMes').val(mesAtual);
        $('#selectAno').val(anoAtual);
        $('#resultCard').html(`
          <div class="result-placeholder" id="resultPlaceholder">
            <i class="bi bi-people"></i>
            <p>Selecione o mês e o ano e clique em <strong>Filtrar</strong> para exibir o faturamento.</p>
          </div>`);
      }

      /* ── Inicialização ───────────────────────────────────────── */
      function init() {
        $('#btnFiltrar').on('click', _filtrar);
        $('#btnLimpar').on('click', _limpar);

        /* Atalho Enter nos selects */
        $('#selectMes, #selectAno').on('keydown', function (e) {
          if (e.key === 'Enter') _filtrar();
        });
      }

      return Object.freeze({ init });
    })();

    $(function () { RepresentantesConsulta.init(); });
  </script>
</body>
</html>
