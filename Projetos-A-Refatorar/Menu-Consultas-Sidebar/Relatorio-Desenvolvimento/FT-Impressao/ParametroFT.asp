<!-- #include file="../includes/variaveis.asp" -->
<!-- #include file="../funcoes/funcoes.asp" -->
<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1046
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet - Ficha Técnica de Produto</title>

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
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

      --color-danger:        hsl(3, 90%, 58%);

      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

      --radius-sm:  6px;
      --radius-md:  10px;
      --radius-lg:  14px;

      --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
      --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

      --transition: 160ms ease;

      --space-1: 4px;
      --space-2: 8px;
      --space-3: 12px;
      --space-4: 16px;
      --space-5: 24px;
      --space-6: 32px;
    }

    /* Reset / base */
    *, *::before, *::after { box-sizing: border-box; }

    body {
      background-color: var(--color-bg);
      font-family: var(--font-sans);
      color: var(--color-text);
      min-height: 100vh;
    }

    /* Centering wrapper */
    .page-wrapper {
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      padding: var(--space-5) var(--space-4);
    }

    /* Card principal */
    .filter-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-md);
      width: 100%;
      max-width: 480px;
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

    /* Labels */
    .form-label {
      font-size: .78rem;
      font-weight: 500;
      color: var(--color-text-muted);
      letter-spacing: .01em;
      margin-bottom: .3rem;
    }

    /* Inputs */
    .form-control,
    .form-select {
      font-size: .85rem;
      border-color: var(--color-border);
      color: var(--color-text);
      background-color: var(--color-surface);
      border-radius: var(--radius-sm);
      transition: border-color var(--transition), box-shadow var(--transition);
    }
    .form-control::placeholder { color: var(--color-text-faint); }
    .form-control:focus,
    .form-select:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
      outline: none;
    }

    /* Botão primário */
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
    }
    .btn-primary:active { transform: scale(.98); }
    .btn-primary:disabled {
      opacity: .5;
      cursor: not-allowed;
      pointer-events: none;
    }

    /* Divider */
    .section-divider {
      border: none;
      border-top: 1px solid var(--color-border);
      margin: .85rem 0;
    }

    /* Título da página */
    .page-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
      font-weight: 600;
      letter-spacing: -0.01em;
      color: var(--color-text);
      margin: 0;
    }

    /* Acessibilidade: reduced motion */
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.01ms !important;
        transition-duration: 0.01ms !important;
      }
    }

    /* Responsividade abaixo de 576px */
    @media (max-width: 576px) {
      .filter-card { max-width: 100%; }
      .filter-card-body { padding: 1rem; }
    }
  </style>
</head>
<body>

  <div class="page-wrapper">
    <div class="filter-card">

      <!-- Cabeçalho -->
      <div class="filter-card-header d-flex align-items-center gap-2">
        <span class="page-title">Impressão de Ficha Técnica</span>
      </div>

      <!-- Corpo -->
      <div class="filter-card-body">
        <div class="row g-3">

          <!-- Filial -->
          <div class="col-12">
            <label for="filial" class="form-label">Filial</label>
            <select name="filial" id="filial" class="form-select form-select-sm">
              <option value=""   <% if filial = ""   then %>selected<% end if %>>Selecione...</option>
              <option value="00" <% if filial = "00" then %>selected<% end if %>>PIRAPETINGA</option>
              <option value="04" <% if filial = "04" then %>selected<% end if %>>UBERABA</option>
              <option value="05" <% if filial = "05" then %>selected<% end if %>>BENTO GONÇALVES</option>
              <option value="06" <% if filial = "06" then %>selected<% end if %>>MARANGUAPE</option>
              <option value="07" <% if filial = "07" then %>selected<% end if %>>SAQUAREMA</option>
            </select>
          </div>

          <!-- Número da FT -->
          <div class="col-12">
            <label for="ft" class="form-label">Número da FT</label>
            <input id="ft" name="ft" type="text"
                   class="form-control form-control-sm"
                   maxlength="20"
                   required
                   placeholder="Digite o número da FT..."
                   autocomplete="off" />
          </div>

        </div><!-- /row -->

        <br>

        <!-- Ação -->
        <div class="d-flex justify-content-end">
          <button type="button" id="btnImprimir"
                  class="btn btn-primary btn-sm"
                  title="Abrir ficha técnica para impressão">
            <i class="bi bi-printer-fill me-2"></i>Imprimir
          </button>
        </div>

      </div><!-- /filter-card-body -->
    </div><!-- /filter-card -->
  </div><!-- /page-wrapper -->

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <!-- jQuery -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
    $(function () {
      $('#btnImprimir').on('click', function () {
        const filial = $('#filial').val();
        const ft     = $('#ft').val().trim();

        if (!filial) {
          alert('Selecione uma filial.');
          $('#filial').focus();
          return;
        }
        if (!ft) {
          alert('Informe o número da FT.');
          $('#ft').focus();
          return;
        }

        // Abre FichaTecnica.asp em nova aba (comportamento original: target="_Blank")
        const url = 'FichaTecnica.asp?filial=' + encodeURIComponent(filial) + '&ft=' + encodeURIComponent(ft);
        window.open(url, '_blank');
      });

      // Limpa o número da FT ao trocar a filial
      $('#filial').on('change', function () {
        $('#ft').val('').focus();
      });

      // Atalho Enter no campo FT dispara impressão
      $('#ft').on('keydown', function (e) {
        if (e.key === 'Enter') $('#btnImprimir').trigger('click');
      });
    });
  </script>

</body>
</html>
