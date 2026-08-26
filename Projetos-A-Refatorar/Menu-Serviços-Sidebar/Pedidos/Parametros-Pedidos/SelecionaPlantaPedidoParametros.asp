<!DOCTYPE html>
<!--#include file="../../includes/variaveis.asp" -->
<!--#Include File="../../funcoes/funcoes.asp"-->
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Parâmetros Pedidos</title>

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

    <style>
      :root {
        /* — Neutros — */
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

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

        --radius-sm: 6px;
        --radius-md: 10px;
        --radius-lg: 14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-2: 8px;
        --space-3: 12px;
        --space-4: 16px;
        --space-5: 24px;

        /* — Cores das plantas (letra escura / hover claro) — */
        --planta-pir:        hsl(210, 70%, 38%);
        --planta-pir-hover:  hsl(210, 90%, 96%);

        --planta-ube:        hsl(160, 55%, 30%);
        --planta-ube-hover:  hsl(160, 70%, 95%);

        --planta-ben:        hsl(280, 50%, 42%);
        --planta-ben-hover:  hsl(280, 70%, 96%);

        --planta-for:        hsl(25, 75%, 38%);
        --planta-for-hover:  hsl(25, 90%, 95%);

        --planta-saq:        hsl(0, 60%, 42%);
        --planta-saq-hover:  hsl(0, 75%, 96%);
      }

      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
          animation-duration: 0.01ms !important;
          transition-duration: 0.01ms !important;
        }
      }

      body {
        background-color: var(--color-bg);
        min-height: 100vh;
        font-family: var(--font-sans);
        color: var(--color-text);
      }

      .page-wrapper {
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
      }

      .planta-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-lg);
        box-shadow: var(--shadow-md);
        max-width: 420px;
        width: 100%;
      }

      .planta-card-header {
        padding: var(--space-4) var(--space-5);
        border-bottom: 1px solid var(--color-border);
        text-align: center;
      }

      .planta-card-body {
        padding: var(--space-3) var(--space-3) var(--space-4);
      }

      .planta-item {
        display: flex;
        align-items: center;
        gap: var(--space-3);
        padding: var(--space-3) var(--space-3);
        border-radius: var(--radius-md);
        border: 1px solid var(--color-border);
        margin-bottom: var(--space-2);
        cursor: pointer;
        transition: background-color var(--transition), border-color var(--transition);
        position: relative;
      }

      .planta-item:last-child {
        margin-bottom: 0;
      }

      .planta-item input[type="radio"] {
        position: absolute;
        opacity: 0;
        pointer-events: none;
      }

      .planta-letra {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 38px;
        height: 38px;
        min-width: 38px;
        border-radius: var(--radius-sm);
        font-weight: 600;
        font-size: 1rem;
        transition: background-color var(--transition);
      }

      .planta-nome {
        font-size: .92rem;
        font-weight: 500;
      }

      .planta-item:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }

      /* Pirapetinga */
      .planta-item[data-planta="00"] { border-color: hsla(210, 70%, 38%, .35); }
      .planta-item[data-planta="00"] .planta-letra,
      .planta-item[data-planta="00"] .planta-nome { color: var(--planta-pir) !important; }
      .planta-item[data-planta="00"] .planta-letra { background-color: hsla(210, 70%, 38%, .1); }
      .planta-item[data-planta="00"]:hover,
      .planta-item[data-planta="00"]:focus-visible { background-color: var(--planta-pir-hover); border-color: var(--planta-pir); }

      /* Uberaba */
      .planta-item[data-planta="04"] { border-color: hsla(160, 55%, 30%, .35); }
      .planta-item[data-planta="04"] .planta-letra,
      .planta-item[data-planta="04"] .planta-nome { color: var(--planta-ube) !important; }
      .planta-item[data-planta="04"] .planta-letra { background-color: hsla(160, 55%, 30%, .1); }
      .planta-item[data-planta="04"]:hover,
      .planta-item[data-planta="04"]:focus-visible { background-color: var(--planta-ube-hover); border-color: var(--planta-ube); }

      /* Bento Gonçalves */
      .planta-item[data-planta="05"] { border-color: hsla(280, 50%, 42%, .35); }
      .planta-item[data-planta="05"] .planta-letra,
      .planta-item[data-planta="05"] .planta-nome { color: var(--planta-ben) !important; }
      .planta-item[data-planta="05"] .planta-letra { background-color: hsla(280, 50%, 42%, .1); }
      .planta-item[data-planta="05"]:hover,
      .planta-item[data-planta="05"]:focus-visible { background-color: var(--planta-ben-hover); border-color: var(--planta-ben); }

      /* Fortaleza */
      .planta-item[data-planta="06"] { border-color: hsla(25, 75%, 38%, .35); }
      .planta-item[data-planta="06"] .planta-letra,
      .planta-item[data-planta="06"] .planta-nome { color: var(--planta-for) !important; }
      .planta-item[data-planta="06"] .planta-letra { background-color: hsla(25, 75%, 38%, .1); }
      .planta-item[data-planta="06"]:hover,
      .planta-item[data-planta="06"]:focus-visible { background-color: var(--planta-for-hover); border-color: var(--planta-for); }

      /* Saquarema */
      .planta-item[data-planta="07"] { border-color: hsla(0, 60%, 42%, .35); }
      .planta-item[data-planta="07"] .planta-letra,
      .planta-item[data-planta="07"] .planta-nome { color: var(--planta-saq) !important; }
      .planta-item[data-planta="07"] .planta-letra { background-color: hsla(0, 60%, 42%, .1); }
      .planta-item[data-planta="07"]:hover,
      .planta-item[data-planta="07"]:focus-visible { background-color: var(--planta-saq-hover); border-color: var(--planta-saq); }
    </style>
  </head>
  <body>

    <div class="page-wrapper container-fluid px-3">
      <div class="planta-card">
        <div class="planta-card-header">
          <span class="page-title">Parâmetros Pedidos</span>
          <div class="mt-1" style="font-size:.8rem; color: var(--color-text-muted);">Selecione a planta</div>
        </div>

        <form name="form" method="post">
          <input type="hidden" name="Planta" id="Planta" />

          <div class="planta-card-body">
            <div class="planta-item" data-planta="00" tabindex="0" role="radio" aria-checked="false">
              <input type="radio" name="selecione" value="00" />
              <span class="planta-letra">P</span>
              <span class="planta-nome">Pirapetinga</span>
            </div>

            <div class="planta-item" data-planta="04" tabindex="0" role="radio" aria-checked="false">
              <input type="radio" name="selecione" value="04" />
              <span class="planta-letra">U</span>
              <span class="planta-nome">Uberaba</span>
            </div>

            <div class="planta-item" data-planta="05" tabindex="0" role="radio" aria-checked="false">
              <input type="radio" name="selecione" value="05" />
              <span class="planta-letra">B</span>
              <span class="planta-nome">Bento Gonçalves</span>
            </div>

            <div class="planta-item" data-planta="06" tabindex="0" role="radio" aria-checked="false">
              <input type="radio" name="selecione" value="06" />
              <span class="planta-letra">F</span>
              <span class="planta-nome">Fortaleza</span>
            </div>

            <div class="planta-item" data-planta="07" tabindex="0" role="radio" aria-checked="false">
              <input type="radio" name="selecione" value="07" />
              <span class="planta-letra">S</span>
              <span class="planta-nome">Saquarema</span>
            </div>
          </div>
        </form>
      </div>
    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <script>
      $(function() {
        const irParaPlanta = (valor) => {
          $('#Planta').val(valor);
          window.location.href = 'PedidosParametros.asp?Planta=' + valor;
        };

        $('.planta-item').on('click', function() {
          const valor = $(this).data('planta');
          $(this).find('input[type="radio"]').prop('checked', true);
          irParaPlanta(valor);
        });

        $('.planta-item').on('keydown', function(e) {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            $(this).trigger('click');
          }
        });
      });
    </script>

  </body>
</html>
