<!DOCTYPE html>
<html lang="pt-br">
  <head>
      <meta charset="utf-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <title>Selecione a Planta</title>

      <!-- Bootstrap CSS -->
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
            rel="stylesheet"
            integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
            crossorigin="anonymous" />

      <!--#include file="../../cnn/cnn.asp" -->
      <!--#include file="../../includes/variaveis.asp" -->
      <!--#include file="../../funcoes/funcoes.asp"-->
      <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>

      <style type="text/css">
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

              /* — Accent único — */
              --color-accent:        hsl(210, 100%, 45%);
              --color-accent-hover:  hsl(210, 100%, 40%);
              --color-accent-soft:   hsl(210, 100%, 96%);

              /* — Tipografia — */
              --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

              /* — Raio — */
              --radius-sm: 6px;
              --radius-md: 10px;
              --radius-lg: 14px;

              /* — Sombra — */
              --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
              --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

              /* — Transição — */
              --transition: 160ms ease;

              /* — Espaçamento — */
              --space-2: 8px;
              --space-4: 16px;
          }

          body {
              background-color: var(--color-bg);
              min-height: 100vh;
              font-family: var(--font-sans);
              color: var(--color-text);
          }

          .selecao-planta-wrapper {
              min-height: 100vh;
              display: flex;
              align-items: center;
              justify-content: center;
              padding: var(--space-4);
          }

          .selecao-planta-card {
              width: 100%;
              max-width: 420px;
              background-color: var(--color-surface);
              border: 1px solid var(--color-border);
              border-radius: var(--radius-lg);
              box-shadow: var(--shadow-md);
              overflow: hidden;
          }

          .selecao-planta-header {
              padding: 1.5rem 1.5rem 1rem;
              text-align: center;
              border-bottom: 1px solid var(--color-border);
          }

          .page-title {
              font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
              font-weight: 600;
              letter-spacing: -0.01em;
              color: var(--color-text);
          }

          .selecao-planta-subtitulo {
              font-size: .8rem;
              color: var(--color-text-muted);
              margin-top: .25rem;
          }

          .selecao-planta-lista {
              list-style: none;
              margin: 0;
              padding: 1rem 1.25rem 1.25rem;
              display: flex;
              flex-direction: column;
              gap: var(--space-2);
          }

          .opcao-planta {
              position: relative;
          }

          .opcao-planta label {
              display: flex;
              align-items: center;
              gap: .75rem;
              width: 100%;
              min-height: 44px;
              padding: .75rem 1rem;
              border: 1px solid var(--color-border);
              border-radius: var(--radius-md);
              font-size: .9rem;
              font-weight: 500;
              cursor: pointer;
              transition: background-color var(--transition), border-color var(--transition), color var(--transition);
          }

          .opcao-planta input[type="radio"] {
              width: 16px;
              height: 16px;
              accent-color: var(--color-accent);
              cursor: pointer;
              flex-shrink: 0;
          }

          .opcao-planta label:hover {
              background-color: var(--color-surface-muted);
              border-color: var(--color-border-strong);
          }

          .opcao-planta label:has(input:focus-visible) {
              outline: none;
              box-shadow: 0 0 0 3px var(--color-accent-soft);
              border-color: var(--color-accent);
          }

          .opcao-planta label:has(input:checked) {
              background-color: var(--color-accent-soft);
              border-color: var(--color-accent);
          }

          /* — Cor própria por planta (mais escura por padrão, mais clara no hover) — */
          .planta-bento label       { color: hsl(280, 70%, 35%); }
          .planta-bento label:hover { color: hsl(280, 70%, 55%); }

          .planta-maranguape label       { color: hsl(35, 90%, 32%); }
          .planta-maranguape label:hover { color: hsl(35, 90%, 50%); }

          .planta-saquarema label       { color: hsl(3, 80%, 40%); }
          .planta-saquarema label:hover { color: hsl(3, 80%, 58%); }

          @media (prefers-reduced-motion: reduce) {
              *, *::before, *::after {
                  animation-duration: 0.01ms !important;
                  transition-duration: 0.01ms !important;
              }
          }
      </style>
  </head>
  <body>
      <div class="selecao-planta-wrapper">
        <form name="form" method="post" action="pedidos.asp" id="pedidoForm" class="selecao-planta-card">
          <input type="hidden" name="Planta" id="Planta" />

          <div class="selecao-planta-header">
              <span class="page-title">Emitir Pedidos</span>
              <p class="selecao-planta-subtitulo">Selecione a planta para continuar</p>
          </div>

          <ul class="selecao-planta-lista" role="radiogroup" aria-label="Selecione a planta">
              <li class="opcao-planta planta-bento">
                  <label for="planta05">
                      <input type="radio" name="selecione" id="planta05" value="05" onclick="submitForm(this.value);" />
                      Bento Gonçalves
                  </label>
              </li>
              <li class="opcao-planta planta-maranguape">
                  <label for="planta06">
                      <input type="radio" name="selecione" id="planta06" value="06" onclick="submitForm(this.value);" />
                      Maranguape
                  </label>
              </li>
              <li class="opcao-planta planta-saquarema">
                  <label for="planta07">
                      <input type="radio" name="selecione" id="planta07" value="07" onclick="submitForm(this.value);" />
                      Saquarema
                  </label>
              </li>
          </ul>
        </form>
      </div>

      <!-- jQuery -->
      <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

      <!-- Bootstrap JS Bundle -->
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
              integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
              crossorigin="anonymous"></script>

      <script type="text/javascript">
          function submitForm(planta) {
              const form = document.getElementById('pedidoForm');
              const plantaInput = document.getElementById('Planta');

              plantaInput.value = planta;
              form.action = "pedidos.asp?paramPlanta=" + planta;
              form.submit();
          }
      </script>
  </body>
</html>
