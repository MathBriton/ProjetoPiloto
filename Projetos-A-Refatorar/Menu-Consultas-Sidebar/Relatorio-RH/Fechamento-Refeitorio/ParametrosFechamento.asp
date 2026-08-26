<%Session.LCID=1046%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Fechamento Refeitório</title>
    <!--#include file="../cnn2/cnn2.asp"-->
    <!--#include file="../funcoes/funcoes.asp"-->
    <!--#include file="../includes/variaveis.asp" -->
    <script type="text/javascript" src="../funcoes/funcoesJAVASCRIPT.js"></script>
    <style>
      :root {
        --background: hsl(0, 0%, 100%);
        --foreground: hsl(240, 10%, 3.9%);
        --card: hsl(0, 0%, 100%);
        --muted: hsl(240, 4.8%, 95.9%);
        --muted-foreground: hsl(240, 3.8%, 46.1%);
        --border: hsl(240, 5.9%, 90%);
        --input: hsl(240, 5.9%, 90%);
        --primary: hsl(221, 83%, 53%);
        --primary-hover: hsl(221, 83%, 47%);
        --primary-foreground: hsl(0, 0%, 98%);
        --ring: hsl(221, 83%, 53%);
        --radius: .5rem;
        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        --shadow-card: 0 1px 2px hsla(240, 10%, 3.9%, .04);
        --transition: 150ms ease;
      }

      *,
      *::before,
      *::after {
        box-sizing: border-box;
      }

      body {
        min-height: 100vh;
        margin: 0;
        background-color: hsl(240, 4.8%, 97.9%);
        color: var(--foreground);
        font-family: var(--font-sans);
        font-size: 14px;
        line-height: 1.5;
        -webkit-font-smoothing: antialiased;
      }

      button,
      input,
      select {
        font: inherit;
      }

      .page-container {
        display: flex;
        width: 100%;
        min-height: 100vh;
        align-items: center;
        flex-direction: column;
        justify-content: center;
        padding: 24px;
      }

      .page-title {
        margin: 0;
        color: var(--foreground);
        font-size: clamp(1.125rem, 1rem + .35vw, 1.375rem);
        font-weight: 600;
        line-height: 1.2;
        letter-spacing: -.025em;
      }

      .form-card {
        width: min(100%, 720px);
        background-color: var(--card);
        border: 1px solid var(--border);
        border-radius: calc(var(--radius) + 2px);
        box-shadow: var(--shadow-card);
      }

      .form-card-header {
        padding: 22px 24px;
        border-bottom: 1px solid var(--border);
      }

      .form-card-body {
        padding: 24px;
      }

      .form-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 20px;
      }

      .form-field {
        display: flex;
        min-width: 0;
        flex-direction: column;
        gap: 8px;
      }

      .field-label {
        color: var(--foreground);
        font-size: .875rem;
        font-weight: 500;
        line-height: 1;
      }

      .field-control {
        width: 100%;
        height: 40px;
        padding: 0 12px;
        background-color: var(--background);
        border: 1px solid var(--input);
        border-radius: var(--radius);
        color: var(--foreground);
        font-size: .875rem;
        line-height: 1;
        outline: none;
        transition: border-color var(--transition), box-shadow var(--transition), background-color var(--transition);
      }

      select.field-control {
        appearance: none;
        padding-right: 36px;
        background-image:
          linear-gradient(45deg, transparent 50%, hsl(240, 3.8%, 46.1%) 50%),
          linear-gradient(135deg, hsl(240, 3.8%, 46.1%) 50%, transparent 50%);
        background-position:
          calc(100% - 16px) 50%,
          calc(100% - 11px) 50%;
        background-repeat: no-repeat;
        background-size: 5px 5px, 5px 5px;
        cursor: pointer;
      }

      input[type="date"].field-control {
        color-scheme: light;
        cursor: pointer;
      }

      input[type="date"].field-control::-webkit-calendar-picker-indicator {
        cursor: pointer;
        opacity: .65;
      }

      .field-control:hover {
        border-color: hsl(240, 5.2%, 76%);
      }

      .field-control:focus-visible {
        border-color: var(--ring);
        box-shadow: 0 0 0 3px hsla(221, 83%, 53%, .15);
      }

      .form-actions {
        display: flex;
        justify-content: flex-end;
        margin-top: 24px;
      }

      .button-primary {
        display: inline-flex;
        min-height: 40px;
        align-items: center;
        justify-content: center;
        gap: 8px;
        padding: 0 20px;
        background-color: var(--primary);
        border: 1px solid var(--primary);
        border-radius: var(--radius);
        box-shadow: 0 1px 2px hsla(240, 10%, 3.9%, .08);
        color: var(--primary-foreground);
        cursor: pointer;
        font-size: .875rem;
        font-weight: 500;
        line-height: 1;
        outline: none;
        transition: background-color var(--transition), box-shadow var(--transition), transform var(--transition), opacity var(--transition);
      }

      .button-primary:hover:not(:disabled) {
        background-color: var(--primary-hover);
        border-color: var(--primary-hover);
      }

      .button-primary:focus-visible {
        box-shadow: 0 0 0 3px hsla(221, 83%, 53%, .22);
      }

      .button-primary:active:not(:disabled) {
        transform: translateY(1px);
      }

      /* ── Estado de loading ── */
      .button-primary:disabled {
        opacity: .75;
        cursor: not-allowed;
      }

      .btn-spinner {
        display: none;
        width: 14px;
        height: 14px;
        border: 2px solid hsla(0, 0%, 100%, .35);
        border-top-color: #fff;
        border-radius: 50%;
        animation: spin .65s linear infinite;
        flex-shrink: 0;
      }

      .button-primary.is-loading .btn-spinner { display: block; }
      .button-primary.is-loading .btn-label   { opacity: .85; }

      @keyframes spin {
        to { transform: rotate(360deg); }
      }

      @media (max-width: 767.98px) {
        .page-container {
          padding: 16px;
        }

        .form-card-body {
          padding: 20px;
        }

        .form-card-header {
          padding: 20px;
        }

        .form-grid {
          grid-template-columns: 1fr;
          gap: 16px;
        }

        .button-primary {
          width: 100%;
          min-height: 44px;
        }
      }

      @media (prefers-reduced-motion: reduce) {
        *,
        *::before,
        *::after {
          animation-duration: .01ms !important;
          transition-duration: .01ms !important;
        }
      }
    </style>
  </head>
  <body>
    <main id="param" class="page-container">
      <section class="form-card" aria-labelledby="pageTitle">
        <form action="FechamentoRefeitorio.asp" method="get" name="form" id="formFechamento">
          <header class="form-card-header">
            <h1 id="pageTitle" class="page-title">Fechamento Refeitório</h1>
          </header>

          <div class="form-card-body">
            <div class="form-grid">
              <div class="form-field">
                <label class="field-label" for="dtini">Data Inicial:</label>
                <input type="date" required="required" name="dtini" id="dtini" value="<%=trim(request("dtini")) %>" class="field-control" />
              </div>

              <div class="form-field">
                <label class="field-label" for="dtfim">Data Final:</label>
                <input type="date" required="required" name="dtfim" id="dtfim" value="<%=trim(request("dtfim")) %>" class="field-control" />
              </div>

              <div class="form-field">
                <label class="field-label" for="tipo">Tipo:</label>
                <select name="tipo" id="tipo" required="required" class="field-control">
                  <option value="1" <%if trim(Request("tipo")) = "1" then %> selected="selected" <%end if %>>Analitico</option>
                  <option value="2" <%if trim(Request("tipo")) = "2" then %> selected="selected" <%end if %>>Sintetico</option>
                </select>
              </div>

              <div class="form-field">
                <label class="field-label" for="Planta">Planta:</label>
                <select name="Planta" id="Planta" required="required" class="field-control">
                  <option value="P" <%if trim(Request("Planta")) = "P" then %> selected="selected" <%end if %>>Pirapetinga</option>
                  <option value="U" <%if trim(Request("Planta")) = "U" then %> selected="selected" <%end if %>>Uberaba</option>
                </select>
              </div>
            </div>

            <div class="form-actions">
              <button type="button" id="btnConsultar" class="button-primary" onclick="consultar()">
                <span class="btn-spinner" aria-hidden="true"></span>
                <span class="btn-label">Consultar</span>
              </button>
            </div>
          </div>
        </form>
      </section>
    </main>

    <script type="text/javascript">
      function consultar() {
        var dtini = document.getElementById('dtini').value;
        var dtfim = document.getElementById('dtfim').value;

        if (dtini === '') {
          alert('Preencha a data inicial');
          return false;
        }
        if (dtfim === '') {
          alert('Preencha a data final');
          return false;
        }

        var btn = document.getElementById('btnConsultar');
        btn.classList.add('is-loading');
        btn.disabled = true;
        btn.querySelector('.btn-label').textContent = 'Aguarde...';

        document.getElementById('formFechamento').submit();
      }
    </script>
  </body>
</html>
