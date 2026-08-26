<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Selecione a Planta</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <!-- Inter font (shadcn/ui usa Inter como fonte padrão) -->
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap"
        rel="stylesheet" />

  <style>
    /* ── Tokens shadcn/ui (tema light) ─────────────────────────── */
    :root {
      --background:   #ffffff;
      --foreground:   #09090b;
      --card:         #ffffff;
      --card-fg:      #09090b;
      --muted:        #f4f4f5;
      --muted-fg:     #71717a;
      --border:       #e4e4e7;
      --ring:         #18181b;
      --radius:       0.5rem;
    }

    * { box-sizing: border-box; }

    body {
      font-family: 'Inter', system-ui, sans-serif;
      background-color: #fafafa;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      color: var(--foreground);
      -webkit-font-smoothing: antialiased;
    }

    /* ── page-title (regra ESTILIZAR) ──────────────────────────── */
    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: var(--foreground);
      letter-spacing: -.025em;
    }

    /* ── Card shadcn ────────────────────────────────────────────── */
    .sw-card {
      background-color: var(--card);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      box-shadow:
        0 1px 3px rgba(0,0,0,.06),
        0 1px 2px rgba(0,0,0,.04);
      width: 100%;
      max-width: 380px;
    }

    .sw-card-header {
      padding: 1.5rem 1.5rem .75rem;
      border-bottom: 1px solid var(--border);
    }

    .sw-card-description {
      font-size: .875rem;
      color: var(--muted-fg);
      font-weight: 400;
      margin-top: .25rem;
      line-height: 1.4;
    }

    .sw-card-body {
      padding: 1.25rem 1.5rem 1.5rem;
      display: flex;
      flex-direction: column;
      gap: .5rem;
    }

    /* ── Itens de planta (shadcn RadioGroup style) ──────────────── */
    .planta-option {
      display: flex;
      align-items: center;
      gap: .75rem;
      padding: .75rem 1rem;
      border-radius: calc(var(--radius) - 2px);
      border: 1px solid var(--border);
      background-color: var(--card);
      cursor: pointer;
      transition:
        background-color .15s ease,
        border-color .15s ease,
        box-shadow .15s ease;
      font-size: .875rem;
      font-weight: 500;
      user-select: none;
    }

    .planta-option:hover {
      background-color: var(--muted);
      border-color: #d4d4d8;
    }

    /* Radio customizado estilo shadcn */
    .planta-option input[type="radio"] {
      appearance: none;
      -webkit-appearance: none;
      width: 1rem;
      height: 1rem;
      border-radius: 50%;
      border: 2px solid var(--border);
      flex-shrink: 0;
      position: relative;
      transition: border-color .15s ease;
      background-color: #fff;
    }

    .planta-option input[type="radio"]:checked::after {
      content: '';
      position: absolute;
      inset: 2px;
      border-radius: 50%;
    }

    /* Cores individuais por planta */
    .planta-bento      { color: #1d4ed8; }
    .planta-bento:hover { border-color: #bfdbfe; background-color: #eff6ff; }
    .planta-bento input[type="radio"]:checked        { border-color: #1d4ed8; }
    .planta-bento input[type="radio"]:checked::after { background-color: #1d4ed8; }

    .planta-maranguape      { color: #b45309; }
    .planta-maranguape:hover { border-color: #fde68a; background-color: #fffbeb; }
    .planta-maranguape input[type="radio"]:checked        { border-color: #b45309; }
    .planta-maranguape input[type="radio"]:checked::after { background-color: #b45309; }

    .planta-saquarema      { color: #15803d; }
    .planta-saquarema:hover { border-color: #bbf7d0; background-color: #f0fdf4; }
    .planta-saquarema input[type="radio"]:checked        { border-color: #15803d; }
    .planta-saquarema input[type="radio"]:checked::after { background-color: #15803d; }

    /* Label de texto dentro do item */
    .planta-label {
      display: flex;
      flex-direction: column;
      gap: .1rem;
    }

    .planta-label-code {
      font-size: .72rem;
      font-weight: 400;
      color: var(--muted-fg);
      letter-spacing: .02em;
    }

    /* ── Overlay de carregamento ────────────────────────────────── */
    .overlay-loading {
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, .85);
      backdrop-filter: blur(2px);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 9999;
      gap: .75rem;
    }

    .overlay-loading .spinner-border {
      width: 1.75rem;
      height: 1.75rem;
      border-width: 2px;
      color: var(--muted-fg);
    }

    .overlay-loading span {
      font-size: .8125rem;
      font-weight: 500;
      color: var(--muted-fg);
    }
  </style>
</head>

<body>

  <!-- Overlay de carregamento -->
  <!-- <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border" role="status">
      <span class="visually-hidden">Carregando...</span>
    </div>
    <span>Aguarde...</span>
  </div> -->

  <!-- Card central -->
  <div class="sw-card">

    <div class="sw-card-header">
      <span class="page-title">Selecione a planta</span>
      <!-- <p class="sw-card-description mb-0">Selecione a Planta :</p> -->
    </div>

    <div class="sw-card-body">

      <label class="planta-option planta-bento">
        <input type="radio" name="opcao" value="05" />
        <span class="planta-label">
          <span>Bento Gonçalves</span>
          <span class="planta-label-code">Rio Grande do Sul</span>
        </span>
      </label>

      <label class="planta-option planta-maranguape">
        <input type="radio" name="opcao" value="06" />
        <span class="planta-label">
          <span>Maranguape</span>
          <span class="planta-label-code">Ceará</span>
        </span>
      </label>

      <label class="planta-option planta-saquarema">
        <input type="radio" name="opcao" value="07" />
        <span class="planta-label">
          <span>Saquarema</span>
          <span class="planta-label-code">Rio de Janeiro</span>
        </span>
      </label>

    </div>
  </div>

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <!-- jQuery -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
    $(function () {
      $('input[name="opcao"]').on('change', function () {
        $('#overlayLoading').removeClass('d-none');
        window.open('ListaPedidosValidar.asp?Planta=' + this.value, '_self');
      });
    });
  </script>

</body>
</html>
