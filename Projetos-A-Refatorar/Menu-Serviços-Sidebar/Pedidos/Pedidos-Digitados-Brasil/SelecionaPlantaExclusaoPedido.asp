<!DOCTYPE html>
<!-- #include file="../../../includes/variaveis.asp" -->
<!-- #include file="../../../funcoes/funcoes.asp"-->
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet - Seleciona Planta Exclusão de Pedidos</title>

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" />

    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />

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
          --space-3: 12px;
          --space-4: 16px;
          --space-5: 24px;
        }

        body {
            background-color: var(--color-bg);
            min-height: 100vh;
            font-family: var(--font-sans);
            color: var(--color-text);
        }

        .page-wrapper {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            padding: var(--space-5);
        }

        .selection-card {
            width: 100%;
            max-width: 380px;
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-md);
            overflow: hidden;
        }

        .selection-card-header {
            padding: var(--space-5) var(--space-5) var(--space-3);
            text-align: center;
        }

        .selection-card-header .page-title {
            display: block;
            font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
            font-weight: 600;
            letter-spacing: -0.01em;
            color: var(--color-text);
        }

        .selection-card-header .page-subtitle {
            display: block;
            margin-top: var(--space-2);
            font-size: .82rem;
            color: var(--color-text-muted);
        }

        .selection-list {
            display: flex;
            flex-direction: column;
            padding: 0 var(--space-3) var(--space-4);
            gap: var(--space-2);
        }

        .selection-option {
            display: flex;
            align-items: center;
            gap: var(--space-3);
            padding: var(--space-3) var(--space-4);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            background-color: var(--color-surface);
            cursor: pointer;
            transition: background-color var(--transition), border-color var(--transition);
        }

        .selection-option:hover {
            background-color: var(--color-surface-muted);
            border-color: var(--color-border-strong);
        }

        /* — Hover colorido por planta — */
        .selection-option.planta-ben:hover {
            background-color: hsl(210, 100%, 96%);
            border-color: hsl(210, 100%, 70%);
        }
        .selection-option.planta-ben:hover i.bi {
            color: hsl(210, 100%, 45%);
        }
        .selection-option.planta-ben:hover .selection-option-label {
            color: hsl(210, 100%, 30%);
        }

        .selection-option.planta-for:hover {
            background-color: hsl(145, 60%, 95%);
            border-color: hsl(145, 55%, 65%);
        }
        .selection-option.planta-for:hover i.bi {
            color: hsl(145, 63%, 32%);
        }
        .selection-option.planta-for:hover .selection-option-label {
            color: hsl(145, 70%, 20%);
        }

        .selection-option.planta-saq:hover {
            background-color: hsl(35, 100%, 95%);
            border-color: hsl(35, 90%, 65%);
        }
        .selection-option.planta-saq:hover i.bi {
            color: hsl(35, 100%, 40%);
        }
        .selection-option.planta-saq:hover .selection-option-label {
            color: hsl(35, 100%, 25%);
        }

        .selection-option:focus-within {
            outline: none;
            box-shadow: 0 0 0 3px var(--color-accent-soft);
            border-color: var(--color-accent);
        }

        .selection-option .form-check-input {
            width: 1.05rem;
            height: 1.05rem;
            margin: 0;
            border-color: var(--color-border-strong);
        }

        .selection-option .form-check-input:checked {
            background-color: var(--color-accent);
            border-color: var(--color-accent);
        }

        .selection-option .form-check-input:focus {
            border-color: var(--color-accent);
            box-shadow: 0 0 0 3px var(--color-accent-soft);
        }

        .selection-option i.bi {
            color: var(--color-text-faint);
            font-size: 1rem;
        }

        .selection-option-label {
            flex: 1;
            font-size: .9rem;
            font-weight: 500;
            color: var(--color-text);
            cursor: pointer;
        }

        /* — Cor do texto por planta, já aplicada no estado padrão — */
        .selection-option.planta-ben .selection-option-label {
            color: hsl(210, 100%, 45%);
        }
        .selection-option.planta-for .selection-option-label {
            color: hsl(145, 63%, 32%);
        }
        .selection-option.planta-saq .selection-option-label {
            color: hsl(35, 100%, 40%);
        }

        @media (prefers-reduced-motion: reduce) {
            *, *::before, *::after {
                animation-duration: 0.01ms !important;
                transition-duration: 0.01ms !important;
            }
        }

        @media (max-width: 576px) {
            .selection-card {
                max-width: 100%;
            }
        }
    </style>
</head>
<body>

    <div class="page-wrapper">
        <div class="selection-card">
            <div class="selection-card-header">
                <span class="page-title">Pedidos Digitados</span>
                <span class="page-subtitle">Selecione a planta com a qual deseja trabalhar</span>
            </div>

            <hr style="color:hsl(240, 1%, 32%);">

            <form name="form" method="post" action="PedidoDigitado.asp">
                <input type="hidden" name="Planta" id="Planta" />

                <div class="selection-list" role="radiogroup" aria-label="Seleção de planta">
                    <label class="selection-option planta-ben" for="plantaBen">
                        <input class="form-check-input" type="radio" name="selecione" id="plantaBen" value="05" />
                        <span class="selection-option-label">Bento Gonçalves</span>
                    </label>

                    <label class="selection-option planta-for" for="plantaFor">
                        <input class="form-check-input" type="radio" name="selecione" id="plantaFor" value="06" />
                        <span class="selection-option-label">Maranguape</span>
                    </label>

                    <label class="selection-option planta-saq" for="plantaSaq">
                        <input class="form-check-input" type="radio" name="selecione" id="plantaSaq" value="07" />
                        <span class="selection-option-label">Saquarema</span>
                    </label>
                </div>
            </form>
        </div>
    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"></script>

    <script>
        $(function () {
            $('input[name="selecione"]').on('change', function () {
                $('#Planta').val($(this).val());
                $('form[name="form"]').trigger('submit');
            });
        });
    </script>
</body>
</html>
