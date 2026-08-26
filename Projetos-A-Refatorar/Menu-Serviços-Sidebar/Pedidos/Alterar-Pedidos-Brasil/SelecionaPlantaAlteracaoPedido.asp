<!-- #include file = "../../../includes/variaveis.asp" -->
<!-- #include file = "../../../funcoes/funcoes.asp"-->
<%
' ------------------------------------------------------------------------------------------------------------------------------------------------
' Empresa        :  Smurfit Kappa
' Sistema        :  SKnet
' Fonte          :  SelecionaPlantaAlteracaoPedido.asp
' Funcionalidade :  Apresentar as plantas (BEN/FOR/SAQ) para que o usuário possa selecionar com qual irá trabalhar.
'
' ATUALIZAÇÕES
'
' CHPL           :  371 - Inserçao de Fortaleza e Saquarema na rotina de alteraçao de pedidos.
' Analista       :  Jorge Queiroz
' Data           :  17/05/2024
' ------------------------------------------------------------------------------------------------------------------------------------------------

  Dim ExibeBento, ExibeMaranguape, ExibeSaquarema

  ExibeBento      = Not (SESSION("FILIAL") <> "05" AND SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "P")
  ExibeMaranguape = Not (SESSION("FILIAL") <> "06" AND SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "P")
  ExibeSaquarema  = Not (SESSION("FILIAL") <> "07" AND SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = "P")
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

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

    <title>Alterar Pedidos - Seleção de Planta</title>

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
            --space-7: 48px;

            /* — Cores específicas por planta — */
            --planta-bento-cor:        hsl(262, 70%, 48%);
            --planta-bento-cor-hover:  hsl(262, 70%, 66%);
            --planta-bento-fundo-hover: hsl(262, 70%, 96%);
            --planta-maranguape-cor:       hsl(165, 70%, 32%);
            --planta-maranguape-cor-hover: hsl(165, 60%, 50%);
            --planta-maranguape-fundo-hover: hsl(165, 60%, 95%);
            --planta-saquarema-cor:        hsl(25, 90%, 46%);
            --planta-saquarema-cor-hover:  hsl(25, 90%, 62%);
            --planta-saquarema-fundo-hover: hsl(25, 90%, 95%);
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
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .pagina-container {
            width: 100%;
            display: flex;
            justify-content: center;
        }

        .selecao-planta-card {
            width: 100%;
            max-width: 560px;
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-sm);
            overflow: hidden;
        }

        .selecao-planta-header {
            background-color: var(--color-surface);
            border-bottom: 1px solid var(--color-border);
            padding: 1.1rem 1.25rem;
            text-align: center;
        }

        .selecao-planta-titulo {
            font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
            font-weight: 600;
            letter-spacing: -0.01em;
            color: var(--color-text);
            margin: 0;
        }

        .selecao-planta-descricao {
            font-size: .82rem;
            font-weight: 400;
            color: var(--color-text-muted);
            margin: .3rem 0 0;
        }

        .selecao-planta-lista {
            display: flex;
            flex-direction: column;
            gap: var(--space-3);
            padding: 1.25rem;
        }

        .planta-item {
            display: flex;
            align-items: center;
            gap: var(--space-3);
            padding: 1rem 1.25rem;
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            cursor: pointer;
            transition: background-color var(--transition), border-color var(--transition), box-shadow var(--transition);
        }

        .planta-item:hover {
            border-color: var(--color-border-strong);
            box-shadow: var(--shadow-sm);
        }

        .planta-item .form-check-input {
            margin: 0;
            cursor: pointer;
        }

        .planta-item .form-check-input:focus {
            border-color: var(--color-accent);
            box-shadow: 0 0 0 3px var(--color-accent-soft);
        }

        .planta-nome {
            font-size: .92rem;
            font-weight: 500;
            transition: color var(--transition);
        }

        .planta-item--bento:hover {
            background-color: var(--planta-bento-fundo-hover);
        }

        .planta-nome--bento {
            color: var(--planta-bento-cor);
        }
        .planta-item:hover .planta-nome--bento {
            color: var(--planta-bento-cor-hover);
        }

        .planta-item--maranguape:hover {
            background-color: var(--planta-maranguape-fundo-hover);
        }

        .planta-nome--maranguape {
            color: var(--planta-maranguape-cor);
        }
        .planta-item:hover .planta-nome--maranguape {
            color: var(--planta-maranguape-cor-hover);
        }

        .planta-item--saquarema:hover {
            background-color: var(--planta-saquarema-fundo-hover);
        }

        .planta-nome--saquarema {
            color: var(--planta-saquarema-cor);
        }
        .planta-item:hover .planta-nome--saquarema {
            color: var(--planta-saquarema-cor-hover);
        }

    </style>
</head>
<body>

    <div class="container-fluid py-3 px-3 px-lg-4 pagina-container">

        <form name="form" method="post" action="ListaSolicitacaoAlteracaoPedido.asp">

                <input type="hidden" name="Planta" id="Planta" />

                <div class="selecao-planta-card">
                    <div class="selecao-planta-header">
                        <p class="selecao-planta-titulo">Alterar Pedidos</p>
                        <p class="selecao-planta-descricao">Selecione a planta com a qual deseja trabalhar</p>
                    </div>

                    <div class="selecao-planta-lista">
<%
                        If ExibeBento Then
%>
                        <label class="planta-item planta-item--bento" for="plantaBento">
                            <input type="radio" class="form-check-input" name="selecione" id="plantaBento" value="05" />
                            <span class="planta-nome planta-nome--bento">Bento Gonçalves</span>
                        </label>
<%
                        End If

                        If ExibeMaranguape Then
%>
                        <label class="planta-item planta-item--maranguape" for="plantaMaranguape">
                            <input type="radio" class="form-check-input" name="selecione" id="plantaMaranguape" value="06" />
                            <span class="planta-nome planta-nome--maranguape">Maranguape</span>
                        </label>
<%
                        End If

                        If ExibeSaquarema Then
%>
                        <label class="planta-item planta-item--saquarema" for="plantaSaquarema">
                            <input type="radio" class="form-check-input" name="selecione" id="plantaSaquarema" value="07" />
                            <span class="planta-nome planta-nome--saquarema">Saquarema</span>
                        </label>
<%
                        End If
%>
                    </div>
                </div>
            </form>
    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

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
