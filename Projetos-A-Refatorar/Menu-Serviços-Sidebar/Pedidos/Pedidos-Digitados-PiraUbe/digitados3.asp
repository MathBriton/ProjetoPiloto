<!--#include file="../../../estatisticas.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../includes/atribui_nivel.inc" -->
<!--#include file="../../../../cnn/cnn.asp" -->

<%
'conecta_banco_siga.inc

'***************************************'
'***        Exclui Pedidos           ***'
'***************************************'

'RECUPERANDO Nº DOS PEDIDOS A SEREM EXCLUÍDOS
checado = session("checado")
Dados.Open( "UPDATE SZF010 SET D_E_L_E_T_ = '*', ZF_ESTATUS = 'EX', ZF_QUEMLIB = '" & Usr & "' WHERE ZF_NUMERO IN(" & checado & ") AND LTRIM(RTRIM(SZF010.ZF_ESTATUS)) NOT IN ('LI','OK') " )

'*************************************'
'** Enviando email de confirmação   **'
'*************************************'
'Set MailCliente = Server.CreateObject("Persits.MailSender")
'MailCliente.Host = "smail.inpa-embalagens.com.br"
'MailCliente.From = "webmaster@inpa-embalagens.com.br"
'MailCliente.FromName = "INPANET - INPA"
'MailCliente.AddAddress "comercial-rj@inpa-embalagens.com.br" , "COMERCIAL-RJ"
'MailCliente.AddCC "jefferson@inpa-embalagens.com.br" , "DTI"
'MailCliente.Subject = "Exclusão de Pedido(s) digitados que não foram liberados"
'MailCliente.Body = "Os seguintes pedidos foram excluídos por " & Usr & " em " & now & chr(13) & chr(10) & "Pedido(s) :" & checado
'on Error Resume Next
'MailCliente.Send
'if Err <> 0 then
'    Response.write "Um erro ocorreu :" & err.description
'end if
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet - Consulta / Exclusão a Pedidos Digitados</title>

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

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

      /* — Estado — */
      --color-success:       hsl(145, 63%, 32%);
      --color-success-soft:  hsl(145, 60%, 95%);

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
      --space-3: 12px;
      --space-4: 16px;
      --space-5: 24px;
    }

    body {
      background-color: var(--color-bg);
      min-height: 100vh;
      font-family: var(--font-sans);
      color: var(--color-text);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: var(--space-5);
    }

    .page-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
      font-weight: 600;
      letter-spacing: -0.01em;
      color: var(--color-text);
    }

    .result-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-md);
      max-width: 460px;
      width: 100%;
      padding: var(--space-5);
      text-align: center;
    }

    .result-icon {
      width: 56px;
      height: 56px;
      border-radius: 50%;
      background-color: var(--color-success-soft);
      color: var(--color-success);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.6rem;
      margin: 0 auto var(--space-4);
    }

    .result-text {
      font-size: .9rem;
      color: var(--color-text-muted);
      margin-bottom: var(--space-5);
    }

    .result-text a {
      color: var(--color-accent);
      text-decoration: none;
      font-weight: 500;
    }

    .result-text a:hover {
      color: var(--color-accent-hover);
      text-decoration: underline;
    }

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

    .btn-primary:focus,
    .btn-primary:focus-visible {
      box-shadow: 0 0 0 3px var(--color-accent-soft);
      outline: none;
    }

    .btn-primary:active {
      transform: scale(.98);
    }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.01ms !important;
        transition-duration: 0.01ms !important;
      }
    }
  </style>
</head>

<body>

  <main class="result-card" role="status" aria-live="polite">
    <div class="mb-3">
      <span class="page-title">Conclusão de Exclusão de Pedidos Digitados</span>
    </div>

    <div class="result-icon">
      <i class="bi bi-check-lg"></i>
    </div>

    <p class="result-text">
      Exclusão efetuada com sucesso!<br />
      <!-- Clique <a href="../../../meio_principal.asp">aqui</a> para voltar à página inicial. -->
    </p>

    <button type="button" class="btn btn-primary btn-sm px-4" id="btnVoltar">
      <i class="bi bi-arrow-left me-2"></i>Voltar à exclusão de pedidos
    </button>
  </main>

  <script>
    $(function () {
      $('#btnVoltar').on('click', function () {
        // Navegação direta (nova requisição ao servidor), evitando que o
        // navegador restaure a página anterior a partir do bfcache com a
        // lista de pedidos desatualizada (itens já excluídos reaparecendo).
        window.location.replace('digitados1.asp');
      });
    });

    // Reforço: se esta própria página for restaurada a partir do bfcache
    // (ex.: usuário navega para frente/voltar pelo navegador), força reload.
    window.addEventListener('pageshow', function (event) {
      if (event.persisted) {
        window.location.reload();
      }
    });
  </script>

</body>
</html>
