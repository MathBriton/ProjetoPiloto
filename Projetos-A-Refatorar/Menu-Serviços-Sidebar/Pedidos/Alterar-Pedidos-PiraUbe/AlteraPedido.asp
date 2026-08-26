<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
<%
' ==============================================================================
' AlteraPedido.asp — Confirmação de alteração de pedido
' ==============================================================================

Dim pedido, id, operacao, descricaoAlteracao
Dim SQL, oRS

pedido = Trim(Request("Pedido"))
id     = Trim(Request("Id"))

SQL = "SELECT " &_
        "Id, " &_
        "Operacao " &_
      "FROM InHistoricoTipoOperacao " &_
      "WHERE Nivel <= " & NIVELALTERAPEDIDO & " " &_
        "AND Id = " & Trim(Request("Tipo"))

Set oRS = Conecta("PIR-HELPDESK", SQL)

If Not oRS.EOF Then
    operacao           = Trim(oRS("Id"))
    descricaoAlteracao = Trim(oRS("Operacao"))
Else
    operacao           = ""
    descricaoAlteracao = "Tipo de operação não encontrada"
End If

oRS.Close()
Set oRS = Nothing

%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet — Alteração de Pedido</title>

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />

    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
      /* ── Tokens de Design ──────────────────────────────────────────────────── */
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
        --color-danger-hover:  hsl(3, 90%, 50%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

        --radius-sm:  6px;
        --radius-md:  10px;
        --radius-lg:  14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-2: 8px;
        --space-3: 12px;
        --space-4: 16px;
        --space-5: 24px;
        --space-6: 32px;
      }

      /* ── Base ───────────────────────────────────────────────────────────────── */
      *,
      *::before,
      *::after { box-sizing: border-box; }

      body {
        background-color: var(--color-bg);
        font-family: var(--font-sans);
        color: var(--color-text);
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: var(--space-5) var(--space-4);
        margin: 0;
      }

      /* ── Card de confirmação ────────────────────────────────────────────────── */
      .confirm-card {
        width: 100%;
        max-width: 520px;
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-lg);
        box-shadow: var(--shadow-md);
        overflow: hidden;
      }

      .confirm-card-header {
        padding: var(--space-4) var(--space-5);
        border-bottom: 1px solid var(--color-border);
        background-color: var(--color-surface);
      }

      .confirm-card-header .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.125rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
        margin: 0;
      }

      .confirm-card-body {
        padding: var(--space-5);
        display: flex;
        flex-direction: column;
        gap: var(--space-4);
      }

      /* ── Linhas de detalhe ──────────────────────────────────────────────────── */
      .detail-row {
        display: flex;
        align-items: baseline;
        gap: var(--space-3);
        padding: var(--space-3) 0;
        border-bottom: 1px solid var(--color-border);
      }

      .detail-row:last-of-type {
        border-bottom: none;
      }

      .detail-row--striped {
        background-color: var(--color-surface-muted);
      }

      .detail-label {
        flex: 0 0 80px;
        font-size: .78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .04em;
        color: var(--color-text-muted);
        white-space: nowrap;
      }

      .detail-value {
        font-size: .9rem;
        color: var(--color-text);
        font-weight: 400;
        word-break: break-word;
        border-left: 2px solid var(--color-border-strong);
        padding-left: var(--space-3);
      }

      /* ── Área de ações ──────────────────────────────────────────────────────── */
      .confirm-card-footer {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: var(--space-2);
        padding: var(--space-4) var(--space-5);
        border-top: 1px solid var(--color-border);
        background-color: var(--color-surface-muted);
      }

      /* ── Botão primário (Alterar) ───────────────────────────────────────────── */
      .btn-confirm {
        background-color: var(--color-accent);
        border: 1px solid var(--color-accent);
        color: hsl(0, 0%, 100%);
        font-weight: 500;
        font-size: .85rem;
        border-radius: var(--radius-sm);
        padding: .45rem .95rem;
        cursor: pointer;
        transition: background-color var(--transition), border-color var(--transition), transform var(--transition);
        display: inline-flex;
        align-items: center;
        gap: var(--space-2);
        text-decoration: none;
      }

      .btn-confirm:hover {
        background-color: var(--color-accent-hover);
        border-color: var(--color-accent-hover);
        color: hsl(0, 0%, 100%);
      }

      .btn-confirm:active {
        transform: scale(.98);
      }

      .btn-confirm:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }

      /* ── Botão secundário (Voltar) ──────────────────────────────────────────── */
      .btn-back {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        color: var(--color-text);
        font-weight: 500;
        font-size: .85rem;
        border-radius: var(--radius-sm);
        padding: .45rem .95rem;
        cursor: pointer;
        transition: background-color var(--transition), border-color var(--transition), transform var(--transition);
        display: inline-flex;
        align-items: center;
        gap: var(--space-2);
        text-decoration: none;
      }

      .btn-back:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
        color: var(--color-text);
      }

      .btn-back:active {
        transform: scale(.98);
      }

      .btn-back:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }

      /* ── Alerta de operação não encontrada ──────────────────────────────────── */
      .alert-op-missing {
        display: flex;
        align-items: center;
        gap: var(--space-3);
        padding: var(--space-3) var(--space-4);
        background-color: hsl(3, 100%, 96%);
        border: 1px solid hsl(3, 80%, 88%);
        border-radius: var(--radius-sm);
        font-size: .85rem;
        color: hsl(3, 70%, 38%);
      }

      .alert-op-missing .bi {
        font-size: 1.1rem;
        flex-shrink: 0;
      }

      /* ── Acessibilidade ─────────────────────────────────────────────────────── */
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
          animation-duration: 0.01ms !important;
          transition-duration: 0.01ms !important;
        }
      }

      /* ── Responsividade abaixo de --bp-md ───────────────────────────────────── */
      @media (max-width: 576px) {
        .confirm-card-footer {
          flex-direction: column-reverse;
          align-items: stretch;
        }
        .btn-confirm,
        .btn-back {
          justify-content: center;
          width: 100%;
          padding: .6rem 1rem;
        }
        .detail-label {
          flex: 0 0 68px;
        }
      }
    </style>
  </head>

  <body>

    <div class="confirm-card" role="main">

      <!-- Cabeçalho -->
      <div class="confirm-card-header">
        <h1 class="page-title">Alteração de Pedido</h1>
      </div>

      <!-- Campos ocultos do formulário -->
      <form id="form" name="form" method="post" action="AlteraPedido02.asp">
        <input type="hidden" name="id"              id="id"              value="<%=id%>" />
        <input type="hidden" name="pedido"           id="pedido"          value="<%=pedido%>" />
        <input type="hidden" name="TPOperacoes"      id="TPOperacoes"     value="<%=operacao%>" />
        <input type="hidden" name="de"               id="de"              value="<%=Trim(Request("de"))%>" />
        <input type="hidden" name="para"             id="para"            value="<%=Trim(Request("para"))%>" />
        <input type="hidden" name="para2"            id="para2"           value="<%=Trim(Request("para2"))%>" />
        <input type="hidden" name="filtro"           id="filtro"          value="<%=Trim(Request("filtro"))%>" />
        <input type="hidden" name="NomeTPOperacoes"  id="NomeTPOperacoes" value="<%=Trim(descricaoAlteracao)%>" />

        <!-- Corpo: detalhes -->
        <div class="confirm-card-body">

          <div class="detail-row">
            <span class="detail-label">Pedido</span>
            <span class="detail-value"><%=pedido%></span>
          </div>

          <div class="detail-row detail-row--striped">
            <span class="detail-label">Tipo</span>
            <span class="detail-value"><%=descricaoAlteracao%></span>
          </div>

          <div class="detail-row">
            <span class="detail-label">De</span>
            <span class="detail-value"><%=Trim(Request("de"))%></span>
          </div>

          <div class="detail-row detail-row--striped">
            <span class="detail-label">Para</span>
            <span class="detail-value">
              <%=Trim(Request("Descricao"))%><%
              If Trim(Request("para2")) <> "" Then
                Response.Write " &nbsp;a&nbsp; " & Trim(Request("para2"))
              End If
              %>
            </span>
          </div>

          <%
          If operacao = "" Then
          %>
          <div class="alert-op-missing" role="alert">
            <i class="bi bi-exclamation-triangle-fill"></i>
            <span>Tipo de operação não encontrada. Não é possível prosseguir com a alteração.</span>
          </div>
          <%
          End If
          %>

        </div>

        <!-- Rodapé: ações -->
        <%
        If operacao <> "" Then
        %>
        <div class="confirm-card-footer">
          <a href="#" class="btn-back" onclick="history.go(-1); return false;" title="Voltar">
            <i class="bi bi-arrow-left"></i>
            Voltar
          </a>
          <button type="submit" class="btn-confirm" title="Confirmar alteração do pedido">
            <i class="bi bi-pencil-square"></i>
            Alterar Pedido
          </button>
        </div>
        <%
        Else
        %>
        <div class="confirm-card-footer">
          <a href="#" class="btn-back" onclick="history.go(-1); return false;" title="Voltar">
            <i class="bi bi-arrow-left"></i>
            Voltar
          </a>
        </div>
        <%
        End If
        %>

      </form>

    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

  </body>

</html>

