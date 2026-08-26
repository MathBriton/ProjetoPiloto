<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../includes/atribui_nivel.inc" -->
<!--#include file="../../../../cnn/cnn.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
<%
response.Expires = -1
response.ExpiresAbsolute = Now() - 1
response.AddHeader "pragma", "no-store"
response.AddHeader "cache-control", "no-store, no-cache, must-revalidate"
session.LCID = 1046

' Verificando o nivel para ver se o usuario tem permissao para deletar pedidos de outros usuarios
' B = TEM PERMISSAO
' L = PERMISSAO SO DE LEITURA
' G = PERMISSAO DE GRAVACAO
PERMISSAO = "B"

Dim vData1, vData2, vEmissao, y, isLiberado

SQL = "Execute SP_CONSULTAPEDIDOSZF2 [" & usr & "],[" & nivel & "],[" & permissao & "]"

Dados.Open( SQL )
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
      --color-success-row:   hsl(145, 55%, 88%);
      --color-danger:        hsl(3, 90%, 58%);
      --color-danger-soft:   hsl(3, 100%, 96%);

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
    }

    .page-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
      font-weight: 600;
      letter-spacing: -0.01em;
      color: var(--color-text);
    }

    .table-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-sm);
      position: relative;
    }

    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 70vh;
    }
    .table-wrapper table { min-width: max-content; }

    .table thead th {
      position: sticky;
      top: 0;
      background-color: var(--color-surface-muted);
      color: var(--color-text-muted);
      font-size: .76rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .03em;
      text-align: center;
      white-space: nowrap;
      z-index: 2;
      border-color: var(--color-border);
      padding: .6rem .75rem;
    }

    .table tbody td {
      font-size: .85rem;
      text-align: center;
      white-space: nowrap;
      color: var(--color-text);
      vertical-align: middle;
      padding: .55rem .75rem;
    }

    .table tbody tr:not([class*="row-"]):hover td {
      background-color: var(--color-surface-muted);
    }

    .table > tbody > tr.row-success > td { background-color: var(--color-success-row); }
    .table > tbody > tr.row-success:hover > td { filter: brightness(.95); }

    .table-footer {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      padding: .7rem 1.25rem;
      border-top: 1px solid var(--color-border);
      flex-wrap: wrap;
      gap: .5rem;
    }
    .info-registros { font-size: .8rem; color: var(--color-text-muted); }

    .table-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 4rem 1rem;
      color: var(--color-text-muted);
      gap: .85rem;
    }
    .table-placeholder .bi { font-size: 2.4rem; color: var(--color-text-faint); }
    .table-placeholder p  { margin: 0; font-size: .88rem; text-align: center; }

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

    .btn-danger {
      font-weight: 500;
      font-size: .85rem;
      transition: background-color var(--transition), border-color var(--transition);
    }

    .form-control:focus,
    .form-select:focus,
    .form-check-input:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
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

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="row">
      <div class="col-12 col-lg-8 offset-lg-2 d-flex align-items-center gap-2">
        <button type="button" class="btn btn-light btn-sm" onclick="history.go(-1);">
          <i class="bi bi-arrow-left me-1"></i>Voltar
        </button>
        <span class="page-title">Pedidos Digitados</span>
      </div>
    </div>

    <div class="row mt-2">
      <div class="col-12 col-lg-8 offset-lg-2">
        <a href="../../../../comercial/formularioV1/SelecionaPlanta.asp" class="btn btn-primary btn-sm">
          <i class="bi bi-plus-lg me-1"></i>Emitir novo pedido
        </a>
      </div>
    </div>

<% If Dados.EOF Then %>

    <div class="row mt-3">
      <div class="col-12 col-lg-8 offset-lg-2">
        <div class="table-card">
          <div class="table-placeholder">
            <i class="bi bi-clipboard-data" aria-hidden="true"></i>
            <p>Nenhum pedido foi encontrado. Verifique na consulta <strong>Pedidos Validados</strong>.</p>
          </div>
        </div>
      </div>
    </div>

<%
  Dados.Close()
Else
%>

    <form method="post" action="digitados2.asp" name="form" id="formPedidos">

      <div class="row mt-3">
        <div class="col-12 col-lg-8 offset-lg-2">
          <div class="table-card">
        <div class="table-wrapper">
          <table class="table table-bordered table-hover table-sm mb-0">
            <thead>
              <tr>
                <th scope="col">
                  <i class="bi bi-trash" style="color: var(--color-danger);" aria-hidden="true" title="Selecione para excluir!"></i>
                </th>
                <th scope="col">Pedido</th>
                <th scope="col">Cliente</th>
                <th scope="col">Nome</th>
                <th scope="col">FT</th>
                <th scope="col">Produto</th>
                <th scope="col">Qtde.</th>
                <th scope="col">Preço</th>
                <th scope="col">Entrega</th>
                <th scope="col">Medida</th>
                <th scope="col">Vincos</th>
                <th scope="col">Emissão</th>
              </tr>
            </thead>
            <tbody>
<%
              y = 0

              Do Until Dados.EOF

                vData1   = mid(Dados("ZF_ENT1"), 7, 2)     & "/" & mid(Dados("ZF_ENT1"), 5, 2)     & "/" & mid(Dados("ZF_ENT1"), 1, 4)
                vData2   = mid(Dados("ZF_ENT2"), 7, 2)     & "/" & mid(Dados("ZF_ENT2"), 5, 2)     & "/" & mid(Dados("ZF_ENT2"), 1, 4)
                vEmissao = mid(Dados("ZF_EMISSAO"), 7, 2)  & "/" & mid(Dados("ZF_EMISSAO"), 5, 2)  & "/" & mid(Dados("ZF_EMISSAO"), 1, 4)

                isLiberado = (Trim(Dados("ZF_ESTATUS")) = "LI" Or Trim(Dados("ZF_ESTATUS")) = "OK")
%>
              <tr<% If isLiberado Then %> class="row-success"<% End If %>>
                <td>
<%                If isLiberado Then %>
                  <span class="badge text-bg-light border">Liberado</span>
<%                Else %>
                  <input type="checkbox" class="form-check-input" name="checados"
                         value="<%="'" & Dados("ZF_NUMERO") & "'"%>"
                         aria-label="Selecionar pedido <%=Dados("ZF_NUMERO")%> para exclusão" />
<%                End If %>
                </td>
                <td><%=Dados("ZF_NUMERO") %></td>
                <td><%=Dados("ZF_CLIENTE") & " - " & Dados("ZF_LOJACLI") %></td>
                <td><%=Dados("CLIENTE") %></td>
                <td><%=Dados("ZF_PRODUTO") %></td>
                <td><%=Dados("REF") %></td>
                <td><%=FormatNumber(Dados("ZF_QUANT"), 0) %></td>
                <td><%=FormatNumber(Dados("ZF_PRECO"), 5) %></td>
                <td><%=vData1 & " - " & vData2 %></td>
                <td><%=FormatNumber(Dados("ZF_LARG"), 0) & " x " & FormatNumber(Dados("ZF_COMP"), 0) %></td>
                <td><%=Dados("ZF_VINCO1") & " - " & FormatNumber(Dados("ZF_VINCO2"), 0) & " - " & FormatNumber(Dados("ZF_VINCO3"), 0) %></td>
                <td><%=vEmissao & " às " & Dados("ZF_HORA") %></td>
              </tr>
<%
                y = y + 1
                Dados.MoveNext()

              Loop
              Dados.Close()
%>
            </tbody>
          </table>
        </div>

        <div class="table-footer">
          <span id="infoRegistros" class="info-registros" aria-live="polite">
<%
            If y = 1 Then
              response.write "Foi encontrado " & y & " pedido digitado."
            Else
              response.write "Foram encontrados " & y & " pedidos digitados."
            End If
%>
          </span>
        </div>
      </div>

      <div class="d-flex justify-content-center gap-2 mt-4">
        <button type="button" class="btn btn-light btn-sm" name="Reset" onclick="document.getElementById('formPedidos').reset();">
          <i class="bi bi-eraser me-1"></i>Limpar Campos
        </button>
        <button type="button" class="btn btn-danger btn-sm" data-bs-toggle="modal" data-bs-target="#modalConfirmaExclusao">
          <i class="bi bi-trash me-1"></i>Excluir Pedidos Selecionados
        </button>
      </div>

        </div>
      </div>

    </form>

    <!-- Modal de confirmação — ação destrutiva -->
    <div class="modal fade" id="modalConfirmaExclusao" tabindex="-1" aria-labelledby="modalConfirmaExclusaoLabel" aria-hidden="true">
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="modalConfirmaExclusaoLabel">Confirmar exclusão</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">
            Tem certeza que deseja excluir os pedidos selecionados? Esta ação não pode ser desfeita.
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Cancelar</button>
            <button type="button" class="btn btn-danger btn-sm" id="btnConfirmarExclusao">
              <i class="bi bi-trash me-1"></i>Sim, excluir pedidos
            </button>
          </div>
        </div>
      </div>
    </div>

    <script>
      document.getElementById('btnConfirmarExclusao').addEventListener('click', function () {
        document.getElementById('formPedidos').submit();
      });
    </script>

<% End If %>

  </div>

</body>
</html>
