<!-- #include file="../../../cnn2/cnn2.asp" -->
<!-- #include file="../../../includes/variaveis.asp" -->
<!-- #include file="../../../funcoes/funcoes.asp" -->
<%
session.LCID = 1046

'IdUsuario = "000022"

Planta = request("Planta")

if trim(request("checados")) <> "" then

    ListaPedidos = request("checados") ' vem como 'pedido1','pedido2',... (checkboxes "checados")

    ' Verifica se algum dos pedidos selecionados já foi validado
    SQL = "Select Pedido from Pedido where Pedido IN (" & ListaPedidos & ") and Item = '1' and Isnull(Status,0) = 1 AND Isnull(Excluido,0) = 0 AND ISNULL(QuemValidou,0) <> 0 "

    select case Planta
        case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
        case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
        case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
    end select

    if not isobject(consulta) then
       response.end
    end if

    PedidosValidados = ""
    Do While Not consulta.EOF
        PedidosValidados = PedidosValidados & consulta("Pedido") & ", "
        consulta.MoveNext()
    Loop
    consulta.Close()
    SET consulta = NOTHING

    if PedidosValidados <> "" then
             response.write "<script type=""text/javascript"" charset=""utf-8"">"
             response.write "alert('Os seguintes pedidos já foram validados e não podem ser excluídos: " & PedidosValidados & "');"
             response.write "location.href='PedidoDigitado.asp?Planta=" & Planta & "'"
             response.write "</script>"
             response.end
    end if


  SQL = "" &_
    "UPDATE Pedido SET " &_
      "Excluido = 1, " &_
      "QuemExcluiu = " & IdUsuario & ", " &_
      "DataExclusao = GETDATE() " &_
    "WHERE Isnull(Status,0) = 0 " &_
    "AND Pedido IN (" & ListaPedidos & ") " &_
    iif( Nivel <> 0, "AND QuemIncluiu = '"&IdUsuario&"' ", "" )

  select case Planta
  case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
  case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
  case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
  end select

  if vartype(consulta) = 0 or vartype(consulta) = 1 then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "Pedido " &_
    "FROM Pedido " &_
    "WHERE " &_
    "Pedido IN (" & ListaPedidos & ") " &_
    "AND Isnull(Excluido,0) = 1 " &_
    "AND QuemExcluiu = " & IdUsuario

  select case Planta
  case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
  case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
  case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
  end select

  if vartype(consulta) = 0 or vartype(consulta) = 1 then
    response.end
  end if

  PedidosExcluidos = ""
  Do While not consulta.Eof
    PedidosExcluidos = PedidosExcluidos & consulta("Pedido") & ", "
    consulta.MoveNext()
  Loop

  if PedidosExcluidos <> "" then
    response.write "<script type=""text/javascript"">alert('Pedido(s) " & PedidosExcluidos & "excluído(s) com sucesso.')</script>"
  else
    response.write "<script type=""text/javascript"">alert('Não foi possível excluir o(s) pedido(s) selecionado(s).')</script>"
  end if

  consulta.Close()
  SET consulta = NOTHING

end if



SQL = "" &_
  "SELECT " &_
    "Pedido, " &_
    "Planta, " &_
    "Status, " &_
    "LotePiloto, " &_
    "ClienteRemessa, " &_
    "Produto, " &_
    "Referencia, " &_
    "Estrutura, " &_
    "Isnull(PrecoUnitario,0) PrecoUnitario, " &_
    "Quantidade, " &_
    "Observacao, " &_
    "Largura, " &_
    "Comprimento, " &_
    "Cliente, " &_
    "DataEmissao, " &_
    "PrimeiraDataEntrega, " &_
    "DescricaoBloqueio, " &_
    "NumeroOrdemCompra, " &_
    "ItemOrdemCompra " &_
  "FROM Pedido " &_
  "WHERE Isnull(Status,0) = 0 " &_
    "AND Isnull(Excluido,0) = 0 " &_
    "AND QuemIncluiu = '" & IdUsuario & "' " &_
  "ORDER BY Pedido DESC "

select case Planta
case "05" set consulta = Conecta("BEN-CORPORATIVO", SQL )
case "06" set consulta = Conecta("FOR-CORPORATIVO", SQL )
case "07" set consulta = Conecta("SAQ-CORPORATIVO", SQL )
end select

if vartype(consulta) = 0 or vartype(consulta) = 1 then
  response.end
end if

temRegistros = NOT consulta.EOF
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet - Exclusão de Pedidos Digitados</title>

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
            --color-danger:        hsl(3, 90%, 58%);
            --color-warning:       hsl(35, 100%, 48%);

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
        }

        body {
            background-color: var(--color-bg);
            min-height: 100vh;
            font-family: var(--font-sans);
        }

        .page-title {
            font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
            font-weight: 600;
            letter-spacing: -0.01em;
            color: var(--color-text);
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

        .btn-outline-danger {
            font-weight: 500;
            font-size: .82rem;
            transition: background-color var(--transition), border-color var(--transition), color var(--transition);
        }

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

        .overlay-loading {
            position: absolute;
            inset: 0;
            background: hsla(0, 0%, 100%, .82);
            backdrop-filter: blur(2px);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            z-index: 10;
            gap: .85rem;
            border-radius: var(--radius-md);
        }
        .overlay-loading .spinner-border {
            width: 1.75rem;
            height: 1.75rem;
            border-width: 2px;
            color: var(--color-text-faint);
        }
        .overlay-loading span {
            font-size: .85rem;
            font-weight: 500;
            color: var(--color-text-muted);
        }

        .table-card {
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            position: relative;
            box-shadow: var(--shadow-sm);
        }

        .table-wrapper {
            overflow-x: auto;
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

        .form-control:focus,
        .form-select:focus {
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

    <script type="text/javascript">
        function abrirConfirmacaoExclusao() {
            const total = document.querySelectorAll('input[name="checados"]:checked').length;
            if (total === 0) {
                return;
            }
            document.getElementById('lblPedidoConfirmacao').textContent = total;
            const modal = new bootstrap.Modal(document.getElementById('modalConfirmaExclusao'));
            modal.show();
        }

        function confirmarExclusao() {
            document.getElementById('overlayLoading').classList.remove('d-none');
            document.getElementById('PedidoDigitado').submit();
        }
    </script>
</head>
<body>
    <div class="container-fluid py-3 px-3 px-lg-4">

        <div class="row">
            <div class="col-12 col-lg-8 offset-lg-2 d-flex align-items-center gap-2">
                <a href="#" onclick="history.back(); return false;" class="btn btn-light btn-sm">
                    <i class="bi bi-arrow-left me-1"></i>Voltar
                </a>
                <span class="page-title">Pedidos Digitados</span>
            </div>
        </div>

        <div class="row mt-2">
            <div class="col-12 col-lg-8 offset-lg-2">
                <a href="../selecionaPlanta.asp" class="btn btn-primary btn-sm">
                    <i class="bi bi-plus-lg me-1"></i>Emitir novo pedido
                </a>
            </div>
        </div>

        <form name="PedidoDigitado" id="PedidoDigitado" action="PedidoDigitado.asp?Planta=<%=Planta%>" method="post">

            <div class="row mt-3">
                <div class="col-12 col-lg-8 offset-lg-2">
                    <div class="table-card">
                        <div id="overlayLoading" class="overlay-loading d-none">
                            <div class="spinner-border" role="status">
                                <span class="visually-hidden">Carregando...</span>
                            </div>
                            <span>Excluindo pedido(s)...</span>
                        </div>
<%
                        if temRegistros then
%>
                        <div class="table-wrapper">
                            <table class="table table-bordered table-hover table-sm mb-0">
                                <thead>
                                    <tr>
                                        <th scope="col">
                                            <i class="bi bi-trash" style="color: var(--color-danger);" aria-hidden="true" title="Selecione para excluir!"></i>
                                        </th>
                                        <th>Pedido</th>
                                        <th>Cliente</th>
                                        <th>Nome</th>
                                        <th>Produto</th>
                                        <th>Referência</th>
                                        <th>Quant.</th>
                                        <th>Preço</th>
                                        <th>Emissão</th>
                                        <th>Entrega</th>
                                    </tr>
                                </thead>
                                <tbody>
<%
                            while NOT consulta.EOF

                              razao = ""

                              if Planta = "05" then 'BENTO GONCALVES --------------------------------------------------------------------------------'

                                 SQL = "" &_
                                  "SELECT " &_
                                    "NOME_CLIENTE As Razao " &_
                                  "FROM VW_CLIENTES_PEDIDO " &_
                                  "WHERE COD_CLIENTE = " & consulta("Cliente") & " "

                                SET BEN_DADOSADV = Conecta( "BEN-DADOSADV", SQL )

                                if vartype(BEN_DADOSADV) = 0 or vartype(BEN_DADOSADV) = 1 then
                                  response.end
                                end if

                                if NOT BEN_DADOSADV.EOF then
                                  razao = trim(BEN_DADOSADV("Razao"))
                                end if

                                BEN_DADOSADV.Close()
                                SET BEN_DADOSADV = NOTHING

                              end if

                              if Planta = "06" then 'FORTALEZA --------------------------------------------------------------------------------------'

                                SQL = "" &_
                                  "SELECT " &_
                                    "NOME_CLIENTE As Razao " &_
                                  "FROM VW_CLIENTES_PEDIDO " &_
                                  "WHERE COD_CLIENTE = " & consulta("Cliente") & " "

                                SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)

                                if vartype(FOR_DADOSADV) = 0 or vartype(FOR_DADOSADV) = 1 then
                                  response.end
                                end if

                                if NOT FOR_DADOSADV.EOF then
                                  razao = Trim(FOR_DADOSADV("Razao"))
                                end if

                                FOR_DADOSADV.Close()
                                set FOR_DADOSADV = NOTHING

                              end if

                              if Planta = "07" then 'SAQUAREMA --------------------------------------------------------------------------------------'

                                SQL = "" &_
                                  "SELECT " &_
                                    "NOME_CLIENTE AS Razao " &_
                                  "FROM VW_BI_CLIENTES " &_
                                  "WHERE COD_CLIENTE = "&consulta("Cliente")&" "

                                SET SAQ_DADOSADV = Conecta( "SAQ-DADOSADV", SQL )

                                if vartype(SAQ_DADOSADV) = 0 or vartype(SAQ_DADOSADV) = 1 then
                                  response.end
                                end if

                                if NOT SAQ_DADOSADV.EOF then
                                  razao = Trim(SAQ_DADOSADV("Razao"))
                                end if

                                SAQ_DADOSADV.Close()
                                SET SAQ_DADOSADV = NOTHING

                              end if
%>
                                    <tr class="<%=classe %>">
                                      <td>
                                        <input type="checkbox" class="form-check-input" name="checados"
                                               value="<%="'" & consulta("Pedido") & "'"%>"
                                               aria-label="Selecionar pedido <%=consulta("Pedido")%> para exclusão" />
                                      </td>
                                      <td><%=consulta("Pedido") %></td>
                                      <td><%=consulta("Cliente") %></td>
                                      <td><%=razao %></td>
                                      <td><%=Trim(consulta("Produto")) %></td>
                                      <td><%=Trim(consulta("Referencia"))%></td>
                                      <td><%=consulta("Quantidade") %></td>
                                      <td><%=FormatNumber(consulta("PrecoUnitario"),6) %></td>
                                      <td><%=consulta("DataEmissao") %></td>
                                      <td><%=consulta("PrimeiraDataEntrega") %></td>
                                    </tr>
<%
                          response.Flush()
                          consulta.MoveNext()

                        wend

                        consulta.Close()
                        SET consulta = NOTHING
%>
                                </tbody>
                            </table>
                        </div>

                        <div class="d-flex justify-content-center gap-2 p-3">
                            <button type="button" class="btn btn-danger btn-sm" onclick="abrirConfirmacaoExclusao()">
                                <i class="bi bi-trash me-1"></i>Excluir Pedidos Selecionados
                            </button>
                        </div>
<%
                        else
%>
                        <div class="table-placeholder">
                            <i class="bi bi-clipboard-data"></i>
                            <p>Nenhum pedido digitado encontrado para esta planta.</p>
                        </div>
<%
                        end if
%>
                    </div>
                </div>
            </div>
        </form>
    </div>

    <!-- Modal de confirmação de exclusão -->
    <div class="modal fade" id="modalConfirmaExclusao" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Confirmar exclusão</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <p class="mb-0">
                        Tem certeza que deseja excluir os <strong id="lblPedidoConfirmacao"></strong> pedido(s) selecionado(s)? Esta ação não pode ser desfeita.
                    </p>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-danger btn-sm" data-bs-dismiss="modal" onclick="confirmarExclusao()">
                        <i class="bi bi-trash3 me-1"></i>Sim, excluir pedidos
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
</body>
</html>
