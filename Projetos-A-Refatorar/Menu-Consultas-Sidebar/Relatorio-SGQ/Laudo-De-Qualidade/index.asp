<%session.lcid = 1046%>

<!--#include File = "../cnn2/cnn2.asp"-->
<!--#include File = "../cnn2/conexoes.asp"-->
<!--#include file = "../funcoes/funcoes.asp"-->
<!--#include file = "../includes/variaveis.asp" -->

<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <title>SWnet — Consulta Laudos</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <!-- Bootstrap 5.3 -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" />
    <!-- Bootstrap Icons -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

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

        /* — Accent — */
        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        /* — Estado — */
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-hover:  hsl(3, 90%, 50%);

        /* — Tipografia — */
        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

        /* — Raio — */
        --radius-sm:  6px;
        --radius-md:  10px;
        --radius-lg:  14px;

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
        --space-6: 32px;
      }

      /* — Base — */
      body {
        background-color: var(--color-bg);
        color: var(--color-text);
        font-family: var(--font-sans);
        font-size: .9375rem;
      }

      /* — Card principal — */
      .page-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-lg);
        box-shadow: var(--shadow-md);
        overflow: hidden;
      }

      /* — Cabeçalho do card — */
      .page-card-header {
        background-color: var(--color-surface-muted);
        border-bottom: 1px solid var(--color-border);
        padding: var(--space-4) var(--space-5);
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: var(--space-3);
      }

      .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.125rem);
        font-weight: 600;
        color: var(--color-text);
        margin: 0;
        letter-spacing: -.015em;
      }

      /* — Corpo do card — */
      .page-card-body {
        padding: var(--space-5);
      }

      /* — Inputs e selects — */
      .form-label {
        font-size: .8125rem;
        font-weight: 500;
        color: var(--color-text-muted);
        margin-bottom: var(--space-2);
      }

      .form-control,
      .form-select {
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        background-color: var(--color-surface);
        color: var(--color-text);
        font-size: .9375rem;
        transition: border-color var(--transition), box-shadow var(--transition);
        height: 40px;
      }

      .form-control::placeholder {
        color: var(--color-text-faint);
      }

      .form-control:focus,
      .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .18);
        outline: none;
      }

      /* — Botão primário (Pesquisar) — */
      .btn-primary-sw {
        display: inline-flex;
        align-items: center;
        gap: var(--space-2);
        padding: 0 var(--space-5);
        height: 40px;
        border-radius: var(--radius-sm);
        border: none;
        background-color: var(--color-accent);
        color: hsl(0, 0%, 100%);
        font-size: .9375rem;
        font-weight: 500;
        cursor: pointer;
        transition: background-color var(--transition);
      }

      .btn-primary-sw:hover {
        background-color: var(--color-accent-hover);
      }

      .btn-primary-sw:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .3);
      }

      /* — Botão Voltar — */
      .btn-ghost {
        display: inline-flex;
        align-items: center;
        gap: var(--space-2);
        padding: 0 var(--space-5);
        height: 40px;
        border-radius: var(--radius-sm);
        border: 1px solid var(--color-border);
        background-color: var(--color-surface);
        color: var(--color-text-muted);
        font-size: .9375rem;
        font-weight: 500;
        cursor: pointer;
        text-decoration: none;
        transition: border-color var(--transition), color var(--transition), background-color var(--transition);
      }

      .btn-ghost:hover {
        border-color: var(--color-border-strong);
        color: var(--color-text);
        background-color: var(--color-surface-muted);
      }

      .btn-ghost:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .25);
      }

      /* — Tabela de resultados — */
      .results-table-wrapper {
        overflow-x: auto;
      }

      .results-table {
        width: 100%;
        border-collapse: collapse;
        font-size: .875rem;
      }

      .results-table thead th {
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        font-weight: 600;
        font-size: .75rem;
        text-transform: uppercase;
        letter-spacing: .06em;
        padding: var(--space-3) var(--space-4);
        text-align: center;
        border-bottom: 1px solid var(--color-border);
        white-space: nowrap;
      }

      .results-table thead th:first-child {
        text-align: left;
      }

      .results-table tbody td {
        padding: var(--space-3) var(--space-4);
        text-align: center;
        color: var(--color-text);
        border-bottom: 1px solid var(--color-border);
        vertical-align: middle;
      }

      .results-table tbody td:first-child {
        text-align: left;
        color: var(--color-text);
      }

      .results-table tbody tr:last-child td {
        border-bottom: none;
      }

      .results-table tbody tr:hover td {
        background-color: var(--color-surface-muted);
      }

      /* — Link de OP — */
      .op-link {
        color: var(--color-accent);
        text-decoration: none;
        font-weight: 500;
        transition: color var(--transition);
      }

      .op-link:hover {
        color: var(--color-accent-hover);
        text-decoration: underline;
      }

      /* — Rodapé do card — */
      .page-card-footer {
        border-top: 1px solid var(--color-border);
        background-color: var(--color-surface-muted);
        padding: var(--space-4) var(--space-5);
        display: flex;
        justify-content: center;
      }

      /* — Spinner do botão Pesquisar — */
      @keyframes sw-spin {
        to { transform: rotate(360deg); }
      }

      .btn-primary-sw .btn-spinner {
        display: none;
        width: 14px;
        height: 14px;
        border: 2px solid hsla(0, 0%, 100%, .35);
        border-top-color: hsl(0, 0%, 100%);
        border-radius: 50%;
        animation: sw-spin 600ms linear infinite;
        flex-shrink: 0;
      }

      .btn-primary-sw.is-loading .btn-spinner { display: block; }
      .btn-primary-sw.is-loading .btn-icon    { display: none; }
      .btn-primary-sw.is-loading              { opacity: .8; pointer-events: none; }

      /* — Divisórias verticais na tabela de resultados — */
      .results-table thead th + th,
      .results-table tbody td + td {
        border-left: 1px solid var(--color-border);
      }

      /* — Reduced motion — */
      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after { transition-duration: 0ms !important; }
        .btn-spinner { animation: none !important; }
      }

      /* — Mobile — */
      @media (max-width: 768px) {
        .page-card-body { padding: var(--space-4); }
        .page-card-header { padding: var(--space-3) var(--space-4); }
        .page-card-footer { padding: var(--space-3) var(--space-4); }
      }
    </style>
  </head>

  <body class="d-flex align-items-center justify-content-center min-vh-100 py-5 px-3">

<%
if (trim(request("serie")) <> "" and trim(request("nota")) <> "" and trim(SESSION("TIPO")) <> "") then

  Dim vSerie

  vSerie = trim(request("serie"))

  'TRIMBOX
  if (vSerie="001") or (vSerie="002") then
    filial = iif( (vSerie="001"), "00", "05" )
  end if

  'TRIMBOXUB
  if (vSerie="003") or (vSerie="004") then
    filial = iif( (vSerie="003"), "04", "06" )
  end if

  SQL = "" &_
    "SELECT " &_
      "r.Pedido, A1_CGC, r.OP, ISNULL(r.Referencia,'') Referencia, ISNULL(r.Composicao,'') Composicao, ISNULL(r.AlturaInterna,'') AlturaInterna, ISNULL(r.ComprimentoInterno,'') ComprimentoInterno, ISNULL(r.LarguraInterna,'') LarguraInterna, ISNULL(r.Gramatura,'') Gramatura, ISNULL(r.ColunaMinimo,'') ColunaMinimo, " &_
      "ISNULL(r.CobbIntMaximo,'') CobbIntMaximo, ISNULL(r.CobbExtMaximo,'') CobbExtMaximo, ISNULL(r.Espessura,'') Espessura, ISNULL(r.MullenMinimo,'') MullenMinimo, ISNULL(r.CrushMinimo,'') CrushMinimo, ISNULL(r.TCrush,'') TCrush, ISNULL(r.TColuna,'') TColuna, ISNULL(r.TCobbInt,'') TCobbInt, ISNULL(r.TCobbExt,'') TCobbExt, ISNULL(r.TEspessura,'') TEspessura, " &_
      "ISNULL(r.TMullen,'') TMullen, ISNULL(r.TGramatura,'') TGramatura, ISNULL(r.TComprimentoInterno,'') TComprimentoInterno, ISNULL(r.TLarguraInterna,'') TLarguraInterna, ISNULL(r.TAlturaInterna,'') TAlturaInterna, ISNULL(r.DPGramatura,'') DPGramatura, ISNULL(r.DPColuna,'') DPColuna, ISNULL(r.Responsavel1,'') Responsavel1, ISNULL(r.NotaFiscal,'') NotaFiscal, " &_
      "ISNULL(r.QtdNotaFiscal,'') QtdNotaFiscal, ISNULL(r.Responsavel2,'') Responsavel2, ISNULL(c.Mullen,'') Mullen, ISNULL(c.CobbExt1,'') CobbExt1, ISNULL(c.CobbInt1,'') CobbInt1, ISNULL(c.Coluna,'') Coluna, ISNULL(c.Crush,'') Crush, ISNULL(c.Gramatura,'') Gramatura, ISNULL(c.Espessura,'') Espessura, ISNULL(le.Municipio,'') Municipio, " &_
      "ISNULL(r.TCompressao,'') TCompressao, ISNULL(r.Compressao,'') Compressao, ISNULL(p.Lap,'') Lap, ISNULL(p.Item,'') Item, ISNULL(r.CodigoBarras,'') CodigoBarras, ISNULL(r.TUmidade,'') TUmidade, ISNULL(r.DataFabricacao,'') DataFabricacao, ISNULL(r.DataCriacao,'') DataCriacao, " &_
      "ISNULL(p.Obs7,'') Obs7, ISNULL(r.ObsOF,'') ObsOF, ISNULL(p.Chapa,'') Chapa, ISNULL(c.CobbInt0,'') CobbInt0, ISNULL(p.ResinaInterna,'') ResinaInterna, ISNULL(c.CobbExt0,'') CobbExt0, ISNULL(p.ResinaExterna,'') ResinaExterna, ISNULL(cl.CodCliente,'') CodCliente, ISNULL(l.RazaoSocial,'') RazaoSocial, ISNULL(c.Onda,'') Onda " &_
    "FROM rellaudo r " &_
      "INNER JOIN composic c ON r.Composicao=c.Composicao " &_
      "INNER JOIN Pedidos p ON (r.Pedido=p.Pedido) AND (r.Item=p.Item) " &_
      "LEFT OUTER JOIN vw_LocaisEntrega le ON (p.IDCliente=le.IDCliente) AND (p.IDLoja=le.IDLoja) AND (p.IDLocalEntrega=le.IDLocalEntrega) " &_
      "INNER JOIN Clientes cl ON p.IDCliente=cl.IDCliente " &_
      "LEFT OUTER JOIN Lojas l ON (p.IDCliente=l.IDCliente) AND (p.IDLoja=l.IDLoja) " &_
      "INNER JOIN " & cnx(skCnx("PIR-DADOSADV")).getServer() & "." & cnx(skCnx("PIR-DADOSADV")).getDataBase() & ".dbo.SF2010 SF2 ON r.NotaFiscal = F2_DOC AND cl.CodCliente = F2_CLIENTE COLLATE Latin1_General_BIN " &_
      "INNER JOIN " & cnx(skCnx("PIR-DADOSADV")).getServer() & "." & cnx(skCnx("PIR-DADOSADV")).getDataBase() & ".dbo.SA1010 SA1 ON F2_CLIENTE = A1_COD AND F2_LOJA = A1_LOJA " &_
    "WHERE SF2.D_E_L_E_T_ = '' " &_
      "AND SA1.D_E_L_E_T_ = '' " &_
      "AND F2_FILIAL = '"&trim(filial)&"' " &_
      iif( trim(request("nota")) <> "", "AND NotaFiscal = '"&trim(request("nota"))&"' ", "" ) &_
      iif( trim(CRMTIPO) = "C", "AND A1_CGC = '"&USR&"' ", "" ) &_
      iif( NIVEL = 2, "AND A1_GRPVEN = '"&CODUSER&"' ", "" ) &_
      "AND 1=1 " &_
    "ORDER BY NotaFiscal, r.OP"

  'TRIMBOX
  if (vSerie="001") or (vSerie="002") then
    SET TRIMBOX = Conecta("PIR-TRIMBOX", SQL )
  end if

  'TRIMBOXUB
  if (vSerie="003") or (vSerie="004") then
    SET TRIMBOX = Conecta("UBE-TRIMBOX", SQL )
  end if
%>

    <!-- ===== TELA DE RESULTADOS ===== -->
    <div style="width: 100%; max-width: 860px;">
      <div class="page-card">

        <div class="page-card-header">
          <h1 class="page-title">Consulta Laudos</h1>
        </div>

        <div class="results-table-wrapper">
          <table class="results-table">
            <thead>
              <tr>
                <th style="color:hsl(240, 2%, 42%);">Cliente</th>
                <th style="color:hsl(240, 2%, 42%);">O.P.</th>
                <th style="color:hsl(240, 2%, 42%);">N.F.</th>
                <th style="color:hsl(240, 2%, 42%);">Pedido</th>
                <th style="color:hsl(240, 2%, 42%);">Qtde.</th>
              </tr>
            </thead>
            <tbody>
<%
              while not TRIMBOX.eof
%>
              <tr>
                <td><%=TRIMBOX("CodCliente")&" — "&TRIMBOX("RazaoSocial") %></td>
                <td>
                  <a class="op-link" href="laudo-envio.asp?id1=<%=trim(request("nota")) %>&id2=<%=trim(request("serie")) %>&id3=<%=TRIMBOX("OP") %>" target="_blank">
                    <%=TRIMBOX("OP") %>
                  </a>
                </td>
                <td><%=TRIMBOX("NotaFiscal") %></td>
                <td><%=TRIMBOX("Pedido") %></td>
                <td><%=TRIMBOX("QtdNotaFiscal") %></td>
              </tr>
<%
                TRIMBOX.MoveNext()
              wend
              TRIMBOX.Close()
              SET TRIMBOX = Nothing
%>
            </tbody>
          </table>
        </div>

        <div class="page-card-footer">
          <button type="button" class="btn-ghost"
           style="color:black;" 
           onclick="history.go(-1)">
            <i class="bi bi-arrow-left"></i> Voltar
          </button>
        </div>

      </div>
    </div>

<%
else
%>

    <!-- ===== TELA DE FILTRO ===== -->
    <div style="width: 100%; max-width: 460px;">
      <div class="page-card">

        <div class="page-card-header">
          <h1 class="page-title">Consulta Laudos</h1>
        </div>

        <div class="page-card-body">
          <form action="index.asp" name="consulta" method="get">
            <div class="row g-3">
              <div class="col-6">
                <label for="nota" class="form-label">Nota Fiscal</label>
                <input
                  type="text"
                  name="nota"
                  id="nota"
                  maxlength="6"
                  class="form-control"
                  placeholder="Ex.: 123456"
                  value="<%=trim(request("nota")) %>"
                />
              </div>
              <div class="col-6">
                <label for="serie" class="form-label">Série</label>
                <select name="serie" id="serie" class="form-select">
                  <option value=""    <%if trim(request("serie")) = "" then    %>selected<%end if %>>Selecione...</option>
                  <option value="001" <%if trim(request("serie")) = "001" then %>selected<%end if %>>001</option>
                  <option value="002" <%if trim(request("serie")) = "002" then %>selected<%end if %>>002</option>
                  <option value="003" <%if trim(request("serie")) = "003" then %>selected<%end if %>>003</option>
                  <option value="004" <%if trim(request("serie")) = "004" then %>selected<%end if %>>004</option>
                </select>
              </div>
            </div>

            <div class="d-flex justify-content-center mt-4">
              <button type="submit" name="ok" id="btnPesquisar" class="btn-primary-sw">
                <span class="btn-spinner" aria-hidden="true"></span>
                <i class="bi bi-search btn-icon"></i> Pesquisar
              </button>
            </div>
          </form>
        </div>

      </div>
    </div>

<%
end if
%>

    <!-- Bootstrap 5 JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- jQuery 4 -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <script src="../funcoes/funcoesJAVASCRIPT.js"></script>

    <script>
      $(function () {
        $('form[name="consulta"]').on('submit', function () {
          $('#btnPesquisar').addClass('is-loading');
        });
      });
    </script>
  </body>
</html>
