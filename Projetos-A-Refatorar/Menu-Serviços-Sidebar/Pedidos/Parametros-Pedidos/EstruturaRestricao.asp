<!--#Include File="../../cnn2/cnn2.asp"-->
<%
  Session.LCID = 1046
  msgSuccess = ""
  msgError = ""

  if trim(request("acao")) = "1" then

    SQL = "" &_
      "INSERT INTO EstruturaRestricao ( " &_
        "Estrutura, " &_
        "LarguraDe, " &_
        "LarguraAte, " &_
        "Quem, " &_
        "Quando, " &_
        "Tipo " &_
      ") VALUES( " &_
        "'" & trim(request("estrutura")) & "', " &_
        trim(request("larguraDe")) & ", " &_
        trim(request("larguraAte")) & ", " &_
        "'" & usr & "', " &_
        "'" & inverteData(Date()) & "', " &_
        "'" & trim(request("tipo")) & "' ) "

    Set Dados11 = Conecta("PIR-CORPORATIVO", SQL)

    SQL = "" &_
      "SELECT " &_
        "Estrutura " &_
      "FROM EstruturaRestricao " &_
      "WHERE Isnull(Excluido,0) = 0 " &_
        "AND Estrutura = '"&trim(request("estrutura"))&"' " &_
        "AND LarguraDe = "&trim(request("larguraDe"))&" " &_
        "AND LarguraAte = "&trim(request("larguraAte"))&" " &_
        "AND Tipo = '"&trim(request("tipo"))&"' "

      Set Dados11 = Conecta("PIR-CORPORATIVO", SQL)

      if (NOT Dados11.EOF) then
        msgSuccess = "Registro inserido com sucesso!"
      else
        msgError = "Houve um erro, não foi possível inserir o registro."
      end if

      Dados11.Close()

  end if

  if trim(request("acao")) = "2" and Trim(Request("id")) <> "" then

    SQL = "" &_
      "UPDATE EstruturaRestricao SET " &_
        "Excluido = 1, " &_
        "Quem = '" & usr & "', " &_
        "Quando = '"&inverteData(Date())&"' " &_
      "WHERE Id = "&Trim(request("id"))

    Set Dados11 = Conecta("PIR-CORPORATIVO", SQL)

    SQL = "" &_
      "SELECT " &_
        "Estrutura " &_
      "FROM EstruturaRestricao " &_
      "WHERE Id = " & Trim(request("id")) & " " &_
        "AND Isnull(Excluido,0) = 0 "

    Set Dados11 = Conecta("PIR-CORPORATIVO", SQL)

    if Dados11.Eof then
      msgSuccess = "Registro excluído com sucesso!"
    else
      msgError = "Houve um erro, não foi possível excluir o registro."
    end if

    Dados11.Close()

  end if

  SQL = "SELECT " &_
    "Id, " &_
    "Estrutura, " &_
    "LarguraDe, " &_
    "ISNULL(LarguraAte,0) LarguraAte, " &_
    "Quem, " &_
    "Quando, " &_
    "Tipo " &_
  "FROM EstruturaRestricao " &_
  "WHERE Isnull(Excluido,0) = 0 "

  Set Dados11 = Conecta("PIR-CORPORATIVO", SQL)

  SQL = "" &_
    "SELECT " &_
      "SUBSTRING(RTRIM(LTRIM(Z1_ESTRUT)), 1, LEN(RTRIM(LTRIM(Z1_ESTRUT)))-1)+'A' AS ESTRUTURA " &_
    "FROM SZ1010 " &_
    "WHERE (Z1_STATUS IN ('I','T')) " &_
      "AND D_E_L_E_T_ = '' " &_
    "UNION " &_
    "SELECT " &_
      "SUBSTRING(RTRIM(LTRIM(Z1_ESTRUT)),1,LEN(RTRIM(LTRIM(Z1_ESTRUT)))-1)+'B' AS ESTRUTURA " &_
    "FROM SZ1010 " &_
    "WHERE (Z1_STATUS IN ('I','T')) " &_
      "AND D_E_L_E_T_ = '' "

  Set Dados2 = Conecta("PIR-DADOSADV", SQL)
%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <title>Restrições por Estrutura</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />

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

        /* — Accent — */
        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        /* — Estado — */
        --color-success:       hsl(145, 63%, 32%);
        --color-success-soft:  hsl(145, 60%, 95%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);
        --color-warning:       hsl(35, 100%, 48%);
        --color-warning-soft:  hsl(40, 100%, 95%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
        --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

        --radius-sm: 6px;
        --radius-md: 10px;
        --radius-lg: 14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-1: 4px;
        --space-2: 8px;
        --space-3: 12px;
        --space-4: 16px;
        --space-5: 24px;

        --z-toast: 1100;
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

      /* — Card de formulário (reaproveita o padrão de Card de Filtros) — */
      .form-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
      }
      .form-card-header {
        background-color: var(--color-surface);
        color: var(--color-text);
        padding: .7rem 1.1rem;
        border-radius: var(--radius-md) var(--radius-md) 0 0;
        border-bottom: 1px solid var(--color-border);
        font-weight: 600;
        font-size: .85rem;
      }
      .form-card-body {
        padding: 1.1rem 1.25rem 1.35rem;
      }

      .form-label-quiet {
        font-size: .78rem;
        font-weight: 500;
        color: var(--color-text-muted);
        letter-spacing: .01em;
        margin-bottom: .3rem;
      }

      .form-control:focus,
      .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px var(--color-accent-soft);
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

      .btn-danger-quiet {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        color: var(--color-danger);
        font-weight: 500;
        font-size: .8rem;
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-danger-quiet:hover {
        background-color: var(--color-danger-soft);
        border-color: var(--color-danger);
      }

      /* — Card de tabela — */
      .table-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        position: relative;
      }
      .table-card-header {
        padding: .7rem 1.1rem;
        border-bottom: 1px solid var(--color-border);
        font-weight: 600;
        font-size: .85rem;
        color: var(--color-text);
      }

      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 50vh;
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

      .table tbody tr:hover td {
        background-color: var(--color-surface-muted);
      }

      .table-placeholder {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 3rem 1rem;
        color: var(--color-text-muted);
        gap: .85rem;
      }
      .table-placeholder .bi { font-size: 2.2rem; color: var(--color-text-faint); }
      .table-placeholder p  { margin: 0; font-size: .88rem; text-align: center; }

      /* — Rodapé da tabela (page size · paginação · total) — */
      .table-footer {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: .7rem 1.25rem;
        border-top: 1px solid var(--color-border);
        flex-wrap: wrap;
        gap: .5rem;
      }
      .table-footer-left {
        min-width: 160px;
        display: flex;
        align-items: center;
        gap: .5rem;
      }
      .table-footer-center {
        flex: 1;
        display: flex;
        justify-content: center;
      }
      .table-footer-right {
        min-width: 140px;
        text-align: right;
        padding-right: .25rem;
      }
      .info-registros { font-size: .8rem; color: var(--color-text-muted); }
      .select-limite  { width: auto; }

      #paginacao .page-link {
        color: var(--color-text-muted);
        border-color: var(--color-border);
        font-size: .82rem;
        transition: background-color var(--transition), border-color var(--transition), color var(--transition);
      }
      #paginacao .page-item.active .page-link {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
        color: var(--color-surface);
      }
      #paginacao .page-link:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
        color: var(--color-text);
      }
      #paginacao .page-item.disabled .page-link {
        color: var(--color-text-faint);
      }

      /* — Overlay de carregamento (envio do formulário) — */
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

      /* — Toast — */
      .toast-container-custom {
        position: fixed; top: 1rem; right: 1rem;
        z-index: var(--z-toast);
        display: flex; flex-direction: column; gap: .5rem;
      }
      .toast-custom {
        min-width: 280px; max-width: 360px;
        border-radius: var(--radius-md); overflow: hidden;
        box-shadow: var(--shadow-md);
        background: var(--color-surface); border: 1px solid var(--color-border);
        animation: toastSlideIn .22s ease;
      }
      @keyframes toastSlideIn {
        from { opacity: 0; transform: translateX(24px); }
        to   { opacity: 1; transform: translateX(0); }
      }
      .toast-custom .toast-progress-bar {
        animation: toastProgressShrink 4.5s linear forwards;
      }
      @keyframes toastProgressShrink {
        from { width: 100%; } to { width: 0%; }
      }

      /* — Modal de confirmação — */
      .modal-content {
        border-radius: var(--radius-lg);
        border: 1px solid var(--color-border);
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
    <div class="container-fluid py-3 px-3 px-lg-4" style="max-width: 760px">

      <span class="page-title d-block mb-3">Cadastro de Restrições por Estrutura</span>

      <form action="EstruturaRestricao.asp" name="frm" id="form" method="post">

        <input type="hidden" name="acao" id="acao" value="0" />
        <input type="hidden" name="id" id="id" value="" />

        <div class="form-card mb-4">
          <div class="form-card-header">
            <i class="bi bi-sliders me-2" style="color: var(--color-text-muted);"></i>Nova restrição
          </div>
          <div class="form-card-body">

            <div class="row g-3">

              <div class="col-12 col-md-6">
                <label for="tipo" class="form-label-quiet">Tipo</label>
                <select name="tipo" id="tipo" class="form-select">
                  <option value="">Selecione</option>
                  <option value="CX">Caixa</option>
                  <option value="CH">Chapa</option>
                </select>
              </div>

              <div class="col-12 col-md-6">
                <label for="estrutura" class="form-label-quiet">Estrutura</label>
                <select name="estrutura" id="estrutura" class="form-select">
                  <option value="">Selecione</option>
<%
      while not Dados2.Eof
%>
                  <option value="<%=Trim(Dados2("ESTRUTURA")) %>" <%if Trim(Dados2("ESTRUTURA")) = Trim(Request("estrutura")) then%>selected<%end if %>><%=Trim(Dados2("ESTRUTURA")) %></option>
<%
        Dados2.MoveNext()
      wend
      Dados2.Close()
%>
                </select>
              </div>

              <div class="col-12 col-md-6">
                <label for="larguraDe" class="form-label-quiet">Largura de</label>
                <input type="number" name="larguraDe" id="larguraDe" min="0" max="2500" value="0" class="form-control" />
              </div>

              <div class="col-12 col-md-6">
                <label for="larguraAte" class="form-label-quiet">Largura até</label>
                <input type="number" name="larguraAte" id="larguraAte" min="0" max="2500" value="0" class="form-control" />
              </div>

            </div>

            <div class="d-flex justify-content-end mt-4">
              <button type="button" id="btnIncluir" class="btn btn-primary btn-sm">
                <i class="bi bi-plus-lg me-2"></i>Adicionar restrição
              </button>
            </div>

          </div>
        </div>

        <div class="table-card">
          <div class="table-card-header">
            <i class="bi bi-table me-2" style="color: var(--color-text-muted);"></i>Restrições cadastradas
          </div>

          <div class="table-wrapper">
            <table class="table table-bordered table-hover table-sm mb-0">
              <thead>
                <tr>
                  <th>Tipo</th>
                  <th>Estrutura</th>
                  <th>Largura de</th>
                  <th>Largura até</th>
                  <th>Quem</th>
                  <th>Quando</th>
                  <th>Ação</th>
                </tr>
              </thead>
              <tbody id="tbodyRestricoes">
<%
      hasRows = not Dados11.Eof
      while not Dados11.Eof
%>
                <tr>
                  <td><%=Dados11("Tipo") %></td>
                  <td><%=Dados11("Estrutura") %></td>
                  <td><%=Dados11("LarguraDe") %></td>
                  <td><%=Dados11("LarguraAte") %></td>
                  <td><%=Dados11("Quem") %></td>
                  <td><%=AcertaData(Dados11("Quando")) %></td>
                  <td>
                    <button type="button" class="btn btn-danger-quiet btn-sm btn-excluir" data-id="<%=Dados11("Id") %>" title="Excluir">
                      <i class="bi bi-trash3"></i>
                    </button>
                  </td>
                </tr>
<%
        Dados11.MoveNext()
      wend
      Dados11.Close()
%>
              </tbody>
            </table>
<%
      if not hasRows then
%>
            <div class="table-placeholder">
              <i class="bi bi-clipboard-data"></i>
              <p>Nenhuma restrição cadastrada até o momento.</p>
            </div>
<%
      end if
%>
          </div>

          <div class="table-footer">

            <div class="table-footer-left">
              <label for="selectLimite" class="form-label mb-0" style="font-size:.8rem;white-space:nowrap;color:var(--color-text-muted);">
                Linhas por página:
              </label>
              <select id="selectLimite" class="form-select form-select-sm select-limite">
                <option value="5">5</option>
                <option value="10" selected>10</option>
                <option value="25">25</option>
                <option value="50">50</option>
              </select>
            </div>

            <div class="table-footer-center">
              <nav aria-label="Navegação de páginas">
                <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
              </nav>
            </div>

            <div class="table-footer-right">
              <span id="infoRegistros" class="info-registros"></span>
            </div>

          </div>

          <div id="overlayLoading" class="overlay-loading d-none">
            <div class="spinner-border" role="status">
              <span class="visually-hidden">Processando...</span>
            </div>
            <span>Processando...</span>
          </div>
        </div>

      </form>
    </div>

    <!-- Modal de confirmação de exclusão -->
    <div class="modal fade" id="modalConfirmaExclusao" tabindex="-1" aria-labelledby="modalConfirmaExclusaoLabel" aria-hidden="true">
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="modalConfirmaExclusaoLabel">Excluir restrição</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">
            <p class="mb-0">Tem certeza que deseja excluir esta restrição de estrutura? Esta ação não pode ser desfeita.</p>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Cancelar</button>
            <button type="button" id="btnConfirmaExclusao" class="btn btn-danger-quiet btn-sm">
              <i class="bi bi-trash3 me-2"></i>Sim, excluir restrição
            </button>
          </div>
        </div>
      </div>
    </div>

    <!-- Toast container -->
    <div class="toast-container-custom" id="toastContainer" aria-live="polite"></div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <script>
      const MSG_SUCCESS = "<%=Replace(msgSuccess, """", "'") %>";
      const MSG_ERROR   = "<%=Replace(msgError, """", "'") %>";

      function exibirToast(tipo, texto) {
        const icones = {
          sucesso: 'bi-check-circle-fill',
          erro:    'bi-x-circle-fill',
          aviso:   'bi-exclamation-triangle-fill',
          info:    'bi-info-circle-fill'
        };
        const cores = {
          sucesso: 'var(--color-success)',
          erro:    'var(--color-danger)',
          aviso:   'var(--color-warning)',
          info:    'var(--color-text-muted)'
        };

        const $toast = $(`
          <div class="toast-custom">
            <div class="d-flex align-items-center px-3 py-2">
              <i class="bi ${icones[tipo]} me-2" style="color:${cores[tipo]};"></i>
              <span class="flex-grow-1" style="font-size:.85rem;">${texto}</span>
              <button type="button" class="btn-close ms-2" style="font-size:.7rem;" aria-label="Fechar"></button>
            </div>
            <div class="toast-progress-bar" style="height:3px;background:${cores[tipo]};"></div>
          </div>
        `);

        $toast.find('.btn-close').on('click', () => $toast.remove());
        $('#toastContainer').append($toast);

        setTimeout(() => $toast.remove(), 4500);
      }

      let paginaAtual = 1;
      let limite = 10;
      let linhasRestricoes = [];

      function _renderPaginacao(paginaAtual, totalPaginas) {
        const $ul = $('#paginacao').empty();
        if (totalPaginas <= 1) return;

        const mkItem = (label, page, disabled, active) => {
          const $li = $('<li>').addClass('page-item');
          if (disabled) $li.addClass('disabled');
          if (active)   $li.addClass('active');
          const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
          if (!disabled && !active) $a.on('click', e => { e.preventDefault(); _irParaPagina(page); });
          return $li.append($a);
        };

        $ul.append(mkItem('&laquo;', 1,              paginaAtual === 1, false));
        $ul.append(mkItem('&lsaquo;', paginaAtual - 1, paginaAtual === 1, false));

        const inicio = Math.max(1, paginaAtual - 2);
        const fim    = Math.min(totalPaginas, paginaAtual + 2);
        for (let p = inicio; p <= fim; p++) {
          $ul.append(mkItem(p, p, false, p === paginaAtual));
        }

        $ul.append(mkItem('&rsaquo;', paginaAtual + 1, paginaAtual === totalPaginas, false));
        $ul.append(mkItem('&raquo;', totalPaginas,    paginaAtual === totalPaginas, false));
      }

      function _aplicarPaginacao() {
        const total = linhasRestricoes.length;
        const totalPaginas = Math.max(1, Math.ceil(total / limite));

        if (paginaAtual > totalPaginas) paginaAtual = totalPaginas;

        const inicio = (paginaAtual - 1) * limite;
        const fim = inicio + limite;

        linhasRestricoes.forEach(function (linha, indice) {
          linha.style.display = (indice >= inicio && indice < fim) ? '' : 'none';
        });

        _renderPaginacao(paginaAtual, totalPaginas);

        $('#infoRegistros').text(`Total: ${total} registro(s)`);
      }

      function _irParaPagina(pagina) {
        paginaAtual = pagina;
        _aplicarPaginacao();
      }

      function validaForm() {
        let msg = '';

        if ($('#tipo').val() === '') {
          msg += 'Por favor, informe o tipo de produto.\n';
        }

        if ($('#estrutura').val() === '') {
          msg += 'Por favor, informe a estrutura.\n';
        }

        const larguraDe  = $('#larguraDe').val();
        const larguraAte = $('#larguraAte').val();

        if (larguraDe === '' || larguraAte === '') {
          msg += 'Por favor, informe a Largura De e a Largura Até.\n';
        } else if ((larguraDe === '0' && larguraAte === '0')) {
          msg += 'É necessário informar, pelo menos, a Largura De.\n';
        }

        if (msg !== '') {
          exibirToast('aviso', msg.replace(/\n/g, '<br>'));
          $('#id').val('');
          $('#acao').val('');
          return false;
        }

        return true;
      }

      $(document).ready(function () {

        linhasRestricoes = $('#tbodyRestricoes tr').toArray();
        limite = parseInt($('#selectLimite').val(), 10);
        _aplicarPaginacao();

        $('#selectLimite').on('change', function () {
          limite = parseInt($(this).val(), 10);
          paginaAtual = 1;
          _aplicarPaginacao();
        });

        if (MSG_SUCCESS !== '') {
          exibirToast('sucesso', MSG_SUCCESS);
        }
        if (MSG_ERROR !== '') {
          exibirToast('erro', MSG_ERROR);
        }

        $('#btnIncluir').on('click', function () {
          $('#id').val('');
          $('#acao').val('1');

          if (validaForm()) {
            $('#overlayLoading').removeClass('d-none');
            $('#form').trigger('submit');
          }
        });

        let idParaExcluir = '';
        const modalExclusao = new bootstrap.Modal(document.getElementById('modalConfirmaExclusao'));

        $(document).on('click', '.btn-excluir', function () {
          idParaExcluir = $(this).data('id');
          modalExclusao.show();
        });

        $('#btnConfirmaExclusao').on('click', function () {
          $('#id').val(idParaExcluir);
          $('#acao').val('2');
          modalExclusao.hide();
          $('#overlayLoading').removeClass('d-none');
          $('#form').trigger('submit');
        });

      });
    </script>
  </body>
</html>
