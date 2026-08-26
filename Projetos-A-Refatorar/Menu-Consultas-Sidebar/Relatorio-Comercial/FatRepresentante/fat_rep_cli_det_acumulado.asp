<%RESPONSE.BUFFER=TRUE%>
<!--#include virtual="/corporativo/cnn2/cnn2.asp" -->
<!--#include virtual="/corporativo/includes/variaveis.asp" -->
<%
'*******************************
'***    Validando Acesso     ***
'*******************************

If Trim(USR) = "" Then
%>
<script>alert("Acesso Negado.\nVoce nao tem permissao para esse modulo."); window.close();</script>
<%
  Response.End
End If

If CInt(NIVEL) <> 1 Then
%>
<script>alert("Acesso Negado.\nConsulta disponivel somente para representante."); window.close();</script>
<%
  Response.End
End If

'*******************************
'***    Parâmetros           ***
'*******************************

Dim dtsp, mes, mesExib, cod_cliente, uf_cliente
dtsp        = Trim(Request.QueryString("dtsp"))
cod_cliente = Trim(Request.QueryString("cod_cli"))
uf_cliente  = Trim(Request.QueryString("uf"))

If dtsp = "" Or Len(dtsp) < 6 Then
%>
<script>alert("Parametros invalidos."); history.back();</script>
<%
  Response.End
End If

mesExib = Mid(dtsp, 5, 2) & "/" & Mid(dtsp, 1, 4)

'*******************************
'***    Query               ***
'*******************************

Dim VarSQL, RS
Dim totquant, tottotal, totpeso, totipi, toticms
totquant = 0 : tottotal = 0 : totpeso = 0 : totipi = 0 : toticms = 0

Dim jsonLinhas
jsonLinhas = "["

VarSQL = "Execute SP_CONFATMESREPRECLIDET [" & dtsp & "],[" & CODUSER & "],[" & cod_cliente & "],[" & uf_cliente & "]"
Set RS = Conecta("PIR-DADOSADV", VarSQL)

Dim temDados, nomeCliente
temDados   = False
nomeCliente = ""

If Not RS Is Nothing Then
  If Not RS.EOF Then
    temDados = True
    Dim primeiro
    primeiro = True
    Do Until RS.EOF
      Dim det, vIPI, vICMS, vRsKg, dataNF

      If RS("TIPO_PRO") = "900" Then
        det = " Kg"
      Else
        det = " Un"
      End If

      ' Formatar data AAAAMMDD → DD/MM/AAAA
      Dim dRaw
      dRaw = CStr(RS("DATANF"))
      If Len(dRaw) >= 8 Then
        dataNF = Mid(dRaw,7,2) & "/" & Mid(dRaw,5,2) & "/" & Mid(dRaw,1,4)
      Else
        dataNF = dRaw
      End If

      If RS("IPI") > 0 And RS("TOTAL") > 0 Then
        vIPI = FormatNumber((RS("ipi") / RS("total") * 100), 1)
      Else
        vIPI = "0,0"
      End If

      If RS("ICMS") > 0 And RS("TOTAL") > 0 Then
        vICMS = FormatNumber((RS("icms") / RS("total") * 100), 1)
      Else
        vICMS = "0,0"
      End If

      If RS("PESO") > 0 Then
        vRsKg = FormatNumber(RS("total") / RS("peso"), 5)
      Else
        vRsKg = "Peso Zerado"
      End If

      If nomeCliente = "" Then nomeCliente = CStr(RS("nome_cli"))

      totquant = totquant + RS("QUANT")
      tottotal = tottotal + RS("total")
      totpeso  = totpeso  + RS("peso")
      totipi   = totipi   + RS("ipi")
      toticms  = toticms  + RS("icms")

      If Not primeiro Then jsonLinhas = jsonLinhas & ","
      primeiro = False

      jsonLinhas = jsonLinhas & "{" & _
        """pedido"":"   & """" & EscapeJSON(CStr(RS("PEDIDO")))    & """," & _
        """datanf"":"   & """" & EscapeJSON(dataNF)                & """," & _
        """produto"":"  & """" & EscapeJSON(CStr(RS("produto")))   & """," & _
        """ref"":"      & """" & EscapeJSON(CStr(RS("ref")))       & """," & _
        """cod_cli"":"  & """" & EscapeJSON(CStr(RS("cod_cli")))   & """," & _
        """nome_cli"":" & """" & EscapeJSON(CStr(RS("nome_cli")))  & """," & _
        """uf"":"       & """" & EscapeJSON(CStr(RS("uf_entrega"))) & """," & _
        """quant"":"    & """" & EscapeJSON(CStr(FormatNumber(RS("quant"), 0)) & det) & """," & _
        """total"":"    & """" & EscapeJSON(CStr(FormatNumber(RS("TOTAL"), 2))) & """," & _
        """peso"":"     & """" & EscapeJSON(CStr(FormatNumber(RS("PESO"),  2))) & """," & _
        """ipi"":"      & """" & EscapeJSON(vIPI)                  & """," & _
        """icms"":"     & """" & EscapeJSON(vICMS)                 & """," & _
        """rskgt"":"    & """" & EscapeJSON(vRsKg)                 & """" & _
      "}"

      RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
  End If
End If

jsonLinhas = jsonLinhas & "]"

Dim jsonTotais
If totpeso > 0 Then
  jsonTotais = "{""totquant"":""" & FormatNumber(totquant, 0) & """," & _
                """tottotal"":""" & FormatNumber(tottotal, 2) & """," & _
                """totpeso"":"""  & FormatNumber(totpeso,  0) & """," & _
                """rskgtot"":""" & FormatNumber(tottotal / totpeso, 5) & """}"
Else
  jsonTotais = "{""totquant"":""" & FormatNumber(totquant, 0) & """," & _
                """tottotal"":""" & FormatNumber(tottotal, 2) & """," & _
                """totpeso"":"""  & FormatNumber(totpeso,  0) & """," & _
                """rskgtot"":""Peso Zerado""}"
End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Detalhe Faturamento — <%=mesExib%></title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

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
      --color-success:       hsl(145, 63%, 32%);
      --color-danger:        hsl(3, 90%, 58%);
      --color-warning:       hsl(35, 100%, 48%);
      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
      --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;
      --radius-sm:  6px;
      --radius-md:  10px;
      --shadow-sm:  0 1px 2px hsla(240, 6%, 10%, .04);
      --shadow-md:  0 4px 16px hsla(240, 6%, 10%, .08);
      --transition: 160ms ease;
      --z-toast:    1100;
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
    .page-subtitle {
      font-size: .82rem;
      color: var(--color-text-muted);
    }

    .table-card {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-sm);
    }

    .form-control:focus, .form-select:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
    }

    /* TOOLBAR */
    .table-toolbar {
      background-color: var(--color-surface);
      border-radius: var(--radius-md) var(--radius-md) 0 0;
    }
    .btn-export-excel {
      background-color: var(--color-success); border-color: var(--color-success);
      color: hsl(0,0%,100%); font-weight: 500; font-size: .8rem;
      transition: background-color var(--transition), border-color var(--transition);
    }
    .btn-export-excel:hover { background-color: hsl(145,63%,27%); border-color: hsl(145,63%,27%); }
    .btn-export-excel .bi   { color: hsl(0,0%,100%); }

    /* SEARCH BOX */
    .search-box { position: relative; max-width: 200px; }
    .search-box .form-control { padding-right: 76px; font-size: .82rem; }
    .search-box .search-hotkey {
      position: absolute; right: 8px; top: 50%; transform: translateY(-50%);
      pointer-events: none; user-select: none; display: flex; gap: 2px;
    }
    .search-box .search-hotkey kbd {
      font-size: .65rem; padding: .1rem .35rem;
      background-color: var(--color-surface-muted);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-sm);
      color: var(--color-text-muted);
      font-family: var(--font-mono);
    }

    /* TABELA */
    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
    .table-wrapper table { min-width: max-content; }

    .table thead th {
      position: sticky; top: 0; z-index: 2;
      background-color: var(--color-surface-muted);
      color: var(--color-text-muted);
      font-size: .76rem; font-weight: 600;
      text-transform: uppercase; letter-spacing: .03em;
      text-align: center; white-space: nowrap;
      border-color: var(--color-border);
      padding: .6rem .75rem;
    }
    .table tbody td {
      font-size: .85rem; text-align: center; white-space: nowrap;
      color: var(--color-text); vertical-align: middle;
      padding: .55rem .75rem; border-color: var(--color-border);
    }
    .table tbody tr:not([class*="row-"]):hover td {
      background-color: var(--color-surface-muted);
    }
    .td-desc { text-align: left !important; max-width: 220px; white-space: normal !important; }
    .td-link {
      color: var(--color-accent); text-decoration: none; font-weight: 500;
      transition: color var(--transition);
    }
    .td-link:hover { color: var(--color-accent-hover); text-decoration: underline; }

    .tr-total td {
      background-color: var(--color-surface-muted) !important;
      font-weight: 600; font-size: .82rem;
      color: var(--color-text) !important;
      border-top: 2px solid var(--color-border-strong);
    }

    /* RODAPÉ */
    .table-footer {
      display: flex; align-items: center; justify-content: space-between;
      padding: .7rem 1.25rem; border-top: 1px solid var(--color-border);
      flex-wrap: wrap; gap: .5rem;
    }
    .table-footer-left  { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
    .table-footer-center { flex: 1; display: flex; justify-content: center; }
    .table-footer-right { min-width: 140px; text-align: right; padding-right: .25rem; }
    .info-registros { font-size: .8rem; color: var(--color-text-muted); }

    /* PAGINAÇÃO */
    #paginacao .page-link {
      color: var(--color-text-muted); border-color: var(--color-border); font-size: .82rem;
      transition: background-color var(--transition), border-color var(--transition), color var(--transition);
    }
    #paginacao .page-item.active .page-link {
      background-color: var(--color-accent); border-color: var(--color-accent); color: var(--color-surface);
    }
    #paginacao .page-link:hover {
      background-color: var(--color-surface-muted); border-color: var(--color-border-strong); color: var(--color-text);
    }
    #paginacao .page-item.disabled .page-link { color: var(--color-text-faint); }

    /* PLACEHOLDER */
    .table-placeholder {
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; padding: 4rem 1rem;
      color: var(--color-text-muted); gap: .85rem;
    }
    .table-placeholder .bi { font-size: 2.4rem; color: var(--color-text-faint); }
    .table-placeholder p  { margin: 0; font-size: .88rem; text-align: center; }

    /* TOAST */
    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: var(--z-toast); display: flex; flex-direction: column; gap: .5rem;
    }
    .toast-custom {
      min-width: 280px; max-width: 360px; border-radius: var(--radius-md);
      overflow: hidden; box-shadow: var(--shadow-md);
      background: var(--color-surface); border: 1px solid var(--color-border);
      animation: toastSlideIn .22s ease;
    }
    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(24px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { animation-duration: .01ms !important; transition-duration: .01ms !important; }
    }
    @media (max-width: 768px) {
      .table-footer { flex-direction: column; align-items: stretch; }
      .table-footer-right { text-align: center; }
    }
  </style>
</head>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center mb-4 gap-3">
      <a href="fat_rep_cli_acumulado.asp?select_mes=<%=Mid(dtsp,5,2)%>&select_ano=<%=Mid(dtsp,1,4)%>"
         class="btn btn-light btn-sm" title="Voltar">
        <i class="bi bi-arrow-left me-1"></i>Voltar
      </a>
      <div>
        <span class="page-title">Detalhe de Faturamento</span>
        <div class="page-subtitle">
          <%=nomeCliente%> &mdash; <%=mesExib%>
        </div>
      </div>
    </div>

    <!-- Card da Tabela -->
    <div class="table-card" id="tableCard">
      <%If Not temDados Then%>
      <div class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Nenhum pedido encontrado para este cliente em <strong><%=mesExib%></strong>.</p>
      </div>
      <%Else%>

      <!-- Toolbar -->
      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          <button id="btnExportCsv" type="button" class="btn btn-export-excel btn-sm">
            <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
          </button>
        </div>
        <div class="search-box">
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" aria-label="Buscar na tabela" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <!-- Tabela -->
      <div class="table-wrapper">
        <table class="table table-bordered table-hover table-sm mb-0" id="tabelaDet">
          <thead>
            <tr>
              <th>Pedido</th>
              <th>Dt Fat</th>
              <th>Produto</th>
              <th>Descrição</th>
              <th>Cód Cliente</th>
              <th>Cliente</th>
              <th>UF</th>
              <th>Quant</th>
              <th>Total</th>
              <th>Peso</th>
              <th>IPI %</th>
              <th>ICMS %</th>
              <th>R$/Kg</th>
            </tr>
          </thead>
          <tbody id="tbodyDet">
            <!-- preenchido pelo JS -->
          </tbody>
        </table>
      </div>

      <!-- Rodapé -->
      <div class="table-footer">
        <div class="table-footer-left">
          <label for="selectLimite" class="form-label mb-0" style="font-size:.8rem;white-space:nowrap;color:var(--color-text-muted);">
            Linhas por página:
          </label>
          <select id="selectLimite" class="form-select form-select-sm" style="width:auto;">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50" selected>50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="table-footer-center">
          <nav aria-label="Navegação de páginas">
            <ul id="paginacao" class="pagination pagination-sm mb-0"></ul>
          </nav>
        </div>
        <div class="table-footer-right">
          <span id="infoRegistros" class="info-registros" aria-live="polite"></span>
        </div>
      </div>
      <%End If%>
    </div>

  </div>

  <div class="toast-container-custom" aria-live="polite" id="toastContainer"></div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
    'use strict';

    const FatRepCliDet = (function () {

      const DADOS  = <%=jsonLinhas%>;
      const TOTAIS = <%=jsonTotais%>;
      const DTSP   = "<%=dtsp%>";

      let _dadosFiltrados = [];
      let _paginaAtual    = 1;
      let _itensPorPagina = 50;

      const DEBOUNCE_MS = 280;
      const TOAST_DELAY = 4500;

      const TOAST_ICON = {
        sucesso: { icon: 'bi-check-circle-fill',         cor: 'var(--color-success)' },
        erro:    { icon: 'bi-x-circle-fill',             cor: 'var(--color-danger)'  },
        aviso:   { icon: 'bi-exclamation-triangle-fill', cor: 'var(--color-warning)' },
        info:    { icon: 'bi-info-circle-fill',          cor: 'var(--color-text-muted)' }
      };

      function exibirToast(tipo, mensagem) {
        const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
        const id  = 'toast_' + Date.now();
        $('#toastContainer').append(`
          <div id="${id}" class="toast-custom">
            <div class="d-flex align-items-center px-3 py-2 border-bottom gap-2">
              <i class="bi ${cfg.icon}" style="color:${cfg.cor};font-size:1rem;"></i>
              <span style="font-size:.85rem;font-weight:600;color:${cfg.cor};flex:1;">${tipo.charAt(0).toUpperCase()+tipo.slice(1)}</span>
              <button type="button" style="background:none;border:none;font-size:1.1rem;cursor:pointer;color:var(--color-text-muted);"
                onclick="this.closest('.toast-custom').remove()">&#215;</button>
            </div>
            <div class="px-3 py-2" style="font-size:.84rem;color:var(--color-text);">${mensagem}</div>
            <div style="height:3px;background:var(--color-border);">
              <div class="toast-progress-bar" style="height:100%;background:${cfg.cor};"></div>
            </div>
          </div>`);
        setTimeout(() => { const el = document.getElementById(id); if (el) el.remove(); }, TOAST_DELAY + 250);
      }

      function _debounce(fn, ms) {
        let t;
        return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
      }

      function _renderTabela() {
        const inicio = (_paginaAtual - 1) * _itensPorPagina;
        const pagina = _dadosFiltrados.slice(inicio, inicio + _itensPorPagina);
        const $tbody = $('#tbodyDet').empty();

        pagina.forEach(function (r) {
          const urlPedido = '/inpanet/consultas/Detalha_NF.asp?Pedido=' + encodeURIComponent(r.pedido);
          const urlFT     = '/inpanet/consultas/folha_ft.asp?FT='        + encodeURIComponent(r.produto);

          $tbody.append(`<tr>
            <td><a href="${urlPedido}" class="td-link">${r.pedido}</a></td>
            <td>${r.datanf}</td>
            <td><a href="${urlFT}" class="td-link">${r.produto}</a></td>
            <td class="td-desc">${r.ref}</td>
            <td>${r.cod_cli}</td>
            <td class="td-desc">${r.nome_cli}</td>
            <td>${r.uf}</td>
            <td>${r.quant}</td>
            <td>R$ ${r.total}</td>
            <td>${r.peso} Kg</td>
            <td>${r.ipi} %</td>
            <td>${r.icms} %</td>
            <td>${r.rskgt}</td>
          </tr>`);
        });

        $tbody.append(`<tr class="tr-total">
          <td colspan="7">Totais</td>
          <td>${TOTAIS.totquant}</td>
          <td>R$ ${TOTAIS.tottotal}</td>
          <td>${TOTAIS.totpeso} Kg</td>
          <td></td>
          <td></td>
          <td>${TOTAIS.rskgtot}</td>
        </tr>`);

        const total    = _dadosFiltrados.length;
        const totalPags = Math.ceil(total / _itensPorPagina);
        $('#infoRegistros').text('Total: ' + total + ' pedido(s)');
        _renderPaginacao(_paginaAtual, totalPags);
      }

      function _renderPaginacao(atual, total) {
        const $ul = $('#paginacao').empty();
        if (total <= 1) return;
        const mk = (label, page, disabled, active) => {
          const $li = $('<li>').addClass('page-item');
          if (disabled) $li.addClass('disabled');
          if (active)   $li.addClass('active');
          const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
          if (!disabled && !active) $a.on('click', e => { e.preventDefault(); _paginaAtual = page; _renderTabela(); });
          return $li.append($a);
        };
        $ul.append(mk('&laquo;', 1, atual === 1, false));
        $ul.append(mk('&lsaquo;', atual - 1, atual === 1, false));
        const ini = Math.max(1, atual - 2), fim = Math.min(total, atual + 2);
        for (let p = ini; p <= fim; p++) $ul.append(mk(p, p, false, p === atual));
        $ul.append(mk('&rsaquo;', atual + 1, atual === total, false));
        $ul.append(mk('&raquo;', total, atual === total, false));
      }

      function _aplicarBusca(termo) {
        const t = termo.toLowerCase().trim();
        _dadosFiltrados = t
          ? DADOS.filter(r =>
              r.pedido.toLowerCase().includes(t)   ||
              r.produto.toLowerCase().includes(t)  ||
              r.ref.toLowerCase().includes(t)      ||
              r.nome_cli.toLowerCase().includes(t) ||
              r.uf.toLowerCase().includes(t)
            )
          : DADOS.slice();
        _paginaAtual = 1;
        _renderTabela();
      }

      function _exportarCSV() {
        const esc = v => {
          const s = v == null ? '' : String(v);
          return (s.includes(';') || s.includes('"') || s.includes('\n'))
            ? '"' + s.replace(/"/g, '""') + '"' : s;
        };
        const cab  = ['Pedido','Dt Fat','Produto','Descrição','Cód Cliente','Cliente','UF','Quant','Total','Peso (Kg)','IPI %','ICMS %','R$/Kg'];
        const linhas = DADOS.map(r => [r.pedido, r.datanf, r.produto, r.ref, r.cod_cli, r.nome_cli, r.uf, r.quant, r.total, r.peso, r.ipi, r.icms, r.rskgt]);
        linhas.push(['', '', '', '', '', '', 'TOTAIS', TOTAIS.totquant, TOTAIS.tottotal, TOTAIS.totpeso, '', '', TOTAIS.rskgtot]);
        const rows = [cab.map(esc).join(';'), ...linhas.map(row => row.map(esc).join(';'))];
        const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href = url;
        a.download = 'detalhe_' + DTSP + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }

      function init() {
        if (!DADOS || !DADOS.length) return;

        _dadosFiltrados = DADOS.slice();
        _renderTabela();

        $('#selectLimite').on('change', function () {
          _itensPorPagina = parseInt($(this).val(), 10);
          _paginaAtual = 1;
          _renderTabela();
        });

        $('#searchBox').on('input', _debounce(function () {
          _aplicarBusca($(this).val());
        }, DEBOUNCE_MS));

        $(document).on('keydown', function (e) {
          if (e.ctrlKey && e.key === 'k') { e.preventDefault(); $('#searchBox').trigger('focus'); }
        });

        $('#btnExportCsv').on('click', _exportarCSV);
      }

      return Object.freeze({ init });
    })();

    $(function () { FatRepCliDet.init(); });
  </script>
</body>
</html>
