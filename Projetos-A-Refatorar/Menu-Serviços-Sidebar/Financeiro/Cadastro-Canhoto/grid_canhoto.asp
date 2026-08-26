<%
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046
%>

<!-- #include file="../../cnn2/cnn2.asp" -->
<!-- #include file="../../funcoes/funcoes.asp" -->
<!-- #include file="../../includes/variaveis.asp" -->

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — Canhotos de Nota Fiscal</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

  <% if Session("NIVEL") <> "0" then %>
  <script src="../../js/sw_noviewsrc.js"></script>
  <% end if %>

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
      font-family: var(--bs-font-sans-serif);
    }

    .filter-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .filter-card-header {
      background-color: #fff;
      color: #212529;
      padding: .6rem 1rem;
      border-radius: 6px 6px 0 0;
      border-bottom: 1px solid #dee2e6;
      font-weight: 600;
      font-size: .9rem;
    }
    .filter-card-body {
      padding: 1rem 1.25rem 1.25rem;
    }
    .filter-card-body label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }

    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 68vh;
    }
    .table-wrapper table { min-width: max-content; }

    .table thead th {
      position: sticky;
      top: 0;
      background-color: #fff;
      color: #212529;
      font-size: .82rem;
      font-weight: 700;
      text-align: center;
      white-space: nowrap;
      z-index: 2;
      border-color: #dee2e6;
      padding: .55rem .75rem;
    }
    .table tbody td {
      font-size: .82rem;
      text-align: center;
      white-space: nowrap;
      color: #212529;
      vertical-align: middle;
      padding: .45rem .75rem;
    }

    .table-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: .6rem 1rem;
      border-top: 1px solid #dee2e6;
      flex-wrap: wrap;
      gap: .5rem;
    }
    .table-footer-left   { min-width: 160px; display: flex; align-items: center; gap: .5rem; }
    .table-footer-center { flex: 1; display: flex; justify-content: center; }
    .table-footer-right  { min-width: 160px; text-align: right; }
    .info-registros { font-size: .8rem; color: #6c757d; white-space: nowrap; }
    .select-limite  { width: 75px; font-size: .82rem; }

    #paginacao .page-link { color: #6c757d; border-color: #dee2e6; }
    #paginacao .page-item.active .page-link {
      background-color: #6c757d;
      border-color: #6c757d;
      color: #fff;
    }
    #paginacao .page-link:hover   { background-color: #e9ecef; color: #495057; }
    #paginacao .page-item.disabled .page-link { color: #adb5bd; }

    .search-box { position: relative; max-width: 180px; }
    .search-box .form-control { padding-right: 72px; font-size: .82rem; }
    .search-box .search-hotkey {
      position: absolute;
      right: 8px;
      top: 50%;
      transform: translateY(-50%);
      pointer-events: none;
      user-select: none;
    }
    .search-box .search-hotkey kbd {
      font-size: .65rem;
      padding: .1rem .3rem;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      color: #6c757d;
    }

    .table-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3.5rem 1rem;
      color: #6c757d;
      gap: .75rem;
    }
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

    .overlay-loading {
      position: fixed;
      inset: 0;
      background: rgba(255,255,255,.72);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 1055;
      gap: 1rem;
    }
    .overlay-loading .spinner-border { width: 3rem; height: 3rem; }
    .overlay-loading span { font-size: .95rem; font-weight: 600; }

    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: 1100;
      display: flex; flex-direction: column; gap: .5rem;
    }
    .toast-custom {
      min-width: 280px; max-width: 360px;
      border-radius: 10px; overflow: hidden;
      box-shadow: 0 4px 18px rgba(0,0,0,.13);
      background: #fff; border: 1px solid #dee2e6;
      animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
    }
    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(40px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-custom .toast-progress-bar {
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; } to { width: 0%; }
    }
  </style>
</head>

<%
Dim strMsg
strMsg = ""

if trim(Request.QueryString("aviso")) = "ok" then
  select case trim(ucase(Request.QueryString("acao")))
  case "INS" strMsg = "Cadastro realizado com sucesso!"
  case "ALT" strMsg = "Cadastro atualizado com sucesso!"
  case "DEL" strMsg = "Registro apagado com sucesso!"
  end select
end if

Dim pesqData, pesqNF, pesqPlaca
pesqData  = ""
pesqNF    = ""
pesqPlaca = ""

if trim(request("cboPesquisa")) = "DATA" and trim(request("pesquisaData")) <> "" then
  ' O input type="date" retorna yyyy-MM-dd; a coluna DATA é yyyyMMdd (sem hifens)
  pesqData = "AND DATA = '" & replace(trim(request("pesquisaData")), "-", "") & "' "
end if
if trim(request("cboPesquisa")) = "NF" and trim(request("pesquisaNF")) <> "" then
  pesqNF = "AND NF = '" & trim(request("pesquisaNF")) & "' "
end if
if trim(request("cboPesquisa")) = "PLACA" and trim(request("pesquisaPlaca")) <> "" then
  pesqPlaca = "AND OBS LIKE '%" & trim(request("pesquisaPlaca")) & "%' "
end if

Dim quantidadeItems
quantidadeItems = "50"
if trim(request("cmbLinhas")) <> "" then
  quantidadeItems = trim(request("cmbLinhas"))
end if

Dim cboPlantaVal
cboPlantaVal = trim(request("cboPlanta"))
if cboPlantaVal = "" then cboPlantaVal = "00"

Dim filialDigit
filialDigit = mid(cboPlantaVal, 2, 1)
if filialDigit = "" then filialDigit = "0"

SQL = "SELECT TOP " & quantidadeItems & " " &_
        "NF, SERIE, DATA, USERDIGIT, USERALTERA, OBS, " &_
        "isnull(Carga,0) Carga, isnull(Veiculo,0) Veiculo, Filial " &_
      "FROM CANHOTOS " &_
      "WHERE D_E_L_E_T_ <> '*' " &_
        "AND Filial = '" & filialDigit & "' " &_
        pesqData & pesqNF & pesqPlaca &_
      "ORDER BY R_E_C_N_O_ DESC "

set RS = Conecta("PIR-CORPORATIVO", SQL)
%>

<body>

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border text-secondary" role="status"></div>
    <span>Carregando...</span>
  </div>

  <div class="toast-container-custom" id="toastContainer"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="mb-3">
      <span class="page-title" style="font-size:1.15rem;font-weight:700;color:#212529;">
        Canhotos de Nota Fiscal
      </span>
    </div>

    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>

      <div class="filter-card-body">
        <form id="filtro" name="filtro" method="post" action="grid_canhoto.asp" novalidate>
          <input type="hidden" name="planta"    id="planta"    value="<%=cboPlantaVal%>" />
          <input type="hidden" name="cmbLinhas" id="hCmbLinhas" value="<%=quantidadeItems%>" />

          <div class="d-flex flex-wrap align-items-end gap-2 justify-content-between">
            <div class="d-flex flex-wrap align-items-end gap-2">
              <div>
                <label for="cboPlanta">Planta</label>
                <select name="cboPlanta" id="cboPlanta" class="form-select form-select-sm"
                        onchange="document.getElementById('planta').value=this.value; document.getElementById('filtro').submit();">
                  <option value="00" <%if cboPlantaVal="00" then%>selected<%end if%>>Pirapetinga</option>
                  <option value="04" <%if cboPlantaVal="04" then%>selected<%end if%>>Uberaba</option>
                  <option value="05" <%if cboPlantaVal="05" then%>selected<%end if%>>Bento Gonçalves</option>
                  <option value="06" <%if cboPlantaVal="06" then%>selected<%end if%>>Fortaleza</option>
                  <option value="07" <%if cboPlantaVal="07" then%>selected<%end if%>>Saquarema</option>
                </select>
              </div>

              <div>
                <label for="cboPesquisa">Pesquisar por</label>
                <select name="cboPesquisa" id="cboPesquisa" class="form-select form-select-sm"
                        onchange="document.getElementById('filtro').submit();">
                  <option value=""      <%if trim(request("cboPesquisa"))=""      then%>selected<%end if%>>Selecione...</option>
                  <option value="PLACA" <%if trim(request("cboPesquisa"))="PLACA" then%>selected<%end if%>>Placa</option>
                  <option value="DATA"  <%if trim(request("cboPesquisa"))="DATA"  then%>selected<%end if%>>Data</option>
                  <option value="NF"    <%if trim(request("cboPesquisa"))="NF"    then%>selected<%end if%>>Nota Fiscal</option>
                </select>
              </div>

              <%
              if request("cboPesquisa") = "DATA" then
              %>
              <div>
                <label for="pesquisaData">Data</label>
                <input id="pesquisaData" name="pesquisaData" type="date"
                       class="form-control form-control-sm" style="width:155px"
                       value="<%=request("pesquisaData")%>" />
              </div>
              <%
              elseif request("cboPesquisa") = "NF" then
              %>
              <div>
                <label for="pesquisaNF">Nota Fiscal</label>
                <input id="pesquisaNF" name="pesquisaNF" type="text"
                       class="form-control form-control-sm" style="width:140px"
                       maxlength="10" value="<%=request("pesquisaNF")%>" />
              </div>
              <%
              elseif request("cboPesquisa") = "PLACA" then
              %>
              <div>
                <label for="pesquisaPlaca">Placa</label>
                <input id="pesquisaPlaca" name="pesquisaPlaca" type="text"
                       class="form-control form-control-sm" style="width:140px"
                       maxlength="10" value="<%=request("pesquisaPlaca")%>" />
              </div>
              <%
              end if
              %>

            </div>

            <div class="d-flex align-items-end gap-2 ms-auto">
              <button type="submit" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" class="btn btn-light btn-sm"
                      onclick="document.getElementById('cboPesquisa').value='';
                               document.getElementById('filtro').submit();">
                <i class="bi bi-eraser"></i>
              </button>
              <button type="button" class="btn btn-danger btn-sm" title="Exportar PDF"
                      onclick="gerarPDF()">
                <i class="bi bi-file-earmark-pdf"></i>
              </button>
              <button type="button" class="btn btn-warning btn-sm" id="btnImprimir" title="Imprimir canhotos selecionados"
                  onclick="imprimirSelecionados()">
                <i class="bi bi-printer me-2 text-black"></i>Etiqueta
              </button>
            </div>

          </div>
        </form>
      </div>
    </div>

    <div class="table-card">

      <div class="d-flex align-items-center justify-content-between px-3 py-2 border-bottom gap-2 flex-wrap">
        <div class="d-flex gap-2">
          <button type="button" class="btn btn-success btn-sm"
                  onclick="location.href='form_canhoto.asp?acao=INS&filial=<%=filialDigit%>'">
            <i class="bi bi-plus-lg me-1"></i>Novo
          </button>
          
        </div>
        <div class="search-box">
          <input type="text" id="buscaMemoria" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <div class="table-wrapper">
        <%
        if RS.EOF then
        %>
        <div class="table-placeholder">
          <i class="bi bi-clipboard-data"></i>
          <p>Nenhum registro encontrado para os filtros selecionados.</p>
        </div>
        <%
        else
        %>
        <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tblCanhotos">
          <thead>
            <tr>
              <th style="width:32px;">
                <input type="checkbox" id="selectAll" class="form-check-input mt-0"
                       title="Selecionar todos" onchange="toggleTodos(this)" />
              </th>
              <th>Ações</th>
              <th>NF / Série</th>
              <th>Data</th>
              <th>Produto</th>
              <th>Serviço</th>
              <th>Incluído por</th>
              <th>Alterado por</th>
              <th>Observação</th>
            </tr>
          </thead>
          <tbody id="tbodyCanhotos">
            <%
            do while not RS.EOF
              Dim nomeChk
              nomeChk = RS("NF") & RS("SERIE") & mid(RS("Filial"),2,1)
            %>
            <tr>
              <td>
                <input class="form-check-input chk-canhoto" type="checkbox"
                       value="<%=nomeChk%>" title="Selecionar" />
              </td>
              <td>
                <div class="d-flex gap-1 justify-content-center">
                  <button type="button" class="btn btn-primary btn-sm" title="Visualizar"
                          onclick="location.href='form_canhoto.asp?acao=CON&nf=<%=RS("NF")%>&serie=<%=RS("SERIE")%>&filial=<%=RS("Filial")%>'">
                    <i class="bi bi-eye"></i>
                  </button>
                  <button type="button" class="btn btn-outline-warning btn-sm" title="Editar"
                          onclick="location.href='form_canhoto.asp?acao=ALT&nf=<%=RS("NF")%>&serie=<%=RS("SERIE")%>&filial=<%=RS("Filial")%>'">
                    <i class="bi bi-pencil"></i>
                  </button>
                  <button type="button" class="btn btn-outline-danger btn-sm" title="Apagar"
                          onclick="location.href='form_canhoto.asp?acao=DEL&nf=<%=RS("NF")%>&serie=<%=RS("SERIE")%>&filial=<%=RS("Filial")%>'">
                    <i class="bi bi-trash"></i>
                  </button>
                </div>
              </td>
              <td><%=RS("NF")%> / <%=RS("SERIE")%></td>
              <td><%=AcertaData(RS("DATA"))%></td>
              <td><%=mostraCarga(RS("Carga"))%></td>
              <td><%=mostraVeiculo(RS("Veiculo"))%></td>
              <td><%=RS("USERDIGIT")%></td>
              <td><%=RS("USERALTERA")%></td>
              <td style="text-align:left;"><%=RS("OBS")%></td>
            </tr>
            <%
              RS.MoveNext()
            loop
            %>
          </tbody>
        </table>
        <%
        end if
        RS.Close()
        %>
      </div>

      <div class="table-footer">
        <div class="table-footer-left">
          <label for="cmbLinhas" class="mb-0" style="font-size:.8rem;color:#6c757d;white-space:nowrap;">Linhas:</label>
          <select id="cmbLinhas" class="form-select form-select-sm select-limite"
                  onchange="alterarLinhas(this.value)">
            <option value="10"  <%if quantidadeItems="10"  then%>selected<%end if%>>10</option>
            <option value="25"  <%if quantidadeItems="25"  then%>selected<%end if%>>25</option>
            <option value="50"  <%if quantidadeItems="50"  then%>selected<%end if%>>50</option>
            <option value="100" <%if quantidadeItems="100" then%>selected<%end if%>>100</option>
            <option value="250" <%if quantidadeItems="250" then%>selected<%end if%>>250</option>
            <option value="500" <%if quantidadeItems="500" then%>selected<%end if%>>500</option>
          </select>
        </div>

        <div class="table-footer-center">
          <nav><ul class="pagination pagination-sm mb-0" id="paginacao"></ul></nav>
        </div>

        <div class="table-footer-right">
          <span class="info-registros" id="infoRegistros"></span>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <script>
  (() => {
    "use strict";

    const TOAST_DELAY = 4500;
    const LOGO_PATH   = "/corporativo/CriarPermissao/Logo/logo-smurfit1.png";
    const TOAST_ICON  = Object.freeze({
      sucesso: { icon: "bi-check-circle-fill",           color: "#198754" },
      erro:    { icon: "bi-x-circle-fill",               color: "#dc3545" },
      aviso:   { icon: "bi-exclamation-triangle-fill",   color: "#856404" },
      info:    { icon: "bi-info-circle-fill",            color: "#6c757d" }
    });

    const $el = (id) => document.getElementById(id);

    function exibirToast(tipo, titulo, mensagem) {
      const cfg  = TOAST_ICON[tipo] || TOAST_ICON.info;
      const cont = $el("toastContainer");
      const el   = document.createElement("div");
      el.className = "toast-custom";
      el.innerHTML = `
        <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
          <i class="bi ${cfg.icon}" style="color:${cfg.color};font-size:1.1rem;"></i>
          <strong class="flex-grow-1" style="font-size:.85rem;color:${cfg.color};">${titulo}</strong>
          <button class="btn-close btn-close-sm" onclick="this.closest('.toast-custom').remove()"></button>
        </div>
        <div class="px-3 pb-2" style="font-size:.82rem;color:#495057;">${mensagem}</div>
        <div style="height:3px;background:#e9ecef;">
          <div class="toast-progress-bar" style="height:100%;background:${cfg.color};width:100%;"></div>
        </div>`;
      cont.appendChild(el);
      setTimeout(() => el.remove(), TOAST_DELAY + 300);
    }

    const PAGINA_SIZE = 10;
    let paginaAtual      = 1;
    let linhasFiltradas  = [];

    function obterLinhas() {
      return Array.from(document.querySelectorAll("#tbodyCanhotos tr"));
    }

    function renderPagina() {
      const total    = linhasFiltradas.length;
      const totalPag = Math.max(1, Math.ceil(total / PAGINA_SIZE));
      if (paginaAtual > totalPag) paginaAtual = totalPag;

      const inicio = (paginaAtual - 1) * PAGINA_SIZE;
      const fim    = inicio + PAGINA_SIZE;

      obterLinhas().forEach(tr => tr.style.display = "none");
      linhasFiltradas.slice(inicio, fim).forEach(tr => tr.style.display = "");

      const infoEl = $el("infoRegistros");
      if (infoEl) {
        infoEl.textContent = total === 0
          ? "Nenhum registro"
          : `Exibindo ${inicio + 1}\u2013${Math.min(fim, total)} de ${total} registros`;
      }

      const ul = $el("paginacao");
      if (!ul) return;
      ul.innerHTML = "";

      const addItem = (label, page, disabled, active) => {
        const li = document.createElement("li");
        li.className = "page-item" + (disabled ? " disabled" : "") + (active ? " active" : "");
        const a = document.createElement("a");
        a.className  = "page-link";
        a.href       = "#";
        a.innerHTML  = label;
        if (!disabled && !active) {
          a.addEventListener("click", e => { e.preventDefault(); paginaAtual = page; renderPagina(); });
        }
        li.appendChild(a);
        ul.appendChild(li);
      };

      addItem("&laquo;", paginaAtual - 1, paginaAtual === 1);
      const janela = 2;
      for (let p = Math.max(1, paginaAtual - janela); p <= Math.min(totalPag, paginaAtual + janela); p++) {
        addItem(p, p, false, p === paginaAtual);
      }
      addItem("&raquo;", paginaAtual + 1, paginaAtual === totalPag);
    }

    function aplicarBusca(termo) {
      const t = termo.toLowerCase().trim();
      const todas = obterLinhas();
      linhasFiltradas = t === "" ? todas : todas.filter(tr => tr.textContent.toLowerCase().includes(t));
      paginaAtual = 1;
      renderPagina();
    }

    window.toggleTodos = function(chk) {
      document.querySelectorAll("#tbodyCanhotos .chk-canhoto").forEach(c => c.checked = chk.checked);
    };

    window.alterarLinhas = function(val) {
      const inp = $el("hCmbLinhas");
      if (inp) inp.value = val;
      $el("filtro").submit();
    };

    window.gerarPDF = function() {
      const { jsPDF } = window.jspdf;
      const doc    = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      const pageW  = doc.internal.pageSize.getWidth();
      const hoje   = new Date().toLocaleDateString("pt-BR");
      const planta = document.querySelector("#cboPlanta option:checked")?.textContent || "";

      const carregarLogo = () => new Promise((resolve) => {
        const img = new Image();
        img.crossOrigin = "anonymous";
        img.onload = () => {
          const escala  = Math.min(300 / img.naturalWidth, 1);
          const canvas  = document.createElement("canvas");
          canvas.width  = img.naturalWidth  * escala;
          canvas.height = img.naturalHeight * escala;
          canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
          resolve({ data: canvas.toDataURL("image/jpeg", 0.75), ratio: canvas.height / canvas.width });
        };
        img.onerror = () => resolve(null);
        img.src = LOGO_PATH;
      });

      carregarLogo().then((logo) => {
        if (logo) {
          const logoW = 38;
          const logoH = logoW * logo.ratio;
          doc.addImage(logo.data, "JPEG", 10, 4, logoW, logoH);
        }

        doc.setFontSize(13);
        doc.setFont("helvetica", "bold");
        doc.text("Canhotos de Nota Fiscal", pageW / 2, 14, { align: "center" });
        doc.setFontSize(8);
        doc.setFont("helvetica", "normal");
        doc.text(planta, pageW / 2, 20, { align: "center" });
        doc.setFontSize(7);
        doc.text(`Emitido em: ${hoje}`, pageW - 10, 8, { align: "right" });
        const cabecalho = ["NF / Série", "Data", "Produto", "Serviço", "Incluído por", "Alterado por", "Observação"];
        const linhas = Array.from(document.querySelectorAll("#tbodyCanhotos tr")).map(tr => {
          const tds = tr.querySelectorAll("td");
          return [
            tds[2]?.textContent.trim() || "",
            tds[3]?.textContent.trim() || "",
            tds[4]?.textContent.trim() || "",
            tds[5]?.textContent.trim() || "",
            tds[6]?.textContent.trim() || "",
            tds[7]?.textContent.trim() || "",
            tds[8]?.textContent.trim() || ""
          ];
        });

        doc.autoTable({
          startY: 26,
          head: [cabecalho],
          body: linhas,
          theme: "grid",
          headStyles:         { fillColor: [74,74,74], textColor: [255,255,255], fontSize: 4.5, cellPadding: 1.5 },
          bodyStyles:         { fontSize: 4.5, cellPadding: 1.5 },
          alternateRowStyles: { fillColor: [245,245,245] },
          tableLineColor:     [180,180,180],
          tableLineWidth:     0.2
        });

        doc.save(`canhotos_${planta.replace(/\s+/g,"_")}_${hoje.replace(/\//g,"-")}.pdf`);
      });
    };

    window.imprimirSelecionados = function() {
      const marcados = Array.from(document.querySelectorAll("#tbodyCanhotos .chk-canhoto:checked"))
        .map(c => "'" + c.value + "'");

      if (marcados.length === 0) {
        exibirToast("aviso", "Nenhum item selecionado", "Marque ao menos um canhoto para imprimir.");
        return;
      }

      const selecionados = marcados.join(",");
      const cboPlanta    = document.getElementById("cboPlanta").value;
      const url = "impr_canhoto.asp?selecionados=" + encodeURIComponent(selecionados)
                + "&cboPlanta=" + encodeURIComponent(cboPlanta);

      window.open(url, "_blank", "width=600,height=700,scrollbars=yes,resizable=yes");
    };

    document.addEventListener("DOMContentLoaded", () => {
      linhasFiltradas = obterLinhas();
      renderPagina();

      const buscaEl = $el("buscaMemoria");
      if (buscaEl) {
        buscaEl.addEventListener("input", () => aplicarBusca(buscaEl.value));
      }

      document.addEventListener("keydown", e => {
        if (e.ctrlKey && e.key === "k") {
          e.preventDefault();
          if (buscaEl) { buscaEl.focus(); buscaEl.select(); }
        }
      });

      <%
      if trim(strMsg) <> "" then
        Dim tipoToast
        select case trim(ucase(Request.QueryString("acao")))
        case "INS" tipoToast = "sucesso"
        case "ALT" tipoToast = "sucesso"
        case "DEL" tipoToast = "info"
        case else  tipoToast = "info"
        end select
        response.write "exibirToast('" & tipoToast & "', 'Operação concluída', '" & strMsg & "');"
      end if
      %>
    });

  })();
  </script>

</body>
</html>

<%
Function AcertaData(InverteData)
  Dim Ano, Mes, Dia
  Ano = mid(InverteData, 1, 4)
  Mes = mid(InverteData, 5, 2)
  Dia = mid(InverteData, 7, 2)
  if Len(Ano) < 4 then Ano = "20" & Ano
  if Len(Mes) = 1 then Mes = "0" & Mes
  if Len(Dia) = 1 then Dia = "0" & Dia
  AcertaData = Dia & "/" & Mes & "/" & Ano
End Function

Function mostraCarga(arg)
  Dim carga
  select case CStr(arg)
  case "1" carga = "Muito insatisfeito"
  case "2" carga = "Insatisfeito"
  case "3" carga = "Indiferente"
  case "4" carga = "Satisfeito"
  case "5" carga = "Muito Satisfeito"
  case else carga = "Não respondido"
  end select
  mostraCarga = carga
End Function

Function mostraVeiculo(arg)
  Dim veiculo
  select case CStr(arg)
  case "1" veiculo = "Muito insatisfeito"
  case "2" veiculo = "Insatisfeito"
  case "3" veiculo = "Indiferente"
  case "4" veiculo = "Satisfeito"
  case "5" veiculo = "Muito Satisfeito"
  case else veiculo = "Não respondido"
  end select
  mostraVeiculo = veiculo
End Function
%>
