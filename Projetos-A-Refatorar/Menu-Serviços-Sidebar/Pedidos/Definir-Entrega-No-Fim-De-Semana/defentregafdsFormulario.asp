<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../cnn2/conexoes.asp"-->
<!--#Include File="../../includes/variaveis.asp"-->
<!-- #include file="../../funcoes/funcoes.asp" -->
<%
'===================================================================================================
' FONTE    : defentregafdsFormulario.asp
' DESCRICAO: Frontend SPA do modulo "Define Entrega no Final de Semana".
'            Substitui grid_defentregafds.asp + form_defentregafds.asp.
'            Criar/Editar/Visualizar acontecem em modal (SPA), Excluir via modal de confirmacao.
'            Consome defentregafdsRepositorio.asp via AJAX (JSON).
'===================================================================================================
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID      = 1046
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Define Entrega no Final de Semana</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
    }

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      position: relative;
      max-width: 1200px;
      margin-left: auto;
      margin-right: auto;
    }

    #tabelaGrid th:nth-child(1),
    #tabelaGrid td:nth-child(1) {
      width: 90px;
    }

    #tabelaGrid th:nth-child(2),
    #tabelaGrid td:nth-child(2) {
      width: 90px;
    }

    #tabelaGrid th:nth-child(4),
    #tabelaGrid td:nth-child(4) {
      width: 70px;
      text-align: center;
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

    .form-label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }

    /* Toolbar do topo da tabela */
    .table-toolbar {
      background-color: #fff;
      border-radius: 6px 6px 0 0;
    }

    /* Search box (Ctrl+K) */
    .search-box { position: relative; max-width: 200px; }
    .search-box .form-control { padding-right: 76px; font-size: .82rem; }
    .search-box .search-hotkey {
      position: absolute;
      right: 8px;
      top: 50%;
      transform: translateY(-50%);
      pointer-events: none;
      user-select: none;
      display: flex;
      gap: 2px;
    }
    .search-box .search-hotkey kbd {
      font-size: .65rem;
      padding: .1rem .3rem;
      background-color: #f8f9fa;
      border: 1px solid #dee2e6;
      border-radius: 3px;
      color: #6c757d;
    }

    /* Rodapé da tabela */
    .table-footer {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: .6rem 1rem;
      border-top: 1px solid #dee2e6;
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
    }
    .info-registros { font-size: .8rem; color: #6c757d; }
    .select-limite  { width: auto; }

    /* Paginação */
    #paginacao .page-link {
      color: #6c757d;
      border-color: #dee2e6;
    }
    #paginacao .page-item.active .page-link {
      background-color: #6c757d;
      border-color: #6c757d;
      color: #fff;
    }
    #paginacao .page-link:hover {
      background-color: #e9ecef;
      color: #495057;
    }
    #paginacao .page-item.disabled .page-link {
      color: #adb5bd;
    }

    /* Toast (padrão SWnet) */
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
    .toast-custom .toast-custom-body {
      display: flex;
      align-items: flex-start;
      padding: .65rem .75rem;
      gap: .5rem;
    }
    .toast-custom .toast-custom-close {
      background: none;
      border: none;
      font-size: 1.1rem;
      line-height: 1;
      color: #adb5bd;
      cursor: pointer;
      padding: 0 0 0 .5rem;
    }
  </style>
</head>
<body>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="row align-items-center mb-3" style="max-width:1200px;margin-left:auto;margin-right:auto;">
      <div class="col">
        <span class="page-title">Lista de Cliente - Entrega no Final de Semana</span>
      </div>
      <div class="col-auto">
        <!-- <a href="../../../manuais/manual_cadastro_de_fornecedores.pdf" target="_blank" title="Manual" class="btn btn-info btn-sm">
          <i class="bi bi-question-circle"></i>
        </a> -->
      </div>
    </div>

    <div class="table-card">

      <!-- Toolbar: search box (Ctrl+K) -->
      <div class="table-toolbar d-flex align-items-center justify-content-end px-3 py-2 border-bottom">
        <div class="search-box">
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <div class="table-wrapper">
        <table id="tabelaGrid" class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead>
            <tr>
              <th>
                <button type="button" class="btn btn-success btn-sm" title="Criar" id="btnNovo">
                  Criar
                </button>
              </th>
              <th>CÓDIGO</th>
              <th>NOME</th>
              <th>ATIVO</th>
            </tr>
          </thead>
          <tbody id="tbodyGrid">
            <tr>
              <td colspan="4" class="py-4">
                <div class="spinner-border text-secondary" role="status">
                  <span class="visually-hidden">Carregando...</span>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Rodapé: page size · paginação · total -->
      <div class="table-footer">

        <div class="table-footer-left">
          <label for="selectLimite" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">
            Linhas por página:
          </label>
          <select id="selectLimite" class="form-select form-select-sm select-limite">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
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

    </div>

  </div>

  <!-- Modal Criar / Editar / Visualizar -->
  <div class="modal fade" id="modalForm" tabindex="-1" aria-labelledby="modalFormTitulo" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">

        <div class="modal-header">
          <h5 class="modal-title" id="modalFormTitulo">Novo</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
        </div>

        <div class="modal-body">

          <div id="modalFormAlerta" class="alert alert-danger d-none" role="alert"></div>

          <form class="needs-validation" id="formCadastro" novalidate>

            <input type="hidden" id="modo" value="criar" />

            <div class="row g-3">
              <div class="col-12">
                <label for="codcli" class="form-label">Cliente</label>
                <div class="input-group">
                  <input id="codcli" type="text" pattern="[0-9]*" inputmode="numeric"
                         class="form-control text-end" style="max-width: 140px;" maxlength="14"
                         required />
                  <input type="hidden" id="lojacli" value="" />
                  <button id="btnPesqCli" type="button" class="btn btn-light"
                          title="Filtrar"
                          onclick="janela('../../buscaCLIENTE.asp?id1=codcli&id2=lojacli&id3=razao&id6=00','Consulta Cliente','width=650,height=400,scrollbars=1');">
                    <i class="bi bi-search"></i>
                  </button>
                  <input id="razao" type="text" class="form-control" maxlength="255" readonly />
                </div>
                <div class="form-text" id="codcliStatus"></div>
              </div>
            </div>

            <div class="row g-3 mt-1">
              <div class="col-12">
                <div class="form-check form-switch">
                  <input id="cbxAtivo" class="form-check-input" type="checkbox" />
                  <label for="cbxAtivo" class="form-check-label">Permitir entrega aos finais de semana.</label>
                </div>
              </div>
            </div>

          </form>

        </div>

        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Fechar</button>
          <button type="button" class="btn btn-primary" id="btnSalvar">
            <span class="spinner-border spinner-border-sm d-none" id="spinnerSalvar"></span>
            Salvar
          </button>
        </div>

      </div>
    </div>
  </div>

  <!-- Modal Excluir (confirmacao) -->
  <div class="modal fade" id="modalExcluir" tabindex="-1" aria-labelledby="modalExcluirTitulo" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="modalExcluirTitulo">Apagar registro</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
        </div>
        <div class="modal-body">
          <p class="mb-1">Deseja realmente apagar o registro do cliente <strong id="excluirCodigo"></strong>?</p>
          <p class="text-muted mb-0" id="excluirNome"></p>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
          <button type="button" class="btn btn-danger" id="btnConfirmarExcluir">
            <span class="spinner-border spinner-border-sm d-none" id="spinnerExcluir"></span>
            Apagar
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- Toast container (padrão SWnet) -->
  <div class="toast-container-custom" id="toastContainer"></div>

  <!-- jQuery -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <script>
    if (typeof janela !== "function") {
      function janela(url, nome, opcoes) {
        window.open(url, nome, opcoes);
      }
    }
  </script>

  <script>
    const DefEntregaFDS = (function () {
      "use strict";

      const REPOSITORIO = "defentregafdsRepositorio.asp";

      const $el = (id) => $(`#${id}`);

      let codigoExcluirSelecionado = "";

      let dadosCompletos = [];
      let dadosFiltrados = [];
      let paginaAtual    = 1;
      let totalPaginas   = 1;
      let limite         = 25;

      const TOAST_ICON = {
        sucesso: { icon: "bi-check-circle-fill", cor: "#198754" },
        erro:    { icon: "bi-x-circle-fill",      cor: "#dc3545" },
        aviso:   { icon: "bi-exclamation-triangle-fill", cor: "#856404", barra: "#ffc107" },
        info:    { icon: "bi-info-circle-fill",   cor: "#6c757d" }
      };
      const TOAST_DELAY = 4500;

      // --------------------------------------------------------------------------------------------
      function exibirToast(tipo, mensagem) {
        const cfg   = TOAST_ICON[tipo] || TOAST_ICON.info;
        const barra = cfg.barra || cfg.cor;
        const id    = "toast" + Date.now();

        const html = `
          <div class="toast-custom" id="${id}">
            <div class="toast-custom-body">
              <i class="bi ${cfg.icon} me-1" style="color:${cfg.cor};font-size:1.1rem;"></i>
              <div class="flex-grow-1" style="font-size:.85rem;color:#212529;">${mensagem}</div>
              <button type="button" class="toast-custom-close" onclick="this.closest('.toast-custom').remove()" aria-label="Close">&times;</button>
            </div>
            <div class="toast-progress-bar" style="height:3px;background-color:${barra};"></div>
          </div>
        `;

        $el("toastContainer").append(html);
        setTimeout(function () { $("#" + id).remove(); }, TOAST_DELAY);
      }

      // --------------------------------------------------------------------------------------------
      function carregarGrid() {
        const $tbody = $el("tbodyGrid");

        $tbody.html(`
          <tr>
            <td colspan="4" class="py-4">
              <div class="spinner-border text-secondary" role="status">
                <span class="visually-hidden">Carregando...</span>
              </div>
            </td>
          </tr>
        `);

        $.ajax({
          url: REPOSITORIO,
          method: "GET",
          dataType: "json",
          data: { acao: "listar" },
          success: function (resp) {
            if (!resp.sucesso) {
              exibirToast("erro", resp.mensagem || "Erro ao carregar a lista.");
              $tbody.html(`<tr><td colspan="4" class="py-4 text-danger">${resp.mensagem || "Erro ao carregar."}</td></tr>`);
              return;
            }
            dadosCompletos = resp.dados || [];
            paginaAtual = 1;
            aplicarFiltroEPaginar();
          },
          error: function () {
            exibirToast("erro", "Falha de comunicação ao carregar a lista.");
            $tbody.html(`<tr><td colspan="4" class="py-4 text-danger">Falha de comunicação.</td></tr>`);
          }
        });
      }

      // --------------------------------------------------------------------------------------------
      function aplicarFiltroEPaginar() {
        const termo = $el("searchBox").val().trim().toLowerCase();

        dadosFiltrados = !termo
          ? dadosCompletos
          : dadosCompletos.filter(function (item) {
              return String(item.codigo).toLowerCase().includes(termo)
                  || String(item.nome).toLowerCase().includes(termo);
            });

        totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / limite));
        if (paginaAtual > totalPaginas) paginaAtual = totalPaginas;

        const inicio = (paginaAtual - 1) * limite;
        const pagina = dadosFiltrados.slice(inicio, inicio + limite);

        renderizarGrid(pagina);
        renderizarPaginacao();
        $el("infoRegistros").text(`Total: ${dadosFiltrados.length} registro(s)`);
      }

      // --------------------------------------------------------------------------------------------
      function renderizarGrid(lista) {
        const $tbody = $el("tbodyGrid");

        if (!lista || lista.length === 0) {
          $tbody.html(`<tr><td colspan="4" class="py-4">Nenhum registro encontrado.</td></tr>`);
          return;
        }

        const linhas = lista.map(function (item) {
          const iconeAtivo = item.ativo
            ? "<i class='text-success bi bi-check-square'></i>"
            : "<i class='text-info bi bi-square'></i>";

          return `
            <tr>
              <td>
                <div class="d-flex justify-content-center gap-1">
                  <button type="button" class="btn btn-primary btn-sm btn-visualizar" title="Visualizar" data-codigo="${item.codigo}">
                    <i class="bi bi-eye"></i>
                  </button>
                  <button type="button" class="btn btn-warning btn-sm btn-editar" title="Editar" data-codigo="${item.codigo}">
                    <i class="bi bi-pencil"></i>
                  </button>
                  <button type="button" class="btn btn-danger btn-sm btn-excluir" title="Apagar" data-codigo="${item.codigo}" data-nome="${item.nome}">
                    <i class="bi bi-trash"></i>
                  </button>
                </div>
              </td>
              <td>${item.codigo}</td>
              <td>${item.nome}</td>
              <td>${iconeAtivo}</td>
            </tr>
          `;
        }).join("");

        $tbody.html(linhas);
      }

      // --------------------------------------------------------------------------------------------
      function renderizarPaginacao() {
        const $ul = $el("paginacao").empty();
        if (totalPaginas <= 1) return;

        const mkItem = (label, page, disabled, active) => {
          const $li = $("<li>").addClass("page-item");
          if (disabled) $li.addClass("disabled");
          if (active)   $li.addClass("active");
          const $a = $("<a>").addClass("page-link").attr("href", "#").html(label);
          if (!disabled && !active) {
            $a.on("click", function (e) {
              e.preventDefault();
              paginaAtual = page;
              aplicarFiltroEPaginar();
            });
          }
          return $li.append($a);
        };

        $ul.append(mkItem("&laquo;", 1, paginaAtual === 1, false));
        $ul.append(mkItem("&lsaquo;", paginaAtual - 1, paginaAtual === 1, false));

        const inicio = Math.max(1, paginaAtual - 2);
        const fim    = Math.min(totalPaginas, paginaAtual + 2);
        for (let p = inicio; p <= fim; p++) {
          $ul.append(mkItem(p, p, false, p === paginaAtual));
        }

        $ul.append(mkItem("&rsaquo;", paginaAtual + 1, paginaAtual === totalPaginas, false));
        $ul.append(mkItem("&raquo;", totalPaginas, paginaAtual === totalPaginas, false));
      }

      // --------------------------------------------------------------------------------------------
      function abrirModalNovo() {
        $el("formCadastro")[0].reset();
        $el("formCadastro").removeClass("was-validated");
        $el("modalFormAlerta").addClass("d-none").text("");
        $el("modo").val("criar");
        $el("modalFormTitulo").text("Novo");
        $el("codcliStatus").text("");

        $el("codcli").prop("readonly", false).val("");
        $el("btnPesqCli").prop("disabled", false);
        $el("razao").val("");
        $el("cbxAtivo").prop("disabled", false).prop("checked", false);

        $el("btnSalvar").removeClass("d-none");

        bootstrap.Modal.getOrCreateInstance(document.getElementById("modalForm")).show();
      }

      // --------------------------------------------------------------------------------------------
      function abrirModalEditarOuVisualizar(codigo, somenteVisualizar) {
        $.ajax({
          url: REPOSITORIO,
          method: "GET",
          dataType: "json",
          data: { acao: "obter", codigo: codigo },
          success: function (resp) {
            if (!resp.sucesso) {
              exibirToast("erro", resp.mensagem || "Não foi possível obter o registro.");
              return;
            }

            $el("formCadastro")[0].reset();
            $el("formCadastro").removeClass("was-validated");
            $el("modalFormAlerta").addClass("d-none").text("");
            $el("modo").val(somenteVisualizar ? "visualizar" : "editar");
            $el("modalFormTitulo").text(somenteVisualizar ? "Consulta" : "Alterar");
            $el("codcliStatus").text("");

            $el("codcli").prop("readonly", true).val(resp.dados.codigo);
            $el("btnPesqCli").prop("disabled", true);
            $el("razao").val(resp.dados.nome);
            $el("cbxAtivo").prop("disabled", somenteVisualizar).prop("checked", resp.dados.ativo === true);

            if (somenteVisualizar) {
              $el("btnSalvar").addClass("d-none");
            } else {
              $el("btnSalvar").removeClass("d-none");
            }

            bootstrap.Modal.getOrCreateInstance(document.getElementById("modalForm")).show();
          },
          error: function () {
            exibirToast("erro", "Falha de comunicação ao obter o registro.");
          }
        });
      }

      // --------------------------------------------------------------------------------------------
      function pesquisarCliente() {
        const codigo = $el("codcli").val().trim();

        if (codigo === "") {
          $el("razao").val("");
          $el("codcliStatus").text("");
          return;
        }

        $el("codcliStatus").text("Pesquisando...");

        $.ajax({
          url: REPOSITORIO,
          method: "GET",
          dataType: "json",
          data: { acao: "validarCliente", codigo: codigo },
          success: function (resp) {
            if (!resp.sucesso) {
              $el("razao").val("");
              $el("codcliStatus").addClass("text-danger").text(resp.mensagem || "Cliente não encontrado.");
              return;
            }
            $el("codcliStatus").removeClass("text-danger").text("");
            $el("razao").val(resp.dados.nome);
          },
          error: function () {
            $el("codcliStatus").addClass("text-danger").text("Falha de comunicação ao validar cliente.");
          }
        });
      }

      // --------------------------------------------------------------------------------------------
      function salvar() {
        const form = $el("formCadastro")[0];

        if (form.checkValidity() === false) {
          form.classList.add("was-validated");
          return;
        }

        const modo   = $el("modo").val();
        const codcli = $el("codcli").val().trim();
        const ativo  = $el("cbxAtivo").is(":checked") ? "1" : "0";

        if ($el("razao").val().trim() === "") {
          $el("modalFormAlerta").removeClass("d-none").text("Cliente não localizado. Verifique o código informado.");
          return;
        }

        const acao = (modo === "criar") ? "criar" : "alterar";

        $el("spinnerSalvar").removeClass("d-none");
        $el("btnSalvar").prop("disabled", true);

        $.ajax({
          url: REPOSITORIO,
          method: "POST",
          dataType: "json",
          data: { acao: acao, codcli: codcli, ativo: ativo },
          success: function (resp) {
            $el("spinnerSalvar").addClass("d-none");
            $el("btnSalvar").prop("disabled", false);

            if (!resp.sucesso) {
              $el("modalFormAlerta").removeClass("d-none").text(resp.mensagem || "Erro ao salvar.");
              return;
            }

            bootstrap.Modal.getOrCreateInstance(document.getElementById("modalForm")).hide();
            exibirToast("sucesso", resp.mensagem || "Registro salvo com sucesso!");
            carregarGrid();
          },
          error: function () {
            $el("spinnerSalvar").addClass("d-none");
            $el("btnSalvar").prop("disabled", false);
            $el("modalFormAlerta").removeClass("d-none").text("Falha de comunicação ao salvar.");
          }
        });
      }

      // --------------------------------------------------------------------------------------------
      function abrirModalExcluir(codigo, nome) {
        codigoExcluirSelecionado = codigo;
        $el("excluirCodigo").text(codigo);
        $el("excluirNome").text(nome || "");

        bootstrap.Modal.getOrCreateInstance(document.getElementById("modalExcluir")).show();
      }

      // --------------------------------------------------------------------------------------------
      function confirmarExcluir() {
        if (codigoExcluirSelecionado === "") return;

        $el("spinnerExcluir").removeClass("d-none");
        $el("btnConfirmarExcluir").prop("disabled", true);

        $.ajax({
          url: REPOSITORIO,
          method: "POST",
          dataType: "json",
          data: { acao: "excluir", codcli: codigoExcluirSelecionado },
          success: function (resp) {
            $el("spinnerExcluir").addClass("d-none");
            $el("btnConfirmarExcluir").prop("disabled", false);

            bootstrap.Modal.getOrCreateInstance(document.getElementById("modalExcluir")).hide();

            if (!resp.sucesso) {
              exibirToast("erro", resp.mensagem || "Erro ao apagar registro.");
              return;
            }

            exibirToast("sucesso", resp.mensagem || "Registro apagado com sucesso!");
            carregarGrid();
          },
          error: function () {
            $el("spinnerExcluir").addClass("d-none");
            $el("btnConfirmarExcluir").prop("disabled", false);
            exibirToast("erro", "Falha de comunicação ao apagar registro.");
          }
        });
      }

      // --------------------------------------------------------------------------------------------
      function init() {
        limite = parseInt($el("selectLimite").val(), 10) || 25;

        carregarGrid();

        $el("btnNovo").on("click", abrirModalNovo);
        $el("btnSalvar").on("click", salvar);
        $el("btnConfirmarExcluir").on("click", confirmarExcluir);
        $el("codcli").on("blur", pesquisarCliente);
        $el("codcli").on("focus", function () { $el("codcliStatus").removeClass("text-danger").text(""); });

        let debounceBusca = null;
        $el("searchBox").on("input", function () {
          clearTimeout(debounceBusca);
          debounceBusca = setTimeout(function () {
            paginaAtual = 1;
            aplicarFiltroEPaginar();
          }, 250);
        });

        $el("selectLimite").on("change", function () {
          limite = parseInt($(this).val(), 10) || 25;
          paginaAtual = 1;
          aplicarFiltroEPaginar();
        });

        $(document).on("keydown", function (e) {
          if (e.ctrlKey && e.key === "k") {
            e.preventDefault();
            $el("searchBox").focus();
          }
        });

        $(document).on("click", ".btn-editar", function () {
          abrirModalEditarOuVisualizar($(this).data("codigo"), false);
        });

        $(document).on("click", ".btn-visualizar", function () {
          abrirModalEditarOuVisualizar($(this).data("codigo"), true);
        });

        $(document).on("click", ".btn-excluir", function () {
          abrirModalExcluir($(this).data("codigo"), $(this).data("nome"));
        });
      }

      return Object.freeze({ init });
    })();

    $(document).ready(function () {
      DefEntregaFDS.init();
    });
  </script>

</body>
</html>
