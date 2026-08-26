<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp"-->
<!--#include file="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp"-->
<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046

Dim IdUsuarioLogado, LoginUsuarioLogado
IdUsuarioLogado    = Session("CODIGO")
LoginUsuarioLogado = Session("USR")

' if ValidaPermissaoRelatorios("",CODUSER,FILIAL) then
'   Response.Write "Você não tem permissão de acesso! Por favor, entre em contato com o administrador do sistema."
'   Response.End
' end if
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet &mdash; Relat&oacute;rio de Pedidos em Carteira</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>

    body {
      background-color: #f5f5f5;
      font-size: 0.8rem;
      padding-bottom: 20px;
    }
    .container {
      max-width: 100%;
    }
    .form-label {
      font-weight: 600;
    }
    .table-wrapper {
      overflow: auto;
      max-height: 70vh;
    }
    .card {
      position: relative;
    }
    #gradeRelatorio {
      width: max-content;
      min-width: 100%;
    }
    #gradeRelatorio th,
    #gradeRelatorio td {
      white-space: nowrap;
      vertical-align: middle;
      padding: 6px 10px !important;
    }
    thead {
      vertical-align: middle !important;
      position: sticky;
      inset-block-start: 0;
      background: #f5f5f5;
      z-index: 10;
    }
    tfoot {
      vertical-align: middle !important;
      position: sticky;
      inset-block-end: 0;
      bottom: 0;
      background: #f5f5f5;
      z-index: 10;
      font-weight: 600;
      color: #495057;
    }
    td,
    th {
      padding: 6px 10px !important;
    }
    th {
      font-weight: 600;
      color: #495057;
    }
    tr:hover {
      background: #f5f5f5;
    }
    .loading-overlay {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.8);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    }

    #pedido {
      text-transform: uppercase;
    }
    .filtro-tabela-wrapper {
      max-width: 320px;
    }
    #inputFiltroTabela:focus {
      border-color: #0d6efd;
      box-shadow: 0 0 0 0.2rem rgba(13, 110, 253, 0.2);
    }
    .linha-oculta {
      display: none !important;
    }
    .highlight-match {
      background-color: #fff3cd;
      border-radius: 2px;
    }
    .pedido-detalhe-link {
      cursor: pointer;
    }

  </style>

</head>

<body onload="RelatorioPedidosEmCarteira.init()">

  <div class="container my-4">

    <div class="d-flex justify-content-between align-items-center mb-4">
      <h3 class="mb-0">Relat&oacute;rio de Pedidos em Carteira</h3>
    </div>

    <form name="PedidosEmCarteira" id="PedidosEmCarteira" onsubmit="return false;">

      <div class="card p-3 mb-4">
        <div class="row g-3">

          <div class="col-12 col-md-4 col-lg-3">
            <label for="cboCidade" class="form-label">Filial</label>
            <select required name="cboCidade" id="cboCidade" class="form-select form-select-sm">
              <option value="" selected>Selecione...</option>
              <option value="PIRAPETINGA">Pirapetinga</option>
              <option value="UBERABA">Uberaba</option>
              <option value="BENTO">Bento Gon&ccedil;alves</option>
              <option value="MARANGUAPE">Maranguape</option>
              <option value="SAQUAREMA">Saquarema</option>
              <option value="ESHOP">ESHOP</option>
            </select>
          </div>

            <div class="col-12 col-md-4 col-lg-3">
              <label for="cboPeriodoPor" class="form-label">Per&iacute;odo por</label>
              <select name="cboPeriodoPor" id="cboPeriodoPor" class="form-select form-select-sm">
                <option value="DTEMI" selected>Data de Emiss&atilde;o</option>
                <option value="DTENT">Data de Entrega</option>
                <option value="DTEMB">Data de Embarque</option>
              </select>
            </div>
          <div class="col-12 col-md-4 col-lg-3">
            <label for="dataini" class="form-label">Data Inicial</label>
            <input id="dataini" name="dataini" type="date" class="form-control form-control-sm" />
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="datafin" class="form-label">Data Final</label>
            <input id="datafin" name="datafin" type="date" class="form-control form-control-sm" />
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="razao" class="form-label">Cliente</label>
            <div class="input-group input-group-sm">
              <button type="button" class="btn btn-secondary" id="btnBuscarCliente" title="Buscar Cliente">
                <i class="bi bi-box-arrow-up-right me-1"></i>Buscar
              </button>
              <input name="razao" id="razao" type="text" class="form-control" readonly placeholder="Raz&atilde;o Social" />
              <input type="text" name="lojacli" id="lojacli" class="form-control" style="max-width:70px" readonly placeholder="Loja" />
              <button type="button" class="btn btn-outline-secondary" id="btnLimparCliente" title="Limpar Cliente">
                <i class="bi bi-x-lg"></i>
              </button>
            </div>
            <input type="text" name="codcli" id="codcli" style="display:none" />
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="cboGrupoProduto" class="form-label">Grupo de Produto</label>
            <select name="cboGrupoProduto" id="cboGrupoProduto" class="form-select form-select-sm">
              <option value="" selected>Todos</option>
              <option value="PAPEL">Papel</option>
              <option value="PAPELÃO">Papel&atilde;o</option>
            </select>
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="cboTipoProduto" class="form-label">Tipo de Produto</label>
            <select name="cboTipoProduto" id="cboTipoProduto" class="form-select form-select-sm">
              <option value="" selected>Todos</option>
              <option value="CAIXA">Caixa</option>
              <option value="CHAPA">Chapa</option>
              <option value="CORTE VINCO">Corte Vinco</option>
              <option value="MALETA">Maleta</option>
            </select>
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="cboTipoFrete" class="form-label">Tipo de Frete</label>
            <select name="cboTipoFrete" id="cboTipoFrete" class="form-select form-select-sm">
              <option value="" selected>Todos</option>
              <option value="CIF">CIF</option>
              <option value="FOB">FOB</option>
            </select>
          </div>


          <div class="col-12 col-md-4 col-lg-3">
            <label for="representante" class="form-label">Representante</label>
            <input type="text" name="representante" id="representante" class="form-control form-control-sm" placeholder="C&oacute;digo" maxlength="6" />
          </div>

          <div class="col-12 col-md-4 col-lg-3">
            <label for="pedido" class="form-label">Pedido</label>
            <input type="text" name="pedido" id="pedido" class="form-control form-control-sm" />
          </div>

          <div class="col-12 col-md-4 col-lg-6 d-flex justify-content-end align-items-end gap-2">
            <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
              <i class="bi bi-x-lg me-1"></i>Limpar
            </button>

            <button type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
              <i class="bi bi-search me-1"></i>Pesquisar
            </button>
          </div>

        </div>
      </div>
    </form>

    <div class="card">
      <div id="overlayLoading" class="loading-overlay d-none">
        <div class="spinner-border text-primary" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
      </div>
      <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
        <div class="d-flex gap-2">
          <button type="button" class="btn btn-success btn-sm" id="btnExcel">
            <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
          </button>
          <button type="button" class="btn btn-danger btn-sm" id="btnPdf">
            <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
          </button>
        </div>
        <div class="input-group input-group-sm filtro-tabela-wrapper">
          <span class="input-group-text"><i class="bi bi-search"></i></span>
          <input type="text" id="inputFiltroTabela" class="form-control" placeholder="Filtrar tabela… (Ctrl+K)" autocomplete="off" />
          <button type="button" class="btn btn-outline-secondary d-none" id="btnLimparFiltroTabela" title="Limpar filtro">
            <i class="bi bi-x-lg"></i>
          </button>
        </div>
      </div>

      <div class="table-wrapper">
        <table class="table table-hover table-striped table-bordered table-sm mb-0" id="gradeRelatorio">
          <thead>
            <tr class="text-center">
              <th>N&uacute;mero</th>
              <th></th>
              <th>Data</th>
              <th>Data</th>
              <th>Data</th>
              <th colspan="5">Cliente</th>
              <th>Tipo</th>
              <th>Cond.de</th>
              <th colspan="6">Produto</th>
              <th colspan="2">Representante</th>
              <th>Qtde.</th>
              <th>Saldo</th>
              <th>Valor</th>
              <th>&Aacute;rea</th>
              <th>Peso</th>
              <th>Valor</th>
              <th>Valor</th>
              <th>&Aacute;rea</th>
              <th>Peso</th>
              <th>Saldo</th>
              <th>Saldo</th>
              <th>Saldo</th>
              <th>Saldo</th>
              <th>Saldo</th>
              <th>Pre&ccedil;o</th>
              <th>Pre&ccedil;o</th>
            </tr>
            <tr class="text-center">
              <th>Pedido</th>
              <th>TES</th>
              <th>Emiss&atilde;o</th>
              <th>Entrega</th>
              <th>Embarque</th>
              <th>C&oacute;digo</th>
              <th>Loja</th>
              <th>Nome</th>
              <th>Cidade</th>
              <th>Pedido</th>
              <th>Frete</th>
              <th>Pagamento</th>
              <th>C&oacute;digo</th>
              <th>Descri&ccedil;&atilde;o</th>
              <th>Estrutura</th>
              <th>Grupo</th>
              <th>Tipo</th>
              <th>M&aacute;quina</th>
              <th>C&oacute;digo</th>
              <th>Nome</th>
              <th>Vendida</th>
              <th>Entregar</th>
              <th>Unit&aacute;rio</th>
              <th>Unit&aacute;ria</th>
              <th>Unit&aacute;rio</th>
              <th>Total</th>
              <th>NET</th>
              <th>Vendida</th>
              <th>Vendido</th>
              <th>&Aacute;rea</th>
              <th>Peso</th>
              <th>Total S/IPI</th>
              <th>Total C/IPI</th>
              <th>Valor NET</th>
              <th>KG NET</th>
              <th>M&sup2;.NET</th>
            </tr>
          </thead>

          <tbody id="corpoTabela">
            <tr>
              <td colspan="36" class="text-center text-muted py-4">
                Utilize os filtros acima para pesquisar
              </td>
            </tr>
          </tbody>

          <tfoot id="rodapeTabela">
          </tfoot>
        </table>
      </div>
    </div>

  </div>

  <!-- Modal de Detalhes de Produção -->
  <div class="modal fade" id="modalDetalhesPedido" tabindex="-1" aria-labelledby="modalDetalhesPedidoTitulo" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="modalDetalhesPedidoTitulo">
            Detalhes de Produção #<span id="numeroPedidoDetalhe"></span>
          </h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <div class="table-responsive">
            <table class="table table-hover table-striped table-bordered table-sm mb-0" id="gradeDetalhesPedido">
              <thead>
                <tr>
                  <th>OP</th>
                  <th>Qtd. Pedida</th>
                  <th>Qtd. Programada</th>
                  <th>Qtd. Ondulada</th>
                  <th>Qtd. Produzida</th>
                  <th>Qtd. Aproveitada Estoque</th>
                  <th>Qtd. Expedida</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td colspan="7" class="text-center text-muted py-4">Nenhum registro encontrado</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.5.31/dist/jspdf.plugin.autotable.min.js"></script>

  <script>
     (() => {
      "use strict";
    
      const REPOSITORIO = "RelatorioPedidosEmCarteiraRepositorio.asp";
      const CAMINHO_LOGO = "/corporativo/comercial/Logo/logo-smurfit1.png";
      const ID_USUARIO = "<%=IdUsuarioLogado%>";
      const LOGIN_USUARIO = "<%=LoginUsuarioLogado%>";
    
      const COLS = Object.freeze([
        "Pedido",
        "TES",
        "Emissão",
        "Entrega",
        "Embarque",
        "Cód. Cliente",
        "Loja",
        "Nome Cliente",
        "Cidade",
        "Pedido Cliente",
        "Tipo Frete",
        "Cond. Pagamento",
        "Cód. Produto",
        "Descrição",
        "Estrutura",
        "Grupo",
        "Tipo",
        "Máquina",
        "Cód. Rep.",
        "Nome Rep.",
        "Qtde. Vendida",
        "Saldo Entregar",
        "Valor Unitário",
        "Área Unitária",
        "Peso Unitário",
        "Valor Total",
        "Valor NET",
        "Área Vendida",
        "Peso Vendido",
        "Área Saldo",
        "Peso Saldo",
        "Saldo S/IPI",
        "Saldo C/IPI",
        "Valor NET Saldo",
        "Preço KG NET",
        "Preço M² NET",
      ]);
    
      // Estado local dos dados retornados pelo repositório
      let _dadosCompletos = [];
      let _totaisCompletos = null;
    
      const $id = (id) => document.getElementById(id);
    
      function exibirOverlay() {
        $id("overlayLoading").classList.remove("d-none");
      }
      function ocultarOverlay() {
        $id("overlayLoading").classList.add("d-none");
      }
    
      // Formata número no padrão pt-BR com N casas decimais
      function fmtNum(valor, decimais) {
        const n = Number(valor) || 0;
        return n.toLocaleString("pt-BR", {
          minimumFractionDigits: decimais,
          maximumFractionDigits: decimais,
        });
      }
    
      function escapeHtml(texto) {
        if (texto === null || texto === undefined) return "";
        return String(texto)
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }
    
      // AJAX/XHR — chamadas ao repositório
      function _xhrPost(params, onSucesso, onErro) {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", REPOSITORIO, true);
        xhr.setRequestHeader(
          "Content-Type",
          "application/x-www-form-urlencoded; charset=UTF-8",
        );
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4) return;
          if (xhr.status >= 200 && xhr.status < 300) {
            try {
              const obj = JSON.parse(xhr.responseText);
              onSucesso(obj);
            } catch (e) {
              if (onErro) onErro("Resposta inválida do servidor.");
            }
          } else {
            if (onErro) onErro("Falha na comunicação (HTTP " + xhr.status + ").");
          }
        };
    
        const corpo = Object.keys(params)
          .map(
            (k) =>
              encodeURIComponent(k) +
              "=" +
              encodeURIComponent(params[k] == null ? "" : params[k]),
          )
          .join("&");
    
        xhr.send(corpo);
      }
    
      function abrirModalDetalhesPedido(pedido) {
        const cidade = $id("cboCidade").value;
        $id("numeroPedidoDetalhe").textContent = pedido;
        const modal = new bootstrap.Modal($id("modalDetalhesPedido"));
        modal.show();
        const tbody = $id("gradeDetalhesPedido").querySelector("tbody");
        tbody.innerHTML =
          '<tr><td colspan="7" class="text-center text-muted py-4">' +
          '<div class="spinner-border spinner-border-sm text-primary me-2" role="status"></div>' +
          "Carregando...</td></tr>";

        _xhrPost(
          {
            acao: "obterdetalhes",
            pedido: pedido,
            cidade: cidade,
          },
          function (response) {
            if (!response || response.sucesso !== true) {
              alert((response && response.mensagem) || "Erro ao buscar detalhes do pedido.");
              tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">Nenhum registro encontrado</td></tr>';
              return;
            }

            const dados = response.dados || [];
            if (dados.length === 0) {
              tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">Nenhum registro encontrado</td></tr>';
              return;
            }

            tbody.innerHTML = dados.map(function (item) {
              return (
                "<tr>" +
                "<td>" + escapeHtml(item.OP || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadePedida || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadeProgramada || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadeOndulada || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadeProduzida || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadeAproveitadaEstoque || "-") + "</td>" +
                '<td class="text-end">' + escapeHtml(item.quantidadeExpedida || "-") + "</td>" +
                "</tr>"
              );
            }).join("");
          },
          function (msg) {
            alert(msg);
            tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">Nenhum registro encontrado</td></tr>';
          },
        );
      }

      // RENDERIZAÇÃO DA TABELA
      function _linhaVazia(mensagem) {
        return (
          '<tr><td colspan="36" class="text-center text-muted py-4">' +
          '<i class="bi bi-inbox fs-4"></i><div class="mt-2">' +
          mensagem +
          "</div>" +
          "</td></tr>"
        );
      }
    
      function _renderizarCorpo(dados) {
        const tbody = $id("corpoTabela");
    
        if (!dados || dados.length === 0) {
          tbody.innerHTML = _linhaVazia(
            "Nenhum registro encontrado para esse filtro.",
          );
          return;
        }
    
        const linhas = dados.map((r) => {
          return (
            "<tr>" +
            '<td><a href="#" class="pedido-detalhe-link" data-pedido="' +

            escapeHtml(r.pedido) +

            '">' +

            escapeHtml(r.pedido) +

            "</a></td>" +
            "<td>" +
            escapeHtml(r.tes) +
            "</td>" +
            "<td>" +
            escapeHtml(r.emissao) +
            "</td>" +
            "<td>" +
            escapeHtml(r.entrega) +
            "</td>" +
            "<td>" +
            escapeHtml(r.embarque) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.codCli) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.lojaCli) +
            "</td>" +
            "<td>" +
            escapeHtml(r.nomeCli) +
            "</td>" +
            "<td>" +
            escapeHtml(r.cidade) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.pedidoCliente) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.tipoFrete) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.condPagamento) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.produto) +
            "</td>" +
            "<td>" +
            escapeHtml(r.descricaoProduto) +
            "</td>" +
            "<td>" +
            escapeHtml(r.estrutura) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.grupoProduto) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.tipoProduto) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.maquina) +
            "</td>" +
            '<td class="text-center">' +
            escapeHtml(r.codRepre) +
            "</td>" +
            "<td>" +
            escapeHtml(r.nomeRepre) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.quantVendida, 0) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.quantSaldoEntregar, 0) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.valorUnitario, 4) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.areaUnitaria, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.pesoUnitario, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.valorTotal, 2) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.valorNet, 2) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.areaVendida, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.pesoVendido, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.areaSaldo, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.pesoSaldo, 3) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.saldoSemIPI, 2) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.saldoComIPI, 2) +
            "</td>" +
            '<td class="text-end">' +
            fmtNum(r.valorNetSaldo, 2) +
            "</td>" +
            '<td class="text-end">' +
            escapeHtml(r.precoKgNet) +
            "</td>" +
            '<td class="text-end">' +
            escapeHtml(r.precoM2Net) +
            "</td>" +
            "</tr>"
          );
        });
    
        tbody.innerHTML = linhas.join("");
      }
    
      function _renderizarRodape(totais) {
        const tfoot = $id("rodapeTabela");
    
        if (!totais) {
          tfoot.innerHTML = "";
          return;
        }
    
        const labelEstilo =
          'class="text-muted me-1 fst-italic" style="font-size:.7rem"';
        const tdVazios = "<td></td>".repeat(23);
    
        tfoot.innerHTML =
          "<tr>" +
          '<td class="text-center fw-bold" colspan="2">T O T A I S</td>' +
          tdVazios +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Valor Total:</span>" +
          fmtNum(totais.valorTotal, 2) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Valor NET:</span>" +
          fmtNum(totais.valorNet, 2) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Área Vendida:</span>" +
          fmtNum(totais.areaVendida, 3) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Peso Vendido:</span>" +
          fmtNum(totais.pesoVendido, 3) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Área Saldo:</span>" +
          fmtNum(totais.areaSaldo, 3) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Peso Saldo:</span>" +
          fmtNum(totais.pesoSaldo, 3) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Saldo S/IPI:</span>" +
          fmtNum(totais.saldoSemIPI, 2) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Saldo C/IPI:</span>" +
          fmtNum(totais.saldoComIPI, 2) +
          "</td>" +
          '<td class="text-end"><span ' +
          labelEstilo +
          ">Valor NET Saldo:</span>" +
          fmtNum(totais.valorNetSaldo, 2) +
          "</td>" +
          "<td></td><td></td>" +
          "</tr>";
      }
    
      // EXPORTAÇÕES
      // Helpers locais (mesmo padrão do funcoesJAVASCRIPT.js usado em outros relatórios)
      function escapeCSV(valor) {
        if (valor === null || valor === undefined) return "";
        const txt = String(valor);
        if (/[;"\r\n]/.test(txt)) {
          return '"' + txt.replace(/"/g, '""') + '"';
        }
        return txt;
      }

      function formatarNumCSV(valor, decimais) {
        const n = Number(valor);
        if (!isFinite(n)) return "0";
        const d = (decimais === undefined || decimais === null) ? 2 : decimais;
        // Sem separador de milhar, vírgula decimal — Excel PT-BR reconhece como número somável
        return n.toFixed(d).replace(".", ",");
      }

      function formatarMoeda(valor, decimais) {
        const n = Number(valor);
        if (!isFinite(n)) return "0,00";
        const d = (decimais === undefined || decimais === null) ? 2 : decimais;
        return n.toLocaleString("pt-BR", {
          minimumFractionDigits: d,
          maximumFractionDigits: d
        });
      }

      function formatarDataCSV(valor) {
        if (!valor || String(valor).trim() === "" || String(valor).trim() === "-") return "-";
        // Prefixo ="..." força o Excel a tratar como texto, evitando conversão para número de série (######)
        return '="' + String(valor).trim().replace(/"/g, '""') + '"';
      }

      function exportarExcel() {
        if (_dadosCompletos.length === 0) {
          alert("Não há dados para exportar");
          return;
        }

        const nomeArquivo = "Pedidos_Em_Carteira.csv";
        const colunas = [
          "Pedido",
          "TES",
          "Emissão",
          "Entrega",
          "Embarque",
          "Cód. Cliente",
          "Loja",
          "Nome Cliente",
          "Cidade",
          "Pedido Cliente",
          "Tipo Frete",
          "Cond. Pagamento",
          "Cód. Produto",
          "Descrição",
          "Estrutura",
          "Grupo",
          "Tipo",
          "Máquina",
          "Cód. Rep.",
          "Nome Rep.",
          "Qtde. Vendida",
          "Saldo Entregar",
          "Valor Unitário",
          "Área Unitária",
          "Peso Unitário",
          "Valor Total",
          "Valor NET",
          "Área Vendida",
          "Peso Vendido",
          "Área Saldo",
          "Peso Saldo",
          "Saldo S/IPI",
          "Saldo C/IPI",
          "Valor NET Saldo",
          "Preço KG NET",
          "Preço M² NET"
        ];

        let csvContent = "\uFEFF";
        csvContent += colunas.join(";") + "\r\n";

        _dadosCompletos.forEach(item => {
          const linha = [
            escapeCSV(item.pedido || "-"),
            escapeCSV(item.tes || "-"),
            escapeCSV(item.emissao),
            escapeCSV(item.entrega),
            escapeCSV(item.embarque),
            escapeCSV(item.codCli || "-"),
            escapeCSV(item.lojaCli || "-"),
            escapeCSV(item.nomeCli || "-"),
            escapeCSV(item.cidade || "-"),
            escapeCSV(item.pedidoCliente || "-"),
            escapeCSV(item.tipoFrete || "-"),
            escapeCSV(item.condPagamento || "-"),
            escapeCSV(item.produto || "-"),
            escapeCSV(item.descricaoProduto || "-"),
            escapeCSV(item.estrutura || "-"),
            escapeCSV(item.grupoProduto || "-"),
            escapeCSV(item.tipoProduto || "-"),
            escapeCSV(item.maquina || "-"),
            escapeCSV(item.codRepre || "-"),
            escapeCSV(item.nomeRepre || "-"),
            escapeCSV(formatarNumCSV(item.quantVendida, 0)),
            escapeCSV(formatarNumCSV(item.quantSaldoEntregar, 0)),
            escapeCSV(formatarNumCSV(item.valorUnitario, 5)),
            escapeCSV(formatarNumCSV(item.areaUnitaria, 5)),
            escapeCSV(formatarNumCSV(item.pesoUnitario, 5)),
            escapeCSV(formatarNumCSV(item.valorTotal, 5)),
            escapeCSV(formatarNumCSV(item.valorNet, 5)),
            escapeCSV(formatarNumCSV(item.areaVendida, 5)),
            escapeCSV(formatarNumCSV(item.pesoVendido, 5)),
            escapeCSV(formatarNumCSV(item.areaSaldo, 5)),
            escapeCSV(formatarNumCSV(item.pesoSaldo, 5)),
            escapeCSV(formatarNumCSV(item.saldoSemIPI, 5)),
            escapeCSV(formatarNumCSV(item.saldoComIPI, 5)),
            escapeCSV(formatarNumCSV(item.valorNetSaldo, 5)),
            escapeCSV(formatarNumCSV(item.precoKgNet, 5)),
            escapeCSV(formatarNumCSV(item.precoM2Net, 5))
          ];
          csvContent += linha.join(";") + "\r\n";
        });

        const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
        const link = document.createElement("a");

        if (navigator.msSaveBlob) {
          navigator.msSaveBlob(blob, nomeArquivo);
        } else {
          const url = URL.createObjectURL(blob);
          link.href = url;
          link.download = nomeArquivo;
          link.style.display = "none";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
        }
      }
    
      // Monta string legível com os filtros que o usuário aplicou na pesquisa
      function _textoFiltrosAtivos() {
        const frm = document.PedidosEmCarteira;
        const partes = [];
    
        const labelPeriodo = {
          DTEMI: "Dt. Emissão",
          DTENT: "Dt. Entrega",
          DTEMB: "Dt. Embarque",
        };
    
        const addFiltro = (label, valor) => {
          if (valor && valor.trim() !== "")
            partes.push(label + ": " + valor.trim());
        };
    
        // Filial
        const cidadeOpt = frm.cboCidade.options[frm.cboCidade.selectedIndex];
        if (frm.cboCidade.value)
          addFiltro("Filial", cidadeOpt ? cidadeOpt.text : frm.cboCidade.value);
    
        // Período
        const periodo =
          labelPeriodo[frm.cboPeriodoPor.value] || frm.cboPeriodoPor.value;
        const di = frm.dataini.value
          ? new Date(frm.dataini.value + "T00:00:00").toLocaleDateString("pt-BR")
          : "";
        const df = frm.datafin.value
          ? new Date(frm.datafin.value + "T00:00:00").toLocaleDateString("pt-BR")
          : "";
        if (di && df) partes.push(periodo + ": " + di + " a " + df);
    
        // Cliente
        if (frm.codcli.value.trim()) {
          const display = frm.razao.value.trim() || frm.codcli.value.trim();
          addFiltro("Cliente", display);
        }
    
        // Selects opcionais
        if (frm.cboGrupoProduto.value)
          addFiltro(
            "Grupo Produto",
            frm.cboGrupoProduto.options[frm.cboGrupoProduto.selectedIndex].text,
          );
        if (frm.cboTipoProduto.value)
          addFiltro(
            "Tipo Produto",
            frm.cboTipoProduto.options[frm.cboTipoProduto.selectedIndex].text,
          );
        if (frm.cboTipoFrete.value) addFiltro("Frete", frm.cboTipoFrete.value);
        addFiltro("Representante", frm.representante.value);
        addFiltro("Pedido", frm.pedido.value);
    
        return partes.length
          ? partes.join("   |   ")
          : "Nenhum filtro adicional aplicado";
      }
    
      function _dadosParaAoA() {
        return _dadosCompletos.map(function (r) {
          return [
            r.pedido           || "",
            r.tes              || "",
            r.emissao          || "",
            r.entrega          || "",
            r.embarque         || "",
            r.codCli           || "",
            r.lojaCli          || "",
            r.nomeCli          || "",
            r.cidade           || "",
            r.pedidoCliente    || "",
            r.tipoFrete        || "",
            r.condPagamento    || "",
            r.produto          || "",
            r.descricaoProduto || "",
            r.estrutura        || "",
            r.grupoProduto     || "",
            r.tipoProduto      || "",
            r.maquina          || "",
            r.codRepre         || "",
            r.nomeRepre        || "",
            fmtNum(r.quantVendida,        0),
            fmtNum(r.quantSaldoEntregar,  0),
            fmtNum(r.valorUnitario,       4),
            fmtNum(r.areaUnitaria,        3),
            fmtNum(r.pesoUnitario,        3),
            fmtNum(r.valorTotal,          2),
            fmtNum(r.valorNet,            2),
            fmtNum(r.areaVendida,         3),
            fmtNum(r.pesoVendido,         3),
            fmtNum(r.areaSaldo,           3),
            fmtNum(r.pesoSaldo,           3),
            fmtNum(r.saldoSemIPI,         2),
            fmtNum(r.saldoComIPI,         2),
            fmtNum(r.valorNetSaldo,       2),
            r.precoKgNet != null ? String(r.precoKgNet) : "",
            r.precoM2Net != null ? String(r.precoM2Net) : "",
          ];
        });
      }

      function _totaisParaAoA() {
        if (!_totaisCompletos) return null;
        const t = _totaisCompletos;
        const vazio = "";
        return [
          "TOTAIS", vazio, vazio, vazio, vazio,
          vazio, vazio, vazio, vazio, vazio,
          vazio, vazio, vazio, vazio, vazio,
          vazio, vazio, vazio, vazio, vazio,
          vazio, vazio, vazio, vazio, vazio,
          fmtNum(t.valorTotal,    2),
          fmtNum(t.valorNet,      2),
          fmtNum(t.areaVendida,   3),
          fmtNum(t.pesoVendido,   3),
          fmtNum(t.areaSaldo,     3),
          fmtNum(t.pesoSaldo,     3),
          fmtNum(t.saldoSemIPI,   2),
          fmtNum(t.saldoComIPI,   2),
          fmtNum(t.valorNetSaldo, 2),
          vazio, vazio,
        ];
      }

      function _gerarAutoTable(doc, pageW, startY) {
        const dadosAoA = _dadosParaAoA();
        const totaisRow = _totaisParaAoA();
        const corpo = totaisRow ? [...dadosAoA, totaisRow] : dadosAoA;
    
        doc.autoTable({
          head: [COLS],
          body: corpo,
          startY: startY,
          theme: "striped",
          styles: { fontSize: 4, cellPadding: 1.2, overflow: "linebreak" },
          headStyles: {
            fillColor: [74, 74, 74],
            textColor: 255,
            fontStyle: "bold",
            fontSize: 4.5,
          },
          alternateRowStyles: { fillColor: [245, 245, 245] },
          didParseCell(data) {
            if (
              totaisRow &&
              data.section === "body" &&
              data.row.index === corpo.length - 1
            ) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [212, 212, 212];
              data.cell.styles.textColor = [0, 0, 0];
            }
          },
          margin: { left: 5, right: 5 },
        });
    
        doc.save("RelatorioPedidosEmCarteira.pdf");
        ocultarOverlay();
      }
    
      function exportarPdf() {
        if (!_dadosCompletos.length) {
          alert("Nenhum dado para exportar.");
          return;
        }
        exibirOverlay();
    
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
          orientation: "landscape",
          unit: "mm",
          format: "a3",
        });
        const pageW = doc.internal.pageSize.getWidth();
        const cx = pageW / 2;
        const hoje = new Date().toLocaleDateString("pt-BR");
    
        const _renderHeader = (logoData) => {
          if (logoData) {
            doc.addImage(logoData, "PNG", 5, 4, 38, 13);
          }
    
          // Título
          doc.setFont("helvetica", "bold");
          doc.setFontSize(13);
          doc.text("Relatório de Pedidos em Carteira", cx, 11, { align: "center" });
    
          // Subtítulo: data + contagem
          doc.setFont("helvetica", "normal");
          doc.setFontSize(8);
          doc.text(
            "Gerado em: " +
              hoje +
              "  |  " +
              _dadosCompletos.length +
              " registro(s)",
            cx,
            18,
            { align: "center" },
          );
    
          // Usuário (canto direito)
          doc.setFontSize(7);
          doc.text(
            "Usuário: " + LOGIN_USUARIO + " [" + ID_USUARIO + "]",
            pageW - 5,
            11,
            { align: "right" },
          );
    
          // Bloco de filtros ativos
          const textoFiltros = _textoFiltrosAtivos();
          const maxLargura = pageW - 10; // margem 5mm de cada lado
          const linhasFiltro = doc.splitTextToSize(
            "Filtros: " + textoFiltros,
            maxLargura,
          );
    
          doc.setFont("helvetica", "italic");
          doc.setFontSize(6.5);
          doc.setTextColor(80, 80, 80);
    
          // Fundo cinza claro atrás do bloco de filtros
          const alturaLinha = 4;
          const paddingV = 1.5;
          const blocoH = linhasFiltro.length * alturaLinha + paddingV * 2;
          const blocoY = 21;
          doc.setFillColor(240, 240, 240);
          doc.rect(5, blocoY, pageW - 10, blocoH, "F");
    
          doc.text(linhasFiltro, 7, blocoY + paddingV + alturaLinha * 0.8);
          doc.setTextColor(0, 0, 0);
          doc.setFont("helvetica", "normal");
    
          // autoTable começa abaixo do bloco de filtros
          _gerarAutoTable(doc, pageW, blocoY + blocoH + 2);
        };
    
        const img = new Image();
        img.crossOrigin = "anonymous";
        img.onload = () => {
          const canvas = document.createElement("canvas");
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          canvas.getContext("2d").drawImage(img, 0, 0);
          _renderHeader(canvas.toDataURL("image/png"));
        };
        img.onerror = () => _renderHeader(null);
        img.src = CAMINHO_LOGO;
      }
    
      // FILTRO INLINE NA TABELA
      function _aplicarFiltroTabela() {
        const termo = $id("inputFiltroTabela").value.toLowerCase().trim();
        const btnX = $id("btnLimparFiltroTabela");
        const linhas = document.querySelectorAll("#gradeRelatorio tbody tr");
    
        btnX.classList.toggle("d-none", termo === "");
    
        linhas.forEach((tr) => {
          tr.querySelectorAll("td .highlight-match").forEach((el) => {
            el.outerHTML = el.textContent;
          });
    
          if (termo === "") {
            tr.classList.remove("linha-oculta");
            return;
          }
    
          const texto = tr.textContent.toLowerCase();
          if (!texto.includes(termo)) {
            tr.classList.add("linha-oculta");
            return;
          }
    
          tr.classList.remove("linha-oculta");
    
          tr.querySelectorAll("td").forEach((td) => {
            const txt = td.textContent;
            if (!txt.toLowerCase().includes(termo)) return;
            const re = new RegExp(
              "(" + termo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")",
              "gi",
            );
            td.innerHTML = txt.replace(
              re,
              '<span class="highlight-match">$1</span>',
            );
          });
        });
      }
    
      // FORMULÁRIO
      function _sincronizaTipoProduto() {
        const grupo = $id("cboGrupoProduto");
        const tipo = $id("cboTipoProduto");
        if (!grupo || !tipo) return;
        if (grupo.value === "PAPELÃO" || grupo.value === "") {
          tipo.disabled = false;
        } else {
          tipo.disabled = true;
          tipo.value = "";
        }
      }
    
      function _validar() {
        const frm = document.PedidosEmCarteira;
        if (frm.cboCidade.value === "") {
          alert("Selecione a filial!");
          return false;
        }
        if (frm.dataini.value === "") {
          alert("Informe a data inicial!");
          return false;
        }
        if (frm.datafin.value === "") {
          alert("Informe a data final!");
          return false;
        }
        return true;
      }
    
      function _coletarFiltros() {
        const frm = document.PedidosEmCarteira;
        return {
          acao: "obterPedidos",
          cboCidade: frm.cboCidade.value,
          dataini: frm.dataini.value,
          datafin: frm.datafin.value,
          codcli: frm.codcli.value.trim(),
          lojacli: frm.lojacli.value.trim(),
          cboGrupoProduto: frm.cboGrupoProduto.value,
          cboTipoProduto: frm.cboTipoProduto.value,
          cboTipoFrete: frm.cboTipoFrete.value,
          cboPeriodoPor: frm.cboPeriodoPor.value,
          representante: frm.representante.value.trim(),
          pedido: frm.pedido.value.trim().toUpperCase(),
        };
      }
    
      function _filtrar() {
        if (!_validar()) return;
    
        exibirOverlay();
        $id("inputFiltroTabela").value = "";
    
        _xhrPost(
          _coletarFiltros(),
          function (resp) {
            ocultarOverlay();
    
            if (!resp || resp.sucesso !== true) {
              alert((resp && resp.mensagem) || "Erro ao carregar dados.");
              _dadosCompletos = [];
              _totaisCompletos = null;
              _renderizarCorpo([]);
              _renderizarRodape(null);
              return;
            }
    
            _dadosCompletos = resp.dados || [];
            _totaisCompletos = resp.totais || null;
    
            _renderizarCorpo(_dadosCompletos);
            _renderizarRodape(_dadosCompletos.length ? _totaisCompletos : null);
          },
          function (msg) {
            ocultarOverlay();
            alert(msg);
          },
        );
      }
    
      function _limpar() {
        const frm = document.PedidosEmCarteira;
        frm.cboCidade.value = "";
        frm.cboGrupoProduto.value = "";
        frm.cboTipoProduto.value = "";
        frm.cboTipoFrete.value = "";
        frm.cboPeriodoPor.value = "DTEMI";
        frm.pedido.value = "";
        frm.codcli.value = "";
        frm.lojacli.value = "";
        frm.razao.value = "";
        frm.representante.value = "";
        frm.dataini.value = "";
        frm.datafin.value = "";
    
        _sincronizaTipoProduto();
    
        _dadosCompletos = [];
        _totaisCompletos = null;
        $id("inputFiltroTabela").value = "";
        $id("btnLimparFiltroTabela").classList.add("d-none");
    
        $id("corpoTabela").innerHTML = _linhaVazia(
          "Utilize os filtros acima para pesquisar",
        );
        $id("rodapeTabela").innerHTML = "";
      }
    
    
      function _consultaCliente() {
        const mapa = {
          PIRAPETINGA: "00",
          UBERABA: "04",
          BENTO: "05",
          MARANGUAPE: "06",
          SAQUAREMA: "07",
        };
        const unid = mapa[document.PedidosEmCarteira.cboCidade.value] || "";
        window.open(
          "../../buscaBICLIENTES.asp?id1=codcli&id2=lojacli&id3=razao&id6=" + unid,
          "Consulta Cliente",
          "left,width=650,height=" + window.screen.height,
        );
      }
    
      // INICIALIZAÇÃO
      function init() {
        ocultarOverlay();
        _sincronizaTipoProduto();
    
        $id("btnPesquisar").addEventListener("click", _filtrar);
        $id("btnLimpar").addEventListener("click", _limpar);
        $id("btnExcel").addEventListener("click", exportarExcel);
        $id("btnPdf").addEventListener("click", exportarPdf);
        $id("btnBuscarCliente").addEventListener("click", _consultaCliente);
        $id("btnLimparCliente").addEventListener("click", function () {
          const frm = document.PedidosEmCarteira;
          frm.codcli.value  = "";
          frm.lojacli.value = "";
          frm.razao.value   = "";
        });
        $id("corpoTabela").addEventListener("click", function (event) {
          const link = event.target.closest(".pedido-detalhe-link");
          if (!link) return;
          event.preventDefault();
          abrirModalDetalhesPedido(link.dataset.pedido);
        });
        $id("cboGrupoProduto").addEventListener("change", _sincronizaTipoProduto);
        $id("pedido").addEventListener("input", function () {
          this.value = this.value.toUpperCase();
        });
    
        $id("inputFiltroTabela").addEventListener("input", _aplicarFiltroTabela);
        $id("btnLimparFiltroTabela").addEventListener("click", () => {
          $id("inputFiltroTabela").value = "";
          _aplicarFiltroTabela();
          $id("inputFiltroTabela").focus();
        });
    
        document.addEventListener("keydown", (e) => {
          if (e.ctrlKey && e.key === "k") {
            e.preventDefault();
            const el = $id("inputFiltroTabela");
            el.focus();
            el.select();
          }
        });

        // Pré-preenche os campos a partir de parâmetros recebidos via URL
        // (ex.: chamada vinda do RelatorioDeClientes)
        _preencherViaQueryString();
      }

      // Lê a querystring e popula o formulário; se &filtra=1, dispara a pesquisa
      function _preencherViaQueryString() {
        const qs  = new URLSearchParams(window.location.search);
        const frm = document.PedidosEmCarteira;

        const get = (nome) => (qs.get(nome) || "").trim();

        const cidade     = get("cboCidade");
        const codcli     = get("codcli");
        const lojacli    = get("lojacli");
        const razao      = get("razao");
        const dataini    = get("dataini");
        const datafin    = get("datafin");
        const periodoPor = get("cboPeriodoPor");
        const filtra     = get("filtra");

        if (cidade)     frm.cboCidade.value     = cidade;
        if (codcli)     frm.codcli.value        = codcli;
        if (lojacli)    frm.lojacli.value       = lojacli;
        if (razao)      frm.razao.value         = razao;
        if (dataini)    frm.dataini.value       = dataini;
        if (datafin)    frm.datafin.value       = datafin;
        if (periodoPor) frm.cboPeriodoPor.value = periodoPor;

        _sincronizaTipoProduto();

        // Dispara a pesquisa automaticamente quando chamado com &filtra=1
        if (filtra === "1" && cidade && dataini && datafin) {
          _filtrar();
        }
      }
    
      // Callback usado pela busca de cliente em popup (buscaBICLIENTES.asp)
      window.objReturn = function (obj) {
        const frm = document.PedidosEmCarteira;
        if (obj.erro === "ErroLogof") {
          alert("Favor se conectar novamente ao SWNET.\nSeu logon EXPIROU.");
          frm.codcli.value = "";
          frm.lojacli.value = "";
          frm.razao.value = "";
          return false;
        }
        if (obj.erro === "ErroNEncontrado") {
          frm.codcli.value = "";
          frm.lojacli.value = "";
          frm.razao.value = "";
          alert(
            "CLIENTE NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente Bloqueado.\nCliente não pertence a sua carteira.\nCliente não Cadastrado.",
          );
          return false;
        }
        if (obj.erro === "ErroCliente") {
          alert("Favor preencher os campos Cliente e Loja.");
          frm.codcli.value = "";
          frm.lojacli.value = "";
          frm.razao.value = "";
          return false;
        }
        if (obj.tipo === "Grupo") {
          frm.razao.value = obj.val1 || "";
        }
      };
    
      window.RelatorioPedidosEmCarteira = Object.freeze({ init });
    })();

  </script>

</body>
</html>
