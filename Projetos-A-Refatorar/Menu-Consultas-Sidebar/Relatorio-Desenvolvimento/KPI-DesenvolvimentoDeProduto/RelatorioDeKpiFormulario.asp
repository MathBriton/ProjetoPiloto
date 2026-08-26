<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2.1/cnn2.asp"-->
<!--#Include File="../../cnn2.1/conexoes.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->

<%
Response.Charset  = "UTF-8"
Response.CodePage = 65001

session.LCID = 1046
server.ScriptTimeout = 900

LCID = 1049

Call VerificaSessaoExpirada()
%>
<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - KPI Desvolvimento de Produto</title>
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

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <style>

      body {
        background-color: #f0f0f0;
        min-height: 100vh;
        padding-bottom: 2rem;
      }

      .page-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #212529;
      }

      /* Card de Filtros ------------------------------------------------ */
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

      .filter-card-body .form-label {
        font-size: .82rem;
        font-weight: 600;
        color: #6c757d;
        margin-bottom: .25rem;
      }

      .input-data { width: 155px; }

      /* Dropdown de seleção múltipla (checkbox) ------------------------ */
      .dd-toggle        { width: 100%; text-align: left; }
      .dd-toggle::after { display: none; }
      .dd-menu          { min-width: 100%; max-height: 260px; overflow-y: auto; }

      /* Card de Tabela --------------------------------------------------*/
      .table-card {
        background-color: #fff;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        position: relative; /* necessário para o overlay de carregamento */
      }

      .table-toolbar {
        background-color: #fff;
        border-radius: 6px 6px 0 0;
      }

      .search-box { position: relative; max-width: 220px; }
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

      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }
      .table-wrapper table { min-width: 3500px; }

      .table thead th {
        position: sticky;
        background-color: #fff;
        color: #212529;
        font-size: .78rem;
        font-weight: 700;
        text-align: center;
        white-space: nowrap;
        z-index: 2;
        border-color: #dee2e6;
        padding: .5rem .6rem;
      }
      .table thead tr:first-child th { top: 0; }
      .table thead tr:nth-child(2) th { top: 27px; }

      .table tbody td {
        font-size: .8rem;
        text-align: center;
        white-space: nowrap;
        color: #212529;
        vertical-align: middle;
        padding: .4rem .6rem;
      }

      /* Overlay de carregamento (escopo da tabela) ---------------------*/
      .overlay-loading {
        position: absolute;
        inset: 0;
        background: rgba(255, 255, 255, .80);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 10;
        gap: 1rem;
        border-radius: 6px;
      }
      .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
      .overlay-loading span            { font-size: .9rem; font-weight: 600; color: #495057; }

      /* Rodapé da tabela ------------------------------------------------*/
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

      /* Mensagem de "nenhum resultado" -----------------------------------*/
      .no-results-message {
        padding: 2.5rem 1rem;
        text-align: center;
        color: #6c757d;
      }
      .no-results-message i {
        font-size: 2rem;
        color: #adb5bd;
        display: block;
        margin-bottom: .5rem;
      }
      .no-results-message span {
        font-size: .95rem;
        font-weight: 700;
        letter-spacing: .03em;
      }

      /* Toast (notificações) -------------------------------------------- */
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

  <body onload="doOnLoad()">

    <div id="toastContainer" class="toast-container-custom"></div>

    <div class="container-fluid py-3 px-3 px-lg-4">
      <div class="row align-items-center">
        <div class="col text-start">
          <span class="page-title">Relatório de KPI Desenvolvimento de Produto</span>
        </div>
      </div>
    </div>

    <div id="filtro" name="filtro" class="container-fluid px-3 px-lg-4">
      <div class="filter-card">
        <div class="filter-card-header">
          <i class="bi bi-funnel me-2"></i>Filtros
        </div>
        <div class="filter-card-body">

          <form id="form" name="form" method="post">

            <div class="row g-2">

              <div class="col-md-2">
                  <label for="filial" class="form-label">Filial</label>
                  <select name="filial" id="filial" class="form-select form-select-sm">
                    <option value="" <% if request("filial") = "" then %> selected <% end if %> >Selecione...</option>
                    <option value="00" <% if request("filial") = "00" then %> selected <% end if %> >Pirapetinga</option>
                    <option value="04" <% if request("filial") = "04" then %> selected <% end if %> >Uberaba</option>
                    <option value="05" <% if request("filial") = "05" then %> selected <% end if %> >Bento Gonçalves</option>
                    <option value="06" <% if request("filial") = "06" then %> selected <% end if %> >Maranguape</option>
                    <option value="07" <% if request("filial") = "07" then %> selected <% end if %> >Saquarema</option>
                  </select>
              </div>

              <div class="col-auto">
                <label for="dtini" class="form-label">De</label>
                <input type="date" name="dtini" id="dtini" class="form-control form-control-sm input-data" placeholder="Data Inicial de Solicitação" value="<%=request("dtini") %>" />
              </div>

              <div class="col-auto">
                <label for="dtfim" class="form-label">Até</label>
                <input type="date" name="dtfim" id="dtfim" class="form-control form-control-sm input-data" placeholder="Data Final de Solicitação" value="<%=request("dtfim") %>" />
              </div>

              <div class="col-1">
                <label for="id" class="form-label">ID</label>
                <input type="text" name="id" class="form-control form-control-sm" value="<%=request("id") %>"/>
              </div>

              <div class="col">
                <label for="cliente" class="form-label">Cliente</label>
                <div class="input-group input-group-sm">
                  <input type="text" name="cliente" id="cliente" class="form-control" placeholder="Pesquisar..." readonly value="<%= request("cliente") %>" />
                  <button type="button" class="btn btn-secondary" onclick="PesquisaCliente()" title="Localizar clientes"><i class="bi bi-box-arrow-up-right"></i> Buscar</button>
                  <button type="button" class="btn btn-outline-secondary" onclick="document.getElementById('cliente').value=''" title="Limpar clientes"><i class="bi bi-x-lg"></i></button>
                </div>
              </div>

              <div class="col">
                <label for="tipo" class="form-label">Tipo</label>
                <select name="tipo" id="tipo" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="1" <% if InStr(request("tipo"),"1") > 0 then %> selected <% end if %> >Projeto a partir de amostra</option>
                  <option value="2" <% if InStr(request("tipo"),"2") > 0 then %> selected <% end if %> >Projeto a partir de especificação</option>
                  <option value="3" <% if InStr(request("tipo"),"3") > 0 then %> selected <% end if %> >Altera / Cancela</option>
                  <option value="4" <% if InStr(request("tipo"),"4") > 0 then %> selected <% end if %> >Inovação</option>
                  <option value="6" <% if InStr(request("tipo"),"6") > 0 then %> selected <% end if %> >Novo Projeto / Primário</option>
                </select>
              </div>

            </div>

            <div class="row g-2 mt-2">

              <div class="col">
                <label for="FT1_QTDAMO" class="form-label">Solicitação de amostras</label>
                <select name="FT1_QTDAMO" id="FT1_QTDAMO" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value= "0" <% if InStr(request("FT1_QTDAMO"),"0,") > 0 or request("FT1_QTDAMO") = "0" then %> selected <%end if %>>Sem amostras</option>
                  <option value="10" <% if InStr(request("FT1_QTDAMO"),"10") > 0 then %> selected <% end if %>>1 a 10</option>
                  <option value="11" <% if InStr(request("FT1_QTDAMO"),"11") > 0 then %> selected <% end if %>>Acima de 10</option>
                </select>
              </div>

              <div class="col">
                <label for="status" class="form-label">Status</label>
                <select name="status" id="status" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="P" <% if InStr(request("status"),"P") > 0 then %> selected <%end if%>>Aprovado</option>
                  <option value="A" <% if InStr(request("status"),"A") > 0 then %> selected <%end if%>>Pendente</option>
                  <option value="X" <% if InStr(request("status"),"X") > 0 then %> selected <%end if%>>Reprovado</option>
                  <option value="E" <% if InStr(request("status"),"E") > 0 then %> selected <%end if%>>Expirado</option>
                  <option value="K" <% if InStr(request("status"),"K") > 0 then %> selected <%end if%>>Cancelado</option>
                </select>
              </div>

              <div class="col">
                <label for="Saving" class="form-label">Saving</label>
                <select name="Saving" class="form-control form-control-sm">
                  <option value="">Selecione...</option>
                  <option value="K" <%if request("Saving") = "K" then%>selected<%end if%>>Com Saving SW</option>
                  <option value="C" <%if request("Saving") = "C" then%>selected<%end if%>>Com Saving Cliente</option>
                  <option value="A" <%if request("Saving") = "A" then%>selected<%end if%>>Ambos</option>
                  <option value="N" <%if request("Saving") = "N" then%>selected<%end if%>>Sem Saving</option>
                </select>
              </div>

              <div class="col">
                <label for="estilo" class="form-label">Estilo</label>
                <select name="estilo" id="estilo" multiple="multiple" class="selectpicker form-control form-control-sm">
                </select>
              </div>

              <div class="col">
                <label for="gerouft" class="form-label">Quem gerou FT</label>
                <select name="gerouft" id="gerouft" multiple="multiple" class="selectpicker form-control form-control-sm">
                </select>
              </div>

              <div class="col">
                <label for="envamostra" class="form-label">Quem enviou amostra</label>
                <select name="envamostra" id="envamostra" multiple="multiple" class="selectpicker form-control form-control-sm">
                </select>
              </div>

            </div>

            <div class="row g-2 mt-2">

              <div class="col">
                <label for="FT1_TipoAlteraCancela" class="form-label">Processo</label>
                <select name="FT1_TipoAlteraCancela" id="FT1_TipoAlteraCancela" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="1" <%if InStr(request("FT1_TipoAlteraCancela"),1) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Inovação</option>
                  <option value="4" <%if InStr(request("FT1_TipoAlteraCancela"),4) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Melhoria Continua - SW</option>
                  <option value="5" <%if InStr(request("FT1_TipoAlteraCancela"),5) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Melhoria Continua - Cliente</option>
                  <option value="6" <%if InStr(request("FT1_TipoAlteraCancela"),6) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Alteração Layout</option>
                  <option value="7" <%if InStr(request("FT1_TipoAlteraCancela"),7) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Alteração Dimensional</option>
                  <option value="8" <%if InStr(request("FT1_TipoAlteraCancela"),8) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Alteração Estrutura</option>
                  <option value="9" <%if InStr(request("FT1_TipoAlteraCancela"),9) > 0 And InStr(request("FT1_TipoAlteraCancela"),0) = 0 then %> selected <%end if %>>Nova Referência conforme especificado</option>
                </select>
              </div>

              <div class="col">
                <label for="representante" class="form-label">Representante</label>
                <select name="representante" id="representante" multiple="multiple" class="selectpicker form-control form-control-sm">
                </select>
              </div>

              <div class="col">
                <label for="solicitante" class="form-label">Solicitante</label>
                <select name="solicitante" id="solicitante" multiple="multiple" class="selectpicker form-control form-control-sm">
                </select>
              </div>

              <div class="col">
                <label for="FT1_STATUS2" class="form-label">Andamento do Processo</label>
                <select name="FT1_STATUS2" id="FT1_STATUS2" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="E" <% if InStr(request("FT1_STATUS2"),"E") > 0 then %> selected <% end if %>>Amostra Enviada</option>
                  <option value="A" <% if InStr(request("FT1_STATUS2"),"A") > 0 then %> selected <% end if %>>Amostra Aprovada</option>
                  <option value="R" <% if InStr(request("FT1_STATUS2"),"R") > 0 then %> selected <% end if %>>Amostra Reprovada</option>
                  <option value="P" <% if InStr(request("FT1_STATUS2"),"P") > 0 then %> selected <% end if %>>Análise do Projeto</option>
                  <option value="G" <% if InStr(request("FT1_STATUS2"),"G") > 0 then %> selected <% end if %>>Envio de Especificação/Amostra</option>
                  <option value="B" <% if InStr(request("FT1_STATUS2"),"B") > 0 then %> selected <% end if %>>Liberação PPCP/Produção/Custo</option>
                  <option value="L" <% if InStr(request("FT1_STATUS2"),"L") > 0 then %> selected <% end if %>>Aprovação de Layout</option>
                  <option value="C" <% if InStr(request("FT1_STATUS2"),"C") > 0 then %> selected <% end if %>>Liberação Comercial - Gerência</option>
                  <option value="D" <% if InStr(request("FT1_STATUS2"),"D") > 0 then %> selected <% end if %>>Liberação Comercial - Diretoria</option>
                  <option value="F" <% if InStr(request("FT1_STATUS2"),"F") > 0 then %> selected <% end if %>>Cadastro de FT</option>
                  <option value="H" <% if InStr(request("FT1_STATUS2"),"H") > 0 then %> selected <% end if %>>Padrão de Cor</option>
                  <option value="J" <% if InStr(request("FT1_STATUS2"),"J") > 0 then %> selected <% end if %>>Confecção do Ferramental</option>
                  <option value="K" <% if InStr(request("FT1_STATUS2"),"K") > 0 then %> selected <% end if %>>FT Liberada</option>
                  <option value="T" <% if InStr(request("FT1_STATUS2"),"T") > 0 then %> selected <% end if %>>Sem Retorno</option>
                </select>
              </div>

              <div class="col">
                <label for="FT1_MOTCANC" class="form-label">Cancelado</label>
                <select name="FT1_MOTCANC" id="FT1_MOTCANC" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="E" <%if InStr(request("FT1_MOTCANC"),"E") > 0 then %> selected <%end if %>>Emissão em Duplicidade</option>
                  <option value="P" <%if InStr(request("FT1_MOTCANC"),"P") > 0 then %> selected <%end if %>>Preço</option>
                  <option value="D" <%if InStr(request("FT1_MOTCANC"),"D") > 0 then %> selected <%end if %>>Descontinuidade do Projeto</option>
                  <option value="R" <%if InStr(request("FT1_MOTCANC"),"R") > 0 then %> selected <%end if %>>Sem Retorno Comercial</option>
                  <option value="C" <%if InStr(request("FT1_MOTCANC"),"C") > 0 then %> selected <%end if %>>Sem Retorno Cliente</option>
                  <option value="T" <%if InStr(request("FT1_MOTCANC"),"T") > 0 then %> selected <%end if %>>Proposta Descartada</option>
                </select>
              </div>

              <div class="col">
                <label for="FT1_MOTIVOREPROVACAOAMOSTRA2" class="form-label">Amostra reprovada</label>
                <select name="FT1_MOTIVOREPROVACAOAMOSTRA2" id="FT1_MOTIVOREPROVACAOAMOSTRA2" multiple="multiple" class="selectpicker form-control form-control-sm">
                  <option value="P" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"P") > 0 then%> selected <%end if %>>Preço</option>
                  <option value="D" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"D") > 0 then%> selected <%end if %>>Dimensional em Desacordo</option>
                  <option value="E" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"E") > 0 then%> selected <%end if %>>Especificação Cliente Desatualizada</option>
                  <option value="S" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"S") > 0 then%> selected <%end if %>>Especificação de Primário Errado</option>
                  <option value="U" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"U") > 0 then%> selected <%end if %>>Descontinuidade do Projeto</option>
                  <option value="K" <%if InStr(request("FT1_MOTIVOREPROVACAOAMOSTRA2"),"K") > 0 then%> selected <%end if %>>Sem Retorno Comercial</option>
                </select>
              </div>

              <div class="col-2">
                <div class="d-flex justify-content-end gap-1 mt-4">
                  <button type="button" class="btn btn-light btn-sm" title="Limpar o filtro" onclick="LimparFiltros()"><i class="bi bi-eraser"></i></button>
                  <button type="button" class="btn btn-primary btn-sm" title="Filtrar" onclick="fsubmit()"><i class="bi bi-funnel-fill me-2"></i>Filtrar</button>
                </div>
              </div>

            </div>

            <div class="row mt-3">

              <div class="col">
                <span id="msgAtualizacao"></span>
              </div>

            </div>

          </form>

        </div>
      </div>
    </div>

    <!-- Modal Localizar Clientes (substitui o antigo popup MPA clientes_busca.asp) -->
    <div class="modal fade" id="modalBuscaCliente" tabindex="-1" aria-labelledby="modalBuscaClienteLabel" aria-hidden="true">
      <div class="modal-dialog modal-xl modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="modalBuscaClienteLabel">Localizar Clientes</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">

            <div class="input-group input-group-sm mb-3">
              <input type="text" id="termoBuscaCliente" class="form-control" placeholder="Digite parte da razão social ou código..." />
              <button type="button" class="btn btn-primary" id="btnBuscarClienteModal"><i class="bi bi-search me-1"></i>Buscar</button>
            </div>

            <div class="mb-2">
              <span class="text-muted" style="font-size:.82rem;">Enviados: </span>
              <span id="clientesEnviadosLabel" style="font-size:.82rem;"></span>
            </div>

            <div class="table-responsive" style="max-height: 50vh;">
              <table class="table table-bordered table-hover table-sm mb-0" style="font-size:.82rem;">
                <thead>
                  <tr>
                    <th style="width:34px;"></th>
                    <th>Código</th>
                    <th>Loja</th>
                    <th>Razão</th>
                    <th>Fantasia</th>
                    <th>Cidade</th>
                    <th>UF</th>
                    <th>Vendedor</th>
                  </tr>
                </thead>
                <tbody id="tbodyBuscaCliente">
                  <tr>
                    <td colspan="8" class="text-center text-muted py-3">Digite um termo e clique em <strong>Buscar</strong>.</td>
                  </tr>
                </tbody>
              </table>
            </div>

          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Fechar</button>
            <button type="button" class="btn btn-primary btn-sm" id="btnEnviarClientesModal"></i>Importar</button>
          </div>
        </div>
      </div>
    </div>

    <div class="table-card mt-3 mb-4 mx-3 mx-lg-4">

      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          <button id="btnExportCsv" type="button" class="btn btn-success btn-sm" onclick="GerarExcel()">
            <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
          </button>
        </div>

        <div class="search-box">
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" />
          <span class="search-hotkey"><kbd>CTRL</kbd><kbd>K</kbd></span>
        </div>
      </div>

      <div id="tableLoader" class="overlay-loading d-none">
        <div class="spinner-border" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <div class="table-wrapper" id="lista" name="lista">
        <table class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead>
          <tr>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center">Andamento</th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center" colspan="7">Amostra</th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center"></th>
            <th class="text-center">Quantidade</th>
            <th class="text-center">Faturamento</th>
            <th class="text-center" colspan="2">Ped. Fat.</th>
            <th class="text-center" colspan="2">FT</th>
            <th class="text-center" colspan="3">Saving</th>
            <th class="text-center">Motivo</th>
            <th class="text-center">Cancelamento</th>
          </tr>
          <tr>
            <th class="text-center">ID</th>
            <th class="text-center">Status</th>
            <th class="text-center">Processo</th>
            <th class="text-center">Processo</th>
            <th class="text-center">Tipo</th>
            <th class="text-center">Data</th>
            <th class="text-center">Solicitante</th>
            <th class="text-center">Referência</th>
            <th class="text-center">Cliente</th>
            <th class="text-center">Representante</th>
            <th class="text-center">FEFCO</th>
            <th class="text-center">Segmento</th>
            <th class="text-center">Unidade</th>
            <th class="text-center">Quantidade</th>
            <th class="text-center">Status</th>
            <th class="text-center">Data envio</th>
            <th class="text-center">Enviou</th>
            <th class="text-center">Aprovou</th>
            <th class="text-center">Data aprov.</th>
            <th class="text-center">Hora aprov.</th>
            <th class="text-center">Município</th>
            <th class="text-center">UF</th>
            <th class="text-center">Leadtime</th>
            <th class="text-center">FT Nova</th>
            <th class="text-center">Pedidos</th>
            <th class="text-center">NET</th>
            <th class="text-center">(m²)</th>
            <th class="text-center">(kg)</th>
            <th class="text-center">Gerou</th>
            <th class="text-center">Data</th>
            <th class="text-center">Saving</th>
            <th class="text-center">SW valor</th>
            <th class="text-center">Cliente valor</th>
            <th class="text-center">Reprov.Amostra</th>
            <th class="text-center">do Processo</th>
          </tr>
        </thead>

        <tbody>
          <tr>
            <td colspan="36" class="text-center text-muted py-4">
              <i class="bi bi-funnel me-2"></i>Utilize os filtros acima e clique em <strong>Filtrar</strong> para visualizar os dados.
            </td>
          </tr>
        </tbody>
        </table>
      </div>

      <div class="table-footer">

        <div class="table-footer-left">
          <label for="selectLimite" class="form-label mb-0 text-muted" style="font-size:.8rem;white-space:nowrap;">
            Linhas por página:
          </label>
          <select id="selectLimite" class="form-select form-select-sm select-limite">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="100">100</option>
            <option value="250">250</option>
            <option value="500" selected>500</option>
            <option value="999999">Todos</option>
          </select>
          <span id="totalRegistrosFooter" class="info-registros"></span>
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

<%
Function VerificaSessaoExpirada()

  if trim(Session("CODUSER")) = "" then
    Call SesseaoExpirada()
  end if

End function
%>


    <script>

      const REPOSITORIO = 'RelatorioDeKpiRepositorio.asp';

      // Valores pré-selecionados vindos da querystring (suporte a links com filtros prontos),
      // usados para marcar as opções como "selected" após o carregamento via AJAX.
      const FILTROS_SELECIONADOS = {
        estilo:        "<%= EscapeJSON(request("estilo")) %>",
        gerouft:       "<%= EscapeJSON(request("gerouft")) %>",
        envamostra:    "<%= EscapeJSON(request("envamostra")) %>",
        representante: "<%= EscapeJSON(request("representante")) %>",
        solicitante:   "<%= EscapeJSON(request("solicitante")) %>"
      };

      // ------------------------------------------------------------------------------------------
      function ShowLoader() {
        const displayLoader = document.getElementById("tableLoader");
        if (displayLoader) displayLoader.classList.remove("d-none");
      }
      // ------------------------------------------------------------------------------------------
      function HideLoader() {
        const displayLoader = document.getElementById("tableLoader");
        if (displayLoader) displayLoader.classList.add("d-none");
      }
      // ------------------------------------------------------------------------------------------
      const TOAST_DELAY = 4500;
      const TOAST_CONFIG = {
        sucesso: { icone: 'bi-check-circle-fill', cor: '#198754', barra: '#198754' },
        erro:    { icone: 'bi-x-circle-fill',     cor: '#dc3545', barra: '#dc3545' },
        aviso:   { icone: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
        info:    { icone: 'bi-info-circle-fill',  cor: '#6c757d', barra: '#6c757d' }
      };
      function exibirToast(tipo, mensagem) {
        const cfg = TOAST_CONFIG[tipo] || TOAST_CONFIG.info;
        const $toast = $(
          '<div class="toast-custom">' +
            '<div class="d-flex align-items-start p-2">' +
              '<i class="bi ' + cfg.icone + ' me-2 mt-1" style="color:' + cfg.cor + ';"></i>' +
              '<div class="flex-grow-1" style="font-size:.85rem;color:#212529;">' + mensagem + '</div>' +
              '<button type="button" class="btn-close ms-2" aria-label="Fechar"></button>' +
            '</div>' +
            '<div class="toast-progress-bar" style="height:3px;background:' + cfg.barra + ';"></div>' +
          '</div>'
        );
        $toast.find('.btn-close').on('click', function () { $toast.remove(); });
        $('#toastContainer').append($toast);
        setTimeout(function () { $toast.remove(); }, TOAST_DELAY);
      }
      // ------------------------------------------------------------------------------------------
      function doOnLoad() {
        HideLoader();
        carregarFiltros(true);

        $('#filial').on('change', function () {
          // Apenas atualiza as opções dos filtros para a nova planta; não
          // dispara a consulta — isso só deve acontecer ao clicar "Filtrar".
          dadosCompletos = [];
          $('#msgAtualizacao').text('');
          renderizarTabela();
          carregarFiltros(false);
        });
      }
      // ------------------------------------------------------------------------------------------
      function fsubmit() {
        if (!validaFilialObrigatoria()) return;
        carregarDados();
      }
      // ------------------------------------------------------------------------------------------
      function LimparFiltros() {
        window.location.href = window.location.pathname;
      }
      // ------------------------------------------------------------------------------------------
      function formatarDataCSV(valor) {
        if (!valor) return '';
        const m = String(valor).match(/^(\d{4})-(\d{2})-(\d{2})/);
        return m ? (m[3] + '/' + m[2] + '/' + m[1]) : String(valor);
      }
      function formatarNumCSV(valor, casas) {
        casas = casas == null ? 2 : casas;
        if (valor == null || valor === '') return '';
        const n = parseFloat(valor);
        if (isNaN(n)) return '';
        return n.toFixed(casas).replace('.', ',');
      }
      function exportarCSV(nomeArquivo, cabecalhos, linhas) {
        const esc = function (v) {
          const s = v == null ? '' : String(v);
          return (s.includes(';') || s.includes('"') || s.includes('\n'))
            ? '"' + s.replace(/"/g, '""') + '"' : s;
        };
        const rows = [cabecalhos.map(esc).join(';')]
          .concat(linhas.map(function (row) { return row.map(esc).join(';'); }));
        const blob = new Blob(['\uFEFF' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = nomeArquivo.endsWith('.csv') ? nomeArquivo : nomeArquivo + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
      function GerarExcel() {

        if (!dadosCompletos || dadosCompletos.length === 0) {
          exibirToast('aviso', 'Não há dados para exportar. Filtre os dados primeiro.');
          return;
        }

        const cabecalhos = [
          'ID', 'Status', 'Processo', 'Andamento', 'Tipo', 'Data', 'Solicitante', 'Referência',
          'Cliente', 'Representante', 'Estilo', 'Segmento', 'Unidade', 'Qtd. Amostra',
          'Status Amostra', 'Data Envio Amostra', 'Enviou', 'Aprovou', 'Data Aprov.', 'Hora Aprov.',
          'Município', 'UF', 'Leadtime', 'FT Nova', 'Pedidos', 'NET', 'Área (m²)', 'Peso (kg)',
          'Gerou FT', 'Data Geração FT', 'Saving', 'Saving SW', 'Saving Cliente',
          'Motivo Reprov. Amostra', 'Motivo Cancelamento'
        ];

        const linhas = dadosCompletos.map(function (item) {
          return [
            item.id ?? '',
            item.status ?? '',
            item.processo ?? '',
            item.andamento ?? '',
            item.tipo ?? '',
            formatarDataCSV(item.data),
            item.solicitante ?? '',
            item.referencia ?? '',
            item.cliente ?? '',
            item.representante ?? '',
            item.estilo ?? '',
            item.segmento ?? '',
            item.unidade ?? '',
            item.qtdAmostra ?? '',
            item.statusAmostra ?? '',
            formatarDataCSV(item.dataEnvioAmostra),
            item.quemEnviouAmostra ?? '',
            item.quemAprovouAmostra ?? '',
            formatarDataCSV(item.dataStatusAmostra),
            item.horaStatusAmostra ?? '',
            item.municipio ?? '',
            item.uf ?? '',
            item.leadtime ?? '',
            item.ftNova ?? '',
            item.pedidos ?? '',
            formatarNumCSV(item.valorNet),
            formatarNumCSV(item.area),
            formatarNumCSV(item.peso),
            item.quemGerouFT ?? '',
            formatarDataCSV(item.dataGeracaoFT),
            item.saving ?? '',
            item.savingSW ?? '',
            item.savingCliente ?? '',
            item.motivoReprovacao ?? '',
            item.motivoCancelamento ?? ''
          ];
        });

        exportarCSV('RelatorioDeKpi.csv', cabecalhos, linhas);
      }
      // ------------------------------------------------------------------------------------------
      // Localizar Clientes — agora como modal Bootstrap 5 nativo (SPA), substituindo o antigo
      // popup MPA "clientes_busca.asp" que abria em window.open() e usava window.opener.
      let clienteSelecionados = [];
      const modalBuscaCliente = new bootstrap.Modal(document.getElementById('modalBuscaCliente'));

      function PesquisaCliente() {

        if (!validaFilialObrigatoria()) return;

        clienteSelecionados = (document.getElementById('cliente').value || '')
          .split(',')
          .map(function (v) { return v.trim(); })
          .filter(function (v) { return v !== ''; });

        atualizarLabelEnviados();
        $('#termoBuscaCliente').val('');
        $('#tbodyBuscaCliente').html('<tr><td colspan="8" class="text-center text-muted py-3">Digite um termo e clique em <strong>Buscar</strong>.</td></tr>');
        modalBuscaCliente.show();
        setTimeout(function () { $('#termoBuscaCliente').trigger('focus'); }, 300);
      }
      // ------------------------------------------------------------------------------------------
      function atualizarLabelEnviados() {
        $('#clientesEnviadosLabel').text(clienteSelecionados.length ? clienteSelecionados.join(', ') : '');
      }
      // ------------------------------------------------------------------------------------------
      function montaLinhaBuscaCliente(item) {
        const chave = item.codigo + '-' + item.loja;
        const checked = clienteSelecionados.indexOf(chave) > -1 ? 'checked' : '';
        return '<tr>' +
          '<td class="text-center"><input type="checkbox" class="form-check-input chk-cliente-modal" data-chave="' + chave + '" ' + checked + ' title="Selecione a loja do cliente."></td>' +
          '<td>' + (item.codigo ?? '') + '</td>' +
          '<td>' + (item.loja ?? '') + '</td>' +
          '<td>' + (item.razao ?? '') + '</td>' +
          '<td>' + (item.fantasia ?? '') + '</td>' +
          '<td>' + (item.municipio ?? '') + '</td>' +
          '<td>' + (item.uf ?? '') + '</td>' +
          '<td>' + (item.vendedor ?? '') + '</td>' +
        '</tr>';
      }
      // ------------------------------------------------------------------------------------------
      function buscarClientesModal() {

        const filial = $('#filial').val();
        const termo = $('#termoBuscaCliente').val().trim();

        if (!filial) {
          exibirToast('aviso', 'Por favor, selecione a filial!');
          return;
        }

        if (!termo) {
          exibirToast('aviso', 'Digite um código ou parte da razão social para buscar.');
          return;
        }

        $('#tbodyBuscaCliente').html('<tr><td colspan="8" class="text-center text-muted py-3"><span class="spinner-border spinner-border-sm me-2"></span>Buscando...</td></tr>');

        $.ajax({
          url: REPOSITORIO,
          type: 'POST',
          data: { action: 'buscarCliente', filial: filial, cliente: termo },
          dataType: 'json'
        }).done(function (resp) {

          if (!resp || resp.success !== true) {
            $('#tbodyBuscaCliente').html('<tr><td colspan="8" class="text-center text-muted py-3">Nenhum resultado encontrado</td></tr>');
            if (resp && resp.message) exibirToast('aviso', resp.message);
            return;
          }

          const dados = resp.dados || [];

          if (dados.length === 0) {
            $('#tbodyBuscaCliente').html('<tr><td colspan="8" class="text-center text-muted py-3">Nenhum resultado encontrado</td></tr>');
            return;
          }

          $('#tbodyBuscaCliente').html(dados.map(montaLinhaBuscaCliente).join(''));

        }).fail(function () {
          $('#tbodyBuscaCliente').html('<tr><td colspan="8" class="text-center text-muted py-3">Nenhum resultado encontrado</td></tr>');
          exibirToast('erro', 'Erro ao buscar clientes. Tente novamente.');
        });

      }
      // ------------------------------------------------------------------------------------------
      function confirmarSelecaoClientes() {

        const $marcados = $('.chk-cliente-modal:checked');

        if ($marcados.length === 0) {
          exibirToast('aviso', 'Selecione ao menos um cliente.');
          return;
        }

        $marcados.each(function () {
          const chave = $(this).data('chave');
          if (clienteSelecionados.indexOf(chave) === -1) {
            clienteSelecionados.push(chave);
          }
        });

        $('#cliente').val(clienteSelecionados.join(','));
        atualizarLabelEnviados();
        modalBuscaCliente.hide();
      }
      // ------------------------------------------------------------------------------------------
      function validaFilialObrigatoria() {
        const vFilial = document.getElementById('filial');
        if (!vFilial || vFilial.value === '') {
          exibirToast('aviso', 'Por favor, selecione a Filial. Esse é o único campo obrigatório do filtro.');
          return false;
        }
        return true;
      }
      // ------------------------------------------------------------------------------------------
      // Converte um <select multiple class="selectpicker"> em dropdown
      // Bootstrap 5 com checkboxes, mantendo o <select> original como fonte
      // de dados (oculto) para o submit do formulário.
      function initMultiSelectDropdown(selectEl) {

        const $select = $(selectEl);
        const id = $select.attr('id');
        const placeholder = 'Selecione...';

        // Remove o wrapper criado em uma chamada anterior (ex.: ao trocar a
        // planta e recarregar os filtros), senão um novo dropdown duplicado
        // é inserido e o antigo, com opções desatualizadas, continua visível.
        $select.next('.dropdown').remove();

        $select.addClass('d-none');

        const $wrapper = $('<div class="dropdown"></div>');
        const $toggle = $(
          `<button id="btnFiltro${id}" type="button" class="form-select form-select-sm dd-toggle dropdown-toggle" data-bs-toggle="dropdown" data-bs-auto-close="outside" aria-expanded="false">${placeholder}</button>`
        );
        const $menu = $('<ul class="dropdown-menu dd-menu py-1"></ul>');

        $select.find('option').each(function () {
          const $opt = $(this);
          const val = $opt.val();
          const text = $opt.text();
          const checked = $opt.is(':selected');

          const $item = $(
            `<li><button class="dropdown-item" data-val="${val}" type="button">` +
            `<input type="checkbox" class="form-check-input me-1" ${checked ? 'checked' : ''} />` +
            `<span></span></button></li>`
          );
          $item.find('span').text(text);
          $menu.append($item);
        });

        $wrapper.append($toggle).append($menu);
        $select.after($wrapper);

        function updateLabel() {
          const $checked = $menu.find('input:checked');
          if ($checked.length === 0) {
            $toggle.text(placeholder);
          } else if ($checked.length === 1) {
            $toggle.text($checked.closest('button').find('span').text());
          } else {
            const first = $checked.first().closest('button').find('span').text();
            $toggle.text(`${first} +${$checked.length - 1}`);
          }
        }

        $menu.on('click', '.dropdown-item', function (e) {
          e.preventDefault();
          const $checkbox = $(this).find('input[type=checkbox]');
          const checked = !$checkbox.prop('checked');
          $checkbox.prop('checked', checked);

          const val = $(this).data('val');
          $select.find(`option[value="${val}"]`).prop('selected', checked);
          updateLabel();
        });

        updateLabel();
      }
      // ------------------------------------------------------------------------------------------
      // Estado dos dados carregados via AJAX e paginação client-side
      let paginaAtual = 1;
      let limitePagina = 500;
      let dadosCompletos = [];

      // IDs dos <select> múltiplos cujas opções vêm da ação "obterFiltros" (antes eram
      // renderizados inline no servidor com loops VBScript embutidos na página).
      const SELECTS_DINAMICOS = ['estilo', 'gerouft', 'envamostra', 'representante', 'solicitante'];

      function _popularSelectDinamico(id, opcoes) {

        const $select = $('#' + id);
        $select.empty();

        const valoresSelecionados = (FILTROS_SELECIONADOS[id] || '')
          .split(',')
          .map(function (v) { return v.trim(); })
          .filter(function (v) { return v !== ''; });

        (opcoes || []).forEach(function (opcao) {
          const selecionado = valoresSelecionados.indexOf(String(opcao.valor)) > -1;
          $select.append(
            $('<option></option>').val(opcao.valor).text(opcao.texto).prop('selected', selecionado)
          );
        });
      }

      function carregarFiltros(recarregarDados) {

        $.ajax({
          url: REPOSITORIO,
          type: 'POST',
          data: { action: 'obterFiltros', filial: $('#filial').val() },
          dataType: 'json'
        }).done(function (resp) {

          if (resp && resp.sucesso === true) {
            _popularSelectDinamico('estilo', resp.estilo);
            _popularSelectDinamico('gerouft', resp.gerouft);
            _popularSelectDinamico('envamostra', resp.envamostra);
            _popularSelectDinamico('representante', resp.representante);
            _popularSelectDinamico('solicitante', resp.solicitante);
          }

        }).always(function () {
          // Converte os selects dinâmicos em dropdown com checkbox somente após
          // as opções estarem disponíveis no DOM.
          SELECTS_DINAMICOS.forEach(function (id) {
            initMultiSelectDropdown(document.getElementById(id));
          });
          if (recarregarDados) {
            carregarDados();
          }
        });

      }

      function fmtNum(v) {
        const n = parseFloat(v || 0);
        return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }

      function montaLinha(item) {
        return '<tr class="data-row">' +
          '<td>' + (item.id ?? '') + '</td>' +
          '<td>' + (item.status ?? '') + '</td>' +
          '<td>' + (item.processo ?? '') + '</td>' +
          '<td>' + (item.andamento ?? '') + '</td>' +
          '<td>' + (item.tipo ?? '') + '</td>' +
          '<td>' + (item.data ?? '') + '</td>' +
          '<td>' + (item.solicitante ?? '') + '</td>' +
          '<td>' + (item.referencia ?? '') + '</td>' +
          '<td>' + (item.cliente ?? '') + '</td>' +
          '<td>' + (item.representante ?? '') + '</td>' +
          '<td>' + (item.estilo ?? '') + '</td>' +
          '<td>' + (item.segmento ?? '') + '</td>' +
          '<td>' + (item.unidade ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.qtdAmostra ?? '') + '</td>' +
          '<td>' + (item.statusAmostra ?? '') + '</td>' +
          '<td>' + (item.dataEnvioAmostra ?? '') + '</td>' +
          '<td>' + (item.quemEnviouAmostra ?? '') + '</td>' +
          '<td>' + (item.quemAprovouAmostra ?? '') + '</td>' +
          '<td>' + (item.dataStatusAmostra ?? '') + '</td>' +
          '<td>' + (item.horaStatusAmostra ?? '') + '</td>' +
          '<td>' + (item.municipio ?? '') + '</td>' +
          '<td>' + (item.uf ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.leadtime ?? '') + '</td>' +
          '<td>' + (item.ftNova ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.pedidos ?? '') + '</td>' +
          '<td style="text-align:right;">' + fmtNum(item.valorNet) + '</td>' +
          '<td style="text-align:right;">' + fmtNum(item.area) + '</td>' +
          '<td style="text-align:right;">' + fmtNum(item.peso) + '</td>' +
          '<td>' + (item.quemGerouFT ?? '') + '</td>' +
          '<td>' + (item.dataGeracaoFT ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.saving ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.savingSW ?? '') + '</td>' +
          '<td style="text-align:right;">' + (item.savingCliente ?? '') + '</td>' +
          '<td>' + (item.motivoReprovacao ?? '') + '</td>' +
          '<td>' + (item.motivoCancelamento ?? '') + '</td>' +
        '</tr>';
      }

      function carregarDados() {

        ShowLoader();

        const dados = $('#form').serialize() + '&envia=s';

        $.ajax({
          url: REPOSITORIO,
          type: 'POST',
          data: dados + '&action=consultar',
          dataType: 'json'
        }).done(function (resp) {

          if (!resp || resp.success !== true) {
            dadosCompletos = [];
            $('#msgAtualizacao').text('');
            renderizarTabela();
            HideLoader();
            return;
          }

          dadosCompletos = resp.dados || [];
          $('#msgAtualizacao').html(resp.ultimaAtualizacao || '');
          $('#totalRegistrosFooter').text('');

          paginaAtual = 1;
          renderizarTabela();

        }).fail(function () {
          dadosCompletos = [];
          renderizarTabela();
          exibirToast('erro', 'Erro ao carregar os dados. Tente novamente.');
        }).always(function () {
          HideLoader();
        });

      }

      function _itensFiltrados() {
        const termo = ($('#searchBox').val() || '').toLowerCase().trim();
        if (!termo) return dadosCompletos;
        return dadosCompletos.filter(function (item) {
          return JSON.stringify(item).toLowerCase().includes(termo);
        });
      }

      function renderizarTabela() {

        const $tbody = $('#lista tbody');
        const itens = _itensFiltrados();
        const total = itens.length;
        const totalPaginas = Math.max(1, Math.ceil(total / limitePagina));

        if (paginaAtual > totalPaginas) paginaAtual = totalPaginas;

        if (total === 0) {
          $tbody.html('<tr><td colspan="36"><div class="no-results-message"><i class="bi bi-clipboard2-data-fill"></i><span>Aguardando pela filtragem</span></div></td></tr>');
          $('#infoRegistros').text('Total: 0 registro(s)');
          _renderPaginacao(1, 1);
          return;
        }

        const inicio = (paginaAtual - 1) * limitePagina;
        const fim = inicio + limitePagina;
        const pagina = itens.slice(inicio, fim);

        $tbody.html(pagina.map(montaLinha).join(''));

        $('#infoRegistros').text('Total: ' + total + ' registro(s)');
        _renderPaginacao(paginaAtual, totalPaginas);
      }

      function _renderPaginacao(pagina, totalPaginas) {
        const $ul = $('#paginacao').empty();
        if (totalPaginas <= 1) return;

        const mkItem = (label, page, disabled, active) => {
          const $li = $('<li>').addClass('page-item');
          if (disabled) $li.addClass('disabled');
          if (active) $li.addClass('active');
          const $a = $('<a>').addClass('page-link').attr('href', '#').html(label);
          if (!disabled && !active) {
            $a.on('click', function (e) {
              e.preventDefault();
              paginaAtual = page;
              renderizarTabela();
            });
          }
          return $li.append($a);
        };

        $ul.append(mkItem('&laquo;', 1, pagina === 1, false));
        $ul.append(mkItem('&lsaquo;', pagina - 1, pagina === 1, false));

        const inicio = Math.max(1, pagina - 2);
        const fim = Math.min(totalPaginas, pagina + 2);
        for (let p = inicio; p <= fim; p++) {
          $ul.append(mkItem(p, p, false, p === pagina));
        }

        $ul.append(mkItem('&rsaquo;', pagina + 1, pagina === totalPaginas, false));
        $ul.append(mkItem('&raquo;', totalPaginas, pagina === totalPaginas, false));
      }

      function initTableSearch() {

        const $search = $('#searchBox');
        let timer = null;

        $search.on('input', function () {
          clearTimeout(timer);
          timer = setTimeout(function () {
            paginaAtual = 1;
            renderizarTabela();
          }, 250);
        });

        $(document).on('keydown', function (e) {
          if (e.ctrlKey && e.key === 'k') {
            e.preventDefault();
            $search.trigger('focus');
          }
        });
      }

      function initTablePagination() {

        $('#selectLimite').on('change', function () {
          limitePagina = parseInt($(this).val(), 10);
          paginaAtual = 1;
          renderizarTabela();
        });

      }
      // ------------------------------------------------------------------------------------------
      $(function () {
        $('.selectpicker').not('#' + SELECTS_DINAMICOS.join(', #')).each(function () {
          initMultiSelectDropdown(this);
        });
        initTableSearch();
        initTablePagination();

        $('#btnBuscarClienteModal').on('click', buscarClientesModal);
        $('#termoBuscaCliente').on('keydown', function (e) {
          if (e.key === 'Enter') {
            e.preventDefault();
            buscarClientesModal();
          }
        });
        $('#btnEnviarClientesModal').on('click', confirmarSelecaoClientes);
      });
    </script>

  </body>
</html>
