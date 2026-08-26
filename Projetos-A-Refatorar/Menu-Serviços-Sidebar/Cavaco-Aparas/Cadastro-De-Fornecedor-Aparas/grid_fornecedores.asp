<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer       = True
Response.Charset      = "UTF-8"
Response.CodePage     = 65001
Response.Expires      = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!-- #include file="../../cnn2/cnn2.asp" -->
<!-- #include file="../../funcoes/funcoes.asp" -->
<!-- #include File="../../includes/variaveis.asp"-->
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SWnet - Lista de Fornecedores</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
  <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
  <style>
    body {
      background: #f0f0f0;
      color: #212529;
      min-height: 100vh;
    }
    .page-title {
      color: #4a4a4a;
      font-size: 1.25rem;
    }
    .form-label {
      font-size: 0.82rem;
      font-weight: 600;
      margin-bottom: 0.25rem;
      color: #212529;
    }
    .form-control,
    .form-select {
      font-size: 0.85rem;
    }
    .form-control:focus,
    .form-select:focus {
      border-color: #4a4a4a;
      box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
    }
    .panel {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
    }
    .panel-header {
      background: #fff;
      color: #212529;
      border-bottom: 1px solid #dee2e6;
      padding: 0.6rem 1rem;
      border-radius: 6px 6px 0 0;
      font-size: 0.8rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.3px;
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }
    .panel-body {
      padding: 1rem;
    }
    .tbl-wrapper {
      overflow: auto;
      max-height: calc(100vh - 180px);
      border-radius: 4px;
    }
    .tbl-wrapper thead th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #f8f9fa;
      color: #212529;
      font-size: 0.71rem;
      font-weight: 600;
      white-space: nowrap;
      vertical-align: middle;
      text-align: center;
      padding: 0.45rem 0.75rem;
      border-color: #dee2e6;
      text-transform: uppercase;
      letter-spacing: 0.2px;
      cursor: default;
    }
    .tbl-wrapper thead th.sortable {
      cursor: pointer;
      user-select: none;
    }
    .tbl-wrapper thead th.sortable:hover {
      background: #eef0f2;
    }
    .tbl-wrapper tbody td {
      font-size: 0.78rem;
      vertical-align: middle;
      padding: 0.32rem 0.75rem;
      border-color: #e8e8e8;
      text-align: center;
    }
    .tbl-wrapper tbody tr:hover td {
      background-color: rgba(0, 0, 0, 0.03);
    }
    .tbl-vazia {
      padding: 3rem 1rem;
      text-align: center;
      color: #888;
      font-size: 0.88rem;
    }
    .btn-acao {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 26px;
      height: 22px;
      padding: 0;
      border-radius: 4px;
      font-size: 0.78rem;
      cursor: pointer;
      transition: opacity 0.15s;
    }
    .btn-acao:hover {
      opacity: 0.82;
    }
    .sort-icon {
      font-size: 0.65rem;
      margin-left: 3px;
      opacity: 0.55;
    }
    .toast-wrapper {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 11000;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
      min-width: 300px;
      max-width: 380px;
      pointer-events: none;
    }
    .toast-item {
      background: #fff;
      border-radius: 6px;
      overflow: hidden;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
      border: 1px solid #e8e8e8;
      pointer-events: all;
      animation: entrar-toast 0.2s ease;
    }
    .toast-item-success { border-left: 4px solid #198754; }
    .toast-item-danger  { border-left: 4px solid #dc3545; }
    .toast-item-info    { border-left: 4px solid #6c757d; }
    @keyframes entrar-toast {
      from { opacity: 0; transform: translateX(20px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    @keyframes encolher {
      from { width: 100%; }
      to   { width: 0%; }
    }
    .toast-corpo {
      display: flex;
      align-items: center;
      gap: 0.6rem;
      padding: 0.65rem 0.85rem;
      font-size: 0.85rem;
      color: #333;
    }
    .toast-fechar {
      margin-left: auto;
      background: none;
      border: none;
      font-size: 1rem;
      color: #999;
      cursor: pointer;
      padding: 0 0 0 0.5rem;
    }
    .toast-fechar:hover { color: #333; }
    .toast-barra { height: 3px; animation: encolher linear forwards; }
    .toast-barra-success { background: #198754; }
    .toast-barra-danger  { background: #dc3545; }
    .toast-barra-info    { background: #6c757d; }
    .overlay-loader {
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, 0.65);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .spinner-circular {
      width: 3rem;
      height: 3rem;
      border: 4px solid #e8e8e8;
      border-top-color: #4a4a4a;
      border-radius: 50%;
      animation: girar 0.8s linear infinite;
    }
    @keyframes girar { to { transform: rotate(360deg); } }
    .rodape-paginacao {
      display: grid;
      grid-template-columns: auto 1fr auto;
      align-items: center;
      padding: 0.5rem 0.75rem;
      background: #fff;
      border: 1px solid #dee2e6;
      border-top: none;
      border-radius: 0 0 6px 6px;
      font-size: 0.78rem;
      color: #212529;
      gap: 0.5rem;
    }
    .rodape-esquerda {
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }
    .rodape-centro {
      display: flex;
      justify-content: center;
    }
    .rodape-direita {
      display: flex;
      justify-content: flex-end;
      white-space: nowrap;
      color: rgb(33, 37, 41);
      font-weight: 500;
    }
    .paginacao .page-item .page-link {
      color: #6c757d;
      font-size: 0.78rem;
    }
    .paginacao .page-item.active .page-link {
      background: #4a4a4a;
      border-color: #4a4a4a;
      color: #fff;
    }
    .paginacao .page-item.disabled .page-link {
      color: #aaa;
      pointer-events: none;
    }
    .filtro-inline {
      position: relative;
      width: 340px;
    }
    .filtro-inline .icone-busca {
      position: absolute;
      left: 0.6rem;
      top: 50%;
      transform: translateY(-50%);
      color: #aaa;
      font-size: 0.8rem;
      pointer-events: none;
    }
    .filtro-inline input {
      padding-left: 2rem;
      padding-right: 4.6rem;
      font-size: 0.82rem;
    }
    .filtro-inline .btn-limpar-filtro {
      position: absolute;
      right: 0.4rem;
      top: 50%;
      transform: translateY(-50%);
      background: none;
      border: none;
      color: #aaa;
      font-size: 0.75rem;
      cursor: pointer;
      padding: 0;
      line-height: 1;
      display: none;
    }
    .filtro-inline .btn-limpar-filtro:hover { color: #555; }
    .kbd-hint {
      position: absolute;
      right: 0.45rem;
      top: 50%;
      transform: translateY(-50%);
      display: flex;
      align-items: center;
      gap: 2px;
      pointer-events: none;
    }
    .kbd-hint kbd {
      display: inline-block;
      padding: 1px 5px;
      font-size: 0.68rem;
      font-family: inherit;
      font-weight: 600;
      color: #212529;
      background: #fff;
      border: 1px solid #ced4da;
      border-radius: 3px;
      box-shadow: inset 0 -1px 0 rgba(0,0,0,.12);
      white-space: nowrap;
      line-height: 1.5;
    }
    .modal-detalhe .modal-header {
      background: #fff;
      color: #212529;
      border-radius: 6px 6px 0 0;
      padding: 0.75rem 1rem;
      border-bottom: 1px solid #dee2e6;
    }
    .modal-detalhe .modal-header .btn-close {
      filter: none;
    }
    .modal-detalhe .modal-title {
      font-size: 0.9rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }
    .detalhe-campo-label {
      font-size: 0.72rem;
      font-weight: 600;
      color: #6c757d;
      text-transform: uppercase;
      letter-spacing: 0.2px;
      margin-bottom: 0.15rem;
    }
    .detalhe-campo-valor {
      font-size: 0.88rem;
      color: #212529;
      padding: 0.3rem 0.5rem;
      background: #f8f9fa;
      border-radius: 4px;
      border: 1px solid #e9ecef;
      min-height: 2rem;
      display: flex;
      align-items: center;
      word-break: break-word;
    }
  </style>
</head>
<body>

<div class="toast-wrapper" id="toastWrapper"></div>
<div id="loader" class="overlay-loader d-none"><div class="spinner-circular"></div></div>

<%
strMsg     = ""
strMsgTipo = "success"

If Trim(Request.QueryString("aviso")) = "ok" Then
  Select Case Trim(UCase(Request.QueryString("acao")))
    Case "INS" : strMsg = "Cadastro realizado com sucesso!"
    Case "ALT" : strMsg = "Cadastro atualizado com sucesso!"
    Case "DEL" : strMsg = "Registro apagado com sucesso!"
  End Select
End If

ordenarPor     = Trim(Request("ordenarpor"))
ordenarAscDesc = Trim(Request("ordenarascdesc"))

SQL = "" & _
  "SELECT " & _
    "R.CODIGO " & _
    ",R.CNPJ " & _
    ",R.NOME " & _
  "FROM RECEBIMENTO_FORNECEDORES R " & _
  "WHERE ISNULL(EXCLUIDO,'') = '' "

If ordenarPor = "CODIGO" Then SQL = SQL & "ORDER BY CODIGO "
If ordenarPor = "CNPJ"   Then SQL = SQL & "ORDER BY CNPJ "
If ordenarPor = "NOME"   Then SQL = SQL & "ORDER BY NOME "
If ordenarPor <> "" And ordenarAscDesc = "ASC"  Then SQL = SQL & "ASC "
If ordenarPor <> "" And ordenarAscDesc = "DESC" Then SQL = SQL & "DESC "

SET RS = Conecta("PIR-CORPORATIVO", SQL)

totalRegistros = 0
If isObject(RS) And Not (RS Is Nothing) Then
  If Not RS.EOF Then
    RS.MoveLast : totalRegistros = RS.RecordCount : RS.MoveFirst
  End If
End If
%>

<div class="container-fluid py-3 pb-5">

  <div class="d-flex align-items-center justify-content-between mb-3">
    <h4 class="fw-bold mb-0 page-title">Lista de Fornecedores</h4>
  </div>

  <div class="panel">
    <div class="panel-header d-flex align-items-center justify-content-between">
      <button type="button" class="btn btn-success btn-sm"
              onclick="abrirModalCriar()">
        <i class="bi bi-plus-lg me-1"></i>Criar
      </button>
      <div class="filtro-inline">
        <i class="bi bi-search icone-busca"></i>
        <input type="text" id="inputFiltro" class="form-control form-control-sm"
               placeholder="Nome, CNPJ ou Código SAP...">
        <button class="btn-limpar-filtro" id="btnLimparFiltro" title="Limpar" style="right:3.8rem">
          <i class="bi bi-x-lg"></i>
        </button>
        <span class="kbd-hint" id="kbdHint">
          <kbd>Ctrl+K</kbd>
        </span>
      </div>
    </div>

    <div class="panel-body p-0">

      <input type="hidden" id="ordenarpor"     value="<%=ordenarPor%>">
      <input type="hidden" id="ordenarascdesc" value="<%=ordenarAscDesc%>">

      <div class="tbl-wrapper">
        <table class="table table-bordered table-striped mb-0" id="tblFornecedores">
          <thead>
            <tr>
              <th style="width:90px;">Ações</th>
              <th class="sortable" onclick="Ordenar('CODIGO')" title="Ordenar pelo código SAP">
                Código SAP <%=ShowOrderIcon("CODIGO", ordenarPor, ordenarAscDesc)%>
              </th>
              <th class="sortable" onclick="Ordenar('CNPJ')" title="Ordenar pelo CNPJ">
                CNPJ <%=ShowOrderIcon("CNPJ", ordenarPor, ordenarAscDesc)%>
              </th>
              <th class="sortable" onclick="Ordenar('NOME')" title="Ordenar pelo Nome">
                Nome <%=ShowOrderIcon("NOME", ordenarPor, ordenarAscDesc)%>
              </th>
            </tr>
          </thead>
          <tbody id="tbodyFornecedores">
<%
If Not isObject(RS) Or (RS Is Nothing) Then
%>
            <tr class="tr-estado">
              <td colspan="4">
                <div class="tbl-vazia">
                  <i class="bi bi-exclamation-triangle" style="font-size:1.4rem;color:#dc3545"></i>
                  <p class="mt-2 mb-0">Erro ao Conectar com o Banco de Dados.</p>
                </div>
              </td>
            </tr>
<%
ElseIf RS.EOF Then
%>
            <tr class="tr-estado">
              <td colspan="4">
                <div class="tbl-vazia">
                  <i class="bi bi-inbox" style="font-size:1.4rem"></i>
                  <p class="mt-2 mb-0">Nenhum Fornecedor cadastrado.</p>
                </div>
              </td>
            </tr>
<%
Else
  Do While Not RS.EOF
%>
            <tr data-codigo="<%=RS("CODIGO")%>"
                data-cnpj="<%=RS("CNPJ")%>"
                data-nome="<%=Server.HTMLEncode(RS("NOME"))%>">
              <td>
                <div class="d-flex align-items-center justify-content-center gap-1">
                  <button type="button"
                          class="btn-acao btn btn-primary btn-sm"
                          title="Visualizar"
                          onclick="abrirModal('<%=RS("CODIGO")%>','<%=Server.HTMLEncode(RS("CNPJ"))%>','<%=Server.HTMLEncode(RS("NOME"))%>')">
                    <i class="bi bi-eye"></i>
                  </button>
                  <button type="button"
                          class="btn-acao btn btn-outline-warning btn-sm"
                          title="Editar"
                          onclick="abrirModalEditar('<%=RS("CODIGO")%>')">
                    <i class="bi bi-pencil"></i>
                  </button>
                  <button type="button"
                          class="btn-acao btn btn-outline-danger btn-sm"
                          title="Apagar"
                          onclick="confirmarApagar('<%=RS("CODIGO")%>','<%=Server.HTMLEncode(RS("NOME"))%>')">
                    <i class="bi bi-trash"></i>
                  </button>
                </div>
              </td>
              <td><%=RS("CODIGO")%></td>
              <td><%=RS("CNPJ")%></td>
              <td><%=RS("NOME")%></td>
            </tr>
<%
    RS.MoveNext()
  Loop
  RS.Close()
End If
%>
          </tbody>
        </table>
      </div>
    </div>

    <div class="rodape-paginacao">
      <div class="rodape-esquerda">
        <label class="form-label mb-0" for="cmbLinhas" style="font-size:0.78rem">Exibir</label>
        <select id="cmbLinhas" class="form-select form-select-sm" style="width:auto">
          <option value="10">10</option>
          <option value="25">25</option>
          <option value="50" selected>50</option>
          <option value="100">100</option>
          <option value="200">200</option>
          <option value="500">500</option>
        </select>
      </div>
      <div class="rodape-centro">
        <nav aria-label="Paginação">
          <ul class="pagination pagination-sm mb-0 paginacao" id="paginacaoGrid"></ul>
        </nav>
      </div>
      <div class="rodape-direita" id="contadorGrid"></div>
    </div>
  </div>

</div>

<div class="modal fade modal-detalhe" id="modalVisualizacao" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content border-0 shadow">
      <div class="modal-header">
        <span class="modal-title">Dados do Fornecedor</span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body">
        <div class="row g-3">
          <div class="col-md-4">
            <div class="detalhe-campo-label">Código SAP</div>
            <div class="detalhe-campo-valor" id="detCodigo"></div>
          </div>
          <div class="col-md-8">
            <div class="detalhe-campo-label">CNPJ</div>
            <div class="detalhe-campo-valor" id="detCnpj"></div>
          </div>
          <div class="col-12">
            <div class="detalhe-campo-label">Razão Social</div>
            <div class="detalhe-campo-valor" id="detNome"></div>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="modalConfirmarApagar" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" style="max-width:420px">
    <div class="modal-content border-0 shadow">
      <div class="modal-header border-0 pb-2" style="border-bottom:1px solid #d1c9c9!important">
        <span class="modal-title fw-bold text-danger" style="font-size:0.95rem">
          Atenção
        </span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body pt-2">
        <p class="mb-0 text-center fw-bold" style="font-size:0.85rem;color:#495057">
          Esta ação é <strong class="text-danger">crítica</strong>.
        </p>
        <p class="mt-1 mb-1 text-center" style="font-size:0.85rem;color:#495057">
          O fornecedor abaixo será excluído:
        </p>
        <div class="p-2 mt-2 rounded border text-center" style="background:#fff5f5;border-color:#f5c6cb!important;font-size:0.83rem">
          <div style="color:#6c757d;font-size:0.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.2px">Fornecedor</div>
          <div id="apagarNome" style="color:#212529;font-weight:600;margin-top:2px"></div>
        </div>
      </div>
      <div class="modal-footer pt-1" style="border-top:1px solid #dee2e6">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
          Fechar
        </button>
        <button type="button" class="btn btn-sm btn-danger" id="btnConfirmarApagar">
          Excluir
        </button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade modal-detalhe" id="modalEditar" tabindex="-1" aria-hidden="true" data-bs-backdrop="static">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content border-0 shadow">
      <div class="modal-header">
        <span class="modal-title">Editar Fornecedor</span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body">

        <div id="editarSkeleton">
          <div class="row g-3">
            <div class="col-md-4">
              <div class="detalhe-campo-label mb-1">Código SAP</div>
              <div class="placeholder-glow"><span class="placeholder col-12 rounded" style="height:2rem"></span></div>
            </div>
            <div class="col-md-8">
              <div class="detalhe-campo-label mb-1">CNPJ</div>
              <div class="placeholder-glow"><span class="placeholder col-12 rounded" style="height:2rem"></span></div>
            </div>
            <div class="col-12">
              <div class="detalhe-campo-label mb-1">Razão Social</div>
              <div class="placeholder-glow"><span class="placeholder col-12 rounded" style="height:2rem"></span></div>
            </div>
          </div>
        </div>

        <div id="editarForm" style="display:none">
          <input type="hidden" id="editAcao"   value="ALT">
          <input type="hidden" id="editCodigo" name="edtCODIGO">
          <div class="row g-3">
            <div class="col-md-4">
              <label class="form-label" for="editCampoCodigoDisplay">Código SAP</label>
              <input type="text" id="editCampoCodigoDisplay" class="form-control form-control-sm"
                     readonly style="background:#f8f9fa;cursor:default">
            </div>
            <div class="col-md-8">
              <label class="form-label" for="editCampoCNPJ">CNPJ</label>
              <input type="text" id="editCampoCNPJ" name="edtCNPJ" class="form-control form-control-sm"
                     readonly style="background:#f8f9fa;cursor:default">
            </div>
            <div class="col-12">
              <label class="form-label" for="editCampoNome">
                Razão Social <span class="text-danger">*</span>
              </label>
              <input type="text" id="editCampoNome" name="edtNOME" class="form-control form-control-sm"
                     maxlength="255" autocomplete="off"
                     oninput="this.value = this.value.toUpperCase();">
              <div class="invalid-feedback">Razão Social é obrigatória.</div>
            </div>
          </div>
        </div>

        <div id="editarErro" style="display:none">
          <div class="text-center text-danger py-3" style="font-size:0.85rem">
            <i class="bi bi-exclamation-triangle me-1"></i>
            Não foi possível carregar os dados. Tente novamente.
          </div>
        </div>

      </div>
      <div class="modal-footer pt-1" style="border-top:1px solid #dee2e6">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
          Fechar
        </button>
        <button type="button" class="btn btn-sm btn-warning" id="btnSalvarEdicao" disabled>
          <span id="btnSalvarSpinner" class="spinner-border spinner-border-sm me-1 d-none" role="status"></span>
          Salvar
        </button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade modal-detalhe" id="modalCriar" tabindex="-1" aria-hidden="true" data-bs-backdrop="static">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content border-0 shadow">
      <div class="modal-header">
        <span class="modal-title">Novo Fornecedor</span>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
      </div>
      <div class="modal-body">

        <div id="criarForm">
          <div class="row g-3">
            <div class="col-md-4">
              <label class="form-label" for="criarCampoCodigo">
                Código SAP <span class="text-danger">*</span>
              </label>
              <input type="text" id="criarCampoCodigo" class="form-control form-control-sm"
                     maxlength="14" pattern="[0-9]*" inputmode="numeric"
                     autocomplete="off" placeholder="Somente números">
              <div class="invalid-feedback">Código SAP é obrigatório.</div>
            </div>
            <div class="col-md-8">
              <label class="form-label d-flex justify-content-between" for="criarCampoCNPJ">
                <span>CNPJ <span class="text-danger">*</span></span>
                <span id="criarCNPJContador" class="text-muted" style="font-size:0.75rem;font-weight:400">0/14</span>
              </label>
              <input type="text" id="criarCampoCNPJ" class="form-control form-control-sm"
                     maxlength="14" autocomplete="off"
                     placeholder="14 caracteres alfanuméricos"
                     oninput="this.value = this.value.toUpperCase();">
              <div class="invalid-feedback" id="criarCNPJFeedback">CNPJ é obrigatório.</div>
            </div>
            <div class="col-12">
              <label class="form-label" for="criarCampoNome">
                Razão Social <span class="text-danger">*</span>
              </label>
              <input type="text" id="criarCampoNome" class="form-control form-control-sm"
                     maxlength="255" autocomplete="off"
                     oninput="this.value = this.value.toUpperCase();">
              <div class="invalid-feedback">Razão Social é obrigatória.</div>
            </div>
          </div>
        </div>

        <div id="criarErro" style="display:none">
          <div class="text-center text-danger py-3" style="font-size:0.85rem">
            <i class="bi bi-exclamation-triangle me-1"></i>
            <span id="criarErroMensagem">Não foi possível realizar o cadastro. Tente novamente.</span>
          </div>
        </div>

      </div>
      <div class="modal-footer pt-1" style="border-top:1px solid #dee2e6">
        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
          Fechar
        </button>
        <button type="button" class="btn btn-sm btn-success" id="btnSalvarCriacao">
          <span id="btnCriarSpinner" class="spinner-border spinner-border-sm me-1 d-none" role="status"></span>
          Salvar
        </button>
      </div>
    </div>
  </div>
</div>

<template id="tmplToast">
  <div class="toast-item">
    <div class="toast-corpo">
      <i class="bi toast-icone"></i>
      <span class="toast-mensagem"></span>
      <button class="toast-fechar">&times;</button>
    </div>
    <div class="toast-barra"></div>
  </div>
</template>

<template id="tmplTrVazia">
  <tr class="tr-filtro-vazia">
    <td colspan="4">
      <div class="tbl-vazia">
        <i class="bi bi-search" style="font-size:1.4rem"></i>
        <p class="mt-2 mb-0">Nenhum resultado para "<strong class="termo-buscado"></strong>".</p>
      </div>
    </td>
  </tr>
</template>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"></script>
<script>
(function () {
  'use strict';

  const TOAST_DELAY = 4500;
  const TOAST_ICONS = {
    success : 'bi-check-circle-fill text-success',
    danger  : 'bi-x-circle-fill text-danger',
    info    : 'bi-info-circle-fill text-secondary',
  };

  function exibirToast(tipo, mensagem) {
    const tmpl  = document.getElementById('tmplToast');
    const el    = tmpl.content.cloneNode(true).querySelector('.toast-item');
    el.classList.add('toast-item-' + tipo);
    el.querySelector('.toast-icone').className      = 'bi ' + (TOAST_ICONS[tipo] || TOAST_ICONS.info);
    el.querySelector('.toast-mensagem').textContent = mensagem;
    el.querySelector('.toast-fechar').addEventListener('click', () => el.remove());
    const barra = el.querySelector('.toast-barra');
    barra.classList.add('toast-barra-' + tipo);
    barra.style.animationDuration = TOAST_DELAY + 'ms';
    document.getElementById('toastWrapper').appendChild(el);
    setTimeout(() => { if (el.parentNode) el.remove(); }, TOAST_DELAY + 300);
  }

  const todasLinhas = Array.from(
    document.querySelectorAll('#tbodyFornecedores tr:not(.tr-estado)')
  );

  let paginaAtual     = 1;
  let linhasFiltradas = todasLinhas;

  function normalizar(str) {
    return str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function filtrar(termo) {
    const t = normalizar(termo.trim());
    linhasFiltradas = t
      ? todasLinhas.filter(tr =>
          normalizar(tr.dataset.codigo || '').includes(t) ||
          normalizar(tr.dataset.cnpj   || '').includes(t) ||
          normalizar(tr.dataset.nome   || '').includes(t)
        )
      : todasLinhas;
    paginaAtual = 1;
    renderizar();
  }

  function renderizar() {
    const porPagina = parseInt(document.getElementById('cmbLinhas').value, 10);
    const total     = linhasFiltradas.length;
    const totalPags = Math.max(1, Math.ceil(total / porPagina));
    if (paginaAtual > totalPags) paginaAtual = totalPags;

    const inicio = (paginaAtual - 1) * porPagina;

    document.querySelectorAll('.tr-filtro-vazia').forEach(el => el.remove());
    todasLinhas.forEach(tr => { tr.style.display = 'none'; });

    const visiveis = linhasFiltradas.slice(inicio, inicio + porPagina);

    if (visiveis.length === 0) {
      const clone = document.getElementById('tmplTrVazia').content.cloneNode(true);
      const tr    = clone.querySelector('tr');
      const termo = document.getElementById('inputFiltro').value.trim();
      if (termo) tr.querySelector('.termo-buscado').textContent = termo;
      document.getElementById('tbodyFornecedores').appendChild(tr);
    } else {
      visiveis.forEach(tr => { tr.style.display = ''; });
    }

    renderizarPaginacao(total, porPagina, totalPags);

    const ini = total === 0 ? 0 : inicio + 1;
    const fim = Math.min(inicio + porPagina, total);
    const suf = total !== 1 ? 's' : '';
    document.getElementById('contadorGrid').innerHTML =
      '<strong>' + ini + '-' + fim + '</strong>&nbsp;de&nbsp;<strong>' + total + '</strong>&nbsp;Registro' + suf;
  }

  function renderizarPaginacao(total, porPagina, totalPags) {
    const ul = document.getElementById('paginacaoGrid');
    ul.innerHTML = '';

    const criarItem = (label, pagina, desabilitado, ativo) => {
      const li = document.createElement('li');
      li.className = 'page-item' +
        (desabilitado ? ' disabled' : '') +
        (ativo        ? ' active'   : '');
      const a = document.createElement('a');
      a.className = 'page-link';
      a.href      = '#';
      a.innerHTML = label;
      if (!desabilitado && !ativo) {
        a.addEventListener('click', e => {
          e.preventDefault();
          paginaAtual = pagina;
          renderizar();
        });
      }
      li.appendChild(a);
      ul.appendChild(li);
    };

    criarItem('<i class="bi bi-chevron-left"></i>', paginaAtual - 1, paginaAtual === 1, false);

    let anterior = null;
    paginasVisiveis(paginaAtual, totalPags).forEach(p => {
      if (anterior !== null && p - anterior > 1) criarItem('…', null, true, false);
      criarItem(p, p, false, p === paginaAtual);
      anterior = p;
    });

    criarItem('<i class="bi bi-chevron-right"></i>', paginaAtual + 1, paginaAtual === totalPags, false);
  }

  function paginasVisiveis(atual, total) {
    const delta  = 2;
    const inicio = Math.max(1, atual - delta);
    const fim    = Math.min(total, atual + delta);
    const pages  = [];
    if (inicio > 1) pages.push(1);
    for (let i = inicio; i <= fim; i++) pages.push(i);
    if (fim < total) pages.push(total);
    return pages;
  }

  window.Ordenar = function (por) {
    const hidAsc = document.getElementById('ordenarascdesc');
    document.getElementById('ordenarpor').value = por;
    hidAsc.value = hidAsc.value === 'ASC' ? 'DESC' : 'ASC';
    const form = document.createElement('form');
    form.method = 'post';
    form.action = 'grid_fornecedores.asp';
    ['ordenarpor', 'ordenarascdesc'].forEach(id => {
      const inp = document.createElement('input');
      inp.type  = 'hidden';
      inp.name  = id;
      inp.value = document.getElementById(id).value;
      form.appendChild(inp);
    });
    document.body.appendChild(form);
    form.submit();
  };

  window.abrirModal = function (codigo, cnpj, nome) {
    document.getElementById('detCodigo').textContent = codigo;
    document.getElementById('detCnpj').textContent   = cnpj;
    document.getElementById('detNome').textContent   = nome;
    new bootstrap.Modal(document.getElementById('modalVisualizacao')).show();
  };

  window.confirmarApagar = function (codigo, nome) {
    document.getElementById('apagarNome').textContent = nome;
    document.getElementById('btnConfirmarApagar').onclick = function () {
      location.href = 'repo_fornecedor.asp?acao=DEL&codigo=' + encodeURIComponent(codigo);
    };
    new bootstrap.Modal(document.getElementById('modalConfirmarApagar')).show();
  };

  const REPOSITORIO     = 'repo_fornecedor.asp';
  let   modalEditarInst = null;

  function elEditar(id) { return document.getElementById(id); }

  function estadoModal(estado) {
    elEditar('editarSkeleton').style.display = estado === 'loading' ? ''     : 'none';
    elEditar('editarForm').style.display     = estado === 'ready'   ? ''     : 'none';
    elEditar('editarErro').style.display     = estado === 'erro'    ? ''     : 'none';
    elEditar('btnSalvarEdicao').disabled     = estado !== 'ready';
  }

  window.abrirModalEditar = function (codigo) {
    if (!modalEditarInst) {
      modalEditarInst = new bootstrap.Modal(elEditar('modalEditar'));
    }
    elEditar('editCampoNome').classList.remove('is-invalid');
    estadoModal('loading');
    modalEditarInst.show();

    const xhr = new XMLHttpRequest();
    xhr.open('GET', REPOSITORIO + '?acao=GET&codigo=' + encodeURIComponent(codigo), true);
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      if (xhr.status === 200) {
        try {
          const resp = JSON.parse(xhr.responseText);
          if (resp.sucesso) {
            elEditar('editCodigo').value            = resp.dados.CODIGO;
            elEditar('editCampoCodigoDisplay').value = resp.dados.CODIGO;
            elEditar('editCampoCNPJ').value          = resp.dados.CNPJ;
            elEditar('editCampoNome').value          = resp.dados.NOME;
            estadoModal('ready');
            setTimeout(function () { elEditar('editCampoNome').focus(); }, 50);
          } else {
            estadoModal('erro');
          }
        } catch (e) {
          estadoModal('erro');
        }
      } else {
        estadoModal('erro');
      }
    };
    xhr.send();
  };

  elEditar('btnSalvarEdicao').addEventListener('click', function () {
    const campoNome = elEditar('editCampoNome');
    if (!campoNome.value.trim()) {
      campoNome.classList.add('is-invalid');
      campoNome.focus();
      return;
    }
    campoNome.classList.remove('is-invalid');

    elEditar('btnSalvarSpinner').classList.remove('d-none');
    elEditar('btnSalvarEdicao').disabled = true;

    const params =
      'acao=ALT' +
      '&edtCODIGO=' + encodeURIComponent(elEditar('editCodigo').value) +
      '&edtCNPJ='   + encodeURIComponent(elEditar('editCampoCNPJ').value) +
      '&edtNOME='   + encodeURIComponent(elEditar('editCampoNome').value);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', REPOSITORIO, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;

      elEditar('btnSalvarSpinner').classList.add('d-none');
      elEditar('btnSalvarEdicao').disabled = false;

      if (xhr.status === 200) {
        try {
          const resp = JSON.parse(xhr.responseText);
          if (resp.sucesso) {
            modalEditarInst.hide();

            const codigo   = elEditar('editCodigo').value;
            const novoNome = elEditar('editCampoNome').value.trim();
            const tr = document.querySelector('#tbodyFornecedores tr[data-codigo="' + codigo + '"]');
            if (tr) {
              tr.dataset.nome = novoNome;
              const tds = tr.querySelectorAll('td');
              if (tds[3]) tds[3].textContent = novoNome;
              const btnVis = tr.querySelector('button[title="Visualizar"]');
              if (btnVis) {
                const cnpj = tr.dataset.cnpj;
                btnVis.setAttribute('onclick',
                  "abrirModal('" + codigo + "','" + cnpj.replace(/'/g, "\\'") + "','" + novoNome.replace(/'/g, "\\'") + "')");
              }
            }

            exibirToast('success', 'Fornecedor atualizado com sucesso!');
          } else {
            exibirToast('danger', resp.mensagem || 'Erro ao salvar. Tente novamente.');
          }
        } catch (e) {
          exibirToast('danger', 'Resposta inesperada do servidor.');
        }
      } else {
        exibirToast('danger', 'Erro de comunicação. Tente novamente.');
      }
    };
    xhr.send(params);
  });

  // ── Modal Criar ──────────────────────────────────────────────
  let modalCriarInst = null;

  function elCriar(id) { return document.getElementById(id); }

  function resetModalCriar() {
    ['criarCampoCodigo', 'criarCampoCNPJ', 'criarCampoNome'].forEach(function (id) {
      const el = elCriar(id);
      el.value = '';
      el.classList.remove('is-invalid');
    });
    const contador = elCriar('criarCNPJContador');
    contador.textContent = '0/14';
    contador.style.color = '';
    elCriar('criarCNPJFeedback').textContent = 'CNPJ é obrigatório.';
    elCriar('criarErro').style.display = 'none';
    elCriar('btnCriarSpinner').classList.add('d-none');
    elCriar('btnSalvarCriacao').disabled = false;
  }

  window.abrirModalCriar = function () {
    if (!modalCriarInst) {
      modalCriarInst = new bootstrap.Modal(elCriar('modalCriar'));
      elCriar('modalCriar').addEventListener('hidden.bs.modal', resetModalCriar);
    }
    resetModalCriar();
    modalCriarInst.show();
    setTimeout(function () { elCriar('criarCampoCodigo').focus(); }, 50);
  };

  // Contador de caracteres do CNPJ em tempo real
  elCriar('criarCampoCNPJ').addEventListener('input', function () {
    this.value = this.value.toUpperCase();
    const chars    = this.value.replace(/[^A-Za-z0-9]/g, '').length;
    const contador = elCriar('criarCNPJContador');
    const completo = chars === 14;
    contador.textContent = chars + '/14';
    contador.style.color = chars === 0 ? '' : completo ? '#198754' : '#dc3545';
    // Remove is-invalid em tempo real ao atingir 14
    if (completo) {
      this.classList.remove('is-invalid');
    }
  });

  elCriar('btnSalvarCriacao').addEventListener('click', function () {
    const campoCodigo = elCriar('criarCampoCodigo');
    const campoCNPJ   = elCriar('criarCampoCNPJ');
    const campoNome   = elCriar('criarCampoNome');
    let valido = true;

    [campoCodigo, campoCNPJ, campoNome].forEach(function (c) { c.classList.remove('is-invalid'); });

    if (!campoCodigo.value.trim()) { campoCodigo.classList.add('is-invalid'); valido = false; }

    const charsCNPJ = campoCNPJ.value.replace(/[^A-Za-z0-9]/g, '').length;
    if (charsCNPJ !== 14) {
      campoCNPJ.classList.add('is-invalid');
      elCriar('criarCNPJFeedback').textContent =
        charsCNPJ === 0 ? 'CNPJ é obrigatório.'
                        : 'CNPJ deve conter exatamente 14 caracteres (' + charsCNPJ + '/14 preenchidos).';
      valido = false;
    }

    if (!campoNome.value.trim())   { campoNome.classList.add('is-invalid');   valido = false; }

    if (!valido) {
      (campoCodigo.classList.contains('is-invalid') ? campoCodigo
        : campoCNPJ.classList.contains('is-invalid') ? campoCNPJ
        : campoNome).focus();
      return;
    }

    elCriar('criarErro').style.display = 'none';
    elCriar('btnCriarSpinner').classList.remove('d-none');
    elCriar('btnSalvarCriacao').disabled = true;

    const params =
      'acao=INS' +
      '&edtCODIGO=' + encodeURIComponent(campoCodigo.value.trim()) +
      '&edtCNPJ='   + encodeURIComponent(campoCNPJ.value.trim()) +
      '&edtNOME='   + encodeURIComponent(campoNome.value.trim());

    const xhr = new XMLHttpRequest();
    xhr.open('POST', REPOSITORIO, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;

      elCriar('btnCriarSpinner').classList.add('d-none');
      elCriar('btnSalvarCriacao').disabled = false;

      if (xhr.status === 200) {
        try {
          const resp = JSON.parse(xhr.responseText);
          if (resp.sucesso) {
            modalCriarInst.hide();
            exibirToast('success', 'Fornecedor cadastrado com sucesso!');
            setTimeout(function () { location.reload(); }, 1200);
          } else {
            elCriar('criarErroMensagem').textContent = resp.mensagem || 'Erro ao cadastrar. Tente novamente.';
            elCriar('criarErro').style.display = '';
          }
        } catch (e) {
          elCriar('criarErroMensagem').textContent = 'Resposta inesperada do servidor.';
          elCriar('criarErro').style.display = '';
        }
      } else {
        elCriar('criarErroMensagem').textContent = 'Erro de comunicação. Tente novamente.';
        elCriar('criarErro').style.display = '';
      }
    };
    xhr.send(params);
  });
  // ── /Modal Criar ─────────────────────────────────────────────

  const inputFiltro = document.getElementById('inputFiltro');
  const btnLimpar   = document.getElementById('btnLimparFiltro');
  const kbdHint     = document.getElementById('kbdHint');

  let debounce;
  inputFiltro.addEventListener('input', () => {
    const temTexto = inputFiltro.value.length > 0;
    btnLimpar.style.display = temTexto ? 'block' : 'none';
    kbdHint.style.display   = temTexto ? 'none'  : '';
    clearTimeout(debounce);
    debounce = setTimeout(() => filtrar(inputFiltro.value), 200);
  });

  btnLimpar.addEventListener('click', () => {
    inputFiltro.value       = '';
    btnLimpar.style.display = 'none';
    kbdHint.style.display   = '';
    filtrar('');
    inputFiltro.focus();
  });

  document.addEventListener('keydown', function (e) {
    if (e.ctrlKey && e.key === 'k') {
      e.preventDefault();
      inputFiltro.focus();
      inputFiltro.select();
    }
  });

  document.getElementById('cmbLinhas').addEventListener('change', () => {
    paginaAtual = 1;
    renderizar();
  });

  <%
  If strMsg <> "" Then
    Response.Write "exibirToast('" & strMsgTipo & "', '" & strMsg & "');"
  End If
  %>

  renderizar();

})();
</script>

</body>
</html>

<%
Function ShowOrderIcon(campo, ordenarPor, ordenarAscDesc)
  If ordenarPor = campo Then
    Dim seta
    If ordenarAscDesc = "ASC" Then seta = "bi-arrow-down" Else seta = "bi-arrow-up"
    ShowOrderIcon = "<i class=""bi " & seta & " sort-icon""></i>"
  Else
    ShowOrderIcon = "<i class=""bi bi-arrow-down-up sort-icon""></i>"
  End If
End Function
%>
