<%
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046
%>

<!-- #include file="../../cnn2/cnn2.asp" -->
<!-- #include file="../../includes/variaveis.asp" -->
<!-- #include file="../../funcoes/funcoes.asp" -->

<%
Dim nf, serie, filial, acao
nf     = trim(request.QueryString("nf"))
serie  = trim(request.QueryString("serie"))
filial = trim(request("filial"))
acao   = trim(request("acao"))

if len(filial) > 2 then
  filial = mid(filial, 2, 1)
end if

Dim erroMensagem, avisoMensagem
erroMensagem  = ""
avisoMensagem = ""

if trim(request("aviso")) = "erro" then
  erroMensagem = "Preenchimento incorreto, tente novamente!"
elseif trim(request("aviso")) = "salvou" then
  select case ucase(acao)
  case "INS" avisoMensagem = "Canhoto incluído com sucesso!"
  case "ALT" avisoMensagem = "Canhoto alterado com sucesso!"
  case "DEL" avisoMensagem = "Canhoto apagado com sucesso!"
  end select
end if

Dim edtNOTA, edtSERIE, edtFILIAL, edtPRODUTO, edtSERVICO, edtOBS

if ucase(acao) <> "INS" then

  Dim SQLform
  SQLform = "SELECT NF, SERIE, DATA, USERDIGIT, OBS, " &_
              "isnull(Carga,0) Carga, isnull(Veiculo,0) Veiculo, Filial " &_
            "FROM CANHOTOS " &_
            "WHERE D_E_L_E_T_ <> '*' " &_
              "AND NF = '" & nf & "' " &_
              "AND SERIE = '" & serie & "' " &_
              "AND Filial = '" & filial & "'"

  set RSform = Conecta("PIR-CORPORATIVO", SQLform)

  if not RSform.EOF then
    edtNOTA    = RSform("NF")
    edtSERIE   = RSform("SERIE")
    edtFILIAL  = RSform("Filial")
    edtPRODUTO = RSform("Carga")
    edtSERVICO = RSform("Veiculo")
    edtOBS     = trim(RSform("OBS"))
  end if
  RSform.Close()

else
  edtNOTA    = ""
  edtSERIE   = ""
  edtFILIAL  = filial
  edtPRODUTO = 0
  edtSERVICO = 0
  edtOBS     = ""
end if

Dim titPlanta
select case filial
case "0" titPlanta = "Pirapetinga"
case "4" titPlanta = "Uberaba"
case "5" titPlanta = "Bento Gonçalves"
case "6" titPlanta = "Fortaleza"
case "7" titPlanta = "Saquarema"
case else titPlanta = "Planta " & filial
end select

Dim titAcao
select case ucase(acao)
case "INS" titAcao = "Novo Canhoto"
case "ALT" titAcao = "Editar Canhoto"
case "CON" titAcao = "Consultar Canhoto"
case "DEL" titAcao = "Apagar Canhoto"
case else  titAcao = "Canhoto"
end select
Dim urlVolta
urlVolta = "grid_canhoto.asp?cboPlanta=0" & filial
Dim soLeitura
soLeitura = (ucase(acao) = "CON" or ucase(acao) = "DEL")
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — <%=titAcao%></title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <% if Session("NIVEL") <> "0" then %>
  <script src="../../js/sw_noviewsrc.js"></script>
  <% end if %>

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
      font-family: var(--bs-font-sans-serif);
    }
    .form-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .form-card-header {
      padding: .75rem 1.25rem;
      border-bottom: 1px solid #dee2e6;
      border-radius: 6px 6px 0 0;
    }
    .form-card-body {
      padding: 1.25rem;
    }
    .form-label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }
    .form-control, .form-select { font-size: .85rem; }
    .form-control:read-only     { background-color: #f8f9fa; }
    .badge-acao                 { font-size: .72rem; font-weight: 600; letter-spacing: .03em; }

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

    #modalConfirmarApagar .modal-header {
      border-bottom: 2px solid #dc3545;
    }
  </style>
</head>

<body>

  <div class="toast-container-custom" id="toastContainer"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">
    <div class="row justify-content-center">
      <div class="col-12 col-md-8 col-lg-6 col-xl-5">
        <div class="d-flex align-items-center mb-3 gap-2">
          <a href="<%=urlVolta%>" class="btn btn-sm btn-outline-secondary">
            <i class="bi bi-arrow-left"></i>
          </a>
          <div class="d-flex align-items-center gap-2 flex-grow-1">
            <span style="font-size:1.15rem;font-weight:700;color:#212529;"><%=titAcao%></span>
            <span class="text-muted" style="font-size:.85rem;"><%=titPlanta%></span>
          </div>
          <%
          select case ucase(acao)
          case "INS" response.write "<span class=""badge bg-success badge-acao"">Novo</span>"
          case "ALT" response.write "<span class=""badge bg-primary badge-acao"">Editar</span>"
          case "CON" response.write "<span class=""badge bg-secondary badge-acao"">Consulta</span>"
          case "DEL" response.write "<span class=""badge bg-danger badge-acao"">Apagar</span>"
          end select
          %>
        </div>

        <% if trim(erroMensagem) <> "" then %>
        <div class="alert alert-danger alert-dismissible fade show d-flex align-items-center gap-2 mb-3" role="alert">
          <i class="bi bi-exclamation-triangle-fill"></i>
          <span><%=erroMensagem%></span>
          <button type="button" class="btn-close ms-auto" data-bs-dismiss="alert"></button>
        </div>
        <% end if %>

        <div class="form-card">
          <div class="form-card-header d-flex align-items-center gap-2">
            <i class="bi bi-file-earmark-text text-secondary"></i>
            <span style="font-weight:600;font-size:.9rem;">Dados do Canhoto</span>
          </div>
          <div class="form-card-body">
            <form id="canhoto" name="canhoto" method="post" action="rpst_canhoto.asp" novalidate>
              <input type="hidden" name="acao"      id="acao"      value="<%=acao%>" />
              <input type="hidden" name="edtFILIAL" id="edtFILIAL" value="<%=edtFILIAL%>" />
              <div class="mb-3">
                <label for="cboPlanta" class="form-label">Planta</label>
                <input type="text" id="cboPlanta" class="form-control form-control-sm"
                       value="<%=titPlanta%>" readonly />
              </div>

              <div class="row g-2 mb-3">
                <div class="col">
                  <label for="edtNOTA" class="form-label">Nota Fiscal</label>
                  <input id="edtNOTA" name="edtNOTA" type="text"
                         class="form-control form-control-sm"
                         minlength="1" maxlength="9"
                         value="<%=edtNOTA%>"
                         <%if ucase(acao)<>"INS" then%>readonly<%end if%>
                         required />
                </div>
                <div class="col">
                  <label for="edtSERIE" class="form-label">Série</label>
                  <input id="edtSERIE" name="edtSERIE" type="text"
                         class="form-control form-control-sm"
                         minlength="1" maxlength="3"
                         value="<%=edtSERIE%>"
                         <%if ucase(acao)<>"INS" then%>readonly<%end if%>
                         required />
                </div>
              </div>

              <div class="row g-2 mb-3">
                <div class="col">
                  <label for="edtPRODUTO" class="form-label">Produto (Carga)</label>
                  <select id="edtPRODUTO" name="edtPRODUTO"
                          class="form-select form-select-sm"
                          <%if soLeitura then%>disabled<%end if%>>
                    <option value="">Selecione...</option>
                    <option value="0" <%if CStr(edtPRODUTO)="0" then%>selected<%end if%>>Não respondido</option>
                    <option value="5" <%if CStr(edtPRODUTO)="5" then%>selected<%end if%>>Muito Satisfeito</option>
                    <option value="4" <%if CStr(edtPRODUTO)="4" then%>selected<%end if%>>Satisfeito</option>
                    <option value="3" <%if CStr(edtPRODUTO)="3" then%>selected<%end if%>>Indiferente</option>
                    <option value="2" <%if CStr(edtPRODUTO)="2" then%>selected<%end if%>>Insatisfeito</option>
                    <option value="1" <%if CStr(edtPRODUTO)="1" then%>selected<%end if%>>Muito insatisfeito</option>
                  </select>
                </div>
                <div class="col">
                  <label for="edtSERVICO" class="form-label">Serviço (Veículo)</label>
                  <select id="edtSERVICO" name="edtSERVICO"
                          class="form-select form-select-sm"
                          <%if soLeitura then%>disabled<%end if%>>
                    <option value="">Selecione...</option>
                    <option value="0" <%if CStr(edtSERVICO)="0" then%>selected<%end if%>>Não respondido</option>
                    <option value="5" <%if CStr(edtSERVICO)="5" then%>selected<%end if%>>Muito Satisfeito</option>
                    <option value="4" <%if CStr(edtSERVICO)="4" then%>selected<%end if%>>Satisfeito</option>
                    <option value="3" <%if CStr(edtSERVICO)="3" then%>selected<%end if%>>Indiferente</option>
                    <option value="2" <%if CStr(edtSERVICO)="2" then%>selected<%end if%>>Insatisfeito</option>
                    <option value="1" <%if CStr(edtSERVICO)="1" then%>selected<%end if%>>Muito insatisfeito</option>
                  </select>
                </div>
              </div>

              <div class="mb-4">
                <label for="edtOBS" class="form-label">Observação</label>
                <input type="text" id="edtOBS" name="edtOBS"
                       class="form-control form-control-sm"
                       maxlength="100"
                       value="<%=edtOBS%>"
                       <%if soLeitura then%>readonly<%end if%> />
              </div>

              <div class="d-flex gap-2 justify-content-end">
                <%
                if ucase(acao) = "INS" then
                %>
                <button type="button" class="btn btn-primary btn-sm" onclick="verificarESalvar()">
                  Salvar
                </button>
                <%
                elseif ucase(acao) = "ALT" then
                %>
                <button type="button" class="btn btn-primary btn-sm" onclick="verificarESalvar()">
                  Salvar
                </button>
                <%
                elseif ucase(acao) = "DEL" then
                %>
                <button type="button" class="btn btn-danger btn-sm" onclick="confirmarApagar()">
                  Apagar
                </button>
                <%
                end if
                %>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="modal fade" id="modalConfirmarApagar" tabindex="-1" aria-labelledby="modalConfirmarApagarLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-sm">
      <div class="modal-content">
        <div class="modal-header py-2">
          <h6 class="modal-title d-flex align-items-center gap-2" id="modalConfirmarApagarLabel">
            <i class="bi bi-exclamation-triangle-fill text-danger"></i>
            Confirmar exclusão
          </h6>
          <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="modal"></button>
        </div>
        <div class="modal-body" style="font-size:.875rem;">
          Deseja realmente apagar este registro? Esta ação não pode ser facilmente desfeita.
        </div>
        <div class="modal-footer py-2 gap-2">
          <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
            Fechar
          </button>
          <button type="button" class="btn btn-sm btn-danger" id="btnConfirmarApagarOk">
            Apagar
          </button>
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
    const TOAST_ICON  = Object.freeze({
      sucesso: { icon: "bi-check-circle-fill",           color: "#198754" },
      erro:    { icon: "bi-x-circle-fill",               color: "#dc3545" },
      aviso:   { icon: "bi-exclamation-triangle-fill",   color: "#856404" },
      info:    { icon: "bi-info-circle-fill",            color: "#6c757d" }
    });

    function exibirToast(tipo, titulo, mensagem) {
      const cfg  = TOAST_ICON[tipo] || TOAST_ICON.info;
      const cont = document.getElementById("toastContainer");
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

    window.verificarESalvar = function() {
      const nota    = document.getElementById("edtNOTA").value.trim();
      const serie   = document.getElementById("edtSERIE").value.trim();
      const produto = document.getElementById("edtPRODUTO").value.trim();
      const servico = document.getElementById("edtSERVICO").value.trim();
      const filial  = document.getElementById("edtFILIAL").value;

      if (!nota)    { exibirToast("erro", "Campo obrigatório", "Número da Nota Fiscal não informado."); return; }
      if (!serie)   { exibirToast("erro", "Campo obrigatório", "Número da Série não informado."); return; }
      if (!produto) { exibirToast("erro", "Campo obrigatório", "Avaliação de Produto não informada."); return; }
      if (!servico) { exibirToast("erro", "Campo obrigatório", "Avaliação de Serviço não informada."); return; }
      if (filial === "7" && serie !== "1" && serie !== "12") {
        exibirToast("erro", "Série inválida", "Para Saquarema a série deve ser 1 ou 12.");
        return;
      }
      document.getElementById("canhoto").submit();
    };

    window.confirmarApagar = function() {
      const modal = new bootstrap.Modal(document.getElementById("modalConfirmarApagar"));
      modal.show();
    };

    document.getElementById("btnConfirmarApagarOk").addEventListener("click", function() {
      bootstrap.Modal.getInstance(document.getElementById("modalConfirmarApagar")).hide();
      document.getElementById("canhoto").submit();
    });

    <%
    if trim(avisoMensagem) <> "" then
      response.write "document.addEventListener('DOMContentLoaded', () => {"
      response.write "  exibirToast('sucesso', 'Operação concluída', '" & avisoMensagem & "');"
      response.write "});"
    end if
    %>

  })();
  </script>

</body>
</html>