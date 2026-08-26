<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include file = "../funcoes/funcoes.asp" -->
<!--#include file = "../includes/variaveis.asp" -->

<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>

<%
Const PARAM_DESATIVAR = "D"
Const PARAM_ATIVAR    = "A"
Const DELETADO        = "*"
Const CONN_LOG        = "PIR-CORPORATIVO"

Dim sCsrfPost, sCsrfSessao
Dim sWhereEmpresa, sFilial
sCsrfPost   = Trim(Request("csrf_token") & "")
sCsrfSessao = Trim(Session("csrf_excusuario") & "")

If Trim(Request("ListaUsuario")) <> "" Then

  If sCsrfPost = "" Or sCsrfPost <> sCsrfSessao Then
    Response.Write "<script>var msgErro='Requisição inválida. Recarregue a página e tente novamente.';</script>"
  Else

    Dim sRecno, sParam
    sRecno = Trim(Request("ListaUsuario"))
    sParam = Trim(Request("btnParam"))

    If sParam = PARAM_DESATIVAR Then

      SQL = "" &_
        "BEGIN TRANSACTION " &_
          "UPDATE CRM_DTI SET " &_
            "CRM_D_E_L_E_T = '" & DELETADO & "' " &_
          "WHERE CRM_R_E_C_N_O = '"& sRecno &"' " &_
            "AND CRM_D_E_L_E_T <> '" & DELETADO & "' " &_
        "IF @@ERROR > 0 " &_
          "ROLLBACK TRANSACTION " &_
        "ELSE " &_
          "COMMIT TRANSACTION "

      SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
      SET PIR_HELPDESK = NOTHING

      SQL = "" &_
        "INSERT INTO dbo.LogAtividade " &_
          "(IdUsuarioAcao, IdUsuarioAlvo, Sistema, Rotina, Acao, Descricao, DataCriacao) " &_
        "VALUES (" &_
          "'"& Replace(Session("USR"), "'", "''") &"', " &_
          "'"& Replace(sRecno, "'", "''") &"', " &_
          "'SWnet', " &_
          "'ExcUsuario', " &_
          "'DESATIVAR', " &_
          "'Usuario RECNO " & sRecno & " desativado pelo operador " & Replace(Session("USR"), "'", "''") & " (IP: " & Request.ServerVariables("REMOTE_ADDR") & ")', " &_
          "GETDATE()" &_
        ")"
      SET PIR_LOG = Conecta(CONN_LOG, SQL)
      SET PIR_LOG = NOTHING

      Response.Write "<script>var msgSucesso='Usuário desativado com sucesso!';</script>"

    Else

      SQL = "" &_
        "UPDATE CRM_DTI SET " &_
          "CRM_D_E_L_E_T = '', " &_
          "DataBloqueio = NULL " &_
        "WHERE CRM_R_E_C_N_O = '"& sRecno &"' " &_
          "AND CRM_D_E_L_E_T = '" & DELETADO & "' " &_
        "IF @@ERROR > 0 " &_
          "ROLLBACK TRANSACTION " &_
        "ELSE " &_
          "COMMIT TRANSACTION "

      SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
      SET PIR_HELPDESK = NOTHING

      SQL = "" &_
        "INSERT INTO dbo.LogAtividade " &_
          "(IdUsuarioAcao, IdUsuarioAlvo, Sistema, Rotina, Acao, Descricao, DataCriacao) " &_
        "VALUES (" &_
          "'"& Replace(Session("USR"), "'", "''") &"', " &_
          "'"& Replace(sRecno, "'", "''") &"', " &_
          "'SWnet', " &_
          "'ExcUsuario', " &_
          "'ATIVAR', " &_
          "'Usuario RECNO " & sRecno & " reativado pelo operador " & Replace(Session("USR"), "'", "''") & " (IP: " & Request.ServerVariables("REMOTE_ADDR") & ")', " &_
          "GETDATE()" &_
        ")"
      SET PIR_LOG = Conecta(CONN_LOG, SQL)
      SET PIR_LOG = NOTHING

      Response.Write "<script>var msgSucesso='Usuário ativado com sucesso!';</script>"

    End If

  End If

End If

Dim sCsrfNovo
sCsrfNovo = Left(Replace(CStr(Now()) & CStr(Rnd()), " ", ""), 8) & _
            Hex(Int(Rnd() * 65535)) & _
            Hex(Int(Rnd() * 65535))
Session("csrf_excusuario") = sCsrfNovo

Dim sNomeOperador, rsOperador
sNomeOperador = ""
SQL = "" &_
  "SELECT TOP 1 CRM_NOME " &_
  "FROM CRM_DTI " &_
  "WHERE RTRIM(LTRIM(CRM_LOGIN)) = '" & Replace(Session("USR"), "'", "''") & "' " &_
    "AND CRM_D_E_L_E_T <> '*'"
SET rsOperador = Conecta("PIR-HELPDESK", SQL)
If IsObject(rsOperador) And Not (rsOperador IS NOTHING) Then
  If Not rsOperador.EOF Then
    sNomeOperador = Trim(rsOperador("CRM_NOME") & "")
  End If
  rsOperador.Close()
  SET rsOperador = NOTHING
End If
%>

<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />
    <title>SWnet - Gerenciar Usuário</title>

    <link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

    <style>
      body {
        background: #f0f0f0;
        color: #212529;
        font-size: 0.8125rem;
        min-height: 100vh;
      }

      .form-label { font-weight: 600; }

      .form-control,
      .form-select {
        font-size: 0.85rem;
      }

      .form-control:focus,
      .form-select:focus {
        border-color: #4a4a4a;
        box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
      }

      .form-control:disabled,
      .form-control[readonly] {
        background-color: #fff;
        opacity: 1;
      }

      .form-card {
        background: #fff;
      }

      .form-card.has-data {
        background: #f7f7f7;
      }

      .form-card.has-data > .card-body {
        background: #f7f7f7;
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
      }

      .info-placeholder {
        padding: 2.5rem 1rem;
        text-align: center;
        color: #6c757d;
      }

      .info-placeholder i {
        font-size: 2.5rem;
        display: block;
        margin-bottom: 0.5rem;
      }

      .info-placeholder p {
        font-size: 0.75rem;
        margin: 0;
      }

      .badge-status {
        font-size: 0.6875rem;
        font-weight: 700;
        letter-spacing: 0.3px;
        padding: 0.25em 0.65em;
        border-radius: 4px;
        display: inline-block;
      }

      .badge-status-ativo {
        background: #d1e7dd;
        color: #0a3622;
      }

      .badge-status-inativo {
        background: #f8d7da;
        color: #842029;
      }

      .modal-content {
        border: 1px solid #dee2e6;
        border-radius: 6px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
      }

      .modal-header-custom {
        background: #fff;
        color: #212529;
        padding: 0.65rem 1rem;
        border-bottom: 1px solid #dee2e6;
        border-radius: 6px 6px 0 0;
      }

      .modal-header-custom .modal-title {
        font-size: 0.8125rem;
        font-weight: 600;
        color: #212529;
      }

      .modal-footer-custom {
        background: #f8f9fa;
        border-top: 1px solid #dee2e6;
        padding: 0.6rem 1rem;
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
      .toast-item-warning { border-left: 4px solid #ffc107; }
      .toast-item-danger  { border-left: 4px solid #dc3545; }
      .toast-item-info    { border-left: 4px solid #6c757d; }

      @keyframes entrar-toast {
        from { opacity: 0; transform: translateX(20px); }
        to   { opacity: 1; transform: translateX(0); }
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
      .toast-barra-warning { background: #ffc107; }
      .toast-barra-danger  { background: #dc3545; }
      .toast-barra-info    { background: #6c757d; }

      @keyframes encolher {
        from { width: 100%; }
        to   { width: 0%; }
      }

      .user-search-wrapper { position: relative; }

      .user-dropdown {
        position: absolute;
        top: 100%;
        left: 0;
        right: 0;
        max-height: 280px;
        overflow-y: auto;
        background: #fff;
        border: 1px solid #dee2e6;
        border-top: none;
        border-radius: 0 0 6px 6px;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.12);
        z-index: 1050;
        display: none;
      }

      .user-dropdown.show { display: block; }

      .user-dropdown-item {
        padding: 9px 14px;
        cursor: pointer;
        border-bottom: 1px solid #f5f5f5;
        transition: background 0.15s;
      }

      .user-dropdown-item:last-child { border-bottom: none; }

      .user-dropdown-item:hover,
      .user-dropdown-item.active { background: #f0f0f0; }

      .user-dropdown-login { font-size: 0.78rem; }

      .card-header-label {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.3px;
      }

      .user-dropdown-counter {
        padding: 5px 14px;
        font-size: 0.72rem;
        background: #fafafa;
        border-top: 1px solid #dee2e6;
        color: #888;
        text-align: right;
        position: sticky;
        bottom: 0;
      }

      .form-select:disabled {
        background-color: #fff;
        background-image: none;
        opacity: 1;
      }

      #divInfos { display: none; }
    </style>
  </head>

  <body>

    <div class="toast-wrapper" id="toastWrapper"></div>


    <div class="modal fade" id="modalConfirma" tabindex="-1" aria-labelledby="modalConfirmaLabel">
      <div class="modal-dialog modal-dialog-centered modal-sm">
        <div class="modal-content">
          <div class="modal-header-custom position-relative">
            <h5 class="modal-title d-flex align-items-center gap-2 mb-0" id="modalConfirmaLabel">
              <span id="modalConfirmaTitulo">Confirmar Ação</span>
            </h5>
            <button type="button" class="btn-close position-absolute top-50 end-0 translate-middle-y me-3" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>

          <div class="modal-body py-3">
            <p class="mb-1 small text-muted" id="modalConfirmaAcaoTexto"></p>
            <p class="fw-bold mb-0" id="modalConfirmaNomeUsuario" style="font-size:0.875rem;color:#4a4a4a;">—</p>
          </div>

          <div class="modal-footer-custom d-flex justify-content-end">
            <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
            <button type="button" class="btn btn-sm btn-danger ms-2" id="btnConfirmaAcao">
              <span id="btnConfirmaTexto">Confirmar</span>
            </button>
          </div>
        </div>
      </div>
    </div>

    <div class="container-fluid px-3 mt-5" style="max-width:54rem;">

        <form name="ExcluirUsuario" id="ExcluirUsuario" method="post" autocomplete="off">

          <input type="hidden" name="csrf_token"  id="csrf_token"  value="<%=sCsrfNovo%>" />
          <input type="hidden" name="FilialCargas" id="FilialCargas" />
          <input type="hidden" name="btnTitle"     id="btnTitle" />
          <input type="hidden" name="btnClass"     id="btnClass" />
          <input type="hidden" name="btnParam"     id="btnParam" />
          <input type="hidden" name="ListaUsuario" id="ListaUsuario" />

<%
SQL = "" &_
  "SELECT " &_
    "CRM_R_E_C_N_O, " &_
    "CRM_COD, " &_
    "CRM_NOME, " &_
    "CRM_LOGIN, " &_
    "CRM_D_E_L_E_T " &_
  "FROM CRM_DTI " &_
  "ORDER BY CRM_NOME"

SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)

If Not IsObject(PIR_HELPDESK1) Or (PIR_HELPDESK1 IS NOTHING) Then
  Response.End
End If
%>

          <script>
            var LoginUsuarioLogado = "<%=Session("USR")%>";
            var NomeUsuarioLogado  = "<%=sNomeOperador%>";
            var listaUsuarios = [
<%
Dim sNome, sLogin, sStatus, sVirgula
sVirgula = ""
While Not PIR_HELPDESK1.EOF
  sNome  = Replace(Replace(Trim(PIR_HELPDESK1("CRM_NOME")  & ""), "\", "\\"), """", "\""")
  sNome  = Replace(Replace(sNome,  "'", "\'"), vbCrLf, " ")
  sLogin = Replace(Replace(Trim(PIR_HELPDESK1("CRM_LOGIN") & ""), "\", "\\"), """", "\""")
  sLogin = Replace(Replace(sLogin, "'", "\'"), vbCrLf, " ")
  If Trim(PIR_HELPDESK1("CRM_D_E_L_E_T") & "") = "*" Then
    sStatus = "INATIVO"
  Else
    sStatus = "ATIVO"
  End If
  Response.Write sVirgula & vbCrLf
  Response.Write "              {" & vbCrLf
  Response.Write "                id: """    & PIR_HELPDESK1("CRM_R_E_C_N_O") & """," & vbCrLf
  Response.Write "                nome: """  & sNome   & """," & vbCrLf
  Response.Write "                login: """ & sLogin  & """," & vbCrLf
  Response.Write "                status: """& sStatus & """" & vbCrLf
  Response.Write "              }"
  sVirgula = ","
  PIR_HELPDESK1.MoveNext()
Wend
PIR_HELPDESK1.Close()
SET PIR_HELPDESK1 = NOTHING
%>
            ];
          </script>

          <div class="card form-card shadow">
            <div class="panel-header">
              Desativar/Ativar Usuário SWNET
            </div>

            <div class="card-body pt-4 pb-3 px-4">

              <div class="mb-4">
                <label class="form-label" for="searchUsuario">
                  <i class="bi bi-person me-2"></i>Usuário
                </label>
                <div class="user-search-wrapper" id="searchWrapper">
                  <div class="input-group">
                    <span class="input-group-text bg-white border-end-0">
                      <i class="bi bi-search text-muted"></i>
                    </span>
                    <input
                      type="text"
                      id="searchUsuario"
                      class="form-control border-start-0 border-end-0"
                      placeholder="Digite o nome ou login do usuário..."
                      autocomplete="off"
                      aria-label="Buscar usuário"
                      aria-autocomplete="list"
                      aria-controls="dropdownUsuario"
                      aria-expanded="false"
                    />
                    <span class="input-group-text bg-white border-start-0">
                      <i class="bi bi-chevron-down text-muted" id="searchChevron" style="font-size:0.7rem;transition:transform .2s;"></i>
                    </span>
                  </div>
                  <div class="user-dropdown" id="dropdownUsuario" role="listbox"></div>
                </div>
              </div>

              <div class="info-placeholder" id="divPlaceholder">
                <i class="bi bi-person-circle"></i>
                <p>Selecione um usuário acima para visualizar os dados e gerenciar a conta.</p>
              </div>

              <div id="divInfos">

                <div class="card border mb-3">
                  <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                    <i class="bi bi-info-circle text-secondary me-2"></i>
                    <span class="fw-semibold card-header-label">Dados Cadastrais</span>
                    <span class="badge-status ms-auto" id="statusBadge">—</span>
                  </div>
                  <div class="card-body px-3 py-3">
                    <div class="row g-3 mb-3">
                      <div class="col-md-2">
                        <label class="form-label" for="Matricula">Matrícula</label>
                        <input id="Matricula" name="Matricula" type="text" class="form-control form-control-sm" maxlength="6" readonly />
                      </div>
                      <div class="col-md-5">
                        <label class="form-label" for="Nome">Nome</label>
                        <input id="Nome" name="Nome" type="text" class="form-control form-control-sm" disabled />
                      </div>
                      <div class="col-md-5">
                        <label class="form-label" for="Email">E-mail</label>
                        <input id="Email" name="Email" type="email" class="form-control form-control-sm" disabled />
                      </div>
                    </div>
                    <div class="row g-3">
                      <div class="col-md-3">
                        <label class="form-label" for="Nascimento">Nascimento</label>
                        <input id="Nascimento" name="Nascimento" type="date" class="form-control form-control-sm" disabled maxlength="10" />
                      </div>
                      <div class="col-md-3">
                        <label class="form-label" for="UltimoAcesso">Último Acesso</label>
                        <input id="UltimoAcesso" name="UltimoAcesso" type="text" class="form-control form-control-sm" disabled />
                      </div>
                      <div class="col-md-6">
                        <label class="form-label">Sexo</label>
                        <div class="d-flex gap-4 pt-1">
                          <div class="form-check">
                            <input class="form-check-input" type="radio" name="radios" id="sexoM" value="M" checked disabled />
                            <label class="form-check-label small" for="sexoM">Masculino</label>
                          </div>
                          <div class="form-check">
                            <input class="form-check-input" type="radio" name="radios" id="sexoF" value="F" disabled />
                            <label class="form-check-label small" for="sexoF">Feminino</label>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div class="card border mb-3">
                  <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                    <i class="bi bi-shield-lock text-secondary me-2"></i>
                    <span class="fw-semibold card-header-label">Dados de Acesso</span>
                  </div>
                  <div class="card-body px-3 py-3">
                    <div class="row g-3">
                      <div class="col-md-6">
                        <label class="form-label" for="Usuario">Usuário</label>
                        <input id="Usuario" name="Usuario" type="text" class="form-control form-control-sm" disabled />
                      </div>
                      <div class="col-md-6">
                        <label class="form-label" for="UserAD">Usuário AD</label>
                        <input id="UserAD" name="UserAD" type="text" class="form-control form-control-sm" disabled />
                      </div>
                    </div>
                  </div>
                </div>

                <div class="card border mb-3">
                  <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                    <i class="bi bi-building text-secondary me-2"></i>
                    <span class="fw-semibold card-header-label">Vínculo</span>
                  </div>
                  <div class="card-body px-3 py-3">

                    <div class="row g-3">
                  <div class="col-md-4">
                    <label class="form-label" for="Tipo">Tipo</label>
<%
SQL = "" &_
  "SELECT " &_
    "Id, " &_
    "Tipo " &_
  "FROM TipoUsuario "

SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)

If Not IsObject(PIR_HELPDESK1) Or (PIR_HELPDESK1 IS NOTHING) Then
  Response.End
End If
%>
                    <select id="Tipo" name="Tipo" class="form-select form-select-sm" disabled>
                      <option value="">Selecione...</option>
<%
While Not PIR_HELPDESK1.Eof
%>
                      <option value="<%=Trim(PIR_HELPDESK1("Id"))%>"><%=Trim(PIR_HELPDESK1("Tipo"))%></option>
<%
  PIR_HELPDESK1.MoveNext()
Wend
PIR_HELPDESK1.Close()
SET PIR_HELPDESK1 = NOTHING
%>
                    </select>
                  </div>

                  <div class="col-md-5">
                    <label class="form-label" for="Filial">Empresa</label>
<%
sFilial = Right("0" & Trim(CStr(Filial & "")), 2)

If Nivel = 0 Or Nivel > 5 Then
  sWhereEmpresa = ""
ElseIf Nivel = 4 And (sFilial = "05" Or sFilial = "06") Then
  sWhereEmpresa = "WHERE Empresa IN ('05','06') "
ElseIf Nivel = 3 Then
  sWhereEmpresa = "WHERE Empresa = '" & sFilial & "' "
Else
  sWhereEmpresa = "WHERE Empresa IN ('00','01','02','04') "
End If

SQL = "" &_
  "SELECT " &_
    "Empresa, " &_
    "Razao " &_
  "FROM Empresas " &_
  sWhereEmpresa &_
  "ORDER BY Empresa"

SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)

If Not IsObject(PIR_HELPDESK1) Or (PIR_HELPDESK1 IS NOTHING) Then
  Response.End
End If
%>
                    <select id="Filial" name="Filial" class="form-select form-select-sm" disabled>
                      <option value="">Selecione...</option>
<%
While Not PIR_HELPDESK1.Eof
%>
                      <option value="<%=Trim(PIR_HELPDESK1("Empresa"))%>"><%=Trim(PIR_HELPDESK1("Razao"))%></option>
<%
  PIR_HELPDESK1.MoveNext()
Wend
PIR_HELPDESK1.Close()
SET PIR_HELPDESK1 = NOTHING
%>
                    </select>
                  </div>

                  <div class="col-md-3">
                    <label class="form-label" for="Ramal">Ramal</label>
                    <input id="Ramal" name="Ramal" type="text" class="form-control form-control-sm" disabled />
                  </div>
                    </div>
                  </div>
                </div>

                <div class="d-flex justify-content-center gap-2 mt-3">
                  <button
                    type="button"
                    class="btn btn-sm btn-danger px-4"
                    id="btnAcao"
                    onclick="abrirModalConfirma()">
                    <span id="btnAcaoText">Desativar</span>
                  </button>
                  <button
                    type="button"
                    class="btn btn-sm btn-danger"
                    id="btnExportarPDF"
                    title="Exportar relatório em PDF"
                    aria-label="Exportar PDF">
                    <i class="bi bi-file-earmark-pdf"></i>
                  </button>
                </div>

              </div>

            </div>
          </div>

        </form>

      </div>

    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>
    <script src="../funcoes/funcoesJAVASCRIPT.js"></script>

    <script>
      (($) => {
        "use strict";

        const TOAST_DELAY    = 4500;
        const LIMITE_DROPDOWN = 50;
        const LOGO_PATH      = "../../corporativo/Orcamento/Relatorio/Logo/logo-smurfit1.png";
        const STATUS_ATIVO   = "ATIVO";
        const STATUS_INATIVO = "INATIVO";
        const PARAM_DESATIVAR = "D";
        const PARAM_ATIVAR    = "A";

        const TOAST_ICON = Object.freeze({
          success: { cls: "bi-check-circle-fill",         cor: "#198754" },
          warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
          danger:  { cls: "bi-x-circle-fill",             cor: "#dc3545" },
          info:    { cls: "bi-info-circle-fill",           cor: "#6c757d" },
        });

        const PDF_TABELA_ESTILO = Object.freeze({
          theme: "grid",
          styles: {
            fontSize: 9.5,
            cellPadding: { top: 3, bottom: 3, left: 4, right: 4 },
            valign: "middle",
          },
          columnStyles: {
            0: { cellWidth: 45, fontStyle: "bold", fillColor: [248, 248, 248], textColor: [74, 74, 74] },
            1: { cellWidth: "auto", textColor: [30, 30, 30] },
          },
        });

        const CFG_STATUS = Object.freeze({
          ATIVO:   { param: "D", btnClass: "btn-danger",  label: "Desativar", badgeLabel: "ATIVO",      badgeClass: "badge-status-ativo",   modalTitulo: "Confirmar Desativação", modalTexto: "O usuário será desativado:", confirmaClass: "btn-danger",  confirmaLabel: "Desativar" },
          INATIVO: { param: "A", btnClass: "btn-success", label: "Ativar",    badgeLabel: "Desativado", badgeClass: "badge-status-inativo", modalTitulo: "Confirmar Ativação",     modalTexto: "O usuário será ativado:",    confirmaClass: "btn-success", confirmaLabel: "Ativar"    },
        });

        const state = {
          debounceTimer: null,
          activeIndex: -1,
        };

        $(document).ready(() => {
          _exibirMensagensIniciais();
          _bindEvents();
        });

        function _exibirMensagensIniciais() {
          if (typeof msgSucesso !== "undefined" && msgSucesso) {
            mostrarNotificacao(msgSucesso, "success");

            // Geração automática de PDF após ativação de usuário
            const pdfPending = sessionStorage.getItem("excusuario_pdf_pending");
            if (pdfPending === "1") {
              sessionStorage.removeItem("excusuario_pdf_pending");
              sessionStorage.removeItem("excusuario_pdf_recno");
              sessionStorage.removeItem("excusuario_pdf_login");
              sessionStorage.removeItem("excusuario_pdf_nome");
              // Após ativar, retorna à tela inicial (campo de busca limpo, sem dados pré-carregados)
              // O toast de sucesso já foi exibido acima; nenhuma ação adicional aqui.
            }
          }
          if (typeof msgErro !== "undefined" && msgErro) mostrarNotificacao(msgErro, "danger");
        }

        function _bindEvents() {
          const $search = $("#searchUsuario");

          $search.on("input",   _onSearchInput);
          $search.on("keydown", _onSearchKeydown);
          $search.on("focus", function () { filtrarUsuarios($search.val().trim()); });

          $(document).on("click", function (e) {
            if (!$(e.target).closest(".user-search-wrapper").length) _fecharDropdown();
          });

          $("#btnExportarPDF").on("click", exportarPDF);
          $("#btnConfirmaAcao").on("click", function () {
            bootstrap.Modal.getInstance(document.getElementById("modalConfirma")).hide();
            const param = $("#btnParam").val();
          
            if (param === PARAM_ATIVAR) {
              // Gera o PDF com os dados já visíveis na tela, ANTES de submeter
              exportarPDF();
              // Pequeno delay para garantir que o PDF foi iniciado antes do navigate
              setTimeout(function () {
                document.getElementById("ExcluirUsuario").submit();
              }, 300);
            } else {
              sessionStorage.removeItem("excusuario_pdf_pending");
              document.getElementById("ExcluirUsuario").submit();
            }
          });
        }

        function _onSearchInput() {
          const termo = $(this).val().trim();
          state.activeIndex = -1;
          clearTimeout(state.debounceTimer);
          state.debounceTimer = setTimeout(() => filtrarUsuarios(termo), 300);
        }

        function _onSearchKeydown(e) {
          const $items = $(".user-dropdown-item:visible");
          if (!$items.length) return;
          if (e.key === "ArrowDown") {
            e.preventDefault();
            state.activeIndex = Math.min(state.activeIndex + 1, $items.length - 1);
            _atualizarItemAtivo($items);
          } else if (e.key === "ArrowUp") {
            e.preventDefault();
            state.activeIndex = Math.max(state.activeIndex - 1, 0);
            _atualizarItemAtivo($items);
          } else if (e.key === "Enter") {
            e.preventDefault();
            if (state.activeIndex >= 0) $items.eq(state.activeIndex).trigger("click");
          } else if (e.key === "Escape") {
            _fecharDropdown();
          }
        }

        function _fecharDropdown() {
          $("#dropdownUsuario").removeClass("show");
          $("#searchChevron").css("transform", "rotate(0deg)");
          state.activeIndex = -1;
        }

        function _atualizarItemAtivo($items) {
          $items.removeClass("active");
          if (state.activeIndex < 0) return;
          const $ativo = $items.eq(state.activeIndex).addClass("active");
          const container = document.getElementById("dropdownUsuario");
          const item = $ativo[0];
          if (item.offsetTop < container.scrollTop) {
            container.scrollTop = item.offsetTop;
          } else if (item.offsetTop + item.offsetHeight > container.scrollTop + container.clientHeight) {
            container.scrollTop = item.offsetTop + item.offsetHeight - container.clientHeight;
          }
        }

        function filtrarUsuarios(termo) {
          const $dropdown  = $("#dropdownUsuario");
          const termoLower = termo.toLowerCase();
          const filtrados  = listaUsuarios.filter(u =>
            !termo ||
            u.nome.toLowerCase().includes(termoLower) ||
            u.login.toLowerCase().includes(termoLower)
          );

          $dropdown.empty();

          if (!filtrados.length) {
            $dropdown
              .html('<div class="p-3 text-center text-muted small"><i class="bi bi-emoji-frown me-1"></i>Nenhum usuário encontrado</div>')
              .addClass("show");
            return;
          }

          const exibidos = filtrados.slice(0, LIMITE_DROPDOWN);
          const limitado = filtrados.length > LIMITE_DROPDOWN;

          exibidos.forEach(usr => {
            const statusAtivo = usr.status === STATUS_ATIVO;
            const statusCls   = statusAtivo ? "text-success" : "text-danger";
            const statusIcon  = statusAtivo ? "bi-check-circle-fill" : "bi-x-circle-fill";

            const $item = $(
              `<div class="user-dropdown-item" role="option">
                <div class="d-flex justify-content-between align-items-center">
                  <div>
                    <div class="fw-semibold small">${usr.nome}</div>
                    <div class="text-muted user-dropdown-login"><i class="bi bi-person-badge me-1"></i>${usr.login}</div>
                  </div>
                  <span class="${statusCls} small fw-semibold">
                    <i class="bi ${statusIcon}"></i> ${usr.status}
                  </span>
                </div>
              </div>`
            ).on("click", function () {
              $("#searchUsuario").val(usr.nome);
              $("#ListaUsuario").val(usr.id);
              _fecharDropdown();
              SelecionouUsuario(usr.id, usr.status);
            });

            $dropdown.append($item);
          });

          if (limitado) {
            $dropdown.append(
              `<div class="user-dropdown-counter">Mostrando ${LIMITE_DROPDOWN} de ${filtrados.length}</div>`
            );
          }

          $dropdown.addClass("show");
          $("#searchChevron").css("transform", "rotate(180deg)");
        }

        function abrirModalConfirma() {
          const param = $("#btnParam").val();
          const nome  = $("#Nome").val() || "este usuário";
          const cfg   = param === PARAM_DESATIVAR ? CFG_STATUS.ATIVO : CFG_STATUS.INATIVO;

          $("#modalConfirmaTitulo").text(cfg.modalTitulo);
          $("#modalConfirmaNomeUsuario").text(nome);
          $("#modalConfirmaAcaoTexto").text(cfg.modalTexto);
          $("#btnConfirmaAcao").removeClass("btn-danger btn-success").addClass(cfg.confirmaClass);
          $("#btnConfirmaTexto").text(cfg.confirmaLabel);

          new bootstrap.Modal(document.getElementById("modalConfirma")).show();
        }

        function limparCampos() {
          $("#Matricula, #Nome, #Email, #Nascimento, #Usuario, #UserAD, #Ramal, #UltimoAcesso").val("");
          $("#sexoM, #sexoF").prop("checked", false);
          $("#Tipo, #Filial").prop("selectedIndex", 0);
          $("#statusBadge").text("—").removeClass("badge-status-ativo badge-status-inativo");
          $("#divInfos").hide();
          $("#divPlaceholder").show();
          $("#searchChevron").css("transform", "rotate(0deg)");
        }

        function _parseDateInput(val) {
          const s = (val || "").toString().trim().replace(/\D/g, "");
          if (s.length === 8) return s.slice(0, 4) + "-" + s.slice(4, 6) + "-" + s.slice(6, 8);
          return "";
        }

        function _aplicarStatus(status) {
          const key    = (status || "").toString().trim().toUpperCase();
          const cfg    = CFG_STATUS[key] || {};
          const $badge = $("#statusBadge");
          $badge.text(cfg.badgeLabel || key || "—").removeClass("badge-status-ativo badge-status-inativo");
          if (cfg.badgeClass) $badge.addClass(cfg.badgeClass);
          $("#btnParam").val(cfg.param || "");
          $("#btnAcao").removeClass("btn-success btn-danger").addClass(cfg.btnClass || "btn-danger");
          $("#btnAcaoText").text(cfg.label || "Desativar");
        }

        function SelecionouUsuario(id, status) {
          if (id) {
            if (status) _aplicarStatus(status);
            ObjAjax("ajax/Usuario.asp?id=" + id);
          } else {
            limparCampos();
          }
        }

        function objReturn(obj) {
          if (obj.erro === "1") {
            mostrarNotificacao("Não foi possível realizar a operação. Tente novamente mais tarde.", "danger");
            return;
          }

          $("#Matricula").val(obj.val1);
          $("#Nome").val(obj.val3);
          $("#Email").val(obj.val4);
          $("#Nascimento").val(_parseDateInput(obj.val5));
          $("#UserAD").val(obj.val14);
          $("#Usuario").val(obj.val6);
          $("#FilialCargas").val(obj.val9);
          $("#Ramal").val(obj.val10);

          const sexo = (obj.val2 || "").toUpperCase();
          $("#sexoM").prop("checked", sexo === "M");
          $("#sexoF").prop("checked", sexo === "F");

          const tipoAlvo   = parseInt(obj.val8, 10);
          const filialAlvo = obj.val9;
          $("#Tipo option").each(function () { $(this).prop("selected", parseInt($(this).val(), 10) === tipoAlvo); });
          $("#Filial option").each(function () { $(this).prop("selected", $(this).val() === filialAlvo); });

          $("#UltimoAcesso").val(obj.val13);

          $("#divPlaceholder").hide();
          $(".form-card").addClass("has-data");
          $("#divInfos").fadeIn(200);

          // Geração automática de PDF após redirect de ativação (ERRO 1 fix)
          if (typeof window._excusuario_pdfLoginPendente !== "undefined") {
            const loginAlvoPendente = window._excusuario_pdfLoginPendente;
            delete window._excusuario_pdfLoginPendente;
            setTimeout(function () { exportarPDF(loginAlvoPendente); }, 100);
          }
        }

        function exportarPDF(loginForcado) {
          const recno = $("#ListaUsuario").val();
          if (!recno) {
            mostrarNotificacao("Selecione um usuário antes de exportar o PDF.", "warning");
            return;
          }

          const loginDOM = $("#Usuario").val() || "";
          const loginStr = (typeof loginForcado === "string" && loginForcado.trim() !== "") ? loginForcado.trim() : loginDOM;

          // Se a ação em curso é ATIVAR, o PDF deve refletir o status resultante (ATIVO),
          // não o status atual do badge (que ainda mostra "Desativado" antes do submit).
          const paramAtual  = $("#btnParam").val();
          const statusFinal = paramAtual === PARAM_ATIVAR ? STATUS_ATIVO : ($("#statusBadge").text() || "—");

          const usr = {
            id:           recno,
            nome:         $("#Nome").val()                         || "—",
            login:        loginStr                                 || "—",
            matricula:    $("#Matricula").val()                    || "—",
            email:        $("#Email").val()                        || "—",
            nascimento:   $("#Nascimento").val()                   || "—",
            sexo:         $("input[name='radios']:checked").val()  || "—",
            ramal:        $("#Ramal").val()                        || "—",
            tipo:         $("#Tipo option:selected").text()        || "—",
            filial:       $("#Filial option:selected").text()      || "—",
            status:       statusFinal,
            userAD:       $("#UserAD").val()                       || "—",
            ultimoAcesso: $("#UltimoAcesso").val()                 || "—",
          };

          mostrarLoading(true);

          const logoImg = new Image();
          logoImg.crossOrigin = "Anonymous";
          let pdfGerado = false;

          function finalizar(logoBase64) {
            if (pdfGerado) return;
            pdfGerado = true;
            _gerarPDFCompleto(logoBase64, usr);
          }

          logoImg.onload = function () {
            const canvas  = document.createElement("canvas");
            canvas.width  = logoImg.width;
            canvas.height = logoImg.height;
            const ctx = canvas.getContext("2d");
            ctx.fillStyle = "#FFFFFF";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(logoImg, 0, 0);
            finalizar(canvas.toDataURL("image/png"));
          };

          logoImg.onerror = () => finalizar(null);
          logoImg.src = LOGO_PATH;
          setTimeout(() => finalizar(null), 2000);
        }

        function _gerarPDFCompleto(logoBase64, usr) {
          const { jsPDF } = window.jspdf;
          const doc      = new jsPDF({ orientation: "p", unit: "mm", format: "a4", compress: true });
          const login    = LoginUsuarioLogado || "Sistema";
          const nomeOp   = NomeUsuarioLogado  || "";
          const operador = nomeOp ? `${login} — ${nomeOp}` : login;
          const dataEmissao = new Date().toLocaleString("pt-BR", {
            day: "2-digit", month: "2-digit", year: "numeric",
            hour: "2-digit", minute: "2-digit", second: "2-digit",
          });

          const pageW   = doc.internal.pageSize.getWidth();
          const pageH   = doc.internal.pageSize.getHeight();
          const margin  = 14;
          const headerH = 34;

          doc.setFillColor(255, 255, 255);
          doc.setDrawColor(200, 200, 200);
          doc.setLineWidth(0.5);
          doc.rect(0, 0, pageW, headerH, "FD");

          if (logoBase64) {
            try { doc.addImage(logoBase64, "PNG", margin, 6, 36, 16); } catch (e) {}
          }

          doc.setFont("helvetica", "bold");
          doc.setFontSize(14);
          doc.setTextColor(0, 0, 0);
          doc.text("Relatório do Usuário", pageW / 2, 14, { align: "center" });

          doc.setFont("helvetica", "normal");
          doc.setFontSize(8.5);
          doc.setTextColor(90, 90, 90);
          doc.text(`Emitido por: ${login}`, pageW - margin, 14, { align: "right" });
          doc.text(`Sistema: SWNET  |  Emitido em: ${dataEmissao}`, pageW / 2, 27, { align: "center" });

          const statusAtivo = usr.status === STATUS_ATIVO;
          const cores = statusAtivo
            ? { fundo: [209, 231, 221], texto: [10, 54, 34] }
            : { fundo: [248, 215, 218], texto: [88, 21, 28] };

          doc.setFillColor(...cores.fundo);
          doc.roundedRect(margin, headerH + 5, pageW - margin * 2, 11, 2, 2, "F");
          doc.setFont("helvetica", "bold");
          doc.setFontSize(10);
          doc.setTextColor(...cores.texto);
          doc.text(`${usr.nome}  ·  ${usr.status}`, pageW / 2, headerH + 12, { align: "center" });

          function renderSecao(titulo, yPos) {
            doc.setFillColor(212, 212, 212);
            doc.rect(margin, yPos, pageW - margin * 2, 7, "F");
            doc.setFont("helvetica", "bold");
            doc.setFontSize(8);
            doc.setTextColor(74, 74, 74);
            doc.text(titulo.toUpperCase(), margin + 3, yPos + 5);
            return yPos + 7;
          }

          function renderTabela(dados, yPos, extraColumnStyles) {
            const colStyles = Object.assign({}, PDF_TABELA_ESTILO.columnStyles, extraColumnStyles || {});
            doc.autoTable({
              theme: PDF_TABELA_ESTILO.theme,
              styles: PDF_TABELA_ESTILO.styles,
              columnStyles: colStyles,
              body: dados,
              startY: yPos,
              margin: { left: margin, right: margin },
              tableWidth: pageW - margin * 2,
            });
            return doc.lastAutoTable.finalY;
          }

          let y = headerH + 22;
          y = renderSecao("Dados Pessoais", y);
          y = renderTabela([
            ["Matrícula",     usr.matricula],
            ["Nome completo", usr.nome],
            ["E-mail",        usr.email],
            ["Nascimento",    (function(v){ const p=v.split("-"); return p.length===3?p[2]+"/"+p[1]+"/"+p[0]:v; })(usr.nascimento)],
            ["Sexo",          usr.sexo === "M" ? "Masculino" : usr.sexo === "F" ? "Feminino" : "—"],
          ], y) + 4;

          y = renderSecao("Dados de Acesso", y);
          y = renderTabela([
            ["Login",         usr.login],
            ["Usuário AD",    usr.userAD],
            ["Ramal",         usr.ramal],
            ["Último Acesso", usr.ultimoAcesso],
          ], y) + 4;

          y = renderSecao("Vínculo", y);
          y = renderTabela([
            ["Tipo",    usr.tipo],
            ["Empresa", usr.filial],
          ], y) + 4;

          y = renderSecao("Status da Conta", y);
          renderTabela([["Status", usr.status]], y, {
            1: { cellWidth: "auto", fontStyle: "bold", textColor: cores.texto },
          });

          // Nome do arquivo: <nome_sanitizado>_<ddMMYYYY_HHmm>.pdf
          // Regex: remove acentos, substitui espaços por hífens, remove caracteres não alfanuméricos/hífen
          const nomeBase = String(usr.nome !== "—" ? usr.nome : usr.login);
          const loginSanitizado = nomeBase
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // remove acentos
            .replace(/\s+/g, "-")                                // espaços → hífens
            .replace(/[^a-zA-Z0-9\-]/g, "")                     // remove não-alfanuméricos (exceto hífen)
            .replace(/-{2,}/g, "-")                              // hífens duplos → simples
            .replace(/^-|-$/g, "")                               // remove hífens nas bordas
            .toLowerCase()
            .slice(0, 60);                                       // limita comprimento

          const d   = new Date();
          const pad = n => n < 10 ? "0" + n : String(n);
          const dataNome = `${pad(d.getDate())}${pad(d.getMonth() + 1)}${d.getFullYear()}_${pad(d.getHours())}${pad(d.getMinutes())}`;
          const blob = doc.output("blob");
          const url  = URL.createObjectURL(blob);
          const $a   = $("<a>")
            .attr("href", url)
            .attr("download", `${loginSanitizado}_${dataNome}.pdf`);
          $("body").append($a);
          $a[0].click();
          $a.remove();
          URL.revokeObjectURL(url);

          mostrarLoading(false);
          mostrarNotificacao("PDF gerado com sucesso!", "success");
        }

        function mostrarLoading(show) {
          const $btn = $("#btnExportarPDF");
          if (show) {
            $btn
              .prop("disabled", true)
              .data("html-original", $btn.html())
              .html('<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>');
          } else {
            const original = $btn.data("html-original");
            $btn
              .prop("disabled", false)
              .html(original || '<i class="bi bi-file-earmark-pdf"></i>');
          }
        }

        function mostrarNotificacao(msg, tipo = "info") {
          const cfg       = TOAST_ICON[tipo] ?? TOAST_ICON.info;
          const tipoClass = tipo === "error" ? "danger" : tipo;
          const item = $(`
            <div class="toast-item toast-item-${tipoClass}">
              <div class="toast-corpo">
                <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
                <span class="flex-grow-1">${msg}</span>
                <button class="toast-fechar" aria-label="Fechar">&times;</button>
              </div>
              <div class="toast-barra toast-barra-${tipoClass}" style="animation-duration:${TOAST_DELAY}ms;"></div>
            </div>
          `);
          item.find(".toast-fechar").on("click", () => item.remove());
          $("#toastWrapper").append(item);
          setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_DELAY);
        }

        Object.assign(window, {
          SelecionouUsuario,
          objReturn,
          abrirModalConfirma,
        });

      })(jQuery);
    </script>

  </body>
</html>
