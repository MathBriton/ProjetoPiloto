<%
Session.LCID = 1046
%>
<!--#include file="../cnn2/cnn2.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../includes/variaveis.asp" -->
<%
Const CONN_LOG = "PIR-CORPORATIVO"
Sub ProcessarAlteracao()

  Dim sBloqueado, sRecno
  Dim sNome, sSexo, sNasc, sUserAD, sLogin, sSenha
  Dim sEmail, sRamal, sTrocarSenha, sTipo, sFilial

  sRecno = Trim(Request("ListaUsuario"))

  If sRecno = "" Then Exit Sub

  If Trim(Request("bloqueado")) = "1" Then
      sBloqueado = "1"
  Else
      sBloqueado = "0"
  End If

  sNome        = Left(UCase(Trim(Request("Nome"))), 60)
  sSexo        = Left(Trim(Request("sexo")), 1)
  sNasc        = Left(Replace(Trim(Request("Nascimento")), "-", ""), 8)
  sUserAD      = Left(Trim(Request("UserAD")), 255)
  sLogin       = Left(Trim(Request("Usuario")), 30)
  sSenha       = Left(Trim(Request("Senha")), 50)
  sEmail       = Left(LCase(Trim(Request("Email"))), 255)
  sRamal       = Left(Trim(Request("Ramal")), 4)
  If Trim(Request("Senha")) <> "" Then
    sTrocarSenha = "S"
  Else
    sTrocarSenha = "N"
  End If
  sTipo        = Left(Trim(Request("Tipo")), 1)
  sFilial      = Left(Trim(Request("Filial")), 6)

  Dim oConn
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.Open BuscaStringConexao("PIR-HELPDESK")

  Dim oCmd
  Set oCmd = Server.CreateObject("ADODB.Command")

  Set oCmd.ActiveConnection = oConn
  oCmd.CommandType = 1

  Dim sSenhaSet
  If sSenha <> "" Then
    sSenhaSet = _
      "CRM_SENHA = ?, " & _
      "Senha = HASHBYTES('SHA1', ?), " & _
      "TROCARSENHA = ?, " & _
      "CRM_DATEXP = CONVERT(varchar(8), DATEADD(day, 60, GETDATE()), 112), "
  ElseIf sBloqueado = "1" Then
    ' Usuário bloqueado manualmente: garantir que TROCARSENHA seja limpo
    ' para não redirecionar para troca de senha no próximo login
    sSenhaSet = "TROCARSENHA = 'N', "
  Else
    sSenhaSet = ""
  End If

  oCmd.CommandText = _
    "BEGIN TRANSACTION " & _
      "UPDATE CRM_DTI SET " & _
        "CRM_NOME = ?, " & _
        "RA_SEXO = ?, " & _
        "RA_NASC = ?, " & _
        "CRM_USERAD = ?, " & _
        "CRM_LOGIN = ?, " & _
        sSenhaSet & _
        "CRM_EMAILPRIMARIO = ?, " & _
        "CRM_RAMAL = ?, " & _
        "Bloqueado = ?, " & _
        "Desbloqueio = CASE WHEN ? = 1 THEN NULL ELSE Desbloqueio END, " & _
        "CRM_TIPO = ?, " & _
        "EMP_CODIGO = ? " & _
      "WHERE CRM_R_E_C_N_O = ? " & _
    "IF @@ERROR <> 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_NOME", 200, 1, 60, sNome)
  oCmd.Parameters.Append oCmd.CreateParameter("@RA_SEXO", 200, 1, 1, sSexo)
  oCmd.Parameters.Append oCmd.CreateParameter("@RA_NASC", 200, 1, 8, sNasc)
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_USERAD", 200, 1, 255, sUserAD)
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_LOGIN", 200, 1, 30, sLogin)
  If sSenha <> "" Then
    oCmd.Parameters.Append oCmd.CreateParameter("@CRM_SENHA", 200, 1, 50, sSenha)
    oCmd.Parameters.Append oCmd.CreateParameter("@SenhaHash", 200, 1, 50, sSenha)
    oCmd.Parameters.Append oCmd.CreateParameter("@TROCARSENHA", 200, 1, 1, sTrocarSenha)
  End If
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_EMAIL", 200, 1, 255, sEmail)
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_RAMAL", 200, 1, 4, sRamal)
  oCmd.Parameters.Append oCmd.CreateParameter("@Bloqueado", 11, 1, , CBool(CLng(sBloqueado)))
  oCmd.Parameters.Append oCmd.CreateParameter("@Bloqueado2", 3, 1, , CLng(sBloqueado))
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_TIPO", 200, 1, 1, sTipo)
  oCmd.Parameters.Append oCmd.CreateParameter("@EMP_CODIGO", 200, 1, 6, sFilial)
  oCmd.Parameters.Append oCmd.CreateParameter("@CRM_R_E_C_N_O", 3, 1, 4, CLng(sRecno))

  oCmd.Execute

  Set oCmd = Nothing
  oConn.Close
  Set oConn = Nothing

  Dim oConnLog
  Set oConnLog = Server.CreateObject("ADODB.Connection")
  oConnLog.Open BuscaStringConexao(CONN_LOG)

  Dim oCmdLog
  Set oCmdLog = Server.CreateObject("ADODB.Command")

  Set oCmdLog.ActiveConnection = oConnLog
  oCmdLog.CommandType = 1

  oCmdLog.CommandText = _
    "INSERT INTO dbo.LogAtividade " & _
      "(IdUsuarioAcao, IdUsuarioAlvo, Sistema, Rotina, Acao, Descricao, DataCriacao) " & _
    "VALUES (?, ?, 'SWnet', 'AltUsuario', 'ALTERAR', ?, GETDATE())"

  oCmdLog.Parameters.Append oCmdLog.CreateParameter("@IdUsuarioAcao", 200, 1, 50, Session("USR"))
  oCmdLog.Parameters.Append oCmdLog.CreateParameter("@IdUsuarioAlvo", 200, 1, 20, sRecno)

  oCmdLog.Parameters.Append oCmdLog.CreateParameter( _
    "@Descricao", 200, 1, 500, _
    "Usuario RECNO " & sRecno & _
    " alterado pelo operador " & Session("USR") & _
    " (IP: " & Request.ServerVariables("REMOTE_ADDR") & ")" _
  )

  oCmdLog.Execute

  Set oCmdLog = Nothing
  oConnLog.Close
  Set oConnLog = Nothing

  If sSenha <> "" Then
    Dim Para, Assunto, Mensagem

    Para    = sEmail
    Assunto = "Acesso liberado ao sistema SWnet - Smurfit Westrock Brasil"

    Mensagem = _
      "<!doctype html><html><head><meta charset='utf-8'>" & _
      "<style>" & _
        "body{font-family:Arial,sans-serif;font-size:15px;color:#333}" & _
        "p{margin-bottom:10px}" & _
        ".cred{background:#f5f5f5;border-left:4px solid #4a4a4a;padding:10px 16px;border-radius:4px}" & _
      "</style></head><body>" & _
      "<p>Parab&eacute;ns, seu acesso foi liberado ao sistema <strong>SWnet - Smurfit Westrock Brasil</strong>.</p>" & _
      "<p>Recomendamos fortemente que, ao acessar, troque sua senha imediatamente.</p>" & _
      "<div class='cred'>" & _
        "<p><strong>Usu&aacute;rio:</strong> " & sLogin & "</p>" & _
        "<p><strong>Senha:</strong> " & sSenha & "</p>" & _
      "</div>" & _
      "<p>Para usu&aacute;rios LSK: http://br-pir1-m010/</p>" & _
      "<p>Para usu&aacute;rios LWR: http://br-pir1-m010.group.wan/</p>" & _
      "<p>Seja bem-vindo ao nosso sistema.<br>At.te,<br>" & _
      "<strong>Equipe de Desenvolvimento &mdash; Smurfit Westrock Brasil</strong></p>" & _
      "</body></html>"

    Call EnviaEmail(Para, "", "", "", "noreply.brasil@smurfitwestrock.com.br", Mensagem, Assunto)

    Response.Write "<script>var msgSucesso='Usuário alterado com sucesso! E-mail enviado.';</script>"
  Else
    Response.Write "<script>var msgSucesso='Usuário alterado com sucesso!';</script>"
  End If

End Sub

If Trim(Request("ListaUsuario")) <> "" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  Call ProcessarAlteracao()
End If
%>

<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />
    <title>SWnet - Alterar Usuário</title>

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

      .form-label {
        font-weight: 600;
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

      .form-control.is-valid,
      .form-control.is-invalid {
        background-image: none;
        padding-right: 0.75rem;
      }

      .toggle-bloqueado {
        position: relative;
        display: inline-flex;
        align-items: center;
        gap: 0.6rem;
        cursor: pointer;
        user-select: none;
      }

      .toggle-bloqueado input[type="checkbox"] {
        position: absolute;
        opacity: 0;
        width: 0;
        height: 0;
      }

      .toggle-track {
        width: 2.75rem;
        height: 1.5rem;
        background: #dee2e6;
        border-radius: 9999px;
        position: relative;
        transition: background 0.2s;
        flex-shrink: 0;
      }

      .toggle-bloqueado input:checked ~ .toggle-track {
        background: #0d6efd;
      }

      .toggle-thumb {
        position: absolute;
        top: 3px;
        left: 3px;
        width: 1.125rem;
        height: 1.125rem;
        background: #fff;
        border-radius: 9999px;
        transition: transform 0.2s, box-shadow 0.2s;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
      }

      .toggle-bloqueado input:checked ~ .toggle-track .toggle-thumb {
        transform: translateX(1.25rem);
      }

      .toggle-thumb .bi {
        font-size: 0.6rem;
        color: transparent;
        transition: color 0.15s;
      }

      .toggle-bloqueado input:checked ~ .toggle-track .toggle-thumb .bi {
        color: #0d6efd;
      }

      .toggle-bloqueado:focus-within .toggle-track {
        box-shadow: 0 0 0 0.2rem rgba(13, 110, 253, 0.25);
      }

      .toggle-label {
        font-size: 0.8125rem;
        font-weight: 600;
        color: #212529;
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

      .card-header-label {
        font-size: 0.78rem;
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

      .senha-feedback {
        font-size: 0.78rem;
        margin-top: 0.25rem;
        min-height: 1.1em;
        font-weight: 500;
      }

      .senha-feedback-ok {
        color: #198754;
      }

      .senha-feedback-erro {
        color: #dc3545;
      }

      .btn-senha-toggle {
        border: 1px solid #ced4da;
        border-left: none;
        border-radius: 0 6px 6px 0;
        background-color: #fff;
        color: #6c757d;
        transition: background-color 0.15s;
      }

      .btn-senha-toggle:hover {
        background-color: #f8f9fa;
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

      .toast-item-success {
        border-left: 4px solid #198754;
      }

      .toast-item-warning {
        border-left: 4px solid #ffc107;
      }

      .toast-item-danger {
        border-left: 4px solid #dc3545;
      }

      .toast-item-info {
        border-left: 4px solid #6c757d;
      }

      @keyframes entrar-toast {
        from {
          opacity: 0;
          transform: translateX(20px);
        }
        to {
          opacity: 1;
          transform: translateX(0);
        }
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

      .toast-fechar:hover {
        color: #333;
      }

      .toast-barra {
        height: 3px;
        animation: encolher linear forwards;
      }

      .toast-barra-success {
        background: #198754;
      }

      .toast-barra-warning {
        background: #ffc107;
      }

      .toast-barra-danger {
        background: #dc3545;
      }

      .toast-barra-info {
        background: #6c757d;
      }

      @keyframes encolher {
        from {
          width: 100%;
        }
        to {
          width: 0%;
        }
      }

      .user-search-wrapper {
        position: relative;
      }

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

      .user-dropdown.show {
        display: block;
      }

      .user-dropdown-item {
        padding: 9px 14px;
        cursor: pointer;
        border-bottom: 1px solid #f5f5f5;
        transition: background 0.15s;
      }

      .user-dropdown-item:last-child {
        border-bottom: none;
      }

      .user-dropdown-item:hover,
      .user-dropdown-item.active {
        background: #f0f0f0;
      }

      .user-dropdown-login {
        font-size: 0.78rem;
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

      @keyframes spin {
        to {
          transform: rotate(360deg);
        }
      }

      .spin {
        display: inline-block;
        animation: spin 0.75s linear infinite;
      }

      .card.card-principal {
        background: #ffffff;
      }

      .card.card-principal > .card-body {
        background: #ffffff;
      }

      .card.card-principal.has-data {
        background: #f7f7f7;
      }

      .card.card-principal.has-data > .card-body {
        background: #f7f7f7;
      }

      #divInfos .card {
        background: #ffffff !important;
      }

      #divInfos .card > .card-header,
      #divInfos .card > .card-body {
        background: #ffffff !important;
      }

      #divInfos {
        display: none;
      }
    </style>
  </head>

  <body>

    <div class="toast-wrapper" id="toastWrapper"></div>

    <div class="container-fluid px-3 mt-5 mb-5" style="max-width: 54rem">

      <form name="CadastroUsuario" id="CadastroUsuario" method="post" autocomplete="off">

        <%
        SQL = "" &_
          "SELECT " &_
            "CRM_R_E_C_N_O, " &_
            "CRM_NOME, " &_
            "CRM_LOGIN " &_
          "FROM CRM_DTI " &_
          "WHERE CRM_D_E_L_E_T <> '*' " &_
          "ORDER BY CRM_NOME"

        SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)

        if not isObject(PIR_HELPDESK1) or (PIR_HELPDESK1 IS NOTHING) then
          response.end
        end if
        %>

        <input type="hidden" name="ListaUsuario" id="ListaUsuario" />
        <input type="hidden" id="hBloqueado" name="bloqueado" value="0" />

        <script>
          var listaUsuarios = [
          <%
          Dim sNomeAlt, sLoginAlt, sVirgulaAlt
          sVirgulaAlt = ""
          while not PIR_HELPDESK1.EOF
            sNomeAlt  = Replace(Replace(Trim(PIR_HELPDESK1("CRM_NOME")  & ""), "\", "\\"), """", "\""")
            sNomeAlt  = Replace(Replace(sNomeAlt,  "'", "\'"), vbCrLf, " ")
            sLoginAlt = Replace(Replace(Trim(PIR_HELPDESK1("CRM_LOGIN") & ""), "\", "\\"), """", "\""")
            sLoginAlt = Replace(Replace(sLoginAlt, "'", "\'"), vbCrLf, " ")
            Response.Write sVirgulaAlt & vbCrLf
            Response.Write "{ id: """ & PIR_HELPDESK1("CRM_R_E_C_N_O") & """, nome: """ & sNomeAlt & """, login: """ & sLoginAlt & """ }"
            sVirgulaAlt = ","
            PIR_HELPDESK1.MoveNext()
          wend
          PIR_HELPDESK1.Close()
          SET PIR_HELPDESK1 = NOTHING
          %>
          ];
        </script>

        <div class="card shadow card-principal">
          <div class="panel-header">
            Alterar Usuário
          </div>

          <div class="card-body pt-4 pb-3 px-4">

            <div class="mb-4">
              <label class="form-label" for="searchUsuario">
                <i class="bi bi-person me-1"></i>Usuário
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
                    <i class="bi bi-chevron-down text-muted" id="searchChevron" style="font-size: 0.7rem; transition: transform 0.2s"></i>
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
                  <i class="bi bi-person-vcard text-secondary"></i>
                  <span class="fw-semibold card-header-label">Dados Pessoais</span>
                </div>
                <div class="card-body px-3 py-3">
                  <div class="row g-3 mb-3">
                    <div class="col-md-2">
                      <label class="form-label" for="Matricula">Matrícula</label>
                      <input id="Matricula" name="Matricula" type="text" class="form-control form-control-sm" maxlength="6" required />
                    </div>
                    <div class="col-md-7">
                      <label class="form-label" for="Nome">Nome</label>
                      <input id="Nome" name="Nome" type="text" placeholder="Nome completo" class="form-control form-control-sm" required />
                    </div>
                    <div class="col-md-3">
                      <label class="form-label" for="Nascimento">Nascimento</label>
                      <input id="Nascimento" name="Nascimento" type="date" class="form-control form-control-sm" />
                    </div>
                  </div>
                  <div class="row g-3">
                    <div class="col-md-6">
                      <label class="form-label">Sexo</label>
                      <div class="d-flex gap-4 pt-1">
                        <div class="form-check">
                          <input class="form-check-input" type="radio" name="sexo" id="sexoM" value="M" required />
                          <label class="form-check-label small" for="sexoM">Masculino</label>
                        </div>
                        <div class="form-check">
                          <input class="form-check-input" type="radio" name="sexo" id="sexoF" value="F" />
                          <label class="form-check-label small" for="sexoF">Feminino</label>
                        </div>
                      </div>
                    </div>
                    <div class="col-md-6">
                      <label class="form-label" for="Email">E-mail</label>
                      <input id="Email" name="Email" type="email" placeholder="exemplo@smurfitwestrock.com" autocomplete="username" class="form-control form-control-sm" required />
                    </div>
                  </div>
                </div>
              </div>

              <div class="card border mb-3">
                <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                  <i class="bi bi-shield-lock text-secondary"></i>
                  <span class="fw-semibold card-header-label">Dados de Acesso</span>
                </div>
                <div class="card-body px-3 py-3">
                  <input type="hidden" id="UserAD" name="UserAD" />

                  <div class="row g-3 mb-3">
                    <div class="col-auto">
                      <label class="form-label" for="PrefixoAD">Usuário AD</label>
                      <select id="PrefixoAD" name="PrefixoAD" class="form-select form-select-sm" style="width:120px;">
                        <option value="GROUP">GROUP</option>
                        <option value="WESTROCK">WESTROCK</option>
                      </select>
                    </div>
                    <div class="col">
                      <label class="form-label" for="Usuario">Usuário</label>
                      <div class="input-group input-group-sm">
                        <span id="PrefixoDisplay" class="input-group-text fw-semibold px-2"
                              style="background:#f8f9fa;color:#4a4a4a;white-space:nowrap;"></span>
                        <input id="Usuario" name="Usuario" type="text" placeholder="Login do Usuário..."
                               class="form-control form-control-sm" required maxlength="30" />
                      </div>
                    </div>
                  </div>
                  <div class="row g-3">
                    <div class="col-md-6">
                      <label class="form-label" for="Senha">Senha</label>
                      <div class="input-group input-group-sm">
                        <input id="Senha" name="Senha" type="password" placeholder="Deixe em branco para não alterar" autocomplete="new-password" class="form-control" />
                        <button type="button" class="btn btn-senha-toggle" data-target="Senha" aria-label="Mostrar senha">
                          <i class="bi bi-eye"></i>
                        </button>
                      </div>
                      <div id="fbSenha" class="senha-feedback"></div>
                    </div>
                    <div class="col-md-6">
                      <label class="form-label" for="ConfirmaSenha">Confirmar Senha</label>
                      <div class="input-group input-group-sm">
                        <input id="ConfirmaSenha" name="ConfirmaSenha" type="password" placeholder="Confirme a nova senha" autocomplete="new-password" class="form-control" />
                        <button type="button" class="btn btn-senha-toggle" data-target="ConfirmaSenha" aria-label="Mostrar senha">
                          <i class="bi bi-eye"></i>
                        </button>
                      </div>
                      <div id="fbConfirma" class="senha-feedback"></div>
                    </div>
                  </div>
                </div>
              </div>

              <div class="card border mb-3">
                <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                  <i class="bi bi-building text-secondary"></i>
                  <span class="fw-semibold card-header-label">Vínculo</span>
                </div>
                <div class="card-body px-3 py-3">
                  <div class="row g-3">
                    <div class="col-md-4">
                      <label class="form-label" for="Tipo">Tipo</label>
                      <select id="Tipo" name="Tipo" class="form-select form-select-sm" required>
                        <option value="">Selecione...</option>
                        <%
                        SQL = "SELECT Id, Tipo FROM TipoUsuario"
                        SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)
                        if not isObject(PIR_HELPDESK1) or (PIR_HELPDESK1 IS NOTHING) then response.end
                        while Not PIR_HELPDESK1.EOF
                        %>
                        <option value="<%=Trim(PIR_HELPDESK1("Id"))%>"><%=Trim(PIR_HELPDESK1("Tipo"))%></option>
                        <%
                          PIR_HELPDESK1.MoveNext()
                        wend
                        PIR_HELPDESK1.Close()
                        SET PIR_HELPDESK1 = NOTHING
                        %>
                      </select>
                    </div>
                    <div class="col-md-5">
                      <label class="form-label" for="Filial">Empresa</label>
                      <select id="Filial" name="Filial" class="form-select form-select-sm" required>
                        <option value="">Selecione...</option>
                        <%
                        SQL = "SELECT Empresa, Razao FROM Empresas ORDER BY Empresa"
                        SET PIR_HELPDESK1 = Conecta("PIR-HELPDESK", SQL)
                        if not isObject(PIR_HELPDESK1) or (PIR_HELPDESK1 IS NOTHING) then response.end
                        while Not PIR_HELPDESK1.EOF
                        %>
                        <option value="<%=Trim(PIR_HELPDESK1("Empresa"))%>"><%=Trim(PIR_HELPDESK1("Razao"))%></option>
                        <%
                          PIR_HELPDESK1.MoveNext()
                        wend
                        PIR_HELPDESK1.Close()
                        SET PIR_HELPDESK1 = NOTHING
                        %>
                      </select>
                    </div>
                    <div class="col-md-3">
                      <label class="form-label" for="Ramal">Ramal</label>
                      <input id="Ramal" name="Ramal" type="text" placeholder="Ramal" class="form-control form-control-sm" maxlength="6" />
                    </div>
                  </div>
                </div>
              </div>

              <div class="card border mb-3">
                <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                  <span class="fw-semibold card-header-label">Status da Conta</span>
                </div>
                <div class="card-body px-3 py-3 d-flex justify-content-center">
                  <label class="toggle-bloqueado">
                    <input type="checkbox" id="Bloqueado" value="1" role="switch" />
                    <span class="toggle-track">
                      <span class="toggle-thumb">
                        <i class="bi bi-check-lg"></i>
                      </span>
                    </span>
                    <span class="toggle-label">Bloqueado</span>
                  </label>
                </div>
              </div>

              <div class="d-flex justify-content-center gap-2 mt-3">
                <button id="Salvar" name="Salvar" type="button" class="btn btn-sm btn-success px-4" disabled aria-label="Salvar alterações do usuário">
                  Salvar
                </button>
              </div>

            </div>

          </div>
        </div>

      </form>
    </div>

    <template id="tplDropdownItem">
      <div class="user-dropdown-item" role="option">
        <div class="d-flex justify-content-between align-items-center">
          <div>
            <div class="fw-semibold small" data-field="nome"></div>
            <div class="text-muted user-dropdown-login">
              <i class="bi bi-person-badge me-1"></i>
              <span data-field="login"></span>
            </div>
          </div>
        </div>
      </div>
    </template>

    <template id="tplDropdownCounter">
      <div class="user-dropdown-counter"></div>
    </template>

    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    <script src="../funcoes/funcoesJAVASCRIPT.js"></script>

    <script>
      (($) => {
        "use strict";

        const TOAST_DELAY = 4500;
        const LIMITE_DROPDOWN = 50;

        const TOAST_ICON = Object.freeze({
          success: { cls: "bi-check-circle-fill", cor: "#198754" },
          warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
          danger: { cls: "bi-x-circle-fill", cor: "#dc3545" },
          info: { cls: "bi-info-circle-fill", cor: "#6c757d" },
        });

        const state = {
          debounceTimer: null,
          activeIndex: -1,
        };

        $(document).ready(() => {
          if (typeof msgSucesso !== "undefined" && msgSucesso) {
            mostrarNotificacao(msgSucesso, "success");
          }
          _bindEvents();
        });

        function _bindEvents() {
          const $search = $("#searchUsuario");

          $search.on("input", _onSearchInput);
          $search.on("keydown", _onSearchKeydown);
          $search.on("focus", function () {
            filtrarUsuarios($search.val().trim());
          });

          $(document).on("click", function (e) {
            if (!$(e.target).closest(".user-search-wrapper").length) {
              _fecharDropdown();
            }
          });

          $(document).on("click", ".btn-senha-toggle", function () {
            _toggleSenha($(this).data("target"), this);
          });

          $("#Senha, #ConfirmaSenha").on("input", _validarSenhaInline);

          $("#Usuario").on("input", _atualizarUserAD);
          $(document).on("change input", "#PrefixoAD", _atualizarUserAD);
          setTimeout(_atualizarUserAD, 0);

          $("#Salvar").on("click", _submeterFormulario);

          $("#CadastroUsuario").on("submit", function (e) {
            if (!_validarSenhaInline()) e.preventDefault();
          });
        }

        function _onSearchInput() {
          const termo = $(this).val().trim();
          state.activeIndex = -1;
          clearTimeout(state.debounceTimer);
          state.debounceTimer = setTimeout(() => filtrarUsuarios(termo), 200);
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

        function _fecharDropdown() {
          $("#dropdownUsuario").removeClass("show");
          $("#searchChevron").css("transform", "rotate(0deg)");
          state.activeIndex = -1;
        }

        function filtrarUsuarios(termo) {
          const $dropdown = $("#dropdownUsuario");
          const termoLower = termo.toLowerCase();
          const filtrados = listaUsuarios.filter(
            (u) => !termo || u.nome.toLowerCase().includes(termoLower) || u.login.toLowerCase().includes(termoLower),
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
          const tpl = document.getElementById("tplDropdownItem");

          exibidos.forEach((usr) => {
            const clone = tpl.content.cloneNode(true);
            clone.querySelector("[data-field='nome']").textContent = usr.nome;
            clone.querySelector("[data-field='login']").textContent = usr.login;
            const $item = $(clone.querySelector(".user-dropdown-item")).on("click", function () {
              $("#searchUsuario").val(usr.nome);
              $("#ListaUsuario").val(usr.id);
              _fecharDropdown();
              SelecionouUsuario(usr.id);
            });
            $dropdown.append($item);
          });

          if (limitado) {
            const tplCounter = document.getElementById("tplDropdownCounter");
            const cloneCounter = tplCounter.content.cloneNode(true);
            cloneCounter.querySelector(".user-dropdown-counter").textContent = `Mostrando ${LIMITE_DROPDOWN} de ${filtrados.length}`;
            $dropdown.append(cloneCounter);
          }

          $dropdown.addClass("show");
          $("#searchChevron").css("transform", "rotate(180deg)");
        }

        function _aplicarEstadoTela(usuarioSelecionado) {
          if (usuarioSelecionado) {
            $("#divPlaceholder").hide();
            $("#divInfos").fadeIn(200);
            $("#Salvar").prop("disabled", false);
            $(".card-principal").addClass("has-data");
          } else {
            $("#divPlaceholder").show();
            $("#divInfos").hide();
            $("#Salvar").prop("disabled", true);
            $(".card-principal").removeClass("has-data");
          }
        }

        function SelecionouUsuario(id) {
          const idStr = String(id || "").trim();
          if (idStr) {
            _aplicarEstadoTela(true);
            ObjAjax("ajax/Usuario.asp?id=" + encodeURIComponent(idStr));
          } else {
            _limparCampos();
          }
        }

        function _parseDateInput(val) {
          const s = (val || "").toString().trim().replace(/\D/g, "");
          if (s.length === 8) return s.slice(0, 4) + "-" + s.slice(4, 6) + "-" + s.slice(6, 8);
          return "";
        }

        window.objReturn = function (obj) {
          if (obj && obj.erro === "1") {
            mostrarNotificacao("Não foi possível realizar a operação. Tente novamente mais tarde.", "danger");
            _limparCampos();
            return;
          }

          const dados = obj || {};
          $("#Matricula").val(dados.val1 || "");
          $("#Nome").val(dados.val3 || "");
          $("#Email").val(dados.val4 || "");
          $("#Nascimento").val(_parseDateInput(dados.val5));
          $("#Usuario").val(dados.val6 || "");

          const userAD    = (dados.val14 || "").trim();
          const sepIdx    = userAD.indexOf("\\");
          const prefixoAD = sepIdx > -1 ? userAD.slice(0, sepIdx).toUpperCase() : "GROUP";
          const loginAD   = sepIdx > -1 ? userAD.slice(sepIdx + 1) : userAD;
          const prefixoValido = ["GROUP", "WESTROCK"].includes(prefixoAD) ? prefixoAD : "GROUP";
          $("#PrefixoAD").val(prefixoValido);
          if (loginAD) $("#Usuario").val(loginAD);
          _atualizarUserAD();
          $("#Ramal").val(dados.val10 || "");
          $("#Senha").val("");
          $("#ConfirmaSenha").val("");
          $("#Bloqueado").prop("checked", dados.val11 === "1");

          const sexo = (dados.val2 || "").toUpperCase();
          $("#sexoM").prop("checked", sexo === "M");
          $("#sexoF").prop("checked", sexo === "F");

          const tipoAlvo = String(dados.val8 || "");
          const filialAlvo = String(dados.val9 || "");
          $("#Tipo option").each(function () {
            $(this).prop("selected", $(this).val() === tipoAlvo);
          });
          $("#Filial option").each(function () {
            $(this).prop("selected", $(this).val() === filialAlvo);
          });

          _validarSenhaInline();
          _aplicarEstadoTela(true);
        };

        function _limparCampos() {
          ["Matricula", "Nome", "Email", "Nascimento", "Usuario", "UserAD", "Senha", "ConfirmaSenha", "Ramal"].forEach((id) => {
            $("#" + id).val("");
          });
          $("#PrefixoAD").val("GROUP");
          _atualizarUserAD();
          $("#sexoM, #sexoF").prop("checked", false);
          $("#Tipo, #Filial").prop("selectedIndex", 0);
          $("#Bloqueado").prop("checked", false);
          $("#Senha, #ConfirmaSenha").removeClass("is-valid is-invalid");
          $("#fbSenha, #fbConfirma").text("").attr("class", "senha-feedback");
          _aplicarEstadoTela(false);
        }

        function _validarSenhaInline() {
          const senha = $("#Senha").val();
          const confirma = $("#ConfirmaSenha").val();

          // Senha é opcional: se ambos os campos estiverem vazios, não há nada a validar
          if (senha.length === 0 && confirma.length === 0) {
            $("#Senha, #ConfirmaSenha").removeClass("is-valid is-invalid");
            $("#fbSenha, #fbConfirma").text("").attr("class", "senha-feedback");
            return true;
          }

          $("#Senha").removeClass("is-valid is-invalid");
          if (senha.length > 0) {
            $("#Senha").addClass(senha.length >= 7 ? "is-valid" : "is-invalid");
          }

          if (senha.length > 0 && senha.length < 7) {
            $("#fbSenha").text("Mínimo de 7 caracteres.").attr("class", "senha-feedback senha-feedback-erro");
          } else {
            $("#fbSenha").text("").attr("class", "senha-feedback");
          }

          if (confirma.length === 0) {
            $("#ConfirmaSenha").removeClass("is-valid is-invalid");
            $("#fbConfirma").text("").attr("class", "senha-feedback");
          } else if (senha !== confirma) {
            $("#ConfirmaSenha").removeClass("is-valid").addClass("is-invalid");
            $("#fbConfirma").text("As senhas não conferem.").attr("class", "senha-feedback senha-feedback-erro");
          } else {
            $("#ConfirmaSenha").removeClass("is-invalid").addClass("is-valid");
            $("#fbConfirma").text("Senhas conferem.").attr("class", "senha-feedback senha-feedback-ok");
          }

          return senha.length >= 7 && senha === confirma;
        }

        function _submeterFormulario() {
          if (!_validarSenhaInline()) {
            mostrarNotificacao("Verifique os campos de senha antes de salvar.", "warning");
            return;
          }

          $("#hBloqueado").val($("#Bloqueado").is(":checked") ? "1" : "0");

          const $btn = $("#Salvar");
          $btn.prop("disabled", true).html(
            '<span class="spinner-border spinner-border-sm me-1" aria-hidden="true"></span>'
            //  +
            // '<span role="status">Salvando...</span>',
          );
          document.getElementById("CadastroUsuario").submit();
        }

        function _atualizarUserAD() {
          const prefixo = $("#PrefixoAD").val();
          const login   = $("#Usuario").val().trim();
          $("#PrefixoDisplay").text(prefixo + "\\");
          $("#UserAD").val(login ? prefixo + "\\" + login : "");
        }

        function _toggleSenha(campoId, btn) {
          const $input = $("#" + campoId);
          const $icone = $(btn).find("i");
          const mostrar = $input.attr("type") === "password";
          $input.attr("type", mostrar ? "text" : "password");
          $icone.attr("class", mostrar ? "bi bi-eye-slash" : "bi bi-eye");
          $(btn).attr("aria-label", mostrar ? "Ocultar senha" : "Mostrar senha");
        }

        function mostrarNotificacao(msg, tipo = "info") {
          const cfg = TOAST_ICON[tipo] ?? TOAST_ICON.info;
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
        });
      })(jQuery);
    </script>

  </body>
</html>
