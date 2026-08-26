<%
Session.LCID = 1046
%>
<!--#include file="../cnn2/cnn2.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../includes/variaveis.asp" -->
<%
Const CONN_LOG = "PIR-CORPORATIVO"

Dim existeDivergencia, msg
existeDivergencia = 0
msg = ""

Dim cadastroSucesso
cadastroSucesso = False

If Trim(Request("Matricula")) <> "" Then

    Dim Bloqueado
    If Trim(Request("Bloqueado")) = "True" Or Trim(Request("Bloqueado")) = "Verdadeiro" Then
        Bloqueado = 1
    Else
        Bloqueado = 0
    End If

    Dim ExisteUsuarioCadastrado
    ExisteUsuarioCadastrado = 0

    Dim oConnChk, oCmdChk, oRsChk
    Set oConnChk = Server.CreateObject("ADODB.Connection")
    oConnChk.Open BuscaStringConexao("PIR-HELPDESK")

    Set oCmdChk = Server.CreateObject("ADODB.Command")
    Set oCmdChk.ActiveConnection = oConnChk
    oCmdChk.CommandType = 1
    oCmdChk.CommandText = _
        "SELECT COUNT(*) AS EXISTE FROM CRM_DTI " & _
        "WHERE EMP_CODIGO = ? AND CRM_COD = ? AND ISNULL(CRM_D_E_L_E_T, '') = ''"
    oCmdChk.Parameters.Append oCmdChk.CreateParameter("@EMP_CODIGO", 200, 1, 6, Trim(Request("Filial")))
    oCmdChk.Parameters.Append oCmdChk.CreateParameter("@CRM_COD",    200, 1, 6, Trim(Request("Matricula")))

    Set oRsChk = oCmdChk.Execute

    If Not oRsChk.EOF And CInt(oRsChk("EXISTE")) > 0 Then
        ExisteUsuarioCadastrado = 1
    End If

    oRsChk.Close()
    Set oRsChk  = Nothing
    Set oCmdChk = Nothing
    oConnChk.Close()
    Set oConnChk = Nothing

    If ExisteUsuarioCadastrado = 1 Then
        existeDivergencia = 1
        msg = msg & "<div class=""alert alert-warning mb-3"">"
        msg = msg & "<h5 class=""mb-0""><i class=""bi bi-exclamation-triangle-fill me-2""></i><strong>Alerta!</strong> Usuário já cadastrado. "
        msg = msg & "<a href=""javascript:history.back()"" class=""alert-link"">Clique aqui</a> para voltar e corrigir.</h5>"
        msg = msg & "</div>"
    End If

    If Len(Trim(Request("Matricula"))) < 6 And Trim(Request("Filial")) <> "05" And Trim(Request("Filial")) <> "06" Then
        existeDivergencia = 1
        msg = msg & "<div class=""alert alert-warning mb-3"">"
        msg = msg & "<h5 class=""mb-0""><i class=""bi bi-exclamation-triangle-fill me-2""></i><strong>Alerta!</strong> O campo matrícula deve conter 6 dígitos. "
        msg = msg & "<a href=""javascript:history.back()"" class=""alert-link"">Clique aqui</a> para voltar e corrigir.</h5>"
        msg = msg & "</div>"
    End If

    If existeDivergencia = 0 Then

        Dim oConnIns, oCmdIns
        Set oConnIns = Server.CreateObject("ADODB.Connection")
        oConnIns.Open BuscaStringConexao("PIR-HELPDESK")

        Set oCmdIns = Server.CreateObject("ADODB.Command")
        Set oCmdIns.ActiveConnection = oConnIns
        oCmdIns.CommandType = 1
        oCmdIns.CommandText = _
            "INSERT INTO CRM_DTI (" & _
                "CRM_COD, CRM_DATCAD, HP_ADMHP, CRM_DATEXP, EST_COD, TRA_COD, " & _
                "CRM_NOME, RA_SEXO, RA_NASC, CRM_LOGIN, Senha, CRM_SENHA, " & _
                "CRM_EMAILPRIMARIO, CRM_RAMAL, EMP_CODIGO, CRM_TIPO, " & _
                "TROCARSENHA, CAT_CODIGO, Bloqueado, CRM_USERAD, CRM_VALNIVEL" & _
            ") VALUES (" & _
                "?, ?, 'N', ?, 1, 1, " & _
                "?, ?, ?, ?, HASHBYTES('SHA1', ?), ?, " & _
                "?, ?, ?, ?, " & _
                "?, 1, ?, ?, ?" & _
            ")"

        Dim sMatricula, sNome, sSexo, sNasc, sLogin, sSenha
        Dim sEmail, sRamal, sFilial, sTipo, sTrocarSenha, sUserAD
        Dim iRepresentante

        sMatricula   = Left(Trim(Request("Matricula")), 6)
        sNome        = Left(UCase(Trim(Request("Nome"))), 60)
        sSexo        = Left(Trim(Request("radios")), 1)
        sNasc        = inverteData(Trim(Request("Nascimento")))
        sLogin       = Left(Trim(Request("Usuario")), 30)
        sSenha       = Left(Trim(Request("Senha")), 50)
        sEmail       = Left(LCase(Trim(Request("Email"))), 255)
        sRamal       = Left(Trim(Request("Ramal")), 4)
        sFilial      = Left(Trim(Request("Filial")), 6)
        sTipo        = Left(Trim(Request("Tipo")), 1)
        sTrocarSenha = Left(Trim(Request("TrocarSenha")), 1)
        sUserAD      = Left(Trim(Request("UserAD")), 255)
        If Request("Representante") = "S" Then iRepresentante = "1" Else iRepresentante = "0"

        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_COD",          200, 1, 6,   sMatricula)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_DATCAD",        200, 1, 8,   inverteData(Date()))
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_DATEXP",        200, 1, 8,   inverteData(DateAdd("d", 42, Date())))
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_NOME",          200, 1, 60,  sNome)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@RA_SEXO",           200, 1, 1,   sSexo)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@RA_NASC",           200, 1, 8,   sNasc)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_LOGIN",         200, 1, 30,  sLogin)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@SenhaHash",         200, 1, 50,  sSenha)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_SENHA",         200, 1, 50,  sSenha)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_EMAILPRIMARIO", 200, 1, 255, sEmail)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_RAMAL",         200, 1, 4,   sRamal)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@EMP_CODIGO",        200, 1, 6,   sFilial)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_TIPO",          200, 1, 1,   sTipo)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@TROCARSENHA",       200, 1, 1,   sTrocarSenha)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@Bloqueado",         3,   1, 4,   CLng(Bloqueado))
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_USERAD",        200, 1, 255, sUserAD)
        oCmdIns.Parameters.Append oCmdIns.CreateParameter("@CRM_VALNIVEL",      200, 1, 1,   iRepresentante)

        oCmdIns.Execute

        Set oCmdIns = Nothing
        oConnIns.Close()
        Set oConnIns = Nothing

        Dim oConnLog, oCmdLog
        Set oConnLog = Server.CreateObject("ADODB.Connection")
        oConnLog.Open BuscaStringConexao(CONN_LOG)

        Set oCmdLog = Server.CreateObject("ADODB.Command")
        Set oCmdLog.ActiveConnection = oConnLog
        oCmdLog.CommandType = 1
        oCmdLog.CommandText = _
            "INSERT INTO dbo.LogAtividade " & _
                "(IdUsuarioAcao, IdUsuarioAlvo, Sistema, Rotina, Acao, Descricao, DataCriacao) " & _
            "VALUES (?, ?, 'SWnet', 'CadUsuario', 'INCLUIR', ?, GETDATE())"

        oCmdLog.Parameters.Append oCmdLog.CreateParameter("@IdUsuarioAcao", 200, 1, 50,  Session("USR"))
        oCmdLog.Parameters.Append oCmdLog.CreateParameter("@IdUsuarioAlvo", 200, 1, 6,   sMatricula)
        oCmdLog.Parameters.Append oCmdLog.CreateParameter("@Descricao",     200, 1, 500, _
            "Usuario " & sMatricula & " - " & sNome & _
            " incluido pelo operador " & Session("USR") & _
            " (IP: " & Request.ServerVariables("REMOTE_ADDR") & ")")

        oCmdLog.Execute

        Set oCmdLog = Nothing
        oConnLog.Close()
        Set oConnLog = Nothing

        If Trim(Request("EnviaEmail")) = "S" Then

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

        End If

        cadastroSucesso = True

    End If

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
    <title>SWnet - Incluir Usuário</title>

    <link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

  <style>
      body {
        background: #f0f0f0;
        font-size: 0.8125rem;
        min-height: 100vh;
      }
      
      .panel-header {
        border-bottom: 1px solid #dee2e6;
        padding: 0.6rem 1rem;
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
      
      @keyframes spin {
        to {
          transform: rotate(360deg);
        }
      }
      .spin {
        display: inline-block;
        animation: spin 0.75s linear infinite;
      }
  </style>
  
</head>

  <body>
    <div class="toast-wrapper" id="toastWrapper"></div>
    <div class="container-fluid px-3 mt-5 mb-5" style="max-width:54rem">
      <%
      If existeDivergencia > 0 Then
          Response.Write msg
          Response.End
      End If
      %>
      <form name="CadastroUsuario" id="CadastroUsuario" method="post" autocomplete="off">
        <input type="hidden" name="Bloqueado"   value="False" />
        <input type="hidden" name="TrocarSenha" value="S" />
        <input type="hidden" name="EnviaEmail"  value="S" />

        <div class="card shadow bg-light">
          <div class="panel-header rounded-top">
            Incluir Usuário
          </div>

          <div class="card-body pt-4 pb-3 px-4">
            <div class="card border mb-3 bg-white">
              <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                <i class="bi bi-person-vcard text-secondary"></i>
                <span class="fw-semibold card-header-label">Dados Pessoais</span>
              </div>
              <div class="card-body px-3 py-3">

                <div class="row g-3 mb-3">
                  <div class="col-md-3">
                    <label class="form-label fw-semibold" for="Matricula">Matrícula</label>
                    <input id="Matricula" name="Matricula" type="text" placeholder="Matrícula"
                           class="form-control form-control-sm" maxlength="6" required />
                  </div>
                  <div class="col-md-5">
                    <label class="form-label fw-semibold d-block">Sexo</label>
                    <div class="d-flex gap-4 pt-1">
                      <div class="form-check">
                        <input class="form-check-input" type="radio" name="radios" id="sexoM" value="M" required />
                        <label class="form-check-label small" for="sexoM">Masculino</label>
                      </div>
                      <div class="form-check">
                        <input class="form-check-input" type="radio" name="radios" id="sexoF" value="F" />
                        <label class="form-check-label small" for="sexoF">Feminino</label>
                      </div>
                    </div>
                  </div>
                  <div class="col-md-4">
                    <label class="form-label fw-semibold" for="Nascimento">Nascimento</label>
                    <input id="Nascimento" name="Nascimento" type="date" class="form-control form-control-sm" />
                  </div>
                </div>

                <div class="row g-3 mb-3">
                  <div class="col-12">
                    <label class="form-label fw-semibold" for="Nome">Nome</label>
                    <input id="Nome" name="Nome" type="text" placeholder="Nome completo"
                           class="form-control form-control-sm" required />
                  </div>
                </div>

                <div class="row g-3">
                  <div class="col-md-8">
                    <label class="form-label fw-semibold" for="Email">E-mail</label>
                    <input id="Email" name="Email" type="email" placeholder="exemplo@smurfitwestrock.com"
                           autocomplete="username" class="form-control form-control-sm" required />
                  </div>
                  <div class="col-md-4">
                    <label class="form-label fw-semibold d-block">Representante</label>
                    <div class="d-flex gap-4 pt-1">
                      <div class="form-check">
                        <input class="form-check-input" type="radio" name="Representante" id="Representante1" value="S" />
                        <label class="form-check-label small" for="Representante1">Sim</label>
                      </div>
                      <div class="form-check">
                        <input class="form-check-input" type="radio" name="Representante" id="Representante2" value="N" checked />
                        <label class="form-check-label small" for="Representante2">Não</label>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div class="card border mb-3 bg-white">
              <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                <i class="bi bi-shield-lock text-secondary"></i>
                <span class="fw-semibold card-header-label">Dados de Acesso</span>
              </div>
              <div class="card-body px-3 py-3">

                <input type="hidden" id="UserAD" name="UserAD" />

                <div class="row g-3 mb-3">
                  <div class="col-auto">
                    <label class="form-label fw-semibold" for="PrefixoAD">Usuário AD</label>
                    <select id="PrefixoAD" name="PrefixoAD" class="form-select form-select-sm" style="width:120px;">
                      <option value="GROUP">GROUP</option>
                      <option value="WESTROCK">WESTROCK</option>
                    </select>
                  </div>
                  <div class="col">
                    <label class="form-label fw-semibold" for="Usuario">Usuário</label>
                    <div class="input-group input-group-sm">
                      <span id="PrefixoDisplay" class="input-group-text fw-semibold px-2"
                            style="background:#f8f9fa;color:#4a4a4a;white-space:nowrap;">GROUP\</span>
                      <input id="Usuario" name="Usuario" type="text" placeholder="Login do Usuário..."
                             class="form-control form-control-sm" required maxlength="30" />
                    </div>
                  </div>
                </div>

                <div class="row g-3">
                  <div class="col-md-6">
                    <label class="form-label fw-semibold" for="Senha">Senha</label>
                    <div class="input-group input-group-sm">
                      <input id="Senha" name="Senha" type="password" placeholder="Nova senha"
                             autocomplete="new-password" class="form-control" required />
                      <button type="button" class="btn btn-senha-toggle" data-target="Senha" aria-label="Mostrar senha">
                        <i class="bi bi-eye"></i>
                      </button>
                    </div>
                    <div id="fbSenha" class="senha-feedback"></div>
                  </div>
                  <div class="col-md-6">
                    <label class="form-label fw-semibold" for="ConfirmaSenha">Confirmar Senha</label>
                    <div class="input-group input-group-sm">
                      <input id="ConfirmaSenha" name="ConfirmaSenha" type="password" placeholder="Confirme a senha"
                             autocomplete="new-password" class="form-control" required />
                      <button type="button" class="btn btn-senha-toggle" data-target="ConfirmaSenha" aria-label="Mostrar senha">
                        <i class="bi bi-eye"></i>
                      </button>
                    </div>
                    <div id="fbConfirma" class="senha-feedback"></div>
                  </div>
                </div>

              </div>
            </div>

            <div class="card border mb-3 bg-white">
              <div class="card-header bg-white py-2 px-3 d-flex align-items-center gap-2 border-bottom">
                <i class="bi bi-building text-secondary"></i>
                <span class="fw-semibold card-header-label">Vínculo</span>
              </div>
              <div class="card-body px-3 py-3">

                <div class="row g-3">
                  <div class="col-md-4">
                    <label class="form-label fw-semibold" for="Tipo">Tipo</label>
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
                    <label class="form-label fw-semibold" for="Filial">Empresa</label>
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
                    <label class="form-label fw-semibold" for="Ramal">Ramal</label>
                    <input id="Ramal" name="Ramal" type="text" placeholder="Ramal"
                           class="form-control form-control-sm" maxlength="6" />
                  </div>
                </div>
              </div>
            </div>

            <div class="d-flex justify-content-center gap-2 mt-3">
              <button id="btnSalvar" name="btnSalvar" type="button"
                      class="btn btn-sm btn-success px-4" aria-label="Salvar novo usuário">
                Salvar
              </button>
            </div>
          </div>
        </div>
      </form>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    <script src="../funcoes/funcoesJAVASCRIPT.js"></script>

    <% If cadastroSucesso Then %>
    <script>var msgSucesso = 'Usuário incluído com sucesso! E-mail enviado.';</script>
    <% End If %>

    <script>
      (($) => {
        "use strict";

        const TOAST_DELAY = 4500;

        const TOAST_ICON = Object.freeze({
          success: { cls: "bi-check-circle-fill",        cor: "#198754" },
          warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
          danger:  { cls: "bi-x-circle-fill",            cor: "#dc3545" },
          info:    { cls: "bi-info-circle-fill",          cor: "#6c757d" },
        });

        $(document).ready(() => {
          if (typeof msgSucesso !== "undefined" && msgSucesso) {
            mostrarNotificacao(msgSucesso, "success");
          }
          _bindEvents();
        });

        function _bindEvents() {
          $(document).on("click", ".btn-senha-toggle", function () {
            _toggleSenha($(this).data("target"), this);
          });
          $("#Senha, #ConfirmaSenha").on("input", _validarSenhaInline);
          $("#Usuario").on("input", _atualizarUserAD);
          $("#PrefixoAD").on("change", _atualizarUserAD);
          _atualizarUserAD();
          $("#btnSalvar").on("click", _submeterFormulario);
          $("#CadastroUsuario").on("submit", function (e) {
            if (!_validarSenhaInline()) e.preventDefault();
          });
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

        function _validarSenhaInline() {
          const senha    = $("#Senha").val();
          const confirma = $("#ConfirmaSenha").val();

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
          const $btn = $("#btnSalvar");
          $btn.prop("disabled", true).html(
            '<span class="spinner-border spinner-border-sm me-1" aria-hidden="true"></span>' +
            '<span role="status">Salvando...</span>',
          );
          document.getElementById("CadastroUsuario").submit();
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

      })(jQuery);
    </script>
  </body>
</html>
