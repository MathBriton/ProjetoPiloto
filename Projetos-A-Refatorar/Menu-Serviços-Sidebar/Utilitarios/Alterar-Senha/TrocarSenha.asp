<!-- #include file = "../cnn2/cnn2.asp" -->
<!-- #include file = "../includes/variaveis.asp" -->
<%
'----------------------------------------------------------------------
' FUNÇÕES AUXILIARES
'----------------------------------------------------------------------
function VerificaMaiuscula(texto)
  Dim alfabetoMaiusculo, listaAlfabetoMaiusculo, existeMaiuscula, i
  alfabetoMaiusculo = UCase("a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z")
  listaAlfabetoMaiusculo = Split(alfabetoMaiusculo, ",")
  existeMaiuscula = 0
  for i = 0 to Ubound(listaAlfabetoMaiusculo)
    if instr(texto, listaAlfabetoMaiusculo(i)) > 0 then existeMaiuscula = 1
  next
  VerificaMaiuscula = existeMaiuscula
end function

function VerificaMinuscula(texto)
  Dim alfabetoMinusculo, listaAlfabetoMinusculo, existeMinuscula, i
  alfabetoMinusculo = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
  listaAlfabetoMinusculo = Split(alfabetoMinusculo, ",")
  existeMinuscula = 0
  for i = 0 to Ubound(listaAlfabetoMinusculo)
    if instr(texto, listaAlfabetoMinusculo(i)) > 0 then existeMinuscula = 1
  next
  VerificaMinuscula = existeMinuscula
end function

function VerificaCaracterEspecial(texto)
  Dim caracterEspecial, listaCaracterEspecial, existeCaracterEspecial, i
  caracterEspecial = "!,@,#,$,%,¨,&,*,(,),-,+,=,§,?,/,°,|,\,{,},[,],º,_,;,¬,."
  listaCaracterEspecial = Split(caracterEspecial, ",")
  existeCaracterEspecial = 0
  for i = 0 to Ubound(listaCaracterEspecial)
    if instr(texto, listaCaracterEspecial(i)) > 0 then existeCaracterEspecial = 1
  next
  VerificaCaracterEspecial = existeCaracterEspecial
end function

function VerificaNumero(texto)
  Dim numeros, listaNumeros, existeNumero, i
  numeros = "1,2,3,4,5,6,7,8,9,0"
  listaNumeros = Split(numeros, ",")
  existeNumero = 0
  for i = 0 to Ubound(listaNumeros)
    if instr(texto, listaNumeros(i)) > 0 then existeNumero = 1
  next
  VerificaNumero = existeNumero
end function

function VerificaSequencia(texto)
  Dim sequenciaNumerica, sequenciaAlfabetica
  Dim listaSequenciaNumerica, listaSequenciaAlfabetica
  Dim existeSequenciaNumerica, existeSequenciaAlfabetica, i
  sequenciaNumerica  = "012,123,234,345,456,567,678,789,8910"
  sequenciaAlfabetica = "abc,bcd,cde,def,efg,fgh,ghi,hij,ijk,jkl,klm,lmn,mno,nop,opq,pqr,qrs,rst,stu,tuv,uvw,vwx,wxy,xyz,yza,zab"
  listaSequenciaNumerica  = Split(sequenciaNumerica, ",")
  listaSequenciaAlfabetica = Split(sequenciaAlfabetica, ",")
  existeSequenciaNumerica  = false
  existeSequenciaAlfabetica = false
  for i = 0 to Ubound(listaSequenciaNumerica)
    if instr(texto, listaSequenciaNumerica(i)) > 0 then existeSequenciaNumerica = true
  next
  for i = 0 to Ubound(listaSequenciaAlfabetica)
    if instr(LCase(texto), listaSequenciaAlfabetica(i)) > 0 then existeSequenciaAlfabetica = true
  next
  VerificaSequencia = Not(existeSequenciaNumerica or existeSequenciaAlfabetica)
end function

function VerificaRepeticao(texto)
  Dim repeticaoNumerica, repeticaoAlfabetica
  Dim listaRepeticaoNumerica, listaRepeticaoAlfabetica
  Dim existeRepeticaoNumerica, existeRepeticaoAlfabetica, i
  repeticaoNumerica  = "000,111,222,333,444,555,666,777,888,999"
  repeticaoAlfabetica = "aaa,bbb,ccc,ddd,eee,fff,ggg,hhh,iii,jjj,kkk,lll,mmm,nnn,ooo,ppp,qqq,rrr,sss,ttt,uuu,vvv,www,xxx,yyy,zzz"
  listaRepeticaoNumerica  = Split(repeticaoNumerica, ",")
  listaRepeticaoAlfabetica = Split(repeticaoAlfabetica, ",")
  existeRepeticaoNumerica  = false
  existeRepeticaoAlfabetica = false
  for i = 0 to Ubound(listaRepeticaoNumerica)
    if instr(texto, listaRepeticaoNumerica(i)) > 0 then existeRepeticaoNumerica = true
  next
  for i = 0 to Ubound(listaRepeticaoAlfabetica)
    if instr(LCase(texto), listaRepeticaoAlfabetica(i)) > 0 then existeRepeticaoAlfabetica = true
  next
  VerificaRepeticao = Not(existeRepeticaoNumerica or existeRepeticaoAlfabetica)
end function

function GravaHistorico(idUsuario, texto)
  Dim sql
  sql = "INSERT INTO HistoricoUsuario(IdUsuario,[Hash],[Data]) VALUES(" & _
        idUsuario & "," & _
        "HASHBYTES('SHA1','" & texto & "')," & _
        "GETDATE())"
  Call Conecta("PIR-HELPDESK", sql)

  sql = "SELECT COUNT(*) EXISTE FROM HistoricoUsuario " & _
        "WHERE IdUsuario=" & idUsuario & " " & _
        "AND Hash=HASHBYTES('SHA1','" & texto & "')"
  Set rsGrava = Conecta("PIR-HELPDESK", sql)
  GravaHistorico = iif((not rsGrava Is Nothing) and Cint(rsGrava("Existe")) > 0, 1, 0)
  if not rsGrava Is Nothing then
    rsGrava.Close()
    Set rsGrava = Nothing
  end if
end function

function ValidarHistorico(idUsuario, texto)
  Dim sql
  sql = "SELECT COUNT(*) EXISTE FROM (" & _
          "SELECT TOP 24 * FROM HistoricoUsuario " & _
          "WHERE IdUsuario=" & idUsuario & " ORDER BY IdHistorico Desc" & _
        ") AS Historico " & _
        "WHERE IdUsuario=" & idUsuario & " " & _
        "AND Hash=HASHBYTES('SHA1','" & texto & "')"
  Set rsHistorico = Conecta("PIR-HELPDESK", sql)
  ValidarHistorico = iif((not rsHistorico Is Nothing) and Cint(rsHistorico("Existe")) > 0, 0, 1)
  if not rsHistorico Is Nothing then
    rsHistorico.Close()
    Set rsHistorico = Nothing
  end if
end function

function ValidaSenhaIgual(atual, nova)
  ValidaSenhaIgual = iif(atual = nova, 0, 1)
end function

function ValidaTamanho(texto, nivel)
  if nivel = 5 then
    ValidaTamanho = iif(Len(trim(texto)) >= 24, 1, 0)
  else
    ValidaTamanho = iif(Len(trim(texto)) >= 10, 1, 0)
  end if
end function

function mensagem(aMsg)
  mensagem = "<div class='alert-inline alert-inline-danger' role='alert'>" & _
             "<i class='bi bi-exclamation-circle-fill me-2' aria-hidden='true'></i>" & _
             aMsg & _
             "</div>"
end function

function nivelDoUsuario(aUser)
  Dim rLevel, sql
  sql = "SELECT CRM_TIPO FROM CRM_DTI " & _
        "WHERE CRM_LOGIN='" & iif(trim(aUser) = "", USR, trim(aUser)) & "'"
  Set rsNivel = Conecta("PIR-HELPDESK", sql)
  rLevel = iif((not rsNivel Is Nothing) and (not rsNivel.EOF), trim(rsNivel("CRM_TIPO")), "")
  if not rsNivel Is Nothing then
    rsNivel.Close()
    Set rsNivel = Nothing
  end if
  nivelDoUsuario = rLevel
end function

'----------------------------------------------------------------------
' BACKEND — lógica de troca de senha
'----------------------------------------------------------------------
session.LCID = 1046

login = iif((trim(request("Usuario")) <> ""), trim(request("Usuario")), SESSION("USR"))
msg = ""
passouValidacao = 0
tipo = nivelDoUsuario(login)

if trim(request("SenhaAtual")) <> "" and trim(request("Senha")) <> "" then

  usuario = ""
  codigoUsuario = ""
  codigoFilial = ""

  sql = "" & _
    "SELECT " & _
      "CRM_R_E_C_N_O, " & _
      "CRM_LOGIN, " & _
      "CRM_COD, " & _
      "EMP_CODIGO, " & _
      "CRM_TIPO, " & _
      "Bloqueado, " & _
      "CRM_D_E_L_E_T " & _
    "FROM CRM_DTI " & _
    "WHERE CRM_LOGIN = '" & trim(request("Usuario")) & "' " & _
      "AND Senha = HASHBYTES('SHA1', '" & trim(request("SenhaAtual")) & "') "
  Set rsAuth = Conecta("PIR-HELPDESK", sql)

  if (not rsAuth Is Nothing) and (not rsAuth.Eof) then
    usuario          = trim(rsAuth("CRM_LOGIN"))
    codigoUsuario    = trim(rsAuth("CRM_COD"))
    codigoFilial     = trim(rsAuth("EMP_CODIGO"))
    idUsuarioHelpdesk = trim(rsAuth("CRM_R_E_C_N_O"))
    nivel            = trim(rsAuth("CRM_TIPO"))
    bloqueado        = trim(rsAuth("Bloqueado"))
    deletado         = trim(rsAuth("CRM_D_E_L_E_T"))
  else
    usuario           = ""
    codigoUsuario     = ""
    codigoFilial      = ""
    idUsuarioHelpdesk = ""
    nivel             = ""
    bloqueado         = ""
    deletado          = ""
  end if

  if not rsAuth Is Nothing then
    rsAuth.Close()
    Set rsAuth = Nothing
  end if

  ' --- Bug 3: usuário desativado (soft-delete) ---
  if deletado = "*" then
%>
    <div class="alert-inline alert-inline-danger mb-4" role="alert">
      <i class="bi bi-person-x-fill me-2"></i>
      Este usuário está desativado. Entre em contato com o suporte.
    </div>
    <div class="d-flex gap-2 mb-4">
      <a href="#" class="btn-sw btn-sw-ghost" onclick="Sair(); return false;">
        <i class="bi bi-box-arrow-right me-1"></i>Sair
      </a>
    </div>
<%
    response.end
  end if

  ' --- Bug 2: usuário bloqueado ---
  if bloqueado = "1" then
%>
    <div class="alert-inline alert-inline-danger mb-4" role="alert">
      <i class="bi bi-lock-fill me-2"></i>
      Usuário bloqueado. Entre em contato com o suporte.
    </div>
    <div class="d-flex gap-2 mb-4">
      <a href="#" class="btn-sw btn-sw-ghost" onclick="Sair(); return false;">
        <i class="bi bi-box-arrow-right me-1"></i>Sair
      </a>
    </div>
<%
    response.end
  end if

  if usuario = trim(request("Usuario")) and usuario <> "" then

    msg = ""

    if ValidaTamanho(trim(request("Senha")), nivel) = 0 then
      msg = msg & mensagem("A senha deve conter, no mínimo, " & iif(nivel = "5", "24", "10") & " caracteres.")
    end if

    if ValidaSenhaIgual(trim(request("SenhaAtual")), trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A nova senha deve ser diferente da senha atual.")
    end if

    if VerificaMaiuscula(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha deve conter, pelo menos, um caracter maiúsculo.")
    end if

    if VerificaMinuscula(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha deve conter, pelo menos, um caracter minúsculo.")
    end if

    if VerificaCaracterEspecial(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha deve conter, pelo menos, um caracter especial.")
    end if

    if VerificaNumero(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha deve conter, pelo menos, um número.")
    end if

    if VerificaSequencia(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha NÃO deve conter caracteres sequenciais. Ex.: 123 ou abc")
    end if

    if VerificaRepeticao(trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A senha NÃO deve conter repetições. Ex.: 111, aaa")
    end if

    if ValidarHistorico(idUsuarioHelpdesk, trim(request("Senha"))) = 0 then
      msg = msg & mensagem("A nova senha NÃO deve ter sido utilizada nas últimas 24 vezes que você a alterou.")
    end if

    if msg = "" then

      if GravaHistorico(idUsuarioHelpdesk, trim(request("SenhaAtual"))) = 1 then

        sql = "" & _
          "UPDATE CRM_DTI SET " & _
            "Senha = HASHBYTES('SHA1', '" & trim(request("Senha")) & "'), " & _
            "CRM_SENHA = HASHBYTES('SHA1', '" & trim(request("Senha")) & "'), " & _
            "CRM_DATEXP = '" & inverteData(DateAdd("d", 90, Date())) & "', " & _
            "TROCARSENHA = 'N' " & _
          "WHERE CRM_LOGIN = '" & trim(request("Usuario")) & "' " & _
            "AND Senha = HASHBYTES('SHA1', '" & trim(request("SenhaAtual")) & "') " & _
            "AND EMP_CODIGO = '" & codigoFilial & "' "
        Call Conecta("PIR-HELPDESK", sql)

        if codigoFilial <> "05" and codigoFilial <> "06" then
          sql = "" & _
            "UPDATE inpanetUSUARIOS SET " & _
              "HASH_SENHA = HASHBYTES('SHA1', '" & trim(request("Senha")) & "'), " & _
              "SENHA = '" & trim(request("Senha")) & "', " & _
              "DataSenhaExpira = '" & inverteData(DateAdd("d", 90, Date())) & "' " & _
            "WHERE CRM_COD = '" & codigoUsuario & "' " & _
              "AND FILIAL = '" & codigoFilial & "' "
          CALL EnviaEmail("jorge.queiroz@smurfitkappa.com.br", "", "", "", "", sql, "TESTE - TrocarSenha")
          Call Conecta("PIR-CORPORATIVO", sql)
        end if

        response.write "<script>alert('Senha alterada com sucesso!');</script>"
        Response.Redirect "../../Default.asp"

      else
        msg = mensagem("Não foi possível alterar sua senha. Por favor, tente novamente mais tarde.")
        response.Write msg
      end if

    else
      response.write msg
    end if

  else
%>
    <!-- Senha atual inválida — exibido inline antes do form -->
    <div class="alert-inline alert-inline-danger mb-4" role="alert">
      <i class="bi bi-exclamation-circle-fill me-2"></i>
      Senha atual inválida.
    </div>
    <div class="d-flex gap-2 mb-4">
      <a href="#" class="btn-sw btn-sw-ghost" onclick="Sair()">
        <i class="bi bi-box-arrow-right me-1"></i>Sair
      </a>
      <a href="TrocarSenha.asp?Usuario=<%=USR%>" class="btn-sw btn-sw-secondary">
        <i class="bi bi-arrow-clockwise me-1"></i>Tentar Novamente
      </a>
    </div>
<%
    response.end
  end if

end if
' FIM do bloco de processamento — abaixo começa o HTML da página

%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trocar Senha — SWnet</title>

  <!-- Bootstrap 5.3.8 -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" />
  <!-- Bootstrap Icons -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

  <style>
    /* ── Tokens de Design ─────────────────────────────────────────── */
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
      --color-success-soft:  hsl(145, 60%, 95%);
      --color-danger:        hsl(3, 90%, 58%);
      --color-danger-soft:   hsl(3, 100%, 96%);
      --color-warning:       hsl(35, 100%, 48%);
      --color-warning-soft:  hsl(40, 100%, 95%);

      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

      --radius-sm: 6px;
      --radius-md: 10px;
      --radius-lg: 14px;

      --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
      --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

      --transition: 160ms ease;

      --space-1: 4px;
      --space-2: 8px;
      --space-3: 12px;
      --space-4: 16px;
      --space-5: 24px;
      --space-6: 32px;
      --space-7: 48px;
    }

    /* ── Base ─────────────────────────────────────────────────────── */
    *, *::before, *::after { box-sizing: border-box; }

    body {
      font-family: var(--font-sans);
      background: var(--color-bg);
      color: var(--color-text);
      font-size: .9375rem;
      line-height: 1.6;
      min-height: 100vh;
      margin: 0;
    }

    /* ── Page wrapper ─────────────────────────────────────────────── */
    .page-wrapper {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      padding: var(--space-5) var(--space-4);
    }

    /* ── Card principal ───────────────────────────────────────────── */
    .pw-card {
      width: 100%;
      max-width: 680px;
      background: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-md);
      overflow: hidden;
    }

    /* ── Cabeçalho do card ────────────────────────────────────────── */
    .pw-header {
      padding: var(--space-5) var(--space-6);
      border-bottom: 1px solid var(--color-border);
      background: var(--color-surface-muted);
    }

    .pw-header-title {
      font-size: clamp(1rem, 0.9rem + 0.4vw, 1.2rem);
      font-weight: 600;
      color: var(--color-text);
      margin: 0 0 var(--space-1);
      display: flex;
      align-items: center;
      justify-content: center;
      gap: var(--space-2);
    }

    .pw-header-title i {
      color: var(--color-accent);
      font-size: 1.1em;
    }

    .pw-header-sub {
      font-size: .8125rem;
      color: var(--color-text-muted);
      margin: 0;
      text-align: center;
    }

    /* ── Corpo do card ────────────────────────────────────────────── */
    .pw-body {
      padding: var(--space-5) var(--space-6);
    }

    /* ── Alerta de regras ─────────────────────────────────────────── */
    .rules-card {
      background: var(--color-warning-soft);
      border: 1px solid hsl(35, 80%, 82%);
      border-left: 3px solid var(--color-warning);
      border-radius: var(--radius-sm);
      padding: var(--space-4) var(--space-5);
      margin-bottom: var(--space-5);
      font-size: .8438rem;
      color: var(--color-text);
    }

    .rules-card-title {
      font-weight: 600;
      font-size: .875rem;
      color: hsl(35, 100%, 30%);
      margin-bottom: var(--space-2);
      display: flex;
      align-items: center;
      justify-content: center;
      gap: var(--space-2);
    }

    .rules-list {
      margin: 0;
      padding: 0;
      list-style: none;
      display: flex;
      flex-direction: column;
      gap: var(--space-1);
      color: var(--color-text-muted);
    }

    .rules-list li {
      display: flex;
      align-items: flex-start;
      gap: var(--space-2);
    }

    .rules-list li i {
      color: var(--color-warning);
      flex-shrink: 0;
      margin-top: 3px;
    }

    .rules-strong {
      font-weight: 600;
      color: var(--color-danger);
    }

    /* ── Labels e inputs ──────────────────────────────────────────── */
    .field-label {
      display: block;
      font-size: .8125rem;
      font-weight: 500;
      color: var(--color-text-muted);
      margin-bottom: var(--space-1);
    }

    .field-input {
      width: 100%;
      padding: var(--space-2) var(--space-3);
      font-size: .9375rem;
      font-family: var(--font-sans);
      color: var(--color-text);
      background: var(--color-surface);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-sm);
      transition: border-color var(--transition), box-shadow var(--transition);
      outline: none;
      -webkit-appearance: none;
    }

    .field-input:hover {
      border-color: var(--color-border-strong);
    }

    .field-input:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .12);
    }

    .field-input[readonly] {
      background: var(--color-surface-muted);
      color: var(--color-text-muted);
      cursor: not-allowed;
    }

    .field-input.is-invalid {
      border-color: var(--color-danger);
      box-shadow: 0 0 0 3px hsla(3, 90%, 58%, .10);
    }

    /* ── Grade de campos ──────────────────────────────────────────── */
    .fields-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: var(--space-4);
      margin-bottom: var(--space-5);
    }

    .field-full { grid-column: 1 / -1; }

    /* ── Indicador de força da senha ──────────────────────────────── */
    .strength-bar-wrap {
      margin-top: var(--space-2);
      display: flex;
      flex-direction: column;
      gap: var(--space-1);
    }

    .strength-segments {
      display: flex;
      gap: var(--space-1);
    }

    .strength-seg {
      flex: 1;
      height: 3px;
      background: var(--color-border);
      border-radius: 99px;
      transition: background var(--transition);
    }

    .strength-label {
      font-size: .75rem;
      color: var(--color-text-faint);
      font-weight: 500;
    }

    /* ── Erro inline de campo ────────────────────────────────────── */
    .field-error-msg {
      display: none;
      align-items: center;
      gap: var(--space-1);
      margin-top: var(--space-1);
      font-size: .8125rem;
      font-weight: 500;
      color: var(--color-danger);
    }

    .field-error-msg.visible {
      display: flex;
    }

    /* ── Alertas inline (erros do backend) ────────────────────────── */
    .alert-inline {
      display: flex;
      align-items: center;
      gap: var(--space-2);
      padding: var(--space-3) var(--space-4);
      border-radius: var(--radius-sm);
      font-size: .875rem;
      font-weight: 500;
      margin-bottom: var(--space-4);
    }

    .alert-inline-danger {
      background: var(--color-danger-soft);
      border: 1px solid hsl(3, 70%, 88%);
      color: hsl(3, 70%, 38%);
    }

    .alert-inline-success {
      background: var(--color-success-soft);
      border: 1px solid hsl(145, 50%, 82%);
      color: var(--color-success);
    }

    /* ── Rodapé do card (ações) ───────────────────────────────────── */
    .pw-footer {
      padding: var(--space-4) var(--space-6);
      border-top: 1px solid var(--color-border);
      background: var(--color-surface-muted);
      display: flex;
      justify-content: center;
      gap: var(--space-3);
    }

    /* ── Botões ───────────────────────────────────────────────────── */
    .btn-sw {
      display: inline-flex;
      align-items: center;
      gap: var(--space-2);
      padding: var(--space-2) var(--space-4);
      font-size: .875rem;
      font-weight: 500;
      font-family: var(--font-sans);
      border-radius: var(--radius-sm);
      border: 1px solid transparent;
      cursor: pointer;
      text-decoration: none;
      transition: background var(--transition), border-color var(--transition), color var(--transition), box-shadow var(--transition);
      line-height: 1.5;
      min-height: 36px;
    }

    .btn-sw-primary {
      background: var(--color-accent);
      color: hsl(0, 0%, 100%);
      border-color: var(--color-accent);
    }

    .btn-sw-primary:hover {
      background: var(--color-accent-hover);
      border-color: var(--color-accent-hover);
      color: hsl(0, 0%, 100%);
    }

    .btn-sw-primary:focus-visible {
      outline: none;
      box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .25);
    }

    .btn-sw-primary:disabled {
      opacity: .45;
      cursor: not-allowed;
      pointer-events: none;
    }

    .btn-sw-ghost {
      background: transparent;
      color: var(--color-text-muted);
      border-color: var(--color-border);
    }

    .btn-sw-ghost:hover {
      background: var(--color-surface-muted);
      border-color: var(--color-border-strong);
      color: var(--color-text);
    }

    .btn-sw-ghost:focus-visible {
      outline: none;
      box-shadow: 0 0 0 3px hsla(240, 6%, 10%, .10);
    }

    .btn-sw-secondary {
      background: var(--color-surface);
      color: var(--color-accent);
      border-color: var(--color-border);
    }

    .btn-sw-secondary:hover {
      background: var(--color-accent-soft);
      border-color: hsl(210, 60%, 80%);
      color: var(--color-accent-hover);
    }

    /* ── Responsividade ───────────────────────────────────────────── */
    @media (max-width: 768px) {
      .pw-card { border-radius: var(--radius-md); }
      .pw-header,
      .pw-body { padding: var(--space-4); }
      .pw-footer { padding: var(--space-4); flex-direction: column; }
      .btn-sw { width: 100%; justify-content: center; }
      .fields-grid { grid-template-columns: 1fr; }
      .field-full { grid-column: 1; }
    }

    /* ── Reduced motion ───────────────────────────────────────────── */
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { transition: none !important; }
    }
  </style>
</head>

<body>
  <div class="page-wrapper">
    <div class="pw-card" role="main">

      <!-- Cabeçalho -->
      <div class="pw-header">
        <h1 class="pw-header-title">
          <i class="bi bi-shield-lock" aria-hidden="true"></i>
          Troca de Senha
        </h1>
        <!-- <p class="pw-header-sub">SWnet — Smurfit Westrock Brasil</p> -->
      </div>

      <!-- Corpo -->
      <div class="pw-body">

        <!-- Regras de senha -->
        <div class="rules-card" role="note" aria-label="Regras de senha">
          <div class="rules-card-title">
            <i class="bi bi-exclamation-triangle-fill" aria-hidden="true"></i>
            Atenção
          </div>
          <hr style="color:hsl(39, 100%, 50%)">
          <ul class="rules-list">
            <li>
              <i class="bi bi-dot" aria-hidden="true"></i>
              <span>
                Mínimo de
                <strong class="rules-strong"><% if nivelDoUsuario(request("Usuario")) = 5 then %>24<% else %>10<% end if %> caracteres.</strong>
              </span>
            </li>
            <li><i class="bi bi-dot" aria-hidden="true"></i><span>Pelo menos uma letra <strong class="rules-strong">maiúscula</strong> e uma <strong class="rules-strong">minúscula.</strong></span></li>
            <li><i class="bi bi-dot" aria-hidden="true"></i><span>Pelo menos um <strong class="rules-strong">número</strong> e um <strong class="rules-strong">caractere especial.</strong></span></li>
            <li><i class="bi bi-dot" aria-hidden="true"></i><span>Sem sequências como <strong class="rules-strong">123</strong> ou <strong class="rules-strong">abc</strong>, nem repetições como <strong class="rules-strong">111</strong> ou <strong class="rules-strong">aaa.</strong></span></li>
            <li><i class="bi bi-dot" aria-hidden="true"></i><span>Sem caracteres <strong class="rules-strong">[</strong> <strong class="rules-strong">]</strong> <strong class="rules-strong">(</strong> <strong class="rules-strong">).</strong></span></li>
            <li><i class="bi bi-dot" aria-hidden="true"></i><span>Não reutilize nenhuma das últimas <strong class="rules-strong">24 senhas.</strong></span></li>
          </ul>
        </div>

        <!-- Formulário -->
        <form id="formTrocarSenha" name="TrocarSenha" method="post" novalidate>
          <input type="hidden" name="nivel" id="Nivel" value="<%=tipo%>" />

          <div class="fields-grid">

            <!-- Usuário (readonly) -->
            <div class="field-full">
              <label for="Usuario" class="field-label">Usuário</label>
              <input
                id="Usuario" name="Usuario" type="text"
                class="field-input" readonly
                value="<%=login%>"
                aria-readonly="true"
              />
            </div>

            <!-- Senha Atual -->
            <div class="field-full">
              <label for="SenhaAtual" class="field-label">Senha Atual</label>
              <input
                id="SenhaAtual" name="SenhaAtual" type="password"
                class="field-input"
                placeholder="Digite sua senha atual"
                required
                autocomplete="current-password"
                aria-required="true"
              />
            </div>

            <!-- Nova Senha -->
            <div>
              <label for="Senha" class="field-label">Nova Senha</label>
              <input
                id="Senha" name="Senha" type="password"
                class="field-input"
                placeholder="Digite a nova senha"
                required
                autocomplete="new-password"
                onblur="validarCampoSenha()"
                oninput="atualizarForca()"
                aria-required="true"
                aria-describedby="forca-label"
              />
              <!-- Indicador de força -->
              <div class="strength-bar-wrap" aria-live="polite">
                <div class="strength-segments" aria-hidden="true">
                  <div class="strength-seg" id="seg1"></div>
                  <div class="strength-seg" id="seg2"></div>
                  <div class="strength-seg" id="seg3"></div>
                  <div class="strength-seg" id="seg4"></div>
                </div>
                <span class="strength-label" id="forca-label"></span>
              </div>
            </div>

            <!-- Confirmar Nova Senha -->
            <div>
              <label for="ConfirmaSenha" class="field-label">Confirmar Nova Senha</label>
              <input
                id="ConfirmaSenha" name="ConfirmaSenha" type="password"
                class="field-input"
                placeholder="Repita a nova senha"
                required
                autocomplete="new-password"
                onblur="validarSenha()"
                aria-required="true"
                aria-describedby="erro-confirma"
              />
              <span class="field-error-msg" id="erro-confirma" role="alert" aria-live="polite">
                <i class="bi bi-exclamation-circle-fill" aria-hidden="true"></i>
                As senhas não conferem.
              </span>
            </div>

          </div>

      </div><!-- /.pw-body -->

      <!-- Rodapé com ações -->
      <div class="pw-footer">
        <a href="#" class="btn-sw btn-sw-ghost" onclick="Sair(); return false;" aria-label="Sair sem salvar">
          <i class="bi bi-box-arrow-right" aria-hidden="true"></i>
          Sair
        </a>
        <button
          id="Salvar" type="submit"
          class="btn-sw btn-sw-primary"
          aria-label="Salvar nova senha"
        >
          <i class="bi bi-check2" aria-hidden="true"></i>
          Salvar
        </button>
      </div>
      </form>

    </div><!-- /.pw-card -->
  </div><!-- /.page-wrapper -->

  <!-- Bootstrap JS -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"></script>

  <script>
    /* ── Constantes injetadas pelo servidor ───────────────────────── */
    const NIVEL_USUARIO = "<%=tipo%>";

    /* ── Helpers ──────────────────────────────────────────────────── */
    function getEl(id) { return document.getElementById(id); }

    function setInvalid(el, invalido) {
      if (invalido) {
        el.classList.add('is-invalid');
        el.setAttribute('aria-invalid', 'true');
      } else {
        el.classList.remove('is-invalid');
        el.removeAttribute('aria-invalid');
      }
    }

    /* ── Indicador de força da senha ──────────────────────────────── */
    function calcularForca(senha) {
      if (!senha) return 0;
      let pontos = 0;
      if (senha.length >= 10)                                      pontos++;
      if (senha.length >= 16)                                      pontos++;
      if (/[A-Z]/.test(senha) && /[a-z]/.test(senha))             pontos++;
      if (/[0-9]/.test(senha) && /[^A-Za-z0-9]/.test(senha))      pontos++;
      return pontos;
    }

    function atualizarForca() {
      const senha = getEl('Senha').value;
      const forca = calcularForca(senha);
      const cores = [
        'hsl(3, 90%, 58%)',
        'hsl(35, 100%, 48%)',
        'hsl(35, 100%, 48%)',
        'hsl(145, 63%, 42%)',
        'hsl(145, 63%, 32%)'
      ];
      const labels = ['', 'Fraca', 'Razoável', 'Boa', 'Forte'];

      for (let i = 1; i <= 4; i++) {
        getEl('seg' + i).style.background =
          i <= forca ? cores[forca] : 'var(--color-border)';
      }

      const labelEl = getEl('forca-label');
      labelEl.textContent = senha.length ? labels[forca] || '' : '';
      labelEl.style.color = forca <= 1 ? 'var(--color-danger)' :
                             forca <= 2 ? 'var(--color-warning)' :
                                          'var(--color-success)';
    }

    /* ── Validação: campo senha (onblur da senha atual / nova) ────── */
    function validarCampoSenha() {
      const nivel = getEl('Nivel').value;
      const senha = getEl('Senha');
      const confirmaSenha = getEl('ConfirmaSenha');
      const minChar = nivel === '5' ? 24 : 10;

      if (senha.value.length > 0 && senha.value.length < minChar) {
        setInvalid(senha, true);
        return false;
      }

      setInvalid(senha, false);
      return true;
    }

    /* ── Validação: confirmação de senha (onblur do confirm) ──────── */
    function setErroConfirma(mostrar) {
      const el = getEl('erro-confirma');
      if (mostrar) {
        el.classList.add('visible');
      } else {
        el.classList.remove('visible');
      }
    }

    function validarSenha() {
      const nivel = getEl('Nivel').value;
      const senha = getEl('Senha');
      const confirmaSenha = getEl('ConfirmaSenha');
      const minChar = nivel === '5' ? 24 : 10;

      if (confirmaSenha.value.length === 0) {
        setErroConfirma(false);
        return false;
      }

      if (senha.value !== confirmaSenha.value) {
        setInvalid(confirmaSenha, true);
        setErroConfirma(true);
        return false;
      }

      if (senha.value.length < minChar) {
        setInvalid(senha, true);
        setErroConfirma(false);
        return false;
      }

      setInvalid(confirmaSenha, false);
      setErroConfirma(false);
      return true;
    }

    /* ── Envio do formulário ──────────────────────────────────────── */
    function enviarFormulario() {
      const nivel        = getEl('Nivel').value;
      const senhaAtual   = getEl('SenhaAtual');
      const senha        = getEl('Senha');
      const confirmaSenha = getEl('ConfirmaSenha');
      const minChar      = nivel === '5' ? 24 : 10;
      let valido         = true;

      /* 1. Senha atual obrigatória */
      if (!senhaAtual.value.trim()) {
        setInvalid(senhaAtual, true);
        if (valido) senhaAtual.focus();
        valido = false;
      } else {
        setInvalid(senhaAtual, false);
      }

      /* 2. Nova senha: tamanho mínimo */
      if (senha.value.length < minChar) {
        setInvalid(senha, true);
        if (valido) senha.focus();
        valido = false;
      } else {
        setInvalid(senha, false);
      }

      /* 3. Confirmação: preenchida e igual à nova senha */
      if (confirmaSenha.value.length === 0 || confirmaSenha.value !== senha.value) {
        setInvalid(confirmaSenha, true);
        setErroConfirma(true);
        if (valido) confirmaSenha.focus();
        valido = false;
      } else {
        setInvalid(confirmaSenha, false);
        setErroConfirma(false);
      }

      return valido;
    }

    /* ── Sair ─────────────────────────────────────────────────────── */
    function Sair() {
      const { protocol, host } = window.location;
      window.open(protocol + '//' + host, '_top');
    }
  </script>

</body>
</html>
