<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
' ── Cabeçalhos HTTP ─────────────────────────────────────────────────────────
Response.Buffer  = True

' ── Tela de expiração de sessão (sem sidebar) ────────────────────────────────
If SESSAO_EXPIRADA Then
  Response.Write "" &_
    "<!DOCTYPE html>" &_
    "<html lang='pt-BR'>" &_
      "<head>" &_
        "<meta charset='utf-8' />" &_
        "<title>SWnet</title>" &_
        "<meta name='viewport' content='width=device-width, initial-scale=1' />" &_
        "<meta http-equiv='PRAGMA' content='no-cache' />" &_
        "<meta http-equiv='cache-control' content='no-cache' />" &_
        "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css' />" &_
        "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap-icons@1.13.1/font/bootstrap-icons.min.css' />" &_
        "<style>" &_
          ":root {" &_
            "--color-bg:          hsl(240, 9%, 98%);" &_
            "--color-surface:     hsl(0, 0%, 100%);" &_
            "--color-border:      hsl(240, 6%, 90%);" &_
            "--color-text:        hsl(240, 6%, 10%);" &_
            "--color-text-muted:  hsl(240, 4%, 46%);" &_
            "--color-accent:      hsl(210, 100%, 45%);" &_
            "--color-accent-soft: hsl(210, 100%, 96%);" &_
            "--color-brand:       hsl(220, 72%, 18%);" &_
            "--font-sans: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Inter', Roboto, Helvetica, Arial, sans-serif;" &_
            "--radius-sm: 6px;" &_
            "--radius-lg: 14px;" &_
            "--shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);" &_
            "--transition: 160ms ease;" &_
          "}" &_
          "*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }" &_
          "body {" &_
            "background: var(--color-bg);" &_
            "font-family: var(--font-sans);" &_
            "min-height: 100vh;" &_
            "display: flex;" &_
            "align-items: center;" &_
            "justify-content: center;" &_
            "padding: 1.5rem;" &_
          "}" &_
          ".sw-shell {" &_
            "display: flex;" &_
            "border-radius: var(--radius-lg);" &_
            "overflow: hidden;" &_
            "border: 1px solid var(--color-border);" &_
            "box-shadow: var(--shadow-md);" &_
            "width: 100%;" &_
            "max-width: 680px;" &_
            "min-height: 400px;" &_
          "}" &_
          ".sw-left {" &_
            "width: 240px;" &_
            "flex-shrink: 0;" &_
            "background: var(--color-brand);" &_
            "padding: 2rem 1.75rem;" &_
            "display: flex;" &_
            "flex-direction: column;" &_
            "justify-content: space-between;" &_
            "position: relative;" &_
            "overflow: hidden;" &_
          "}" &_
          ".sw-left::after {" &_
            "content: '';" &_
            "position: absolute;" &_
            "bottom: -60px; right: -60px;" &_
            "width: 220px; height: 220px;" &_
            "border-radius: 50%;" &_
            "border: 40px solid hsla(210, 80%, 60%, 0.12);" &_
            "pointer-events: none;" &_
          "}" &_
          ".sw-left::before {" &_
            "content: '';" &_
            "position: absolute;" &_
            "bottom: 30px; right: -40px;" &_
            "width: 160px; height: 160px;" &_
            "border-radius: 50%;" &_
            "border: 28px solid hsla(210, 80%, 60%, 0.08);" &_
            "pointer-events: none;" &_
          "}" &_
          ".sw-logo-wrap {" &_
            "background: hsl(0, 0%, 100%);" &_
            "border-radius: 10px;" &_
            "padding: 0.6rem 0.75rem;" &_
            "display: inline-flex;" &_
            "align-items: center;" &_
          "}" &_
          ".sw-logo-wrap img { display: block; width: 140px; height: auto; }" &_
          ".sw-left-footer {" &_
            "font-size: 0.68rem;" &_
            "color: hsla(0, 0%, 100%, 0.4);" &_
            "line-height: 1.5;" &_
          "}" &_
          ".sw-right {" &_
            "flex: 1;" &_
            "background: var(--color-surface);" &_
            "display: flex;" &_
            "align-items: center;" &_
            "justify-content: center;" &_
            "padding: 2.5rem 2rem;" &_
          "}" &_
          ".sw-content { max-width: 280px; text-align: center; }" &_
          ".sw-icon-wrap {" &_
            "width: 64px; height: 64px;" &_
            "border-radius: 50%;" &_
            "background: var(--color-accent-soft);" &_
            "display: flex;" &_
            "align-items: center;" &_
            "justify-content: center;" &_
            "margin: 0 auto 1.5rem;" &_
          "}" &_
          ".sw-icon-wrap .bi { font-size: 1.75rem; color: var(--color-accent); }" &_
          ".sw-title {" &_
            "font-size: 1.2rem;" &_
            "font-weight: 600;" &_
            "color: var(--color-text);" &_
            "margin-bottom: 0.5rem;" &_
            "line-height: 1.3;" &_
          "}" &_
          ".sw-body {" &_
            "font-size: 0.82rem;" &_
            "color: var(--color-text-muted);" &_
            "line-height: 1.6;" &_
            "margin-bottom: 1.75rem;" &_
          "}" &_
          ".sw-btn {" &_
            "display: inline-flex;" &_
            "align-items: center;" &_
            "gap: 0.4rem;" &_
            "background: var(--color-brand);" &_
            "color: hsl(0, 0%, 100%);" &_
            "font-size: 0.82rem;" &_
            "font-weight: 500;" &_
            "padding: 0.6rem 1.4rem;" &_
            "border-radius: var(--radius-sm);" &_
            "border: none;" &_
            "text-decoration: none;" &_
            "cursor: pointer;" &_
            "transition: opacity var(--transition);" &_
            "min-height: 44px;" &_
          "}" &_
          ".sw-btn:hover { opacity: 0.85; color: hsl(0, 0%, 100%); }" &_
          ".sw-btn:focus-visible { outline: none; box-shadow: 0 0 0 3px hsla(210, 100%, 45%, 0.35); }" &_
          "@media (prefers-reduced-motion: reduce) { * { transition: none !important; } }" &_
          "@media (max-width: 540px) {" &_
            ".sw-shell { flex-direction: column; }" &_
            ".sw-left { width: 100%; min-height: auto; padding: 1.5rem; }" &_
            ".sw-left::after, .sw-left::before { display: none; }" &_
          "}" &_
        "</style>" &_
      "</head>" &_
      "<body>" &_
        "<script>function sair() { window.top.location.href = 'http://' + location.hostname; }</script>" &_
        "<div class='sw-shell'>" &_
          "<div class='sw-left'>" &_
            "<div>" &_
              "<div class='sw-logo-wrap'>" &_
                "<img src='http://" & Request.ServerVariables("SERVER_NAME") & "/corporativo/images/SWlogo500w.png' alt='Smurfit Westrock' />" &_
              "</div>" &_
            "</div>" &_
            "<div class='sw-left-footer' style='color:white'>Smurfit Westrock Brasil</div>" &_
          "</div>" &_
          "<div class='sw-right'>" &_
            "<div class='sw-content'>" &_
              "<div class='sw-icon-wrap'>" &_
                "<i class='bi bi-lock' aria-hidden='true'></i>" &_
              "</div>" &_
              "<h1 class='sw-title'>Sua sess&#227;o expirou</h1>" &_
              "<p class='sw-body'>Efetue novamente o login para continuar utilizando o sistema.</p>" &_
              "<a href='#' onclick='sair()' class='sw-btn'>" &_
                "<i class='bi bi-box-arrow-in-right me-1' aria-hidden='true'></i>" &_
                "Ir para o login" &_
              "</a>" &_
            "</div>" &_
          "</div>" &_
        "</div>" &_
      "</body>" &_
    "</html>"
  Response.End
End If

' ── Resposta JSON para o polling de sessão ───────────────────────────────────
If Request.QueryString("check") = "1" Then
  Response.ContentType = "application/json"
  Response.Expires     = 0
  Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
  Response.AddHeader "Pragma", "no-cache"
  If SESSAO_EXPIRADA Then
    Response.Write "{""expirada"": true}"
  Else
    Response.Write "{""expirada"": false}"
  End If
  Response.End
End If

Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Cache-Control", "post-check=0, pre-check=0"

Session.LCID = 1046

' ── Helpers SQL ─────────────────────────────────────────────────────────────
Function SqlSafe(ByVal v)
  SqlSafe = Replace(Trim(CStr(v & "")), "'", "''")
End Function

Function nomeFilial(ByVal codigoFilial)
  Dim Retorno
  Retorno = ""
  Select Case CStr(codigoFilial & "")
    Case "00": Retorno = "Pirapetinga"
    Case "01": Retorno = "Rio de Janeiro"
    Case "02": Retorno = "São Paulo"
    Case "03": Retorno = "Corporativo"
    Case "04": Retorno = "Uberaba"
    Case "05": Retorno = "Bento Gonçalves"
    Case "06": Retorno = "Fortaleza"
    Case "07": Retorno = "Saquarema"
    Case "08": Retorno = "Aracatuba"
    Case "09": Retorno = "Blumenau"
    Case "10": Retorno = "Pacajus"
    Case "11": Retorno = "Porto Feliz"
    Case "12": Retorno = "Tres Barras"
    Case "13": Retorno = "Florestal"
  End Select
  nomeFilial = Retorno
End Function

' ── Dados de sessão ──────────────────────────────────────────────────────────
Dim idUsuarioLogado, nomeUsuario
idUsuarioLogado = Session("CODIGO")
nomeUsuario     = Session("NOME")

Dim codUserSafe, filialSafe
codUserSafe = SqlSafe(CODUSER)
filialSafe  = SqlSafe(UNIDADE_FILIAL)

' ── Queries do menu ──────────────────────────────────────────────────────────
On Error Resume Next

Dim SQL, PIR_MNUSRV, PIR_MNURLT

SQL = "SELECT " & _
        iif((CODUSER="101100"), "DISTINCT LINKS.LIN_COD, ", "MONTAMENU.CRM_COD AS CODUSER, ") & _
        "LINKS.LIN_NOME AS LINK, " & _
        "SERVICOS.SRV_DESC AS DESCSERV, " & _
        "LINKS.LIN_URL AS URL, " & _
        "LINKS.LIN_EXTEN AS EXTPAG, " & _
        "MONTAMENU.LIN_COD AS LINCOD, " & _
        "LINKS.LIN_DESC AS LINDESC, " & _
        "SERVICOS.SRV_CODSERV AS CODSERV, " & _
        "SUBMENU.SUB_MENU AS SBDESC, " & _
        "LINKS.SUB_COD AS SBCOD " & _
      "FROM MONTAMENU " & _
        "INNER JOIN LINKS ON MONTAMENU.LIN_COD = LINKS.LIN_COD " & _
        "INNER JOIN SERVICOS ON MONTAMENU.SRV_CODSERV = SERVICOS.SRV_CODSERV " & _
        "INNER JOIN SUBMENU ON LINKS.SUB_COD = SUBMENU.SUB_COD " & _
      "WHERE " & _
        "(MONTAMENU.CRM_COD = '" & codUserSafe & "') AND " & _
        "(MONTAMENU.FILIAL = '" & filialSafe & "') " & _
        "AND (MONTAMENU.MON_D_E_L_E_T <> N'*') " & _
        "AND (LINKS.LIN_D_E_L_E_T <> N'*') " & _
        "AND (SERVICOS.SRV_D_E_L_E_T <> N'*') " & _
      "ORDER BY SBDESC, LINK "

Set PIR_MNUSRV = Conecta("PIR-HELPDESK", SQL)
If (Not IsObject(PIR_MNUSRV)) Or (PIR_MNUSRV Is Nothing) Then
  Response.End
End If

SQL = "SELECT " & _
        iif((CODUSER="101100"), "DISTINCT RELATORIOSCAPA.REL_COD, ", "MONTARELAT.SRV_CODSERV AS CODSERV, ") & _
        "MONTARELAT.REL_COD AS CODRELATORIO, " & _
        "RELATORIOSCAPA.REL_NOME AS NOMERELAT, " & _
        "RELATORIOSCAPA.REL_URL AS URL, " & _
        "RELATORIOSCAPA.REL_EXTEN AS EXTEN, " & _
        "RELATORIOSCAPA.REL_BANCO AS BANCO, " & _
        "RELATORIOSCAPA.REL_QUERY AS DESCRI, " & _
        "CAT_RELATORIOS.CAT_NOME AS CATNOME, " & _
        "CAT_RELATORIOS.CAT_COD AS CATCOD " & _
      "FROM MONTARELAT " & _
        "INNER JOIN RELATORIOSCAPA ON MONTARELAT.REL_COD = RELATORIOSCAPA.REL_COD " & _
        "INNER JOIN CAT_RELATORIOS ON RELATORIOSCAPA.CAT_COD = CAT_RELATORIOS.CAT_COD " & _
      "WHERE " & _
        "(MONTARELAT.CRM_COD = '" & codUserSafe & "') AND " & _
        "(MONTARELAT.FILIAL = '" & filialSafe & "') " & _
        "AND (MONTARELAT.MON_D_E_L_E_T_E_D <> '*') " & _
      "ORDER BY CAT_NOME, REL_NOME"

Set PIR_MNURLT = Conecta("PIR-HELPDESK", SQL)
If (Not IsObject(PIR_MNURLT)) Or (PIR_MNURLT Is Nothing) Then
  Response.End
End If

' ── Serializa serviços em JSON ────────────────────────────────────────────────
Dim jsonServicos
jsonServicos = "["
Dim curGrupoSrv, gSrv, urlSrv, linkSrv, primeiroGrupoSrv, primeiroItemSrv
curGrupoSrv      = ""
primeiroGrupoSrv = True
primeiroItemSrv  = True
While Not PIR_MNUSRV.EOF
  gSrv    = Trim(PIR_MNUSRV("SBDESC"))
  urlSrv  = Trim(PIR_MNUSRV("URL")) & Trim(PIR_MNUSRV("EXTPAG"))
  linkSrv = Trim(PIR_MNUSRV("LINK"))
  If curGrupoSrv <> gSrv Then
    If curGrupoSrv <> "" Then jsonServicos = jsonServicos & "]},"
    jsonServicos = jsonServicos & "{""grupo"":""" & Replace(gSrv, """", "\""") & """,""itens"":["
    curGrupoSrv      = gSrv
    primeiroGrupoSrv = False
    primeiroItemSrv  = True
  End If
  If Not primeiroItemSrv Then jsonServicos = jsonServicos & ","
  jsonServicos = jsonServicos & "{""nome"":""" & Replace(linkSrv, """", "\""") & """,""url"":""" & Replace(urlSrv, """", "\""") & """}"
  primeiroItemSrv = False
  PIR_MNUSRV.MoveNext()
Wend
If curGrupoSrv <> "" Then jsonServicos = jsonServicos & "]}"
jsonServicos = jsonServicos & "]"
PIR_MNUSRV.Close()
Set PIR_MNUSRV = Nothing

' ── Serializa consultas/relatórios em JSON ────────────────────────────────────
Dim jsonConsultas
jsonConsultas = "["
Dim curGrupoRel, gRel, urlRel, nomeRel, primeiroGrupoRel, primeiroItemRel
curGrupoRel      = ""
primeiroGrupoRel = True
primeiroItemRel  = True
While Not PIR_MNURLT.EOF
  gRel    = Trim(PIR_MNURLT("CATNOME"))
  urlRel  = Trim(PIR_MNURLT("URL")) & Trim(PIR_MNURLT("EXTEN"))
  nomeRel = Trim(PIR_MNURLT("NOMERELAT"))
  If curGrupoRel <> gRel Then
    If curGrupoRel <> "" Then jsonConsultas = jsonConsultas & "]},"
    jsonConsultas = jsonConsultas & "{""grupo"":""" & Replace(gRel, """", "\""") & """,""itens"":["
    curGrupoRel      = gRel
    primeiroGrupoRel = False
    primeiroItemRel  = True
  End If
  If Not primeiroItemRel Then jsonConsultas = jsonConsultas & ","
  jsonConsultas = jsonConsultas & "{""nome"":""" & Replace(nomeRel, """", "\""") & """,""url"":""" & Replace(urlRel, """", "\""") & """}"
  primeiroItemRel = False
  PIR_MNURLT.MoveNext()
Wend
If curGrupoRel <> "" Then jsonConsultas = jsonConsultas & "]}"
jsonConsultas = jsonConsultas & "]"
PIR_MNURLT.Close()
Set PIR_MNURLT = Nothing

Dim nomeFilialAtual, nomeUsuarioSessao
nomeFilialAtual   = nomeFilial(UNIDADE_FILIAL)
nomeUsuarioSessao = Session("USR")
%>
<!doctype html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Smurfit Westrock</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="apple-touch-icon" sizes="180x180" href="../images/favicon/apple-touch-icon.png" />
    <link rel="icon" type="image/png" sizes="32x32" href="../images/favicon/favicon-32x32.png" />
    <link rel="icon" type="image/png" sizes="16x16" href="../images/favicon/favicon-16x16.png" />

    <link
      rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
      integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH"
      crossorigin="anonymous"
    />
    <link
      rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.13.1/font/bootstrap-icons.min.css"
    />

    <style>
      /* ── Reset global ───────────────────────────────────────────────────── */
      html, body {
        height: 100%;
        margin: 0;
        padding: 0;
        overflow: hidden;
      }

      /* ── Select2 — estilos globais (dropdown renderizado no body do pai) ── */
      .select2-container--open,
      .select2-dropdown {
        z-index: 999999 !important;
      }
      .select2-container--bootstrap-5 .select2-dropdown {
        border: 1px solid #e8e8e8;
        border-radius: 0.85rem;
        overflow: hidden;
        background: #ffffff;
      }
      .select2-search--dropdown {
        padding: 0.6rem 0.75rem;
        border-bottom: 1px solid #e8e8e8;
      }
      .select2-search__field {
        border: 1px solid #e8e8e8 !important;
        border-radius: 0.75rem !important;
        padding: 0.5rem 0.65rem !important;
        box-shadow: none !important;
        font-size: 0.85rem !important;
      }
      .select2-container--bootstrap-5 .select2-results__options {
        padding: 0.35rem 0.35rem 0.5rem;
        max-height: 70vh !important;
        overflow-y: auto !important;
        -webkit-overflow-scrolling: touch;
      }
      .select2-results__option ul,
      .select2-results__option .select2-results__options,
      .select2-results__option--group ul {
        max-height: none !important;
        overflow: visible !important;
      }
      .select2-results__group {
        background: #2f2f2f;
        color: #ffffff !important;
        font-weight: 700;
        text-transform: uppercase;
        margin: 0.55rem 0.25rem 0.35rem;
        border-radius: 0.65rem;
        box-shadow: 0 6px 14px rgba(0, 0, 0, 0.12);
        font-size: 0.7rem !important;
        padding: 0.35rem 0.55rem !important;
      }
      .select2-results__option {
        border-radius: 0.65rem;
        margin: 0.2rem 0.15rem;
        font-size: 0.82rem !important;
        padding: 0.35rem 0.55rem !important;
      }
      .select2-results__option--highlighted {
        background: #f2f2f2 !important;
        color: #111 !important;
      }
      .select2-results__option--selected {
        background: #e9ecef !important;
        font-weight: 600;
      }

      /* ── Layout principal ───────────────────────────────────────────────── */
      .app-shell {
        height: 100vh;
        display: flex;
        flex-direction: row;
        overflow: hidden;
        background: #ffffff;
      }

      /* ── Shell-nav: wrapper da sidebar inline ───────────────────────────── */
      .shell-nav {
        flex: 0 0 auto;
        height: 100%;
        display: flex;
        flex-direction: column;
        transition: background 0.15s ease;
        position: relative;
        z-index: 100;
      }
      /* Dark mode propagado ao shell-nav pelo JS (data-theme no aside) */
      .shell-nav[data-theme="dark"] { background: hsl(240 10% 3.9%); }
      .shell-nav[data-theme="light"] { background: hsl(0 0% 100%); }
      /* shell-nav espelha a largura da sidebar com as mesmas transicoes */
      .shell-nav { width: 256px; transition: width 0.30s cubic-bezier(0.4, 0, 0.2, 1), background 0.15s ease; }
      .shell-nav.is-collapsed  { width: 52px; }
      .shell-nav.is-collapsing { transition: width 0.85s cubic-bezier(0.4, 0, 0.2, 1), background 0.15s ease; }
      .shell-nav.is-hover-expanded { width: 256px; }

      /* ── Shell-mid: iframe de conteúdo (principal_mid.asp) ─────────────── */
      .shell-mid {
        flex: 1 1 auto;
        min-width: 0;
        height: 100%;
        overflow: hidden;
        background: #ffffff;
      }
      .shell-frame {
        width: 100%;
        height: 100%;
        border: 0;
        display: block;
      }

      @media print {
        .shell-nav { display: none !important; }
      }

      /* ================================================================
         SIDEBAR (portado de navbar_left.asp)
         ================================================================ */

      /* ── Tokens shadcn/ui ─────────────────────────────────────────── */
      .sidebar {
        --background:         0 0% 100%;
        --foreground:         240 10% 3.9%;
        --muted:              240 4.8% 95.9%;
        --muted-foreground:   240 3.8% 46.1%;
        --accent:             240 4.8% 95.9%;
        --accent-foreground:  240 5.9% 10%;
        --primary:            240 5.9% 10%;
        --primary-foreground: 0 0% 98%;
        --border:             240 5.9% 90%;
        --input:              240 5.9% 90%;
        --ring:               240 5.9% 10%;
        --destructive:        0 72% 51%;
        --radius:             0.5rem;
      }
      .sidebar[data-theme="dark"] {
        --background:         240 10% 3.9%;
        --foreground:         0 0% 98%;
        --muted:              240 3.7% 15.9%;
        --muted-foreground:   240 5% 64.9%;
        --accent:             240 3.7% 15.9%;
        --accent-foreground:  0 0% 98%;
        --primary:            0 0% 98%;
        --primary-foreground: 240 5.9% 10%;
        --border:             240 3.7% 15.9%;
        --input:              240 3.7% 15.9%;
        --ring:               240 4.9% 83.9%;
        --destructive:        0 72% 60%;
      }

      * { box-sizing: border-box; }

      .is-hidden { display: none !important; }
      /* Context oculto na página inicial (elemento independente do footer) */
      .context.context-home-hidden { display: none !important; }

      /* Sidebar */
      .sidebar {
        width: 256px;
        height: 100%;
        display: flex;
        flex-direction: column;
        background: hsl(var(--background));
        color: hsl(var(--foreground));
        border-right: 1px solid hsl(var(--border));
        font-size: 0.875rem;
        font-family: system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
        overflow: hidden;
        transition: width 0.30s cubic-bezier(0.4, 0, 0.2, 1), background 0.15s ease, color 0.15s ease;
      }
      /* Colapsado: rail de ícones de 52px */
      .sidebar.is-collapsed { width: 52px; }
      /* Transicao lenta apenas ao recolher */
      .sidebar.is-collapsing {
        transition: width 0.85s cubic-bezier(0.4, 0, 0.2, 1), background 0.15s ease, color 0.15s ease;
      }
      /* Header colapsado: só toggle centralizado */
      .sidebar.is-collapsed .sidebar-header {
        justify-content: center;
        padding: 0;
      }
      .sidebar.is-collapsed .brand { display: none; }
      .sidebar.is-collapsed .toggle-area {
        width: 100%; border-left: 0;
        padding: 10px 0;
      }

      /* Header */
      .sidebar-header {
        display: flex; align-items: center; gap: 6px;
        padding: 8px;
        flex-shrink: 0; min-height: 68px;
      }
      .brand {
        flex: 1 1 auto; min-width: 0; display: flex; align-items: center;
        border: 0; background: transparent; cursor: pointer; padding: 4px;
        border-radius: var(--radius);
        overflow: hidden;
      }
      /* 14) Sem highlight de fundo no hover — apenas a logo escala suavemente */
      .brand:hover { background: transparent; }
      .brand-plate {
        flex: 1 1 auto; width: 100%; height: 52px;
        display: flex; align-items: center; justify-content: center;
        overflow: hidden; padding: 4px 6px;
      }
      .brand-img {
        height: 40px; width: auto; max-width: 100%;
        object-fit: contain; display: block;
        transition: transform 0.25s ease;
      }
      .brand:hover  .brand-img { transform: scale(1.08); }
      .brand:not(:hover) .brand-img { transform: scale(1); }
      .brand-img-dark { display: none; }
      .sidebar[data-theme="dark"] .brand-img-light { display: none; }
      .sidebar[data-theme="dark"] .brand-img-dark  { display: block; }

      /* Toggle */
      .toggle-area {
        flex-shrink: 0; display: flex; align-items: center; justify-content: center;
        width: 44px;
        align-self: stretch;
      }
      .toggle-btn { flex-shrink: 0; }

      /* btnToggle — troca ícone via JS conforme estado; visual = btn-icon padrão */
      #btnToggle svg { display: block; }

      /* Pulso do unpin quando sidebar está recolhida */
      @keyframes pulse-unpin {
        0%, 100% { opacity: 1;   transform: scale(1); }
        50%       { opacity: 0.5; transform: scale(0.82); }
      }
      .sidebar.is-collapsed #btnToggle i {
        animation: pulse-unpin 2s ease-in-out infinite;
      }
      .sidebar.is-collapsed:hover #btnToggle i {
        animation: none;
        opacity: 1;
        transform: scale(1);
      }

      /* Botões icon */
      .btn-icon {
        display: inline-flex; align-items: center; justify-content: center;
        height: 32px; width: 32px; min-width: 32px; padding: 0;
        border: 0; background: transparent; color: hsl(var(--foreground));
        border-radius: var(--radius); cursor: pointer;
        transition: background 0.15s ease, color 0.15s ease; flex-shrink: 0;
      }
      .btn-icon:hover { background: hsl(var(--accent)); color: hsl(var(--accent-foreground)); }
      .btn-icon:focus-visible { outline: none; box-shadow: 0 0 0 2px hsl(var(--background)), 0 0 0 4px hsl(var(--ring)); }
      .btn-destructive { color: hsl(var(--destructive)); }
      .btn-destructive:hover { color: #fff; background: hsl(var(--destructive)); }
      .sidebar[data-theme="dark"] .btn-destructive:hover { background: #fff; color: hsl(var(--destructive)); }
      .btn-icon-outline { border: 1px solid hsl(var(--border)); }
      .icon { width: 16px; height: 16px; display: block; flex: 0 0 auto; }

      /* Busca */
      .sidebar-search { padding: 8px; flex-shrink: 0; }
      .search-input {
        display: flex; align-items: center; gap: 8px; height: 36px; padding: 0 10px;
        border: 1px solid hsl(var(--input)); border-radius: var(--radius);
        background: hsl(var(--muted));
        animation: cardBorderLight 2s ease-in-out infinite;
      }
      .sidebar[data-theme="dark"] .search-input { animation: cardBorderDark 2s ease-in-out infinite; }
      .search-input:focus-within { animation: none; border-color: hsl(var(--ring)); box-shadow: 0 0 0 2px hsl(var(--ring) / 0.25); }
      .search-input .icon { color: hsl(var(--muted-foreground)); flex-shrink: 0; }
      #filtro { flex: 1 1 auto; min-width: 0; border: 0; outline: 0; background: transparent; color: hsl(var(--foreground)); font-size: 0.8125rem; }
      #filtro::placeholder { color: hsl(var(--muted-foreground)); }
      .kbd {
        display: inline-flex; align-items: center; height: 20px; padding: 0 6px;
        font-size: 0.6875rem; font-family: inherit; color: hsl(var(--muted-foreground));
        background: hsl(var(--muted)); border: 1px solid hsl(var(--border)); border-radius: 4px;
        white-space: nowrap; flex-shrink: 0;
      }
      /* Navegação */
      .sidebar-content {
        flex: 1 1 auto; min-height: 0; overflow-y: auto; overflow-x: hidden;
        padding: 4px 8px 8px;
        scrollbar-width: thin; scrollbar-color: hsl(var(--border)) transparent;
      }
      .sidebar-content::-webkit-scrollbar { width: 10px; }
      .sidebar-content::-webkit-scrollbar-thumb { background: hsl(var(--border)); border-radius: 8px; border: 3px solid hsl(var(--background)); }

      /* Seções */
      .section + .section { margin-top: 18px; }
      .section-label {
        width: 100%; display: flex; align-items: center; gap: 6px;
        height: 30px; padding: 0 8px; margin-top: 6px;
        border: 0; background: transparent; cursor: pointer;
        border-radius: var(--radius); font-size: 0.72rem; font-weight: 600;
        letter-spacing: 0.04em; text-transform: uppercase;
        color: hsl(var(--muted-foreground)); white-space: nowrap;
      }
      /* 18) "Serviços" — label do dropdown em azul HSL */
      .section--servico  .section-label { color: hsl(217 91% 45%); }
      .section--consulta .section-label { color: #f97316; }
      /* Dark mode: "Serviços" azul claro HSL; "Consultas" laranja */
      .sidebar[data-theme="dark"] .section--servico  .section-label { color: hsl(213 94% 68%); }
      .sidebar[data-theme="dark"] .section--consulta .section-label { color: #fb923c; }
      .section-label-txt { flex: 1 1 auto; text-align: left; overflow: hidden; text-overflow: ellipsis; }
      .section-chevron { display: flex; transition: transform 0.2s ease; transform: rotate(90deg); }
      .section:not(.is-open) .section-chevron { transform: rotate(0deg); }
      .section-toggle-all {
        flex-shrink: 0; display: flex; align-items: center; justify-content: center;
        width: 20px; height: 20px; padding: 0; border: 0; border-radius: 4px;
        background: transparent; cursor: pointer; color: inherit; opacity: 0.45;
        transition: opacity 0.15s ease, background 0.15s ease;
      }
      .section-toggle-all:hover { opacity: 1; background: hsl(var(--accent)); }
      .sidebar[data-theme="dark"] .section-toggle-all:hover { background: hsl(var(--accent)); }
      .sidebar.is-collapsed .section-toggle-all { display: none; }
      .section-body { max-height: 4000px; overflow: hidden; transition: max-height 0.2s ease; }
      .section:not(.is-open) .section-body { max-height: 0; }
      .nav-empty { display: none; padding: 16px 12px; color: hsl(var(--muted-foreground)); font-size: 0.8125rem; text-align: center; }
      .nav-empty.is-visible { display: block; }

      /* Grupos */
      .group-header {
        width: 100%; display: flex; align-items: center; gap: 8px; height: 32px; padding: 0 8px;
        border: 0; background: transparent; color: hsl(var(--foreground));
        border-radius: var(--radius); cursor: pointer; text-align: left;
        transition: background 0.15s ease, color 0.15s ease;
      }
      .group-header:hover { background: hsl(var(--accent)); color: hsl(var(--accent-foreground)); }
      .group.has-active > .group-header,
      .group.has-active > .group-header:hover {
        background: hsl(240 5% 88%);
        color: hsl(var(--accent-foreground));
      }
      .sidebar[data-theme="dark"] .group.has-active > .group-header,
      .sidebar[data-theme="dark"] .group.has-active > .group-header:hover {
        background: hsl(240 4% 22%);
        color: hsl(var(--accent-foreground));
      }
      .group-header:focus-visible { outline: none; box-shadow: 0 0 0 2px hsl(var(--ring) / 0.5); }
      .group-icon    { display: flex; color: hsl(var(--muted-foreground)); flex-shrink: 0; }
      .group-label-txt { flex: 1 1 auto; font-size: 0.875rem; font-weight: 400; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .group-badge   { font-size: 0.6875rem; color: hsl(var(--muted-foreground)); font-variant-numeric: tabular-nums; padding: 0 2px; flex-shrink: 0; }
      .group-chevron { display: flex; color: hsl(var(--muted-foreground)); transition: transform 0.15s ease; flex-shrink: 0; }
      .group.is-open .group-chevron { transform: rotate(90deg); }
      .group-body { max-height: 0; overflow: hidden; transition: max-height 0.2s ease; }
      .group.is-open .group-body { max-height: 3000px; }
      .group-sub { margin: 2px 0 4px 15px; padding-left: 10px; border-left: 1px solid hsl(var(--border)); display: flex; flex-direction: column; gap: 1px; }

      /* ── Bolinha pulsante no item ativo (sobre a linha vertical do group-sub) */
      .group-sub { margin: 2px 0 4px 15px; padding-left: 10px; border-left: 1px solid hsl(var(--border)); display: flex; flex-direction: column; gap: 1px; position: relative; }

      /* O item ativo gera uma bolinha sobre a linha vertical via ::before */
      .item.is-active { position: relative; }
      .item.is-active::before {
        content: "";
        position: absolute;
        left: -15px;          /* alinha sobre o border-left do group-sub */
        top: 50%;
        transform: translateY(-50%);
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: #18181b;  /* light mode: preta */
        animation: dotPulseLight 1.6s ease-in-out infinite;
        z-index: 1;
      }
      .sidebar[data-theme="dark"] .item.is-active::before {
        background: #ffffff;  /* dark mode: branca */
        animation: dotPulseDark 1.6s ease-in-out infinite;
      }

      @keyframes dotPulseLight {
        0%, 100% { box-shadow: 0 0 0 0 rgba(24,24,27,0.6); }
        50%       { box-shadow: 0 0 0 5px rgba(24,24,27,0); }
      }
      @keyframes dotPulseDark {
        0%, 100% { box-shadow: 0 0 0 0 rgba(255,255,255,0.7); }
        50%       { box-shadow: 0 0 0 5px rgba(255,255,255,0); }
      }

      /* Itens */
      .item {
        width: 100%; display: flex; align-items: center; min-height: 28px; padding: 5px 8px;
        border: 0; background: transparent; color: hsl(var(--muted-foreground));
        border-radius: var(--radius); cursor: pointer; text-align: left; line-height: 1.3;
        transition: background 0.15s ease, color 0.15s ease;
      }
      .item:hover       { background: hsl(var(--accent)); color: hsl(var(--accent-foreground)); }
      .item:focus-visible { outline: none; box-shadow: 0 0 0 2px hsl(var(--ring) / 0.5); }
      .item.is-active   { background: hsl(var(--accent)); color: hsl(var(--accent-foreground)); font-weight: 400; }
      .item-label { flex: 1 1 auto; font-size: 0.8125rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .mark { background: hsl(var(--foreground) / 0.14); border-radius: 3px; color: inherit; padding: 0 1px; }

      /* ── Expansão temporária via hover (sidebar recolhida) ──────────────── */

      /* Wrapper: sobrepõe o conteúdo sem empurrar o layout */
      .sidebar.is-collapsed.is-hover-expanded {
        width: 256px !important;
        box-shadow: 4px 0 24px rgba(0,0,0,0.12);
      }

      /* Restaura elementos ocultos pela rail */
      .sidebar.is-collapsed.is-hover-expanded .sidebar-search { display: block !important; }
      .sidebar.is-collapsed.is-hover-expanded .section-label  { display: flex !important; }
      .sidebar.is-collapsed.is-hover-expanded .group-label-txt,
      .sidebar.is-collapsed.is-hover-expanded .group-badge,
      .sidebar.is-collapsed.is-hover-expanded .group-chevron  { display: flex !important; }
      .sidebar.is-collapsed.is-hover-expanded .group-body     { display: block !important; }
      .sidebar.is-collapsed.is-hover-expanded .group-sub      { display: flex !important; }
      .sidebar.is-collapsed.is-hover-expanded .nav-empty      { display: none; }
      .sidebar.is-collapsed.is-hover-expanded .nav-empty.is-visible { display: block !important; }
      .sidebar.is-collapsed.is-hover-expanded .context        { display: flex !important; }
      .sidebar.is-collapsed.is-hover-expanded .context.context-home-hidden { display: none !important; }
      .sidebar.is-collapsed.is-hover-expanded .user-card      { display: flex !important; }

      /* Section-body: respeita is-open (não força aberto como a rail faz) */
      .sidebar.is-collapsed.is-hover-expanded .section-body { max-height: 4000px; transition: max-height 0.2s ease; }
      .sidebar.is-collapsed.is-hover-expanded .section:not(.is-open) .section-body { max-height: 0; }
      .sidebar.is-collapsed.is-hover-expanded .section + .section { margin-top: 18px !important; }

      /* Header: logo + toggle lado a lado, igual ao expandido */
      .sidebar.is-collapsed.is-hover-expanded .sidebar-header {
        justify-content: space-between !important;
        padding: 8px !important;
      }
      .sidebar.is-collapsed.is-hover-expanded .brand      { display: flex !important; }
      .sidebar.is-collapsed.is-hover-expanded .toggle-area {
        width: 44px !important;
        border-left: 0 !important;
        padding: 10px 0 !important;
        align-self: stretch;
      }

      /* Conteúdo: padding normal */
      .sidebar.is-collapsed.is-hover-expanded .sidebar-content {
        padding: 4px 8px 8px !important;
        overflow-y: auto;
      }

      /* Group-headers: layout normal (ícone + texto + badge + chevron) */
      .sidebar.is-collapsed.is-hover-expanded .group-header {
        justify-content: flex-start !important;
        padding: 0 8px !important;
        height: 32px !important;
        border-radius: var(--radius) !important;
      }
      .sidebar.is-collapsed.is-hover-expanded .group-icon {
        color: hsl(var(--muted-foreground)) !important;
      }

      /* Footer: coluna normal */
      .sidebar.is-collapsed.is-hover-expanded .sidebar-footer {
        flex-direction: column !important;
        align-items: stretch !important;
        padding: 8px !important;
        gap: 8px !important;
        flex: 0 0 auto !important;
      }
      .sidebar.is-collapsed.is-hover-expanded .footer-actions {
        flex-direction: row !important;
        justify-content: space-between !important;
        align-items: center !important;
        width: 100% !important;
        gap: 6px !important;
      }

      /* ── Tooltip nos group-headers da rail ──────────────────────────────── */
      .sidebar.is-collapsed:not(.is-hover-expanded) .group-header[title]:hover::after {
        content: attr(title);
        position: fixed;
        left: 60px;
        background: hsl(var(--foreground));
        color: hsl(var(--background));
        font-size: 0.75rem;
        padding: 4px 8px;
        border-radius: 4px;
        white-space: nowrap;
        pointer-events: none;
        z-index: 9999;
        box-shadow: 0 2px 8px rgba(0,0,0,0.18);
      }

      /* ── Rail colapsada (52px) ────────────────────────────────────────────
         Visível: toggle no header, ícone de cada group-header, btnTheme, btnLogout.
         Textos, badges, chevrons, corpos dos grupos e busca ficam ocultos.
         Footer sempre disponível com theme + logout centralizados.
      ────────────────────────────────────────────────────────────────────── */
      .sidebar.is-collapsed .sidebar-search,
      .sidebar.is-collapsed .section-label,
      .sidebar.is-collapsed .group-label-txt,
      .sidebar.is-collapsed .group-badge,
      .sidebar.is-collapsed .group-chevron,
      .sidebar.is-collapsed .group-body,
      .sidebar.is-collapsed .group-sub,
      .sidebar.is-collapsed .nav-empty,
      .sidebar.is-collapsed .context,
      .sidebar.is-collapsed .user-card { display: none; }

      /* sidebar-content visível na rail (mostra os group-headers com ícones) */
      .sidebar.is-collapsed .sidebar-content {
        padding: 4px 0;
        overflow-y: auto;
        overflow-x: hidden;
      }
      /* section-body sempre aberto na rail (is-open ou não) */
      .sidebar.is-collapsed .section-body { max-height: none !important; }
      /* section sem margin extra */
      .sidebar.is-collapsed .section + .section { margin-top: 4px; }
      /* group-header centralizado — só o ícone */
      .sidebar.is-collapsed .group-header {
        justify-content: center;
        padding: 0;
        height: 40px;
        border-radius: 0;
      }
      .sidebar.is-collapsed .group-icon { color: hsl(var(--foreground)); }

      /* Footer colapsado: theme + logout em coluna centralizada */
      .sidebar.is-collapsed .sidebar-footer {
        flex: 0 0 auto;
        align-items: center;
        padding: 8px 0;
        gap: 8px;
      }
      .sidebar.is-collapsed .footer-actions {
        flex-direction: column;
        gap: 8px;
        justify-content: center;
        align-items: center;
        width: 100%;
      }

      /* Footer */
      .sidebar-footer { flex: 0 0 auto; border-top: 1px solid hsl(var(--border)); padding: 8px; display: flex; flex-direction: column; gap: 8px; }
      .context {
        display: flex; flex-direction: column; align-items: center; text-align: center;
        gap: 1px; padding: 6px 8px; border: 1px solid hsl(var(--border));
        border-radius: var(--radius); background: hsl(var(--muted)); min-height: 40px;
        animation: cardBorderLight 2s ease-in-out infinite;
        flex-shrink: 0; margin: 0 8px 0 8px;
      }
      .sidebar[data-theme="dark"] .context { animation: cardBorderDark 2s ease-in-out infinite; }
      .context-kind  { font-size: 0.625rem; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(217 91% 45%); font-weight: 700; min-height: 0.75rem; }
      .sidebar[data-theme="dark"] .context-kind { color: hsl(213 94% 68%); }
      .context-label { font-size: 0.8125rem; color: hsl(var(--foreground)); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%; }
      .user-card { position: relative; display: flex; align-items: center; gap: 8px; padding: 6px; border-radius: var(--radius); }
      .avatar {
        flex: 0 0 auto; width: 32px; height: 32px; border-radius: var(--radius);
        display: flex; align-items: center; justify-content: center;
        background: hsl(var(--primary)); color: hsl(var(--primary-foreground));
        font-size: 0.75rem; font-weight: 600; letter-spacing: -0.3px; text-transform: uppercase; flex-shrink: 0;
      }
      .user-meta  { display: flex; flex-direction: column; min-width: 0; line-height: 1.2; flex: 1 1 auto; overflow: hidden; }
      .user-name  { font-size: 0.8125rem; font-weight: 600; color: hsl(var(--foreground)); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .user-sub   { display: flex; align-items: center; font-size: 0.6875rem; color: hsl(var(--muted-foreground)); overflow: hidden; }
      .filial-dot  { flex: 0 0 auto; width: 8px; height: 8px; border-radius: 50%; margin-right: 5px; }
      .filial-nome { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .user-menu-btn { margin-left: auto; height: 28px; width: 28px; min-width: 28px; color: hsl(var(--muted-foreground)); }
      .user-menu {
        position: fixed;   /* escapa do overflow:hidden da sidebar */
        min-width: 176px; z-index: 9999;
        background: hsl(var(--background)); border: 1px solid hsl(var(--border));
        border-radius: var(--radius);
        box-shadow: 0 10px 30px rgba(0,0,0,0.18), 0 2px 6px rgba(0,0,0,0.12);
        padding: 4px;
      }
      .menu-item {
        display: flex; align-items: center; gap: 8px; width: 100%; padding: 7px 8px;
        border: 0; background: transparent; color: hsl(var(--foreground));
        border-radius: calc(var(--radius) - 2px);
        cursor: pointer; font-size: 0.8125rem; text-align: left;
        transition: background 0.12s ease;
      }
      .menu-item:hover { background: hsl(var(--accent)); }
      .menu-item .icon { color: hsl(var(--muted-foreground)); }
      .footer-actions { display: flex; align-items: center; justify-content: space-between; gap: 6px; }

      /* Animações */
      @keyframes cardBorderLight {
        0%, 100% { border-color: #d4d4d8; box-shadow: 0 0 0 0 rgba(82,82,91,0); }
        50%       { border-color: #52525b; box-shadow: 0 0 0 3px rgba(82,82,91,0.16); }
      }
      @keyframes cardBorderDark {
        0%, 100% { border-color: rgba(255,255,255,0.25); box-shadow: 0 0 0 0 rgba(255,255,255,0); }
        50%       { border-color: #fff; box-shadow: 0 0 0 3px rgba(255,255,255,0.22); }
      }
      @media (prefers-reduced-motion: reduce) { .context, .search-input { animation: none; } }

      /* Mobile */
      .topbar-mobile {
        display: none; height: 52px; align-items: center; gap: 10px; padding: 0 12px;
        background: hsl(var(--background)); border-bottom: 1px solid hsl(var(--border));
        color: hsl(var(--foreground)); flex-shrink: 0;
      }
      .backdrop { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 1030; }
      .backdrop.is-visible { display: block; }

      @media (max-width: 768px) {
        .topbar-mobile { display: flex; }
        .sidebar {
          position: fixed; top: 0; left: 0; height: 100%;
          width: 256px !important; transform: translateX(-100%);
          transition: transform 0.2s ease; z-index: 1040;
          box-shadow: 4px 0 24px rgba(0,0,0,0.18);
        }
        .sidebar.is-mobile-open { transform: translateX(0); }
        .sidebar.is-collapsed { width: 256px !important; }
        .sidebar.is-collapsed .sidebar-header { justify-content: flex-start; padding: 8px; }
        .sidebar.is-collapsed .sidebar-search,
        .sidebar.is-collapsed .section-label,
        .sidebar.is-collapsed .group-label-txt,
        .sidebar.is-collapsed .group-badge,
        .sidebar.is-collapsed .group-chevron,
        .sidebar.is-collapsed .group-body,
        .sidebar.is-collapsed .group-sub,
        .sidebar.is-collapsed .nav-empty,
        .sidebar.is-collapsed .context,
        .sidebar.is-collapsed .user-card { display: revert; }
        .sidebar.is-collapsed .section-body { max-height: 4000px !important; }
        .sidebar.is-collapsed .brand { display: flex; }
        .sidebar.is-collapsed .toggle-area { width: 44px; padding: 0; }
        .sidebar.is-collapsed .sidebar-footer { flex: 0 0 auto; justify-content: flex-start; align-items: stretch; padding: 8px; gap: 8px; }
        .sidebar.is-collapsed .footer-actions { flex-direction: row; justify-content: space-between; width: auto; }
        .toggle-btn { display: none; }
      }
    </style>
  </head>

  <body>

    <div class="app-shell">

      <!-- ================================================================
           SHELL-NAV: sidebar inline (portada de navbar_left.asp)
           ================================================================ -->
      <div class="shell-nav" id="shellNav">

        <!-- Topbar mobile (só em ≤768px) -->
        <div class="topbar-mobile" id="topbarMobile">
          <button class="btn-icon" id="btnHamburger" aria-label="Abrir menu" title="Abrir menu"></button>
          <strong>SWnet</strong>
        </div>

        <div class="backdrop" id="backdrop"></div>

        <aside class="sidebar" id="sidebar" data-theme="light" aria-label="Navegação principal">

          <!-- Header: Logo + Toggle -->
          <div class="sidebar-header">
            <button class="brand" id="btnHome"
                    title="Voltar ao Menu Principal" aria-label="Voltar ao Menu Principal">
              <span class="brand-plate">
                <img class="brand-img brand-img-light"
                     src="Logo/LOGO-LETRA-PRETA.png" alt="Smurfit Westrock" />
                <img class="brand-img brand-img-dark"
                     src="Logo/LOGO-LETRA-BRANCA.png" alt="Smurfit Westrock" />
              </span>
            </button>
            <div class="toggle-area">
              <button class="btn-icon toggle-btn" id="btnToggle"
                      title="Recolher menu" aria-label="Recolher menu"></button>
            </div>
          </div>

          <!-- Context: breadcrumb de navegação (acima do filtro) -->
          <div class="context" id="context" title="Menu Principal">
            <span class="context-kind"  id="ctxKind"></span>
            <span class="context-label" id="ctxLabel">Menu Principal</span>
          </div>

          <!-- Busca -->
          <div class="sidebar-search">
            <div class="search-input">
              <span id="searchIcon"></span>
              <input id="filtro" type="text" placeholder="Filtrar menu..."
                     autocomplete="off" spellcheck="false" aria-label="Filtrar menu" />
              <kbd class="kbd" id="searchKbd">Ctrl K</kbd>
            </div>
          </div>

          <!-- Menu dinâmico -->
          <nav class="sidebar-content" id="nav" aria-label="Menu de navegação"></nav>
          <div class="nav-empty" id="navEmpty">Nenhum item encontrado.</div>

          <!-- Footer: user-card + ações (sempre no final/baixo da sidebar) -->
          <div class="sidebar-footer" id="sidebarFooter">
            <div class="user-card">
              <span class="avatar" id="avatar"></span>
              <span class="user-meta">
                <span class="user-name" id="uUsuario"><%=nomeUsuarioSessao%></span>
                <span class="user-sub">
                  <span class="filial-dot"  id="filialDot"></span>
                  <span class="filial-nome" id="uFilial"><%=nomeFilialAtual%></span>
                </span>
              </span>
              <button class="btn-icon user-menu-btn" id="userMenuBtn"
                      title="Opções da conta" aria-label="Opções da conta"
                      aria-haspopup="true" aria-expanded="false"></button>
              <div class="user-menu is-hidden" id="userMenu" role="menu">
                <button class="menu-item" id="btnTrocarSenha" role="menuitem">
                  <span id="lockIcon"></span>Trocar sua senha
                </button>
              </div>
            </div>

            <div class="footer-actions">
              <button class="btn-icon btn-icon-outline" id="btnTheme"
                      title="Alternar tema" aria-label="Alternar tema"></button>
              <button class="btn-icon btn-destructive" id="btnLogout"
                      title="Sair" aria-label="Sair"></button>
            </div>
          </div>
        </aside>

      </div><!-- /.shell-nav -->

      <!-- ── Conteúdo principal (principal_mid.asp — mantido como iframe) ── -->
      <div class="shell-mid">
        <iframe
          class="shell-frame"
          src="principal_mid.asp"
          name="meio"
          id="meio"
          title="Conteúdo"
          scrolling="auto"
        ></iframe>
      </div>

    </div><!-- /.app-shell -->

    <!-- Modal de aviso -->
    <div class="modal fade" id="warningModal" tabindex="-1" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title">Aviso</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">
            Esta ação não é recomendada dentro do ambiente corporativo.
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-primary" data-bs-dismiss="modal">Entendi</button>
          </div>
        </div>
      </div>
    </div>

    <script
      src="https://code.jquery.com/jquery-3.7.1.min.js"
      integrity="sha256-/JqT3SQfawRcv/BIHPThkBvs0OEvtFFmqPF/lYI/Cxo="
      crossorigin="anonymous"
    ></script>
    <script
      src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"
      integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz"
      crossorigin="anonymous"
    ></script>

    <script>
    (() => {
      "use strict";

      /* ── Constantes ──────────────────────────────────────────────────────── */
      const LS_COLLAPSED     = "swnet.sidebar.collapsed";
      const LS_THEME         = "swnet.sidebar.theme";
      const URL_TROCAR_SENHA = "/corporativo/usuario/TrocarSenha.asp";

      /* ── Dados do menu (gerados pelo VBScript) ───────────────────────────── */
      const SWNET_MENU = {
        servicos:  <%=jsonServicos%>,
        consultas: <%=jsonConsultas%>
      };

      /* ── Cores das filiais ───────────────────────────────────────────────── */
      const FILIAL_COLOR = {
        "Pirapetinga":"#22c55e","Rio de Janeiro":"#ef4444","São Paulo":"#3b82f6",
        "Corporativo":"#a855f7","Uberaba":"#f97316","Bento Gonçalves":"#14b8a6",
        "Fortaleza":"#eab308","Saquarema":"#06b6d4","Aracatuba":"#ec4899",
        "Blumenau":"#8b5cf6","Pacajus":"#f59e0b","Porto Feliz":"#10b981",
        "Tres Barras":"#0ea5e9","Florestal":"#84cc16",
      };

      /* ── Ícones por grupo ────────────────────────────────────────────────── */
      const GROUP_ICON = {
        "Administrador":"shield","Cliente":"idcard","Compras":"cart",
        "Consultas":"search","Desenvolvimento":"code","Financeiro":"dollar",
        "Pedidos":"box","RH":"users","Relatórios":"barChart","SGQ":"clipboard",
        "Utilitários":"tool","Validaçoes":"checkCircle","Comercial":"briefcase",
        "Ondulado":"wave","Papel":"file",
        /* 11) Grupos com ícones dedicados */
        "Cadastros FNC":"alertTriangle",
        "Cadastros Orçamento":"calculator",
        "Cavaco/Aparas":"leaf",
        "Expedição":"truck",
        "Controles Internos":"lock",
      };

      /* ── SVG paths ───────────────────────────────────────────────────────── */
      const ICON_PATHS = {
        list:        "M4 6h16M4 12h16M4 18h16",
        file:        "M6 3h9l4 4v14H6zM15 3v4h4",
        search:      "M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-5-5",
        checkCircle: "M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM8 12l3 3 5-6",
        box:         "M3 8l9-5 9 5v8l-9 5-9-5zM3 8l9 5 9-5M12 13v8",
        coins:       "M12 4a8 8 0 1 0 0 16 8 8 0 0 0 0-16zM9 12h6M12 9v6",
        cart:        "M3 4h2l2 12h11l2-8H6M9 20a1 1 0 1 0 0 2 1 1 0 0 0 0-2zM17 20a1 1 0 1 0 0 2 1 1 0 0 0 0-2z",
        shield:      "M12 3l7 3v5c0 4.4-3 7.6-7 8.7-4-1.1-7-4.3-7-8.7V6zM9 12l2 2 4-4",
        barChart:    "M4 20V10M10 20V4M16 20v-7M3 20h18",
        users:       "M16 20v-1a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v1M9.5 11a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7M17 11a3 3 0 0 0 0-6",
        idcard:      "M3 5h18v14H3zM8 11a2 2 0 1 0 0-4 2 2 0 0 0 0 4M5 16c0-1.7 1.3-3 3-3s3 1.3 3 3M14 9h5M14 13h4",
        code:        "M8 6l-6 6 6 6M16 6l6 6-6 6",
        clipboard:   "M9 4h6v3H9zM8 4H6a1 1 0 0 0-1 1v15a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V5a1 1 0 0 0-1-1h-2M9 14l2 2 4-4",
        tool:        "M14.7 6.3a4 4 0 0 0-5.4 5.4l-6 6 3 3 6-6a4 4 0 0 0 5.4-5.4l-2.7 2.7-2.7-2.7z",
        briefcase:   "M4 8h16v11H4zM9 8V6a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2M4 13h16",
        wave:        "M2 8c2-2 4-2 6 0s4 2 6 0 4-2 6 0M2 14c2-2 4-2 6 0s4 2 6 0 4-2 6 0",
        logout:      "M9 21H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h4M16 17l5-5-5-5M21 12H9",
        panelLeft:   "M4 5h16a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1zM9 5v14",
        chevronRight:"M9 6l6 6-6 6",
        dots:        "M12 5a1 1 0 1 0 0-2 1 1 0 0 0 0 2M12 13a1 1 0 1 0 0-2 1 1 0 0 0 0 2M12 21a1 1 0 1 0 0-2 1 1 0 0 0 0 2",
        gear:        "M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z",
        lock:        "M6 10h12v11H6zM9 10V7a3 3 0 0 1 6 0v3",
        sun:         "M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM12 2v2M12 20v2M4 12H2M22 12h-2M5.6 5.6L4.2 4.2M19.8 19.8l-1.4-1.4M18.4 5.6l1.4-1.4M4.2 19.8l1.4-1.4",
        moon:        "M21 12.8A8 8 0 1 1 11.2 3a6 6 0 0 0 9.8 9.8z",
        /* 11) Novos ícones */
        alertTriangle:"M10.3 3.4L1.5 18a2 2 0 0 0 1.7 3h17.6a2 2 0 0 0 1.7-3L13.7 3.4a2 2 0 0 0-3.4 0zM12 9v4M12 17h.01",
        calculator:  "M4 2h16a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2zM8 6h2M14 6h2M8 10h2M14 10h2M8 14h2M14 14h4M14 18h4",
        leaf:        "M2 22c1-1 1.5-2 1.5-3.5C3.5 15 5 13 8 11c3-2 6-3 8-3-1 3-2 6-4 8-2.5 2.5-5 3-7.5 3.5C3 20 2.5 21 2 22zM12 8a4 4 0 0 0-4 4",
        truck:       "M1 3h15v13H1zM16 8h4l3 3v5h-7V8zM5.5 19a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3zM18.5 19a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3z",
        dollar:           "M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6",
        chevronDoubleDown:"M7 6l5 5 5-5M7 13l5 5 5-5",
        chevronDoubleUp:  "M7 18l5-5 5 5M7 11l5-5 5 5",

      };

      /* ── Helpers ─────────────────────────────────────────────────────────── */
      const $  = (sel, ctx = document) => ctx.querySelector(sel);
      const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));
      const fold = (s) => String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();

      function makeIcon(key, size, opts) {
        const NS = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(NS, "svg");
        svg.setAttribute("viewBox","0 0 24 24");
        svg.setAttribute("fill",   (opts && opts.fill)   ? opts.fill   : "none");
        svg.setAttribute("stroke", (opts && opts.stroke) ? opts.stroke : "currentColor");
        svg.setAttribute("stroke-width","2");
        svg.setAttribute("stroke-linecap","round"); svg.setAttribute("stroke-linejoin","round");
        svg.setAttribute("class","icon");
        const s = size || 16;
        svg.style.width = s+"px"; svg.style.height = s+"px";
        const p = document.createElementNS(NS,"path");
        p.setAttribute("d", ICON_PATHS[key] || ICON_PATHS.list);
        if (opts && opts.pathFill) p.setAttribute("fill", opts.pathFill);
        svg.appendChild(p);
        return svg;
      }

      function setHighlight(labelEl, text, rawQuery) {
        labelEl.textContent = "";
        if (!rawQuery) { labelEl.textContent = text; return; }
        const idx = fold(text).indexOf(fold(rawQuery));
        if (idx < 0) { labelEl.textContent = text; return; }
        const mark = document.createElement("mark"); mark.className = "mark";
        mark.textContent = text.slice(idx, idx + rawQuery.length);
        labelEl.append(document.createTextNode(text.slice(0,idx)), mark, document.createTextNode(text.slice(idx+rawQuery.length)));
      }

      const sidebar  = document.getElementById("sidebar");
      const nav      = document.getElementById("nav");
      const REGISTRY = { sections: [] };

      /* ── Hover-expand (expande temporariamente a sidebar recolhida) ─────── */
      let _hoverLeaveTimer = null;

      function iniciarHoverExpand() {
        const RAIL = "#btnToggle, #btnTheme, #btnLogout";
        let _enterTimer = null;

        sidebar.addEventListener("mouseenter", (e) => {
          if (!sidebar.classList.contains("is-collapsed")) return;
          if (e.target.closest(RAIL)) return;
          clearTimeout(_hoverLeaveTimer);
          _enterTimer = setTimeout(() => {
            if (!sidebar.matches(":hover")) return;
            const h = sidebar.querySelector(":hover");
            if (h && h.closest(RAIL)) return;
            sidebar.classList.add("is-hover-expanded");
            const sn = document.getElementById("shellNav");
            if (sn) sn.classList.add("is-hover-expanded");
          }, 600);
        });
        sidebar.addEventListener("mouseleave", () => {
          if (!sidebar.classList.contains("is-collapsed")) return;
          clearTimeout(_enterTimer);
          _hoverLeaveTimer = setTimeout(() => {
            const sn = document.getElementById("shellNav");
            const sn2 = document.getElementById("shellNav");
            sidebar.classList.add("is-collapsing");
            sidebar.classList.remove("is-hover-expanded");
            if (sn2) { sn2.classList.add("is-collapsing"); sn2.classList.remove("is-hover-expanded"); }
            sidebar.addEventListener("transitionend", function _done(e) {
              if (e.propertyName !== "width") return;
              sidebar.classList.remove("is-collapsing");
              if (sn2) sn2.classList.remove("is-collapsing");
              sidebar.removeEventListener("transitionend", _done);
            });
          }, 120);
        });
        [$("#btnToggle"), $("#btnTheme"), $("#btnLogout")].forEach(btn => {
          if (btn) btn.addEventListener("mouseenter", () => clearTimeout(_enterTimer));
        });
      }

      /* ── Collapsed / Tema ────────────────────────────────────────────────── */
      function setCollapsed(collapsed) {
        const shellNav = document.getElementById("shellNav");
        if (collapsed) {
          sidebar.classList.add("is-collapsing");
          sidebar.classList.add("is-collapsed");
          if (shellNav) { shellNav.classList.add("is-collapsing"); shellNav.classList.add("is-collapsed"); }
          sidebar.addEventListener("transitionend", function _done(e) {
            if (e.propertyName !== "width") return;
            sidebar.classList.remove("is-collapsing");
            if (shellNav) shellNav.classList.remove("is-collapsing");
            sidebar.removeEventListener("transitionend", _done);
          });
        } else {
          sidebar.classList.remove("is-collapsing");
          sidebar.classList.remove("is-collapsed");
          if (shellNav) { shellNav.classList.remove("is-collapsing"); shellNav.classList.remove("is-collapsed"); }
        }
        /* Sempre limpa o hover-expand ao mudar estado */
        sidebar.classList.remove("is-hover-expanded");
        clearTimeout(_hoverLeaveTimer);
        try { localStorage.setItem(LS_COLLAPSED, collapsed ? "1" : "0"); } catch(e) {}

        const b = $("#btnToggle");
        b.title = collapsed ? "Expandir menu" : "Recolher menu";
        b.setAttribute("aria-label", b.title);
        b.textContent = "";
        const pinEl = document.createElement("i");
        pinEl.className = collapsed ? "bi bi-pin-angle" : "bi bi-pin";
        pinEl.style.fontSize = "16px";
        b.appendChild(pinEl);
        toggleUserMenu(false);

        /* Ao recolher: apenas limpa filtro de busca (não fecha menus) */
        if (collapsed) {
          const filtro = $("#filtro");
          if (filtro && filtro.value !== "") { filtro.value = ""; aplicarFiltro(""); }
        }
      }

      function setTheme(theme) {
        sidebar.setAttribute("data-theme", theme);
        /* Propaga o tema ao wrapper shell-nav para sincronizar o fundo */
        const sn = document.getElementById("shellNav");
        if (sn) sn.setAttribute("data-theme", theme);
        try { localStorage.setItem(LS_THEME, theme); } catch(e) {}
        const b = $("#btnTheme");
        b.textContent = "";
        if (theme === "dark") {
          /* 12) Tema claro → ícone sol laranja */
          const ico = makeIcon("sun", 16, { stroke: "hsl(30 100% 55%)" });
          ico.style.color = "hsl(30 100% 55%)";
          b.style.color = "hsl(30 100% 55%)";
          b.appendChild(ico);
        } else {
          /* 13) Tema escuro → ícone lua totalmente roxa (fill + stroke roxo) */
          const ico = makeIcon("moon", 16, { fill: "hsl(270 70% 55%)", stroke: "hsl(270 70% 55%)" });
          b.style.color = "hsl(270 70% 55%)";
          b.appendChild(ico);
        }
        b.title = theme==="dark" ? "Tema claro" : "Tema escuro";
        b.setAttribute("aria-label", b.title);
      }

      /* ── Mobile ──────────────────────────────────────────────────────────── */
      function abrirMobile()  { sidebar.classList.add("is-mobile-open");    $("#backdrop").classList.add("is-visible"); }
      function fecharMobile() { sidebar.classList.remove("is-mobile-open"); $("#backdrop").classList.remove("is-visible"); }

      /* ── Logout ──────────────────────────────────────────────────────────── */
      function logout() {
        const d = window.location.protocol+"//"+window.location.hostname+"/";
        try { window.top.location.href = d; } catch(e) { window.open(d,"_top"); }
      }

      /* ── Dropdown da conta ───────────────────────────────────────────────── */
      let _userMenuOpen = false;
      function toggleUserMenu(force) {
        const menu = $("#userMenu");
        if (!menu) return;
        const show = force !== undefined ? force : menu.classList.contains("is-hidden");
        menu.classList.toggle("is-hidden", !show);
        _userMenuOpen = show;
        $("#userMenuBtn").setAttribute("aria-expanded", show ? "true" : "false");

        /* Posiciona o dropdown via fixed baseado na posição do botão */
        if (show) {
          const btn = $("#userMenuBtn");
          const r = btn.getBoundingClientRect();
          /* Abre à direita do botão; se não couber, abre acima */
          const menuW = 180;
          let left = r.right + 8;
          let top  = r.top;
          if (left + menuW > window.innerWidth) left = r.left - menuW - 8;
          menu.style.left = left + "px";
          menu.style.top  = top  + "px";
          menu.style.bottom = "auto";
        }
      }
      function trocarSenha() {
        toggleUserMenu(false);
        abrirNoFrameMeio(URL_TROCAR_SENHA);
      }

      /* ── Navegação (abre no iframe de conteúdo) ──────────────────────────── */
      function abrirNoFrameMeio(url) {
        if (!url) return;
        const frame = document.getElementById("meio");
        if (frame) frame.src = url;
      }

      function atualizarContexto(kind, label) {
        const ek = $("#ctxKind"), el = $("#ctxLabel"), ctx = $("#context");
        if (!kind || !label) {
          ek.textContent = ""; ek.style.color = "";
          el.textContent = "Menu Principal"; ctx.title = "Menu Principal"; return;
        }
        ek.textContent = kind;
        /* "Serviço": azul claro fixo em ambos os temas; "Consulta": laranja */
        const isDark = sidebar.getAttribute("data-theme") === "dark";
        if (kind === "Serviço") {
          ek.style.color = "hsl(217 91% 45%)";
        } else if (kind === "Consulta") {
          ek.style.color = isDark ? "#fb923c" : "#f97316";
        } else {
          ek.style.color = "";
        }
        el.textContent = label; ctx.title = kind+" — "+label;
      }

      function selecionarItem(btn, kind, item) {
        $$(".item.is-active").forEach(el => el.classList.remove("is-active"));
        $$(".group.has-active").forEach(el => el.classList.remove("has-active"));
        btn.classList.add("is-active");
        const grupo = btn.closest(".group");
        if (grupo) grupo.classList.add("has-active");
        atualizarContexto(kind, item.nome);
        abrirNoFrameMeio(item.url);
        fecharMobile();
        /* Exibe o context ao navegar para uma página */
        const ctx = document.getElementById("context");
        if (ctx) ctx.classList.remove("context-home-hidden");
      }

      function voltarMenuPrincipal(expandir = true) {
        $$(".item.is-active").forEach(el => el.classList.remove("is-active"));
        $$(".group.has-active").forEach(el => el.classList.remove("has-active"));
        atualizarContexto("","");
        abrirNoFrameMeio("principal_mid.asp");
        fecharMobile();
        /* Oculta o context na página inicial */
        const ctx = document.getElementById("context");
        if (ctx) ctx.classList.add("context-home-hidden");
        /* Ao navegar para o menu principal, sempre expande a sidebar */
        if (expandir && sidebar.classList.contains("is-collapsed")) {
          setCollapsed(false);
        }
      }

      /* ── Filtro ──────────────────────────────────────────────────────────── */
      function aplicarFiltro(raw) {
        const q = fold(raw); const filtering = q.length > 0; let total = 0;
        REGISTRY.sections.forEach(sec => {
          let secVis = false;
          sec.groups.forEach(grp => {
            let grpVis = false;
            grp.items.forEach(it => {
              const match = !filtering || fold(it.nome).includes(q);
              it.btn.style.display = match ? "" : "none";
              setHighlight(it.labelEl, it.nome, filtering ? raw : "");
              if (match) { grpVis = true; total++; }
            });
            grp.el.style.display = grpVis ? "" : "none";
            grp.el.classList.toggle("is-open", filtering && grpVis);
            if (grpVis) secVis = true;
          });
          sec.sectionEl.style.display = secVis ? "" : "none";
          if (filtering) sec.sectionEl.classList.add("is-open");
        });
        $("#navEmpty").classList.toggle("is-visible", filtering && total === 0);
      }

      /* ── Render do menu ──────────────────────────────────────────────────── */
      function criarItem(item, kind) {
        const btn = document.createElement("button");
        btn.type = "button"; btn.className = "item"; btn.title = item.nome;
        const label = document.createElement("span"); label.className = "item-label"; label.textContent = item.nome;
        btn.appendChild(label);
        btn.addEventListener("click", () => selecionarItem(btn, kind, item));
        return { btn, nome: item.nome, labelEl: label };
      }

      function criarGrupo(g, kind) {
        const grupo = document.createElement("div"); grupo.className = "group";
        const header = document.createElement("button"); header.type = "button"; header.className = "group-header"; header.title = g.grupo;
        const icon = document.createElement("span"); icon.className = "group-icon"; icon.appendChild(makeIcon(GROUP_ICON[g.grupo]||"list",16));
        const lbl  = document.createElement("span"); lbl.className = "group-label-txt"; lbl.textContent = g.grupo;
        const bdg  = document.createElement("span"); bdg.className = "group-badge"; bdg.textContent = g.itens.length;
        const chev = document.createElement("span"); chev.className = "group-chevron"; chev.appendChild(makeIcon("chevronRight",15));
        header.append(icon,lbl,bdg,chev);
        const body = document.createElement("div"); body.className = "group-body";
        const sub  = document.createElement("div"); sub.className = "group-sub";
        const items = g.itens.map(it => { const r = criarItem(it,kind); sub.appendChild(r.btn); return r; });
        body.appendChild(sub);
        header.addEventListener("click", () => {
          if (sidebar.classList.contains("is-collapsed")) setCollapsed(false);
          grupo.classList.toggle("is-open");
        });
        grupo.append(header,body);
        return { el: grupo, items };
      }

      function renderSecao(titulo, dados, kind) {
        const sec = document.createElement("div");
        sec.className = "section is-open "+(kind==="Serviço"?"section--servico":"section--consulta");
        const lbl = document.createElement("button"); lbl.type = "button"; lbl.className = "section-label"; lbl.title = titulo;
        const txt  = document.createElement("span"); txt.className = "section-label-txt"; txt.textContent = titulo;
        const chev = document.createElement("span"); chev.className = "section-chevron"; chev.appendChild(makeIcon("chevronRight",14));
        lbl.append(txt,chev);
        const body   = document.createElement("div"); body.className = "section-body";
        const groups = (dados||[]).map(g => { const r = criarGrupo(g,kind); body.appendChild(r.el); return r; });
        lbl.addEventListener("click", () => sec.classList.toggle("is-open"));

        const btnAll = document.createElement("button");
        btnAll.type = "button"; btnAll.className = "section-toggle-all";
        btnAll.title = "Expandir todos";
        btnAll.setAttribute("aria-label", "Expandir todos");
        btnAll.appendChild(makeIcon("chevronDoubleDown", 13));
        btnAll.addEventListener("click", (e) => {
          e.stopPropagation();
          if (!sec.classList.contains("is-open")) sec.classList.add("is-open");
          const allOpen = groups.every(g => g.el.classList.contains("is-open"));
          groups.forEach(g => g.el.classList.toggle("is-open", !allOpen));
          const next = allOpen ? "Expandir todos" : "Recolher todos";
          btnAll.title = next; btnAll.setAttribute("aria-label", next);
          btnAll.textContent = "";
          btnAll.appendChild(makeIcon(allOpen ? "chevronDoubleDown" : "chevronDoubleUp", 13));
        });
        lbl.appendChild(btnAll);
        sec.append(lbl,body); nav.appendChild(sec);
        REGISTRY.sections.push({ sectionEl: sec, groups });
      }

      function renderMenu() {
        nav.textContent = ""; REGISTRY.sections = [];
        renderSecao("Serviços",               SWNET_MENU.servicos,  "Serviço");
        renderSecao("Consultas / Relatórios", SWNET_MENU.consultas, "Consulta");
      }

      /* ── Init ────────────────────────────────────────────────────────────── */
      document.addEventListener("DOMContentLoaded", () => {

        /* ícone do toggle é definido por setCollapsed() abaixo */
        $("#btnHamburger").appendChild(makeIcon("list",18));
        $("#searchIcon")  .appendChild(makeIcon("search",16));
        $("#btnLogout")   .appendChild(makeIcon("logout",16));
        $("#userMenuBtn") .appendChild(makeIcon("gear",16));
        const lockEl = $("#lockIcon");
        if (lockEl) lockEl.appendChild(makeIcon("lock",16));

        /* Avatar */
        const partes = String($("#uUsuario").textContent||"").split(/[^A-Za-zÀ-ÿ]+/).filter(Boolean);
        $("#avatar").textContent = ((partes[0]||"")[0]||"") + ((partes[1]||"")[0]||"");

        /* Dot filial */
        const filialNome = ($("#uFilial").textContent||"").trim();
        $("#filialDot").style.background = FILIAL_COLOR[filialNome] || "#a1a1aa";

        /* Mac kbd */
        if (navigator.platform && /mac/i.test(navigator.platform)) $("#searchKbd").textContent = "⌘K";

        renderMenu();
        iniciarHoverExpand();

        let collapsed = false, theme = "light";
        try { collapsed = localStorage.getItem(LS_COLLAPSED)==="1"; } catch(e) {}
        try { theme = localStorage.getItem(LS_THEME)||"light"; } catch(e) {}
        try {
          if (sessionStorage.getItem("swnet.fromLogin") === "1") {
            sessionStorage.removeItem("swnet.fromLogin");
            collapsed = false;
          }
        } catch(e) {}
        setCollapsed(collapsed);
        setTheme(theme);
        voltarMenuPrincipal(false); /* init: não força expansão, respeita localStorage */

        /* Listeners */
        $("#btnToggle")   .addEventListener("click", () => setCollapsed(!sidebar.classList.contains("is-collapsed")));
        $("#btnTheme")    .addEventListener("click", () => setTheme(sidebar.getAttribute("data-theme")==="dark"?"light":"dark"));
        $("#btnHome")     .addEventListener("click", voltarMenuPrincipal);
        $("#btnLogout")   .addEventListener("click", logout);
        $("#btnHamburger").addEventListener("click", abrirMobile);
        $("#backdrop")    .addEventListener("click", fecharMobile);

        /* Dropdown conta */
        $("#userMenuBtn").addEventListener("click", (e) => { e.stopPropagation(); toggleUserMenu(); });
        $("#btnTrocarSenha").addEventListener("click", trocarSenha);
        document.addEventListener("click", (e) => {
          const menu = $("#userMenu"), btn = $("#userMenuBtn");
          if (!menu || menu.classList.contains("is-hidden")) return;
          if (!menu.contains(e.target) && !btn.contains(e.target)) toggleUserMenu(false);
        });

        /* Filtro */
        const filtro = $("#filtro");
        filtro.addEventListener("input",   () => aplicarFiltro(filtro.value));
        filtro.addEventListener("keydown", (e) => {
          if (e.key==="Escape") { filtro.value=""; aplicarFiltro(""); filtro.blur(); }
        });
        document.addEventListener("keydown", (e) => {
          if (e.key==="Escape") { toggleUserMenu(false); }
          if ((e.ctrlKey||e.metaKey) && (e.key==="k"||e.key==="K")) {
            e.preventDefault();
            if (sidebar.classList.contains("is-collapsed")) setCollapsed(false);
            if (window.innerWidth<=768) abrirMobile();
            filtro.focus(); filtro.select();
          }
        });

        /* No estado collapsed ao clicar no group-header na rail, expande */
        nav.addEventListener("click", (e) => {
          if (sidebar.classList.contains("is-collapsed")) {
            const gh = e.target.closest(".group-header");
            if (gh) { e.stopPropagation(); setCollapsed(false); }
          }
        });
      });

    })();

    /* ── Polling de sessão ───────────────────────────────────────────────────
       Verifica a cada 30 segundos se a sessão ainda está ativa.
       Quando o servidor retornar expirada:true, redireciona automaticamente
       para o principal.asp, que renderiza a tela de expiração.
    ── */
    (function () {
      var INTERVALO_MS = 30000;

      function verificarSessao() {
        fetch("principal.asp?check=1", { method: "GET", cache: "no-store" })
          .then(function (res) { return res.json(); })
          .then(function (data) {
            if (data.expirada) {
              window.top.location.href = "principal.asp";
            }
          })
          .catch(function () {
            /* falha silenciosa — tenta novamente no próximo intervalo */
          });
      }

      setInterval(verificarSessao, INTERVALO_MS);
    })();
    </script>

    <noscript>
      <div class="alert alert-danger m-3">
        Para usar o sistema, habilite o JavaScript no navegador.
      </div>
    </noscript>
  </body>
</html>
