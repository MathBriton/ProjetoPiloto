<!-- #include file="../../cnn2/cnn2.asp" -->
<!-- #include file="../../includes/variaveis.asp" -->
<!-- #include file="../../funcoes/funcoes.asp" -->
<%
Response.Buffer  = True
Response.Expires = 0
Session.lcId     = 1046

vldcod = Request("vldcod")
acao   = Request("acao")

If Trim(Request("acao")) <> "INS" Then

  SQL = "" & _
    "SELECT " & _
      "IdValidador  AS VLD_COD, " & _
      "U.CRM_NOME   AS VLD_NOM, " & _
      "Planta       AS VLD_PLT, " & _
      "Nivel        AS VLD_NIV, " & _
      "Minimo       AS VLD_VALMIN, " & _
      "Maximo       AS VLD_VALMAX " & _
    "FROM ValidadoresCredito " & _
    "INNER JOIN HELPDESK..CRM_DTI U ON IdValidador = U.CRM_R_E_C_N_O " & _
    "INNER JOIN HELPDESK..Empresas E ON U.EMP_CODIGO = E.Empresa " & _
    "WHERE IdValidador = '" & vldcod & "' " & _
      "AND (ISNULL(Excluido,0) = 0 OR Excluido = '0') "

  SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )

  If Not PIR_CORPORATIVO.EOF Then
    cboVLD_COD              = Trim(vldcod)
    edtValidadorCodigo      = Trim(vldcod)
    edtValidadorNome        = iif( (PIR_CORPORATIVO("VLD_NOM")    = "" Or IsNull(PIR_CORPORATIVO("VLD_NOM")))   , "", PIR_CORPORATIVO("VLD_NOM") )
    cboValidadorNivel       = iif( (PIR_CORPORATIVO("VLD_NIV")    = "" Or IsNull(PIR_CORPORATIVO("VLD_NIV")))   , "", PIR_CORPORATIVO("VLD_NIV") )
    edtValidadorValorMinimo = iif( (PIR_CORPORATIVO("VLD_VALMIN") = "" Or IsNull(PIR_CORPORATIVO("VLD_VALMIN"))),  0, PIR_CORPORATIVO("VLD_VALMIN") )
    edtValidadorValorMaximo = iif( (PIR_CORPORATIVO("VLD_VALMAX") = "" Or IsNull(PIR_CORPORATIVO("VLD_VALMAX"))),  0, PIR_CORPORATIVO("VLD_VALMAX") )
  End If

  PIR_CORPORATIVO.Close()
  Set PIR_CORPORATIVO = Nothing

Else
  edtValidadorCodigo      = 0
  edtValidadorNome        = ""
  cboValidadorNivel       = ""
  edtValidadorValorMinimo = 0
  edtValidadorValorMaximo = 0
End If

' Dicionario de plantas
Dim p
Set p = CreateObject("Scripting.Dictionary")
p.Add "00","0"
p.Add "01","0"
p.Add "02","0"
p.Add "03","0"
p.Add "04","0"
p.Add "05","0"
p.Add "06","0"
p.Add "07","0"
p.Add "08","0"
p.Add "09","0"
p.Add "10","0"
p.Add "11","0"
p.Add "12","0"
p.Add "13","0"

If acao <> "INS" Then
  SQL = "" & _
    "SELECT IdValidador, Planta " & _
    "FROM ValidadoresCredito " & _
    "WHERE IdValidador = '" & vldcod & "' " & _
      "AND (ISNULL(Excluido,0) = 0 OR Excluido = '0') "

  SET PIR_CORPORATIVO1 = Conecta( "PIR-CORPORATIVO", SQL )

  While Not PIR_CORPORATIVO1.EOF
    Select Case PIR_CORPORATIVO1("Planta")
      Case "00" p.Item("00") = "1"
      Case "01" p.Item("01") = "1"
      Case "02" p.Item("02") = "1"
      Case "03" p.Item("03") = "1"
      Case "04" p.Item("04") = "1"
      Case "05" p.Item("05") = "1"
      Case "06" p.Item("06") = "1"
      Case "07" p.Item("07") = "1"
      Case "08" p.Item("08") = "1"
      Case "09" p.Item("09") = "1"
      Case "10" p.Item("10") = "1"
      Case "11" p.Item("11") = "1"
      Case "12" p.Item("12") = "1"
      Case "13" p.Item("13") = "1"
    End Select
    PIR_CORPORATIVO1.MoveNext()
  Wend

  PIR_CORPORATIVO1.Close()
  Set PIR_CORPORATIVO1 = Nothing
End If

' Query de usuarios (executada antes do HTML)
SQL = "SELECT " & _
        "U.CRM_R_E_C_N_O AS USR_REC, " & _
        "U.CRM_NOME      AS USR_NOM, " & _
        "U.EMP_CODIGO    AS RAZ_SOC, " & _
        "E.Razao         AS EMP_RAZ " & _
      "FROM CRM_DTI U " & _
      "INNER JOIN Empresas E ON U.EMP_CODIGO = E.Empresa " & _
      "WHERE CRM_D_E_L_E_T <> '*' " & _
      iif( (acao<>"INS"), "AND U.CRM_R_E_C_N_O = " & vldcod & " ", " ") & _
      "ORDER BY CRM_NOME "

SET PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )

' Titulo da acao
Select Case Trim(Request("acao"))
  Case "INS" titulo = "Novo"
  Case "ALT" titulo = "Alterar"
  Case "CON" titulo = "Consultar"
  Case "DEL" titulo = "Apagar"
End Select
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet - Validador de Credito</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
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
      --color-danger:        hsl(3, 90%, 58%);
      --color-warning:       hsl(35, 100%, 48%);
      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
      --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;
      --radius-sm:  6px;
      --radius-md:  10px;
      --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);
      --transition: 160ms ease;
    }

    body {
      background-color: transparent;
      font-family: var(--font-sans);
      color: var(--color-text);
      margin: 0;
    }

    .form-wrapper   { background-color: var(--color-surface); min-height: 100%; }
    .form-card-body { padding: 1.25rem 1.25rem 1.5rem; }

    .form-label {
      font-size: .78rem;
      font-weight: 500;
      color: var(--color-text-muted);
      letter-spacing: .01em;
      margin-bottom: .3rem;
    }

    .form-control,
    .form-select {
      font-size: .85rem;
      border-color: var(--color-border);
      border-radius: var(--radius-sm);
      color: var(--color-text);
      background-color: var(--color-surface);
      transition: border-color var(--transition), box-shadow var(--transition);
    }
    .form-control:focus,
    .form-select:focus {
      border-color: var(--color-accent);
      box-shadow: 0 0 0 3px var(--color-accent-soft);
      outline: none;
    }
    .form-control[readonly],
    .form-select[disabled] {
      background-color: var(--color-surface-muted);
      color: var(--color-text-muted);
      cursor: not-allowed;
    }

    input[type='number'] { -moz-appearance: textfield; }
    input::-webkit-outer-spin-button,
    input::-webkit-inner-spin-button { -webkit-appearance: none; }

    .input-monetario {
      text-align: right;
      font-family: var(--font-mono);
      font-variant-numeric: tabular-nums;
    }

    .plantas-section {
      border: 1px solid var(--color-border);
      border-radius: var(--radius-sm);
      padding: .85rem 1rem;
      background-color: var(--color-surface-muted);
    }
    .plantas-section .section-label {
      font-size: .72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: var(--color-text-muted);
      margin-bottom: .65rem;
    }
    .plantas-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: .35rem .5rem;
    }
    .form-check-label { font-size: .83rem; color: var(--color-text); }
    .form-check-input:checked        { background-color: var(--color-accent); border-color: var(--color-accent); }
    .form-check-input:focus          { box-shadow: 0 0 0 3px var(--color-accent-soft); }

    .btn-primary {
      background-color: var(--color-accent);
      border-color: var(--color-accent);
      color: #fff;
      font-weight: 500;
      font-size: .85rem;
      border-radius: var(--radius-sm);
      transition: background-color var(--transition), border-color var(--transition), transform var(--transition);
    }
    .btn-primary:hover  { background-color: var(--color-accent-hover); border-color: var(--color-accent-hover); }
    .btn-primary:active { transform: scale(.98); }

    .btn-light {
      background-color: var(--color-surface);
      border: 1px solid var(--color-border);
      color: var(--color-text);
      font-weight: 500;
      font-size: .85rem;
      border-radius: var(--radius-sm);
      transition: background-color var(--transition), border-color var(--transition), transform var(--transition);
    }
    .btn-light:hover  { background-color: var(--color-surface-muted); border-color: var(--color-border-strong); }
    .btn-light:active { transform: scale(.98); }

    /* Toast */
    .toast-container-custom {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 1100;
      display: flex;
      flex-direction: column;
      gap: .5rem;
    }
    .toast-custom {
      min-width: 280px;
      max-width: 360px;
      border-radius: var(--radius-md);
      overflow: hidden;
      box-shadow: var(--shadow-md);
      background: var(--color-surface);
      border: 1px solid var(--color-border);
      animation: toastSlideIn .22s ease;
    }
    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(24px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-progress-bar {
      height: 3px;
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    /* Spinner no botao */
    .btn-spinner {
      display: none;
      width: .85rem;
      height: .85rem;
      border: 2px solid rgba(255,255,255,.4);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin .6s linear infinite;
      vertical-align: middle;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.01ms !important;
        transition-duration: 0.01ms !important;
      }
    }
  </style>
</head>

<body>
  <!-- Toast container — populado via JS -->
  <div class="toast-container-custom" id="toastContainer" aria-live="polite"></div>

  <div class="form-wrapper">
    <div class="form-card-body">

      <form id="validadordecredito" name="validadordecredito" class="needs-validation" novalidate>

        <!-- Contexto da requisicao (nao enviados no JSON diretamente — lidos por ID) -->
        <input type="hidden" id="edtValidadorCodigo" value="<%=vldcod%>" />
        <input type="hidden" id="acao"               value="<%=acao%>" />
        <input type="hidden" id="planta"             value="<%=planta%>" />
        <input type="hidden" id="page"               value="<%=page%>" />

        <!-- Validador -->
        <div class="mb-3">
          <label for="cboVLD_COD" class="form-label">Validador</label>
          <select id="cboVLD_COD"
                  class="form-select form-select-sm"
                  required
                  <% If Trim(Request("acao")) <> "INS" Then %> disabled <% End If %>>
            <% If acao = "INS" Then %>
            <option value="">Selecione...</option>
            <% End If %>
            <%
            While Not PIR_HELPDESK.EOF
            %>
            <option value="<%=PIR_HELPDESK("USR_REC")%>"
              <% If PIR_HELPDESK("USR_REC") = edtValidadorCodigo Then %> selected <% End If %>>
              <%=Trim(PIR_HELPDESK("USR_NOM")) & " - (" & Trim(PIR_HELPDESK("EMP_RAZ")) & ")"%>
            </option>
            <%
              PIR_HELPDESK.MoveNext()
            Wend
            PIR_HELPDESK.Close()
            Set PIR_HELPDESK = Nothing
            %>
          </select>
        </div>

        <!-- Nivel + Minimo + Maximo -->
        <div class="row g-3 mb-3">
          <div class="col-4">
            <label for="cboValidadorNivel" class="form-label">Nivel</label>
            <select id="cboValidadorNivel"
                    class="form-select form-select-sm"
                    required
                    <% If Trim(Request("acao")) = "CON" Or Trim(Request("acao")) = "DEL" Then %> disabled <% End If %>>
              <option value=""   <% If cboValidadorNivel = ""   Then %> selected <% End If %>>Selecione...</option>
              <option value="10" <% If cboValidadorNivel = "10" Then %> selected <% End If %>>Analista de Credito</option>
              <option value="20" <% If cboValidadorNivel = "20" Then %> selected <% End If %>>Coordenador de Credito</option>
              <option value="30" <% If cboValidadorNivel = "30" Then %> selected <% End If %>>Gerente de Tesouraria</option>
              <option value="40" <% If cboValidadorNivel = "40" Then %> selected <% End If %>>Divisional Controller</option>
              <option value="50" <% If cboValidadorNivel = "50" Then %> selected <% End If %>>Gerente de Planta</option>
              <option value="60" <% If cboValidadorNivel = "60" Then %> selected <% End If %>>CFO</option>
            </select>
          </div>

          <div class="col-4">
            <label for="edtValidadorValorMinimo" class="form-label">Valor Minimo</label>
            <input id="edtValidadorValorMinimo"
                   type="text"
                   value="<%=edtValidadorValorMinimo%>"
                   class="form-control form-control-sm input-monetario"
                   onkeypress="return(moeda(this,'.',',',event))"
                   required
                   <% If Trim(Request("acao")) = "CON" Or Trim(Request("acao")) = "DEL" Then %> readonly <% End If %> />
          </div>

          <div class="col-4">
            <label for="edtValidadorValorMaximo" class="form-label">Valor Maximo</label>
            <input id="edtValidadorValorMaximo"
                   type="text"
                   value="<%=edtValidadorValorMaximo%>"
                   class="form-control form-control-sm input-monetario"
                   onkeypress="return(moeda(this,'.',',',event))"
                   required
                   <% If Trim(Request("acao")) = "CON" Or Trim(Request("acao")) = "DEL" Then %> readonly <% End If %> />
          </div>
        </div>

        <!-- Plantas -->
        <div class="mb-4">
          <div class="plantas-section">
            <div class="section-label">
              <i class="bi bi-building me-1"></i>Plantas
            </div>
            <div class="plantas-grid">
              <%
              Dim plantasDados(13, 1)
              plantasDados(0,0)  = "selecionaPlanta00" : plantasDados(0,1)  = "00 - Pirapetinga"
              plantasDados(1,0)  = "selecionaPlanta01" : plantasDados(1,1)  = "01 - Rio de Janeiro"
              plantasDados(2,0)  = "selecionaPlanta02" : plantasDados(2,1)  = "02 - Sao Paulo"
              plantasDados(3,0)  = "selecionaPlanta03" : plantasDados(3,1)  = "03 - Corporativo"
              plantasDados(4,0)  = "selecionaPlanta04" : plantasDados(4,1)  = "04 - Uberaba"
              plantasDados(5,0)  = "selecionaPlanta05" : plantasDados(5,1)  = "05 - Bento Goncalves"
              plantasDados(6,0)  = "selecionaPlanta06" : plantasDados(6,1)  = "06 - Fortaleza"
              plantasDados(7,0)  = "selecionaPlanta07" : plantasDados(7,1)  = "07 - Saquarema"
              plantasDados(8,0)  = "selecionaPlanta08" : plantasDados(8,1)  = "08 - Aracatuba"
              plantasDados(9,0)  = "selecionaPlanta09" : plantasDados(9,1)  = "09 - Blumenau"
              plantasDados(10,0) = "selecionaPlanta10" : plantasDados(10,1) = "10 - Pacajus"
              plantasDados(11,0) = "selecionaPlanta11" : plantasDados(11,1) = "11 - Porto Feliz"
              plantasDados(12,0) = "selecionaPlanta12" : plantasDados(12,1) = "12 - Tres Barras"
              plantasDados(13,0) = "selecionaPlanta13" : plantasDados(13,1) = "13 - Florestal"

              Dim i, chave
              For i = 0 To 13
                chave = Right("0" & CStr(i), 2)
              %>
              <div class="form-check form-switch">
                <input class="form-check-input planta-check"
                       type="checkbox"
                       id="<%=plantasDados(i,0)%>"
                       data-planta="<%=chave%>"
                       <% If Trim(Request("acao")) = "CON" Or Trim(Request("acao")) = "DEL" Then %> disabled <% End If %>
                       <% If p.Item(chave) = "1" Then %> checked <% End If %>>
                <label class="form-check-label" for="<%=plantasDados(i,0)%>"><%=plantasDados(i,1)%></label>
              </div>
              <%
              Next
              %>
            </div>
          </div>
        </div>

        <!-- Botoes -->
        <div class="d-flex gap-2 justify-content-center">
          <% If Trim(Request("acao")) = "INS" Then %>
          <button type="button" id="btnSalvar" class="btn btn-primary btn-sm" title="Salvar novo registro" onclick="VerificaDados('INS')">
            <span class="btn-spinner" id="spinnerBtn"></span>
            <i class="bi bi-check-lg me-2" id="iconBtn"></i>Salvar
          </button>
          <% End If %>

          <% If Trim(Request("acao")) = "ALT" Then %>
          <button type="button" id="btnSalvar" class="btn btn-primary btn-sm" title="Salvar alteracoes" onclick="VerificaDados('ALT')">
            <span class="btn-spinner" id="spinnerBtn"></span>
            <i class="bi bi-check-lg me-2" id="iconBtn"></i>Salvar
          </button>
          <% End If %>

          <% If Trim(Request("acao")) = "DEL" Then %>
          <button type="button" id="btnSalvar" class="btn btn-danger btn-sm" title="Confirmar exclusao" onclick="VerificaDados('DEL')">
            <span class="btn-spinner" id="spinnerBtn"></span>
            Apagar
          </button>
          <% End If %>
        </div>

      </form>
    </div>
  </div>

  <!-- Modal de confirmacao de exclusao -->
  <div class="modal fade" id="modalConfirmarExclusao" tabindex="-1" aria-labelledby="modalExclusaoLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content" style="border-radius: var(--radius-md); border: 1px solid var(--color-border); box-shadow: var(--shadow-md);">
        <div class="modal-header" style="border-bottom: 1px solid var(--color-border); padding: 1rem 1.25rem .75rem;">
          <div class="d-flex align-items-center gap-2">
            <span style="display:inline-flex;align-items:center;justify-content:center;width:2rem;height:2rem;border-radius:50%;background:hsl(3,90%,95%);">
              <i class="bi bi-trash3-fill" style="color:var(--color-danger);font-size:.9rem;"></i>
            </span>
            <h6 class="modal-title mb-0" id="modalExclusaoLabel" style="font-size:.9rem;font-weight:600;color:var(--color-text);">
              Confirmar exclusao
            </h6>
          </div>
          <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body" style="padding: .9rem 1.25rem; font-size:.84rem; color:var(--color-text-muted);">
          <p class="mb-1">Deseja apagar este validador?</p>
          <p class="mb-0" style="color:var(--color-danger);font-size:.8rem;">
            <i class="bi bi-exclamation-triangle-fill me-2"></i>Esta acao nao pode ser desfeita.
          </p>
        </div>
        <div class="modal-footer" style="border-top: 1px solid var(--color-border); padding: .75rem 1.25rem; gap:.5rem;">
          <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">
            <i class="bi bi-x-lg me-2"></i>Cancelar
          </button>
          <button type="button" class="btn btn-danger btn-sm" id="btnConfirmarExclusao">
            <span class="btn-spinner" id="spinnerDel" style="border-color:rgba(255,255,255,.4);border-top-color:#fff;"></span>
            <i class="bi bi-trash3 me-2" id="iconDel"></i>Apagar
          </button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>
  <script src="../../funcoes/funcoesJAVASCRIPT.js?id=2345652"></script>

  <script>
    // ── Toast helper ──────────────────────────────────────────────
    // tipo: 'sucesso' | 'erro' | 'aviso'
    function exibeToast(tipo, titulo, mensagem) {
      var map = {
        sucesso: { bar: 'var(--color-success)', icon: 'bi-check-circle-fill', cor: 'var(--color-success)' },
        erro:    { bar: 'var(--color-danger)',  icon: 'bi-x-circle-fill',     cor: 'var(--color-danger)'  },
        aviso:   { bar: 'var(--color-warning)', icon: 'bi-exclamation-circle-fill', cor: 'var(--color-warning)' }
      };
      var c = map[tipo] || map.aviso;
      var $t = $(
        '<div class="toast-custom" role="alert">' +
          '<div class="toast-progress-bar" style="background-color:' + c.bar + ';"></div>' +
          '<div class="d-flex align-items-start p-3 gap-2">' +
            '<i class="bi ' + c.icon + '" style="color:' + c.cor + ';font-size:1rem;flex-shrink:0;margin-top:.1rem;"></i>' +
            '<div class="flex-grow-1">' +
              '<div style="font-size:.82rem;font-weight:600;color:' + c.cor + ';">' + titulo + '</div>' +
              '<div style="font-size:.82rem;color:var(--color-text-muted);margin-top:.15rem;">' + mensagem + '</div>' +
            '</div>' +
            '<button type="button" style="background:none;border:none;cursor:pointer;color:var(--color-text-faint);font-size:1rem;line-height:1;" ' +
              'onclick="$(this).closest(\'.toast-custom\').remove()">&times;</button>' +
          '</div>' +
        '</div>'
      );
      $('#toastContainer').append($t);
      setTimeout(function() { $t.remove(); }, 4500);
    }

    // ── Mascara de moeda ──────────────────────────────────────────
    function moeda(a, e, r, t) {
      var n = '', h = 0, j = 0, u = 0, tamanho2 = 0, l = '', ajd2 = '';
      var o = window.Event ? t.which : t.keyCode;
      if (o === 13 || o === 8) return true;
      n = String.fromCharCode(o);
      if ('0123456789'.indexOf(n) === -1) return false;
      for (u = a.value.length, h = 0; h < u && (a.value.charAt(h) === '0' || a.value.charAt(h) === r); h++);
      for (l = ''; h < u; h++) {
        if ('0123456789'.indexOf(a.value.charAt(h)) !== -1) l += a.value.charAt(h);
      }
      l += n;
      u = l.length;
      if (u === 0)      { a.value = ''; }
      else if (u === 1) { a.value = '0' + r + '0' + l; }
      else if (u === 2) { a.value = '0' + r + l; }
      else {
        for (ajd2 = '', j = 0, h = u - 3; h >= 0; h--) {
          if (j === 3) { ajd2 += e; j = 0; }
          ajd2 += l.charAt(h);
          j++;
        }
        a.value = '';
        for (tamanho2 = ajd2.length, h = tamanho2 - 1; h >= 0; h--) a.value += ajd2.charAt(h);
        a.value += r + l.substr(u - 2, u);
      }
      return false;
    }

    // ── "1.250,75" => 1250.75 ─────────────────────────────────────
    function parseMoeda(val) {
      if (!val) return 0;
      return parseFloat(val.replace(/\./g, '').replace(',', '.')) || 0;
    }

    // ── Coleta codigos das plantas marcadas ───────────────────────
    function coletaPlantas() {
      var plantas = [];
      $('.planta-check:checked').each(function() {
        plantas.push($(this).data('planta'));
      });
      return plantas;
    }

    // ── Estado do botao ───────────────────────────────────────────
    function setBtnLoading(ativo) {
      $('#btnSalvar').prop('disabled', ativo);
      if (ativo) { $('#spinnerBtn').show(); $('#iconBtn').hide(); }
      else       { $('#spinnerBtn').hide(); $('#iconBtn').show(); }
    }

    // ── Envia requisicao AJAX ─────────────────────────────────────
    function enviaRequisicao(acao) {
      var payload = {
        acao:   $('#acao').val(),
        codigo: acao === 'INS'
                  ? $('#cboVLD_COD').val()
                  : $('#edtValidadorCodigo').val(),
        nivel:  $('#cboValidadorNivel').val(),
        valMin: parseMoeda($('#edtValidadorValorMinimo').val()),
        valMax: parseMoeda($('#edtValidadorValorMaximo').val()),
        plantas: coletaPlantas(),
        planta: $('#planta').val(),
        page:   $('#page').val()
      };

      setBtnLoading(true);

      $.ajax({
        url:         'rpst_validadoresdecredito.asp',
        method:      'POST',
        contentType: 'application/json; charset=utf-8',
        data:        JSON.stringify(payload),
        dataType:    'json',
        success: function(resp) {
          if (resp && resp.sucesso) {
            exibeToast('sucesso', 'Aviso!', resp.mensagem || 'Operacao realizada com sucesso!');
            // Notifica o grid pai via postMessage para fechar o modal e recarregar
            setTimeout(function () {
              window.parent.postMessage(
                { tipo: 'vldcredito:ok', mensagem: resp.mensagem || '', acao: acao },
                window.location.origin
              );
            }, 1200);
          } else {
            exibeToast('erro', 'Erro!',
              (resp && resp.mensagem) ? resp.mensagem : 'Ocorreu um erro ao processar a solicitacao.');
          }
        },
        error: function(xhr) {
          var msg = 'Erro de comunicacao com o servidor.';
          try {
            var r = JSON.parse(xhr.responseText);
            if (r && r.mensagem) msg = r.mensagem;
          } catch(e) {}
          exibeToast('erro', 'Erro!', msg);
        },
        complete: function() {
          setBtnLoading(false);
        }
      });
    }

    // ── Validacao e envio via $.ajax + JSON ───────────────────────
    function VerificaDados(acao) {
      var ok = true;

      if (acao === 'INS' || acao === 'ALT') {
        if (!$('#cboVLD_COD').val()) {
          exibeToast('aviso', 'Atencao', 'Validador nao informado!');
          ok = false;
        }
        if (!$('#cboValidadorNivel').val()) {
          exibeToast('aviso', 'Atencao', 'Nivel nao informado!');
          ok = false;
        }

        var valMin = parseMoeda($('#edtValidadorValorMinimo').val());
        var valMax = parseMoeda($('#edtValidadorValorMaximo').val());

        if (valMax === 0) {
          exibeToast('aviso', 'Atencao', 'Valor maximo nao informado!');
          ok = false;
        }
        if (valMax > 0 && valMax < valMin) {
          exibeToast('aviso', 'Atencao', 'Valor Maximo deve ser maior que o Valor Minimo!');
          ok = false;
        }
      }

      if (!ok) return;

      // Exclusao: abre modal de confirmacao em vez de confirm() nativo
      if (acao === 'DEL') {
        var modalEl = document.getElementById('modalConfirmarExclusao');
        var modal   = bootstrap.Modal.getOrCreateInstance(modalEl);
        modal.show();
        return; // execucao continua no handler do botao de confirmacao abaixo
      }

      enviaRequisicao(acao);
    }

    // ── Confirmacao de exclusao via modal ─────────────────────────
    $(document).ready(function() {
      $('#btnConfirmarExclusao').on('click', function() {
        var modalEl = document.getElementById('modalConfirmarExclusao');
        var modal   = bootstrap.Modal.getInstance(modalEl);

        // Spinner no botao do modal
        $('#spinnerDel').show();
        $('#iconDel').hide();
        $('#btnConfirmarExclusao').prop('disabled', true);

        modal.hide();

        enviaRequisicao('DEL');

        // Restaura botao do modal para proxima abertura
        $(modalEl).one('hidden.bs.modal', function() {
          $('#spinnerDel').hide();
          $('#iconDel').show();
          $('#btnConfirmarExclusao').prop('disabled', false);
        });
      });
    });
  </script>

</body>
</html>
