<!-- #include file="../../includes/variaveis.asp" -->
<!-- #include file="../../funcoes/funcoes.asp"-->
<!-- #include file="../../cnn2/cnn2.asp"-->

<!DOCTYPE html>
<!-- V1.1 - Data:28/09/2020 - CHPL:165-->

<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>SWnet - Parâmetros para emissão de pedido</title>
    <!-- Bootstrap -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
    <!-- jQuery -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>
    <!-- jQuery Mask -->
    <script src="https://cdn.jsdelivr.net/npm/jquery-mask-plugin@1.14.16/dist/jquery.mask.min.js"></script>
    <!-- Moment.js (validação de datas) -->
    <script src="https://cdn.jsdelivr.net/npm/moment@2.30.1/moment.min.js"></script>

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>
    <script src="../../js/sw_logs.js"></script>

    <style>
      :root {
        /* — Neutros — */
        --color-bg:            hsl(240, 9%, 98%);
        --color-surface:       hsl(0, 0%, 100%);
        --color-surface-muted: hsl(240, 9%, 97%);
        --color-border:        hsl(240, 14%, 92%);
        --color-border-strong: hsl(240, 19%, 92%);

        --color-text:          hsl(240, 6%, 10%);
        --color-text-muted:    hsl(240, 4%, 46%);
        --color-text-faint:    hsl(240, 4%, 64%);

        /* — Accent único — */
        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        /* — Estado — */
        --color-success:       hsl(145, 63%, 32%);
        --color-success-soft:  hsl(145, 60%, 95%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);
        --color-warning:       hsl(35, 100%, 48%);
        --color-warning-soft:  hsl(40, 100%, 95%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
        --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

        --radius-sm: 6px;
        --radius-md: 10px;
        --radius-lg: 14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;
      }

      body {
        background-color: var(--color-bg);
        min-height: 100vh;
        font-family: var(--font-sans);
        color: var(--color-text);
      }

      .page-shell {
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      .page-title {
        font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
        font-weight: 600;
        letter-spacing: -0.01em;
        color: var(--color-text);
      }

      .plant-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: .25rem .75rem;
        border-radius: var(--radius-sm);
        font-size: .8rem;
        font-weight: 600;
        border: 1px solid transparent;
        margin-top: .4rem;
      }
      .plant-pir { color: hsl(210, 90%, 32%); background-color: hsl(210, 100%, 95%); border-color: hsl(210, 90%, 60%); }
      .plant-ube { color: hsl(145, 65%, 24%); background-color: hsl(145, 60%, 94%); border-color: hsl(145, 65%, 50%); }
      .plant-ben { color: hsl(35, 90%, 30%);  background-color: hsl(40, 100%, 93%); border-color: hsl(35, 90%, 55%); }
      .plant-for { color: hsl(3, 80%, 38%);   background-color: hsl(3, 100%, 95%);  border-color: hsl(3, 80%, 62%); }
      .plant-saq { color: hsl(270, 55%, 38%); background-color: hsl(270, 60%, 96%); border-color: hsl(270, 55%, 62%); }

      .param-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
      }

      .param-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1rem 1.25rem;
        border-bottom: 1px solid var(--color-border);
      }

      /* — Segmented control (pill buttons) para as abas — */
      .nav-tabs {
        border-bottom: none;
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        border-radius: 999px;
        padding: .25rem;
        gap: .25rem;
        flex-wrap: nowrap;
      }
      .nav-tabs .nav-item {
        flex: 1 1 0;
        text-align: center;
        display: flex;
      }
      .nav-tabs .btnTab {
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        border: none;
        border-radius: 999px;
        font-size: .74rem;
        font-weight: 500;
        color: var(--color-text-muted);
        padding: .45rem .35rem;
        background-color: transparent;
        transition: background-color var(--transition), color var(--transition), box-shadow var(--transition);
      }
      .nav-tabs .btnTab:hover {
        color: var(--color-text);
      }
      .nav-tabs .btnTab.active {
        color: hsl(240, 6%, 10%);
        font-weight: 600;
        background-color: var(--color-surface);
        border: 1px solid var(--color-border-strong);
        box-shadow: var(--shadow-sm);
      }

      .tab-content {
        padding: 1.5rem 1.25rem;
      }

      /* — Cabeçalho de seção (substitui .alert-secondary) — */
      .section-header {
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        color: var(--color-text);
        font-weight: 600;
        font-size: .85rem;
        text-align: center;
        padding: .55rem 1rem;
        margin-bottom: 1.25rem;
      }

      /* — Divisor de subgrupo (substitui .divider cinza/preto) — */
      .divider {
        color: var(--color-text-muted);
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        align-items: center;
        justify-content: center;
        display: flex;
        margin-top: 1.5rem;
        margin-bottom: .75rem;
        border-radius: var(--radius-sm);
        padding: .4rem;
        font-size: .82rem;
        font-weight: 600;
        width: 100%;
      }

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
      }
      .form-control:focus,
      .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }

      .input-group-text {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border);
        color: var(--color-text-muted);
        font-size: .82rem;
      }

      .form-check-input:checked {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
      }
      .form-check-input:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px var(--color-accent-soft);
      }
      .form-check-label {
        font-size: .85rem;
        color: var(--color-text);
      }

      /* — Botões (estilo "quiet") — */
      .btn-primary {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
        font-weight: 500;
        font-size: .85rem;
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-primary:hover {
        background-color: var(--color-accent-hover);
        border-color: var(--color-accent-hover);
        color: hsl(0, 0%, 100%);
      }

      .btn-light {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        color: var(--color-text);
        font-weight: 500;
        font-size: .85rem;
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-light:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
      }

      .btn-save {
        background-color: var(--color-success);
        border-color: var(--color-success);
        color: hsl(0, 0%, 100%);
        font-weight: 500;
        font-size: .85rem;
        transition: background-color var(--transition), border-color var(--transition);
      }
      .btn-save:hover {
        background-color: hsl(145, 63%, 27%);
        border-color: hsl(145, 63%, 27%);
        color: hsl(0, 0%, 100%);
      }

      .form-actions {
        display: flex;
        justify-content: center;
        gap: .5rem;
        padding: 1.25rem;
      }

      @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
          animation-duration: 0.01ms !important;
          transition-duration: 0.01ms !important;
        }
      }
    </style>

    <script type="text/javascript">

      $(document).ready(function(){
        $('edtPrcKgCX').mask("#.##0,00000", {reverse: true});
        $('edtPrcKgCX').on('blur').mask("#.##0,00000", {reverse: true});      

        $('edtPrcKgCH').mask("#.##0,00000", {reverse: true});
        $('edtPrcKgCH').on('blur').mask("#.##0,00000", {reverse: true});      
        
        $('edtPrcKgCV').mask("#.##0,00000", {reverse: true});
        $('edtPrcKgCV').on('blur').mask("#.##0,00000", {reverse: true});      
      });

      // ------------------------------------------------------------------------------------------
      function doOnLoad()
      {
        //console.log('doOnLoad');

        var cargaMaquinaAtiva = document.pedidoparametro.cbxCargaMaquina.checked;

        // Controla disabled dos sub-switches conforme estado inicial de Carga Maquina
        document.pedidoparametro.cbxBlqMaqCX.disabled = !cargaMaquinaAtiva;
        document.pedidoparametro.cbxBlqMaqCH.disabled = !cargaMaquinaAtiva;
        document.pedidoparametro.cbxBlqMaqCV.disabled = !cargaMaquinaAtiva;

        document.pedidoparametro.edtBlqMaqCXDe.readOnly  = (document.pedidoparametro.cbxBlqMaqCX.checked) ? false : true;
        document.pedidoparametro.edtBlqMaqCXAte.readOnly = (document.pedidoparametro.cbxBlqMaqCX.checked) ? false : true;
        
        document.pedidoparametro.edtBlqMaqCHDe.readOnly  = (document.pedidoparametro.cbxBlqMaqCH.checked) ? false : true;
        document.pedidoparametro.edtBlqMaqCHAte.readOnly = (document.pedidoparametro.cbxBlqMaqCH.checked) ? false : true;
        
        document.pedidoparametro.edtBlqMaqCVDe.readOnly  = (document.pedidoparametro.cbxBlqMaqCV.checked) ? false : true;
        document.pedidoparametro.edtBlqMaqCVAte.readOnly = (document.pedidoparametro.cbxBlqMaqCV.checked) ? false : true;

      }
      // ------------------------------------------------------------------------------------------
      function onclickBloqueioMaquina(obj)
      {
        //console.log( 'obj.name: ' + obj.name + ' - ' + obj.value );

        if (obj.name == 'cbxBlqMaqCX') {

          document.pedidoparametro.edtBlqMaqCXDe.readOnly = (obj.checked) ? false : true;
          document.pedidoparametro.edtBlqMaqCXAte.readOnly = (obj.checked) ? false : true;

          if (!obj.checked) {

            document.pedidoparametro.cbxBlqMaqCX.value = 0;
            document.pedidoparametro.edtBlqMaqCXDe.value = '';
            document.pedidoparametro.edtBlqMaqCXAte.value = '';

            document.getElementById('edtBlqMaqCXDe').dispatchEvent(new Event('change'));
            document.getElementById('edtBlqMaqCXAte').dispatchEvent(new Event('change'));

          } else {

            document.pedidoparametro.cbxBlqMaqCX.value = 1;

          }

        }

        if (obj.name == 'cbxBlqMaqCH') {

          document.pedidoparametro.edtBlqMaqCHDe.readOnly = (obj.checked) ? false : true;
          document.pedidoparametro.edtBlqMaqCHAte.readOnly = (obj.checked) ? false : true;

          if (!obj.checked) {

            document.pedidoparametro.cbxBlqMaqCH.value = 0;
            document.pedidoparametro.edtBlqMaqCHDe.value = '';
            document.pedidoparametro.edtBlqMaqCHAte.value = '';

            document.getElementById('edtBlqMaqCHDe').dispatchEvent(new Event('change'));
            document.getElementById('edtBlqMaqCHAte').dispatchEvent(new Event('change'));

          } else {

            document.pedidoparametro.cbxBlqMaqCH.value = 1;

          }

        }

        if (obj.name == 'cbxBlqMaqCV') {

          document.pedidoparametro.edtBlqMaqCVDe.readOnly = (obj.checked) ? false : true;
          document.pedidoparametro.edtBlqMaqCVAte.readOnly = (obj.checked) ? false : true;

          if (!obj.checked) {
            
            document.pedidoparametro.cbxBlqMaqCV.value = 0;
            document.pedidoparametro.edtBlqMaqCVDe.value = '';
            document.pedidoparametro.edtBlqMaqCVAte.value = '';

            document.getElementById('edtBlqMaqCVDe').dispatchEvent(new Event('change'));
            document.getElementById('edtBlqMaqCVAte').dispatchEvent(new Event('change'));

          } else {

            document.pedidoparametro.cbxBlqMaqCV.value = 1;

          }

        }
      }
      // ------------------------------------------------------------------------------------------
      function BloqueiaMaquina(aCheck) {

        if (document.pedidoparametro.cbxCargaMaquina.checked) {
          document.pedidoparametro.cbxCargaMaquina.value = 1;
        } else {
          document.pedidoparametro.cbxCargaMaquina.value = 0;
        }

        // Quando Carga Maquina é desmarcado, apenas desabilita e desmarca os sub-switches
        // mantendo-os sempre visíveis na tela
        if (!aCheck) {

          document.pedidoparametro.cbxBlqMaqCX.value   = 0;
          document.pedidoparametro.cbxBlqMaqCX.checked = false;
          document.pedidoparametro.cbxBlqMaqCX.disabled = true;

          document.pedidoparametro.edtBlqMaqCXDe.value    = '';
          document.pedidoparametro.edtBlqMaqCXAte.value   = '';
          document.pedidoparametro.edtBlqMaqCXDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCXAte.readOnly = true;

          document.pedidoparametro.cbxBlqMaqCH.value   = 0;
          document.pedidoparametro.cbxBlqMaqCH.checked = false;
          document.pedidoparametro.cbxBlqMaqCH.disabled = true;

          document.pedidoparametro.edtBlqMaqCHDe.value    = '';
          document.pedidoparametro.edtBlqMaqCHAte.value   = '';
          document.pedidoparametro.edtBlqMaqCHDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCHAte.readOnly = true;

          document.pedidoparametro.cbxBlqMaqCV.value   = 0;
          document.pedidoparametro.cbxBlqMaqCV.checked = false;
          document.pedidoparametro.cbxBlqMaqCV.disabled = true;

          document.pedidoparametro.edtBlqMaqCVDe.value    = '';
          document.pedidoparametro.edtBlqMaqCVAte.value   = '';
          document.pedidoparametro.edtBlqMaqCVDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCVAte.readOnly = true;

        } else {

          // Reabilita os sub-switches quando Carga Maquina é marcado
          document.pedidoparametro.cbxBlqMaqCX.disabled = false;
          document.pedidoparametro.cbxBlqMaqCH.disabled = false;
          document.pedidoparametro.cbxBlqMaqCV.disabled = false;

          // Campos de data permanecem readonly até o usuário marcar cada sub-switch
          document.pedidoparametro.edtBlqMaqCXDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCXAte.readOnly = true;
          document.pedidoparametro.edtBlqMaqCHDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCHAte.readOnly = true;
          document.pedidoparametro.edtBlqMaqCVDe.readOnly  = true;
          document.pedidoparametro.edtBlqMaqCVAte.readOnly = true;

        }

      }
      // ------------------------------------------------------------------------------------------
      function objReturn(objeto) {
          
        switch (objeto.erro) {
        case "ErroCONEXAO": 
        case "ErroTABELA": 
          alert(objeto.tipo);
          break;
        case "SUCESSO":
          console.log( 'LOG SALVO COM SUCESSO!' );
          break;
        }          

      }      

    </script>

  </head>
<%
  On Error GoTo 0
  On Error Resume Next

  ' -------------------------------------------------------------------------------------------------
  function dataYYYYMMDD( data )
    '2023-02-28 -> 20230228
    Ano = left(data,4)
    Mes = mid(data,6,2)
    Dia = right(data,2)
    dataYYYYMMDD = Ano & Mes & Dia
  end function

  '--------------------------------------------------------------------------------------------------
  function dataYYYY_MM_DD( data )
    '20230228 -> 2023-02-28
    Ano = left(data,4)
    Mes = mid(data,5,2)
    Dia = right(data,2)
    dataYYYY_MM_DD = Ano & "-" & Mes & "-" & Dia
  end function 

  planta = trim(request("Planta"))

  if trim(request("gravar")) = "1" then

    DIM cbxDiasViagem, cbxCargaMaquina, cbxLimiteDeCredito, cbxDiasUteis, cbxPrzMinBlqP1, cbxBlqMaqCX, cbxBlqMaqCH, cbxBlqMaqCV, cbxAprovacaoComercial     

    cbxDiasViagem      = iif( request("cbxDiasViagem") = 1, 1, 0 )
    cbxCargaMaquina    = iif( request("cbxCargaMaquina") = 1, 1, 0 )
    cbxLimiteDeCredito = iif( request("cbxLimiteDeCredito") = 1, 1, 0 )
    cbxDiasUteis       = iif( request("cbxDiasUteis") = 1, 1, 0 )
    cbxPrzMinBlqP1     = iif( request("cbxPrzMinBlqP1") = 1, 1, 0 )
    cbxBlqMaqCX        = iif( request("cbxBlqMaqCX") = 1, 1, 0 )
    cbxBlqMaqCH        = iif( request("cbxBlqMaqCH") = 1, 1, 0 )
    cbxBlqMaqCV        = iif( request("cbxBlqMaqCV") = 1, 1, 0 )
    cbxAprovacaoComercial = iif( request("cbxAprovacaoComercial") = 1, 1, 0 )

    'DeMaleta e AteMaleta ------------------------------------------------------------------------'
    if trim(request("edtBlqMaqCXDe")) <> "" and trim(request("edtBlqMaqCXAte")) <> "" then

      edtBlqMaqCXDe = trim(request("edtBlqMaqCXDe"))
      edtBlqMaqCXAte = trim(request("edtBlqMaqCXAte"))

    else

      edtBlqMaqCXDe = "" 
      edtBlqMaqCXAte = ""

    end if

    'DeCorteVinco e AteCorteVinco ----------------------------------------------------------------'
    if trim(request("edtBlqMaqCVDe")) <> "" and trim(request("edtBlqMaqCVAte")) <> "" then

      edtBlqMaqCVDe = trim(request("edtBlqMaqCVDe"))
      edtBlqMaqCVAte = trim(request("edtBlqMaqCVAte"))

    else

      edtBlqMaqCVDe = "" 
      edtBlqMaqCVAte = ""

    end if

    'DeChapa e AteChapa --------------------------------------------------------------------------'
    if trim(request("edtBlqMaqCHDe")) <> "" and trim(request("edtBlqMaqCHDe")) <> "" then

      edtBlqMaqCHDe = trim(request("edtBlqMaqCHDe"))
      edtBlqMaqCHAte = trim(request("edtBlqMaqCHAte"))

    else

      edtBlqMaqCHDe = "" 
      edtBlqMaqCHAte = ""

    end if

    edtPrcKgCV = iif( trim(request("edtPrcKgCV")) <> "", replace(replace(trim(request("edtPrcKgCV")),".",""),",","."), 0 )
    edtPrcKgCX = iif( trim(request("edtPrcKgCX")) <> "", replace(replace(trim(request("edtPrcKgCX")),".",""),",","."), 0 )
    edtPrcKgCH = iif( trim(request("edtPrcKgCH")) <> "", replace(replace(trim(request("edtPrcKgCH")),".",""),",","."), 0 )

    edtLargMin = iif( request("edtLargMin") = "", 0, request("edtLargMin") )
    edtLargMax = iif( request("edtLargMax") = "", 0, request("edtLargMax") )

    edtCompMin = iif( request("edtCompMin") = "", 0, request("edtCompMin") )
    edtCompMax = iif( request("edtCompMax") = "", 0, request("edtCompMax") )

    edtMtsCVmin = iif( request("edtMtsCVmin") = "", 0, request("edtMtsCVmin") )
    edtMtsCVmax = iif( request("edtMtsCVmax") = "", 0, request("edtMtsCVmax") )

    edtMtsCXmin = iif( request("edtMtsCXmin") = "", 0, request("edtMtsCXmin") )
    edtMtsCXmax = iif( request("edtMtsCXmax") = "", 0, request("edtMtsCXmax") )

    edtMtsCHmin = iif( request("edtMtsCHmin") = "", 0, request("edtMtsCHmin") )
    edtMtsCHmax = iif( request("edtMtsCHmax") = "", 0, request("edtMtsCHmax") )

    SQL = "" &_
      "UPDATE PARAMETROS_PEDIDO SET " &_
        "DiaUtil = " & cbxDiasUteis & ", " &_
        "DiasBloqueiaP1 = " & cbxPrzMinBlqP1 & ", " &_
        "DiasViagem = " & cbxDiasViagem & ", " &_
        "CargaMaquina = " & cbxCargaMaquina & ", " &_
        "LimiteDeCredito = " & cbxLimiteDeCredito & ", " &_
        "DiasAcrescimoPA = " & iif(trim(request("edtAcrescimoPA"))<>"",trim(request("edtAcrescimoPA")),0) & ", " &_
        "DiasAcrescimoCX = " & iif(trim(request("edtAcrescimoCX"))<>"",trim(request("edtAcrescimoCX")),0) & ", " &_
        "DiasAcrescimoCH = " & iif(trim(request("edtAcrescimoCH"))<>"",trim(request("edtAcrescimoCH")),0) & ", " &_
        "PesoMN = " & trim(request("edtPesoCX")) & ", " &_
        "PesoCV = " & trim(request("edtPesoCV")) & ", " &_
        "PesoCH = " & trim(request("edtPesoCH")) & ", " &_
        "BloqueiaMaquinaCaixa = " & cbxBlqMaqCX & ", " &_
        "BloqueiaMaquinaChapa = " & cbxBlqMaqCH & ", " &_
        "BloqueiaMaquinaCorteVinco = " & cbxBlqMaqCV & ", " &_
        "DeMaleta = '" & dataYYYYMMDD(edtBlqMaqCXDe) & "', " &_
        "AteMaleta = '" & dataYYYYMMDD(edtBlqMaqCXAte) & "', " &_
        "DeCorteVinco = '" & dataYYYYMMDD(edtBlqMaqCVDe) & "', " &_
        "AteCorteVinco = '" & dataYYYYMMDD(edtBlqMaqCVAte) & "', " &_
        "DeChapa = '" & dataYYYYMMDD(edtBlqMaqCHDe) & "', " &_
        "AteChapa = '" & dataYYYYMMDD(edtBlqMaqCHAte) & "', " &_
        "PrecoKgCorteVinco = " & edtPrcKgCV & ", " &_
        "PrecoKgMaleta = " & edtPrcKgCX & ", " &_
        "PrecoKgChapa = " & edtPrcKgCH & ", " &_
        "LarguraMin = " & edtLargMin & ", " &_
        "LarguraMax = " & edtLargMax & ", " &_
        "ComprimentoMin = " & edtCompMin & ", " &_
        "ComprimentoMax = " & edtCompMax & ", " &_
        "MtsCVmin = " & edtMtsCVmin & ", " &_
        "MtsCVmax = " & edtMtsCVmax & ", " &_
        "MtsCXmin = " & edtMtsCXmin & ", " &_
        "MtsCXmax = " & edtMtsCXmax & ", " &_
        "MtsCHmin = " & edtMtsCHmin & ", " &_
        "MtsCHmax = " & edtMtsCHmax & ", " &_
        "SPDAprovacaoComercial = " & cbxAprovacaoComercial & " " &_
      "WHERE Planta = '" & planta & "' "

    SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )

    if NOT isObject(PIR_CORPORATIVO) then
      response.end
    end if

    SET PIR_CORPORATIVO = Nothing
    
    if err.number = 0 then

      response.write "<script type=""text/javascript"" charset=""utf-8"">"
      response.write   "alert('Parâmetros salvos com sucesso !'); "
      response.write   "history.go(-1);"
      response.write "</script>"
      response.End

    end if

  end if 

  SQL = "" &_
    "SELECT " &_
      "Dias, " &_
      "DiaUtil, " &_
      "LimiteDeCredito, " &_
      "DiasBloqueiaP1, " &_
      "ISNULL(BloqueiaMaquinaChapa,0) BloqueiaMaquinaChapa, " &_
      "ISNULL(BloqueiaMaquinaCaixa,0) BloqueiaMaquinaCaixa, " &_
      "ISNULL(BloqueiaMaquinaCorteVinco,0) BloqueiaMaquinaCorteVinco, " &_
      "ISNULL(DeMaleta,'') DeMaleta, " &_
      "ISNULL(AteMaleta,'') AteMaleta, " &_
      "ISNULL(DeCorteVinco,'') DeCorteVinco, " &_
      "ISNULL(AteCorteVinco,'') AteCorteVinco, " &_
      "ISNULL(DeChapa,'') DeChapa, " &_
      "ISNULL(AteChapa,'') AteChapa, " &_
      "DiasViagem, " &_
      "CargaMaquina, " &_
      "DiasAcrescimoPA, " &_
      "DiasAcrescimoCX, " &_
      "DiasAcrescimoCH, " &_
      "PesoMN, " &_
      "PesoCV, " &_
      "PesoCH, " &_
      "PrecoKgCorteVinco, " &_
      "PrecoKgMaleta, " &_
      "PrecoKgChapa, " &_
      "LarguraMin, " &_
      "LarguraMax, " &_
      "ComprimentoMin, " &_
      "ComprimentoMax, " &_
      "MtsCVmin, " &_
      "MtsCVmax, " &_
      "MtsCXmin, " &_
      "MtsCXmax, " &_
      "MtsCHmin, " &_
      "MtsCHmax " &_
    "FROM PARAMETROS_PEDIDO " &_
    "WHERE Planta = '" & planta & "' "

  'RESPONSE.WRITE SQL
  'RESPONSE.END

  SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )

  if NOT isObject(PIR_CORPORATIVO) then
    response.end
  end if

  edtBlqMaqCXDe  = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("DeMaleta")))
  edtBlqMaqCXAte = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("AteMaleta")))

  edtBlqMaqCVDe  = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("DeCorteVinco")))
  edtBlqMaqCVAte = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("AteCorteVinco")))

  edtBlqMaqCHDe  = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("DeChapa")))
  edtBlqMaqCHAte = dataYYYY_MM_DD(trim(PIR_CORPORATIVO("AteChapa")))

  ' response.write "edtBlqMaqCXDe : " & edtBlqMaqCXDe & "<br>"
  ' response.write "edtBlqMaqCXAte: " & edtBlqMaqCXAte & "<br>"
  ' response.write "edtBlqMaqCVDe : " & edtBlqMaqCVDe & "<br>"
  ' response.write "edtBlqMaqCVAte: " & edtBlqMaqCVAte & "<br>"
  ' response.write "edtBlqMaqCHDe : " & edtBlqMaqCHDe & "<br>"
  ' response.write "edtBlqMaqCHAte: " & edtBlqMaqCHAte & "<br>"
  ' response.end

  edtPrcKgCV = PIR_CORPORATIVO("PrecoKgCorteVinco")
  edtPrcKgCX = PIR_CORPORATIVO("PrecoKgMaleta")
  edtPrcKgCH = PIR_CORPORATIVO("PrecoKgChapa")
  
  select case planta
  case "00": SubTitulo = "Pirapetinga"     : PlantaClasse = "plant-pir"
  case "04": SubTitulo = "Uberaba"         : PlantaClasse = "plant-ube"
  case "05": SubTitulo = "Bento Gonçalves" : PlantaClasse = "plant-ben"
  case "06": SubTitulo = "Fortaleza"       : PlantaClasse = "plant-for"
  case "07": SubTitulo = "Saquarema"       : PlantaClasse = "plant-saq"
  end select
  '------------------------------------------------------------------------------------------------'
%>
  <body onload="doOnLoad(); log_load_fields( 'pedidoparametro' )">
    <div class="page-shell">
      <div class="container-fluid py-4 px-3 px-lg-4" style="max-width: 900px;">

      <div class="param-card">

        <div class="param-header justify-content-center">
          <div class="text-center">
            <span class="page-title d-block">Parâmetros para Emissão de Pedido</span>
            <span class="plant-badge <%=PlantaClasse%>"><%=SubTitulo%></span>
          </div>
        </div>

          <form name="pedidoparametro" id="pedidoparametro" action="" method="post" class="needs-validation">

            <input type="hidden" name="gravar" id="gravar" value="" />
            <input type="hidden" name="planta" id="planta" value="<%=planta%>" />

            <!-- Nav tabs -->
            <ul class="nav nav-tabs mx-3 mt-3" role="tablist">
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab active" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabped" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabped" 
                        aria-selected="true">Prazos</button>
              </li>
              <!--
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabdia" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabdia" 
                        aria-selected="true">Dias de Acréscimo</button>
              </li>
              -->
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabpes" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabpes" 
                        aria-selected="true">Peso</button>
              </li>
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabmts" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabmts" 
                        aria-selected="true">M²</button>
              </li>
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tablim" 
                        type="button" 
                        role="tab" 
                        aria-controls="tablim" 
                        aria-selected="true">Largura X Comprimento</button>
              </li>
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabmaq" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabmaq" 
                        aria-selected="true">Bloqueio de Máquina</button>
              </li>
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabprc" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabprc" 
                        aria-selected="true">Preço Mínimo / Quilo</button>
              </li>
              <li class="nav-item" role="presentation">
                <button class="nav-link btnTab" 
                        data-bs-toggle="tab" 
                        data-bs-target="#tabspd" 
                        type="button" 
                        role="tab" 
                        aria-controls="tabspd" 
                        aria-selected="true">SPD</button>
              </li>
            </ul>
        
            <div class="tab-content">

              <!-- Prazos ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade show active" id="tabped" role="tabpanel" aria-labelledby="nav-tabped-tab">

                <div class="section-header">Pedido</div>

                  <div class="row">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxDiasUteis" name="cbxDiasUteis" class="form-check-input" type="checkbox" 
                          <% if PIR_CORPORATIVO("DiaUtil") then %> value="on" <% else %> value="off" <% end if %> 
                          <% if PIR_CORPORATIVO("DiaUtil") then %> checked <% end if %> >
                        <label for="cbxDiasUteis" class="form-check-label">Considerar apenas dias uteis</label>
                      </div>                   
                    </div>
                  </div>

                  <div class="row mt-2">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxPrzMinBlqP1" name="cbxPrzMinBlqP1" class="form-check-input" type="checkbox" 
                          <% if PIR_CORPORATIVO("DiasBloqueiaP1") and (planta = "00" or planta = "04") then %> value="on" <% else %> value="off" <% end if %> 
                          <% if PIR_CORPORATIVO("DiasBloqueiaP1") and (planta = "00" or planta = "04") then %> checked <% end if %> 
                          <% if (planta <> "00" and planta <> "04") then %> disabled <% end if %> >
                        <label for="cbxPrzMinBlqP1" class="form-check-label">Prazo mínimo bloqueia P1 (Clientes diferenciados)</label>
                      </div>
                    </div>
                  </div>

                  <div class="row mt-2">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxDiasViagem" name="cbxDiasViagem" class="form-check-input" type="checkbox" 
                          <% if PIR_CORPORATIVO("DiasViagem") then %> value="on" <% else %> value="off" <% end if %> 
                          <% if PIR_CORPORATIVO("DiasViagem") then %> checked <% end if %> >
                        <label for="cbxDiasViagem" class="form-check-label">Considerar dias de viagem</label>
                      </div>
                    </div>
                  </div>

                  <div class="row mt-2 mb-3">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxLimiteDeCredito" name="cbxLimiteDeCredito" class="form-check-input" type="checkbox" 
                          <% if PIR_CORPORATIVO("LimiteDeCredito") then %> value="on" <% else %> value="off" <% end if %> 
                          <% if PIR_CORPORATIVO("LimiteDeCredito") then %> checked <% end if %> >
                        <label for="cbxLimiteDeCredito" class="form-check-label">Limite de crédito</label>
                      </div>
                    </div>
                  </div>

                <div class="section-header mt-4">Dias de Acréscimo</div>

                  <div class="row justify-content-center g-3">

                    <div class="col-2">
                      <label for="edtAcrescimoPA" class="form-label">Papel</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Mais número de dias">+</span>
                        <input name="edtAcrescimoPA" id="edtAcrescimoPA" type="number" pattern="[0-9]+$" class="form-control" value="<%=trim(PIR_CORPORATIVO("DiasAcrescimoPA"))%>"/>
                      </div>
                    </div>

                    <div class="col-2">
                      <label for="edtAcrescimoCX" class="form-label">Caixa</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Mais número de dias">+</span>
                        <input name="edtAcrescimoCX" id="edtAcrescimoCX" type="number" pattern="[0-9]+$" class="form-control" value="<%=trim(PIR_CORPORATIVO("DiasAcrescimoCX"))%>"/>
                      </div>
                    </div>

                    <div class="col-2">
                      <label for="edtAcrescimoCH" class="form-label">Chapa</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Mais número de dias">+</span>
                        <input name="edtAcrescimoCH" id="edtAcrescimoCH" type="number" pattern="[0-9]+$" class="form-control" value="<%=trim(PIR_CORPORATIVO("DiasAcrescimoCH"))%>"/>
                      </div>
                    </div>
                  </div>

                <div class="row mt-4">
                  <div class="col text-center">
                    <a href="../formulario/EstruturaRestricao.asp" target="_blank" class="btn btn-light btn-sm">Restrições por estrutura</a>
                  </div>
                </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- Peso --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade show" id="tabpes" role="tabpanel" aria-labelledby="nav-tabpes-tab">

                <div class="section-header">Peso</div>

                  <div class="row justify-content-center g-3">

                    <div class="col-2">
                      <label for="edtPesoCV" class="form-label">Corte e vinco</label>
                      <input type="number" pattern="[0-9]+$" name="edtPesoCV" id="edtPesoCV" class="form-control" value="<%=trim(PIR_CORPORATIVO("PesoCV"))%>" />
                    </div>

                    <div class="col-2">
                      <label for="edtPesoCX" class="form-label">Maleta</label>
                      <input type="number" pattern="[0-9]+$" name="edtPesoCX" id="edtPesoCX" class="form-control" value="<%=trim(PIR_CORPORATIVO("PesoMN"))%>" />
                    </div>

                    <div class="col-2">
                      <label for="edtPesoCH" class="form-label">Chapa</label>
                      <input type="number" pattern="[0-9]+$" name="edtPesoCH" id="edtPesoCH" class="form-control" value="<%=trim(PIR_CORPORATIVO("PesoCH"))%>" />
                    </div>

                  </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- Metros Quadrados --------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade show" id="tabmts" role="tabpanel" aria-labelledby="nav-tabmts-tab">

                <div class="section-header">Metros Quadrados</div>

                  <div class="row justify-content-center">

                    <div class="divider">Corte e Vinco</div>

                    <div class="col-3">
                      <label for="edtMtsCVmin" class="form-label">Mínimo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Maior ou igual">>=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCVmin" id="edtMtsCVmin" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCVmin"))%>" />
                      </div>
                    </div>
<!--
                    <div class="col-3">
                      <label for="edtMtsCVmin" class="form-label">Máximo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Menor ou igual"><=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCVmax" id="edtMtsCVmax" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCVmax"))%>" />
                      </div>
                    </div>
-->
                  </div>

                  <div class="row justify-content-center">

                    <div class="divider">Maleta</div>

                    <div class="col-3">
                      <label for="edtMtsCXmin" class="form-label">Mínimo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Maior ou igual">>=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCXmin" id="edtMtsCXmin" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCXmin"))%>" />
                      </div>
                    </div>
<!--
                    <div class="col-3">
                      <label for="edtMtsCXmax" class="form-label">Máximo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Menor ou igual"><=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCXmax" id="edtMtsCXmax" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCXmax"))%>" />
                      </div>
                    </div>
-->
                  </div>

                  <div class="row justify-content-center">

                    <div class="divider">Chapa</div>

                    <div class="col-3">
                      <label for="edtMtsCHmin" class="form-label">Mínimo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Maior ou igual">>=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCHmin" id="edtMtsCHmin" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCHmin"))%>" />
                      </div>
                    </div>
<!--
                    <div class="col-3">
                      <label for="edtMtsCHmax" class="form-label">Máximo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Menor ou igual"><=</span>
                        <input type="number" pattern="[0-9]+$" name="edtMtsCHmax" id="edtMtsCHmax" class="form-control" value="<%=trim(PIR_CORPORATIVO("MtsCHmax"))%>" />
                      </div>
                    </div>
-->
                  </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- Tab LARGURA X COMPRIMENTO ------------------------------------------------------------------------------------------------------------------------------------------------------ -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade" id="tablim" role="tabpanel" aria-labelledby="nav-tablim-tab">

                <div class="section-header">Largura x Comprimento</div>

                  <div class="row justify-content-center">

                    <div class="divider">Largura</div>

                    <div class="col-3">
                      <label for="edtLargMin" class="form-label">Mínimo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Maior ou igual">>=</span>
                        <input type="number" pattern="[0-9]+$" name="edtLargMin" id="edtLargMin" class="form-control" value="<%=trim(PIR_CORPORATIVO("LarguraMin"))%>" />
                      </div>
                    </div>

                    <div class="col-3">
                      <label for="edtLargMax" class="form-label">Máximo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Menor ou igual"><=</span>
                        <input type="number" pattern="[0-9]+$" name="edtLargMax" id="edtLargMax" class="form-control" value="<%=trim(PIR_CORPORATIVO("LarguraMax"))%>" />
                      </div>
                    </div>

                  </div>

                  <div class="row justify-content-center">

                    <div class="divider">Comprimento</div>

                    <div class="col-3">
                      <label for="edtCompMin" class="form-label">Mínimo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Maior ou igual">>=</span>
                        <input type="number" pattern="[0-9]+$" name="edtCompMin" id="edtCompMin" class="form-control" value="<%=trim(PIR_CORPORATIVO("ComprimentoMin"))%>" />
                      </div>
                    </div>

                    <div class="col-3">
                      <label for="edtCompMax" class="form-label">Máximo</label>
                      <div class="input-group">
                        <span class="input-group-text" id="basic-addon1" title="Menor ou igual"><=</span>
                        <input type="number" pattern="[0-9]+$" name="edtCompMax" id="edtCompMax" class="form-control" value="<%=trim(PIR_CORPORATIVO("ComprimentoMax"))%>" />
                      </div>
                    </div>
                  </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- TabBloqueioMaquina ------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade" id="tabmaq" role="tabpanel" aria-labelledby="nav-tabmaq-tab">

                <div class="section-header">Bloqueio Máquina</div>

                <div class="row">
                  <div class="col">
                    <div class="form-check form-switch">
                      <input id="cbxCargaMaquina" name="cbxCargaMaquina" class="form-check-input" type="checkbox" 
                        <% if PIR_CORPORATIVO("CargaMaquina") then %> value="on" <% else %> value="off" <% end if %> 
                        <% if PIR_CORPORATIVO("CargaMaquina") then %> checked <% end if %> 
                        onclick="BloqueiaMaquina(this.checked)">
                      <label for="cbxCargaMaquina" class="form-check-label">Carga Maquina</label>
                    </div>
                  </div>
                </div>

                <div class="row mt-3" id="divBlqMaq">

                  <div class="row g-2 align-items-center">

                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxBlqMaqCX" name="cbxBlqMaqCX" class="form-check-input" type="checkbox" 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaCaixa") then %> value="on" <% else %> value="off" <% end if %> 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaCaixa") then %> checked <% end if %> 
                        onchange="onclickBloqueioMaquina(this)" >
                        <label class="form-check-label" for="cbxBlqMaqCX">Bloquear máquina caixa</label>
                      </div>
                    </div>

                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCXDe" id="edtBlqMaqCXDe" value="<%=edtBlqMaqCXDe %>" maxlength="10" placeholder="De" title="De" class="form-control"/>
                    </div>
                    
                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCXAte" id="edtBlqMaqCXAte" value="<%=edtBlqMaqCXAte %>" maxlength="10" placeholder="Até" title="Até" class="form-control"/>
                    </div>

                  </div>

                  <div class="row g-2 align-items-center mt-3">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxBlqMaqCH" name="cbxBlqMaqCH" class="form-check-input" type="checkbox" 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaChapa") then %> value="on" <% else %> value="off" <% end if %> 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaChapa") then %> checked <% end if %> 
                        onchange="onclickBloqueioMaquina(this)"  >
                        <label class="form-check-label" for="cbxBlqMaqCH">Bloquear máquina chapa</label>
                      </div>
                    </div>

                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCHDe" id="edtBlqMaqCHDe" value="<%=edtBlqMaqCHDe %>" maxlength="10" placeholder="De" class="form-control"/>
                    </div>

                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCHAte" id="edtBlqMaqCHAte" value="<%=edtBlqMaqCHAte %>" maxlength="10" placeholder="Até" class="form-control"/>
                    </div>
                  </div>

                  <div class="row g-2 align-items-center mt-3">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxBlqMaqCV" name="cbxBlqMaqCV" class="form-check-input" type="checkbox" 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaCorteVinco") then %> value="on" <% else %> value="off" <% end if %> 
                        <% if PIR_CORPORATIVO("BloqueiaMaquinaCorteVinco") then %> checked <% end if %> 
                        onchange="onclickBloqueioMaquina(this)"  >
                        <label class="form-check-label" for="cbxBlqMaqCV">Bloquear máquina corte vinco</label>
                      </div>
                    </div>

                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCVDe" id="edtBlqMaqCVDe" value="<%=edtBlqMaqCVDe %>" maxlength="10" placeholder="De" class="form-control"/>
                    </div>

                    <div class="col-3">
                      <input type="date" name="edtBlqMaqCVAte" id="edtBlqMaqCVAte" value="<%=edtBlqMaqCVAte %>" maxlength="10" placeholder="Até" class="form-control"/>
                    </div>
                  </div>

                </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- Preço Mínimo por Quilo --------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade" id="tabprc" role="tabpanel" aria-labelledby="nav-tabprc-tab">

                <div class="section-header">Preço Mínimo por Quilo</div>

                  <div class="row justify-content-center g-3">

                    <div class="col-2">
                      <label class="form-label" for="edtPrcKgCV">Corte Vinco</label>
                      <input 
                        id="edtPrcKgCV" 
                        name="edtPrcKgCV" 
                        type="text inputVal" 
                        value="<%=edtPrcKgCV %>" 
                        class="form-control inputVal"                         
                        onkeypress="return(moeda(this,'.',',',event))"
                        aria-describedby="inputGroupPrepend" 
                        required="required" />

                    </div>
                    <div class="col-2">
                      <label class="form-label" for="edtPrcKgCX">Caixa</label>
                      <input 
                        id="edtPrcKgCX" 
                        name="edtPrcKgCX" 
                        type="text inputVal" 
                        value="<%=edtPrcKgCX %>" 
                        class="form-control inputVal"                         
                        onkeypress="return(moeda(this,'.',',',event))"
                        aria-describedby="inputGroupPrepend" 
                        required="required" />                    
                    </div>
                    <div class="col-2">
                      <label class="form-label" for="edtPrcKgCH">Chapa</label>
                      <input 
                        id="edtPrcKgCH" 
                        name="edtPrcKgCH" 
                        type="text inputVal" 
                        value="<%=edtPrcKgCH %>" 
                        class="form-control inputVal"                         
                        onkeypress="return(moeda(this,'.',',',event))"
                        aria-describedby="inputGroupPrepend" 
                        required="required" />                    
                    </div>

                  </div>

              </div>

              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- SPD ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->
              <!-- -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -->

              <div class="tab-pane fade" id="tabspd" role="tabpanel" aria-labelledby="nav-tabspd-tab">

                <div class="section-header">SPD</div>

                  <div class="row">
                    <div class="col">
                      <div class="form-check form-switch">
                        <input id="cbxAprovacaoComercial" name="cbxAprovacaoComercial" class="form-check-input" type="checkbox" 
                          <% if PIR_CORPORATIVO("SPDAprovacaoComercial") then %> value="on" <% else %> value="off" <% end if %> 
                          <% if PIR_CORPORATIVO("SPDAprovacaoComercial") then %> checked <% end if %> >
                        <label for="cbxAprovacaoComercial" class="form-check-label">Aprovação comercial</label>
                      </div>
                    </div>
                  </div>

              </div>

            </div>

            <div class="form-actions border-top">
              <a name="btnVoltar" id="btnVoltar" class="btn btn-light" onclick="location.href='SelecionaPlantaPedidoParametros.asp'">
                <i class="bi bi-arrow-left me-1"></i>Voltar
              </a>
              <a name="btnGrava" id="btnGrava" class="btn btn-save" onclick="ValidaForm()">Gravar</a>
            </div>

            <script>
              // ------------------------------------------------------------------------------------------
              function ValidaForm() 
              {
                debugger;

                vCV = (document.getElementById('edtPrcKgCV').value != "0" && document.getElementById('edtPrcKgCV').value != "")
                vCX = (document.getElementById('edtPrcKgCX').value != "0" && document.getElementById('edtPrcKgCX').value != "")
                vCH = (document.getElementById('edtPrcKgCH').value != "0" && document.getElementById('edtPrcKgCH').value != "")

/*
                if (document.getElementById('edtPrzMin').value == ''){
                  document.getElementById('edtPrzMin').value = '0';
                }
*/
                document.pedidoparametro.cbxDiasViagem.value = (document.pedidoparametro.cbxDiasViagem.checked) ? 1 : 0;
                document.getElementById('cbxDiasViagem').dispatchEvent(new Event('change'));

                document.pedidoparametro.cbxCargaMaquina.value = (document.pedidoparametro.cbxCargaMaquina.checked) ? 1 : 0;
                document.getElementById('cbxCargaMaquina').dispatchEvent(new Event('change'));

                document.pedidoparametro.cbxLimiteDeCredito.value = (document.pedidoparametro.cbxLimiteDeCredito.checked) ? 1 : 0;
                document.getElementById('cbxLimiteDeCredito').dispatchEvent(new Event('change'));

                document.pedidoparametro.cbxDiasUteis.value = (document.pedidoparametro.cbxDiasUteis.checked) ? 1 : 0;
                document.getElementById('cbxDiasUteis').dispatchEvent(new Event('change'));

                document.pedidoparametro.cbxPrzMinBlqP1.value = (document.pedidoparametro.cbxPrzMinBlqP1.checked) ? 1 : 0;
                document.getElementById('cbxPrzMinBlqP1').dispatchEvent(new Event('change'));

                document.pedidoparametro.cbxAprovacaoComercial.value = (document.pedidoparametro.cbxAprovacaoComercial.checked) ? 1 : 0;
                document.getElementById('cbxAprovacaoComercial').dispatchEvent(new Event('change'));

                // BLOQUEIO DE CARGA MAQUINA CAIXA ------------------------------------------------
                if (document.pedidoparametro.cbxBlqMaqCX.checked) {

                  document.pedidoparametro.cbxBlqMaqCX.value = 1;

                  


                  const data1 = moment(document.pedidoparametro.edtBlqMaqCXDe.value.trim());
                  const data2 = moment(document.pedidoparametro.edtBlqMaqCXAte.value.trim());

                  console.log('document.pedidoparametro.edtBlqMaqCXDe.value',document.pedidoparametro.edtBlqMaqCXDe.value)
                  console.log('document.pedidoparametro.edtBlqMaqCXAte.value',document.pedidoparametro.edtBlqMaqCXAte.value)

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCXDe.value.trim())) {
                    alert("Data inicial INVÁLIDA, para bloqueio de máquina na produção de caixa!");
                    return false;
                  }

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCXAte.value.trim())) {
                    alert("Data final INVÁLIDA, para bloqueio de máquina na produção de caixa!");
                    return false;
                  }

                  if (data2.isBefore(data1)) {
                    alert("Data final é menor que a data inicial, para bloqueio de máquina na produção de caixa!");
                    return false;
                  }

                } else {
                  document.pedidoparametro.cbxBlqMaqCX.value = 0;
                }
                // BLOQUEIO DE CARGA MAQUINA CHAPA ------------------------------------------------
                if (document.pedidoparametro.cbxBlqMaqCH.checked) {
            

                  document.pedidoparametro.cbxBlqMaqCH.value = 1;

                  console.log('document.pedidoparametro.edtBlqMaqCHDe.value',document.pedidoparametro.edtBlqMaqCHDe.value)
                  console.log('document.pedidoparametro.edtBlqMaqCHAte.value',document.pedidoparametro.edtBlqMaqCHAte.value)

                  const data1 = moment(document.pedidoparametro.edtBlqMaqCHDe.value.trim());
                  const data2 = moment(document.pedidoparametro.edtBlqMaqCHAte.value.trim());

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCHDe.value.trim())) {
                    alert("Data inicial INVÁLIDA, para bloqueio de máquina na produção de chapa!");
                    return false;
                  }

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCHAte.value.trim())) {
                    alert("Data final INVÁLIDA, para bloqueio de máquina na produção de chapa!");
                    return false;
                  }

                  if (data2.isBefore(data1)) {
                    alert("Data final é menor que a data inicial, para bloqueio de máquina na produção de chapa!");
                    return false;
                  }

                } else {
                  document.pedidoparametro.cbxBlqMaqCH.value = 0;
                }
                // BLOQUEIO DE CARGA MAQUINA CORTE VINCO ------------------------------------------
                if (document.pedidoparametro.cbxBlqMaqCV.checked) {

                  document.pedidoparametro.cbxBlqMaqCV.value = 1;

                  const data1 = moment(document.pedidoparametro.edtBlqMaqCVDe.value.trim());
                  const data2 = moment(document.pedidoparametro.edtBlqMaqCVAte.value.trim());

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCVDe.value.trim())) {
                    alert("Data inicial INVÁLIDA, para bloqueio de máquina na produção de corte e vinco!");
                    return false;
                  }

                  if (!isValidDateWithMoment(document.pedidoparametro.edtBlqMaqCVAte.value.trim())) {
                    alert("Data final INVÁLIDA, para bloqueio de máquina na produção de corte e vinco!");
                    return false;
                  }

                  if (data2.isBefore(data1)) {
                    alert("Data final é menor que a data inicial, para bloqueio de máquina na produção de corte e vinco!");
                    return false;
                  }

                } else {
                  document.pedidoparametro.cbxBlqMaqCV.value = 0;
                }
                // VERIFICANDO CAMPO NUMERICO EM BRANCO -------------------------------------------
                if (document.pedidoparametro.edtAcrescimoPA.value.length == 0) {
                  document.pedidoparametro.edtAcrescimoPA.value = 0;
                }

                if (document.pedidoparametro.edtAcrescimoCX.value.length == 0) {
                  document.pedidoparametro.edtAcrescimoCX.value = 0;
                }

                if (document.pedidoparametro.edtAcrescimoCH.value.length == 0) {
                  document.pedidoparametro.edtAcrescimoCH.value = 0;
                }

                if (document.pedidoparametro.edtPesoCV.value.length == 0) {
                  document.pedidoparametro.edtPesoCV.value = 0;
                }

                if (document.pedidoparametro.edtPesoCX.value.length == 0) {
                  document.pedidoparametro.edtPesoCX.value = 0;
                }

                if (document.pedidoparametro.edtPesoCH.value.length == 0) {
                  document.pedidoparametro.edtPesoCH.value = 0;
                }

                if (document.pedidoparametro.edtMtsCVmin.value.length == 0) {
                  document.pedidoparametro.edtMtsCVmin.value = 0;
                }

                if (document.pedidoparametro.edtMtsCXmin.value.length == 0) {
                  document.pedidoparametro.edtMtsCXmin.value = 0;
                }

                if (document.pedidoparametro.edtMtsCHmin.value.length == 0) {
                  document.pedidoparametro.edtMtsCHmin.value = 0;
                }
                // LARGURA ------------------------------------------------------------------------
                if (document.pedidoparametro.edtLargMin.value.length == 0) {
                  document.pedidoparametro.edtLargMin.value = 0;
                } 

                if (document.pedidoparametro.edtLargMax.value.length == 0) {
                  document.pedidoparametro.edtLargMax.value = 0;
                }

                LargMin = document.pedidoparametro.edtLargMin.value;
                LargMax = document.pedidoparametro.edtLargMax.value;

                if (Number(LargMax) < Number(LargMin)) {
                  alert("A larguna máxima deve ser maior que a largura mínima!");
                  return false;
                }
                // COMPRIMENTO --------------------------------------------------------------------
                if (document.pedidoparametro.edtCompMin.value.length == 0) {
                  document.pedidoparametro.edtCompMin.value = 0;
                }

                if (document.pedidoparametro.edtCompMax.value.length == 0) {
                  document.pedidoparametro.edtCompMax.value = 0;
                }

                CompMin = document.pedidoparametro.edtCompMin.value.trim();
                CompMax = document.pedidoparametro.edtCompMax.value.trim();

                console.log( CompMin );
                console.log( CompMax );

                if (Number(CompMax) < Number(CompMin)) {
                  alert("O comprimento máximo deve ser maior que a comprimento mínimo!");
                  return false;
                }

                if (document.pedidoparametro.edtPrcKgCV.value.length == 0) {
                  document.pedidoparametro.edtPrcKgCV.value = 0;
                }

                if (document.pedidoparametro.edtPrcKgCX.value.length == 0) {
                  document.pedidoparametro.edtPrcKgCX.value = 0;
                }

                if (document.pedidoparametro.edtPrcKgCH.value.length == 0) {
                  document.pedidoparametro.edtPrcKgCH.value = 0;
                }

                document.getElementById('gravar').value = '1';

                log_register( 'PIR-CORPORATIVO', 'LOG_PARAMETROS_PEDIDOS', '<%=mid(GetFileName(),1,len(GetFileName())-4)%>', '', '<%=session("CODIGO")%>', '<%=session("USR")%>', 'U');

                pedidoparametro.submit();
              }
              // ----------------------------------------------------------------------------------
              function isValidDateWithMoment(dateString) {
                return moment(dateString, 'YYYY-MM-DD').isValid();
              }

            </script>

          </form>

      </div>

      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

  </body>
</html>