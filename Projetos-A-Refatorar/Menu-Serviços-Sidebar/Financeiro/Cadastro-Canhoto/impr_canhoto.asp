<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
  Response.Buffer  = True
  Response.Expires = 0
  Session.LCID     = 1046
  On Error Resume Next

  Dim cboPlantaImp, filialImp
  cboPlantaImp = trim(Request("cboPlanta"))
  filialImp    = mid(cboPlantaImp, 2, 1)
  if filialImp = "" then filialImp = "0"
%>

  <!--#include file="../../cnn2/cnn2.asp" -->
  <!--#include file="../../includes/variaveis.asp" -->
  <!--#include file="../../funcoes/funcoes.asp" -->

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — Imprimir Canhoto</title>

  <link rel="icon" type="image/png" href="/corporativo/CriarPermissao/Logo/logo-smurfit1.png" />

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <script>
  (() => {
    "use strict";

    const PLANTA = "<%=cboPlantaImp%>";

    window.handleImprimir = function() {
      document.getElementById("barraAcoes").style.display = "none";
      window.print();
      document.getElementById("barraAcoes").style.display = "";
      if (window.opener) {
        window.close();
      } else {
        window.location.href = "grid_canhoto.asp?acao=IMP&aviso=ok&cboPlanta=" + encodeURIComponent(PLANTA);
      }
    };

    window.handleVoltar = function() {
      if (window.opener) {
        window.close();
      } else {
        window.location.href = "grid_canhoto.asp?cboPlanta=" + encodeURIComponent(PLANTA);
      }
    };

  })();
  </script>

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
      font-family: var(--bs-font-sans-serif);
    }

    .cupom-wrapper {
      font-family: 'Courier New', Courier, monospace;
    }

    .cupom {
      width: 260px;
      margin-bottom: 1.5rem;
      font-size: 12px;
      line-height: 1.6;
    }
    .cupom-titulo {
      font-weight: 700;
      font-size: 13px;
      margin-bottom: .25rem;
    }
    .cupom-linha {
      border: none;
      border-top: 1px dashed #555;
      margin: .3rem 0;
    }
    .cupom-campo {
      display: flex;
      gap: .25rem;
    }
    .cupom-campo .label {
      white-space: nowrap;
      font-weight: 700;
    }
    .cupom-rodape {
      display: flex;
      align-items: center;
      gap: .4rem;
      margin-top: .3rem;
      font-size: 10px;
    }
    .cupom-rodape img {
      height: 20px;
      width: auto;
    }
    .cupom-tesoura {
      color: #555;
      margin-top: .25rem;
    }

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

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }

    .overlay-loading {
      position: fixed;
      inset: 0;
      background: rgba(255,255,255,.72);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 1055;
      gap: 1rem;
    }
    .overlay-loading .spinner-border { width: 3rem; height: 3rem; }
    .overlay-loading span { font-size: .95rem; font-weight: 600; }

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

    @media print {
      @page {
        size: auto;
        margin: 8mm;
      }
      html, body {
        height: auto !important;
        background-color: #fff !important;
      }
      #barraAcoes,
      .filter-card { display: none !important; }
      .container-fluid { padding: 0 !important; }
      .table-card {
        border: none !important;
        box-shadow: none !important;
      }
      .cupom-wrapper {
        padding: 0 !important;
      }
      .cupom {
        page-break-inside: avoid;
        break-inside: avoid;
        margin-bottom: 0 !important;
      }
    }
  </style>
</head>
<body>

  <div id="overlayLoading" class="overlay-loading d-none">
    <div class="spinner-border text-secondary" role="status"></div>
    <span>Carregando...</span>
  </div>

  <div class="toast-container-custom" id="toastContainer"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- <div class="mb-3">
      <span class="page-title" style="font-size:1.15rem;font-weight:700;color:#212529;">
        Imprimir Canhoto
      </span>
    </div> -->

    <div class="filter-card mb-3" id="barraAcoes">
      <div class="filter-card-header">
        <i class="bi bi-printer me-2"></i>Ações
      </div>
      <div class="filter-card-body">
        <div class="d-flex gap-2">
          <button class="btn btn-light btn-sm" onclick="handleVoltar()">
            <i class="bi bi-arrow-left me"></i>
          </button>
          <button class="btn btn-success btn-sm" onclick="handleImprimir()">
            <i class="bi bi-printer me-2"></i>Imprimir
          </button>
        </div>
      </div>
    </div>

    <div class="table-card">
      <div class="p-4 cupom-wrapper">

<%

  SQL = "SELECT " &_
          "NF, " &_
          "SERIE, " &_
          "Filial, " &_
          "DATA, " &_
          "USERDIGIT, " &_
          "OBS " &_
        "FROM CANHOTOS " &_
        "WHERE concat(NF, SERIE) IN ( " & trim(Request("selecionados")) & " ) " &_
          "AND Filial = '" & filialImp & "' " &_
          "AND D_E_L_E_T_ <> '*' "

  Dim rsCanhoto
  Set rsCanhoto = Conecta("PIR-CORPORATIVO", SQL)

  if not isObject(rsCanhoto) or (rsCanhoto Is Nothing) or rsCanhoto.EOF then
    Response.Write "<div class='table-placeholder'>" &_
                   "<i class='bi bi-clipboard-x' style='font-size:2.8rem;opacity:.35;color:#6c757d;'></i>" &_
                   "<p style='color:#6c757d;font-size:.9rem;margin:.75rem 0 0;'>Nenhum canhoto encontrado para os itens selecionados.</p>" &_
                   "</div>"
    Response.End
  end if

  do while not rsCanhoto.EOF

    Dim titFilial
    select case CStr(rsCanhoto("Filial"))
    case "0" titFilial = "Pirapetinga"
    case "4" titFilial = "Uberaba"
    case "5" titFilial = "Bento Goncalves"
    case "6" titFilial = "Fortaleza"
    case "7" titFilial = "Saquarema"
    case else titFilial = "Filial " & rsCanhoto("Filial")
    end select

    Dim obsTexto
    obsTexto = trim(rsCanhoto("OBS"))

    Dim dataHoje
    dataHoje = Right("0" & DatePart("d", Now()), 2) & "/" &_
               Right("0" & DatePart("m", Now()), 2) & "/" &_
               DatePart("yyyy", Now()) & " " &_
               FormatDateTime(Now(), vbShortTime)

%>
        <div class="cupom">
          <div class="cupom-titulo">Recibo de Canhoto</div>
          <hr class="cupom-linha" />
          <div class="cupom-campo"><span class="label">Nota Fiscal:</span> <%=rsCanhoto("NF")%></div>
          <div class="cupom-campo"><span class="label">Serie......:</span> <%=rsCanhoto("SERIE")%></div>
          <div class="cupom-campo"><span class="label">Filial.....:</span> <%=titFilial%></div>
          <div class="cupom-campo"><span class="label">Data.......:</span> <%=AcertaData(CStr(rsCanhoto("DATA")))%></div>
          <div class="cupom-campo"><span class="label">Usuario....:</span> <%=rsCanhoto("USERDIGIT")%></div>
          <%
          if Len(obsTexto) > 0 then
            Dim obsPos, obsBloco, obsPrimeiro
            obsPos      = 1
            obsPrimeiro = True
            do while obsPos <= Len(obsTexto)
              obsBloco = mid(obsTexto, obsPos, 20)
              if obsPrimeiro then
          %>
          <div class="cupom-campo"><span class="label">Observacao.:</span> <%=obsBloco%></div>
          <%
                obsPrimeiro = False
              else
          %>
          <div class="cupom-campo"><span class="label">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span> <%=obsBloco%></div>
          <%
              end if
              obsPos = obsPos + 20
            loop
          end if
          %>
          <hr class="cupom-linha" />
          <div class="cupom-rodape">
            <img src="../../images/swlogo-90w30h.png" alt="Smurfit Westrock" />
            <span><%=dataHoje%></span>
          </div>
          <div class="cupom-tesoura">&#9988;- - - - - - - - - - - - - -</div>
        </div>
<%

    rsCanhoto.MoveNext()

  loop

  rsCanhoto.Close()

%>

      </div>
    </div>

  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

</body>
</html>

<%
Function AcertaData(InverteData)
  Dim Ano, Mes, Dia
  Ano = mid(InverteData, 1, 4)
  Mes = mid(InverteData, 5, 2)
  Dia = mid(InverteData, 7, 2)
  if Len(Ano) < 4 then Ano = "20" & Ano
  if Len(Mes) = 1 then Mes = "0" & Mes
  if Len(Dia) = 1 then Dia = "0" & Dia
  AcertaData = Dia & "/" & Mes & "/" & Ano
End Function
%>
