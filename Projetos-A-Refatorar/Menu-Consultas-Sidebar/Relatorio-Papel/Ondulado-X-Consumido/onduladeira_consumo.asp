<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
response.Buffer=true
Server.ScriptTimeout = 360
Session.LCID=1046
%>

<!--#include file="../../cnn/cnn.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#Include File="../../includes/variaveis.asp"-->

<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <meta charset="iso-8859-1" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet &#8212; Consumo de Papel &#8212; Onduladeira</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />

    <style>
      body { background-color: #f0f0f0; min-height: 100vh; }
      .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

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
      .filter-card-body { padding: 1rem 1.25rem 1.25rem; }
      .filter-card label {
        font-size: .82rem;
        font-weight: 600;
        color: #6c757d;
        margin-bottom: .25rem;
      }
      .input-data { width: 155px; }
      .filial-select { width: 180px; }

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
    </style>
  </head>

  <body>
    <div id="overlayLoading" class="overlay-loading d-none">
      <div class="spinner-border text-secondary" role="status">
        <span class="visually-hidden">Carregando...</span>
      </div>
      <span>Carregando...</span>
    </div>

    <div class="container-fluid py-3 px-3 px-lg-4">
      <div class="d-flex align-items-center mb-3">
        <span class="page-title">Consumo de Papel &#8212; Onduladeira</span>
      </div>

      <div class="filter-card">
        <div class="filter-card-header">
          <i class="bi bi-funnel me-2"></i>Filtros
        </div>
        <div class="filter-card-body">
          <form id="frmFiltro" action="detalha_oxc.asp" method="get">
            <div class="row g-3 align-items-end">
              <div class="col-auto">
                <label for="filial" class="form-label">Filial</label>
                <select name="filial" id="filial" class="form-select form-select-sm filial-select">
                  <option value="00" <%if request("filial") = "00" or request("filial") = "" then%>selected="selected"<%end if%>>Pirapetinga</option>
                  <option value="04" <%if request("filial") = "04" then%>selected="selected"<%end if%>>Uberaba</option>
                </select>
              </div>

              <div class="col-auto">
                <label for="dataini" class="form-label">Data Inicial</label>
                <input type="date" class="form-control form-control-sm input-data" name="dataini" id="dataini" value="<%=FormatInputDate(request("dataini"))%>" />
              </div>

              <div class="col-auto">
                <label for="datafim" class="form-label">Data Final</label>
                <input type="date" class="form-control form-control-sm input-data" name="datafim" id="datafim" value="<%=FormatInputDate(request("datafim"))%>" />
              </div>

              <div class="col d-flex gap-2 justify-content-end">
                <button type="submit" name="enviar" class="btn btn-primary btn-sm">
                  <i class="bi bi-funnel-fill me-2"></i>Filtrar
                </button>
              </div>
            </div>
          </form>
        </div>
      </div>
    </div>

<%
    Dim Ano, Mes, Dia

    Function FormatInputDate(strData)
      If strData <> "" Then
        Dim partes
        partes = Split(strData, "/")
        If UBound(partes) = 2 Then
          FormatInputDate = partes(2) & "-" & partes(1) & "-" & partes(0)
        Else
          FormatInputDate = strData
        End If
      Else
        FormatInputDate = ""
      End If
    End Function

    Function InverteDataSiga(InverteData)
      Ano = Year(InverteData)
      Mes = Month(InverteData)
      Dia = Day(InverteData)
      If Len(Ano) > 4 Then
        Ano = "20" & Ano
      End If
      If Len(Mes) = 1 Then
        Mes = "0" & Mes
      End If
      If Len(Dia) = 1 Then
        Dia = "0" & Dia
      End If
      InverteDataSiga = Ano & Mes & Dia
    End Function

    Function AcertaData(InverteData)
      Ano = Year(InverteData)
      Mes = Month(InverteData)
      Dia = Day(InverteData)
      If Len(Ano) > 4 Then
        Ano = "20" & Ano
      End If
      If Len(Mes) = 1 Then
        Mes = "0" & Mes
      End If
      If Len(Dia) = 1 Then
        Dia = "0" & Dia
      End If
      AcertaData = Dia & "/" & Mes & "/" & Ano
    End Function
%>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <script>
      (function () {
        "use strict";

        const frm = document.getElementById("frmFiltro");
        const overlay = document.getElementById("overlayLoading");

        frm.addEventListener("submit", function (e) {
          const dataini = document.getElementById("dataini").value;
          const datafim = document.getElementById("datafim").value;

          if (!dataini || !datafim) {
            e.preventDefault();
            alert("Preencha a Data Inicial e a Data Final.");
            return;
          }

          overlay.classList.remove("d-none");
        });
      })();
    </script>
  </body>
</html>
