<%
  SESSION.LCID=1046
  response.Buffer=true
%>

<!--#include file="../../cnn/cnn.asp"-->
<!--#include file="../../includes/variaveis.asp" -->

<!DOCTYPE html>
<html lang="pt-br">

  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet — Detalhado: Consumo de Papel</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />

    <style>
      body { background-color: #f0f0f0; min-height: 100vh; }
      .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

      .filter-card,
      .table-card {
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

      .info-label {
        font-size: .72rem;
        font-weight: 600;
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: .03rem;
        margin-bottom: .15rem;
      }
      .info-value {
        font-size: .9rem;
        font-weight: 600;
        color: #212529;
      }

      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }
      .table-wrapper table { min-width: max-content; margin-bottom: 0; }

      .table thead th {
        position: sticky;
        top: 0;
        background-color: #fff;
        color: #212529;
        font-size: .82rem;
        font-weight: 700;
        text-align: center;
        white-space: nowrap;
        z-index: 2;
        border-color: #dee2e6;
        padding: .55rem .75rem;
      }
      .table tbody td {
        font-size: .82rem;
        text-align: center;
        white-space: nowrap;
        color: #212529;
        vertical-align: middle;
        padding: .45rem .75rem;
      }
      .table tbody th {
        font-size: .82rem;
        text-align: center;
        white-space: nowrap;
        vertical-align: middle;
        padding: .45rem .75rem;
      }

      .tr-total th {
        background-color: #4a4a4a !important;
        color: #fff !important;
        font-weight: 700;
      }
      .tr-spacer td {
        background-color: transparent !important;
        border: none !important;
        padding: .35rem 0 !important;
      }
      .tr-distribuicao th {
        background-color: #d4d4d4 !important;
        color: #212529 !important;
        font-weight: 600;
      }
      .tr-distribuicao-total th {
        background-color: #4a4a4a !important;
        color: #fff !important;
        font-weight: 700;
      }
      .cel-num { text-align: right !important; }
      .cel-empty { background-color: #fff !important; border-color: transparent !important; }

      .table-placeholder {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 3.5rem 1rem;
        color: #6c757d;
        gap: .75rem;
      }
      .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
      .table-placeholder p { margin: 0; font-size: .9rem; text-align: center; }

      .filial-badge-pira {
        background-color: #d1e7ff;
        color: #084298;
        font-size: .88rem;
        font-weight: 700;
        padding: .35em .7em;
        border-radius: 4px;
      }
      .filial-badge-uber {
        background-color: #d1f0e0;
        color: #0a4d28;
        font-size: .88rem;
        font-weight: 700;
        padding: .35em .7em;
        border-radius: 4px;
      }
    </style>
  </head>

  <body>
<%
    DIM datafimpap, datainipap, SQL, celula, papel
    if (request("dataini") <> "") and (request("datafim") <> "") then

      if request("excel") = 1 then
        response.ContentType="application/vnd.ms-excel"
        response.AddHeader "content-disposition", "inline; filename=arquivo.xls"
      end if

      dataInicial = dateadd("d",-1,request("dataini"))

      if request("filial") = "00" then

        if time() > "22:29:00" then
          datafimpap = InverteDataSiga(request("datafim"))&" 22:29:00"
        else
          datafimpap = InverteDataSiga(request("datafim"))&" "&time()
        end if

      else

        if time() > "22:29:00" then
          datafimpap = InverteDataSiga(dateadd("d",1,request("datafim")))&" 22:29:00"
        else
          datafimpap = InverteDataSiga(request("datafim"))&" "&time()
        end if

      end if

      datainipap = InverteDataSiga(dataInicial)&" 22:30:00"

      if request("filial") = "00" then

        Set Lista = Dados51
        SQL = "EXEC SP_ONDULADO_CONSUMIDO '"&InverteDataSiga(request("dataini"))&"','"&InverteDataSiga(request("datafim"))&"','"&datainipap&"','"&datafimpap&"',''"
        Lista.Open(SQL)

        Set Lista2 = Dados52

        SQL = "SELECT " & _
                "Papel, " & _
                "SUM(KG_Teorico_Pedido) KG_Teorico_Pedido, " & _
                "SUM(KG_Teorico_Ondulado) KG_Teorico_Ondulado, " & _
                "SUM(KG_Bobinas_Consumidas) KG_Bobinas_Consumidas, " & _
                "SUM(TeoricoPedido_X_Consumido) TeoricoPedido_X_Consumido, " & _
                "SUM(TeoricoOndu_X_Consumido) TeoricoOndu_X_Consumido, " & _
                "MAX(Gram_Media) Gram_Media, " & _
                "MAX(Gram_Teorica) Gram_Teorica " & _
              "FROM ONDULADO_CONSUMIDO " & _
              "GROUP BY Papel "
        Lista2.Open(SQL)

      else

        Set Lista = Dados61
        SQL = "EXEC SP_ONDULADO_CONSUMIDO '"&InverteDataSiga(request("dataini"))&"','"&InverteDataSiga(request("datafim"))&"','"&datainipap&"','"&datafimpap&"',''"
        Lista.Open(SQL)

        Set Lista2 = Dados62

        SQL = "SELECT " & _
                "Papel, " & _
                "SUM(KG_Teorico_Pedido) KG_Teorico_Pedido, " & _
                "SUM(KG_Teorico_Ondulado) KG_Teorico_Ondulado, " & _
                "SUM(KG_Bobinas_Consumidas) KG_Bobinas_Consumidas, " & _
                "SUM(TeoricoPedido_X_Consumido) TeoricoPedido_X_Consumido, " & _
                "SUM(TeoricoOndu_X_Consumido) TeoricoOndu_X_Consumido, " & _
                "MAX(Gram_Media) Gram_Media, " & _
                "MAX(Gram_Teorica) Gram_Teorica " & _
              "FROM ONDULADO_CONSUMIDO " & _
              "GROUP BY Papel "
        Lista2.Open(SQL)

      end if
%>

    <div class="container-fluid py-3 px-3 px-lg-4">

      <div class="d-flex align-items-center mb-3">
        <a href="onduladeira_consumo.asp" class="btn btn-outline-secondary btn-sm me-3" title="Voltar">
          <i class="bi bi-arrow-left"></i>
        </a>
        <span class="page-title">Detalhado — Consumo de Papel</span>
      </div>

      <div class="d-none">
        <span id="infoDataEmissao"><%=date()%></span>
        <span id="infoHoraEmissao"><%=time()%></span>
        <span id="infoDataInicial"><%=request("dataini")%></span>
        <span id="infoDataFinal"><%=request("datafim")%></span>
        <span id="infoFilial"><%if request("filial") = "00" then%>Pirapetinga<%else%>Uberaba<%end if%></span>
      </div>

      <div class="filter-card mb-3">
        <div class="filter-card-header d-flex align-items-center justify-content-between">
          Informações do Relatório
          <div class="d-flex gap-2">
            <button type="button" class="btn btn-success btn-sm" id="btnExcel" title="Exportar Excel">
              <i class="bi bi-file-earmark-excel me-1"></i>Excel
            </button>
            <button type="button" class="btn btn-danger btn-sm" id="btnPdf" title="Exportar PDF">
              <i class="bi bi-file-earmark-pdf me-1"></i>PDF
            </button>
          </div>
        </div>
        <div class="filter-card-body">
          <div class="row g-4">
            <div class="col-auto">
              <div class="info-label">Filial</div>
              <%if request("filial") = "00" then%>
              <div class="info-value">
                <span class="badge filial-badge-pira">Pirapetinga</span>
              </div>
              <%else%>
              <div class="info-value">
                <span class="badge filial-badge-uber">Uberaba</span>
              </div>
              <%end if%>
            </div>
            <div class="col-auto">
              <div class="info-label">Data Inicial</div>
              <div class="info-value"><%=FormatDisplayDate(request("dataini"))%></div>
            </div>
            <div class="col-auto">
              <div class="info-label">Data Final</div>
              <div class="info-value"><%=FormatDisplayDate(request("datafim"))%></div>
            </div>
            <div class="col-auto">
              <div class="info-label">Emissão</div>
              <div class="info-value"><%=date()%>&nbsp;<%=time()%></div>
            </div>
          </div>
        </div>
      </div>

<%
      if Lista2.eof then
%>
      <div class="table-card">
        <div class="table-placeholder">
          <i class="bi bi-clipboard-data"></i>
          <p>Não existem informações para gerar este relatório.</p>
        </div>
      </div>
    </div>
<%
        response.End
      end if
%>

      <div class="table-card">
        <div class="table-wrapper">
          <table id="tblRelatorio" class="table table-bordered table-striped table-hover table-sm mb-0">
            <thead>
              <tr>
                <th rowspan="2">Papel</th>
                <th rowspan="2">Gram. Média<br />Especificada</th>
                <th rowspan="2">Gram. Média<br />Consumida</th>
                <th colspan="4">Peso (kg)</th>
              </tr>
              <tr>
                <th>Teórico Espec.</th>
                <th>Teórico Ondul.</th>
                <th>Real (Bobinas<br />Consumidas)</th>
                <th>Teórico Espec.<br />x Consumido</th>
              </tr>
            </thead>
            <tbody>
<%
      up_papel = 0
      up_prog = 0
      celula = ""
      papel = ""
      Total_consumo = 0

      while not Lista2.eof

        if trim(Lista2("Papel")) = "_Total" then

          celula = "th"
          papel = "Total"
          Total_consumo = Lista2("KG_Bobinas_Consumidas")
          Total_KG_Teorico_Pedido = Lista2("KG_Teorico_Pedido")
          Total_Pedidoxconsumido = Lista2("TeoricoPedido_X_Consumido")

        else

          celula = "td"
          papel = trim(Lista2("Papel"))

          if (Lista2("Gram_Media") <> 0) and (Lista2("Gram_Teorica") <> 0) then
            up_papel = up_papel+(((Lista2("Gram_Media")/Lista2("Gram_Teorica"))-1)* Lista2("KG_Bobinas_Consumidas"))
          end if

          up_prog = up_prog+(Lista2("KG_Teorico_Ondulado") - Lista2("KG_Teorico_Pedido"))

        end if
%>
              <tr<% if papel = "Total" then %> class="tr-total"<% end if %>>
                <<%=celula%>><%=papel%></<%=celula %>>
                <%if papel <> "Total" then %>
                <<%=celula%> class="cel-num"><%=Lista2("Gram_Teorica")%></<%=celula%>>
                <<%=celula%> class="cel-num"><%=Lista2("Gram_Media")%></<%=celula%>>
                <%else %>
                <<%=celula%>></<%=celula%>>
                <<%=celula%>></<%=celula%>>
                <%end if %>

                <% if Trim(Lista2("KG_Teorico_Pedido")) <> "" then%>
                <<%=celula%> class="cel-num"><%=formatnumber(Lista2("KG_Teorico_Pedido"),0)%></<%=celula%>>
                <%else %>
                <<%=celula%> class="cel-num">0</<%=celula%>>
                <%end if %>

                <% if Trim(Lista2("KG_Teorico_Ondulado")) <> "" then%>
                <<%=celula%> class="cel-num"><%=formatnumber(Lista2("KG_Teorico_Ondulado"),0)%></<%=celula%>>
                <%else %>
                <<%=celula%> class="cel-num">0</<%=celula%>>
                <%end if %>

                <% if Trim(Lista2("KG_Bobinas_Consumidas")) <> "" then%>
                <<%=celula%> class="cel-num"><%=formatnumber(Lista2("KG_Bobinas_Consumidas"),0)%></<%=celula%>>
                <%else %>
                <<%=celula%> class="cel-num">0</<%=celula%>>
                <%end if %>

                <% if Trim(Lista2("TeoricoPedido_X_Consumido")) <> "" then%>
                <<%=celula%> class="cel-num"><%=formatnumber(Lista2("TeoricoPedido_X_Consumido"),0)%></<%=celula%>>
                <%else %>
                <<%=celula%> class="cel-num">0</<%=celula%>>
                <%end if %>
              </tr>
<%
        Lista2.MoveNext

        response.Flush

      wend

      Lista2.close
%>
              <tr class="tr-spacer">
                <td colspan="7"></td>
              </tr>
              <tr class="tr-distribuicao">
                <th colspan="3" class="cel-empty"></th>
                <th colspan="2">Distribuição do consumo a maior</th>
                <th>Percentual<br />(Peso Real)</th>
                <th>Peso (kg)</th>
              </tr>
              <tr>
                <th colspan="3" class="cel-empty"></th>
                <th colspan="2">Papel</th>
                <td class="cel-num"><%if Total_KG_Teorico_Pedido > 0 then%><%=formatnumber((up_papel*100)/Total_KG_Teorico_Pedido,2) %><%else %>0<%end if %>%</td>
                <td class="cel-num"><%=formatnumber(up_papel,0)%></td>
              </tr>
              <tr>
                <th colspan="3" class="cel-empty"></th>
                <th colspan="2">Programação</th>
                <td class="cel-num"><%if Total_KG_Teorico_Pedido > 0 then%><%=formatnumber((up_prog*100)/Total_KG_Teorico_Pedido,2) %><%else %>0<%end if %>%</td>
                <td class="cel-num"><%=formatnumber(up_prog,0)%></td>
              </tr>
              <tr>
                <th colspan="3" class="cel-empty"></th>
                <th colspan="2">Produção</th>
                <td class="cel-num"><%if Total_KG_Teorico_Pedido > 0 then%><%=formatnumber(((Total_Pedidoxconsumido-up_papel-up_prog)*100)/Total_KG_Teorico_Pedido,2) %><%else %>0<%end if %>%</td>
                <td class="cel-num"><%=formatnumber(Total_Pedidoxconsumido-up_papel-up_prog,0) %></td>
              </tr>
              <tr class="tr-distribuicao-total">
                <th colspan="3" class="cel-empty"></th>
                <th colspan="2">Total</th>
                <th class="cel-num"><%if Total_KG_Teorico_Pedido > 0 then%><%=formatnumber((Total_Pedidoxconsumido*100)/Total_KG_Teorico_Pedido,2) %><%else %>0<%end if %>%</th>
                <th class="cel-num"><%=formatnumber(Total_Pedidoxconsumido,0) %></th>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>

    <script>
      (function () {
        "use strict";

        const LOGO_PATH = "Logo/logo-smurfit1.png";

        const $ = (sel) => document.querySelector(sel);

        async function carregarLogo(src) {
          return new Promise((resolve) => {
            const img = new Image();
            img.crossOrigin = "anonymous";
            img.onload = () => {
              const canvas = document.createElement("canvas");
              const targetW = 900;
              const ratio = img.height / img.width;
              canvas.width = targetW;
              canvas.height = Math.round(targetW * ratio);
              const ctx = canvas.getContext("2d");
              ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
              resolve({ dataUrl: canvas.toDataURL("image/png"), ratio: ratio });
            };
            img.onerror = () => resolve(null);
            img.src = src;
          });
        }

        function lerTabela() {
          const tabela = document.getElementById("tblRelatorio");
          if (!tabela) return { head: [], body: [] };
          const head = [];
          tabela.querySelectorAll("thead tr").forEach((tr) => {
            const linha = [];
            tr.querySelectorAll("th").forEach((th) => {
              linha.push({
                content: th.textContent.replace(/\s+/g, " ").trim(),
                colSpan: parseInt(th.getAttribute("colspan") || "1", 10),
                rowSpan: parseInt(th.getAttribute("rowspan") || "1", 10)
              });
            });
            head.push(linha);
          });
          const body = [];
          tabela.querySelectorAll("tbody tr").forEach((tr) => {
            const linha = [];
            const ehTotal = tr.classList.contains("tr-total");
            const ehDist = tr.classList.contains("tr-distribuicao");
            const ehDistTotal = tr.classList.contains("tr-distribuicao-total");
            const ehSpacer = tr.classList.contains("tr-spacer");
            tr.querySelectorAll("td, th").forEach((cel) => {
              const styles = {};
              if (ehTotal || ehDistTotal) {
                styles.fillColor = [74, 74, 74];
                styles.textColor = [255, 255, 255];
                styles.fontStyle = "bold";
              } else if (ehDist) {
                styles.fillColor = [212, 212, 212];
                styles.textColor = [33, 37, 41];
                styles.fontStyle = "bold";
              } else if (ehSpacer) {
                styles.fillColor = [255, 255, 255];
                styles.lineWidth = 0;
              }
              linha.push({
                content: cel.textContent.replace(/\s+/g, " ").trim(),
                colSpan: parseInt(cel.getAttribute("colspan") || "1", 10),
                rowSpan: parseInt(cel.getAttribute("rowspan") || "1", 10),
                styles: styles
              });
            });
            body.push(linha);
          });
          return { head: head, body: body };
        }

        async function gerarPdf() {
          const { jsPDF } = window.jspdf;
          const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
          const pageW = doc.internal.pageSize.getWidth();

          const logo = await carregarLogo(LOGO_PATH);
          if (logo) {
            const larguraMm = 38;
            const alturaMm = larguraMm * logo.ratio;
            doc.addImage(logo.dataUrl, "PNG", 10, 4, larguraMm, alturaMm);
          }

          doc.setFontSize(13);
          doc.setFont("helvetica", "bold");
          doc.text("Detalhado — Consumo de Papel", pageW / 2, 14, { align: "center" });

          const dataIni = ($("#infoDataInicial")?.textContent || "").trim();
          const dataFim = ($("#infoDataFinal")?.textContent || "").trim();
          const filial = ($("#infoFilial")?.textContent || "").trim();
          const dataEmissao = ($("#infoDataEmissao")?.textContent || "").trim();
          const horaEmissao = ($("#infoHoraEmissao")?.textContent || "").trim();

          doc.setFontSize(8);
          doc.setFont("helvetica", "normal");
          doc.text(
            `Período: ${dataIni} a ${dataFim} — Filial: ${filial}`,
            pageW / 2,
            20,
            { align: "center" }
          );

          doc.setFontSize(7);
          doc.text(
            `Emitido em ${dataEmissao} ${horaEmissao}`,
            pageW - 10,
            8,
            { align: "right" }
          );

          const { head, body } = lerTabela();

          doc.autoTable({
            head: head,
            body: body,
            startY: 26,
            theme: "grid",
            styles: {
              fontSize: 4.5,
              cellPadding: 1.5,
              lineColor: [180, 180, 180],
              lineWidth: 0.2,
              halign: "center",
              valign: "middle"
            },
            headStyles: {
              fillColor: [74, 74, 74],
              textColor: [255, 255, 255],
              fontStyle: "bold",
              halign: "center"
            },
            alternateRowStyles: { fillColor: [245, 245, 245] }
          });

          const nomeArquivo = `consumo_papel_${dataIni.replace(/\//g, "-")}_a_${dataFim.replace(/\//g, "-")}.pdf`;
          doc.save(nomeArquivo);
        }

        function gerarExcel() {
          const tabela = document.getElementById("tblRelatorio");
          if (!tabela) return;

          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.table_to_sheet(tabela, { raw: false });

          const dataIni = (document.getElementById("infoDataInicial")?.textContent || "").trim();
          const dataFim = (document.getElementById("infoDataFinal")?.textContent || "").trim();
          const filial = (document.getElementById("infoFilial")?.textContent || "").trim();

          XLSX.utils.sheet_add_aoa(ws, [
            ["Detalhado — Consumo de Papel"],
            [`Período: ${dataIni} a ${dataFim} — Filial: ${filial}`],
            []
          ], { origin: "A1", shiftRows: true });

          XLSX.utils.book_append_sheet(wb, ws, "Consumo de Papel");

          const nomeArquivo = `consumo_papel_${dataIni.replace(/\//g, "-")}_a_${dataFim.replace(/\//g, "-")}.xlsx`;
          XLSX.writeFile(wb, nomeArquivo);
        }

        document.addEventListener("DOMContentLoaded", () => {
          const btnPdf = document.getElementById("btnPdf");
          if (btnPdf) {
            btnPdf.addEventListener("click", () => {
              btnPdf.disabled = true;
              gerarPdf().finally(() => { btnPdf.disabled = false; });
            });
          }

          const btnExcel = document.getElementById("btnExcel");
          if (btnExcel) {
            btnExcel.addEventListener("click", () => {
              gerarExcel();
            });
          }
        });
      })();
    </script>
<%
    else
%>
    <div class="container-fluid py-3 px-3 px-lg-4">
      <div class="d-flex align-items-center mb-3">
        <a href="onduladeira_consumo.asp" class="btn btn-outline-secondary btn-sm me-3" title="Voltar">
          <i class="bi bi-arrow-left me-1"></i>
        </a>
        <span class="page-title">Detalhado — Consumo de Papel</span>
      </div>
      <div class="alert alert-warning d-flex align-items-center mb-0" role="alert">
        <i class="bi bi-exclamation-triangle-fill me-2"></i>
        <div>Clique em <strong>Voltar</strong> e preencha os campos <strong>Data Inicial</strong> e <strong>Data Final</strong>.</div>
      </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
<%
    end if
%>
  </body>
</html>
<%
Dim Ano, Mes, Dia

function InverteDataSiga(InverteData)
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
  InverteDataSiga = Ano &"-"& Mes &"-"& Dia
end function

function FormatDisplayDate(strData)
  if strData = "" then
    FormatDisplayDate = ""
    exit function
  end if
  Dim partes
  partes = Split(strData, "-")
  if UBound(partes) = 2 then
    FormatDisplayDate = partes(2) & "/" & partes(1) & "/" & partes(0)
  else
    FormatDisplayDate = strData
  end if
end function

function Matricula(mat)
  if len(mat) < 6 then
    zero = 6 - len(mat)
    cont = 0
    matricula = mat
    while cont < zero
      matricula = "0"&matricula
      cont = cont + 1
    wend
  else
    matricula = mat
  end if
end function

function AcertaData(InverteData)
  Ano = Mid(InverteData,1,4)
  Mes = Mid(InverteData,5,2)
  Dia = Mid(InverteData,7,2)
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
end function
%>
