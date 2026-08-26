<!--#include file = "funcoes/funcoes.asp"-->
<!--#include file = "cnn2/cnn2.asp"-->
<%
if trim(Request("id4")) <> "1" then
%>
<!--#include file = "includes/variaveis.asp"-->
<%
end if
%>
<%
Function EscapeJS(str)
  Dim i, c, code, result
  result = ""
  For i = 1 To Len(str)
    c    = Mid(str, i, 1)
    code = AscW(c)
    If code < 0 Then code = code + 65536  ' corrige AscW negativo em arquivos Windows-1252
    If code > 127 Then
      result = result & "\u" & Right("0000" & Hex(code), 4)
    ElseIf c = "'" Then
      result = result & "\'"
    ElseIf c = "\" Then
      result = result & "\\"
    ElseIf code = 13 Or code = 10 Then
      ' ignora CR/LF
    Else
      result = result & c
    End If
  Next
  EscapeJS = result
End Function
%>
<%
SESSION.LCID = 1046

if trim(request("id4")) = "1" then
  Nivel = 0
end if

if CODUSER = "" then
  codsiga = " "
else
  codsiga = CODUSER
end if

if trim(request("parametro")) <> "" then

  select case request("id6")

  case "00", "04" 'PIRAPETINGA, UBERABA

    SQL = "" &_
      "SELECT " &_
        "A1_COD, " &_
        "A1_LOJA, " &_
        "A1_CODSAP, " &_
        "A1_MSBLQL, " &_
        "A1_NOME, " &_
        "A1_NREDUZ, " &_
        "A1_MUN, " &_
        "A1_EST, " &_
        "A1_BAIRRO, " &_
        "A1_VEND, " &_
        "ISNULL(A3_NOME, '') AS A3_NOME, " &_
        "A1_CGC Cnpj, " &_
        "0 as CONDICAO_PAGAMENTO, " &_
        "A1_DIASVIA AS DIAS_TRANSPORTE, " &_
        "'' AS TIPO_FRETE " &_
      "FROM SA1010 s1 " &_
      "LEFT JOIN SA3010 s3 on s3.A3_COD = s1.A1_VEND " &_
      "WHERE ((A1_NREDUZ like '%" & Ucase(trim(Request("parametro"))) & "%') OR (A1_NOME like '%" & Ucase(trim(Request("parametro"))) & "%') OR (A1_COD = '" & trim(Request("parametro")) & "')) " &_
        "AND s1.D_E_L_E_T_ = '' " &_
        "AND A1_MSBLQL <> '1' " &_
        iif( (Nivel <> 0), "AND A1_VEND = '" & codsiga & "' ", "" ) &_
      "ORDER BY A1_NREDUZ "

  case "05", "06" 'BENTO - MARANGUAPE

    SQL = "" &_
      "SELECT " &_
        "T1.CODIGO_ERP AS A1_COD, " &_
        "T1.LOJA_CLIENTE AS A1_LOJA, " &_
        "'COD_SAP_CLIENTE' AS A1_CODSAP, " &_
        "'2' AS A1_MSBLQL, " &_
        "T1.NOME_CLIENTE AS A1_NOME, " &_
        "T1.FANTASIA_CLIENTE AS A1_NREDUZ, " &_
        "T1.MUNICIPIO_CLIENTE AS A1_MUN, " &_
        "T1.ESTADO_CLIENTE AS A1_EST, " &_
        "T1.A1_BAIRRO, " &_
        "T1.COD_VENDEDOR AS A1_VEND, " &_
        "ISNULL(T1.NOME_VENDEDOR, '') AS A3_NOME, " &_
        "T1.CNPJ_CLIENTE AS Cnpj, " &_
        "T1.CONDICAO_PAGAMENTO, " &_
        "T1.DIAS_TRANSPORTE, " &_
        "CASE WHEN T2.TIPO_FRETE = 'C' THEN '0' ELSE '1' END as TIPO_FRETE " &_
      "FROM VW_BI_CLIENTES T1 " &_
      "LEFT JOIN VW_CLIENTES_PEDIDO T2 ON T2.COD_CLIENTE = T1.CODIGO_ERP " &_
      "WHERE ((T1.FANTASIA_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') " &_
        "OR (T1.NOME_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') " &_
        iif( IsNumeric(Request("parametro")), " OR (T1.CODIGO_ERP = " & trim(Request("parametro")) & ")", "" ) &_
        " )" &_
        "AND T1.STATUS = 'ATIVO' " &_
        iif( (Nivel <> 0), "AND T1.COD_VENDEDOR = '" & codsiga & "' ", "" ) &_
      "ORDER BY T1.FANTASIA_CLIENTE "

  case "07" 'SAQUAREMA

    SQL = "" &_
      "SELECT " &_
        "COD_CLIENTE AS A1_COD, " &_
        "LOJA_CLIENTE AS A1_LOJA, " &_
        "'COD_SAP_CLIENTE' AS A1_CODSAP, " &_
        "'2' AS A1_MSBLQL, " &_
        "NOME_CLIENTE AS A1_NOME, " &_
        "FANTASIA_CLIENTE AS A1_NREDUZ, " &_
        "MUNICIPIO_CLIENTE AS A1_MUN, " &_
        "ESTADO_CLIENTE AS A1_EST, " &_
        "'' AS A1_BAIRRO, " &_
        "COD_VENDEDOR AS A1_VEND, " &_
        "ISNULL(NOME_VENDEDOR, '') AS A3_NOME, " &_
        "CNPJ_CLIENTE AS Cnpj, " &_
        "CONDICAO_PAGAMENTO, " &_
        "'' AS TIPO_FRETE, " &_
        "DIAS_TRANSPORTE " &_
      "FROM VW_BI_CLIENTES " &_
      "WHERE ((FANTASIA_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') OR (NOME_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') OR (COD_CLIENTE = '" & trim(Request("parametro")) & "')) " &_
        "AND STATUS = 'ATIVO' " &_
        iif( (Nivel <> 0), "AND COD_VENDEDOR = '" & codsiga & "' ", "" ) &_
      "ORDER BY FANTASIA_CLIENTE "

  ' Plantas Westrock
  case else

    SQL = "" &_
      "SELECT " &_
        "CODIGO_SAP AS A1_COD, " &_
        "'01' AS A1_LOJA, " &_
        "CODIGO_SAP AS A1_CODSAP, " &_
        "'2' AS A1_MSBLQL, " &_
        "NOME_CLIENTE AS A1_NOME, " &_
        "NOME_REDUZIDO_CLI AS A1_NREDUZ, " &_
        "CIDADE_CLIENTE AS A1_MUN, " &_
        "ESTADO_CLIENTE AS A1_EST, " &_
        "'' AS A1_BAIRRO, " &_
        "'' AS A1_VEND, " &_
        "'' AS A3_NOME, " &_
        "CNPJ AS Cnpj, " &_
        "CONDICAO_PAGAMENTO, " &_
        "'' AS TIPO_FRETE, " &_
        "'' as DIAS_TRANSPORTE " &_
      "FROM BI_CLIENTES_WRK " &_
      "WHERE ((NOME_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') OR (NOME_CLIENTE like '%" & Ucase(trim(Request("parametro"))) & "%') OR (CODIGO_SAP = '" & trim(Request("parametro")) & "')) " &_
        "AND STATUS_CLI = 'Active' " &_
        iif( (Nivel <> 0), "AND CODIGO_SAP = '" & codsiga & "' ", "" ) &_
      "ORDER BY NOME_CLIENTE"

  end select

  select case request("id6")
  case "00" SET Consulta = Conecta( "PIR-DADOSADV", SQL )
  case "04" SET Consulta = Conecta( "PIR-DADOSADV", SQL )
  case "05" SET Consulta = Conecta( "BEN-DADOSADV", SQL )
  case "06" SET Consulta = Conecta( "FOR-DADOSADV", SQL )
  case "07" SET Consulta = Conecta( "SAQ-DADOSADV", SQL )
  ' Plantas Westrock
  case else SET Consulta = Conecta( "SPO-BI", SQL )
  end select

  if not isObject(Consulta) then
    response.end
  end if

end if
%>
<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />
    <title>SWnet - Busca Cliente</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
      body {
        background-color: #f0f0f0;
        min-height: 100vh;
      }

      /* Título */
      .page-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #212529;
      }

      /* Card de filtro */
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
      .filter-card-body label {
        font-size: .82rem;
        font-weight: 600;
        color: #6c757d;
        margin-bottom: .25rem;
      }

      /* ── Card de tabela ── */
      .table-card {
        background-color: #fff;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,.07);
        overflow: hidden; /* garante border-radius nos cantos da tabela */
      }

      /* ── Wrapper de scroll ── */
      .table-wrapper {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 68vh;
      }
      .table-wrapper table {
        min-width: max-content;
        margin-bottom: 0;
      }

      /* ── Cabeçalho sticky ── */
      .table thead th {
        position: sticky;
        top: 0;
        z-index: 2;
        background-color: #4a4a4a;
        color: #fff;
        font-size: .78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .04em;
        white-space: nowrap;
        text-align: center;
        vertical-align: middle;
        padding: .6rem .75rem;
        border-color: #5a5a5a;
      }
      /* sombra separadora quando há scroll vertical */
      .table thead th::after {
        content: '';
        position: absolute;
        left: 0; right: 0; bottom: -3px;
        height: 3px;
        background: linear-gradient(to bottom, rgba(0,0,0,.12), transparent);
        pointer-events: none;
      }

      /* ── Células ── */
      .table tbody td {
        font-size: .82rem;
        text-align: center;
        white-space: nowrap;
        color: #212529;
        vertical-align: middle;
        padding: .45rem .75rem;
        border-color: #e8e8e8;
      }

      /* ── Hover nas linhas ── */
      .table-hover tbody tr:hover td {
        background-color: #f0f4ff !important;
        color: #1a1a2e;
      }

      /* ── Linha de alerta: cliente bloqueado / sem SAP ── */
      .table tbody tr.Alert td {
        background-color: #fff8e1 !important;
        color: #7d5a00 !important;
        border-color: #ffe082 !important;
      }
      .table tbody tr.Alert:hover td {
        background-color: #fff3cd !important;
      }

      /* ── Botão de seleção dentro da célula ── */
      .btn-selecionar {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        padding: 0;
        border-radius: 5px;
        background-color: #0d6efd;
        color: #fff;
        border: none;
        cursor: pointer;
        transition: background-color .15s ease, transform .1s ease;
        font-size: .85rem;
      }
      .btn-selecionar:hover {
        background-color: #0b5ed7;
        transform: translateY(-1px);
      }
      .btn-selecionar:active {
        transform: scale(.95);
      }

      /* ── Placeholder (sem busca / sem resultado) ── */
      .table-placeholder {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 4rem 1rem;
        color: #adb5bd;
        gap: .75rem;
      }
      .table-placeholder .bi {
        font-size: 3rem;
        opacity: .4;
      }
      .table-placeholder p {
        margin: 0;
        font-size: .88rem;
        text-align: center;
        color: #6c757d;
      }
    </style>
  </head>

  <body>
    <div class="container-fluid py-3 px-3 px-lg-4">

      <!-- Título -->
      <div class="mb-3">
        <span class="page-title">Consulta Cliente</span>
      </div>

      <!-- Card de Filtros -->
      <div class="filter-card mb-3">
        <div class="filter-card-header">
          <i class="bi bi-funnel me-2"></i>Filtros
        </div>
        <div class="filter-card-body">
          <form id="frmBuscaCliente" name="frmBuscaCliente" method="post" action="buscaCLIENTE.asp">

            <input type="hidden" name="id1" id="id1" value="<%=trim(request("id1"))%>" />
            <input type="hidden" name="id2" id="id2" value="<%=trim(request("id2"))%>" />
            <input type="hidden" name="id3" id="id3" value="<%=trim(request("id3"))%>" />
            <input type="hidden" name="id4" id="id4" value="<%=trim(request("id4"))%>" />
            <input type="hidden" name="id5" id="id5" value="<%=trim(request("id5"))%>" />
            <input type="hidden" name="id6" id="id6" value="<%=trim(request("id6"))%>" />
            <input type="hidden" name="id7" id="id7" value="<%=trim(request("id7"))%>" />
            <input type="hidden" name="id8" id="id8" value="<%=trim(request("id8"))%>" />
            <input type="hidden" name="id9" id="id9" value="<%=trim(request("id9"))%>" />
            <input type="hidden" name="id10" id="id10" value="<%=trim(request("id10"))%>" />

            <label for="parametro" class="form-label">Nome ou código</label>
            <div class="d-flex gap-2">
              <input type="text"
                     class="form-control form-control-sm flex-grow-1"
                     id="parametro"
                     name="parametro"
                     autocomplete="off" />
              <button type="button" class="btn btn-primary btn-sm flex-shrink-0" onclick="filtra()">
                <i class="bi bi-funnel-fill me-1"></i>Filtrar
              </button>
            </div>

          </form>
        </div>
      </div>

      <!-- Card de Tabela -->
      <div class="table-card">
        <div class="table-wrapper">
<%
if trim(request("parametro")) <> "" then
%>
          <table class="table table-bordered table-striped table-hover table-sm mb-0">
            <thead>
              <tr>
                <th colspan="2">Código</th>
                <th>Loja</th>
                <th>Nome</th>
                <th>CNPJ</th>
                <th>Bairro</th>
                <th>Município</th>
                <th>U.F</th>
              </tr>
            </thead>
            <tbody>
<%
  if not Consulta.Eof then

    while not Consulta.Eof
%>
              <tr>
<%
      if Trim(Consulta("A1_CODSAP")) <> "" and Trim(Consulta("A1_MSBLQL")) <> "1" then
        alert = ""
%>
                <td>
                  <%if request("id6")="06" then%>
                    <button type="button"
                            class="btn-selecionar"
                            title="Selecionar cliente"
                            onclick="valida('<%=EscapeJS(Trim(Consulta("A1_COD")))%>','<%=EscapeJS(Trim(Consulta("A1_LOJA")))%>','<%=EscapeJS(Trim(Consulta("A1_NOME")))%>','<%=EscapeJS(Trim(Consulta("A1_VEND")))%>','<%=EscapeJS(Trim(Consulta("CONDICAO_PAGAMENTO")))%>','<%=EscapeJS(Trim(Consulta("DIAS_TRANSPORTE")))%>','<%=EscapeJS(Trim(Consulta("TIPO_FRETE")))%>','<%=EscapeJS(Trim(Consulta("A1_MUN")))%>','<%=EscapeJS(Trim(Consulta("A1_EST")))%>', '<%=EscapeJS(Trim(Consulta("A1_VEND")))%>', '<%=EscapeJS(Trim(Consulta("A3_NOME")))%>')">
                      <i class="bi bi-download"></i>
                    </button>
                  <%else%>
                    <button type="button"
                            class="btn-selecionar"
                            title="Selecionar cliente"
                            onclick="valida('<%=EscapeJS(Trim(Consulta("A1_COD")))%>','<%=EscapeJS(Trim(Consulta("A1_LOJA")))%>','<%=EscapeJS(Trim(Consulta("A1_NOME")))%>','<%=EscapeJS(Trim(Consulta("A1_VEND")))%>','<%=EscapeJS(Trim(Consulta("CONDICAO_PAGAMENTO")))%>','<%=EscapeJS(Trim(Consulta("DIAS_TRANSPORTE")))%>','<%=EscapeJS(Trim(Consulta("TIPO_FRETE")))%>','<%=EscapeJS(Trim(Consulta("A1_MUN")))%>','<%=EscapeJS(Trim(Consulta("A1_EST")))%>', '<%=EscapeJS(Trim(Consulta("A1_VEND")))%>', '<%=EscapeJS(Trim(Consulta("A3_NOME")))%>')">
                      <i class="bi bi-download"></i>
                    </button>
                  <%end if%>
                </td>
<%
      end if

      if Trim(Consulta("A1_CODSAP")) = "" and Trim(Consulta("A1_MSBLQL")) <> "1" then
        alert = "class='Alert'"
%>
                <td>
                  <i class="bi bi-slash-circle text-danger"
                     title="Este cliente não possui código SAP cadastrado. Favor solicitar o cadastro com o comercial."></i>
                </td>
<%
      end if

      if Trim(Consulta("A1_MSBLQL")) = "1" then
        alert = "class='Alert'"
%>
                <td>
                  <i class="bi bi-slash-circle text-danger"
                     title="Este cliente está bloqueado. Favor solicitar o desbloqueio com o comercial."></i>
                </td>
<%
      end if
%>
                <td <%=alert%>><%=Consulta("A1_COD")%></td>
                <td <%=alert%>><%=Consulta("A1_LOJA")%></td>
                <td <%=alert%>><%=Consulta("A1_NOME")%></td>
                <td <%=alert%>><%=Consulta("Cnpj")%></td>
                <td <%=alert%>><%=Consulta("A1_BAIRRO")%></td>
                <td <%=alert%>><%=Consulta("A1_MUN")%></td>
                <td <%=alert%>><%=Consulta("A1_EST")%></td>
              </tr>
<%
      response.Flush()
      Consulta.MoveNext()

    wend

  else
%>
              <tr>
                <td colspan="8" class="text-center text-muted py-3">
                  <i class="bi bi-exclamation-circle me-1"></i>
                  Cliente não encontrado. Possíveis causas: cliente bloqueado ou não pertence à sua carteira.
                </td>
              </tr>
<%
  end if

  Consulta.Close()
  SET Consulta = NOTHING
%>
            </tbody>
          </table>
<%
else
%>
          <div class="table-placeholder">
            <i class="bi bi-clipboard-data"></i>
            <p>Digite um nome ou código e clique em <strong>Filtrar</strong> para buscar clientes.</p>
          </div>
<%
end if
%>
        </div>
      </div>

    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="funcoes/funcoesJAVASCRIPT.js"></script>

    <script>
      // TODO 01/11/2024 - Carlos Souza
      // Foi observado que o id6 recebe o conteúdo da Filial (não o id3 como estava anteriormente)
      // id1=codCliente | id2=codLoja | id3=txtNomeCliente
      // id5=hidCondicoesFaturamento | id6=07 | id7=1 | id8=hidDiasTransporte

      function valida(cliente, loja, nome, codrepre, condfaturamento, diastransporte, tipofrete, cidade, estado, codVendedor, nomeVendedor) {
        diastransporte = diastransporte || '';
        tipofrete      = tipofrete      || '';
        cidade         = cidade         || '';
        estado         = estado         || '';

        const vPlanta = document.getElementById('id6');
        let planta = '';

        if (vPlanta) {
          const vValue = vPlanta.value;
          console.log('value: ' + vValue);
          planta = vValue.trim() !== '' ? vValue : '';
        }

        <% if trim(request("id8")) <> "" then %>
          opener.document.getElementById('<%=trim(request("id8"))%>').value = diastransporte;
        <% end if %>

        <% if trim(request("id9")) <> "" then %>
          opener.document.getElementById('<%=trim(request("id9"))%>').value = `${cidade} - ${estado}`;
        <% end if %>

        <% if trim(request("id10")) <> "" then %>
          opener.document.getElementById('<%=trim(request("id10"))%>').value = `${codVendedor} - ${nomeVendedor}`;
        <% end if %>

        if (document.getElementById('id4').value != cliente.trim() + '' + loja.trim()) {

          opener.document.getElementById('<%=trim(request("id1"))%>').value = cliente;
          opener.document.getElementById('<%=trim(request("id2"))%>').value = loja;

          if (planta === '06') {
            const element = opener.document.getElementById('hidTipoFrete');
            if (opener.document.contains(element)) {
              opener.document.getElementById('hidTipoFrete').value = tipofrete;
              opener.tipoFrete();
            }
          }

          console.log('planta', planta);

          if (planta !== '00' && planta !== '05' && planta !== '06' && planta !== '04') {

            <% if trim(request("id3")) <> "" then %>
              opener.document.getElementById('<%=trim(request("id3"))%>').value = codrepre;
              console.log('codrepre', codrepre);
              opener.document.getElementById('<%=trim(request("id3"))%>').value = nome;
            <% end if %>

            <% if trim(request("id5")) <> "" then %>
              opener.document.getElementById('<%=trim(request("id5"))%>').value = condfaturamento;
            <% end if %>

          } else {

            console.log('nome', nome);

            opener.document.getElementById('<%=trim(request("id1"))%>').value = cliente;

            <% if trim(request("id3")) <> "" then %>
              opener.document.getElementById('<%=trim(request("id3"))%>').value = nome;
            <% end if %>

            <% if trim(request("id5")) <> "" then %>
              opener.document.getElementById('<%=trim(request("id5"))%>').value = condfaturamento;
            <% end if %>

          }

          opener.document.getElementById('<%=trim(request("id1"))%>').focus();
          window.close();

        } else {
          alert('Cliente e loja de Origem devem ser diferentes\ndo cliente e Loja de Destino.');
        }
      }

      function filtra() {
        frmBuscaCliente.submit();
      }
    </script>

  </body>
</html>
