<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../funcoes/funcoes.asp"-->
<!--#Include File="../../../includes/variaveis.asp"-->
<%
Response.Charset  = "UTF-8"
Response.CodePage = 65001
Session.LCID      = 1046

Dim pagina
pagina = Trim(Request("pagina"))


' ════════════════════════════════════════════════════════════════════════
'  ROTA: buscaCliente
' ════════════════════════════════════════════════════════════════════════
If pagina = "buscaCliente" Then

  Dim bc_STATUS, bc_codsiga, bc_codigoDePesquisa, bc_SQL, bc_RS

  bc_STATUS = Trim(Request("id9"))

  If Trim(Request("id4")) = "1" Then Nivel = 0

  bc_codsiga = IIf(CODUSER = "", " ", CODUSER)
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Busca Cliente</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <style>
    body{background-color:#f0f0f0;min-height:100vh}
    .filter-card{background:#fff;border:1px solid #dee2e6;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .filter-card-header{background:#fff;color:#212529;padding:.6rem 1rem;border-radius:6px 6px 0 0;border-bottom:1px solid #dee2e6;font-weight:600;font-size:.9rem}
    .filter-card-body{padding:1rem 1.25rem 1.25rem}
    .filter-card-body label{font-size:.82rem;font-weight:600;color:#6c757d;margin-bottom:.25rem}
    .table-card{background:#fff;border:1px solid #dee2e6;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .table-wrapper{overflow-x:auto;overflow-y:auto;max-height:68vh}
    .table thead th{position:sticky;top:0;background:#fff;color:#212529;font-size:.82rem;font-weight:700;text-align:center;white-space:nowrap;z-index:2;border-color:#dee2e6;padding:.55rem .75rem}
    .table tbody td{font-size:.82rem;text-align:center;white-space:nowrap;color:#212529;vertical-align:middle;padding:.45rem .75rem}
    .btn-selecionar{cursor:pointer;font-size:1rem}
    .btn-selecionar.ok{color:#198754}.btn-selecionar.ok:hover{color:#146c43}
    .btn-selecionar.ban{color:#dc3545}
    .table-placeholder{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:3.5rem 1rem;color:#6c757d;gap:.75rem}
    .table-placeholder .bi{font-size:2.8rem;opacity:.35}
    .table-placeholder p{margin:0;font-size:.9rem;text-align:center}
  </style>
</head>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="filter-card mb-3">
      <div class="filter-card-header"><i class="bi bi-funnel me-2"></i>Filtros</div>
      <div class="filter-card-body">
        <form id="frmBuscaCliente" name="frmBuscaCliente" method="post"
              action="ConsultaTransporteDeCarga.asp?pagina=buscaCliente">
          <input type="hidden" name="pagina" value="buscaCliente" />
          <input type="hidden" name="id1" id="id1" value="<%=Trim(Request("id1"))%>" />
          <input type="hidden" name="id2" id="id2" value="<%=Trim(Request("id2"))%>" />
          <input type="hidden" name="id3" id="id3" value="<%=Trim(Request("id3"))%>" />
          <input type="hidden" name="id4" id="id4" value="<%=Trim(Request("id4"))%>" />
          <input type="hidden" name="id5" id="id5" value="<%=Trim(Request("id5"))%>" />
          <input type="hidden" name="id6" id="id6" value="<%=Trim(Request("id6"))%>" />
          <input type="hidden" name="id7" id="id7" value="<%=Trim(Request("id7"))%>" />
          <input type="hidden" name="id8" id="id8" value="<%=Trim(Request("id8"))%>" />
          <input type="hidden" name="id9"       value="<%=Trim(Request("id9"))%>" />
          <div class="row g-2 align-items-end">
            <div class="col-auto flex-grow-1">
              <label for="parametro" class="form-label">Nome ou código</label>
              <input type="text" class="form-control form-control-sm" id="parametro" name="parametro"
                     value="<%=Trim(Request("parametro"))%>" autocomplete="off" />
            </div>
            <div class="col-auto d-flex gap-2">
              <button type="submit" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" class="btn btn-light btn-sm"
                      onclick="document.getElementById('parametro').value=''">
                <i class="bi bi-eraser"></i>
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>

    <div class="table-card">
<%
  If Trim(Request("parametro")) <> "" Then

    bc_codigoDePesquisa = IIf(InArray(Request("id6"), Array("BENTO", "MARANGUAPE")), "CODIGO_ERP", "COD_CLIENTE")

    bc_SQL = "" &_
      "SELECT " &_
        "PLANTA,CIDADE,COD_CLIENTE,LOJA_CLIENTE,NOME_CLIENTE,FANTASIA_CLIENTE," &_
        "CNPJ_CLIENTE,MUNICIPIO_CLIENTE,ESTADO_CLIENTE,COD_VENDEDOR,NOME_VENDEDOR," &_
        "CNPJ_VENDEDOR,COD_SAP_CLIENTE,CHAVE_CLIENTE,END_CLIENTE,CEP_CLIENTE,STATUS," &_
        "CLIENTE_GRUPO,GRUPO_DO_CLIENTE,LIMITE_CREDITO,INSCRICAO_ESTADUAL," &_
        "COD_MUNICIPIOS_IBGE,TES,EMAIL,BAIRRO,CONDICAO_PAGAMENTO,DIAS_TRANSPORTE," &_
        "CODIGO_SEGMENTO,DESCRICAO_SEGMENTO,CONTATO,TELEFONE,CODIGO_ERP " &_
      "FROM BI_CLIENTES " &_
      "WHERE ((FANTASIA_CLIENTE LIKE '%" & UCase(Trim(Request("parametro"))) & "%') " &_
          "OR (NOME_CLIENTE LIKE '%" & UCase(Trim(Request("parametro"))) & "%') " &_
          "OR (" & bc_codigoDePesquisa & " LIKE '%" & Trim(Request("parametro")) & "%')) " &_
        "AND CIDADE LIKE '" & Request("id6") & "' " &_
        IIf(bc_STATUS <> "", "AND STATUS = '" & bc_STATUS & "' ", "") &_
        IIf(Nivel <> 0, "AND COD_VENDEDOR = '" & bc_codsiga & "' ", "") &_
      "ORDER BY NOME_CLIENTE, LOJA_CLIENTE"

    Set bc_RS = Conecta("SPO-BI", bc_SQL)

    If Not IsObject(bc_RS) Then Response.End
%>
      <div class="table-wrapper">
        <table class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead>
            <tr>
              <th colspan="2">Código</th>
              <th>Loja</th>
              <th>Nome</th>
              <th>CNPJ</th>
              <th>Bairro</th>
              <th>Município</th>
              <th>U.F.</th>
            </tr>
          </thead>
          <tbody>
<%
    If Not bc_RS.EOF Then
      While Not bc_RS.EOF
%>
            <tr>
              <td>
<%
        If Request("id6") = "06" Then
%>
                <span onclick="valida('<%=IIf(InArray(bc_RS("CIDADE"),Array("BENTO","MARANGUAPE")),bc_RS("CODIGO_ERP"),bc_RS("COD_CLIENTE"))%>','<%=Trim(bc_RS("LOJA_CLIENTE"))%>','<%=Trim(bc_RS("NOME_CLIENTE"))%>','<%=Trim(bc_RS("COD_VENDEDOR"))%>','<%=Trim(bc_RS("CONDICAO_PAGAMENTO"))%>','<%=Trim(bc_RS("DIAS_TRANSPORTE"))%>')">
                  <i class="bi bi-check-circle-fill btn-selecionar ok" title="Selecione o cliente."></i>
                </span>
<%
        Else
%>
                <span onclick="valida('<%=IIf(InArray(bc_RS("CIDADE"),Array("BENTO","MARANGUAPE")),bc_RS("CODIGO_ERP"),bc_RS("COD_CLIENTE"))%>','<%=Trim(bc_RS("LOJA_CLIENTE"))%>','<%=Trim(bc_RS("NOME_CLIENTE"))%>','<%=Trim(bc_RS("COD_VENDEDOR"))%>','<%=Trim(bc_RS("CONDICAO_PAGAMENTO"))%>','')">
<%
          If Trim(bc_RS("COD_SAP_CLIENTE")) = "" And Trim(bc_RS("STATUS")) <> "INATIVO" Then
%>
                  <i class="bi bi-slash-circle-fill btn-selecionar ban" title="Este cliente não possui código SAP cadastrado."></i>
<%
          ElseIf Trim(bc_RS("STATUS")) = "INATIVO" Then
%>
                  <i class="bi bi-slash-circle-fill btn-selecionar ban" title="Cliente INATIVO."></i>
<%
          Else
%>
                  <i class="bi bi-check-circle-fill btn-selecionar ok" title="Selecione o cliente."></i>
<%
          End If
%>
                </span>
<%
        End If
%>
              </td>
              <td><%=IIf(InArray(bc_RS("CIDADE"),Array("BENTO","MARANGUAPE")),bc_RS("CODIGO_ERP"),bc_RS("COD_CLIENTE"))%></td>
              <td><%=bc_RS("LOJA_CLIENTE")%></td>
              <td><%=bc_RS("NOME_CLIENTE")%></td>
              <td><%=bc_RS("CNPJ_CLIENTE")%></td>
              <td><%=bc_RS("BAIRRO")%></td>
              <td><%=bc_RS("MUNICIPIO_CLIENTE")%></td>
              <td><%=bc_RS("ESTADO_CLIENTE")%></td>
            </tr>
<%
        bc_RS.MoveNext()
      Wend
    Else
%>
            <tr>
              <td colspan="8">
                <div class="table-placeholder">
                  <i class="bi bi-clipboard-x"></i>
                  <p>Cliente não encontrado. Possíveis causas: cliente bloqueado ou não pertence à sua carteira.</p>
                </div>
              </td>
            </tr>
<%
    End If
    bc_RS.Close()
    Set bc_RS = Nothing
%>
          </tbody>
        </table>
      </div>
<%
  Else
%>
      <div class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Informe o nome ou código e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>
<%
  End If
%>
    </div>

  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    "use strict";
    function valida(cliente, loja, nome, codrepre, condfaturamento, diastransporte, tipofrete) {
      diastransporte = diastransporte || "";
      tipofrete      = tipofrete      || "";
      var planta = document.getElementById("id6").value;

      <%If Trim(Request("id8")) <> "" Then%>
      if (window.opener) opener.document.getElementById("<%=Trim(Request("id8"))%>").value = diastransporte;
      <%End If%>

      if (document.getElementById("id4").value !== cliente.trim() + loja.trim()) {
        if (window.opener) {
          opener.document.getElementById("<%=Trim(Request("id1"))%>").value = cliente;
          opener.document.getElementById("<%=Trim(Request("id2"))%>").value = loja;
        }
        if (planta === "06") {
          if (window.opener) {
            var elFrete = opener.document.getElementById("hidTipoFrete");
            if (elFrete) { elFrete.value = tipofrete; if (opener.tipoFrete) opener.tipoFrete(); }
          }
        }
        if (planta !== "00" && planta !== "05" && planta !== "06" && planta !== "04") {
          <%If Trim(Request("id3")) <> "" Then%>
          if (window.opener) opener.document.getElementById("<%=Trim(Request("id3"))%>").value = nome;
          <%End If%>
          <%If Trim(Request("id5")) <> "" Then%>
          if (window.opener) opener.document.getElementById("<%=Trim(Request("id5"))%>").value = condfaturamento;
          <%End If%>
        } else {
          if (window.opener) {
            opener.document.getElementById("<%=Trim(Request("id1"))%>").value = cliente;
            opener.document.getElementById("<%=Trim(Request("id3"))%>").value = nome;
            <%If Trim(Request("id5")) <> "" Then%>
            opener.document.getElementById("<%=Trim(Request("id5"))%>").value = condfaturamento;
            <%End If%>
          }
        }
        if (window.opener) window.close();
        /* Se chamado via iframe, esta função será sobrescrita pelo pai */
      } else {
        alert("Cliente e loja de Origem, devem ser diferentes\ndo cliente e Loja de Destino.");
      }
    }
  </script>
</body>
</html>
<%
  Response.End
End If ' pagina = "buscaCliente"


' ════════════════════════════════════════════════════════════════════════
'  ROTA: buscaProduto
' ════════════════════════════════════════════════════════════════════════
If pagina = "buscaProduto" Then

  Dim bp_codcli, bp_lojcli, bp_codsiga, bp_SQL, bp_RS, bp_CONECTADO
  bp_codcli = Trim(Request("id3"))
  bp_lojcli = Trim(Request("id4"))
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Busca Produto</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    body{background-color:#f0f0f0;min-height:100vh}
    .filter-card{background:#fff;border:1px solid #dee2e6;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .filter-card-header{background:#fff;color:#212529;padding:.6rem 1rem;border-radius:6px 6px 0 0;border-bottom:1px solid #dee2e6;font-weight:600;font-size:.9rem}
    .filter-card-body{padding:1rem 1.25rem 1.25rem}
    .filter-card-body label{font-size:.82rem;font-weight:600;color:#6c757d;margin-bottom:.25rem}
    .table-card{background:#fff;border:1px solid #dee2e6;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .table-wrapper{overflow-x:auto;overflow-y:auto;max-height:68vh}
    .table thead th{position:sticky;top:0;background:#fff;color:#212529;font-size:.82rem;font-weight:700;text-align:center;white-space:nowrap;z-index:2;border-color:#dee2e6;padding:.55rem .75rem}
    .table tbody td{font-size:.82rem;text-align:center;white-space:nowrap;color:#212529;vertical-align:middle;padding:.45rem .75rem}
    .btn-selecionar{cursor:pointer;color:#198754;font-size:1rem}.btn-selecionar:hover{color:#146c43}
    .table-placeholder{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:3.5rem 1rem;color:#6c757d;gap:.75rem}
    .table-placeholder .bi{font-size:2.8rem;opacity:.35}
    .table-placeholder p{margin:0;font-size:.9rem;text-align:center}
  </style>
</head>
<%
  bp_CONECTADO = ConectaDB("PIR-DADOSADV")
  If bp_CONECTADO = False Then
%>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">
    <div class="alert alert-warning d-flex align-items-center gap-2" role="alert">
      <i class="bi bi-exclamation-triangle-fill"></i>
      <div>Erro ao conectar ao banco de dados. Verifique com a equipe de suporte.</div>
      <button type="button" class="btn-close ms-auto" onclick="window.close()"></button>
    </div>
  </div>
</body>
</html>
<%
    Response.End
  End If

  bp_codsiga = CODUSER
  If bp_codsiga = "" Then bp_codsiga = " "
%>
<body>
  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="filter-card mb-3">
      <div class="filter-card-header"><i class="bi bi-funnel me-2"></i>Filtros</div>
      <div class="filter-card-body">
        <form id="form1" name="form1" method="post"
              action="ConsultaTransporteDeCarga.asp?pagina=buscaProduto">
          <input type="hidden" name="pagina" value="buscaProduto" />
          <input type="hidden" name="id1" id="id1" value="<%=Trim(Request("id1"))%>" />
          <input type="hidden" name="id2" id="id2" value="<%=Trim(Request("id2"))%>" />
          <input type="hidden" name="id3" id="id3" value="<%=Trim(Request("id3"))%>" />
          <input type="hidden" name="id4" id="id4" value="<%=Trim(Request("id4"))%>" />
          <div class="row g-2 align-items-end">
            <div class="col-auto flex-grow-1">
              <label for="parametro" class="form-label">Descrição ou código</label>
              <input type="text" class="form-control form-control-sm" id="parametro" name="parametro"
                     value="<%=Request("parametro")%>" autocomplete="off" />
            </div>
            <div class="col-auto d-flex gap-2">
              <button type="submit" class="btn btn-primary btn-sm">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
              <button type="button" class="btn btn-light btn-sm"
                      onclick="document.getElementById('parametro').value=''">
                <i class="bi bi-eraser"></i>
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>

    <div class="table-card">
<%
  If Trim(Request("parametro")) <> "" Then

    bp_SQL = "" &_
      "SELECT DISTINCT CODIGO_CLIENTE, LOJA_CLIENTE, CODIGO, REFERENCIA " &_
      "FROM VW_QUANTIDADE_CAMINHAO " &_
      "WHERE CODIGO_CLIENTE = '" & bp_codcli & "' " &_
      "AND LOJA_CLIENTE     = '" & bp_lojcli & "' " &_
      "AND (CODIGO = '" & Trim(Request("parametro")) & "' OR UPPER(REFERENCIA) LIKE UPPER('%" & Trim(Request("parametro")) & "%')) "

    Set bp_RS = Conecta("PIR-DADOSADV", bp_SQL)

    If Not IsObject(bp_RS) Or (bp_RS Is Nothing) Then Response.End
%>
      <div class="table-wrapper">
        <table class="table table-bordered table-striped table-hover table-sm mb-0">
          <thead>
            <tr>
              <th></th>
              <th>Código</th>
              <th>Descrição</th>
            </tr>
          </thead>
          <tbody>
<%
    If Not bp_RS.EOF Then
      While Not bp_RS.EOF
%>
            <tr>
              <td>
                <i class="bi bi-check-circle-fill btn-selecionar"
                   title="Clique aqui para selecionar o Produto."
                   onclick='selecionou(`<%=Trim(bp_RS("CODIGO"))%>`,`<%=Replace(Replace(Trim(bp_RS("REFERENCIA")),".",""  ),",",".")%>`)'>
                </i>
              </td>
              <td><%=bp_RS("CODIGO")%></td>
              <td><%=Replace(Replace(bp_RS("REFERENCIA"),".",""  ),",",".")%></td>
            </tr>
<%
        bp_RS.MoveNext()
      Wend
    Else
%>
            <tr>
              <td colspan="3">
                <div class="table-placeholder">
                  <i class="bi bi-clipboard-x"></i>
                  <p>Produto não encontrado. Possíveis causas: produto bloqueado, não pertence à sua carteira ou não está cadastrado.</p>
                </div>
              </td>
            </tr>
<%
    End If
    bp_RS.Close()
    Set bp_RS = Nothing
%>
          </tbody>
        </table>
      </div>
<%
  Else
%>
      <div class="table-placeholder">
        <i class="bi bi-clipboard-data"></i>
        <p>Informe a descrição ou código e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>
<%
  End If
%>
    </div>

  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script>
    "use strict";
    function selecionou(aCod, aRef) {
      if (window.opener) {
        opener.document.getElementById("<%=Trim(Request("id1"))%>").value = aCod;
        opener.document.getElementById("<%=Trim(Request("id2"))%>").value = aRef;
        window.close();
      }
      /* Se chamado via iframe, esta função será sobrescrita pelo pai */
    }
  </script>
</body>
</html>
<%
  Response.End
End If ' pagina = "buscaProduto"


' ════════════════════════════════════════════════════════════════════════
'  PÁGINA PRINCIPAL
' ════════════════════════════════════════════════════════════════════════
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Consulta de Transporte de Carga</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    .filter-card { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .filter-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
    .filter-card-body { padding: 1rem 1.25rem 1.25rem; }
    .filter-label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    .table-card { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); position: relative; }
    .table-toolbar { background-color: #fff; border-radius: 6px 6px 0 0; }
    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 68vh; }
    .table-wrapper table { min-width: max-content; }
    .table thead th { position: sticky; top: 0; background-color: #fff; color: #212529; font-size: .82rem; font-weight: 700; text-align: center; white-space: nowrap; z-index: 2; border-color: #dee2e6; padding: .55rem .75rem; }
    .table tbody td { font-size: .82rem; text-align: center; white-space: nowrap; color: #212529; vertical-align: middle; padding: .45rem .75rem; }

    .table-footer { display: flex; align-items: center; justify-content: space-between; padding: .6rem 1rem; border-top: 1px solid #dee2e6; flex-wrap: wrap; gap: .5rem; }
    .info-registros { font-size: .8rem; color: #6c757d; }

    .table-placeholder { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 3.5rem 1rem; color: #6c757d; gap: .75rem; }
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p   { margin: 0; font-size: .9rem; text-align: center; }

    .overlay-loading { position: absolute; inset: 0; background: rgba(255,255,255,.80); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 10; gap: 1rem; border-radius: 6px; }
    .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
    .overlay-loading span { font-size: .9rem; font-weight: 600; color: #495057; }

    .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
    .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
    @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
    .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

    #modalBuscaCliente .modal-dialog,
    #modalBuscaProduto .modal-dialog { max-width: 860px; }
    #modalBuscaCliente .modal-body,
    #modalBuscaProduto .modal-body   { padding: 0; height: 65vh; }
    #iframeBuscaCliente,
    #iframeBuscaProduto { width: 100%; height: 100%; border: none; border-radius: 0 0 .375rem .375rem; }
    .iframe-loading { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; background: #fff; border-radius: 0 0 .375rem .375rem; z-index: 5; }
  </style>
</head>

<body>

  <div class="toast-container-custom" id="toastContainer"></div>

  <!-- Modal — Busca de Cliente -->
  <div class="modal fade" id="modalBuscaCliente" tabindex="-1" aria-labelledby="lblModalCliente" aria-hidden="true">
    <div class="modal-dialog modal-lg modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header py-2">
          <h6 class="modal-title fw-bold" id="lblModalCliente">Busca de Cliente</h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body position-relative">
          <div class="iframe-loading" id="loadingCliente">
            <div class="spinner-border text-secondary" role="status">
              <span class="visually-hidden">Carregando...</span>
            </div>
          </div>
          <iframe id="iframeBuscaCliente" src="about:blank"
                  onload="document.getElementById('loadingCliente').style.display='none'">
          </iframe>
        </div>
      </div>
    </div>
  </div>

  <!-- Modal — Busca de Produto -->
  <div class="modal fade" id="modalBuscaProduto" tabindex="-1" aria-labelledby="lblModalProduto" aria-hidden="true">
    <div class="modal-dialog modal-lg modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header py-2">
          <h6 class="modal-title fw-bold" id="lblModalProduto">Busca de Produto</h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body position-relative">
          <div class="iframe-loading" id="loadingProduto">
            <div class="spinner-border text-secondary" role="status">
              <span class="visually-hidden">Carregando...</span>
            </div>
          </div>
          <iframe id="iframeBuscaProduto" src="about:blank"
                  onload="document.getElementById('loadingProduto').style.display='none'">
          </iframe>
        </div>
      </div>
    </div>
  </div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="d-flex align-items-center justify-content-between mb-3">
      <span class="page-title">Consulta de Transporte de Carga</span>
    </div>

    <!-- Filtros -->
    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">
        <form name="frmFiltro" id="frmFiltro">
          <div class="d-flex flex-wrap align-items-end gap-3">

            <!-- Filial -->
            <div>
              <label for="cboPlanta" class="filter-label form-label">Filial</label>
              <select name="cboPlanta" id="cboPlanta" class="form-select form-select-sm"
                      style="width:140px;"
                      onchange="ConsultaTransporteDeCarga.habilitaCodLoja(this)">
                <option value="">Selecione...</option>
                <option value="PIRAPETINGA">Pirapetinga</option>
                <!--
                <option value="UBERABA">Uberaba</option>
                <option value="BENTO">Bento Gonçalves</option>
                <option value="MARANGUAPE">Maranguape</option>
                <option value="SAQUAREMA">Saquarema</option>
                -->
              </select>
            </div>

            <!-- Cliente -->
            <div>
              <label class="filter-label form-label">Cliente</label>
              <div class="d-flex gap-1 align-items-center">
                <input type="text" name="codcli" id="codcli"
                       class="form-control form-control-sm" style="width:72px;" />
                <input type="text" name="lojacli" id="lojacli"
                       class="form-control form-control-sm" style="width:46px;" />
                <button type="button" class="btn btn-secondary btn-sm"
                        onclick="ConsultaTransporteDeCarga.abrirModalCliente()">
                  Buscar <i class="bi bi-box-arrow-up-right ms-1"></i>
                </button>
                <input name="razao" id="razao" type="text"
                       class="form-control form-control-sm" style="width:310px;" readonly />
              </div>
            </div>

            <!-- Produto -->
            <div>
              <label class="filter-label form-label">Produto</label>
              <div class="d-flex gap-1 align-items-center">
                <input type="text" name="produto" id="produto"
                       class="form-control form-control-sm" style="width:90px;"
                       onblur="ConsultaTransporteDeCarga.validarProduto(this)" />
                <button type="button" class="btn btn-secondary btn-sm"
                        onclick="ConsultaTransporteDeCarga.abrirModalProduto()">
                  Buscar <i class="bi bi-box-arrow-up-right ms-1"></i>
                </button>
                <input name="referencia" id="referencia" type="text"
                       class="form-control form-control-sm" style="width:310px;" readonly />
              </div>
            </div>

            <!-- Ações -->
            <div class="ms-auto">
              <label class="filter-label form-label d-block">&nbsp;</label>
              <div class="d-flex gap-2">
                <button type="button" class="btn btn-light btn-sm"
                        onclick="ConsultaTransporteDeCarga.limparFiltro()">
                  <i class="bi bi-eraser me-2"></i>Limpar
                </button>
                <button type="button" class="btn btn-primary btn-sm"
                        onclick="ConsultaTransporteDeCarga.filtrar()">
                  <i class="bi bi-funnel-fill me-2"></i>Filtrar
                </button>
              </div>
            </div>

          </div>
        </form>
      </div>
    </div>

    <!-- Tabela -->
    <div class="table-card">

      <div id="overlayLoading" class="overlay-loading d-none">
        <div class="spinner-border" role="status"><span class="visually-hidden">Carregando...</span></div>
        <span>Carregando...</span>
      </div>

      <!-- Toolbar -->
      <div class="table-toolbar d-flex align-items-center justify-content-between px-3 py-2 border-bottom">
        <div class="d-flex gap-2">
          <button id="btnExportCsv" type="button" class="btn btn-success btn-sm" disabled
                  onclick="ConsultaTransporteDeCarga.exportarCSV()">
            <i class="bi bi-file-earmark-excel me-1"></i>Exportar Excel
          </button>
          <button id="btnExportPdf" type="button" class="btn btn-danger btn-sm" disabled
                  onclick="ConsultaTransporteDeCarga.exportarPDF()">
            <i class="bi bi-file-earmark-pdf me-1"></i>Exportar PDF
          </button>
        </div>
        <div>
          <input type="text" id="searchBox" class="form-control form-control-sm"
                 placeholder="Buscar..." autocomplete="off" style="width:200px;" />
        </div>
      </div>

      <!-- Área de resultado -->
      <div id="areaResultado">
        <div class="table-placeholder">
          <i class="bi bi-clipboard-data"></i>
          <p>Selecione uma filial e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
        </div>
      </div>

    </div><!-- /table-card -->

  </div><!-- /container-fluid -->

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.5.31/dist/jspdf.plugin.autotable.min.js"></script>

  <script>
    "use strict";

    const ConsultaTransporteDeCarga = (() => {

      const REPOSITORIO = "ConsultaTransporteDeCargaRepositorio.asp";
      const SELF        = "ConsultaTransporteDeCarga.asp";
      const LOGO_PATH   = "Logo/logo-smurfit1.png";
      const TOAST_DELAY = 4500;
      const NIVEL       = <%=Nivel%>;
      const CODUSER     = "<%=CODUSER%>";

      const TOAST_ICON = Object.freeze({
        sucesso : { icon: "bi-check-circle-fill",         cor: "#198754", barra: "#198754" },
        erro    : { icon: "bi-x-circle-fill",             cor: "#dc3545", barra: "#dc3545" },
        aviso   : { icon: "bi-exclamation-triangle-fill", cor: "#856404", barra: "#ffc107" },
        info    : { icon: "bi-info-circle-fill",          cor: "#6c757d", barra: "#6c757d" },
      });

      const CABECALHOS_TABELA = Object.freeze([
        "Tipo",
        "Caixas 2700", "Palete 2700", "M² Total 2700",
        "Caixas 2800", "Palete 2800", "M² Total 2800",
      ]);

      const $el = id => $(`#${id}`);

      let modalCliente = null;
      let modalProduto = null;

      /* ── Toast ──────────────────────────────────────────────────────── */
      function exibirToast(tipo, mensagem) {
        const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
        const el  = document.createElement("div");
        el.className = "toast-custom p-2";
        el.innerHTML = `
          <div class="d-flex align-items-start gap-2 px-1 pt-1">
            <i class="bi ${cfg.icon} fs-5" style="color:${cfg.cor};flex-shrink:0;margin-top:2px;"></i>
            <div class="flex-grow-1" style="font-size:.85rem;">${mensagem}</div>
            <button style="background:none;border:none;font-size:1rem;line-height:1;cursor:pointer;color:#6c757d;"
                    onclick="this.closest('.toast-custom').remove()">&times;</button>
          </div>
          <div style="height:3px;background:#e9ecef;margin-top:6px;border-radius:0 0 10px 10px;">
            <div class="toast-progress-bar" style="height:100%;background:${cfg.barra};border-radius:inherit;"></div>
          </div>`;
        $el("toastContainer")[0].appendChild(el);
        setTimeout(() => el.remove(), TOAST_DELAY);
      }

      /* ── AJAX nativo ────────────────────────────────────────────────── */
      function ajax(params, onSuccess, onError) {
        const qs = Object.entries(params)
          .map(([k, v]) => encodeURIComponent(k) + "=" + encodeURIComponent(v))
          .join("&");
        const xhr = new XMLHttpRequest();
        xhr.open("POST", REPOSITORIO, true);
        xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4) return;
          if (xhr.status === 200) {
            try { onSuccess(JSON.parse(xhr.responseText)); }
            catch (e) { onError("Resposta inválida do servidor."); }
          } else {
            onError("Erro HTTP " + xhr.status);
          }
        };
        xhr.send(qs);
      }

      /* ── Filtro principal ───────────────────────────────────────────── */
      function filtrar() {
        const planta  = $el("cboPlanta").val();
        const codcli  = $el("codcli").val().trim();
        const lojacli = $el("lojacli").val().trim();
        const produto = $el("produto").val().trim();

        if (planta  === "") { exibirToast("aviso", "Selecione a filial."); return; }
        if (codcli  === "" || lojacli === "") { exibirToast("aviso", "Informe o cliente e a loja."); return; }
        if (produto === "") { exibirToast("aviso", "Informe o produto."); return; }

        $el("overlayLoading").removeClass("d-none");
        $el("btnExportCsv").prop("disabled", true);
        $el("btnExportPdf").prop("disabled", true);

        ajax(
          { acao: "consultar", planta, codcli, lojacli, produto },
          function (res) {
            $el("overlayLoading").addClass("d-none");
            if (!res.sucesso) { exibirToast("erro", res.mensagem || "Erro ao consultar."); renderizarVazio(); return; }
            renderizarTabela(res.dados || []);
          },
          function (msg) {
            $el("overlayLoading").addClass("d-none");
            exibirToast("erro", msg);
            renderizarVazio();
          }
        );
      }

      /* ── Renderização ───────────────────────────────────────────────── */
      function renderizarVazio() {
        $el("areaResultado").html(`
          <div class="table-placeholder">
            <i class="bi bi-clipboard-x"></i>
            <p>Nenhum registro encontrado para o filtro selecionado.</p>
          </div>`);
        $el("btnExportCsv").prop("disabled", true);
        $el("btnExportPdf").prop("disabled", true);
      }

      function renderizarTabela(dados) {
        if (!dados || dados.length === 0) { renderizarVazio(); return; }

        let linhas = "";
        dados.forEach(r => {
          linhas += `<tr>
            <td class="text-start">${r.tipo}</td>
            <td class="text-end">${r.totalCaixas}</td>
            <td class="text-end">${r.totalPalete}</td>
            <td class="text-end">${r.m2Total}</td>
            <td class="text-end">${r.totalCaixas2800}</td>
            <td class="text-end">${r.totalPalete2800}</td>
            <td class="text-end">${r.m2Total2800}</td>
          </tr>`;
        });

        $el("areaResultado").html(`
          <div class="table-wrapper">
            <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tblResultado">
              <thead>
                <tr>
                  <th rowspan="2">Tipo</th>
                  <th colspan="3">Total Altura 2700</th>
                  <th colspan="3">Total Altura 2800</th>
                </tr>
                <tr>
                  <th>Caixas</th><th>Palete</th><th>M² Total</th>
                  <th>Caixas</th><th>Palete</th><th>M² Total</th>
                </tr>
              </thead>
              <tbody>${linhas}</tbody>
            </table>
          </div>
          <div class="table-footer">
            <div></div>
            <div class="info-registros">
              <i class="bi bi-list-ul me-1"></i>
              <span>${dados.length} registro(s)</span>
            </div>
          </div>`);

        $el("btnExportCsv").prop("disabled", false);
        $el("btnExportPdf").prop("disabled", false);
        initSearchBox();
      }

      /* ── Limpar ─────────────────────────────────────────────────────── */
      function limparFiltro() {
        $el("cboPlanta").val("");
        $el("codcli").val("").prop("disabled", false);
        $el("lojacli").val("").prop("disabled", false);
        $el("razao").val("");
        $el("produto").val("");
        $el("referencia").val("");
        $el("areaResultado").html(`
          <div class="table-placeholder">
            <i class="bi bi-clipboard-data"></i>
            <p>Selecione uma filial e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
          </div>`);
        $el("btnExportCsv").prop("disabled", true);
        $el("btnExportPdf").prop("disabled", true);
      }

      /* ── habilitaCodLoja ────────────────────────────────────────────── */
      function habilitaCodLoja(obj) {
        const val = obj.value;
        if (val === "BENTO" || val === "MARANGUAPE") {
          $el("lojacli").val("01").prop("disabled", true);
          $el("codcli").off("blur.tc").on("blur.tc", () => validarCliente($el("lojacli")[0], "ATIVO"));
          $el("lojacli").off("blur.tc");
        } else {
          $el("lojacli").val("").prop("disabled", false);
          $el("lojacli").off("blur.tc").on("blur.tc", () => validarCliente($el("lojacli")[0], "ATIVO"));
          $el("codcli").off("blur.tc");
        }
      }

      /* ── validarCliente (blur inline) ──────────────────────────────── */
      function validarCliente(obj, pStatus) {
        if (!obj.value || obj.value.trim() === "") { obj.value = ""; return; }
        ajax(
          { acao: "validaCliente", planta: $el("cboPlanta").val(), cliente: $el("codcli").val(),
            loja: obj.value, codsiga: CODUSER, nivel: NIVEL, tipo: "Grupo", status: pStatus },
          function (res) {
            if (!res.sucesso) {
              const erro = res.erro || "";
              if (erro === "ErroLogof") {
                exibirToast("erro", "Seu logon expirou. Conecte-se novamente ao SWnet.");
              } else {
                exibirToast("aviso", "Cliente não encontrado. Verifique bloqueio, carteira ou cadastro.");
              }
              $el("codcli").val(""); $el("lojacli").val(""); $el("razao").val("");
              return;
            }
            $el("razao").val(res.dados.nomeCliente);
          },
          function (msg) { exibirToast("erro", msg); }
        );
      }

      /* ── validarProduto (blur inline) ──────────────────────────────── */
      function validarProduto(obj) {
        if (!obj.value || obj.value.trim() === "") { obj.value = ""; return; }
        ajax(
          { acao: "validaProduto", cliente: $el("codcli").val(), loja: $el("lojacli").val(),
            produto: obj.value, nivel: NIVEL, tipo: "Produto" },
          function (res) {
            if (!res.sucesso) { $el("produto").val(""); $el("referencia").val(""); exibirToast("aviso", "Produto não encontrado."); return; }
            $el("referencia").val(res.dados.referencia);
          },
          function (msg) { exibirToast("erro", msg); }
        );
      }

      /* ── Modal busca cliente ────────────────────────────────────────── */
      function abrirModalCliente() {
        if ($el("cboPlanta").val() === "") { exibirToast("aviso", "Selecione a Filial."); return; }
        document.getElementById("loadingCliente").style.display = "flex";
        const lnk = `${SELF}?pagina=buscaCliente&id1=codcli&id2=lojacli&id3=razao&id6=${$el("cboPlanta").val()}&id9=ATIVO`;
        const ifr = document.getElementById("iframeBuscaCliente");
        ifr.src = lnk;
        ifr.onload = function () {
          document.getElementById("loadingCliente").style.display = "none";
          try {
            ifr.contentWindow.valida = function (cliente, loja, nome) {
              document.getElementById("codcli").value  = cliente;
              document.getElementById("lojacli").value = loja;
              document.getElementById("razao").value   = nome;
              modalCliente.hide();
            };
          } catch (err) { console.error("Erro ao interceptar iframe cliente:", err); }
        };
        modalCliente.show();
      }

      /* ── Modal busca produto ────────────────────────────────────────── */
      function abrirModalProduto() {
        if ($el("cboPlanta").val() === "") { exibirToast("aviso", "Selecione a Filial."); return; }
        document.getElementById("loadingProduto").style.display = "flex";
        const lnk = `${SELF}?pagina=buscaProduto&id1=produto&id2=referencia&id3=${$el("codcli").val()}&id4=${$el("lojacli").val()}`;
        const ifr = document.getElementById("iframeBuscaProduto");
        ifr.src = lnk;
        ifr.onload = function () {
          document.getElementById("loadingProduto").style.display = "none";
          try {
            ifr.contentWindow.selecionou = function (aCod, aRef) {
              document.getElementById("produto").value    = aCod;
              document.getElementById("referencia").value = aRef;
              modalProduto.hide();
            };
          } catch (err) { console.error("Erro ao interceptar iframe produto:", err); }
        };
        modalProduto.show();
      }

      /* ── Search box ─────────────────────────────────────────────────── */
      function initSearchBox() {
        $el("searchBox").off("input.tc").on("input.tc", function () {
          const term = $(this).val().toLowerCase();
          $("#tblResultado tbody tr").each(function () {
            $(this).toggleClass("d-none", term !== "" && !$(this).text().toLowerCase().includes(term));
          });
        });
      }

      /* ── Linhas visíveis ────────────────────────────────────────────── */
      function linhasVisiveis() {
        const linhas = [];
        $("#tblResultado tbody tr:not(.d-none)").each(function () {
          const tds = $(this).find("td");
          linhas.push([
            tds.eq(0).text().trim(), tds.eq(1).text().trim(), tds.eq(2).text().trim(),
            tds.eq(3).text().trim(), tds.eq(4).text().trim(), tds.eq(5).text().trim(),
            tds.eq(6).text().trim(),
          ]);
        });
        return linhas;
      }

      /* ── Exportar CSV ───────────────────────────────────────────────── */
      function exportarCSV() {
        const esc = v => {
          const s = v == null ? "" : String(v);
          return (s.includes(";") || s.includes('"') || s.includes("\n")) ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const rows = [
          CABECALHOS_TABELA.map(esc).join(";"),
          ...linhasVisiveis().map(row => row.map(esc).join(";")),
        ];
        const blob = new Blob(["\uFEFF" + rows.join("\r\n")], { type: "text/csv;charset=utf-8;" });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        a.href = url; a.download = "ConsultaTransporteDeCarga.csv";
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
      }

      /* ── Exportar PDF ───────────────────────────────────────────────── */
      function exportarPDF() {
        const { jsPDF } = window.jspdf;
        const doc   = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
        const pageW = doc.internal.pageSize.getWidth();
        const img   = new Image();
        img.src = LOGO_PATH;
        img.onload = () => {
          const canvas  = document.createElement("canvas");
          canvas.width  = 900;
          canvas.height = Math.round(900 * (img.height / img.width));
          canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
          doc.addImage(canvas.toDataURL("image/png"), "PNG", 14, 5, 38, 13);
          _gerarPDF(doc, pageW);
        };
        img.onerror = () => _gerarPDF(doc, pageW);
      }

      function _gerarPDF(doc, pageW) {
        doc.setFontSize(14).setFont(undefined, "bold");
        doc.text("Consulta de Transporte de Carga", pageW / 2, 24, { align: "center" });
        doc.setFontSize(8).setFont(undefined, "normal");
        doc.text(`Filial: ${$el("cboPlanta").val()}`, pageW / 2, 30, { align: "center" });
        doc.setFontSize(7);
        doc.text(`Emitido em: ${new Date().toLocaleString("pt-BR")}`, pageW - 10, 8, { align: "right" });
        doc.autoTable({
          head      : [CABECALHOS_TABELA],
          body      : linhasVisiveis(),
          startY    : 34,
          theme     : "striped",
          styles    : { fontSize: 4, cellPadding: 1.5, lineColor: [180, 180, 180], lineWidth: { top: 0, bottom: 0, left: 0.2, right: 0.2 } },
          headStyles: { fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: "bold", halign: "center", lineColor: [180, 180, 180], lineWidth: { top: 0, bottom: 0, left: 0.2, right: 0.2 } },
        });
        doc.save("ConsultaTransporteDeCarga.pdf");
      }

      /* ── Init ───────────────────────────────────────────────────────── */
      function init() {
        modalCliente = new bootstrap.Modal(document.getElementById("modalBuscaCliente"));
        modalProduto = new bootstrap.Modal(document.getElementById("modalBuscaProduto"));
        $(document).on("keydown", function (e) {
          if ((e.ctrlKey || e.metaKey) && e.key === "k") { e.preventDefault(); $el("searchBox").trigger("focus"); }
        });
      }

      return Object.freeze({
        init, filtrar, limparFiltro, habilitaCodLoja,
        validarCliente, validarProduto,
        abrirModalCliente, abrirModalProduto,
        exportarCSV, exportarPDF,
      });

    })();

    $(document).ready(() => ConsultaTransporteDeCarga.init());
  </script>

</body>
</html>
