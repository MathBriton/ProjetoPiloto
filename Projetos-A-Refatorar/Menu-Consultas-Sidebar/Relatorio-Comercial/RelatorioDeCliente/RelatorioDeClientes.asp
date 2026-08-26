<!--#include file="../../../cnn2/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp"-->
<%
Function dataEhoraDeHoje(tipo)
  Dim dia, mes, ano, hor, min, seg
  Dim data, hora, agora
  agora = Now()
  dia   = Right("00" & Day(agora),    2)
  mes   = Right("00" & Month(agora),  2)
  ano   = Year(agora)
  hor   = Right("00" & Hour(agora),   2)
  min   = Right("00" & Minute(agora), 2)
  seg   = Right("00" & Second(agora), 2)
  data  = dia & "/" & mes & "/" & ano
  hora  = hor & ":" & min & ":" & seg
  Select Case tipo
    Case "D"    dataEhoraDeHoje = data
    Case "H"    dataEhoraDeHoje = hora
    Case Else   dataEhoraDeHoje = data & " " & hora
  End Select
End Function

Function formataDataUStoBR(data)
  Dim dia, mes, ano, ret
  dia = "" : mes = "" : ano = "" : ret = ""
  If (Trim(data) <> "") Then
    If (Len(data) = 10) Then
      ano = Mid(data, 1, 4)
      mes = Mid(data, 6, 2)
      dia = Mid(data, 9, 2)
      ret = dia & "/" & mes & "/" & ano
    End If
  End If
  formataDataUStoBR = ret
End Function

Function formataDataSigatoBR(data)
  Dim dia, mes, ano, ret
  dia = "" : mes = "" : ano = "" : ret = ""
  If (Trim(data) <> "") Then
    If (Len(data) = 8) Then
      ano = Mid(data, 1, 4)
      mes = Mid(data, 5, 2)
      dia = Mid(data, 7, 2)
      ret = dia & "/" & mes & "/" & ano
    End If
  End If
  formataDataSigatoBR = ret
End Function

Session.LCID = 1046

vPedido = Request("pedido")

Dim vLimpando
vLimpando = (Trim(Request("limpar")) = "1")

Dim vPlanta, vStatus, vSegmento, vTipoFrete, vCodcli, vLojacli, vRazao, vRepresentante, vRegPorPage
If vLimpando Then
  vPlanta        = ""
  vStatus        = ""
  vSegmento      = ""
  vTipoFrete     = ""
  vCodcli        = ""
  vLojacli       = ""
  vRazao         = ""
  vRepresentante = ""
  vRegPorPage    = ""
Else
  vPlanta        = Trim(Request("cboPlanta"))
  vStatus        = Trim(Request("cboStatus"))
  vSegmento      = Trim(Request("cboSegmentoEconomico"))
  vTipoFrete     = Trim(Request("cboTipoFrete"))
  vCodcli        = Trim(Request("codcli"))
  vLojacli       = Trim(Request("lojacli"))
  vRazao         = Trim(Request("razao"))
  vRepresentante = Trim(Request("cboRepresentante"))
  vRegPorPage    = Trim(Request("cboRegPorPage"))
End If

' --- Modo "Todos": retorna JSON sem paginação para exportação client-side ---
If Request("todos") = "1" Then
  Server.ScriptTimeout = 300
  On Error Resume Next

  Dim flagAtu, codPesq, sqlTodos, rsTodos
  Dim jsonArr, sep, vCod, vCond

  SQL = "SELECT TOP 1 VALIDACAO FROM BI_ATUALIZACAO"
  Set rsTodos = Conecta("SPO-BI", SQL)
  If Err.Number <> 0 Or Not IsObject(rsTodos) Then
    Response.Clear
    Response.ContentType = "application/json"
    Response.CharSet     = "utf-8"
    Response.Write "{""erro"":""Falha ao conectar: " & Replace(Err.Description, """", "'") & """}"
    Response.End
  End If
  flagAtu = False
  If Not rsTodos.EOF Then flagAtu = rsTodos("VALIDACAO")
  rsTodos.Close()
  Set rsTodos = Nothing

  codPesq = iif(inArray(Request("cboPlanta"), Array("BENTO", "MARANGUAPE")), "C.CODIGO_ERP", "C.COD_CLIENTE")

  sqlTodos = "" &_
    "SELECT " &_
      "C.CIDADE, C.COD_CLIENTE, C.LOJA_CLIENTE, C.NOME_CLIENTE, C.FANTASIA_CLIENTE, " &_
      "C.CNPJ_CLIENTE, C.INSCRICAO_ESTADUAL, C.MUNICIPIO_CLIENTE, C.ESTADO_CLIENTE, " &_
      "C.END_CLIENTE, C.BAIRRO, C.TELEFONE, " &_
      "CASE WHEN C.TIPO_FRETE = 'C' THEN 'CIF' ELSE 'FOB' END TIPO_FRETE, " &_
      "C.CEP_CLIENTE, C.LIMITE_CREDITO, C.RISCO, C.TES, " &_
      "CP.CODIGO AS COD_COND_PAGTO, CP.DESCRICAO AS CONDICAO_PAGAMENTO, " &_
      "C.COD_VENDEDOR, C.NOME_VENDEDOR, C.DESCRICAO_SEGMENTO, C.CODIGO_ERP, " &_
      "C.DATA_CADASTRO, C.STATUS, C.COD_SAP_CLIENTE, C.DATA_MAIOR, C.DATA_MENOR, " &_
      "ISNULL(C.TOLMAIS,0) AS TOLMAIS, ISNULL(C.TOLMENOS,0) AS TOLMENOS " &_
    "FROM " & iif(flagAtu, "BI_CLIENTES_TEMP C ", "BI_CLIENTES C ") &_
      "LEFT JOIN BI_CONDICAO_PAGAMENTO CP ON C.CIDADE = CP.CIDADE AND C.CONDICAO_PAGAMENTO = CP.CODIGO " &_
    "WHERE 1=1 " &_
      iif(Nivel > 0,                                        "AND C.COD_VENDEDOR = '"        & CODUSER                         & "' ", "") &_
      iif(Trim(Request("cboPlanta"))            <> "",      "AND C.CIDADE = '"               & Request("cboPlanta")            & "' ", "") &_
      iif(Trim(Request("cboStatus"))            <> "",      "AND C.STATUS = '"               & Request("cboStatus")            & "' ", "") &_
      iif(Trim(Request("cboSegmentoEconomico")) <> "",      "AND C.CODIGO_SEGMENTO = '"      & Request("cboSegmentoEconomico") & "' ", "") &_
      iif(Trim(Request("codcli"))               <> "",      "AND " & codPesq & " = '"        & Request("codcli")               & "' ", "") &_
      iif(Trim(Request("lojacli"))              <> "",      "AND C.LOJA_CLIENTE = '"          & Request("lojacli")              & "' ", "") &_
      iif(Trim(Request("cboRepresentante"))     <> "",      "AND C.COD_VENDEDOR = '"          & Request("cboRepresentante")     & "' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "CIF",    "AND C.TIPO_FRETE = 'C' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "FOB",    "AND C.TIPO_FRETE <> 'C' ", "") &_
    "ORDER BY C.COD_CLIENTE"

  Set rsTodos = Conecta("SPO-BI", sqlTodos)
  If Err.Number <> 0 Or Not IsObject(rsTodos) Then
    Response.Clear
    Response.ContentType = "application/json"
    Response.CharSet     = "utf-8"
    Response.Write "{""erro"":""Falha na consulta: " & Replace(Err.Description, """", "'") & """}"
    Response.End
  End If

  Response.Clear
  Response.ContentType = "application/json"
  Response.CharSet     = "utf-8"

  Function fNull(v)
    If IsNull(v) Or IsEmpty(v) Then
      fNull = ""
    Else
      Dim s
      s = CStr(v)
      s = Replace(s, Chr(9),  "")
      s = Replace(s, Chr(10), "")
      s = Replace(s, Chr(13), "")
      s = Replace(s, "&amp;",  "&")
      s = Replace(s, "&lt;",   "<")
      s = Replace(s, "&gt;",   ">")
      s = Replace(s, "&quot;", """")
      fNull = s
    End If
  End Function

  Dim rowJson, flushCount
  flushCount = 0
  Response.Write "["
  sep = ""
  Do While Not rsTodos.EOF
    If fNull(rsTodos("CIDADE")) = "MARANGUAPE" Then
      vCod = fNull(rsTodos("CODIGO_ERP"))
    Else
      vCod = fNull(rsTodos("COD_CLIENTE"))
    End If
    If fNull(rsTodos("COD_COND_PAGTO")) <> "" Then
      vCond = "[" & Trim(fNull(rsTodos("COD_COND_PAGTO"))) & "] " & fNull(rsTodos("CONDICAO_PAGAMENTO"))
    Else
      vCond = ""
    End If
    rowJson = sep & "{"
    rowJson = rowJson & """STATUS"":"            & """" & fNull(rsTodos("STATUS"))                                         & ""","
    rowJson = rowJson & """CIDADE"":"            & """" & fNull(rsTodos("CIDADE"))                                         & ""","
    rowJson = rowJson & """COD_CLIENTE"":"       & """" & vCod                                                             & ""","
    rowJson = rowJson & """LOJA_CLIENTE"":"      & """" & fNull(rsTodos("LOJA_CLIENTE"))                                   & ""","
    rowJson = rowJson & """COD_SAP_CLIENTE"":"   & """" & fNull(rsTodos("COD_SAP_CLIENTE"))                                & ""","
    rowJson = rowJson & """NOME_CLIENTE"":"      & """" & Replace(fNull(rsTodos("NOME_CLIENTE")),     """", "\""")  & ""","
    rowJson = rowJson & """FANTASIA_CLIENTE"":"  & """" & Replace(fNull(rsTodos("FANTASIA_CLIENTE")), """", "\""")  & ""","
    rowJson = rowJson & """CNPJ_CLIENTE"":"      & """" & fNull(rsTodos("CNPJ_CLIENTE"))                                   & ""","
    rowJson = rowJson & """INSCRICAO_ESTADUAL"":"& """" & fNull(rsTodos("INSCRICAO_ESTADUAL"))                             & ""","
    rowJson = rowJson & """TELEFONE"":"          & """" & fNull(rsTodos("TELEFONE"))                                       & ""","
    rowJson = rowJson & """TIPO_FRETE"":"        & """" & fNull(rsTodos("TIPO_FRETE"))                                     & ""","
    rowJson = rowJson & """END_CLIENTE"":"       & """" & Replace(fNull(rsTodos("END_CLIENTE")),      """", "\""")  & ""","
    rowJson = rowJson & """BAIRRO"":"            & """" & Replace(fNull(rsTodos("BAIRRO")),            """", "\""")  & ""","
    rowJson = rowJson & """CEP_CLIENTE"":"       & """" & fNull(rsTodos("CEP_CLIENTE"))                                    & ""","
    rowJson = rowJson & """ESTADO_CLIENTE"":"    & """" & fNull(rsTodos("ESTADO_CLIENTE"))                                 & ""","
    rowJson = rowJson & """MUNICIPIO_CLIENTE"":"  & """" & Replace(fNull(rsTodos("MUNICIPIO_CLIENTE")),       """", "\""")  & ""","
    rowJson = rowJson & """LIMITE_CREDITO"":"    & """" & FormatNumber(iif(IsNull(rsTodos("LIMITE_CREDITO")), 0, rsTodos("LIMITE_CREDITO")), 2) & ""","
    rowJson = rowJson & """RISCO"":"             & """" & fNull(rsTodos("RISCO"))                                          & ""","
    rowJson = rowJson & """TES"":"               & """" & fNull(rsTodos("TES"))                                            & ""","
    rowJson = rowJson & """CONDICAO_PAGAMENTO"":"& """" & Replace(vCond, """", "\""")                                      & ""","
    rowJson = rowJson & """COD_VENDEDOR"":"      & """" & fNull(rsTodos("COD_VENDEDOR"))                                   & ""","
    rowJson = rowJson & """NOME_VENDEDOR"":"     & """" & Replace(fNull(rsTodos("NOME_VENDEDOR")),    """", "\""")  & ""","
    rowJson = rowJson & """DESCRICAO_SEGMENTO"":"& """" & Replace(fNull(rsTodos("DESCRICAO_SEGMENTO")), """", "\""") & ""","
    rowJson = rowJson & """DATA_CADASTRO"":"     & """" & formataDataSigatoBR(fNull(rsTodos("DATA_CADASTRO")))             & ""","
    rowJson = rowJson & """TOLMAIS"":"           & """" & fNull(rsTodos("TOLMAIS"))                                        & ""","
    rowJson = rowJson & """TOLMENOS"":"          & """" & fNull(rsTodos("TOLMENOS"))                                       & ""","
    rowJson = rowJson & """DATA_MAIOR"":"        & """" & fNull(rsTodos("DATA_MAIOR"))                                     & ""","
    rowJson = rowJson & """DATA_MENOR"":"        & """" & fNull(rsTodos("DATA_MENOR"))                                     & """"
    rowJson = rowJson & "}"
    Response.Write rowJson
    sep = ","
    flushCount = flushCount + 1
    If flushCount Mod 200 = 0 Then Response.Flush()
    rsTodos.MoveNext()
  Loop
  Response.Write "]"

  rsTodos.Close()
  Set rsTodos = Nothing

  Response.End
End If

' --- Modo normal: paginação para exibição em tela ---
If Request("filtra") = "1" Then
  Dim conn, rs, sql, pageSize, pageNumber, totalRecords, totalPages

  pageSize   = Request("cboRegPorPage")
  pageNumber = Request("page")

  If pageSize   = "" Or LCase(Trim(pageSize)) = "todos" Then pageSize   = 10 : End If
  If pageNumber = "" Then pageNumber = 1  : End If

  SQL = "SELECT TOP 1 VALIDACAO FROM BI_ATUALIZACAO"
  Set SPO_VALIDACAO = Conecta("SPO-BI", SQL)
  If Not IsObject(SPO_VALIDACAO) Then Response.End

  flagAtualizando = False
  If Not SPO_VALIDACAO.EOF Then flagAtualizando = SPO_VALIDACAO("VALIDACAO")
  SPO_VALIDACAO.Close()
  Set SPO_VALIDACAO = Nothing

  codigoDePesquisa = iif(inArray(Request("cboPlanta"), Array("BENTO", "MARANGUAPE")), "CODIGO_ERP", "COD_CLIENTE")

  SQL = "" &_
    "SELECT COUNT(*) AS TOTAL FROM " & iif(flagAtualizando, "BI_CLIENTES_TEMP ", "BI_CLIENTES ") &_
    "WHERE 1=1 " &_
      iif(Nivel > 0,                                        "AND COD_VENDEDOR = '"    & CODUSER                         & "' ", "") &_
      iif(Trim(Request("cboPlanta"))            <> "",      "AND CIDADE = '"          & Request("cboPlanta")            & "' ", "") &_
      iif(Trim(Request("cboStatus"))            <> "",      "AND STATUS = '"          & Request("cboStatus")            & "' ", "") &_
      iif(Trim(Request("cboSegmentoEconomico")) <> "",      "AND CODIGO_SEGMENTO = '" & Request("cboSegmentoEconomico") & "' ", "") &_
      iif(Trim(Request("codcli"))               <> "",      "AND " & codigoDePesquisa & " = '" & Request("codcli")     & "' ", "") &_
      iif(Trim(Request("lojacli"))              <> "",      "AND LOJA_CLIENTE = '"    & Request("lojacli")              & "' ", "") &_
      iif(Trim(Request("cboRepresentante"))     <> "",      "AND COD_VENDEDOR = '"    & Request("cboRepresentante")     & "' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "CIF",    "AND TIPO_FRETE = 'C' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "FOB",    "AND TIPO_FRETE <> 'C' ", "")

  Set SPO_TOTALREGS = Conecta("SPO-BI", SQL)
  If Not IsObject(SPO_TOTALREGS) Then Response.End

  totalRecords = SPO_TOTALREGS("TOTAL")
  totalPages   = Int((totalRecords + pageSize - 1) / pageSize)

  SPO_TOTALREGS.Close()
  Set SPO_TOTALREGS = Nothing
End If
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <title>SWnet - Relatório de Clientes</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />
  <script src="../../../funcoes/funcoesJAVASCRIPT.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.5.31/dist/jspdf.plugin.autotable.min.js"></script>
  <style>
    body { background: #f0f0f0; color: #212529; min-height: 100vh; }

    .page-title { color: #000000; }

    .form-label { font-size: 0.82rem; font-weight: 600; margin-bottom: 0.25rem; }

    .form-control:focus,
    .form-select:focus {
      border-color: #4a4a4a;
      box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
    }

    #cboPlanta    { min-width: 160px; }
    #cboStatus    { min-width: 110px; }
    #cboTipoFrete { min-width: 100px; }

    .overlay-loader {
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, 0.65);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .spinner-circular {
      width: 3rem;
      height: 3rem;
      border: 4px solid #e8e8e8;
      border-top-color: #4a4a4a;
      border-radius: 50%;
      animation: girar 0.8s linear infinite;
    }

    @keyframes girar {
      to { transform: rotate(360deg); }
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

    .toast-item-success { border-left: 4px solid #198754; }
    .toast-item-warning { border-left: 4px solid #ffc107; }
    .toast-item-danger  { border-left: 4px solid #dc3545; }
    .toast-item-info    { border-left: 4px solid #6c757d; }

    @keyframes entrar-toast {
      from { opacity: 0; transform: translateX(20px); }
      to   { opacity: 1; transform: translateX(0); }
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

    .toast-fechar:hover { color: #333; }

    .toast-barra { height: 3px; animation: encolher linear forwards; }
    .toast-barra-success { background: #198754; }
    .toast-barra-warning { background: #ffc107; }
    .toast-barra-danger  { background: #dc3545; }
    .toast-barra-info    { background: #6c757d; }

    @keyframes encolher {
      from { width: 100%; }
      to   { width: 0%; }
    }

    .filtro-container { box-shadow: 0 1px 4px rgba(0,0,0,.06); }

    .tbl-container {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 1px 4px rgba(0,0,0,.06);
      overflow: hidden;
    }

    .tbl-wrapper {
      overflow: auto;
      max-height: 72vh;
      border-bottom: 2px solid #dee2e6;
    }

    .tbl-wrapper thead th {
      position: sticky;
      top: 0;
      z-index: 3;
      background: #f8f9fa;
      color: #212529;
      font-size: 0.72rem;
      font-weight: 600;
      text-align: center;
      white-space: nowrap;
      vertical-align: middle;
      padding: 0.5rem 0.4rem;
      border-color: #dee2e6;
      text-transform: uppercase;
      letter-spacing: 0.2px;
    }

    .tbl-wrapper thead tr:first-child th { top: 0; }

    .tbl-wrapper tbody td {
      font-size: 0.78rem;
      vertical-align: middle;
      padding: 0.35rem 0.4rem;
      border-color: #e8e8e8;
    }

    .tbl-wrapper tbody tr:hover td {
      background-color: rgba(0, 0, 0, 0.03);
    }

    .tbl-vazia {
      padding: 5rem 1rem;
      text-align: center;
      color: #888;
      font-size: 0.88rem;
      background: #fff !important;
    }

    .tbl-aguardando {
      color: #adb5bd !important;
    }

    .tbl-empty-icon {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 3.5rem;
      height: 3.5rem;
      border-radius: 50%;
      background: #f1f3f5;
      margin: 0 auto 0.75rem;
      font-size: 1.5rem;
      color: #adb5bd;
    }

    .tbl-empty-title {
      font-size: 0.9rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: 0.25rem;
    }

    .tbl-empty-sub {
      font-size: 0.8rem;
      color: #adb5bd;
    }

    .rodape-tabela {
      position: relative;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 0.5rem 0.75rem;
      background: #fff;
      border-top: 1px solid #dee2e6;
      font-size: 0.8rem;
      color: #555;
    }

    .rodape-tabela #contador        { position: absolute; right: 0.75rem; }
    .rodape-tabela #seletorRegistros {
      position: absolute;
      left: 0.75rem;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      z-index: 2;
    }

    .seletor-label  { font-size: 0.78rem; }
    .seletor-select { width: auto; }

    .paginacao .page-link                     { font-size: 0.8rem; padding: 0.25rem 0.55rem; color: #4a4a4a; border-color: #dee2e6; }
    .paginacao .page-item.active .page-link   { background: #4a4a4a; border-color: #4a4a4a; color: #fff; }
    .paginacao .page-item.disabled .page-link { color: #aaa; pointer-events: none; }

    .loja-cli-input { max-width: 70px; }

    .btn-linha {
      font-size: 0.7rem;
      padding: 2px 7px;
      display: inline-flex;
      align-items: center;
      gap: 0.25rem;
      white-space: nowrap;
      font-weight: 500;
      letter-spacing: 0.1px;
    }

    .badge-cli-inativo {
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
      font-size: 0.7rem;
      font-weight: 600;
      padding: 0.2rem 0.5rem;
      border-radius: 20px;
      background: #fff0f0;
      color: #c0392b;
      border: 1px solid #f5c6cb;
      white-space: nowrap;
    }

    .badge-cli-ativo {
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
      font-size: 0.7rem;
      font-weight: 600;
      padding: 0.2rem 0.5rem;
      border-radius: 20px;
      background: #f0fff4;
      color: #1a7a3c;
      border: 1px solid #b7e4c7;
      white-space: nowrap;
    }

    .tbl-wrapper td.col-fix-0,
    .tbl-wrapper td.col-fix-1,
    .tbl-wrapper td.col-fix-2 { white-space: nowrap; }
    .tbl-wrapper td.col-fix-0 { text-align: center; }
    .tbl-wrapper thead th.col-fix-0,
    .tbl-wrapper thead th.col-fix-1,
    .tbl-wrapper thead th.col-fix-2 { }

    .tbl-wrapper th.col-razao-social,
    .tbl-wrapper td.col-razao-social { min-width: 260px; }

    .respiro-rodape { height: 2rem; }

    .dropdown-item i { min-width: 1.1rem; }
  </style>
</head>
<body>

  <div class="toast-wrapper" id="toastWrapper"></div>
  <div id="loader" class="overlay-loader d-none"><div class="spinner-circular"></div></div>

  <div class="container-fluid px-3 pt-2">

    <div class="d-flex justify-content-between align-items-center mb-2">
      <h5 class="mb-0 fw-bold page-title">Relatório de Clientes</h5>
      <a href="../../../manuais/manual_relatorio_de_clientes.pdf" target="_blank" title="Manual" class="btn btn-info btn-sm text-dark">
        <i class="bi bi-question-circle"></i>
      </a>
    </div>

    <div class="bg-white border rounded p-3 mb-2 filtro-container">
<%
If Trim(Request("oculta")) <> "1" Then
%>
      <form name="frmFiltroDeCliente" id="frmFiltroDeCliente" method="post" action="RelatorioDeClientes.asp">
        <input type="hidden" name="oculta" id="oculta" value="0" />
        <input type="hidden" name="filtra" id="filtra" value="0" />
        <input type="hidden" name="pdf"    id="pdf"    value="0" />
        <input type="hidden" name="page"   id="page"   value="0" />
        <input type="hidden" name="limpar" id="limpar" value="0" />
        <input type="hidden" name="selRep" id="selRep" value="<%=iif(vLimpando, "", trim(request("selRep")))%>" />
        <input type="hidden" name="cboRegPorPage" id="hRegPorPage" value="<%=iif(vRegPorPage = "", "10", vRegPorPage)%>" />

        <div class="row g-2 align-items-end mb-1">

          <div class="col-auto">
            <label for="cboPlanta" class="form-label">Filial <span class="text-danger">*</span></label>
            <select name="cboPlanta" id="cboPlanta" class="form-select form-select-sm" onchange="limparFiltro(this); atualizarRepresentantes(''); habilitaCodLoja(this);">
              <option value=""            <% if vPlanta = ""            then %> selected <% end if %>>Selecione...</option>
              <option value="PIRAPETINGA" <% if vPlanta = "PIRAPETINGA" then %> selected <% end if %>>Pirapetinga</option>
              <option value="UBERABA"     <% if vPlanta = "UBERABA"     then %> selected <% end if %>>Uberaba</option>
              <option value="BENTO"       <% if vPlanta = "BENTO"       then %> selected <% end if %>>Bento Gonçalves</option>
              <option value="MARANGUAPE"  <% if vPlanta = "MARANGUAPE"  then %> selected <% end if %>>Maranguape</option>
              <option value="SAQUAREMA"   <% if vPlanta = "SAQUAREMA"   then %> selected <% end if %>>Saquarema</option>
              <option value="ESHOP"       <% if vPlanta = "ESHOP"       then %> selected <% end if %>>ESHOP</option>
            </select>
          </div>

          <div class="col-auto">
            <label for="cboStatus" class="form-label">Status</label>
            <select name="cboStatus" id="cboStatus" class="form-select form-select-sm">
              <option value=""        <% if vStatus = ""        then %> selected <% end if %>>Todos</option>
              <option value="ATIVO"   <% if vStatus = "ATIVO"   then %> selected <% end if %>>Ativo</option>
              <option value="INATIVO" <% if vStatus = "INATIVO" then %> selected <% end if %>>Inativo</option>
            </select>
          </div>

          <div class="col">
            <label for="cboSegmentoEconomico" class="form-label">Segmento Econômico</label>
            <select name="cboSegmentoEconomico" id="cboSegmentoEconomico" class="form-select form-select-sm">
              <option value="" <% if vSegmento = "" then %> selected <% end if %>>Selecione...</option>
<%
              SQLSegEco = "" &_
                "SELECT CODIGO_SEGMENTO, DESCRICAO_SEGMENTO " &_
                "FROM BI_CLIENTES " &_
                "WHERE CODIGO_SEGMENTO <> '' AND DESCRICAO_SEGMENTO <> '' " &_
                "GROUP BY CODIGO_SEGMENTO, DESCRICAO_SEGMENTO " &_
                "ORDER BY DESCRICAO_SEGMENTO"

              Set TABSEGECO = Conecta("SPO-BI", SQLSegEco)
              If Not IsObject(TABSEGECO) Then Response.End

              Do While Not TABSEGECO.EOF
%>
              <option value="<%=Trim(TABSEGECO("CODIGO_SEGMENTO"))%>" <% If vSegmento = Trim(TABSEGECO("CODIGO_SEGMENTO")) Then %> selected <% End If %>><%=TABSEGECO("DESCRICAO_SEGMENTO") & " [" & TABSEGECO("CODIGO_SEGMENTO") & "]" %></option>
<%
                TABSEGECO.MoveNext()
              Loop
              TABSEGECO.Close()
              Set TABSEGECO = Nothing
%>
            </select>
          </div>

          <%
          Dim arrValores, arrTextos, i
          %>

          <div class="col-auto">
            <label for="cboTipoFrete" class="form-label">Tipo Frete</label>
            <select name="cboTipoFrete" id="cboTipoFrete" class="form-select form-select-sm">
              <%
              arrValores = Array("", "CIF", "FOB")
              arrTextos  = Array("Todos", "CIF", "FOB")
              For i = 0 To UBound(arrValores)
                Response.Write "<option value=""" & arrValores(i) & """"
                If vTipoFrete = arrValores(i) Then Response.Write " selected" End If
                Response.Write ">" & arrTextos(i) & "</option>" & vbCrLf
              Next
              %>
            </select>
          </div>

          <div class="col">
            <label for="filtroCodClienteDisplay" class="form-label">Cód. Cliente</label>
            <div class="input-group input-group-sm">
              <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarCliente()">
              Buscar <i class="bi bi-box-arrow-up-right me-1" title="Buscar Cliente"></i>
            </button>
              <input type="text" class="form-control form-control-sm" id="filtroCodClienteDisplay" placeholder="Buscar Cliente" readonly>
              <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimparCliente" title="Limpar cliente"><i class="bi bi-x-lg"></i></button>
            </div>
          </div>

          <input type="hidden" name="codcli"   id="codcli"   value="<%=vCodcli%>" />
          <input type="hidden" name="lojacli" id="lojacli" value="<%=vLojacli%>" />
          <input type="hidden" name="razao"   id="razao"   value="<%=vRazao%>" />

          <div class="col">
            <label for="cboRepresentante" class="form-label">Representante</label>
            <select name="cboRepresentante" id="cboRepresentante" class="form-select form-select-sm" onchange="obterRepresentanteSelecionado(this)">
              <option value="" <% if vRepresentante = "" then %> selected <% end if %>>Selecione a planta primeira...</option>
            </select>
          </div>

          <div class="col-auto d-flex gap-1 align-items-end">
            <button type="button" name="consultar" class="btn btn-primary btn-sm" onclick="Filtro()" title="Filtrar">
              <i class="bi bi-funnel me-1"></i>Filtrar
            </button>
            <button type="button" class="btn btn-light btn-sm border" title="Limpar o filtro" onclick="limparFiltro()">
              <i class="bi bi-eraser"></i>
            </button>
          </div>

          <div class="col-auto d-flex gap-1 align-items-end">
            <button type="button" class="btn btn-success btn-sm" title="Exportar Excel" onclick="exportarExcel()">
              <i class="bi bi-file-earmark-excel"></i>
            </button>
            <button type="button" class="btn btn-danger btn-sm" title="Exportar PDF" onclick="exportarPDF()">
              <i class="bi bi-file-earmark-pdf"></i>
            </button>
          </div>

        </div>

      </form>
<%
End If
%>
    </div>

    <div class="tbl-container mb-1">
      <div class="tbl-wrapper">
        <table id="GradeClientes" class="table table-bordered table-sm table-hover table-striped mb-0">
        <thead>
          <tr>
            <th rowspan="2" class="col-fix-0">Ação</th>
            <th rowspan="2" class="col-fix-1">Status</th>
            <th rowspan="2" class="col-fix-2">Planta</th>
            <th colspan="9">Cliente</th>
            <th colspan="5">Endereço</th>
            <th rowspan="2">Limite de<br>Crédito</th>
            <th rowspan="2">Risco<br>Financeiro</th>
            <th rowspan="2">TES</th>
            <th rowspan="2">Cond. de<br>Pagamento</th>
            <th colspan="2">Representante</th>
            <th rowspan="2">Segmento<br>Econômico</th>
            <th rowspan="2">Data<br>Cadastro</th>
            <th rowspan="2">Tolerância<br>Mais</th>
            <th rowspan="2">Tolerância<br>Menos</th>
            <th rowspan="2">Tolerância<br>Data Maior</th>
            <th rowspan="2">Tolerância<br>Data Menor</th>
          </tr>
          <tr>
            <th>Código</th>
            <th>Loja</th>
            <th>Código SAP</th>
            <th class="col-razao-social">Razão Social</th>
            <th>Nome Fantasia</th>
            <th>CNPJ</th>
            <th>I.E</th>
            <th>Telefone</th>
            <th>Frete</th>
            <th>Endereço</th>
            <th>Bairro</th>
            <th>CEP</th>
            <th>U.F</th>
            <th>Município</th>
            <th>Código</th>
            <th>Nome</th>
          </tr>
        </thead>
        <tbody>
<%
If Request("filtra") <> "1" Then
%>
          <tr>
            <td colspan="29" class="tbl-vazia tbl-aguardando">
              <div class="tbl-empty-icon"><i class="bi bi-funnel"></i></div>
              <div class="tbl-empty-title">Nenhum resultado exibido</div>
              <div class="tbl-empty-sub">Aplique os filtros acima para obter os resultados.</div>
            </td>
          </tr>
<%
End If
If Request("filtra") = "1" Then

  If pageSize   = "" Or pageSize   = 0 Then pageSize   = Request("cboRegPorPage") : End If
  If pageSize   = "" Or pageSize   = 0 Then pageSize   = 10                       : End If
  If LCase(Trim(pageSize)) = "todos"   Then pageSize   = 10                       : End If
  If pageNumber = "" Or pageNumber = 0 Then pageNumber = Request("page")          : End If
  If pageNumber = "" Or pageNumber = 0 Then pageNumber = 1                        : End If

  SQL = "SELECT TOP 1 VALIDACAO FROM BI_ATUALIZACAO"
  Set SPO_VALIDACAO = Conecta("SPO-BI", SQL)
  If Not IsObject(SPO_VALIDACAO) Then Response.End

  flagAtualizando = False
  If Not SPO_VALIDACAO.EOF Then flagAtualizando = SPO_VALIDACAO("VALIDACAO")
  SPO_VALIDACAO.Close()
  Set SPO_VALIDACAO = Nothing

  codigoDePesquisa = iif(inArray(Request("cboPlanta"), Array("BENTO", "MARANGUAPE")), "C.CODIGO_ERP", "C.COD_CLIENTE")

  SQL = "" &_
    "SELECT " &_
      "C.CIDADE, C.COD_CLIENTE, C.LOJA_CLIENTE, C.NOME_CLIENTE, C.FANTASIA_CLIENTE, " &_
      "C.CNPJ_CLIENTE, C.INSCRICAO_ESTADUAL, C.MUNICIPIO_CLIENTE, C.ESTADO_CLIENTE, " &_
      "C.END_CLIENTE, C.BAIRRO, C.TELEFONE, " &_
      "CASE WHEN C.TIPO_FRETE = 'C' THEN 'CIF' ELSE 'FOB' END TIPO_FRETE, " &_
      "C.CEP_CLIENTE, C.LIMITE_CREDITO, C.RISCO, C.TES, " &_
      "CP.CODIGO AS COD_COND_PAGTO, CP.DESCRICAO AS CONDICAO_PAGAMENTO, " &_
      "C.COD_VENDEDOR, C.NOME_VENDEDOR, C.DESCRICAO_SEGMENTO, C.CODIGO_ERP, " &_
      "C.DATA_CADASTRO, C.STATUS, C.COD_SAP_CLIENTE, C.DATA_MAIOR, C.DATA_MENOR, " &_
      "ISNULL(C.TOLMAIS,0) AS TOLMAIS, ISNULL(C.TOLMENOS,0) AS TOLMENOS " &_
    "FROM " & iif(flagAtualizando, "BI_CLIENTES_TEMP C ", "BI_CLIENTES C ") &_
      "LEFT JOIN BI_CONDICAO_PAGAMENTO CP ON C.CIDADE = CP.CIDADE AND C.CONDICAO_PAGAMENTO = CP.CODIGO " &_
    "WHERE 1=1 " &_
      iif(Nivel > 0,                                        "AND C.COD_VENDEDOR = '"        & CODUSER                         & "' ", "") &_
      iif(Trim(Request("cboPlanta"))            <> "",      "AND C.CIDADE = '"               & Request("cboPlanta")            & "' ", "") &_
      iif(Trim(Request("cboStatus"))            <> "",      "AND C.STATUS = '"               & Request("cboStatus")            & "' ", "") &_
      iif(Trim(Request("cboSegmentoEconomico")) <> "",      "AND C.CODIGO_SEGMENTO = '"      & Request("cboSegmentoEconomico") & "' ", "") &_
      iif(Trim(Request("codcli"))               <> "",      "AND " & codigoDePesquisa & " = '" & Request("codcli")            & "' ", "") &_
      iif(Trim(Request("lojacli"))              <> "",      "AND C.LOJA_CLIENTE = '"          & Request("lojacli")             & "' ", "") &_
      iif(Trim(Request("cboRepresentante"))     <> "",      "AND C.COD_VENDEDOR = '"          & Request("cboRepresentante")    & "' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "CIF",    "AND C.TIPO_FRETE = 'C' ", "") &_
      iif(Trim(Request("cboTipoFrete"))         = "FOB",    "AND C.TIPO_FRETE <> 'C' ", "") &_
    "ORDER BY COD_CLIENTE " &_
    "OFFSET " & (pageSize * (pageNumber - 1)) & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"

  Set CONSULTA = Conecta("SPO-BI", SQL)
  If Not IsObject(CONSULTA) Then Response.End

  If Not CONSULTA.EOF Then
    Do While Not CONSULTA.EOF
%>
          <tr>
            <td class="text-center col-fix-0">
<%
              Dim vCidadeAcao, vCodCliAcao
              vCidadeAcao = CONSULTA("CIDADE")
              If vCidadeAcao = "MARANGUAPE" Then
                vCodCliAcao = CONSULTA("CODIGO_ERP")
              Else
                vCodCliAcao = CONSULTA("COD_CLIENTE")
              End If
%>
              <button class="btn btn-warning btn-sm btn-linha" title="Lista as NFs" data-bs-toggle="modal" data-bs-target="#modalFatDet" onclick="openModalFatDet(`<%=vCidadeAcao%>`,`<%=vCodCliAcao%>`,`<%=CONSULTA("LOJA_CLIENTE")%>`,`<%=CONSULTA("NOME_CLIENTE")%>`)"><i class="bi bi-receipt"></i>NFs</button>
              <button class="btn btn-info btn-sm btn-linha"    title="Lista os Pedidos" data-bs-toggle="modal" data-bs-target="#modalPedCart" onclick="openModalPedCart(`<%=vCidadeAcao%>`,`<%=vCodCliAcao%>`,`<%=CONSULTA("LOJA_CLIENTE")%>`,`<%=CONSULTA("NOME_CLIENTE")%>`)"><i class="bi bi-cart3"></i>Ped.</button>
            </td>
            <td class="text-center col-fix-1">
              <% If CONSULTA("STATUS") = "INATIVO" Then %>
                <span class="badge-cli-inativo"><i class="bi bi-slash-circle-fill"></i>Inativo</span>
              <% ElseIf CONSULTA("STATUS") = "ATIVO" Then %>
                <span class="badge-cli-ativo"><i class="bi bi-check-circle-fill"></i>Ativo</span>
              <% End If %>
            </td>
            <td class="col-fix-2"><%=CONSULTA("CIDADE")%></td>
            <td class="text-end"><%=vCodCliAcao%></td>
            <td class="text-end"><%=CONSULTA("LOJA_CLIENTE")%></td>
            <td class="text-end"><%=CONSULTA("COD_SAP_CLIENTE")%></td>
            <td class="col-razao-social"><%=CONSULTA("NOME_CLIENTE")%></td>
            <td><%=CONSULTA("FANTASIA_CLIENTE")%></td>
            <td><%=CONSULTA("CNPJ_CLIENTE")%></td>
            <td><%=CONSULTA("INSCRICAO_ESTADUAL")%></td>
            <td><%=CONSULTA("TELEFONE")%></td>
            <td><%=CONSULTA("TIPO_FRETE")%></td>
            <td><%=CONSULTA("END_CLIENTE")%></td>
            <td><%=CONSULTA("BAIRRO")%></td>
            <td><%=CONSULTA("CEP_CLIENTE")%></td>
            <td><%=CONSULTA("ESTADO_CLIENTE")%></td>
            <td><%=CONSULTA("MUNICIPIO_CLIENTE")%></td>
            <td class="text-end"><%=FormatNumber(CONSULTA("LIMITE_CREDITO"), 2)%></td>
            <td class="text-center"><%=CONSULTA("RISCO")%></td>
            <td class="text-end"><%=CONSULTA("TES")%></td>
<%
              Dim vCondPagto
              If Trim(CONSULTA("COD_COND_PAGTO")) <> "" Then
                vCondPagto = "[" & Trim(CONSULTA("COD_COND_PAGTO")) & "] " & CONSULTA("CONDICAO_PAGAMENTO")
              Else
                vCondPagto = ""
              End If
%>
            <td><%=vCondPagto%></td>
            <td class="text-end"><%=CONSULTA("COD_VENDEDOR")%></td>
            <td><%=CONSULTA("NOME_VENDEDOR")%></td>
            <td><%=CONSULTA("DESCRICAO_SEGMENTO")%></td>
            <td><%=formataDataSigatoBR(CONSULTA("DATA_CADASTRO"))%></td>
            <td class="text-center"><%=CONSULTA("TOLMAIS")%></td>
            <td class="text-center"><%=CONSULTA("TOLMENOS")%></td>
            <td class="text-center"><%=CONSULTA("DATA_MAIOR")%></td>
            <td class="text-center"><%=CONSULTA("DATA_MENOR")%></td>
          </tr>
<%
      CONSULTA.MoveNext()
    Loop
  Else
%>
          <tr>
            <td colspan="29" class="tbl-vazia">
              <i class="bi bi-exclamation-circle me-1"></i>Nenhum registro encontrado para esse filtro.
            </td>
          </tr>
<%
  End If

  CONSULTA.Close()
  Set CONSULTA = Nothing
End If
%>
        </tbody>
      </table>
    </div>

    <div class="rodape-tabela d-none" id="rodapeTabela">
      <div id="seletorRegistros">
        <label class="form-label mb-0 me-1 seletor-label" for="cboRegPorPage">Mostrar</label>
        <select name="cboRegPorPage" id="cboRegPorPage" class="form-select form-select-sm seletor-select" title="Registros por página">
          <option value="10"   <% If vRegPorPage =   "10" Or vRegPorPage = "" Then %> selected <% End If %>>10</option>
          <option value="25"   <% If vRegPorPage =   "25" Then %> selected <% End If %>>25</option>
          <option value="50"   <% If vRegPorPage =   "50" Then %> selected <% End If %>>50</option>
          <option value="100"  <% If vRegPorPage =  "100" Then %> selected <% End If %>>100</option>
          <option value="500"  <% If vRegPorPage =  "500" Then %> selected <% End If %>>500</option>
          <option value="1000" <% If vRegPorPage = "1000" Then %> selected <% End If %>>1000</option>
          <option value="todos"<% If vRegPorPage = "todos" Then %> selected <% End If %>>Todos</option>
        </select>
      </div>
      <nav>
        <ul class="pagination pagination-sm mb-0 paginacao" id="paginacao">
<%
If Request("filtra") = "1" Then
  Dim pgAtual, pgTotal
  pgAtual = CInt(Request("page"))
  If pgAtual < 1 Then pgAtual = 1
  pgTotal = totalPages
%>
          <li class="page-item <%=iif(pgAtual <= 1, "disabled", "")%>"><a class="page-link" href="javascript:void(0)" onclick="pagination(1)"><i class="bi bi-chevron-bar-left"></i></a></li>
          <li class="page-item <%=iif(pgAtual <= 1, "disabled", "")%>"><a class="page-link" href="javascript:void(0)" onclick="pagination(<%=pgAtual - 1%>)"><i class="bi bi-chevron-left"></i></a></li>
<%
  Dim pgIni, pgFim, p
  pgIni = pgAtual - 2
  pgFim = pgAtual + 2
  If pgIni < 1       Then pgFim = pgFim + (1 - pgIni)       : pgIni = 1       : End If
  If pgFim > pgTotal Then pgIni = pgIni - (pgFim - pgTotal) : pgFim = pgTotal : End If
  If pgIni < 1       Then pgIni = 1 : End If
  For p = pgIni To pgFim
%>
          <li class="page-item <%=iif(p = pgAtual, "active", "")%>"><a class="page-link" href="javascript:void(0)" onclick="pagination(<%=p%>)"><%=p%></a></li>
<%
  Next
%>
          <li class="page-item <%=iif(pgAtual >= pgTotal, "disabled", "")%>"><a class="page-link" href="javascript:void(0)" onclick="pagination(<%=pgAtual + 1%>)"><i class="bi bi-chevron-right"></i></a></li>
          <li class="page-item <%=iif(pgAtual >= pgTotal, "disabled", "")%>"><a class="page-link" href="javascript:void(0)" onclick="pagination(<%=pgTotal%>)"><i class="bi bi-chevron-bar-right"></i></a></li>
<%
End If
%>
        </ul>
      </nav>
      <span id="contador">
        <i class="bi bi-table me-1"></i>
        Página <b id="pgAtualSpan"><%=iif(Request("filtra") = "1", CStr(pgAtual), "0")%></b> de <span id="pgTotalSpan"><%=iif(Request("filtra") = "1", CStr(pgTotal), "0")%></span>
      </span>
    </div>
    </div>

  </div>

  <div class="respiro-rodape"></div>

<%
If Trim(Request("oculta")) <> "1" Then
%>
  <div class="modal fade" id="modalFatDet" tabindex="-1" aria-labelledby="ModalLabelFatDet" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <form name="frmModalFatDet" id="frmModalFatDet" method="post">
        <div class="modal-content border-0 shadow">
          <div class="modal-header">
            <h6 class="modal-title" id="ModalLabelFatDet">Relatório de Faturamento Detalhado</h6>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">
            <p class="text-muted small">Informe o período para gerar o Relatório de Faturamento Detalhado.</p>
            <input type="hidden" id="m1Cidade" name="m1Cidade" />
            <input type="hidden" id="m1CodCli" name="m1CodCli" />
            <input type="hidden" id="m1LojCli" name="m1LojCli" />
            <input type="hidden" id="m1Razao"  name="m1Razao" />
            <div class="mb-2">
              <label class="form-label">Cliente</label>
              <input type="text" id="m1RazaoExibicao" class="form-control form-control-sm" readonly />
            </div>
            <div class="row g-2">
              <div class="col-md-4">
                <label for="dataini" class="form-label">Data Inicial</label>
                <input id="dataini" name="dataini" type="date" class="form-control form-control-sm" maxlength="10" />
              </div>
              <div class="col-md-4">
                <label for="datafin" class="form-label">Data Final</label>
                <input id="datafin" name="datafin" type="date" class="form-control form-control-sm" maxlength="10" />
              </div>
              <div class="col-md-4">
                <label for="cboTipoNotaFiscal" class="form-label">Tipo da Nota Fiscal</label>
                <select name="cboTipoNotaFiscal" id="cboTipoNotaFiscal" class="form-select form-select-sm">
                  <option value=""          <% If Trim(Request("cboTipoNotaFiscal")) = ""          Then %> selected="selected" <% End If %>>Todos</option>
                  <option value="VENDA"     <% If Trim(Request("cboTipoNotaFiscal")) = "VENDA"     Then %> selected="selected" <% End If %>>Sem devolução</option>
                  <option value="DEVOLUCAO" <% If Trim(Request("cboTipoNotaFiscal")) = "DEVOLUCAO" Then %> selected="selected" <% End If %>>Somente devolução</option>
                </select>
              </div>
            </div>
          </div>
          <div class="modal-footer bg-white">
            <button type="button" class="btn btn-outline-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
            <button type="button" class="btn btn-primary btn-sm" onclick="notasfiscais()">Confirmar</button>
          </div>
        </div>
      </form>
    </div>
  </div>

  <div class="modal fade" id="modalPedCart" tabindex="-1" aria-labelledby="ModalLabelPedCart" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <form name="frmModalPedCart" id="frmModalPedCart" method="post">
        <div class="modal-content border-0 shadow">
          <div class="modal-header">
            <h6 class="modal-title" id="ModalLabelPedCart">Relatório de Pedidos em Carteira</h6>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
          <div class="modal-body">
            <p class="text-muted small">Informe o período para gerar o Relatório de Pedidos em Carteira.</p>
            <input type="hidden" id="m2Cidade" name="m2Cidade" />
            <input type="hidden" id="m2CodCli" name="m2CodCli" />
            <input type="hidden" id="m2LojCli" name="m2LojCli" />
            <input type="hidden" id="m2Razao"  name="m2Razao" />
            <div class="mb-2">
              <label class="form-label">Cliente</label>
              <input type="text" id="m2RazaoExibicao" class="form-control form-control-sm" readonly />
            </div>
            <div class="row g-2">
              <div class="col-md-4">
                <label for="cboPeriodoPor" class="form-label">Período por</label>
                <select name="cboPeriodoPor" id="cboPeriodoPor" class="form-select form-select-sm" onclick="habilitacboTipoProduto(this)">
                  <option value="DTEMI" <% If Trim(Request("cboPeriodoPor")) = "DTEMI" Then %> selected="selected" <% End If %>>Data de Emissão</option>
                  <option value="DTENT" <% If Trim(Request("cboPeriodoPor")) = "DTENT" Then %> selected="selected" <% End If %>>Data de Entrega</option>
                  <option value="DTEMB" <% If Trim(Request("cboPeriodoPor")) = "DTEMB" Then %> selected="selected" <% End If %>>Data de Embarque</option>
                </select>
              </div>
              <div class="col-md-4">
                <label for="m2Dataini" class="form-label">Data Inicial</label>
                <input id="m2Dataini" name="m2Dataini" type="date" class="form-control form-control-sm" maxlength="10" placeholder="Data inicial" />
              </div>
              <div class="col-md-4">
                <label for="m2Datafin" class="form-label">Data Final</label>
                <input id="m2Datafin" name="m2Datafin" type="date" class="form-control form-control-sm" maxlength="10" placeholder="Data final" />
              </div>
            </div>
          </div>
          <div class="modal-footer bg-white">
            <button type="button" class="btn btn-outline-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
            <button type="button" class="btn btn-primary btn-sm" onclick="PedidosEmCarteira()">Confirmar</button>
          </div>
        </div>
      </form>
    </div>
  </div>
<%
End If
%>

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

  <script>
    (($) => {
      "use strict";

      const TOAST_CFG = Object.freeze({ delay: 4500, borda: true });

      const TOAST_ICON = Object.freeze({
        success: { cls: "bi-check-circle-fill",         cor: "#198754" },
        warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
        danger:  { cls: "bi-x-circle-fill",             cor: "#dc3545" },
        info:    { cls: "bi-info-circle-fill",          cor: "#6c757d" },
      });

      const UI = Object.freeze({
        exibirLoader()  { $("#loader").removeClass("d-none"); },
        ocultarLoader() { $("#loader").addClass("d-none"); },
        notificar(mensagem, tipo = "info") {
          const cfg  = TOAST_ICON[tipo] ?? TOAST_ICON.info;
          const item = $(`
            <div class="toast-item toast-item-${tipo}">
              <div class="toast-corpo">
                <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
                <span class="flex-grow-1">${mensagem}</span>
                <button class="toast-fechar" aria-label="Fechar">&times;</button>
              </div>
              <div class="toast-barra toast-barra-${tipo}" style="animation-duration:${TOAST_CFG.delay}ms;"></div>
            </div>
          `);
          item.find(".toast-fechar").on("click", () => item.remove());
          $("#toastWrapper").append(item);
          setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_CFG.delay);
        },
      });

      function pagination(aPage) {
        const enc = encodeURIComponent;
        const url = window.location.pathname +
          "?filtra=1" +
          "&page="                + aPage +
          "&cboRegPorPage="       + enc($("#hRegPorPage").val()) +
          "&cboPlanta="           + enc($("#cboPlanta").val()) +
          "&cboStatus="           + enc($("#cboStatus").val()) +
          "&cboSegmentoEconomico="+ enc($("#cboSegmentoEconomico").val()) +
          "&cboTipoFrete="        + enc($("#cboTipoFrete").val()) +
          "&codcli="              + enc($("#codcli").val()) +
          "&lojacli="             + enc($("#lojacli").val()) +
          "&razao="               + enc($("#razao").val()) +
          "&cboRepresentante="    + enc($("#cboRepresentante").val());
        UI.exibirLoader();
        window.location.href = url;
      }

      const PLANTAS_ERP = Object.freeze(["BENTO", "MARANGUAPE"]);

      const Formato = Object.freeze({
        dataStamp() {
          const d = new Date();
          return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
        },
        slugTexto(txt, limite = 30) {
          return (txt || "").replace(/[^a-zA-Z0-9À-ÿ\s]/g, "").replace(/\s+/g, "_").substring(0, limite);
        }
      });

      function _nomePdfArquivo() {
        const planta = $("#cboPlanta").val() || "TODOS";
        const razao  = ($("#razao").val() || "").trim();
        const cliente = razao
          ? razao.replace(/[^a-zA-Z0-9À-ÿ\s]/g, "").replace(/\s+/g, "_").substring(0, 30)
          : "CLIENTES";
        const d  = new Date();
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `${planta}_${cliente}_${d.getFullYear()}${mm}${dd}.pdf`;
      }

      function _gerarPdf(dados) {
        const { jsPDF } = window.jspdf;
        const doc       = new jsPDF({ orientation: "landscape", format: "a3" });
        const PAG_L     = doc.internal.pageSize.getWidth();
        const PAG_H     = doc.internal.pageSize.getHeight();
        const MARGEM_L  = 10;
        const MARGEM_TB = 4;
        const HDR_H     = 26;
        const dataAtual = new Date().toLocaleDateString("pt-BR");

        const colunas = ["Status","Planta","Cód. Cliente","Cód. Loja","Cód. SAP","Razão Social","Nome Fantasia","CNPJ","I.E.","Telefone","Frete","Endereço","Bairro","CEP","UF","Município","Limite de Crédito","Risco Financeiro","TES","Cond. Pagamento","Cód. Representante","Nome Representante","Segmento Econômico","Data Cadastro","Tolerância Mais","Tolerância Menos","Tolerância Data Maior","Tolerância Data Menor"];

        const _desenharLogo = (callback) => {
          const img = new Image();
          img.crossOrigin = "anonymous";
          img.onload = () => {
            const canvas  = document.createElement("canvas");
            canvas.width  = img.naturalWidth;
            canvas.height = img.naturalHeight;
            canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
            callback(canvas.toDataURL("image/jpeg", 0.7));
          };
          img.onerror = () => callback(null);
          img.src = "/corporativo/comercial/relatorio/RelatorioDeClientes/Logo/logo-smurfit1.png";
        };

        const _cabecalho = (logoData) => {
          doc.setFillColor(255, 255, 255);
          doc.rect(0, 0, PAG_L, HDR_H, "F");
          doc.setDrawColor(200, 200, 200);
          doc.setLineWidth(0.4);
          doc.line(0, HDR_H, PAG_L, HDR_H);
          if (logoData) {
            doc.addImage(logoData, "JPEG", 14, 5, 38, 13);
          }
          doc.setFontSize(13);
          doc.setFont(undefined, "bold");
          doc.setTextColor(0, 0, 0);
          doc.text("Relatório de Clientes", PAG_L / 2, 11, { align: "center" });
          doc.setFontSize(7.5);
          doc.setFont(undefined, "normal");
          doc.setTextColor(100, 100, 100);
          doc.text(`${dataAtual} | ${dados.length} registro(s)`, PAG_L / 2, 18, { align: "center" });
        };

        _desenharLogo((logoData) => {
          _cabecalho(logoData);
          doc.autoTable({
            head: [colunas],
            body: dados,
            startY: HDR_H + 3,
            tableWidth: PAG_L - MARGEM_TB * 2,
            theme: "striped",
            styles: { fontSize: 4, cellPadding: 1, valign: "middle", overflow: "ellipsize", lineColor: [200, 200, 200], lineWidth: { right: 0.2 } },
            headStyles: { fillColor: [74, 74, 74], textColor: [255, 255, 255], fontStyle: "bold", halign: "center", fontSize: 4, lineColor: [120, 120, 120], lineWidth: { right: 0.2 } },
            columnStyles: { 0: { cellWidth: 14, halign: "center" }, 1: { cellWidth: 12 } },
            alternateRowStyles: { fillColor: [248, 248, 248] },
            margin: { left: MARGEM_TB, right: MARGEM_TB },
            compress: true,
            didParseCell: (data) => {
              if (data.section === "body" && data.column.index === 0) {
                const v = (data.cell.raw || "").toString().toUpperCase();
                if (v === "ATIVO") {
                  data.cell.styles.textColor = [26, 122, 60];
                  data.cell.styles.fillColor = [240, 255, 244];
                  data.cell.styles.fontStyle = "bold";
                } else if (v === "INATIVO") {
                  data.cell.styles.textColor = [192, 57, 43];
                  data.cell.styles.fillColor = [255, 240, 240];
                  data.cell.styles.fontStyle = "bold";
                }
              }
            },
            didDrawPage: () => _cabecalho(logoData),
          });
          const totalPags = doc.getNumberOfPages();
          for (let p = 1; p <= totalPags; p++) {
            doc.setPage(p);
            doc.setFontSize(7);
            doc.setTextColor(150);
            doc.text(`Página ${p} de ${totalPags}`, PAG_L - MARGEM_L, PAG_H - 5, { align: "right" });
          }
          doc.save(_nomePdfArquivo());
          UI.ocultarLoader();
          UI.notificar(`PDF gerado com ${dados.length} registro(s).`, "success");
        });
      }

      function exportarExcel() {
        if ($("#cboRegPorPage").val() === "todos") {
          _buscarTodosEExportar("excel");
        } else {
          _exportarExcelPagina();
        }
      }

      function exportarPDF() {
        if ($("#cboPlanta").val() === "") {
          UI.notificar("Selecione a filial antes de exportar!", "warning");
          return;
        }
        if ($("#cboRegPorPage").val() === "todos") {
          _buscarTodosEExportar("pdf");
        } else {
          _exportarPdfPagina();
        }
      }

      function _exportarExcelPagina() {
        const planta    = $("#cboPlanta").val() || "TODOS";
        const razao     = ($("#razao").val() || "").trim();
        const cliente   = razao ? Formato.slugTexto(razao) : "CLIENTES";
        const cabecalho = [
          "Status","Planta","Cód. Cliente","Loja","Cód. SAP","Razão Social","Nome Fantasia",
          "CNPJ","I.E.","Telefone","Frete","Endereço","Bairro","CEP","UF","Município",
          "Limite Crédito","Risco","TES","Cond. Pagamento",
          "Cód. Representante","Nome Representante","Segmento Econômico",
          "Data Cadastro","Tol. Mais","Tol. Menos","Tol. Data Maior","Tol. Data Menor"
        ];
        const linhas = [cabecalho];
        $("#GradeClientes tbody tr").each(function () {
          const celulas = $(this).find("td");
          if (celulas.length < 5) return;
          const linha = [];
          linha.push(celulas.eq(1).text().trim());
          celulas.each(function (i) {
            if (i <= 1) return;
            linha.push($(this).text().trim());
          });
          linhas.push(linha);
        });
        if (linhas.length <= 1) {
          UI.notificar("Nenhum dado na página para exportar.", "warning");
          return;
        }
        const wb         = XLSX.utils.book_new();
        const ws         = XLSX.utils.aoa_to_sheet(linhas);
        const maxLargura = cabecalho.map((h, i) => {
          const maxConteudo = Math.max(h.length, ...linhas.slice(1).map(r => (r[i] || "").toString().length));
          return { wch: Math.max(maxConteudo, h === "Razão Social" ? 40 : 0) };
        });
        ws["!cols"]   = maxLargura;
        ws["!freeze"] = { xSplit: 0, ySplit: 1 };
        XLSX.utils.book_append_sheet(wb, ws, "Clientes");
        XLSX.writeFile(wb, `RelatorioDeClientes_${planta}_${cliente}_${Formato.dataStamp()}.xlsx`);
        UI.notificar(`Planilha gerada com ${linhas.length - 1} registro(s).`, "success");
      }

      function _exportarPdfPagina() {
        if ($("#cboPlanta").val() === "") {
          UI.notificar("Selecione a filial antes de exportar!", "warning");
          return;
        }
        const dados  = [];
        const linhas = document.querySelector("#GradeClientes tbody").querySelectorAll("tr");
        linhas.forEach((tr) => {
          const celulas = tr.querySelectorAll("td");
          if (celulas.length < 3) return;
          const linha     = [];
          const statusTxt = celulas[1] ? celulas[1].innerText.trim() : "";
          linha.push(statusTxt);
          for (let j = 2; j < celulas.length; j++) {
            linha.push(celulas[j].innerText.trim());
          }
          dados.push(linha);
        });
        if (!dados.length) {
          UI.notificar("Nenhum dado na página para exportar.", "warning");
          return;
        }
        UI.exibirLoader();
        _gerarPdf(dados);
      }

      function _buscarTodosEExportar(formato) {
        if ($("#cboPlanta").val() === "") {
          UI.notificar("Selecione a filial antes de exportar!", "warning");
          return;
        }
        UI.exibirLoader();
        const enc = encodeURIComponent;
        const url = "RelatorioDeClientes.asp" +
          "?todos=1" +
          "&cboPlanta="            + enc($("#cboPlanta").val()) +
          "&cboStatus="            + enc($("#cboStatus").val()) +
          "&cboSegmentoEconomico=" + enc($("#cboSegmentoEconomico").val()) +
          "&cboTipoFrete="         + enc($("#cboTipoFrete").val()) +
          "&codcli="               + enc($("#codcli").val()) +
          "&lojacli="              + enc($("#lojacli").val()) +
          "&razao="                + enc($("#razao").val()) +
          "&cboRepresentante="     + enc($("#cboRepresentante").val());

        const xhr = new XMLHttpRequest();
        xhr.open("GET", url, true);
        xhr.timeout = 300000;
        xhr.ontimeout = function () {
          UI.ocultarLoader();
          UI.notificar("Tempo esgotado. Tente aplicar mais filtros para reduzir o volume.", "warning");
        };
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4) return;
          UI.ocultarLoader();
          if (xhr.status !== 200) {
            UI.notificar("Erro ao buscar os dados. Status: " + xhr.status, "danger");
            return;
          }
          let registros;
          try {
            registros = JSON.parse(xhr.responseText);
          } catch (e) {
            UI.notificar("Erro ao processar a resposta do servidor.", "danger");
            return;
          }
          if (!Array.isArray(registros)) {
            const msg = (registros && registros.erro) ? registros.erro : "Resposta inesperada do servidor.";
            UI.notificar(msg, "danger");
            return;
          }
          if (registros.length === 0) {
            UI.notificar("Nenhum registro encontrado para exportar.", "warning");
            return;
          }
          const cabecalho = [
            "Status","Planta","Cód. Cliente","Loja","Cód. SAP","Razão Social","Nome Fantasia",
            "CNPJ","I.E.","Telefone","Frete","Endereço","Bairro","CEP","UF","Município",
            "Limite Crédito","Risco","TES","Cond. Pagamento",
            "Cód. Representante","Nome Representante","Segmento Econômico",
            "Data Cadastro","Tol. Mais","Tol. Menos","Tol. Data Maior","Tol. Data Menor"
          ];
          const linhas = registros.map((r) => [
            r.STATUS             || "",
            r.CIDADE             || "",
            r.COD_CLIENTE        || "",
            r.LOJA_CLIENTE       || "",
            r.COD_SAP_CLIENTE    || "",
            r.NOME_CLIENTE       || "",
            r.FANTASIA_CLIENTE   || "",
            r.CNPJ_CLIENTE       || "",
            r.INSCRICAO_ESTADUAL || "",
            r.TELEFONE           || "",
            r.TIPO_FRETE         || "",
            r.END_CLIENTE        || "",
            r.BAIRRO             || "",
            r.CEP_CLIENTE        || "",
            r.ESTADO_CLIENTE     || "",
            r.MUNICIPIO_CLIENTE  || "",
            r.LIMITE_CREDITO     || "",
            r.RISCO              || "",
            r.TES                || "",
            r.CONDICAO_PAGAMENTO || "",
            r.COD_VENDEDOR       || "",
            r.NOME_VENDEDOR      || "",
            r.DESCRICAO_SEGMENTO || "",
            r.DATA_CADASTRO      || "",
            r.TOLMAIS            || "",
            r.TOLMENOS           || "",
            r.DATA_MAIOR         || "",
            r.DATA_MENOR         || ""
          ]);

          UI.exibirLoader();
          if (formato === "excel") {
            const planta  = $("#cboPlanta").val() || "TODOS";
            const wb      = XLSX.utils.book_new();
            const ws      = XLSX.utils.aoa_to_sheet([cabecalho, ...linhas]);
            const maxLarg = cabecalho.map((h, i) => {
              const maxC = Math.max(h.length, ...linhas.map(r => (r[i] || "").toString().length));
              return { wch: Math.max(maxC, h === "Razão Social" ? 40 : 0) };
            });
            ws["!cols"]   = maxLarg;
            ws["!freeze"] = { xSplit: 0, ySplit: 1 };
            XLSX.utils.book_append_sheet(wb, ws, "Clientes");
            XLSX.writeFile(wb, `RelatorioDeClientes_${planta}_TODOS_${Formato.dataStamp()}.xlsx`);
            UI.ocultarLoader();
            UI.notificar(`Planilha gerada com ${linhas.length} registro(s).`, "success");
          } else {
            _gerarPdf(linhas);
          }
        };
        xhr.send();
      }

      function notasfiscais() {
        const url = window.location.protocol + "//" +
          window.location.hostname + "/" +
          "corporativo/comercial/relatorio/RelatorioFaturamentoDetalhado.asp" +
          "?cboCidade="          + document.frmModalFatDet.m1Cidade.value +
          "&codcli="             + document.frmModalFatDet.m1CodCli.value +
          "&lojacli="            + document.frmModalFatDet.m1LojCli.value +
          "&razao="              + document.frmModalFatDet.m1Razao.value +
          "&dataini="            + document.frmModalFatDet.dataini.value +
          "&datafin="            + document.frmModalFatDet.datafin.value +
          "&cboTipoNotaFiscal="  + document.frmModalFatDet.cboTipoNotaFiscal.value +
          "&filtra=1";
        window.open(url, "_blank");
      }

      function PedidosEmCarteira() {
        const url = window.location.protocol + "//" +
          window.location.hostname + "/" +
          "corporativo/comercial/relatorio/RelatorioPedidosEmCarteira.asp" +
          "?cboCidade="      + document.frmModalPedCart.m2Cidade.value +
          "&codcli="         + document.frmModalPedCart.m2CodCli.value +
          "&lojacli="        + document.frmModalPedCart.m2LojCli.value +
          "&razao="          + document.frmModalPedCart.m2Razao.value +
          "&dataini="        + document.frmModalPedCart.m2Dataini.value +
          "&datafin="        + document.frmModalPedCart.m2Datafin.value +
          "&cboPeriodoPor="  + document.frmModalPedCart.cboPeriodoPor.value +
          "&filtra=1";
        window.open(url, "_blank");
      }

      function Filtro() {
        if ($("#cboPlanta").val() === "") {
          UI.notificar("Selecione a filial!", "warning");
          return;
        }
        $("#filtra").val("1");
        $("#page").val(1);
        UI.exibirLoader();
        document.frmFiltroDeCliente.submit();
      }

      function limparFiltro(obj) {
        if (obj && obj.name === "cboPlanta") return;
        $("#filtra").val("0");
        $("#limpar").val("1");
        document.frmFiltroDeCliente.submit();
      }

      function pesquisaCliente(obj, pStatus) {
        if (typeof obj.value !== "undefined" && obj.value.trim() !== "") {
          const lnk = "validaBICLIENTES_ajax.asp?planta=" + $("#cboPlanta").val() +
            "&cliente=" + $("#codcli").val() +
            "&loja="    + obj.value +
            "&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Grupo&status=" + pStatus;
          ObjAjax(lnk);
        } else {
          obj.value = "";
        }
      }

      function objReturn(obj) {
        if (obj.erro === "ErroLogof") {
          UI.notificar("Favor se conectar novamente ao SWNET. Seu logon EXPIROU.", "danger");
          if (obj.tipo === "Grupo") {
            $("#codcli").val("");
            $("#filtroCodClienteDisplay").val("");
            $("#lojacli").val("");
            $("#razao").val("");
          }
          return false;
        }
        if (obj.erro === "ErroNEncontrado") {
          if (obj.tipo === "Grupo") {
            $("#codcli").val("");
            $("#filtroCodClienteDisplay").val("");
            $("#lojacli").val("");
            $("#razao").val("");
            UI.notificar("Cliente não encontrado. Possíveis causas: bloqueado, fora da sua carteira ou não cadastrado.", "warning");
          }
          return false;
        }
        if (obj.erro === "ErroCliente") {
          UI.notificar("Favor preencher os campos Cliente e Loja.", "warning");
          if (obj.tipo === "Grupo") {
            $("#codcli").val("");
            $("#filtroCodClienteDisplay").val("");
            $("#lojacli").val("");
            $("#razao").val("");
          }
          return false;
        }
        if (obj.tipo === "Grupo") {
          $("#razao").val((obj.val1 || "").replace(/^[\d.\/-]+\s*/, "").trim());
          $("#filtroCodClienteDisplay").val($("#codcli").val());
        }
      }

      function ConsultaCliente() {
        if ($("#cboPlanta").val() !== "") {
          const screenHeight = window.screen.availHeight;
          const lnk = "buscaBICLIENTES.asp?id1=codcli&id2=lojacli&id3=razao&id6=" +
            $("#cboPlanta").val() + "&id9=" + $("#cboStatus").val();
          window.open(lnk, "Consulta Cliente", `left,width=650,height=${screenHeight}`);
        } else {
          UI.notificar("Selecione a Filial", "warning");
        }
      }

      function pesquisarCliente() {
        if ($("#cboPlanta").val() !== "") {
          const screenHeight = window.screen.availHeight;
          const lnk = "buscaBICLIENTES.asp?id1=codcli&id2=lojacli&id3=razao&id6=" +
            $("#cboPlanta").val() + "&id9=" + $("#cboStatus").val();
          window.open(lnk, "Consulta Cliente", `left,width=650,height=${screenHeight}`);
        } else {
          UI.notificar("Selecione a Filial", "warning");
        }
      }

      function _decodeHtml(str) {
        const tmp = document.createElement("textarea");
        tmp.innerHTML = str;
        return tmp.value;
      }

      function atualizarRepresentantes(sel) {
        const codPlanta          = $("#cboPlanta").val();
        const comboRepresentante = document.getElementById("cboRepresentante");
        comboRepresentante.innerHTML = "";

        if (!codPlanta) {
          const opt   = document.createElement("option");
          opt.value   = "";
          opt.text    = "Selecione a planta primeira...";
          comboRepresentante.add(opt);
          return;
        }

        const xhr = new XMLHttpRequest();
        xhr.open("GET", "getRepresentantes.asp?codPlanta=" + codPlanta, true);
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4 || xhr.status !== 200) return;
          const representantes = JSON.parse(xhr.responseText);
          if (representantes.length > 0) {
            const optPadrao   = document.createElement("option");
            optPadrao.value   = "";
            optPadrao.text    = "Selecione...";
            comboRepresentante.add(optPadrao);
            representantes.forEach((rep) => {
              const opt   = document.createElement("option");
              opt.value   = rep.COD_VENDEDOR;
              opt.text    = _decodeHtml(rep.NOME_VENDEDOR);
              if (sel === rep.COD_VENDEDOR) opt.selected = true;
              comboRepresentante.add(opt);
            });
          } else {
            const opt   = document.createElement("option");
            opt.value   = "";
            opt.text    = "Não foram encontrados representantes para essa planta.";
            comboRepresentante.add(opt);
          }
        };
        xhr.send();
      }

      function obterRepresentanteSelecionado(obj) {
        $("#selRep").val(obj.value);
      }

      function habilitaCodLoja(obj) {
        const inputCodCli = document.getElementById("filtroCodClienteDisplay");
        const inputLojCli = document.getElementById("lojacli");
        if (!inputCodCli || !inputLojCli) return;
        if (obj.value === "BENTO" || obj.value === "MARANGUAPE") {
          $("#lojacli").val("01");
          inputCodCli.addEventListener("blur", onBlurHandler);
          inputLojCli.removeEventListener("blur", onBlurHandler);
        } else {
          $("#lojacli").val("");
          inputLojCli.addEventListener("blur", onBlurHandler);
          inputCodCli.removeEventListener("blur", onBlurHandler);
        }
      }

      function onBlurHandler() {
        pesquisaCliente(this, $("#cboStatus").val());
      }

      function primeiroDiaDoMesFormatado() {
        const d  = new Date();
        const dd = String(new Date(d.getFullYear(), d.getMonth(), 1).getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `${d.getFullYear()}-${mm}-01`;
      }

      function diaAtualFormatado() {
        const d  = new Date();
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `${d.getFullYear()}-${mm}-${dd}`;
      }

      function openModalFatDet(pCidade, pCodCli, pLojCli, pRazao) {
        $("#m1Cidade").val(pCidade);
        $("#m1CodCli").val(pCodCli);
        $("#m1LojCli").val(pLojCli);
        $("#m1Razao").val(pRazao.trim());
        $("#m1RazaoExibicao").val(pRazao.trim());
        $("#dataini").val(primeiroDiaDoMesFormatado());
        $("#datafin").val(diaAtualFormatado());
      }

      function openModalPedCart(pCidade, pCodCli, pLojCli, pRazao) {
        $("#m2Cidade").val(pCidade);
        $("#m2CodCli").val(pCodCli);
        $("#m2LojCli").val(pLojCli);
        $("#m2Razao").val(pRazao.trim());
        $("#m2RazaoExibicao").val(pRazao.trim());
        $("#m2Dataini").val(primeiroDiaDoMesFormatado());
        $("#m2Datafin").val(diaAtualFormatado());
      }

      function revelarRodape() {
        const temDados = document.querySelector("#GradeClientes tbody tr td[colspan]") === null &&
                         document.querySelectorAll("#GradeClientes tbody tr").length > 0;
        if (temDados) {
          $("#rodapeTabela").removeClass("d-none");
        }
      }

      function anexarEventos() {
        $("#cboRegPorPage").on("change", function () {
          if ($(this).val() === "todos") {
            $("#hRegPorPage").val("todos");
            UI.notificar("Modo <b>Todos</b> ativo. Use os botões de exportação para baixar os dados.", "info");
            return;
          }
          $("#hRegPorPage").val($(this).val());
          pagination(1);
        });

        $("#btnLimparCliente").on("click", function () {
          $("#codcli").val("");
          $("#filtroCodClienteDisplay").val("");
          $("#lojacli").val("");
          $("#razao").val("");
        });

        $("#filtroCodClienteDisplay").on("input", function () {
          $("#codcli").val($(this).val());
        });

        if ($("#codcli").val()) {
          $("#filtroCodClienteDisplay").val($("#razao").val() || $("#codcli").val());
        }

        $("#modalFatDet").on("show.bs.modal", function () {
          $("#cboTipoNotaFiscal").val("");
        });

        $("#modalPedCart").on("show.bs.modal", function () {
        });
      }

      function fixarCabecalhoDuplo() {
        const theadRows = document.querySelectorAll("#GradeClientes thead tr");
        if (theadRows.length < 2) return;
        const alturaLinha1 = theadRows[0].getBoundingClientRect().height;
        theadRows[1].querySelectorAll("th").forEach(th => {
          th.style.top = alturaLinha1 + "px";
        });
      }

      $(function () {
        UI.ocultarLoader();

        const vlrPlanta        = $("#cboPlanta").val();
        const vlrRepresentante = '<%=iif(vLimpando, "", Trim(Request("selRep")))%>';

        if (vlrPlanta !== "") {
          atualizarRepresentantes(vlrRepresentante !== "" ? vlrRepresentante : "");
        }

        $("#oculta").val(0);

        if ($("#cboPlanta").val() === "BENTO" || $("#cboPlanta").val() === "MARANGUAPE") {
          habilitaCodLoja(document.getElementById("cboPlanta"));
        }

        revelarRodape();
        anexarEventos();
        fixarCabecalhoDuplo();
      });

      window.pagination                    = pagination;
      window.exportarExcel                 = exportarExcel;
      window.exportarPDF                   = exportarPDF;
      window.notasfiscais                  = notasfiscais;
      window.PedidosEmCarteira             = PedidosEmCarteira;
      window.Filtro                        = Filtro;
      window.limparFiltro                  = limparFiltro;
      window.ConsultaCliente               = ConsultaCliente;
      window.pesquisarCliente              = pesquisarCliente;
      window.atualizarRepresentantes       = atualizarRepresentantes;
      window.obterRepresentanteSelecionado = obterRepresentanteSelecionado;
      window.habilitaCodLoja               = habilitaCodLoja;
      window.objReturn                     = objReturn;
      window.openModalFatDet               = openModalFatDet;
      window.openModalPedCart              = openModalPedCart;

    })(jQuery);
  </script>

</body>
</html>
