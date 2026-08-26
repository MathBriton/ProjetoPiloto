<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../cnn2/cnn2.asp"-->
<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Session.LCID     = 1046

' -------------------------------------------------------------------------------------------------
' CAMINHO DO ARQUIVO EXCEL (mantido igual ao original)
' -------------------------------------------------------------------------------------------------
Dim nomeCaminhoExcel
nomeCaminhoExcel = "C:\inetpub\wwwroot\corporativo\suprimentos\Arquivos\Fornecedores.xlsx"

' -------------------------------------------------------------------------------------------------
' HELPERS DE LEITURA DO EXCEL via ADODB + ACE OLEDB
' Substitui o Excel.asp do cnn/ que dependia de variáveis globais do cnn.asp legado.
' -------------------------------------------------------------------------------------------------
Dim cnnExcel, rsExcel

Sub AbrirExcel()
    Dim strConnExcel
    strConnExcel = "Provider=Microsoft.ACE.OLEDB.12.0;" &_
                   "Data Source=" & nomeCaminhoExcel & ";" &_
                   "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

    On Error Resume Next
    Set cnnExcel = Server.CreateObject("ADODB.Connection")
    cnnExcel.Open strConnExcel
    If Err.Number <> 0 Then
        Set cnnExcel = Nothing
        Set rsExcel  = Nothing
        Err.Clear
        Exit Sub
    End If

    Set rsExcel = Server.CreateObject("ADODB.Recordset")
    rsExcel.Open "SELECT * FROM [Planilha1$A1:O100000]", cnnExcel, 3, 1
    If Err.Number <> 0 Then
        Set rsExcel = Nothing
        cnnExcel.Close
        Set cnnExcel = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub FecharExcel()
    On Error Resume Next
    If IsObject(rsExcel)  Then rsExcel.Close  : Set rsExcel  = Nothing
    If IsObject(cnnExcel) Then cnnExcel.Close : Set cnnExcel = Nothing
    On Error GoTo 0
End Sub

Function ExcelOk()
    ExcelOk = IsObject(rsExcel) And Not (rsExcel Is Nothing)
End Function

' -------------------------------------------------------------------------------------------------
' HELPERS JSON
' EscapeJSON de funcoes.asp tem bug (usa variavel "resultado" em vez de "result" -- apaga aspas).
' Usar versao local. Barra invertida via Chr(92) para evitar erro de parse do VBScript.
' -------------------------------------------------------------------------------------------------
Function EscapeJSONLocal(str)
    If IsNull(str) Then str = ""
    Dim result
    result = CStr(str)
    result = Replace(result, Chr(92), Chr(92) & Chr(92))
    result = Replace(result, """",    Chr(92) & """")
    result = Replace(result, vbCrLf,  Chr(92) & "n")
    result = Replace(result, vbCr,    Chr(92) & "n")
    result = Replace(result, vbLf,    Chr(92) & "n")
    result = Replace(result, vbTab,   Chr(92) & "t")
    EscapeJSONLocal = Trim(result)
End Function

Function JsonStr(v)
    If IsNull(v) Then
        JsonStr = """"""
    Else
        JsonStr = """" & EscapeJSONLocal(CStr(v)) & """"
    End If
End Function

' -------------------------------------------------------------------------------------------------
' DISPATCHER DE ACOES -- toda resposta JSON sai aqui
' -------------------------------------------------------------------------------------------------
Dim acao
acao = Trim(Request("acao"))

If acao <> "" Then

    Response.ContentType = "application/json"
    Response.Expires     = 0

    Dim resultado
    resultado = ""

    Select Case acao

        Case "importarCompleto"
            resultado = ImportarCompleto()

        Case "atualizarIQS"
            resultado = AtualizarIQS()

        Case "previsualizar"
            resultado = Previsualizar()

        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Acao invalida: " & EscapeJSONLocal(acao) & """}"

    End Select

    Response.Write resultado
    Response.End()

End If

' =================================================================================================
' FUNCAO: Previsualizar
' Le o Excel e retorna as primeiras linhas como JSON (sem gravar no banco)
' =================================================================================================
Function Previsualizar()

    Dim jsonLinhas, jsonCabecalhos, primeiroItem, primeiroCampo
    Dim x, curValue, y

    AbrirExcel()

    If Not ExcelOk() Then
        Previsualizar = "{""sucesso"": false, ""mensagem"": ""Nao foi possivel abrir o arquivo Excel. Verifique se ele existe em: " & EscapeJSONLocal(nomeCaminhoExcel) & """}"
        Exit Function
    End If

    ' Monta cabecalhos
    jsonCabecalhos = "["
    primeiroItem = True
    For x = 0 To rsExcel.Fields.Count - 1
        If Not primeiroItem Then jsonCabecalhos = jsonCabecalhos & ","
        jsonCabecalhos = jsonCabecalhos & JsonStr(rsExcel.Fields(x).Name)
        primeiroItem = False
    Next
    jsonCabecalhos = jsonCabecalhos & "]"

    ' Monta linhas (maximo 200 para preview)
    jsonLinhas   = "["
    primeiroItem = True
    y = 0

    Do While Not rsExcel.EOF And y < 200

        If Not primeiroItem Then jsonLinhas = jsonLinhas & ","

        jsonLinhas = jsonLinhas & "["
        primeiroCampo = True

        For x = 0 To rsExcel.Fields.Count - 1
            curValue = rsExcel.Fields(x).Value
            If IsNull(curValue) Then curValue = ""
            If Not primeiroCampo Then jsonLinhas = jsonLinhas & ","
            jsonLinhas = jsonLinhas & JsonStr(CStr(curValue))
            primeiroCampo = False
        Next

        jsonLinhas = jsonLinhas & "]"

        primeiroItem = False
        y = y + 1
        rsExcel.MoveNext

    Loop

    jsonLinhas = jsonLinhas & "]"

    FecharExcel()

    Previsualizar = "{""sucesso"": true, ""mensagem"": """", ""dados"": {""cabecalhos"": " & jsonCabecalhos & ", ""linhas"": " & jsonLinhas & "}}"

End Function

' =================================================================================================
' FUNCAO: ImportarCompleto
' Apaga FornecedoresInsumos e reinsere todos os dados do Excel
' =================================================================================================
Function ImportarCompleto()

    Dim usr, data, hora, dataHora
    Dim SQL1, SQL2, SQL3, SQL
    Dim x, curValue
    Dim cnn, rsVerifica

    usr      = Session("USR")
    data     = InverteData(Date())
    hora     = Time()
    dataHora = data & " " & hora

    ' --- Abre conexao direta (cnn2 Conecta() escreveria HTML em caso de erro, corrompendo o JSON)
    On Error Resume Next
    Set cnn = Server.CreateObject("ADODB.Connection")
    cnn.CommandTimeout = 300
    cnn.Open BuscaStringConexao("PIR-CORPORATIVO")
    If Err.Number <> 0 Or (cnn Is Nothing) Then
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Erro ao conectar ao banco: " & EscapeJSONLocal(Err.Description) & """}"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' --- Limpa tabela
    On Error Resume Next
    cnn.Execute "DELETE FROM FornecedoresInsumos"
    If Err.Number <> 0 Then
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Erro ao limpar tabela: " & EscapeJSONLocal(Err.Description) & """}"
        Err.Clear
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    ' --- Le Excel
    AbrirExcel()
    If Not ExcelOk() Then
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Nao foi possivel abrir o arquivo Excel em: " & EscapeJSONLocal(nomeCaminhoExcel) & """}"
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If

    SQL1 = "" &_
        "INSERT INTO dbo.FornecedoresInsumos ( " &_
            "Planta, " &_
            "Codigo, " &_
            "Razao, " &_
            "Endereco, " &_
            "UF, " &_
            "CEP, " &_
            "Cidade, " &_
            "Pais, " &_
            "CNPJ, " &_
            "IQS, " &_
            "DataHoraIntegracao, " &_
            "Usuario " &_
        ") VALUES "

    SQL  = "BEGIN TRANSACTION "
    SQL2 = ""

    Do While Not rsExcel.EOF

        SQL2 = SQL2 & SQL1

        For x = 0 To rsExcel.Fields.Count - 1

            curValue = rsExcel.Fields(x).Value
            If IsNull(curValue) Then curValue = "NA"
            curValue = CStr(curValue)

            ' Remove aspas simples E duplas de campos string (ambas quebram o SQL)
            Dim curValueStr
            curValueStr = Replace(Replace(curValue, "'", ""), Chr(34), "")

            If x = 1 Then
                SQL2 = SQL2 & "('" & curValueStr & "',"
            Else
                If x <> 0 And x <> 7 And x <> 8 And x <> 11 And x <> 12 Then
                    If x < 14 Then
                        SQL2 = SQL2 & "'" & curValueStr & "', "
                    Else
                        ' Campo numérico (IQS): valida antes de injetar — nunca concatenar texto livre
                        Dim curValueNum
                        curValueNum = Replace(curValue, Chr(34), "")
                        If IsNumeric(curValueNum) Then
                            SQL2 = SQL2 & CDbl(curValueNum) & ", "
                        Else
                            SQL2 = SQL2 & "NULL, "
                        End If
                    End If
                End If
            End If

        Next

        SQL2 = SQL2 & "'" & dataHora & "','" & usr & "') "
        rsExcel.MoveNext

    Loop

    FecharExcel()

    SQL = SQL & SQL2
    SQL = SQL & " IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION "

    On Error Resume Next
    cnn.Execute SQL
    If Err.Number <> 0 Then
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Erro ao executar importacao: " & EscapeJSONLocal(Err.Description) & """}"
        Err.Clear
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    ' --- Verifica resultado
    SQL3 = "" &_
        "SELECT COUNT(*) AS Numero " &_
        "FROM FornecedoresInsumos " &_
        "WHERE DataHoraIntegracao = '" & dataHora & "' " &_
            "AND Usuario = '" & usr & "'"

    On Error Resume Next
    Set rsVerifica = Server.CreateObject("ADODB.Recordset")
    rsVerifica.Open SQL3, cnn
    If Err.Number <> 0 Or (rsVerifica Is Nothing) Then
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Importacao executada mas nao foi possivel confirmar o resultado.""}"
        Err.Clear
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Dim numero
    numero = 0
    If Not rsVerifica.EOF Then numero = CInt(rsVerifica("Numero"))
    rsVerifica.Close : Set rsVerifica = Nothing
    cnn.Close        : Set cnn        = Nothing

    If numero > 0 Then
        ImportarCompleto = "{""sucesso"": true, ""mensagem"": ""Fornecedores importados com sucesso!"", ""dados"": {""registros"": " & numero & "}}"
    Else
        Call EnviaEmail("jorge.queiroz@smurfitkappa.com.br", "", "", "", "", "Pagina: corporativo/suprimentos/ImportacaoFornecedores.asp<br /><p>SQL: " & SQL & "</p>", "Erro Importacao de Fornecedores")
        ImportarCompleto = "{""sucesso"": false, ""mensagem"": ""Houve um erro e nenhuma alteracao foi realizada no cadastro de fornecedores!""}"
    End If

End Function

' =================================================================================================
' FUNCAO: AtualizarIQS
' Atualiza somente o campo IQS dos fornecedores existentes
' =================================================================================================
Function AtualizarIQS()

    Dim usr, data, hora, dataHora
    Dim SQL1, SQL2, SQL3, SQL
    Dim x, curValue, fornecedor
    Dim cnn, rsVerifica

    usr      = Session("USR")
    data     = InverteData(Date())
    hora     = Time()
    dataHora = data & " " & hora

    On Error Resume Next
    Set cnn = Server.CreateObject("ADODB.Connection")
    cnn.CommandTimeout = 300
    cnn.Open BuscaStringConexao("PIR-CORPORATIVO")
    If Err.Number <> 0 Or (cnn Is Nothing) Then
        AtualizarIQS = "{""sucesso"": false, ""mensagem"": ""Erro ao conectar ao banco: " & EscapeJSONLocal(Err.Description) & """}"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    AbrirExcel()
    If Not ExcelOk() Then
        AtualizarIQS = "{""sucesso"": false, ""mensagem"": ""Nao foi possivel abrir o arquivo Excel em: " & EscapeJSONLocal(nomeCaminhoExcel) & """}"
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If

    SQL1 = "UPDATE FornecedoresInsumos SET IQS = "
    SQL2 = ""
    SQL  = "BEGIN TRANSACTION "

    Do While Not rsExcel.EOF

        SQL2 = SQL2 & SQL1

        For x = 0 To rsExcel.Fields.Count - 1

            curValue = rsExcel.Fields(x).Value
            If IsNull(curValue) Then curValue = "NA"
            curValue = CStr(curValue)

            If x = 2 Then fornecedor = Replace(Replace(curValue, "'", ""), Chr(34), "")

            If x = 14 Then
                ' Campo numérico (IQS): valida antes de injetar — nunca concatenar texto livre
                Dim iqs
                iqs = Replace(curValue, Chr(34), "")
                If IsNumeric(iqs) Then
                    iqs = CDbl(iqs)
                Else
                    iqs = "NULL"
                End If
                SQL2 = SQL2 & iqs & ", DataAtualizacao = '" & dataHora & "', UsuarioAtualizou = '" & usr & "' WHERE Codigo = '" & fornecedor & "'; "
            End If

        Next

        rsExcel.MoveNext

    Loop

    FecharExcel()

    SQL = SQL & SQL2
    SQL = SQL & " IF @@ERROR > 0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION "

    On Error Resume Next
    cnn.Execute SQL
    If Err.Number <> 0 Then
        AtualizarIQS = "{""sucesso"": false, ""mensagem"": ""Erro ao executar atualizacao de IQS: " & EscapeJSONLocal(Err.Description) & """}"
        Err.Clear
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    SQL3 = "" &_
        "SELECT COUNT(*) AS Numero " &_
        "FROM FornecedoresInsumos " &_
        "WHERE DataAtualizacao = '" & dataHora & "' " &_
            "AND UsuarioAtualizou = '" & usr & "'"

    On Error Resume Next
    Set rsVerifica = Server.CreateObject("ADODB.Recordset")
    rsVerifica.Open SQL3, cnn
    If Err.Number <> 0 Or (rsVerifica Is Nothing) Then
        AtualizarIQS = "{""sucesso"": false, ""mensagem"": ""Atualizacao executada mas nao foi possivel confirmar o resultado.""}"
        Err.Clear
        cnn.Close : Set cnn = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Dim numero
    numero = 0
    If Not rsVerifica.EOF Then numero = CInt(rsVerifica("Numero"))
    rsVerifica.Close : Set rsVerifica = Nothing
    cnn.Close        : Set cnn        = Nothing

    If numero > 0 Then
        AtualizarIQS = "{""sucesso"": true, ""mensagem"": ""IQS dos Fornecedores atualizado com sucesso!"", ""dados"": {""registros"": " & numero & "}}"
    Else
        Call EnviaEmail("jorge.queiroz@smurfitkappa.com.br", "", "", "", "", "Pagina: corporativo/suprimentos/ImportacaoFornecedores.asp<br /><p>SQL: " & SQL & "</p>", "Erro Atualizacao IQS Fornecedores")
        AtualizarIQS = "{""sucesso"": false, ""mensagem"": ""Houve um erro e nenhuma alteracao foi realizada no IQS de fornecedores!""}"
    End If

End Function
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet - Importar Fornecedores</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .action-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .action-card-header {
      background-color: #d5d5d5;
      color: #000000;
      padding: .65rem 1rem;
      border-radius: 6px 6px 0 0;
      font-weight: 600;
      font-size: .9rem;
    }
    .action-card-body { padding: 1.25rem; }

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      position: relative;
    }
    .table-toolbar {
      background-color: #fff;
      border-radius: 6px 6px 0 0;
      padding: .6rem 1rem;
      border-bottom: 1px solid #dee2e6;
    }
    .table-wrapper { overflow-x: auto; overflow-y: auto; max-height: 60vh; }
    .table-wrapper table { min-width: max-content; }

    .table thead th {
      position: sticky; top: 0;
      background-color: #fff; color: #212529;
      font-size: .82rem; font-weight: 700;
      text-align: center; white-space: nowrap;
      z-index: 2; border-color: #dee2e6;
      padding: .55rem .75rem;
    }
    .table tbody td {
      font-size: .82rem; text-align: center;
      white-space: nowrap; color: #212529;
      vertical-align: middle; padding: .45rem .75rem;
    }

    .table-placeholder {
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      padding: 3.5rem 1rem; color: #6c757d; gap: .75rem;
    }
    .table-placeholder .bi { font-size: 2.8rem; opacity: .35; }
    .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

    .overlay-loading {
      position: absolute; inset: 0;
      background: rgba(255,255,255,.80);
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      z-index: 10; gap: 1rem; border-radius: 6px;
    }
    .overlay-loading .spinner-border { width: 2.5rem; height: 2.5rem; color: #6c757d; }
    .overlay-loading span            { font-size: .9rem; font-weight: 600; color: #495057; }

    .toast-container-custom {
      position: fixed; top: 1rem; right: 1rem;
      z-index: 1100; display: flex; flex-direction: column; gap: .5rem;
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
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }

    .badge-total {
      font-size: .82rem; font-weight: 600;
      background-color: #e9ecef; color: #495057;
      border-radius: 20px; padding: .3rem .8rem;
    }
  </style>
</head>
<body>

<div class="container-fluid py-3 px-3 px-lg-4">

  <div class="d-flex align-items-center gap-3 mb-3 flex-wrap">
    <button type="button" class="btn btn-light btn-sm" onclick="history.go(-1)">
      <i class="bi bi-arrow-left me-1"></i>Voltar
    </button>
    <span class="page-title" style="font-size:1.15rem;font-weight:700;color:#212529;">
      Importar Cadastro de Fornecedores SAP
    </span>
  </div>

  <div class="action-card mb-3">
    <div class="action-card-header">
      <i class="bi bi-cloud-upload me-2"></i>Operacoes de Importacao
    </div>
    <div class="action-card-body">
      <p class="text-muted mb-3" style="font-size:.88rem;">
        Selecione a operacao desejada. O arquivo de origem é
        <code>Fornecedores.xlsx</code> no diretorio configurado no servidor.
      </p>
      <div class="d-flex flex-wrap gap-2 align-items-center">
        <button id="btnPrevisualizar" type="button" class="btn btn-secondary btn-sm">
          <i class="bi bi-eye me-2"></i>Pre-visualizar
        </button>
        <button id="btnImportarCompleto" type="button" class="btn btn-primary btn-sm">
          <i class="bi bi-database-fill-up me-2"></i>Importar Completo
        </button>
        <button id="btnAtualizarIQS" type="button" class="btn btn-warning btn-sm">
          <i class="bi bi-arrow-repeat me-2"></i>Atualizar somente IQS
        </button>
      </div>
    </div>
  </div>

  <div class="table-card d-none" id="cardTabela">

    <div id="overlayLoading" class="overlay-loading d-none">
      <div class="spinner-border" role="status">
        <span class="visually-hidden">Carregando...</span>
      </div>
      <span>Processando...</span>
    </div>

    <div class="table-toolbar d-flex align-items-center justify-content-between">
      <span style="font-weight:600;font-size:.9rem;color:#212529;">
        <i class="bi bi-table me-2"></i>Pre-visualizacao do Arquivo
      </span>
      <span id="badgeTotal" class="badge-total d-none"></span>
    </div>

    <div class="table-wrapper">
      <div id="tablePlaceholder" class="table-placeholder">
        <i class="bi bi-clipboard-data me-2"></i>
        <p>Clique em <strong>Pre-visualizar</strong> para exibir os dados do arquivo.</p>
      </div>
      <table id="tabelaPreview" class="table table-bordered table-striped table-hover table-sm mb-0 d-none">
        <thead id="theadPreview"></thead>
        <tbody id="tbodyPreview"></tbody>
      </table>
    </div>

  </div>

</div>

<div class="toast-container-custom" id="toastContainer"></div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

<script>
"use strict";

const ImportacaoFornecedores = (function () {

  const REPOSITORIO = "ImportacaoFornecedores.asp";
  const TOAST_DELAY = 4500;

  const TOAST_CFG = Object.freeze({
    sucesso : { icone: "bi-check-circle-fill",         cor: "#198754", barra: "#198754" },
    erro    : { icone: "bi-x-circle-fill",             cor: "#dc3545", barra: "#dc3545" },
    aviso   : { icone: "bi-exclamation-triangle-fill", cor: "#856404", barra: "#ffc107" },
    info    : { icone: "bi-info-circle-fill",          cor: "#6c757d", barra: "#6c757d" }
  });

  function exibirToast(tipo, mensagem) {
    const cfg     = TOAST_CFG[tipo] || TOAST_CFG.info;
    const titulos = { sucesso: "Sucesso", erro: "Erro", aviso: "Atencao", info: "Info" };
    const el      = document.createElement("div");
    el.className  = "toast-custom";
    el.innerHTML  = `
      <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
        <i class="bi ${cfg.icone}" style="color:${cfg.cor};font-size:1.1rem;"></i>
        <strong style="color:${cfg.cor};font-size:.9rem;">${titulos[tipo] || "Info"}</strong>
        <button type="button" class="btn-close btn-close-sm ms-auto"
                onclick="this.closest('.toast-custom').remove()" aria-label="Fechar"></button>
      </div>
      <div class="px-3 pb-2" style="font-size:.85rem;color:#212529;">${mensagem}</div>
      <div style="height:3px;background:#e9ecef;border-radius:0 0 10px 10px;">
        <div class="toast-progress-bar" style="height:100%;background:${cfg.barra};border-radius:inherit;"></div>
      </div>`;
    document.getElementById("toastContainer").appendChild(el);
    setTimeout(() => { if (el.parentNode) el.remove(); }, TOAST_DELAY + 400);
  }

  function mostrarOverlay() { $("#overlayLoading").removeClass("d-none"); }
  function ocultarOverlay() { $("#overlayLoading").addClass("d-none");    }

  function previsualizar() {
    $("#cardTabela").removeClass("d-none");
    mostrarOverlay();

    $.ajax({
      url      : REPOSITORIO,
      method   : "GET",
      dataType : "json",
      data     : { acao: "previsualizar" },
      success  : function (data) {
        if (!data.sucesso) {
          exibirToast("erro", data.mensagem || "Erro ao pre-visualizar.");
          return;
        }

        const cabecalhos = data.dados.cabecalhos;
        const linhas     = data.dados.linhas;

        let thHtml = "<tr><th>#</th>";
        cabecalhos.forEach(function (c) {
          thHtml += "<th>" + $("<div>").text(c).html() + "</th>";
        });
        thHtml += "</tr>";
        $("#theadPreview").html(thHtml);

        const frag = document.createDocumentFragment();
        linhas.forEach(function (linha, idx) {
          const tr   = document.createElement("tr");
          const tdN  = document.createElement("td");
          tdN.textContent = idx + 1;
          tr.appendChild(tdN);
          linha.forEach(function (cel) {
            const td = document.createElement("td");
            td.textContent = cel;
            tr.appendChild(td);
          });
          frag.appendChild(tr);
        });
        const tbody = document.getElementById("tbodyPreview");
        tbody.innerHTML = "";
        tbody.appendChild(frag);

        $("#tablePlaceholder").addClass("d-none");
        $("#tabelaPreview").removeClass("d-none");
        $("#badgeTotal").text(linhas.length + " linha(s) exibida(s)").removeClass("d-none");
        exibirToast("sucesso", "Arquivo carregado com " + linhas.length + " linha(s).");
      },
      error    : function (xhr) {
        exibirToast("erro", "Erro de comunicacao ao pre-visualizar.");
        console.error(xhr.responseText);
      },
      complete : ocultarOverlay
    });
  }

  function executarImportacao(acao, labelOperacao) {
    if (!confirm("Confirma a operacao: " + labelOperacao + "?\n\nEsta acao nao pode ser desfeita.")) {
      return;
    }

    $("#cardTabela").removeClass("d-none");
    mostrarOverlay();

    $.ajax({
      url      : REPOSITORIO,
      method   : "POST",
      dataType : "json",
      data     : { acao: acao },
      success  : function (data) {
        if (data.sucesso) {
          const qtd = (data.dados && data.dados.registros) ? data.dados.registros : 0;
          exibirToast("sucesso", data.mensagem + (qtd > 0 ? " (" + qtd + " registro(s))" : ""));
        } else {
          exibirToast("erro", data.mensagem || "Ocorreu um erro na operacao.");
        }
      },
      error    : function (xhr) {
        exibirToast("erro", "Erro de comunicacao durante a operacao.");
        console.error(xhr.responseText);
      },
      complete : ocultarOverlay
    });
  }

  function init() {
    $("#btnPrevisualizar").on("click", function () {
      previsualizar();
    });
    $("#btnImportarCompleto").on("click", function () {
      executarImportacao("importarCompleto", "Importacao Completa de Fornecedores");
    });
    $("#btnAtualizarIQS").on("click", function () {
      executarImportacao("atualizarIQS", "Atualizacao de IQS");
    });
  }

  return Object.freeze({ init });

})();

$(function () {
  ImportacaoFornecedores.init();
});
</script>

</body>
</html>
