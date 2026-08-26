<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#Include File="../includes/variaveis.asp"-->
<%
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046

' ================================================================
' BLOCO DE API — só executa quando vier o parâmetro "acao".
' ================================================================
Dim acao
acao = Trim("" & Request("acao"))

If acao <> "" Then

    Response.Charset     = "UTF-8"
    Response.CodePage    = 65001
    Response.ContentType = "application/json"

    Dim resultado
    Select Case acao
        Case "obterDadosIniciais" : resultado = ObterDadosIniciais()
        Case "buscarPlaca"        : resultado = BuscarPlaca()
        Case "salvar"             : resultado = Salvar()
        Case Else
            resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & EscapeJSON(acao) & """}"
    End Select

    Response.Write resultado
    Response.End()

End If

' ════════════════════════════════════════════════════════════════
' FUNÇÃO: ObterDadosIniciais
' ════════════════════════════════════════════════════════════════
Function ObterDadosIniciais()

    Dim jsonTipos, primeiroTipo
    primeiroTipo = True
    jsonTipos    = "["

    Dim rsTipos
    Set rsTipos = Conecta("PIR-CORPORATIVO", "SELECT ID, DESCRICAO FROM CHECKLIST_TIPOVEICULO WHERE ATIVO = 1 ORDER BY DESCRICAO")
    If rsTipos Is Nothing Then
        ObterDadosIniciais = "{""sucesso"": false, ""mensagem"": ""Erro ao carregar tipos de veículo.""}"
        Exit Function
    End If

    Do While Not rsTipos.EOF
        If Not primeiroTipo Then jsonTipos = jsonTipos & ","
        jsonTipos = jsonTipos & "{"
        jsonTipos = jsonTipos & """id"": "         & (rsTipos("ID")        & "") & ", "
        jsonTipos = jsonTipos & """descricao"": """ & EscapeJSON(rsTipos("DESCRICAO") & "") & """"
        jsonTipos = jsonTipos & "}"
        primeiroTipo = False
        rsTipos.MoveNext()
    Loop
    rsTipos.Close()
    Set rsTipos = Nothing

    jsonTipos = jsonTipos & "]"

    Dim jsonItens, primeiroItem
    primeiroItem = True
    jsonItens    = "["

    Dim rsItens
    Set rsItens = Conecta("PIR-CORPORATIVO", "SELECT ID, DESCRICAO FROM CHECKLIST_LISTAITENS WHERE ATIVO = 1 ORDER BY ID")
    If rsItens Is Nothing Then
        ObterDadosIniciais = "{""sucesso"": false, ""mensagem"": ""Erro ao carregar itens avaliados.""}"
        Exit Function
    End If

    Do While Not rsItens.EOF
        If Not primeiroItem Then jsonItens = jsonItens & ","
        jsonItens = jsonItens & "{"
        jsonItens = jsonItens & """id"": "         & (rsItens("ID")        & "") & ", "
        jsonItens = jsonItens & """descricao"": """ & EscapeJSON(rsItens("DESCRICAO") & "") & """"
        jsonItens = jsonItens & "}"
        primeiroItem = False
        rsItens.MoveNext()
    Loop
    rsItens.Close()
    Set rsItens = Nothing

    jsonItens = jsonItens & "]"

    Dim json
    json = "{"
    json = json & """sucesso"": true, "
    json = json & """mensagem"": """", "
    json = json & """dados"": {"
    json = json & """tiposVeiculo"": "   & jsonTipos & ", "
    json = json & """itensAvaliados"": " & jsonItens
    json = json & "}}"

    ObterDadosIniciais = json
End Function

' ════════════════════════════════════════════════════════════════
' FUNÇÃO: BuscarPlaca
' ════════════════════════════════════════════════════════════════
Function BuscarPlaca()

    Dim placa
    placa = Trim("" & Request("placa"))

    If placa = "" Then
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Placa não informada.""}"
        Exit Function
    End If

    placa = Replace(placa, "'", "''")

    Dim sql
    sql = "SELECT TOP 1 "
    sql = sql & "  NOMEPROPRIETARIO, TELEFONE, DOCUMENTO, "
    sql = sql & "  NOMEMOTORISTA, ID_TIPOVEICULO, QTDEEIXO, "
    sql = sql & "  ALTURA, COMPRIMENTO, LARGURA "
    sql = sql & "FROM CHECKLIST_TRANSPORTES "
    sql = sql & "WHERE UPPER(PLACAVEICULO) = UPPER('" & placa & "') "
    sql = sql & "  AND ATIVO = 1 "
    sql = sql & "ORDER BY ID DESC"

    On Error Resume Next
    Dim rsPlaca
    Set rsPlaca = Conecta("PIR-CORPORATIVO", sql)
    If Err.Number <> 0 Or rsPlaca Is Nothing Then
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar placa: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    If rsPlaca.EOF Then
        rsPlaca.Close()
        Set rsPlaca = Nothing
        BuscarPlaca = "{""sucesso"": false, ""mensagem"": ""Placa não encontrada.""}"
        Exit Function
    End If

    Dim json
    json = "{"
    json = json & """sucesso"": true, "
    json = json & """mensagem"": """", "
    json = json & """dados"": {"
    json = json & """nomeProprietario"": """ & EscapeJSON(rsPlaca("NOMEPROPRIETARIO") & "") & """, "
    json = json & """telefone"":         """ & EscapeJSON(rsPlaca("TELEFONE")         & "") & """, "
    json = json & """cnpjCpf"":          """ & EscapeJSON(rsPlaca("DOCUMENTO")        & "") & """, "
    json = json & """nomeMotorista"":    """ & EscapeJSON(rsPlaca("NOMEMOTORISTA")    & "") & """, "
    json = json & """idTipoVeiculo"":    "   & (rsPlaca("ID_TIPOVEICULO")             & "") & ", "
    json = json & """quantidadeEixo"":   "   & (rsPlaca("QTDEEIXO")                  & "") & ", "
    json = json & """altura"":           "   & Replace(rsPlaca("ALTURA")      & "0", ",", ".") & ", "
    json = json & """comprimento"":      "   & Replace(rsPlaca("COMPRIMENTO") & "0", ",", ".") & ", "
    json = json & """largura"":          "   & Replace(rsPlaca("LARGURA")     & "0", ",", ".")
    json = json & "}}"

    rsPlaca.Close()
    Set rsPlaca = Nothing

    BuscarPlaca = json
End Function

' ════════════════════════════════════════════════════════════════
' FUNÇÃO: Salvar
' ════════════════════════════════════════════════════════════════
Function Salvar()

    Dim nomeProprietario, telefone, nomeMotorista, cnpjCpf
    Dim placa, dataPreenchimento, idTipoVeiculo, quantidadeEixo
    Dim horarioInicio, horarioTermino, tipoTransporte, destinoCarga
    Dim altura, comprimento, largura, manutencao
    Dim nomeConferente, portariaResponsavel, status
    Dim itensAvaliados

    nomeProprietario    = Trim("" & Request("nomeProprietario"))
    telefone            = Trim("" & Request("telefone"))
    nomeMotorista       = Trim("" & Request("nomeMotorista"))
    cnpjCpf             = Trim("" & Request("cnpjCpf"))
    placa               = Trim("" & Request("placa"))
    dataPreenchimento   = Trim("" & Request("dataPreenchimento"))
    idTipoVeiculo       = Trim("" & Request("idTipoVeiculo"))
    quantidadeEixo      = Trim("" & Request("quantidadeEixo"))
    horarioInicio       = Trim("" & Request("horarioInicio"))
    horarioTermino      = Trim("" & Request("horarioTermino"))
    tipoTransporte      = Trim("" & Request("tipoTransporte"))
    destinoCarga        = Trim("" & Request("destinoCarga"))
    altura              = Trim("" & Request("altura"))
    comprimento         = Trim("" & Request("comprimento"))
    largura             = Trim("" & Request("largura"))
    manutencao          = Trim("" & Request("manutencao"))
    nomeConferente      = Trim("" & Request("nomeConferente"))
    portariaResponsavel = Trim("" & Request("portariaResponsavel"))
    status              = Trim("" & Request("status"))
    itensAvaliados      = Trim("" & Request("itensAvaliados"))

    Dim erros
    erros = ""
    If nomeProprietario  = ""            Then erros = erros & "Nome do Proprietário não informado. "
    If nomeMotorista     = ""            Then erros = erros & "Nome do Motorista não informado. "
    If cnpjCpf           = ""            Then erros = erros & "CPF/CNPJ não informado. "
    If placa             = ""            Then erros = erros & "Placa não informada. "
    If dataPreenchimento = ""            Then erros = erros & "Data de Preenchimento não informada. "
    If idTipoVeiculo     = "" _
       Or idTipoVeiculo  = "0"           Then erros = erros & "Tipo de Veículo não informado. "
    If quantidadeEixo    = ""            Then erros = erros & "Quantidade de Eixos não informada. "
    If horarioInicio     = ""            Then erros = erros & "Horário de Início não informado. "
    If destinoCarga      = ""            Then erros = erros & "Destino da Carga não informado. "
    If altura            = ""            Then erros = erros & "Altura não informada. "
    If comprimento       = ""            Then erros = erros & "Comprimento não informado. "
    If largura           = ""            Then erros = erros & "Largura não informada. "
    If nomeConferente    = ""            Then erros = erros & "Nome do Conferente não informado. "
    If status            = "" _
       Or status          = "0"          Then erros = erros & "Status não informado. "

    If erros <> "" Then
        Salvar = "{""sucesso"": false, ""mensagem"": """ & EscapeJSON(Trim(erros)) & """}"
        Exit Function
    End If

    ' ── Validação de formato de horário (HH:MM) ──────────────────────
    Dim reHora
    Set reHora = New RegExp
    reHora.Pattern = "^\d{2}:\d{2}$"

    If Not reHora.Test(horarioInicio) Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Horário de Início inválido. Use o formato HH:MM.""}"
        Exit Function
    End If

    If horarioTermino <> "" And Not reHora.Test(horarioTermino) Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Horário de Término inválido. Use o formato HH:MM.""}"
        Exit Function
    End If

    Dim partesData, dataSQL
    partesData = Split(dataPreenchimento, "/")
    If UBound(partesData) = 2 Then
        dataSQL = partesData(2) & "-" & partesData(1) & "-" & partesData(0)
    Else
        Salvar = "{""sucesso"": false, ""mensagem"": ""Formato de data inválido.""}"
        Exit Function
    End If

    nomeProprietario    = Replace(nomeProprietario,    "'", "''")
    telefone            = Replace(telefone,            "'", "''")
    nomeMotorista       = Replace(nomeMotorista,       "'", "''")
    cnpjCpf             = Replace(cnpjCpf,             "'", "''")
    placa               = Replace(placa,               "'", "''")
    horarioInicio       = Replace(horarioInicio,       "'", "''")
    horarioTermino      = Replace(horarioTermino,      "'", "''")
    destinoCarga        = Replace(destinoCarga,        "'", "''")
    nomeConferente      = Replace(nomeConferente,      "'", "''")
    portariaResponsavel = Replace(portariaResponsavel, "'", "''")
    status              = Replace(status,              "'", "''")

    altura      = Replace(altura,      ",", ".")
    comprimento = Replace(comprimento, ",", ".")
    largura     = Replace(largura,     ",", ".")

    Dim sqlInsert
    sqlInsert = "INSERT INTO CHECKLIST_TRANSPORTES ("
    sqlInsert = sqlInsert & "  NOMEPROPRIETARIO, TELEFONE, NOMEMOTORISTA, DOCUMENTO, "
    sqlInsert = sqlInsert & "  PLACAVEICULO, DATAPREENCHIMENTO, ID_TIPOVEICULO, QTDEEIXO, "
    sqlInsert = sqlInsert & "  HORARIOINICIO, HORARIOTERMINO, ID_TIPOTRANSPORTADORA, DESTINOCARGA, "
    sqlInsert = sqlInsert & "  ALTURA, COMPRIMENTO, LARGURA, MANUTENCAOVEICULO, "
    sqlInsert = sqlInsert & "  CONFERENTE, RESPONSAVELPORTARIA, STATUS, ATIVO"
    sqlInsert = sqlInsert & ") OUTPUT INSERTED.ID VALUES ("
    sqlInsert = sqlInsert & "  '" & nomeProprietario    & "', "
    sqlInsert = sqlInsert & "  '" & telefone            & "', "
    sqlInsert = sqlInsert & "  '" & nomeMotorista       & "', "
    sqlInsert = sqlInsert & "  '" & cnpjCpf             & "', "
    sqlInsert = sqlInsert & "  '" & placa               & "', "
    sqlInsert = sqlInsert & "  '" & dataSQL             & "', "
    sqlInsert = sqlInsert & "  "  & idTipoVeiculo       & ", "
    sqlInsert = sqlInsert & "  "  & quantidadeEixo      & ", "
    sqlInsert = sqlInsert & "  '" & horarioInicio       & "', "
    sqlInsert = sqlInsert & "  '" & horarioTermino      & "', "
    sqlInsert = sqlInsert & "  "  & tipoTransporte      & ", "
    sqlInsert = sqlInsert & "  '" & destinoCarga        & "', "
    sqlInsert = sqlInsert & "  "  & altura              & ", "
    sqlInsert = sqlInsert & "  "  & comprimento         & ", "
    sqlInsert = sqlInsert & "  "  & largura             & ", "
    sqlInsert = sqlInsert & "  "  & manutencao          & ", "
    sqlInsert = sqlInsert & "  '" & nomeConferente      & "', "
    sqlInsert = sqlInsert & "  '" & portariaResponsavel & "', "
    sqlInsert = sqlInsert & "  '" & status              & "', "
    sqlInsert = sqlInsert & "  1"
    sqlInsert = sqlInsert & ")"

    On Error Resume Next
    Dim rsInsert
    Set rsInsert = Conecta("PIR-CORPORATIVO", sqlInsert)
    If Err.Number <> 0 Or rsInsert Is Nothing Then
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro ao gravar checklist: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error GoTo 0

    If rsInsert.EOF Then
        rsInsert.Close()
        Set rsInsert = Nothing
        Salvar = "{""sucesso"": false, ""mensagem"": ""Erro inesperado: registro não foi inserido.""}"
        Exit Function
    End If

    Dim idNovoChecklist
    idNovoChecklist = rsInsert("ID") & ""
    rsInsert.Close()
    Set rsInsert = Nothing

    ' Insere apenas os itens marcados como "Atende" (pipe ou vírgula-separated vindo do front)
    If itensAvaliados <> "" Then
        Dim arrItensNovos, itemIDNovo, i
        ' Front envia vírgula-separated: "37,38,39"
        arrItensNovos = Split(itensAvaliados, ",")
        For i = 0 To UBound(arrItensNovos)
            itemIDNovo = Trim(arrItensNovos(i))
            If IsNumeric(itemIDNovo) And CLng(itemIDNovo) > 0 Then
                Dim sqlItemNovo
                sqlItemNovo = "INSERT INTO CHECKLIST_TRANSPORTESITENS " & _
                              "(ID_CHEKLIST_TRANSPORTE, ID_CHEKLIST_ITENSAVALIADOS) " & _
                              "VALUES (" & CLng(idNovoChecklist) & ", " & CLng(itemIDNovo) & ")"
                On Error Resume Next
                Dim rsItem
                Set rsItem = Conecta("PIR-CORPORATIVO", sqlItemNovo)
                If Err.Number = 0 And Not rsItem Is Nothing Then rsItem.Close() : Set rsItem = Nothing
                On Error GoTo 0
            End If
        Next
    End If

    Call RegistrarLog(idNovoChecklist)

    Dim jsonRetorno
    jsonRetorno = "{"
    jsonRetorno = jsonRetorno & """sucesso"": true, "
    jsonRetorno = jsonRetorno & """mensagem"": ""CheckList gravado com sucesso."", "
    jsonRetorno = jsonRetorno & """dados"": {""id"": " & idNovoChecklist & "}"
    jsonRetorno = jsonRetorno & "}"

    Salvar = jsonRetorno
End Function

Sub RegistrarLog(idRef)
    Dim descricao, sqlLog
    descricao = "Usuário " & Session("USR") & " gravou o checklist de transporte id " & idRef
    descricao = Replace(descricao, "'", "''")

    sqlLog = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    sqlLog = sqlLog & Session("CODIGO") & ", "
    sqlLog = sqlLog & "'CheckListTransporte', 'SWNET', "
    sqlLog = sqlLog & "'criacao_checklist', "
    sqlLog = sqlLog & "'" & descricao & "', "
    sqlLog = sqlLog & "GETDATE())"

    On Error Resume Next
    Dim rsLog
    Set rsLog = Conecta("PIR-CORPORATIVO", sqlLog)
    If Err.Number = 0 And Not rsLog Is Nothing Then rsLog.Close() : Set rsLog = Nothing
    On Error GoTo 0
End Sub
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — Novo CheckList Transporte</title>

  <link href="css/bootstrap.min.css" rel="stylesheet" />
  <link href="css/bootstrap-icons.min.css" rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    .form-section {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
      margin-bottom: 1.25rem;
    }
    .form-section-header {
      background-color: #cfcfcf;
      color: #050000;
      padding: .6rem 1rem;
      border-radius: 6px 6px 0 0;
      font-weight: 600;
      font-size: .88rem;
      display: flex;
      align-items: center;
      gap: .5rem;
    }
    .form-section-body { padding: 1.25rem; }

    .form-label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    .table-itens thead th {
      background-color: #909090; color: #fff;
      font-size: .8rem; font-weight: 600;
      text-align: center; padding: .5rem .75rem;
    }
    .table-itens tbody td {
      font-size: .82rem; vertical-align: middle; padding: .4rem .75rem;
    }
    .table-itens tbody tr:hover { background-color: #f8f9fa; }

    .notas-lista { font-size: .82rem; color: #495057; line-height: 1.6; }
    .notas-lista li { margin-bottom: .35rem; }

    .radio-opcao { cursor: pointer; display: flex; align-items: center; gap: .4rem; }
    .radio-opcao input[type="radio"] { cursor: pointer; }

    /* Overlay */
    #overlayCarregando {
      position: fixed; inset: 0;
      background: rgba(255,255,255,.75);
      display: flex; align-items: center; justify-content: center;
      z-index: 1200;
    }
    #overlayCarregando.d-none { display: none !important; }

    /* Skeleton */
    .skeleton-bar {
      height: 14px; border-radius: 4px;
      background: linear-gradient(90deg, #e9ecef 25%, #dee2e6 50%, #e9ecef 75%);
      background-size: 200% 100%;
      animation: shimmer 1.2s infinite;
    }
    @keyframes shimmer { to { background-position: -200% 0; } }

    /* ── Toast canônico (igual ao CheckListTransporteGrade.asp) ── */
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
  </style>
</head>
<body>

  <!-- Overlay de carregamento inicial -->
  <div id="overlayCarregando">
    <div class="text-center">
      <div class="spinner-border text-secondary mb-2" role="status"></div>
      <div style="font-size:.85rem;color:#6c757d;">Carregando formulário...</div>
    </div>
  </div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <!-- Cabeçalho -->
    <div class="d-flex align-items-center justify-content-between mb-3">
      <span class="page-title">Novo CheckList Transporte</span>
      <button type="button" class="btn btn-outline-secondary btn-sm" id="btnVoltar">
        <i class="bi bi-arrow-left me-2"></i>Voltar
      </button>
    </div>

    <div id="frmChecklist">

      <!-- Seção: Dados do Veículo / Proprietário -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-truck me-2"></i>Dados do Veículo / Proprietário
        </div>
        <div class="form-section-body">

          <div class="row g-3">
            <div class="col-md-6">
              <label for="txtNomeProprietario" class="form-label">
                Nome do Proprietário <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtNomeProprietario" maxlength="150" />
            </div>
            <div class="col-md-3">
              <label for="txtTelefone" class="form-label">Telefone</label>
              <input type="text" class="form-control form-control-sm"
                     id="txtTelefone" maxlength="15" />
            </div>
          </div>

          <div class="row g-3 mt-1">
            <div class="col-md-6">
              <label for="txtNomeMotorista" class="form-label">
                Nome do Motorista <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtNomeMotorista" maxlength="150" />
            </div>
          </div>

          <div class="row g-3 mt-1">
            <div class="col-md-3">
              <label for="txtCnpj" class="form-label">
                CPF / CNPJ <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtCnpj" maxlength="18" />
            </div>
            <div class="col-md-3">
              <label for="txtPlaca" class="form-label">
                Placa do Veículo <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtPlaca" maxlength="8" />
            </div>
            <div class="col-md-3">
              <label for="txtDataPreenchimento" class="form-label">
                Data de Preenchimento <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtDataPreenchimento" maxlength="10"
                     placeholder="DD/MM/AAAA" readonly />
            </div>
          </div>

          <div class="row g-3 mt-1">
            <div class="col-md-5">
              <label for="cboTipoVeiculo" class="form-label">
                Tipo de Veículo <span class="text-danger">*</span>
              </label>
              <select id="cboTipoVeiculo" class="form-select form-select-sm">
                <option value="0">Carregando...</option>
              </select>
            </div>
            <div class="col-md-2">
              <label for="txtQuantidadeEixo" class="form-label">
                Qtd. Eixos <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtQuantidadeEixo" maxlength="5" />
            </div>
          </div>

          <div class="row g-3 mt-1">
            <div class="col-md-3">
              <label for="txtHorarioInicio" class="form-label">
                Horário Início Carregamento <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm"
                     id="txtHorarioInicio" maxlength="5" placeholder="HH:MM" />
            </div>
            <div class="col-md-3">
              <label for="txtHorarioTermino" class="form-label">Horário Término Carregamento</label>
              <input type="text" class="form-control form-control-sm"
                     id="txtHorarioTermino" maxlength="5" placeholder="HH:MM" />
            </div>
          </div>

        </div>
      </div>

      <!-- Seção: Tipo de Transporte -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-signpost-split me-2"></i>Tipo de Transporte
        </div>
        <div class="form-section-body">

          <div class="d-flex flex-wrap gap-4">
            <label class="radio-opcao">
              <input type="radio" name="rdTipoTransporte" value="1" />
              Smurfit Westrock
            </label>
            <label class="radio-opcao">
              <input type="radio" name="rdTipoTransporte" value="2" checked />
              Transportadora
            </label>
            <label class="radio-opcao">
              <input type="radio" name="rdTipoTransporte" value="3" />
              Freteiros
            </label>
          </div>

          <div class="row g-3 mt-2">
            <div class="col-md-6">
              <label for="txtDestinoCarga" class="form-label">
                Destino da Carga <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm" id="txtDestinoCarga" />
            </div>
          </div>

        </div>
      </div>

      <!-- Seção: Itens Avaliados -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-list-check me-2"></i>Itens Avaliados
        </div>
        <div class="form-section-body p-0">
          <div class="table-responsive">
            <table class="table table-bordered table-sm mb-0 table-itens">
              <thead>
                <tr>
                  <th class="text-start" style="width:70%">Item Avaliado</th>
                  <th>Atende</th>
                  <th>Não Atende</th>
                </tr>
              </thead>
              <tbody id="tbodyItens">
                <tr class="skeleton-row">
                  <td><div class="skeleton-bar" style="width:80%"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                </tr>
                <tr class="skeleton-row">
                  <td><div class="skeleton-bar" style="width:62%"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                </tr>
                <tr class="skeleton-row">
                  <td><div class="skeleton-bar" style="width:71%"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                  <td class="text-center"><div class="skeleton-bar mx-auto" style="width:20px"></div></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Seção: Dimensões do Veículo -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-rulers me-2"></i>Dimensões do Veículo
        </div>
        <div class="form-section-body">

          <div class="row g-3">
            <div class="col-md-3">
              <label for="txtAltura" class="form-label">Altura (m) <span class="text-danger">*</span></label>
              <input type="number" class="form-control form-control-sm" id="txtAltura" step=".01" />
            </div>
            <div class="col-md-3">
              <label for="txtComprimento" class="form-label">Comprimento (m) <span class="text-danger">*</span></label>
              <input type="number" class="form-control form-control-sm" id="txtComprimento" step=".01" />
            </div>
            <div class="col-md-3">
              <label for="txtLargura" class="form-label">Largura (m) <span class="text-danger">*</span></label>
              <input type="number" class="form-control form-control-sm" id="txtLargura" step=".01" />
            </div>
          </div>

          <div class="mt-3">
            <label class="form-label d-block">
              Manutenção do veículo em dia? <span class="text-danger">*</span>
            </label>
            <div class="d-flex gap-4">
              <label class="radio-opcao">
                <input type="radio" name="rdManutencao" value="1" checked />
                <strong>Sim</strong>
              </label>
              <label class="radio-opcao">
                <input type="radio" name="rdManutencao" value="0" />
                <strong>Não</strong>
              </label>
            </div>
          </div>

        </div>
      </div>

      <!-- Seção: Conferência -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-person-check me-2"></i>Conferência
        </div>
        <div class="form-section-body">

          <div class="row g-3">
            <div class="col-md-6">
              <label for="txtNomeConferente" class="form-label">
                Nome do Conferente <span class="text-danger">*</span>
              </label>
              <input type="text" class="form-control form-control-sm" id="txtNomeConferente" />
            </div>
            <div class="col-md-6">
              <label for="txtPortariaResponsavel" class="form-label">Responsável na Portaria</label>
              <input type="text" class="form-control form-control-sm" id="txtPortariaResponsavel" />
            </div>
          </div>

          <div class="row g-3 mt-1">
            <div class="col-md-3">
              <label for="cboStatus" class="form-label">Status <span class="text-danger">*</span></label>
              <select id="cboStatus" class="form-select form-select-sm">
                <option value="0">Selecione...</option>
                <option value="PATIO HJ">Pátio HJ</option>
                <option value="PORTARIA">Portaria</option>
                <option value="EXPEDIÇÃO">Expedição</option>
                <option value="FINALIZADO">Finalizado</option>
              </select>
            </div>
          </div>

          <p class="mt-3 mb-0" style="font-size:.8rem;color:#6c757d;">
            Este formulário deverá ser preenchido corretamente pelos responsáveis do transporte
            da Smurfit Westrock conforme as notas abaixo.
          </p>

        </div>
      </div>

      <!-- Seção: Notas -->
      <div class="form-section">
        <div class="form-section-header">
          <i class="bi bi-info-circle me-2"></i>Notas
        </div>
        <div class="form-section-body">
          <ol class="notas-lista mb-0">
            <li>O não cumprimento a qualquer um dos itens de avaliação exclui o motorista de carregar o caminhão.</li>
            <li>Será impossibilitado de viajar o motorista que apresentar sintomas de embriaguez ou estado físico alterado.</li>
            <li>Também será excluído do cadastro o motorista que causar qualquer tipo de problemas com os clientes da Smurfit Westrock.</li>
          </ol>
          <p class="mt-3 mb-0 text-muted" style="font-size:.75rem;">
            RGQ105rev11 · Edição original: 26/09/2023 · Revisão: 18/12/2023
          </p>
        </div>
      </div>

      <!-- Ações -->
      <div class="d-flex gap-2 mb-4">
        <button type="button" class="btn btn-primary btn-sm" id="btnGravar">
          Gravar
        </button>
        <button type="button" class="btn btn-outline-secondary btn-sm" id="btnCancelar">
          Cancelar
        </button>
      </div>

    </div><!-- /#frmChecklist -->
  </div><!-- /.container-fluid -->

  <!-- Toast container -->
  <div id="toastContainer" class="toast-container-custom"></div>

  <!-- Template de linha de item avaliado -->
  <template id="tmplItemAvaliado">
    <tr>
      <td class="text-start item-descricao"></td>
      <td class="text-center">
        <input type="radio" class="form-check-input radio-sim" value="" checked />
      </td>
      <td class="text-center">
        <input type="radio" class="form-check-input radio-nao" value="0" />
      </td>
    </tr>
  </template>

   <script src="js/jquery.min.js"></script>
   <script src="js/bootstrap.bundle.min.js"></script>

  <script>
    'use strict';

    var CheckListTransporte = (function () {

      /* ── Constantes ── */
      const ENDPOINT    = 'CheckListTransporte.asp';
      const TOAST_DELAY = 4500;
      /* ── Toast canônico ── */
      const TOAST_ICON = {
        sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754' },
        erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545' },
        aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404' },
        info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d' }
      };

      /* ── Toast ── */
      function exibirToast(tipo, mensagem) {
        const t      = TOAST_ICON[tipo] || TOAST_ICON.info;
        const barCor = tipo === 'aviso' ? '#ffc107' : t.cor;
        const el = $(
          '<div class="toast-custom p-3">' +
            '<div class="d-flex align-items-start gap-2">' +
              '<i class="bi ' + t.icon + ' mt-1" style="color:' + t.cor + ';font-size:1.1rem;"></i>' +
              '<div class="flex-grow-1">' +
                '<div style="font-size:.85rem;font-weight:600;color:' + t.cor + ';">' + mensagem + '</div>' +
              '</div>' +
              '<button class="btn-close btn-close-sm ms-auto" style="font-size:.7rem;" ' +
                      'onclick="this.closest(\'.toast-custom\').remove()"></button>' +
            '</div>' +
            '<div class="mt-2" style="height:3px;background:#e9ecef;border-radius:2px;">' +
              '<div class="toast-progress-bar" ' +
                   'style="height:100%;background:' + barCor + ';border-radius:2px;"></div>' +
            '</div>' +
          '</div>'
        );
        $('#toastContainer').append(el);
        setTimeout(function () { el.remove(); }, TOAST_DELAY);
      }

      /* ── Overlay ── */
      function mostrarOverlay() { $('#overlayCarregando').removeClass('d-none'); }
      function ocultarOverlay() { $('#overlayCarregando').addClass('d-none');    }

      /* ── Navegação ── */
      function volta() { window.location.href = 'CheckListTransporteGrade.asp'; }

      /* ── Data automática ── */
      function preencherData() {
        const hoje = new Date();
        const dd   = String(hoje.getDate()).padStart(2, '0');
        const mm   = String(hoje.getMonth() + 1).padStart(2, '0');
        $('#txtDataPreenchimento').val(dd + '/' + mm + '/' + hoje.getFullYear());
      }

      /* ── Máscara horário ── */
      function aplicarMascaraHorario(seletor) {
        $(seletor).on('input', function () {
          let v = $(this).val().replace(/\D/g, '').substring(0, 4);
          if (v.length >= 3) v = v.substring(0, 2) + ':' + v.substring(2);
          $(this).val(v);
        });
      }

      /* ── Máscara telefone ── */
      function mphone(v) {
        let r = v.replace(/\D/g, '').replace(/^0/, '');
        if      (r.length > 10) r = r.replace(/^(\d{2})(\d{5})(\d{4}).*/, '($1) $2-$3');
        else if (r.length > 5)  r = r.replace(/^(\d{2})(\d{4})(\d{0,4}).*/, '($1) $2-$3');
        else if (r.length > 2)  r = r.replace(/^(\d{2})(\d{0,5})/, '($1) $2');
        else                    r = r.replace(/^(\d*)/, '($1');
        return r;
      }

      /* ── Máscara CPF/CNPJ ── */
      function formatarDocumento(campo) {
        let v = $(campo).val().replace(/\D/g, '');
        if (v.length <= 11) v = v.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
        else                v = v.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
        $(campo).val(v);
      }

      /* ── Carregar dados iniciais ── */
      function carregarDadosIniciais() {
        $.ajax({
          url    : ENDPOINT,
          method : 'GET',
          data   : { acao: 'obterDadosIniciais' },
          success: function (resp) {
            if (!resp.sucesso) {
              exibirToast('erro', resp.mensagem || 'Erro ao carregar formulário.');
              ocultarOverlay();
              return;
            }

            const $cbo = $('#cboTipoVeiculo').empty();
            $cbo.append('<option value="0">Selecione...</option>');
            $.each(resp.dados.tiposVeiculo, function (_, tv) {
              $cbo.append('<option value="' + tv.id + '">' + tv.descricao + '</option>');
            });

            const $tbody = $('#tbodyItens').empty();
            const tmpl   = document.getElementById('tmplItemAvaliado');
            $.each(resp.dados.itensAvaliados, function (_, item) {
              const clone = tmpl.content.cloneNode(true);
              $(clone).find('.item-descricao').text(item.descricao);
              $(clone).find('.radio-sim').val(item.id).attr('name', 'marca' + item.id);
              $(clone).find('.radio-nao').attr('name', 'marca' + item.id);
              $tbody.append(clone);
            });

            ocultarOverlay();
          },
          error: function () {
            exibirToast('erro', 'Não foi possível carregar o formulário.');
            ocultarOverlay();
          }
        });
      }

      /* ── Buscar dados pela placa ── */
      function buscarPlaca(placa) {
        $.ajax({
          url    : ENDPOINT,
          method : 'GET',
          data   : { acao: 'buscarPlaca', placa: placa },
          success: function (resp) {
            if (!resp.sucesso) {
              exibirToast('aviso', resp.mensagem || 'Placa não encontrada.');
              return;
            }
            const d = resp.dados;
            $('#txtNomeProprietario').val(d.nomeProprietario);
            $('#txtTelefone').val(d.telefone);
            $('#txtCnpj').val(d.cnpjCpf);
            $('#txtNomeMotorista').val(d.nomeMotorista);
            $('#cboTipoVeiculo').val(d.idTipoVeiculo);
            $('#txtQuantidadeEixo').val(d.quantidadeEixo);
            $('#txtAltura').val(d.altura);
            $('#txtComprimento').val(d.comprimento);
            $('#txtLargura').val(d.largura);
          },
          error: function () {
            exibirToast('aviso', 'Não foi possível buscar dados da placa.');
          }
        });
      }

      /* ── Coletar IDs dos itens que ATENDEM ── */
      function coletarItensAtende() {
        const ids = [];
        $('#tbodyItens input.radio-sim:checked').each(function () {
          ids.push($(this).val());
        });
        return ids.join(',');
      }

      /* ── Validação client-side ── */
      function validar() {
        const erros = [];
        if (!$('#txtNomeProprietario').val().trim())  erros.push('Nome do Proprietário não informado');
        if (!$('#txtNomeMotorista').val().trim())     erros.push('Nome do Motorista não informado');
        if (!$('#txtCnpj').val().trim())              erros.push('CPF/CNPJ não informado');
        if (!$('#txtPlaca').val().trim())             erros.push('Placa não informada');
        if (!$('#txtDataPreenchimento').val().trim()) erros.push('Data de Preenchimento não informada');
        if ($('#cboTipoVeiculo').val() === '0')       erros.push('Tipo de Veículo não informado');
        if (!$('#txtQuantidadeEixo').val().trim())    erros.push('Quantidade de Eixos não informada');
        if (!$('#txtHorarioInicio').val().trim())     erros.push('Horário de Início não informado');
        if (!$('#txtDestinoCarga').val().trim())      erros.push('Destino da Carga não informado');
        if (!$('#txtAltura').val().trim())            erros.push('Altura não informada');
        if (!$('#txtComprimento').val().trim())       erros.push('Comprimento não informado');
        if (!$('#txtLargura').val().trim())           erros.push('Largura não informada');
        if (!$('#txtNomeConferente').val().trim())    erros.push('Nome do Conferente não informado');
        if ($('#cboStatus').val() === '0')            erros.push('Status não informado');
        var reHora = /^\d{2}:\d{2}$/;
        if (!reHora.test($('#txtHorarioInicio').val().trim()))   erros.push('Horário de Início inválido (use HH:MM)');
        var hTermino = $('#txtHorarioTermino').val().trim();
        if (hTermino !== '' && !reHora.test(hTermino))           erros.push('Horário de Término inválido (use HH:MM)');
        return erros;
      }

      /* ── Gravar via AJAX ── */
      function gravar() {
        const erros = validar();
        if (erros.length > 0) {
          exibirToast('aviso', erros.join('<br>'));
          return;
        }
        if (!confirm('Confirmar gravação deste CheckList?')) return;

        mostrarOverlay();

        $.ajax({
          url    : ENDPOINT,
          method : 'POST',
          data   : {
            acao               : 'salvar',
            nomeProprietario   : $('#txtNomeProprietario').val().trim(),
            telefone           : $('#txtTelefone').val().trim(),
            nomeMotorista      : $('#txtNomeMotorista').val().trim(),
            cnpjCpf            : $('#txtCnpj').val().trim(),
            placa              : $('#txtPlaca').val().trim(),
            dataPreenchimento  : $('#txtDataPreenchimento').val().trim(),
            idTipoVeiculo      : $('#cboTipoVeiculo').val(),
            quantidadeEixo     : $('#txtQuantidadeEixo').val().trim(),
            horarioInicio      : $('#txtHorarioInicio').val().trim(),
            horarioTermino     : $('#txtHorarioTermino').val().trim(),
            tipoTransporte     : $('input[name="rdTipoTransporte"]:checked').val(),
            destinoCarga       : $('#txtDestinoCarga').val().trim(),
            altura             : $('#txtAltura').val().trim(),
            comprimento        : $('#txtComprimento').val().trim(),
            largura            : $('#txtLargura').val().trim(),
            manutencao         : $('input[name="rdManutencao"]:checked').val(),
            nomeConferente     : $('#txtNomeConferente').val().trim(),
            portariaResponsavel: $('#txtPortariaResponsavel').val().trim(),
            status             : $('#cboStatus').val(),
            itensAvaliados     : coletarItensAtende()
          },
          success: function (resp) {
            ocultarOverlay();
            if (resp.sucesso) {
              exibirToast('sucesso', resp.mensagem || 'CheckList gravado com sucesso.');
              setTimeout(function () { volta(); }, 1800);
            } else {
              exibirToast('erro', resp.mensagem || 'Erro ao gravar.');
            }
          },
          error: function () {
            ocultarOverlay();
            exibirToast('erro', 'Erro de comunicação com o servidor.');
          }
        });
      }

      /* ── Init ── */
      function init() {
        preencherData();
        aplicarMascaraHorario('#txtHorarioInicio');
        aplicarMascaraHorario('#txtHorarioTermino');

        $('#txtHorarioInicio').on('blur', function () {
          var v = $(this).val().trim();
          if (v !== '' && !/^\d{2}:\d{2}$/.test(v)) {
            exibirToast('aviso', 'Horário de Início inválido. Use o formato HH:MM.');
            $(this).val('').trigger('focus');
          }
        });

        $('#txtHorarioTermino').on('blur', function () {
          var v = $(this).val().trim();
          if (v !== '' && !/^\d{2}:\d{2}$/.test(v)) {
            exibirToast('aviso', 'Horário de Término inválido. Use o formato HH:MM.');
            $(this).val('');
          }
        });

        $('#txtTelefone').on('input', function () {
          $(this).val(mphone($(this).val()));
        });

        $('#txtCnpj')
          .on('focus', function () { $(this).val($(this).val().replace(/[.\-\/]/g, '')); })
          .on('blur',  function () { formatarDocumento(this); });

        $('#txtQuantidadeEixo').on('keypress', function (e) {
          const c = e.charCode || e.keyCode;
          if (c !== 8 && c !== 9 && (c < 48 || c > 57)) return false;
        });

        $('#txtPlaca').on('blur', function () {
          const placa = $(this).val().trim();
          if (placa !== '') buscarPlaca(placa);
        });

        $('#btnGravar').on('click',   function () { gravar(); });
        $('#btnCancelar').on('click', function () { volta();  });
        $('#btnVoltar').on('click',   function () { volta();  });

        carregarDadosIniciais();
      }

      return Object.freeze({ init: init });

    }());

    $(function () { CheckListTransporte.init(); });
  </script>
</body>
</html>
