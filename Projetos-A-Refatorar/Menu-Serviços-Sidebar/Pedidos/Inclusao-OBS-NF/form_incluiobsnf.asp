<!-- #include file="../cnn/cnn.asp" -->
<!-- #include file="../includes/variaveis.asp" -->
<!-- #include file="../funcoes/funcoes.asp" -->
<%
Response.Buffer  = True
Response.Expires = 0
Session.LCID     = 1046

On Error GoTo 0
On Error Resume Next

' ------------------------------------------------------------------
' Funções auxiliares
' ------------------------------------------------------------------
Function dataYYYYMMDD(data)
  Dim dia, mes, ano
  dia = Day(data)
  mes = Month(data)
  ano = Year(data)
  If CInt(dia) < 10 Then dia = "0" & dia
  If CInt(mes) < 10 Then mes = "0" & mes
  dataYYYYMMDD = ano & mes & dia
End Function

' ------------------------------------------------------------------
' Variáveis de controle
' ------------------------------------------------------------------
Dim buscar, salvar, edtPEDIDO
Dim edtOBS1, edtOBS2, edtOBS3, edtOBS4
Dim toastMsg, toastTipo

buscar    = Request("buscar")
salvar    = Request("salvar")
edtPEDIDO = Trim(Request("edtPEDIDO"))
toastMsg  = ""
toastTipo = ""

' ------------------------------------------------------------------
' Buscar pedido
' ------------------------------------------------------------------
If buscar = "S" Then

  Dim SQL
  SQL = "" &_
    "SELECT " &_
      "C5_NUM, " &_
      "C5_MENNOTA, " &_
      "C5_MENNOT2, " &_
      "C5_MENNOT3, " &_
      "C5_MENNOT4 " &_
    "FROM SC5010 " &_
    "WHERE C5_NUM = '" & edtPEDIDO & "'"

  Dados.Open(SQL)

  If Err.Number <> 0 Then
    toastMsg  = "Não foi possível pesquisar o pedido. Tente mais tarde."
    toastTipo = "erro"
  Else
    If Not Dados.EOF Then
      edtOBS1 = Trim(Dados("C5_MENNOTA"))
      edtOBS2 = Trim(Dados("C5_MENNOT2"))
      edtOBS3 = Trim(Dados("C5_MENNOT3"))
      edtOBS4 = Trim(Dados("C5_MENNOT4"))
    Else
      toastMsg  = "Número de pedido não encontrado."
      toastTipo = "aviso"
    End If
  End If

  Dados.Close()

End If

' ------------------------------------------------------------------
' Salvar observações
' ------------------------------------------------------------------
If salvar = "S" And edtPEDIDO <> "" Then

  Dim dataHoje, horaAgora
  dataHoje  = Date()
  horaAgora = Time()

  SQL = "" &_
    "UPDATE SC5010 SET " &_
      "C5_MENNOTA = '" & Trim(Request("edtOBS1")) & "', " &_
      "C5_MENNOT2 = '" & Trim(Request("edtOBS2")) & "', " &_
      "C5_MENNOT3 = '" & Trim(Request("edtOBS3")) & "', " &_
      "C5_MENNOT4 = '" & Trim(Request("edtOBS4")) & "' " &_
    "WHERE C5_NUM = '" & edtPEDIDO & "'"

  Dados.Open(SQL)

  If Err.Number <> 0 Then
    toastMsg  = "Não foi possível salvar as observações. Tente mais tarde."
    toastTipo = "erro"
  Else

    Dim camposLog(3), valoresLog(3), i
    camposLog(0) = "c5_mennota" : valoresLog(0) = Trim(Request("edtOBS1"))
    camposLog(1) = "c5_mennot2" : valoresLog(1) = Trim(Request("edtOBS2"))
    camposLog(2) = "c5_mennot3" : valoresLog(2) = Trim(Request("edtOBS3"))
    camposLog(3) = "c5_mennot4" : valoresLog(3) = Trim(Request("edtOBS4"))

    Dim logOk : logOk = True
    For i = 0 To 3
      SQL = "INSERT INTO InHistoricoAlteracoesPedidos " &_
            "(Pedido, Para, Login, DataAlteracao, HoraAlteracao) " &_
            "VALUES (" &_
              "'" & edtPEDIDO & "', " &_
              "'" & camposLog(i) & " = " & valoresLog(i) & "', " &_
              "'" & Session("USR") & "', " &_
              "'" & dataYYYYMMDD(dataHoje) & "', " &_
              "'" & horaAgora & "')"
      Dados21.Open(SQL)
      If Err.Number <> 0 Then
        toastMsg  = "Observações salvas, mas houve erro ao registrar o histórico."
        toastTipo = "aviso"
        logOk = False
        Exit For
      End If
    Next

    If logOk Then
      toastMsg  = "Observações gravadas com sucesso!"
      toastTipo = "sucesso"
      edtPEDIDO = ""
      edtOBS1   = ""
      edtOBS2   = ""
      edtOBS3   = ""
      edtOBS4   = ""
    End If

  End If

End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>SWnet — Observação na Nota Fiscal</title>

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <style>
    body { background-color: #f0f0f0; min-height: 100vh; }

    /* Page title */
    .page-title { font-size: 1.15rem; font-weight: 700; color: #212529; }

    /* Filter card */
    .filter-card        { background-color: #fff; border: 1px solid #dee2e6; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); max-width: 820px; }
    .filter-card-header { background-color: #fff; color: #212529; padding: .6rem 1rem; border-radius: 6px 6px 0 0; border-bottom: 1px solid #dee2e6; font-weight: 600; font-size: .9rem; }
    .filter-card-body   { padding: 1rem 1.25rem 1.25rem; }

    /* Labels */
    .filter-card-body .form-label { font-size: .82rem; font-weight: 600; color: #6c757d; margin-bottom: .25rem; }

    /* Toast */
    .toast-container-custom { position: fixed; top: 1rem; right: 1rem; z-index: 1100; display: flex; flex-direction: column; gap: .5rem; }
    .toast-custom { min-width: 280px; max-width: 360px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 18px rgba(0,0,0,.13); background: #fff; border: 1px solid #dee2e6; animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1); }
    @keyframes toastSlideIn { from { opacity: 0; transform: translateX(40px); } to { opacity: 1; transform: translateX(0); } }
    .toast-custom .toast-progress-bar { animation: toastProgressShrink 4.5s linear forwards; }
    @keyframes toastProgressShrink { from { width: 100%; } to { width: 0%; } }
  </style>
</head>
<body>

  <!-- Toast container -->
  <div class="toast-container-custom" id="toastContainer"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="mb-3 mx-auto" style="max-width: 820px;">
      <span class="page-title">Observação na Nota Fiscal</span>
    </div>

    <div class="filter-card mx-auto mb-0">
      <div>

      <div class="filter-card-header">
        Dados do Pedido
      </div>

      <div class="filter-card-body">

        <form id="formObs" name="formObs" method="post" action="form_incluiobsnf.asp" novalidate>

          <input type="hidden" name="buscar" id="buscar" value="" />
          <input type="hidden" name="salvar" id="salvar" value="" />

          <!-- Número do pedido + botão buscar -->
          <div class="row g-2 mb-3">
            <div class="col-sm-4">
              <label for="edtPEDIDO" class="form-label">Número do Pedido</label>
              <div class="input-group input-group-sm">
                <input id="edtPEDIDO" name="edtPEDIDO" type="text"
                       class="form-control"
                       placeholder="Ex.: 000123"
                       value="<%=edtPEDIDO%>"
                       maxlength="6"
                       autocomplete="off" />
                <button type="button" class="btn btn-primary btn-sm" onclick="Buscar()">
                  <i class="bi bi-search"></i>
                </button>
              </div>
            </div>
          </div>

          <hr class="my-3" />

          <!-- Observações -->
          <div class="row g-2 mb-2">
            <div class="col-12">
              <label for="edtOBS1" class="form-label">Observação 1 — Constar na NF (60 pos.)</label>
              <input id="edtOBS1" name="edtOBS1" type="text"
                     class="form-control form-control-sm"
                     value="<%=edtOBS1%>"
                     maxlength="60"
                     autocomplete="off" />
            </div>
          </div>

          <div class="row g-2 mb-2">
            <div class="col-12">
              <label for="edtOBS2" class="form-label">Observação 2 — Constar na NF (130 pos.)</label>
              <input id="edtOBS2" name="edtOBS2" type="text"
                     class="form-control form-control-sm"
                     value="<%=edtOBS2%>"
                     maxlength="130"
                     autocomplete="off" />
            </div>
          </div>

          <div class="row g-2 mb-2">
            <div class="col-12">
              <label for="edtOBS3" class="form-label">Observação 3 — Constar na NF (130 pos.)</label>
              <input id="edtOBS3" name="edtOBS3" type="text"
                     class="form-control form-control-sm"
                     value="<%=edtOBS3%>"
                     maxlength="130"
                     autocomplete="off" />
            </div>
          </div>

          <div class="row g-2 mb-3">
            <div class="col-12">
              <label for="edtOBS4" class="form-label">Observação 4 — Constar na NF (130 pos.)</label>
              <input id="edtOBS4" name="edtOBS4" type="text"
                     class="form-control form-control-sm"
                     value="<%=edtOBS4%>"
                     maxlength="130"
                     autocomplete="off" />
            </div>
          </div>

          <hr class="my-3" />

          <!-- Botões de ação -->
          <div class="d-flex gap-2 justify-content-end">
            <button type="button" class="btn btn-light btn-sm" onclick="limparCampos()">
              <i class="bi bi-eraser me-2"></i>Limpar
            </button>
            <button type="button" class="btn btn-primary btn-sm" onclick="Salvar()">
              Salvar
            </button>
          </div>

        </form>

      </div>
    </div>

  </div>

  <!-- Bootstrap JS Bundle -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <!-- jQuery 4 -->
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script src="../funcoes/funcoesJAVASCRIPT.js?id=2345652"></script>

  <script>
    'use strict';

    /* ------------------------------------------------------------------ */
    /* Toast                                                                */
    /* ------------------------------------------------------------------ */
    const TOAST_DELAY = 4500;

    const TOAST_CFG = {
      sucesso: { icon: 'bi-check-circle-fill', cor: '#198754', barra: '#198754', titulo: 'Sucesso'    },
      erro:    { icon: 'bi-x-circle-fill',     cor: '#dc3545', barra: '#dc3545', titulo: 'Erro'       },
      aviso:   { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107', titulo: 'Atenção' },
      info:    { icon: 'bi-info-circle-fill',   cor: '#6c757d', barra: '#6c757d', titulo: 'Informação' }
    };

    function exibirToast(tipo, mensagem) {
      const cfg = TOAST_CFG[tipo] || TOAST_CFG.info;
      const el  = document.createElement('div');
      el.className = 'toast-custom';
      el.innerHTML = `
        <div class="d-flex align-items-center px-3 pt-2 pb-1 gap-2">
          <i class="bi ${cfg.icon}" style="color:${cfg.cor};font-size:1.1rem;"></i>
          <strong style="color:${cfg.cor};font-size:.88rem;">${cfg.titulo}</strong>
          <button type="button" class="btn-close btn-close ms-auto" style="font-size:.7rem;"
                  onclick="this.closest('.toast-custom').remove()"></button>
        </div>
        <div class="px-3 pb-2" style="font-size:.85rem;color:#212529;">${mensagem}</div>
        <div style="height:3px;background:#e9ecef;">
          <div class="toast-progress-bar" style="height:100%;background:${cfg.barra};"></div>
        </div>`;
      document.getElementById('toastContainer').appendChild(el);
      setTimeout(() => el.remove(), TOAST_DELAY);
    }

    /* ------------------------------------------------------------------ */
    /* Toast do servidor (injetado via ASP)                                 */
    /* ------------------------------------------------------------------ */
    <%
    If toastMsg <> "" Then
      Response.Write "exibirToast('" & toastTipo & "', '" & Replace(toastMsg, "'", "\'") & "');"
    End If
    %>

    /* ------------------------------------------------------------------ */
    /* Ações do formulário                                                  */
    /* ------------------------------------------------------------------ */
    function Buscar() {
      $('#buscar').val('S');
      $('#salvar').val('N');
      document.getElementById('formObs').submit();
    }

    function Salvar() {
      if ($('#edtPEDIDO').val().trim() === '') {
        exibirToast('aviso', 'Informe o número do pedido antes de salvar.');
        return;
      }
      $('#buscar').val('N');
      $('#salvar').val('S');
      document.getElementById('formObs').submit();
    }

    function limparCampos() {
      $('#edtPEDIDO').val('');
      $('#edtOBS1').val('');
      $('#edtOBS2').val('');
      $('#edtOBS3').val('');
      $('#edtOBS4').val('');
    }
  </script>

</body>
</html>
