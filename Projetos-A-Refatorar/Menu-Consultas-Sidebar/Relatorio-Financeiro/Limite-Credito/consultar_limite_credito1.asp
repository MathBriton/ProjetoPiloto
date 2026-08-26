<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->

<%
function AddZerosLeft(originalString, totalLength)

  Dim currentLength
  currentLength = Len(originalString)
  
  if currentLength < totalLength then
    Dim zerosToAdd
    zerosToAdd = totalLength - currentLength
    originalString = String(zerosToAdd, "0") & originalString
  end if
  
  AddZerosLeft = originalString

end function
%>

<%
codsiga = iif( nivel = 0, 0, CODUSER )
codfil = trim(SESSION("FILIAL"))

SQL = "" &_
  "SELECT " &_
    "isnull(NivelLimiteCredito,0) AS NivelLimiteCredito "  &_
  "FROM CRM_DTI " &_
  "WHERE CRM_COD = '" & SESSION("CODUSER")   &_
    "' AND CRM_D_E_L_E_T <> '*' "

SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)

if not isobject( PIR_HELPDESK ) or ( PIR_HELPDESK is NOTHING) then
  response.end
end if

NivelAcesso       = iif( not PIR_HELPDESK.EOF, PIR_HELPDESK("NivelLimiteCredito"), false )
NivelAcessoLimite = iif( not PIR_HELPDESK.EOF and PIR_HELPDESK("NivelLimiteCredito")=1, 1, 0 )

NivelAcesso = True

PIR_HELPDESK.Close()
SET PIR_HELPDESK = NOTHING

if NivelAcesso = false then

  select case codfil
    case "00" : subTitulo = "Pirapetinga / Uberaba / E-Shop"
    case "05" : subTitulo = "Bento Gonçalves"
    case "06" : subTitulo = "Maranguape"
    case "07" : subTitulo = "Saquarema"
  end select

end if
%>

<!DOCTYPE html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <title>SWnet — Consultar Limite de Crédito</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <!-- Bootstrap 5.3 -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" />
    <!-- Bootstrap Icons -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

    <style>
      /* ── Tokens de design ─────────────────────────────────────────── */
      :root {
        --color-bg:            hsl(240, 9%, 98%);
        --color-surface:       hsl(0, 0%, 100%);
        --color-surface-muted: hsl(240, 9%, 97%);
        --color-border:        hsl(240, 6%, 90%);
        --color-border-strong: hsl(240, 6%, 82%);

        --color-text:          hsl(240, 6%, 10%);
        --color-text-muted:    hsl(240, 4%, 46%);
        --color-text-faint:    hsl(240, 4%, 64%);

        --color-accent:        hsl(210, 100%, 45%);
        --color-accent-hover:  hsl(210, 100%, 40%);
        --color-accent-soft:   hsl(210, 100%, 96%);

        --color-success:       hsl(145, 63%, 32%);
        --color-success-soft:  hsl(145, 60%, 95%);
        --color-danger:        hsl(3, 90%, 58%);
        --color-danger-soft:   hsl(3, 100%, 96%);

        --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

        --radius-sm:  6px;
        --radius-md:  10px;
        --radius-lg:  14px;

        --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
        --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

        --transition: 160ms ease;

        --space-1: 4px;
        --space-2: 8px;
        --space-3: 12px;
        --space-4: 16px;
        --space-5: 24px;
        --space-6: 32px;
      }

      /* ── Reset base ───────────────────────────────────────────────── */
      *, *::before, *::after { box-sizing: border-box; }

      body {
        font-family: var(--font-sans);
        background-color: var(--color-bg);
        color: var(--color-text);
        font-size: .875rem;
        margin: 0;
      }

      /* ── Card do formulário ───────────────────────────────────────── */
      .form-card {
        background-color: var(--color-surface);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        box-shadow: var(--shadow-sm);
        padding: var(--space-5);
      }

      /* ── Labels ───────────────────────────────────────────────────── */
      .form-label {
        font-size: .75rem;
        font-weight: 500;
        color: var(--color-text-muted);
        margin-bottom: var(--space-1);
        text-transform: uppercase;
        letter-spacing: .04em;
      }

      /* ── Inputs ───────────────────────────────────────────────────── */
      .form-control,
      .form-select {
        border: 1px solid var(--color-border);
        border-radius: var(--radius-sm);
        background-color: var(--color-surface);
        color: var(--color-text);
        font-size: .875rem;
        padding: .45rem .75rem;
        transition: border-color var(--transition), box-shadow var(--transition);
      }

      .form-control:focus,
      .form-select:focus {
        border-color: var(--color-accent);
        box-shadow: 0 0 0 3px hsla(210, 100%, 45%, .12);
        outline: none;
      }

      .form-control[readonly] {
        background-color: var(--color-surface-muted);
        color: var(--color-text-muted);
        cursor: default;
      }

      .input-text-right {
        text-align: right;
      }

      /* ── Campos de valor (Limite / Aberto / Saldo) ────────────────── */
      .value-group {
        background-color: var(--color-surface-muted);
        border: 1px solid var(--color-border);
        border-radius: var(--radius-md);
        padding: var(--space-4) var(--space-5);
      }

      .value-label {
        font-size: .7rem;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: .05em;
        color: var(--color-text-muted);
        margin-bottom: var(--space-1);
      }

      .value-field {
        font-size: 1rem;
        font-weight: 600;
        color: var(--color-text);
        background: none;
        border: none;
        padding: 0;
        width: 100%;
        text-align: right;
        font-family: var(--font-sans);
      }

      .value-field:focus {
        outline: none;
      }

      /* ── Botões ───────────────────────────────────────────────────── */
      .btn {
        font-size: .8125rem;
        font-weight: 500;
        border-radius: var(--radius-sm);
        padding: .45rem 1rem;
        transition: background-color var(--transition), border-color var(--transition), color var(--transition);
        display: inline-flex;
        align-items: center;
        gap: var(--space-1);
        min-height: 36px;
        border: 1px solid transparent;
      }

      .btn-primary {
        background-color: var(--color-accent);
        border-color: var(--color-accent);
        color: hsl(0, 0%, 100%);
      }

      .btn-primary:hover,
      .btn-primary:focus {
        background-color: var(--color-accent-hover);
        border-color: var(--color-accent-hover);
        color: hsl(0, 0%, 100%);
      }

      .btn-ghost {
        background-color: transparent;
        border-color: var(--color-border);
        color: var(--color-text-muted);
      }

      .btn-ghost:hover {
        background-color: var(--color-surface-muted);
        border-color: var(--color-border-strong);
        color: var(--color-text);
      }

      .btn-success-outline {
        background-color: var(--color-success-soft);
        border-color: hsl(145, 40%, 75%);
        color: var(--color-success);
      }

      .btn-success-outline:hover {
        background-color: hsl(145, 50%, 88%);
        color: var(--color-success);
      }

      .btn-danger-outline {
        background-color: var(--color-danger-soft);
        border-color: hsl(3, 60%, 80%);
        color: var(--color-danger);
      }

      .btn-danger-outline:hover {
        background-color: hsl(3, 70%, 90%);
        color: var(--color-danger);
      }

      /* ── Título interno do formulário ────────────────────────────── */
      .form-section-title {
        font-size: clamp(.9375rem, 0.85rem + 0.35vw, 1.125rem);
        font-weight: 600;
        color: var(--color-text);
        letter-spacing: -.015em;
        margin: 0;
        padding-bottom: var(--space-4);
        border-bottom: 1px solid var(--color-border);
      }

      /* ── Botão de lupa (apenas ícone, quadrado) ───────────────────── */
      .btn-search-icon {
        width: 36px;
        min-height: 36px;
        padding: 0;
        justify-content: center;
        flex-shrink: 0;
      }

      /* ── Barra de ações (aparecem após consultar) ─────────────────── */
      .action-bar {
        display: none;
        flex-direction: column;
        align-items: center;
        gap: var(--space-3);
        background-color: var(--color-accent-soft);
        border: 1px solid hsl(210, 60%, 82%);
        border-radius: var(--radius-md);
        padding: var(--space-4) var(--space-5);
      }

      .action-bar.visible {
        display: flex;
      }

      .action-bar-title {
        font-size: .7rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .06em;
        color: var(--color-text-muted);
      }

      .action-bar-buttons {
        display: flex;
        gap: var(--space-2);
        flex-wrap: wrap;
        justify-content: center;
      }

      /* ── Overlay de loading ───────────────────────────────────────── */
      .overlay-loading {
        display: none;
        position: fixed;
        inset: 0;
        background-color: hsla(240, 6%, 10%, .35);
        z-index: 1030;
        align-items: center;
        justify-content: center;
      }

      .overlay-loading.active {
        display: flex;
      }

      .spinner-ring {
        width: 48px;
        height: 48px;
        border: 4px solid hsla(0, 0%, 100%, .3);
        border-top-color: hsl(0, 0%, 100%);
        border-radius: 50%;
        animation: spin 700ms linear infinite;
      }

      @keyframes spin {
        to { transform: rotate(360deg); }
      }

      @media (prefers-reduced-motion: reduce) {
        .spinner-ring { animation: none; border-top-color: hsl(0, 0%, 100%); }
        .btn { transition: none; }
        .form-control, .form-select { transition: none; }
      }

      /* ── Separador de seção ───────────────────────────────────────── */
      .section-divider {
        border: none;
        border-top: 1px solid var(--color-border);
        margin: var(--space-5) 0;
      }

      /* ── Responsividade ───────────────────────────────────────────── */
      @media (max-width: 768px) {
        .value-group {
          margin-bottom: var(--space-3);
        }

        .action-bar-buttons {
          flex-direction: column;
          align-items: stretch;
        }

        .action-bar-buttons .btn {
          justify-content: center;
        }
      }
    </style>
  </head>

  <body>

    <!-- Overlay de loading -->
    <div id="overlayLoading" class="overlay-loading" role="status" aria-live="polite" aria-label="Carregando...">
      <div class="spinner-ring"></div>
    </div>

    <!-- Conteúdo principal -->
    <main class="container-fluid px-3 px-lg-4" style="max-width: 860px; min-height: 100vh; display: flex; flex-direction: column; justify-content: center; padding-top: var(--space-6); padding-bottom: var(--space-6);">

      <div class="form-card">

        <form id="formLimiteCredito" name="formLimiteCredito" method="post" autocomplete="off">

          <!-- Título interno do formulário -->
          <div class="mb-4">
            <h2 class="form-section-title">Consultar Limite de Crédito</h2>
          </div>

          <!-- Filial (visível apenas para NivelAcesso = true) -->
          <% if NivelAcesso = true then %>
          <div class="mb-4">
            <label for="cboFilial" class="form-label">Filial</label>
            <select name="cboFilial" id="cboFilial" class="form-select" onchange="habilitaCodLoja(this)">
              <option value="">Selecione...</option>
              <option value="00" <% if request("cboFilial") = "00" then %>selected<% end if %>>Pirapetinga / Uberaba / E-Shop</option>
              <option value="05" <% if request("cboFilial") = "05" then %>selected<% end if %>>Bento Gonçalves</option>
              <option value="06" <% if request("cboFilial") = "06" then %>selected<% end if %>>Maranguape</option>
              <option value="07" <% if request("cboFilial") = "07" then %>selected<% end if %>>Saquarema</option>
            </select>
          </div>
          <% end if %>

          <!-- Cliente -->
          <div class="row g-3 mb-4">
            <div class="col-12 col-md-5">
              <label class="form-label">Cliente</label>
              <div class="d-flex gap-2 align-items-center">
                <input type="text"
                       name="CodCli" id="CodCli"
                       class="form-control"
                       style="max-width: 100px;"
                       maxlength="10"
                       title="Código do cliente"
                       onchange="verificaLimiteCredito()"
                       aria-label="Código do cliente" />
                <input type="text"
                       name="CodLoja" id="CodLoja"
                       class="form-control"
                       style="max-width: 56px;"
                       maxlength="2"
                       title="Código da loja"
                       aria-label="Código da loja" />
                <button type="button"
                        class="btn btn-primary btn-search-icon"
                        onclick="procuraCliente()"
                        title="Pesquisar cliente"
                        aria-label="Pesquisar cliente">
                  <i class="bi bi-search" aria-hidden="true"></i>
                </button>
              </div>
            </div>

            <div class="col-12 col-md-7">
              <label for="RazaoSocial" class="form-label">Razão Social</label>
              <input type="text"
                     name="RazaoSocial" id="RazaoSocial"
                     class="form-control"
                     title="Razão Social"
                     readonly
                     aria-label="Razão Social" />
            </div>
          </div>

          <hr class="section-divider" />

          <!-- Valores: Limite / Aberto / Saldo -->
          <div class="row g-3 mb-4">
            <div class="col-12 col-md-4">
              <div class="value-group">
                <div class="value-label">Limite</div>
                <input type="text"
                       name="Limite" id="Limite"
                       class="value-field input-text-right"
                       value="<%=Limite%>"
                       readonly
                       tabindex="-1"
                       aria-label="Limite de crédito" />
              </div>
            </div>
            <div class="col-12 col-md-4">
              <div class="value-group">
                <div class="value-label">Aberto</div>
                <input type="text"
                       name="Aberto" id="Aberto"
                       class="value-field input-text-right"
                       value="<%=Aberto%>"
                       readonly
                       tabindex="-1"
                       aria-label="Valor em aberto" />
              </div>
            </div>
            <div class="col-12 col-md-4">
              <div class="value-group">
                <div class="value-label">Saldo</div>
                <input type="text"
                       name="Saldo" id="Saldo"
                       class="value-field input-text-right"
                       value="<%=Credito%>"
                       readonly
                       tabindex="-1"
                       aria-label="Saldo disponível" />
              </div>
            </div>
          </div>

          <!-- Barra de ações secundárias (visível após resultado) -->
          <div id="btnsVer" class="action-bar mt-3">
            <div class="action-bar-title"><label style="color:black;">Ações</label></div>
            <div class="action-bar-buttons">
              <button type="button" class="btn btn-success-outline" onclick="VerDetalhes()">
                <i class="bi bi-card-list me-1" aria-hidden="true"></i>
                Ver Detalhes
              </button>
              <% if NIVEL = "0" then %>
              <button type="button" class="btn btn-primary" onclick="acessaRelatorio()">
                <i class="bi bi-clock-history me-1" aria-hidden="true"></i>
                Histórico de Limites Liberados
              </button>
              <% end if %>
            </div>
          </div>

        </form>

      </div><!-- /.form-card -->

      <!-- Botões principais centralizados no body -->
      <div class="d-flex justify-content-center gap-3 mt-4">
        <button type="button"
                class="btn btn-ghost"
                style="color:black;"
                onclick="limparFiltro()"
                title="Limpar filtro">
          <i class="bi bi-eraser me-1" aria-hidden="true"></i>
          Limpar
        </button>
        <button type="button"
                class="btn btn-primary"
                onclick="verificaLimiteCredito()"
                title="Consultar limite de crédito">
          <i class="bi bi-search me-1" aria-hidden="true"></i>
          Consultar
        </button>
      </div>

    </main>

    <!-- jQuery + Bootstrap 5 JS -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>

    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>

    <script>
      let atuFilial = '';
      let atuCodCli = '';
      let atuCodLoj = '';

      let argPlanta = '<%=AddZerosLeft(codfil,2)%>';

      // ── Inicialização ───────────────────────────────────────────────
      function doOnLoad() {
        ocultaLoader();
      }

      // ── Loading overlay ─────────────────────────────────────────────
      function mostraLoader() {
        document.getElementById('overlayLoading').classList.add('active');
      }

      function ocultaLoader() {
        document.getElementById('overlayLoading').classList.remove('active');
      }

      // ── Pesquisa cliente (popup) ─────────────────────────────────────
      function procuraCliente() {
        const largura = 650;
        const altura  = 600;
        const left    = window.screenX + (window.outerWidth  - largura) / 2;
        const top     = window.screenY + (window.outerHeight - altura)  / 2;

        document.getElementById('btnsVer').classList.remove('visible');

        const mapaPlanta = { '00': 'PIRAPETINGA', '04': 'UBERABA', '05': 'BENTO', '06': 'MARANGUAPE', '07': 'SAQUAREMA' };
        const nPlanta    = mapaPlanta[argPlanta] || '';

        const url = 'buscaBICLIENTES.asp?id1=CodCli&id2=CodLoja&id3=RazaoSocial&id6=' + nPlanta;

        window.open(url, 'Pesquisa Cliente', `width=${largura},height=${altura},left=${left},top=${top},resizable,scrollbars`);
      }

      // ── Verifica limite de crédito ───────────────────────────────────
      function verificaLimiteCredito() {
        mostraLoader();

        const aFil = document.getElementById('cboFilial').value;
        const aCli = document.getElementById('CodCli').value;
        const aLoj = document.getElementById('CodLoja').value;

        if (atuFilial !== aFil || atuCodCli !== aCli || atuCodLoj !== aLoj) {
          atuFilial = aFil;
          atuCodCli = aCli;
          atuCodLoj = aLoj;
          document.getElementById('Limite').value = '0,00';
          document.getElementById('Aberto').value = '0,00';
          document.getElementById('Saldo').value  = '0,00';
        }

        const lnk = window.location.protocol + '//' +
                    window.location.hostname  + '/' +
                    'corporativo/financeiro/consulta_limite_credito/' +
                    'buscaLimiteCredito.asp' +
                    '?filial='  + encodeURIComponent(atuFilial) +
                    '&cliente=' + encodeURIComponent(atuCodCli) +
                    '&loja='    + encodeURIComponent(atuCodLoj);

        ObjAjax(lnk);
      }

      // ── Limpar filtro ───────────────────────────────────────────────
      function limparFiltro() {
        ocultaLoader();

        document.getElementById('btnsVer').classList.remove('visible');
        document.getElementById('cboFilial').value    = '<%=trim(SESSION("FILIAL"))%>';
        document.getElementById('CodCli').value       = '';
        document.getElementById('CodLoja').value      = '';
        document.getElementById('RazaoSocial').value  = '';
        document.getElementById('Limite').value       = '';
        document.getElementById('Aberto').value       = '';
        document.getElementById('Saldo').value        = '';
      }

      // ── Callback AJAX (objReturn) ───────────────────────────────────
      function objReturn(obj) {

        if (obj.erro === 'ErroLogof') {
          alert('Favor se conectar novamente ao SKNET.\nSeu logon EXPIROU.');
          if (obj.tipo === 'Grupo') {
            document.getElementById('CodCli').value    = '';
            document.getElementById('CodLoja').value   = '';
            document.getElementById('RazaoSocial').value = '';
          }
          return false;
        }

        if (obj.erro === 'ErroNEncontrado') {
          if (obj.tipo === 'Grupo') {
            document.getElementById('CodCli').value    = '';
            document.getElementById('CodLoja').value   = '';
            document.getElementById('RazaoSocial').value = '';
            alert('CLIENTE NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente Bloqueado.\nCliente não pertence a sua carteira.\nCliente não Cadastrado.');
          }
          return false;
        }

        if (obj.erro === 'ErroCliente') {
          alert('Favor preencher os campos Cliente e Loja.');
          if (obj.tipo === 'Grupo') {
            document.getElementById('CodCli').value    = '';
            document.getElementById('CodLoja').value   = '';
            document.getElementById('RazaoSocial').value = '';
          }
          return false;
        }

        if (obj.tipo === 'Grupo') {
          document.getElementById('RazaoSocial').value = obj.val1;
        }

        if (obj.tipo === 'LimiteCredito') {
          document.getElementById('Limite').value = obj.val1;
          document.getElementById('Aberto').value = obj.val2;
          document.getElementById('Saldo').value  = obj.val3;
          document.getElementById('btnsVer').classList.add('visible');
        }

        ocultaLoader();
      }

      // ── Ver Detalhes ────────────────────────────────────────────────
      function VerDetalhes() {
        const aFil = document.getElementById('cboFilial').value;
        const aCli = document.getElementById('CodCli').value;
        const aLoj = document.getElementById('CodLoja').value;
        const aNom = document.getElementById('RazaoSocial').value;

        if (aCli || aLoj || aNom) {
          const url = window.location.protocol + '//' +
                      window.location.hostname  + '/' +
                      'corporativo/financeiro/analisa_credito/analisacreditoV2.asp' +
                      '?codcli='  + aCli +
                      '&codLoja=' + aLoj +
                      '&Filial='  + aFil +
                      '&nomcli='  + aNom;

          const screenHeight = window.screen.availHeight;
          window.open(url, 'AnalisaCredito', `width=900,height=${screenHeight},top=0,left=0`);
        } else {
          ocultaLoader();
          alert('Selecione o cliente para visualizar os detalhes.');
        }
      }

      // ── Acessar relatório ───────────────────────────────────────────
      function acessaRelatorio() {
        const aFil = document.getElementById('cboFilial').value;
        const aCli = document.getElementById('CodCli').value;
        const aLoj = document.getElementById('CodLoja').value;
        const aNom = document.getElementById('RazaoSocial').value;

        const url = window.location.protocol + '//' +
                    window.location.hostname  + '/corporativo/SolicitacaoCredito/SolicitacaoCreditoRelatorio.asp' +
                    '?filtroDataSolicitacaoDe='  + primeiroDiaDoMesFormatado() +
                    '&filtroDataSolicitacaoAte=' + diaAtualFormatado() +
                    '&filtroPlanta='    + aFil +
                    '&filtroCodCliente=' + aCli +
                    '&lojacli='         + aLoj +
                    '&razao='           + aNom;

        window.open(url, 'meio');
      }

      // ── Helpers de data ─────────────────────────────────────────────
      function primeiroDiaDoMesFormatado() {
        const d = new Date();
        const p = new Date(d.getFullYear(), d.getMonth(), 1);
        return `${p.getFullYear()}-${String(p.getMonth() + 1).padStart(2, '0')}-01`;
      }

      function diaAtualFormatado() {
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      }

      // ── Habilita/desabilita campo Loja conforme filial ───────────────
      function habilitaCodLoja(obj) {
        if (obj.value) argPlanta = obj.value;

        document.getElementById('CodLoja').value      = '';
        document.getElementById('CodCli').value       = '';
        document.getElementById('RazaoSocial').value  = '';
        document.getElementById('Limite').value        = '0,00';
        document.getElementById('Aberto').value        = '0,00';
        document.getElementById('Saldo').value         = '0,00';

        const inputCodCli = document.getElementById('CodCli');
        const inputLojCli = document.getElementById('CodLoja');

        if (obj.value === '05' || obj.value === '06') {
          document.getElementById('CodLoja').value    = '01';
          document.getElementById('CodLoja').disabled = true;
          inputCodCli.addEventListener('blur', onBlurHandler);
          inputLojCli.removeEventListener('blur', onBlurHandler);
        } else {
          document.getElementById('CodLoja').value    = '';
          document.getElementById('CodLoja').disabled = false;
          inputLojCli.addEventListener('blur', onBlurHandler);
          inputCodCli.removeEventListener('blur', onBlurHandler);
        }

        function onBlurHandler() {
          pesquisaCliente(this);
        }
      }

      // ── Pesquisa cliente via AJAX ────────────────────────────────────
      function pesquisaCliente(obj) {
        if (typeof obj.value !== 'undefined' && obj.value.trim() !== '') {
          const loja = (formLimiteCredito.cboFilial.value === '05' || formLimiteCredito.cboFilial.value === '06')
                       ? '01'
                       : obj.value;

          const lnk = 'validaClienteLC_ajax.asp?planta=' + formLimiteCredito.cboFilial.value +
                      '&cliente=' + formLimiteCredito.CodCli.value +
                      '&loja='    + loja +
                      '&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Grupo';

          ObjAjax(lnk);
        } else {
          obj.value = '';
        }
      }
    </script>

  </body>
</html>
