<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp"-->
<!--#include file="../../../funcoes/funcoes.asp"-->
<!--#include file="../../../includes/variaveis.asp"-->
<%
Response.Buffer       = True
Response.Charset      = "UTF-8"
Response.CodePage     = 65001
Response.Expires      = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>

  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA"        content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet - Relatório de FT</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
        crossorigin="anonymous" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
        rel="stylesheet" />

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
          integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
          crossorigin="anonymous"></script>

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.15rem;
      font-weight: 700;
      color: #212529;
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
    .filter-card-body { padding: 1rem 1.25rem 1.25rem; }
    .filter-card-body label {
      font-size: .82rem;
      font-weight: 600;
      color: #6c757d;
      margin-bottom: .25rem;
    }

    .table-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .table-wrapper {
      overflow-x: auto;
      overflow-y: auto;
      max-height: 62vh;
    }
    .table-wrapper table { min-width: max-content; }
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
    .table-placeholder p  { margin: 0; font-size: .9rem; text-align: center; }

    .table-card { position: relative; min-height: 120px; }

    .table-loading {
      position: absolute;
      inset: 0;
      background: rgba(255,255,255,.82);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 10;
      gap: .75rem;
      border-radius: 6px;
      min-height: 120px;
    }
    .table-loading .spinner-border { width: 2.25rem; height: 2.25rem; }
    .table-loading span            { font-size: .88rem; font-weight: 600; color: #495057; }

    .ft-section-header {
      background-color: #f8f9fa;
      font-size: .82rem;
      font-weight: 700;
      color: #495057;
      padding: .5rem .75rem;
      border-top: 2px solid #dee2e6;
      border-bottom: 1px solid #dee2e6;
      text-align: center;
    }

    .table-footer {
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      padding: .5rem 1rem;
      border-top: 1px solid #dee2e6;
      background: #f8f9fa;
      font-size: .82rem;
      gap: .5rem;
    }
    .table-footer .zone-left   { display: flex; align-items: center; gap: .4rem; }
    .table-footer .zone-right  { display: flex; justify-content: flex-end; align-items: center; color: #6c757d; }
    .table-footer .pagination         { margin-bottom: 0; }
    .table-footer .pagination .page-link { font-size: .78rem; padding: .25rem .5rem; color: #6c757d; border-color: #dee2e6; }
    .table-footer .pagination .page-link:hover { color: #fff; background-color: #6c757d; border-color: #6c757d; }
    .table-footer .pagination .page-item.active .page-link { background-color: #6c757d; border-color: #6c757d; color: #fff; }
    .table-footer .pagination .page-item.disabled .page-link { color: #adb5bd; }

    #modalImgFT .modal-body   { text-align: center; background: #1a1a1a; padding: .5rem; }
    #modalImgFT .modal-body img { max-width: 100%; max-height: 80vh; border-radius: 4px; }
    #modalImgFT .modal-header { background: #212529; color: #fff; border-bottom: 1px solid #444; display: flex; align-items: center; justify-content: space-between; }
    #modalImgFT .modal-header .btn-close { filter: invert(1); }

    .info-card {
      background-color: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    .info-card-header {
      background-color: #fff;
      color: #212529;
      padding: .6rem 1rem;
      border-radius: 6px 6px 0 0;
      border-bottom: 1px solid #dee2e6;
      font-weight: 600;
      font-size: .9rem;
    }
    .info-card-body { padding: .75rem 1.25rem; }

    .cliente-badge {
      display: inline-flex;
      align-items: center;
      gap: .35rem;
      background: #f0f4ff;
      border: 1px solid #c7d4f0;
      border-radius: 4px;
      padding: .3rem .7rem;
      font-size: .82rem;
    }
    .cliente-badge .cliente-label { color: #6c757d; font-weight: 600; }
    .cliente-badge .cliente-sep   { color: #adb5bd; }
    .cliente-badge .cliente-cod   { font-weight: 700; color: #1a4dbf; }
    .cliente-badge .cliente-loja  { font-weight: 700; color: #1a4dbf; }
    .cliente-badge .cliente-razao { color: #1a4dbf; font-weight: 700; }

    .legenda-ft {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: .5rem;
      font-size: .78rem;
      color: #6c757d;
    }
    .legenda-ft .legenda-group { display: flex; align-items: center; gap: .4rem; flex-wrap: wrap; }
    .legenda-ft .legenda-sep {
      display: inline-block;
      width: 1px;
      height: 1rem;
      background: #dee2e6;
    }

    .toast-container-custom {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 1100;
      display: flex;
      flex-direction: column;
      gap: .5rem;
    }
    .toast-custom {
      min-width: 280px;
      max-width: 360px;
      border-radius: 10px;
      overflow: hidden;
      box-shadow: 0 4px 18px rgba(0,0,0,.13);
      background: #fff;
      border: 1px solid #dee2e6;
      animation: toastSlideIn .28s cubic-bezier(.22,1,.36,1);
    }
    @keyframes toastSlideIn {
      from { opacity: 0; transform: translateX(40px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-custom .toast-inner {
      display: flex;
      align-items: flex-start;
      gap: .6rem;
      padding: .65rem .85rem;
    }
    .toast-custom .toast-icon {
      font-size: 1.25rem;
      line-height: 1;
      margin-top: .1rem;
    }
    .toast-custom .toast-body-text { flex: 1; min-width: 0; }
    .toast-custom .toast-title {
      font-size: .82rem;
      font-weight: 700;
      margin-bottom: .1rem;
    }
    .toast-custom .toast-msg {
      font-size: .8rem;
      color: #495057;
      word-wrap: break-word;
    }
    .toast-custom .toast-close {
      background: none;
      border: 0;
      color: #6c757d;
      font-size: 1.2rem;
      line-height: 1;
      cursor: pointer;
      padding: 0 .2rem;
    }
    .toast-custom .toast-progress {
      height: 3px;
      background: #f0f0f0;
      overflow: hidden;
    }
    .toast-custom .toast-progress-bar {
      height: 100%;
      animation: toastProgressShrink 4.5s linear forwards;
    }
    @keyframes toastProgressShrink {
      from { width: 100%; } to { width: 0%; }
    }
    .toast-sucesso .toast-icon,
    .toast-sucesso .toast-title { color: #198754; }
    .toast-sucesso .toast-progress-bar { background: #198754; }
    .toast-erro .toast-icon,
    .toast-erro .toast-title { color: #dc3545; }
    .toast-erro .toast-progress-bar { background: #dc3545; }
    .toast-aviso .toast-icon,
    .toast-aviso .toast-title { color: #856404; }
    .toast-aviso .toast-progress-bar { background: #ffc107; }
    .toast-info .toast-icon,
    .toast-info .toast-title { color: #6c757d; }
    .toast-info .toast-progress-bar { background: #6c757d; }
  </style>

</head>
<body>

  <div id="toastMsg" class="toast-container-custom"></div>

  <div class="modal fade" id="modalImgFT" tabindex="-1"
       aria-labelledby="modalImgFTLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg modal-dialog-centered">
      <div class="modal-content border-0">
        <div class="modal-header py-2">
          <span class="modal-title fw-bold" id="modalImgFTLabel">
            <i class="bi bi-image me-2"></i>Imagem da FT
          </span>
          <div class="d-flex align-items-center gap-2">
            <button type="button" class="btn btn-primary btn-sm d-none" id="btnDownloadImgFT"
                    title="Download da imagem">
              <i class="bi bi-download me-2"></i>Download
            </button>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>
        </div>
        <div class="modal-body p-2">
          <div id="imgFTLoading" class="text-center py-4 text-white">
            <div class="spinner-border spinner-border-sm text-light me-2"></div>Carregando imagem...
          </div>
          <img id="imgFTSrc" src="" alt="Imagem da FT" class="d-none" />
          <div id="imgFTError" class="d-none text-center py-4 text-secondary">
            <i class="bi bi-image-alt" style="font-size:2rem;opacity:.4;"></i>
            <p class="mt-2 mb-0">Imagem não encontrada.</p>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="mb-3">
      <span class="page-title">Relatório de FT</span>
    </div>

    <div class="filter-card mb-3">
      <div class="filter-card-header">
        <i class="bi bi-funnel me-2"></i>Filtros
      </div>
      <div class="filter-card-body">

        <div class="row g-2 align-items-end">

          <div class="col-sm-auto">
            <label for="cboPlanta">Filial</label>
            <select id="cboPlanta" class="form-select form-select-sm" style="min-width:140px">
              <option value="">Selecione...</option>
              <option value="00">Pirapetinga</option>
              <option value="04">Uberaba</option>
              <option value="05">Bento Gonçalves</option>
              <option value="06">Maranguape</option>
              <option value="07">Saquarema</option>
            </select>
          </div>

          <div class="col-sm-auto">
            <label for="cboStatus">Status</label>
            <select id="cboStatus" class="form-select form-select-sm" style="min-width:110px">
              <option value="">Todos</option>
              <option value="ATIVO">Ativo</option>
              <option value="INATIVO">Inativo</option>
            </select>
          </div>

          <div class="col-sm-auto d-none" id="colGrupo">
            <label for="cboGrupo">Grupo</label>
            <select id="cboGrupo" class="form-select form-select-sm" style="min-width:130px">
              <option value="">Todos</option>
            </select>
          </div>

          <div class="col-sm-auto d-none" id="colTipo">
            <label for="cboTipo">Tipo</label>
            <select id="cboTipo" class="form-select form-select-sm" style="min-width:130px">
              <option value="">Todos</option>
            </select>
          </div>

          <div class="col-sm">
            <label for="termoCli">Cliente</label>
            <div class="input-group input-group-sm" style="max-width:480px">
              <button type="button" class="btn btn-secondary btn-sm" id="btnConsultaCliente"
                      title="Buscar pelo popup">
                Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
              </button>
              <input type="text" id="termoCli"
                     class="form-control form-control-sm"
                     placeholder="Código ou fantasia" />
              <span class="input-group-text px-2" style="font-size:.78rem;color:#6c757d;">Loja</span>
              <input type="text" id="lojaCliVis"
                     class="form-control form-control-sm"
                     placeholder="Todas"
                     maxlength="4"
                     style="max-width:60px"
                     title="Loja específica (opcional). Deixe vazio para trazer todas as lojas." />
              <input type="hidden" id="codcli"  value="" />
              <input type="hidden" id="lojacli" value="" />
              <input type="hidden" id="razao"   value="" />
            </div>
          </div>

          <div class="col-sm-auto ms-auto">
            <!-- <label>&nbsp;</label> -->
            <div class="d-flex gap-2">
              <button type="button" class="btn btn-light btn-sm" id="btnLimpar">
                <i class="bi bi-eraser">&nbsp; Limpar </i>
              </button>
              <button type="button" class="btn btn-primary btn-sm" id="btnFiltrar">
                <i class="bi bi-funnel-fill me-2"></i>Filtrar
              </button>
            </div>
          </div>

        </div>

      </div>
    </div>

    <div class="info-card mb-3 d-none" id="cardCliente">
      <div class="info-card-header">
        <i class="bi bi-person me-2"></i>Cliente importado
      </div>
      <div class="info-card-body">
        <div class="d-flex flex-wrap align-items-center gap-3">

          <!-- Badge preenchido pelo popup (código + loja + fantasia) -->
          <div class="cliente-badge d-none" id="clienteBadgeDyn">
            <span class="cliente-label">CÓDIGO:</span>
            <span class="cliente-cod"></span>
            <span class="cliente-sep">-</span>
            <span class="cliente-label">LOJA:</span>
            <span class="cliente-loja"></span>
            <span class="cliente-sep">-</span>
            <span class="cliente-label">FANTASIA:</span>
            <span class="cliente-razao"></span>
          </div>
          <!-- Badge para digitação livre (código ou fantasia) -->
          <div class="cliente-badge d-none" id="clienteBadgeTexto">
            <span class="cliente-label">FILTRANDO POR:</span>
            <span class="cliente-cod-texto" style="color:#212529;font-weight:600;"></span>
          </div>
          <span id="clienteSemSelecao" style="font-size:.82rem;color:#6c757d;font-style:italic;">Nenhum cliente selecionado &mdash; exibindo todos.</span>

          <div class="legenda-ft ms-auto d-none" id="legendaPirUbe">
            <span class="legenda-group">
              <strong>Status:</strong>
              <span><strong>E</strong> &ndash; Espera</span>
              <span><strong>P</strong> &ndash; Projeto</span>
              <span><strong>V</strong> &ndash; Verificado</span>
              <span><strong>A</strong> &ndash; Aprovado</span>
              <span><strong>W</strong> &ndash; Validado</span>
            </span>
            <span class="legenda-sep"></span>
            <span class="legenda-group">
              <strong>Colunas:</strong>
              <span><strong>C</strong> &ndash; Comprimento</span>
              <span><strong>L</strong> &ndash; Largura</span>
              <span><strong>A</strong> &ndash; Altura</span>
            </span>
          </div>

        </div>
      </div>
    </div>

    <div class="table-card" id="tableCard">
      <div class="table-loading d-none" id="tableLoading">
        <div class="spinner-border text-primary" role="status">
          <span class="visually-hidden">Carregando...</span>
        </div>
        <span>Carregando...</span>
      </div>

      <div class="table-placeholder" id="placeholderInicial">
        <i class="bi bi-clipboard-data"></i>
        <p>Selecione uma filial e clique em <strong>Filtrar</strong> para exibir os resultados.</p>
      </div>

      <div class="table-wrapper d-none" id="tableWrapper">
        <table class="table table-bordered table-striped table-hover table-sm mb-0" id="tblFT">
          <thead id="tblHead"></thead>
          <tbody id="tblBody"></tbody>
        </table>
      </div>

      <div class="table-footer d-none" id="tableFooter">
        <div class="zone-left">
          <label for="selWrapper" class="mb-0" style="color:#6c757d;">Exibir:</label>
          <select id="selWrapper" class="form-select form-select-sm" style="width:auto;">
            <option value="1000" selected>1.000</option>
            <option value="2500">2.500</option>
            <option value="5000">5.000</option>
            <option value="10000">10.000</option>
            <option value="0">Todos</option>
          </select>
        </div>

        <div class="zone-right" id="lblCounter">
          <span id="spanVisivel">&ndash;</span>&nbsp;de&nbsp;<strong id="spanTotal">0</strong>&nbsp;registros
        </div>
      </div>
    </div>

  </div>

  <script>
    const RelatorioFT = (() => {

      // ── Constantes ──────────────────────────────────────────
      const REPOSITORIO = 'RelatorioFTRepositorio.asp';
      const NIVEL_USR   = <%=CInt(0 & Trim(NIVEL))%>;
      const CODSIGA_USR = '<%=Trim(CODUSER)%>';
      const TOAST_DELAY = 4500;
      const TOAST_CFG   = Object.freeze({
        sucesso: { icon: 'bi-check-circle-fill',         title: 'Sucesso'    },
        erro:    { icon: 'bi-x-circle-fill',             title: 'Erro'       },
        aviso:   { icon: 'bi-exclamation-triangle-fill', title: 'Atenção'    },
        info:    { icon: 'bi-info-circle-fill',          title: 'Informação' },
      });
      const SESSION_KEY = 'RelatorioFT_filtros';

      // ── Estado global ────────────────────────────────────────
      let tipoAtual     = '';
      let nivelResp     = NIVEL_USR;
      let cancelado     = false;
      let paginacao     = { pagina: 1, totalCan: 0, maxPag: 1, hasMore: false, pageSize: 500 };
      let secoes        = [];
      let filtrosAtivos = null;
      let tamanhoPag    = 1000;

      // ── Utilitários ──────────────────────────────────────────
      const $el = (id) => $('#' + id);

      const escapeHtml = (s) => $('<div>').text(s == null ? '' : String(s)).html();

      const exibirToast = (msg, tipo) => {
        const t  = TOAST_CFG[tipo] ? tipo : 'info';
        const c  = TOAST_CFG[t];
        const el = document.createElement('div');
        el.className = `toast-custom toast-${t}`;
        el.innerHTML = `
          <div class="toast-inner">
            <i class="bi ${c.icon} toast-icon"></i>
            <div class="toast-body-text">
              <div class="toast-title">${c.title}</div>
              <div class="toast-msg">${escapeHtml(msg)}</div>
            </div>
            <button class="toast-close" onclick="this.closest('.toast-custom').remove()" aria-label="Fechar">&times;</button>
          </div>
          <div class="toast-progress"><div class="toast-progress-bar"></div></div>`;
        document.getElementById('toastMsg').appendChild(el);
        setTimeout(() => { if (el.parentNode) el.remove(); }, TOAST_DELAY);
      };

      // ── Filtros / estado da UI ───────────────────────────────
      const getFiltros = () => ({
        cboPlanta: $el('cboPlanta').val() || '',
        cboStatus: $el('cboStatus').val() || '',
        cboGrupo:  $el('cboGrupo').val()  || '',
        cboTipo:   $el('cboTipo').val()   || '',
        termoCli:  $el('termoCli').val().trim(),
        codcli:    $el('codcli').val().trim(),
        lojacli:   $el('lojaCliVis').val().trim().replace(/^\d+$/, (v) => v.padStart(2, '0')),
        razao:     $el('razao').val().trim(),
      });

      const updateStatus = () => {
        const planta  = $el('cboPlanta').val();
        const current = $el('cboStatus').val();
        const isPirUbe  = (planta === '00' || planta === '04');
        const isTrimbox = (planta === '05' || planta === '06' || planta === '07');

        const optsDefault = [
          { value: '',        label: 'Todos'   },
          { value: 'ATIVO',   label: 'Ativo'   },
          { value: 'INATIVO', label: 'Inativo' },
        ];

        const optsPIRUBE = [
          { value: '',  label: 'Todos'          },
          { value: 'A', label: 'A – Aprovado'   },
          { value: 'C', label: 'C – Cancelado'  },
          { value: 'E', label: 'E – Espera'     },
          { value: 'P', label: 'P – Projeto'    },
          { value: 'V', label: 'V – Verificado' },
          { value: 'W', label: 'W – Validado'   },
        ];

        const opts = isPirUbe ? optsPIRUBE : optsDefault;
        const sel  = $el('cboStatus')[0];
        sel.innerHTML = '';
        opts.forEach(({ value, label }) => {
          const opt       = document.createElement('option');
          opt.value       = value;
          opt.textContent = label;
          if (value === current) opt.selected = true;
          sel.appendChild(opt);
        });

        // Grupo/Tipo só se aplicam a TRIMBOX
        $el('colGrupo').toggleClass('d-none', !isTrimbox);
        $el('colTipo').toggleClass('d-none', !isTrimbox);
        if (!isTrimbox) {
          $el('cboGrupo').val('');
          $el('cboTipo').val('');
        }
      };

      const limparFiltro = () => {
        try { sessionStorage.removeItem(SESSION_KEY); } catch (_) {}
        location.reload();
      };

      // ── Badge de cliente ─────────────────────────────────────
      // modoPopup = true  → popup preencheu código + loja + fantasia (badge completo)
      // modoPopup = false → usuário digitou livremente (badge simples "FILTRANDO POR: x")
      const atualizarBadgeCliente = (cod, loja, razao, modoPopup) => {
        const badgeDyn   = $el('clienteBadgeDyn')[0];
        const badgeTexto = $el('clienteBadgeTexto')[0];
        const semCli     = $el('clienteSemSelecao')[0];

        // Esconde tudo primeiro
        if (badgeDyn)   badgeDyn.classList.add('d-none');
        if (badgeTexto) badgeTexto.classList.add('d-none');
        if (semCli)     semCli.classList.remove('d-none');

        if (!cod) return;

        if (semCli) semCli.classList.add('d-none');

        if (modoPopup) {
          // Popup/código resolvido: exibe código + loja + fantasia
          if (badgeDyn) {
            badgeDyn.classList.remove('d-none');
            badgeDyn.querySelector('.cliente-cod').textContent   = cod;
            badgeDyn.querySelector('.cliente-loja').textContent  = loja || 'Todos';
            badgeDyn.querySelector('.cliente-razao').textContent = razao || '-';
          }
        } else {
          // Digitação livre: exibe só o termo buscado
          if (badgeTexto) {
            badgeTexto.classList.remove('d-none');
            badgeTexto.querySelector('.cliente-cod-texto').textContent = cod;
          }
        }
      };

      // ── Busca de cliente (janela externa) ────────────────────
      const consultaCliente = () => {
        const planta = $el('cboPlanta').val();
        if (!planta) { exibirToast('Selecione a Filial.', 'aviso'); return; }
        const isPirUbe    = (planta === '00' || planta === '04');
        const statusParam = isPirUbe ? '' : $el('cboStatus').val();
        const lnk = `${location.protocol}//${location.host}/corporativo/buscaBICLIENTES.asp` +
                    `?id1=codcli&id2=lojacli&id3=razao&id6=${planta}&id9=${statusParam}`;
        window.open(lnk, 'Consulta Cliente', 'width=650,height=750');
      };

      // ── Combos (Grupo / Tipo) ────────────────────────────────
      const carregarCombos = () => {
        $.ajax({
          url:      REPOSITORIO,
          type:     'POST',
          dataType: 'json',
          data:     { acao: 'obterCombos' },
        }).done((data) => {
          const resp = (typeof data === 'string') ? JSON.parse(data) : data;
          if (!resp || !resp.sucesso) return;
          const grp = $el('cboGrupo')[0];
          (resp.grupos || []).forEach((v) => {
            const o = document.createElement('option');
            o.value = v; o.textContent = v; grp.appendChild(o);
          });
          const tp = $el('cboTipo')[0];
          (resp.tipos || []).forEach((v) => {
            const o = document.createElement('option');
            o.value = v; o.textContent = v; tp.appendChild(o);
          });
        });
      };

      // ── Consulta principal (obterFTs) ────────────────────────
      const carregarFTs = (pagCan) => {
        const f = getFiltros();
        if (!f.cboPlanta) { exibirToast('Selecione a Filial para emitir o Relatório.', 'aviso'); return; }
        if (f.lojacli && !f.termoCli) {
          exibirToast('Informe o Cliente antes de filtrar por Loja.', 'aviso');
          $el('lojaCliVis').val('');
          return;
        }

        filtrosAtivos = f;
        salvarFiltros(f);
        if (pagCan) paginacao.pagina = pagCan;

        $el('tableLoading').removeClass('d-none');

        const payload = {
          acao:      'obterFTs',
          cboPlanta: f.cboPlanta,
          cboStatus: f.cboStatus,
          cboGrupo:  f.cboGrupo,
          cboTipo:   f.cboTipo,
          termoCli:  f.termoCli,
          lojacli:   f.lojacli,
          codsiga:   CODSIGA_USR,
          pagCan:    paginacao.pagina,
          pagSize:   tamanhoPag === 0 ? 99999 : tamanhoPag,
          // Páginas 2+: reenvia o total já conhecido para dispensar o COUNT no servidor.
          // Na página 1 paginacao.totalCan ainda é 0 — o servidor executa o COUNT normalmente.
          totalCan:  paginacao.pagina > 1 ? paginacao.totalCan : 0,
        };

        const isPirUbe = (f.cboPlanta === '00' || f.cboPlanta === '04');

        const exec = () => {
          $.ajax({
            url:      REPOSITORIO,
            type:     'POST',
            dataType: 'text',
            data:     payload,
          }).done((raw) => {
            $el('tableLoading').addClass('d-none');
            let resp;
            try {
              resp = JSON.parse(raw);
            } catch (e) {
              console.error('[RelatorioFT] JSON inválido (primeiros 500 chars):', raw.substring(0, 500));
              render([], 'Erro ao processar resposta do servidor');
              return;
            }
            if (!resp || !resp.sucesso) {
              const msg = resp?.mensagem || 'Erro ao consultar o servidor';
              console.error('[RelatorioFT] sucesso:false —', msg);
              render([], msg);
              return;
            }
            tipoAtual   = resp.tipo;
            nivelResp   = resp.nivel;
            cancelado   = resp.cancelado;
            paginacao   = resp.paginacao;
            secoes      = resp.secoes || [];
            render(secoes, '');
            atualizarCardCliente(resp, f);
          }).fail((xhr) => {
            $el('tableLoading').addClass('d-none');
            console.error('[RelatorioFT] HTTP fail:', xhr.status, xhr.responseText?.substring(0, 300));
            render([], 'Erro ao consultar o servidor');
          });
        };

        // Valida permissão do representante antes de consultar
        // Só aplica quando o termo é numérico (código exato); busca por fantasia não usa SP de validação
        const termoEhCodigo = f.termoCli !== '' && /^\d+$/.test(f.termoCli);
        if (isPirUbe && NIVEL_USR === 1 && termoEhCodigo) {
          $.ajax({
            url:      REPOSITORIO,
            type:     'POST',
            dataType: 'json',
            data:     { acao: 'validarCliente', termoCli: f.termoCli, codsiga: CODSIGA_USR },
          }).done((data) => {
            const resp = (typeof data === 'string') ? JSON.parse(data) : data;
            if (!resp || !resp.sucesso || !resp.permitido) {
              $el('tableLoading').addClass('d-none');
              exibirToast('Você não tem permissão para efetuar consulta deste cliente.', 'erro');
              render([], 'Você não tem permissão para efetuar consulta deste cliente.');
              return;
            }
            exec();
          }).fail(() => {
            $el('tableLoading').addClass('d-none');
            exibirToast('Erro ao validar permissão do representante.', 'erro');
            render([], 'Erro ao validar permissão');
          });
        } else {
          exec();
        }
      };

      const atualizarCardCliente = (resp, filtros) => {
        $el('cardCliente').removeClass('d-none');
        const isPirUbe = (resp.tipo === 'PIRUBE');
        $el('legendaPirUbe').toggleClass('d-none', !isPirUbe);

        if (!filtros.termoCli) {
          atualizarBadgeCliente('', '', '', false);
          return;
        }

        // Tenta extrair código + loja + fantasia dos dados retornados pelo servidor
        let codResp = '', lojaResp = '', fantasiaResp = '';

        if (resp.cliente && resp.cliente.codigo && resp.cliente.codigo.trim()) {
          // Envelope traz cliente resolvido pelo servidor (PIRUBE ou TRIMBOX com filtro de cliente)
          codResp      = resp.cliente.codigo || '';
          fantasiaResp = resp.cliente.nome   || '';
          lojaResp     = resp.cliente.loja   || filtros.lojacli || '';
        } else {
          // Fallback: extrai do primeiro registro da resposta
          const primeiroReg = resp.secoes && resp.secoes[0] && resp.secoes[0].registros && resp.secoes[0].registros[0];
          if (primeiroReg) {
            codResp      = primeiroReg.codcli  || '';
            fantasiaResp = primeiroReg.fantasia || '';
            lojaResp     = primeiroReg.lojacli  || filtros.lojacli || '';
          }
        }

        const temDadosCompletos = !!(codResp && fantasiaResp);
        if (temDadosCompletos) {
          // Servidor resolveu o cliente — badge completo
          atualizarBadgeCliente(codResp, lojaResp, fantasiaResp, true);
        } else {
          // Servidor não encontrou vínculo único (ex: fantasia sem registros ou filtro sem cliente)
          atualizarBadgeCliente(filtros.termoCli, '', '', false);
        }
      };

      // ── Renderização ─────────────────────────────────────────
      const fmtArea = (n, dec) => {
        if (n == null || isNaN(n)) return '0' + (dec > 0 ? ',' + '0'.repeat(dec) : '');
        return Number(n).toLocaleString('pt-BR', { minimumFractionDigits: dec, maximumFractionDigits: dec });
      };

      // fmtPreco: exibe o preço com no mínimo 2 casas decimais e até 10,
      // sem truncar nem arredondar dígitos que venham do banco.
      const fmtPreco = (n) => {
        if (n == null || isNaN(n)) return '0,00';
        return Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 10 });
      };

      const cellCor = (v) => {
        const s = (v == null ? '' : String(v)).trim();
        if (!s || s[0] === ' ') return '<td>-</td>';
        return `<td>${escapeHtml(s)}</td>`;
      };

      const cellEstrutura = (cod, val) => {
        if (nivelResp !== 2 && String(cod).substring(0, 3) === '700') {
          return '<td>&nbsp;</td>';
        }
        return `<td>${escapeHtml(val)}</td>`;
      };

      const cellImg = (r) => {
        if (r.imgExiste) {
          return `<td><a href="javascript:void(0)" class="link-primary" title="Ver imagem"
                     onclick="RelatorioFT.abrirImgFT('${encodeURI(r.urlImg)}','${escapeHtml(r.codigo)}')">
                     <i class="bi bi-image text-primary"></i></a></td>`;
        }
        return `<td><span class="text-secondary" title="Não há imagem" style="opacity:.4;cursor:default">
                  <i class="bi bi-image"></i></span></td>`;
      };

      const montarCabecalho = () => {
        const head       = $el('tblHead')[0];
        const semCliente = !(filtrosAtivos && filtrosAtivos.termoCli);
        const thCli      = semCliente ? '<th>Código</th><th>Loja</th><th>Fantasia</th>' : '';

        if (tipoAtual === 'PIRUBE') {
          const tdPreco = (nivelResp !== 2) ? '<th>Preço</th>'     : '';
          const tdEstr  = (nivelResp !== 2) ? '<th>Estrutura</th>' : '';
          head.innerHTML = `
            <tr>
              <th>Img</th>
              <th>FT</th>
              <th>Referência</th>
              <th>Fabricação</th>
              ${thCli}
              <th>Sap</th>
              ${tdPreco}
              <th>Entrega</th>
              <th>Status</th>
              <th>C</th>
              <th>L</th>
              <th>A</th>
              <th>Área</th>
              ${tdEstr}
              <th>Cor1</th>
              <th>Cor2</th>
              <th>Cor3</th>
              <th>Cor4</th>
            </tr>`;
        } else if (tipoAtual === 'TRIMBOX') {
          head.innerHTML = `
            <tr>
              <th>Img</th>
              <th>FT</th>
              <th>Referência</th>
              <th>Cidade</th>
              ${thCli}
              <th>Sap</th>
              <th>Grupo</th>
              <th>Tipo</th>
              <th>Preço</th>
              <th>Status</th>
              <th>C</th>
              <th>L</th>
              <th>A</th>
              <th>Área</th>
              <th>Estrutura</th>
            </tr>`;
        }
      };

      const montarLinhaPirUbe = (r) => {
        const lnk        = `folha_ft.asp?FT=${encodeURIComponent(r.codigo)}&filial=${encodeURIComponent(r.filial)}`;
        const tdPreco    = (nivelResp !== 2)
          ? `<td class="text-end">${fmtPreco(r.preco)}</td>`
          : '';
        const semCliente = !(filtrosAtivos && filtrosAtivos.termoCli);
        const tdCli      = semCliente
          ? `<td>${escapeHtml(r.codcli)}</td><td>${escapeHtml(r.lojacli)}</td><td class="text-start">${escapeHtml(r.fantasia)}</td>`
          : '';
        return `
          <tr class="ft-data-row">
            ${cellImg(r)}
            <td><a href="${lnk}" class="link-primary text-decoration-none">${escapeHtml(r.codigo)}</a></td>
            <td class="text-start">${escapeHtml(r.descricao)}</td>
            <td>${escapeHtml(r.fabricacao)}</td>
            ${tdCli}
            <td>${escapeHtml(r.sap)}</td>
            ${tdPreco}
            <td>${escapeHtml(r.entrega)}</td>
            <td>${escapeHtml(r.status)}</td>
            <td class="text-end">${escapeHtml(r.compr)}</td>
            <td class="text-end">${escapeHtml(r.larg)}</td>
            <td class="text-end">${escapeHtml(r.altu)}</td>
            <td class="text-end">${fmtArea(r.area, 3)}</td>
            ${cellEstrutura(r.codigo, r.estrutura)}
            ${cellCor(r.cor1)}
            ${cellCor(r.cor2)}
            ${cellCor(r.cor3)}
            ${cellCor(r.cor4)}
          </tr>`;
      };

      const montarLinhaTrimbox = (r) => {
        const planta     = filtrosAtivos ? filtrosAtivos.cboPlanta : '';
        const lnk        = `DetalhesFT.asp?id=${encodeURIComponent(r.codigo)}&planta=${encodeURIComponent(planta)}`;
        const semCliente = !(filtrosAtivos && filtrosAtivos.termoCli);
        const tdCli      = semCliente
          ? `<td>${escapeHtml(r.codcli)}</td><td>${escapeHtml(r.lojacli)}</td><td class="text-start">${escapeHtml(r.fantasia)}</td>`
          : '';
        return `
          <tr class="ft-data-row">
            ${cellImg(r)}
            <td><a href="${lnk}" class="link-primary text-decoration-none" onclick="try{sessionStorage.setItem('ft_origem_ativa','1');}catch(_){}">${escapeHtml(r.codigo)}</a></td>
            <td class="text-start">${escapeHtml(r.referencia)}</td>
            <td>${escapeHtml(r.cidade)}</td>
            ${tdCli}
            <td>${escapeHtml(r.sap)}</td>
            <td>${escapeHtml(r.grupo)}</td>
            <td>${escapeHtml(r.tipo)}</td>
            <td class="text-end">${fmtPreco(r.preco)}</td>
            <td>${escapeHtml(r.status)}</td>
            <td class="text-end">${escapeHtml(r.compr)}</td>
            <td class="text-end">${escapeHtml(r.larg)}</td>
            <td class="text-end">${escapeHtml(r.altu)}</td>
            <td class="text-end">${fmtArea(r.area, 4)}</td>
            <td>${escapeHtml(r.estrutura)}</td>
          </tr>`;
      };

      const getColSpan = () => {
        const extra = (filtrosAtivos && !filtrosAtivos.termoCli) ? 3 : 0;
        if (tipoAtual === 'PIRUBE')  return 18 + extra;
        if (tipoAtual === 'TRIMBOX') return 14 + extra;
        return 14 + extra;
      };

      const render = (listaSecoes, msgErro) => {
        const wrap        = $el('tableWrapper');
        const footer      = $el('tableFooter');
        const placeholder = $el('placeholderInicial');
        const body        = $el('tblBody')[0];

        if (msgErro) {
          wrap.addClass('d-none');
          footer.addClass('d-none');
          placeholder.html(`
            <i class="bi bi-shield-x text-danger"></i>
            <p class="text-danger fw-bold">${escapeHtml(msgErro)}</p>`).removeClass('d-none');
          $el('cardCliente').addClass('d-none');
          return;
        }

        montarCabecalho();

        const colSpan  = getColSpan();
        const totalReg = listaSecoes.reduce((acc, s) => acc + (s.registros ? s.registros.length : 0), 0);

        if (totalReg === 0) {
          wrap.addClass('d-none');
          footer.addClass('d-none');
          body.innerHTML = '';
          const msg = (filtrosAtivos && filtrosAtivos.termoCli)
            ? 'Não existem FTs para este cliente' + (filtrosAtivos.lojacli ? ' / loja informada.' : '.')
            : 'Nenhum registro encontrado.';
          placeholder.html(`
            <i class="bi bi-clipboard-data"></i>
            <p>${escapeHtml(msg)}</p>`).removeClass('d-none');
          return;
        }

        let html = '';
        listaSecoes.forEach((sec) => {
          if (sec.titulo) {
            html += `<tr class="ft-section-header"><td colspan="${colSpan}">${escapeHtml(sec.titulo)}</td></tr>`;
          }
          (sec.registros || []).forEach((r) => {
            html += (tipoAtual === 'PIRUBE') ? montarLinhaPirUbe(r) : montarLinhaTrimbox(r);
          });
        });

        body.innerHTML = html;
        placeholder.addClass('d-none');
        wrap.removeClass('d-none');
        footer.removeClass('d-none');

        // Exibe total geral quando há paginação server-side
        const temPaginacaoSvr = (paginacao && paginacao.totalCan > paginacao.pageSize);
        const totalExibido    = temPaginacaoSvr ? paginacao.totalCan : totalReg;
        $el('spanTotal').text(totalExibido);
        $el('spanVisivel').text(
          tamanhoPag === 0 ? totalExibido : Math.min(tamanhoPag, totalReg)
        );
      };

      // ── Modal da imagem ──────────────────────────────────────
      const abrirImgFT = (url, codigo) => {
        const modal   = new bootstrap.Modal(document.getElementById('modalImgFT'));
        const img     = document.getElementById('imgFTSrc');
        const loading = document.getElementById('imgFTLoading');
        const error   = document.getElementById('imgFTError');
        const label   = document.getElementById('modalImgFTLabel');
        const btnDl   = document.getElementById('btnDownloadImgFT');

        label.innerHTML = `<i class="bi bi-image me-2"></i>Imagem da FT &ndash; ${escapeHtml(codigo)}`;
        img.classList.add('d-none');
        error.classList.add('d-none');
        loading.classList.remove('d-none');
        if (btnDl) btnDl.classList.add('d-none');
        img.src = '';

        modal.show();

        img.onload = () => {
          loading.classList.add('d-none');
          img.classList.remove('d-none');
          if (btnDl) {
            btnDl.classList.remove('d-none');
            btnDl.onclick = () => {
              const a    = document.createElement('a');
              a.href     = url;
              a.download = 'FT_' + codigo + '.jpg';
              a.target   = '_blank';
              document.body.appendChild(a);
              a.click();
              document.body.removeChild(a);
            };
          }
        };
        img.onerror = () => {
          loading.classList.add('d-none');
          img.classList.add('d-none');
          error.classList.remove('d-none');
        };
        img.src = url;
      };

      // ── Observador do retorno da janela buscaBICLIENTES ──────
      // O popup preenche #codcli, #lojacli e #razao via atribuição direta de .value (cross-window).
      // MutationObserver não detecta mudanças de .value em inputs — usamos Object.defineProperty
      // para interceptar as três escritas e espelhar nos campos visíveis.
      const instalarObservadorCliente = () => {
        const fTermo  = $el('termoCli')[0];
        const fCod    = $el('codcli')[0];
        const fLoja   = $el('lojacli')[0];
        const fRazao  = $el('razao')[0];
        const fLojaVis = document.getElementById('lojaCliVis');
        if (!fTermo) return;

        // Digitação livre em #termoCli: limpa hiddens e o campo visível de loja (cancela seleção anterior do popup)
        fTermo.addEventListener('input', () => {
          if (fCod)    fCod.value    = '';
          if (fLoja)   fLoja.value   = '';
          if (fRazao)  fRazao.value  = '';
          if (fLojaVis) fLojaVis.value = '';
        });

        const origDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');

        // Intercepta #codcli → espelha código em #termoCli
        let _valCod = fCod ? fCod.value : '';
        if (fCod) {
          Object.defineProperty(fCod, 'value', {
            get() { return _valCod; },
            set(v) {
              origDesc.set.call(this, v);
              _valCod = v;
              if (v && fTermo) {
                origDesc.set.call(fTermo, v);
              }
            },
            configurable: true,
          });
        }

        // Intercepta #lojacli → espelha loja em #lojaCliVis
        let _valLoja = fLoja ? fLoja.value : '';
        if (fLoja) {
          Object.defineProperty(fLoja, 'value', {
            get() { return _valLoja; },
            set(v) {
              origDesc.set.call(this, v);
              _valLoja = v;
              if (fLojaVis) {
                origDesc.set.call(fLojaVis, v);
              }
            },
            configurable: true,
          });
        }

        // Intercepta #razao → apenas armazena o valor preenchido pelo popup.
        // Não dispara consulta automática; o usuário decide quando clicar em Filtrar.
        let _valRazao = fRazao ? fRazao.value : '';
        if (fRazao) {
          Object.defineProperty(fRazao, 'value', {
            get() { return _valRazao; },
            set(v) {
              origDesc.set.call(this, v);
              _valRazao = v;
            },
            configurable: true,
          });
        }
      };

      // ── Persistência de filtros (sessionStorage) ─────────────
      const salvarFiltros = (f) => {
        try { sessionStorage.setItem(SESSION_KEY, JSON.stringify(f)); } catch (_) {}
      };

      const restaurarFiltros = () => {
        const flagAtiva = sessionStorage.getItem('ft_origem_ativa') === '1';
        sessionStorage.removeItem('ft_origem_ativa');

        if (!flagAtiva) {
          try { sessionStorage.removeItem(SESSION_KEY); } catch (_) {}
          return false;
        }

        let salvo;
        try { salvo = JSON.parse(sessionStorage.getItem(SESSION_KEY)); } catch (_) {}
        if (!salvo || !salvo.cboPlanta) return false;

        $el('cboPlanta').val(salvo.cboPlanta);
        updateStatus();
        $el('cboStatus').val(salvo.cboStatus || '');
        $el('cboGrupo').val(salvo.cboGrupo   || '');
        $el('cboTipo').val(salvo.cboTipo     || '');
        $el('termoCli').val(salvo.termoCli   || '');
        $el('codcli').val(salvo.codcli       || '');
        $el('lojacli').val(salvo.lojacli     || '');
        $el('lojaCliVis').val(salvo.lojacli  || '');
        $el('razao').val(salvo.razao         || '');

        return true;
      };

      // ── Centralização de eventos ─────────────────────────────
      const anexarEventos = () => {
        $el('cboPlanta').on('change', updateStatus);
        $el('btnFiltrar').on('click',         () => carregarFTs(1));
        $el('btnLimpar').on('click',          limparFiltro);
        $el('btnConsultaCliente').on('click', consultaCliente);

        $el('selWrapper').on('change', function () {
          tamanhoPag = parseInt(this.value, 10);
          carregarFTs(1);
        });
      };

      // ── Inicialização ────────────────────────────────────────
      const init = () => {
        carregarCombos();
        instalarObservadorCliente();
        anexarEventos();

        if (restaurarFiltros()) {
          carregarFTs(1);
        } else {
          updateStatus();
        }
      };

      return Object.freeze({ init, abrirImgFT });
    })();

    $(document).ready(() => RelatorioFT.init());
  </script>

</body>
</html>
