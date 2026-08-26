<%@ Language="VBScript" %>
<!--#include virtual="/corporativo/cnn2/cnn2.asp"-->
<!--#include virtual="/corporativo/funcoes/funcoes.asp"-->
<!--#include virtual="/corporativo/includes/variaveis.asp"-->
<%
  Session.LCID = 1046

  If CInt(NIVEL) > 0 Then
%>
<script>
  alert("*** Acesso Negado.\nVoce nao tem permissao para esse modulo");
  window.close();
</script>
<%
    Response.End
  End If

  Dim swnet_nivel   : swnet_nivel   = Session("NIVEL")
  Dim swnet_coduser : swnet_coduser = Session("CODUSER")
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="PRAGMA" content="no-cache" />
  <meta http-equiv="cache-control" content="no-cache" />
  <title>SWnet - Alterar Risco Cliente</title>

  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />

  <style>
    body {
      background-color: #f0f0f0;
      min-height: 100vh;
    }

    input[readonly] {
      background-color: #f8f9fa;
    }

    .field-hint {
      font-size: .75rem;
      color: #6c757d;
      margin-top: .2rem;
    }

    /* ── Toast ─────────────────────────────────────────────────────────── */
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

    .toast-custom .toast-progress-bar {
      height: 3px;
      animation: toastProgressShrink 4.5s linear forwards;
    }

    @keyframes toastProgressShrink {
      from { width: 100%; }
      to   { width: 0%; }
    }

    /* ── Modal de busca de clientes ─────────────────────────────────────── */
    #modalBuscaCliente .modal-dialog {
      max-width: 820px;
    }

    #tblBuscaCliente {
      font-size: .8rem;
    }

    #tblBuscaCliente thead th {
      background: #fff;
      color: #212529;
      font-weight: 600;
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    #tblBuscaCliente tbody tr {
      cursor: pointer;
      transition: background .15s;
    }

    #tblBuscaCliente tbody tr:hover td {
      background: #e8f0fe;
    }

    .btn-selecionar-cli {
      background: none;
      border: none;
      padding: 0 4px;
      color: #0d6efd;
      cursor: pointer;
      font-size: .95rem;
    }

    .btn-selecionar-cli:hover {
      color: #0a58ca;
    }

    .modal-busca-scroll {
      max-height: 340px;
      overflow-y: auto;
    }

    #busca-spinner {
      display: none;
    }
  </style>
</head>

<body onload="AlterarRisco.init()">

  <!-- Toast container -->
  <div class="toast-container-custom" id="toastContainer"></div>

  <div class="container-fluid py-3 px-3 px-lg-4">

    <div class="row justify-content-center">
      <div class="col-12 col-lg-10 col-xl-9">

        <!-- Cabeçalho -->
        <div class="d-flex align-items-center mb-3">
          <span class="page-title fw-bold" style="font-size:1.15rem;">Alterar Risco Cliente</span>
        </div>

        <!-- Card principal -->
        <div class="card border-0 shadow-sm position-relative">
          <div class="card-header border-bottom" style="background:#fff;">
            <span class="fw-semibold" style="font-size:.9rem;">Dados do Cliente</span>
          </div>
          <div class="card-body">

            <!-- Loading overlay do card -->
            <div id="loadingOverlay" class="d-none position-absolute top-0 start-0 w-100 h-100 align-items-center justify-content-center" style="background:rgba(255,255,255,.7);z-index:10;border-radius:inherit;">
              <div class="spinner-border text-secondary" role="status"></div>
            </div>

            <input id="hidCodCliente" type="hidden" value="" />
            <input id="hidRiscoSelecionado" type="hidden" value="0" />

            <!-- Linha 1: Planta + Código do Cliente + Loja -->
            <div class="row mb-1 g-2 align-items-end">

              <div class="col-12 col-md-4">
                <label for="Planta" class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Planta</label>
                <select id="Planta" class="form-select form-select-sm">
                  <option value="">Selecione...</option>
                  <option value="05">Bento Gonçalves</option>
                  <option value="06">Maranguape</option>
                  <option value="00">Pirapetinga</option>
                  <option value="07">Saquarema</option>
                  <option value="04">Uberaba</option>
                </select>
              </div>

              <div class="col-12 col-md-4">
                <label for="codcli" class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Código do Cliente</label>
                <div class="input-group input-group-sm">
                  <button type="button" class="btn btn-secondary" id="btnPesquisaCliente" title="Pesquisar cliente">
                    <i class="bi bi-box-arrow-up-right me-1"></i>Buscar
                  </button>
                  <input id="codcli" type="text" class="form-control" placeholder="Código..." />
                </div>
              </div>

              <div class="col-12 col-md-4">
                <label for="codloja" class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Loja</label>
                <div class="input-group input-group-sm">
                  <input id="codloja" type="text" class="form-control" maxlength="2" placeholder="Nº..." />
                  <button type="button" class="btn btn-warning" id="btnFiltrar" title="Filtrar">
                    <i class="bi bi-funnel-fill"></i>
                  </button>
                </div>
              </div>

            </div>

            <!-- Legenda Loja -->
            <div class="row g-2 mb-3">
              <div class="col-12 col-md-4 offset-md-8">
                <div class="field-hint">Informe a Loja para alterar apenas uma loja específica.</div>
              </div>
            </div>

            <!-- Linha 2: Razão Social + Risco Atual + Alterar Risco -->
            <div class="row mb-4 g-2 align-items-end">

              <div class="col-12 col-md-4">
                <label for="nomecli" class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Razão Social</label>
                <input id="nomecli" type="text" class="form-control form-control-sm" readonly />
              </div>

              <div class="col-12 col-md-4">
                <label for="riscoatual" class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Risco Atual</label>
                <input id="riscoatual" type="text" class="form-control form-control-sm" readonly />
              </div>

              <div class="col-12 col-md-4">
                <label class="form-label form-label-sm fw-semibold text-secondary" style="font-size:.82rem;">Alterar o Risco</label>
                <div id="divRisco"></div>
              </div>

            </div>

            <!-- Botões -->
            <div class="d-flex gap-2 justify-content-end">
              <button type="button" class="btn btn-light btn-sm" id="btnLimpar">
                <i class="bi bi-eraser me-2"></i>Limpar
              </button>
              <button type="button" class="btn btn-success btn-sm" id="btnConfirmar" title="Alterar o Risco">
                Confirmar Alteração
              </button>
            </div>

          </div>
        </div>

      </div>
    </div>

  </div>

  <!-- ════════════════════════════════════════════════════════════════════════
       Modal de Confirmação de Alteração
  ════════════════════════════════════════════════════════════════════════ -->
  <div class="modal fade" id="modalConfirmacao" tabindex="-1" aria-labelledby="modalConfirmacaoLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header border-bottom">
          <h6 class="modal-title" id="modalConfirmacaoLabel">Confirmar Alteração</h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <p id="modalConfirmacaoMsg" class="mb-0" style="font-size:.9rem;"></p>
        </div>
        <div class="modal-footer border-top">
          <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Cancelar</button>
          <button type="button" class="btn btn-success btn-sm" id="btnConfirmarModal">Confirmar</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ════════════════════════════════════════════════════════════════════════
       Modal de Busca de Clientes (substitui o popup externo)
  ════════════════════════════════════════════════════════════════════════ -->
  <div class="modal fade" id="modalBuscaCliente" tabindex="-1" aria-labelledby="modalBuscaClienteLabel" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <div class="modal-content">

        <div class="modal-header border-bottom">
          <h6 class="modal-title" id="modalBuscaClienteLabel">Buscar Cliente</h6>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>

        <div class="modal-body pb-2">

          <!-- Campo de pesquisa -->
          <div class="d-flex gap-2 mb-3">
            <input type="text"
                   id="inputBuscaCliente"
                   class="form-control form-control-sm"
                   placeholder="Digite nome ou código do cliente..."
                   autocomplete="off" />
            <button type="button" class="btn btn-primary btn-sm flex-shrink-0" id="btnExecutarBusca">
              <i class="bi bi-funnel-fill me-1"></i>Filtrar
            </button>
          </div>

          <!-- Spinner -->
          <div id="busca-spinner" class="text-center py-3">
            <div class="spinner-border spinner-border-sm text-secondary" role="status"></div>
            <span class="ms-2 text-secondary" style="font-size:.85rem;">Buscando...</span>
          </div>

          <!-- Tabela de resultados -->
          <div class="modal-busca-scroll">
            <table class="table table-bordered table-hover table-sm mb-0" id="tblBuscaCliente">
              <thead>
                <tr>
                  <th style="width:32px;"></th>
                  <th>Código</th>
                  <th>Loja</th>
                  <th>Nome</th>
                  <th>CNPJ</th>
                  <th>Município</th>
                  <th>UF</th>
                </tr>
              </thead>
              <tbody id="tbodyBuscaCliente">
                <tr>
                  <td colspan="7" class="text-center text-muted py-3" style="font-size:.82rem;">
                    <i class="bi bi-clipboard-data me-1"></i>
                    Digite um nome ou código e clique em Filtrar.
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

        </div>

        <div class="modal-footer border-top py-2">
          <span id="busca-contador" class="text-muted me-auto" style="font-size:.78rem;"></span>
          <button type="button" class="btn btn-light btn-sm" data-bs-dismiss="modal">Fechar</button>
        </div>

      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

  <script>
    window.SWNET_NIVEL   = '<%= swnet_nivel %>';
    window.SWNET_CODUSER = '<%= swnet_coduser %>';

    const AlterarRisco = (function () {

      // ── Constantes ──────────────────────────────────────────────────────
      const REPOSITORIO = 'alterar_riscoRepositorio.asp';
      const TOAST_DELAY = 4500;

      const TOAST_ICON = {
        sucesso : { icon: 'bi-check-circle-fill',         cor: '#198754', barra: '#198754' },
        erro    : { icon: 'bi-x-circle-fill',             cor: '#dc3545', barra: '#dc3545' },
        aviso   : { icon: 'bi-exclamation-triangle-fill', cor: '#856404', barra: '#ffc107' },
        info    : { icon: 'bi-info-circle-fill',          cor: '#6c757d', barra: '#6c757d' }
      };

      // ── Referência ao modal de busca ─────────────────────────────────────
      let _modalBusca = null;

      // ── Atalho ──────────────────────────────────────────────────────────
      const $el = (id) => $(`#${id}`);

      // ── Init ────────────────────────────────────────────────────────────
      function init() {
        _iniciaComboRiscos();

        _modalBusca = new bootstrap.Modal(document.getElementById('modalBuscaCliente'));

        $el('Planta').on('change', function () {
          _limpaCampos('Planta');
          _comboRiscos();
          // Maranguape não usa loja — desabilita e zera o campo
          if ($(this).val() === '06') {
            _tratarCampoLoja(true);
          }
        });

        $el('codcli').on('focus', function () {
          $el('nomecli').val('');
          $el('riscoatual').val('');
        });

        $el('codloja').on('blur', function () {
          if ($el('nomecli').val() === '') {
            _comboRiscos();
          }
        });

        $el('btnPesquisaCliente').on('click', _abrirModalBusca);
        $el('btnFiltrar').on('click', _carregaCliente);
        $el('btnLimpar').on('click', function () { _limpaCampos(); });
        $el('btnConfirmar').on('click', _gravarSolicitacao);

        // ── Botão Filtrar no modal de busca ──────────────────────────────
        $el('btnExecutarBusca').on('click', _executarBuscaModal);

        // Enter no campo de busca do modal dispara a pesquisa
        $el('inputBuscaCliente').on('keydown', function (e) {
          if (e.key === 'Enter') _executarBuscaModal();
        });

        // Limpa o modal ao fechar
        document.getElementById('modalBuscaCliente').addEventListener('hidden.bs.modal', function () {
          $el('inputBuscaCliente').val('');
          $el('tbodyBuscaCliente').html(
            '<tr><td colspan="7" class="text-center text-muted py-3" style="font-size:.82rem;">' +
            '<i class="bi bi-clipboard-data me-1"></i>Digite um nome ou código e clique em Filtrar.' +
            '</td></tr>'
          );
          $el('busca-contador').text('');
        });

        // Foca o campo ao abrir o modal
        document.getElementById('modalBuscaCliente').addEventListener('shown.bs.modal', function () {
          document.getElementById('inputBuscaCliente').focus();
        });
      }

      // ── Abre o modal de busca ────────────────────────────────────────────
      function _abrirModalBusca() {
        const planta = $el('Planta').val();
        if (planta === '') {
          _toast('aviso', 'Selecione a Planta antes de buscar um cliente.');
          return;
        }
        _modalBusca.show();
      }

      // ── AJAX: executa a busca de clientes no modal ───────────────────────
      function _executarBuscaModal() {
        const planta    = $el('Planta').val();
        const parametro = $el('inputBuscaCliente').val().trim();

        if (parametro.length < 2) {
          _toast('aviso', 'Digite ao menos 2 caracteres para pesquisar.');
          return;
        }

        // Exibe spinner, oculta contador
        document.getElementById('busca-spinner').style.display = 'block';
        $el('busca-contador').text('');
        $el('tbodyBuscaCliente').html(
          '<tr><td colspan="7" class="text-center py-3"><span class="text-muted" style="font-size:.82rem;">Pesquisando...</span></td></tr>'
        );

        $.ajax({
          url      : REPOSITORIO,
          method   : 'GET',
          dataType : 'json',
          cache    : false,
          data     : { acao: 'buscarClientes', Planta: planta, Parametro: parametro },
          success  : function (resp) {
            document.getElementById('busca-spinner').style.display = 'none';
            if (!resp.sucesso) {
              $el('tbodyBuscaCliente').html(
                '<tr><td colspan="7" class="text-center text-danger py-3" style="font-size:.82rem;">' +
                '<i class="bi bi-x-circle me-1"></i>' + (resp.erro || 'Erro ao buscar clientes.') +
                '</td></tr>'
              );
              return;
            }
            _renderizarResultados(resp.dados);
          },
          error: function (xhr) {
            document.getElementById('busca-spinner').style.display = 'none';
            $el('tbodyBuscaCliente').html(
              '<tr><td colspan="7" class="text-center text-danger py-3" style="font-size:.82rem;">' +
              '<i class="bi bi-x-circle me-1"></i>Erro de comunicação com o servidor.' +
              '</td></tr>'
            );
            console.error(xhr.responseText);
          }
        });
      }

      // ── Renderiza os resultados na tabela do modal ───────────────────────
      function _renderizarResultados(dados) {
        if (!dados || dados.length === 0) {
          $el('tbodyBuscaCliente').html(
            '<tr><td colspan="7" class="text-center text-muted py-3" style="font-size:.82rem;">' +
            '<i class="bi bi-exclamation-circle me-1"></i>Nenhum cliente encontrado.' +
            '</td></tr>'
          );
          $el('busca-contador').text('0 registros');
          return;
        }

        const plantaModal = $el('Planta').val();
        const usaCodext   = (plantaModal === '05' || plantaModal === '06');

        let html = '';
        for (const cli of dados) {
          // Escapa aspas simples para uso no onclick inline
          const cod  = (cli.cod  || '').replace(/'/g, "\\'");
          const loja = (cli.loja || '').replace(/'/g, "\\'");
          const nome = (cli.nome || '').replace(/'/g, "\\'");

          // Para Bento/Maranguape exibe CODIGO_ERP na coluna Código (visualmente),
          // mas cod (COD_CLIENTE) é o que vai para _selecionarCliente e para o SA1010.
          const codExibido = usaCodext ? _esc(cli.codext || cli.cod) : _esc(cli.cod);

          html +=
            '<tr onclick="AlterarRisco._selecionarCliente(\'' + cod + '\',\'' + loja + '\',\'' + nome + '\')">' +
              '<td class="text-center">' +
                '<button type="button" class="btn btn-primary btn-sm py-0 px-1" title="Selecionar">' +
                  '<i class="bi bi-download"></i>' +
                '</button>' +
              '</td>' +
              '<td>' + codExibido         + '</td>' +
              '<td>' + _esc(cli.loja)     + '</td>' +
              '<td>' + _esc(cli.nome)     + '</td>' +
              '<td>' + _esc(cli.cnpj)     + '</td>' +
              '<td>' + _esc(cli.municipio)+ '</td>' +
              '<td>' + _esc(cli.estado)   + '</td>' +
            '</tr>';
        }
        $el('tbodyBuscaCliente').html(html);

        const sufixo = dados.length === 100 ? ' (máximo atingido)' : '';
        $el('busca-contador').text(dados.length + ' registro(s)' + sufixo);
      }

      // ── Chamado ao clicar em uma linha do resultado ──────────────────────
      function _selecionarCliente(cod, loja, nome) {
        const planta = $el('Planta').val();
        $el('codcli').val(cod);
        // Maranguape não usa loja — nunca popula o campo
        $el('codloja').val(planta === '06' ? '' : loja);
        $el('nomecli').val('');
        $el('riscoatual').val('');
        _limpaCombo();
        _iniciaComboRiscos();

        _modalBusca.hide();

        // Carrega nome e risco atual via AJAX
        setTimeout(_carregaCliente, 150);
      }

      // ── AJAX: carrega nome e risco atual do cliente ──────────────────────
      function _carregaCliente() {
        const codPlanta = $el('Planta').val();
        const codcli    = $el('codcli').val();
        const codloja   = $el('codloja').val();

        if (codPlanta === '' || codcli === '') {
          _toast('aviso', 'Planta ou Cliente não informado!');
          return;
        }

        $.ajax({
          url      : REPOSITORIO,
          method   : 'GET',
          dataType : 'json',
          cache    : false,
          data     : { acao: 'carregarCliente', Planta: codPlanta, Cliente: codcli, Loja: codloja },
          success  : function (data) {
            const resp = (typeof data === 'string') ? JSON.parse(data) : data;
            if (!resp.sucesso) {
              _limpaCampos('Planta');
              _toast('erro', resp.erro || 'Cliente não encontrado. Verifique os dados informados!');
              return;
            }
            $el('nomecli').val(resp.dados.nomeCliente);
            $el('riscoatual').val(resp.dados.riscoAtual);
            _comboRiscos();
          },
          error: function (xhr) {
            _toast('erro', 'Erro ao carregar dados do cliente.');
            console.error(xhr.responseText);
          }
        });
      }

      // ── AJAX: grava a alteração de risco ─────────────────────────────────
      function _gravarSolicitacao() {
        const codPlanta     = $el('Planta').val();
        const codcli        = $el('codcli').val();
        const codloja       = $el('codloja').val();
        const nomecli       = $el('nomecli').val();
        const riscoanterior = $el('riscoatual').val();

        const cboEl    = document.getElementById('cboRisco');
        const cborisco = cboEl ? cboEl.value : '';

        const erros = [];
        if (codPlanta === '') erros.push('Planta não informada.');
        if (codcli    === '') erros.push('Código do Cliente não informado.');
        if (nomecli   === '') erros.push('Razão Social não informada. Clique no botão Filtrar.');
        if (cborisco  === '0' || cborisco === '') erros.push('Risco não informado.');

        if (erros.length > 0) {
          erros.forEach(msg => _toast('aviso', msg));
          return;
        }

        const msgConfirm = (codloja === '' && codPlanta !== '06')
          ? 'Deseja alterar o risco de <strong>TODAS as lojas</strong> deste cliente?'
          : 'Deseja alterar o risco deste cliente?';

        $el('modalConfirmacaoMsg').html(msgConfirm);

        const modal = new bootstrap.Modal(document.getElementById('modalConfirmacao'));
        modal.show();

        $el('btnConfirmarModal').off('click').on('click', function () {
          modal.hide();
          _loading(true);

          $.ajax({
            url      : REPOSITORIO,
            method   : 'POST',
            dataType : 'json',
            cache    : false,
            data     : {
              acao          : 'gravarRisco',
              Planta        : codPlanta,
              Cliente       : codcli,
              Loja          : codloja,
              Risco         : cborisco,
              RiscoAnterior : riscoanterior
            },
            success: function (resp) {
              if (resp.sucesso) {
                _toast('sucesso', 'Risco atualizado com sucesso!');
                _limpaCampos();
              } else {
                _toast('erro', 'Erro ao gravar: ' + (resp.erro || 'erro desconhecido.'));
              }
            },
            error: function (xhr) {
              _toast('erro', 'Erro ao gravar a alteração de risco.');
              console.error(xhr.responseText);
            },
            complete: function () {
              _loading(false);
            }
          });
        });
      }

      // ── Combos de risco ──────────────────────────────────────────────────
      function _comboRiscos(argFlag) {
        if (argFlag) _limpaCampos('Planta');
        const planta = $el('Planta').val();
        switch (planta) {
          case '00': case '04': case '05': case '06': case '07': _riscos1(); break;
        }
      }

      function _riscos1() {
        _tratarCampoLoja(false);
        const riscocliente = $el('riscoatual').val();
        const codigoloja   = $el('codloja').val();

        let opcoes;
        if (riscocliente === null || riscocliente === '') {
          opcoes = [
            { value: 'A', text: 'RISCO A' },
            { value: 'E', text: 'RISCO E' }
          ];
        } else if (codigoloja === '') {
          opcoes = [
            { value: 'A', text: 'RISCO A' },
            { value: 'B', text: 'RISCO B' },
            { value: 'C', text: 'RISCO C' },
            { value: 'D', text: 'RISCO D' },
            { value: 'E', text: 'RISCO E' },
            { value: 'F', text: 'RISCO F - Bloqueia só Faturamento' },
            { value: 'P', text: 'RISCO P - Bloqueia Novos Pedidos' }
          ];
        } else {
          opcoes = [
            { value: 'J', text: 'RISCO J - Bloqueia Novos Pedidos desta loja' },
            { value: 'A', text: 'RISCO A' }
          ];
        }
        _criaComboRisco(opcoes);
      }

      function _criaComboRisco(opcoes) {
        _limpaCombo();

        const cbo = document.createElement('select');
        cbo.classList.add('form-select', 'form-select-sm');
        cbo.name = 'cboRisco';
        cbo.id   = 'cboRisco';

        const optDefault = document.createElement('option');
        optDefault.value = '0';
        optDefault.text  = 'Selecione...';
        cbo.appendChild(optDefault);

        for (const opcao of opcoes) {
          const opt = document.createElement('option');
          opt.value = opcao.value;
          opt.text  = opcao.text;
          cbo.appendChild(opt);
        }

        cbo.addEventListener('change', function () {
          $el('hidRiscoSelecionado').val(this.value);
        });

        document.getElementById('divRisco').appendChild(cbo);
      }

      function _iniciaComboRiscos() {
        _limpaCombo();
        const cbo = document.createElement('select');
        cbo.classList.add('form-select', 'form-select-sm');
        cbo.name = 'cboRisco';
        cbo.id   = 'cboRisco';
        const opt = document.createElement('option');
        opt.value = '0';
        opt.text  = 'Selecione...';
        cbo.appendChild(opt);
        document.getElementById('divRisco').appendChild(cbo);
      }

      function _limpaCombo() {
        const el = document.getElementById('cboRisco');
        if (el) {
          el.remove();
          $el('hidRiscoSelecionado').val('');
        }
      }

      function _tratarCampoLoja(desabilitar) {
        const campoLoja = document.getElementById('codloja');
        if (desabilitar) {
          campoLoja.value = '';
          campoLoja.setAttribute('readonly', '');
        } else {
          campoLoja.removeAttribute('readonly');
        }
      }

      function _limpaCampos(exceto) {
        if (exceto !== 'Planta') $el('Planta').val('');
        $el('codcli').val('');
        $el('codloja').val('');
        $el('nomecli').val('');
        $el('riscoatual').val('');
        $el('hidCodCliente').val('');
        _limpaCombo();
        _iniciaComboRiscos();
        if (exceto !== 'Planta') {
          const campo = document.getElementById('codloja');
          if (campo) campo.removeAttribute('readonly');
        }
      }

      // ── Helpers ──────────────────────────────────────────────────────────
      function _loading(ativo) {
        const overlay = document.getElementById('loadingOverlay');
        if (ativo) {
          overlay.classList.remove('d-none');
          overlay.classList.add('d-flex');
          $el('btnConfirmar').prop('disabled', true);
          $el('btnFiltrar').prop('disabled', true);
        } else {
          overlay.classList.remove('d-flex');
          overlay.classList.add('d-none');
          $el('btnConfirmar').prop('disabled', false);
          $el('btnFiltrar').prop('disabled', false);
        }
      }

      function _toast(tipo, mensagem) {
        const cfg = TOAST_ICON[tipo] || TOAST_ICON.info;
        const el  = document.createElement('div');
        el.className = 'toast-custom';
        el.innerHTML = `
          <div class="d-flex align-items-start p-3 gap-2">
            <i class="bi ${cfg.icon} mt-1" style="color:${cfg.cor};font-size:1.1rem;flex-shrink:0;"></i>
            <span style="font-size:.875rem;color:#212529;flex:1;">${mensagem}</span>
            <button type="button" style="background:none;border:none;cursor:pointer;padding:0;line-height:1;color:#6c757d;font-size:1.1rem;" onclick="this.closest('.toast-custom').remove()">&times;</button>
          </div>
          <div class="toast-progress-bar" style="background:${cfg.barra};"></div>`;
        document.getElementById('toastContainer').appendChild(el);
        setTimeout(() => el.remove(), TOAST_DELAY + 300);
      }

      // Escapa HTML para exibição segura na tabela
      function _esc(str) {
        if (!str) return '';
        return String(str)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      }

      return Object.freeze({ init, _selecionarCliente });

    }());
  </script>

</body>
</html>
