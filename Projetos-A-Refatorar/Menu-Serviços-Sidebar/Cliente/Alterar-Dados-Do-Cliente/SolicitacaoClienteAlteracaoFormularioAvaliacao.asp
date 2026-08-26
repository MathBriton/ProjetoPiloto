<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../cnn2/cnn2.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<!--#Include File="../includes/variaveis.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link
        href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC"
        crossorigin="anonymous"
    />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }

        .container {
            max-width: 85%;
        }

        .page-actions {
            gap: .5rem;
        }

        .btn-back {
            font-size: 1.25rem;
        }

        .loading-overlay {
            align-items: center;
            background: rgba(255, 255, 255, .8);
            height: 100%;
            justify-content: center;
            left: 0;
            position: fixed;
            top: 0;
            width: 100%;
            z-index: 9999;
        }

        .form-container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, .1);
            overflow: hidden;
        }

        .form-layout {
            display: grid;
            grid-template-columns: 270px minmax(0, 1fr);
        }

        .form-label {
            font-weight: 600;
        }

        .form-content {
            overflow: hidden;
        }

        .required-field {
            color: #dc3545;
        }

        .tabs-container {
            background-color: #f8f9fa;
            border-right: 1px solid #e9ecef;
            overflow-y: auto;
        }

        .tabs {
            display: flex;
            flex-direction: column;
            gap: 3px;
            padding: 14px 10px;
        }

        .tab-group-title {
            color: #9aa3af;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .06em;
            margin: 14px 10px 6px;
            text-transform: uppercase;
        }

        .tab-group-title:first-child {
            margin-top: 0;
        }

        .tab-button {
            align-items: center;
            background: transparent;
            border: none;
            border-radius: 7px;
            color: #495057;
            cursor: pointer;
            display: flex;
            font-weight: 500;
            gap: 9px;
            justify-content: flex-start;
            padding: 9px 11px;
            text-align: left;
            transition: all .2s ease;
            white-space: nowrap;
            width: 100%;
        }

        .tab-button:hover {
            background-color: #eef2f7;
        }

        .tab-button.active {
            background-color: #e7f1ff;
            color: #0d6efd;
            font-weight: 700;
        }

        .tab-button.completed {
            color: #198754;
        }

        .tab-button.invalid {
            background-color: #fff5f5;
            color: #dc3545;
        }

        .tab-button:disabled {
            cursor: not-allowed;
            opacity: .5;
        }

        .tab-icon {
            flex-shrink: 0;
            font-size: 16px;
            width: 18px;
        }

        .tab-title {
            flex: 1;
            font-size: 14px;
        }

        .tab-meta {
            background: #e9ecef;
            border-radius: 999px;
            color: #6c757d;
            font-size: 10px;
            font-weight: 700;
            margin-left: auto;
            padding: 1px 7px;
        }

        .tab-button.active .tab-meta {
            background: #d7eaff;
            color: #0d6efd;
        }

        .tab-content-wrapper {
            background: #fff;
            min-width: 0;
        }

        .tab-content {
            animation: fadeIn .25s ease;
            display: none;
            height: 80vh;
            overflow: auto;
            padding: 30px;
        }

        .tab-content.active {
            display: block;
        }

        .section-header {
            border-bottom: 1px solid #e9ecef;
            margin-bottom: 22px;
            padding-bottom: 14px;
        }

        .section-title {
            color: #343a40;
            font-size: 1.05rem;
            font-weight: 700;
            margin-bottom: 3px;
        }

        .section-subtitle {
            color: #6c757d;
            font-size: .825rem;
            margin: 0;
        }

        .readonly-alert {
            background: #f8f9fa;
            border: 1px dashed #ced4da;
            border-radius: 8px;
            color: #6c757d;
            padding: 12px 14px;
        }

        .nav-footer {
            background: #fff;
            border-top: 1px solid #dee2e6;
            bottom: 0;
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            padding: 10px 20px;
            position: sticky;
            z-index: 10;
        }

        #btnVoltarPagina:disabled,
        #btnProximaPagina:disabled {
            cursor: not-allowed;
            pointer-events: auto;
        }

        .just-validate-error-label {
            color: #f87171 !important;
            display: block;
            font-size: .8rem;
            margin-top: 3px;
        }

        .just-validate-error-field {
            border-color: #f87171 !important;
        }

        th,
        td {
            vertical-align: top;
            white-space: nowrap;
        }

        .nav-tabs .contador-grupo {
            font-size: .68rem;
            font-weight: 600;
        }

        .campo-cabecalho,
        .campo-alteracao-linha {
            align-items: center;
            column-gap: .75rem;
            display: grid;
            grid-template-columns: minmax(150px, 1.1fr) minmax(150px, 1.1fr) 24px minmax(170px, 1.1fr);
            padding: .55rem .5rem;
        }

        .campo-cabecalho {
            color: #adb5bd;
            font-size: .7rem;
            font-weight: 700;
            letter-spacing: .05em;
            padding-bottom: .35rem;
            padding-top: .35rem;
            text-transform: uppercase;
        }

        .campo-alteracao-linha {
            border-bottom: 1px solid #f1f3f5;
            border-radius: 4px;
        }

        .campo-alteracao-linha:last-child {
            border-bottom-color: transparent;
        }

        .campo-alteracao-linha:hover {
            background-color: #fafbfc;
        }

        .campo-nome {
            color: #212529;
            font-weight: 600;
        }

        .valor-atual,
        .valor-novo {
            white-space: normal;
            word-break: break-word;
        }

        .valor-atual {
            color: #6c757d;
            text-decoration: line-through;
        }

        .valor-novo {
            color: #146c43;
            font-weight: 600;
        }

        .campo-seta {
            color: #ced4da;
            text-align: center;
        }

        .campo-rotulo-mobile {
            color: #adb5bd;
            display: none;
            font-size: .68rem;
            font-weight: 700;
            letter-spacing: .05em;
            text-transform: uppercase;
        }

        .estado-vazio {
            background-color: #f8f9fa;
            border: 1px dashed #dee2e6;
            border-radius: 8px;
            color: #6c757d;
            padding: 2.5rem 1rem;
            text-align: center;
        }

        .badge-grupo {
            background-color: #e7f1ff;
            color: #0a58ca;
            font-weight: 600;
        }
        .tab-content-grupos > .tab-pane {
            display: none;
        }
        .tab-content-grupos > .tab-pane.active {
            display: block;
        }


        @media (max-width: 767.98px) {
            .campo-cabecalho,
            .campo-seta {
                display: none;
            }

            .campo-alteracao-linha {
                grid-template-columns: 1fr;
                row-gap: .35rem;
            }

            .campo-rotulo-mobile {
                display: block;
            }
        }

        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(6px);
            }

            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @media (max-width: 1400px) {
            .container {
                max-width: 95%;
            }

            .form-layout {
                grid-template-columns: 1fr;
            }

            .tabs-container {
                border-bottom: 1px solid #e9ecef;
                border-right: none;
                overflow-x: auto;
            }

            .tabs {
                flex-direction: row;
                min-width: max-content;
                padding: 8px;
            }

            .tab-group-title,
            .tab-meta {
                display: none;
            }

            .tab-button {
                justify-content: center;
                min-width: 135px;
                text-align: center;
                width: auto;
            }

            .tab-title {
                flex: initial;
            }
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script
        src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js"
        integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM"
        crossorigin="anonymous"
    ></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let acao = '';
        let abaAtual = 1;
        let totalAbas = 3;
        let dadosSolicitacao = null;
        let camposAlteracao = [];
        function repoUrl() {
            return 'SolicitacaoClienteAlteracaoRepositorio.asp';
        }

        function obterMetodoAvaliacao() {
            const metodosAvaliacao = {
                AVALIACAO_DIRETOR_COMERCIAL: 'avaliacaoDiretorComercial',
                AVALIACAO_FINANCEIRO: 'avaliacaoFinanceiro',
                AVALIACAO_CONTROLADORIA: 'avaliacaoControladoria'
            };

            return metodosAvaliacao[acao] || '';
        }

        function exibirLoading(exibir) {
            $('#loadingOverlay').toggleClass('d-none', !exibir).toggleClass('d-flex', exibir);
        }

        function navegarParaGrade() {
            window.location.href = 'SolicitacaoClienteAlteracaoGrade.asp';
        }

        function tipoSolicitacaoPossuiAlteracao(tipoSolicitacao) {
            return tipoSolicitacao === 'ALTERACAO_DADOS'
                || tipoSolicitacao === 'DESBLOQUEIO_COM_ALTERACAO_DADOS';
        }

        function configurarAbaCamposAlterados(tipoSolicitacao) {
            const possuiAlteracao = tipoSolicitacaoPossuiAlteracao(tipoSolicitacao);

            $('#abaCamposAlterados, #conteudoCamposAlterados').toggleClass('d-none', !possuiAlteracao);
            totalAbas = possuiAlteracao ? 3 : 2;

            if (!possuiAlteracao && abaAtual === 3) {
                navegarParaAba(2);
            }
        }

        function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) {
                return;
            }

            abaAtual = numeroAba;

            $('.tab-button').removeClass('active');
            $(`.tab-button[data-tab="${numeroAba}"]`).addClass('active');

            $('.tab-content').removeClass('active');
            $(`.tab-content[data-tab-content="${numeroAba}"]`).addClass('active');

            atualizarBotoesNavegacao();
        }

        function atualizarBotoesNavegacao() {
            $('#btnVoltarPagina').prop('disabled', abaAtual <= 1);
            $('#btnProximaPagina').prop('disabled', abaAtual >= totalAbas);
        }

        function marcarAbaComoCompleta(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if (!$button.length) {
                return;
            }

            $button.addClass('completed').removeClass('invalid');
            $button.find('.tab-icon').attr('class', 'tab-icon bi bi-check-circle-fill');
        }

        function marcarAbaComoInvalida(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if (!$button.length) {
                return;
            }

            $button.addClass('invalid').removeClass('completed');
            $button.find('.tab-icon').attr('class', 'tab-icon bi bi-exclamation-circle-fill');
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function carregarDadosCadastro() {
            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: {
                    acao: 'obterPorId',
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    preencherFormulario(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados');
                }
            });
        }

        function preencherFormulario(dados) {
            if (!dados) {
                return;
            }

            dadosSolicitacao = dados;

            $('#codigoPlanta').val(dados.codigoPlanta || '').trigger('change.select2');
            $('#tipoSolicitacao').val(dados.tipoSolicitacao || '').trigger('change.select2');
            configurarAbaCamposAlterados(dados.tipoSolicitacao || '');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            carregarCamposSolicitacao(dados);
        }

        function buscarCamposDisponiveis(codigoPlanta, codigoCliente, codigoLoja) {
            return $.ajax({
                url: repoUrl() + '?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarCamposAlteracaoCliente',
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja
                },
                dataType: 'json'
            });
        }

        function mapearCampoAlteracao(campo, salvo) {
            return {
                idCampo: campo.idCampo,
                codigoCampo: campo.codigoCampo,
                nomeExibicao: campo.nomeExibicao,
                grupoCampo: campo.grupoCampo,
                tipoDado: campo.tipoDado,
                opcoes: campo.opcoes || [],
                ordem: campo.ordem || 999,
                valorAtual: salvo ? (salvo.valorAtual || '') : (campo.valorAtual || ''),
                valorNovo: salvo ? (salvo.valorNovo || '') : ''
            };
        }

        function carregarCamposSolicitacao(dados) {
            if (!tipoSolicitacaoPossuiAlteracao(dados.tipoSolicitacao || '')) {
                camposAlteracao = [];
                return;
            }

            const camposSalvos = dados.campos || [];

            if (!camposSalvos.length) {
                camposAlteracao = [];
                renderizarCamposAlteracao();
                return;
            }

            $('#containerCamposAlteracao').html(`
                <div class="text-center text-muted py-4">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                    <p class="mt-2 mb-0">Carregando campos alterados...</p>
                </div>
            `);

            buscarCamposDisponiveis(dados.codigoPlanta, dados.codigoCliente, dados.codigoLoja)
                .then(function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao carregar campos alterados.');
                        camposAlteracao = camposSalvos.map(function(campo) {
                            return mapearCampoAlteracao({
                                idCampo: campo.idCampo,
                                nomeExibicao: 'Campo #' + campo.idCampo,
                                grupoCampo: 'GERAL'
                            }, campo);
                        });
                        renderizarCamposAlteracao();
                        return;
                    }

                    camposAlteracao = camposSalvos.map(function(salvo) {
                        const campoCatalogo = (response.dados || []).find(function(campo) {
                            return String(campo.idCampo) === String(salvo.idCampo);
                        }) || {
                            idCampo: salvo.idCampo,
                            nomeExibicao: 'Campo #' + salvo.idCampo,
                            grupoCampo: 'GERAL'
                        };

                        return mapearCampoAlteracao(campoCatalogo, salvo);
                    });

                    renderizarCamposAlteracao();
                })
                .fail(function(xhr, status, error) {
                    console.error('Erro ao carregar campos alterados:', error);
                    Toast.error('Erro ao carregar campos alterados.');
                    camposAlteracao = camposSalvos.map(function(campo) {
                        return mapearCampoAlteracao({
                            idCampo: campo.idCampo,
                            nomeExibicao: 'Campo #' + campo.idCampo,
                            grupoCampo: 'GERAL'
                        }, campo);
                    });
                    renderizarCamposAlteracao();
                });
        }

        function renderizarCamposAlteracao() {
            if (!camposAlteracao.length) {
                $('#containerCamposAlteracao').html(`
                    <div class="estado-vazio">
                        <i class="bi bi-inbox fs-3 d-block mb-2"></i>
                        Nenhum campo alterado informado para esta solicitação.
                    </div>
                `);
                return;
            }

            const camposOrdenados = camposAlteracao.slice().sort(function(a, b) {
                return a.ordem - b.ordem;
            });

            let conteudo = `
                <div class="campo-cabecalho">
                    <div>Campo</div>
                    <div>Valor atual</div>
                    <div></div>
                    <div>Novo valor</div>
                </div>
            `;

            camposOrdenados.forEach(function(campo) {
                conteudo += renderizarLinhaCampoAlteracao(campo);
            });

            $('#containerCamposAlteracao').html(conteudo);
        }

        function renderizarLinhaCampoAlteracao(campo) {
            const textoValorAtual = obterTextoExibicaoValor(campo, campo.valorAtual);
            const textoValorNovo = obterTextoExibicaoValor(campo, campo.valorNovo);

            return `
                <div class="campo-alteracao-linha" data-id-campo="${escapeAttribute(campo.idCampo)}">
                    <div class="campo-nome">
                        ${escapeHtml(campo.nomeExibicao || ('Campo #' + campo.idCampo))}
                    </div>
                    <div>
                        <span class="campo-rotulo-mobile">Valor atual</span>
                        <div class="valor-atual">
                            ${normalizarValor(textoValorAtual) ? escapeHtml(textoValorAtual) : '<span class="text-muted fst-italic text-decoration-none">Sem valor</span>'}
                        </div>
                    </div>
                    <div class="campo-seta">
                        <i class="bi bi-arrow-right"></i>
                    </div>
                    <div>
                        <span class="campo-rotulo-mobile">Novo valor</span>
                        <div class="valor-novo">
                            ${normalizarValor(textoValorNovo) ? escapeHtml(textoValorNovo) : '<span class="text-muted fst-italic fw-normal">Não informado</span>'}
                        </div>
                    </div>
                </div>
            `;
        }

        function obterValorOpcao(opcao) {
            return (opcao !== null && typeof opcao === 'object') ? opcao.valor : opcao;
        }

        function obterTextoOpcao(opcao) {
            if (opcao !== null && typeof opcao === 'object') {
                return (opcao.descricao !== undefined && opcao.descricao !== null && opcao.descricao !== '')
                    ? opcao.descricao
                    : opcao.valor;
            }

            return opcao;
        }

        function obterTextoExibicaoValor(campo, valor) {
            if (normalizarValor(valor) === '') return valor;

            if (campo.tipoDado === 'BOOLEANO') {
                if (String(valor) === '1') return 'Sim';
                if (String(valor) === '0') return 'Não';
            }

            if (campo.tipoDado === 'LISTA') {
                const opcao = (campo.opcoes || []).find(function(item) {
                    return String(obterValorOpcao(item)) === String(valor);
                });

                if (opcao) return obterTextoOpcao(opcao);
            }

            return valor;
        }

        function normalizarValor(valor) {
            return String(valor === null || valor === undefined ? '' : valor).trim();
        }

        function escapeHtml(value) {
            return String(value === null || value === undefined ? '' : value)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }

        function escapeAttribute(value) {
            return escapeHtml(value).replace(/`/g, '&#096;');
        }

        function validarAvaliacao(flagAprovar) {
            const observacao = ($('#observacaoAvaliacao').val() || '').trim();

            $('.just-validate-error-label').remove();
            $('.just-validate-error-field').removeClass('just-validate-error-field');

            if (!flagAprovar && !observacao) {
                $('#observacaoAvaliacao')
                    .addClass('just-validate-error-field')
                    .after('<div class="just-validate-error-label">Observação obrigatória para rejeição</div>');

                marcarAbaComoInvalida(1);
                navegarParaAba(1);
                Toast.warning('Informe o motivo da reprovação.');

                return false;
            }

            marcarAbaComoCompleta(1);
            return true;
        }

        function avaliarSolicitacao(flagAprovar) {
            if (!validarAvaliacao(flagAprovar)) {
                return;
            }

            const metodoAvaliacao = obterMetodoAvaliacao();

            if (!metodoAvaliacao) {
                Toast.warning('Ação de avaliação inválida.');
                return;
            }

            const observacao = ($('#observacaoAvaliacao').val() || '').trim();

            exibirLoading(true);

            $.ajax({
                url: repoUrl(),
                type: 'POST',
                dataType: 'json',
                data: {
                    acao: metodoAvaliacao,
                    id: id,
                    flagAprovar: flagAprovar ? 1 : 0,
                    observacao: observacao.replaceAll("'", '"')
                },
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem || (
                        flagAprovar
                            ? 'Solicitação aprovada com sucesso!'
                            : 'Solicitação reprovada com sucesso!'
                    ));

                    setTimeout(function() {
                        navegarParaGrade();
                    }, 500);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao avaliar solicitação:', error);
                    Toast.error('Erro ao avaliar solicitação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function anexarEventos() {
            $('#formularioAvaliacao').on('submit', function(e) {
                e.preventDefault();
                return false;
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));

                if (tabNumber === abaAtual) {
                    return;
                }

                navegarParaAba(tabNumber);
            });

            $('#btnVoltarGrade').on('click', navegarParaGrade);

            $('#btnAprovar').on('click', function() {
                avaliarSolicitacao(true);
            });

            $('#btnRejeitar').on('click', function() {
                avaliarSolicitacao(false);
            });
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id');
            acao = (urlParams.get('acao') || '').toUpperCase();

            anexarEventos();
            navegarParaAba(1);
            inicializarSelect2();

            if (!id) {
                $('#titulo').text('Avaliação de Solicitação');
                Toast.warning('Informe uma solicitação para os ados da solicitação.');
                return;
            }

            $('#titulo').text(`Avaliação de Solicitação #${id}`);
            await carregarDadosCadastro();
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex my-4 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" class="btn btn-sm btn-light me-2 btn-back" id="btnVoltarGrade">
                    <i class="bi bi-arrow-left"></i>
                </button>

                <h3 id="titulo" class="mb-0"></h3>
            </div>
        </div>

        <div id="loadingOverlay" class="loading-overlay d-none">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="form-container">
            <div class="form-layout">
                <div class="tabs-container">
                    <div class="tabs">
                        <div class="tab-group-title">Ação requerida</div>

                        <button type="button" class="tab-button active" data-tab="1">
                            <i class="tab-icon bi bi-check2-square"></i>
                            <span class="tab-title">Avaliação</span>
                        </button>

                        <div class="tab-group-title">Consulta</div>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Dados Gerais</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="3" id="abaCamposAlterados">
                            <i class="tab-icon bi bi-list-check"></i>
                            <span class="tab-title">Campos Alterados</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioAvaliacao" name="formularioAvaliacao" class="form-content" novalidate>
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Avaliação da solicitação</div>
                                <p class="section-subtitle">Registre a aprovação ou rejeição da solicitação enviada.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12">
                                    <label for="observacaoAvaliacao" class="form-label">
                                        Observações
                                        <span class="text-muted fw-normal">(obrigatória para rejeição)</span>
                                    </label>

                                    <textarea
                                        class="form-control form-control-sm"
                                        id="observacaoAvaliacao"
                                        rows="5"
                                        maxlength="1000"
                                    ></textarea>
                                </div>

                                <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                                    <button type="button" class="btn btn-sm btn-danger" id="btnRejeitar">
                                        <i class="bi bi-x-lg me-1"></i>Rejeitar
                                    </button>

                                    <button type="button" class="btn btn-sm btn-success" id="btnAprovar">
                                        <i class="bi bi-check-lg me-1"></i>
                                        Aprovar
                                    </button>
                                </div>
                            </div>

                            <div class="readonly-alert mt-3">
                                <i class="bi bi-info-circle me-1"></i>
                                Consulte as demais abas para revisar os dados gerais e os campos alterados antes de concluir.
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Dados Solicitação</div>
                                <p class="section-subtitle">Informações básicas da solicitação selecionada em modo consulta.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="codigoPlanta" class="form-label">Planta</label>
                                    <select disabled class="form-select form-select-sm select2" name="codigoPlanta" id="codigoPlanta">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="tipoSolicitacao" class="form-label">Tipo de Solicitação</label>
                                    <select disabled class="form-select form-select-sm select2" name="tipoSolicitacao" id="tipoSolicitacao">
                                        <option value="">Selecione um item</option>
                                        <option value="DESBLOQUEIO">Desbloqueio Simples</option>
                                        <option value="DESBLOQUEIO_COM_ALTERACAO_DADOS">Desbloqueio c/ Alteração de Dados</option>
                                        <option value="BLOQUEIO">Bloqueio Simples</option>
                                        <option value="ALTERACAO_DADOS">Alteração de Dados</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoCliente" class="form-label">Cód. Cliente</label>
                                    <input disabled onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="codigoLoja" class="form-label">Cód. Loja</label>
                                    <input disabled type="text" onblur="obterCliente()" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja">
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="razaoSocial" class="form-label">Razão Social</label>
                                    <input disabled type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" maxlength="160" />
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="cpfCnpj" class="form-label" id="labelCpfCnpj">CNPJ/CPF</label>
                                    <input disabled type="text" class="form-control form-control-sm" name="cpfCnpj" id="cpfCnpj" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="3" id="conteudoCamposAlterados">
                            <div class="section-header">
                                <div class="section-title">Campos Alterados</div>
                                <p class="section-subtitle">Valores atuais e novos informados na solicitação para revisão da análise.</p>
                            </div>

                            <div id="containerCamposAlteracao">
                                <div class="estado-vazio">
                                    <i class="bi bi-inbox fs-3 d-block mb-2"></i>
                                    Nenhum campo alterado carregado.
                                </div>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
