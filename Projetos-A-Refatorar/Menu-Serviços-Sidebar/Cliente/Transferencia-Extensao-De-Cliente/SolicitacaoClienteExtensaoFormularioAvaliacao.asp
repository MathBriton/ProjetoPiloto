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

        .anexos-list {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
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

        .anexo-item {
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            margin-bottom: 10px;
            padding: 12px;
            transition: background-color .2s;
        }

        .anexo-item:hover {
            background-color: #f9fafb;
        }

        .anexo-icon {
            color: #3b82f6;
            font-size: 1.5rem;
            margin-right: 12px;
        }

        .anexo-info {
            flex: 1;
        }

        .anexo-nome {
            color: #1f2937;
            font-weight: 600;
            margin-bottom: 4px;
        }

        .anexo-data {
            color: #6b7280;
            font-size: .85rem;
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
        let listaAnexos = [];

        function repoUrl() {
            return 'SolicitacaoClienteExtensaoRepositorio.asp';
        }

        function obterMetodoAvaliacao() {
            const metodosAvaliacao = {
                AVALIACAO_FISCAL: 'avaliacaoFiscal',
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
            window.location.href = 'SolicitacaoClienteExtensaoGrade.asp';
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
            listaAnexos = dados.anexos || [];

            $('#codigoPlantaOrigem').val(dados.codigoPlantaOrigem || '').trigger('change.select2');
            $('#codigoPlantaDestino').val(dados.codigoPlantaDestino || '').trigger('change.select2');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#codigoVendedor').val(dados.codigoVendedor || '');
            $('#nomeVendedor').val(dados.nomeVendedor || '');
            $('#tes').val(dados.tes || '');
            $('#toleranciaQuantidade').val(dados.toleranciaQuantidade || '');
            $('#toleranciaDias').val(dados.toleranciaDias || '');

            Promise.all([
                carregarNaturezas(dados.codigoPlantaOrigem)
            ]).then(function() {
                $('#natureza').val(dados.natureza).trigger('change.select2');
            });

            $('#menuContadorAnexos').text(listaAnexos.length);

            renderizarAnexos();
        }

        function carregarNaturezas(codigoPlanta) {
            const $select = $('#natureza');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarNaturezas', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar naturezas: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as naturezas:', error);
                }
            });
        }

        function renderizarAnexos() {
            const $container = $('#anexosListaContainer');

            $('#menuContadorAnexos').text(listaAnexos.length);

            if (listaAnexos.length === 0) {
                $container.html(`
                    <div class="alert alert-info mb-0">
                        <i class="bi bi-info-circle me-1"></i>
                        Nenhum anexo ativo encontrado para esta solicitação.
                    </div>
                `);
                return;
            }

            const html = listaAnexos.map(function(anexo) {
                const nome = anexo.nome;
                const caminho = anexo.caminho;
                const dataInclusao = anexo.dataInclusao;
                const usuario = anexo.loginUsuarioInclusao || '-';

                return `
                    <div class="anexo-item d-flex align-items-center">
                        <i class="bi bi-file-earmark anexo-icon"></i>

                        <div class="anexo-info">
                            <div class="anexo-nome">${nome}</div>
                            <div class="anexo-data">
                                Adicionado por <strong>${usuario}</strong> em: ${dataInclusao || '-'}
                            </div>
                        </div>

                        <a href="/${caminho}" target="_blank" class="btn btn-sm btn-primary">
                            <i class="bi bi-box-arrow-up-right me-1"></i>
                            Abrir
                        </a>
                    </div>
                `;
            }).join('');

            $container.html(`<div class="anexos-list">${html}</div>`);
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

                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
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
                                Consulte as demais abas para revisar os dados gerais e anexos antes de concluir.
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Dados Solicitação</div>
                                <p class="section-subtitle">Informações básicas da solicitação selecionada em modo consulta.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-md-6">
                                    <label for="codigoPlantaOrigem" class="form-label">Planta Origem <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="codigoPlantaOrigem" id="codigoPlantaOrigem">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-md-6">
                                    <label for="codigoPlantaDestino" class="form-label">Planta Destino <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="codigoPlantaDestino" id="codigoPlantaDestino">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="codigoCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                                    <input disabled onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                                </div>
                                
                                <div class="col-12 col-md-2">
                                    <label for="codigoLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                                    <input disabled type="text" onblur="obterCliente()" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja">
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="razaoSocial" class="form-label">Razão Social <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" maxlength="160" />
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="cpfCnpj" class="form-label" id="labelCpfCnpj">CNPJ/CPF <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="cpfCnpj" id="cpfCnpj" />
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="codigoVendedor" class="form-label">Cód. Vendedor <span class="required-field me-2">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="codigoVendedor" name="codigoVendedor">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="nomeVendedor" class="form-label">Nome do Vendedor <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="nomeVendedor" name="nomeVendedor">
                                </div>

                                <div class="col-md-4">
                                    <label for="tes" class="form-label">TES <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="tes" id="tes" maxlength="20" />
                                </div>

                                <div class="col-md-4">
                                    <label for="natureza" class="form-label">Natureza <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="natureza" id="natureza"></select>
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaQuantidade" class="form-label">Tolerância de Quantidade +/- <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm percentual" name="toleranciaQuantidade" id="toleranciaQuantidade" />
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaDias" class="form-label">Tolerância de Dias +/- <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="toleranciaDias" id="toleranciaDias" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Documentos ativos da solicitação para conferência.</p>
                            </div>

                            <div id="anexosListaContainer"></div>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
