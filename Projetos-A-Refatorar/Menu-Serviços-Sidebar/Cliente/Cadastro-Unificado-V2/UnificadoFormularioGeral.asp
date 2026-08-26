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
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
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

        .page-subtitle {
            color: #6c757d;
            font-size: .85rem;
        }

        .page-actions {
            gap: .5rem;
        }

        .form-container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .form-label {
            font-weight: 600;
        }

        .required-field {
            color: #dc3545;
        }

        .form-switch {
            display: flex;
            align-items: end;
        }

        .form-switch .form-check-input {
            width: 3em;
            height: 1.5em;
            cursor: pointer;
        }
        .form-switch .form-check-label {
            cursor: pointer;
            font-weight: 500;
            color: #495057;
            padding-left: 8px;
        }

        .form-layout {
            display: grid;
            grid-template-columns: 270px minmax(0, 1fr);
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

        #btnVoltarPagina:disabled,
        #btnProximaPagina:disabled {
            cursor: not-allowed;
            pointer-events: auto;
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
            padding: 30px;
            height: 80vh;
            overflow: auto;
        }

        .tab-content.active {
            display: block;
        }

        .section-header {
            border-bottom: 1px solid #e9ecef;
            margin-bottom: 14px;
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

        .z-index-9999 {
            z-index: 9999 !important;
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

        .just-validate-error-field:focus {
            border-color: #ff5353 !important;
            box-shadow: none !important;
            outline: none;
        }
        
        .nav-footer {
            position: sticky;
            bottom: 0;
            background: #fff;
            border-top: 1px solid #dee2e6;
            padding: 10px 20px;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            z-index: 10;
        }

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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

        .btn-selecionar {
            align-items: baseline;
            background-color: #fff;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            color: #6c757d;
            cursor: pointer;
            display: flex;
            font-size: .875rem;
            justify-content: center;
            padding: .25rem .5rem;
        }

        .btn-selecionar:hover {
            background-color: #0d6efd;
            border-color: #0d6efd;
            color: #fff;
        }

        .btn-selecionar.selecionado {
            background-color: #198754;
            border-color: #198754;
            color: #fff;
            cursor: default;
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
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-maskmoney/3.0.2/jquery.maskMoney.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/documentUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let codigo = null;
        let preCadastro = '';
        let abaAtual = 1;
        let totalAbas = 6;
        let abasValidadas = new Set();
        let validacoes = {};
        let acao = 'CRIAR';
        let flagValidaDocumentos = true

        // Anexos
        let listaAnexos = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function normalizarAcao(valor) {
            const acaoNormalizada = (valor || '').toString().trim().toUpperCase();
            return ['CRIAR', 'EDITAR', 'VISUALIZAR'].includes(acaoNormalizada) ? acaoNormalizada : 'CRIAR';
        }

        function formularioSomenteVisualizacao() {
            return acao === 'VISUALIZAR';
        }

        function configurarModoFormulario() {
            if (!formularioSomenteVisualizacao()) return;

            $('#divAcoes').addClass('d-none');
            $('#formularioUnificado').find('input, select, textarea').prop('disabled', true);
            $('#formularioUnificado').find('.select2').trigger('change.select2');

            // Em modo visualização, o componente de documento continua visível para permitir abrir o anexo.
            // Apenas bloqueamos ações de anexar/trocar/remover.
            $('document-upload').each(function() {
                if (typeof this.disable === 'function') {
                    this.disable();
                } else {
                    $(this).attr('disabled', 'disabled');
                }
            });

            // O file-upload apaga a lista ao chamar disable(). Por isso bloqueamos
            // a interação sem limpar os anexos opcionais existentes.
            $('file-upload').each(function() {
                bloquearFileUploadSemLimpar(this);
            });
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function inicializarMascaras() {
            $('#telefone').mask('(00) 00000-0000');
            $('#cep, #cepEntrega, #cepCobranca').mask('00000-000');
            $('#cpf').mask('000.000.000-00');
            $('.money').maskMoney({ thousands: '.', decimal: ',', allowZero: true, prefix: 'R$ ' });

            $('#cnpj').mask('AA.AAA.AAA/AAAA-00', {
                translation: { 'A': { pattern: /[A-Za-z0-9]/ } },
                onKeyPress: function(value, event, field) {
                    field.val(value.toUpperCase());
                }
            });
            
            $('#cnpj').on('input', function() {
                this.value = this.value.toUpperCase();
            });
        }

        function carregarTiposIndustria() {
            const $select = $('#tipoIndustria');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposIndustria' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao tipos de industria: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar os tipos de industria:', error);
                }
            });
        }

        function carregarGruposConta() {
            const $select = $('#grupoContas');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarGruposConta' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar grupos de contas: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar grupos de contas:', error);
                }
            });
        }

        function carregarTiposContribuinte() {
            const $select = $('#contribuinteIcms');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposContribuinte' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar tipos de contribuinte: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de contribuinte:', error);
                }
            });
        }

        function carregarTiposDeclarante() {
            const $select = $('#declarantePisCofins');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposDeclarante' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar tipos de declarante: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de declarante:', error);
                }
            });
        }

        function carregarMetodosPagamento() {
            const $select = $('#metodoPagamento');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarMetodosPagamento' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar métodos de pagamento: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar métodos de pagamento:', error);
                }
            });
        }

        function carregarTiposFrete() {
            const $select = $('#tipoFrete');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposFrete' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar tipos de frete: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de frete:', error);
                }
            });
        }

        function carregarEstados() {
            const $select = $('#estado, #estadoEntrega, #estadoCobranca');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarEstados' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar estados: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar os estados:', error);
                }
            });
        }

        function carregarCidades(cidadeSelecionada) {
            const estado = $('#estado').val();
            const $select = $('#cidade');
            $select.empty().append('<option value="">Selecione um item</option>');

            if (!estado) {
                $select.trigger('change.select2');
                return $.Deferred().resolve().promise();
            }

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarCidades', codigoEstado: estado },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar cidades: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.nome }));
                    });

                    if (cidadeSelecionada) {
                        if (!$select.find('option').filter(function() { return $(this).val() == cidadeSelecionada; }).length) {
                            $select.append($('<option>', { value: cidadeSelecionada, text: cidadeSelecionada }));
                        }
                        $select.val(cidadeSelecionada);
                    }

                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as cidades:', error);
                }
            });
        }

        function carregarCidadesEntrega(cidadeSelecionada) {
            const estado = $('#estadoEntrega').val();
            const $select = $('#cidadeEntrega');
            $select.empty().append('<option value="">Selecione um item</option>');

            if (!estado) {
                $select.trigger('change.select2');
                return $.Deferred().resolve().promise();
            }

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarCidades', estado: estado },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar cidades de entrega: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.nome }));
                    });

                    if (cidadeSelecionada) {
                        if (!$select.find('option').filter(function() { return $(this).val() == cidadeSelecionada; }).length) {
                            $select.append($('<option>', { value: cidadeSelecionada, text: cidadeSelecionada }));
                        }
                        $select.val(cidadeSelecionada);
                    }

                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as cidades de entrega:', error);
                }
            });
        }

        function carregarCidadesCobranca(cidadeSelecionada) {
            const estado = $('#estadoCobranca').val();
            const $select = $('#cidadeCobranca');
            $select.empty().append('<option value="">Selecione um item</option>');

            if (!estado) {
                $select.trigger('change.select2');
                return $.Deferred().resolve().promise();
            }

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarCidades', estado: estado },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar cidades de cobrança: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.nome }));
                    });

                    if (cidadeSelecionada) {
                        if (!$select.find('option').filter(function() { return $(this).val() == cidadeSelecionada; }).length) {
                            $select.append($('<option>', { value: cidadeSelecionada, text: cidadeSelecionada }));
                        }
                        $select.val(cidadeSelecionada);
                    }

                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as cidades de cobrança:', error);
                }
            });
        }

        function carregarCondicoesPagamento() {
            const $select = $('#condicaoPagamento');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarCondicoesPagamento' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar condições de pagamento: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as condições de pagamento:', error);
                }
            });
        }

        function carregarSegmentos() {
            const $select = $('#segmento');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarSegmentos' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar segmentos: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar os segmentos:', error);
                }
            });
        }

        function carregarContasConciliacao() {
            const $select = $('#contaConciliacao');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarContasConciliacao' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar contas: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar os contas:', error);
                }
            });
        }

        function carregarNaturezas(codigoPlanta) {
            const $select = $('#natureza');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
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

        function carregarTiposCliente() {
            const $select = $('#tipoCliente');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposCliente' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tipos de cliente: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de cliente:', error);
                }
            });
        }

        function carregarVendedores(codigoPlanta) {
            const $select = $('#codigoVendedor');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarVendedores', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar vendedores: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar os vendedores:', error);
                }
            });
        }

        function preencherEnderecoPorCep(cep, prefixo) {
            const cepLimpo = String(cep || '').replace(/\D/g, '');
            if (cepLimpo.length !== 8) return;

            $.ajax({
                url: 'https://viacep.com.br/ws/' + cepLimpo + '/json/',
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    if (response.erro) {
                        Toast.error('CEP não encontrado.');
                        return;
                    }

                    const sufixo = prefixo || '';
                    const estadoId = '#estado' + sufixo;
                    const cidadeId = '#cidade' + sufixo;
                    const cidade = response.localidade || '';
                    const uf = response.uf || '';

                    $('#rua' + sufixo).val(response.logradouro || '');
                    $('#bairro' + sufixo).val(response.bairro || '');

                    if (uf) {
                        if (!$(estadoId).find('option').filter(function() { return $(this).val() == uf; }).length) {
                            $(estadoId).append($('<option>', { value: uf, text: uf }));
                        }
                        $(estadoId).val(uf).trigger('change.select2');
                    }

                    if (prefixo === 'Entrega') {
                        carregarCidadesEntrega(cidade);
                    } else if (prefixo === 'Cobranca') {
                        carregarCidadesCobranca(cidade);
                    } else {
                        carregarCidades(cidade);
                    }

                    if (cidade && !$(cidadeId).find('option').filter(function() { return $(this).val() == cidade; }).length) {
                        $(cidadeId).append($('<option>', { value: cidade, text: cidade })).val(cidade).trigger('change.select2');
                    }
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao consultar CEP:', error);
                    Toast.error('Erro ao consultar CEP.');
                }
            });
        }

        function aplicarCamposCondicionais() {
            const tipoSolicitacao = $('#tipoSolicitacao').val();
            const tipoPessoa = $('#tipoPessoa').val();
            const exibirUnificacaoCliente = !!codigo && preCadastro === 'P';
            const pessoaFisica = tipoPessoa === 'F';
            const pessoaJuridica = tipoPessoa === 'J';
            const clienteIntercompany = ['94100000', '41200100', '41200000'].includes($('#contaConciliacao').val())
            const icmsDiferido = $('#icmsDiferido').val() == '1';
            const isencaoIpi = $('#isencaoIpi').val() == '1';
            const entregaDiferente = $('#enderecoEntregaDiferente').is(':checked');
            const cobrancaDiferente = $('#enderecoCobrancaDiferente').is(':checked');
            const flagUnificarCliente = $('#flagUnificarCliente').is(':checked');

            $('#linhaUnificarCliente').toggleClass('d-none', !exibirUnificacaoCliente);
            if (!exibirUnificacaoCliente) {
                $('#flagUnificarCliente').prop('checked', false);
            }

            $('#avisoUnificar').toggle(flagUnificarCliente);
            $('.campos-pessoa-fisica').toggle(pessoaFisica);
            $('.campos-pessoa-juridica').toggle(pessoaJuridica);
            $('.campo-intercompany').toggle(clienteIntercompany);
            $('.campo-diferimento').toggle(icmsDiferido);
            $('.campo-isencao-ipi').toggle(isencaoIpi);
            $('.grupo-endereco-entrega').toggle(entregaDiferente);
            $('.grupo-endereco-cobranca').toggle(cobrancaDiferente);
            
            atualizarMarcacoesObrigatorias();
            atualizarObrigatoriedadeDocumentos();
        }

        function documentosObrigatoriosPorTipoSolicitacao() {
            return $('#tipoSolicitacao').val() === '1' || $('#flagUnificarCliente').is(':checked');
        }

        function valorSelectNormalizado(seletor) {
            const $select = $(seletor);
            const valor = ($select.val() || '').toString().trim().toUpperCase();
            const texto = ($select.find('option:selected').text() || '').toString().trim().toUpperCase();
            return { valor, texto };
        }

        function inscricaoEstadualInformada() {
            return String($('#inscricaoEstadual').val() || '').replace(/\D/g, '').length > 0;
        }

        function contribuinteIcmsSelecionado() {
            const selecionado = valorSelectNormalizado('#contribuinteIcms');
            if (!selecionado.valor && !selecionado.texto) return false;
            if (selecionado.texto.includes('NÃO') || selecionado.texto.includes('NAO')) return false;
            if (selecionado.texto.includes('CONTRIBUINTE')) return true;
            return ['1', 'S', 'SIM', 'CONTRIBUINTE'].includes(selecionado.valor);
        }
        
        function definirDocumentoObrigatorio(seletor, obrigatorio) {
            const $documento = $(seletor);
            if (!$documento.length) return;
            const requerido = !!obrigatorio;
            $documento.toggle(requerido);
            $documento.each(function() {
                if (typeof this.setRequired === 'function') {
                    this.setRequired(requerido);
                } else {
                    this.required = requerido;
                    if (requerido) {
                        this.setAttribute('required', '');
                    } else {
                        this.removeAttribute('required');
                    }
                }
            });
        }
        function atualizarObrigatoriedadeDocumentos() {
            const validarDocumentos = documentosObrigatoriosPorTipoSolicitacao();
            const pessoaFisica = $('#tipoPessoa').val() === 'F';
            const pessoaJuridica = $('#tipoPessoa').val() === 'J';
            const contribuinteIcms = contribuinteIcmsSelecionado();
            const isencaoIpi = $('#isencaoIpi').val() == '1';
            const icmsDiferido = $('#icmsDiferido').val() == '1';

            $('document-upload').each(function() {
                definirDocumentoObrigatorio(this, false);
            });

            definirDocumentoObrigatorio('#doc_cartao_cnpj', validarDocumentos && pessoaJuridica);
            definirDocumentoObrigatorio('#doc_consulta_sintegra', validarDocumentos && pessoaJuridica);
            definirDocumentoObrigatorio('#doc_contrato_social', validarDocumentos && pessoaJuridica);

            definirDocumentoObrigatorio('#doc_documento_foto', validarDocumentos && pessoaFisica);
            definirDocumentoObrigatorio('#doc_declaracao_ir', validarDocumentos && pessoaFisica);

            definirDocumentoObrigatorio('#doc_comprovante_endereco', validarDocumentos && (pessoaFisica || pessoaJuridica));
            definirDocumentoObrigatorio('#doc_inscricao_estadual', validarDocumentos && contribuinteIcms);
            definirDocumentoObrigatorio('#doc_carta_isencao_ipi', validarDocumentos && isencaoIpi && (pessoaFisica || pessoaJuridica));
            definirDocumentoObrigatorio('#doc_comprovatorio_icms_diferido', validarDocumentos && icmsDiferido);
        }
        
        function definirMarcacaoObrigatoria(campoId, obrigatorio) {
            const $label = $(`label[for="${campoId}"]`);
            if (!$label.length) return;

            let $span = $label.find('.required-field');
            if (obrigatorio) {
                if (!$span.length) {
                    $label.append(' <span class="required-field">*</span>');
                } else {
                    $span.show();
                }
            } else if ($span.length) {
                $span.hide();
            }
        }

        function atualizarMarcacoesObrigatorias() {
            const tipoSolicitacaoCadastro = $('#tipoSolicitacao').val() === '1';
            const unificarCliente = $('#flagUnificarCliente').is(':checked');
            const obrigatorioCadastroCompleto = tipoSolicitacaoCadastro || unificarCliente;
            const pessoaFisica = $('#tipoPessoa').val() == 'F';
            const pessoaJuridica = $('#tipoPessoa').val() == 'J';
            const icmsDiferido = $('#icmsDiferido').val() == '1';
            const entregaDiferente = $('#enderecoEntregaDiferente').is(':checked');
            const cobrancaDiferente = $('#enderecoCobrancaDiferente').is(':checked');
            const clienteIntercompany = ['94100000', '41200100', '41200000'].includes($('#contaConciliacao').val())

            // Obrigatórios tanto para cadastro completo como também pra prospect
            const camposObrigatorios = [
                'codigoPlanta', 'tipoSolicitacao', 'eshop', 'tipoPessoa', 'razaoSocial',
                'cep', 'rua', 'numero', 'bairro', 'estado', 'cidade', 'idioma'
            ];

            // Obrigatórios apenas para cadastro completo (Tipo de solicitação Cadastro de Cliente)
            const camposObrigatoriosCadastroCompleto = [
                'grupoContas', 'nomeFantasia', 
                'icmsDiferido', 'isencaoIpi', 'tipoIndustria', 'natureza',
                'contaConciliacao', 'chaveOrdenacao', 'condicaoPagamento', 'metodoPagamento',
                'codigoJuros', 'limiteCredito', 'volumeVendaMensalM2', 'precoVendaNetM2',
                'precoVendaNetKg', 'segmento', 'tipoFrete', 'codigoVendedor', 'diasViagem',
                'toleranciaDiasMaior', 'toleranciaDiasMenor', 'pais', 'kmDestino',
                'emailXml', 'emailOrdemPagamento', 'toleranciaQuantidadeMaior', 'toleranciaQuantidadeMenor'
            ];

            // Obrigatório apenas quando marcado 'Endereço de Entrega diferente do cadastro'
            const camposEnderecoEntrega = [
                'cepEntrega', 'ruaEntrega', 'numeroEntrega', 'bairroEntrega',
                'estadoEntrega', 'cidadeEntrega', 'kmEntrega'
            ];

            // Obrigatório apenas quando marcado 'Endereço de Cobrança diferente do cadastro'
            const camposEnderecoCobranca = [
                'cepCobranca', 'ruaCobranca', 'numeroCobranca', 'bairroCobranca',
                'estadoCobranca', 'cidadeCobranca'
            ];

            $('.required-field').hide();

            camposObrigatorios.forEach(campo => definirMarcacaoObrigatoria(campo, true));
            camposObrigatoriosCadastroCompleto.forEach(campo => definirMarcacaoObrigatoria(campo, obrigatorioCadastroCompleto));

            // Obrigatório para pessoa juridica
            definirMarcacaoObrigatoria('cnpj', pessoaJuridica);

            // Obrigatório para pessoa física
            definirMarcacaoObrigatoria('cpf', pessoaFisica);

            // Obrigatório quando ICMS Diferido = Sim
            definirMarcacaoObrigatoria('percentualDiferimento', icmsDiferido);

            // Codigo codigoHfm para clientes intercompany
            definirMarcacaoObrigatoria('codigoHfm', clienteIntercompany);

            camposEnderecoEntrega.forEach(campo => definirMarcacaoObrigatoria(campo, entregaDiferente));
            camposEnderecoCobranca.forEach(campo => definirMarcacaoObrigatoria(campo, cobrancaDiferente));
        }

        function renderizarAnexos() {
            carregarAnexosExistentesNosUploads();
            carregarAnexosOpcionaisExistentes();
            atualizarContadorAnexos();
        }

        function atualizarContadorAnexos() {
            let total = 0;

            $('document-upload').each(function() {
                if (!documentoUploadEstaDisponivel(this)) return;

                if (typeof this.hasFile === 'function' && this.hasFile()) {
                    total++;
                }
            });

            const uploadAnexos = document.getElementById('uploadAnexos');
            if (uploadAnexos && typeof uploadAnexos.getFiles === 'function') {
                total += uploadAnexos.getFiles().length;
            }

            $('#menuContadorAnexos').text(total);
        }

        function normalizarTextoComparacao(valor) {
            return String(valor || '').trim().toLowerCase();
        }

        function obterNomeDocumentoAnexo(anexo) {
            return anexo.nome || '';
        }

        function obterIdDocumentoAnexo(anexo) {
            return anexo.id || '';
        }

        function localizarAnexoPorNome(nomeDocumento) {
            const nomeNormalizado = normalizarTextoComparacao(nomeDocumento);

            return (listaAnexos || []).find(anexo =>
                normalizarTextoComparacao(obterNomeDocumentoAnexo(anexo)) === nomeNormalizado
            );
        }

        function montarItemAnexoParaUpload(anexo) {
            if (!anexo) return null;

            const caminho = anexo.caminho || '';

            return {
                id: obterIdDocumentoAnexo(anexo),
                nome: obterNomeDocumentoAnexo(anexo),
                caminho: caminho,
                dataInclusao: anexo.dataInclusao || '',
                invalidated: parseInt(anexo.flagInvalido) || false,
                invalidationReason: anexo.motivoInvalido || '',
                url: caminho ? `/Clientes_Docs/${caminho}` : ''
            };
        }

        function obterNomesDocumentosObrigatorios() {
            return $('document-upload').map(function() {
                return normalizarTextoComparacao(this.getAttribute('name'));
            }).get().filter(Boolean);
        }

        function anexoPertenceADocumentoObrigatorio(anexo) {
            const nomeAnexo = normalizarTextoComparacao(obterNomeDocumentoAnexo(anexo));
            return obterNomesDocumentosObrigatorios().includes(nomeAnexo);
        }

        function obterAnexosOpcionais() {
            return (listaAnexos || []).filter(anexo => !anexoPertenceADocumentoObrigatorio(anexo));
        }

        function montarItemAnexoOpcionalParaFileUpload(anexo) {
            if (!anexo) return null;

            const caminho = anexo.caminho || anexo.anexo || anexo.attachment || '';
            const nome = obterNomeDocumentoAnexo(anexo) || caminho;

            return {
                id: obterIdDocumentoAnexo(anexo),
                descricao: nome,
                description: nome,
                anexo: caminho,
                attachment: caminho
            };
        }

        function carregarAnexosOpcionaisExistentes() {
            const componente = document.getElementById('uploadAnexos');
            if (!componente || typeof componente.loadExistingFiles !== 'function') return;

            const anexosOpcionais = obterAnexosOpcionais()
                .map(montarItemAnexoOpcionalParaFileUpload)
                .filter(Boolean);

            componente.loadExistingFiles(anexosOpcionais, '../Clientes_Docs');
        }

        function bloquearFileUploadSemLimpar(componente) {
            if (!componente) return;

            componente.disabled = true;
            componente.setAttribute('disabled', '');

            if (typeof componente.updateDisabledState === 'function') {
                componente.updateDisabledState();
            }
        }

        function normalizarIdOriginalAnexoOpcional(file) {
            const id = String(file?.originalId || file?.id || '');
            return id.startsWith('existing-') ? id.replace('existing-', '') : id;
        }

        function montarPayloadAnexoOpcional(file, action) {
            const nome = file.name || file.uploadedName || '';
            const arquivo = file.uploadedName || (file.url ? file.url.split('/').pop() : '');
            const originalId = normalizarIdOriginalAnexoOpcional(file);

            return {
                action: action,
                originalId: originalId,
                name: nome,
                uploadedName: arquivo
            };
        }

        function coletarAnexosOpcionais() {
            const componente = document.getElementById('uploadAnexos');
            if (!componente || typeof componente.getFiles !== 'function') return [];

            const anexosAtivos = componente.getFiles().map(file =>
                montarPayloadAnexoOpcional(file, file.existing ? 'unchanged' : 'added')
            );

            const anexosRemovidos = typeof componente.getRemovedFiles === 'function'
                ? componente.getRemovedFiles().map(file => montarPayloadAnexoOpcional(file, 'removed'))
                : [];

            return [...anexosAtivos, ...anexosRemovidos];
        }

        function carregarAnexosExistentesNosUploads() {
            $('document-upload').each(function() {
                const componente = this;
                const nomeDocumento = componente.getAttribute('name');
                const anexo = localizarAnexoPorNome(nomeDocumento);

                if (typeof componente.isDirty === 'function' && componente.isDirty()) {
                    return;
                }

                if (anexo && typeof componente.loadExistingFile === 'function') {
                    componente.loadExistingFile(montarItemAnexoParaUpload(anexo));
                }
            });
        }

        function anexoJaExiste(nome) {
            return listaAnexos.some(anexo =>
                String(anexo.nome || '').trim().toLowerCase() === String(nome || '').trim().toLowerCase()
            );
        }


        function documentoUploadEstaDisponivel(componente) {
            const $componente = $(componente);

            if ($componente.hasClass('documento-ja-anexado') || $componente.hasClass('d-none') || $componente.attr('disabled')) {
                return false;
            }

            return $componente.parents().addBack().filter(function() {
                const $elemento = $(this);

                if ($elemento.hasClass('tab-content')) {
                    return false;
                }

                return $elemento.hasClass('d-none') || this.style.display === 'none';
            }).length === 0;
        }

        function coletarDocumentos() {
            const componentes = $('document-upload').filter(function() {
                return documentoUploadEstaDisponivel(this);
            }).toArray();
            
            const documentos = [];
            let tudoOk = true;

            componentes.forEach(c => {
                const nomeDocumento = c.getAttribute('name');

                if (typeof c.validate === 'function' && !c.validate()) {
                    tudoOk = false;
                    return;
                }

                if (typeof c.getValue !== 'function') {
                    return;
                }

                const valor = c.getValue();

                if (valor.action === 'added' || valor.action === 'replaced') {
                    valor.name = nomeDocumento;
                }

                documentos.push(valor);
            });

            if (!tudoOk) {
                return null;
            }

            documentos.push(...coletarAnexosOpcionais());
            return documentos;
        }

        function inicializarValidacoes() {
            const obrigatorio = (msg) => ({
                validator: (value, fields) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const obrigatorioQuando = (condicao, msg) => ({
                validator: (value, fields) => !condicao() || (value !== null && value !== undefined && String(value).trim() !== ''),
                errorMessage: msg
            });

            const numeroEntreQuandoInformado = (min, max, msg) => ({
                validator: (value, fields) => {
                    if (value === null || value === undefined || String(value).trim() === '') return true;
                    const numero = parseFloat(String(value).replace(',', '.'));
                    return !isNaN(numero) && numero >= min && numero <= max;
                },
                errorMessage: msg
            });

            const menorQueQuandoInformado = (limite, msg) => ({
                validator: (value, fields) => {
                    const numero = dinheiroParaNumero(value);
                    if (numero === null) return true;
                    return !isNaN(numero) && numero < limite;
                },
                errorMessage: msg
            });

            const somenteNumerosMaximoDigitosQuandoInformado = (maximoDigitos, msg) => ({
                validator: (value, fields) => {
                    if (value === null || value === undefined || String(value).trim() === '') return true;

                    const valor = String(value).trim();

                    return /^\d+$/.test(valor) && valor.length <= maximoDigitos;
                },
                errorMessage: msg
            });

            const tipoSolicitacaoCadastro = () => $('#tipoSolicitacao').val() == '1';
            const unificarCliente = () => $('#flagUnificarCliente').is(':checked');
            const cadastroCompletoOuUnificacao = () => tipoSolicitacaoCadastro() || unificarCliente();
            const pessoaJuridica = () => $('#tipoPessoa').val() == 'J';
            const pessoaFisica = () => $('#tipoPessoa').val() == 'F';
            const icmsDiferidoSim = () => $('#icmsDiferido').val() == '1';
            const enderecoEntregaDiferente = () => $('#enderecoEntregaDiferente').is(':checked');
            const enderecoCobrancaDiferente = () => $('#enderecoCobrancaDiferente').is(':checked');
            const clienteInterCompany = () => ['94100000', '41200100', '41200000'].includes($('#contaConciliacao').val())

            validacoes.aba1 = new JustValidate('#formularioUnificado', { lockForm: false });
            validacoes.aba1
                .addField('#codigoPlanta', [obrigatorio('Planta é obrigatória')])
                .addField('#tipoSolicitacao', [obrigatorio('Tipo de solicitação é obrigatório')])
                .addField('#eshop', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Eshop é obrigatório')])
                .addField('#grupoContas', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Grupo de contas é obrigatório')])
                .addField('#codigoHfm', [obrigatorioQuando(clienteInterCompany, 'Código HFM é obrigatório para clientes intercompany')])
                .addField('#tipoIndustria', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tipo de indústria é obrigatório')])
                .addField('#segmento', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Segmento é obrigatório')])
                .addField('#natureza', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Natureza é obrigatória')])
                .addField('#tipoPessoa', [obrigatorio('Tipo de pessoa é obrigatório')])
                .addField('#cpf', [obrigatorioQuando(pessoaFisica, 'CPF é obrigatório para pessoa física')])
                .addField('#cnpj', [obrigatorioQuando(pessoaJuridica, 'CNPJ é obrigatório para pessoa jurídica')])
                .addField('#inscricaoEstadual', [
                    somenteNumerosMaximoDigitosQuandoInformado(15, 'Inscrição estadual deve conter apenas números e no máximo 15 dígitos')
                ])
                .addField('#inscricaoMunicipal', [
                    somenteNumerosMaximoDigitosQuandoInformado(15, 'Inscrição municipal deve conter apenas números e no máximo 15 dígitos')
                ])
                .addField('#razaoSocial', [obrigatorio('Razão social é obrigatória')])
                .addField('#nomeFantasia', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Nome fantasia é obrigatório')])
                .addField('#codigoSuframa', [somenteNumerosMaximoDigitosQuandoInformado(9, 'Código suframa deve conter apenas números e no máximo 9 dígitos')])
                .addField('#idioma', [obrigatorio('Idioma é obrigatório')]);

            validacoes.aba2 = new JustValidate('#formularioUnificado', { lockForm: false });
            validacoes.aba2
                .addField('#icmsDiferido', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'ICMS diferido é obrigatório')])
                .addField('#percentualDiferimento', [obrigatorioQuando(icmsDiferidoSim, 'Percentual de diferimento é obrigatório')])
                .addField('#isencaoIpi', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Suspensão/Isenção do IPI é obrigatória')])
                .addField('#tes', [somenteNumerosMaximoDigitosQuandoInformado(3, 'TES deve conter apenas números e no máximo 3 dígitos')]);

            validacoes.aba3 = new JustValidate('#formularioUnificado', { lockForm: false });
            validacoes.aba3
                .addField('#contaConciliacao', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Conta conciliação é obrigatória')])
                .addField('#chaveOrdenacao', [
                    obrigatorioQuando(cadastroCompletoOuUnificacao, 'Chave de ordenação é obrigatória'),
                    somenteNumerosMaximoDigitosQuandoInformado(4, 'Chave de ordenação deve conter apenas números e no máximo 4 dígitos')
                ])
                .addField('#codigoJuros', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Código de juros é obrigatório'),])
                .addField('#condicaoPagamento', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Condição de pagamento é obrigatória')])
                .addField('#metodoPagamento', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Método de pagamento é obrigatório')])
                .addField('#limiteCredito', [
                    obrigatorioQuando(cadastroCompletoOuUnificacao, 'Limite de crédito é obrigatório'),
                    menorQueQuandoInformado(180000, 'Limite de crédito deve ser menor que 180000')
                ])
                .addField('#volumeVendaMensalM2', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Volume de venda projetado mensal é obrigatório')])
                .addField('#precoVendaNetM2', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Preço de venda net em R$/M² é obrigatório')])
                .addField('#precoVendaNetKg', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Preço de venda net em R$/KG é obrigatório')]);

            validacoes.aba4 = new JustValidate('#formularioUnificado', { lockForm: false });
            validacoes.aba4
                .addField('#tipoFrete', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tipo de frete é obrigatório')])
                .addField('#codigoVendedor', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Código do vendedor é obrigatório')])
                .addField('#diasViagem', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Transit time é obrigatório')])
                .addField('#toleranciaDiasMaior', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tolerância de dias para entrega a maior é obrigatória')])
                .addField('#toleranciaDiasMenor', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tolerância de dias para entrega a menor é obrigatória')])
                .addField('#toleranciaQuantidadeMaior', [
                    obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tolerância de quantidade a maior é obrigatória'),
                    numeroEntreQuandoInformado(0, 10, 'Tolerância de quantidade a maior deve estar entre 0 e 10')
                ])
                .addField('#toleranciaQuantidadeMenor', [
                    obrigatorioQuando(cadastroCompletoOuUnificacao, 'Tolerância de quantidade a menor é obrigatória'),
                    numeroEntreQuandoInformado(0, 10, 'Tolerância de quantidade a menor deve estar entre 0 e 10')
                ]);

            validacoes.aba5 = new JustValidate('#formularioUnificado', { lockForm: false });
            validacoes.aba5
                .addField('#cep', [obrigatorio('CEP é obrigatório')])
                .addField('#rua', [obrigatorio('Rua é obrigatória')])
                .addField('#numero', [obrigatorio('Número é obrigatório')])
                .addField('#bairro', [obrigatorio('Bairro é obrigatório')])
                .addField('#estado', [obrigatorio('Estado é obrigatório')])
                .addField('#cidade', [obrigatorio('Cidade é obrigatória')])
                .addField('#pais', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'País é obrigatório')])
                .addField('#kmDestino', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'Km até destino é obrigatório')])
                .addField('#emailXml', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'E-mail para envio do XML é obrigatório')])
                .addField('#emailOrdemPagamento', [obrigatorioQuando(cadastroCompletoOuUnificacao, 'E-mail para envio da ordem de pagamento é obrigatório')])

                // Se endereço de entrega diferente
                .addField('#cepEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'CEP de entrega é obrigatório')])
                .addField('#ruaEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'Rua de entrega é obrigatória')])
                .addField('#numeroEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'Número de entrega é obrigatório')])
                .addField('#bairroEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'Bairro de entrega é obrigatório')])
                .addField('#estadoEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'Estado de entrega é obrigatório')])
                .addField('#cidadeEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'Cidade de entrega é obrigatória')])
                .addField('#kmEntrega', [obrigatorioQuando(enderecoEntregaDiferente, 'KM de entrega é obrigatório')])

                // Se endereço de cobrança diferente
                .addField('#cepCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'CEP de cobrança é obrigatório')])
                .addField('#ruaCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'Rua de cobrança é obrigatória')])
                .addField('#numeroCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'Número de cobrança é obrigatório')])
                .addField('#bairroCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'Bairro de cobrança é obrigatório')])
                .addField('#estadoCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'Estado de cobrança é obrigatório')])
                .addField('#cidadeCobranca', [obrigatorioQuando(enderecoCobrancaDiferente, 'Cidade de cobrança é obrigatória')]);
        }

        function obterValorCampo(campo) {
            if (!campo || !campo.elem) return '';

            const elem = campo.elem;

            if (elem.type === 'checkbox') {
                return elem.checked ? elem.value : '';
            }

            if (elem.type === 'radio') {
                const selecionado = document.querySelector(`input[name="${elem.name}"]:checked`);
                return selecionado ? selecionado.value : '';
            }

            return $(elem).val();
        }

        function exibirErrosValidacao(validator) {
            if (!validator || !validator.fields) {
                return;
            }

            const errorsParaExibir = {};

            Object.entries(validator.fields).forEach(([index, campo]) => {
                if (!campo || !campo.elem || !campo.elem.id || !campo.rules) {
                    return;
                }

                const valorAtual = obterValorCampo(campo);

                const primeiraRegraInvalida = campo.rules.find(rule => {
                    if (typeof rule.validator !== 'function') {
                        return rule.isValid === false;
                    }

                    try {
                        return !rule.validator(valorAtual, validator.fields);
                    } catch (e) {
                        console.error(`Erro ao validar campo #${campo.elem.id}`, e);
                        return true;
                    }
                });

                if (primeiraRegraInvalida) {
                    errorsParaExibir[`#${campo.elem.id}`] =
                        primeiraRegraInvalida.errorMessage || 'Campo inválido';
                }
            });

            validator.showErrors(errorsParaExibir);
        }

        function navegarParaGrade() {
            window.location.href = 'UnificadoGrade.asp';
        }

        function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) return;

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

            if ($button.length) {
                $button.addClass('completed').removeClass('invalid');
                const $icon = $button.find('.tab-icon');
                if ($icon.length) {
                    $icon.attr('class', 'tab-icon bi bi-check-circle-fill');
                }
            }

            abasValidadas.add(numeroAba);
        }

        function marcarAbaComoInvalida(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);
            if ($button.length && !$button.hasClass('completed')) {
                $button.addClass('invalid');
                const $icon = $button.find('.tab-icon');
                if ($icon.length) {
                    $icon.attr('class', 'tab-icon bi bi-exclamation-circle-fill');
                }
            }
        }

        function carregarDadosCadastro() {
            const params = {
                acao: 'obterPorCodigo',
                codigo: codigo
            };

            return $.ajax({
                url: 'UnificadoRepositorio.asp',
                type: 'GET',
                data: params,
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
            if (!dados) return;

            preCadastro = dados.preCadastro || '';

            $('#codigoPlanta').val(dados.codigoPlanta).prop('disabled', true).trigger('change.select2');
            $('#tipoSolicitacao').val(dados.tipoSolicitacao).prop('disabled', true).trigger('change.select2');
            $('#eshop').val(dados.eshop).trigger('change.select2');
            $('#grupoContas').val(dados.grupoContas).trigger('change.select2');
            $('#codigoHfm').val(dados.codigoHfm);
            $('#tipoPessoa').val(dados.tipoPessoa).trigger('change.select2');
            $('#tipoIndustria').val(dados.tipoIndustria).trigger('change.select2');
            $('#segmento').val(dados.segmento).trigger('change.select2');
            $('#cpf').val(dados.cpf).trigger('input');
            $('#cnpj').val(dados.cnpj).trigger('input');
            $('#inscricaoEstadual').val(dados.inscricaoEstadual);
            $('#inscricaoMunicipal').val(dados.inscricaoMunicipal);
            $('#razaoSocial').val(dados.razaoSocial);
            $('#nomeFantasia').val(dados.nomeFantasia);
            $('#codigoSuframa').val(dados.codigoSuframa);
            $('#idioma').val(dados.idioma).trigger('change.select2');
            $('#codigoSap').val(dados.codigoSap);
            $('#codigoProtheus').val(dados.codigoProtheus);
            $('#codigoLoja').val(dados.codigoLoja);

            $('#declarantePisCofins').val(dados.declarantePisCofins).trigger('change.select2');
            $('#contribuinteIcms').val(dados.contribuinteIcms).trigger('change.select2');
            $('#icmsDiferido').val(dados.icmsDiferido).trigger('change.select2');
            $('#percentualDiferimento').val(dados.percentualDiferimento);
            $('#isencaoIpi').val(dados.isencaoIpi).trigger('change.select2');
            $('#tipoCliente').val(dados.tipoCliente).trigger('change.select2');
            $('#tes').val(dados.tes);

            $('#contaConciliacao').val(dados.contaConciliacao).trigger('change.select2');
            $('#chaveOrdenacao').val(dados.chaveOrdenacao);
            $('#condicaoPagamento').val(dados.condicaoPagamento).trigger('change.select2');
            $('#metodoPagamento').val(dados.metodoPagamento).trigger('change.select2');
            $('#codigoJuros').val(dados.codigoJuros);
            
            $('#limiteCredito').val(parseFloat(dados.limiteCredito).toFixed(2));
            $('#volumeVendaMensalM2').val(parseFloat(dados.volumeVendaMensalM2).toFixed(2));
            $('#precoVendaNetM2').val(parseFloat(dados.precoVendaNetM2).toFixed(2));
            $('#precoVendaNetKg').val(parseFloat(dados.precoVendaNetKg).toFixed(2));

            $('#tipoFrete').val(dados.tipoFrete).trigger('change.select2');
            $('#kmDestino').val(dados.kmDestino);
            $('#diasViagem').val(dados.diasViagem);
            $('#toleranciaDiasMaior').val(dados.toleranciaDiasMaior).trigger('change.select2');
            $('#toleranciaDiasMenor').val(dados.toleranciaDiasMenor).trigger('change.select2');
            $('#toleranciaQuantidadeMaior').val(dados.toleranciaQuantidadeMaior);
            $('#toleranciaQuantidadeMenor').val(dados.toleranciaQuantidadeMenor);

            $('#cep').val(dados.cep).trigger('input');
            $('#rua').val(dados.rua);
            $('#numero').val(dados.numero);
            $('#bairro').val(dados.bairro);
            $('#estado').val(dados.estado).trigger('change.select2');
            $('#pais').val(dados.pais).trigger('change.select2');
            $('#telefone').val(dados.telefone).trigger('input');
            $('#emailXml').val(dados.emailXml);
            $('#emailOrdemPagamento').val(dados.emailOrdemPagamento);
            $('#enderecoEntregaDiferente').prop('checked', dados.enderecoEntregaDiferente == '1');
            $('#cepEntrega').val(dados.cepEntrega).trigger('input');
            $('#ruaEntrega').val(dados.ruaEntrega);
            $('#numeroEntrega').val(dados.numeroEntrega);
            $('#bairroEntrega').val(dados.bairroEntrega);
            $('#estadoEntrega').val(dados.estadoEntrega).trigger('change.select2');
            $('#kmEntrega').val(dados.kmEntrega);
            $('#enderecoCobrancaDiferente').prop('checked', dados.enderecoCobrancaDiferente == '1');
            $('#cepCobranca').val(dados.cepCobranca).trigger('input');
            $('#ruaCobranca').val(dados.ruaCobranca);
            $('#numeroCobranca').val(dados.numeroCobranca);
            $('#bairroCobranca').val(dados.bairroCobranca);
            $('#estadoCobranca').val(dados.estadoCobranca).trigger('change.select2');

            $('#observacaoEmissor').val(dados.observacaoEmissor);
            $('#observacaoCdag').val(dados.observacaoCdag);

            $('.money').each(function() {
                if ($(this).val() !== '') {
                    $(this).maskMoney('mask');
                }
            });

            if(acao === 'EDITAR') {
                $('#flagUnificarCliente').prop('checked', true);
            }

            Promise.all([
                carregarVendedores(dados.codigoPlanta),
                carregarNaturezas(dados.codigoPlanta),
                carregarCidades(dados.cidade),
                carregarCidadesEntrega(dados.cidadeEntrega),
                carregarCidadesCobranca(dados.cidadeCobranca)
            ]).then(function() {
                $('#cidade').val(dados.cidade).trigger('change.select2');
                $('#cidadeEntrega').val(dados.cidadeEntrega).trigger('change.select2');
                $('#cidadeCobranca').val(dados.cidadeCobranca).trigger('change.select2');
                $('#codigoVendedor').val(dados.codigoVendedor).trigger('change.select2');
                $('#natureza').val(dados.natureza).trigger('change.select2');
            });

            if(new Date(`${dados.dataInclusao} 23:59:59`) <= new Date('2026-06-22 00:00:00')) {
                flagValidaDocumentos = false
            }

            aplicarCamposCondicionais();

            listaAnexos = dados.anexos || [];
            renderizarAnexos();
            configurarModoFormulario();
        }

        function irProximo() {
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior(){
            navegarParaAba(abaAtual - 1);
        }

        function calcularDiasViagem(distanciaKM = 0) {            
            let resultado = Math.ceil(distanciaKM / 35)
            resultado = Math.ceil(resultado / 8);

            $('#diasViagem').val(resultado);
        }

        function anexarEventos() {
            $('#codigoPlanta').on('change', function() {
                const valor = $(this).val();
                carregarVendedores(valor);
                carregarNaturezas(valor);
            })

            $('#tipoSolicitacao, #contaConciliacao, #tipoPessoa, #grupoContas, #contribuinteIcms, #icmsDiferido, #isencaoIpi, #enderecoEntregaDiferente, #enderecoCobrancaDiferente, #flagUnificarCliente').on('change', function() {
                if (this.id === 'enderecoEntregaDiferente') {
                    $(`#cepEntrega, #ruaEntrega, #numeroEntrega, #bairroEntrega, #kmEntrega`).val('');
                    $(`#estadoEntrega, #cidadeEntrega`).val('').trigger('change');
                }

                if (this.id === 'enderecoCobrancaDiferente') {
                    $(`#cepCobranca, #ruaCobranca, #numeroCobranca, #bairroCobranca, #kmCobranca`).val('');
                    $(`#estadoCobranca, #cidadeCobranca`).val('').trigger('change');
                }

                aplicarCamposCondicionais();
            });

            $('#estado').on('change', function() {
                carregarCidades();
            });

            $('#estadoEntrega').on('change', function() {
                carregarCidadesEntrega();
            });

            $('#estadoCobranca').on('change', function() {
                carregarCidadesCobranca();
            });

            $('#cep').on('blur', function() {
                preencherEnderecoPorCep($(this).val(), '');
            });

            $('#cepEntrega').on('blur', function() {
                preencherEnderecoPorCep($(this).val(), 'Entrega');
            });

            $('#cepCobranca').on('blur', function() {
                preencherEnderecoPorCep($(this).val(), 'Cobranca');
            });

            $('#kmDestino').on('blur',  function() {
                calcularDiasViagem($(this).val());
            });
            $('#inscricaoEstadual').on('input change', function() {
                atualizarObrigatoriedadeDocumentos();
            });

            $('#btnSalvar').on('click', function() {
                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });

            $('#formularioUnificado').on('submit', function(e) {
                e.preventDefault();
                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));
                if (tabNumber === abaAtual) return;

                navegarParaAba(tabNumber);
            });

            $('#uploadAnexos').on('upload-success delete-success', function() {
                atualizarContadorAnexos();
            });

            $('#uploadAnexos').on('error', function(e) {
                Toast.error(e.detail?.message || 'Erro ao processar anexo opcional.');
            });
        }

        async function validarFormulario() {
            if (formularioSomenteVisualizacao()) return;

            const abas = [1, 2, 3, 4, 5];
            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                validacoes.aba2.revalidate(),
                validacoes.aba3.revalidate(),
                validacoes.aba4.revalidate(),
                validacoes.aba5.revalidate(),
            ]);

            resultados.forEach((valida, index) => {
                const numeroAba = abas[index];

                if (!valida) {
                    marcarAbaComoInvalida(numeroAba);
                    exibirErrosValidacao(validacoes[`aba${numeroAba}`]);
                } else {
                    marcarAbaComoCompleta(numeroAba);
                }
            });

            if (!resultados.every(Boolean)) {
                const primeiraInvalida = abas[resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return;
            }

            const documentos = coletarDocumentos();
            if (documentos === null && flagValidaDocumentos) {
                marcarAbaComoInvalida(6);
                navegarParaAba(6);
                Toast.error('Existem documentos obrigatórios pendentes. Verifique a seção Anexos.');
                return;
            }

            marcarAbaComoCompleta(6);
            enviarFormulario(documentos);
        }

        function enviarFormulario(documentos) {
            if (formularioSomenteVisualizacao()) return;

            const dadosFormulario = {
                codigoPlanta: $('#codigoPlanta').val(),
                tipoSolicitacao: $('#tipoSolicitacao').val(),
                eshop: $('#eshop').val(),
                grupoContas: $('#grupoContas').val(),
                codigoHfm: $('#codigoHfm').val(),
                tipoPessoa: $('#tipoPessoa').val(),
                razaoSocial: $('#razaoSocial').val(),
                nomeFantasia: $('#nomeFantasia').val(),
                cnpj: $('#cnpj').val().replace(/[^a-zA-Z0-9\s]/g, ''),
                cpf: $('#cpf').val().replace(/\D/g, ''),
                inscricaoEstadual: $('#inscricaoEstadual').val(),
                inscricaoMunicipal: $('#inscricaoMunicipal').val(),
                rua: $('#rua').val(),
                numero: $('#numero').val(),
                bairro: $('#bairro').val(),
                cep: $('#cep').val().replace(/\D/g, ''),
                estado: $('#estado').val(),
                cidade: $('#cidade').val(),
                pais: $('#pais').val(),
                kmDestino: $('#kmDestino').val(),
                idioma: $('#idioma').val(),
                telefone: $('#telefone').val().replace(/\D/g, ''),
                emailXml: $('#emailXml').val(),
                emailOrdemPagamento: $('#emailOrdemPagamento').val(),
                declarantePisCofins: $('#declarantePisCofins').val(),
                contribuinteIcms: $('#contribuinteIcms').val(),
                icmsDiferido: $('#icmsDiferido').val(),
                percentualDiferimento: $('#percentualDiferimento').val(),
                isencaoIpi: $('#isencaoIpi').val(),
                tipoIndustria: $('#tipoIndustria').val(),
                codigoSuframa: $('#codigoSuframa').val(),
                contaConciliacao: $('#contaConciliacao').val(),
                chaveOrdenacao: $('#chaveOrdenacao').val(),
                condicaoPagamento: $('#condicaoPagamento').val(),
                metodoPagamento: $('#metodoPagamento').val(),
                codigoJuros: $('#codigoJuros').val(),
                limiteCredito: dinheiroParaNumero($('#limiteCredito').val() || 0),
                segmento: $('#segmento').val(),
                volumeVendaMensalM2: $('#volumeVendaMensalM2').val(),
                precoVendaNetM2: dinheiroParaNumero($('#precoVendaNetM2').val() || 0),
                precoVendaNetKg: dinheiroParaNumero($('#precoVendaNetKg').val() || 0),
                codigoSap: $('#codigoSap').val(),
                natureza: $('#natureza').val(),
                tipoCliente: $('#tipoCliente').val(),
                tipoFrete: $('#tipoFrete').val(),
                tes: $('#tes').val(),
                codigoVendedor: $('#codigoVendedor').val(),
                diasViagem: $('#diasViagem').val(),
                toleranciaDiasMaior: $('#toleranciaDiasMaior').val(),
                toleranciaDiasMenor: $('#toleranciaDiasMenor').val(),
                enderecoEntregaDiferente: $('#enderecoEntregaDiferente').is(':checked') ? 'S' : 'N',
                ruaEntrega: $('#ruaEntrega').val(),
                numeroEntrega: $('#numeroEntrega').val(),
                bairroEntrega: $('#bairroEntrega').val(),
                cepEntrega: $('#cepEntrega').val().replace(/\D/g, ''),
                estadoEntrega: $('#estadoEntrega').val(),
                cidadeEntrega: $('#cidadeEntrega').val(),
                kmEntrega: $('#kmEntrega').val(),
                enderecoCobrancaDiferente: $('#enderecoCobrancaDiferente').is(':checked') ? 'S' : 'N',
                ruaCobranca: $('#ruaCobranca').val(),
                numeroCobranca: $('#numeroCobranca').val(),
                bairroCobranca: $('#bairroCobranca').val(),
                cepCobranca: $('#cepCobranca').val().replace(/\D/g, ''),
                estadoCobranca: $('#estadoCobranca').val(),
                cidadeCobranca: $('#cidadeCobranca').val(),
                toleranciaQuantidadeMaior: $('#toleranciaQuantidadeMaior').val(),
                toleranciaQuantidadeMenor: $('#toleranciaQuantidadeMenor').val(),
                codigoProtheus: $('#codigoProtheus').val(),
                codigoLoja: $('#codigoLoja').val(),
                observacaoEmissor: $('#observacaoEmissor').val(),
                observacaoCdag: $('#observacaoCdag').val(),
                anexos: JSON.stringify(documentos)
            };

            if (codigo) {
                dadosFormulario.flagUnificarCliente = $('#flagUnificarCliente').is(':checked') ? '1' : '0',
                dadosFormulario.codigo = codigo;
                dadosFormulario.acao = 'editar';
            } else {
                dadosFormulario.acao = 'criar';
            }

            exibirLoading(true);

            $.ajax({
                url: 'UnificadoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem);

                    setTimeout(function() {
                        navegarParaGrade();
                    }, 500);
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function obterNivelUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'obterNivelUsuario',
                    },
                    dataType: 'json',
                    success: function(response) {
                        exibirLoading(false);
                        resolve(response); 
                    },
                    error: function(xhr, status, error) {
                        exibirLoading(false);
                        reject(error); 
                    }
                });
            });
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            codigo = urlParams.get('codigo');
            acao = normalizarAcao(urlParams.get('acao'));

            resposta = await obterNivelUsuario();
            if (!resposta.sucesso) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = 'index.asp';
            }
        
            nivelUsuario = resposta.dados?.nivel;
            if (!nivelUsuario) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = 'index.asp';
            }
       
            anexarEventos();
            navegarParaAba(1);

            inicializarSelect2();
            inicializarMascaras();

            await Promise.all([
                carregarEstados(),
                carregarSegmentos(),
                carregarContasConciliacao(),
                carregarCondicoesPagamento(),
                carregarTiposIndustria(),
                carregarGruposConta(),
                carregarTiposCliente(),
                carregarTiposContribuinte(),
                carregarTiposDeclarante(),
                carregarMetodosPagamento(),
                carregarTiposFrete()
            ]);

            if (codigo) {
                await carregarDadosCadastro();

                if (formularioSomenteVisualizacao()) {
                    $('#titulo').text(`Visualizar Cadastro #${codigo}`);
                } else {
                    $('#titulo').text(`Editar Cadastro #${codigo}`);
                }
            } else {
                $('#titulo').text('Cadastro de Cliente');
            }

            inicializarValidacoes();
            aplicarCamposCondicionais();
            configurarModoFormulario();
        });
        
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex my-4 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5">
                    <i class="bi bi-arrow-left"></i>
                </button>
                <h3 id="titulo" class="mb-0"></h3>
            </div>

            <div id="divAcoes" class="d-flex mb-0 page-actions flex-wrap">
                <button type="button" class="btn btn-sm btn-success" id="btnSalvar">
                    <i class="bi bi-check-lg me-1"></i>
                    Enviar Formulário
                </button>
            </div>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="form-container">
            <div class="form-layout">
                <div class="tabs-container">
                    <div class="tabs">
                        <div class="tab-group-title">Geral</div>

                        <button type="button" class="tab-button active" data-tab="1">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Dados Gerais</span>
                        </button>

                        <div class="tab-group-title">Cadastro</div>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-receipt"></i>
                            <span class="tab-title">Fiscal</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-cash-coin"></i>
                            <span class="tab-title">Financeiro</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="4">
                            <i class="tab-icon bi bi-truck"></i>
                            <span class="tab-title">Logística</span>
                        </button>
                        
                        <button type="button" class="tab-button" data-tab="5">
                            <i class="tab-icon bi bi-geo-alt"></i>
                            <span class="tab-title">Endereço e Contatos</span>
                        </button>
                        
                        <!-- <button type="button" class="tab-button" data-tab="6">
                            <i class="tab-icon bi bi-person-lines-fill"></i>
                            <span class="tab-title">Protheus</span>
                        </button> -->

                        <div class="tab-group-title">Adicionais</div>

                        <button type="button" class="tab-button" data-tab="6">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form style="overflow:hidden;" id="formularioUnificado" name="formularioUnificado" novalidate>  
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Dados Gerais</div>
                                <p class="section-subtitle">Dados da solicitação e identificação principal do cliente.</p>

                                <div id="linhaUnificarCliente" class="my-2 d-none">
                                    <div class="form-check form-switch">
                                        <input class="form-check-input" type="checkbox" id="flagUnificarCliente" name="flagUnificarCliente" value="1">
                                        <label class="form-check-label" for="flagUnificarCliente">Unificar Cliente</label>
                                    </div>
                                </div>
                            </div>

                            <div id="avisoUnificar" class="alert alert-primary alert-dismissible py-2 px-3 mb-3 align-items-center" role="alert" style="font-size: 0.85rem; display: none;">
                                <i class="bi bi-info-circle me-2"></i>
                                Ao salvar o formulário o cliente será automaticamente unificado
                            </div>

                            <div class="row g-3">
                                <div class="col-12">
                                    <label for="codigoPlanta" class="form-label">Planta <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="codigoPlanta" id="codigoPlanta">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="tipoSolicitacao" class="form-label">Tipo de Cadastro <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoSolicitacao" id="tipoSolicitacao">
                                        <option value="">Selecione um item</option>
                                        <option value="1">Cadastro de Cliente</option>
                                        <option value="5">Cadastro de Cliente Prospect</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="eshop" class="form-label">Eshop <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="eshop" id="eshop">
                                        <option selected value="0">Não</option>
                                        <option value="1">Sim</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="grupoContas" class="form-label">Grupo de Contas <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="grupoContas" id="grupoContas">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="contaConciliacao" class="form-label">Conta Conciliação <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="contaConciliacao" id="contaConciliacao">
                                    </select>
                                </div>
                                <div class="col-12 col-md-4 campo-intercompany" style="display:none;">
                                    <label for="codigoHfm" class="form-label">Código HFM para clientes Intercompany <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="codigoHfm" id="codigoHfm" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="tipoIndustria" class="form-label">Tipo de Indústria <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoIndustria" id="tipoIndustria">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="segmento" class="form-label">Segmento <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="segmento" id="segmento"></select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="natureza" class="form-label">Natureza <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="natureza" id="natureza"></select>
                                </div>

                                <div class="col-12"><hr/></div>

                                <div class="col-12 col-md-4">
                                    <label for="tipoPessoa" class="form-label">Tipo de Pessoa <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoPessoa" id="tipoPessoa">
                                        <option selected value="J">Pessoa Jurídica</option>
                                        <option value="F">Pessoa Física</option>
                                    </select>
                                </div>
                                <div style="display:none;" class="campos-pessoa-fisica col-12 col-md-4">
                                    <label for="cpf" class="form-label">CPF <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="cpf" id="cpf" />
                                </div>
                                <div class="campos-pessoa-juridica col-12 col-md-4">
                                    <label for="cnpj" class="form-label">CNPJ <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="cnpj" id="cnpj" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="inscricaoEstadual" class="form-label">Inscrição Estadual</label>
                                    <input type="number" class="form-control form-control-sm" name="inscricaoEstadual" id="inscricaoEstadual" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="inscricaoMunicipal" class="form-label">Inscrição Municipal</label>
                                    <input type="number" class="form-control form-control-sm" name="inscricaoMunicipal" id="inscricaoMunicipal" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="razaoSocial" class="form-label">Razão Social <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="nomeFantasia" class="form-label">Nome Fantasia <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="nomeFantasia" id="nomeFantasia" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoSuframa" class="form-label">Código Suframa</label>
                                    <input type="number" class="form-control form-control-sm" name="codigoSuframa" id="codigoSuframa" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="idioma" class="form-label">Idioma <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="idioma" id="idioma">
                                        <option selected value="PT-BR">Português</option>
                                        <option value="ES">Espanhol</option>
                                        <option value="EN">Inglês</option>
                                    </select>
                                </div>
                                 <div class="col-12 col-md-4">
                                    <label for="codigoSap" class="form-label">Código SAP</label>
                                    <input type="text" class="form-control form-control-sm" name="codigoSap" id="codigoSap" readonly />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoProtheus" class="form-label">Código Protheus</label>
                                    <input type="text" class="form-control form-control-sm" name="codigoProtheus" id="codigoProtheus" readonly />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoLoja" class="form-label">Loja</label>
                                    <input type="text" class="form-control form-control-sm" name="codigoLoja" id="codigoLoja" readonly />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Fiscal</div>
                                <p class="section-subtitle">Informações fiscais, tributárias e parâmetros fiscais do Protheus.</p>
                            </div>

                            <div class="row g-3">
                                 <div class="col-md-4">
                                    <label for="declarantePisCofins" class="form-label">Declarante de PIS/COFINS</label>
                                    <select class="form-select form-select-sm select2" name="declarantePisCofins" id="declarantePisCofins">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="contribuinteIcms" class="form-label">Contribuinte de ICMS</label>
                                    <select class="form-select form-select-sm select2" name="contribuinteIcms" id="contribuinteIcms">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-2 campo-cliente-final">
                                    <label for="icmsDiferido" class="form-label">ICMS Diferido <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="icmsDiferido" id="icmsDiferido">
                                        <option value="0">Não</option>
                                        <option value="1">Sim</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-2 campo-diferimento" style="display:none;">
                                    <label for="percentualDiferimento" class="form-label">% Diferimento</label>
                                    <input type="number" class="form-control form-control-sm" name="percentualDiferimento" id="percentualDiferimento" min="0" max="100" step="0.01" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="isencaoIpi" class="form-label">Suspensão/Isenção IPI <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="isencaoIpi" id="isencaoIpi">
                                        <option value="0">Não</option>
                                        <option value="1">Sim</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="tipoCliente" class="form-label">Tipo</label>
                                    <select class="form-select form-select-sm select2" name="tipoCliente" id="tipoCliente">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="tes" class="form-label">TES</label>
                                    <input type="number" class="form-control form-control-sm" name="tes" id="tes" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">Financeiro</div>
                                <p class="section-subtitle">Dados de pagamento, crédito, condições comerciais.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="chaveOrdenacao" class="form-label">Chave de Ordenação <span class="required-field">*</span></label>
                                    <input disabled type="text" value="0009" class="form-control form-control-sm" name="chaveOrdenacao" id="chaveOrdenacao" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="condicaoPagamento" class="form-label">Condição de Pagamento <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="condicaoPagamento" id="condicaoPagamento"></select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="metodoPagamento" class="form-label">Método de Pagamento <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="metodoPagamento" id="metodoPagamento">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoJuros" class="form-label">Código de Juros <span class="required-field">*</span></label>
                                    <input value="BR" disabled type="text" class="form-control form-control-sm" name="codigoJuros" id="codigoJuros" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="limiteCredito" class="form-label">Limite de Crédito <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm money" name="limiteCredito" id="limiteCredito" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="volumeVendaMensalM2" class="form-label">Volume mensal (M²) <span class="required-field">*</span></label>
                                    <input type="number" class="form-control form-control-sm" name="volumeVendaMensalM2" id="volumeVendaMensalM2" min="0" step="0.01" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="precoVendaNetM2" class="form-label">Preço net R$/M² <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm money" name="precoVendaNetM2" id="precoVendaNetM2" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="precoVendaNetKg" class="form-label">Preço net R$/KG <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm money" name="precoVendaNetKg" id="precoVendaNetKg" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="4">
                            <div class="section-header">
                                <div class="section-title">Logística</div>
                                <p class="section-subtitle">Dados de logística e tolerâncias.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="tipoFrete" class="form-label">Tipo de Frete <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoFrete" id="tipoFrete">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoVendedor" class="form-label">Código do Vendedor <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="codigoVendedor" id="codigoVendedor"></select>
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="kmDestino" class="form-label">KM até o destino <span class="required-field">*</span></label>
                                    <input type="number" class="form-control form-control-sm" name="kmDestino" id="kmDestino" min="0" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="diasViagem" class="form-label">Dias Viagem 
                                        <span class="required-field">*</span>
                                        <button type="button" 
                                            class="btn btn-link p-0 m-0 align-baseline" 
                                            data-toggle="tooltip" 
                                            data-placement="right" 
                                            title="Refere-se ao período necessário para que uma carga seja transportada de um ponto de origem até o seu destino final.   Esse tempo inclui todas as etapas do processo de transporte, desde a saída do local inicial, passando por eventuais pontos intermiediarios, ate a chegada ao destino.">
                                            <i class="bi bi-question-circle"></i>
                                        </button>
                                    </label>
                                    <input disabled type="number" class="form-control form-control-sm" name="diasViagem" id="diasViagem" min="0" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="toleranciaDiasMaior" class="form-label">Tolerância de Dias para Entrega Maior
                                        <span class="required-field">*</span>
                                            <button type="button" 
                                                class="btn btn-link p-0 m-0 align-baseline" 
                                                data-toggle="tooltip" 
                                                data-placement="right" 
                                                title="Período adicional permitido além do prazo inicialmente previsto para a entrega de um produto quando acordado com o cliente. Essas margem serve para acomodar possíveis atrasos inesperados no processo logístico, como imprevistos no transporte, condições climáticas adversas, ou dificuldades operacionais, sem que isso seja considerado uma falha no cumprimento do prazo.">
                                                <i class="bi bi-question-circle"></i>
                                            </button>
                                    </label>
                                    <select class="form-select form-select-sm select2" name="toleranciaDiasMaior" id="toleranciaDiasMaior">
                                        <option value="">Selecione um item</option>
                                        <option>0</option>
                                        <option>1</option>
                                        <option>2</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="toleranciaDiasMenor" class="form-label">Tolerância de Dias para Entrega Menor
                                        <span class="required-field">*</span>
                                        <button type="button" 
                                            class="btn btn-link p-0 m-0 align-baseline" 
                                            data-toggle="tooltip" 
                                            data-placement="right" 
                                            title="Período adicional permitido além do prazo inicialmente previsto para a entrega de um produto quando acordado com o cliente. Essas margem serve para acomodar possíveis atrasos inesperados no processo logístico, como imprevistos no transporte, condições climáticas adversas, ou dificuldades operacionais, sem que isso seja considerado uma falha no cumprimento do prazo.">
                                            <i class="bi bi-question-circle"></i>
                                        </button>
                                    </label>
                                   
                                    <select class="form-select form-select-sm select2" name="toleranciaDiasMenor" id="toleranciaDiasMenor">
                                        <option value="">Selecione um item</option>
                                        <option value="0">0</option>
                                        <option value="1">-1</option>
                                        <option value="2">-2</option>
                                        <option value="3">-3</option>
                                        <option value="4">-4</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="toleranciaQuantidadeMaior" class="form-label">Tolerância qtd. maior (%) <span class="required-field">*</span></label>
                                    <input value="10" type="number" class="form-control form-control-sm" name="toleranciaQuantidadeMaior" id="toleranciaQuantidadeMaior" step="0.01" />
                                </div>
                                <div class="col-12 col-md-4 campo-cliente-final">
                                    <label for="toleranciaQuantidadeMenor" class="form-label">Tolerância qtd. menor (%) <span class="required-field">*</span></label>
                                    <input value="10" type="number" class="form-control form-control-sm" name="toleranciaQuantidadeMenor" id="toleranciaQuantidadeMenor" step="0.01" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="5">
                            <div class="section-header">
                                <div class="section-title">Endereço e Contatos</div>
                                <p class="section-subtitle">Endereço principal e contatos operacionais do cliente.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="cep" class="form-label">CEP <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="cep" id="cep" />
                                </div>
                                <div class="col-12 col-md-5">
                                    <label for="rua" class="form-label">Rua <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="rua" id="rua" />
                                </div>
                                <div class="col-12 col-md-3">
                                    <label for="numero" class="form-label">Número <span class="required-field">*</span></label>
                                    <input type="number" class="form-control form-control-sm" name="numero" id="numero" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="bairro" class="form-label">Bairro <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="bairro" id="bairro" />
                                </div>
                                <div class="col-12 col-md-3">
                                    <label for="estado" class="form-label">Estado <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="estado" id="estado" data-cidade-target="cidade"></select>
                                </div>
                                <div class="col-12 col-md-3">
                                    <label for="cidade" class="form-label">Cidade <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="cidade" id="cidade"></select>
                                </div>
                                <div class="col-12 col-md-2">
                                    <label for="pais" class="form-label">País <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="pais" id="pais">
                                        <option value="">Selecione um item</option>
                                        <option value="Brasil" selected>Brasil</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="telefone" class="form-label">Telefone</label>
                                    <input type="text" class="form-control form-control-sm" name="telefone" id="telefone" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="emailXml" class="form-label">E-mail XML <span class="required-field">*</span></label>
                                    <input type="email" class="form-control form-control-sm" name="emailXml" id="emailXml" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="emailOrdemPagamento" class="form-label">E-mail Ordem Pagamento <span class="required-field">*</span></label>
                                    <input type="email" class="form-control form-control-sm" name="emailOrdemPagamento" id="emailOrdemPagamento" />
                                </div>

                                <div class="col-12"><hr/></div>

                                <div class="col-12">
                                    <div class="form-check form-switch">
                                        <input class="form-check-input" type="checkbox" name="enderecoEntregaDiferente" id="enderecoEntregaDiferente" />
                                        <label class="form-check-label" for="enderecoEntregaDiferente">Endereço de Entrega diferente do cadastro</label>
                                    </div>
                                </div>
                                <div class="grupo-endereco-entrega" style="display:none;">
                                    <div class="row g-3">
                                        <div class="col-12 col-md-4">
                                            <label for="cepEntrega" class="form-label">CEP <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="cepEntrega" id="cepEntrega" />
                                            <small class="text-muted">Digite o CEP para buscar o endereço</small>
                                        </div>
                                        <div class="col-12 col-md-5">
                                            <label for="ruaEntrega" class="form-label">Rua <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="ruaEntrega" id="ruaEntrega" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="numeroEntrega" class="form-label">Número <span class="required-field">*</span></label>
                                            <input type="number" class="form-control form-control-sm" name="numeroEntrega" id="numeroEntrega" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="bairroEntrega" class="form-label">Bairro <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="bairroEntrega" id="bairroEntrega" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="estadoEntrega" class="form-label">Estado <span class="required-field">*</span></label>
                                            <select class="form-select form-select-sm select2" name="estadoEntrega" id="estadoEntrega" data-cidade-target="cidadeEntrega"></select>
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="cidadeEntrega" class="form-label">Cidade <span class="required-field">*</span></label>
                                            <select class="form-select form-select-sm select2" name="cidadeEntrega" id="cidadeEntrega"></select>
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="kmEntrega" class="form-label">KM <span class="required-field">*</span></label>
                                            <input type="number" class="form-control form-control-sm" name="kmEntrega" id="kmEntrega" min="0" />
                                        </div>
                                    </div>
                                </div>
                                <div class="col-12">
                                    <div class="form-check form-switch">
                                        <input class="form-check-input" type="checkbox" name="enderecoCobrancaDiferente" id="enderecoCobrancaDiferente" />
                                        <label class="form-check-label" for="enderecoCobrancaDiferente">Endereço de Cobrança diferente do cadastro</label>
                                    </div>
                                </div>
                                <div class="grupo-endereco-cobranca" style="display:none;">
                                    <div class="row g-3">
                                        <div class="col-12 col-md-4">
                                            <label for="cepCobranca" class="form-label">CEP <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="cepCobranca" id="cepCobranca" />
                                            <small class="text-muted">Digite o CEP para buscar o endereço</small>
                                        </div>
                                        <div class="col-12 col-md-5">
                                            <label for="ruaCobranca" class="form-label">Rua <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="ruaCobranca" id="ruaCobranca" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="numeroCobranca" class="form-label">Número <span class="required-field">*</span></label>
                                            <input type="number" class="form-control form-control-sm" name="numeroCobranca" id="numeroCobranca" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="bairroCobranca" class="form-label">Bairro <span class="required-field">*</span></label>
                                            <input type="text" class="form-control form-control-sm" name="bairroCobranca" id="bairroCobranca" />
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="estadoCobranca" class="form-label">Estado <span class="required-field">*</span></label>
                                            <select class="form-select form-select-sm select2" name="estadoCobranca" id="estadoCobranca" data-cidade-target="cidadeCobranca"></select>
                                        </div>
                                        <div class="col-12 col-md-3">
                                            <label for="cidadeCobranca" class="form-label">Cidade <span class="required-field">*</span></label>
                                            <select class="form-select form-select-sm select2" name="cidadeCobranca" id="cidadeCobranca"></select>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="6">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Evidências e documentos relacionados ao cadastro. Para ICMS diferido ou isenção de IPI, anexe os documentos comprobatórios.</p>
                            </div>

                            <div style="display: grid; grid-template-columns: 1fr 1fr;" class="row g-3">
                                <document-upload
                                    id="doc_cartao_cnpj"
                                    class="campos-pessoa-juridica"
                                    name="cartao_cnpj"
                                    label="Cartão CNPJ"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_documento_foto"
                                    class="campos-pessoa-fisica"
                                    name="documento_foto"
                                    label="Documento com Foto (RG ou carteira de motorista)"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_comprovante_endereco"
                                    name="comprovante_endereco"
                                    label="Comprovante de Endereço"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_consulta_sintegra"
                                    class="campos-pessoa-juridica"
                                    name="consulta_sintegra"
                                    label="Consulta SINTEGRA"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_contrato_social"
                                    class="campos-pessoa-juridica"
                                    name="contrato_social"
                                    label="Contrato Social e Alterações Contratuais"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_declaracao_ir"
                                    class="campos-pessoa-fisica"
                                    name="declaracao_ir"
                                    label="Declaração do Imposto de renda"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_inscricao_estadual"
                                    name="inscricao_estadual"
                                    label="Inscrição estadual"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_carta_isencao_ipi"
                                    name="carta_isencao_ipi"
                                    label="Carta de isenção de IPI"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <document-upload
                                    id="doc_comprovatorio_icms_diferido"
                                    class="campo-diferimento"
                                    name="comprovatorio_icms_diferido"
                                    label="Documento comprobatório de deferimento"
                                    file-code="UNIFICADO_"
                                    required>
                                </document-upload>

                                <div class="col-12" style="grid-column: 1 / -1;">
                                    <hr/>
                                    <label class="form-label mb-1" for="uploadAnexos">Outros documentos</label>
                                    <p class="section-subtitle mb-2">Anexe documentos complementares, se necessário. Este envio é opcional.</p>
                                    <file-upload id="uploadAnexos" file-code="UNIFICADO_" max-files="10" multiple></file-upload>
                                </div>

                                <!--  
                                <div class="col-12">
                                    <label for="observacaoEmissor" class="form-label">Observação de arquivos enviados (Emissor)</label>
                                    <textarea class="form-control form-control-sm" name="observacaoEmissor" id="observacaoEmissor" rows="3" readonly></textarea>
                                </div>
                                <div class="col-12">
                                    <label for="observacaoCdag" class="form-label">Observação de arquivos recebidos (CDAG)</label>
                                    <textarea class="form-control form-control-sm" name="observacaoCdag" id="observacaoCdag" rows="3" readonly></textarea>
                                </div> 
                                -->
                            </div>
                        </div>
                    </form>

                    <div class="nav-footer">
                        <button type="button" id="btnVoltarPagina" class="btn btn-sm btn-outline-secondary" onclick="irAnterior()">
                            <i class="bi bi-arrow-left"></i> Voltar
                        </button>
                        <button type="button" id="btnProximaPagina" class="btn btn-sm btn-primary" onclick="irProximo()">
                            Próximo <i class="bi bi-arrow-right"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
