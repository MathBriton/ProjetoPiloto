<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
<!--#Include File="../../includes/variaveis.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

nivelSelecionado = Request("nivelSelecionado")
If nivelSelecionado <> "" Then
    SESSION("NIVELSPD") = nivelSelecionado
End If

nivelUsuario = SESSION("NIVELSPD")
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

         .z-index-9999 {
            z-index: 9999 !important;
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

        .barra-teste {
            background-color: #fef9c3;
            border-bottom: 2px solid #facc15;
            padding: 6px 16px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: .8rem;
            font-weight: 600;
            color: #713f12;
            position: sticky;
            top: 0;
            z-index: 1050;
        }

        .barra-teste select {
            font-size: .8rem;
            font-weight: 600;
            padding: 2px 8px;
            border-radius: 6px;
            border: 1.5px solid #facc15;
            background-color: #fff;
            color: #713f12;
            cursor: pointer;
        }

        .barra-teste select:focus {
            outline: none;
            box-shadow: 0 0 0 2px #fde68a;
        }

        .barra-teste .nivel-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            background-color: #fde68a;
            border-radius: 20px;
            padding: 2px 10px;
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
        .modal select.is-invalid + .select2-container .select2-selection {
            border-color: #dc3545 !important;
        }
        .modal .invalid-feedback {
            display: none;
            font-size: .8rem;
        }
        .modal .is-invalid ~ .invalid-feedback,
        .modal select.is-invalid + .select2-container + .invalid-feedback {
            display: block;
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
        let abaAtual = 1;
        let totalAbas = 7;
        let dadosSolicitacao = null;

        const eNivel = Object.freeze({
            COMERCIAL: 'C',
            DESENVOLVIMENTO: 'D',
            PCP: 'P'
        })

        let codPlanta = ''
        let nivelUsuario = '<%= nivelUsuario %>';

        function alterarNivelTeste(valor) {
            nivelUsuario = valor;
            const label = valor === eNivel.COMERCIAL ? 'Comercial' :
                valor === eNivel.DESENVOLVIMENTO ? 'Desenvolvimento' :
                valor === eNivel.PCP ? 'PCP' : '';
            $('#labelNivelAtivo').text(label);
            configurarCampoProcessoAprovacao();
        }

        function exibirLoading(exibir) {
            $('#loadingOverlay').toggleClass('d-none', !exibir).toggleClass('d-flex', exibir);
        }

        function navegarParaGrade() {
            window.location.href = `SolicitacaoProjetoGrade.asp?codPlanta=${codPlanta}`;
        }

        function navegarParaLotePiloto() {
            window.location.href = '/corporativo/Lotepiloto/LotePilotoGrade.asp?paramPlanta=<%=SESSION("FILIAL")%>'
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

        function carregarDadosSolicitacao() {
            const params = { 
                acao: 'obterPorId', 
                id: id, 
                codPlanta: codPlanta
            }

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp',
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
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }
            });
        }

        function carregarDesenvolvedores() {
            const $select = $('#idDesenvolvedor');
            $select.empty().append('<option value="">Selecione o desenvolvedor...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarDesenvolvedores',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar desenvolvedores: " + response.mensagem)
                        Toast.error("Erro ao carregar desenvolvedores: " + response.mensagem)
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar desenvolvedores:', error);
                    Toast.error('Erro ao carregar desenvolvedores');
                }
            });
        }

        function carregarEstilos() {
            const $select = $('#estilo');
            $select.empty().append('<option value="">Selecione o estilo...</option>');

            const $selectModal = $('#acessorioEstilo');
            $selectModal.empty().append('<option value="">Selecione o estilo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarEstilos',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar estilos: " + response.mensagem)
                        Toast.error("Erro ao carregar estilos: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        const option = { value: item.codigo, text: `${item.codigo} - ${item.descricao}` }
                        $select.append($('<option>', option));
                        $selectModal.append($('<option>', option));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar estilos:', error);
                    Toast.error('Erro ao carregar estilos');
                }
            });
        }

        function carregarEstruturasPapelao() {
            const $select = $('#estruturaPapelao');
            $select.empty().append('<option value="">Selecione a estrutura...</option>');
            
            const $selectModal = $('#acessorioEstrutura');
            $selectModal.empty().append('<option value="">Selecione a estrutura...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarEstruturasPapelao',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar estruturas de papelão:', response.mensagem);
                        Toast.error('Erro ao carregar estruturas de papelão');
                        return;
                    }
                    response.dados.forEach(function(item) {
                        const option = { value: item.codigo, text: item.codigo, 'data-gramatura': item.gramatura }
                        $select.append($('<option>', option));
                        $selectModal.append($('<option>', option));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar estruturas de papelão:', error);
                    Toast.error('Erro ao carregar estruturas de papelão');
                }
            });
        }

        function carregarConteudosEmbalado() {
            const $select = $('#conteudoEmbalado');
            $select.empty().append('<option value="">Selecione o conteúdo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarConteudosEmbalado',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar conteúdos: " + response.mensagem)
                        Toast.error("Erro ao carregar conteúdos: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar conteúdos:', error);
                    Toast.error('Erro ao carregar conteúdos');
                }
            });
        }

        function carregarTiposProdutoPrimario() {
            const $select = $('#tipoProdutoPrimario');
            $select.empty().append('<option value="">Selecione o tipo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarTiposProdutoPrimario',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar tipos de produto primário: " + response.mensagem)
                        Toast.error("Erro ao carregar tipos de produto primário: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de produto primário:', error);
                    Toast.error('Erro ao carregar tipos de produto primário');
                }
            });
        }

        function carregarPaletes() {
            const $select = $('#palete');
            $select.empty().append('<option value="">Selecione o palete...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarPaletes',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar paletes: " + response.mensagem)
                        Toast.error("Erro ao carregar paletes: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.chave, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar paletes:', error);
                    Toast.error('Erro ao carregar paletes');
                }
            });
        }

        function carregarDefeitos() {
            const $select = $('#defeitoId');
            $select.empty().append('<option value="">Todos os defeitos</option>');

            const params = {
                acao: 'listarDefeitos',
                tipoDefeito: $('#tipoDefeito').val()
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar defeitos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar defeitos:', error);
                }
            });
        }

        function carregarDisposicoes() {
            const $select = $('#disposicaoId');
            $select.empty().append('<option value="">Todas as disposições</option>');

            const params = {
                acao: 'listarDisposicoes'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar disposições: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar disposições:', error);
                }
            });
        }

        function carregarMotivosRejeicao() {
            const $select = $('#motivoRejeicao');
            $select.empty().append('<option value="">Selecione o motivo...</option>');

            const params = {
                acao: 'listarMotivosRejeicao'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar motivos de rejeição: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar motivos de rejeição:', error);
                }
            });
        }

        function carregarProcessosAprovacao(){
            const $select = $('#processoAprovacao');
            $select.empty().append('<option value="">Selecione o processo...</option>');

            const params = {
                acao: 'listarProcessosAprovacao'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar processos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar processos:', error);
                }
            });
        }
     
        function limparValidacaoModal($modal) {
            $modal.find('.is-invalid').removeClass('is-invalid');
            $modal.find('.invalid-feedback').hide();
        }

        function marcarCampoModalInvalido($campo, mensagem) {
            $campo.addClass('is-invalid');
            const $feedback = $campo.nextAll('.invalid-feedback').first();
            if ($feedback.length) {
                $feedback.text(mensagem).show();
            }
            if ($campo.hasClass('select2-hidden-accessible')) {
                $campo.next('.select2-container').find('.select2-selection').attr('aria-invalid', 'true');
            }
        }

        function limparCampoModalInvalido($campo) {
            $campo.removeClass('is-invalid');
            $campo.nextAll('.invalid-feedback').first().hide();
            if ($campo.hasClass('select2-hidden-accessible')) {
                $campo.next('.select2-container').find('.select2-selection').removeAttr('aria-invalid');
            }
        }

        function validarCampoObrigatorioModal(seletor, mensagem) {
            const $campo = $(seletor);
            const campoVisivel = $campo.hasClass('select2-hidden-accessible')
                ? $campo.next('.select2-container').is(':visible')
                : $campo.is(':visible');
            if (!$campo.length || !campoVisivel || $campo.prop('disabled')) {
                return true;
            }
            if (($campo.val() || '').toString().trim() === '') {
                marcarCampoModalInvalido($campo, mensagem);
                if ($campo.hasClass('select2-hidden-accessible')) {
                    $campo.select2('open');
                } else {
                    $campo.trigger('focus');
                }
                return false;
            }
            limparCampoModalInvalido($campo);
            return true;
        }

        function configurarCampoProcessoAprovacao() {
            const ehDesenvolvimento = nivelUsuario == eNivel.DESENVOLVIMENTO;
            const $grupo = $('#grupoProcessoAprovacao');
            const $campo = $('#processoAprovacao');

            $grupo.toggle(ehDesenvolvimento);
            $campo.prop('required', ehDesenvolvimento);
            $campo.prop('disabled', !ehDesenvolvimento);

            if (!ehDesenvolvimento) {
                $campo.val('').trigger('change.select2');
                limparCampoModalInvalido($campo);
            }
        }

        function validarModalAprovacao() {
            const $modal = $('#modalAprovacao');
            limparValidacaoModal($modal);
            if (nivelUsuario == eNivel.DESENVOLVIMENTO) {
                return validarCampoObrigatorioModal('#processoAprovacao', 'Informe o processo de aprovação.');
            }

            return true;
        }

        function validarModalRejeicao() {
            const $modal = $('#modalRejeicao');
            limparValidacaoModal($modal);
            return validarCampoObrigatorioModal('#motivoRejeicao', 'Informe o motivo da rejeição.');
        }

        function preencherFormulario(dados) {
            $('#codPlanta').val(dados.codPlanta).trigger('change.select2');

            var p1 = carregarDesenvolvedores();
            var p2 = carregarEstilos();
            var p3 = carregarEstruturasPapelao();
            var p4 = carregarConteudosEmbalado();
            var p5 = carregarTiposProdutoPrimario();
            var p6 = carregarPaletes();
        
            $.when(p1, p2, p3, p4, p5, p6).then(function() {
                $('#idDesenvolvedor').val(dados.idDesenvolvedor).trigger('change');
                $('#codCliente').val(dados.codCliente);
                $('#nomeCliente').val(dados.nomeCliente);
                $('#codLoja').val(dados.codLoja);
                $('#tipoSolicitacao').val(dados.tipoSolicitacao).trigger('change');
                $('#referencia').val(dados.referencia);
                $('#toleranciaMaxima').val(dados.toleranciaMaxima);
                $('#toleranciaMinima').val(dados.toleranciaMinima);
                $('#estilo').val(dados.estilo).trigger('change');
                $('#observacoes').val(dados.observacoes).trigger('change');
                $('#comprimento').val(dados.comprimento || 0);
                $('#largura').val(dados.largura || 0);
                $('#altura').val(dados.altura || 0);
                $('#numeroMatriz').val(dados.numeroMatriz || 0);
                $('#inkJet').val(dados.inkJet).trigger('change');

                $('#necessitaAcompanhamento').val(dados.necessitaAcompanhamento).trigger('change');
                if (dados.necessitaAcompanhamento === 'S') {
                    $('#divNivel').show();
                    $('#nivelLotePiloto').val(dados.nivelLotePiloto).trigger('change');
                }

                $('#clienteUsaTara').val(dados.clienteUsaTara).trigger('change');
                if (dados.clienteUsaTara == '1') {
                    $('#divToleranciaPeso').show();
                    $('#toleranciaPeso').val(dados.toleranciaPeso || 0);
                }

                $('#estruturaPapelao').val(dados.estruturaPapelao).trigger('change');
                $('#gramatura').val(dados.gramatura || 0);
                $('#larguraChapa').val(dados.larguraChapa || 0);
                $('#comprimentoChapa').val(dados.comprimentoChapa || 0);
                $('#areaLiquida').val(dados.areaLiquida || 0);

                $('#gramaturaReferencial').val(dados.gramaturaReferencial || 0);
                $('#esmagamento').val(dados.esmagamento || 0);
                $('#ectColuna').val(dados.ectColuna || 0);
                $('#mullenReferencial').val(dados.mullenReferencial || 0);
                $('#espessura').val(dados.espessura || 0);
                $('#bctCompressao').val(dados.bctCompressao || 0);
                $('#cobbInterno').val(dados.cobbInterno || 0);
                $('#cobbExterno').val(dados.cobbExterno || 0);

                $('#tipoFechamento').val(dados.tipoFechamento).trigger('change');
                $('#pecasPorUnidade').val(dados.pecasPorUnidade).trigger('change');

                $('#ehAmarrado').val(dados.ehAmarrado).trigger('change');
                if (dados.ehAmarrado === 'S') {
                    $('#divQuantidadeFitilhos').show();
                    $('#divTipoFitilho').show();
                    $('#quantidadeFitilhos').val(dados.quantidadeFitilhos).trigger('change');
                    $('#tipoFitilho').val(dados.tipoFitilho).trigger('change');
                }

                $('#ehPaletizado').val(dados.ehPaletizado).trigger('change');
                if (dados.ehPaletizado === 'S') {
                    $('#camposPaletizado').show();
                    $('#alturaMaximaPalete').val(dados.alturaMaximaPalete);
                    $('#palete').val(dados.palete).trigger('change');
                    $('#temFacesOpostas').val(dados.temFacesOpostas).trigger('change');
                    $('#temCantinho').val(dados.temCantinho).trigger('change');
                    $('#temGradeMadeira').val(dados.temGradeMadeira).trigger('change');
                    $('#temFilmeStrech').val(dados.temFilmeStrech).trigger('change');
                    $('#clienteUsaRack').val(dados.clienteUsaRack).trigger('change');
                    $('#overhang').val(dados.overhang).trigger('change');
                    $('#quantidadeEspelhos').val(dados.quantidadeEspelhos).trigger('change');
                }

                if (dados.acessorios) {
                    listaAcessorios = dados.acessorios;
                } 
      
                renderizarListaAcessorios();

                $('#conteudoEmbalado').val(dados.conteudoEmbalado).trigger('change');
                $('#consumoMensal').val(dados.consumoMensal || 0);
                $('#tipoLinhaProducaoCliente').val(dados.tipoLinhaProducaoCliente).trigger('change');
                $('#tipoProdutoPrimario').val(dados.tipoProdutoPrimario).trigger('change');
                $('#pesoPrimario').val(dados.pesoPrimario || 0);
                $('#quantidadePrimarios').val(dados.quantidadePrimarios || 0);
                $('#suportacaoPrimario').val(dados.suportacaoPrimario).trigger('change');
                $('#tempoEstocagem').val(dados.tempoEstocagem || 0);
                $('#umidadeRelativa').val(dados.umidadeRelativa || 0);
                $('#quantidadeCores').val(dados.quantidadeCores || 0);
                $('#precoKG').val(dados.precoKG || 0);
                $('#orcamento').val(dados.orcamento).trigger('change');
                $('#margemContribuicao').val(dados.margemContribuicao);
                $('#resultado').val(dados.resultado || 0);
                $('#tipoSavings').val(dados.tipoSavings).trigger('change');
                $('#valorSavingSK').val(dados.valorSavingSK || 0);
                $('#valorSavingCliente').val(dados.valorSavingCliente || 0);
                $('#valorEstimadoCliche').val(dados.valorEstimadoCliche || 0);
                $('#valorEstimadoMatriz').val(dados.valorEstimadoMatriz || 0);
                $('#informacoesAdicionais').val(dados.informacoesAdicionais);

                $('#quantidadeAmostras').val(dados.quantidadeAmostras).trigger('change');
                var qtdAmostras = parseInt(dados.quantidadeAmostras) || 0;
                if (qtdAmostras >= 1) {
                    $('#camposAmostra').show();
                    $('#amostraCep').val(dados.amostraCep).trigger('input');
                    $('#amostraResponsavel').val(dados.amostraResponsavel);
                    $('#amostraEndereco').val(dados.amostraEndereco);
                    $('#amostraBairro').val(dados.amostraBairro);
                    $('#amostraMunicipio').val(dados.amostraMunicipio);
                    $('#amostraUF').val(dados.amostraUF).trigger('change');
                    $('#amostraCNPJ').val(dados.amostraCNPJ).trigger('input');
                    $('#amostraTelefone').val(dados.amostraTelefone).trigger('input');
                    $('#amostraEmail').val(dados.amostraEmail);
                    $('#amostraObservacoes').val(dados.amostraObservacoes);
                }
            });
        }

        
        function renderizarListaAcessorios() {
            const $container = $('#acessoriosListaContainer');
            const $contador = $('#contadorAcessorios');
            
            $contador.text(listaAcessorios.length);

            if (listaAcessorios.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhum acessório adicionado.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div class="table-responsive">
                    <table class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th class="text-center" style="width: 50px;">#</th>
                                <th class="text-center">Qtd</th>
                                <th>Estilo</th>
                                <th class="text-center">Nº Cortes</th>
                                <th class="text-center">Largura (mm)</th>
                                <th class="text-center">Comprimento (mm)</th>
                                <th class="text-center">Estrutura</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaAcessorios.forEach((acessorio, index) => {
                html += `
                    <tr>
                        <td class="text-center">${index + 1}</td>
                        <td class="text-center">
                            <span class="badge bg-secondary">${acessorio.quantidade}</span>
                        </td>
                        <td>${acessorio.estiloDescricao || acessorio.estilo}</td>
                        <td class="text-center">${acessorio.numeroCortes}</td>
                        <td class="text-center">${acessorio.largura}</td>
                        <td class="text-center">${acessorio.comprimento}</td>
                        <td class="text-center"><code>${acessorio.estrutura}</code></td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $container.html(html);
        }

        function avaliarSolicitacao(flagRejeitar) {
            const modalValido = flagRejeitar ? validarModalRejeicao() : validarModalAprovacao();
            if (!modalValido) {
                return;
            }
            const dadosFormulario = {
                id: id,
                codPlanta: codPlanta,
                motivoRejeicao: $('#motivoRejeicao')?.val() || '',
                processoAprovacao: $('#processoAprovacao')?.val() || ''
            };

            if (nivelUsuario == eNivel.COMERCIAL) {
                dadosFormulario.acao = flagRejeitar ? 'rejeicaoComercial' : 'aprovacaoComercial'
            } else if (nivelUsuario == eNivel.DESENVOLVIMENTO) {
                dadosFormulario.acao = flagRejeitar ? 'rejeicaoDesenvolvimento' : 'aprovacaoDesenvolvimento'
            } else if (nivelUsuario == eNivel.PCP) {
                dadosFormulario.acao = flagRejeitar ? 'rejeicaoPCP' : 'aprovacaoPCP'
            }
                        
            exibirLoading(true);

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                    
                    Toast.success(response.mensagem);
                    if(dadosFormulario.acao == 'aprovacaoDesenvolvimento') {
                        confirmarFinalizacao();
                    } else {
                        navegarParaGrade(); 
                    }
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao aprovar solicitação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function confirmarFinalizacao() {
            ConfirmationModal.show({
                message: 'Sap aprovada com sucesso! Deseja criar um novo Lote Piloto?',
                type: 'info',
                confirmText: 'Sim',
                cancelText: 'Não',
                onConfirm: () => navegarParaLotePiloto(),
                onCancel: () => navegarParaGrade()
            });
        }

        function anexarEventos() {
            $('#formularioAvaliacao').on('submit', function(e) {
                e.preventDefault();
                return false;
            });
            $('#formularioAprovacao').on('submit', function(e) {
                e.preventDefault();
                avaliarSolicitacao(false);
                return false;
            });
            $('#formularioRejeicao').on('submit', function(e) {
                e.preventDefault();
                avaliarSolicitacao(true);
                return false;
            });
            $('#processoAprovacao, #motivoRejeicao').on('change select2:select', function() {
                if (($(this).val() || '').toString().trim() !== '') {
                    limparCampoModalInvalido($(this));
                }
            });
            
            $('#acao').on('change', function() { 
                campoAcaoAlterado()
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
                abrirModalAprovacao();
            });

            $('#btnRejeitar').on('click', function() {
                abrirModalRejeicao();
            });
        }

        function abrirModalAprovacao() {
            $('#idSolicitacaoAprovacao').text(id);
            limparValidacaoModal($('#modalAprovacao'));
            $('#processoAprovacao').val('').trigger('change.select2');
            configurarCampoProcessoAprovacao();
            if (nivelUsuario == eNivel.DESENVOLVIMENTO) {
                carregarProcessosAprovacao();
            }

            const modal = new bootstrap.Modal(document.getElementById('modalAprovacao'));
            modal.show();
        }

        function abrirModalRejeicao() {
            carregarMotivosRejeicao();

            $('#idSolicitacaoRejeicao').text(id);
            limparValidacaoModal($('#modalRejeicao'));
            $('#motivoRejeicao').val('').trigger('change.select2');

            const modal = new bootstrap.Modal(document.getElementById('modalRejeicao'));
            modal.show();
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id');
            codPlanta = urlParams.get('codPlanta');

            anexarEventos();
            navegarParaAba(1);
            inicializarSelect2();

            if (!id || !codPlanta) {
                $('#titulo').text('Avaliação de Solicitação SPD');
                Toast.warning('Id ou Planta não informada');
                return;
            }

            
            $('#titulo').text(`Avaliação de Solicitação SPD #${id}`);
            await carregarDadosSolicitacao();
        });
    </script>
</head>
<body>
    <div class="barra-teste">
        <i class="bi bi-wrench-adjustable"></i>
        <span>Modo de teste — Visualizando como:</span>
        <span id="labelNivelAtivo">
            <% 
                If nivelUsuario = "D" Then
                    Response.Write "Desenvolvimento"
                ElseIf nivelUsuario = "C" Then
                    Response.Write "Comercial"
                ElseIf nivelUsuario = "P" Then
                    Response.Write "PCP"
                End If
            %>
        </span>
        <select onchange="alterarNivelTeste(this.value)">
            <option value="C" <%= Iif(nivelUsuario = "C", "selected", "") %>>Comercial</option>
            <option value="D" <%= Iif(nivelUsuario = "D", "selected", "") %>>Desenvolvimento</option>
            <option value="P" <%= Iif(nivelUsuario = "P", "selected", "") %>>PCP</option>
        </select>
    </div>

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
                        <button type="button" class="tab-button active" data-tab="1"><i class="tab-icon bi bi-check2-square"></i> <span class="tab-title">Avaliação</span> </button>

                        <div class="tab-group-title">Consulta</div>
                        <button type="button" class="tab-button active" data-tab="2"><i class="tab-icon bi bi-box"></i><span class="tab-title">Produto</span></button>
                        <button type="button" class="tab-button" data-tab="3"><i class="tab-icon bi bi-list-ul"></i><span class="tab-title">Especificações</span></button>
                        <button type="button" class="tab-button" data-tab="4"><i class="tab-icon bi bi-brush"></i><span class="tab-title">Acabamento</span></button>
                        <button type="button" class="tab-button" data-tab="5"><i class="tab-icon bi bi-puzzle"></i><span class="tab-title">Acessórios</span></button>
                        <button type="button" class="tab-button" data-tab="6"><i class="tab-icon bi bi-archive"></i><span class="tab-title">Embalagem</span></button>
                        <button type="button" class="tab-button" data-tab="7"><i class="tab-icon bi bi-clipboard-check"></i><span class="tab-title">Amostra</span></button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioAvaliacao" name="formularioAvaliacao" class="form-content" novalidate>
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Avaliação da solicitação</div>
                                <p class="section-subtitle">Registre a aprovação ou rejeição da solicitação enviada.</p>
                            </div>

                            <div class="readonly-alert mt-3">
                                <i class="bi bi-info-circle me-1"></i>
                                Consulte as demais abas para revisar os dados gerais antes de concluir.
                            </div>

                            <div class="row g-3 mt-4">
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
                        </div>

                        <!-- Aba Produto -->
                        <div class="tab-content active" data-tab-content="2">
                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="codPlanta" class="form-label">Planta <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="codPlanta" id="codPlanta">
                                        <option value="">Selecione a planta</option>
                                        <option value="00">Pirapetinga</option>
                                        <option value="04">Uberaba</option>
                                        <option value="05">Bento</option>
                                        <option value="06">Maranguape</option>
                                        <option value="07">Saquarema</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="nomeCliente" class="form-label">Nome do Cliente</label>
                                    <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" disabled>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="codCliente" class="form-label">Cód. Cliente</label>
                                    <input disabled type="text" class="form-control form-control-sm" id="codCliente" name="codCliente">
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="codLoja" class="form-label">Loja</label>
                                    <input type="text" class="form-control form-control-sm" id="codLoja" name="codLoja" disabled>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="tipoSolicitacao" class="form-label">Tipo de Solicitação <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="tipoSolicitacao" id="tipoSolicitacao">
                                        <option value="">Selecione o tipo de solicitação</option>
                                        <option value="1">Projetar/alterar exatamente conforme desenho ou amostra enviado ao Desenvolvimento</option>
                                        <option value="2">Projetar/alterar exatamente conforme especificações identificadas nesta Solicitação</option>
                                        <option value="3">Alterar ou Cancelar uma Ficha Técnica</option>
                                        <option value="6">Novo Projeto / Primário</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="idDesenvolvedor" class="form-label">Técnico/Desenvolvedor <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="idDesenvolvedor" name="idDesenvolvedor"><option value="">Selecione o técnico...</option></select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="referencia" class="form-label">Referência <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="referencia" name="referencia">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="necessitaAcompanhamento" class="form-label">Necessita Acompanhamento <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="necessitaAcompanhamento" name="necessitaAcompanhamento">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6" id="divNivel" style="display: none;">
                                    <label for="nivelLotePiloto" class="form-label">Nível <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="nivelLotePiloto" name="nivelLotePiloto">
                                        <option value="">Selecione...</option>
                                        <option value="1">Nível 1</option>
                                        <option value="2">Nível 2</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="toleranciaMaxima" class="form-label">Tolerância Máxima (%) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="toleranciaMaxima" name="toleranciaMaxima" min="0" max="100">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="toleranciaMinima" class="form-label">Tolerância Mínima (%) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="toleranciaMinima" name="toleranciaMinima" min="0" max="100">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <div class="d-flex align-items-center gap-2">
                                        <label for="estilo" class="form-label mb-0">Estilo <span class="required-field">*</span></label>
                                    </div>
                                    <select disabled class="form-select form-select-sm select2" id="estilo" name="estilo">
                                        <option value="">Selecione o estilo...</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="observacoes" class="form-label">Observações</label>
                                    <select disabled class="form-select form-select-sm select2" id="observacoes" name="observacoes">
                                        <option value="">Selecione...</option>
                                        <option value="F">Com Furos</option>
                                        <option value="C">Com Chanfros</option>
                                        <option value="E">Com Furos/Chanfros</option>
                                        <option value="D">Com Furo Alça</option>
                                        <option value="A">Caixa com Display</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="comprimento" class="form-label">Comprimento (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="comprimento" name="comprimento" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="largura" class="form-label">Largura (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="largura" name="largura" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="altura" class="form-label">Altura (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="altura" name="altura" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="numeroMatriz" class="form-label">Número da Matriz (Faca Rotativa)<span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="numeroMatriz" name="numeroMatriz" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="clienteUsaTara" class="form-label">Cliente usa Tara <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="clienteUsaTara" name="clienteUsaTara">
                                        <option value="">Selecione...</option>
                                        <option value="1">Sim</option>
                                        <option value="2">Não</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6" id="divToleranciaPeso" style="display: none;">
                                    <label for="toleranciaPeso" class="form-label">Tolerância de Peso ( ± %) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="toleranciaPeso" name="toleranciaPeso" min="0" max="100">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="estruturaPapelao" class="form-label">Estrutura de Papelão Ondulado <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="estruturaPapelao" name="estruturaPapelao">
                                        <option value="">Selecione a estrutura...</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="gramatura" class="form-label">Gramatura <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="gramatura" name="gramatura" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="larguraChapa" class="form-label">Largura da Chapa (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="larguraChapa" name="larguraChapa" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="comprimentoChapa" class="form-label">Comprimento da Chapa (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="comprimentoChapa" name="comprimentoChapa" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="areaBlank" class="form-label">Área Blank (m²) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="areaBlank" name="areaBlank" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="areaLiquida" class="form-label">Área Líquida (m²) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="areaLiquida" name="areaLiquida" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="peso" class="form-label">Peso (g) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="peso" name="peso" min="0">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="inkJet" class="form-label">Ink Jet <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="inkJet" name="inkJet">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                            </div>
                        </div>

                        <!-- Aba Especificações -->
                        <div class="tab-content" data-tab-content="3">
                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="gramaturaReferencial" class="form-label">Gramatura Referencial (g/m²) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="gramaturaReferencial" name="gramaturaReferencial" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="esmagamento" class="form-label">Esmagamento (kgf/cm²) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="esmagamento" name="esmagamento" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="ectColuna" class="form-label">ECT/Coluna (kgf/cm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="ectColuna" name="ectColuna" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="mullenReferencial" class="form-label">Mullen Referencial (kgf/cm²) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="mullenReferencial" name="mullenReferencial" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="espessura" class="form-label">Espessura (mm) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="espessura" name="espessura" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="bctCompressao" class="form-label">BCT/Compressão (kgf) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="bctCompressao" name="bctCompressao" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="cobbInterno" class="form-label">COBB Interno (g/m²H₂O) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="cobbInterno" name="cobbInterno" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="cobbExterno" class="form-label">COBB Externo (g/m²H₂O) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="cobbExterno" name="cobbExterno" min="0" step="1">
                                </div>
                            </div>

                            <div class="card bg-light p-2 mt-4">
                                <div class="d-flex gap-2 align-items-center">
                                    <code>1 kgf = 9,80665 N</code>
                                    <span disabled class="text-muted">Para converter de kgf para Newton, multiplique por 9,80665</span>
                                </div>
                                <div class="d-flex gap-2 align-items-center">
                                    <code>1 lbf/in² = 0,07031 kgf/cm²</code>
                                    <span disabled class="text-muted">Para converter de lbf/in² para kgf/cm², multiplique por 0,07031</span>
                                </div>
                                <div class="d-flex gap-2 align-items-center">
                                    <code>1 in = 25,4 mm</code>
                                    <span disabled class="text-muted">Para converter de polegadas para mm, multiplique por 25,4</span>
                                </div>
                            </div>
                        </div>

                        <!-- Aba Acabamento -->
                        <div class="tab-content" data-tab-content="4">
                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="tipoFechamento" class="form-label">Fechamento <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="tipoFechamento" name="tipoFechamento">
                                        <option value="">Selecione...</option>
                                        <option value="C">C - Colado</option>
                                        <option value="G">G - Grampeado</option>
                                        <option value="S">S - Sem Fechamento</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="ehAmarrado" class="form-label">Amarrado <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="ehAmarrado" name="ehAmarrado">
                                        <option value="">Selecione...</option>
                                        <option value="S">F - Fitilhado</option>
                                        <option value="N">S - Sem Amarrar</option>
                                        <option value="C">C - Caixa Master</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6" id="divQuantidadeFitilhos" style="display: none;">
                                    <label for="quantidadeFitilhos" class="form-label">Quantidade de Fitilhos <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="quantidadeFitilhos" name="quantidadeFitilhos">
                                        <option value="">Selecione...</option>
                                        <option value="1">1</option>
                                        <option value="2">2</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6" id="divTipoFitilho" style="display: none;">
                                    <label for="tipoFitilho" class="form-label">Fitilhos na Largura ou Comprimento <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="tipoFitilho" name="tipoFitilho">
                                        <option value="">Selecione...</option>
                                        <option value="L">L - Na largura</option>
                                        <option value="C">C - No comprimento</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="pecasPorUnidade" class="form-label">Pacotes de <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="pecasPorUnidade" name="pecasPorUnidade">
                                        <option value="">Selecione...</option>
                                        <option value="5">5</option>
                                        <option value="10">10</option>
                                        <option value="15">15</option>
                                        <option value="20">20</option>
                                        <option value="25">25</option>
                                        <option value="30">30</option>
                                        <option value="35">35</option>
                                        <option value="40">40</option>
                                        <option value="45">45</option>
                                        <option value="50">50</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="ehPaletizado" class="form-label">Paletizado <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="ehPaletizado" name="ehPaletizado">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>

                                <div id="camposPaletizado" style="display: none;" class="mt-3">
                                    <div class="row g-3">
                                        <div class="col-lg-4 col-md-6">
                                            <label for="palete" class="form-label">
                                                Palete <span class="required-field">*</span>
                                            </label>
                                            <select disabled class="form-select form-select-sm select2" id="palete">
                                                <option value="">Selecione o palete...</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="alturaMaximaPalete" class="form-label">Altura Máxima do Palete (mm) <span class="required-field">*</span></label>
                                            <input disabled type="number" class="form-control form-control-sm" id="alturaMaximaPalete" name="alturaMaximaPalete" min="0">
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="temFacesOpostas" class="form-label">Faces Opostas <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="temFacesOpostas" name="temFacesOpostas">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="temCantinho" class="form-label">Cantinho <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="temCantinho" name="temCantinho">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="temGradeMadeira" class="form-label">Grade de Madeira <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="temGradeMadeira" name="temGradeMadeira">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="temFilmeStrech" class="form-label">Filme Strech <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="temFilmeStrech" name="temFilmeStrech">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="clienteUsaRack" class="form-label">Cliente usa Rack <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="clienteUsaRack" name="clienteUsaRack">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="overhang" class="form-label">Cliente aceita caixa para fora do palete (overhang) <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="overhang" name="overhang">
                                                <option value="">Selecione...</option>
                                                <option value="S">Sim</option>
                                                <option value="N">Não</option>
                                            </select>
                                        </div>
                                        <div class="col-lg-4 col-md-6">
                                            <label for="quantidadeEspelhos" class="form-label">Espelho <span class="required-field">*</span></label>
                                            <select disabled class="form-select form-select-sm select2" id="quantidadeEspelhos" name="quantidadeEspelhos">
                                                <option value="">Selecione...</option>
                                                <option value="0">0</option>
                                                <option value="1">1</option>
                                                <option value="2">2</option>
                                                <option value="3">3</option>
                                                <option value="4">4</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Aba Acessórios -->
                        <div class="tab-content" data-tab-content="5">
                            <div class="d-flex justify-content-between align-items-center mb-3 pb-3 border-bottom">
                                <div class="d-flex align-items-center gap-2">
                                    <h5 class="mb-0 text-secondary">Lista de Acessórios</h5>
                                    <span class="badge bg-secondary" id="contadorAcessorios">0</span>
                                </div>
                            </div>

                            <div id="acessoriosListaContainer">
                            </div>
                        </div>

                        <!-- Aba Embalagem -->
                        <div class="tab-content" data-tab-content="6">
                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="conteudoEmbalado" class="form-label">Conteúdo a ser Embalado <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="conteudoEmbalado" name="conteudoEmbalado">
                                        <option value="">Selecione o conteúdo...</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="consumoMensal" class="form-label">Consumo Mensal (unidades) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="consumoMensal" name="consumoMensal" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="tipoLinhaProducaoCliente" class="form-label">Linha de Produção do Cliente <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="tipoLinhaProducaoCliente" name="tipoLinhaProducaoCliente">
                                        <option value="">Selecione...</option>
                                        <option value="M">Manual</option>
                                        <option value="A">Automática</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <div class="d-flex align-items-center gap-2">
                                        <label for="tipoProdutoPrimario" class="form-label mb-0">Tipo de Produto Primário <span class="required-field">*</span></label>
                                    </div>
                                    <select disabled class="form-select form-select-sm select2" id="tipoProdutoPrimario" name="tipoProdutoPrimario">
                                        <option value="">Selecione o tipo...</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="pesoPrimario" class="form-label">Peso do Primário (kg) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="pesoPrimario" name="pesoPrimario" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadePrimarios" class="form-label">Quantidade de Primários <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="quantidadePrimarios" name="quantidadePrimarios" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="suportacaoPrimario" class="form-label">Suportação do Primário <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="suportacaoPrimario" name="suportacaoPrimario">
                                        <option value="">Selecione...</option>
                                        <option value="2">Suportante</option>
                                        <option value="1">Semi Suportante</option>
                                        <option value="0">Não Suportante</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="tempoEstocagem" class="form-label">Tempo de Estocagem (meses) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="tempoEstocagem" name="tempoEstocagem" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="umidadeRelativa" class="form-label">Umidade Relativa (%) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="umidadeRelativa" name="umidadeRelativa" min="0" max="100" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeCores" class="form-label">Quantidade de Cores <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="quantidadeCores" name="quantidadeCores" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="precoKG" class="form-label">Preço KG (R$) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="precoKG" name="precoKG" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="orcamento" class="form-label">Orçamento</label>
                                    <select disabled class="form-select form-select-sm select2" id="orcamento" name="orcamento">
                                        <option value="">Selecione...</option>
                                        <option value="1">Sim</option>
                                        <option value="0">Não</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="margemContribuicao" class="form-label">Margem de Contribuição (%) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="margemContribuicao" name="margemContribuicao" min="0" max="100" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="resultado" class="form-label">Resultado (%) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="resultado" name="resultado" min="0" max="100" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="tipoSavings" class="form-label">Savings <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="tipoSavings" name="tipoSavings">
                                        <option value="">Selecione...</option>
                                        <option value="K">Saving SK</option>
                                        <option value="C">Saving Cliente</option>
                                        <option value="A">Ambos</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="valorSavingSK" class="form-label">Valor Saving SK (R$) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="valorSavingSK" name="valorSavingSK" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="valorSavingCliente" class="form-label">Valor Saving Cliente (R$) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="valorSavingCliente" name="valorSavingCliente" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="valorEstimadoCliche" class="form-label">Valor Estimado Clichê (R$) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="valorEstimadoCliche" name="valorEstimadoCliche" min="0" step="1">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="valorEstimadoMatriz" class="form-label">Valor Estimado Matriz (R$) <span class="required-field">*</span></label>
                                    <input disabled type="number" class="form-control form-control-sm" id="valorEstimadoMatriz" name="valorEstimadoMatriz" min="0" step="1">
                                </div>

                                <div class="col-12">
                                    <label for="informacoesAdicionais" class="form-label">Informações Adicionais</label>
                                    <textarea disabled class="form-control form-control-sm" id="informacoesAdicionais" name="informacoesAdicionais" rows="4"></textarea>
                                </div>
                            </div>
                        </div>

                        <!-- Aba Amostra -->
                        <div class="tab-content" data-tab-content="7">
                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeAmostras" class="form-label">Quantidade de Amostras <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="quantidadeAmostras" name="quantidadeAmostras">
                                        <option value="0">0 - Não necessita amostra</option>
                                        <option value="1">1</option>
                                        <option value="2">2</option>
                                    </select>
                                </div>
                            </div>

                            <div id="camposAmostra" style="display: none;" class="mt-3">
                                <div class="row g-3">
                                    <div class="col-lg-3 col-md-4">
                                        <label for="amostraCep" class="form-label">CEP <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraCep" 
                                            name="amostraCep" 
                                            placeholder="00000-000"
                                            data-mask="00000-000"
                                            maxlength="9">
                                    </div>

                                    <div class="col-lg-5 col-md-8">
                                        <label for="amostraResponsavel" class="form-label">Aos Cuidados <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraResponsavel" 
                                            name="amostraResponsavel"
                                            placeholder="Nome do responsável pelo recebimento">
                                    </div>

                                    <div class="col-lg-4 col-md-12">
                                        <label for="amostraTelefone" class="form-label">Telefone <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraTelefone" 
                                            name="amostraTelefone"
                                            placeholder="(00) 00000-0000"
                                            maxlength="15">
                                    </div>

                                    <div class="col-12">
                                        <label for="amostraEndereco" class="form-label">Endereço de Entrega <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraEndereco" 
                                            name="amostraEndereco"
                                            placeholder="Rua, Avenida, número, complemento...">
                                    </div>

                                    <div class="col-lg-4 col-md-6">
                                        <label for="amostraBairro" class="form-label">Bairro <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraBairro" 
                                            name="amostraBairro"
                                            placeholder="Bairro">
                                    </div>

                                    <div class="col-lg-4 col-md-6">
                                        <label for="amostraMunicipio" class="form-label">Município <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraMunicipio" 
                                            name="amostraMunicipio"
                                            placeholder="Município">
                                    </div>

                                    <div class="col-lg-4 col-md-4">
                                        <label for="amostraUF" class="form-label">UF <span class="required-field">*</span></label>
                                        <select disabled class="form-select form-select-sm select2" id="amostraUF" name="amostraUF">
                                            <option value="">UF</option>
                                            <option value="AC">AC</option>
                                            <option value="AL">AL</option>
                                            <option value="AP">AP</option>
                                            <option value="AM">AM</option>
                                            <option value="BA">BA</option>
                                            <option value="CE">CE</option>
                                            <option value="DF">DF</option>
                                            <option value="ES">ES</option>
                                            <option value="GO">GO</option>
                                            <option value="MA">MA</option>
                                            <option value="MT">MT</option>
                                            <option value="MS">MS</option>
                                            <option value="MG">MG</option>
                                            <option value="PA">PA</option>
                                            <option value="PB">PB</option>
                                            <option value="PR">PR</option>
                                            <option value="PE">PE</option>
                                            <option value="PI">PI</option>
                                            <option value="RJ">RJ</option>
                                            <option value="RN">RN</option>
                                            <option value="RS">RS</option>
                                            <option value="RO">RO</option>
                                            <option value="RR">RR</option>
                                            <option value="SC">SC</option>
                                            <option value="SP">SP</option>
                                            <option value="SE">SE</option>
                                            <option value="TO">TO</option>
                                        </select>
                                    </div>

                                    <div class="col-lg-6 col-md-8">
                                        <label for="amostraCNPJ" class="form-label">CNPJ <span class="required-field">*</span></label>
                                        <input disabled type="text" 
                                            class="form-control form-control-sm" 
                                            id="amostraCNPJ" 
                                            name="amostraCNPJ"
                                            placeholder="00.000.000/0000-00"
                                            maxlength="15">
                                    </div>

                                    <div class="col-lg-6 col-md-12">
                                        <label for="amostraEmail" class="form-label">E-mail <span class="required-field">*</span></label>
                                        <input disabled type="email" 
                                            class="form-control form-control-sm" 
                                            id="amostraEmail" 
                                            name="amostraEmail"
                                            placeholder="email@exemplo.com.br">
                                    </div>

                                    <div class="col-12">
                                        <label for="amostraObservacoes" class="form-label">Observações</label>
                                        <textarea disabled class="form-control form-control-sm" 
                                            id="amostraObservacoes" 
                                            name="amostraObservacoes" 
                                            rows="3"
                                            placeholder="Informações adicionais sobre a entrega da amostra..."></textarea>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de confirmação aprovação -->
    <div class="modal fade" id="modalAprovacao" aria-labelledby="modalAprovacaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <form id="formularioAprovacao" name="formularioAprovacao" novalidate>  
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalAprovacaoTitulo">
                            Confirmar aprovação #<span id="idSolicitacaoAprovacao"></span>
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                    </div>
                    <div class="modal-body">
                        <!-- Campo apenas para nivelUsuario == eNivel.DESENVOLVIMENTO -->
                        <div class="col-12" id="grupoProcessoAprovacao">
                            <label for="processoAprovacao" class="form-label">Processo <span class="required-field">*</span></label>
                            <select required class="form-select form-select-sm select2" id="processoAprovacao" name="processoAprovacao">
                                <option value="">Selecione o processo...</option>
                            </select>
                            <div class="invalid-feedback">Informe o processo de aprovação.</div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="button" onclick="avaliarSolicitacao(false)" class="btn btn-sm btn-success" id="btnEnviarAprovacao">
                            <i class="bi bi-check-lg me-1"></i>
                            <span>Enviar</span>
                        </button>
                    </div>
                </div>
            </form>
        </div>
    </div>

    <!-- Modal de confirmação rejeição -->
    <div class="modal fade" id="modalRejeicao" aria-labelledby="modalRejeicaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <form id="formularioRejeicao" name="formularioRejeicao" novalidate>  
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalRejeicaoTitulo">
                            Confirmar rejeição #<span id="idSolicitacaoRejeicao"></span>
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                    </div>
                    <div class="modal-body">
                        <div class="col-12">
                            <label for="motivoRejeicao" class="form-label">Motivo Rejeição <span class="required-field">*</span></label>
                            <select required class="form-select form-select-sm select2" id="motivoRejeicao" name="motivoRejeicao">
                                <option value="">Selecione o motivo...</option>
                            </select>
                            <div class="invalid-feedback">Informe o motivo da rejeição.</div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="button" onclick="avaliarSolicitacao(true)" class="btn btn-sm btn-success" id="btnEnviarRejeicao">
                            <i class="bi bi-check-lg me-1"></i>
                            <span>Enviar</span>
                        </button>
                    </div>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
