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
            max-width: 90%;
        }
        .form-label {
            font-weight: 600;
        }
        .tab-pane {
            display: block;
            max-height: 50vh;
            overflow: auto;
        }
        .required-field {
            color: #dc3545;
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
        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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
        .form-switch {
            display: flex;
            align-items: center;
        }
        .form-switch .form-check-input {
            width: 3em;
            height: 1.5em;
            cursor: pointer;
        }
        .form-switch .form-check-label {
            cursor: pointer;
            font-weight: 600;
            color: #495057;
            padding-left: 8px;
        }
        /* ---------- Campos para alteração ---------- */
        .nav-tabs .contador-grupo {
            font-size: .68rem;
            font-weight: 600;
        }
        .tab-sem-resultado {
            opacity: .45;
        }
        .barra-acoes {
            background-color: #fff;
            border-radius: 0 0 calc(.25rem - 1px) calc(.25rem - 1px);
            border-top: 1px solid #dee2e6;
            bottom: 0;
            box-shadow: 0 -6px 16px rgba(33, 37, 41, .05);
            margin: 1.5rem -1rem -1rem;
            padding: .75rem 1rem;

        }
        .campo-cabecalho,
        .campo-alteracao-linha {
            align-items: center;
            column-gap: .75rem;
            display: grid;
            grid-template-columns: 56px minmax(150px, 1.1fr) minmax(150px, 1.1fr) 24px minmax(170px, 1.1fr);
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
            transition: background-color .15s ease, border-color .15s ease;
        }
        .campo-alteracao-linha:last-child {
            border-bottom-color: transparent;
        }
        .campo-alteracao-linha:hover {
            background-color: #fafbfc;
        }
        .campo-alteracao-linha.selecionado {
            background-color: #f0f6ff;
        }
        .campo-alteracao-linha .form-check {
            margin: 0;
            min-height: 0;
        }
        .campo-alteracao-linha .form-check-input {
            cursor: pointer;
        }
        .campo-nome {
            color: #212529;
            font-weight: 600;
        }
        .valor-atual {
            color: #6c757d;
            white-space: normal;
            word-break: break-word;
        }
        .campo-seta {
            color: #ced4da;
            text-align: center;
            transition: color .15s ease;
        }
        .campo-alteracao-linha.selecionado .campo-seta {
            color: #0d6efd;
        }
        .campo-rotulo-mobile {
            color: #adb5bd;
            display: none;
            font-size: .68rem;
            font-weight: 700;
            letter-spacing: .05em;
            text-transform: uppercase;
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
        .estado-vazio {
            background-color: #f8f9fa;
            border: 1px dashed #dee2e6;
            border-radius: 8px;
            color: #6c757d;
            padding: 2.5rem 1rem;
            text-align: center;
        }
        /* ---------- Resumo das alterações ---------- */
        .resumo-alteracoes {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: .5rem 1rem;
        }
        .resumo-item {
            border-bottom: 1px solid #e9ecef;
            padding: .65rem 0;
        }
        .resumo-item:last-child {
            border-bottom: none;
        }
        .resumo-valores {
            align-items: baseline;
            display: flex;
            flex-wrap: wrap;
            font-size: .85rem;
            gap: .5rem;
            margin-top: .2rem;
        }
        .resumo-de {
            color: #6c757d;
            text-decoration: line-through;
        }
        .resumo-para {
            color: #146c43;
            font-weight: 600;
        }
        .badge-grupo {
            background-color: #e7f1ff;
            color: #0a58ca;
            font-weight: 600;
        }
        .campo-nome mark {
            background-color: #fff3cd;
            border-radius: 2px;
            padding: 0 1px;
        }
        .campo-alteracao-linha.linha-invalida {
            background-color: #fff5f5;
        }
        .campo-alteracao-linha .invalid-feedback {
            font-size: .75rem;
            margin-top: 2px;
        }
        .input-valor-novo:disabled {
            background-color: #f8f9fa;
            color: #adb5bd;
        }
        #clienteResultados tbody tr {
            cursor: pointer;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let validacoes = {};
        let acao = 'CRIAR';
        let clienteSelecionado = null;
        let resultadosCliente = [];
        let camposAlteracao = [];
        let grupoAtivo = null;

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function formularioSomenteVisualizacao() {
            return acao === 'VISUALIZAR';
        }

        function tipoSolicitacaoPossuiAlteracao() {
            const tipoSolicitacao = $('#tipoSolicitacao').val();

            return tipoSolicitacao === 'ALTERACAO_DADOS'
                || tipoSolicitacao === 'DESBLOQUEIO_COM_ALTERACAO_DADOS';
        }

        function atualizarInterfaceTipoSolicitacao() {
            const possuiAlteracao = tipoSolicitacaoPossuiAlteracao();

            $('#areaCamposAlteracao').toggleClass('d-none', !possuiAlteracao);

            if (!possuiAlteracao) {
                camposAlteracao = [];
                grupoAtivo = null;
                $('#filtroCampos').val('');
                $('#filtroCamposWrapper').addClass('d-none');
                $('#containerCamposAlteracao').empty();
            }

            atualizarBarraAcoes();
        }

        function configurarModoFormulario() {
            if (!formularioSomenteVisualizacao()) return;

            $('#btnSalvar').addClass('d-none');
            $('#formularioSolicitacao').find('input, select, textarea, button').prop('disabled', true);
            $('#btnVoltar, #filtroCampos, #btnLimparFiltro').prop('disabled', false);
            $('#tabsGrupos button').prop('disabled', false);
            $('#formularioSolicitacao').find('.select2').trigger('change.select2');
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function aplicarCamposCondicionais() {
            const tipoPessoa = $('#hidTipoPessoa').val();

            try {
                $('#cpfCnpj').unmask();
            } catch (e) {}

            if (tipoPessoa === 'J') {
                $('#cpfCnpj').mask('AA.AAA.AAA/AAAA-00', {
                    translation: {
                        'A': { pattern: /[A-Za-z0-9]/ }
                    },
                    onKeyPress: function(value, event, field) {
                        field.val(value.toUpperCase());
                    }
                });

                $('#cpfCnpj').off('input.apenasMaiusculo').on('input.apenasMaiusculo', function() {
                    this.value = this.value.toUpperCase();
                });

                $('#labelCpfCnpj').html('CNPJ <span class="required-field">*</span>');
            }

            if (tipoPessoa === 'F') {
                $('#cpfCnpj').mask('000.000.000-00');
                $('#labelCpfCnpj').html('CPF <span class="required-field">*</span>');
            }
        }

        function inicializarValidacoes() {
            const obrigatorio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            validacoes = new JustValidate('#formularioSolicitacao', { lockForm: false });

            validacoes
                .addField('#codigoPlanta', [obrigatorio('Planta é obrigatória')])
                .addField('#tipoSolicitacao', [obrigatorio('Tipo de solicitação é obrigatório')])
                .addField('#justificativa', [obrigatorio('Justificativa é obrigatória')])
                .addField('#codigoCliente', [obrigatorio('Código do cliente é obrigatório')])
                .addField('#codigoLoja', [obrigatorio('Loja do cliente é obrigatória')])
                .addField('#razaoSocial', [obrigatorio('Razão Social é obrigatória')])
                .addField('#cpfCnpj', [obrigatorio('CNPJ/CPF é obrigatório')]);
        }

        function navegarParaGrade() {
            window.location.href = 'SolicitacaoClienteAlteracaoGrade.asp';
        }

        function carregarDadosCadastro() {
            exibirLoading(true);

            return $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp',
                type: 'GET',
                data: {
                    acao: 'obterPorId',
                    id: id
                },
                dataType: 'json'
            }).then(function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }

                return preencherFormularioEdicao(response.dados);
            }).fail(function(xhr, status, error) {
                console.error('Erro ao carregar dados:', error);
                Toast.error('Erro ao carregar dados');
            }).always(function() {
                exibirLoading(false);
            });
        }

        function buscarCamposDisponiveis(codigoPlanta, codigoCliente, codigoLoja) {
            return $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
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
                // Na edição, preserva o valor atual capturado no momento da solicitação
                valorAtual: salvo ? (salvo.valorAtual || '') : (campo.valorAtual || ''),
                valorNovo: salvo ? (salvo.valorNovo || '') : '',
                selecionado: !!salvo
            };
        }

        function preencherFormularioEdicao(dados) {
            if (!dados) return;

            $('#btnBuscarCliente, #codigoPlanta, #tipoSolicitacao').prop('disabled', true);

            $('#hidTipoPessoa').val(dados.tipoPessoa || '');
            $('#codigoPlanta').val(dados.codigoPlanta || '').trigger('change.select2');
            $('#tipoSolicitacao').val(dados.tipoSolicitacao || dados.tipoSolicitacao || '').trigger('change.select2');
            
            atualizarInterfaceTipoSolicitacao();
            
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#justificativa').val(dados.justificativa || '');

            aplicarCamposCondicionais();

            const camposSalvos = dados.campos || [];

            if (!tipoSolicitacaoPossuiAlteracao()) {
                limparCamposAlteracao();
                atualizarInterfaceTipoSolicitacao();
                configurarModoFormulario();
                return Promise.resolve();
            }

            // O retorno de obterPorId traz apenas idCampo/valorAtual/valorNovo.
            // Busca o catálogo completo de campos (com nome, grupo, tipo, opções)
            // e mescla com os campos gravados na solicitação.
            return buscarCamposDisponiveis(dados.codigoPlanta, dados.codigoCliente, dados.codigoLoja)
                .then(function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao carregar campos.');
                        limparCamposAlteracao();
                        return;
                    }

                    camposAlteracao = (response.dados || []).map(function(campo) {
                        const salvo = camposSalvos.find(function(item) {
                            return String(item.idCampo) === String(campo.idCampo);
                        });

                        return mapearCampoAlteracao(campo, salvo);
                    });

                    renderizarCamposAlteracao();
                    configurarModoFormulario();
                })
                .fail(function(xhr, status, error) {
                    console.error('Erro ao carregar campos:', error);
                    Toast.error('Erro ao carregar campos disponíveis.');
                    limparCamposAlteracao();
                });
        }

        function anexarEventos() {
            $('#hidTipoPessoa').on('change', function() {
                $('#cpfCnpj').val('');
                aplicarCamposCondicionais();
            });

            $('#codigoPlanta').on('change', function() {
                limparClienteSelecionado();
                limparCamposAlteracao();
            });

            $('#tipoSolicitacao').on('change', function() {
                atualizarInterfaceTipoSolicitacao();

                if (tipoSolicitacaoPossuiAlteracao()
                        && $('#codigoPlanta').val()
                        && $('#codigoCliente').val()) {
                    carregarCamposAlteracaoCliente();
                }
            });

            $('#filtroCampos').on('input', function() {
                $('#btnLimparFiltro').toggleClass('d-none', $(this).val() === '');
                renderizarCamposAlteracao();
            });

            $('#filtroCampos').on('keydown', function(e) {
                if (e.key === 'Escape') {
                    $(this).val('').trigger('input');
                }
            });

            $('#btnLimparFiltro').on('click', function() {
                $('#filtroCampos').val('').trigger('input').trigger('focus');
            });

            $('#btnLimparSelecao').on('click', function() {
                limparSelecaoCampos();
            });

            $('#btnConfirmarEnvio').on('click', function() {
                const modal = bootstrap.Modal.getInstance(document.getElementById('modalConfirmacao'));

                if (modal) {
                    modal.hide();
                }

                enviarFormulario();
            });

            let buscaClienteTimeout = null;

            $('#clienteBusca').on('input', function() {
                clearTimeout(buscaClienteTimeout);
                buscaClienteTimeout = setTimeout(function() {
                    pesquisarCliente();
                }, 500);
            });

            $('#clienteBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaClienteTimeout);
                    pesquisarCliente();
                }
            });

            $('#btnSalvar').on('click', function() {
                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });

            $('#formularioSolicitacao').on('submit', function(e) {
                e.preventDefault();

                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });
        }

        async function validarFormulario() {
            if (formularioSomenteVisualizacao()) return;

            const resultados = await Promise.all([
                validacoes.revalidate()
            ]);

            if (typeof exibirErrosValidacao === 'function') {
                exibirErrosValidacao(validacoes);
            }

            if (!resultados.every(Boolean)) {
                return;
            }

            if (!tipoSolicitacaoPossuiAlteracao()) {
                enviarFormulario();
                return;
            }

            const camposSelecionados = obterCamposSelecionados();

            if (camposSelecionados.length === 0) {
                Toast.warning('Selecione ao menos um campo para alteração.');
                return;
            }

            camposAlteracao.forEach(function(campo) {
                campo.erro = '';
            });

            let possuiErro = false;

            camposSelecionados.forEach(function(campo) {
                if (normalizarValor(campo.valorNovo) === '') {
                    campo.erro = 'Informe o novo valor deste campo.';
                    possuiErro = true;
                } else if (normalizarValor(campo.valorAtual) === normalizarValor(campo.valorNovo)) {
                    campo.erro = 'O novo valor é igual ao valor atual.';
                    possuiErro = true;
                }
            });

            if (possuiErro) {
                // Remove o filtro para garantir que todos os campos com erro fiquem visíveis
                if (normalizarValor($('#filtroCampos').val()) !== '') {
                    $('#filtroCampos').val('');
                    $('#btnLimparFiltro').addClass('d-none');
                }

                renderizarCamposAlteracao();
                focarPrimeiroCampoComErro();
                Toast.warning('Revise os campos destacados antes de enviar.');
                return;
            }

            abrirModalConfirmacao(camposSelecionados);
        }

        function abrirModalConfirmacao(camposSelecionados) {
            $('#confirmacaoQuantidade').text(
                camposSelecionados.length === 1
                    ? '1 alteração'
                    : camposSelecionados.length + ' alterações'
            );
            $('#confirmacaoCliente').text($('#razaoSocial').val() || '');
            $('#resumoConfirmacao').html(gerarHtmlResumo(camposSelecionados));

            const modalEl = document.getElementById('modalConfirmacao');
            const modal = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl);
            modal.show();
        }

        function enviarFormulario() {
            if (formularioSomenteVisualizacao()) return;

            const possuiAlteracao = tipoSolicitacaoPossuiAlteracao();
            const camposSelecionados = possuiAlteracao ? obterCamposSelecionados() : [];

            const dadosFormulario = {
                codigoPlanta: $('#codigoPlanta').val(),
                tipoSolicitacao: $('#tipoSolicitacao').val(),
                tipoPessoa: $('#hidTipoPessoa').val(),
                codigoCliente: $('#codigoCliente').val(),
                codigoLoja: $('#codigoLoja').val(),
                razaoSocial: $('#razaoSocial').val(),
                cpfCnpj: $('#cpfCnpj').val(),
                justificativa: $('#justificativa').val().replaceAll("'", ''),
                detalhamento: possuiAlteracao
                    ? JSON.stringify(camposSelecionados.map(function(campo) {
                        return {
                            idCampo: campo.idCampo,
                            valorAtual: campo.valorAtual,
                            valorNovo: campo.valorNovo
                        };
                    }))
                    : ''
            };

            if (id) {
                dadosFormulario.id = id;
                dadosFormulario.acao = 'editar';
            } else {
                dadosFormulario.acao = 'criar';
            }

            exibirLoading(true);

            $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp',
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
                    console.error('Erro ao salvar formulário:', error);
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function abrirModalCliente() {
            const codigoPlanta = $('#codigoPlanta').val();

            if (!codigoPlanta) {
                Toast.warning('Informe a planta antes de buscar por clientes.');
                return;
            }

            clienteSelecionado = null;
            resultadosCliente = [];

            $('#clienteBusca').val('');
            $('#clienteResultados').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar clientes existentes.</p>
                </div>
            `);

            $('#btnConfirmarCliente').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalCliente'));
            modal.show();

            setTimeout(function() {
                $('#clienteBusca').focus();
            }, 500);
        }

        function pesquisarCliente() {
            const busca = $('#clienteBusca').val().trim();
            const codigoPlanta = $('#codigoPlanta').val();

            if (!busca) {
                $('#clienteResultados').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar clientes existentes.</p>
                    </div>
                `);
                return;
            }

            clienteSelecionado = null;
            $('#btnConfirmarCliente').prop('disabled', true);

            $('#clienteResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando clientes...</p>
                </div>
            `);

            $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarClientes',
                    codigoPlanta: codigoPlanta,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#clienteResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0">${escapeHtml(response.mensagem || 'Erro ao buscar clientes.')}</p>
                            </div>
                        `);
                        return;
                    }

                    resultadosCliente = response.dados || [];

                    if (resultadosCliente.length === 0) {
                        $('#clienteResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0">Nenhum cliente encontrado com os termos pesquisados.</p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosCliente();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar clientes:', error);

                    $('#clienteResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0">Nenhum cliente encontrado com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosCliente() {
            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Nome</th>
                                <th>Código</th>
                                <th>Loja</th>
                                <th>CNPJ</th>
                                <th>Município</th>
                                <th>UF</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosCliente.forEach(function(cliente, index) {
                const documento = cliente.tipoPessoa == 'F'
                    ? formatarCPF(cliente.cpfCnpj || '')
                    : formatarCNPJ(cliente.cpfCnpj || '');

                html += `
                    <tr data-index="${index}" onclick="selecionarCliente(${index})" ondblclick="confirmarCliente()">
                        <td class="text-center">
                            <button type="button" onclick="selecionarCliente(${index})"
                                class="btn-selecionar"
                                title="Selecionar este cliente">
                                <i class="bi bi-circle"></i>
                            </button>
                        </td>
                        <td>${escapeHtml(cliente.razaoSocial || '')}</td>
                        <td>${escapeHtml(cliente.codigoCliente || '')}</td>
                        <td>${escapeHtml(cliente.codigoLoja || '')}</td>
                        <td>${escapeHtml(documento)}</td>
                        <td>${escapeHtml(cliente.municipio || '')}</td>
                        <td>${escapeHtml(cliente.uf || '')}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#clienteResultados').html(html);
        }

        function selecionarCliente(index) {
            clienteSelecionado = resultadosCliente[index];

            $('#clienteResultados tbody tr').removeClass('table-primary');
            $('#clienteResultados tr[data-index="' + index + '"]').addClass('table-primary');

            $('#clienteResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');

            $('#clienteResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');

            $('#btnConfirmarCliente').prop('disabled', false);
        }

        function confirmarCliente() {
            if (!clienteSelecionado) {
                Toast.warning('Selecione um cliente');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalCliente'));
            modal.hide();

            preencherDadosCliente(clienteSelecionado);
            aplicarCamposCondicionais();

            if (tipoSolicitacaoPossuiAlteracao()) {
                carregarCamposAlteracaoCliente();
            } else {
                atualizarInterfaceTipoSolicitacao();
            }
        }

        function preencherDadosCliente(dados) {
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#hidTipoPessoa').val(dados.tipoPessoa || '');
        }

        function carregarCamposAlteracaoCliente() {
            if (!tipoSolicitacaoPossuiAlteracao()) {
                atualizarInterfaceTipoSolicitacao();
                return;
            }

            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!codigoPlanta || !codigoCliente) return;

            $('#containerCamposAlteracao').html(`
                <div class="text-center text-muted py-4">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                    <p class="mt-2 mb-0">Carregando campos disponíveis...</p>
                </div>
            `);

            buscarCamposDisponiveis(codigoPlanta, codigoCliente, codigoLoja)
                .then(function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao carregar campos.');
                        limparCamposAlteracao();
                        return;
                    }

                    camposAlteracao = (response.dados || []).map(function(campo) {
                        return mapearCampoAlteracao(campo, null);
                    });

                    renderizarCamposAlteracao();
                })
                .fail(function(xhr, status, error) {
                    console.error('Erro ao carregar campos:', error);
                    Toast.error('Erro ao carregar campos disponíveis.');
                    limparCamposAlteracao();
                });
        }

        function renderizarCamposAlteracao() {
            if (!camposAlteracao.length) {
                $('#filtroCamposWrapper').addClass('d-none');
                $('#containerCamposAlteracao').html(`
                    <div class="estado-vazio">
                        <i class="bi bi-inbox fs-3 d-block mb-2"></i>
                        Nenhum campo disponível para alteração nesta planta.
                    </div>
                `);

                atualizarBarraAcoes();
                return;
            }

            $('#filtroCamposWrapper').removeClass('d-none');

            const filtro = normalizarValor($('#filtroCampos').val()).toLowerCase();
            const grupos = ['GERAL', 'CONTATOS', 'FISCAL', 'FINANCEIRO', 'ENDERECO', 'LOGISTICA'];
            const gruposRenderizados = [];

            grupos.forEach(function(grupo) {
                const camposGrupo = camposAlteracao
                    .filter(function(campo) {
                        return campo.grupoCampo === grupo;
                    })
                    .sort(function(a, b) {
                        return a.ordem - b.ordem;
                    });

                if (camposGrupo.length === 0) return;

                const camposExibidos = camposGrupo.filter(function(campo) {
                    return !filtro || String(campo.nomeExibicao || '').toLowerCase().indexOf(filtro) !== -1;
                });

                gruposRenderizados.push({
                    grupo: grupo,
                    camposExibidos: camposExibidos
                });
            });

            if (!gruposRenderizados.length) {
                $('#containerCamposAlteracao').html(`
                    <div class="estado-vazio">
                        <i class="bi bi-inbox fs-3 d-block mb-2"></i>
                        Nenhum campo disponível para alteração nesta planta.
                    </div>
                `);

                atualizarBarraAcoes();
                return;
            }

            const grupoAtivoExiste = gruposRenderizados.some(function(item) {
                return item.grupo === grupoAtivo;
            });

            if (!grupoAtivoExiste) {
                grupoAtivo = gruposRenderizados[0].grupo;
            }

            let tabs = '';
            let panes = '';

            gruposRenderizados.forEach(function(item) {
                const ativo = item.grupo === grupoAtivo;
                const semResultado = filtro && item.camposExibidos.length === 0;

                tabs += `
                    <li class="nav-item" role="presentation">
                        <button class="nav-link ${ativo ? 'active' : ''} ${semResultado ? 'tab-sem-resultado' : ''}"
                                id="tab-${item.grupo}"
                                data-bs-toggle="tab"
                                data-bs-target="#pane-${item.grupo}"
                                data-grupo="${item.grupo}"
                                type="button"
                                role="tab">
                            ${item.grupo}
                            <span class="badge rounded-pill ms-1 contador-grupo" data-grupo="${item.grupo}"></span>
                        </button>
                    </li>
                `;

                let conteudo = '';

                if (item.camposExibidos.length === 0) {
                    conteudo = `
                        <div class="estado-vazio">
                            <i class="bi bi-search fs-3 d-block mb-2"></i>
                            Nenhum campo deste grupo corresponde ao filtro informado.
                        </div>
                    `;
                } else {
                    conteudo = `
                        <div class="campo-cabecalho">
                            <div>Alterar</div>
                            <div>Campo</div>
                            <div>Valor atual</div>
                            <div></div>
                            <div>Novo valor</div>
                        </div>
                    `;

                    item.camposExibidos.forEach(function(campo) {
                        conteudo += renderizarLinhaCampoAlteracao(campo);
                    });
                }

                panes += `
                    <div class="tab-pane fade ${ativo ? 'show active' : ''}" id="pane-${item.grupo}" role="tabpanel">
                        ${conteudo}
                    </div>
                `;
            });

            $('#containerCamposAlteracao').html(`
                <ul class="nav nav-tabs" id="tabsGrupos" role="tablist">${tabs}</ul>
                <div class="tab-content pt-2">${panes}</div>
            `);

            $('#tabsGrupos button[data-bs-toggle="tab"]').on('shown.bs.tab', function(e) {
                grupoAtivo = $(e.target).data('grupo');
            });

            anexarEventosCamposAlteracao();
            atualizarContadoresGrupos();
            atualizarBarraAcoes();

            if (formularioSomenteVisualizacao()) {
                configurarModoFormulario();
            }
        }

        function atualizarContadoresGrupos() {
            $('.contador-grupo').each(function() {
                const grupo = $(this).data('grupo');

                const camposGrupo = camposAlteracao.filter(function(campo) {
                    return campo.grupoCampo === grupo;
                });

                const selecionados = camposGrupo.filter(function(campo) {
                    return campo.selecionado;
                });

                $(this)
                    .text(selecionados.length + '/' + camposGrupo.length)
                    .attr('title', selecionados.length + ' de ' + camposGrupo.length + ' campos selecionados')
                    .toggleClass('bg-primary', selecionados.length > 0)
                    .toggleClass('bg-light text-secondary border', selecionados.length === 0);
            });

        }

        function renderizarLinhaCampoAlteracao(campo) {
            const selecionado = campo.selecionado ? 'checked' : '';
            const disabled = campo.selecionado ? '' : 'disabled';
            const classeSelecionado = campo.selecionado ? 'selecionado' : '';

            const textoValorAtual = obterTextoExibicaoValor(campo, campo.valorAtual);
            const tituloValorAtual = (campo.valorAtual && textoValorAtual !== campo.valorAtual)
                ? ` title="Código: ${escapeAttribute(campo.valorAtual)}"`
                : '';

            return `
                <div class="campo-alteracao-linha ${classeSelecionado} ${campo.erro ? 'linha-invalida' : ''}" data-id-campo="${escapeAttribute(campo.idCampo)}">
                    <div class="form-check form-switch">
                        <input
                            type="checkbox"
                            role="switch"
                            class="form-check-input check-alterar-campo"
                            data-id-campo="${escapeAttribute(campo.idCampo)}"
                            title="Alterar este campo"
                            aria-label="Alterar o campo ${escapeAttribute(campo.nomeExibicao)}"
                            ${selecionado}
                        />
                    </div>

                    <div class="campo-nome">${destacarFiltro(campo.nomeExibicao)}</div>

                    <div>
                        <span class="campo-rotulo-mobile">Valor atual</span>
                        <div class="valor-atual"${tituloValorAtual}>
                            ${campo.valorAtual ? escapeHtml(textoValorAtual) : '<span class="text-muted fst-italic">Sem valor</span>'}
                        </div>
                    </div>

                    <div class="campo-seta">
                        <i class="bi bi-arrow-right"></i>
                    </div>

                    <div>
                        <span class="campo-rotulo-mobile">Novo valor</span>
                        ${renderizarInputValorNovo(campo, disabled)}
                        ${campo.erro ? '<div class="invalid-feedback d-block mensagem-erro-campo">' + escapeHtml(campo.erro) + '</div>' : ''}
                    </div>
                </div>
            `;
        }

        function destacarFiltro(texto) {
            const nome = String(texto === null || texto === undefined ? '' : texto);
            const filtro = normalizarValor($('#filtroCampos').val());
            const indice = filtro ? nome.toLowerCase().indexOf(filtro.toLowerCase()) : -1;

            if (indice === -1) return escapeHtml(nome);

            return escapeHtml(nome.substring(0, indice))
                + '<mark>' + escapeHtml(nome.substring(indice, indice + filtro.length)) + '</mark>'
                + escapeHtml(nome.substring(indice + filtro.length));
        }

        function renderizarInputValorNovo(campo, disabled) {
            const value = escapeAttribute(campo.valorNovo || '');

            const attrs = `
                class="form-control form-control-sm input-valor-novo${campo.erro ? ' is-invalid' : ''}"
                data-id-campo="${escapeAttribute(campo.idCampo)}"
                aria-label="Novo valor para ${escapeAttribute(campo.nomeExibicao)}"
                ${disabled}
            `;

            if (campo.tipoDado === 'EMAIL') {
                return `<input type="email" value="${value}" ${attrs} />`;
            }

            if (campo.tipoDado === 'NUMERO') {
                return `<input type="number" value="${value}" ${attrs} />`;
            }

            if (campo.tipoDado === 'BOOLEANO') {
                return `
                    <select ${attrs}>
                        <option value="">Selecione</option>
                        <option value="1" ${campo.valorNovo === '1' ? 'selected' : ''}>Sim</option>
                        <option value="0" ${campo.valorNovo === '0' ? 'selected' : ''}>Não</option>
                    </select>
                `;
            }

            if (campo.tipoDado === 'LISTA') {
                let options = '<option value="">Selecione</option>';

                (campo.opcoes || []).forEach(function(opcao) {
                    const opcaoValor = obterValorOpcao(opcao);
                    const opcaoTexto = obterTextoOpcao(opcao);
                    const selected = String(campo.valorNovo) === String(opcaoValor) ? 'selected' : '';

                    options += `<option value="${escapeAttribute(opcaoValor)}" ${selected}>${escapeHtml(opcaoTexto)}</option>`;
                });

                return `<select ${attrs}>${options}</select>`;
            }

            return `<input type="text" value="${value}" ${attrs} />`;
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

        function encontrarCampoPorId(idCampo) {
            return camposAlteracao.find(function(item) {
                return String(item.idCampo) === String(idCampo);
            });
        }

        function anexarEventosCamposAlteracao() {
            $('.check-alterar-campo').off('change').on('change', function() {
                const campo = encontrarCampoPorId($(this).attr('data-id-campo'));

                if (!campo) return;

                const selecionado = $(this).is(':checked');
                const linha = $(this).closest('.campo-alteracao-linha');
                const input = linha.find('.input-valor-novo');

                campo.selecionado = selecionado;

                if (selecionado) {
                    linha.addClass('selecionado');
                    input.prop('disabled', false).trigger('focus');
                } else {
                    // Mantém o valor digitado para não perder o trabalho caso o usuário reative o campo.
                    // Campos desmarcados não são enviados (obterCamposSelecionados filtra por 'selecionado').
                    linha.removeClass('selecionado');
                    input.prop('disabled', true);
                    definirErroCampo(campo, linha, '');
                }

                atualizarContadoresGrupos();
                atualizarBarraAcoes();
            });

            $('.input-valor-novo').off('input change').on('input change', function() {
                const campo = encontrarCampoPorId($(this).attr('data-id-campo'));

                if (!campo) return;

                campo.valorNovo = $(this).val();
                definirErroCampo(campo, $(this).closest('.campo-alteracao-linha'), '');
                atualizarBarraAcoes();
            });
        }

        function gerarHtmlResumo(selecionados) {
            let html = '';

            selecionados.forEach(function(campo) {
                html += `
                    <div class="resumo-item">
                        <div class="d-flex align-items-center flex-wrap gap-2">
                            <strong>${escapeHtml(campo.nomeExibicao)}</strong>
                            <span class="badge badge-grupo">${escapeHtml(campo.grupoCampo)}</span>
                        </div>
                        <div class="resumo-valores">
                            <span class="resumo-de">${campo.valorAtual ? escapeHtml(obterTextoExibicaoValor(campo, campo.valorAtual)) : 'Sem valor'}</span>
                            <i class="bi bi-arrow-right text-muted"></i>
                            <span class="resumo-para">${campo.valorNovo ? escapeHtml(obterTextoExibicaoValor(campo, campo.valorNovo)) : '<span class="text-muted fst-italic fw-normal">Não informado</span>'}</span>
                        </div>
                    </div>
                `;
            });

            return html;
        }

        function atualizarBarraAcoes() {
            const selecionados = obterCamposSelecionados();
            const somenteVisualizacao = formularioSomenteVisualizacao();
            const possuiAlteracao = tipoSolicitacaoPossuiAlteracao();

            $('#btnSalvar').prop('disabled', somenteVisualizacao || (possuiAlteracao && selecionados.length === 0));
            $('#btnLimparSelecao').toggleClass('d-none', somenteVisualizacao || !possuiAlteracao || selecionados.length === 0);
            $('#contadorBarraAcoes').toggleClass('d-none', !possuiAlteracao);

            if (!possuiAlteracao) {
                return;
            }

            if (selecionados.length === 0) {
                $('#contadorBarraAcoes').text('Nenhum campo selecionado');
                return;
            }

            $('#contadorBarraAcoes').html(
                (selecionados.length === 1
                    ? '<strong>1</strong> alteração selecionada'
                    : '<strong>' + selecionados.length + '</strong> alterações selecionadas')

            );
        }

        function definirErroCampo(campo, linha, mensagem) {
            campo.erro = mensagem || '';

            if (!linha || !linha.length) return;

            linha.toggleClass('linha-invalida', !!campo.erro);
            linha.find('.input-valor-novo').toggleClass('is-invalid', !!campo.erro);
            linha.find('.mensagem-erro-campo').remove();

            if (campo.erro) {
                linha.find('.input-valor-novo').closest('div').append(
                    '<div class="invalid-feedback d-block mensagem-erro-campo">' + escapeHtml(campo.erro) + '</div>'
                );
            }
        }

        function limparSelecaoCampos() {
            camposAlteracao.forEach(function(campo) {
                campo.selecionado = false;
                campo.erro = '';
            });

            renderizarCamposAlteracao();
        }

        function focarPrimeiroCampoComErro() {
            const campoComErro = camposAlteracao.find(function(campo) {
                return campo.erro;
            });

            if (!campoComErro) return;

            const tab = document.getElementById('tab-' + campoComErro.grupoCampo);

            if (tab) {
                const instanciaTab = bootstrap.Tab.getInstance(tab) || new bootstrap.Tab(tab);
                instanciaTab.show();
                grupoAtivo = campoComErro.grupoCampo;
            }

            setTimeout(function() {
                const linha = $('.campo-alteracao-linha[data-id-campo="' + campoComErro.idCampo + '"]');

                if (!linha.length) return;

                linha[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
                linha.find('.input-valor-novo').trigger('focus');
            }, 200);
        }

        function obterCamposSelecionados() {
            return camposAlteracao.filter(function(campo) {
                return campo.selecionado;
            });
        }

        function limparClienteSelecionado() {
            clienteSelecionado = null;
            resultadosCliente = [];

            $('#hidTipoPessoa').val('');
            $('#codigoCliente').val('');
            $('#codigoLoja').val('');
            $('#razaoSocial').val('');
            $('#cpfCnpj').val('');
        }

        function limparCamposAlteracao() {
            camposAlteracao = [];
            grupoAtivo = null;

            $('#filtroCampos').val('');
            $('#filtroCamposWrapper').addClass('d-none');

            $('#containerCamposAlteracao').html(`
                <div class="estado-vazio">
                    <i class="bi bi-person-lines-fill fs-3 d-block mb-2"></i>
                    Selecione a planta e busque um cliente para carregar os campos disponíveis.
                </div>
            `);

            atualizarBarraAcoes();
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

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);

            id = urlParams.get('id');
            acao = urlParams.get('acao') || 'CRIAR';

            anexarEventos();
            inicializarSelect2();

            if (id) {
                try {
                    await carregarDadosCadastro();
                } catch (e) {
                    console.error('Falha ao carregar a solicitação:', e);
                }

                if (formularioSomenteVisualizacao()) {
                    $('#titulo').text(`Visualizar Solicitação de Alteração de Cliente #${id}`);
                } else {
                    $('#titulo').text(`Editar Solicitação de Alteração de Cliente #${id}`);
                }
            } else {
                $('#titulo').text('Solicitação de Alteração de Cliente');
                limparCamposAlteracao();
            }

            inicializarValidacoes();
            aplicarCamposCondicionais();
            atualizarInterfaceTipoSolicitacao();
            configurarModoFormulario();
        });
    </script>
</head>

<body>
    <div class="container">
        <div class="d-flex my-4 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" id="btnVoltar" onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5">
                    <i class="bi bi-arrow-left"></i>
                </button>
                <h3 id="titulo" class="mb-0"></h3>
            </div>
        </div>

        <div id="loadingOverlay"
             class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center"
             style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="card p-3">
            <form id="formularioSolicitacao" name="formularioSolicitacao" novalidate>
                <input type="hidden" name="hidTipoPessoa" id="hidTipoPessoa" value="" />

                <div class="row g-3">
                    <div class="col-12 col-md-3">
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

                    <div class="col-12 col-md-3">
                        <label for="tipoSolicitacao" class="form-label">Tipo de Solicitação <span class="required-field">*</span></label>
                        <select required class="form-select form-select-sm select2" id="tipoSolicitacao" name="tipoSolicitacao" data-placeholder="Todos">
                            <option value="">Selecione um item</option>
                            <option value="DESBLOQUEIO">Desbloqueio Simples</option>
                            <option value="DESBLOQUEIO_COM_ALTERACAO_DADOS">Desbloqueio c/ Alteração de Dados</option>
                            <option value="BLOQUEIO">Bloqueio Simples</option>
                            <option value="ALTERACAO_DADOS">Alteração de Dados</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-3">
                        <label for="codigoCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                        <div class="input-group">
                            <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarCliente" onclick="abrirModalCliente()">
                                Buscar
                                <i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input disabled type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                        </div>
                    </div>

                    <div class="col-12 col-md-3">
                        <label for="codigoLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                        <input disabled type="text" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja">
                    </div>

                    <div class="col-12 col-md-8">
                        <label for="razaoSocial" class="form-label">Razão Social <span class="required-field">*</span></label>
                        <input disabled type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" maxlength="160" />
                    </div>

                    <div class="col-12 col-md-4">
                        <label for="cpfCnpj" class="form-label" id="labelCpfCnpj">CNPJ/CPF <span class="required-field">*</span></label>
                        <input disabled type="text" class="form-control form-control-sm" name="cpfCnpj" id="cpfCnpj" />
                    </div>
                    <div class="col-12">
                        <label for="justificativa" class="form-label">Justificativa <span class="required-field">*</span></label>
                        <textarea class="form-control form-control-sm" name="justificativa" id="justificativa" rows="3" maxlength="1000" placeholder="Informe a justificativa da solicitação"></textarea>
                    </div>
                </div>

                <div id="areaCamposAlteracao" class="d-none">
                    <hr class="my-4" />

                    <div class="d-flex justify-content-between align-items-end flex-wrap gap-3 mb-3">
                    <div class="input-group input-group-sm d-none" id="filtroCamposWrapper" style="max-width: 280px;">
                        <span class="input-group-text bg-white"><i class="bi bi-search"></i></span>
                        <input type="text"
                               class="form-control"
                               id="filtroCampos"
                               placeholder="Filtrar campos..."
                               autocomplete="off" />
                        <button type="button"
                                class="btn btn-outline-secondary d-none"
                                id="btnLimparFiltro"
                                title="Limpar filtro (Esc)">
                            <i class="bi bi-x-lg"></i>
                        </button>
                    </div>
                </div>

                    <div id="containerCamposAlteracao"></div>
                </div>

                <div class="barra-acoes">
                    <div class="d-flex justify-content-between align-items-center flex-wrap gap-2">
                        <span class="text-muted" id="contadorBarraAcoes">Nenhum campo selecionado</span>

                        <div class="d-flex align-items-center gap-2 ms-auto">
                            <button type="button" class="btn btn-sm btn-outline-secondary d-none" id="btnLimparSelecao">
                                <i class="bi bi-eraser me-1"></i>
                                Limpar seleção
                            </button>
                            <button type="button" class="btn btn-sm btn-success" id="btnSalvar" disabled>
                                <i class="bi bi-check-lg me-1"></i>
                                Enviar Formulário
                            </button>
                        </div>
                    </div>
                </div>
            </form>
        </div>
    </div>

    <div class="modal fade" id="modalCliente" aria-labelledby="modalClienteLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalClienteLabel">
                        <i class="bi bi-search me-2"></i>Buscar Cliente
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body">
                    <div class="mb-3">
                        <label for="clienteBusca" class="form-label">Pesquisar Cliente</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text"
                                   class="form-control form-control-sm"
                                   id="clienteBusca"
                                   placeholder="Digite o código, nome ou CNPJ para buscar..."
                                   autocomplete="off">
                        </div>
                    </div>
                    <div id="clienteResultados" style="max-height: 400px; overflow-y: auto;"></div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarCliente" onclick="confirmarCliente()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Selecionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="modalConfirmacao" aria-labelledby="modalConfirmacaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalConfirmacaoLabel">
                        <i class="bi bi-send-check me-2"></i>Confirmar solicitação
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body">
                    <p class="mb-3">
                        Você está solicitando <strong id="confirmacaoQuantidade"></strong> para o cliente
                        <strong id="confirmacaoCliente"></strong>. Revise os valores abaixo antes de enviar.
                    </p>
                    <div class="resumo-alteracoes" id="resumoConfirmacao"></div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Voltar e revisar
                    </button>
                    <button type="button" class="btn btn-sm btn-success" id="btnConfirmarEnvio">
                        <i class="bi bi-check-lg me-1"></i>Confirmar e enviar
                    </button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>