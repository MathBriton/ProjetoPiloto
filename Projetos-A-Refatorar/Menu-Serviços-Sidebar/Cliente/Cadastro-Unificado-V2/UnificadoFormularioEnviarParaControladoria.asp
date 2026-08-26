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
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }

        .container {
            max-width: 80%;
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

        .form-content {
            animation: fadeIn .25s ease;
            padding: 30px;
            height: 80vh;
            overflow: auto;
        }

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
        }

        .table thead {
            vertical-align: middle !important;
            position: sticky;
            inset-block-start: 0;
            background: #f8f9fa;
            z-index: 1;
        }

        .table td,
        .table th {
            padding: 8px 12px !important;
        }

        .table th {
            color: #495057;
            font-weight: 600;
        }

        .truncate-text {
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            max-width: 360px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: normal;
            word-break: break-word;
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
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-maskmoney/3.0.2/jquery.maskMoney.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let codigo = 0;
        let validacoes = {}

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function inicializarValidacoes() {
            const icmsDiferidoSim = () => $('#icmsDiferido').val() == '1';

            const obrigatorio = (msg) => ({
                validator: (value, fields) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const obrigatorioQuando = (condicao, msg) => ({
                validator: (value, fields) => !condicao() || (value !== null && value !== undefined && String(value).trim() !== ''),
                errorMessage: msg
            });

            const somenteNumerosMaximoDigitos = (maximoDigitos, msg) => ({
                validator: (value, fields) => {
                    const valor = String(value).trim();
                    return /^\d+$/.test(valor) && valor.length <= maximoDigitos;
                },
                errorMessage: msg
            });
            
            validacoes = new JustValidate('#formularioAvaliacao', { lockForm: false });
            validacoes
                .addField('#declarantePisCofins', [obrigatorio('Declarante de PIS/COFINS é obrigatório')])
                .addField('#contribuinteIcms', [obrigatorio('Contribuinte de ICMS é obrigatório')])
                .addField('#icmsDiferido', [obrigatorio('ICMS diferido é obrigatório')])
                .addField('#percentualDiferimento', [obrigatorioQuando(icmsDiferidoSim, 'Percentual de diferimento é obrigatório')])
                .addField('#isencaoIpi', [obrigatorio('Suspensão/Isenção do IPI é obrigatória')])
                .addField('#tipoCliente', [obrigatorio('Tipo de Cliente é obrigatório')])
                .addField('#tes', [somenteNumerosMaximoDigitos(3, 'TES deve conter apenas números e no máximo 3 dígitos')]);
        }

        function navegarParaGrade() {
            window.location.href = 'UnificadoGrade.asp';
        }

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function anexarEventos() {
            $('#btnSalvar').on('click', function() {
                enviarFormulario();
            });

            $('#icmsDiferido').on('change', function() {
                aplicarCamposCondicionais();
            });
        }

        function carregarDadosCadastro() {
            exibirLoading(true);

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'obterPorCodigo',
                    codigo: codigo
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    preencherDados(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados da FNC');
                },
                complete: function() {
                    exibirLoading(false);
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

        function preencherDados(dados) {
            if (!dados) return;
            $('#declarantePisCofins').val(dados.declarantePisCofins).trigger('change.select2');
            $('#contribuinteIcms').val(dados.contribuinteIcms).trigger('change.select2');
            $('#icmsDiferido').val(dados.icmsDiferido).trigger('change.select2');
            $('#percentualDiferimento').val(dados.percentualDiferimento);
            $('#isencaoIpi').val(dados.isencaoIpi).trigger('change.select2');
            $('#tipoCliente').val(dados.tipoCliente).trigger('change.select2');
            $('#tes').val(dados.tes);

            aplicarCamposCondicionais();
        }
        
        async function validarFormulario() {
            const isValid = await validacoes.revalidate()

            if (!isValid) {
                const errorsParaExibir = {};
                Object.entries(validacoes.fields).forEach(([index, campo]) => {
                    if(campo.isValid) return;
                    const primeiroErro = campo.rules.find(rule => !rule.isValid);
                    if (primeiroErro) errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                });

                validacoes.showErrors(errorsParaExibir);
                return false;
            }

            return true;
        }

        async function enviarFormulario() {            
            if(!await validarFormulario()) {
                return;
            }

            exibirLoading(true);

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'POST',
                dataType: 'json',
                data: {
                    codigo: codigo,
                    acao: 'enviarParaControladoria',
                    declarantePisCofins: $('#declarantePisCofins').val(),
                    contribuinteIcms: $('#contribuinteIcms').val(),
                    icmsDiferido: $('#icmsDiferido').val(),
                    percentualDiferimento: $('#icmsDiferido').val() == '1' ? $('#percentualDiferimento').val() : 0,
                    isencaoIpi: $('#isencaoIpi').val(),
                    tipoCliente: $('#tipoCliente').val(),
                    tes: $('#tes').val(),
                    observacao: ($('#observacao').val() || '').replaceAll("'", '"'),
                },
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success('Avaliação registrada com sucesso');

                    setTimeout(function() {
                        navegarParaGrade();
                    }, 1000);
                },
                error: function(xhr, status, error) {
                    console.error('Erro:', error);
                    Toast.error('Erro ao processar solicitação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function aplicarCamposCondicionais() {
            const icmsDiferido = $('#icmsDiferido').val() == '1';
            $('.campo-diferimento').toggle(icmsDiferido);
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            codigo = urlParams.get('codigo');

            if (!codigo) {
                Toast.error('ID inválido');
                return;
            }

            inicializarSelect2();
            inicializarValidacoes();
            anexarEventos();

            await Promise.all([
                carregarTiposCliente(),
                carregarTiposDeclarante(),
                carregarTiposContribuinte()
            ]);


            $('#titulo').text(`Enviar para Controladoria #${codigo}`)
            carregarDadosCadastro();
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
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="form-container">
            <form id="formularioAvaliacao" name="formularioAvaliacao" class="form-content" novalidate>
                <div class="row g-3">
                    <div class="col-md-4">
                        <label for="declarantePisCofins" class="form-label">Declarante de PIS/COFINS <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="declarantePisCofins" id="declarantePisCofins">
                            <option value="">Selecione um item</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="contribuinteIcms" class="form-label">Contribuinte de ICMS <span class="required-field">*</span></label>
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
                        <label for="percentualDiferimento" class="form-label">% Diferimento <span class="required-field">*</span></label>
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
                        <label for="tipoCliente" class="form-label">Tipo <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="tipoCliente" id="tipoCliente">
                            <option value="">Selecione um item</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="tes" class="form-label">TES <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" name="tes" id="tes" />
                    </div>
                    <div class="col-12">
                        <label for="observacao" class="form-label">
                            Observações
                        </label>
                        <textarea class="form-control form-control-sm" id="observacao" rows="5"></textarea>
                    </div>
                </div>

                <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                    <button type="button" class="btn btn-sm btn-success" id="btnSalvar">
                        <i class="bi bi-check-lg me-1"></i>Enviar para Controladoria
                    </button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
