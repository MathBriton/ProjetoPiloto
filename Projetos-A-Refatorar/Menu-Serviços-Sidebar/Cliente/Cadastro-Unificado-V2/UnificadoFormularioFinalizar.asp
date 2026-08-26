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

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }

        .container {
            max-width: 85%;
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

        .form-layout {
            display: block;
        }

        .form-content-wrapper {
            background: #fff;
            min-width: 0;
        }

        .form-content {
            padding: 30px;
        }

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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

            .form-content {
                padding: 20px;
            }
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>


    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let codigo = 0;
        let validacoes = {}
        
        function inicializarValidacoes() {
            const obrigatorio = (msg) => ({
                validator: (value, fields) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });
            
            validacoes = new JustValidate('#formularioAvaliacao', { lockForm: false });
            validacoes
                .addField('#codigoSap', [obrigatorio('Código SAP é obrigatório')])
                .addField('#codigoProtheus', [obrigatorio('Código Protheus é obrigatório')])
                .addField('#codigoLoja', [obrigatorio('Código Loja é obrigatório')]);
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
            $('#btnEnviar').on('click', function() {
                enviarFormulario(true);
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

        function preencherDados(dados) {
            if (!dados) return;

            $('#codigoPlanta').val(dados.codigoPlanta);
            $('#codigoSap').val(dados.codigoSap);
            $('#codigoProtheus').val(dados.codigoProtheus);
            $('#codigoLoja').val(dados.codigoLoja);

            // Campos escondidos
            $('#limiteCredito').val(dados.limiteCredito);
            $('#volumeVendaMensalM2').val(dados.volumeVendaMensalM2);
            $('#precoVendaNetM2').val(dados.precoVendaNetM2);
            $('#precoVendaNetKg').val(dados.precoVendaNetKg);
            $('#razaoSocial').val(dados.razaoSocial);
            $('#cpfCnpj').val(dados.tipoPessoa == 'F' ? dados.cpf : dados.cnpj);
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

        async function enviarFormulario(flagEnviar) {            
            if(!await validarFormulario()) {
                return;
            }

            exibirLoading(true);

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'POST',
                dataType: 'json',
                data: {
                    acao: 'finalizar',
                    codigo: codigo,
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoSap: $('#codigoSap').val(),
                    codigoProtheus: $('#codigoProtheus').val(),
                    codigoLoja: $('#codigoLoja').val(),
                    observacao: ($('#observacao').val() || '').replaceAll("'", '"'),
                    limiteCredito: $('#limiteCredito').val() || 0,
                    volumeVendaMensalM2: $('#volumeVendaMensalM2').val(),
                    precoVendaNetM2: $('#precoVendaNetM2').val() || 0,
                    precoVendaNetKg: $('#precoVendaNetKg').val() || 0,
                    razaoSocial: $('#razaoSocial').val() || 0,
                    cpfCnpj: $('#cpfCnpj').val() || 0
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

        $(document).ready(function() {
            const urlParams = new URLSearchParams(window.location.search);
            codigo = urlParams.get('codigo');

            if (!codigo) {
                Toast.error('ID inválido');
                return;
            }

            inicializarValidacoes();
            anexarEventos();

            $('#titulo').text(`Finalizar Cadastro Unificado #${codigo}`);
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
            <div class="form-layout">
                <div class="form-content-wrapper">
                    <form id="formularioAvaliacao" name="formularioAvaliacao" novalidate>
                        <input type="hidden" name="codigoPlanta" id="codigoPlanta" />
                        <input type="hidden" name="limiteCredito" id="limiteCredito" />
                        <input type="hidden" name="volumeVendaMensalM2" id="volumeVendaMensalM2" />
                        <input type="hidden" name="precoVendaNetM2" id="precoVendaNetM2" />
                        <input type="hidden" name="precoVendaNetKg" id="precoVendaNetKg" />
                        <input type="hidden" name="razaoSocial" id="razaoSocial" />
                        <input type="hidden" name="cpfCnpj" id="cpfCnpj" />

                        <div class="form-content">
                            <div class="row g-3"> 
                                <div class="col-12 col-md-4">
                                    <label for="codigoSap" class="form-label">Código SAP <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="codigoSap" id="codigoSap" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoProtheus" class="form-label">Código Protheus <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="codigoProtheus" id="codigoProtheus" />
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="codigoLoja" class="form-label">Loja <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="codigoLoja" id="codigoLoja" />
                                </div>
                                <div class="col-12">
                                    <label for="observacao" class="form-label">
                                        Observações
                                    </label>
                                    <textarea class="form-control form-control-sm" id="observacao" rows="5"></textarea>
                                </div>
                                <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                                    <button type="button" class="btn btn-sm btn-success" id="btnEnviar">
                                        <i class="bi bi-check-lg me-1"></i> Finalizar
                                    </button>
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
