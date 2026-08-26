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

        .tab-group-title:first-child {
            margin-top: 0;
        }

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
        }

        .form-content-wrapper {
            background: #fff;
            min-width: 0;
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

        .texto-detalhe-cell {
            min-width: 260px;
            vertical-align: top;
        }

        .texto-detalhe-link {
            cursor: pointer;
            transition: color .15s ease;
        }

        .texto-detalhe-link:hover,
        .texto-detalhe-link:focus {
            color: #0d6efd;
            outline: none;
            text-decoration: underline;
        }

        .detalhe-conteudo {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            max-height: 60vh;
            overflow: auto;
            padding: 14px;
            white-space: pre-wrap;
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

        @media (max-width: 800px) {
            .container {
                max-width: 95%;
            }

            .tab-title {
                flex: initial;
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

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
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

                    $('#hidCodigoPlanta').val(response.dados.codigoPlanta);
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

        function enviarFormulario(flagEnviar) {         
            const anexos = (typeof uploadAnexos !== 'undefined' && uploadAnexos) ? uploadAnexos.getFiles() : [];
            if(anexos.length == 0) {
                Toast.error('Você não adicionou nenhum anexo');
                return;
            }
   
            exibirLoading(true);

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'POST',
                dataType: 'json',
                data: {
                    acao: 'anexarDocumentos',
                    codigo: codigo,
                    codigoPlanta: $('#hidCodigoPlanta').val(),
                    anexos: JSON.stringify(anexos)
                },
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success('Documentos anexados com sucesso!');

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

        function anexarEventos() {
            $('#btnEnviar').on('click', function() {
                enviarFormulario();
            });
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            codigo = urlParams.get('codigo');

            if (!codigo) {
                Toast.error('ID inválido');
                return;
            }

            anexarEventos();
            inicializarSelect2();
            carregarDadosCadastro();
            $('#titulo').text(`Anexar Documentos Cadastro Unificado #${codigo}`)
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
            <div class="form-content-wrapper">
                <form id="formularioAvaliacao" name="formularioAvaliacao" novalidate>
                    <div class="p-4">
                        <div class="section-header">
                            <div class="section-title">Anexos</div>
                            <p class="section-subtitle">Evidências e documentos relacionados ao cadastro.</p>
                        </div>
                        <div class="row g-3">                                
                            <file-upload id="uploadAnexos" file-code="UNIFICADO_" max-files="10" multiple></file-upload>
                            <input type="hidden" id="hidCodigoPlanta" name="hidCodigoPlanta">

                            <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                                <button type="button" class="btn btn-sm btn-success" id="btnEnviar">
                                    <i class="bi bi-check-lg me-1"></i> Salvar Anexos
                                </button>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
        </div>

    <div class="modal fade" id="modalDetalheTexto" tabindex="-1" aria-labelledby="modalDetalheTituloTexto" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalDetalheTituloTexto">
                        Detalhes
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="modalDetalheConteudo" class="detalhe-conteudo"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
