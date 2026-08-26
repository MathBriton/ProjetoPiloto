<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-datepicker@1.10.0/dist/css/bootstrap-datepicker3.min.css">

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
            padding-bottom: 20px;
        }
        .form-label {
            font-weight: 600;
        }
        .required-field {
            color: #dc3545;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap-datepicker@1.10.0/dist/js/bootstrap-datepicker.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap-datepicker@1.10.0/dist/locales/bootstrap-datepicker.pt-BR.min.js"></script>

    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>

    <script>
        let id = null;
        let codPlanta = ''

        function obterSolicitacaoProjeto() {
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

        function preencherFormulario(dados) {
           $('#testeColuna').val(dados.testeColuna);
           $('#testeGramatura').val(dados.testeGramatura);
           $('#testeCobInterno').val(dados.testeCobInterno);
           $('#testeCobExterno').val(dados.testeCobExterno);
           $('#testeDescolamentoCapaInterna').val(dados.testeDescolamentoCapaInterna);
           $('#testeDescolamentoMiolo1').val(dados.testeDescolamentoMiolo1);
           $('#testeDescolamentoMiolo2').val(dados.testeDescolamentoMiolo2);
           $('#testeDescolamentoMiolo3').val(dados.testeDescolamentoMiolo3);
           $('#testeDescolamentoCapaExterna').val(dados.testeDescolamentoCapaExterna);
           $('#testeBctCompressao1').val(dados.testeBctCompressao1);
           $('#testeBctCompressao2').val(dados.testeBctCompressao2);
           $('#testeBctCompressao3').val(dados.testeBctCompressao3);
           $('#testePinInterno').val(dados.testePinInterno);
           $('#testePinExterno').val(dados.testePinExterno);
           $('#testeForcaVinco1').val(dados.testeForcaVinco1);
           $('#testeForcaVinco2').val(dados.testeForcaVinco2);
           $('#testeForcaVinco3').val(dados.testeForcaVinco3);
        }
    
        function navegarParaGrade() {
            window.location.href = `SolicitacaoProjetoGrade.asp?codPlanta=${codPlanta}`;
        }

        function salvarFormulario() {
            const $btnSalvar = $('#btnSalvar');
            const $loadingIcon = $btnSalvar.find('#loadingIcon');
            const $checkIcon = $btnSalvar.find('#checkIcon');

            $btnSalvar.prop('disabled', true);
            $checkIcon.css('display', 'none');
            $loadingIcon.css('display', 'inline-block');

            let dadosFormulario = {
                acao: 'editarTestesFisicos',
                id: id,
                codPlanta: codPlanta,
                testeColuna: $('#testeColuna').val() || 0,
                testeGramatura: $('#testeGramatura').val() || 0,
                testeCobInterno: $('#testeCobInterno').val() || 0,
                testeCobExterno: $('#testeCobExterno').val() || 0,
                testeDescolamentoCapaInterna: $('#testeDescolamentoCapaInterna').val() || 0,
                testeDescolamentoMiolo1: $('#testeDescolamentoMiolo1').val() || 0,
                testeDescolamentoMiolo2: $('#testeDescolamentoMiolo2').val() || 0,
                testeDescolamentoMiolo3: $('#testeDescolamentoMiolo3').val() || 0,
                testeDescolamentoCapaExterna: $('#testeDescolamentoCapaExterna').val() || 0,
                testeBctCompressao1: $('#testeBctCompressao1').val() || 0,
                testeBctCompressao2: $('#testeBctCompressao2').val() || 0,
                testeBctCompressao3: $('#testeBctCompressao3').val() || 0,
                testePinInterno: $('#testePinInterno').val() || 0,
                testePinExterno: $('#testePinExterno').val() || 0,
                testeForcaVinco1: $('#testeForcaVinco1').val() || 0,
                testeForcaVinco2: $('#testeForcaVinco2').val() || 0,
                testeForcaVinco3: $('#testeForcaVinco3').val() || 0,
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        $btnSalvar.prop('disabled', false);
                        $loadingIcon.css('display', 'none');
                        $checkIcon.css('display', 'inline-block');
                        return;
                    }
                    
                    Toast.success(response.mensagem);
                    navegarParaGrade()
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro:', error);
                    Toast.error('Erro ao salvar formulário');
                    $btnSalvar.prop('disabled', false);
                    $loadingIcon.css('display', 'none');
                    $checkIcon.css('display', 'inline-block');
                }
            });
        }

        function cancelarFormulario() {
            ConfirmationModal.show({
                message: 'Deseja realmente cancelar? Os dados não salvos serão perdidos.',
                type: 'info',
                confirmText: 'Confirmar',
                cancelText: 'Continuar editando',
                onConfirm: () => navegarParaGrade(),
                onCancel: () => {}
            });
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: "z-index-9999",
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        function anexarEventos() {
            $('#formularioSolicitacao').on('submit', function(e) {
                e.preventDefault();
                salvarFormulario();
            });
        }

        $(document).ready(function() {
            inicializarSelect2();
            anexarEventos();

            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');
            const paramCodPlanta = urlParams.get('codPlanta');
            
            if (paramId && paramCodPlanta) {
                id = paramId;
                codPlanta = paramCodPlanta;

                $('#idTitulo').text(id);
                obterSolicitacaoProjeto();
            }
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                <i class="bi bi-arrow-left"></i> 
            </button>
            <h3 class="mb-0">Alteração de Testes Físicos de Projeto (SPD) #<span id="idTitulo"></span></h3>
        </div>

        <div class="card p-3">
            <form id="formularioSolicitacao" name="formularioSolicitacao">
                <div class="row g-3">
                    <div class="col-lg-4 col-md-6">
                        <label for="testeColuna" class="form-label">ECT/Coluna (kgf/cm) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeColuna" name="testeColuna" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeGramatura" class="form-label">Gramatura(g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeGramatura" name="testeGramatura" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeCobInterno" class="form-label">Cob Interno(g/m²H2O) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeCobInterno" name="testeCobInterno" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeCobExterno" class="form-label">Cob Externo (g/m²H2O) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeCobExterno" name="testeCobExterno" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeDescolamentoCapaInterna" class="form-label">Descolamento Capa Interna (g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeDescolamentoCapaInterna" name="testeDescolamentoCapaInterna" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeDescolamentoMiolo1" class="form-label">Descolamento Miolo 1 (g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeDescolamentoMiolo1" name="testeDescolamentoMiolo1" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeDescolamentoMiolo2" class="form-label">Descolamento Miolo 2 (g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeDescolamentoMiolo2" name="testeDescolamentoMiolo2" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeDescolamentoMiolo3" class="form-label">Descolamento Miolo 3 (g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeDescolamentoMiolo3" name="testeDescolamentoMiolo3" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeDescolamentoCapaExterna" class="form-label">Descolamento Capa Externa (g/m²) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeDescolamentoCapaExterna" name="testeDescolamentoCapaExterna" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeBctCompressao1" class="form-label">BCT/Compressão 1 (kgf) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeBctCompressao1" name="testeBctCompressao1" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeBctCompressao2" class="form-label">BCT/Compressão 2 (kgf) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeBctCompressao2" name="testeBctCompressao2" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeBctCompressao3" class="form-label">BCT/Compressão 3 (kgf) <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeBctCompressao3" name="testeBctCompressao3" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testePinInterno" class="form-label">PIN INTERNO <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testePinInterno" name="testePinInterno" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testePinExterno" class="form-label">PIN EXTERNO <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testePinExterno" name="testePinExterno" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeForcaVinco1" class="form-label">Força de Vinco 1 <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeForcaVinco1" name="testeForcaVinco1" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeForcaVinco2" class="form-label">Força de Vinco 2 <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeForcaVinco2" name="testeForcaVinco2" min="0" step="1" required>
                    </div>
                    <div class="col-lg-4 col-md-6">
                        <label for="testeForcaVinco3" class="form-label">Força de Vinco 3 <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="testeForcaVinco3" name="testeForcaVinco3" min="0" step="1" required>
                    </div>
                </div>

                <div class="d-flex justify-content-end mt-4 gap-2">
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="cancelarFormulario()">
                        Cancelar
                    </button>
                    <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar" >
                        <i id="loadingIcon" class="spinner-border spinner-border-sm me-1" style="display: none;"></i>
                        <i id="checkIcon" class="bi bi-check-lg me-1"></i>
                        <span>Salvar</span>
                    </button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
