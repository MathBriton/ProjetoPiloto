<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Solicitar Alteração de Pedido</title>
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

        #divMotivoReprovacaoAmostra, #divCamposFerramental, #divCamposAmostra {
            display: none;
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
        let solicitacao = {}

        function obterSolicitacaoProjeto() {
            const params = {
                acao: 'obterPorId',
                id: id,
                codPlanta: codPlanta
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    solicitacao = response.dados
                    $('#idTitulo').text(solicitacao.id);
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao obter solicitação:', error);
                }
            });
        }


        function carregarListaStatusAndamento() {
            const $select = $('#tipoAlteracao');
            $select.empty().append('<option value="">Selecione o tipo de alteração...</option>');

            const params = {
                acao: 'listarStatusAndamento'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar tipos de alteração: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar tipos de alteração:', error);
                }
            });
        }

        function carregarMotivosReprovacaoAmostra() {
            const $select = $('#motivoReprovacaoAmostra');
            $select.empty().append('<option value="">Selecione o motivo...</option>');

            const params = {
                acao: 'listarMotivosReprovacaoAmostra'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar motivos de reprovação: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar motivos de reprovação:', error);
                }
            });
        }

        function campoTipoAlteracaoAlterado() {
            const valorTipoAlteracao = $('#tipoAlteracao').val();

            // Esconde Campos
            $('#divMotivoReprovacaoAmostra').css('display', 'none');
            $('#divCamposFerramental').css('display', 'none');
            $('#divCamposAmostra').css('display', 'none');

            // Marca campos como obrigatório = false
            $('#motivoReprovacaoAmostra').val('').prop('required', false);
            $('#dataConfeccaoFerramental').val('').prop('required', false);
            $('#valorCliche').val('').prop('required', false);
            $('#valorMatriz').val('').prop('required', false);
            $('#dataEnvioAmostra').val('').prop('required', false);
            $('#quemEnviouAmostra').val('').prop('required', false);

            // Amostra Reprovada
            if (valorTipoAlteracao == 'R') {
                $('#divMotivoReprovacaoAmostra').css('display', 'block');
                $('#motivoReprovacaoAmostra').prop('required', true);
                carregarMotivosReprovacaoAmostra();
            }
            
            // Confeccao Ferramental
            if (valorTipoAlteracao == 'J') {
                $('#divCamposFerramental').css('display', 'flex');
                $('#dataConfeccaoFerramental').prop('required', true);
                $('#valorCliche').prop('required', true);
                $('#valorMatriz').prop('required', true);
            }
            
            // Amostra Enviada
            if (valorTipoAlteracao == 'E') {
                $('#divCamposAmostra').css('display', 'flex');
                $('#dataEnvioAmostra').prop('required', true);
                $('#quemEnviouAmostra').prop('required', true);
            } 
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
                acao: 'alterarStatusAndamento',
                id: id,
                codPlanta: codPlanta,
                status: $('#tipoAlteracao').val() || '',
                observacoes: $('#observacoes').val() || ''
            };

            if (dadosFormulario.status == 'E') {
                dadosFormulario.dataEnvioAmostra = $('#dataEnvioAmostra').val() || ''
                dadosFormulario.quemEnviouAmostra = $('#quemEnviouAmostra').val() || ''
            }

            if (dadosFormulario.status == 'J') {
                dadosFormulario.dataConfeccaoFerramental = $('#dataConfeccaoFerramental').val() || ''
                dadosFormulario.valorCliche = $('#valorCliche').val() || 0
                dadosFormulario.valorMatriz = $('#valorMatriz').val() || 0
            }

            if (dadosFormulario.status == 'R') {
                dadosFormulario.motivoReprovacaoAmostra = $('#motivoReprovacaoAmostra').val() || ''
            }

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
                    $btnSalvar.prop('disabled', false);
                    $loadingIcon.css('display', 'none');
                    $checkIcon.css('display', 'inline-block');

                    setTimeout(() => {
                        navegarParaGrade()
                    }, 1000); 
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

            $('#tipoAlteracao').on('change', campoTipoAlteracaoAlterado);
        }

        $(document).ready(function() {
            inicializarSelect2();
            carregarListaStatusAndamento();
            anexarEventos();
           
            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');
            const paramCodPlanta = urlParams.get('codPlanta');
            
            if (paramId && paramCodPlanta) {
                id = paramId;
                codPlanta = paramCodPlanta;

                obterSolicitacaoProjeto()
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
            <h3 class="mb-0">Alteração de Andamento de Projeto (SPD) #<span id="idTitulo"></span></h3>
        </div>

        <div class="card p-3">
            <form id="formularioSolicitacao" name="formularioSolicitacao">
                <div class="row g-3">
                    <div class="col-12">
                        <label for="tipoAlteracao" class="form-label">Tipo de Alteração <span class="required-field">*</span></label>
                        <select required class="form-select form-select-sm select2" id="tipoAlteracao" name="tipoAlteracao">
                            <option value="">Selecione o tipo...</option>
                        </select>
                    </div>

                    <div class="col-12" id="divMotivoReprovacaoAmostra">
                        <label for="motivoReprovacaoAmostra" class="form-label">Motivo Reprovação Amostra <span class="required-field">*</span></label>
                        <select required class="form-select form-select-sm select2" id="motivoReprovacaoAmostra" name="motivoReprovacaoAmostra">
                            <option value="">Selecione o motivo...</option>
                        </select>
                    </div>

                    <!-- Campos Ferramental -->
                    <div class="row p-0 m-0 g-3" id="divCamposFerramental">
                        <div class="col-12">
                            <label for="dataConfeccaoFerramental" class="form-label">Data de Confecção do Ferramental <span class="required-field">*</span> </label>
                            <input required type="date" class="form-control form-control-sm" id="dataConfeccaoFerramental" name="dataConfeccaoFerramental">
                        </div>

                        <div class="col-md-6">
                            <label for="valorCliche" class="form-label">Valor Clichê <span class="required-field">*</span> </label>
                            <input required type="number" class="form-control form-control-sm" id="valorCliche" name="valorCliche">
                        </div>

                        <div class="col-md-6">
                            <label for="valorMatriz" class="form-label">Valor Matriz <span class="required-field">*</span> </label>
                            <input required type="number" class="form-control form-control-sm" id="valorMatriz" name="valorMatriz">
                        </div>
                    </div>

                    <!-- Campos Envio Amostra -->
                    <div class="row p-0 m-0 g-3" id="divCamposAmostra">
                        <div class="col-md-6">
                            <label for="dataEnvioAmostra" class="form-label">Data de Envio da Amostra</label>
                            <input required type="date" class="form-control form-control-sm" id="dataEnvioAmostra" name="dataEnvioAmostra">
                        </div>

                        <div class="col-md-6">
                            <label for="quemEnviouAmostra" class="form-label">Quem Enviou a Amostra</label>
                            <input required type="text" class="form-control form-control-sm" id="quemEnviouAmostra" name="quemEnviouAmostra">
                        </div>
                    </div>

                    <div class="col-12">
                        <label for="observacoes" class="form-label">Observações</label>
                        <textarea class="form-control form-control-sm" id="observacoes" name="observacoes" rows="4"></textarea>
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
