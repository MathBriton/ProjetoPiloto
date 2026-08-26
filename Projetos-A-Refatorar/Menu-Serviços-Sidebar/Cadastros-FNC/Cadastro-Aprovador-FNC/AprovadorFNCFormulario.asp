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
        .container {
            max-width: 90%;
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

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>

    <script>
        function anexarEventos() {
            $('#formularioAprovador').on('submit', function(e) {
                e.preventDefault();
                salvarFormulario();
            });
        }

        function navegarParaGrade() {
            window.location.href = `AprovadorFNCGrade.asp`;
        }

        function carregarUsuarios() {
            const $select = $('#idUsuario');
            $select.empty().append('<option value="">Selecione o usuário</option>');

            const params = {
                acao: 'listarUsuarios'
            };

            $.ajax({
                url: 'AprovadorFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar usuários: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar usuários:', error);
                }
            });
        }

        function carregarSetores() {
            const $select = $('#setores');
            $select.empty().append('<option value="">Selecione o usuário</option>');

            const params = {
                acao: 'listarSetores'
            };

            $.ajax({
                url: 'AprovadorFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar setores: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar setores:', error);
                }
            });
        }

        function salvarFormulario() {
            const $btnSalvar = $('#btnSalvar');
            const $loadingIcon = $btnSalvar.find('#loadingIcon');
            const $checkIcon = $btnSalvar.find('#checkIcon');

            $btnSalvar.prop('disabled', true);
            $checkIcon.css('display', 'none');
            $loadingIcon.css('display', 'inline-block');

            const dadosFormulario = {
                acao: 'criar',
                idUsuario: $('#idUsuario').val(),
            };

            $.ajax({
                url: 'AprovadorFNCRepositorio.asp',
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
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        $(document).ready(function() {
            inicializarSelect2();
            carregarUsuarios()
            carregarSetores();
            anexarEventos();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex">
            <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                <i class="bi bi-arrow-left"></i> 
            </button>
            <h3 class="my-4">Novo Aprovador de FNC</h3>
        </div>

        <div class="card p-3">
            <form id="formularioAprovador" name="formularioAprovador">
                <div class="row g-3">
                    <div class="col-12">
                        <label required for="idUsuario" class="form-label">Usuário <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="idUsuario" id="idUsuario">
                            <option value="">Selecione um usuário</option>
                        </select>
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
