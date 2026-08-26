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
        .just-validate-error-label {
            color: #f87171 !important;
            font-size: .8rem;
            margin-top: 3px;
            display: block;
        }
        .just-validate-error-field {
            border-color: #f87171 !important;
        }
        .just-validate-error-field:focus {
            border-color: #ff5353 !important;
            outline: none;
            box-shadow: none !important;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;

        function exibirLoading(exibir) {            
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function anexarEventos() {
            $('#idUsuario').on('change', function() {
                campoIdUsuarioAlterado();
            });

            $('#codPlanta').on('change', function() {
                campoPlantaAlterado();
            });

            $('#formularioAtendente').on('submit', function(e) {
                e.preventDefault();
                salvarFormulario();
            });
        }

        function campoIdUsuarioAlterado() {
            const $select = $('#idUsuario').find(':selected');

            $('#nome').val($select.data('nome') || '');
            $('#email').val($select.data('email') || '');
        }

        function navegarParaGrade() {
            window.location.href = `AtendenteGrade.asp`;
        }

        function carregarUsuarios() {
            const $select = $('#idUsuario');
            $select.empty().append('<option value="">Selecione o usuário</option>');

            const params = {
                acao: 'listar'
            };

            $.ajax({
                url: '../Usuario/UsuarioRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar usuários: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $('<option>', { value: item.id, text: item.nome })
                            .data("nome",     item.nome)
                            .data("email", item.email)
                            .appendTo($select);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar usuários:', error);
                }
            });
        }

        function campoPlantaAlterado() {
            const codPlanta = $('#codPlanta').val();
            carregarDepartamentos(codPlanta);
        }

        function carregarDepartamentos(codPlanta) {
            if (!codPlanta) {
                return;
            }

            const $select = $('#idsDepartamento');
            $select.empty();

            const params = {
                acao: 'listar',
                codPlanta: codPlanta
            };

            $.ajax({
                url: '../Departamento/DepartamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar departamentos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar departamentos:', error);
                }
            });
        }

        function obterAtendente() {
            $.ajax({
                url: 'AtendenteRepositorio.asp',
                type: 'GET',
                data: { 
                    acao: 'obterPorId',
                    id: id 
                },
                dataType: 'json',
                success: (response) => {
                    if(!response.sucesso) {
                        Toast.erro('Erro ao carregar atendente: ' + response.mensagem);
                        return;
                    }

                    preencherFormularioEdicao(response.dados); 
                },
                error: (xhr, status, error) => {
                    console.error('Erro:', error);
                    Toast.erro('Erro ao carregar dados da atendente');
                }
            });
        }

        function preencherFormularioEdicao(dados) {      
            $('#nome').val(dados.nome);
            $('#email').val(dados.email);

            var p1 = carregarDepartamentos(dados.codPlanta);

            $.when(p1).then(function() {
                setTimeout(() => {
                    const idsDepartamento = dados.departamentos.map(departamento => departamento.Id);

                    $('#idUsuario').val(dados.idUsuario).prop('disabled', true).trigger('change.select2');
                    $('#codPlanta').val(dados.codPlanta).trigger('change.select2');
                    $('#idsDepartamento').val(idsDepartamento).trigger('change.select2');
                }, 500);
            });
        }

        async function validarFormulario() {
            const isValid = await validacao.revalidate();

            if (!isValid) {
                const errorsParaExibir = {};
                Object.entries(validacao.fields).forEach(([index, campo]) => {
                    if (campo.isValid) return;
                    const primeiroErro = campo.rules.find(rule => !rule.isValid);
                    if (primeiroErro) {
                        errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                    }
                });
                validacao.showErrors(errorsParaExibir);
            }

            return isValid;
        }

        async function salvarFormulario() {
            if (!await validarFormulario()) return;

            const dadosFormulario = {
                acao: id ? 'editar' : 'criar',
                codPlanta: $('#codPlanta').val(),
                idUsuario: $('#idUsuario').val(),
                idsDepartamento: JSON.stringify($('#idsDepartamento').val()),
                nome: $('#nome').val(),
                email: $('#email').val(),
            };

            if (id) dadosFormulario.id = id;

            exibirLoading(true)

            $.ajax({
                url: 'AtendenteRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                    
                    Toast.success(response.mensagem);
                    setTimeout(() => navegarParaGrade(), 1000);
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                }, 
                complete: function() {
                    exibirLoading(false)
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

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            validacao = new JustValidate('#formularioAtendente', {
                lockForm: false, 
            });

            validacao
                .addField('#codPlanta', [naoVazio('Planta é obrigatória')])
                .addField('#idUsuario', [naoVazio('Usuário é obrigatório')])
                .addField('#nome', [naoVazio('Nome é obrigatório')])
                .addField('#email', [naoVazio('Email é obrigatório')])
                .addField('#idsDepartamento', [naoVazio('Departamento é obrigatório')]);
        }

        function inicializarSelect2() {
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999" });
        }

        $(document).ready(function() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id') || 0;
            
            if (id) {
                $('#tituloPagina').text(`Editar Atendente de Chamado #${id}`);
                obterAtendente();
            } else {
                $('#tituloPagina').text('Novo Atendente de Chamado');
            }

            carregarUsuarios();
            inicializarValidacoes();
            inicializarSelect2();
            anexarEventos();
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex my-4">
            <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                <i class="bi bi-arrow-left"></i> 
            </button>
            <h3 id="tituloPagina">Atendente</h3>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="card p-3">
            <form id="formularioAtendente" name="formularioAtendente">
                <div class="row g-3">
                    <div class="col-12">
                        <label for="codPlanta" class="form-label">Planta <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="codPlanta" id="codPlanta">
                            <option value="">Selecione a planta</option>
                            <option value="00">Pirapetinga</option>
                            <option value="04">Uberaba</option>
                            <option value="05">Bento</option>
                            <option value="06">Maranguape</option>
                            <option value="07">Saquarema</option>
                            <option value="08">Araçatuba</option>
                            <option value="09">Blumenau</option>
                            <option value="10">Pacajus</option>
                            <option value="11">Porto Feliz</option>
                            <option value="12">Três Barras</option>
                            <option value="13">Florestal</option>
                        </select>
                    </div>
                    <div class="col-12">
                        <label for="idUsuario" class="form-label">Usuário <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="idUsuario" id="idUsuario">
                            <option value="">Selecione o usuário</option>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label for="nome" class="form-label">Nome <span class="required-field">*</span></label>
                        <input disabled type="text" class="form-control form-control-sm" id="nome" name="nome" placeholder="Digite o nome">
                    </div>
                    <div class="col-md-6">
                        <label for="email" class="form-label">Email <span class="required-field">*</span></label>
                        <input disabled type="email" class="form-control form-control-sm" id="email" name="email" placeholder="Digite o email">
                    </div>
                    <div class="col-12">
                        <label for="idsDepartamento" class="form-label">Departamentos <span class="required-field">*</span></label>
                        <select multiple class="form-select form-select-sm select2" name="idsDepartamento" id="idsDepartamento">
                            <option value="">Selecione os departamentos</option>
                        </select>
                    </div>
                </div>

                <div class="d-flex justify-content-end mt-4 gap-2">
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="cancelarFormulario()">
                        Cancelar
                    </button>
                    <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar" >
                        <i id="checkIcon" class="bi bi-check-lg me-1"></i>
                        <span>Salvar</span>
                    </button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
