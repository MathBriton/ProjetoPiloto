
<%
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "nivel", Session("NIVEL")
usuarioLogado.Add "cod", Session("CODUSER")
usuarioLogado.Add "email", Session("EMAIL")

%>
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
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
    <link rel="stylesheet" href="../Compartilhado/estilos.css?v=<%=Now()%>">

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
        .just-validate-error-label {
            color: #dc3545 !important;
            font-size: .8rem;
            margin-top: 3px;
            display: block;
        }
        .just-validate-error-field {
            border-color: #dc3545 !important;
        }
        .just-validate-error-field:focus {
            border-color: #ff5353 !important;
            outline: none;
            box-shadow: none !important;
        }
        .success-message {
            background-color: #d1e7dd;
            border: 1px solid #198754;
            color: #0f5132;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            margin: 40px auto;
            max-width: 600px;
        }
        .success-message i {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        .success-message h4 {
            margin-bottom: 10px;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        window.FormularioChamado = (function () {
            var registroPendente = null;

            function registrar(modulo) {
                registroPendente = modulo;
            }

            // Consumido pelo formulário principal logo apos o getScript do arquivo de tipo.
            function obterRegistro() {
                var m = registroPendente;
                registroPendente = null;
                return m;
            }

            return { registrar: registrar, obterRegistro: obterRegistro };
        })();

        // Estado do formulário principal
        var validacao = null;
        var moduloAtual = null;       // modulo do tipo atualmente renderizado
        var tipoSelecionado = null;   // { id, nome, caminho }
        var cacheModulos = {};        // id do tipo -> modulo ja carregado

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function anexarEventos() {
            $('#formularioChamado').on('submit', function (e) {
                e.preventDefault();
                salvarFormulario();
            });

            $('#seletorTipo').on('change', function () {
                var id = $(this).val();
                if (!id) {
                    limparTipo();
                    return;
                }
                var tipo = (window.__tiposChamado || []).find(function (t) {
                    return String(t.id) === String(id);
                });
                if (tipo) selecionarTipo(tipo);
            });
        }

        // --- Tipos de chamado: carregados via AJAX (sem hardcode) -----------
        function carregarTipos() {
            var $select = $('#seletorTipo');
            $select.empty().append('<option value="">Selecione um tipo</option>');

            return $.ajax({
                url: '../Tipo/TipoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listar' },
                dataType: 'json',
                success: function (response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tipos: ' + response.mensagem);
                        return;
                    }
                    window.__tiposChamado = response.dados || [];
                    window.__tiposChamado.forEach(function (item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function (xhr, status, error) {
                    console.error('Erro ao carregar tipos:', error);
                }
            });
        }

        // --- Carregamento sob demanda do arquivo do tipo --------------------
        function carregarModuloTipo(tipo) {
            if (cacheModulos[tipo.id]) {
                return Promise.resolve(cacheModulos[tipo.id]);
            }

            return $.getScript(tipo.caminho).then(function () {
                var modulo = FormularioChamado.obterRegistro();
                if (!modulo || typeof modulo.render !== 'function') {
                    throw new Error('O arquivo do tipo nao registrou um modulo valido: ' + tipo.caminho);
                }
                cacheModulos[tipo.id] = modulo;
                return modulo;
            });
        }

        function limparTipo() {
            if (moduloAtual && typeof moduloAtual.destruir === 'function') {
                try { moduloAtual.destruir(); } catch (e) { console.warn(e); }
            }
            moduloAtual = null;
            tipoSelecionado = null;
            $('#detalhamento').empty();
            $('#campos-comuns').hide();
        }

        function selecionarTipo(tipo) {
            exibirLoading(true);

            // descarrega o tipo anterior antes de montar o novo
            if (moduloAtual && typeof moduloAtual.destruir === 'function') {
                try { moduloAtual.destruir(); } catch (e) { console.warn(e); }
            }
            $('#detalhamento').empty();

            carregarModuloTipo(tipo)
                .then(function (modulo) {
                    moduloAtual = modulo;
                    tipoSelecionado = tipo;
                    modulo.render($('#detalhamento'), { tipo: tipo, usuario: window.__usuarioLogado || {} });
                    $('#campos-comuns').show();
                })
                .catch(function (err) {
                    console.error(err);
                    Toast.error('Nao foi possivel carregar o formulario deste tipo.');
                    limparTipo();
                })
                .always(function () {
                    exibirLoading(false);
                });
        }

        function carregarDepartamentos() {
            const $select = $('#idDepartamento');
            $select.empty().append('<option value="">Selecione um departamento</option>');

            const params = { acao: 'listar' };

            return $.ajax({
                url: '../Departamento/DepartamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function (response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar departamentos: ' + response.mensagem);
                        return;
                    }
                    response.dados.forEach(function (item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function (xhr, status, error) {
                    console.error('Erro ao carregar departamentos:', error);
                }
            });
        }

        function carregarCategorias(idDepartamento) {
            const $select = $('#idCategoria');
            $select.empty().append('<option value="">Selecione uma categoria</option>');

            if (!idDepartamento) { return; }

            const params = { acao: 'listar', idDepartamento: idDepartamento };

            return $.ajax({
                url: '../Categoria/CategoriaRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function (response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar categorias: ' + response.mensagem);
                        return;
                    }
                    response.dados.forEach(function (item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function (xhr, status, error) {
                    console.error('Erro ao carregar categorias:', error);
                }
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

            if (!moduloAtual || !tipoSelecionado) {
                Toast.warning('Selecione um tipo de chamado');
                return;
            }

            const anexos = uploadAnexos ? uploadAnexos.getFiles() : [];

            // A validacao especifica do tipo fica dentro do proprio arquivo de tipo.
            const faltas = (typeof moduloAtual.validar === 'function') ? (moduloAtual.validar() || []) : [];
            if (faltas.length) {
                Toast.warning(faltas[0] || 'Preencha os campos faltantes');
                return;
            }

            // O detalhamento continua sendo um JSON, obtido aqui no principal
            // a partir do coletar() do arquivo de tipo.
            const detalhamento = (typeof moduloAtual.coletar === 'function') ? moduloAtual.coletar() : {};

            const dadosFormulario = {
                acao: 'criar',
                tipo: tipoSelecionado.id,
                nome: $('#nome').val(),
                email: $('#email').val(),
                telefone: $('#telefone').val(),
                idDepartamento: $('#idDepartamento').val(),
                titulo: $('#titulo').val(),
                descricao: $('#descricao').val(),
                anexos: JSON.stringify(anexos),
                detalhamento: JSON.stringify(detalhamento)
            };

            exibirLoading(true);

            $.ajax({
                url: 'ChamadoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function (response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                    Toast.success(response.mensagem);
                    mostrarMensagemSucesso();
                },
                error: function (xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function () {
                    exibirLoading(false);
                }
            });
        }

        function mostrarMensagemSucesso() {
            $('body').html(`
                <div class="success-message">
                    <i class="bi bi-check-circle"></i>
                    <h4>Chamado criado com sucesso!</h4>
                    <p>Em instantes você receberá um e-mail com mais detalhes</p>
                    <div class="mt-3"><a href="ChamadoFormulario.asp">Clique aqui para criar um novo chamado</a></div>
                </div>
            `);
        }

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const maiorQueZero = (msg) => ({
                validator: (value) => value > 0,
                errorMessage: msg
            });

            validacao = new JustValidate('#formularioChamado', {
                lockForm: false,
            });

            validacao
                .addField('#nome', [naoVazio('Nome é obrigatório')])
                .addField('#email', [naoVazio('E-mail é obrigatório')])
                .addField('#telefone', [naoVazio('Telefone é obrigatório')])
                .addField('#idDepartamento', [naoVazio('Departamento é obrigatório')])
                .addField('#titulo', [naoVazio('Assunto é obrigatório')])
                .addField('#descricao', [naoVazio('Descrição é obrigatória')]);
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999"
            });
        }

        $(document).ready(function () {
            window.__usuarioLogado = {
                id: '<%=usuarioLogado("id")%>',
                nome: '<%=usuarioLogado("nome")%>',
                email: '<%=usuarioLogado("email")%>'
            };

            inicializarValidacoes();
            inicializarSelect2();
            anexarEventos();

            carregarDepartamentos();
            carregarTipos();

            const usuarioEstaLogado = window.__usuarioLogado.id;
            if (usuarioEstaLogado) {
                $('#nome').val(window.__usuarioLogado.nome);
                $('#email').val(window.__usuarioLogado.email);
            }
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex my-4">
            <h3>Novo Chamado</h3>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <form id="formularioChamado" name="formularioChamado">
            <div class="mb-4">
                <label for="seletorTipo" class="form-label">Tipo <span class="required-field">*</span></label>
                <select class="form-select form-select-sm select2" name="seletorTipo" id="seletorTipo">
                    <option value="">Selecione um tipo</option>
                </select>
            </div>

            <div style="display: none;" id="campos-comuns" class="secao">
                <div class="secao-head">
                    <span class="ic"> <i class="fa-solid fa-user"></i> </span>
                    <h3>Dados Gerais</h3>
                </div>

                <div class="row g-3 p-3">
                    <div class="col-12 col-md-6">
                        <label for="nome" class="form-label">Nome Completo <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="nome" name="nome" placeholder="Insira o nome completo">
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="email" class="form-label">Email <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="email" name="email" placeholder="Insira o email">
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="telefone" class="form-label">Telefone <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="telefone" name="telefone" placeholder="Insira o telefone">
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="idDepartamento" class="form-label">Departamento <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" name="idDepartamento" id="idDepartamento">
                            <option value="">Selecione o Departamento</option>
                        </select>
                    </div>
                    <div class="col-12">
                        <label for="titulo" class="form-label">Assunto <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="titulo" name="titulo" placeholder="Insira o assunto">
                    </div>
                    <div class="col-12">
                        <label for="centroCusto" class="form-label">Centro de Custo (opcional)</label>
                        <input type="text" class="form-control form-control-sm" id="centroCusto" name="centroCusto" placeholder="Insira o centro de custo">
                    </div>
                    <div class="col-12">
                        <label for="descricao" class="form-label">Descrição <span class="required-field">*</span></label>
                        <textarea placeholder="Descreva de maneira geral sua necessidade" class="form-control" id="descricao" name="descricao" rows="4"></textarea>
                    </div>
                    <div class="col-12">
                        <label for="uploadAnexos" class="form-label">Anexos</label>
                        <file-upload id="uploadAnexos" file-code="CHAMADOS" max-files="10" multiple></file-upload>
                    </div>
                </div>
            </div>

            <!-- Container onde o arquivo de tipo renderiza o formulario especifico -->
            <div id="detalhamento" class="d-flex flex-column gap-3 mt-3"></div>

            <div class="d-flex justify-content-end mt-4 gap-2">
                <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar">
                    <i id="checkIcon" class="bi bi-check-lg me-1"></i>
                    <span>Salvar</span>
                </button>
            </div>
        </form>
    </div>
</body>
</html>
